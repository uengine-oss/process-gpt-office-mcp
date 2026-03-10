import os
from xml.etree import ElementTree as ET

from ..config import SMALL_CELL_HEIGHT_MM, SMALL_CELL_WIDTH_MM
from ..models import TextNode, _HWP_UNITS_PER_MM
from .style_mapper import resolve_style_for_runs, summarize_style, StyleMaps
from .xml_utils import tag, ns, register_namespaces


def scan_header_charpr(header_path: str) -> tuple[dict[str, dict], list[str]]:
    if not os.path.exists(header_path):
        return {}, []

    register_namespaces(header_path)
    tree = ET.parse(header_path)
    root = tree.getroot()
    charpr_map: dict[str, dict] = {}
    warnings: list[str] = []

    for elem in root.iter():
        if tag(elem) != "charPr":
            continue
        charpr_id = None
        for k, v in elem.attrib.items():
            if "id" in k.lower():
                charpr_id = v
                break
        if not charpr_id:
            continue

        info = dict(elem.attrib)
        for child in elem:
            ctag = tag(child)
            if ctag in ("spacing", "condense", "sz", "height"):
                if child.attrib:
                    for k, v in child.attrib.items():
                        info[f"{ctag}_{k}"] = v
                elif child.text:
                    info[ctag] = child.text
        charpr_map[charpr_id] = info

        spacing_val = info.get("spacing") or info.get("spacing_val")
        condense_val = info.get("condense") or info.get("condense_val")
        if spacing_val:
            try:
                if int(spacing_val) < 0:
                    warnings.append(f"charPrID={charpr_id} spacing={spacing_val}")
            except ValueError:
                pass
        if condense_val:
            try:
                if int(condense_val) != 0:
                    warnings.append(f"charPrID={charpr_id} condense={condense_val}")
            except ValueError:
                warnings.append(f"charPrID={charpr_id} condense={condense_val}")

    return charpr_map, warnings


def collect_runs_and_texts(elem):
    runs = []
    t_elems = []

    for child in elem:
        if tag(child) == "tbl":
            continue
        if tag(child) == "run":
            run_has_t = False
            for gc in child:
                if tag(gc) == "t":
                    t_elems.append(gc)
                    run_has_t = True
            runs.append((child, run_has_t))
        elif tag(child) == "p":
            sub_runs, sub_ts = collect_runs_and_texts(child)
            runs.extend(sub_runs)
            t_elems.extend(sub_ts)
        elif tag(child) == "subList":
            for sub_p in child:
                if tag(sub_p) == "p":
                    sub_runs, sub_ts = collect_runs_and_texts(sub_p)
                    runs.extend(sub_runs)
                    t_elems.extend(sub_ts)
        else:
            sub_runs, sub_ts = collect_runs_and_texts(child)
            runs.extend(sub_runs)
            t_elems.extend(sub_ts)

    return runs, t_elems


def parse_section(
    section_path: str,
    style_maps: StyleMaps | None = None,
) -> tuple[list[TextNode], ET.ElementTree, dict, str]:
    register_namespaces(section_path)
    tree = ET.parse(section_path)
    root = tree.getroot()
    nodes: list[TextNode] = []
    nid = [0]
    tbl_counter = [0]
    parent_map = {c: p for p in root.iter() for c in p}

    sample_ns = ""
    for elem in root.iter():
        if tag(elem) == "t":
            sample_ns = ns(elem)
            break

    def _get_text(t_elems, strip=True):
        joined = "".join((t.text or "") for t in t_elems)
        return joined.strip() if strip else joined

    def _calc_depth(raw_text: str) -> int:
        if not raw_text:
            return 0
        leading = len(raw_text) - len(raw_text.lstrip(" "))
        return leading // 2

    def _process_table(tbl_elem, tbl_idx):
        for tr in tbl_elem:
            if tag(tr) != "tr":
                continue
            for tc in tr:
                if tag(tc) != "tc":
                    continue
                row = col = -1
                cell_width = 0
                cell_height = 0
                cell_col_span = 1
                cell_row_span = 1
                for cc in tc:
                    if tag(cc) == "cellAddr":
                        for k, v in cc.attrib.items():
                            if "colAddr" in k:
                                col = int(v)
                            if "rowAddr" in k:
                                row = int(v)
                    elif tag(cc) == "cellSz":
                        for k, v in cc.attrib.items():
                            if "width" in k:
                                try:
                                    cell_width = int(v)
                                except ValueError:
                                    pass
                            if "height" in k:
                                try:
                                    cell_height = int(v)
                                except ValueError:
                                    pass
                    elif tag(cc) == "cellSpan":
                        try:
                            cell_col_span = int(cc.attrib.get("colSpan", cell_col_span))
                        except (TypeError, ValueError):
                            pass
                        try:
                            cell_row_span = int(cc.attrib.get("rowSpan", cell_row_span))
                        except (TypeError, ValueError):
                            pass

                runs, t_elems = collect_runs_and_texts(tc)
                raw_text = _get_text(t_elems, strip=False)
                text = _get_text(t_elems, strip=True)
                height_mm = round(cell_height / _HWP_UNITS_PER_MM) if cell_height > 0 else 0
                width_mm = round(cell_width / _HWP_UNITS_PER_MM) if cell_width > 0 else 0
                is_spacer = (
                    (not text) and (not raw_text.strip()) and
                    (
                        (height_mm > 0 and height_mm <= SMALL_CELL_HEIGHT_MM) or
                        (width_mm > 0 and width_mm <= SMALL_CELL_WIDTH_MM)
                    )
                )
                run_elems = [r for r, _ in runs]
                style_refs, style_info = resolve_style_for_runs(run_elems, parent_map, style_maps)
                style_missing = {}
                if style_maps:
                    if style_refs.get("style_id") and style_refs["style_id"] not in style_maps.styles:
                        style_missing["style"] = style_refs["style_id"]
                    if style_refs.get("para_id") and style_refs["para_id"] not in style_maps.paraprs:
                        style_missing["para"] = style_refs["para_id"]
                    if style_refs.get("char_id") and style_refs["char_id"] not in style_maps.charprs:
                        style_missing["char"] = style_refs["char_id"]
                style_summary = summarize_style(style_info)

                nodes.append(TextNode(
                    nid=nid[0], ntype="table_cell", text=text,
                    raw_text=raw_text, depth=0, skip_fill=is_spacer,
                    t_elements=t_elems, run_elements=run_elems,
                    table_idx=tbl_idx, row=row, col=col,
                    cell_width=cell_width,
                    cell_height=cell_height,
                    cell_col_span=cell_col_span,
                    cell_row_span=cell_row_span,
                    style_refs=style_refs,
                    style_info=style_info,
                    style_summary=style_summary,
                    style_missing=style_missing,
                ))
                nid[0] += 1

    def _process_paragraph(p_elem):
        for run in p_elem:
            if tag(run) != "run":
                continue
            for child in run:
                if tag(child) == "tbl":
                    _process_table(child, tbl_counter[0])
                    tbl_counter[0] += 1

        runs, t_elems = collect_runs_and_texts(p_elem)
        run_elems = [r for r, _ in runs]
        style_refs, style_info = resolve_style_for_runs(run_elems, parent_map, style_maps)
        style_missing = {}
        if style_maps:
            if style_refs.get("style_id") and style_refs["style_id"] not in style_maps.styles:
                style_missing["style"] = style_refs["style_id"]
            if style_refs.get("para_id") and style_refs["para_id"] not in style_maps.paraprs:
                style_missing["para"] = style_refs["para_id"]
            if style_refs.get("char_id") and style_refs["char_id"] not in style_maps.charprs:
                style_missing["char"] = style_refs["char_id"]
        style_summary = summarize_style(style_info)

        if runs:
            raw_text = _get_text(t_elems, strip=False)
            text = _get_text(t_elems, strip=True)
            depth = _calc_depth(raw_text)
            skip_fill = (not text) and (not raw_text.strip()) and depth == 0
            nodes.append(TextNode(
                nid=nid[0], ntype="body_text", text=text,
                raw_text=raw_text, depth=depth, skip_fill=skip_fill,
                t_elements=t_elems, run_elements=run_elems,
                style_refs=style_refs,
                style_info=style_info,
                style_summary=style_summary,
                style_missing=style_missing,
            ))
            nid[0] += 1

    for child in root:
        if tag(child) == "p":
            _process_paragraph(child)

    return nodes, tree, parent_map, sample_ns
