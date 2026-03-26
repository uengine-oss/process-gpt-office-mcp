from __future__ import annotations

from xml.etree import ElementTree as ET

from ..models import TextNode, TableSummary, _HWP_UNITS_PER_MM
from .xml_utils import tag


def build_table_summaries(
    tree: ET.ElementTree,
    nodes: list[TextNode],
) -> list[TableSummary]:
    root = tree.getroot()
    summaries: list[TableSummary] = []
    tbl_idx = -1

    local_parent = {c: p for p in root.iter() for c in p}

    for tbl_elem in root.iter():
        if tag(tbl_elem) != "tbl":
            continue
        parent = local_parent.get(tbl_elem)
        while parent is not None:
            if tag(parent) == "tbl":
                break
            parent = local_parent.get(parent)
        else:
            parent = None
        if parent is not None and tag(parent) == "tbl":
            continue

        tbl_idx += 1
        row_cnt = int(tbl_elem.get("rowCnt", "0") or tbl_elem.attrib.get(
            next((k for k in tbl_elem.attrib if "rowCnt" in k), ""), "0"))
        col_cnt = int(tbl_elem.get("colCnt", "0") or tbl_elem.attrib.get(
            next((k for k in tbl_elem.attrib if "colCnt" in k), ""), "0"))
        if row_cnt == 0:
            for k, v in tbl_elem.attrib.items():
                if "rowCnt" in k:
                    try:
                        row_cnt = int(v)
                    except ValueError:
                        pass
        if col_cnt == 0:
            for k, v in tbl_elem.attrib.items():
                if "colCnt" in k:
                    try:
                        col_cnt = int(v)
                    except ValueError:
                        pass

        rows_data = _extract_rows(tbl_elem)
        header_texts = _extract_header_texts(rows_data, max_header_rows=3)
        data_min_w, empty_ratio, bfid_alt = _calc_data_metrics(rows_data, max_header_rows=3)

        summaries.append(TableSummary(
            table_idx=tbl_idx,
            row_cnt=row_cnt,
            col_cnt=col_cnt,
            header_texts=header_texts,
            data_min_width_mm=data_min_w,
            empty_cell_ratio=empty_ratio,
            bfid_alternating=bfid_alt,
        ))

    return summaries


def _extract_rows(tbl_elem: ET.Element) -> list[list[dict]]:
    rows: list[list[dict]] = []
    for child in tbl_elem:
        if tag(child) != "tr":
            continue
        row_cells: list[dict] = []
        for tc in child:
            if tag(tc) != "tc":
                continue
            info = _parse_cell(tc)
            row_cells.append(info)
        rows.append(row_cells)
    return rows


def _parse_cell(tc: ET.Element) -> dict:
    bfid = ""
    for k, v in tc.attrib.items():
        if "borderFillIDRef" in k:
            bfid = v
            break

    col_addr = row_addr = 0
    col_span = row_span = 1
    width = height = 0
    texts: list[str] = []

    for child in tc:
        t = tag(child)
        if t == "cellAddr":
            for k, v in child.attrib.items():
                if "colAddr" in k:
                    try:
                        col_addr = int(v)
                    except ValueError:
                        pass
                if "rowAddr" in k:
                    try:
                        row_addr = int(v)
                    except ValueError:
                        pass
        elif t == "cellSpan":
            for k, v in child.attrib.items():
                if "colSpan" in k:
                    try:
                        col_span = int(v)
                    except ValueError:
                        pass
                if "rowSpan" in k:
                    try:
                        row_span = int(v)
                    except ValueError:
                        pass
        elif t == "cellSz":
            for k, v in child.attrib.items():
                if "width" in k:
                    try:
                        width = int(v)
                    except ValueError:
                        pass
                if "height" in k:
                    try:
                        height = int(v)
                    except ValueError:
                        pass

    for elem in tc.iter():
        if tag(elem) == "t" and elem.text and elem.text.strip():
            texts.append(elem.text.strip())

    return {
        "tc_elem": tc,
        "bfid": bfid,
        "col_addr": col_addr,
        "row_addr": row_addr,
        "col_span": col_span,
        "row_span": row_span,
        "width": width,
        "height": height,
        "texts": texts,
    }


def _extract_header_texts(rows: list[list[dict]], max_header_rows: int = 3) -> list[str]:
    result: list[str] = []
    for row_idx, row in enumerate(rows[:max_header_rows]):
        for cell in row:
            text = " ".join(cell["texts"]) if cell["texts"] else ""
            if not text:
                continue
            cs = cell["col_span"]
            if cs > 1:
                result.append(f"{text}({cs}열병합)")
            else:
                result.append(text)
    return result


def _calc_data_metrics(
    rows: list[list[dict]],
    max_header_rows: int = 3,
) -> tuple[float, float, bool]:
    data_rows = rows[max_header_rows:]
    if not data_rows:
        data_rows = rows[1:] if len(rows) > 1 else rows

    all_widths: list[int] = []
    total_cells = 0
    empty_cells = 0
    bfid_seq: list[str] = []

    for row in data_rows:
        for cell in row:
            total_cells += 1
            all_widths.append(cell["width"])
            if not cell["texts"]:
                empty_cells += 1
            bfid_seq.append(cell["bfid"])

    min_w_mm = 0.0
    if all_widths:
        min_w = min(w for w in all_widths if w > 0) if any(w > 0 for w in all_widths) else 0
        min_w_mm = round(min_w / _HWP_UNITS_PER_MM, 1) if min_w > 0 else 0.0

    empty_ratio = empty_cells / total_cells if total_cells > 0 else 0.0
    bfid_alt = _check_alternating(bfid_seq)

    return min_w_mm, empty_ratio, bfid_alt


def _check_alternating(seq: list[str]) -> bool:
    if len(seq) < 6:
        return False
    uniq = list(dict.fromkeys(seq))
    if len(uniq) != 2:
        return False
    a, b = uniq
    expected = [a if i % 2 == 0 else b for i in range(len(seq))]
    return seq[:10] == expected[:10]
