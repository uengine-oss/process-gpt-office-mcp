import argparse
import html
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}


def _css_escape(value: str) -> str:
    return value.replace("'", "\\'")


def _build_style(style_map: dict) -> str:
    items = [f"{k}:{v}" for k, v in style_map.items() if v]
    return "; ".join(items)


def _hwpunit_to_px(value: str | int | None) -> str | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number >= 4294967295:
        return None
    px = number / 7200 * 96
    return f"{px:.2f}px"


def _mm_to_px(value: str | None) -> str | None:
    if not value:
        return None
    try:
        num = float(value.replace(" mm", ""))
    except ValueError:
        return None
    px = num / 25.4 * 96
    return f"{px:.2f}px"


def _normalize_color(value: str | None) -> str | None:
    if not value:
        return None
    if len(value) == 9 and value.startswith("#"):
        return f"#{value[-6:]}"
    return value


def _bgr_to_rgb(color: str | None) -> str | None:
    if not color or not color.startswith("#") or len(color) != 7:
        return color
    r = color[1:3]
    g = color[3:5]
    b = color[5:7]
    return f"#{b}{g}{r}"


def _normalize_color_bgr(value: str | None) -> str | None:
    return _bgr_to_rgb(_normalize_color(value))


def _parse_header(zipf: zipfile.ZipFile):
    head_xml = zipf.read("Contents/header.xml").decode("utf-8")
    root = ET.fromstring(head_xml)

    fontfaces = {}
    for fontface in root.findall(".//hh:fontface", NS):
        lang = fontface.attrib.get("lang")
        for font in fontface.findall("hh:font", NS):
            fontfaces[(lang, font.attrib.get("id"))] = font.attrib.get("face")

    border_fills = {}
    for bf in root.findall(".//hh:borderFill", NS):
        bid = bf.attrib.get("id")
        edges = {}
        for edge in ("leftBorder", "rightBorder", "topBorder", "bottomBorder"):
            node = bf.find(f"hh:{edge}", NS)
            if node is None:
                continue
            edges[edge] = {
                "type": node.attrib.get("type"),
                "width": _mm_to_px(node.attrib.get("width")),
                "color": _normalize_color_bgr(node.attrib.get("color")),
            }
        fill_color = None
        fill = bf.find("hc:fillBrush", NS)
        if fill is not None:
            win = fill.find("hc:winBrush", NS)
            if win is not None:
                color = _normalize_color_bgr(win.attrib.get("faceColor"))
                if color and color.lower() not in ("#ffffff", "#ffffffff"):
                    fill_color = color
        border_fills[bid] = {"edges": edges, "fill": fill_color}

    char_styles = {}
    for char_pr in root.findall(".//hh:charPr", NS):
        char_id = char_pr.attrib.get("id")
        height = char_pr.attrib.get("height")
        color = char_pr.attrib.get("textColor")
        font_ref = char_pr.find("hh:fontRef", NS)
        font_family = None
        if font_ref is not None:
            hangul_id = font_ref.attrib.get("hangul")
            font_family = (
                fontfaces.get(("HANGUL", hangul_id))
                or fontfaces.get(("LATIN", hangul_id))
            )

        style_map = {}
        if height and height.isdigit():
            size_pt = int(height) / 100
            style_map["font-size"] = f"{size_pt:g}pt"
        if color:
            style_map["color"] = _normalize_color_bgr(color)
        if font_family:
            style_map["font-family"] = f"'{_css_escape(font_family)}'"
        shade = _normalize_color_bgr(char_pr.attrib.get("shadeColor"))
        if shade and shade.lower() not in ("#ffffff", "#ffffffff"):
            style_map["background-color"] = shade

        if char_pr.find("hh:bold", NS) is not None:
            style_map["font-weight"] = "700"
        if char_pr.find("hh:italic", NS) is not None:
            style_map["font-style"] = "italic"
        underline = char_pr.find("hh:underline", NS)
        if underline is not None:
            style_map["text-decoration-line"] = "underline"
            style_map["text-decoration-color"] = _normalize_color_bgr(
                underline.attrib.get("color")
            )
            shape = underline.attrib.get("shape")
            if shape == "DASHED":
                style_map["text-decoration-style"] = "dashed"
            elif shape == "DOTTED":
                style_map["text-decoration-style"] = "dotted"
            else:
                style_map["text-decoration-style"] = "solid"
        if char_pr.find("hh:strikeout", NS) is not None:
            style_map["text-decoration-line"] = "line-through"

        char_styles[char_id] = style_map

    para_styles = {}
    for para_pr in root.findall(".//hh:paraPr", NS):
        para_id = para_pr.attrib.get("id")
        align = para_pr.find("hh:align", NS)
        style_map = {}
        if align is not None:
            horiz = align.attrib.get("horizontal", "LEFT")
            if horiz == "CENTER":
                style_map["text-align"] = "center"
            elif horiz == "RIGHT":
                style_map["text-align"] = "right"
            else:
                style_map["text-align"] = "left"
        line_spacing = para_pr.find("hh:lineSpacing", NS)
        if line_spacing is not None:
            if line_spacing.attrib.get("type") == "PERCENT":
                value = line_spacing.attrib.get("value")
                if value and value.isdigit():
                    style_map["line-height"] = f"{int(value)}%"
            elif line_spacing.attrib.get("type") == "FIXED":
                value = _hwpunit_to_px(line_spacing.attrib.get("value"))
                if value:
                    style_map["line-height"] = value
        border = para_pr.find("hh:border", NS)
        if border is not None:
            border_fill = border_fills.get(border.attrib.get("borderFillIDRef"))
            if border_fill:
                style_map.update(_border_css(border_fill))
            left = _hwpunit_to_px(border.attrib.get("offsetLeft"))
            right = _hwpunit_to_px(border.attrib.get("offsetRight"))
            top = _hwpunit_to_px(border.attrib.get("offsetTop"))
            bottom = _hwpunit_to_px(border.attrib.get("offsetBottom"))
            if any([left, right, top, bottom]):
                style_map["padding"] = " ".join(
                    [top or "0", right or "0", bottom or "0", left or "0"]
                )
        para_styles[para_id] = style_map

    styles = {}
    styles_root = root.find(".//hh:styles", NS)
    if styles_root is not None:
        for style in styles_root.findall("hh:style", NS):
            styles[style.attrib.get("id")] = {
                "type": style.attrib.get("type"),
                "paraPrIDRef": style.attrib.get("paraPrIDRef"),
                "charPrIDRef": style.attrib.get("charPrIDRef"),
            }

    return char_styles, para_styles, border_fills, styles


def _merge_styles(base: dict | None, override: dict | None) -> dict:
    merged = {}
    if base:
        merged.update(base)
    if override:
        merged.update(override)
    return merged


def _style_dict_to_key(style_map: dict) -> tuple:
    return tuple(sorted(style_map.items()))


def _collect_styled_chars(paragraph, char_styles, base_char_id):
    styled_chars = []
    for run in paragraph.findall("hp:run", NS):
        if run.find("hp:tbl", NS) is not None:
            continue
        run_char_id = run.attrib.get("charPrIDRef") or base_char_id
        run_style_map = _merge_styles(
            char_styles.get(base_char_id), char_styles.get(run_char_id)
        )
        for t in run.findall("hp:t", NS):
            text = t.text or ""
            for ch in text:
                styled_chars.append((ch, run_style_map))
    return styled_chars


def _extract_box_style(para_style: dict) -> dict:
    box_style = {}
    for key, value in para_style.items():
        if key.startswith("border-") or key.startswith("padding") or key.startswith(
            "background"
        ):
            box_style[key] = value
    return box_style


def _render_lineseg_paragraph(
    paragraph, char_styles, para_styles, border_fills, styles
):
    blocks = []
    style_id = paragraph.attrib.get("styleIDRef")
    base_para_id = None
    base_char_id = None
    if style_id and style_id in styles and styles[style_id]["type"] == "PARA":
        base_para_id = styles[style_id].get("paraPrIDRef")
        base_char_id = styles[style_id].get("charPrIDRef")

    para_pr = paragraph.attrib.get("paraPrIDRef") or base_para_id
    para_style = _merge_styles(para_styles.get(base_para_id), para_styles.get(para_pr))

    line_seg_array = paragraph.find("hp:linesegarray", NS)
    if line_seg_array is None:
        return []
    line_segs = line_seg_array.findall("hp:lineseg", NS)
    if not line_segs:
        return []

    styled_chars = _collect_styled_chars(paragraph, char_styles, base_char_id)
    text_len = len(styled_chars)

    def seg_range(idx):
        start = int(line_segs[idx].attrib.get("textpos", 0))
        end = text_len
        if idx + 1 < len(line_segs):
            end = int(line_segs[idx + 1].attrib.get("textpos", text_len))
        return max(start, 0), max(end, 0)

    min_left = None
    min_top = None
    max_right = None
    max_bottom = None

    for idx, seg in enumerate(line_segs):
        start, end = seg_range(idx)
        span_chunks = []
        if start < end and start < text_len:
            slice_chars = styled_chars[start : min(end, text_len)]
            current_style = None
            current_text = []
            for ch, style in slice_chars:
                key = _style_dict_to_key(style)
                if current_style is None:
                    current_style = key
                    current_text = [ch]
                elif key == current_style:
                    current_text.append(ch)
                else:
                    style_str = _build_style(dict(current_style))
                    span_chunks.append(
                        f"<span style=\"{style_str}\">{html.escape(''.join(current_text))}</span>"
                    )
                    current_style = key
                    current_text = [ch]
            if current_text:
                style_str = _build_style(dict(current_style))
                span_chunks.append(
                    f"<span style=\"{style_str}\">{html.escape(''.join(current_text))}</span>"
                )
        else:
            span_chunks.append("&nbsp;")

        left = _hwpunit_to_px(seg.attrib.get("horzpos")) or "0px"
        top = _hwpunit_to_px(seg.attrib.get("vertpos")) or "0px"
        width = _hwpunit_to_px(seg.attrib.get("horzsize")) or "0px"
        height = _hwpunit_to_px(seg.attrib.get("vertsize")) or "0px"
        text_height = _hwpunit_to_px(seg.attrib.get("textheight"))
        line_style = {
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }
        if text_height:
            line_style["line-height"] = text_height
        line_style_str = _build_style(line_style)
        blocks.append(
            f"<div class=\"line\" style=\"{line_style_str}\">"
            + "".join(span_chunks)
            + "</div>"
        )

        left_px = float(left.replace("px", ""))
        top_px = float(top.replace("px", ""))
        right_px = left_px + float(width.replace("px", ""))
        bottom_px = top_px + float(height.replace("px", ""))
        min_left = left_px if min_left is None else min(min_left, left_px)
        min_top = top_px if min_top is None else min(min_top, top_px)
        max_right = right_px if max_right is None else max(max_right, right_px)
        max_bottom = bottom_px if max_bottom is None else max(max_bottom, bottom_px)

    box_style = _extract_box_style(para_style)
    if box_style and min_left is not None:
        box_style.update(
            {
                "left": f"{min_left:.2f}px",
                "top": f"{min_top:.2f}px",
                "width": f"{(max_right - min_left):.2f}px",
                "height": f"{(max_bottom - min_top):.2f}px",
            }
        )
        blocks.insert(0, f"<div class=\"para-box\" style=\"{_build_style(box_style)}\"></div>")

    return blocks


def _render_runs(
    paragraph,
    char_styles,
    para_styles,
    border_fills,
    styles,
    use_lineseg,
    spacing_px=None,
    in_table=False,
):
    blocks = []
    current_segments = []
    has_table = False
    style_id = paragraph.attrib.get("styleIDRef")
    base_para_id = None
    base_char_id = None
    if style_id and style_id in styles and styles[style_id]["type"] == "PARA":
        base_para_id = styles[style_id].get("paraPrIDRef")
        base_char_id = styles[style_id].get("charPrIDRef")

    para_pr = paragraph.attrib.get("paraPrIDRef") or base_para_id
    para_style = _merge_styles(para_styles.get(base_para_id), para_styles.get(para_pr))
    if in_table:
        line_seg_array = paragraph.find("hp:linesegarray", NS)
        if line_seg_array is not None:
            line_segs = line_seg_array.findall("hp:lineseg", NS)
            text = "".join(t.text or "" for t in paragraph.findall(".//hp:t", NS))
            if len(line_segs) == 1 and "\n" not in text:
                para_style["white-space"] = "nowrap"
    if spacing_px:
        para_style["margin-bottom"] = spacing_px
    para_style_str = _build_style(para_style)

    if use_lineseg and not paragraph.findall(".//hp:tbl", NS):
        line_blocks = _render_lineseg_paragraph(
            paragraph, char_styles, para_styles, border_fills, styles
        )
        if line_blocks:
            return line_blocks

    def flush_paragraph():
        nonlocal current_segments
        if current_segments:
            span_html = "".join(current_segments)
            blocks.append(f"<p style=\"{para_style_str}\">{span_html}</p>")
            current_segments = []

    for run in paragraph.findall("hp:run", NS):
        tbl = run.find("hp:tbl", NS)
        if tbl is not None:
            has_table = True
            flush_paragraph()
            blocks.append(
                _render_table(tbl, char_styles, para_styles, border_fills, styles)
            )
            continue

        run_char_id = run.attrib.get("charPrIDRef") or base_char_id
        run_style_map = _merge_styles(
            char_styles.get(base_char_id), char_styles.get(run_char_id)
        )
        run_style = _build_style(run_style_map)
        for t in run.findall("hp:t", NS):
            raw = t.text or ""
            if raw:
                lead = len(raw) - len(raw.lstrip(" "))
                trail = len(raw) - len(raw.rstrip(" "))
                if lead or trail:
                    mid = raw[lead: len(raw) - trail if trail else len(raw)]
                    text = ("&nbsp;" * lead) + html.escape(mid) + ("&nbsp;" * trail)
                else:
                    text = html.escape(raw)
            else:
                text = ""
            if not text:
                continue
            if run_style:
                current_segments.append(f"<span style=\"{run_style}\">{text}</span>")
            else:
                current_segments.append(text)

    if not current_segments and not has_table:
        # Use lineseg metrics for empty paragraph height.
        line_seg_array = paragraph.find("hp:linesegarray", NS)
        if line_seg_array is not None:
            seg = line_seg_array.find("hp:lineseg", NS)
            if seg is not None:
                textheight = _hwpunit_to_px(seg.attrib.get("textheight"))
                vertsize = _hwpunit_to_px(seg.attrib.get("vertsize"))
                if textheight:
                    para_style["line-height"] = textheight
                if vertsize:
                    para_style["min-height"] = vertsize
                para_style_str = _build_style(para_style)

        # Apply explicit run style from XML when paragraph is empty.
        run_style = ""
        first_run = paragraph.find("hp:run", NS)
        if first_run is not None and first_run.attrib.get("charPrIDRef"):
            run_char_id = first_run.attrib.get("charPrIDRef")
            run_style_map = char_styles.get(run_char_id, {})
            run_style = _build_style(run_style_map)
        if run_style:
            blocks.append(
                f"<p style=\"{para_style_str}\"><span style=\"{run_style}\">&nbsp;</span></p>"
            )
        else:
            blocks.append(f"<p style=\"{para_style_str}\">&nbsp;</p>")
    elif current_segments:
        flush_paragraph()

    return blocks


def _border_css(border_fill):
    edges = border_fill.get("edges") if border_fill else None

    def edge_to_css(edge_name):
        edge = edges.get(edge_name) if edges else None
        if not edge or edge.get("type") in (None, "NONE"):
            return "none"
        width = edge.get("width") or "1px"
        style = "solid"
        if edge.get("type") == "DASHED":
            style = "dashed"
        elif edge.get("type") == "DOTTED":
            style = "dotted"
        color = edge.get("color") or "#000"
        return f"{width} {style} {color}"

    return {
        "border-left": edge_to_css("leftBorder"),
        "border-right": edge_to_css("rightBorder"),
        "border-top": edge_to_css("topBorder"),
        "border-bottom": edge_to_css("bottomBorder"),
    }


def _row_first_pos(tr):
    first_p = tr.find(".//hp:p", NS)
    if first_p is None:
        return None
    metrics = _lineseg_metrics(first_p)
    if not metrics:
        return None
    return metrics["first_pos"]


def _row_bottom_pos(tr):
    max_bottom = None
    for p in tr.findall(".//hp:p", NS):
        metrics = _lineseg_metrics(p)
        if not metrics:
            continue
        bottom = metrics["last_pos"] + metrics["last_size"]
        if max_bottom is None or bottom > max_bottom:
            max_bottom = bottom
    return max_bottom


def _table_row_chunks(tbl):
    rows = tbl.findall("hp:tr", NS)
    chunks = []
    current = []
    prev_pos = None
    for tr in rows:
        has_break = any(
            p.attrib.get("pageBreak") == "1" for p in tr.findall(".//hp:p", NS)
        )
        row_pos = _row_first_pos(tr)
        pos_drop = (
            prev_pos is not None and row_pos is not None and row_pos < prev_pos - 1000
        )
        if (has_break or pos_drop) and current:
            chunks.append(current)
            current = []
            prev_pos = None
        current.append(tr)
        if row_pos is not None:
            prev_pos = row_pos
    if current:
        chunks.append(current)
    return chunks


def _table_max_bottom(tbl):
    max_bottom = None
    for tr in tbl.findall("hp:tr", NS):
        row_bottom = _row_bottom_pos(tr)
        if row_bottom is None:
            continue
        if max_bottom is None or row_bottom > max_bottom:
            max_bottom = row_bottom
    return max_bottom


def _render_table(tbl, char_styles, para_styles, border_fills, styles, rows=None):
    tbl_style = {}
    sz = tbl.find("hp:sz", NS)
    if sz is not None:
        tbl_style["width"] = _hwpunit_to_px(sz.attrib.get("width"))
    # If all cells have no borders, suppress table outer border to match HWPX rendering.
    cell_borders = []
    for tc in tbl.findall(".//hp:tc", NS):
        cell_border = border_fills.get(tc.attrib.get("borderFillIDRef"))
        cell_borders.append(cell_border)
    all_cells_borderless = False
    if cell_borders:
        all_cells_borderless = True
        for b in cell_borders:
            edges = b.get("edges") if b else None
            if not edges:
                continue
            if any(edge.get("type") not in (None, "NONE") for edge in edges.values()):
                all_cells_borderless = False
                break

    tbl_border = border_fills.get(tbl.attrib.get("borderFillIDRef"))
    if not all_cells_borderless:
        tbl_style.update(_border_css(tbl_border))
    pos = tbl.find("hp:pos", NS)
    if pos is not None and pos.attrib.get("horzAlign") == "CENTER":
        tbl_style["margin-left"] = "auto"
        tbl_style["margin-right"] = "auto"
    in_margin = tbl.find("hp:inMargin", NS)
    table_padding = None
    if in_margin is not None:
        pad = _hwpunit_to_px(in_margin.attrib.get("left"))
        table_padding = pad
    spacing = _hwpunit_to_px(tbl.attrib.get("cellSpacing"))
    if spacing and spacing != "0.00px":
        tbl_style["border-collapse"] = "separate"
        tbl_style["border-spacing"] = spacing
    else:
        # If borders have mixed widths/types, avoid collapse to reduce artifacts.
        border_widths = set()
        border_types = set()
        for b in cell_borders + [tbl_border]:
            if not b:
                continue
            edges = b.get("edges") or {}
            for edge in edges.values():
                if edge.get("type") in (None, "NONE"):
                    continue
                if edge.get("width"):
                    border_widths.add(edge.get("width"))
                border_types.add(edge.get("type"))
        if len(border_widths) > 1 or len(border_types) > 1:
            tbl_style["border-collapse"] = "separate"
            tbl_style["border-spacing"] = "0px"
    scale_y = tbl.attrib.get("_scale_y")
    scaled_height = tbl.attrib.get("_scaled_height_px")
    if scale_y:
        tbl_style["transform"] = f"scaleY({scale_y})"
        tbl_style["transform-origin"] = "top left"
        tbl_style["display"] = "inline-block"

    html_rows = []
    table_rows = rows or tbl.findall("hp:tr", NS)
    for tr in table_rows:
        cells_html = []
        for tc in tr.findall("hp:tc", NS):
            cell_style = {}
            cell_border = border_fills.get(tc.attrib.get("borderFillIDRef"))
            cell_style.update(_border_css(cell_border))
            if cell_border and cell_border.get("fill"):
                cell_style["background-color"] = cell_border.get("fill")
            cell_sz = tc.find("hp:cellSz", NS)
            if cell_sz is not None:
                cell_style["width"] = _hwpunit_to_px(cell_sz.attrib.get("width"))
            cell_margin = tc.find("hp:cellMargin", NS)
            if cell_margin is not None:
                padding = _hwpunit_to_px(cell_margin.attrib.get("left"))
                if padding:
                    cell_style["padding"] = padding
            elif table_padding:
                cell_style["padding"] = table_padding

            sub_list = tc.find("hp:subList", NS)
            if sub_list is None:
                cell_content = "&nbsp;"
            else:
                blocks = _render_block_list(
                    sub_list,
                    char_styles,
                    para_styles,
                    border_fills,
                    styles,
                    use_lineseg=False,
                    in_table=True,
                )
                cell_content = "".join(blocks)
                vert = sub_list.attrib.get("vertAlign")
                if vert == "CENTER":
                    cell_style["vertical-align"] = "middle"
                elif vert == "BOTTOM":
                    cell_style["vertical-align"] = "bottom"
            span = tc.find("hp:cellSpan", NS)
            span_attrs = ""
            if span is not None:
                col_span = span.attrib.get("colSpan")
                row_span = span.attrib.get("rowSpan")
                if col_span and col_span != "1":
                    span_attrs += f" colspan=\"{col_span}\""
                if row_span and row_span != "1":
                    span_attrs += f" rowspan=\"{row_span}\""
            cells_html.append(
                f"<td style=\"{_build_style(cell_style)}\"{span_attrs}>{cell_content}</td>"
            )
        html_rows.append("<tr>" + "".join(cells_html) + "</tr>")
    table_html = (
        f"<table style=\"{_build_style(tbl_style)}\">" + "".join(html_rows) + "</table>"
    )
    if scale_y and scaled_height:
        return (
            f"<div class=\"table-scale\" style=\"height:{scaled_height}; overflow:hidden\">"
            + table_html
            + "</div>"
        )
    return table_html


def _lineseg_metrics(paragraph):
    line_seg_array = paragraph.find("hp:linesegarray", NS)
    if line_seg_array is None:
        return None
    line_segs = line_seg_array.findall("hp:lineseg", NS)
    if not line_segs:
        return None
    first = line_segs[0]
    last = line_segs[-1]
    try:
        first_pos = int(first.attrib.get("vertpos", "0"))
        first_size = int(first.attrib.get("vertsize", "0"))
        last_pos = int(last.attrib.get("vertpos", "0"))
        last_size = int(last.attrib.get("vertsize", "0"))
    except ValueError:
        return None
    return {
        "first_pos": first_pos,
        "first_size": first_size,
        "last_pos": last_pos,
        "last_size": last_size,
    }


def _child_first_pos(child):
    tag = child.tag.split("}")[-1]
    if tag == "p":
        metrics = _lineseg_metrics(child)
        return metrics["first_pos"] if metrics else None
    if tag == "tbl":
        first_p = child.find(".//hp:p", NS)
        if first_p is not None:
            metrics = _lineseg_metrics(first_p)
            return metrics["first_pos"] if metrics else None
    return None


def _child_last_bottom(child):
    tag = child.tag.split("}")[-1]
    if tag == "p":
        metrics = _lineseg_metrics(child)
        if not metrics:
            return None
        return metrics["last_pos"] + metrics["last_size"]
    if tag == "tbl":
        return _table_max_bottom(child)
    return None


def _resolve_para_style(paragraph, para_styles, styles):
    style_id = paragraph.attrib.get("styleIDRef")
    base_para_id = None
    if style_id and style_id in styles and styles[style_id]["type"] == "PARA":
        base_para_id = styles[style_id].get("paraPrIDRef")
    para_pr = paragraph.attrib.get("paraPrIDRef") or base_para_id
    return _merge_styles(para_styles.get(base_para_id), para_styles.get(para_pr))


def _parse_padding_px(style_map: dict) -> tuple[float, float]:
    padding = style_map.get("padding")
    if not padding:
        return 0.0, 0.0
    parts = padding.split()
    if len(parts) == 1:
        top = bottom = parts[0]
    elif len(parts) == 2:
        top = bottom = parts[0]
    elif len(parts) == 3:
        top = parts[0]
        bottom = parts[2]
    else:
        top = parts[0]
        bottom = parts[2]
    try:
        return float(top.replace("px", "")), float(bottom.replace("px", ""))
    except ValueError:
        return 0.0, 0.0


def _render_children(
    children,
    char_styles,
    para_styles,
    border_fills,
    styles,
    use_lineseg=True,
    in_table=False,
):
    blocks = []
    for idx, child in enumerate(children):
        tag = child.tag.split("}")[-1]
        if tag == "p":
            spacing_px = None
            if not use_lineseg:
                current = _lineseg_metrics(child)
                next_child = None
                next_kind = None
                for j in range(idx + 1, len(children)):
                    tag_name = children[j].tag.split("}")[-1]
                    if tag_name in ("p", "tbl"):
                        next_child = children[j]
                        next_kind = tag_name
                        break
                if current and next_child is not None:
                    nxt = None
                    table_top_adjust = 0.0
                    next_has_table = next_child.find(".//hp:tbl", NS) is not None
                    if next_kind == "p":
                        nxt = _lineseg_metrics(next_child)
                    elif next_kind == "tbl":
                        first_p = next_child.find(".//hp:p", NS)
                        if first_p is not None:
                            nxt = _lineseg_metrics(first_p)
                        out_margin = next_child.find("hp:outMargin", NS)
                        if out_margin is not None:
                            top = _hwpunit_to_px(out_margin.attrib.get("top"))
                            if top:
                                try:
                                    table_top_adjust += float(top.replace("px", ""))
                                except ValueError:
                                    pass
                    if nxt:
                        gap = (
                            nxt["first_pos"]
                            - current["last_pos"]
                            - current["last_size"]
                        )
                        if gap > 0:
                            spacing_px = _hwpunit_to_px(gap)
                            current_style = _resolve_para_style(
                                child, para_styles, styles
                            )
                            _, current_bottom = _parse_padding_px(current_style)
                            next_top = 0.0
                            if next_kind == "p":
                                next_style = _resolve_para_style(
                                    next_child, para_styles, styles
                                )
                                next_top, _ = _parse_padding_px(next_style)
                            if spacing_px:
                                try:
                                    value = float(spacing_px.replace("px", ""))
                                except ValueError:
                                    value = None
                                if value is not None:
                                    adjusted = max(
                                        0.0,
                                        value - current_bottom - next_top - table_top_adjust,
                                    )
                                    spacing_px = f"{adjusted:.2f}px"
                    if next_kind == "tbl" or next_has_table:
                        spacing_px = "0px"
            blocks.extend(
                _render_runs(
                    child,
                    char_styles,
                    para_styles,
                    border_fills,
                    styles,
                    use_lineseg,
                    spacing_px=spacing_px,
                    in_table=in_table,
                )
            )
        elif tag == "tbl":
            blocks.append(
                _render_table(child, char_styles, para_styles, border_fills, styles)
            )
    return blocks


def _render_block_list(
    parent,
    char_styles,
    para_styles,
    border_fills,
    styles,
    use_lineseg=True,
    in_table=False,
):
    return _render_children(
        list(parent),
        char_styles,
        para_styles,
        border_fills,
        styles,
        use_lineseg=use_lineseg,
        in_table=in_table,
    )


def _sorted_section_names(zipf: zipfile.ZipFile):
    section_re = re.compile(r"Contents/section(\d+)\.xml$")
    sections = []
    for name in zipf.namelist():
        m = section_re.search(name)
        if m:
            sections.append((int(m.group(1)), name))
    return [name for _, name in sorted(sections)]


def _section_page_style(section_root):
    sec_pr = section_root.find(".//hp:secPr", NS)
    if sec_pr is None:
        return {}
    page_pr = sec_pr.find("hp:pagePr", NS)
    if page_pr is None:
        return {}
    style = {}
    page_width = _hwpunit_to_px(page_pr.attrib.get("width"))
    page_height = _hwpunit_to_px(page_pr.attrib.get("height"))
    style["width"] = page_width
    style["min-height"] = page_height
    style["height"] = page_height
    margin = page_pr.find("hp:margin", NS)
    if margin is not None:
        top = _hwpunit_to_px(margin.attrib.get("top")) or "0"
        right = _hwpunit_to_px(margin.attrib.get("right")) or "0"
        bottom = _hwpunit_to_px(margin.attrib.get("bottom")) or "0"
        left = _hwpunit_to_px(margin.attrib.get("left")) or "0"
        style["padding"] = f"{top} {right} {bottom} {left}"
    return style


def hwpx_to_html(hwpx_path: Path, output_path: Path, use_lineseg: bool):
    with zipfile.ZipFile(hwpx_path) as zipf:
        char_styles, para_styles, border_fills, styles = _parse_header(zipf)
        section_names = _sorted_section_names(zipf)
        body_blocks = []
        for section in section_names:
            sec_xml = zipf.read(section).decode("utf-8")
            root = ET.fromstring(sec_xml)
            page_style = _section_page_style(root)
            page_height_hwp = None
            margin_top_hwp = 0
            margin_bottom_hwp = 0
            page_pr = root.find(".//hp:secPr/hp:pagePr", NS)
            if page_pr is not None:
                try:
                    page_height_hwp = int(page_pr.attrib.get("height", "0"))
                except ValueError:
                    page_height_hwp = None
                margin = page_pr.find("hp:margin", NS)
                if margin is not None:
                    try:
                        margin_top_hwp = int(margin.attrib.get("top", "0"))
                        margin_bottom_hwp = int(margin.attrib.get("bottom", "0"))
                    except ValueError:
                        margin_top_hwp = 0
                        margin_bottom_hwp = 0
            children = list(root)
            page_children = []
            prev_first_pos = None
            prev_last_bottom = None
            for child in children:
                tag = child.tag.split("}")[-1]
                if tag == "tbl":
                    table_first_pos = _child_first_pos(child)
                    if (
                        table_first_pos is not None
                        and prev_first_pos is not None
                        and table_first_pos < prev_first_pos - 1000
                    ):
                        if page_children:
                            page_block = _render_children(
                                page_children,
                                char_styles,
                                para_styles,
                                border_fills,
                                styles,
                                use_lineseg=use_lineseg,
                            )
                            body_blocks.append(
                                f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                                + "".join(page_block)
                                + "</div>"
                            )
                            page_children = []
                        prev_first_pos = None
                    max_bottom = _table_max_bottom(child)
                    if page_height_hwp and max_bottom:
                        available = page_height_hwp - margin_top_hwp - margin_bottom_hwp
                        if available > 0 and max_bottom > available:
                            scale = max(0.5, available / max_bottom)
                            child.attrib["_scale_y"] = f"{scale:.4f}"
                            available_px = _hwpunit_to_px(available)
                            if available_px:
                                child.attrib["_scaled_height_px"] = available_px
                    chunks = _table_row_chunks(child)
                    if len(chunks) > 1:
                        if page_children:
                            page_block = _render_children(
                                page_children,
                                char_styles,
                                para_styles,
                                border_fills,
                                styles,
                                use_lineseg=use_lineseg,
                            )
                            body_blocks.append(
                                f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                                + "".join(page_block)
                                + "</div>"
                            )
                            page_children = []
                        for rows in chunks:
                            table_html = _render_table(
                                child,
                                char_styles,
                                para_styles,
                                border_fills,
                                styles,
                                rows=rows,
                            )
                            body_blocks.append(
                                f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                                + table_html
                                + "</div>"
                            )
                        prev_first_pos = None
                        continue
                child_first_pos = _child_first_pos(child)
                child_last_bottom = _child_last_bottom(child)
                if (
                    child_first_pos is not None
                    and prev_first_pos is not None
                    and child_first_pos < prev_first_pos - 1000
                ):
                    if page_children:
                        page_block = _render_children(
                            page_children,
                            char_styles,
                            para_styles,
                            border_fills,
                            styles,
                            use_lineseg=use_lineseg,
                        )
                        body_blocks.append(
                            f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                            + "".join(page_block)
                            + "</div>"
                        )
                        page_children = []
                if (
                    child_first_pos is not None
                    and child_first_pos <= 1000
                    and prev_last_bottom is not None
                    and prev_last_bottom > 2000
                ):
                    if page_children:
                        page_block = _render_children(
                            page_children,
                            char_styles,
                            para_styles,
                            border_fills,
                            styles,
                            use_lineseg=use_lineseg,
                        )
                        body_blocks.append(
                            f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                            + "".join(page_block)
                            + "</div>"
                        )
                        page_children = []
                if tag == "p" and child.attrib.get("pageBreak") == "1":
                    if page_children:
                        page_block = _render_children(
                            page_children,
                            char_styles,
                            para_styles,
                            border_fills,
                            styles,
                            use_lineseg=use_lineseg,
                        )
                        body_blocks.append(
                            f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                            + "".join(page_block)
                            + "</div>"
                        )
                        page_children = []
                page_children.append(child)
                if child_first_pos is not None:
                    prev_first_pos = child_first_pos
                if child_last_bottom is not None:
                    prev_last_bottom = child_last_bottom

            if page_children:
                page_block = _render_children(
                    page_children,
                    char_styles,
                    para_styles,
                    border_fills,
                    styles,
                    use_lineseg=use_lineseg,
                )
                body_blocks.append(
                    f"<div class=\"page\" style=\"{_build_style(page_style)}\">"
                    + "".join(page_block)
                    + "</div>"
                )

    html_doc = """<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>HWPX Export</title>
  <style>
    body { font-family: sans-serif; line-height: 1.3; background: #eee; }
    .page { position: relative; box-sizing: border-box; background: #fff; margin: 18px auto; border: 1px solid #d0d0d0; box-shadow: 0 2px 10px rgba(0,0,0,.15); page-break-after: always; overflow: auto; }
    .page:last-child { page-break-after: auto; }
    .line { position: absolute; white-space: pre; }
    .para-box { position: absolute; box-sizing: border-box; }
    .table-scale { display: inline-block; }
    table { border-collapse: collapse; margin: 6px 0; }
    td { vertical-align: top; }
    p { margin: 0; }
  </style>
</head>
<body>
"""
    html_doc += "".join(body_blocks)
    html_doc += "\n</body>\n</html>\n"
    output_path.write_text(html_doc, encoding="utf-8")


def _convert_all_in_folder(folder: Path, use_lineseg: bool):
    hwpx_files = sorted(folder.glob("*.hwpx"))
    if not hwpx_files:
        print(f"No .hwpx files found in: {folder}")
        return
    for path in hwpx_files:
        output = path.with_suffix(".html")
        hwpx_to_html(path, output, use_lineseg=use_lineseg)
        print(f"Saved: {output}")


def main():
    parser = argparse.ArgumentParser(description="Convert HWPX to HTML (basic).")
    parser.add_argument("hwpx", type=Path, nargs="?", help="Path to .hwpx file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output HTML path (default: same name with .html)",
    )
    parser.add_argument(
        "--folder",
        type=Path,
        default=Path("file"),
        help="Folder to scan when no hwpx is provided (default: file)",
    )
    parser.add_argument(
        "--layout",
        choices=["flow", "lineseg"],
        default="flow",
        help="Layout mode: flow (stable) or lineseg (absolute)",
    )
    args = parser.parse_args()
    use_lineseg = args.layout == "lineseg"
    if args.hwpx:
        output = args.output or args.hwpx.with_suffix(".html")
        hwpx_to_html(args.hwpx, output, use_lineseg=use_lineseg)
        print(f"Saved: {output}")
    else:
        _convert_all_in_folder(args.folder, use_lineseg=use_lineseg)


if __name__ == "__main__":
    main()
