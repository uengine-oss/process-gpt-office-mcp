"""DOCX → HTML conversion using pure stdlib (zipfile + xml.etree.ElementTree)."""
import base64
import html
import mimetypes
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

R_EMBED = f"{{{NS['r']}}}embed"
EMU_PER_PX = 914400 / 96  # 1 inch = 914400 EMU, 96 px/inch


def _css_escape(value: str) -> str:
    return value.replace("'", "\\'")


def _build_style(style_map: dict) -> str:
    items = [f"{k}:{v}" for k, v in style_map.items() if v]
    return "; ".join(items)


def _twips_to_px(value) -> str | None:
    if value is None:
        return None
    try:
        px = float(value) / 1440 * 96
        return f"{px:.2f}px"
    except (TypeError, ValueError):
        return None


def _half_pt_to_pt(value: str | None) -> str | None:
    if not value:
        return None
    try:
        return f"{int(value) / 2:g}pt"
    except (TypeError, ValueError):
        return None


def _eighth_pt_to_px(value: str | None) -> str | None:
    if not value:
        return None
    try:
        pt = int(value) / 8
        return f"{max(pt, 0.5):.1f}px"
    except (TypeError, ValueError):
        return None


def _normalize_color(value: str | None) -> str | None:
    if not value or value.lower() == "auto":
        return None
    return value if value.startswith("#") else f"#{value}"


# ---------------------------------------------------------------------------
# Style parsing
# ---------------------------------------------------------------------------

def _parse_styles(zipf: zipfile.ZipFile):
    try:
        data = zipf.read("word/styles.xml").decode("utf-8")
    except KeyError:
        return {}, {}, {}
    root = ET.fromstring(data)

    doc_defaults_rpr = {}
    doc_defaults_ppr = {}
    rpr_default = root.find(".//w:docDefaults/w:rPrDefault/w:rPr", NS)
    if rpr_default is not None:
        doc_defaults_rpr = _parse_rpr(rpr_default)
    ppr_default = root.find(".//w:docDefaults/w:pPrDefault/w:pPr", NS)
    if ppr_default is not None:
        doc_defaults_ppr = _parse_ppr(ppr_default)

    styles_map = {}
    for style_el in root.findall("w:style", NS):
        style_id = style_el.attrib.get(f"{{{NS['w']}}}styleId")
        if not style_id:
            continue
        based_on_el = style_el.find("w:basedOn", NS)
        based_on = based_on_el.attrib.get(f"{{{NS['w']}}}val") if based_on_el is not None else None

        rpr = {}
        rpr_el = style_el.find("w:rPr", NS)
        if rpr_el is not None:
            rpr = _parse_rpr(rpr_el)

        ppr = {}
        ppr_el = style_el.find("w:pPr", NS)
        if ppr_el is not None:
            ppr = _parse_ppr(ppr_el)

        tbl_pr = {}
        tbl_pr_el = style_el.find("w:tblPr", NS)
        if tbl_pr_el is not None:
            tbl_pr = _parse_tbl_pr(tbl_pr_el)

        styles_map[style_id] = {
            "type": style_el.attrib.get(f"{{{NS['w']}}}type"),
            "basedOn": based_on,
            "rPr": rpr,
            "pPr": ppr,
            "tblPr": tbl_pr,
        }

    return doc_defaults_rpr, doc_defaults_ppr, styles_map


def _resolve_style_rpr(style_id: str | None, styles_map: dict, visited=None) -> dict:
    if not style_id or style_id not in styles_map:
        return {}
    if visited is None:
        visited = set()
    if style_id in visited:
        return {}
    visited.add(style_id)
    style = styles_map[style_id]
    base = _resolve_style_rpr(style.get("basedOn"), styles_map, visited)
    base.update(style.get("rPr", {}))
    return base


def _resolve_style_ppr(style_id: str | None, styles_map: dict, visited=None) -> dict:
    if not style_id or style_id not in styles_map:
        return {}
    if visited is None:
        visited = set()
    if style_id in visited:
        return {}
    visited.add(style_id)
    style = styles_map[style_id]
    base = _resolve_style_ppr(style.get("basedOn"), styles_map, visited)
    base.update(style.get("pPr", {}))
    return base


# ---------------------------------------------------------------------------
# Run properties (character-level)
# ---------------------------------------------------------------------------

def _parse_rpr(rpr) -> dict:
    style = {}
    if rpr is None:
        return style

    fonts = rpr.find("w:rFonts", NS)
    if fonts is not None:
        face = (
            fonts.attrib.get(f"{{{NS['w']}}}eastAsia")
            or fonts.attrib.get(f"{{{NS['w']}}}ascii")
            or fonts.attrib.get(f"{{{NS['w']}}}hAnsi")
        )
        if face:
            style["font-family"] = f"'{_css_escape(face)}'"

    sz = rpr.find("w:sz", NS)
    if sz is not None:
        pt = _half_pt_to_pt(sz.attrib.get(f"{{{NS['w']}}}val"))
        if pt:
            style["font-size"] = pt

    b = rpr.find("w:b", NS)
    if b is not None:
        if b.attrib.get(f"{{{NS['w']}}}val", "true") not in ("false", "0"):
            style["font-weight"] = "700"

    bcs = rpr.find("w:bCs", NS)
    if bcs is not None and "font-weight" not in style:
        if bcs.attrib.get(f"{{{NS['w']}}}val", "true") not in ("false", "0"):
            style["font-weight"] = "700"

    i = rpr.find("w:i", NS)
    if i is not None:
        if i.attrib.get(f"{{{NS['w']}}}val", "true") not in ("false", "0"):
            style["font-style"] = "italic"

    color = rpr.find("w:color", NS)
    if color is not None:
        c = _normalize_color(color.attrib.get(f"{{{NS['w']}}}val"))
        if c:
            style["color"] = c

    u = rpr.find("w:u", NS)
    if u is not None:
        u_val = u.attrib.get(f"{{{NS['w']}}}val", "single")
        if u_val != "none":
            style["text-decoration-line"] = "underline"
            u_color = _normalize_color(u.attrib.get(f"{{{NS['w']}}}color"))
            if u_color:
                style["text-decoration-color"] = u_color
            if u_val == "dash":
                style["text-decoration-style"] = "dashed"
            elif u_val == "dotted":
                style["text-decoration-style"] = "dotted"
            elif u_val == "wave":
                style["text-decoration-style"] = "wavy"
            elif u_val == "double":
                style["text-decoration-style"] = "double"

    strike = rpr.find("w:strike", NS)
    if strike is not None:
        if strike.attrib.get(f"{{{NS['w']}}}val", "true") not in ("false", "0"):
            style["text-decoration-line"] = "line-through"

    highlight = rpr.find("w:highlight", NS)
    if highlight is not None:
        style["background-color"] = highlight.attrib.get(f"{{{NS['w']}}}val", "yellow")

    shd = rpr.find("w:shd", NS)
    if shd is not None and "background-color" not in style:
        fill = _normalize_color(shd.attrib.get(f"{{{NS['w']}}}fill"))
        if fill:
            style["background-color"] = fill

    vert = rpr.find("w:vertAlign", NS)
    if vert is not None:
        v = vert.attrib.get(f"{{{NS['w']}}}val")
        if v == "superscript":
            style["vertical-align"] = "super"
            style["font-size"] = "smaller"
        elif v == "subscript":
            style["vertical-align"] = "sub"
            style["font-size"] = "smaller"

    return style


# ---------------------------------------------------------------------------
# Paragraph properties
# ---------------------------------------------------------------------------

def _parse_ppr(ppr) -> dict:
    style = {}
    if ppr is None:
        return style

    jc = ppr.find("w:jc", NS)
    if jc is not None:
        val = jc.attrib.get(f"{{{NS['w']}}}val", "left")
        if val == "center":
            style["text-align"] = "center"
        elif val == "right":
            style["text-align"] = "right"
        elif val in ("both", "distribute"):
            style["text-align"] = "justify"

    spacing = ppr.find("w:spacing", NS)
    if spacing is not None:
        before = spacing.attrib.get(f"{{{NS['w']}}}before")
        after = spacing.attrib.get(f"{{{NS['w']}}}after")
        line = spacing.attrib.get(f"{{{NS['w']}}}line")
        line_rule = spacing.attrib.get(f"{{{NS['w']}}}lineRule", "auto")
        if before:
            px = _twips_to_px(before)
            if px:
                style["margin-top"] = px
        if after:
            px = _twips_to_px(after)
            if px:
                style["margin-bottom"] = px
        if line:
            if line_rule == "auto":
                try:
                    style["line-height"] = f"{int(line) / 240:.2f}"
                except ValueError:
                    pass
            elif line_rule in ("exact", "atLeast"):
                px = _twips_to_px(line)
                if px:
                    style["line-height"] = px

    ind = ppr.find("w:ind", NS)
    if ind is not None:
        left = ind.attrib.get(f"{{{NS['w']}}}left")
        right = ind.attrib.get(f"{{{NS['w']}}}right")
        first_line = ind.attrib.get(f"{{{NS['w']}}}firstLine")
        hanging = ind.attrib.get(f"{{{NS['w']}}}hanging")
        if left:
            px = _twips_to_px(left)
            if px:
                style["margin-left"] = px
        if right:
            px = _twips_to_px(right)
            if px:
                style["margin-right"] = px
        if first_line:
            px = _twips_to_px(first_line)
            if px:
                style["text-indent"] = px
        elif hanging:
            px = _twips_to_px(hanging)
            if px:
                style["text-indent"] = f"-{px}"

    pbdr = ppr.find("w:pBdr", NS)
    if pbdr is not None:
        for side in ("top", "left", "bottom", "right"):
            bdr = pbdr.find(f"w:{side}", NS)
            if bdr is not None:
                border_css = _parse_border_element(bdr)
                if border_css:
                    style[f"border-{side}"] = border_css
                space = bdr.attrib.get(f"{{{NS['w']}}}space")
                if space:
                    px = _twips_to_px(int(space) * 20)
                    if px:
                        style[f"padding-{side}"] = px

    shd = ppr.find("w:shd", NS)
    if shd is not None:
        fill = _normalize_color(shd.attrib.get(f"{{{NS['w']}}}fill"))
        if fill:
            style["background-color"] = fill

    return style


def _parse_border_element(bdr) -> str | None:
    val = bdr.attrib.get(f"{{{NS['w']}}}val", "none")
    if val in ("none", "nil"):
        return None
    sz = _eighth_pt_to_px(bdr.attrib.get(f"{{{NS['w']}}}sz"))
    color = _normalize_color(bdr.attrib.get(f"{{{NS['w']}}}color"))
    width = sz or "1px"
    css_color = color or "#000000"
    if val == "dashed":
        css_style = "dashed"
    elif val == "dotted":
        css_style = "dotted"
    elif val == "double":
        css_style = "double"
    else:
        css_style = "solid"
    return f"{width} {css_style} {css_color}"


# ---------------------------------------------------------------------------
# Table properties
# ---------------------------------------------------------------------------

def _parse_tbl_pr(tbl_pr) -> dict:
    style = {}
    if tbl_pr is None:
        return style
    tbl_w = tbl_pr.find("w:tblW", NS)
    if tbl_w is not None:
        w_type = tbl_w.attrib.get(f"{{{NS['w']}}}type", "dxa")
        w_val = tbl_w.attrib.get(f"{{{NS['w']}}}w")
        if w_type == "dxa" and w_val:
            style["width"] = _twips_to_px(w_val)
        elif w_type == "pct" and w_val:
            try:
                style["width"] = f"{int(w_val) / 50:.1f}%"
            except ValueError:
                pass
        elif w_type == "auto":
            style["width"] = "auto"
    return style


def _parse_tbl_borders(tbl_pr) -> dict:
    borders = {}
    if tbl_pr is None:
        return borders
    tbl_borders = tbl_pr.find("w:tblBorders", NS)
    if tbl_borders is None:
        return borders
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        bdr = tbl_borders.find(f"w:{side}", NS)
        if bdr is not None:
            css = _parse_border_element(bdr)
            if css:
                borders[side] = css
    return borders


def _parse_tc_borders(tc_pr) -> dict:
    borders = {}
    if tc_pr is None:
        return borders
    tc_borders = tc_pr.find("w:tcBorders", NS)
    if tc_borders is None:
        return borders
    for side in ("top", "left", "bottom", "right"):
        bdr = tc_borders.find(f"w:{side}", NS)
        if bdr is not None:
            val = bdr.attrib.get(f"{{{NS['w']}}}val", "none")
            if val in ("none", "nil"):
                borders[f"border-{side}"] = "none"
            else:
                css = _parse_border_element(bdr)
                if css:
                    borders[f"border-{side}"] = css
    return borders


# ---------------------------------------------------------------------------
# Section properties
# ---------------------------------------------------------------------------

def _parse_sect_pr(sect_pr) -> dict:
    style = {}
    if sect_pr is None:
        return style
    pg_sz = sect_pr.find("w:pgSz", NS)
    if pg_sz is not None:
        style["width"] = _twips_to_px(pg_sz.attrib.get(f"{{{NS['w']}}}w"))
        style["min-height"] = _twips_to_px(pg_sz.attrib.get(f"{{{NS['w']}}}h"))
    pg_mar = sect_pr.find("w:pgMar", NS)
    if pg_mar is not None:
        top = _twips_to_px(pg_mar.attrib.get(f"{{{NS['w']}}}top")) or "0"
        right = _twips_to_px(pg_mar.attrib.get(f"{{{NS['w']}}}right")) or "0"
        bottom = _twips_to_px(pg_mar.attrib.get(f"{{{NS['w']}}}bottom")) or "0"
        left = _twips_to_px(pg_mar.attrib.get(f"{{{NS['w']}}}left")) or "0"
        style["padding"] = f"{top} {right} {bottom} {left}"
    return style


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def _merge_styles(base: dict | None, override: dict | None) -> dict:
    merged = {}
    if base:
        merged.update(base)
    if override:
        merged.update(override)
    return merged


class _Counter:
    """Sequential ID counter shared across rendering calls."""
    def __init__(self):
        self.value = 0

    def next(self) -> int:
        v = self.value
        self.value += 1
        return v


def _parse_relationships(zipf: zipfile.ZipFile) -> dict:
    """Parse word/_rels/document.xml.rels → {rId: 'word/media/...'} for images."""
    try:
        data = zipf.read("word/_rels/document.xml.rels").decode("utf-8")
    except KeyError:
        return {}
    root = ET.fromstring(data)
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    image_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    result = {}
    for rel in root.findall(f"{{{rel_ns}}}Relationship"):
        if rel.attrib.get("Type") == image_type:
            rid = rel.attrib.get("Id", "")
            target = rel.attrib.get("Target", "")
            if target and not target.startswith("/"):
                target = "word/" + target
            if rid:
                result[rid] = target
    return result


def _render_drawing(drawing, images_map: dict, zipf: zipfile.ZipFile) -> str:
    """Render a <w:drawing> element as an <img> tag."""
    # Find extent for dimensions
    width_px = height_px = None
    extent = drawing.find(
        f".//{{{NS['wp']}}}extent"
    )
    if extent is not None:
        cx = extent.attrib.get("cx")
        cy = extent.attrib.get("cy")
        try:
            if cx:
                width_px = int(cx) / EMU_PER_PX
            if cy:
                height_px = int(cy) / EMU_PER_PX
        except (ValueError, TypeError):
            pass

    # Find blip rEmbed
    blip = drawing.find(f".//{{{NS['a']}}}blip")
    if blip is None:
        return ""
    r_embed = blip.attrib.get(R_EMBED, "")
    if not r_embed:
        return ""

    media_path = images_map.get(r_embed, "")
    if not media_path:
        return ""

    try:
        img_bytes = zipf.read(media_path)
    except KeyError:
        return ""

    mime, _ = mimetypes.guess_type(media_path)
    if not mime:
        mime = "image/png"

    b64 = base64.b64encode(img_bytes).decode("ascii")
    src = f"data:{mime};base64,{b64}"

    style_parts = ["max-width:100%"]
    if width_px:
        style_parts.append(f"width:{width_px:.1f}px")
    if height_px:
        style_parts.append(f"height:{height_px:.1f}px")

    return f'<img src="{src}" style="{"; ".join(style_parts)}" />'


def _render_run(run, doc_defaults_rpr, styles_map, parent_style_id=None, images_map=None, zipf=None):
    drawing = run.find("w:drawing", NS)
    if drawing is not None:
        if images_map and zipf:
            return _render_drawing(drawing, images_map, zipf)
        return ""

    parts = []
    for child in run:
        tag = child.tag.split("}")[-1]
        if tag == "t":
            text = child.text or ""
            if text:
                lead = len(text) - len(text.lstrip(" "))
                trail = len(text) - len(text.rstrip(" "))
                mid = text[lead: len(text) - trail if trail else len(text)]
                parts.append(("&nbsp;" * lead) + html.escape(mid) + ("&nbsp;" * trail))
        elif tag == "br":
            br_type = child.attrib.get(f"{{{NS['w']}}}type", "textWrapping")
            parts.append('<br class="page-break">' if br_type == "page" else "<br>")
        elif tag == "tab":
            parts.append("&emsp;")
        elif tag == "sym":
            char_code = child.attrib.get(f"{{{NS['w']}}}char", "")
            if char_code:
                try:
                    parts.append(chr(int(char_code, 16)))
                except ValueError:
                    parts.append("?")

    if not parts:
        return ""

    content = "".join(parts)

    rpr = run.find("w:rPr", NS)
    base_rpr = dict(doc_defaults_rpr)
    if parent_style_id:
        base_rpr.update(_resolve_style_rpr(parent_style_id, styles_map))
    inline_rpr = _parse_rpr(rpr)
    final_rpr = _merge_styles(base_rpr, inline_rpr) if inline_rpr else base_rpr

    style_str = _build_style(final_rpr)
    return f'<span style="{style_str}">{content}</span>' if style_str else content


def _render_paragraph(paragraph, doc_defaults_rpr, doc_defaults_ppr, styles_map, counter=None, images_map=None, zipf=None):
    ppr = paragraph.find("w:pPr", NS)

    style_id = None
    if ppr is not None:
        style_ref = ppr.find("w:pStyle", NS)
        if style_ref is not None:
            style_id = style_ref.attrib.get(f"{{{NS['w']}}}val")

    base_ppr = dict(doc_defaults_ppr)
    if style_id:
        base_ppr.update(_resolve_style_ppr(style_id, styles_map))
    inline_ppr = _parse_ppr(ppr)
    final_ppr = _merge_styles(base_ppr, inline_ppr)

    run_htmls = []
    for child in paragraph:
        tag = child.tag.split("}")[-1]
        if tag == "r":
            run_html = _render_run(child, doc_defaults_rpr, styles_map, style_id, images_map=images_map, zipf=zipf)
            if run_html:
                run_htmls.append(run_html)
        elif tag == "hyperlink":
            for sub_run in child.findall("w:r", NS):
                run_html = _render_run(sub_run, doc_defaults_rpr, styles_map, style_id, images_map=images_map, zipf=zipf)
                if run_html:
                    run_htmls.append(run_html)

    content = "".join(run_htmls) if run_htmls else "&nbsp;"
    id_attr = f' data-id="{counter.next()}"' if counter is not None else ""
    return f'<p{id_attr} style="{_build_style(final_ppr)}">{content}</p>'


def _count_vmerge_rowspan(tbl, tr_index, tc_index):
    rows = tbl.findall("w:tr", NS)
    count = 1
    for i in range(tr_index + 1, len(rows)):
        cells = rows[i].findall("w:tc", NS)
        if tc_index >= len(cells):
            break
        tc_pr = cells[tc_index].find("w:tcPr", NS)
        if tc_pr is None:
            break
        v_merge = tc_pr.find("w:vMerge", NS)
        if v_merge is None:
            break
        if v_merge.attrib.get(f"{{{NS['w']}}}val", "continue") == "continue":
            count += 1
        else:
            break
    return count


def _render_table(tbl, doc_defaults_rpr, doc_defaults_ppr, styles_map, counter=None, images_map=None, zipf=None):
    tbl_pr = tbl.find("w:tblPr", NS)
    tbl_style = {}
    tbl_style.update(_parse_tbl_pr(tbl_pr))

    if tbl_pr is not None:
        jc = tbl_pr.find("w:jc", NS)
        if jc is not None:
            val = jc.attrib.get(f"{{{NS['w']}}}val")
            if val == "center":
                tbl_style["margin-left"] = "auto"
                tbl_style["margin-right"] = "auto"
            elif val == "right":
                tbl_style["margin-left"] = "auto"

    tbl_borders = _parse_tbl_borders(tbl_pr)
    for side in ("top", "left", "bottom", "right"):
        if side in tbl_borders:
            tbl_style[f"border-{side}"] = tbl_borders[side]

    tbl_cell_spacing = None
    if tbl_pr is not None:
        cs = tbl_pr.find("w:tblCellSpacing", NS)
        if cs is not None:
            tbl_cell_spacing = _twips_to_px(cs.attrib.get(f"{{{NS['w']}}}w"))
    tbl_style["border-collapse"] = "separate" if tbl_cell_spacing else "collapse"
    if tbl_cell_spacing:
        tbl_style["border-spacing"] = tbl_cell_spacing

    default_cell_margin = {}
    if tbl_pr is not None:
        tcm = tbl_pr.find("w:tblCellMar", NS)
        if tcm is not None:
            for side in ("top", "left", "bottom", "right"):
                m = tcm.find(f"w:{side}", NS)
                if m is not None:
                    px = _twips_to_px(m.attrib.get(f"{{{NS['w']}}}w"))
                    if px:
                        default_cell_margin[f"padding-{side}"] = px

    rows = tbl.findall("w:tr", NS)
    html_rows = []
    for tr_idx, tr in enumerate(rows):
        cells_html = []
        for tc_idx, tc in enumerate(tr.findall("w:tc", NS)):
            tc_pr = tc.find("w:tcPr", NS)
            cell_style = dict(default_cell_margin)

            # vMerge: skip continuation cells
            if tc_pr is not None:
                v_merge = tc_pr.find("w:vMerge", NS)
                if v_merge is not None:
                    if v_merge.attrib.get(f"{{{NS['w']}}}val", "continue") == "continue":
                        continue

            # Width
            if tc_pr is not None:
                tc_w = tc_pr.find("w:tcW", NS)
                if tc_w is not None:
                    w_type = tc_w.attrib.get(f"{{{NS['w']}}}type", "dxa")
                    w_val = tc_w.attrib.get(f"{{{NS['w']}}}w")
                    if w_type == "dxa" and w_val:
                        cell_style["width"] = _twips_to_px(w_val)

            # Borders
            if tc_pr is not None:
                tc_borders = _parse_tc_borders(tc_pr)
                if tc_borders:
                    cell_style.update(tc_borders)
                else:
                    if "insideH" in tbl_borders:
                        cell_style.setdefault("border-top", tbl_borders["insideH"])
                        cell_style.setdefault("border-bottom", tbl_borders["insideH"])
                    if "insideV" in tbl_borders:
                        cell_style.setdefault("border-left", tbl_borders["insideV"])
                        cell_style.setdefault("border-right", tbl_borders["insideV"])
            else:
                if "insideH" in tbl_borders:
                    cell_style.setdefault("border-top", tbl_borders["insideH"])
                    cell_style.setdefault("border-bottom", tbl_borders["insideH"])
                if "insideV" in tbl_borders:
                    cell_style.setdefault("border-left", tbl_borders["insideV"])
                    cell_style.setdefault("border-right", tbl_borders["insideV"])

            # Shading
            if tc_pr is not None:
                shd = tc_pr.find("w:shd", NS)
                if shd is not None:
                    fill = _normalize_color(shd.attrib.get(f"{{{NS['w']}}}fill"))
                    if fill:
                        cell_style["background-color"] = fill

            # Cell margins override
            if tc_pr is not None:
                tcm = tc_pr.find("w:tcMar", NS)
                if tcm is not None:
                    for side in ("top", "left", "bottom", "right"):
                        m = tcm.find(f"w:{side}", NS)
                        if m is not None:
                            px = _twips_to_px(m.attrib.get(f"{{{NS['w']}}}w"))
                            if px:
                                cell_style[f"padding-{side}"] = px

            # Vertical alignment
            if tc_pr is not None:
                v_align = tc_pr.find("w:vAlign", NS)
                if v_align is not None:
                    va = v_align.attrib.get(f"{{{NS['w']}}}val")
                    if va == "center":
                        cell_style["vertical-align"] = "middle"
                    elif va == "bottom":
                        cell_style["vertical-align"] = "bottom"

            # Span attributes
            span_attrs = ""
            if tc_pr is not None:
                grid_span = tc_pr.find("w:gridSpan", NS)
                if grid_span is not None:
                    cs = grid_span.attrib.get(f"{{{NS['w']}}}val")
                    if cs and cs != "1":
                        span_attrs += f' colspan="{cs}"'
                v_merge = tc_pr.find("w:vMerge", NS)
                if v_merge is not None:
                    if v_merge.attrib.get(f"{{{NS['w']}}}val", "continue") == "restart":
                        rs = _count_vmerge_rowspan(tbl, tr_idx, tc_idx)
                        if rs > 1:
                            span_attrs += f' rowspan="{rs}"'

            # Cell content
            cell_id_attr = f' data-id="{counter.next()}"' if counter is not None else ""
            cell_blocks = []
            for child in tc:
                tag = child.tag.split("}")[-1]
                if tag == "p":
                    cell_blocks.append(_render_paragraph(child, doc_defaults_rpr, doc_defaults_ppr, styles_map, images_map=images_map, zipf=zipf))
                elif tag == "tbl":
                    cell_blocks.append(_render_table(child, doc_defaults_rpr, doc_defaults_ppr, styles_map, counter=counter, images_map=images_map, zipf=zipf))
            cell_content = "".join(cell_blocks) if cell_blocks else "&nbsp;"
            cells_html.append(f'<td{cell_id_attr} style="{_build_style(cell_style)}"{span_attrs}>{cell_content}</td>')

        if cells_html:
            html_rows.append("<tr>" + "".join(cells_html) + "</tr>")

    return f'<table style="{_build_style(tbl_style)}">' + "".join(html_rows) + "</table>"


# ---------------------------------------------------------------------------
# Main conversion
# ---------------------------------------------------------------------------

def docx_to_html(docx_path: Path, output_path: Path, inject_ids: bool = False) -> None:
    """Convert a DOCX file to an HTML file."""
    with zipfile.ZipFile(docx_path) as zipf:
        doc_defaults_rpr, doc_defaults_ppr, styles_map = _parse_styles(zipf)

        doc_xml = zipf.read("word/document.xml").decode("utf-8")
        doc_root = ET.fromstring(doc_xml)
        body = doc_root.find("w:body", NS)
        if body is None:
            output_path.write_text("<html><body>Empty document</body></html>", encoding="utf-8")
            return

        page_style = {}
        sect_pr = body.find("w:sectPr", NS)
        if sect_pr is not None:
            page_style = _parse_sect_pr(sect_pr)
        paragraphs = body.findall("w:p", NS)
        if paragraphs:
            last_ppr = paragraphs[-1].find("w:pPr", NS)
            if last_ppr is not None:
                last_sect = last_ppr.find("w:sectPr", NS)
                if last_sect is not None and not page_style:
                    page_style = _parse_sect_pr(last_sect)

        images_map = _parse_relationships(zipf)
        counter = _Counter() if inject_ids else None
        pages = []
        current_blocks = []

        for child in body:
            tag = child.tag.split("}")[-1]
            if tag == "p":
                ppr = child.find("w:pPr", NS)
                sect_in_ppr = ppr.find("w:sectPr", NS) if ppr is not None else None
                current_blocks.append(_render_paragraph(child, doc_defaults_rpr, doc_defaults_ppr, styles_map, counter=counter, images_map=images_map, zipf=zipf))
                if sect_in_ppr is not None:
                    pages.append((_parse_sect_pr(sect_in_ppr) or page_style, list(current_blocks)))
                    current_blocks = []
            elif tag == "tbl":
                current_blocks.append(_render_table(child, doc_defaults_rpr, doc_defaults_ppr, styles_map, counter=counter, images_map=images_map, zipf=zipf))
            elif tag == "sectPr":
                page_style = _parse_sect_pr(child)

        if current_blocks:
            pages.append((page_style, current_blocks))

    body_html = []
    for ps, blocks in pages:
        style_str = _build_style(ps)
        body_html.append(f'<div class="page" style="{style_str}">' + "".join(blocks) + "</div>")

    html_doc = (
        '<!doctype html>\n<html lang="ko">\n<head>\n'
        '  <meta charset="utf-8">\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '  <title>DOCX Export</title>\n'
        "  <style>\n"
        "    body { font-family: sans-serif; line-height: 1.3; background: #eee; }\n"
        "    .page { position: relative; box-sizing: border-box; background: #fff; margin: 18px auto; border: 1px solid #d0d0d0; box-shadow: 0 2px 10px rgba(0,0,0,.15); page-break-after: always; overflow: auto; }\n"
        "    .page:last-child { page-break-after: auto; }\n"
        "    table { margin: 6px 0; }\n"
        "    td { vertical-align: top; }\n"
        "    p { margin: 0; }\n"
        "    br.page-break { page-break-before: always; }\n"
        "  </style>\n"
        "</head>\n<body>\n"
        + "".join(body_html)
        + "\n</body>\n</html>\n"
    )
    output_path.write_text(html_doc, encoding="utf-8")
