from __future__ import annotations

import logging
from collections import Counter
from xml.etree import ElementTree as ET

from .xml_utils import tag, register_namespaces, find_parent


logger = logging.getLogger("process-gpt-office-mcp")


class StyleMaps:
    def __init__(self, charprs: dict, paraprs: dict, styles: dict) -> None:
        self.charprs = charprs
        self.paraprs = paraprs
        self.styles = styles


def load_style_maps(header_path: str) -> StyleMaps:
    register_namespaces(header_path)
    tree = ET.parse(header_path)
    root = tree.getroot()

    charprs: dict[str, dict] = {}
    paraprs: dict[str, dict] = {}
    styles: dict[str, dict] = {}

    for elem in root.iter():
        t = tag(elem)
        if t == "charPr":
            info = _parse_charpr(elem)
            if info.get("id") is not None:
                charprs[str(info["id"])] = info
        elif t == "paraPr":
            info = _parse_parapr(elem)
            if info.get("id") is not None:
                paraprs[str(info["id"])] = info
        elif t == "style":
            info = _parse_style(elem)
            if info.get("id") is not None:
                styles[str(info["id"])] = info

    return StyleMaps(charprs=charprs, paraprs=paraprs, styles=styles)


def resolve_style_for_runs(
    run_elems: list[ET.Element],
    parent_map: dict,
    style_maps: StyleMaps | None,
) -> tuple[dict, dict | None]:
    refs = {
        "style_id": None,
        "para_id": None,
        "char_id": None,
    }

    if run_elems:
        char_ids = []
        for r in run_elems:
            for k, v in r.attrib.items():
                if "charpridref" in k.lower():
                    char_ids.append(str(v))
                    break
        if char_ids:
            refs["char_id"] = Counter(char_ids).most_common(1)[0][0]

        p = find_parent(run_elems[0], parent_map, "p")
        if p is not None:
            for k, v in p.attrib.items():
                lk = k.lower()
                if "parapridref" in lk:
                    refs["para_id"] = str(v)
                elif "styleidref" in lk:
                    refs["style_id"] = str(v)

    if style_maps is None:
        return refs, None
    return refs, _compose_style_info(refs, style_maps)


def summarize_style(style_info: dict | None) -> str:
    if not style_info:
        return ""
    parts = []
    style_name = style_info.get("style_name")
    if style_name:
        parts.append(f"style={style_name}")

    char = style_info.get("char", {})
    if char.get("height"):
        parts.append(f"size={char['height']}")
    if char.get("bold"):
        parts.append("bold")
    if char.get("italic"):
        parts.append("italic")
    if char.get("textColor"):
        parts.append(f"color={char['textColor']}")

    para = style_info.get("para", {})
    align = para.get("align_horizontal")
    if align:
        parts.append(f"align={align}")

    if not parts:
        return ""
    return "S:" + ",".join(parts)


def log_style_summary(nodes) -> None:
    style_ids = Counter()
    para_ids = Counter()
    char_ids = Counter()
    missing_style = missing_para = missing_char = 0
    summaries = Counter()

    for n in nodes:
        refs = getattr(n, "style_refs", {}) or {}
        if refs.get("style_id") is not None:
            style_ids[refs["style_id"]] += 1
        if refs.get("para_id") is not None:
            para_ids[refs["para_id"]] += 1
        if refs.get("char_id") is not None:
            char_ids[refs["char_id"]] += 1

        miss = getattr(n, "style_missing", {}) or {}
        if miss.get("style"):
            missing_style += 1
        if miss.get("para"):
            missing_para += 1
        if miss.get("char"):
            missing_char += 1

        summary = getattr(n, "style_summary", "")
        if summary:
            summaries[summary] += 1

    if summaries or missing_style or missing_para or missing_char:
        lines = [
            "[스타일 매핑] 요약",
            f"- 참조 ID: style {len(style_ids)}종, para {len(para_ids)}종, char {len(char_ids)}종",
        ]
        if missing_style or missing_para or missing_char:
            lines.append(
                f"- 매핑 실패: style {missing_style}, para {missing_para}, char {missing_char}"
            )
        if summaries:
            lines.append("- 대표 스타일")
            for s, cnt in summaries.most_common(6):
                lines.append(f"  • {s} ({cnt}노드)")
        logger.info("\n".join(lines))


def _parse_charpr(elem: ET.Element) -> dict:
    info = dict(elem.attrib)
    if "id" not in info:
        for k, v in elem.attrib.items():
            if "id" in k.lower():
                info["id"] = v
                break

    for child in elem:
        ct = tag(child)
        if ct in ("bold", "italic", "underline", "strike"):
            info[ct] = True
        elif ct == "fontRef":
            for k, v in child.attrib.items():
                info[f"fontRef_{k}"] = v
    return info


def _parse_parapr(elem: ET.Element) -> dict:
    info = dict(elem.attrib)
    if "id" not in info:
        for k, v in elem.attrib.items():
            if "id" in k.lower():
                info["id"] = v
                break

    for child in elem:
        ct = tag(child)
        if ct == "align":
            info["align_horizontal"] = child.attrib.get("horizontal")
            info["align_vertical"] = child.attrib.get("vertical")
        elif ct == "lineSpacing":
            info["lineSpacing_type"] = child.attrib.get("type")
            info["lineSpacing_value"] = child.attrib.get("value")
        elif ct == "heading":
            info["heading_level"] = child.attrib.get("level")
        elif ct == "border":
            info["borderFillIDRef"] = child.attrib.get("borderFillIDRef")
    return info


def _parse_style(elem: ET.Element) -> dict:
    info = dict(elem.attrib)
    if "id" not in info:
        for k, v in elem.attrib.items():
            if "id" in k.lower():
                info["id"] = v
                break
    if "paraPrIDRef" in info:
        info["para_id"] = info["paraPrIDRef"]
    if "charPrIDRef" in info:
        info["char_id"] = info["charPrIDRef"]
    return info


def _compose_style_info(refs: dict, style_maps: StyleMaps) -> dict:
    style_info = {
        "style_id": refs.get("style_id"),
        "style_name": None,
        "char": {},
        "para": {},
    }

    base_style = None
    if refs.get("style_id") and refs["style_id"] in style_maps.styles:
        base_style = style_maps.styles[refs["style_id"]]
        style_info["style_name"] = base_style.get("name") or base_style.get("engName")
        base_char_id = str(base_style.get("char_id")) if base_style.get("char_id") is not None else None
        base_para_id = str(base_style.get("para_id")) if base_style.get("para_id") is not None else None
        if base_char_id and base_char_id in style_maps.charprs:
            style_info["char"].update(style_maps.charprs[base_char_id])
        if base_para_id and base_para_id in style_maps.paraprs:
            style_info["para"].update(style_maps.paraprs[base_para_id])

    char_id = refs.get("char_id")
    para_id = refs.get("para_id")
    if char_id and char_id in style_maps.charprs:
        style_info["char"].update(style_maps.charprs[char_id])
    if para_id and para_id in style_maps.paraprs:
        style_info["para"].update(style_maps.paraprs[para_id])

    return style_info
