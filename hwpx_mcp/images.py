from __future__ import annotations

import base64
import concurrent.futures
import logging
import os
import re
from pathlib import Path
from typing import Dict, List
from xml.etree import ElementTree as ET

from .config import GEMINI_IMAGE_TIMEOUT_SECONDS

_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tiff"}
_GOV_STYLE_GUIDE = (
    "한국 정부 R&D 제안서 스타일의 깔끔한 인포그래픽. "
    "플랫 벡터, 파스텔 톤, 과도한 장식 없음. "
    "아이콘/도형 위주, 텍스트 최소화, 읽기 쉬운 레이아웃."
)
logger = logging.getLogger("process-gpt-office-mcp")


def _list_bindata_images(extract_dir: str) -> List[Path]:
    bindata_dir = Path(extract_dir) / "BinData"
    if not bindata_dir.exists():
        return []
    files = [
        p
        for p in bindata_dir.iterdir()
        if p.is_file() and p.suffix.lower() in _IMAGE_EXTS
    ]
    files.sort(key=lambda p: p.name.lower())
    return files


def count_bindata_images(extract_dir: str) -> int:
    return len(_list_bindata_images(extract_dir))


def apply_image_prompts_to_hwpx(extract_dir: str, image_prompts: List[Dict[str, str]]) -> int:
    if not image_prompts:
        logger.info("[이미지] 프롬프트 없음 — 건너뜀")
        return 0

    targets = _list_bindata_images(extract_dir)
    if not targets:
        logger.info("[이미지] BinData 이미지 없음 — 템플릿에 이미지 슬롯이 없습니다")
        return 0

    replaced = 0
    for idx, target in enumerate(targets):
        if idx >= len(image_prompts):
            break
        prompt = (image_prompts[idx].get("prompt") or "").strip()
        if not prompt:
            continue
        if generate_image_gemini(prompt, target):
            replaced += 1

    logger.info(
        "[이미지] 교체 완료: %d/%d",
        replaced,
        min(len(targets), len(image_prompts)),
    )
    return replaced


def _next_image_index(bindata_dir: Path) -> int:
    if not bindata_dir.exists():
        return 1
    max_idx = 0
    for path in bindata_dir.iterdir():
        if not path.is_file():
            continue
        m = re.match(r"image(\d+)", path.stem, re.IGNORECASE)
        if not m:
            continue
        try:
            max_idx = max(max_idx, int(m.group(1)))
        except ValueError:
            continue
    return max_idx + 1


def _write_xml_with_header(path: Path, tree: ET.ElementTree) -> None:
    tree.write(path, encoding="utf-8", xml_declaration=True)
    raw = path.read_text(encoding="utf-8")
    fixed = raw.replace(
        "<?xml version='1.0' encoding='utf-8'?>\r\n",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
        1,
    )
    if fixed == raw:
        fixed = raw.replace(
            "<?xml version='1.0' encoding='utf-8'?>\n",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
            1,
        )
    if fixed != raw:
        path.write_text(fixed, encoding="utf-8")


def _ensure_manifest_item(content_hpf: Path, item_id: str, href: str, media_type: str) -> None:
    if not content_hpf.exists():
        return
    tree = ET.parse(content_hpf)
    root = tree.getroot()
    ns = ""
    if root.tag.startswith("{") and "}" in root.tag:
        ns = root.tag.split("}")[0].strip("{")
    manifest_tag = f"{{{ns}}}manifest" if ns else "manifest"
    item_tag = f"{{{ns}}}item" if ns else "item"
    manifest = root.find(manifest_tag)
    if manifest is None:
        manifest = ET.SubElement(root, manifest_tag)
    for item in manifest.findall(item_tag):
        if item.get("href") == href or item.get("id") == item_id:
            return
    manifest.append(
        ET.Element(
            item_tag,
            {
                "id": item_id,
                "href": href,
                "media-type": media_type,
                "isEmbeded": "1",
            },
        )
    )
    _write_xml_with_header(content_hpf, tree)


def _ensure_meta_manifest_entry(meta_manifest: Path, full_path: str, media_type: str) -> None:
    if not meta_manifest.exists():
        return
    tree = ET.parse(meta_manifest)
    root = tree.getroot()
    ns = ""
    if root.tag.startswith("{") and "}" in root.tag:
        ns = root.tag.split("}")[0].strip("{")
    entry_tag = f"{{{ns}}}file-entry" if ns else "file-entry"
    for entry in root.findall(entry_tag):
        if entry.get(f"{{{ns}}}full-path") == full_path or entry.get("full-path") == full_path:
            return
    attrib = {}
    if ns:
        attrib[f"{{{ns}}}full-path"] = full_path
        attrib[f"{{{ns}}}media-type"] = media_type
    else:
        attrib["full-path"] = full_path
        attrib["media-type"] = media_type
    root.append(ET.Element(entry_tag, attrib))
    _write_xml_with_header(meta_manifest, tree)


def _build_pic_run(
    t_ns: str,
    bin_item_id: str,
    width_hwp: int = 33780,
    height_hwp: int = 18480,
    char_pr_id_ref: str | None = None,
    pic_id: str | None = None,
    inst_id: str | None = None,
    clip_scale: float = 3.125,
    caption_text: str | None = None,
    caption_para_pr_id_ref: str | None = None,
    caption_char_pr_id_ref: str | None = None,
) -> ET.Element:
    run_tag = f"{t_ns}run" if t_ns else "run"
    pic_tag = f"{t_ns}pic" if t_ns else "pic"
    hp_offset = f"{t_ns}offset" if t_ns else "offset"
    hp_orgsz = f"{t_ns}orgSz" if t_ns else "orgSz"
    hp_cursz = f"{t_ns}curSz" if t_ns else "curSz"
    hp_flip = f"{t_ns}flip" if t_ns else "flip"
    hp_rot = f"{t_ns}rotationInfo" if t_ns else "rotationInfo"
    hp_render = f"{t_ns}renderingInfo" if t_ns else "renderingInfo"
    hp_imgrect = f"{t_ns}imgRect" if t_ns else "imgRect"
    hp_imgclip = f"{t_ns}imgClip" if t_ns else "imgClip"
    hp_inmargin = f"{t_ns}inMargin" if t_ns else "inMargin"
    hp_effects = f"{t_ns}effects" if t_ns else "effects"
    hp_sz = f"{t_ns}sz" if t_ns else "sz"
    hp_pos = f"{t_ns}pos" if t_ns else "pos"
    hp_outmargin = f"{t_ns}outMargin" if t_ns else "outMargin"
    hp_shape_comment = f"{t_ns}shapeComment" if t_ns else "shapeComment"
    hp_caption = f"{t_ns}caption" if t_ns else "caption"
    hp_sublist = f"{t_ns}subList" if t_ns else "subList"
    hp_p = f"{t_ns}p" if t_ns else "p"
    hp_run = f"{t_ns}run" if t_ns else "run"
    hp_t = f"{t_ns}t" if t_ns else "t"
    hc_ns = "{http://www.hancom.co.kr/hwpml/2011/core}"
    hc_img = f"{hc_ns}img"
    hc_trans = f"{hc_ns}transMatrix"
    hc_sca = f"{hc_ns}scaMatrix"
    hc_rot = f"{hc_ns}rotMatrix"
    hc_pt0 = f"{hc_ns}pt0"
    hc_pt1 = f"{hc_ns}pt1"
    hc_pt2 = f"{hc_ns}pt2"
    hc_pt3 = f"{hc_ns}pt3"

    outer_run = ET.Element(run_tag)
    if char_pr_id_ref:
        outer_run.set("charPrIDRef", char_pr_id_ref)

    pic_attrib = {
        "zOrder": "12",
        "numberingType": "PICTURE",
        "textWrap": "SQUARE",
        "textFlow": "BOTH_SIDES",
        "lock": "0",
        "dropcapstyle": "None",
        "href": "",
        "groupLevel": "0",
        "reverse": "0",
    }
    if pic_id:
        pic_attrib["id"] = pic_id
    if inst_id:
        pic_attrib["instid"] = inst_id
    pic_elem = ET.SubElement(outer_run, pic_tag, pic_attrib)
    ET.SubElement(pic_elem, hp_offset, {"x": "0", "y": "0"})
    ET.SubElement(
        pic_elem, hp_orgsz, {"width": str(width_hwp), "height": str(height_hwp)}
    )
    ET.SubElement(pic_elem, hp_cursz, {"width": "0", "height": "0"})
    ET.SubElement(pic_elem, hp_flip, {"horizontal": "0", "vertical": "0"})
    ET.SubElement(pic_elem, hp_rot, {"angle": "0"})
    render = ET.SubElement(pic_elem, hp_render)
    ET.SubElement(render, hc_trans, {"e1": "1", "e2": "0", "e3": "0", "e4": "0", "e5": "1", "e6": "0"})
    ET.SubElement(render, hc_sca, {"e1": "1", "e2": "0", "e3": "0", "e4": "0", "e5": "1", "e6": "0"})
    ET.SubElement(render, hc_rot, {"e1": "1", "e2": "0", "e3": "0", "e4": "0", "e5": "1", "e6": "0"})
    img_rect = ET.SubElement(pic_elem, hp_imgrect)
    ET.SubElement(img_rect, hc_pt0, {"x": "0", "y": "0"})
    ET.SubElement(img_rect, hc_pt1, {"x": str(width_hwp), "y": "0"})
    ET.SubElement(img_rect, hc_pt2, {"x": str(width_hwp), "y": str(height_hwp)})
    ET.SubElement(img_rect, hc_pt3, {"x": "0", "y": str(height_hwp)})
    clip_right = int(width_hwp * clip_scale)
    clip_bottom = int(height_hwp * clip_scale)
    ET.SubElement(
        pic_elem,
        hp_imgclip,
        {"left": "0", "right": str(clip_right), "top": "0", "bottom": str(clip_bottom)},
    )
    ET.SubElement(pic_elem, hp_inmargin, {"left": "0", "right": "0", "top": "0", "bottom": "0"})
    ET.SubElement(pic_elem, hc_img, {"binaryItemIDRef": bin_item_id, "bright": "0", "contrast": "0", "effect": "REAL_PIC", "alpha": "0"})
    ET.SubElement(pic_elem, hp_effects)
    ET.SubElement(
        pic_elem,
        hp_sz,
        {
            "width": str(width_hwp),
            "widthRelTo": "ABSOLUTE",
            "height": str(height_hwp),
            "heightRelTo": "ABSOLUTE",
            "protect": "0",
        },
    )
    ET.SubElement(
        pic_elem,
        hp_pos,
        {
            "treatAsChar": "1",
            "affectLSpacing": "0",
            "flowWithText": "1",
            "allowOverlap": "0",
            "holdAnchorAndSO": "0",
            "vertRelTo": "PARA",
            "horzRelTo": "PARA",
            "vertAlign": "TOP",
            "horzAlign": "CENTER",
            "vertOffset": "0",
            "horzOffset": "0",
        },
    )
    ET.SubElement(pic_elem, hp_outmargin, {"left": "0", "right": "0", "top": "0", "bottom": "0"})
    if caption_text:
        caption_width = max(1000, int(width_hwp * 0.25))
        caption_elem = ET.SubElement(
            pic_elem,
            hp_caption,
            {
                "side": "BOTTOM",
                "fullSz": "0",
                "width": str(caption_width),
                "gap": "850",
                "lastWidth": str(width_hwp),
            },
        )
        sublist_elem = ET.SubElement(
            caption_elem,
            hp_sublist,
            {
                "id": "",
                "textDirection": "HORIZONTAL",
                "lineWrap": "BREAK",
                "vertAlign": "TOP",
                "linkListIDRef": "0",
                "linkListNextIDRef": "0",
                "textWidth": "0",
                "textHeight": "0",
                "hasTextRef": "0",
                "hasNumRef": "0",
            },
        )
        p_elem = ET.SubElement(
            sublist_elem,
            hp_p,
            {
                "id": "0",
                "paraPrIDRef": caption_para_pr_id_ref or "0",
                "styleIDRef": "0",
                "pageBreak": "0",
                "columnBreak": "0",
                "merged": "0",
            },
        )
        caption_run = ET.SubElement(
            p_elem,
            hp_run,
            {"charPrIDRef": caption_char_pr_id_ref or "0"},
        )
        t_elem = ET.SubElement(caption_run, hp_t)
        t_elem.text = caption_text
    comment = ET.SubElement(pic_elem, hp_shape_comment)
    comment.text = "Generated image"
    return outer_run


def _next_max_attr(root: ET.Element, attr_name: str) -> int:
    max_val = 0
    for elem in root.iter():
        raw = elem.get(attr_name)
        if not raw:
            continue
        try:
            max_val = max(max_val, int(raw))
        except ValueError:
            continue
    return max_val


def _detect_image_format(data: bytes) -> str | None:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8"):
        return "jpeg"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return "gif"
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return "webp"
    if data.startswith(b"BM"):
        return "bmp"
    return None


def _image_size_from_bytes(data: bytes) -> tuple[int | None, int | None]:
    fmt = _detect_image_format(data)
    if fmt == "png" and len(data) >= 24:
        width = int.from_bytes(data[16:20], "big")
        height = int.from_bytes(data[20:24], "big")
        return width, height
    if fmt == "jpeg":
        idx = 2
        while idx + 9 < len(data):
            if data[idx] != 0xFF:
                idx += 1
                continue
            marker = data[idx + 1]
            idx += 2
            if marker in (0xD8, 0xD9):
                continue
            if idx + 1 >= len(data):
                break
            seg_len = int.from_bytes(data[idx:idx + 2], "big")
            if seg_len < 2:
                break
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                if idx + 7 < len(data):
                    height = int.from_bytes(data[idx + 3:idx + 5], "big")
                    width = int.from_bytes(data[idx + 5:idx + 7], "big")
                    return width, height
                break
            idx += seg_len
    return None, None


def _next_paragraph_id(root: ET.Element) -> str:
    max_id = 0
    for p in root.iter():
        if p.tag.endswith("}p") or p.tag == "p":
            raw = p.get("id")
            if not raw:
                continue
            try:
                max_id = max(max_id, int(raw))
            except ValueError:
                continue
    return str(max_id + 1) if max_id else "1"


def _find_center_para_pr_id(header_path: Path) -> str | None:
    if not header_path.exists():
        return None
    try:
        tree = ET.parse(header_path)
        root = tree.getroot()
    except Exception:
        return None
    for elem in root.iter():
        if not elem.tag.endswith("paraPr"):
            continue
        pid = elem.get("id")
        if not pid:
            continue
        align = None
        for child in elem:
            if child.tag.endswith("align"):
                align = child.get("horizontal")
                break
        if align == "CENTER":
            return pid
    return None


def apply_image_markers_to_section(
    tree: ET.ElementTree,
    section_path: str,
    parent_map: dict,
    image_inserts: list[dict],
    t_ns: str = "",
    write_back: bool = False,
    image_export_dir: str | None = None,
    export_prefix: str | None = None,
) -> int:
    if not image_inserts:
        return 0

    extract_dir = Path(section_path).parent.parent
    bindata_dir = extract_dir / "BinData"
    bindata_dir.mkdir(parents=True, exist_ok=True)
    content_hpf = extract_dir / "Contents" / "content.hpf"
    meta_manifest = extract_dir / "META-INF" / "manifest.xml"

    inserted = 0
    next_idx = _next_image_index(bindata_dir)
    root = tree.getroot()
    next_pic_id = _next_max_attr(root, "id") + 1
    next_inst_id = _next_max_attr(root, "instid") + 1
    local_parent = {c: p for p in root.iter() for c in p}
    center_para_pr_id = _find_center_para_pr_id(extract_dir / "Contents" / "header.xml")

    image_bytes_by_index: list[bytes | None] = [None] * len(image_inserts)
    tasks: list[tuple[int, str, str]] = []
    for idx, item in enumerate(image_inserts):
        prompt = (item.get("prompt") or "").strip()
        ratio = (item.get("ratio") or "").strip()
        if ratio not in ("16:9", "4:3"):
            ratio = "16:9"
        if not prompt:
            continue
        tasks.append((idx, prompt, ratio))

    if tasks:
        logger.info("[이미지] Gemini 병렬 생성 시작 %d건 (max=3)", len(tasks))
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=3)
        try:
            future_map = {
                executor.submit(generate_image_gemini_bytes, prompt, ratio=ratio): idx
                for idx, prompt, ratio in tasks
            }
            done, not_done = concurrent.futures.wait(
                future_map.keys(),
                timeout=GEMINI_IMAGE_TIMEOUT_SECONDS,
            )
            for future in done:
                idx = future_map[future]
                try:
                    image_bytes_by_index[idx] = future.result()
                except Exception as exc:
                    logger.warning("[이미지] 생성 예외 — idx=%d (%s)", idx, exc)
            if not_done:
                logger.warning(
                    "[이미지] 생성 타임아웃 — 미완료 %d건 (%.0fs)",
                    len(not_done),
                    GEMINI_IMAGE_TIMEOUT_SECONDS,
                )
                for future in not_done:
                    future.cancel()
        finally:
            executor.shutdown(wait=False, cancel_futures=True)

    for idx, item in enumerate(image_inserts):
        node = item.get("node")
        prompt = (item.get("prompt") or "").strip()
        caption = (item.get("caption") or "").strip()
        ratio = (item.get("ratio") or "").strip()
        if ratio not in ("16:9", "4:3"):
            ratio = "16:9"
        if not node or not prompt:
            continue

        image_bytes = image_bytes_by_index[idx]
        if not image_bytes:
            logger.warning("[이미지] 생성 실패 — 삽입 건너뜀 (idx=%d)", idx)
            continue
        fmt = _detect_image_format(image_bytes) or "png"
        ext = "jpg" if fmt == "jpeg" else fmt
        media_type = f"image/{fmt}"
        bin_name = f"image{next_idx}.{ext}"
        bin_path = bindata_dir / bin_name
        bin_path.write_bytes(image_bytes)

        bin_item_id = f"image{next_idx}"
        _ensure_manifest_item(
            content_hpf, bin_item_id, f"BinData/{bin_name}", media_type
        )
        _ensure_meta_manifest_entry(
            meta_manifest, f"BinData/{bin_name}", media_type
        )

        target_run = None
        if node.t_elements:
            from .core.xml_utils import find_parent  # local import to avoid cycle
            target_run = find_parent(node.t_elements[0], local_parent, "run")
        if target_run is None and node.run_elements:
            target_run = node.run_elements[0]
        if target_run is None:
            next_idx += 1
            continue

        from .core.xml_utils import find_parent

        parent_p = find_parent(target_run, local_parent, "p")
        if parent_p is None:
            logger.warning("[이미지] 대상 문단 없음 — 삽입 건너뜀")
            next_idx += 1
            continue

        char_pr = target_run.get("charPrIDRef")
        width_px, height_px = _image_size_from_bytes(image_bytes)
        if ratio == "4:3":
            width_hwp = width_px * 24 if width_px else 52000
            height_hwp = int(width_hwp * 3 / 4)
        else:
            width_hwp = width_px * 24 if width_px else 60000
            height_hwp = int(width_hwp * 9 / 16)
        clip_scale = 3.125
        actual_ratio = "-"
        if width_px and height_px:
            actual_ratio = f"{width_px}:{height_px}"
        logger.info(
            "[이미지] fmt=%s px=%sx%s reqRatio=%s actualRatio=%s hwp=%sx%s clipScale=%s",
            fmt,
            width_px,
            height_px,
            ratio,
            actual_ratio,
            width_hwp,
            height_hwp,
            clip_scale,
        )
        pic_run = _build_pic_run(
            t_ns,
            bin_item_id,
            width_hwp=width_hwp,
            height_hwp=height_hwp,
            char_pr_id_ref=char_pr,
            pic_id=str(next_pic_id),
            inst_id=str(next_inst_id),
            clip_scale=clip_scale,
            caption_text=caption,
            caption_para_pr_id_ref=center_para_pr_id or parent_p.get("paraPrIDRef"),
            caption_char_pr_id_ref=char_pr,
        )

        p_tag = f"{t_ns}p" if t_ns else "p"
        pic_p = ET.Element(p_tag)
        pic_p.attrib.update(parent_p.attrib)
        if center_para_pr_id:
            pic_p.set("paraPrIDRef", center_para_pr_id)
        pic_p.set("id", _next_paragraph_id(root))
        pic_p.append(pic_run)

        parent = local_parent.get(parent_p)
        if parent is None:
            logger.warning("[이미지] 부모 노드 탐색 실패 — 삽입 건너뜀")
            next_idx += 1
            continue
        p_siblings = list(parent)
        try:
            p_idx = p_siblings.index(parent_p)
        except ValueError:
            p_idx = len(p_siblings) - 1
        parent.insert(p_idx + 1, pic_p)

        inserted += 1
        next_idx += 1
        next_pic_id += 1
        next_inst_id += 1

        if image_export_dir:
            export_dir = Path(image_export_dir)
            export_dir.mkdir(parents=True, exist_ok=True)
            prefix = f"{export_prefix}_" if export_prefix else ""
            export_name = f"{prefix}{bin_name}"
            (export_dir / export_name).write_bytes(image_bytes)

    if inserted:
        pic_count = sum(1 for elem in root.iter() if elem.tag.endswith("pic"))
        if pic_count == 0:
            logger.warning("[이미지] 경고: 삽입 후 pic 요소가 보이지 않습니다")
    if inserted and write_back:
        _write_xml_with_header(Path(section_path), tree)
    return inserted


def _get_client():
    try:
        from google import genai
    except Exception as exc:
        logger.warning("[이미지] google-genai import 실패: %s", exc)
        return None
    api_key = (os.getenv("GOOGLE_API_KEY") or "").strip()
    if not api_key:
        logger.warning("[이미지] GOOGLE_API_KEY 없음 — 이미지 생성 불가")
        return None
    return genai.Client(api_key=api_key)


def _get_image_model() -> str:
    from .config import GEMINI_IMAGE_MODEL

    return GEMINI_IMAGE_MODEL


def generate_image_gemini(prompt: str, output_path: Path) -> bool:
    image_bytes = generate_image_gemini_bytes(prompt)
    if not image_bytes:
        return False
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(image_bytes)
    return True


def generate_image_gemini_bytes(prompt: str, ratio: str | None = None) -> bytes | None:
    client = _get_client()
    if client is None:
        return None
    model = _get_image_model()
    style_prompt = f"{prompt}\n\n스타일 가이드: {_GOV_STYLE_GUIDE}"

    def _call_generate_content(config: dict | None):
        if config is None:
            return client.models.generate_content(model=model, contents=[style_prompt])
        return client.models.generate_content(model=model, contents=[style_prompt], config=config)

    aspect_ratio = ratio if ratio in ("16:9", "4:3") else None
    image_configs = []
    if aspect_ratio:
        image_configs = [
            {"aspect_ratio": aspect_ratio},
            {"aspect_ratio": aspect_ratio, "image_size": "2K"},
        ]
    else:
        image_configs = [{"image_size": "1K"}]

    attempts = [{"label": "default", "config": None}]
    for rc in (["IMAGE"], ["TEXT", "IMAGE"]):
        for cfg in image_configs:
            attempts.append(
                {
                    "label": f"modalities_{rc}_{cfg}",
                    "config": {
                        "response_modalities": rc,
                        "image_config": cfg,
                    },
                }
            )

    response = None
    last_error = None
    for attempt in attempts:
        try:
            response = _call_generate_content(attempt["config"])
            if attempt["label"] != "default":
                logger.info("[이미지] Gemini 생성 성공: %s", attempt["label"])
            break
        except Exception as exc:
            msg = str(exc).upper()
            logger.warning("[이미지] Gemini 생성 실패: %s (%s)", attempt["label"], msg)
            last_error = msg
            if "INVALID_ARGUMENT" not in msg and "400" not in msg:
                break
            continue

    if response is None:
        if last_error:
            logger.warning("[이미지] 최종 실패: %s", last_error)
        return None

    parts = getattr(response, "parts", None)
    if not parts:
        candidates = getattr(response, "candidates", []) or []
        if candidates and getattr(candidates[0], "content", None):
            parts = candidates[0].content.parts
    if not parts:
        return None

    for part in parts:
        inline_data = getattr(part, "inline_data", None)
        if not inline_data:
            continue
        data = getattr(inline_data, "data", None)
        if not data:
            continue
        if isinstance(data, str):
            return base64.b64decode(data)
        return data
    return None
