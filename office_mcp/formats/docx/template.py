"""DOCX template parsing, schema extraction, content application, and Supabase upload."""
import json
import logging
import os
import re
import tempfile
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph
from supabase import create_client

from ...images import generate_image_gemini
from ...config import IMAGE_GENERATION_ENABLED

logger = logging.getLogger(__name__)

PLACEHOLDER_PATTERN = re.compile(r"\[[^\[\]]+?\]")
HEADING_NUMBER_PATTERN = re.compile(r"^\s*\d+(\.\d+)*\s+")
PARA_RANGE_PATTERN = re.compile(r"(\d+)\s*[~\-]\s*(\d+)\s*문단")
PARA_COUNT_PATTERN = re.compile(r"(\d+)\s*문단")
CHAR_COUNT_PATTERN = re.compile(r"(\d+)\s*자")
DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
STORAGE_BUCKET = "deep_research_files"


def _get_supabase():
    from ...config import SUPABASE_URL, SUPABASE_KEY
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("SUPABASE_URL 또는 SUPABASE_KEY가 설정되지 않았습니다.")
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def _extract_public_url(response: Any) -> Optional[str]:
    if not response:
        return None
    if isinstance(response, dict):
        if response.get("publicUrl"):
            return response.get("publicUrl")
        if response.get("public_url"):
            return response.get("public_url")
        data = response.get("data")
        if isinstance(data, dict) and data.get("publicUrl"):
            return data.get("publicUrl")
    return None


def _download_docx(url: str) -> Path:
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.write(resp.content)
    tmp.flush()
    tmp.close()
    return Path(tmp.name)


# 업로드 자료가 저장되는 supabase storage bucket (memento 와 동일).
# memento/storage/supabase_loader.py 의 ``storage.from_("files")`` 와 같은 bucket.
STORAGE_FILES_BUCKET = "files"


def _download_docx_from_storage(file_id: str) -> Path:
    """Supabase storage 에서 file_id (= ``knowledge_files.source_ref`` = storage path)
    로 직접 docx 본문을 받아 tempfile 로 저장.

    URL 변환·HTTP 라운드트립 없이 ``storage.from_("files").download(path)`` 한 번으로 끝.
    한글 파일명 URL 인코딩 이슈도 없음. signed URL 만료도 없음.
    """
    supabase = _get_supabase()
    try:
        data = supabase.storage.from_(STORAGE_FILES_BUCKET).download(file_id)
    except Exception as exc:
        raise RuntimeError(
            f"supabase storage download failed (bucket={STORAGE_FILES_BUCKET}, path={file_id!r}): {exc}"
        ) from exc
    if not data:
        raise RuntimeError(
            f"supabase storage download returned empty (bucket={STORAGE_FILES_BUCKET}, path={file_id!r})"
        )
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.write(data)
    tmp.flush()
    tmp.close()
    return Path(tmp.name)


def _has_outline_level(para) -> bool:
    try:
        ppr = para._p.pPr
        if ppr is None or ppr.outlineLvl is None:
            return False
        return int(ppr.outlineLvl.val) <= 1
    except Exception:
        return False


def _is_heading(para) -> bool:
    style_name = para.style.name if para.style else ""
    style_lower = str(style_name or "").lower()
    if style_lower.startswith("heading"):
        return True
    if "제목" in style_name or "표제" in style_name:
        return True
    return _has_outline_level(para)


def _heading_level(para) -> str:
    return str(para.style.name if para.style else "").strip()


def _heading_depth(para, para_text: str) -> Optional[int]:
    text = (para_text or "").strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)*)\s+", text)
    if m:
        return len(m.group(1).split("."))
    style_name = para.style.name if para.style else ""
    sm = re.search(r"(\d+)", str(style_name or ""))
    if sm:
        try:
            return int(sm.group(1))
        except Exception:
            return None
    return None


def _set_paragraph_text(para, text: str) -> None:
    if para.runs:
        para.runs[0].text = text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(text)


def _tag_sources_on_para(para, sources: Optional[List[Dict[str, Any]]]) -> None:
    """파라그래프 CT_P 요소에 data-sources 속성을 JSON으로 박는다.

    docx_to_html가 렌더링 시 이 속성을 읽어 HTML 데이터 속성으로 전달한다.
    Word 자체는 미지 속성을 무시하므로 안전.
    """
    if not sources:
        return
    try:
        para._p.set("data-sources", json.dumps(sources, ensure_ascii=False))
    except Exception as exc:
        logger.warning("data-sources 태깅 실패 (paragraph): %s", exc)


def _tag_sources_on_cell(cell, sources: Optional[List[Dict[str, Any]]]) -> None:
    """표 셀의 첫 파라그래프에 data-sources를 박는다."""
    if not sources:
        return
    paragraphs = list(getattr(cell, "paragraphs", []) or [])
    if not paragraphs:
        return
    _tag_sources_on_para(paragraphs[0], sources)


def _insert_paragraph_after(paragraph, text: str = ""):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = paragraph.__class__(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def _iter_block_items(parent):
    for child in parent.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _looks_like_heading(para) -> bool:
    if not para.text:
        return False
    return bool(HEADING_NUMBER_PATTERN.match(para.text.strip()))


def _normalize_paragraph_text(text: str) -> str:
    if text is None:
        return ""
    normalized = str(text).replace("\u00a0", " ")
    normalized = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def _extract_guidance(text: str) -> List[str]:
    if not text:
        return []
    guidance = []
    for match in PLACEHOLDER_PATTERN.findall(text):
        cleaned = match.strip("[]").strip()
        if cleaned:
            guidance.append(cleaned)
    if PARA_RANGE_PATTERN.search(text) or PARA_COUNT_PATTERN.search(text) or CHAR_COUNT_PATTERN.search(text):
        guidance.append(text.strip())
    return guidance


def _merge_length_hints(target: Dict[str, Any], guidance_texts: List[str]) -> None:
    for text in guidance_texts:
        rm = PARA_RANGE_PATTERN.search(text)
        if rm:
            target["min_paragraphs"] = int(rm.group(1))
            target["max_paragraphs"] = int(rm.group(2))
            continue
        cm = PARA_COUNT_PATTERN.search(text)
        if cm:
            count = int(cm.group(1))
            target.setdefault("min_paragraphs", count)
            target.setdefault("max_paragraphs", count)
        chm = CHAR_COUNT_PATTERN.search(text)
        if chm:
            target["max_chars"] = int(chm.group(1))


def _remove_table(table) -> None:
    try:
        tbl = table._element
        tbl.getparent().remove(tbl)
        tbl._tbl = tbl._element = None
    except Exception:
        pass


def extract_template_schema(doc: Document) -> Dict[str, Any]:
    sections: List[Dict[str, Any]] = []
    tables: List[Dict[str, Any]] = []
    current = None
    index_map = {para._p: idx for idx, para in enumerate(doc.paragraphs)}
    cover_paragraphs: List[Dict[str, Any]] = []
    cover_tables: List[Dict[str, Any]] = []
    cover_active = True

    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            para = block
            raw_text = para.text or ""
            para_text = _normalize_paragraph_text(raw_text)
            is_heading = _is_heading(para) or _looks_like_heading(para)
            if is_heading:
                cover_active = False
                section_id = f"section_{len(sections)+1}"
                current = {
                    "id": section_id,
                    "title": para_text,
                    "level": _heading_level(para) or "Heading",
                    "depth": _heading_depth(para, para_text),
                    "heading_index": index_map.get(para._p),
                    "paragraph_indices": [],
                    "template_texts": [],
                    "optional": False,
                }
                sections.append(current)
                continue
            if cover_active and para_text:
                cover_paragraphs.append({
                    "index": index_map.get(para._p),
                    "text": para_text,
                    "style": para.style.name if para.style else "",
                })
            if current is not None:
                idx = index_map.get(para._p)
                if idx is not None and para_text:
                    current["paragraph_indices"].append(idx)
                    current["template_texts"].append(para_text)
                    guidance = _extract_guidance(raw_text)
                    if guidance:
                        current.setdefault("guidance", [])
                        current["guidance"].extend(guidance)
                        _merge_length_hints(current, guidance)
        elif isinstance(block, Table):
            table = block
            headers = [cell.text.strip() for cell in table.rows[0].cells] if table.rows else []
            row_samples = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            if cover_active:
                cover_tables.append({"index": len(tables), "rows": row_samples})
            key_value_no_header = False
            if len(table.columns) == 2 and row_samples:
                label_rows = sum(
                    1 for r in row_samples
                    if isinstance(r, list) and len(r) >= 2
                    and _normalize_paragraph_text(r[0])
                    and (not _normalize_paragraph_text(r[1]) or _normalize_paragraph_text(r[1]) == _normalize_paragraph_text(r[0]))
                )
                if label_rows >= max(2, len(row_samples) - 1):
                    key_value_no_header = True
            header_is_data = not headers
            if not header_is_data:
                for cell_text in headers:
                    if PLACEHOLDER_PATTERN.search(cell_text) or re.search(r"YYYY|MM|DD|\\d{4}", cell_text):
                        header_is_data = True
                        break
            if key_value_no_header:
                header_is_data = True
            table_id = f"table_{len(tables)+1}"
            tables.append({
                "id": table_id,
                "index": len(tables),
                "headers": headers,
                "columns": len(table.columns),
                "rows": len(table.rows),
                "section_id": current.get("id") if current else None,
                "section_title": current.get("title") if current else None,
                "row_samples": row_samples,
                "header_is_data": header_is_data,
                "key_value_no_header": key_value_no_header,
            })

    section_ids_with_tables = {t.get("section_id") for t in tables if t.get("section_id")}
    for sec in sections:
        sec["has_tables"] = sec.get("id") in section_ids_with_tables
        template_texts = sec.get("template_texts") or []
        if template_texts:
            excerpt = "\n\n".join(str(t) for t in template_texts if t).strip()
            sec["template_excerpt"] = excerpt

    for i, sec in enumerate(sections):
        current_depth = sec.get("depth")
        has_children: Optional[bool] = None
        if isinstance(current_depth, int):
            has_children = False
            if i + 1 < len(sections):
                next_depth = sections[i + 1].get("depth")
                if isinstance(next_depth, int):
                    has_children = next_depth > current_depth
                else:
                    has_children = None
        sec["has_children"] = has_children

    return {
        "sections": sections,
        "tables": tables,
        "cover": {"paragraphs": cover_paragraphs, "tables": cover_tables},
    }


def summarize_template_schema(schema: Dict[str, Any], max_chars: int = 2000) -> str:
    lines: List[str] = []
    for sec in schema.get("sections", []):
        title = sec.get("title") or sec.get("id")
        optional = " optional" if sec.get("optional") else ""
        guidance = sec.get("guidance") or []
        guidance_text = f" guidance={'; '.join(guidance[:2])}" if guidance else ""
        lines.append(f"- SECTION {sec.get('id')}: {title}{optional}{guidance_text}")
    for tbl in schema.get("tables", []):
        headers = ", ".join(h for h in (tbl.get("headers") or []) if h)
        sec_title = tbl.get("section_title") or ""
        sec_hint = f" section={sec_title}" if sec_title else ""
        lines.append(f"- TABLE {tbl.get('id')}: headers=[{headers}]{sec_hint}")
    return "\n".join(lines)[:max_chars]


def load_template_schema_from_url(template_url: str) -> Dict[str, Any]:
    if not template_url:
        return {"sections": [], "tables": []}
    try:
        doc = Document(str(_download_docx(template_url)))
        return extract_template_schema(doc)
    except Exception as e:
        logger.error("템플릿 스키마 로드 실패: %s", e)
        return {"sections": [], "tables": []}


def load_template_schema_summary_from_url(template_url: str) -> str:
    if not template_url:
        return ""
    try:
        doc = Document(str(_download_docx(template_url)))
        schema = extract_template_schema(doc)
        return summarize_template_schema(schema)
    except Exception as e:
        logger.error("템플릿 스키마 요약 실패: %s", e)
        return ""


def apply_schema_output(
    doc: Document,
    schema: Dict[str, Any],
    output: Dict[str, Any],
    report_id: Optional[str] = None,
) -> None:
    cover_output = output.get("cover") if isinstance(output, dict) else None
    if isinstance(cover_output, dict):
        title_index = cover_output.get("title_index")
        subtitle_index = cover_output.get("subtitle_index")
        title_text = (cover_output.get("title_text") or "").strip()
        subtitle_text = (cover_output.get("subtitle_text") or "").strip()
        if isinstance(title_index, int) and 0 <= title_index < len(doc.paragraphs) and title_text:
            _set_paragraph_text(doc.paragraphs[title_index], title_text)
        if isinstance(subtitle_index, int) and 0 <= subtitle_index < len(doc.paragraphs) and subtitle_text:
            _set_paragraph_text(doc.paragraphs[subtitle_index], subtitle_text)

    sections_output = output.get("sections") if isinstance(output, dict) else None
    if isinstance(sections_output, dict):
        section_index = {s.get("id"): s for s in schema.get("sections", [])}
        for sec in schema.get("sections", []):
            new_title = sec.get("mapped_title")
            heading_index = sec.get("heading_index")
            if new_title and isinstance(heading_index, int) and heading_index < len(doc.paragraphs):
                _set_paragraph_text(doc.paragraphs[heading_index], str(new_title))
        # section_sources: section_id → [meta]. generation output의 section_sources 키에서 옴.
        section_sources_map = output.get("section_sources") if isinstance(output, dict) else None
        if not isinstance(section_sources_map, dict):
            section_sources_map = {}
        for sec_id, content in sections_output.items():
            sec = section_index.get(sec_id)
            if not sec:
                continue
            status = None
            text = ""
            if isinstance(content, dict):
                status = (content.get("status") or "").strip().lower()
                text = str(content.get("content") or "").strip()
            elif isinstance(content, str):
                text = content.strip()
            if status == "omit":
                status = "partial"
            if not text:
                continue
            indices = sec.get("paragraph_indices") or []
            target_para = None
            if indices:
                target_para = doc.paragraphs[indices[0]]
                _set_paragraph_text(target_para, text)
                # 형제 placeholder는 건드리지 않는다 — LLM이 섹션 1개로 처리한 경우
                # 나머지 placeholder는 원본 안내문 그대로 남겨 사용자가 편집할 수 있게 한다.
            else:
                target_para = doc.add_paragraph(text)
            # 출처 메타 태깅 (있을 때만)
            _tag_sources_on_para(target_para, section_sources_map.get(sec_id))

    tables_output = output.get("tables") if isinstance(output, dict) else None
    if isinstance(tables_output, dict):
        table_index = {t.get("id"): t for t in schema.get("tables", [])}
        # table_sources: table_id → [meta].
        table_sources_map = output.get("table_sources") if isinstance(output, dict) else None
        if not isinstance(table_sources_map, dict):
            table_sources_map = {}
        for tbl_id, tbl_content in tables_output.items():
            tbl_meta = table_index.get(tbl_id)
            if not tbl_meta:
                continue
            idx = tbl_meta.get("index")
            if idx is None or idx >= len(doc.tables):
                continue
            table = doc.tables[idx]
            tbl_sources = table_sources_map.get(tbl_id)
            # 표의 첫 셀에 출처 태깅 (tooltip이 여기 달림)
            if tbl_sources and table.rows and table.rows[0].cells:
                _tag_sources_on_cell(table.rows[0].cells[0], tbl_sources)
            header_is_data = bool(tbl_meta.get("header_is_data"))
            status = None
            rows = []
            headers_override = None
            if isinstance(tbl_content, dict):
                status = (tbl_content.get("status") or "").strip().lower()
                rows = tbl_content.get("rows") or []
                headers_override = tbl_content.get("headers")
            if not isinstance(rows, list):
                continue
            if status == "omit":
                _remove_table(table)
                continue

            header_row = table.rows[0] if table.rows else None
            data_template_row = table.rows[1] if len(table.rows) > 1 else header_row
            rows_for_data = rows

            if not header_is_data:
                header_values = None
                if isinstance(headers_override, list) and headers_override:
                    header_values = headers_override
                elif rows:
                    header_values = rows[0] if isinstance(rows[0], list) else []
                if header_row is not None and header_values is not None:
                    for i, cell in enumerate(header_row.cells):
                        value = header_values[i] if i < len(header_values) else ""
                        if cell.paragraphs:
                            _set_paragraph_text(cell.paragraphs[0], str(value))
                            for extra in cell.paragraphs[1:]:
                                _set_paragraph_text(extra, "")
                        else:
                            cell.text = str(value)
                    rows_for_data = rows if header_values is headers_override else rows[1:]

            template_row_xml = deepcopy(data_template_row._tr) if data_template_row is not None else None
            while len(table.rows) > 1:
                table._tbl.remove(table.rows[1]._tr)

            for row_index, row_data in enumerate(rows_for_data):
                if not isinstance(row_data, list):
                    continue
                if row_index == 0 and header_is_data and header_row is not None:
                    new_row = header_row
                elif template_row_xml is None:
                    table.add_row()
                    new_row = table.rows[-1]
                else:
                    table._tbl.append(deepcopy(template_row_xml))
                    new_row = table.rows[-1]
                for i, cell in enumerate(new_row.cells):
                    value = row_data[i] if i < len(row_data) else ""
                    if cell.paragraphs:
                        _set_paragraph_text(cell.paragraphs[0], str(value))
                        for extra in cell.paragraphs[1:]:
                            _set_paragraph_text(extra, "")
                    else:
                        cell.text = str(value)

    images_output = output.get("images") if isinstance(output, dict) else None
    if isinstance(images_output, list):
        section_index = {s.get("id"): s for s in schema.get("sections", [])}
        image_jobs = []
        for index, item in enumerate(images_output, start=1):
            if not isinstance(item, dict):
                continue
            section_id = item.get("section_id")
            prompt = (item.get("prompt") or "").strip()
            caption = (item.get("caption") or "").strip()
            if not section_id or not prompt or section_id not in section_index:
                continue
            image_jobs.append((section_id, prompt, caption, f"image-{index}.png"))

        def _render_image(job):
            section_id, prompt, caption, filename = job
            try:
                tmp_dir = Path(tempfile.mkdtemp())
                img_path = tmp_dir / filename
                if generate_image_gemini(prompt, img_path):
                    public_url = _upload_image_to_storage(img_path, report_id or "default", filename) if report_id else None
                    return (section_id, img_path, caption, public_url)
            except Exception:
                pass
            return None

        if image_jobs and IMAGE_GENERATION_ENABLED:
            with ThreadPoolExecutor(max_workers=5) as executor:
                results = [r for r in executor.map(_render_image, image_jobs) if r]
            for section_id, img_path, caption, _url in results:
                sec = section_index.get(section_id)
                if not sec:
                    continue
                indices = sec.get("paragraph_indices") or []
                new_para = _insert_paragraph_after(doc.paragraphs[indices[-1]]) if indices else doc.add_paragraph()
                try:
                    new_para.add_run().add_picture(str(img_path), width=Inches(5.5))
                    if caption:
                        doc.add_paragraph(caption)
                except Exception:
                    continue


def upload_docx_to_storage(file_path: Path, storage_path: str) -> Optional[str]:
    supabase = _get_supabase()
    file_bytes = file_path.read_bytes()
    safe_path = storage_path.lstrip("/")
    try:
        resp = supabase.storage.from_(STORAGE_BUCKET).upload(
            safe_path, file_bytes, {"content-type": DOCX_CONTENT_TYPE, "upsert": "true"}
        )
        if hasattr(resp, "path") and not resp.path:
            logger.error("storage 업로드 실패: path 없음 %s", resp)
            return None
        public = supabase.storage.from_(STORAGE_BUCKET).get_public_url(safe_path)
        url = _extract_public_url(public)
        if url:
            return url
    except Exception as e:
        logger.error("storage 업로드 실패: %s", e)
        return None
    from ...config import SUPABASE_URL as _sb_url
    base_url = _sb_url.rstrip("/")
    if base_url:
        return f"{base_url}/storage/v1/object/public/{STORAGE_BUCKET}/{quote(safe_path, safe='/-_.')}"
    return None


def _upload_image_to_storage(file_path: Path, report_id: str, filename: str) -> Optional[str]:
    supabase = _get_supabase()
    file_bytes = file_path.read_bytes()
    storage_path = f"deep-research/{report_id}/{Path(filename).name}"
    try:
        supabase.storage.from_(STORAGE_BUCKET).upload(
            storage_path, file_bytes, {"content-type": "image/png", "upsert": "true"}
        )
        public = supabase.storage.from_(STORAGE_BUCKET).get_public_url(storage_path)
        url = _extract_public_url(public)
        if url:
            return url
    except Exception as e:
        logger.error("image 업로드 실패: %s", e)
    from ...config import SUPABASE_URL as _sb_url
    base_url = _sb_url.rstrip("/")
    if base_url:
        return f"{base_url}/storage/v1/object/public/{STORAGE_BUCKET}/{quote(storage_path, safe='/-_.')}"
    return None
