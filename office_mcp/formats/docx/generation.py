"""DOCX content generation via LLM — adapted from deep-research-custom, without deep-research dependencies."""
import asyncio
import json
import logging
import re
from typing import Any, Dict, List, Optional, Tuple

from ...agent.agent import _call_llm_json
from ...config import IMAGE_GENERATION_ENABLED

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Source / table formatting helpers
# ---------------------------------------------------------------------------

def _source_fingerprint(item: Dict[str, Any]) -> str:
    """Dedup key for a retrieved chunk — title + first 120 chars of content."""
    title = (item.get("title") or "").strip()
    content = (item.get("content") or "").strip()
    return f"{title}::{content[:120]}"


def _merge_sources(primary: List[Dict[str, Any]], extra: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """primary는 순서 유지, extra에서 fingerprint가 겹치지 않는 항목만 뒤에 덧붙임."""
    seen = {_source_fingerprint(s) for s in primary}
    merged = list(primary)
    for s in extra:
        fp = _source_fingerprint(s)
        if fp in seen:
            continue
        seen.add(fp)
        merged.append(s)
    return merged


def _format_sources_for_docx(sources: List[Dict[str, Any]], limit: int = 100) -> str:
    if not sources:
        return ""
    memento = [s for s in sources if s.get("source") == "memento"]
    others = [s for s in sources if s.get("source") != "memento"]
    blocks = []
    for item in (memento + others)[:limit]:
        title = item.get("title") or "Untitled"
        url = item.get("url") or ""
        content = (item.get("content") or "").strip()
        source_type = item.get("source") or "unknown"
        meta_lines = [f"source: {source_type}"]
        if url:
            meta_lines.append(f"url: {url}")
        blocks.append(f"[{title}]\n{chr(10).join(meta_lines)}\n{content}")
    return "\n\n".join(blocks)


def _format_table_template(headers: List, row_samples: List) -> str:
    if not row_samples:
        return ("| " + " | ".join(str(c) for c in headers) + " |") if headers else "N/A"
    cols = max(len(row_samples[0]) if row_samples and row_samples[0] else 0, len(headers) or 1)
    lines = []
    for row in row_samples:
        cells = (list(row) + [""] * cols)[:cols]
        lines.append("| " + " | ".join(str(c) for c in cells) + " |")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Structured output schemas (OpenAI json_schema / Gemini response_schema 호환)
# ---------------------------------------------------------------------------

_CONFIDENCE_NUMBER = {"type": "number", "minimum": 0, "maximum": 1}

_SCHEMA_TABLE_TYPE = {
    "name": "table_type_classification",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "type": {"type": "string", "enum": ["meta", "analytical", "mixed"]},
            "confidence": _CONFIDENCE_NUMBER,
            "rationale": {"type": "string"},
        },
        "required": ["type", "confidence", "rationale"],
    },
}

_SCHEMA_KEY_VALUE_NO_HEADER = {
    "name": "key_value_no_header_classification",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key_value_no_header": {"type": "boolean"},
            "confidence": _CONFIDENCE_NUMBER,
            "rationale": {"type": "string"},
        },
        "required": ["key_value_no_header", "confidence", "rationale"],
    },
}

_SCHEMA_OPTIONAL_SECTIONS = {
    "name": "optional_section_classification",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "sections": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "id": {"type": "string"},
                        "optional": {"type": "boolean"},
                        "explicit_optional": {"type": "boolean"},
                        "confidence": _CONFIDENCE_NUMBER,
                    },
                    "required": ["id", "optional", "explicit_optional", "confidence"],
                },
            },
        },
        "required": ["sections"],
    },
}

_SCHEMA_SECTION_ROLE = {
    "name": "section_role_classification",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "role": {"type": "string", "enum": ["container", "table_only", "body"]},
            "confidence": _CONFIDENCE_NUMBER,
            "rationale": {"type": "string"},
        },
        "required": ["role", "confidence", "rationale"],
    },
}

_SCHEMA_COVER = {
    "name": "cover_extraction",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "title_index": {"type": "integer"},
            "subtitle_index": {"type": "integer"},
            "title_text": {"type": "string"},
            "subtitle_text": {"type": "string"},
            "confidence": _CONFIDENCE_NUMBER,
            "rationale": {"type": "string"},
        },
        "required": ["title_index", "subtitle_index", "title_text", "subtitle_text", "confidence", "rationale"],
    },
}

_SCHEMA_SECTION_FILL = {
    "name": "section_fill",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "status": {"type": "string", "enum": ["fill", "partial", "omit"]},
            "content": {"type": "string"},
        },
        "required": ["status", "content"],
    },
}

_SCHEMA_TABLE_FILL = {
    "name": "table_fill",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "status": {"type": "string", "enum": ["fill", "partial", "omit"]},
            "rows": {
                "type": "array",
                "items": {"type": "array", "items": {"type": "string"}},
            },
        },
        "required": ["status", "rows"],
    },
}

_SCHEMA_IMAGES = {
    "name": "image_suggestions",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "images": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "section_id": {"type": "string"},
                        "prompt": {"type": "string"},
                        "caption": {"type": "string"},
                    },
                    "required": ["section_id", "prompt", "caption"],
                },
            },
        },
        "required": ["images"],
    },
}


# ---------------------------------------------------------------------------
# Async LLM wrapper
# ---------------------------------------------------------------------------

async def _chat_json(
    system_prompt: str,
    user_prompt: str,
    context: str = "",
    schema: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    logger.debug("LLM json call [%s] schema=%s", context, (schema or {}).get("name"))
    result = await asyncio.to_thread(_call_llm_json, system_prompt, user_prompt, 0.2, schema)
    return result if isinstance(result, dict) else {}


# ---------------------------------------------------------------------------
# Classification helpers
# ---------------------------------------------------------------------------

async def _classify_table_type(table: Dict[str, Any]) -> Dict[str, Any]:
    headers = table.get("headers") or []
    columns = table.get("columns") or len(headers)
    samples = table.get("row_samples") or []
    section_title = table.get("section_title") or ""
    data = await _chat_json(
        "You are a document analyst. Classify table type. Return JSON only.",
        (
            "다음 표 샘플을 보고 유형을 분류하세요.\n"
            "- JSON만 출력\n- keys: type, confidence, rationale\n"
            "- type은 meta | analytical | mixed 중 하나\n- confidence는 0~1\n\n"
            "분류 기준:\n- meta: 문서번호·작성일자·작성부서 등 2열 키-값 형식\n"
            "- analytical: 구분+항목 A/B/C, 연도+지표 A/B/C\n- mixed: 위 기준에 해당하지 않을 때\n\n"
            f"section_title: {section_title}\nheaders: {headers}\ncolumns: {columns}\n"
            f"row_samples: {json.dumps(samples, ensure_ascii=False)}\n"
        ),
        context=f"table_type:{table.get('id') or 'unknown'}",
        schema=_SCHEMA_TABLE_TYPE,
    )
    table_type = data.get("type")
    if table_type not in ("meta", "analytical", "mixed"):
        table_type = "mixed"
    return {"type": table_type, "confidence": float(data.get("confidence") or 0), "rationale": str(data.get("rationale") or "")}


async def _classify_key_value_no_header(table: Dict[str, Any]) -> Dict[str, Any]:
    headers = table.get("headers") or []
    columns = table.get("columns") or len(headers)
    samples = table.get("row_samples") or []
    section_title = table.get("section_title") or ""
    data = await _chat_json(
        "You are a document analyst. Decide if table has no header. Return JSON only.",
        (
            "다음 표가 헤더 없는 2열 키-값 표인지 판정하세요.\n"
            "- JSON만 출력\n- keys: key_value_no_header, confidence, rationale\n\n"
            f"section_title: {section_title}\nheaders: {headers}\ncolumns: {columns}\n"
            f"row_samples: {json.dumps(samples, ensure_ascii=False)}\n"
        ),
        context=f"kv_header:{table.get('id') or 'unknown'}",
        schema=_SCHEMA_KEY_VALUE_NO_HEADER,
    )
    return {
        "key_value_no_header": bool(data.get("key_value_no_header")),
        "confidence": float(data.get("confidence") or 0),
        "rationale": str(data.get("rationale") or ""),
    }


async def _classify_optional_sections(sections: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not sections:
        return []
    payload = [{"id": s.get("id"), "title": s.get("title"), "guidance": s.get("guidance") or []} for s in sections]
    data = await _chat_json(
        "You are a document analyst. Decide which sections are optional. Return JSON only.",
        (
            "아래 섹션 목록을 보고 선택/생략 가능한 섹션의 id만 추출하세요.\n"
            "- JSON만 출력\n- keys: sections (array of objects)\n"
            "- object keys: id, optional, explicit_optional, confidence\n\n"
            f"{json.dumps(payload, ensure_ascii=False)}\n"
        ),
        context="optional_sections",
        schema=_SCHEMA_OPTIONAL_SECTIONS,
    )
    sections_out = data.get("sections") if isinstance(data, dict) else None
    if not isinstance(sections_out, list):
        return []
    normalized = []
    for item in sections_out:
        if not isinstance(item, dict) or not item.get("id"):
            continue
        normalized.append({
            "id": str(item.get("id")),
            "optional": bool(item.get("optional")),
            "explicit_optional": bool(item.get("explicit_optional")),
            "confidence": float(item.get("confidence") or 0),
        })
    return normalized


async def _classify_section_role(section: Dict[str, Any], prev_title: str, next_title: str) -> Dict[str, Any]:
    section_id = section.get("id") or ""
    title = section.get("title") or section.get("id")
    guidance = section.get("guidance") or []
    template_excerpt = (section.get("template_excerpt") or "").strip()
    data = await _chat_json(
        "You are a document analyst. Classify the section role. Return JSON only.",
        (
            "다음 섹션이 본문을 생성해야 하는지 분류하세요.\n"
            "- JSON만 출력\n- keys: role, confidence, rationale\n"
            "- role: container | table_only | body\n\n"
            f"섹션 제목: {title}\n이전 섹션: {prev_title or 'N/A'}\n다음 섹션: {next_title or 'N/A'}\n"
            f"작성 지침: {(' / '.join(guidance)) if guidance else 'N/A'}\n"
            f"템플릿 예시: {template_excerpt or 'N/A'}\n"
            f"has_paragraphs: {str(bool(section.get('paragraph_indices'))).lower()}\n"
            f"has_tables: {str(bool(section.get('has_tables'))).lower()}\n"
        ),
        context=f"section_role:{section_id}:{title}",
        schema=_SCHEMA_SECTION_ROLE,
    )
    role = str(data.get("role") or "").strip().lower()
    if role not in ("container", "table_only", "body"):
        role = "body"
    return {"role": role, "confidence": float(data.get("confidence") or 0)}


# ---------------------------------------------------------------------------
# Content building helpers
# ---------------------------------------------------------------------------

async def _build_cover_output(cover: Dict[str, Any], query: str, outline: List[str]) -> Dict[str, Any]:
    paragraphs = cover.get("paragraphs") if isinstance(cover, dict) else None
    if not isinstance(paragraphs, list) or not paragraphs:
        return {}
    data = await _chat_json(
        "You are a document editor. Identify cover title/subtitle. Return JSON only.",
        (
            "다음은 문서 1페이지 내용입니다. 표지의 제목/부제를 판별하고 제목·부제를 생성하세요.\n"
            "- JSON만 출력\n"
            "- keys: title_index, subtitle_index, title_text, subtitle_text, confidence, rationale\n"
            "- title_index/subtitle_index는 반드시 제공된 paragraphs의 index 중에서 선택\n\n"
            f"[문서 1페이지]\n{json.dumps({'paragraphs': paragraphs, 'tables': cover.get('tables') or []}, ensure_ascii=False)}\n\n"
            f"[사용자 요청]\n{query}\n\n"
            f"[전체 개요]\n{json.dumps(outline, ensure_ascii=False)}\n"
        ),
        context="cover_title_subtitle",
        schema=_SCHEMA_COVER,
    )
    return data if isinstance(data, dict) else {}


async def _build_section_output(
    section: Dict[str, Any],
    query: str,
    base_sources: List[Dict[str, Any]],
    outline: Optional[List[str]] = None,
    section_rag=None,
) -> Tuple[str, Dict[str, Any]]:
    section_id = section.get("id") or ""
    title = section.get("title") or section_id
    optional = bool(section.get("optional"))
    guidance = section.get("guidance") or []
    template_excerpt = (section.get("template_excerpt") or "").strip()
    min_paragraphs = section.get("min_paragraphs")
    max_paragraphs = section.get("max_paragraphs")
    max_chars = section.get("max_chars")
    outline_text = "\n".join(outline or [])

    effective_sources = base_sources
    if section_rag is not None:
        section_query = " ".join(
            part for part in [title, " ".join(str(g) for g in guidance)] if part
        ).strip()
        try:
            extra = await section_rag(section_query)
        except Exception as exc:
            logger.warning("section RAG 검색 실패 (%s): %s", section_id, exc)
            extra = []
        if extra:
            effective_sources = _merge_sources(base_sources, extra)
            logger.info(
                "[docx섹션RAG] %s (+%d 청크, 총 %d)",
                section_id, len(extra), len(effective_sources),
            )
    sources_text = _format_sources_for_docx(effective_sources)
    data = await _chat_json(
        (
            "You are a report template filler. The user uploaded a docx report template and ran "
            "a deep research query. Write the given section to match the template tone and layout. "
            "Return JSON only."
        ),
        (
            f"사용자 요청:\n{query}\n\n섹션 제목: {title}\noptional: {str(optional).lower()}\n\n"
            f"작성 지침: {(' / '.join(guidance)) if guidance else 'N/A'}\n"
            + (
                f"길이 제한: {min_paragraphs or '제한 없음'}~{max_paragraphs or '제한 없음'}문단"
                + (f", {max_chars}자 이내" if max_chars else "")
                + "\n\n"
                if (min_paragraphs or max_paragraphs or max_chars)
                else ""
            ) +
            f"템플릿 섹션 예시(어투/형식 참고, 내용 복붙 금지):\n{template_excerpt or 'N/A'}\n\n"
            f"전체 개요(참고):\n{outline_text or 'N/A'}\n\n"
            "참고 소스는 source 구분(memento/web)이 포함되어 있습니다.\n"
            "memento(내부 문서) 소스를 우선 참고하고, 웹 소스는 보조로만 사용하세요.\n\n"
            f"참고 소스:\n{sources_text or 'N/A'}\n\n"
            "이 섹션에 들어갈 내용을 작성하세요.\n"
            "- JSON만 출력\n- keys: status, content\n"
            "- status는 fill | partial | omit 중 하나\n- optional=true 섹션은 자료 부족 시 omit\n"
        ),
        context=f"section:{section_id}:{title}",
        schema=_SCHEMA_SECTION_FILL,
    )
    if not isinstance(data, dict):
        data = {}
    if not optional and data.get("status") == "omit":
        data["status"] = "partial"
    if data.get("status") not in ("fill", "partial", "omit"):
        data["status"] = "omit" if optional else "partial"
    content_raw = data.get("content")
    content = "\n\n".join(str(i).strip() for i in content_raw if str(i).strip()) if isinstance(content_raw, list) else str(content_raw or "").strip()
    if content:
        paragraphs = [p.strip() for p in re.split(r"\n{2,}", content) if p.strip()]
        if isinstance(max_paragraphs, int) and max_paragraphs > 0 and len(paragraphs) > max_paragraphs:
            paragraphs = paragraphs[:max_paragraphs]
        content = "\n\n".join(paragraphs)
    if not content and not optional:
        content = "자료가 제한적이어서 간략 요약만 제공합니다."
    data["content"] = content
    return section_id, data


async def _build_table_output(
    table: Dict[str, Any],
    query: str,
    base_sources: List[Dict[str, Any]],
    outline: Optional[List[str]] = None,
    user_info: Optional[List[Dict[str, Any]]] = None,
    section_rag=None,
) -> Tuple[str, Dict[str, Any]]:
    table_id = table.get("id") or ""
    headers = table.get("headers") or []
    columns = table.get("columns") or len(headers)
    template_rows = table.get("row_samples") or []
    header_is_data = bool(table.get("header_is_data"))
    section_title = table.get("section_title") or ""
    outline_text = "\n".join(outline or [])

    effective_sources = base_sources
    if section_rag is not None:
        header_text = " ".join(str(h) for h in headers if h)
        section_query = " ".join(
            part for part in [section_title, header_text] if part
        ).strip()
        try:
            extra = await section_rag(section_query) if section_query else []
        except Exception as exc:
            logger.warning("table RAG 검색 실패 (%s): %s", table_id, exc)
            extra = []
        if extra:
            effective_sources = _merge_sources(base_sources, extra)
            logger.info(
                "[docx표RAG] %s (+%d 청크, 총 %d)",
                table_id, len(extra), len(effective_sources),
            )
    sources_text = _format_sources_for_docx(effective_sources)
    table_type = table.get("table_type") or "mixed"
    key_value_no_header = bool(table.get("key_value_no_header"))
    if key_value_no_header:
        header_is_data = True
    headers_str = " ".join(str(h) for h in headers)
    if not key_value_no_header:
        if any(kw in section_title for kw in ("지표", "분석", "비교")):
            table_type = "analytical"
        elif re.search(r"항목\s+[A-Z]|지표\s+[A-Z]|대상\s+[A-Z]", headers_str):
            table_type = "analytical"

    def _user_info_hint():
        if not user_info:
            return ""
        lines = []
        for u in user_info:
            name = (u.get("name") or "").strip()
            email = (u.get("email") or "").strip()
            if name:
                lines.append(f"- 이름(작성자): {name}")
            if email:
                lines.append(f"- 이메일: {email}")
        return "\n".join(lines)

    row_guidance = ""
    if key_value_no_header:
        hint = _user_info_hint()
        row_guidance = "- 이 표는 헤더 없는 2열 키-값 표입니다.\n- 첫 번째 열의 항목명은 템플릿 그대로 유지하세요.\n"
        if hint:
            row_guidance += f"\n- 아래 [작성자 정보]를 우선 활용하세요:\n{hint}\n"
    elif table_type == "meta":
        hint = _user_info_hint()
        row_guidance = "- 이 표는 문서 메타데이터(키-값) 형식입니다.\n- 첫 번째 열 항목명은 템플릿 그대로 유지하고, 두 번째 열만 채우세요.\n"
        if hint:
            row_guidance += f"\n- 아래 [작성자 정보]를 우선 활용하세요:\n{hint}\n"
    elif table_type == "analytical":
        row_guidance = "- 이 표는 분석/지표 표입니다.\n- 자리표시(연도, 지표 A/B/C 등)는 반드시 실제 값으로 교체하세요.\n"
    else:
        row_guidance = "- 이 표는 혼합 유형입니다. 자리표시는 실제 내용으로 교체하세요.\n"

    table_template_text = _format_table_template(headers, template_rows) if template_rows else f"헤더: {headers}\n열 수: {columns}"

    data = await _chat_json(
        (
            "You are a report template filler. The user uploaded a docx report template and ran "
            "a deep research query. Fill the given table with relevant data. Return JSON only."
        ),
        (
            f"사용자 요청:\n{query}\n\n표 위치 섹션: {section_title or 'N/A'}\n\n"
            f"표 양식(템플릿):\n{table_template_text}\n\n"
            f"{row_guidance}"
            f"전체 개요(참고):\n{outline_text or 'N/A'}\n\n"
            "memento(내부 문서) 소스를 우선 참고하고, 웹 소스는 보조로만 사용하세요.\n\n"
            f"참고 소스:\n{sources_text or 'N/A'}\n\n"
            "위 양식을 연구 결과로 채우세요.\n"
            "- JSON만 출력\n- keys: status, rows\n"
            "- status는 fill | partial | omit 중 하나\n- rows는 2차원 배열\n"
        ),
        context=f"table:{table_id}:{section_title}",
        schema=_SCHEMA_TABLE_FILL,
    )
    if not isinstance(data, dict):
        data = {}
    if data.get("status") not in ("fill", "partial", "omit"):
        data["status"] = "partial"
    rows = data.get("rows") or []
    if not isinstance(rows, list):
        rows = []

    if key_value_no_header and template_rows:
        if len(rows) > len(template_rows):
            rows = rows[:len(template_rows)]
        elif len(rows) < len(template_rows):
            rows.extend([["", ""]] * (len(template_rows) - len(rows)))
        for i, tmpl_row in enumerate(template_rows):
            if not isinstance(rows[i], list):
                rows[i] = ["", ""]
            if len(rows[i]) < columns:
                rows[i] = (rows[i] + [""] * columns)[:columns]
            rows[i][0] = tmpl_row[0]

    normalized_rows = []
    for row in rows:
        if not isinstance(row, list):
            continue
        new_row = []
        for cell in row[:columns]:
            text = re.sub(r"<br\s*/?>", " ", str(cell or "").strip(), flags=re.IGNORECASE)
            new_row.append(text)
        while len(new_row) < columns:
            new_row.append("")
        normalized_rows.append(new_row)

    data["rows"] = normalized_rows
    return table_id, data


async def _finalize_image_outputs(
    sections: List[Dict[str, Any]], sections_output: Dict[str, Dict[str, Any]],
    query: str, outline: List[str], sources_text: str,
    image_hints: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    if not IMAGE_GENERATION_ENABLED or not sections:
        return []
    section_ids = set()
    section_items = []
    for sec in sections:
        sec_id = sec.get("id") or ""
        if not sec_id:
            continue
        section_ids.add(sec_id)
        content_item = sections_output.get(sec_id) if isinstance(sections_output, dict) else None
        status = ""
        content = ""
        if isinstance(content_item, dict):
            status = str(content_item.get("status") or "").strip().lower()
            content = str(content_item.get("content") or "").strip()
        if status == "omit" or not content:
            continue
        section_items.append({"id": sec_id, "title": sec.get("title") or "", "content": content[:800]})
    if not section_items:
        return []

    data = await _chat_json(
        "You are a visual editor. Return JSON only.",
        (
            f"사용자 요청:\n{query}\n\n개요:\n{json.dumps(outline or [], ensure_ascii=False)}\n\n"
            f"이미지 힌트(참고):\n{json.dumps(image_hints or [], ensure_ascii=False)}\n\n"
            f"섹션 본문:\n{json.dumps(section_items, ensure_ascii=False)}\n\n"
            f"참고 소스 요약:\n{sources_text[:2000] if sources_text else 'N/A'}\n\n"
            "시각화 이미지가 필요한 경우에만 제안하세요.\n"
            "- JSON만 출력\n- key: images (array)\n- 각 항목: {section_id, prompt, caption}\n"
            "- section_id는 섹션 본문 목록의 id 중에서만 선택\n- 0~3개로 제한\n"
        ),
        context="image_finalize",
        schema=_SCHEMA_IMAGES,
    )
    images = data.get("images") if isinstance(data, dict) else None
    if not isinstance(images, list):
        return []
    return [
        {"section_id": item.get("section_id"), "prompt": str(item.get("prompt") or "").strip(), "caption": str(item.get("caption") or "").strip()}
        for item in images
        if isinstance(item, dict) and item.get("section_id") in section_ids and (item.get("prompt") or "").strip()
    ]


# ---------------------------------------------------------------------------
# Pre-classification
# ---------------------------------------------------------------------------

async def _apply_optional_sections(sections: List[Dict[str, Any]]) -> None:
    optional_info = await _classify_optional_sections(sections)
    if not optional_info:
        return
    optional_by_id = {item.get("id"): item for item in optional_info if item.get("id")}
    for sec in sections:
        meta = optional_by_id.get(sec.get("id"))
        if not meta:
            continue
        if bool(meta.get("explicit_optional")) or float(meta.get("confidence") or 0) >= 0.8:
            sec["optional"] = True


async def _apply_section_roles(sections: List[Dict[str, Any]]) -> None:
    def _has_paragraphs(s):
        return bool(s.get("paragraph_indices"))
    classify_targets = [
        s for s in sections
        if (not _has_paragraphs(s) and not s.get("has_tables") and s.get("has_children") is None)
        or (s.get("has_tables") and _has_paragraphs(s))
    ]
    if not classify_targets:
        return
    index_by_id = {s.get("id"): i for i, s in enumerate(sections) if s.get("id")}
    tasks = []
    for sec in classify_targets:
        idx = index_by_id.get(sec.get("id"))
        prev_title = sections[idx - 1].get("title") or "" if isinstance(idx, int) and idx > 0 else ""
        next_title = sections[idx + 1].get("title") or "" if isinstance(idx, int) and idx + 1 < len(sections) else ""
        tasks.append(_classify_section_role(sec, prev_title, next_title))
    results = await asyncio.gather(*tasks)
    for sec, meta in zip(classify_targets, results):
        sec["role"] = meta.get("role") or "body"


async def _apply_table_classification(tables: List[Dict[str, Any]]) -> None:
    if not tables:
        return
    table_classifications = await asyncio.gather(*[_classify_table_type(tbl) for tbl in tables])
    kv_candidates = [t for t in tables if not t.get("key_value_no_header") and (t.get("columns") == 2 or len(t.get("headers") or []) == 2)]
    kv_results = await asyncio.gather(*[_classify_key_value_no_header(t) for t in kv_candidates]) if kv_candidates else []
    kv_by_id = {(t.get("id") or ""): meta for t, meta in zip(kv_candidates, kv_results)}
    for tbl, meta in zip(tables, table_classifications):
        tbl_id = tbl.get("id") or ""
        kv_meta = kv_by_id.get(tbl_id)
        if kv_meta and kv_meta.get("key_value_no_header") and float(kv_meta.get("confidence") or 0) >= 0.7:
            tbl["key_value_no_header"] = True
            tbl["header_is_data"] = True
            tbl["table_type"] = "meta"
            tbl["table_type_confidence"] = 1.0
        elif tbl.get("key_value_no_header"):
            tbl["table_type"] = "meta"
            tbl["table_type_confidence"] = 1.0
        else:
            tbl["table_type"] = meta.get("type") or "mixed"
            tbl["table_type_confidence"] = float(meta.get("confidence") or 0)


def _should_skip_section_by_structure(sec: Dict[str, Any]) -> bool:
    title_text = str(sec.get("title") or "").strip()
    has_paragraphs = bool(sec.get("paragraph_indices"))
    if not title_text and not has_paragraphs and not sec.get("has_tables"):
        return True
    if has_paragraphs:
        return False
    if sec.get("has_children") is True:
        return True
    if sec.get("has_tables") is True:
        return True
    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

async def build_docx_output_from_schema(
    query: str,
    outline: List[str],
    sources: List[Dict[str, Any]],
    schema: Dict[str, Any],
    user_info: Optional[List[Dict[str, Any]]] = None,
    image_hints: Optional[List[Dict[str, Any]]] = None,
    tenant_id: Optional[str] = None,
    room_id: Optional[str] = None,
    proc_inst_id: Optional[str] = None,
) -> Dict[str, Any]:
    """사전 분류(계획) → 노드별 병렬 fill → 결과 취합."""
    await asyncio.gather(
        _apply_optional_sections(schema.get("sections") or []),
        _apply_section_roles(schema.get("sections") or []),
        _apply_table_classification(schema.get("tables") or []),
    )

    sections = schema.get("sections") or []
    tables = schema.get("tables") or []
    sections_to_fill = [
        s for s in sections
        if not _should_skip_section_by_structure(s) and s.get("role") not in ("container", "table_only")
    ]
    logger.info(
        "[docx병렬] fill 시작: 섹션 %d개(전체 %d), 표 %d개",
        len(sections_to_fill), len(sections), len(tables),
    )

    section_rag = None
    if (tenant_id or "").strip() and ((room_id or "").strip() or (proc_inst_id or "").strip()):
        from ...memento import search_section_context

        async def section_rag(section_query: str) -> List[Dict[str, Any]]:
            return await search_section_context(
                query=section_query,
                tenant_id=tenant_id.strip(),
                room_id=(room_id or "").strip() or None,
                proc_inst_id=(proc_inst_id or "").strip() or None,
                top_k=5,
            )

    cover_output = await _build_cover_output(schema.get("cover") or {}, query, outline)
    semaphore = asyncio.Semaphore(6)

    async def _guarded_section(sec):
        async with semaphore:
            return await _build_section_output(sec, query, sources, outline, section_rag=section_rag)

    async def _guarded_table(tbl):
        async with semaphore:
            return await _build_table_output(
                tbl, query, sources, outline, user_info=user_info, section_rag=section_rag,
            )

    section_results, table_results = await asyncio.gather(
        asyncio.gather(*[_guarded_section(s) for s in sections_to_fill]),
        asyncio.gather(*[_guarded_table(t) for t in tables]),
    )
    sections_output = {sec_id: data for sec_id, data in section_results if sec_id}
    tables_output = {tbl_id: data for tbl_id, data in table_results if tbl_id}
    sources_text = _format_sources_for_docx(sources)
    images = await _finalize_image_outputs(sections, sections_output, query, outline, sources_text, image_hints=image_hints)
    logger.info(
        "[docx병렬] fill 완료: 섹션 %d, 표 %d, 이미지 %d",
        len(sections_output), len(tables_output), len(images),
    )
    return {"sections": sections_output, "tables": tables_output, "images": images, "cover": cover_output}
