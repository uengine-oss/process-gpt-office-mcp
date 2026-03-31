import asyncio
import json
import logging
import re
import time
from typing import Any

from ..config import (
    MODEL_NAME, LLM_PROVIDER,
    OPENAI_API_KEY, OPENAI_TIMEOUT_SECONDS,
    GOOGLE_API_KEY,
)
from ..models import TextNode, TableSummary

logger = logging.getLogger("process-gpt-office-mcp")

# ── Provider 초기화 ──
_openai_client = None
_gemini_client = None

if LLM_PROVIDER == "gemini":
    if not GOOGLE_API_KEY:
        raise RuntimeError("LLM_PROVIDER=gemini이지만 GOOGLE_API_KEY가 없습니다")
    from google import genai
    from google.genai import types as genai_types
    _gemini_client = genai.Client(api_key=GOOGLE_API_KEY)
    logger.info("[LLM] Provider=gemini, model=%s", MODEL_NAME)
else:
    if not OPENAI_API_KEY:
        raise RuntimeError("LLM_PROVIDER=openai이지만 OPENAI_API_KEY가 없습니다")
    from openai import OpenAI
    _openai_client = OpenAI(api_key=OPENAI_API_KEY, timeout=OPENAI_TIMEOUT_SECONDS)
    logger.info("[LLM] Provider=openai, model=%s", MODEL_NAME)


_CHUNK_PLAN_SCHEMA = {
    "name": "chunk_plan",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "chunks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "chunk_id": {"type": "integer"},
                        "node_ids": {
                            "type": "array",
                            "items": {"type": "integer"},
                        },
                        "rationale": {"type": "string"},
                    },
                    "required": ["chunk_id", "node_ids", "rationale"],
                },
            }
        },
        "required": ["chunks"],
    },
}


_MAX_RETRIES = 3


def _extract_json_from_text(text: str) -> str:
    """Gemini 응답에서 JSON 블록을 추출한다 (```json ... ``` 또는 전체 텍스트)."""
    # ```json ... ``` 패턴
    m = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    # 전체가 JSON이면 그대로
    text = text.strip()
    if text.startswith("{"):
        return text
    return text


# ── OpenAI 호출 ──

def _call_openai_json(prompt_sys: str, prompt_user: str, temperature: float, user_content=None) -> dict:
    started = time.perf_counter()
    messages = [{"role": "system", "content": prompt_sys}]
    if user_content:
        messages.append({"role": "user", "content": user_content})
    else:
        messages.append({"role": "user", "content": prompt_user})

    last_exc = None
    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            resp = _openai_client.chat.completions.create(
                model=MODEL_NAME,
                messages=messages,
                temperature=temperature,
                response_format={"type": "json_object"},
            )
            elapsed = time.perf_counter() - started
            content = resp.choices[0].message.content or "{}"
            data = json.loads(content)
            data["_elapsed_s"] = round(elapsed, 3)
            return data
        except Exception as exc:
            last_exc = exc
            logger.warning("[LLM RETRY] attempt=%d/%d error=%s", attempt, _MAX_RETRIES, exc)
            if attempt < _MAX_RETRIES:
                time.sleep(1)
    raise last_exc


def _call_openai_text(prompt_sys: str, prompt_user: str, temperature: float) -> str:
    started = time.perf_counter()
    resp = _openai_client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": prompt_sys},
            {"role": "user", "content": prompt_user},
        ],
        temperature=temperature,
    )
    elapsed = time.perf_counter() - started
    return resp.choices[0].message.content or ""


# ── Gemini 호출 ──

def _call_gemini_json(prompt_sys: str, prompt_user: str, temperature: float, images_b64: list[str] | None = None) -> dict:
    started = time.perf_counter()

    contents = []
    if images_b64:
        import base64
        contents.append(prompt_user)
        for img in images_b64:
            img_bytes = base64.b64decode(img)
            contents.append(genai_types.Part.from_bytes(data=img_bytes, mime_type="image/png"))
    else:
        contents.append(prompt_user)

    config = genai_types.GenerateContentConfig(
        system_instruction=prompt_sys,
        temperature=temperature,
        response_mime_type="application/json",
    )

    last_exc = None
    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            resp = _gemini_client.models.generate_content(
                model=MODEL_NAME,
                contents=contents,
                config=config,
            )
            elapsed = time.perf_counter() - started
            raw_text = resp.text or "{}"
            json_text = _extract_json_from_text(raw_text)
            data = json.loads(json_text)
            data["_elapsed_s"] = round(elapsed, 3)
            return data
        except Exception as exc:
            last_exc = exc
            logger.warning("[LLM RETRY] attempt=%d/%d error=%s", attempt, _MAX_RETRIES, exc)
            if attempt < _MAX_RETRIES:
                time.sleep(1)
    raise last_exc


def _call_gemini_text(prompt_sys: str, prompt_user: str, temperature: float) -> str:
    config = genai_types.GenerateContentConfig(
        system_instruction=prompt_sys,
        temperature=temperature,
    )
    resp = _gemini_client.models.generate_content(
        model=MODEL_NAME,
        contents=prompt_user,
        config=config,
    )
    return resp.text or ""


# ── 통합 인터페이스 ──

def _call_llm_json(prompt_sys: str, prompt_user: str, temperature: float = 0.2) -> dict:
    started = time.perf_counter()
    logger.info("[LLM REQUEST] temp=%.2f\n[SYSTEM]\n%s\n[USER]\n%s", temperature, prompt_sys, prompt_user)

    if LLM_PROVIDER == "gemini":
        data = _call_gemini_json(prompt_sys, prompt_user, temperature)
    else:
        data = _call_openai_json(prompt_sys, prompt_user, temperature)

    elapsed = time.perf_counter() - started
    logger.info("[LLM RESPONSE] elapsed=%.2fs\n%s", elapsed, json.dumps(data, ensure_ascii=False, indent=2))
    return data


def _call_llm_vision_json(
    prompt_sys: str,
    prompt_user: str,
    images_b64: list[str],
    temperature: float = 0.2,
) -> dict:
    started = time.perf_counter()
    logger.info(
        "[LLM VISION REQUEST] temp=%.2f images=%d\n[SYSTEM]\n%s\n[USER text]\n%s...",
        temperature, len(images_b64), prompt_sys, prompt_user[:500],
    )

    if LLM_PROVIDER == "gemini":
        data = _call_gemini_json(prompt_sys, prompt_user, temperature, images_b64=images_b64)
    else:
        user_content: list[dict] = [{"type": "text", "text": prompt_user}]
        for img in images_b64:
            user_content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img}", "detail": "high"},
            })
        data = _call_openai_json(prompt_sys, prompt_user, temperature, user_content=user_content)

    elapsed = time.perf_counter() - started
    logger.info("[LLM VISION RESPONSE] elapsed=%.2fs\n%s", elapsed, json.dumps(data, ensure_ascii=False, indent=2))
    return data


def _call_llm_text(prompt_sys: str, prompt_user: str, temperature: float = 0.2) -> str:
    started = time.perf_counter()
    logger.info("[LLM REQUEST] temp=%.2f\n[SYSTEM]\n%s\n[USER]\n%s", temperature, prompt_sys, prompt_user)

    if LLM_PROVIDER == "gemini":
        content = _call_gemini_text(prompt_sys, prompt_user, temperature)
    else:
        content = _call_openai_text(prompt_sys, prompt_user, temperature)

    elapsed = time.perf_counter() - started
    logger.info("[LLM RESPONSE] elapsed=%.2fs\n%s", elapsed, content[:2000])
    return content


def _filter_llm_nodes(nodes: list[TextNode]) -> list[TextNode]:
    filtered: list[TextNode] = []
    for n in nodes:
        if n.type == "table_cell" and n.skip_fill:
            raw = (n.raw_text or "").strip()
            txt = (n.text or "").strip()
            if not raw and not txt:
                continue
        if n.type == "body_text" and n.skip_fill:
            raw = (n.raw_text or "").strip()
            txt = (n.text or "").strip()
            if not raw and not txt:
                continue
        filtered.append(n)
    return filtered


def _escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _detect_spacer_cells(table_nodes: list[TextNode]) -> set[int]:
    """[라벨][작은 빈 셀] + [큰 파란/가이드 셀] 패턴에서 작은 빈 셀 ID를 반환.

    정부과제 양식에서 소제목 옆 작은 빈 셀은 실제로 내용을 쓰는 곳이 아니라,
    그 아래 큰 셀(파란 가이드 텍스트)에 내용을 쓴다. 작은 빈 셀을 여백셀로 표시하여
    LLM이 채우지 않도록 한다.
    """
    spacer_ids: set[int] = set()
    # 행별로 그룹
    rows: dict[int, list[TextNode]] = {}
    for n in table_nodes:
        rows.setdefault(n.row, []).append(n)

    sorted_row_nums = sorted(rows.keys())
    for i, row_num in enumerate(sorted_row_nums):
        row_cells = rows[row_num]
        # 이 행에 bold 라벨 + 빈 셀이 있는지
        has_bold_label = any(
            "bold" in (n.style_summary or "") and (n.text or "").strip()
            for n in row_cells
        )
        empty_cells = [
            n for n in row_cells
            if not (n.text or n.raw_text or "").strip()
            and n.cell_height_mm and n.cell_height_mm <= 12
        ]
        if not has_bold_label or not empty_cells:
            continue

        # 다음 행(들)에 큰 셀(높이 ≥18mm) 또는 파란색 셀이 있는지
        for next_row_num in sorted_row_nums[i + 1:i + 3]:
            next_cells = rows.get(next_row_num, [])
            for nc in next_cells:
                h = nc.cell_height_mm or 0
                style = nc.style_summary or ""
                is_large = h >= 18
                is_blue = "#0000FF" in style or "color=#0000FF" in style.replace(" ", "")
                if is_large or is_blue:
                    for ec in empty_cells:
                        spacer_ids.add(ec.id)
                    break

    return spacer_ids


def _render_table_html(nodes: list[TextNode]) -> str:
    table_nodes = [n for n in nodes if n.type == "table_cell"]
    if not table_nodes:
        return ""
    tables: dict[int, list[TextNode]] = {}
    for n in table_nodes:
        tables.setdefault(n.table_idx, []).append(n)
    # 전체 테이블에서 spacer 셀 감지
    spacer_ids: set[int] = set()
    for t_nodes in tables.values():
        spacer_ids.update(_detect_spacer_cells(t_nodes))

    parts: list[str] = []
    for t_idx in sorted(tables.keys()):
        t_nodes = tables[t_idx]
        origin_map: dict[tuple[int, int], TextNode] = {}
        covered: set[tuple[int, int]] = set()
        max_row = 0
        max_col = 0
        for n in t_nodes:
            if n.row < 0 or n.col < 0:
                continue
            origin_map[(n.row, n.col)] = n
            row_span = max(1, n.cell_row_span)
            col_span = max(1, n.cell_col_span)
            for rr in range(n.row, n.row + row_span):
                for cc in range(n.col, n.col + col_span):
                    covered.add((rr, cc))
            max_row = max(max_row, n.row + row_span - 1)
            max_col = max(max_col, n.col + col_span - 1)

        if not origin_map:
            continue

        parts.append(f"<table data-table-idx=\"{t_idx}\">")
        for r in range(max_row + 1):
            row_has_origin = any(rr == r for (rr, _cc) in origin_map.keys())
            if not row_has_origin:
                continue
            parts.append("  <tr>")
            c = 0
            while c <= max_col:
                cell = origin_map.get((r, c))
                if cell is None:
                    if (r, c) in covered:
                        c += 1
                        continue
                    parts.append("    <td data-id=\"\"></td>")
                    c += 1
                    continue
                text = _escape_html((cell.text or cell.raw_text or "").strip())
                if not text:
                    h_mm = cell.cell_height_mm
                    if h_mm and h_mm <= 5:
                        text = "(여백셀)"
                    elif cell.id in spacer_ids:
                        text = "(여백셀)"
                    else:
                        text = f"(빈칸 {cell.cell_width_mm}x{cell.cell_height_mm}mm)" if cell.cell_width_mm and cell.cell_height_mm else "(빈칸)"
                size = ""
                if cell.cell_width_mm or cell.cell_height_mm:
                    w = cell.cell_width_mm
                    h = cell.cell_height_mm
                    size = f" data-size=\"{w}x{h}mm\""
                style = f' data-style="{cell.style_summary}"' if cell.style_summary else ""
                attrs = [f"data-id=\"{cell.id}\"{size}{style}"]
                if cell.cell_col_span > 1:
                    attrs.append(f"colspan=\"{cell.cell_col_span}\"")
                if cell.cell_row_span > 1:
                    attrs.append(f"rowspan=\"{cell.cell_row_span}\"")
                parts.append(f"    <td {' '.join(attrs)}>{text}</td>")
                c += max(1, cell.cell_col_span)
            parts.append("  </tr>")
        parts.append("</table>")
    return "\n".join(parts)


def _render_nodes_html(nodes: list[TextNode]) -> str:
    if not nodes:
        return ""
    table_idxs = {n.table_idx for n in nodes if n.type == "table_cell"}
    table_html_map = {
        t: _render_table_html([n for n in nodes if n.table_idx == t])
        for t in table_idxs if t >= 0
    }
    emitted_tables: set[int] = set()
    parts: list[str] = []
    for n in nodes:
        if n.type == "table_cell":
            if n.table_idx in table_html_map and n.table_idx not in emitted_tables:
                parts.append(table_html_map[n.table_idx])
                emitted_tables.add(n.table_idx)
            continue
        text = _escape_html((n.text or n.raw_text or "").strip())
        if not text:
            text = "(빈칸)"
        depth = n.depth if n.depth is not None else 0
        style = f' data-style="{n.style_summary}"' if n.style_summary else ""
        parts.append(f"<p data-id=\"{n.id}\" data-depth=\"{depth}\"{style}>{text}</p>")
    return "\n".join(parts)


render_nodes_html = _render_nodes_html
filter_llm_nodes = _filter_llm_nodes


def _render_context_summary(nodes: list[TextNode]) -> str:
    """앞뒤 청크 컨텍스트를 간결한 텍스트로 요약한다.

    data-id, data-style 등 분류용 속성을 제거하고,
    구조와 텍스트 내용만 간략히 보여준다.
    """
    if not nodes:
        return ""
    filtered = _filter_llm_nodes(nodes)
    if not filtered:
        return ""

    # 테이블별 요약
    table_summaries: dict[int, list[str]] = {}
    for n in filtered:
        if n.type == "table_cell" and n.table_idx >= 0:
            text = (n.text or n.raw_text or "").strip()
            if text and text not in ("(빈칸)", "(여백셀)"):
                lst = table_summaries.setdefault(n.table_idx, [])
                if len(lst) < 5:  # 셀 텍스트 최대 5개
                    truncated = text[:40] + ("..." if len(text) > 40 else "")
                    lst.append(truncated)

    emitted_tables: set[int] = set()
    lines: list[str] = []
    for n in filtered:
        if n.type == "table_cell":
            if n.table_idx not in emitted_tables and n.table_idx >= 0:
                emitted_tables.add(n.table_idx)
                samples = table_summaries.get(n.table_idx, [])
                count = sum(1 for nd in filtered if nd.type == "table_cell" and nd.table_idx == n.table_idx)
                if samples:
                    lines.append(f"[표] {count}셀 — {', '.join(samples)}")
                else:
                    lines.append(f"[표] {count}셀 (빈 표)")
            continue
        text = (n.text or n.raw_text or "").strip()
        if not text:
            continue
        truncated = text[:60] + ("..." if len(text) > 60 else "")
        lines.append(f"[본문] {truncated}")

    return "\n".join(lines)


def _render_nodes_for_plan(nodes: list[TextNode]) -> str:
    return "\n".join(n.display() for n in nodes)


async def agent_chunk_plan(
    nodes: list[TextNode],
    table_summaries: list[TableSummary] | None = None,
) -> list[dict]:
    if not nodes:
        return []
    doc_view = _render_nodes_for_plan(nodes)
    heading_candidates = []
    for n in nodes:
        if n.type != "body_text":
            continue
        text = (n.text or n.raw_text or "").strip()
        if not text:
            continue
        if "Heading" in (n.style_summary or ""):
            heading_candidates.append(text)
            continue
        if len(text) <= 30 and text[:1].isdigit():
            heading_candidates.append(text)
    if heading_candidates:
        heading_hint = "\n".join(f"- {h}" for h in heading_candidates[:20])
        heading_section = f"\n## 제목 후보\n{heading_hint}\n"
    else:
        heading_section = ""

    table_summary_section = ""
    if table_summaries:
        summaries_text = "\n".join(s.summary_text() for s in table_summaries)
        table_summary_section = f"""
## 표 구조 분석 정보 (코드로 사전 측정)
{summaries_text}
"""
    prompt_sys = "당신은 문서 구조 기반 청킹 전문가입니다."
    prompt_user = f"""다음은 HWPX 문서에서 추출한 노드 목록입니다.

## 목적
문서 흐름을 유지하면서 청크를 계획합니다.

## 규칙
1. 같은 table_idx(표)는 절대 분할하지 마세요.
2. 모든 노드 ID는 정확히 한 번 포함되어야 합니다.
3. 청크 순서는 문서 흐름을 따라야 합니다.

{table_summary_section}{heading_section}
## 노드 목록
{doc_view}

## 출력(JSON)
{{"chunks":[{{"chunk_id":0,"node_ids":[1,2,3],"rationale":"..."}}]}}
"""
    result = await asyncio.to_thread(
        _call_llm_json, prompt_sys, prompt_user, 0.2
    )
    chunks = []
    if isinstance(result, dict):
        chunks = result.get("chunks", [])
    if not isinstance(chunks, list):
        chunks = []
    return chunks


async def agent_analyze_chunk(
    nodes: list[TextNode],
    chunk_idx: int = 0,
    report_description: str = "",
    table_summaries: list[TableSummary] | None = None,
    chunk_image_b64: str = "",
    prev_chunk: list[TextNode] | None = None,
    next_chunk: list[TextNode] | None = None,
) -> dict:
    llm_nodes = _filter_llm_nodes(nodes)
    doc_view = _render_nodes_html(llm_nodes)
    prompt_sys = (
        "당신은 HWPX(한글) 양식 분석 전문가입니다. "
        "사용자가 HWPX 양식을 업로드했으며, 각 노드를 분석해 실제 값을 채울 계획을 만드는 단계입니다.\n\n"
        "★ 스크린샷이 제공되면 반드시 시각적 레이아웃을 우선 참고하세요. "
        "실제 문서가 어떻게 보이는지(셀 크기, 색상, 빈 영역, 테두리 유무)를 눈으로 확인하고 판단하세요.\n"
        "★ 앞뒤 청크가 참고용으로 제공될 수 있습니다. 현재 청크의 역할을 맥락적으로 파악하세요.\n"
        "★ 분류 결과는 현재 청크의 노드에 대해서만 출력하세요.\n"
        "★ 중요: 응답의 nodes 배열에서 id 값은 반드시 HTML의 data-id 속성값을 그대로 사용하세요. "
        "0부터 시작하는 순번이 아닙니다. 예: data-id=\"307\"이면 id: 307."
    )
    table_summary_section = ""
    if table_summaries:
        chunk_table_idxs = {n.table_idx for n in nodes if n.type == "table_cell"}
        relevant = [s for s in table_summaries if s.table_idx in chunk_table_idxs]
        if relevant:
            summaries_text = "\n".join(s.summary_text() for s in relevant)
            table_summary_section = f"""
## 표 구조 분석 정보 (코드로 사전 측정)
{summaries_text}
"""
    # 앞뒤 청크 컨텍스트 (요약만 — data-id 없이)
    context_section = ""
    if prev_chunk:
        prev_summary = _render_context_summary(prev_chunk)
        if prev_summary:
            context_section += f"""
## 이전 청크 (참고용 — 분류 대상 아님)
{prev_summary}
"""
    if next_chunk:
        next_summary = _render_context_summary(next_chunk)
        if next_summary:
            context_section += f"""
## 다음 청크 (참고용 — 분류 대상 아님)
{next_summary}
"""

    prompt_user = f"""다음은 HWPX 문서의 일부입니다.

## 프로젝트 정보
{report_description}
{context_section}
## 현재 청크 — 분류 대상 (이 청크의 노드만 분류하세요)
{doc_view}
{table_summary_section}

## 분류 카테고리
- label: 라벨/제목/여백용 빈 셀 (수정 불가). 항목명 옆의 작은 빈 셀(높이 ≤12mm)도 label.
- fixed: 이미 채워진 실제 고유값 (수정 불가)
- fill: 실제 내용을 채울 빈 필드. 본문의 빈칸이나 표의 넓은 빈 셀(높이 15mm+)이 해당.
- placeholder: 실제 내용으로 교체할 대상. 서식 마커(□, 가., ㅇ), 파란색/빨간색 가이드 텍스트 등. 일반 표 안의 파란 가이드 텍스트는 여기에 해당 — instruction이 아님.
- instruction: **독립된 "< 작성요령 >" 표** 안의 텍스트만 해당. tables_to_remove로 표 전체를 삭제할 때만 사용.
- image_placeholder: 이미지/그림/도식 자리
- index: 행 번호/인덱스

## 핵심 원칙
1) 스크린샷을 보고 판단하세요. 시각적으로 비어있는 영역, 색이 다른 텍스트, 셀 크기 비율을 확인하세요.
2) "작성 요령", "< 작성요령 >" 등이 포함된 **독립된 표**는 instruction + tables_to_remove에 추가. tables_to_remove에는 data-table-idx 값을 사용하세요.
3) 파란색/빨간색 가이드 텍스트가 일반 표(삭제하지 않는 표) 안에 있으면 → placeholder + replace. 이 텍스트는 instruction이 아님.
4) "□", "가.", "ㅇ" 같은 서식 마커는 placeholder + replace로 분류하세요.
5) 빈 본문 노드: 서식 마커나 소제목 뒤의 빈칸은 fill(내용 채울 자리), 그 외 여백용 빈 줄은 label(유지).

## few-shot 예시

예시1: 독립된 작성요령 표 → instruction + tables_to_remove
입력: <table data-table-idx="12"><tr><td>< 작성요령 ></td></tr><tr><td data-style="...color=#0000FF...">ㅇ 수요기관이 본 과제를 제안하게 된 배경...</td></tr></table>
출력: tables_to_remove=[12], 모든 셀 → instruction/keep/skip_fill=true

예시2: 서식 마커 + 빈칸 (본문)
입력:
<p>□</p>
<p>(빈칸)</p>
<p>가.</p>
<p>ㅇ</p>
출력: □,가.,ㅇ → placeholder/replace, (빈칸) → label

예시3: 소제목 + 빈칸 (내용 채울 자리)
입력:
<p data-style="...bold...">가. 시장 현황</p>
(작성요령 표 — 삭제됨)
<p>(빈칸)</p>
<p>(빈칸)</p>
<p data-style="...bold...">나. 시장 동향</p>
출력: 소제목 → label, 빈칸들 → fill/write, 다음 소제목 → label

예시4: [소제목][여백셀] + [color=#0000FF 셀] 패턴
입력:
<table>
  <tr>
    <td data-size="34x10mm" data-style="S:style=바탕글,size=1000,bold,color=#000000,align=JUSTIFY">6. 과제목표</td>
    <td data-size="131x10mm" data-style="S:style=바탕글,size=1000,color=#000000,align=JUSTIFY">(여백셀)</td>
  </tr>
  <tr>
    <td data-size="165x26mm" data-style="S:style=바탕글,size=1000,italic,color=#0000FF,align=JUSTIFY" colspan="3">* 과제의 목표 및 ... 정량적, 정성적으로 기술</td>
  </tr>
</table>
출력:
  bold 소제목 → label/keep/skip_fill=true
  (여백셀) → label/keep/skip_fill=true  ★ (여백셀)은 절대 채우지 않음
  color=#0000FF 셀 → placeholder/replace/skip_fill=false  ★ 이 셀의 가이드를 지우고 실제 내용 작성

## 출력(JSON)
★ tables_to_remove 배열에는 반드시 data-table-idx 속성값(절대 인덱스)을 넣으세요. 0부터 시작하는 상대 순번이 아닙니다.
★ tables_to_remove는 "작성요령/안내문" 표에만 사용하세요. 데이터를 채워야 하는 표에는 절대 사용하지 마세요.
{{"tables_to_remove":[],"nodes":[{{"id":1,"category":"fill","action":"write","skip_fill":false,"reason":"...","role":"summary_header","detail_node":2,"max_chars":30}}]}}
"""
    if chunk_image_b64:
        return await asyncio.to_thread(
            _call_llm_vision_json, prompt_sys, prompt_user, [chunk_image_b64], 0.2
        )
    return await asyncio.to_thread(
        _call_llm_json, prompt_sys, prompt_user, 0.2
    )


async def agent_generate_rag_queries(
    nodes: list[TextNode],
    analysis: dict,
    report_topic: str,
    report_description: str = "",
    num_queries: int = 5,
) -> list[str]:
    """청크 내용과 사용자 질문을 기반으로 RAG 검색 쿼리를 생성한다."""
    # 채워야 할 노드의 정보만 추출
    fill_nodes_info = []
    node_map = {n.id: n for n in nodes}
    for item in (analysis.get("nodes") or []):
        if item.get("skip_fill"):
            continue
        action = (item.get("action") or "").lower()
        if action not in ("write", "replace"):
            continue
        nid = item.get("id")
        node = node_map.get(nid)
        if node:
            label = item.get("reason") or item.get("category") or ""
            text = (node.text or "")[:50]
            fill_nodes_info.append(f"- [{label}] {text}" if text else f"- [{label}] (빈칸)")

    if not fill_nodes_info:
        return []

    fill_summary = "\n".join(fill_nodes_info[:20])
    prompt_sys = (
        "당신은 RAG 검색 쿼리 생성 전문가입니다. "
        "사용자가 문서 양식을 채우려 합니다. "
        "회사 내부 지식공간(memento)에서 필요한 정보를 검색하기 위한 쿼리를 생성하세요. "
        "검색 대상은 회사 내부 문서(기업 정보, 제안서, 사업계획서, 실적 등)입니다."
    )
    prompt_user = f"""## 사용자 요청
{report_topic}

## 보고서 설명
{report_description}

## 이 청크에서 작성해야 할 항목들
{fill_summary}

## 작업
위 항목들을 작성하기 위해 회사 내부 지식공간에서 검색할 쿼리를 정확히 {num_queries}개 생성하세요.
- 사용자 요청에 언급된 기업명, 기술명, 사업명 등을 쿼리에 포함하세요.
- 각 쿼리는 구체적이고 검색에 최적화된 형태로 작성하세요.
- 이건 RAG 검색용 쿼리입니다. 자연어 질문이 아니라 키워드 중심으로 작성하세요.

## 출력(JSON)
{{"queries": ["쿼리1", "쿼리2", ...]}}
"""
    result = await asyncio.to_thread(_call_llm_json, prompt_sys, prompt_user, 0.2)
    queries = result.get("queries") or []
    if not isinstance(queries, list):
        return []
    return [str(q).strip() for q in queries if str(q).strip()][:num_queries]


async def agent_fill_chunk(
    analysis: dict,
    nodes: list[TextNode],
    report_topic: str,
    report_description: str = "",
    reference_text: str = "",
) -> dict:
    llm_nodes = _filter_llm_nodes(nodes)
    doc_view = _render_nodes_html(llm_nodes)
    analysis_view = json.dumps(analysis, ensure_ascii=False, indent=2)
    prompt_sys = "당신은 한국어 보고서 양식 작성 전문가입니다."
    prompt_user = f"""아래 보고서 주제와 설명에 맞춰 HWPX 양식을 채우세요.

## 보고서 주제
{report_topic}

## 보고서 설명
{report_description}

## 참고 텍스트
{reference_text}

## 문서(HTML)
{doc_view}

## 분석 결과
{analysis_view}

## 작업
action이 write/replace이고 skip_fill=false인 노드만 작성합니다.
skip_fill=true인 노드는 작성하지 마세요.

★ 서식 마커("□", "가.", "ㅇ") 작성법:
  마커를 포함한 전체 텍스트를 작성하세요.
  "가." 노드에는 가/나/다 소제목과 각 소제목 아래 ㅇ 불릿 내용을 전부 포함하세요.
  ★ 들여쓰기로 계층 구조를 반드시 표현하세요:
    □ 는 최상위 (들여쓰기 없음)
    가./나./다. 는 2칸 들여쓰기
    ㅇ 는 4칸 들여쓰기
  예:
  "□ 추진배경 및 필요성\n\n  가. 현장 운영 환경\n    ㅇ 내용...\n    ㅇ 내용...\n\n  나. 디지털 전환 요구\n    ㅇ 내용...\n\n  다. 혁신 방향\n    ㅇ 내용..."

★ 표 셀은 data-size(WxH mm)에 맞게 분량을 조절하세요. 작은 셀은 한 문장, 큰 셀은 충분히.
★ 정부 R&D 제안서 수준의 분량으로 구체적이고 충실하게 작성하세요.

★ 이미지 삽입: 시스템 구성도, 아키텍처, 추진체계도, 기대효과 비교 등 시각자료가 효과적인 곳에는
  텍스트 중간이나 끝에 [IMAGE:이미지 설명 프롬프트] 마커를 삽입하세요.
  예: "... 시스템 구성은 다음과 같다.\n[IMAGE:AI 비전검사 시스템 아키텍처 구성도]\n위 구성을 통해..."
  단, 단순 텍스트(인력, 예산, 일정 등)에는 이미지 마커를 넣지 마세요.
  본문 노드(큰 영역)에만 넣고, 작은 표 셀에는 넣지 마세요.

## 출력(JSON)
★ id는 반드시 HTML의 data-id 속성값을 그대로 사용하세요 (0부터 시작하는 순번 아님).
{{"fills":[{{"id":307,"new_text":"..."}}]}}
"""
    return await asyncio.to_thread(
        _call_llm_json, prompt_sys, prompt_user, 0.3
    )


async def agent_judge_image_reference(
    nodes: list[TextNode],
    analysis: dict,
    report_topic: str,
    report_description: str = "",
    num_queries: int = 3,
) -> dict:
    """청크를 보고 회사 내부 지식공간의 기존 이미지를 첨부하면 좋을지 판단한다.

    Returns:
        {"need_images": bool, "queries": ["검색쿼리1", ...], "reason": "판단 이유"}
    """
    fill_nodes_info = []
    node_map = {n.id: n for n in nodes}
    for item in (analysis.get("nodes") or []):
        if item.get("skip_fill"):
            continue
        action = (item.get("action") or "").lower()
        if action not in ("write", "replace"):
            continue
        nid = item.get("id")
        node = node_map.get(nid)
        if node:
            label = item.get("reason") or item.get("category") or ""
            text = (node.text or "")[:80]
            fill_nodes_info.append(f"- [{label}] {text}" if text else f"- [{label}] (빈칸)")

    if not fill_nodes_info:
        return {"need_images": False, "queries": [], "reason": "작성할 노드 없음"}

    fill_summary = "\n".join(fill_nodes_info[:20])
    prompt_sys = (
        "당신은 보고서 작성 보조 AI입니다. "
        "회사 내부 지식공간에는 솔루션/제품 소개 문서에서 추출된 이미지(스크린샷, 아키텍처 다이어그램, 차트 등)가 저장되어 있습니다. "
        "보고서 청크를 보고, 기존 이미지를 첨부하면 설득력이나 이해도가 높아질지 판단하세요."
    )
    prompt_user = f"""## 보고서 주제
{report_topic}

## 보고서 설명
{report_description}

## 이 청크에서 작성해야 할 항목들
{fill_summary}

## 작업
위 항목들의 내용을 고려할 때, 회사 내부 지식공간에 저장된 기존 이미지(솔루션 스크린샷, 시스템 구성도, 아키텍처 다이어그램, 성과 차트 등)를 첨부하면 좋을지 판단하세요.

판단 기준:
- 기술 설명, 시스템 구성, 솔루션 소개 등 시각 자료가 효과적인 내용인가?
- 단순 텍스트(일정, 인력구성, 예산 등)만으로 충분한 내용이면 이미지 불필요
- 이미지가 필요하다고 판단되면, 적절한 이미지를 찾기 위한 검색 쿼리를 {num_queries}개 이내로 생성하세요

## 출력(JSON)
{{"need_images": true/false, "queries": ["쿼리1", ...], "reason": "판단 이유"}}
"""
    result = await asyncio.to_thread(_call_llm_json, prompt_sys, prompt_user, 0.2)
    need = result.get("need_images", False)
    queries = result.get("queries") or []
    if not isinstance(queries, list):
        queries = []
    queries = [str(q).strip() for q in queries if str(q).strip()][:num_queries]
    return {
        "need_images": bool(need),
        "queries": queries,
        "reason": result.get("reason", ""),
    }


async def agent_select_reference_images(
    candidates: list[dict],
    nodes: list[TextNode],
    analysis: dict,
    report_topic: str,
    max_select: int = 2,
    chunk_image_b64: str = "",
) -> list[dict]:
    """검색된 이미지 후보 중 청크에 삽입할 이미지를 AI가 선택한다.

    LLM이 스크린샷과 노드 정보를 보고, 어떤 이미지를 어떤 노드에 넣을지까지 결정한다.

    Args:
        candidates: [{"image_id", "image_url", "caption", ...}, ...]
        chunk_image_b64: 청크 스크린샷 (base64)
    Returns:
        선택된 이미지 목록 [{"image_url", "caption", "reason", "target_node_id"}, ...]
    """
    if not candidates:
        return []

    # 채울 노드 목록 (data-id, 타입, 크기, 텍스트 포함)
    fill_nodes_info = []
    node_map = {n.id: n for n in nodes}
    for item in (analysis.get("nodes") or []):
        action = (item.get("action") or "").lower()
        if action in ("write", "replace") and not item.get("skip_fill"):
            nid = item.get("id")
            node = node_map.get(nid)
            if not node:
                continue
            text = (node.text or "")[:60]
            ntype = getattr(node, "type", "")
            size_info = ""
            if ntype == "table_cell":
                w = round((getattr(node, "cell_width", 0) or 0) / 283.46) if getattr(node, "cell_width", 0) else 0
                h = round((getattr(node, "cell_height", 0) or 0) / 283.46) if getattr(node, "cell_height", 0) else 0
                if w > 0 and h > 0:
                    size_info = f" ({w}x{h}mm)"
            label = f"- data-id={nid} [{ntype}{size_info}] {text}" if text else f"- data-id={nid} [{ntype}{size_info}] (빈칸)"
            fill_nodes_info.append(label)

    candidates_text = ""
    for i, img in enumerate(candidates[:15]):
        caption_preview = (img.get("caption") or "")
        folder_name = img.get("drive_folder_name") or ""
        source_file = img.get("source_file_name") or ""
        file_name = img.get("file_name") or ""
        header = f"[{i}]"
        if folder_name:
            header += f" 폴더: {folder_name}"
        if source_file:
            header += f" | 출처 문서: {source_file}"
        if file_name:
            header += f" ({file_name})"
        candidates_text += f"{header}\n캡션: {caption_preview}\n\n"

    prompt_sys = (
        "당신은 보고서 작성 보조 AI입니다. "
        "검색된 이미지 후보 목록에서 보고서 청크에 삽입하면 효과적인 이미지를 선택하세요.\n"
        "★ 스크린샷이 제공되면 실제 양식 레이아웃을 반드시 참고하세요. "
        "작은 셀(요약, 기관명 등)에는 이미지를 넣지 마세요. "
        "이미지는 넓고 충분한 공간이 있는 노드(본문 영역, 큰 표 셀)에만 삽입하세요."
    )
    prompt_user = f"""## 보고서 주제
{report_topic}

## 이 청크에서 작성할 노드 목록 (data-id, 타입, 크기, 텍스트)
{chr(10).join(fill_nodes_info[:15])}

## 이미지 후보 목록
{candidates_text}

## 작업
위 후보 중 이 청크에 삽입하면 보고서의 설득력·이해도를 높일 이미지를 최대 {max_select}개 선택하세요.
- 출처 문서명과 캡션 내용을 함께 고려하여 청크 주제와 관련 있는 이미지만 선택하세요.
- 관련 없거나 품질이 낮아 보이면 하나도 선택하지 않아도 됩니다.
- 선택한 이미지마다:
  1. 보고서에 표시할 **짧은 캡션**(15~30자, 한국어)을 작성하세요.
  2. **삽입할 노드의 data-id**를 지정하세요.
     ★ 이미지를 넣기에 충분한 공간이 있는 노드만 선택하세요.
     ★ 작은 셀(요약 3줄, 기관명, 수행기간 등)에는 절대 넣지 마세요.
     ★ 본문 영역이나 큰 서술형 셀(추진전략, 사업범위, 목표 등)에 넣으세요.

## 출력(JSON)
{{"selected": [{{"index": 0, "target_node_id": 297, "reason": "선택 이유", "caption": "보고서용 짧은 캡션"}}]}}
"""
    if chunk_image_b64:
        result = await asyncio.to_thread(
            _call_llm_vision_json, prompt_sys, prompt_user, [chunk_image_b64], 0.2
        )
    else:
        result = await asyncio.to_thread(
            _call_llm_json, prompt_sys, prompt_user, 0.2
        )
    selected = result.get("selected") or []
    if not isinstance(selected, list):
        return []

    chosen = []
    for item in selected[:max_select]:
        if not isinstance(item, dict):
            continue
        idx = item.get("index")
        if not isinstance(idx, int) or idx < 0 or idx >= len(candidates):
            continue
        img = candidates[idx]
        short_caption = (item.get("caption") or "").strip()
        if not short_caption:
            short_caption = (item.get("reason") or "").strip()
        target_nid = item.get("target_node_id")
        chosen.append({
            "image_url": img.get("image_url") or "",
            "caption": short_caption,
            "image_id": img.get("image_id") or "",
            "reason": item.get("reason") or "",
            "target_node_id": target_nid,
        })
    return chosen
