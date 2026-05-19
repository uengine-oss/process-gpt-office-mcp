import asyncio
import json
import logging
import time
from typing import Any

from ..models import TextNode, TableSummary
from .llm_provider import create_provider

logger = logging.getLogger(__name__)

# ── Provider 초기화 (LLM_PROVIDER 설정에 따라 자동 선택) ──
_provider = create_provider()


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


# ── 통합 인터페이스 (provider에 위임) ──

def _call_llm_json(
    prompt_sys: str,
    prompt_user: str,
    temperature: float = 0.2,
    schema: dict | None = None,
) -> dict:
    started = time.perf_counter()
    schema_tag = f" schema={schema.get('name')}" if schema else ""
    logger.info("[LLM REQUEST]%s temp=%.2f\n[SYSTEM]\n%s\n[USER]\n%s", schema_tag, temperature, prompt_sys, prompt_user)
    data = _provider.call_json(prompt_sys, prompt_user, temperature, schema=schema)
    elapsed = time.perf_counter() - started
    logger.info("[LLM RESPONSE] elapsed=%.2fs\n%s", elapsed, json.dumps(data, ensure_ascii=False, indent=2))
    return data


def _call_llm_vision_json(
    prompt_sys: str,
    prompt_user: str,
    images_b64: list[str],
    temperature: float = 0.2,
    schema: dict | None = None,
) -> dict:
    from ..config import LLM_VISION_ENABLED

    if not LLM_VISION_ENABLED:
        logger.info("[LLM VISION → TEXT fallback] LLM_VISION_ENABLED=false, 이미지 %d장 무시", len(images_b64))
        return _call_llm_json(prompt_sys, prompt_user, temperature, schema=schema)

    started = time.perf_counter()
    schema_tag = f" schema={schema.get('name')}" if schema else ""
    logger.info(
        "[LLM VISION REQUEST]%s temp=%.2f images=%d\n[SYSTEM]\n%s\n[USER text]\n%s...",
        schema_tag, temperature, len(images_b64), prompt_sys, prompt_user[:500],
    )
    data = _provider.call_vision_json(prompt_sys, prompt_user, images_b64, temperature, schema=schema)
    elapsed = time.perf_counter() - started
    logger.info("[LLM VISION RESPONSE] elapsed=%.2fs\n%s", elapsed, json.dumps(data, ensure_ascii=False, indent=2))
    return data


def _call_llm_text(prompt_sys: str, prompt_user: str, temperature: float = 0.2) -> str:
    started = time.perf_counter()
    logger.info("[LLM REQUEST] temp=%.2f\n[SYSTEM]\n%s\n[USER]\n%s", temperature, prompt_sys, prompt_user)
    content = _provider.call_text(prompt_sys, prompt_user, temperature)
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


# ─────────────────────────────────────────────────────────────────
# 글로벌 outline pre-pass
# ─────────────────────────────────────────────────────────────────
#
# 청크별 fill은 다른 청크에서 어떤 내용이 들어갈지 모르기 때문에
# - 같은 고유명사/수치를 청크마다 다르게 표기
# - 결론 청크가 본문 청크의 핵심 메시지를 회수하지 못함
# - 표·본문 간 횡적 정합성(예: 본문에서 GPT-6 언급 → 표에서 GPT-5 등장) 깨짐
# 같은 문제가 흔하다.
#
# 이를 막기 위해 fill 시작 전에 LLM이 한 번 전체 노드 윤곽을 보고
# "이 보고서를 어떻게 끌고 갈지" 미리 결정하도록 한다. 결과는 짧은 텍스트로
# 모든 chunk analyze/fill 프롬프트에 주입된다.

_OUTLINE_SCHEMA = {
    "name": "report_outline",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "title": {"type": "string"},
            "subtitle": {"type": "string"},
            "key_message": {"type": "string"},
            "narrative_arc": {"type": "string"},
            "key_entities": {"type": "array", "items": {"type": "string"}},
            "section_plans": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "heading": {"type": "string"},
                        "intent": {"type": "string"},
                        "must_mention": {"type": "array", "items": {"type": "string"}},
                    },
                    "required": ["heading", "intent", "must_mention"],
                },
            },
        },
        "required": [
            "title", "subtitle", "key_message",
            "narrative_arc", "key_entities", "section_plans",
        ],
    },
}


def _render_nodes_for_outline(nodes_per_section: list[list[TextNode]], max_chars: int = 12000) -> str:
    """outline 프롬프트용 — 전체 노드를 압축해서 보여준다.

    너무 길면 자르되, 표/제목/본문이 골고루 살아남도록 한다.
    """
    parts: list[str] = []
    for sec_i, nodes in enumerate(nodes_per_section):
        parts.append(f"### Section {sec_i}")
        for n in nodes:
            text = (n.text or n.raw_text or "").strip()
            text = text[:80] + ("..." if len(text) > 80 else "")
            if n.type == "table_cell":
                parts.append(f"  [T{n.table_idx} R{n.row}C{n.col}] {text or '<빈>'}")
            else:
                parts.append(f"  [본문] {text}" if text else "  [본문] <빈>")
    rendered = "\n".join(parts)
    if len(rendered) > max_chars:
        # 앞 70% + 뒤 30% 보존 (도입과 결론을 모두 보기 위해)
        head = int(max_chars * 0.7)
        tail = max_chars - head - 20
        rendered = rendered[:head] + "\n... (중략) ...\n" + rendered[-tail:]
    return rendered


def render_outline_for_prompt(outline: dict | None) -> str:
    """outline dict를 chunk 프롬프트에 끼워 넣을 텍스트로 변환."""
    if not isinstance(outline, dict) or not outline:
        return ""
    title = (outline.get("title") or "").strip()
    subtitle = (outline.get("subtitle") or "").strip()
    key_msg = (outline.get("key_message") or "").strip()
    arc = (outline.get("narrative_arc") or "").strip()
    entities = outline.get("key_entities") or []
    plans = outline.get("section_plans") or []
    lines = ["## 문서 전체 스토리라인 (모든 청크 공통 — 반드시 정합성 유지)"]
    if title:
        lines.append(f"- 제목: {title}")
    if subtitle:
        lines.append(f"- 부제: {subtitle}")
    if key_msg:
        lines.append(f"- 핵심 메시지: {key_msg}")
    if arc:
        lines.append(f"- 전체 흐름: {arc}")
    if isinstance(entities, list) and entities:
        ents = ", ".join(str(e) for e in entities[:20])
        lines.append(f"- 일관 사용 고유명사/수치: {ents}")
    if isinstance(plans, list) and plans:
        lines.append("- 섹션별 계획:")
        for p in plans[:20]:
            if not isinstance(p, dict):
                continue
            heading = (p.get("heading") or "").strip()
            intent = (p.get("intent") or "").strip()
            must = p.get("must_mention") or []
            must_str = (" | 필수: " + ", ".join(str(m) for m in must[:5])) if must else ""
            lines.append(f"  · {heading}: {intent}{must_str}")
    lines.append(
        "★ 위 스토리라인을 벗어나거나 모순되는 표현은 금지. "
        "고유명사/수치는 위 목록 그대로 사용하고, 새 사실을 추가로 만들어내지 말 것."
    )
    return "\n".join(lines)


async def agent_build_report_outline(
    nodes_per_section: list[list[TextNode]],
    report_topic: str,
    report_description: str = "",
    reference_text: str = "",
    domain_type: str = "generic",
) -> dict:
    """문서 전체를 한 번 훑어 일관된 스토리라인을 결정한다.

    fill 단계 전에 1회 호출. 결과 dict는 render_outline_for_prompt()로
    렌더링되어 모든 analyze/fill 프롬프트에 컨텍스트로 주입된다.
    """
    if not nodes_per_section:
        return {}
    doc_view = _render_nodes_for_outline(nodes_per_section)

    ref_section = ""
    if reference_text:
        snippet = reference_text[:4000]
        ref_section = f"\n## 참고 텍스트 (요약 작성에 활용)\n{snippet}\n"

    desc_section = f"\n## 사용자 보고서 설명\n{report_description}\n" if report_description else ""

    prompt_sys = (
        "당신은 한국어 보고서 기획 편집자입니다. "
        "주어진 양식의 전체 구조를 파악하고, 본문을 채우기 전에 "
        "보고서 전체에서 일관되게 유지할 스토리라인·고유명사·핵심 메시지를 결정합니다. "
        "JSON만 출력하세요."
    )
    prompt_user = f"""## 보고서 주제
{report_topic}
{desc_section}{ref_section}
## 양식 전체 구조 (노드 요약)
{doc_view}

## 작업
이 양식이 모두 채워졌을 때의 보고서가 일관된 한 편의 글이 되도록 "스토리라인"을 먼저 정하세요.
- title/subtitle: 표지에 들어갈 제목·부제 (양식이 비워둔 경우 생성)
- key_message: 보고서 전체가 전달하려는 한 문장 결론
- narrative_arc: 도입→본문→결론으로 어떻게 흘러갈지 2-3문장 요약
- key_entities: 본문 전체에서 일관되게 등장시킬 고유명사·핵심 수치 (회사명·모델명·기간·지표 등). 청크별로 표기가 흔들리지 않도록 여기서 못 박는다.
- section_plans: 양식의 주요 섹션마다 어떤 내용을 다룰지 1-2줄 의도 + must_mention(반드시 포함할 키워드 1-3개)

지침:
- 도메인: {domain_type}
- 참고 텍스트가 있다면 거기서 실제 사실을 우선 사용하라. 참고 텍스트가 비어있으면 합리적으로 가정해 일관성을 유지하라.
- 양식의 표·소제목 구조에서 추론할 수 있는 섹션을 모두 section_plans에 넣어라.
- 새 사실을 만들기보다, 양식 자체와 사용자 요청에서 자연스럽게 도출되는 내용에 집중하라.

## 출력(JSON만)
{{"title":"...","subtitle":"...","key_message":"...","narrative_arc":"...","key_entities":["..."],"section_plans":[{{"heading":"...","intent":"...","must_mention":["..."]}}]}}
"""
    try:
        result = await asyncio.to_thread(
            _call_llm_json, prompt_sys, prompt_user, 0.3, _OUTLINE_SCHEMA,
        )
    except Exception as exc:
        logger.warning("[outline] LLM 호출 실패: %s — outline 없이 진행", exc)
        return {}
    if not isinstance(result, dict):
        return {}
    return result


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


# ─── 도메인별 analyze 분류 규칙 ─────────────────────────────────────

_ANALYZE_RULES_PROPOSAL = """## 분류 카테고리
- label: 라벨/제목/여백용 빈 셀.
- fill: 내용을 채울 빈 필드. 넓은 빈 셀(15mm+).
- placeholder: 교체 대상. 서식 마커(□, 가., ㅇ), 파란/빨간 가이드 텍스트.
- instruction: 독립된 "< 작성요령 >" 표. tables_to_remove로 삭제.

## 규칙
- "< 작성요령 >" 독립 표 → category:"instruction", action:"keep", skip_fill:true + tables_to_remove
- 파란/빨간 가이드 텍스트 → category:"placeholder", action:"replace", skip_fill:false
- □, 가., ㅇ 서식 마커 → category:"placeholder", action:"replace", skip_fill:false
- (여백셀) → category:"label", action:"keep", skip_fill:true

## 예시
입력: <table data-table-idx="12"><tr><td>< 작성요령 ></td></tr></table>
출력: {"id":12,"category":"instruction","action":"keep","skip_fill":true}

입력: <td data-style="...color=#0000FF...">* 과제의 목표를 기술</td>
출력: {"id":5,"category":"placeholder","action":"replace","skip_fill":false,"reason":"가이드 교체"}"""

_ANALYZE_RULES_PROJECT_STATUS = """## 분류 카테고리
- label: 섹션 제목("1. 사업 개요" 등). 수정 불가.
- placeholder: [대괄호] 플레이스홀더. 교체 대상.
- fill: 빈 본문 노드.

## 규칙
- [대괄호] 포함 노드 → category:"placeholder", action:"replace", skip_fill:false
- 섹션 제목 → category:"label", action:"keep", skip_fill:true
- (빈칸) → category:"label", action:"keep", skip_fill:true

## 예시
입력: <td data-id="0">[사업명] 사업현황</td>
출력: {"id":0,"category":"placeholder","action":"replace","skip_fill":false,"reason":"사업명 교체"}

입력: <p data-id="3">1. 사업 개요</p>
출력: {"id":3,"category":"label","action":"keep","skip_fill":true,"reason":"섹션 제목"}"""

_ANALYZE_RULES_GENERIC = """## 분류 카테고리
- label: 라벨/제목/여백. 수정 불가.
- fill: 내용을 채울 빈 필드.
- placeholder: 교체 대상 (가이드 텍스트, [대괄호] 등).

## 규칙
- 볼드 제목 → category:"label", action:"keep", skip_fill:true
- [대괄호] 텍스트 → category:"placeholder", action:"replace", skip_fill:false
- 빈 필드 → category:"fill", action:"write", skip_fill:false
- 큰 제목(폰트≥14pt) 위·아래에 붙어있는 연속된 빈 문단은 **레이아웃 여백(spacer)** → category:"label", action:"keep", skip_fill:true
- 제목 문단 자체가 비어있어도(사용자가 채울 표지 제목칸) 주변 여백 문단과 구분할 것: 여백은 절대 본문으로 채우지 말 것"""


# 제목페이지(1페이지/표지) 전용 추가 지침 — chunk_idx == 0 일 때 주입
_ANALYZE_TITLE_PAGE_HINT = """
## ★ 제목페이지(표지) 특별 지침 — 이 청크가 문서 첫 페이지일 가능성이 높음
한국 보고서의 1페이지는 대개 **표지(cover/title page)** 구조입니다. 다음을 반드시 지키세요:

1. **표지 레이아웃 요소** (일반적인 표지 구성):
   - 최상단·중앙의 큰 제목 (폰트 크게, 굵게, 가운데 정렬)
   - 부제/설명 (제목 바로 아래)
   - 하단의 작성자·기관·날짜 정보
   - 위 요소들 **사이를 벌리기 위한 다수의 빈 문단** (레이아웃용 여백)

2. **여백 문단 처리 — 가장 중요**:
   - 제목·부제·작성자 정보 **사이 또는 위아래에 연속으로 붙어있는 빈 문단**은 시각적 간격을 위한 **spacer**임.
   - 이 spacer 문단에는 **절대로 본문 내용을 채우지 말 것** → category:"label", action:"keep", skip_fill:true
   - 폰트가 8pt든 10pt든 12pt든, 용도가 spacer이면 skip_fill:true
   - 판단 근거: 빈 문단이 2개 이상 연속되거나, 큰 폰트의 제목/중앙정렬 요소 주변에 있다면 spacer로 간주.

3. **표지에서 실제로 채워야 할 곳**:
   - 제목 칸(큰 폰트의 빈 문단 단 1개) → fill
   - 부제/설명 칸 (제목 바로 아래 1~2개) → fill
   - 작성자·날짜 등 메타정보 칸 → fill 또는 placeholder

4. **의심스러우면 skip_fill:true로 처리**. 표지에 본문을 쏟아붓는 것보다 비워두는 편이 훨씬 낫습니다.
"""


async def agent_analyze_chunk(
    nodes: list[TextNode],
    chunk_idx: int = 0,
    report_description: str = "",
    table_summaries: list[TableSummary] | None = None,
    chunk_image_b64: str = "",
    prev_chunk: list[TextNode] | None = None,
    next_chunk: list[TextNode] | None = None,
    domain_type: str = "generic",
    outline_text: str = "",
) -> dict:
    llm_nodes = _filter_llm_nodes(nodes)
    doc_view = _render_nodes_html(llm_nodes)

    # 시스템 프롬프트 — 스크린샷 관련 설명은 실제 스크린샷이 있을 때만
    sys_parts = [
        "당신은 HWPX(한글) 양식 분석 전문가입니다.",
        "각 노드를 분석해 카테고리를 분류하세요.",
        "id 값은 반드시 HTML의 data-id 속성값을 그대로 사용하세요.",
    ]
    if chunk_image_b64:
        sys_parts.append("스크린샷이 제공됩니다. 시각적 레이아웃(셀 크기, 색상, 빈 영역)을 우선 참고하세요.")
    prompt_sys = " ".join(sys_parts)

    # 표 구조 정보
    table_summary_section = ""
    if table_summaries:
        chunk_table_idxs = {n.table_idx for n in nodes if n.type == "table_cell"}
        relevant = [s for s in table_summaries if s.table_idx in chunk_table_idxs]
        if relevant:
            summaries_text = "\n".join(s.summary_text() for s in relevant)
            table_summary_section = f"\n## 표 구조\n{summaries_text}\n"

    # 앞뒤 청크 컨텍스트
    context_section = ""
    if prev_chunk:
        prev_summary = _render_context_summary(prev_chunk)
        if prev_summary:
            context_section += f"\n## 이전 청크 (참고용)\n{prev_summary}\n"
    if next_chunk:
        next_summary = _render_context_summary(next_chunk)
        if next_summary:
            context_section += f"\n## 다음 청크 (참고용)\n{next_summary}\n"

    # 도메인별 분류 규칙 + few-shot
    if domain_type == "proposal":
        category_and_examples = _ANALYZE_RULES_PROPOSAL
    elif domain_type == "project_status":
        category_and_examples = _ANALYZE_RULES_PROJECT_STATUS
    else:
        category_and_examples = _ANALYZE_RULES_GENERIC

    # 첫 청크(표지 가능성)에만 제목페이지 전용 지침 주입
    title_page_section = _ANALYZE_TITLE_PAGE_HINT if chunk_idx == 0 else ""

    # 프로젝트 정보
    project_section = f"## 프로젝트 정보\n{report_description}\n\n" if report_description else ""

    # 글로벌 스토리라인 (모든 청크 공통)
    outline_section = f"\n{outline_text}\n" if outline_text else ""

    prompt_user = f"""{project_section}{outline_section}## 문서 청크
{doc_view}
{table_summary_section}{context_section}
{category_and_examples}
{title_page_section}

## 출력(JSON만, 설명 없이)
{{"tables_to_remove":[],"nodes":[{{"id":1,"category":"fill","action":"write","skip_fill":false,"reason":"..."}}]}}
"""
    # 분석 결과 검증 + 재시도: nodes 배열이 비어있으면 재시도
    _ANALYZE_MAX_RETRIES = 2
    for _analyze_attempt in range(1, _ANALYZE_MAX_RETRIES + 1):
        if chunk_image_b64:
            result = await asyncio.to_thread(
                _call_llm_vision_json, prompt_sys, prompt_user, [chunk_image_b64], 0.2
            )
        else:
            result = await asyncio.to_thread(
                _call_llm_json, prompt_sys, prompt_user, 0.2
            )
        # 검증: nodes 배열이 있고 비어있지 않은지
        result_nodes = result.get("nodes") if isinstance(result, dict) else None
        if isinstance(result_nodes, list) and len(result_nodes) > 0:
            return result
        logger.warning(
            "[analyze 검증실패] 청크 %d: attempt=%d/%d — nodes 배열 비어있음 (keys=%s). 재시도.",
            chunk_idx, _analyze_attempt, _ANALYZE_MAX_RETRIES,
            list(result.keys()) if isinstance(result, dict) else type(result).__name__,
        )
    # 모든 재시도 실패 — 마지막 결과라도 반환
    logger.error(
        "[analyze 최종실패] 청크 %d: %d회 시도 후에도 유효한 nodes 없음. 빈 분석 결과 반환.",
        chunk_idx, _ANALYZE_MAX_RETRIES,
    )
    return result


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
    domain_guide: str = "",
    outline_text: str = "",
    target_ids: list[int] | None = None,
) -> dict:
    """청크의 fillable 노드에 들어갈 텍스트를 LLM으로 작성.

    target_ids가 주어지면 그 ID들만 작성한다 (sub-batch 분할 시 사용).
    None이면 청크 내 모든 fillable 노드를 작성.
    """
    llm_nodes = _filter_llm_nodes(nodes)
    doc_view = _render_nodes_html(llm_nodes)
    analysis_view = json.dumps(analysis, ensure_ascii=False, indent=2)
    domain_section = ""
    if domain_guide:
        domain_section = f"""
## 도메인 가이드 (문서 유형별 작성 지침)

{domain_guide}

★ 도메인 가이드가 제공되었으므로 반드시 그 지침을 따르세요.
도메인 가이드의 섹션 패턴, 분량 가이드, 작성 원칙이 일반 규칙보다 우선합니다.
"""

    from ..config import IMAGE_GENERATION_ENABLED

    prompt_sys = "당신은 한국어 보고서 양식 작성 전문가입니다. JSON만 출력하세요."

    # 참고 텍스트 섹션 (있을 때만)
    ref_section = f"\n## 참고 텍스트\n{reference_text}\n" if reference_text else ""

    # 각주 규칙 (도메인 가이드에 이미 포함된 경우 생략 가능하지만 간결하게 유지)
    footnote_rule = "각주: [FN:본문텍스트|각주설명] 형식. 약어 첫 등장 시에만."

    # 이미지 마커 규칙 (이미지 생성이 켜져있을 때만)
    image_rule = ""
    if IMAGE_GENERATION_ENABLED:
        image_rule = "\n이미지가 효과적인 곳에 [IMAGE:설명] 마커 삽입. 작은 표 셀에는 넣지 마세요."

    # 출처 emit 규칙 — 참고 텍스트에 [출처#N: ...] 블록이 있을 때만 활성화
    source_rule = ""
    example_fill = '{"id":307,"new_text":"..."}'
    if reference_text and "[출처#" in reference_text:
        source_rule = (
            "\n출처: 위 '참고 텍스트'에서 실제로 근거로 삼은 내용이 있으면, "
            "해당 fill 항목에 `source_refs` 배열로 [출처#N]의 N 번호를 기재. "
            "참고하지 않은 항목에는 source_refs를 넣지 말 것(빈 배열도 생략). "
            "상상·일반상식으로 쓴 문장에는 절대 source_refs를 붙이지 말 것."
        )
        example_fill = '{"id":307,"new_text":"...","source_refs":[0,3]}'

    # 글로벌 스토리라인 (모든 청크 공통)
    outline_section = f"\n{outline_text}\n" if outline_text else ""

    # sub-batch 분할 모드: 이번 호출에서 작성할 ID만 제한
    target_section = ""
    target_count_hint = ""
    if target_ids:
        ids_str = ", ".join(str(i) for i in sorted(set(target_ids)))
        target_section = (
            f"\n## ★ 이번 호출에서 작성할 노드 ID (반드시 이 ID들만, 빠짐없이)\n[{ids_str}]\n"
        )
        target_count_hint = f" 정확히 {len(set(target_ids))}개 항목을 작성해야 합니다."

    # 빈 응답 방지 — gpt-oss류가 "This is huge → empty array" 도망가는 것을 막음
    no_skip_rule = (
        "\n절대 빈 배열을 반환하지 마세요. 분량이 많아 보여도 모든 대상 노드를 빠짐없이 작성하세요."
        " 각 셀/문단의 분량은 data-size에 맞게 짧게 써도 되지만, 빈칸으로 두거나 생략은 금지."
        f"{target_count_hint}"
    )

    prompt_user = f"""## 보고서 주제
{report_topic}
{domain_section}{outline_section}{ref_section}
## 문서(HTML)
{doc_view}

## 분석 결과
{analysis_view}
{target_section}
## 작업
action=write/replace이고 skip_fill=false인 노드만 작성. id는 data-id 값 그대로 사용.
표 셀은 data-size에 맞게 분량 조절.{no_skip_rule}
{footnote_rule}{image_rule}{source_rule}

## 출력(JSON만, 설명 없이)
{{"fills":[{example_fill}]}}
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


# ─── 소스 참고자료 청크 선택 ─────────────────────────────────────────

_BATCH_SIZE = 200  # 배치당 최대 청크 수
_MAX_SELECT_PER_BATCH = 4  # 배치당 최대 선택 수


async def agent_select_source_chunks(
    nodes: list[TextNode],
    analysis: dict,
    source_chunks: list[dict],
    report_topic: str,
    report_description: str = "",
    max_select: int = 4,
) -> list[int]:
    """템플릿 청크 내용을 기반으로 관련 소스 참고자료 청크를 선택한다.

    - 200개 이하: 한번에 선택
    - 200개 초과: 200개씩 배치로 나눠서 각 배치에서 최대 5개 선택 → 전부 합침

    Returns:
        선택된 소스 청크의 인덱스 리스트 (source_chunks 리스트 기준)
    """
    if not source_chunks or not nodes:
        return []

    fill_nodes_info = _build_fill_nodes_info(nodes, analysis)
    if not fill_nodes_info:
        return []

    fill_summary = "\n".join(fill_nodes_info[:20])
    total = len(source_chunks)

    if total <= _BATCH_SIZE:
        return await _select_batch(
            source_chunks, 0, fill_summary, report_topic, report_description, max_select,
        )

    # 배치 페이징: 200개씩 끊어서 각 배치에서 5개씩 선택, 전부 합침
    all_selected: list[int] = []
    batch_count = (total + _BATCH_SIZE - 1) // _BATCH_SIZE
    logger.info("[소스참고] 청크 %d개 → %d개 배치 (배치당 %d개)", total, batch_count, _BATCH_SIZE)

    for batch_idx in range(batch_count):
        start = batch_idx * _BATCH_SIZE
        end = min(start + _BATCH_SIZE, total)
        batch = source_chunks[start:end]
        logger.info("[소스참고] 배치 %d/%d (청크 %d~%d)", batch_idx + 1, batch_count, start, end - 1)
        selected = await _select_batch(
            batch, start, fill_summary, report_topic, report_description, _MAX_SELECT_PER_BATCH,
        )
        all_selected.extend(selected)

    logger.info("[소스참고] 전체 배치 결과: %d개 청크 선택", len(all_selected))
    return all_selected


def _build_fill_nodes_info(nodes: list[TextNode], analysis: dict) -> list[str]:
    """분석 결과에서 작성 대상 노드 정보를 추출한다."""
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
    return fill_nodes_info


async def _select_batch(
    batch_chunks: list[dict],
    offset: int,
    fill_summary: str,
    report_topic: str,
    report_description: str,
    max_select: int,
) -> list[int]:
    """배치 내 청크 summary를 보고 선택. offset은 원본 인덱스 기준."""
    summary_list = []
    for i, chunk in enumerate(batch_chunks):
        orig_idx = offset + i
        fn = chunk.get("file_name", "")
        ci = chunk.get("chunk_index", 0)
        summary = chunk.get("summary", "")
        summary_list.append(f"[{orig_idx}] 파일: {fn} | 청크#{ci} | 요약: {summary}")

    source_summary = "\n".join(summary_list)

    prompt_sys = (
        "당신은 보고서 작성을 위한 참고자료 선택 전문가입니다. "
        "사용자가 보고서 양식의 특정 부분을 작성하려 합니다. "
        "제공된 참고자료 청크 요약을 검토하고, "
        "현재 작성할 부분과 관련된 청크만 선택하세요."
    )
    prompt_user = f"""## 보고서 주제
{report_topic}

## 보고서 설명
{report_description}

## 현재 작성할 항목들
{fill_summary}

## 참고자료 청크 목록 (요약)
{source_summary}

## 작업
위 참고자료 중에서 현재 작성할 항목에 도움이 되는 청크를 최대 {max_select}개 선택하세요.
- 대괄호 안의 숫자가 인덱스입니다. 그 숫자를 그대로 반환하세요.
- 관련 없는 청크는 선택하지 마세요.
- 관련 청크가 없으면 빈 배열을 반환하세요.

## 출력(JSON)
{{"selected_indices": [0, 3, 7]}}
"""
    result = await asyncio.to_thread(_call_llm_json, prompt_sys, prompt_user, 0.2)
    indices = result.get("selected_indices") or []
    if not isinstance(indices, list):
        return []
    valid_range = set(range(offset, offset + len(batch_chunks)))
    return [idx for idx in indices[:max_select] if isinstance(idx, int) and idx in valid_range]
