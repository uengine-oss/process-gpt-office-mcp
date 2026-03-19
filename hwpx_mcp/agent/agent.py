import asyncio
import json
import logging
import time
from typing import Any

from openai import OpenAI

from ..config import MODEL_NAME, OPENAI_API_KEY, OPENAI_TIMEOUT_SECONDS
from ..models import TextNode, TableSummary


_client = OpenAI(api_key=OPENAI_API_KEY, timeout=OPENAI_TIMEOUT_SECONDS)
logger = logging.getLogger("process-gpt-office-mcp")


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


def _call_llm_json(prompt_sys: str, prompt_user: str, temperature: float = 0.2) -> dict:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is missing")

    started = time.perf_counter()
    logger.info("[LLM REQUEST] temp=%.2f\n[SYSTEM]\n%s\n[USER]\n%s", temperature, prompt_sys, prompt_user)
    resp = _client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": prompt_sys},
            {"role": "user", "content": prompt_user},
        ],
        temperature=temperature,
        response_format={"type": "json_object"},
    )
    elapsed = time.perf_counter() - started
    content = resp.choices[0].message.content or "{}"
    data = json.loads(content)
    data["_elapsed_s"] = round(elapsed, 3)
    logger.info("[LLM RESPONSE] elapsed=%.2fs\n%s", elapsed, json.dumps(data, ensure_ascii=False, indent=2))
    return data


def _call_llm_text(prompt_sys: str, prompt_user: str, temperature: float = 0.2) -> str:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is missing")

    started = time.perf_counter()
    logger.info("[LLM REQUEST] temp=%.2f\n[SYSTEM]\n%s\n[USER]\n%s", temperature, prompt_sys, prompt_user)
    resp = _client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": prompt_sys},
            {"role": "user", "content": prompt_user},
        ],
        temperature=temperature,
    )
    elapsed = time.perf_counter() - started
    content = resp.choices[0].message.content or ""
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


def _render_table_html(nodes: list[TextNode]) -> str:
    table_nodes = [n for n in nodes if n.type == "table_cell"]
    if not table_nodes:
        return ""
    tables: dict[int, list[TextNode]] = {}
    for n in table_nodes:
        tables.setdefault(n.table_idx, []).append(n)

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
        depth = n.depth if n.depth is not None else 0
        style = f' data-style="{n.style_summary}"' if n.style_summary else ""
        parts.append(f"<p data-id=\"{n.id}\" data-depth=\"{depth}\"{style}>{text}</p>")
    return "\n".join(parts)


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
) -> dict:
    llm_nodes = _filter_llm_nodes(nodes)
    doc_view = _render_nodes_html(llm_nodes)
    prompt_sys = (
        "당신은 HWPX 양식 분석 전문가입니다. "
        "현재 상황: 사용자가 HWPX 양식을 업로드했으며, 우리는 이를 분석해 "
        "실제 값을 채우기 위한 계획(분류 결과)을 만드는 단계입니다. "
        "가능하면 개별 규칙에 과도하게 집착하지 말고, 표의 전체 구조와 행/열의 역할을 "
        "종합적으로 보고 자율적으로 판단하세요."
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
    prompt_user = f"""다음은 HWPX 문서의 일부입니다.

## 프로젝트 정보
{report_description}

## 문서(HTML)
{doc_view}
{table_summary_section}

## 예시 (few-shot)
아래와 같이 "작성 요령" 표가 나오면, 해당 표의 모든 셀은 instruction으로 분류하고
tables_to_remove에 그 표 인덱스를 추가하세요.

예시 입력(HTML):
<table data-table-idx="5">
  <tr>
    <td data-id="147" data-size="27x8mm" data-style="S:style=바탕글,size=1000,bold,color=#FF0000,align=CENTER">작성 요령</td>
    <td data-id="148" data-size="132x8mm" data-style="S:style=바탕글,size=1000,color=#FF0000,align=JUSTIFY">※ 본 작성요령은 본문 작성 후 삭제</td>
  </tr>
  <tr>
    <td data-id="149" data-size="159x20mm" data-style="S:style=바탕글,size=1000,color=#FF0000,align=JUSTIFY" colspan="2">ㅇ 수요 산업(현장) 핵심 문제 정의 ...</td>
  </tr>
</table>

예시 출력(JSON):
{{"tables_to_remove":[5],"nodes":[
  {{"id":147,"category":"instruction","action":"keep","skip_fill":true,"reason":"작성 요령 표(삭제 대상)"}},
  {{"id":148,"category":"instruction","action":"keep","skip_fill":true,"reason":"작성 요령 표(삭제 대상)"}},
  {{"id":149,"category":"instruction","action":"keep","skip_fill":true,"reason":"작성 요령 표(삭제 대상)"}}
]}}

또 다른 예시(간단): 데이터 행에서 첫 열만 반복적으로 채워지고
나머지 열이 비어 있으며 번호/항목 패턴이면 placeholder로 분류합니다.
예: <tr><td data-id="213">기술 지표 1</td><td></td><td></td></tr>
출력: {{"id":213,"category":"placeholder","action":"replace","skip_fill":false}}

## 작업
1) 각 노드를 분류하세요:
   - label: 라벨/제목 (수정 불가)
   - fixed: 이미 내용이 있는 고정값 (수정 불가)
   - fill: 빈 필드 — 내용을 채워야 함
   - placeholder: "o", "-", "⋅" 같은 기호, 또는 양식 예시 텍스트 — 실제 내용으로 교체
   - checkbox: 체크박스 ([  ] 포함) — 선택 필요
   - instruction: 작성요령/주석/안내문 (수정 불가)
   - image_placeholder: 이미지/그림/개념도/도식 자리 (텍스트 작성 금지)
   - index: 표 행 번호/인덱스 — 수정 불가
2) action은 write|replace|keep|skip|insert_image 중 하나로 지정하세요.
3) skip_fill은 true/false로 지정하세요.
4) "작성 요령", "작성요령", "삭제" 같은 문구는 instruction으로 분류하세요.
5) 작성요령 안내 문구만 담긴 표는 tables_to_remove에 추가하세요.
6) 파란/빨간색 안내문구가 있고 셀 크기가 크면 가이드 문구로 보고 placeholder + replace로 분류하세요.
   - data-style에 color 정보가 있거나, data-size가 큰 셀(예: 150mm 이상)인 경우
7) 표의 데이터 행에서 "용어", "정의", "설명" 같은 일반 라벨이 반복되면
   instruction이 아니라 placeholder( action=replace )로 분류하세요.
8) 표의 데이터 행에서 "첫 열만 채워지고 나머지 열이 비어 있음"이 여러 행 반복되고,
   첫 열 텍스트가 짧거나 번호/항목 패턴(예: 1, 2, 3, 1), 2), 1., 2.) 등)이면
   예시/임시 라벨로 보고 placeholder( action=replace, skip_fill=false )로 분류하세요.
9) fixed는 "실제 문맥 고유값"일 때만 사용하세요. 아래 유형은 기본적으로
   placeholder로 분류하세요:
   - 순번이 붙은 항목명(예: 항목 1/2/3, 지표 1/2/3, 항목 A/B/C)
   - 너무 일반적인 라벨(예: 항목, 세부항목, 내용, 기타)
   - 프로젝트/문서 맥락 없이도 성립하는 샘플/예시 텍스트
   특히 표 헤더가 "지표명/항목명/성과지표명"처럼 이름을 받는 열이면,
   데이터 행 첫 열의 텍스트가 존재하더라도 fixed로 잠그지 말고
   placeholder로 분류하세요.
10) 표 셀 중 높이(H)가 10mm 이하이고, 바로 아래에 상세 설명 셀이 이어지는 경우
   위 셀은 role="summary_header", 아래 셀은 role="detail_body"로 지정하고
   summary_header에는 detail_node(아래 셀 id), detail_body에는 header_node(위 셀 id)를 기록하세요.
   요약 셀의 권장 글자수는 max_chars에 적어주세요.
11) 셀 내용에 "그림", "이미지", "개념도", "도식", "차트", "그래프" 등이 포함되면
   image_placeholder로 분류하고 action=keep, skip_fill=true로 지정하세요.
   image_prompt, image_caption, image_ratio(16:9 또는 4:3)를 추가로 작성하세요.
12) "삭제" 또는 "없을 시 표 삭제" 같은 문구가 있으면 delete=true를 추가하세요.
   (delete=true는 최종 단계에서 해당 노드를 제거하는 신호입니다.)

## 출력(JSON)
{{"tables_to_remove":[0],"nodes":[{{"id":1,"category":"fill","action":"write","skip_fill":false,"reason":"...","role":"summary_header","detail_node":2,"max_chars":30}}]}}
"""
    return await asyncio.to_thread(
        _call_llm_json, prompt_sys, prompt_user, 0.2
    )


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
checkbox는 반드시 "[✓] 원본라벨" 형식으로 반환하세요.
표 HTML이 주어진 경우 <td data-size=\"W×H(mm)\"> 크기에 맞게 분량을 조절하세요.
한글 기준 1글자 ≈ 3~4mm (10pt 기준). 높이/너비가 작은 셀은 한 문장 이내로 작성하세요.
특히 높이(H)가 10mm 이하인 셀은 **핵심 요약 1문장**으로 제한하세요.
표 행의 C0 열에 번호가 있으면 같은 번호를 C1에 반복하지 마세요.
분석 결과의 role이 summary_header이면 max_chars 이내 1문장으로 요약하고,
role이 detail_body이면 해당 header 요약을 확장한 상세 설명(2~5문장)을 작성하세요.
summary_header와 detail_body는 같은 내용을 길이만 다르게 반복하지 말고 요약-상세로 구분하세요.
시각화 자료가 필요하면 텍스트 안에 아래 형식으로 이미지 마커를 삽입하세요.
- `[IMAGE: 이미지 프롬프트 | 캡션 | 16:9 또는 4:3]`
- 프롬프트는 한글로 구체적으로(차트/인포그래픽 중심), 캡션은 짧게 작성
- 정부 R&D 제안서 스타일(깔끔한 인포그래픽, 플랫 벡터, 과도한 장식 없음, 텍스트 최소화)을 반영
- 반드시 16:9 또는 4:3 비율을 선택해 넣으세요.
- 보고서에 분석/현황/지표/요약 내용이 있으면 최소 1개 이상 마커를 포함하세요.
- 보고서 제목/부제 같은 표지 영역은 피하고, 본문 섹션에 넣으세요.

## 출력(JSON)
{{"fills":[{{"id":1,"new_text":"..."}}]}}
"""
    return await asyncio.to_thread(
        _call_llm_json, prompt_sys, prompt_user, 0.3
    )
