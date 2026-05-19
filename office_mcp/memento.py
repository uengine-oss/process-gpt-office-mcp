"""Memento RAG 모듈 — process-gpt-memento 서비스에서 내부 지식 자료를 검색한다.

deep-research-custom의 memento.py를 process-gpt-office-mcp 내부에서 독립적으로 사용할 수 있도록 포팅.
tenant_id만 있으면 memento에서 RAG 소스를 가져올 수 있다.
"""

import asyncio
import logging
import os
import re
from typing import Any, Dict, List, Optional

import httpx

from .agent.agent import _call_llm_json

logger = logging.getLogger(__name__)

# memento 동시 요청 제한 (과도한 병렬 요청으로 인한 타임아웃 방지)
_MEMENTO_SEM = asyncio.Semaphore(5)


def _get_memento_url() -> str:
    from .config import MEMENTO_SERVICE_URL
    return MEMENTO_SERVICE_URL


def _get_drive_folder_param() -> Dict[str, str]:
    """Deprecated — drive_folder_id 기반 필터는 더 이상 사용하지 않는다.
    현재는 room_id 또는 proc_inst_id 로만 필터링한다. 빈 dict 반환.
    """
    return {}


# ---------------------------------------------------------------------------
# 내부 유틸
# ---------------------------------------------------------------------------

def _docs_to_sources(raw_docs: List[Any]) -> List[Dict[str, Any]]:
    """memento /retrieve 응답을 소스 포맷으로 변환."""
    sources: List[Dict[str, Any]] = []
    for doc in raw_docs:
        if not isinstance(doc, dict):
            continue
        content = (doc.get("page_content") or "").strip()
        if not content:
            continue
        metadata = doc.get("metadata") or {}
        file_name = metadata.get("file_name") or "내부 문서"
        folder_name = metadata.get("drive_folder_name") or ""
        # 제목에 폴더명 포함: "폴더명/파일명" 형태
        title = f"{folder_name}/{file_name}" if folder_name else file_name
        sources.append(
            {
                "title": title,
                "url": metadata.get("web_view_link") or "",
                "content": content,
                "source": "memento",
                "_chunk_index": metadata.get("chunk_index"),
                "_file_name": file_name,
                "_file_id": metadata.get("file_id") or "",
                "_drive_folder_name": folder_name,
                "_section_title": metadata.get("section_title") or "",
                # PDF: 1-based page_number (없으면 page+1로 보강됨). 다른 포맷은 None.
                "_page_number": metadata.get("page_number"),
                # PDF 청크의 bbox 유니온(JSON 문자열). 값 예: '[{"page":11,"bbox":[50,100,550,200]}]'.
                "_bboxes_json": metadata.get("bboxes_json") or "",
            }
        )
    return sources


# ---------------------------------------------------------------------------
# memento HTTP 호출
# ---------------------------------------------------------------------------

async def _broad_search(
    query: str,
    tenant_id: str,
    top_k: int = 15,
    proc_inst_id: Optional[str] = None,
    room_id: Optional[str] = None,
    file_ids: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    url = f"{_get_memento_url()}/retrieve"
    base_params: List[tuple] = [
        ("query", query),
        ("tenant_id", tenant_id),
    ]
    # 우선순위: file_ids > room_id > proc_inst_id > all_docs (폴백).
    # file_ids 가 있으면 사용자가 명시 선택한 자료에만 검색을 제한한다.
    if file_ids:
        for fid in file_ids:
            if fid:
                base_params.append(("file_ids", str(fid)))
    elif room_id:
        base_params.append(("room_id", room_id))
    elif proc_inst_id:
        base_params.append(("proc_inst_id", proc_inst_id))
    else:
        base_params.append(("all_docs", "true"))
    async with _MEMENTO_SEM:
        async with httpx.AsyncClient(timeout=60) as client:
            try:
                response = await client.get(url, params=[*base_params, ("top_k", top_k)])
                if response.status_code == 422:
                    logger.warning("memento가 top_k 파라미터를 지원하지 않음 → top_k 없이 재시도")
                    response = await client.get(url, params=base_params)
                response.raise_for_status()
                data = response.json()
                return _docs_to_sources(data.get("response") or [])
            except Exception as exc:
                logger.warning("memento 브로드 검색 실패: %s", exc)
                return []


async def _process_search(
    query: str,
    tenant_id: str,
    proc_inst_id: str,
    top_k: int = 15,
    file_ids: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """프로세스 인스턴스(proc_inst_id)에 ingest된 문서만 대상으로 단순 retrieve.

    ``file_ids`` 가 같이 주어지면 그 자료들에만 추가로 제한된다.
    """
    return await _broad_search(
        query, tenant_id, top_k=top_k, proc_inst_id=proc_inst_id, file_ids=file_ids,
    )


async def _room_search(
    query: str,
    tenant_id: str,
    room_id: str,
    top_k: int = 15,
    file_ids: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """채팅방(room_id)에 업로드된 문서만 대상으로 memento /retrieve 호출.

    ``file_ids`` 가 같이 주어지면 *사용자가 명시 선택한* 자료에 한정 검색한다
    (file_ids 가 있으면 room_id 보다 우선 — _broad_search 와 동일 정책).
    """
    if file_ids:
        return await _broad_search(query, tenant_id, top_k=top_k, file_ids=file_ids)
    url = f"{_get_memento_url()}/retrieve"
    async with _MEMENTO_SEM:
        async with httpx.AsyncClient(timeout=60) as client:
            try:
                response = await client.get(
                    url,
                    params={
                        "query": query,
                        "tenant_id": tenant_id,
                        "room_id": room_id,
                        "top_k": top_k,
                    },
                )
                if response.status_code == 422:
                    response = await client.get(
                        url,
                        params={
                            "query": query,
                            "tenant_id": tenant_id,
                            "room_id": room_id,
                        },
                    )
                response.raise_for_status()
                data = response.json()
                return _docs_to_sources(data.get("response") or [])
            except Exception as exc:
                logger.warning("memento room 검색 실패 (room_id=%s): %s", room_id, exc)
                return []


async def _list_documents(tenant_id: str) -> List[str]:
    url = f"{_get_memento_url()}/documents/list"
    params = {"tenant_id": tenant_id, **_get_drive_folder_param()}
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.get(url, params=params)
            response.raise_for_status()
            data = response.json()
        files = data.get("files") or []
        return [str(name) for name in files if name]
    except Exception as exc:
        logger.warning("documents/list 호출 실패: %s", exc)
        return []


async def _list_documents_with_folders(tenant_id: str) -> List[Dict[str, str]]:
    """문서 목록을 폴더 정보와 함께 반환한다."""
    url = f"{_get_memento_url()}/documents/list"
    params = {"tenant_id": tenant_id, **_get_drive_folder_param()}
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.get(url, params=params)
            response.raise_for_status()
            data = response.json()
        # file_details가 있으면 사용, 없으면 files로 폴백
        details = data.get("file_details")
        if details and isinstance(details, list):
            return [
                {
                    "file_name": str(d.get("file_name", "")),
                    "drive_folder_name": str(d.get("drive_folder_name", "")),
                }
                for d in details
                if d.get("file_name")
            ]
        files = data.get("files") or []
        return [{"file_name": str(name), "drive_folder_name": ""} for name in files if name]
    except Exception as exc:
        logger.warning("documents/list (with folders) 호출 실패: %s", exc)
        return []


async def _get_chunks_metadata(
    tenant_id: str, file_name: str, room_id: Optional[str] = None, proc_inst_id: Optional[str] = None,
) -> List[Dict[str, Any]]:
    url = f"{_get_memento_url()}/documents/chunks-metadata"
    params: Dict[str, Any] = {"tenant_id": tenant_id, "file_name": file_name}
    if room_id:
        params["room_id"] = room_id
    elif proc_inst_id:
        params["proc_inst_id"] = proc_inst_id
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.get(url, params=params)
            if response.status_code in (404, 422):
                logger.warning(
                    "memento가 /documents/chunks-metadata를 지원하지 않음 (status=%d)",
                    response.status_code,
                )
                return []
            response.raise_for_status()
            data = response.json()
        return data.get("chunks") or []
    except Exception as exc:
        logger.warning("chunks-metadata 호출 실패 (%s): %s", file_name, exc)
        return []


async def _retrieve_by_indices(
    tenant_id: str, file_name: str, chunk_indices: List[int],
    room_id: Optional[str] = None, proc_inst_id: Optional[str] = None,
) -> List[Dict[str, Any]]:
    if not chunk_indices:
        return []
    url = f"{_get_memento_url()}/retrieve-by-indices"
    payload: Dict[str, Any] = {
        "tenant_id": tenant_id,
        "file_name": file_name,
        "chunk_indices": chunk_indices,
    }
    if room_id:
        payload["room_id"] = room_id
    elif proc_inst_id:
        payload["proc_inst_id"] = proc_inst_id
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(url, json=payload)
            if response.status_code in (404, 422):
                logger.warning(
                    "memento가 /retrieve-by-indices를 지원하지 않음 (status=%d)",
                    response.status_code,
                )
                return []
            response.raise_for_status()
            data = response.json()
        raw_docs = data.get("response") or []
        sources: List[Dict[str, Any]] = []
        for item in raw_docs:
            if not isinstance(item, dict):
                continue
            content = (item.get("page_content") or "").strip()
            if not content:
                continue
            metadata = item.get("metadata") or {}
            src_file = metadata.get("file_name") or file_name
            folder_name = metadata.get("drive_folder_name") or ""
            title = f"{folder_name}/{src_file}" if folder_name else src_file
            sources.append(
                {
                    "title": title,
                    "url": metadata.get("web_view_link") or "",
                    "content": content,
                    "source": "memento",
                    "_chunk_index": metadata.get("chunk_index"),
                    "_file_name": src_file,
                    "_file_id": metadata.get("file_id") or "",
                    "_drive_folder_name": folder_name,
                    "_section_title": metadata.get("section_title") or "",
                    "_page_number": metadata.get("page_number"),
                    "_bboxes_json": metadata.get("bboxes_json") or "",
                }
            )
        return sources
    except Exception as exc:
        logger.warning("retrieve-by-indices 실패 (%s): %s", file_name, exc)
        return []


# ---------------------------------------------------------------------------
# LLM 기반 선택 함수들
# ---------------------------------------------------------------------------

def _select_documents_with_llm(
    query: str, file_names: List[str], max_docs: int
) -> List[str]:
    if not file_names or max_docs <= 0:
        return []

    file_list = "\n".join(f"- {name}" for name in file_names)
    system_prompt = (
        "당신은 문서 검색 보조 AI입니다. "
        "사용자의 요청과 관련 있는 문서를 여러 개 선택해 JSON 형식으로 반환합니다."
    )
    user_prompt = (
        f"사용자 요청: {query}\n\n"
        f"검색된 문서 목록:\n{file_list}\n\n"
        f"위 문서 중 사용자 요청에 필요한 문서들을 최대 {max_docs}개까지 선택하세요. "
        "반드시 목록에 있는 문서명을 그대로 JSON으로 반환하세요. "
        '예: {"selected_files": ["회사소개서.pdf", "프로젝트_수행실적.txt"]}'
    )

    def _normalize(selected: List[str]) -> List[str]:
        normalized: List[str] = []
        for name in selected:
            name = (name or "").strip()
            if not name:
                continue
            if name in file_names:
                normalized.append(name)
                continue
            for candidate in file_names:
                if name in candidate or candidate in name:
                    normalized.append(candidate)
                    break
        return list(dict.fromkeys(normalized))

    try:
        result = _call_llm_json(system_prompt, user_prompt)
        selected = result.get("selected_files") or []
        if isinstance(selected, str):
            selected = [selected]
        if not isinstance(selected, list):
            return []
        cleaned = _normalize([str(s) for s in selected])
        return cleaned[:max_docs]
    except Exception as exc:
        logger.warning("LLM 문서 선택 실패: %s", exc)
        return []


def _select_chunks_with_llm(
    outline: List[str],
    chunks_metadata: List[Dict[str, Any]],
    file_name: str,
) -> List[int]:
    if not chunks_metadata:
        return []

    chunks_summary = "\n".join(
        f"- index {c['chunk_index']}: {c.get('section_title') or '(제목 없음)'}"
        for c in chunks_metadata
        if c.get("chunk_index") is not None
    )

    system_prompt = (
        "당신은 문서 검색 보조 AI입니다. "
        "주어진 보고서 아웃라인(섹션 목록)을 작성하는 데 필요한 문서 청크를 골라야 합니다."
    )
    user_prompt = (
        f"문서명: {file_name}\n\n"
        f"보고서 아웃라인(섹션):\n"
        + "\n".join(f"- {s}" for s in outline)
        + f"\n\n청크 목록:\n{chunks_summary}\n\n"
        "위 아웃라인의 각 섹션을 작성하는 데 유용한 청크의 index 번호만 JSON 배열로 반환하세요. "
        '예: {"selected": [0, 3, 7, 12]}'
    )

    try:
        result = _call_llm_json(system_prompt, user_prompt)
        selected = result.get("selected") or []
        cleaned = [int(i) for i in selected if str(i).isdigit() or isinstance(i, int)]
        return cleaned[:30]
    except Exception as exc:
        logger.warning("LLM 청크 선택 실패: %s", exc)
        return []


def _final_review_chunks_with_llm(
    query: str,
    outline: List[str],
    sources: List[Dict[str, Any]],
    max_select: int = 10,
) -> List[Dict[str, Any]]:
    if not sources:
        return sources

    max_candidates = max_select * 3
    limited_sources = sources[:max_candidates]
    max_prompt_chars = 12000

    chunks_text_parts = []
    total_len = 0
    for pos, src in enumerate(limited_sources):
        section_title = src.get("_section_title") or ""
        content_preview = (src.get("content") or "")[:300].replace("\n", " ")
        header = f"[{pos}] {section_title}" if section_title else f"[{pos}]"
        block = f"{header}\n내용: {content_preview}"
        total_len += len(block) + 2
        if total_len > max_prompt_chars:
            break
        chunks_text_parts.append(block)
    chunks_text = "\n\n".join(chunks_text_parts)

    system_prompt = (
        "당신은 보고서 작성 보조 AI입니다. "
        "제공된 문서 청크들의 실제 내용을 검토하고, "
        "보고서 작성에 진짜 필요한 청크만 JSON 형식으로 선택합니다."
    )
    user_prompt = (
        f"사용자 요청: {query}\n\n"
        f"보고서 아웃라인:\n" + "\n".join(f"- {s}" for s in outline)
        + f"\n\n아래 {len(chunks_text_parts)}개 청크의 실제 내용을 검토하여 "
        f"보고서 작성에 실제로 필요한 청크를 최대 {max_select}개만 선택하세요.\n"
        "제목만 보고 선택한 게 아니라 실제 내용을 읽고 판단하세요.\n\n"
        f"[청크 목록]\n{chunks_text}\n\n"
        f"위 청크 번호([0], [1], ...) 중 보고서에 실제로 쓸 것을 최대 {max_select}개만 골라 JSON으로 반환하세요. "
        '예: {"selected_indices": [0, 2, 5]}'
    )

    def _extract_selected_positions(result: Any) -> List[int]:
        if not isinstance(result, dict):
            return []
        candidates = (
            result.get("selected_indices")
            or result.get("selected")
            or result.get("indices")
            or result.get("chunk_indices")
            or result.get("chunks")
            or []
        )
        if isinstance(candidates, str):
            candidates = re.findall(r"\d+", candidates)
        if not isinstance(candidates, list):
            return []
        return [
            int(p) for p in candidates
            if isinstance(p, (int, str)) and str(p).isdigit()
        ]

    try:
        result = _call_llm_json(system_prompt, user_prompt)
        selected_positions = _extract_selected_positions(result)
        selected_positions = selected_positions[:max_select]
        if not selected_positions:
            logger.warning("LLM 최종 검수 결과 빈 리스트 → 상위 %d개로 제한", max_select)
            return limited_sources[:max_select]
        filtered = [limited_sources[p] for p in selected_positions if 0 <= p < len(limited_sources)]
        logger.info("최종 검수 완료: %d → %d 청크", len(sources), len(filtered))
        return filtered
    except Exception as exc:
        logger.warning("LLM 최종 검수 실패 → 상위 %d개로 제한: %s", max_select, exc)
        return limited_sources[:max_select]


# ---------------------------------------------------------------------------
# 공개 API
# ---------------------------------------------------------------------------

async def search_memento(
    query: str,
    tenant_id: str,
    proc_inst_id: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """단순 memento 검색 — 유사 청크를 소스 포맷으로 반환."""
    if not tenant_id:
        return []
    return await _broad_search(query, tenant_id, proc_inst_id=proc_inst_id)


async def search_section_context(
    query: str,
    tenant_id: str,
    room_id: Optional[str] = None,
    proc_inst_id: Optional[str] = None,
    top_k: int = 5,
) -> List[Dict[str, Any]]:
    """섹션별 보조 검색. room_id 또는 proc_inst_id 범위로 단순 /retrieve 한 번."""
    if not tenant_id or not query.strip():
        return []
    room = (room_id or "").strip()
    proc = (proc_inst_id or "").strip()
    if room:
        sources = await _room_search(query, tenant_id, room, top_k=top_k)
    elif proc:
        sources = await _process_search(query, tenant_id, proc, top_k=top_k)
    else:
        return []
    for s in sources:
        s.pop("_chunk_index", None)
        s.pop("_file_name", None)
        s.pop("_drive_folder_name", None)
        s.pop("_section_title", None)
    return sources


async def search_memento_smart(
    query: str,
    outline: List[str],
    tenant_id: str,
    room_id: Optional[str] = None,
    proc_inst_id: Optional[str] = None,
    file_ids: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """문서-우선 스마트 Memento 검색.

    우선순위:
      1. file_ids 있음 → 사용자가 명시 선택한 자료에만 검색 (가장 우선)
      2. room_id 있음 → 채팅방 업로드 문서만 단순 /retrieve
      3. proc_inst_id 있음 → 해당 프로세스 인스턴스 문서만 단순 /retrieve
      4. 그 외 → 드라이브 기반 스마트 플로우(Step 1~6, 현재 비활성화 상태)

    드라이브 기반 스마트 플로우는 LLM 3단계 선택(문서 → 청크 index → 내용 검수)을
    수행하지만 비용이 크고 proc_inst_id 범위 필터가 보편화된 이후 기본 경로에서는
    호출하지 않는다. 필요 시 `force_smart_drive=True` 의도로 직접 호출하거나
    아래 드라이브 분기 블록의 `if False:` 게이트를 풀어 활성화한다.
    """
    if not tenant_id:
        return []

    # 명시 선택 자료 시나리오: file_ids 가 주어지면 그 자료에만 검색을 제한.
    # room_id / proc_inst_id 보다 우선 — 사용자 의도가 가장 좁은 필터이기 때문.
    if file_ids:
        logger.info(
            "search_memento_smart(file_ids) 시작 (query=%s, tenant_id=%s, file_ids=%d개)",
            query, tenant_id, len(file_ids),
        )
        file_sources = await _broad_search(
            query, tenant_id, top_k=15, file_ids=file_ids,
        )
        for s in file_sources:
            s.pop("_chunk_index", None)
            s.pop("_file_name", None)
            s.pop("_drive_folder_name", None)
            s.pop("_section_title", None)
        logger.info("search_memento_smart(file_ids) 완료: %d 청크", len(file_sources))
        return file_sources

    # 채팅방 시나리오: room에 올린 문서만 참조 → 드라이브 선택 단계 스킵
    if room_id:
        logger.info("search_memento_smart(room) 시작 (query=%s, tenant_id=%s, room_id=%s)", query, tenant_id, room_id)
        room_sources = await _room_search(query, tenant_id, room_id, top_k=15)
        for s in room_sources:
            s.pop("_chunk_index", None)
            s.pop("_file_name", None)
            s.pop("_drive_folder_name", None)
            s.pop("_section_title", None)
        logger.info("search_memento_smart(room) 완료: %d 청크", len(room_sources))
        return room_sources

    # 프로세스 시나리오: proc_inst_id에 ingest된 문서만 참조 → 드라이브 선택 단계 스킵
    if proc_inst_id:
        logger.info(
            "search_memento_smart(process) 시작 (query=%s, tenant_id=%s, proc_inst_id=%s)",
            query, tenant_id, proc_inst_id,
        )
        proc_sources = await _process_search(query, tenant_id, proc_inst_id, top_k=15)
        for s in proc_sources:
            s.pop("_chunk_index", None)
            s.pop("_file_name", None)
            s.pop("_drive_folder_name", None)
            s.pop("_section_title", None)
        logger.info("search_memento_smart(process) 완료: %d 청크", len(proc_sources))
        return proc_sources

    # 드라이브 기반 스마트 플로우는 현재 기본 경로에서 비활성화.
    # 아래 DRIVE_SMART_ENABLED 를 True 로 바꾸면 Step 1~6 LLM 3단계 선택을 다시 활성화한다.
    DRIVE_SMART_ENABLED = False
    if not DRIVE_SMART_ENABLED:
        logger.info("search_memento_smart: 드라이브 스마트 검색 비활성화 → 빈 결과 반환 (room_id/proc_inst_id 없음)")
        return []

    logger.info("search_memento_smart 시작 (query=%s, tenant_id=%s)", query, tenant_id)

    # Step 1: 문서 목록
    unique_file_names: List[str] = await _list_documents(tenant_id)
    broad_sources: List[Dict[str, Any]] = []
    if not unique_file_names:
        broad_sources = await _broad_search(query, tenant_id, top_k=15)
        if not broad_sources:
            logger.info("memento 브로드 검색 결과 없음")
            return []
        unique_file_names = list(
            dict.fromkeys(s["_file_name"] for s in broad_sources if s.get("_file_name"))
        )
        if not unique_file_names:
            return broad_sources
        logger.info("문서 후보 목록(브로드): %s", unique_file_names)
    else:
        logger.info("문서 후보 목록(전체): %s", unique_file_names)

    # Step 2: LLM으로 문서 선택
    max_docs = len(unique_file_names)
    selected_docs: List[str] = await asyncio.to_thread(
        _select_documents_with_llm, query, unique_file_names, max_docs
    )
    if not selected_docs:
        selected_docs = unique_file_names[:max_docs]
        logger.info("LLM 문서 선택 실패 → 상위 %d개 후보 사용", len(selected_docs))
    else:
        logger.info("LLM 선택 문서: %s", selected_docs)

    # Step 3~5: 청크 선택 및 조회
    max_total_chunks = 30
    precise_sources: List[Dict[str, Any]] = []

    for file_name in selected_docs:
        if len(precise_sources) >= max_total_chunks:
            break
        chunks_metadata = await _get_chunks_metadata(tenant_id, file_name)
        if not chunks_metadata:
            logger.info("chunks-metadata 없음 (%s) → 건너뜀", file_name)
            continue

        selected_indices = await asyncio.to_thread(
            _select_chunks_with_llm, outline, chunks_metadata, file_name
        )
        if not selected_indices:
            continue

        logger.info("LLM 선택 chunk_indices (%s): %s", file_name, selected_indices)

        doc_sources = await _retrieve_by_indices(tenant_id, file_name, selected_indices)
        if not doc_sources:
            logger.info("retrieve-by-indices 결과 없음 (%s) → 건너뜀", file_name)
            continue

        precise_sources.extend(doc_sources)

    if not precise_sources:
        logger.info("문서별 청크 선택 결과 없음 → 브로드 결과 사용")
        return broad_sources

    logger.info("1차 선택(title 기반): %d 청크", len(precise_sources))

    # Step 6: content 기반 최종 선택
    final_sources = await asyncio.to_thread(
        _final_review_chunks_with_llm, query, outline, precise_sources, max_total_chunks
    )

    # 내부 전용 키 제거
    for s in final_sources:
        s.pop("_chunk_index", None)
        s.pop("_file_name", None)
        s.pop("_drive_folder_name", None)
        s.pop("_section_title", None)

    logger.info("search_memento_smart 완료: 최종 %d 청크", len(final_sources))
    return final_sources


async def search_memento_by_documents(
    query: str,
    outline: List[str],
    tenant_id: str,
    file_names: List[str],
    room_id: Optional[str] = None,
    proc_inst_id: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """사용자가 선택한 특정 문서들만 대상으로 스마트 검색을 수행한다.

    search_memento_smart와 동일한 로직이지만 Step 1~2(문서 목록 조회+LLM 선택)를
    건너뛰고, 사용자가 직접 선택한 file_names로 바로 청크 검색을 수행한다.
    room_id 또는 proc_inst_id 로 범위를 제한한다 (chunks-metadata / retrieve-by-indices 둘 다).
    """
    if not tenant_id or not file_names:
        return []

    logger.info(
        "search_memento_by_documents 시작 (query=%s, tenant_id=%s, files=%s, room_id=%s, proc_inst_id=%s)",
        query, tenant_id, file_names, room_id, proc_inst_id,
    )

    max_total_chunks = 30
    precise_sources: List[Dict[str, Any]] = []

    for file_name in file_names:
        if len(precise_sources) >= max_total_chunks:
            break
        chunks_metadata = await _get_chunks_metadata(
            tenant_id, file_name, room_id=room_id, proc_inst_id=proc_inst_id,
        )
        if not chunks_metadata:
            logger.info("chunks-metadata 없음 (%s) → 건너뜀", file_name)
            continue

        selected_indices = await asyncio.to_thread(
            _select_chunks_with_llm, outline, chunks_metadata, file_name
        )
        if not selected_indices:
            # 청크 선택 실패 시 전체 청크 사용 (최대 10개)
            selected_indices = [
                c["chunk_index"] for c in chunks_metadata
                if c.get("chunk_index") is not None
            ][:10]

        logger.info("LLM 선택 chunk_indices (%s): %s", file_name, selected_indices)

        doc_sources = await _retrieve_by_indices(
            tenant_id, file_name, selected_indices,
            room_id=room_id, proc_inst_id=proc_inst_id,
        )
        if doc_sources:
            precise_sources.extend(doc_sources)

    if not precise_sources:
        logger.info("선택 문서에서 청크 조회 결과 없음")
        return []

    logger.info("1차 선택(title 기반): %d 청크", len(precise_sources))

    # content 기반 최종 선택
    final_sources = await asyncio.to_thread(
        _final_review_chunks_with_llm, query, outline, precise_sources, max_total_chunks
    )

    # 내부 전용 키 제거
    for s in final_sources:
        s.pop("_chunk_index", None)
        s.pop("_file_name", None)
        s.pop("_drive_folder_name", None)
        s.pop("_section_title", None)

    logger.info("search_memento_by_documents 완료: 최종 %d 청크", len(final_sources))
    return final_sources


def sources_to_reference_text(sources: List[Dict[str, Any]]) -> str:
    """소스 목록을 reference_text 문자열로 변환한다.

    generate_hwpx는 structured sources가 아닌 reference_text(str)를 사용하므로
    memento 소스를 텍스트로 변환해야 한다.
    """
    if not sources:
        return ""
    parts: List[str] = []
    for src in sources:
        title = src.get("title") or "Untitled"
        content = (src.get("content") or "").strip()
        if not content:
            continue
        parts.append(f"[{title}]\n{content}")
    return "\n\n---\n\n".join(parts)


async def prefetch_pdf_highlight_url(
    tenant_id: str,
    file_id: str,
    page: int,
    bbox: List[float],
    dpi: int = 120,
) -> Optional[str]:
    """memento `/preview/pdf-highlight`를 호출해 하이라이트 PNG의 public URL을 받아온다.

    실패 시 None. 문서 생성 시점에 미리 렌더링 + Supabase 업로드까지 끝내두고
    HTML의 data-sources에 URL을 직접 박아 넣기 위한 헬퍼.
    """
    if not tenant_id or not file_id or page is None or not bbox or len(bbox) != 4:
        return None
    url = f"{_get_memento_url()}/preview/pdf-highlight"
    params: Dict[str, Any] = {
        "tenant_id": tenant_id,
        "file_id": file_id,
        "page": int(page),
        "bbox": ",".join(f"{float(v):.2f}" for v in bbox),
        "dpi": dpi,
    }
    try:
        async with httpx.AsyncClient(timeout=90) as client:
            resp = await client.get(url, params=params)
            resp.raise_for_status()
            data = resp.json()
            result = (data.get("url") or "").strip()
            return result or None
    except Exception as exc:
        logger.warning("prefetch_pdf_highlight 실패 (file=%s p=%s): %s", file_id, page, exc)
        return None


async def prefetch_preview_urls_for_sources(
    tenant_id: str, sources_meta_list: List[Dict[str, Any]], dpi: int = 120,
) -> None:
    """여러 소스 메타에 대해 preview_url을 병렬 fetch해서 각 메타 dict에 in-place 주입.

    source meta 형태(process-gpt-office-mcp 내부):
        {file_id: str, bboxes_json: str, ...}
    성공 시 `preview_url` 키가 dict에 추가된다. 실패는 조용히 스킵.
    같은 (file_id, page, bbox) 조합은 재사용(중복 호출 방지).
    """
    if not tenant_id or not sources_meta_list:
        return
    import json as _json_mod

    # 중복 제거: 같은 (file_id, page, bbox) → 한 번만 호출
    dedup_tasks: Dict[tuple, asyncio.Task] = {}
    # 각 src와 그에 필요한 key 매핑
    src_keys: List[tuple] = []  # list of (src_dict, key_tuple_or_None)

    for src in sources_meta_list:
        file_id = src.get("file_id") or ""
        bboxes_raw = src.get("bboxes_json") or ""
        if not file_id or not bboxes_raw:
            src_keys.append((src, None))
            continue
        try:
            bboxes = _json_mod.loads(bboxes_raw)
        except Exception:
            src_keys.append((src, None))
            continue
        if not isinstance(bboxes, list) or not bboxes:
            src_keys.append((src, None))
            continue
        first = bboxes[0]
        bbox = first.get("bbox")
        page = first.get("page")
        if not isinstance(bbox, list) or len(bbox) != 4 or page is None:
            src_keys.append((src, None))
            continue
        key = (file_id, int(page), tuple(round(float(v), 2) for v in bbox))
        src_keys.append((src, key))
        if key not in dedup_tasks:
            dedup_tasks[key] = asyncio.create_task(
                prefetch_pdf_highlight_url(tenant_id, file_id, page, bbox, dpi=dpi)
            )

    if not dedup_tasks:
        return

    # 모든 태스크 완료 대기
    await asyncio.gather(*dedup_tasks.values(), return_exceptions=True)

    # 각 src에 결과 주입
    for src, key in src_keys:
        if not key:
            continue
        task = dedup_tasks.get(key)
        if task is None or not task.done():
            continue
        try:
            url = task.result()
        except Exception:
            url = None
        if url:
            src["preview_url"] = url


def sources_to_numbered_reference_text(
    sources: List[Dict[str, Any]], start_idx: int = 0
) -> tuple[str, List[Dict[str, Any]]]:
    """번호를 부여한 reference_text + 출처 메타 리스트를 반환.

    LLM이 fill 결과에 `source_refs: [N, ...]`로 인용할 수 있도록
    `[출처#N: title]\\n원문` 포맷으로 변환한다.
    메타 리스트는 N 순서로 정렬되어 반환되며, 각 항목은
    {file_name, title, snippet, section_title, chunk_index, url} 형태.
    """
    if not sources:
        return "", []
    parts: List[str] = []
    meta: List[Dict[str, Any]] = []
    n = start_idx
    for src in sources:
        content = (src.get("content") or src.get("original_text") or "").strip()
        if not content:
            continue
        title = src.get("title") or src.get("file_name") or "Untitled"
        file_name = src.get("file_name") or src.get("_file_name") or title
        snippet = content.replace("\n", " ")
        if len(snippet) > 240:
            snippet = snippet[:240] + "…"
        meta.append({
            "file_name": file_name,
            "title": title,
            "snippet": snippet,
            "section_title": src.get("section_title") or src.get("_section_title") or "",
            "chunk_index": src.get("chunk_index") or src.get("_chunk_index"),
            "page": src.get("page_number") or src.get("_page_number"),
            "url": src.get("url") or "",
            # PDF 하이라이트 프리뷰용 원본 식별자.
            "file_id": src.get("file_id") or src.get("_file_id") or "",
            "bboxes_json": src.get("bboxes_json") or src.get("_bboxes_json") or "",
        })
        parts.append(f"[출처#{n}: {title}]\n{content}")
        n += 1
    return "\n\n---\n\n".join(parts), meta


async def search_memento_images(
    query: str,
    tenant_id: str,
    top_k: int = 5,
) -> List[Dict[str, Any]]:
    """memento /retrieve-images 로 캡션 기반 이미지 검색.

    Returns:
        [{"image_id", "image_url", "caption", "file_name", "metadata"}, ...]
    """
    if not tenant_id or not query:
        return []
    url = f"{_get_memento_url()}/retrieve-images"
    async with _MEMENTO_SEM:
        async with httpx.AsyncClient(timeout=60) as client:
            try:
                response = await client.get(
                    url,
                    params={
                        "query": query,
                        "tenant_id": tenant_id,
                        "top_k": top_k,
                        **_get_drive_folder_param(),
                    },
                )
                if response.status_code in (404, 422):
                    logger.warning(
                        "memento가 /retrieve-images를 지원하지 않음 (status=%d)",
                        response.status_code,
                    )
                    return []
                response.raise_for_status()
                data = response.json()
                return data.get("images") or []
            except Exception as exc:
                logger.warning("memento 이미지 검색 실패: %s", exc)
                return []


async def search_memento_images_multi_query(
    queries: List[str],
    tenant_id: str,
    top_k: int = 5,
) -> List[Dict[str, Any]]:
    """여러 쿼리로 memento 이미지를 검색하고 중복 제거해 반환한다."""
    if not tenant_id or not queries:
        return []

    logger.info("[이미지RAG] %d개 쿼리로 memento 이미지 검색 시작", len(queries))

    tasks = [search_memento_images(q, tenant_id, top_k=top_k) for q in queries]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    seen_ids: set = set()
    unique_images: List[Dict[str, Any]] = []
    for i, result in enumerate(results):
        if isinstance(result, Exception):
            logger.warning("[이미지RAG] 쿼리 %d 실패: %s", i, result)
            continue
        for img in result:
            image_id = img.get("image_id") or img.get("image_url") or ""
            if not image_id or image_id in seen_ids:
                continue
            seen_ids.add(image_id)
            unique_images.append(img)

    logger.info("[이미지RAG] 검색 완료: 총 %d개 (중복제거 후)", len(unique_images))
    return unique_images


async def search_memento_multi_query(
    queries: List[str],
    tenant_id: str,
    top_k: int = 5,
    file_names: List[str] | None = None,
    proc_inst_id: Optional[str] = None,
    room_id: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """여러 쿼리로 memento를 검색하고 중복 제거해 반환한다.

    각 쿼리별 top_k개를 가져온 뒤 content 기준으로 중복을 제거한다.
    file_names가 주어지면 해당 문서에서 나온 결과만 반환한다.
    room_id 또는 proc_inst_id 로 검색 범위를 제한한다.
    """
    if not tenant_id or not queries:
        return []

    filter_desc = ""
    if file_names:
        filter_desc += f", 문서필터={len(file_names)}개"
    if room_id:
        filter_desc += f", room_id={room_id}"
    elif proc_inst_id:
        filter_desc += f", proc_inst_id={proc_inst_id}"
    logger.info("[청크RAG] %d개 쿼리로 memento 검색 시작 (top_k=%d%s)", len(queries), top_k, filter_desc)

    # 모든 쿼리를 병렬로 검색
    tasks = [
        _broad_search(q, tenant_id, top_k=top_k, proc_inst_id=proc_inst_id, room_id=room_id)
        for q in queries
    ]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    # 파일명 필터 준비 (부분 매칭: 선택 문서명이 결과 파일명에 포함되면 통과)
    def _matches_filter(src: Dict[str, Any]) -> bool:
        if not file_names:
            return True
        src_file = src.get("_file_name") or ""
        if not src_file:
            return False
        for selected in file_names:
            if selected in src_file or src_file in selected:
                return True
        return False

    # 결과 합산 + 중복 제거 (content 해시 기준)
    seen_contents: set = set()
    unique_sources: List[Dict[str, Any]] = []
    filtered_count = 0
    for i, result in enumerate(results):
        if isinstance(result, Exception):
            logger.warning("[청크RAG] 쿼리 %d 실패: %s", i, result)
            continue
        for src in result:
            content = (src.get("content") or "").strip()
            if not content:
                continue
            if not _matches_filter(src):
                filtered_count += 1
                continue
            content_key = content[:200]  # 앞 200자로 중복 판별
            if content_key in seen_contents:
                continue
            seen_contents.add(content_key)
            unique_sources.append(src)

    if filtered_count:
        logger.info("[청크RAG] 검색 완료: 총 %d개 (중복제거 후, 문서필터로 %d개 제외)", len(unique_sources), filtered_count)
    else:
        logger.info("[청크RAG] 검색 완료: 총 %d개 (중복제거 후)", len(unique_sources))
    return unique_sources
