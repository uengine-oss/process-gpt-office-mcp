"""Tavily 웹 검색 모듈 — generate_slides에서 자동 리서치용"""

import os
import logging
from typing import Any, Dict, List

import requests

logger = logging.getLogger("process-gpt-office-mcp")


def search_tavily(query: str, max_results: int = 5) -> List[Dict[str, Any]]:
    """Tavily API로 웹 검색. 결과 리스트 반환."""
    from .config import WEB_SEARCH_ENABLED
    if not WEB_SEARCH_ENABLED:
        logger.info("[검색] WEB_SEARCH_ENABLED=false — 웹 검색 건너뜀")
        return []

    from .config import TAVILY_API_KEY
    api_key = TAVILY_API_KEY
    if not api_key:
        logger.warning("[검색] TAVILY_API_KEY 없음 — 웹 검색 건너뜀")
        return []

    payload = {
        "api_key": api_key,
        "query": query,
        "search_depth": "basic",
        "max_results": max_results,
        "include_answer": False,
        "include_raw_content": False,
    }
    from .config import TAVILY_API_URL, HTTP_TIMEOUT_SHORT
    try:
        resp = requests.post(TAVILY_API_URL, json=payload, timeout=HTTP_TIMEOUT_SHORT)
        resp.raise_for_status()
        results = resp.json().get("results", [])
        logger.info("[검색] Tavily '%s' → %d건", query, len(results))
        return results
    except Exception as exc:
        logger.warning("[검색] Tavily 검색 실패: %s", exc)
        return []


def research_for_slides(topic: str, max_results: int = 5) -> tuple:
    """주제로 웹 검색 → (outline, sources) 튜플 반환.

    outline: List[str] — 검색 결과 제목들
    sources: List[Dict] — [{ "title": ..., "url": ..., "content": ... }]
    """
    results = search_tavily(topic, max_results=max_results)
    if not results:
        return [], []

    outline = []
    sources = []
    for r in results:
        title = r.get("title", "").strip()
        url = r.get("url", "").strip()
        content = r.get("content", "").strip()
        if title:
            outline.append(title)
        sources.append({
            "title": title,
            "url": url,
            "content": content[:500],  # 너무 길면 잘라냄
        })

    return outline, sources
