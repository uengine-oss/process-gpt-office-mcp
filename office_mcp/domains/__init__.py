"""Domain detection and guide loading for form-specific prompts.

Each .md file in this package directory defines writing guidelines for a
specific document type (e.g. proposal, project_status).  ``detect_domain``
asks the LLM to inspect the parsed document structure and choose the best
matching domain guide.
"""

from __future__ import annotations

import asyncio
import json
import logging
import time
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..models import TextNode

logger = logging.getLogger(__name__)

_DOMAINS_DIR = Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _load_domain_catalog() -> dict[str, dict]:
    """Scan domains/ for .md files and build a catalog.

    Returns
    -------
    { "proposal": {"name": "proposal", "summary": "첫 3줄 요약", "path": Path}, ... }
    """
    catalog: dict[str, dict] = {}
    for md in sorted(_DOMAINS_DIR.glob("*.md")):
        name = md.stem                       # e.g. "proposal"
        lines = md.read_text(encoding="utf-8").splitlines()
        # 제목(# 줄) + 첫 비어있지 않은 본문 2줄을 요약으로 사용
        summary_lines: list[str] = []
        for line in lines:
            stripped = line.strip()
            if not stripped or stripped == "---":
                continue
            summary_lines.append(stripped)
            if len(summary_lines) >= 4:
                break
        summary = "\n".join(summary_lines)
        catalog[name] = {"name": name, "summary": summary, "path": md}
    return catalog


def _build_document_preview(nodes: list[TextNode], max_nodes: int = 80) -> str:
    """Build a concise text preview of the document for LLM analysis."""
    lines: list[str] = []
    for n in nodes[:max_nodes]:
        text = (n.text or n.raw_text or "").strip()
        ntype = getattr(n, "type", "")
        if not text:
            continue
        # 간결하게: [타입] 텍스트
        prefix = "표" if ntype == "table_cell" else "본문"
        lines.append(f"[{prefix}] {text[:120]}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

async def detect_domain(nodes: list[TextNode]) -> tuple[str, str]:
    """Detect document type using LLM analysis and return domain guide text.

    The LLM sees a preview of the document structure and a catalog of
    available domain guides, then chooses the best match (or "generic").

    Parameters
    ----------
    nodes : list[TextNode]
        All TextNode objects parsed from the HWPX section.

    Returns
    -------
    (domain_type, guide_text)
        domain_type : domain name (e.g. "proposal") or "generic"
        guide_text  : full contents of the matching ``domains/<type>.md``,
                      or empty string for "generic".
    """
    catalog = _load_domain_catalog()
    if not catalog:
        logger.info("[도메인] 도메인 가이드 파일 없음 → generic")
        return ("generic", "")

    doc_preview = _build_document_preview(nodes)
    if not doc_preview.strip():
        logger.info("[도메인] 문서 미리보기 비어있음 → generic")
        return ("generic", "")

    # 도메인 목록 텍스트 구성
    domain_list_parts: list[str] = []
    for name, info in catalog.items():
        domain_list_parts.append(f"### {name}\n{info['summary']}")
    domain_list_text = "\n\n".join(domain_list_parts)

    domain_names = list(catalog.keys())

    prompt_sys = (
        "당신은 문서 양식 분류 전문가입니다. "
        "HWPX 문서의 구조를 분석하여 가장 적합한 도메인 가이드를 선택하세요."
    )
    prompt_user = f"""아래는 HWPX 문서에서 추출한 텍스트 미리보기입니다.

## 문서 미리보기
{doc_preview}

## 사용 가능한 도메인 가이드
{domain_list_text}

## 작업
위 문서가 어떤 유형의 양식인지 분석하고, 가장 적합한 도메인 가이드를 선택하세요.
적합한 가이드가 없으면 "generic"을 선택하세요.

## 출력(JSON)
{{"domain": "<도메인명 또는 generic>", "reason": "<선택 이유 1줄>"}}

선택 가능한 값: {json.dumps(domain_names + ["generic"], ensure_ascii=False)}
"""

    from ..agent.agent import _call_llm_json

    try:
        started = time.perf_counter()
        result = await asyncio.to_thread(
            _call_llm_json, prompt_sys, prompt_user, 0.1,
        )
        elapsed = time.perf_counter() - started

        domain_type = result.get("domain", "generic")
        reason = result.get("reason", "")

        # 유효성 검증: 카탈로그에 없는 이름이면 generic
        if domain_type not in catalog and domain_type != "generic":
            logger.warning(
                "[도메인] LLM이 알 수 없는 도메인 반환: %s → generic", domain_type
            )
            domain_type = "generic"

        logger.info(
            "[도메인] LLM 감지: %s (%.1fs) — %s", domain_type, elapsed, reason
        )
    except Exception as exc:
        logger.warning("[도메인] LLM 감지 실패 (generic 폴백): %s", exc)
        return ("generic", "")

    if domain_type == "generic":
        return ("generic", "")

    guide_path = catalog[domain_type]["path"]
    guide_text = guide_path.read_text(encoding="utf-8")
    logger.info("[도메인] 가이드 로드: %s (%.1f KB)", domain_type, len(guide_text) / 1024)
    return (domain_type, guide_text)
