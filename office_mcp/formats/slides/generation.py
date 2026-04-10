"""Slide (presentation) content and image generation — adapted from deep-research-custom."""
import logging
import os
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

from supabase import create_client

from ...agent.agent import _call_llm_text
from ...images import generate_image_gemini

logger = logging.getLogger("office-mcp-slides")

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
        logger.warning("slide image 업로드 실패: %s", e)
    from ...config import SUPABASE_URL as _sb_url
    base_url = _sb_url.rstrip("/")
    if base_url:
        return f"{base_url}/storage/v1/object/public/{STORAGE_BUCKET}/{quote(storage_path, safe='/-_.')}"
    return None


# ---------------------------------------------------------------------------
# Slide markdown builders
# ---------------------------------------------------------------------------

def build_slide_markdown(
    report_markdown: str, slide_count: Optional[int] = None, style: Optional[str] = None
) -> str:
    count_line = f"슬라이드 개수: {slide_count} (정확히)" if slide_count else "슬라이드 개수: 적절히"
    style_line = f"스타일 가이드: {style}" if style else ""
    prompt = (
        "당신은 프레젠테이션 디자이너입니다. 아래 보고서 마크다운을 PPT 슬라이드용 마크다운으로 재구성하세요.\n"
        "- 형식: # 덱 제목 1개, 이후 각 슬라이드마다 ## 슬라이드 제목 + 핵심 불릿 2~4개\n"
        "- 불릿은 행동/팩트 중심으로 간결히\n- 디자인 지시문은 넣지 말 것\n"
        f"- {count_line}\n- {style_line}\n\n보고서 마크다운:\n{report_markdown}\n"
    )
    return _call_llm_text("You are a presentation writer. Return markdown only.", prompt) or ""


def build_slide_markdown_from_research(
    research_goal: str,
    outline: List[str],
    sources: List[Dict[str, Any]],
    deck_title: str = "",
    slide_count: Optional[int] = None,
    style: Optional[str] = None,
) -> str:
    count_line = f"슬라이드 개수: {slide_count} (정확히)" if slide_count else "슬라이드 개수: 적절히"
    style_line = f"스타일 가이드: {style}" if style else ""
    outline_text = "\n".join(f"- {item}" for item in (outline or [])) or "N/A"
    sources_text = "\n".join(
        f"- {item.get('title', 'Untitled')} | {item.get('url', '')}"
        for item in (sources or [])[:5]
    )
    prompt = (
        "당신은 프레젠테이션 디자이너입니다. 아래 리서치 목표와 개요/소스를 바탕으로 "
        "PPT 슬라이드용 마크다운을 작성하세요.\n"
        "- 형식: # 덱 제목 1개, 이후 각 슬라이드마다 ## 슬라이드 제목 + 핵심 불릿 2~4개\n"
        "- 불릿은 행동/팩트 중심으로 간결히\n- 디자인 지시문은 넣지 말 것\n"
        f"- {count_line}\n- {style_line}\n\n"
        f"리서치 목표:\n{research_goal}\n\n개요:\n{outline_text}\n\n참고 소스 요약:\n{sources_text or 'N/A'}\n"
    )
    if deck_title:
        prompt += f"\n덱 제목 후보: {deck_title}\n"
    return _call_llm_text("You are a presentation writer. Return markdown only.", prompt) or ""


def parse_slides(markdown: str, max_slides: int = 10) -> List[Dict[str, Any]]:
    slides: List[Dict[str, Any]] = []
    current: Dict[str, Any] = {}
    for line in markdown.splitlines():
        if line.startswith("## "):
            if current.get("title"):
                slides.append(current)
            current = {"title": line.replace("##", "").strip(), "bullets": []}
            if len(slides) >= max_slides:
                break
        elif line.strip().startswith("- ") and current.get("title"):
            current.setdefault("bullets", []).append(line.strip("- ").strip())
    if current.get("title") and len(slides) < max_slides:
        slides.append(current)
    return slides[:max_slides]


def build_style_guide(
    slides: List[Dict[str, Any]], deck_title: str = "", user_style: str = ""
) -> str:
    outline = "\n".join(
        f"- {s.get('title','')}: {', '.join(s.get('bullets', [])[:2])}" for s in slides
    )
    prompt = (
        "당신은 프레젠테이션 아트디렉터입니다. 슬라이드 전체에 일관되게 적용할 스타일 가이드를 한국어로 작성하세요.\n"
        "- 3~5색 팔레트, 폰트 톤, 배경/질감, 조명/톤, 모티프/프레이밍 규칙 포함\n"
        "- JSON이나 코드블럭 없이 문단으로 짧게 요약\n"
        "- 사용자 스타일/색감을 최우선으로 반드시 그대로 반영\n"
        f"- 최대 슬라이드 개요:\n{outline}\n- 덱 제목: {deck_title or 'N/A'}\n"
        f"- 사용자 스타일/색감(반드시 준수): {user_style or 'N/A'}\n"
    )
    return _call_llm_text("You are a concise art director. Respond in Korean.", prompt) or ""


def _build_slide_image_prompt(slide: Dict[str, Any], style_guide: str, deck_outline: str) -> str:
    bullets = ", ".join(slide.get("bullets", [])[:3])
    return (
        "당신은 프레젠테이션용 이미지 디자이너입니다. 한 장의 이미지로 슬라이드 메시지가 전달되도록 작성하세요.\n"
        f"- 슬라이드 제목: {slide.get('title','')}\n- 핵심 내용: {bullets}\n"
        f"- 전체 슬라이드 개요: {deck_outline}\n- 스타일 가이드: {style_guide}\n"
        "- 사용자 스타일/색감을 최우선으로 반드시 준수.\n"
        "- 동일한 조명/팔레트/질감을 유지하고 과도한 텍스트는 넣지 말 것.\n"
    )


def generate_slide_images(
    slide_md: str,
    report_id: str,
    style_guide: str,
    deck_title: str = "",
    slide_count: Optional[int] = None,
) -> Tuple[str, List[str]]:
    max_slides = slide_count if slide_count else 10
    slides = parse_slides(slide_md, max_slides=max_slides)
    if not slides:
        return slide_md, []

    deck_outline = "; ".join(f"{idx+1}. {s.get('title','')}" for idx, s in enumerate(slides))
    url_map: Dict[int, Optional[str]] = {i: None for i in range(len(slides))}

    def _render_and_upload(idx: int, slide: Dict[str, Any]) -> Tuple[int, Optional[str]]:
        prompt = _build_slide_image_prompt(slide, style_guide, deck_outline)
        filename = f"slide-{idx+1}.png"
        tmp_dir = Path(tempfile.mkdtemp())
        local_path = tmp_dir / filename
        url = None
        if generate_image_gemini(prompt, local_path):
            url = _upload_image_to_storage(local_path, report_id, filename)
        return idx, url

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(_render_and_upload, idx, slide): idx for idx, slide in enumerate(slides)}
        for future in as_completed(futures):
            idx, url = future.result()
            url_map[idx] = url

    md_blocks: List[str] = []
    image_urls: List[str] = []
    for idx, slide in enumerate(slides):
        url = url_map.get(idx)
        lines = []
        if idx == 0:
            title_line = deck_title.strip() if deck_title else slide.get("title", "").strip()
            if title_line:
                lines.append(f"# {title_line}")
        slide_title = slide.get("title", "").strip()
        if slide_title:
            lines.append(f"## {slide_title}")
        if url:
            lines.append(f"![{slide_title or 'slide'}]({url})")
            image_urls.append(url)
        for b in slide.get("bullets", []):
            lines.append(f"- {b}")
        md_blocks.append("\n".join(lines))

    return "\n\n---\n\n".join(md_blocks), image_urls
