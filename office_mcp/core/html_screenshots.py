"""HTML을 PNG 스크린샷으로 변환한다 (playwright 사용).

청크별로 해당 노드의 bounding box를 계산하여 crop한 스크린샷을 생성한다.
"""

import base64
import io
import logging
import time
from dataclasses import dataclass, field

logger = logging.getLogger("process-gpt-office-mcp")


@dataclass
class FullPageCapture:
    """전체 HTML + 브라우저 세션 + 풀페이지 스크린샷 캐시."""
    full_html: str = ""
    _browser: object = None
    _playwright: object = None
    _full_screenshot: bytes = b""
    _page: object = None  # 페이지 재사용 (bbox 계산용)


async def init_capture(full_html: str) -> FullPageCapture:
    """hwpx_to_html 결과 HTML을 저장하고 Playwright 브라우저를 준비한다."""
    if not full_html:
        return FullPageCapture()

    from playwright.async_api import async_playwright

    pw = await async_playwright().start()
    browser = await pw.chromium.launch(headless=True)
    logger.info("[캡처] Playwright 브라우저 시작")

    cap = FullPageCapture(
        full_html=full_html,
        _browser=browser,
        _playwright=pw,
    )

    # 전체 페이지를 한 번만 렌더링하고 스크린샷 캐시
    try:
        page = await browser.new_page(viewport={"width": 1024, "height": 800})
        wrapped_html = (
            '<!doctype html><html><head><meta charset="utf-8">'
            "<style>"
            "body{margin:0;padding:0;font-family:sans-serif;line-height:1.3;background:#fff}"
            ".page{position:relative;box-sizing:border-box;"
            "height:auto!important;min-height:auto!important;overflow:visible!important;"
            "margin:0!important;border:none!important;box-shadow:none!important}"
            "table{border-collapse:collapse}td{vertical-align:top}p{margin:0}"
            "</style></head><body>"
            f"{full_html}"
            "</body></html>"
        )
        await page.set_content(wrapped_html, wait_until="networkidle")
        cap._full_screenshot = await page.screenshot(full_page=True)
        cap._page = page
        logger.info("[캡처] 풀페이지 스크린샷 캐시 완료 (%.1f KB)", len(cap._full_screenshot) / 1024)
    except Exception as exc:
        logger.warning("[캡처] 풀페이지 스크린샷 실패: %s", exc)

    return cap


async def close_capture(capture: FullPageCapture):
    """Playwright 브라우저를 닫는다."""
    if capture._page:
        try:
            await capture._page.close()
        except Exception:
            pass
    if capture._browser:
        await capture._browser.close()
    if capture._playwright:
        await capture._playwright.stop()
    logger.info("[캡처] Playwright 브라우저 종료")


_PADDING = 10  # crop 여백 (px)


async def screenshot_chunk(
    capture: FullPageCapture,
    node_ids: list[int],
    width: int = 1024,
) -> str:
    """매칭 노드의 bounding box를 계산하여 풀페이지 스크린샷에서 crop.

    테이블 셀이 포함되면 해당 테이블 전체의 bbox를 사용한다.

    Returns:
        base64 인코딩된 PNG 문자열. 실패 시 빈 문자열.
    """
    if not capture._full_screenshot or not capture._page or not node_ids:
        return ""

    started = time.perf_counter()
    try:
        # JS로 매칭 노드들의 통합 bounding box 계산
        bbox = await capture._page.evaluate(f"""(nodeIds) => {{
            const ids = new Set(nodeIds);
            let minX = Infinity, minY = Infinity, maxX = 0, maxY = 0;
            let found = 0;

            // 매칭되는 data-id 요소 수집
            const matchedEls = [];
            document.querySelectorAll('[data-id]').forEach(el => {{
                const id = parseInt(el.getAttribute('data-id'), 10);
                if (ids.has(id)) matchedEls.push(el);
            }});

            // 테이블 셀이면 테이블 전체의 bbox 사용
            const elements = new Set();
            matchedEls.forEach(el => {{
                let tbl = el.closest('table');
                if (tbl) {{
                    elements.add(tbl);
                }} else {{
                    elements.add(el);
                }}
            }});

            elements.forEach(el => {{
                const rect = el.getBoundingClientRect();
                if (rect.width === 0 && rect.height === 0) return;
                found++;
                minX = Math.min(minX, rect.left + window.scrollX);
                minY = Math.min(minY, rect.top + window.scrollY);
                maxX = Math.max(maxX, rect.right + window.scrollX);
                maxY = Math.max(maxY, rect.bottom + window.scrollY);
            }});

            if (found === 0) return null;
            return {{ x: minX, y: minY, width: maxX - minX, height: maxY - minY }};
        }}""", node_ids)

        if not bbox:
            logger.warning("[청크스크린샷] 매칭 노드 bbox 없음 (node_ids=%d개)", len(node_ids))
            return ""

        # PIL로 crop
        from PIL import Image
        full_img = Image.open(io.BytesIO(capture._full_screenshot))
        img_w, img_h = full_img.size

        x = max(0, int(bbox["x"]) - _PADDING)
        y = max(0, int(bbox["y"]) - _PADDING)
        right = min(img_w, int(bbox["x"] + bbox["width"]) + _PADDING)
        bottom = min(img_h, int(bbox["y"] + bbox["height"]) + _PADDING)

        if right <= x or bottom <= y:
            logger.warning("[청크스크린샷] 유효하지 않은 crop 영역: (%d,%d)-(%d,%d)", x, y, right, bottom)
            return ""

        cropped = full_img.crop((x, y, right, bottom))
        buf = io.BytesIO()
        cropped.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")

        elapsed = time.perf_counter() - started
        logger.info(
            "[청크스크린샷] node %d개, crop (%d,%d %dx%d) → %.1f KB (%.1fs)",
            len(node_ids), x, y, right - x, bottom - y,
            len(buf.getvalue()) / 1024, elapsed,
        )
        return b64
    except Exception as exc:
        logger.warning("[청크스크린샷] 실패: %s", exc)
        return ""


async def html_to_page_images(
    full_html: str,
    width: int = 1024,
    page_height: int = 1448,
) -> list[str]:
    """전체 HTML을 렌더링한 뒤 page_height 단위로 잘라 base64 PNG 리스트로 반환한다."""
    if not full_html:
        return []

    from playwright.async_api import async_playwright

    started = time.perf_counter()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": width, "height": page_height})

        wrapped_html = (
            '<!doctype html><html><head><meta charset="utf-8">'
            "<style>"
            "body{margin:0;padding:0;font-family:sans-serif;line-height:1.3;background:#fff}"
            ".page{position:relative;box-sizing:border-box;"
            "height:auto!important;min-height:auto!important;overflow:visible!important;"
            "margin:0!important;border:none!important;box-shadow:none!important}"
            "table{border-collapse:collapse}td{vertical-align:top}p{margin:0}"
            "</style></head><body>"
            f"{full_html}"
            "</body></html>"
        )
        await page.set_content(wrapped_html, wait_until="networkidle")
        full_screenshot = await page.screenshot(full_page=True)
        await browser.close()

    from PIL import Image
    full_img = Image.open(io.BytesIO(full_screenshot))
    img_width, img_height = full_img.size

    images: list[str] = []
    y = 0
    while y < img_height:
        bottom = min(y + page_height, img_height)
        cropped = full_img.crop((0, y, img_width, bottom))
        buf = io.BytesIO()
        cropped.save(buf, format="PNG")
        images.append(base64.b64encode(buf.getvalue()).decode("ascii"))
        y = bottom

    elapsed = time.perf_counter() - started
    logger.info("[스크린샷] %d 페이지 캡처 완료 (%.1fs)", len(images), elapsed)
    return images
