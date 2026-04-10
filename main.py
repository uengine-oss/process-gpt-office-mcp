import base64
import uvicorn
from starlette.middleware.cors import CORSMiddleware
from starlette.middleware import Middleware
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from office_mcp.mcp_server import mcp


async def edit_slide_image(request: Request) -> JSONResponse:
    """슬라이드 이미지 편집 REST 엔드포인트.

    Body JSON:
        image_url (str): 원본 이미지 URL
        instruction (str): 수정 지시
    Returns:
        { "image_base64": "<base64>", "mime_type": "image/png" }
    """
    import httpx
    from office_mcp.images import edit_image_gemini_bytes

    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"error": "Invalid JSON"}, status_code=400)

    image_url = (body.get("image_url") or "").strip()
    instruction = (body.get("instruction") or "").strip()
    if not image_url or not instruction:
        return JSONResponse({"error": "image_url and instruction are required"}, status_code=400)

    selection = body.get("selection")  # { x1, y1, x2, y2, width, height } or None
    annotated_b64 = body.get("annotated_image_base64")
    annotated_mime = body.get("annotated_image_mime_type", "image/png")
    ref_b64 = body.get("reference_image_base64")
    ref_mime = body.get("reference_image_mime_type", "image/png")

    # 원본 이미지 다운로드
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.get(image_url)
            resp.raise_for_status()
            image_bytes = resp.content
            mime_type = resp.headers.get("content-type", "image/png").split(";")[0]
    except Exception as exc:
        return JSONResponse({"error": f"이미지 다운로드 실패: {exc}"}, status_code=502)

    # 오버레이 이미지 디코딩
    annotated_bytes = None
    if annotated_b64:
        try:
            annotated_bytes = base64.b64decode(annotated_b64)
        except Exception:
            pass

    # 참조 이미지 디코딩
    ref_bytes = None
    if ref_b64:
        try:
            ref_bytes = base64.b64decode(ref_b64)
        except Exception:
            pass

    # Gemini 편집
    result = edit_image_gemini_bytes(
        image_bytes, instruction, mime_type,
        selection=selection,
        annotated_image_bytes=annotated_bytes,
        annotated_mime_type=annotated_mime,
        reference_image_bytes=ref_bytes,
        reference_image_mime_type=ref_mime,
    )
    if result is None:
        return JSONResponse({"error": "Gemini 이미지 편집 실패"}, status_code=500)

    return JSONResponse({
        "image_base64": base64.b64encode(result).decode(),
        "mime_type": "image/png",
    })


async def enhance_image(request: Request) -> JSONResponse:
    """HWPX 뷰어용 이미지 AI 개선 REST 엔드포인트.

    Body JSON:
        image_base64 (str): base64 인코딩된 원본 이미지
        mime_type    (str): 이미지 MIME 타입 (기본값: image/png)
        instruction  (str): 개선 지시사항 (선택, 기본값 제공)
    Returns:
        { "image_base64": "<base64>", "mime_type": "image/png" }
    """
    from office_mcp.images import edit_image_gemini_bytes

    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"error": "Invalid JSON"}, status_code=400)

    image_b64 = (body.get("image_base64") or "").strip()
    mime_type = (body.get("mime_type") or "image/png").strip()
    instruction = (body.get("instruction") or
                   "이 이미지를 더 선명하고 고품질로 개선해줘. "
                   "세부 디테일을 살리고 색감을 풍부하게, 깔끔한 인포그래픽 스타일로 다시 그려줘.").strip()

    if not image_b64:
        return JSONResponse({"error": "image_base64 is required"}, status_code=400)

    try:
        image_bytes = base64.b64decode(image_b64)
    except Exception:
        return JSONResponse({"error": "Invalid base64 data"}, status_code=400)

    result = edit_image_gemini_bytes(image_bytes, instruction, mime_type)
    if result is None:
        return JSONResponse({"error": "Gemini 이미지 개선 실패"}, status_code=500)

    return JSONResponse({
        "image_base64": base64.b64encode(result).decode(),
        "mime_type": "image/png",
    })


# REST 라우트
rest_routes = [
    Route("/api/edit-slide-image", edit_slide_image, methods=["POST"]),
    Route("/api/enhance-image", enhance_image, methods=["POST"]),
]


app = mcp.http_app(
    transport="http",
    json_response=True,
    stateless_http=True,
    middleware=[
        Middleware(
            CORSMiddleware,
            allow_origins=["*"],
            allow_methods=["*"],
            allow_headers=["*"],
        )
    ],
)
# MCP 앱에 REST 라우트 추가
app.routes.extend(rest_routes)

if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO)
    from office_mcp.config import log_config_summary
    log_config_summary()
    from office_mcp.config import SERVER_PORT
    uvicorn.run(app, host="0.0.0.0", port=SERVER_PORT, lifespan="on")
