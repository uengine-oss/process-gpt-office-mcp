import uvicorn
from starlette.middleware.cors import CORSMiddleware
from starlette.middleware import Middleware

from hwpx_mcp.mcp_server import mcp


if __name__ == "__main__":
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
    uvicorn.run(app, host="0.0.0.0", port=1192, lifespan="on")
