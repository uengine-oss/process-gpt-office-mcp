import os
from pathlib import Path


def _load_env_file() -> None:
    root = Path(__file__).resolve().parents[1]
    env_path = root / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#") or "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        if key and key not in os.environ:
            os.environ[key] = value


_load_env_file()


def _get_bool_env(name: str, default: bool) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "y", "on")


def _get_float_env(name: str, default: float) -> float:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()

MODEL_NAME = "gpt-5.1"
GEMINI_IMAGE_MODEL = "gemini-3.1-flash-image-preview"
OPENAI_TIMEOUT_SECONDS = _get_float_env("OPENAI_TIMEOUT_SECONDS", 300.0)
GEMINI_IMAGE_TIMEOUT_SECONDS = _get_float_env("GEMINI_IMAGE_TIMEOUT_SECONDS", 120.0)

MAX_CONCURRENT_LLM = 4
SMALL_CELL_HEIGHT_MM = 3
SMALL_CELL_WIDTH_MM = 8
IMAGE_MIN_WIDTH_MM = 80
IMAGE_MIN_HEIGHT_MM = 25

LOG_PATH = "./run.log"
IMAGE_GENERATION_ENABLED = _get_bool_env("IMAGE_GENERATION_ENABLED", False)
DEBUG_OUTPUT_ENABLED = False
DEBUG_OUTPUT_DIR = "debug_outputs"
