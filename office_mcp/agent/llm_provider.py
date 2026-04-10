"""LLM Provider abstraction layer.

Supports swappable backends via LLM_PROVIDER env var:
  openai     - OpenAI direct
  openrouter - OpenRouter (OpenAI-compatible)
  gemini     - Google Gemini
  custom     - Any OpenAI-compatible endpoint (air-gapped / on-prem)
"""

from __future__ import annotations

import json
import logging
import re
import time
from abc import ABC, abstractmethod

logger = logging.getLogger("process-gpt-office-mcp")

_MAX_RETRIES = 3

from ..config import REASONING_EFFORT


# ── helpers ──

def _extract_json_from_text(text: str) -> str:
    """Gemini 응답에서 JSON 블록을 추출한다 (```json ... ``` 또는 전체 텍스트)."""
    m = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    text = text.strip()
    if text.startswith("{"):
        return text
    return text


# ── Abstract base ──

class LLMProvider(ABC):
    """Three-method contract for LLM backends."""

    @abstractmethod
    def call_json(self, prompt_sys: str, prompt_user: str, temperature: float) -> dict:
        ...

    @abstractmethod
    def call_vision_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        images_b64: list[str],
        temperature: float,
    ) -> dict:
        ...

    @abstractmethod
    def call_text(self, prompt_sys: str, prompt_user: str, temperature: float) -> str:
        ...


# ── OpenAI-compatible (OpenAI / OpenRouter / Custom) ──

class OpenAIProvider(LLMProvider):
    """Covers OpenAI, OpenRouter, and any OpenAI-compatible endpoint."""

    def __init__(
        self,
        api_key: str,
        model_name: str,
        timeout: float = 300.0,
        base_url: str | None = None,
        default_headers: dict | None = None,
        reasoning_effort: str | None = None,
    ):
        from openai import OpenAI

        kwargs: dict = {"api_key": api_key, "timeout": timeout}
        if base_url:
            kwargs["base_url"] = base_url
        if default_headers:
            kwargs["default_headers"] = default_headers
        self._client = OpenAI(**kwargs)
        self._model = model_name
        self._reasoning_effort = reasoning_effort  # "low", "medium", "high", or None

    # -- json --
    def call_json(self, prompt_sys: str, prompt_user: str, temperature: float) -> dict:
        return self._do_json(
            messages=[
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": prompt_user},
            ],
            temperature=temperature,
        )

    # -- vision json --
    def call_vision_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        images_b64: list[str],
        temperature: float,
    ) -> dict:
        user_content: list[dict] = [{"type": "text", "text": prompt_user}]
        for img in images_b64:
            user_content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img}", "detail": "high"},
            })
        return self._do_json(
            messages=[
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": user_content},
            ],
            temperature=temperature,
        )

    # -- text --
    def call_text(self, prompt_sys: str, prompt_user: str, temperature: float) -> str:
        kwargs: dict = {
            "model": self._model,
            "messages": [
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": prompt_user},
            ],
            "temperature": temperature,
            "max_tokens": 16384,
        }
        if self._reasoning_effort:
            kwargs["extra_body"] = {
                "reasoning": {"enabled": True, "effort": self._reasoning_effort},
            }
        resp = self._client.chat.completions.create(**kwargs)
        return resp.choices[0].message.content or ""

    # -- internal --
    def _do_json(self, messages: list[dict], temperature: float) -> dict:
        started = time.perf_counter()
        last_exc = None
        for attempt in range(1, _MAX_RETRIES + 1):
            try:
                kwargs: dict = {
                    "model": self._model,
                    "messages": messages,
                    "temperature": temperature,
                    "response_format": {"type": "json_object"},
                    "max_tokens": 16384,
                }
                if self._reasoning_effort:
                    kwargs["extra_body"] = {
                        "reasoning": {"enabled": True, "effort": self._reasoning_effort},
                    }
                # 스트리밍으로 호출 — reasoning 토큰을 실시간 콘솔 출력
                kwargs["stream"] = True
                stream = self._client.chat.completions.create(**kwargs)
                content_parts: list[str] = []
                reasoning_parts: list[str] = []
                finish_reason = "stop"
                _reasoning_started = False
                _content_started = False
                for chunk in stream:
                    delta = chunk.choices[0].delta if chunk.choices else None
                    if delta:
                        # reasoning 토큰
                        rc = getattr(delta, "reasoning_content", None) or getattr(delta, "reasoning", None)
                        if rc:
                            if not _reasoning_started:
                                print(f"\n{'='*60}\n[LLM REASONING 시작]", flush=True)
                                _reasoning_started = True
                            print(rc, end="", flush=True)
                            reasoning_parts.append(rc)
                        # 본문 토큰
                        if delta.content:
                            if not _content_started:
                                if _reasoning_started:
                                    print(f"\n[LLM REASONING 끝]\n{'='*60}", flush=True)
                                print(f"[LLM CONTENT 시작]", flush=True)
                                _content_started = True
                            print(delta.content, end="", flush=True)
                            content_parts.append(delta.content)
                    if chunk.choices and chunk.choices[0].finish_reason:
                        finish_reason = chunk.choices[0].finish_reason
                if _content_started:
                    print(f"\n[LLM CONTENT 끝]\n{'='*60}", flush=True)

                elapsed = time.perf_counter() - started
                content = "".join(content_parts) or "{}"
                reasoning_text = "".join(reasoning_parts)
                if reasoning_text:
                    logger.info("[LLM REASONING] len=%d", len(reasoning_text))
                logger.info(
                    "[LLM RAW] attempt=%d finish_reason=%s content_len=%d\n%s",
                    attempt, finish_reason, len(content), content,
                )
                # finish_reason이 length면 응답이 잘렸을 가능성
                if finish_reason == "length":
                    logger.warning(
                        "[LLM TRUNCATED] 응답이 max_tokens에 의해 잘림 (content_len=%d). 재시도.",
                        len(content),
                    )
                    if attempt < _MAX_RETRIES:
                        time.sleep(1)
                        continue
                # 모델이 ```json ... ``` 으로 감싸는 경우 처리
                data = json.loads(_extract_json_from_text(content))
                # 빈 JSON 감지: _elapsed_s 외에 유의미한 키가 없으면 재시도
                meaningful_keys = [k for k in data if k != "_elapsed_s"]
                if not meaningful_keys:
                    logger.warning(
                        "[LLM EMPTY] 빈 JSON 응답 감지 (attempt=%d, raw=%s). 재시도.",
                        attempt, content[:500],
                    )
                    if attempt < _MAX_RETRIES:
                        time.sleep(1)
                        continue
                data["_elapsed_s"] = round(elapsed, 3)
                return data
            except Exception as exc:
                last_exc = exc
                raw_preview = content[:500] if "content" in dir() else "N/A"
                logger.warning(
                    "[LLM RETRY] attempt=%d/%d error=%s raw=%s",
                    attempt, _MAX_RETRIES, exc, raw_preview,
                )
                if attempt < _MAX_RETRIES:
                    time.sleep(1)
        # 모든 재시도 실패 — 마지막 예외가 있으면 raise, 아니면 빈 dict 반환
        if last_exc:
            raise last_exc
        return {"_elapsed_s": round(time.perf_counter() - started, 3)}


# ── Gemini ──

class GeminiProvider(LLMProvider):
    """Google Gemini via google-genai SDK."""

    def __init__(self, api_key: str, model_name: str):
        from google import genai
        from google.genai import types as genai_types

        self._client = genai.Client(api_key=api_key)
        self._model = model_name
        self._types = genai_types

    # -- json --
    def call_json(self, prompt_sys: str, prompt_user: str, temperature: float) -> dict:
        return self._do_json(prompt_sys, prompt_user, temperature)

    # -- vision json --
    def call_vision_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        images_b64: list[str],
        temperature: float,
    ) -> dict:
        return self._do_json(prompt_sys, prompt_user, temperature, images_b64=images_b64)

    # -- text --
    def call_text(self, prompt_sys: str, prompt_user: str, temperature: float) -> str:
        config = self._types.GenerateContentConfig(
            system_instruction=prompt_sys,
            temperature=temperature,
        )
        resp = self._client.models.generate_content(
            model=self._model,
            contents=prompt_user,
            config=config,
        )
        return resp.text or ""

    # -- internal --
    def _do_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        temperature: float,
        images_b64: list[str] | None = None,
    ) -> dict:
        import base64

        contents: list = []
        contents.append(prompt_user)
        if images_b64:
            for img in images_b64:
                img_bytes = base64.b64decode(img)
                contents.append(
                    self._types.Part.from_bytes(data=img_bytes, mime_type="image/png")
                )

        config = self._types.GenerateContentConfig(
            system_instruction=prompt_sys,
            temperature=temperature,
            response_mime_type="application/json",
        )

        started = time.perf_counter()
        last_exc = None
        for attempt in range(1, _MAX_RETRIES + 1):
            try:
                resp = self._client.models.generate_content(
                    model=self._model,
                    contents=contents,
                    config=config,
                )
                elapsed = time.perf_counter() - started
                raw_text = resp.text or "{}"
                logger.info(
                    "[LLM RAW] attempt=%d content_len=%d content_preview=%s",
                    attempt, len(raw_text), raw_text[:500],
                )
                json_text = _extract_json_from_text(raw_text)
                data = json.loads(json_text)
                # 빈 JSON 감지: _elapsed_s 외에 유의미한 키가 없으면 재시도
                meaningful_keys = [k for k in data if k != "_elapsed_s"]
                if not meaningful_keys:
                    logger.warning(
                        "[LLM EMPTY] 빈 JSON 응답 감지 (attempt=%d, raw=%s). 재시도.",
                        attempt, raw_text[:500],
                    )
                    if attempt < _MAX_RETRIES:
                        time.sleep(1)
                        continue
                data["_elapsed_s"] = round(elapsed, 3)
                return data
            except Exception as exc:
                last_exc = exc
                raw_preview = raw_text[:500] if "raw_text" in dir() else "N/A"
                logger.warning(
                    "[LLM RETRY] attempt=%d/%d error=%s raw=%s",
                    attempt, _MAX_RETRIES, exc, raw_preview,
                )
                if attempt < _MAX_RETRIES:
                    time.sleep(1)
        if last_exc:
            raise last_exc
        return {"_elapsed_s": round(time.perf_counter() - started, 3)}


# ── Factory ──

def create_provider() -> LLMProvider:
    """LLM_PROVIDER 설정에 따라 적절한 provider를 생성한다.

    키 우선순위: LLM_API_KEY > provider별 키 (OPENAI_API_KEY 등)
    URL 우선순위: LLM_BASE_URL > provider별 기본값
    """
    from ..config import (
        LLM_PROVIDER, MODEL_NAME,
        LLM_API_KEY, LLM_BASE_URL, LLM_TIMEOUT_SECONDS,
        OPENAI_API_KEY,
        OPENROUTER_API_KEY,
        GOOGLE_API_KEY,
    )

    if LLM_PROVIDER == "gemini":
        api_key = GOOGLE_API_KEY
        if not api_key:
            raise RuntimeError("LLM_PROVIDER=gemini이지만 API 키가 없습니다 (LLM_API_KEY 또는 GOOGLE_API_KEY)")
        logger.info("[LLM] Provider=gemini, model=%s", MODEL_NAME)
        return GeminiProvider(api_key=api_key, model_name=MODEL_NAME)

    if LLM_PROVIDER == "openrouter":
        api_key = OPENROUTER_API_KEY
        if not api_key:
            raise RuntimeError("LLM_PROVIDER=openrouter이지만 API 키가 없습니다 (LLM_API_KEY 또는 OPENROUTER_API_KEY)")
        base_url = LLM_BASE_URL or "https://openrouter.ai/api/v1"
        logger.info("[LLM] Provider=openrouter, model=%s, reasoning_effort=%s", MODEL_NAME, REASONING_EFFORT)
        return OpenAIProvider(
            api_key=api_key,
            base_url=base_url,
            timeout=LLM_TIMEOUT_SECONDS,
            model_name=MODEL_NAME,
            reasoning_effort=REASONING_EFFORT,
        )

    if LLM_PROVIDER == "custom":
        base_url = LLM_BASE_URL
        if not base_url:
            raise RuntimeError("LLM_PROVIDER=custom이지만 LLM_BASE_URL이 없습니다")
        api_key = LLM_API_KEY or OPENAI_API_KEY or "not-needed"
        logger.info("[LLM] Provider=custom, base_url=%s, model=%s", base_url, MODEL_NAME)
        return OpenAIProvider(
            api_key=api_key,
            base_url=base_url,
            timeout=LLM_TIMEOUT_SECONDS,
            model_name=MODEL_NAME,
        )

    # default: openai
    api_key = OPENAI_API_KEY
    if not api_key:
        raise RuntimeError("LLM_PROVIDER=openai이지만 API 키가 없습니다 (LLM_API_KEY 또는 OPENAI_API_KEY)")
    logger.info("[LLM] Provider=openai, model=%s", MODEL_NAME)
    return OpenAIProvider(
        api_key=api_key,
        base_url=LLM_BASE_URL,
        timeout=LLM_TIMEOUT_SECONDS,
        model_name=MODEL_NAME,
    )
