"""Rate-limit-safe Gemini Flash intelligence for YT-Bot.

* Hook Analysis: first ~900 chars of the transcript (first 45-60 s).
* Audience Pulse: top 20 comments.
* Audio-to-text backup: only when every caption tier failed - downloads the
  smallest audio stream (worstaudio, max 15 MB), transcribes with Gemini Flash
  into ``[MM:SS] text`` lines and deletes the temp audio immediately.
* No video summaries are ever generated.
* Batch/channel mode groups 10-15 videos per JSON request with a 4-5 s delay and
  ``HTTP 429`` backoff. Falls back to OPENROUTER_API_KEY, then gracefully skips.
"""

from __future__ import annotations

import base64
import json
import logging
import os
import random
import re
import tempfile
import time
from pathlib import Path
from typing import Any

import requests
import yt_dlp

logger = logging.getLogger("yt-bot.ai")

DEFAULT_GEMINI_MODEL = "gemini-3-flash-preview"
GEMINI_FALLBACK_MODELS = (
    "gemini-flash-latest",
    "gemini-3.5-flash",
    "gemini-2.5-flash",
)
GEMINI_MODEL = (os.getenv("GEMINI_MODEL") or DEFAULT_GEMINI_MODEL).strip()
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL") or "openai/gpt-4o-mini"

HOOK_CHARS = 900  # ~45-60 seconds of transcript
TOP_COMMENTS = 20
BATCH_GROUP_SIZE = 12  # 10-15 videos per single JSON request
BATCH_DELAY_MIN, BATCH_DELAY_MAX = 4.0, 5.0
MAX_AUDIO_BYTES = 15 * 1024 * 1024
REQUEST_TIMEOUT = 90
RETRY_DELAYS = (5, 10, 20)


def gemini_key() -> str:
    return (os.getenv("GEMINI_API_KEY") or "").strip()


def openrouter_key() -> str:
    return (os.getenv("OPENROUTER_API_KEY") or "").strip()


def available() -> bool:
    return bool(gemini_key() or openrouter_key())


def _parse_json_block(content: str) -> Any:
    text = (content or "").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.IGNORECASE | re.MULTILINE).strip()
    starts: list[tuple[int, str, str]] = []
    for opener, closer in (("{", "}"), ("[", "]")):
        index = text.find(opener)
        if index != -1:
            starts.append((index, opener, closer))
    if not starts:
        return None
    start, opener, closer = min(starts, key=lambda item: item[0])
    depth = 0
    in_string = False
    escaped = False
    for index in range(start, len(text)):
        char = text[index]
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
        elif char == '"':
            in_string = True
        elif char == opener:
            depth += 1
        elif char == closer:
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(text[start : index + 1])
                except json.JSONDecodeError:
                    return None
    return None


def _post_with_backoff(url: str, *, headers: dict[str, str], payload: dict[str, Any]) -> dict[str, Any]:
    last_error = ""
    for attempt, delay in enumerate((0, *RETRY_DELAYS), start=1):
        if delay:
            time.sleep(delay)
        try:
            response = requests.post(
                url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT
            )
        except requests.RequestException as exc:
            last_error = str(exc)
            logger.warning("AI request failed (attempt %d): %s", attempt, exc)
            continue
        if response.status_code == 429:
            logger.warning("AI rate-limited (attempt %d), backing off", attempt)
            last_error = "HTTP 429"
            continue
        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            last_error = str(exc)[:200]
            logger.warning("AI HTTP error: %s", last_error)
            continue
        return response.json()
    raise RuntimeError(last_error or "AI provider unavailable")


def _gemini_models() -> list[str]:
    """Configured model first, then known-good fallbacks (some are retired/503)."""
    configured = (os.getenv("GEMINI_MODEL") or GEMINI_MODEL or DEFAULT_GEMINI_MODEL).strip()
    models = [configured, *GEMINI_FALLBACK_MODELS]
    seen: list[str] = []
    for model in models:
        if model and model not in seen:
            seen.append(model)
    return seen


def _gemini_request(
    parts: list[dict[str, Any]], *, json_output: bool = False, temperature: float = 0.2
) -> str:
    key = gemini_key()
    if not key:
        raise RuntimeError("GEMINI_API_KEY not set")
    payload: dict[str, Any] = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": temperature, "maxOutputTokens": 2048},
    }
    if json_output:
        payload["generationConfig"]["responseMimeType"] = "application/json"
    last_error = ""
    for model in _gemini_models():
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"{model}:generateContent"
        )
        try:
            data = _post_with_backoff(
                url,
                headers={
                    "Content-Type": "application/json",
                    "x-goog-api-key": key,
                },
                payload=payload,
            )
        except Exception as exc:
            last_error = f"{model}: {str(exc)[:120]}"
            logger.warning("Gemini model %s unavailable: %s", model, str(exc)[:120])
            continue
        candidates = data.get("candidates") or []
        if not candidates:
            last_error = f"{model}: no candidates"
            continue
        content_parts = (candidates[0].get("content") or {}).get("parts") or []
        return "".join(str(part.get("text") or "") for part in content_parts)
    raise RuntimeError(f"Gemini models unavailable ({last_error})")


def _openrouter_request(prompt: str, *, json_output: bool = False) -> str:
    key = openrouter_key()
    if not key:
        raise RuntimeError("OPENROUTER_API_KEY not set")
    payload = {
        "model": OPENROUTER_MODEL,
        "temperature": 0.2,
        "max_tokens": 2048,
        "messages": [
            {
                "role": "system",
                "content": "You are a precise YouTube growth analyst. Reply with JSON only when asked.",
            },
            {"role": "user", "content": prompt},
        ],
    }
    if json_output:
        payload["response_format"] = {"type": "json_object"}
    data = _post_with_backoff(
        OPENROUTER_URL,
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/vashlimusale-cloud/yt-intel-bot",
            "X-Title": "YT-Bot",
        },
        payload=payload,
    )
    choices = data.get("choices") or []
    if not choices:
        raise RuntimeError("OpenRouter returned no choices")
    return str((choices[0].get("message") or {}).get("content") or "")


def generate(prompt: str, *, json_output: bool = False) -> tuple[str, str]:
    """Generate text with Gemini first, then OpenRouter. Returns (text, provider)."""
    errors: list[str] = []
    if gemini_key():
        try:
            return _gemini_request([{"text": prompt}], json_output=json_output), "gemini"
        except Exception as exc:
            errors.append(f"gemini: {str(exc)[:120]}")
    if openrouter_key():
        try:
            return _openrouter_request(prompt, json_output=json_output), "openrouter"
        except Exception as exc:
            errors.append(f"openrouter: {str(exc)[:120]}")
    raise RuntimeError("; ".join(errors) or "No AI provider configured")


def analyze_hook_and_pulse(
    transcript: str, comments: list[dict[str, Any]]
) -> dict[str, Any]:
    """Hook Analysis (first ~900 chars) + Audience Pulse (top 20 comments)."""
    excerpt = (transcript or "")[:HOOK_CHARS].strip()
    comment_payload = [
        {
            "author": str(comment.get("author") or ""),
            "text": str(comment.get("text") or "")[:220],
            "likes": str(comment.get("votes") or "0"),
        }
        for comment in (comments or [])[:TOP_COMMENTS]
        if str(comment.get("text") or "").strip()
    ]
    if not excerpt and not comment_payload:
        return {"skipped": True, "reason": "no transcript and no comments yet"}

    prompt = (
        "Analyse the OPENING of this YouTube video (hook) and the audience reaction. "
        "Do NOT summarise the video. Reply with JSON only:\n"
        '{"hook": {"type": "question|shock|promise|story|number|demo|other", '
        '"analysis": "max 35 words", "strength": 1-10}, '
        '"audience": {"sentiment": "positive|mixed|negative", '
        '"top_themes": ["theme1", "theme2", "theme3"], '
        '"notable_insight": "max 30 words"}}\n\n'
        f"TRANSCRIPT OPENING (~45-60s):\n{excerpt or '(no transcript)'}\n\n"
        "TOP COMMENTS:\n"
        + json.dumps(comment_payload, ensure_ascii=False)
    )
    try:
        content, provider = generate(prompt, json_output=True)
        parsed = _parse_json_block(content)
        if not isinstance(parsed, dict):
            return {"error": "AI returned unparseable JSON", "provider": provider}
        parsed["provider"] = provider
        return parsed
    except Exception as exc:
        logger.warning("Hook/Audience analysis skipped: %s", exc)
        return {"error": str(exc)[:200]}


def analyze_videos_batch(videos: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Group 10-15 videos per Gemini JSON request (4-5 s between requests)."""
    results: list[dict[str, Any]] = []
    if not videos or not available():
        return results
    for start in range(0, len(videos), BATCH_GROUP_SIZE):
        group = videos[start : start + BATCH_GROUP_SIZE]
        payload = [
            {
                "index": start + offset,
                "title": str(item.get("title") or "")[:140],
                "views": item.get("views"),
                "likes": item.get("likes"),
                "comments": item.get("comment_count"),
                "opening": str(item.get("transcript") or "")[:300],
                "top_comment": str(item.get("top_comment") or "")[:160],
            }
            for offset, item in enumerate(group)
        ]
        prompt = (
            "For each YouTube video below, give a short growth read. Do NOT summarise "
            "the video. Reply with ONLY a JSON array; one object per input index:\n"
            '{"index": <int>, "hook_type": "question|shock|promise|story|number|demo|other", '
            '"hook_strength": 1-10, "audience_sentiment": "positive|mixed|negative", '
            '"recommendation": "max 20 words"}\n\n' + json.dumps(payload, ensure_ascii=False)
        )
        parsed: Any = None
        try:
            content, provider = generate(prompt, json_output=True)
            parsed = _parse_json_block(content)
        except Exception as exc:
            logger.warning("Batch AI analysis skipped for group %d: %s", start, exc)
            provider = "none"
        if isinstance(parsed, dict):
            parsed = parsed.get("videos") or parsed.get("results") or []
        if not isinstance(parsed, list):
            parsed = []
        by_index = {
            entry.get("index"): entry
            for entry in parsed
            if isinstance(entry, dict) and isinstance(entry.get("index"), int)
        }
        for offset, item in enumerate(group):
            entry = by_index.get(start + offset) or {}
            results.append(
                {
                    "index": start + offset,
                    "title": item.get("title"),
                    "hook_type": entry.get("hook_type", ""),
                    "hook_strength": entry.get("hook_strength", ""),
                    "audience_sentiment": entry.get("audience_sentiment", ""),
                    "recommendation": entry.get("recommendation", ""),
                    "provider": provider,
                }
            )
        if start + BATCH_GROUP_SIZE < len(videos):
            time.sleep(random.uniform(BATCH_DELAY_MIN, BATCH_DELAY_MAX))
    return results


def _audio_mime(suffix: str) -> str:
    return {
        ".m4a": "audio/mp4",
        ".mp4": "audio/mp4",
        ".webm": "audio/webm",
        ".opus": "audio/ogg",
        ".ogg": "audio/ogg",
        ".mp3": "audio/mpeg",
        ".wav": "audio/wav",
    }.get(suffix.lower(), "audio/mp4")


def _ensure_timestamps(text: str) -> str:
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    if not lines:
        return ""
    if all(re.match(r"^\[\d{1,2}:\d{2}\]", line) for line in lines):
        return "\n".join(lines)
    stamped: list[str] = []
    for index, line in enumerate(lines):
        seconds = index * 5
        text_value = re.sub(r"^\[\d{1,2}:\d{2}\]\s*", "", line)
        stamped.append(f"[{seconds // 60:02d}:{seconds % 60:02d}] {text_value}")
    return "\n".join(stamped)


def audio_transcription_enabled() -> bool:
    """Opt-in audio-to-text backup (off by default to keep single jobs fast)."""
    return (os.getenv("ENABLE_AUDIO_TRANSCRIPTION") or "false").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def transcribe_video_audio(video_id: str) -> str:
    """Gemini Flash audio transcription backup for videos with no captions."""
    if not audio_transcription_enabled():
        logger.info("Audio transcription disabled (ENABLE_AUDIO_TRANSCRIPTION=false)")
        return ""
    if not gemini_key():
        logger.info("Audio transcription skipped: GEMINI_API_KEY not set")
        return ""
    with tempfile.TemporaryDirectory(prefix="yt_audio_") as temp_dir:
        options = {
            "quiet": True,
            "no_warnings": True,
            "format": "worstaudio/worst",
            "outtmpl": str(Path(temp_dir) / "audio.%(ext)s"),
            "noplaylist": True,
            "socket_timeout": 40,
            "max_filesize": MAX_AUDIO_BYTES,
            "retries": 2,
        }
        try:  # reuse the decoded YOUTUBE_COOKIES_B64 file when present
            from bot import apply_cookies

            options = apply_cookies(options)
        except Exception:
            pass
        try:
            with yt_dlp.YoutubeDL(options) as ydl:
                ydl.download([f"https://www.youtube.com/watch?v={video_id}"])
        except Exception as exc:
            logger.warning("Audio download failed for %s: %s", video_id, exc)
            return ""
        files = [path for path in Path(temp_dir).iterdir() if path.is_file()]
        if not files:
            return ""
        audio_path = max(files, key=lambda path: path.stat().st_size)
        if audio_path.stat().st_size > MAX_AUDIO_BYTES:
            logger.warning("Audio too large for %s (%d bytes)", video_id, audio_path.stat().st_size)
            return ""
        try:
            encoded = base64.b64encode(audio_path.read_bytes()).decode("ascii")
        except OSError:
            return ""
        mime = _audio_mime(audio_path.suffix)
        prompt = (
            "Transcribe the spoken content of this audio. Output ONLY transcript lines "
            "in the format '[MM:SS] text', one line per short phrase, English if spoken "
            "in English, otherwise original language. No commentary."
        )
        try:
            content = _gemini_request(
                [
                    {"text": prompt},
                    {"inline_data": {"mime_type": mime, "data": encoded}},
                ]
            )
        except Exception as exc:
            logger.warning("Gemini transcription failed for %s: %s", video_id, exc)
            return ""
        finally:
            try:
                audio_path.unlink(missing_ok=True)
            except OSError:
                pass
    return _ensure_timestamps(content)


def token_note(parts: list[str]) -> str:
    """Human readable note for plugin responses about what was fetched."""
    return (
        "Fetched/used: " + ", ".join(dict.fromkeys(parts)) +
        ". Available but not used (to stay under ChatGPT's ~50KB action limit): "
        "full transcript, all comments, all channel videos, community posts, "
        "collaborations, AI audio transcription."
    )
