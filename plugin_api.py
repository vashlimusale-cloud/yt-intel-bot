"""ChatGPT Custom GPT Actions / Claude MCP plugin layer for YT-Bot.

Serves an OpenAPI 3.1 schema at /openapi.json plus four JSON endpoints:
  POST /api/smart_query  {"url": "...", "request": "what the AI wants"}
  POST /api/video        {"url": "..."}
  POST /api/channel      {"url": "..."}
  POST /api/similar      {"url": "..."}

Responses are deliberately lightweight (transcript excerpt ~900 chars, top 20
comments, top videos) - well under ChatGPT's ~50KB action limit.
Optional auth: set PLUGIN_API_KEY and send X-API-Key.
"""

from __future__ import annotations

import json
import logging
import os
from typing import Any
from urllib.parse import urlparse

import ai_analyzer
import bot
import recommender

logger = logging.getLogger("yt-bot.plugin")

MODULES = ["video", "transcript", "comments", "hook_audience", "channel", "similar", "viral"]
KEYWORD_INTENTS = {
    "transcript": ("transcript", "caption", "said", "words", "transcribe", "quote"),
    "comments": ("comment", "audience", "reaction", "reply", "like"),
    "hook_audience": ("hook", "opening", "intro", "retention", "audience", "sentiment", "think"),
    "channel": ("channel", "profile", "subscriber", "about", "banner", "logo", "stats"),
    "similar": ("similar", "competitor", "peer", "rival", "alternative", "niche"),
    "viral": ("viral", "top video", "best video", "most viewed", "popular"),
    "video": ("video", "metadata", "thumbnail", "title", "duration", "views"),
}
MAX_TRANSCRIPT = 900
MAX_COMMENTS = 20
MAX_COMMENT_CHARS = 220


def _auto_failover(url: str, reason: str = "") -> tuple[int, dict[str, Any]]:
    """AUTO mode: route blocked Render extraction to GitHub Actions (HTTP 202)."""
    try:
        mode = bot.execution_mode.current_mode()
    except Exception:
        mode = "AUTO"
    if mode != "AUTO":
        return 200, {
            "ok": False,
            "warnings": [
                f"YouTube extraction blocked on this host ({reason[:160]}); "
                f"mode={mode} does not auto-failover"
            ],
        }
    chat_id = (os.getenv("TELEGRAM_CHAT_ID") or "").strip()
    if not chat_id:
        return 200, {
            "ok": False,
            "warnings": [f"extraction blocked ({reason[:160]}); TELEGRAM_CHAT_ID missing"],
        }
    ok, detail = bot.dispatch_github_batch([url], chat_id)
    if ok:
        logger.info("Plugin extraction blocked; dispatched %s to GitHub Actions", url)
        return 202, {
            "ok": True,
            "routed_to": "github_actions",
            "message": (
                "Render blocked; auto-dispatched to GitHub Actions and results "
                "will be delivered to Telegram."
            ),
            "warnings": [f"deep extraction blocked on this host: {reason[:200]}"],
        }
    logger.warning("Plugin failover dispatch failed: %s", detail)
    return 200, {
        "ok": False,
        "warnings": [
            f"extraction blocked: {reason[:160]}",
            f"failover unavailable: {detail[:160]}",
        ],
    }


def public_base_url() -> str:
    base = (os.getenv("PUBLIC_BASE_URL") or os.getenv("RENDER_EXTERNAL_URL") or "").strip()
    return base.rstrip("/") or "https://your-service.onrender.com"


def api_key() -> str:
    return (os.getenv("PLUGIN_API_KEY") or "").strip()


def authorized(headers: Any) -> bool:
    key = api_key()
    if not key:
        return True
    return str(headers.get("X-API-Key") or headers.get("x-api-key") or "") == key


def openapi_schema() -> dict[str, Any]:
    schema: dict[str, Any] = {
        "openapi": "3.1.0",
        "info": {
            "title": "YT-Bot Intelligence API",
            "version": "1.0.0",
            "description": (
                "Lightweight YouTube intelligence for ChatGPT Custom GPT Actions. "
                "Use /api/smart_query with a natural-language request to fetch only the "
                "modules you need. Responses are capped under ~50KB."
            ),
        },
        "servers": [{"url": public_base_url()}],
        "paths": {
            "/api/smart_query": {
                "post": {
                    "operationId": "smartQuery",
                    "summary": "Ask for any YouTube insight on a URL",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["url"],
                                    "properties": {
                                        "url": {"type": "string"},
                                        "request": {"type": "string"},
                                    },
                                }
                            }
                        },
                    },
                    "responses": {"200": {"description": "Structured intelligence JSON"}},
                }
            },
            "/api/video": {
                "post": {
                    "operationId": "videoIntelligence",
                    "summary": "Video metadata, transcript excerpt, top comments, hook analysis",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["url"],
                                    "properties": {"url": {"type": "string"}},
                                }
                            }
                        },
                    },
                    "responses": {"200": {"description": "Video intelligence JSON"}},
                }
            },
            "/api/channel": {
                "post": {
                    "operationId": "channelOverview",
                    "summary": "Channel profile, stats, avatar/banner, top videos",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["url"],
                                    "properties": {"url": {"type": "string"}},
                                }
                            }
                        },
                    },
                    "responses": {"200": {"description": "Channel overview JSON"}},
                }
            },
            "/api/similar": {
                "post": {
                    "operationId": "similarChannels",
                    "summary": "Find 5-8 peer channels in the same niche",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["url"],
                                    "properties": {"url": {"type": "string"}},
                                }
                            }
                        },
                    },
                    "responses": {"200": {"description": "Similar channels JSON"}},
                }
            },
        },
    }
    if api_key():
        schema["components"] = {
            "securitySchemes": {
                "ApiKeyAuth": {"type": "apiKey", "in": "header", "name": "X-API-Key"}
            }
        }
        schema["security"] = [{"ApiKeyAuth": []}]
    return schema


# --------------------------------------------------------------------------- #
# Lightweight payload builders
# --------------------------------------------------------------------------- #
def _video_id(url: str) -> str:
    video_id = bot.extract_video_id(url)
    if not video_id:
        raise ValueError("Not a YouTube video URL")
    return video_id


def _channel_meta(channel_url: str) -> dict[str, Any]:
    listing, _ = bot.fetch_channel_listing(channel_url, "videos", 1)
    if isinstance(listing, dict) and listing.get("_fallback"):
        raise bot.DatacenterBlockError("Channel listing blocked on this host")
    profile = bot.build_channel_profile(
        channel_url, listing, bot.fetch_channel_about(channel_url)
    )
    about = bot.fetch_channel_about(channel_url)
    banner_url = bot.channel_banner_url(listing, about)
    return {
        "name": profile.name,
        "handle": profile.handle,
        "channel_url": profile.url,
        "channel_id": profile.channel_id,
        "subscribers": profile.subscribers,
        "total_views": profile.total_views,
        "country": profile.country,
        "joined_date": profile.joined_date,
        "avatar_url": profile.avatar_url,
        "banner_url": banner_url,
        "email": profile.email,
        "keywords": profile.keywords[:10],
    }


def video_payload(url: str) -> dict[str, Any]:
    video_id = _video_id(url)
    watch_url = f"https://www.youtube.com/watch?v={video_id}"
    info = bot.fetch_metadata(watch_url)
    blocked = bool(info.get("_fallback"))
    warnings = list(info.get("_warnings") or [])
    transcript = bot.fetch_transcript(video_id, bot._caption_tracks(info))
    audio_used = False
    if not transcript and ai_analyzer.available():
        transcript = ai_analyzer.transcribe_video_audio(video_id)
        audio_used = bool(transcript)
    comments = bot.fetch_comments(watch_url, limit=MAX_COMMENTS)
    hook = ai_analyzer.analyze_hook_and_pulse(transcript, comments)
    channel: dict[str, Any] = {}
    try:
        channel = recommender.resolve_channel(watch_url)
    except Exception as exc:
        warnings.append(f"channel lookup failed: {str(exc)[:120]}")
    channel_meta: dict[str, Any] = {
        "name": channel.get("name"),
        "handle": channel.get("handle"),
        "channel_url": channel.get("channel_url"),
    }
    if channel.get("channel_url"):
        try:
            channel_meta = _channel_meta(channel["channel_url"])
        except bot.DatacenterBlockError as exc:
            blocked = True
            warnings.append(f"channel deep extraction blocked: {exc}")
            logger.warning("Plugin channel lookup blocked: %s", exc)
        except Exception as exc:  # metadata is best effort
            warnings.append(f"channel metadata unavailable: {str(exc)[:120]}")
            logger.warning("Plugin channel lookup failed: %s", exc)

    transcript_used = transcript[:MAX_TRANSCRIPT]
    return {
        "video": {
            "video_id": video_id,
            "title": str(info.get("title") or ""),
            "url": watch_url,
            "upload_date": bot._iso_upload_date(info),
            "views": info.get("view_count"),
            "likes": info.get("like_count"),
            "comments": info.get("comment_count"),
            "duration_sec": info.get("duration"),
            "categories": info.get("categories") or [],
            "tags": (info.get("tags") or [])[:15],
            "thumbnail_url": str(info.get("thumbnail") or bot.thumbnail_url_for(video_id)),
            "description_excerpt": str(info.get("description") or "")[:600],
        },
        "transcript_excerpt": transcript_used,
        "transcript_chars_total": len(transcript),
        "transcript_audio_backup_used": audio_used,
        "comments_sample": [
            {
                "author": str(comment.get("author") or ""),
                "likes": str(comment.get("votes") or "0"),
                "text": str(comment.get("text") or "")[:MAX_COMMENT_CHARS],
            }
            for comment in comments[:MAX_COMMENTS]
        ],
        "hook_analysis": hook,
        "channel": channel_meta,
        "blocked": blocked,
        "warnings": warnings,
        "token_usage_note": ai_analyzer.token_note(
            [
                "video metadata",
                "transcript excerpt (first ~900 chars)" if transcript_used else "transcript (unavailable)",
                f"top {len(comments[:MAX_COMMENTS])} comments",
                "hook + audience analysis" if hook and not hook.get("skipped") else "no AI hook analysis",
                "parent channel overview",
            ]
        ),
    }


def channel_payload(url: str) -> dict[str, Any]:
    channel = recommender.resolve_channel(url)
    channel_url = str(channel.get("channel_url") or "")
    if not channel_url:
        raise RuntimeError("Could not resolve that channel")
    meta = _channel_meta(channel_url)
    info = recommender._flat_info(f"{channel_url.rstrip('/')}/videos", 20)
    blocked = bool(info.get("_fallback"))
    warnings = list(info.get("_warnings") or [])
    entries = recommender._entries(info)
    ranked = sorted(entries, key=lambda entry: entry.get("view_count") or 0, reverse=True)[:5]
    return {
        "channel": meta,
        "top_videos": [
            {
                "title": str(entry.get("title") or ""),
                "url": f"https://www.youtube.com/watch?v={entry['id']}",
                "views": entry.get("view_count"),
                "duration": entry.get("duration"),
            }
            for entry in ranked
        ],
        "blocked": blocked,
        "warnings": warnings,
        "token_usage_note": ai_analyzer.token_note(
            ["channel profile", "channel about", "banner", "avatar", "top 5 videos"]
        ),
    }


def similar_payload(url: str) -> dict[str, Any]:
    result = recommender.similar_channels(url, limit=8)
    return {
        "channel": result.get("channel"),
        "search_query": result.get("query"),
        "similar_channels": result.get("channels", []),
        "blocked": False,
        "warnings": [],
        "token_usage_note": ai_analyzer.token_note(
            ["parent channel", "niche keywords", "YouTube search results (5-8 peers)"]
        ),
    }


def _ai_intent(request: str) -> list[str] | None:
    if not ai_analyzer.available():
        return None
    prompt = (
        "Pick the lightweight YouTube data modules needed for this request. "
        f"Allowed modules: {MODULES}. Reply JSON only: "
        '{"modules": ["..."]}. Choose the minimum needed (1-3 modules).\n'
        f"Request: {request}"
    )
    try:
        content, _provider = ai_analyzer.generate(prompt, json_output=True)
        parsed = ai_analyzer._parse_json_block(content)
        if isinstance(parsed, dict):
            raw = parsed.get("modules") or []
            chosen = [module for module in raw if module in MODULES]
            if chosen:
                return chosen[:3]
    except Exception as exc:
        logger.warning("AI intent detection failed, using keywords: %s", exc)
    return None


def detect_modules(request: str) -> list[str]:
    modules = _ai_intent(request)
    if modules:
        return modules
    lowered = (request or "").lower()
    chosen = [
        module
        for module, words in KEYWORD_INTENTS.items()
        if any(word in lowered for word in words)
    ]
    if not chosen:
        chosen = ["video", "transcript"]
    # Keep at most 3 modules so the action stays fast.
    return chosen[:3]


def smart_query_payload(url: str, request: str) -> dict[str, Any]:
    modules = detect_modules(request)
    data: dict[str, Any] = {}
    blocked = False
    warnings: list[str] = []
    if any(module in modules for module in ("video", "transcript", "comments", "hook_audience")):
        video = video_payload(url)
        blocked = bool(video.get("blocked"))
        warnings.extend(video.get("warnings") or [])
        payload = {
            "video": video["video"],
            "channel": video["channel"],
        }
        if "transcript" in modules or "hook_audience" in modules:
            payload["transcript_excerpt"] = video["transcript_excerpt"]
            payload["transcript_chars_total"] = video["transcript_chars_total"]
        if "comments" in modules or "hook_audience" in modules:
            payload["comments_sample"] = video["comments_sample"]
        if "hook_audience" in modules:
            payload["hook_analysis"] = video["hook_analysis"]
        data = payload
    if "channel" in modules:
        try:
            data["channel"] = channel_payload(url)["channel"]
        except Exception as exc:
            data["channel_error"] = str(exc)[:160]
            if bot.is_datacenter_block(exc) or isinstance(exc, bot.DatacenterBlockError):
                blocked = True
                warnings.append(f"channel blocked: {str(exc)[:160]}")
    if "similar" in modules:
        try:
            data["similar_channels"] = similar_payload(url)
        except Exception as exc:
            data["similar_error"] = str(exc)[:160]
            if bot.is_datacenter_block(exc) or isinstance(exc, bot.DatacenterBlockError):
                blocked = True
                warnings.append(f"similar blocked: {str(exc)[:160]}")
    if "viral" in modules:
        try:
            data["viral_videos"] = recommender.viral_videos(url, limit=5)
        except Exception as exc:
            data["viral_error"] = str(exc)[:160]
            if bot.is_datacenter_block(exc) or isinstance(exc, bot.DatacenterBlockError):
                blocked = True
                warnings.append(f"viral blocked: {str(exc)[:160]}")
    deep_modules = {"video", "channel", "hook_audience", "similar", "viral"}
    return {
        "url": url,
        "request": request,
        "modules_used": modules,
        "data": data,
        "blocked": blocked and bool(deep_modules.intersection(modules)),
        "warnings": warnings,
        "token_usage_note": ai_analyzer.token_note(modules),
    }


# --------------------------------------------------------------------------- #
# HTTP dispatch (called from bot.HealthHandler)
# --------------------------------------------------------------------------- #
def is_plugin_path(path: str) -> bool:
    return path in ("/openapi.json", "/api/smart_query", "/api/video", "/api/channel", "/api/similar")


def handle_request(
    method: str, path: str, headers: Any, body: bytes
) -> tuple[int, dict[str, Any]]:
    if not authorized(headers):
        return 401, {"error": "unauthorized", "hint": "Send the configured X-API-Key header"}
    if method == "GET" and path == "/openapi.json":
        return 200, openapi_schema()
    if method != "POST":
        return 405, {"error": "method not allowed"}
    try:
        payload = json.loads(body.decode("utf-8") or "{}")
    except ValueError:
        return 400, {"error": "invalid JSON body"}
    if not isinstance(payload, dict):
        return 400, {"error": "JSON body must be an object"}
    url = str(payload.get("url") or "").strip()
    if not url:
        return 400, {"error": "missing 'url'"}
    try:
        if path == "/api/smart_query":
            payload = smart_query_payload(url, str(payload.get("request") or "analyse this"))
        elif path == "/api/video":
            payload = video_payload(url)
        elif path == "/api/channel":
            payload = channel_payload(url)
        elif path == "/api/similar":
            payload = similar_payload(url)
        else:
            return 404, {"error": "unknown endpoint"}
        if payload.pop("blocked", False):
            reason = "; ".join(payload.get("warnings") or []) or "extraction blocked"
            return _auto_failover(url, reason)
        return 200, payload
    except ValueError as exc:
        return 400, {"error": str(exc)[:300]}
    except bot.DatacenterBlockError as exc:
        return _auto_failover(url, str(exc))
    except Exception as exc:
        logger.exception("Plugin endpoint extraction issue: %s", exc)
        if bot.is_datacenter_block(exc):
            return _auto_failover(url, str(exc))
        # Never surface an unhandled 500 for YouTube extraction problems.
        return 200, {
            "ok": False,
            "warnings": [f"extraction failed: {str(exc)[:200]}"],
        }
    return 404, {"error": "unknown endpoint"}
