"""Viral-video and similar-channel discovery using yt-dlp (no API keys)."""

from __future__ import annotations

import logging
import re
from typing import Any

import yt_dlp

import bot

logger = logging.getLogger("yt-bot.recommender")

FLAT_OPTS = {
    "quiet": True,
    "no_warnings": True,
    "skip_download": True,
    "extract_flat": "in_playlist",
    "socket_timeout": 30,
    "retries": 2,
}
STOPWORDS = {
    "the", "and", "for", "you", "your", "with", "this", "that", "from", "have",
    "are", "was", "were", "will", "would", "can", "could", "about", "into", "our",
    "their", "they", "them", "how", "what", "why", "when", "who", "its", "all",
    "more", "most", "new", "video", "videos", "channel", "subscribe", "youtube",
    "official", "every", "everyone", "here", "there", "then", "than", "also",
}


def _flat_info(url: str, limit: int | None = None) -> dict[str, Any]:
    """Flat yt-dlp extraction with the datacenter-safe client + HTML fallback."""
    return bot.extract_yt_info_resilient(url, flat=True, playlistend=limit)


def _entries(info: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        entry
        for entry in (info.get("entries") or [])
        if isinstance(entry, dict) and entry.get("id")
    ]


def resolve_channel(video_or_channel_url: str) -> dict[str, Any]:
    """Return the parent channel of a video URL, or the channel itself."""
    channel_url = bot.extract_channel_url(video_or_channel_url)
    if channel_url:
        info = _flat_info(f"{channel_url.rstrip('/')}/videos", 1)
        resolved = bot._resolve_channel_base(channel_url, info)
        return {
            "name": str(info.get("channel") or info.get("title") or ""),
            "handle": str(info.get("uploader_id") or ""),
            "channel_url": resolved,
            "channel_id": str(info.get("channel_id") or ""),
            "origin": "channel",
        }
    video_id = bot.extract_video_id(video_or_channel_url)
    if not video_id:
        raise ValueError("Send a YouTube video or channel URL")
    meta = bot.fetch_metadata(f"https://www.youtube.com/watch?v={video_id}")
    return {
        "name": str(meta.get("channel") or meta.get("uploader") or ""),
        "handle": str(meta.get("uploader_id") or ""),
        "channel_url": str(meta.get("channel_url") or ""),
        "channel_id": str(meta.get("channel_id") or ""),
        "origin": "video",
        "video_id": video_id,
        "video_title": str(meta.get("title") or ""),
    }


def viral_videos(video_or_channel_url: str, limit: int = 5) -> dict[str, Any]:
    """Top N most-viewed long-form videos on the parent channel."""
    channel = resolve_channel(video_or_channel_url)
    if not channel.get("channel_url"):
        raise RuntimeError("Could not resolve the parent channel")
    info = _flat_info(f"{channel['channel_url'].rstrip('/')}/videos", 60)
    if isinstance(info, dict) and info.get("_fallback"):
        raise bot.DatacenterBlockError(
            "Channel video listing blocked on this datacenter IP"
        )
    entries = _entries(info)
    ranked = sorted(
        entries,
        key=lambda entry: entry.get("view_count") or 0,
        reverse=True,
    )[: max(1, limit)]
    videos = [
        {
            "title": str(entry.get("title") or ""),
            "url": f"https://www.youtube.com/watch?v={entry['id']}",
            "video_id": entry["id"],
            "views": entry.get("view_count"),
            "duration": entry.get("duration"),
            "uploaded": entry.get("upload_date") or "",
        }
        for entry in ranked
    ]
    logger.info("Viral videos for %s: %d found from %d entries", channel.get("name"), len(videos), len(entries))
    return {"channel": channel, "videos": videos}


def similar_channels(video_or_channel_url: str, limit: int = 8) -> dict[str, Any]:
    """Find 5-8 peer channels in the same niche via a YouTube search."""
    channel = resolve_channel(video_or_channel_url)
    if not channel.get("channel_url"):
        raise RuntimeError("Could not resolve the parent channel")
    base = channel["channel_url"].rstrip("/")
    info = _flat_info(f"{base}/videos", 1)
    profile = bot.build_channel_profile(base, info, bot.fetch_channel_about(base))
    words = list(profile.keywords[:4])
    words += [
        word
        for word in re.findall(r"[A-Za-z][A-Za-z0-9+]{3,}", profile.description or "")
        if word.lower() not in STOPWORDS
    ][:6]
    query = " ".join(dict.fromkeys(words))[:120] or profile.name
    logger.info("Similar channels query for %s: %r", channel.get("name"), query)

    search_info = _flat_info(f"ytsearch20:{query}", 20)
    found: dict[str, dict[str, Any]] = {}
    own_id = str(channel.get("channel_id") or "")
    own_name = str(channel.get("name") or "").lower()
    for entry in _entries(search_info):
        name = str(entry.get("channel") or entry.get("uploader") or "").strip()
        handle = str(entry.get("uploader_id") or "").strip()
        channel_url = str(entry.get("channel_url") or "").strip()
        if not name or name.lower() == own_name:
            continue
        if own_id and own_id in channel_url:
            continue
        key = (handle or channel_url or name).lower()
        if key in found:
            continue
        found[key] = {
            "name": name,
            "handle": handle,
            "channel_url": channel_url,
            "subscribers": entry.get("channel_follower_count"),
        }
        if len(found) >= max(1, limit):
            break
    return {"channel": channel, "query": query, "channels": list(found.values())}
