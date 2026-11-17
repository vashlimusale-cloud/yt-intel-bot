"""YT-Bot -- turn YouTube links into research dossiers.

Send the bot a YouTube video, Shorts or channel URL and it answers with:
  * Video/Shorts mode: the thumbnail plus a `.md` file with metadata,
    description, transcript and every comment.
  * Channel mode: the channel avatar, a multi-sheet Excel intelligence workbook
    (`<Channel>_Channel_Intel.xlsx`) and a Markdown SEO audit report.

Usage: paste your @BotFather token into `.env`, then double-click `start_bot.bat`.
"""

from __future__ import annotations

import asyncio
import html
import io
import json
import logging
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse

import requests
import yt_dlp
from dotenv import load_dotenv
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
    Update,
)
from telegram.constants import ChatAction
from telegram.error import TelegramError, TimedOut
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
from youtube_comment_downloader import SORT_BY_POPULAR, YoutubeCommentDownloader
from youtube_transcript_api import NoTranscriptFound, YouTubeTranscriptApi

import ai_analyzer
import execution_mode
from progress import ProgressCard

try:  # channel audits need openpyxl; single-video mode still works without it
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    OPENPYXL_AVAILABLE = True
except ImportError:  # pragma: no cover
    Workbook = None  # type: ignore[assignment]
    Alignment = Font = PatternFill = None  # type: ignore[assignment]
    get_column_letter = None  # type: ignore[assignment]
    OPENPYXL_AVAILABLE = False

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
TOKEN_PLACEHOLDER = "PASTE_YOUR_BOTFATHER_TOKEN_HERE"
TOKEN_RE = re.compile(r"^\d{5,}:[A-Za-z0-9_-]{20,}$")
PREFERRED_LANGUAGES = ["hi", "en", "en-IN"]
TRANSCRIPT_LANGUAGE_PRIORITY = ["hi", "en", "en-IN", "en-US", "en-GB"]
MAX_COMMENTS = 0  # 0 = every comment; set e.g. 500 to cap very large threads
MAX_CAPTION = 1024  # Telegram's hard limit for photo/document captions
HTTP_TIMEOUT = 30

# Telegram HTTP timeouts (seconds): uploading a big Markdown file can be slow.
TELEGRAM_CONNECT_TIMEOUT = 30
TELEGRAM_READ_TIMEOUT = 120
TELEGRAM_WRITE_TIMEOUT = 120
TELEGRAM_POOL_TIMEOUT = 30
DOCUMENT_SEND_ATTEMPTS = 2  # one retry if the .md upload times out
DOCUMENT_RETRY_DELAY = 2.0  # seconds to wait before retrying

OUTPUT_DIR = BASE_DIR / "output"  # every generated .md/.xlsx is kept here permanently
MAX_CHANNEL_VIDEOS = 100  # 0 = every long-form video on the channel
MAX_CHANNEL_SHORTS = 100  # 0 = every short on the channel
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

# Supports youtube.com/watch?v=ID, youtu.be/ID, youtube.com/shorts/ID,
# plus /embed/ and /live/ variants.
YOUTUBE_ID_RE = re.compile(
    r"(?:youtube\.com/(?:watch\?(?:.*?&)?v=|shorts/|embed/|live/)|youtu\.be/)"
    r"([A-Za-z0-9_-]{11})(?![A-Za-z0-9_-])",
    re.IGNORECASE,
)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s | %(levelname)-8s | %(message)s"
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger("yt-bot")


# --------------------------------------------------------------------------- #
# Small formatting helpers
# --------------------------------------------------------------------------- #
def extract_video_id(text: str) -> str | None:
    """Return the 11-character video id from a YouTube watch/shorts/youtu.be link."""
    if not text:
        return None
    match = YOUTUBE_ID_RE.search(text)
    return match.group(1) if match else None


def safe_stem(title: str, fallback: str = "YouTube") -> str:
    """Build a Windows-safe filename stem from a title."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", title or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:80].strip(" .") or fallback


def safe_filename(title: str, video_id: str) -> str:
    """Build a Windows-safe `<Title>_<video_id>.md` filename."""
    return f"{safe_stem(title, 'YouTube_Video')}_{video_id}.md"


def _fmt_int(value: Any) -> str:
    if value is None:
        return "N/A"
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return str(value)


def _fmt_duration(seconds: Any) -> str:
    if seconds is None:
        return "N/A"
    try:
        total = int(seconds)
    except (TypeError, ValueError):
        return str(seconds)
    if total <= 0:
        return "N/A"
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours}:{minutes:02d}:{secs:02d}" if hours else f"{minutes}:{secs:02d}"


def _fmt_timestamp(seconds: Any) -> str:
    try:
        total = max(0, int(seconds or 0))
    except (TypeError, ValueError):
        total = 0
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"


def _fmt_upload_date(info: dict[str, Any]) -> str:
    raw = info.get("upload_date")
    if raw:
        try:
            return datetime.strptime(str(raw), "%Y%m%d").strftime("%d %B %Y")
        except ValueError:
            return str(raw)
    timestamp = info.get("timestamp") or info.get("release_timestamp")
    if timestamp:
        try:
            moment = datetime.fromtimestamp(int(timestamp), tz=timezone.utc)
            return moment.strftime("%d %B %Y")
        except (TypeError, ValueError, OSError, OverflowError):
            pass
    return "N/A"


# --------------------------------------------------------------------------- #
# Scraping (all of this runs in a worker thread)
# --------------------------------------------------------------------------- #
def fetch_metadata(url: str) -> dict[str, Any]:
    """Fetch full video metadata with yt-dlp without downloading the video."""
    options = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "noplaylist": True,
        "socket_timeout": HTTP_TIMEOUT,
        "retries": 3,
    }
    with yt_dlp.YoutubeDL(options) as ydl:
        info = ydl.extract_info(url, download=False)
    if not isinstance(info, dict):
        raise RuntimeError("yt-dlp returned no metadata for this video.")
    if info.get("_type") == "playlist":
        entries = [entry for entry in (info.get("entries") or []) if entry]
        if not entries:
            raise RuntimeError("No video found at that link.")
        info = entries[0]
    if not info.get("title"):
        raise RuntimeError("This video's metadata could not be read (private, removed or age-restricted).")
    return info


def download_thumbnail(
    video_id: str, info: dict[str, Any], dest: Path
) -> tuple[Path | None, str]:
    """Download the biggest available thumbnail to `dest`.

    Returns the local file (or None) and the URL that should be linked in the
    Markdown file.
    """
    candidates = [
        f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg",
        f"https://i.ytimg.com/vi/{video_id}/hqdefault.jpg",
    ]
    fallback = info.get("thumbnail")
    if fallback and fallback not in candidates:
        candidates.append(str(fallback))

    for candidate in candidates:
        try:
            response = requests.get(
                candidate, timeout=HTTP_TIMEOUT, headers={"User-Agent": USER_AGENT}
            )
        except requests.RequestException as exc:
            logger.warning("Thumbnail request failed (%s): %s", candidate, exc)
            continue
        # Missing maxresdefault.jpg still answers 200/404 with a tiny placeholder.
        if response.ok and len(response.content) > 10_000:
            dest.write_bytes(response.content)
            return dest, candidate
    logger.warning("No thumbnail could be downloaded for %s", video_id)
    return None, str(info.get("thumbnail") or candidates[0])


def _caption_tracks(info: dict[str, Any] | None) -> dict[str, Any]:
    """Extract subtitle/auto-caption track maps from yt-dlp info."""
    if not isinstance(info, dict):
        return {}
    return {
        "subtitles": info.get("subtitles") or {},
        "automatic_captions": info.get("automatic_captions") or {},
    }


def _transcript_from_api(video_id: str) -> str:
    """Primary extractor: youtube-transcript-api (v1.2 instance API or legacy)."""
    api = YouTubeTranscriptApi()
    if hasattr(api, "list"):  # youtube-transcript-api >= 1.2 (instance API)
        transcript_list = api.list(video_id)
    else:  # older releases expose class/static helpers
        transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)

    try:
        transcript = transcript_list.find_transcript(PREFERRED_LANGUAGES)
    except NoTranscriptFound:
        transcript = next(iter(transcript_list), None)
        if transcript is None:
            raise

    fetched = transcript.fetch()
    snippets = getattr(fetched, "snippets", fetched)  # v1.2+ object vs legacy list
    lines: list[str] = []
    previous = ""
    for snippet in snippets:
        if isinstance(snippet, dict):
            text, start = snippet.get("text", ""), snippet.get("start", 0)
        else:
            text, start = snippet.text, snippet.start
        # Captions often wrap mid-sentence; keep one transcript line per snippet.
        text = re.sub(r"\s+", " ", text or "").strip()
        if text and text != previous:
            previous = text
            lines.append(f"[{_fmt_timestamp(start)}] {text}")
    return "\n".join(lines)


def _best_caption_entry(entries: Any) -> dict[str, Any] | None:
    """Pick a json3 (preferred) or vtt caption track entry."""
    if not isinstance(entries, list):
        return None
    for wanted_ext in ("json3", "vtt"):
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            if str(entry.get("ext") or "").lower() == wanted_ext and entry.get("url"):
                return entry
    return None


def _ordered_caption_entries(captions: dict[str, Any] | None) -> list[dict[str, Any]]:
    """All usable caption tracks ordered by language priority (original first)."""
    if not captions:
        return []
    manual = captions.get("subtitles") or {}
    automatic = captions.get("automatic_captions") or {}
    ordered: list[dict[str, Any]] = []
    seen: set[str] = set()

    def push(entries: Any) -> None:
        entry = _best_caption_entry(entries)
        if entry and entry.get("url") not in seen:
            seen.add(entry["url"])
            ordered.append(entry)

    for language in TRANSCRIPT_LANGUAGE_PRIORITY:
        for tracks in (manual, automatic):
            entries = tracks.get(language) or []
            originals = [
                item
                for item in entries
                if isinstance(item, dict) and "tlang=" not in str(item.get("url") or "")
            ]
            push(originals)
            push(entries)
    for tracks in (manual, automatic):
        for language in sorted(tracks):
            if str(language).lower().startswith(("hi", "en")):
                push(tracks.get(language))
    for tracks in (manual, automatic):
        for language in sorted(tracks):
            push(tracks.get(language))
    return ordered


def _pick_caption_entry(captions: dict[str, Any] | None) -> dict[str, Any] | None:
    """Choose a caption track: hi/en/en-IN/en-US/en-GB, then any hi/en, then any."""
    entries = _ordered_caption_entries(captions)
    return entries[0] if entries else None


def _extract_info_with_clients(video_id: str) -> dict[str, Any]:
    """Targeted yt-dlp extraction using data-centre friendly player clients."""
    options = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "socket_timeout": HTTP_TIMEOUT,
        "extractor_args": {
            "youtube": {"player_client": ["android", "ios", "tv_embedded", "web"]}
        },
    }
    with yt_dlp.YoutubeDL(options) as ydl:
        info = ydl.extract_info(
            f"https://www.youtube.com/watch?v={video_id}", download=False
        )
    return info if isinstance(info, dict) else {}


def _parse_json3_captions(payload: dict[str, Any]) -> list[tuple[float, str]]:
    cues: list[tuple[float, str]] = []
    for event in payload.get("events") or []:
        if not isinstance(event, dict):
            continue
        segments = event.get("segs") or []
        text = "".join(
            str(segment.get("utf8") or "")
            for segment in segments
            if isinstance(segment, dict)
        )
        text = re.sub(r"\s+", " ", text).strip()
        if not text:
            continue
        start = float(event.get("tStartMs") or 0) / 1000.0
        cues.append((start, text))
    return cues


def _parse_vtt_time(value: str) -> float:
    parts = value.replace(",", ".").split(":")
    try:
        if len(parts) == 3:
            hours, minutes, seconds = parts
        elif len(parts) == 2:
            hours, minutes, seconds = "0", parts[0], parts[1]
        else:
            return 0.0
        return int(hours) * 3600 + int(minutes) * 60 + float(seconds)
    except (TypeError, ValueError):
        return 0.0


def _parse_vtt_captions(body: str) -> list[tuple[float, str]]:
    cues: list[tuple[float, str]] = []
    lines = body.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index].strip()
        if "-->" not in line:
            index += 1
            continue
        start = _parse_vtt_time(line.split("-->")[0].strip())
        index += 1
        parts: list[str] = []
        while index < len(lines) and lines[index].strip():
            text = re.sub(r"<[^>]+>", "", lines[index]).strip()
            if text:
                parts.append(text)
            index += 1
        text = re.sub(r"\s+", " ", " ".join(parts)).strip()
        if text:
            cues.append((start, text))
    return cues


def _format_transcript_cues(cues: list[tuple[float, str]]) -> str:
    lines: list[str] = []
    previous = ""
    for start, text in cues:
        text = re.sub(r"\s+", " ", text).strip()
        if not text or text == previous:
            continue
        previous = text
        lines.append(f"[{_fmt_timestamp(start)}] {text}")
    return "\n".join(lines)


def _download_caption_track(url: str, attempts: int = 3) -> str:
    """Download a YouTube timedtext track, retrying on 403/429 throttling."""
    headers = {
        "User-Agent": USER_AGENT,
        "Referer": "https://www.youtube.com/",
        "Origin": "https://www.youtube.com",
        "Accept-Language": "en-US,en;q=0.9",
    }
    for attempt in range(1, attempts + 1):
        try:
            response = requests.get(url, timeout=HTTP_TIMEOUT, headers=headers)
            if response.ok and response.text.strip():
                return response.text
            if response.status_code in (403, 429):
                logger.info(
                    "Caption track throttled (%s), attempt %d/%d",
                    response.status_code,
                    attempt,
                    attempts,
                )
        except requests.RequestException as exc:
            logger.warning("Caption track request failed: %s", exc)
        if attempt < attempts:
            time.sleep(1.5 * attempt)
    return ""


def _parse_caption_body(body: str, ext: str = "") -> str:
    if not body.strip():
        return ""
    if ext.lower() == "json3" or body.lstrip().startswith("{"):
        try:
            payload = json.loads(body)
        except ValueError:
            return ""
        return _format_transcript_cues(_parse_json3_captions(payload))
    return _format_transcript_cues(_parse_vtt_captions(body))


def _subtitles_via_ytdlp(video_id: str) -> str:
    """Last resort: let yt-dlp download the caption files itself, then parse them."""
    with tempfile.TemporaryDirectory(prefix="yt_subs_") as temp_dir:
        options = {
            "quiet": True,
            "no_warnings": True,
            "skip_download": True,
            "writesubtitles": True,
            "writeautomaticsub": True,
            "subtitleslangs": [*TRANSCRIPT_LANGUAGE_PRIORITY, "en-orig"],
            "subtitlesformat": "json3/vtt",
            "outtmpl": str(Path(temp_dir) / "%(id)s.%(ext)s"),
            "socket_timeout": HTTP_TIMEOUT,
            "extractor_args": {
                "youtube": {"player_client": ["android", "ios", "tv_embedded", "web"]}
            },
        }
        try:
            with yt_dlp.YoutubeDL(options) as ydl:
                ydl.download([f"https://www.youtube.com/watch?v={video_id}"])
        except Exception as exc:
            logger.warning("yt-dlp subtitle download failed for %s: %s", video_id, exc)
            return ""
        files = sorted(
            Path(temp_dir).glob("*"),
            key=lambda path: _language_rank(path.name),
        )
        for path in files:
            try:
                body = path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            text = _parse_caption_body(body, path.suffix.lstrip("."))
            if text.strip():
                return text
    return ""


def _language_rank(filename: str) -> int:
    lowered = filename.lower()
    for rank, language in enumerate(TRANSCRIPT_LANGUAGE_PRIORITY):
        if f".{language.lower()}." in lowered:
            return rank
    return len(TRANSCRIPT_LANGUAGE_PRIORITY) + 1


def _transcript_from_ytdlp(
    video_id: str, captions: dict[str, Any] | None = None
) -> str:
    """Fallback: download yt-dlp caption tracks (json3/vtt) and parse them."""
    entries = _ordered_caption_entries(captions)
    if not entries:
        try:
            entries = _ordered_caption_entries(
                _caption_tracks(_extract_info_with_clients(video_id))
            )
        except Exception as exc:
            logger.warning("Targeted caption fetch failed for %s: %s", video_id, exc)
            entries = []

    for entry in entries[:6]:
        body = _download_caption_track(str(entry.get("url") or ""))
        text = _parse_caption_body(body, str(entry.get("ext") or ""))
        if text.strip():
            return text

    # Everything above can be throttled from datacenter IPs: let yt-dlp fetch it.
    return _subtitles_via_ytdlp(video_id)


def fetch_transcript(video_id: str, captions: dict[str, Any] | None = None) -> str:
    """Return the full timestamped transcript.

    Primary: youtube-transcript-api. Fallback for datacenter IPs (where YouTube
    blocks it): parse the yt-dlp caption tracks (json3/vtt), then run a targeted
    yt-dlp extraction with android/ios/tv_embedded/web clients.
    """
    try:
        text = _transcript_from_api(video_id)
        if text.strip():
            return text
        logger.warning(
            "youtube-transcript-api returned no text for %s; using yt-dlp captions",
            video_id,
        )
    except Exception as exc:
        logger.warning(
            "youtube-transcript-api failed for %s (%s); using yt-dlp captions",
            video_id,
            exc,
        )
    return _transcript_from_ytdlp(video_id, captions)


def fetch_comments(video_url: str, limit: int | None = None) -> list[dict[str, Any]]:
    """Fetch YouTube comments sorted by popularity.

    ``limit=None`` uses the global MAX_COMMENTS setting; ``0`` means unlimited.
    """
    cap = MAX_COMMENTS if limit is None else limit
    downloader = YoutubeCommentDownloader()
    comments: list[dict[str, Any]] = []
    for comment in downloader.get_comments_from_url(video_url, sort_by=SORT_BY_POPULAR):
        comments.append(comment)
        if cap and len(comments) >= cap:
            break
    return comments


def format_comments(comments: list[dict[str, Any]]) -> str:
    blocks: list[str] = []
    for index, comment in enumerate(comments, start=1):
        author = (comment.get("author") or "Unknown").strip()
        relative_time = (comment.get("time") or "").strip()
        votes = (comment.get("votes") or "0").strip()
        details = [part for part in (relative_time, f"👍 {votes}") if part]
        if comment.get("reply"):
            details.append("↳ reply")
        if comment.get("heart"):
            details.append("❤️ creator heart")
        text = (comment.get("text") or "").strip() or "_empty comment_"
        quoted = "\n".join(f"   > {line}" for line in text.splitlines())
        blocks.append(f"{index}. **{author}** — {' · '.join(details)}\n{quoted}")
    return "\n\n".join(blocks)


def build_markdown(
    info: dict[str, Any],
    video_id: str,
    thumbnail_url: str,
    transcript: str,
    comments: list[dict[str, Any]],
) -> str:
    url = f"https://www.youtube.com/watch?v={video_id}"
    title = info.get("title") or "YouTube Video"
    channel = info.get("channel") or info.get("uploader") or "Unknown channel"
    channel_url = info.get("channel_url") or info.get("uploader_url") or ""
    channel_md = f"[{channel}]({channel_url})" if channel_url else channel
    categories = ", ".join(info.get("categories") or []) or "N/A"
    tags = ", ".join(info.get("tags") or []) or "N/A"
    description = (info.get("description") or "").strip() or "_No description provided._"
    transcript_md = transcript.strip() or "_No transcript available for this video._"

    lines = [
        f"# {title}",
        "",
        f"[![Thumbnail]({thumbnail_url})]({url})",
        "",
        "## 📺 Video Details",
        "",
        f"- **Video ID:** `{video_id}`",
        f"- **URL:** {url}",
        f"- **Channel:** {channel_md}",
        f"- **Upload Date:** {_fmt_upload_date(info)}",
        f"- **Views:** {_fmt_int(info.get('view_count'))}",
        f"- **Likes:** {_fmt_int(info.get('like_count'))}",
        f"- **Duration:** {_fmt_duration(info.get('duration'))}",
        f"- **YouTube Comment Count:** {_fmt_int(info.get('comment_count'))}",
        f"- **Comments Fetched:** {len(comments):,}",
        f"- **Categories:** {categories}",
        f"- **Tags:** {tags}",
        "",
        "## 📝 Description",
        "",
        description,
        "",
        "## 🗣️ Transcript",
        "",
        transcript_md,
        "",
        f"## 💬 Comments ({len(comments):,} fetched)",
        "",
        format_comments(comments) if comments else "_No comments could be fetched._",
        "",
    ]
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Channel intelligence & SEO audit (all blocking; runs in a worker thread)
# --------------------------------------------------------------------------- #
CHANNEL_URL_RE = re.compile(
    r"(?:https?://)?(?:www\.|m\.)?youtube\.com/"
    r"(?:@[A-Za-z0-9_.\-]{2,}"
    r"|channel/UC[A-Za-z0-9_\-]{22}"
    r"|c/[A-Za-z0-9_.\-]{2,}"
    r"|user/[A-Za-z0-9_.\-]{2,})",
    re.IGNORECASE,
)
HANDLE_RE = re.compile(r"(?<![A-Za-z0-9._%+\-])@([A-Za-z0-9_.\-]{2,30})")
FEAT_RE = re.compile(r"(?i)\b(?:ft|feat|featuring)\b\.?[\s:]*([A-Za-z0-9_@.\-]{2,40})?")
COLLAB_WORD_RE = re.compile(r"(?i)\b(?:collab(?:oration)?s?|collaborated|collaborating)\b")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
CHANNEL_TAB_SUFFIXES = (
    "/videos",
    "/shorts",
    "/posts",
    "/about",
    "/featured",
    "/streams",
    "/playlists",
    "/community",
    "/channels",
)

CHANNEL_LIST_OPTS = {
    "quiet": True,
    "no_warnings": True,
    "skip_download": True,
    "extract_flat": "in_playlist",
    "socket_timeout": HTTP_TIMEOUT,
    "retries": 2,
}

FULL_ITEM_OPTS = {
    "quiet": True,
    "no_warnings": True,
    "skip_download": True,
    "noplaylist": True,
    "socket_timeout": HTTP_TIMEOUT,
    "retries": 2,
}


@dataclass
class ChannelLink:
    title: str
    url: str


@dataclass
class ChannelProfile:
    name: str
    handle: str
    url: str
    channel_id: str
    avatar_url: str
    subscribers: int | None
    subscribers_text: str
    total_views: int | None
    total_views_text: str
    video_count: int | None
    video_count_text: str
    joined_date: str
    country: str
    description: str
    links: list[ChannelLink]
    email: str
    keywords: list[str]


@dataclass
class ChannelItem:
    """One long-form video or one Short."""

    kind: str  # "Video" or "Short"
    title: str
    url: str
    upload_date: str
    views: int | None
    likes: int | None
    comment_count: int | None
    duration_sec: int | None
    category: str
    tags: list[str]
    collaborations: str
    description: str
    # Filled by the cloud batch runner (channel mode keeps them empty locally).
    transcript: str = ""
    comments: list[dict[str, Any]] = field(default_factory=list)
    captions: dict[str, Any] = field(default_factory=dict)
    thumbnail_url: str = ""
    thumbnail_file: str = ""
    ai: dict[str, Any] = field(default_factory=dict)


@dataclass
class CommunityPost:
    post_id: str
    published_time: str
    text: str
    likes: int | None
    comment_count: int | None
    attachment_type: str
    collaborations: str
    url: str


@dataclass
class ChannelAuditResult:
    profile: ChannelProfile
    videos: list[ChannelItem]
    shorts: list[ChannelItem]
    posts: list[CommunityPost]
    collaborators: list[str]
    avatar_path: Path | None
    xlsx_path: Path
    markdown_path: Path


def extract_channel_url(text: str) -> str | None:
    """Return the channel base URL (no /videos-style suffix) from a message."""
    if not text:
        return None
    match = CHANNEL_URL_RE.search(text)
    if not match:
        return None
    url = match.group(0)
    if not re.match(r"^https?://", url, re.IGNORECASE):
        url = "https://" + url
    url = url.rstrip("/")
    for suffix in CHANNEL_TAB_SUFFIXES:
        if url.lower().endswith(suffix):
            return url[: -len(suffix)]
    return url


# ---- parsing helpers ------------------------------------------------------ #
def _text_of(node: Any) -> str:
    """Flatten YouTube's simpleText / runs / content text nodes."""
    if isinstance(node, str):
        return node
    if not isinstance(node, dict):
        return ""
    simple = node.get("simpleText")
    if isinstance(simple, str):
        return simple
    runs = node.get("runs")
    if isinstance(runs, list):
        return "".join(str(run.get("text", "")) for run in runs if isinstance(run, dict))
    content = node.get("content")
    return content if isinstance(content, str) else ""


def _walk_values(node: Any, key: str) -> list[Any]:
    """Collect every value stored under `key` anywhere in a nested JSON blob."""
    found: list[Any] = []
    stack: list[Any] = [node]
    while stack:
        current = stack.pop()
        if isinstance(current, dict):
            for name, value in current.items():
                if name == key:
                    found.append(value)
                stack.append(value)
        elif isinstance(current, list):
            stack.extend(current)
    return found


def _balanced_json(text: str, start: int) -> str | None:
    """Return the balanced JSON object starting at `start`, honouring strings."""
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
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start:index + 1]
    return None


def _initial_data_from_html(page: str, marker: str = "ytInitialData") -> dict[str, Any]:
    index = page.find(marker)
    while index != -1:
        brace = page.find("{", index)
        if brace != -1:
            raw = _balanced_json(page, brace)
            if raw:
                try:
                    return json.loads(raw)
                except json.JSONDecodeError:
                    pass
        index = page.find(marker, index + len(marker))
    raise RuntimeError("Could not read YouTube page data (CAPTCHA or a YouTube layout change).")


def _fetch_youtube_page(url: str) -> str:
    session = requests.Session()
    session.headers.update(
        {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
    )
    session.cookies.set("CONSENT", "YES+cb", domain=".youtube.com")
    response = session.get(url, timeout=HTTP_TIMEOUT)
    response.raise_for_status()
    return response.text


def _parse_compact_number(value: Any) -> int | None:
    """Parse '1.2K', '198,068' or '27K comments' into an int."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*([KMB]?)", str(value).replace(",", ""), re.I)
    if not match:
        return None
    try:
        number = float(match.group(1))
    except ValueError:
        return None
    multiplier = {"": 1, "K": 1_000, "M": 1_000_000, "B": 1_000_000_000}[
        match.group(2).upper()
    ]
    return int(round(number * multiplier))


def _as_int(value: Any) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _average(values: list[int | None]) -> float | None:
    numbers = [value for value in values if value is not None]
    return sum(numbers) / len(numbers) if numbers else None


def _iso_upload_date(info: dict[str, Any]) -> str:
    raw = str(info.get("upload_date") or "")
    if len(raw) == 8 and raw.isdigit():
        return f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"
    timestamp = info.get("timestamp")
    if timestamp:
        try:
            moment = datetime.fromtimestamp(int(timestamp), tz=timezone.utc)
            return moment.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OSError, OverflowError):
            pass
    return "N/A"


def _find_email(*texts: str) -> str:
    for text in texts:
        if not text:
            continue
        match = EMAIL_RE.search(text)
        if match:
            return match.group(0)
    return ""


def _split_keywords(raw: Any) -> list[str]:
    if isinstance(raw, (list, tuple)):
        return [str(item).strip().strip('"') for item in raw if str(item).strip()]
    text = str(raw or "").strip()
    if not text:
        return []
    keywords = [part.strip().strip('"') for part in re.split(r"[\s,]+", text)]
    return [keyword for keyword in keywords if keyword]


# ---- collaboration detection ---------------------------------------------- #
def _normalise_people(value: Any) -> list[str]:
    people: list[str] = []
    if isinstance(value, str):
        people.append(value)
    elif isinstance(value, (list, tuple)):
        for item in value:
            if isinstance(item, dict):
                people.append(str(item.get("name") or item.get("title") or item.get("id") or ""))
            else:
                people.append(str(item))
    return [person.strip() for person in people if person and person.strip()]


def detect_collaborations(
    info: dict[str, Any] | None,
    *,
    title: str = "",
    description: str = "",
    own_handles: tuple[str, ...] = (),
) -> str:
    """Detect collaborators from yt-dlp metadata plus @handles / ft. / collab text."""
    candidates: list[str] = []

    def add(value: str) -> None:
        value = value.strip().strip(",").strip()
        lowered = {candidate.lower() for candidate in candidates}
        if value and len(value) <= 80 and value.lower() not in lowered:
            candidates.append(value)

    for field in ("creators", "cast", "artists", "collaborators"):
        for person in _normalise_people((info or {}).get(field)):
            add(person)

    text = " ".join(part for part in (title, description) if part)
    own = {handle.lower().strip("@") for handle in own_handles if handle}
    for handle in HANDLE_RE.findall(text):
        handle = handle.rstrip(".")
        if handle.lower() not in own:
            add("@" + handle)
    for match in FEAT_RE.finditer(text):
        person = (match.group(1) or "").strip().rstrip(".")
        if person and not person.startswith("@") and person.lower() not in own:
            add("feat. " + person)
    if not candidates and COLLAB_WORD_RE.search(text):
        candidates.append("collab mention")
    return ", ".join(candidates[:10])


COLLAB_CHANNEL_ID_RE = re.compile(r"@(UC[A-Za-z0-9_\-]{22})")
MAX_COLLABORATOR_LOOKUPS = 25


def _channel_label(channel_id: str, ydl: Any) -> str:
    """Resolve a channel id to 'Name (@handle)' using one cheap yt-dlp lookup."""
    try:
        info = ydl.extract_info(
            f"https://www.youtube.com/channel/{channel_id}", download=False
        )
    except Exception as exc:
        logger.warning("Could not resolve collaborator channel %s: %s", channel_id, exc)
        return ""
    if not isinstance(info, dict):
        return ""
    name = str(info.get("channel") or info.get("title") or "").strip()
    handle = str(info.get("uploader_id") or "").strip()
    if name and handle.startswith("@"):
        return f"{name} ({handle})"
    return name


def resolve_collaborator_ids(
    items: list[ChannelItem],
    posts: list[CommunityPost],
    *,
    own_channel_id: str,
) -> None:
    """Replace '@UC...' multi-creator attribution links with readable channel names."""
    pending: list[str] = []
    for blob in [item.collaborations for item in items] + [
        post.collaborations for post in posts
    ]:
        for channel_id in COLLAB_CHANNEL_ID_RE.findall(blob):
            if channel_id != own_channel_id and channel_id not in pending:
                pending.append(channel_id)
    if not pending:
        return

    labels: dict[str, str] = {}
    options = dict(CHANNEL_LIST_OPTS)
    options["playlistend"] = 1
    with yt_dlp.YoutubeDL(options) as ydl:
        for channel_id in pending[:MAX_COLLABORATOR_LOOKUPS]:
            labels[channel_id] = _channel_label(channel_id, ydl)
    logger.info("Resolved %d/%d collaborator channels", len(labels), len(pending))

    def rewrite(blob: str) -> str:
        if not blob:
            return blob
        return COLLAB_CHANNEL_ID_RE.sub(
            lambda match: labels.get(match.group(1)) or match.group(0), blob
        )

    for item in items:
        item.collaborations = rewrite(item.collaborations)
    for post in posts:
        post.collaborations = rewrite(post.collaborations)


# ---- channel profile ------------------------------------------------------ #
def _upgrade_avatar_url(url: str) -> str:
    cleaned = url.strip()
    if cleaned.startswith("//"):
        cleaned = "https:" + cleaned
    return re.sub(r"=s\d+[^/]*$", "=s900", cleaned.split("?")[0])


def _pick_avatar(metadata: dict[str, Any]) -> str:
    thumbs = (metadata.get("avatar") or {}).get("thumbnails") or []
    usable = [thumb for thumb in thumbs if isinstance(thumb, dict) and thumb.get("url")]
    if usable:
        best = max(usable, key=lambda thumb: thumb.get("width") or 0)
        return _upgrade_avatar_url(str(best["url"]))
    return ""


def _parse_channel_links(raw_links: Any) -> list[ChannelLink]:
    links: list[ChannelLink] = []
    for item in raw_links or []:
        view_model = (
            item.get("channelExternalLinkViewModel") if isinstance(item, dict) else None
        )
        if not isinstance(view_model, dict):
            continue
        title = _text_of(view_model.get("title")).strip() or "Link"
        link_node = view_model.get("link") or {}
        url = ""
        for run in link_node.get("commandRuns") or []:
            command = ((run or {}).get("onTap") or {}).get("innertubeCommand") or {}
            url = str((command.get("urlEndpoint") or {}).get("url") or "").strip()
            if not url:
                metadata = command.get("commandMetadata") or {}
                url = str((metadata.get("webCommandMetadata") or {}).get("url") or "").strip()
            if url:
                break
        if not url:
            url = str(link_node.get("content") or "").strip()
        if not url:
            continue
        if url.startswith("//"):
            url = "https:" + url
        elif not re.match(r"^https?://", url, re.IGNORECASE):
            url = "https://" + url.lstrip("/")
        # YouTube wraps description links in its own redirector; unwrap the target.
        if "youtube.com/redirect" in url:
            target = (parse_qs(urlparse(url).query).get("q") or [""])[0]
            if target:
                url = target
        links.append(ChannelLink(title=title, url=url))
    return links


def _resolve_channel_base(base_url: str, listing: dict[str, Any]) -> str:
    handle = str(listing.get("uploader_id") or "").strip()
    if handle.startswith("@"):
        return f"https://www.youtube.com/{handle}"
    channel_url = str(listing.get("channel_url") or "").strip()
    return channel_url.rstrip("/") or base_url


def build_channel_profile(
    base_url: str, listing: dict[str, Any], about_data: dict[str, Any]
) -> ChannelProfile:
    about = about_data.get("about") or {}
    metadata = about_data.get("metadata") or {}

    name = (
        _text_of(metadata.get("title")).strip()
        or str(listing.get("channel") or listing.get("title") or "YouTube Channel").strip()
    )
    handle = str(listing.get("uploader_id") or "").strip()
    if not handle.startswith("@"):
        vanity = str(metadata.get("vanityChannelUrl") or "")
        match = re.search(r"/@([^/?#]+)", vanity)
        handle = "@" + match.group(1) if match else handle
    url = str(about.get("canonicalChannelUrl") or listing.get("channel_url") or base_url).strip()
    url = re.sub(r"^http://", "https://", url, flags=re.IGNORECASE)
    channel_id = str(
        about.get("channelId") or metadata.get("externalId") or listing.get("channel_id") or ""
    ).strip()

    subscribers_text = _text_of(about.get("subscriberCountText")).strip()
    subscribers = _parse_compact_number(subscribers_text)
    follower_count = _as_int(listing.get("channel_follower_count"))
    if subscribers is None and follower_count is not None:
        subscribers = follower_count
        subscribers_text = f"{follower_count:,} subscribers"

    total_views_text = _text_of(about.get("viewCountText")).strip()
    total_views = _parse_compact_number(total_views_text)
    video_count_text = _text_of(about.get("videoCountText")).strip()
    video_count = _parse_compact_number(video_count_text)
    joined_date = re.sub(
        r"^Joined\s+", "", _text_of(about.get("joinedDateText")).strip(), flags=re.IGNORECASE
    )
    country = str(about.get("country") or "").strip()
    description = str(
        about.get("description")
        or metadata.get("description")
        or listing.get("description")
        or ""
    ).strip()
    links = _parse_channel_links(about.get("links"))
    keywords = _split_keywords(metadata.get("keywords") or listing.get("tags"))
    email = _find_email(description, " ".join(f"{link.title} {link.url}" for link in links))
    if not email:
        email = "Not publicly listed (or behind YouTube CAPTCHA)"

    return ChannelProfile(
        name=name or "YouTube Channel",
        handle=handle,
        url=url,
        channel_id=channel_id,
        avatar_url=_pick_avatar(metadata),
        subscribers=subscribers,
        subscribers_text=subscribers_text or _fmt_int(subscribers),
        total_views=total_views,
        total_views_text=total_views_text or _fmt_int(total_views),
        video_count=video_count,
        video_count_text=video_count_text or _fmt_int(video_count),
        joined_date=joined_date,
        country=country,
        description=description,
        links=links,
        email=email,
        keywords=keywords,
    )


def fetch_channel_about(channel_base_url: str) -> dict[str, Any]:
    """Scrape the channel /about page; returns {} when it cannot be read."""
    try:
        page = _fetch_youtube_page(f"{channel_base_url}/about")
        data = _initial_data_from_html(page)
    except Exception as exc:
        logger.warning("Channel About page unavailable: %s", exc)
        return {}
    about = next(iter(_walk_values(data, "aboutChannelViewModel")), {})
    metadata = next(iter(_walk_values(data, "channelMetadataRenderer")), {})
    result = {
        "about": about if isinstance(about, dict) else {},
        "metadata": metadata if isinstance(metadata, dict) else {},
    }
    banner = _banner_from_initial_data(data)
    if banner:
        result["banner_url"] = banner
    return result


# --------------------------------------------------------------------------- #
# Channel banner (cover art)
# --------------------------------------------------------------------------- #
BANNER_NODE_KEYS = ("banner", "mobileBanner", "tvBanner", "imageBannerViewModel")


def _upgrade_banner_url(url: str) -> str:
    cleaned = str(url or "").strip()
    if not cleaned:
        return ""
    if cleaned.startswith("//"):
        cleaned = "https:" + cleaned
    cleaned = cleaned.split("?")[0]
    return re.sub(r"=w\d+[^/]*$", "=w2560", cleaned)


def _collect_image_urls(node: Any, found: list[tuple[int, str]], depth: int = 0) -> None:
    if depth > 40:
        return
    if isinstance(node, dict):
        url = node.get("url")
        if isinstance(url, str) and "yt3" in url:
            found.append((int(node.get("width") or 0), url))
        for value in node.values():
            _collect_image_urls(value, found, depth + 1)
    elif isinstance(node, list):
        for item in node:
            _collect_image_urls(item, found, depth + 1)


def _banner_from_initial_data(data: dict[str, Any]) -> str:
    """Find banner/mobileBanner/tvBanner image sources in the About page data."""
    candidates: list[tuple[int, str]] = []

    def scan(node: Any, depth: int = 0) -> None:
        if depth > 50:
            return
        if isinstance(node, dict):
            for key, value in node.items():
                if key in BANNER_NODE_KEYS:
                    _collect_image_urls(value, candidates)
                scan(value, depth + 1)
        elif isinstance(node, list):
            for item in node:
                scan(item, depth + 1)

    scan(data)
    if not candidates:
        return ""
    best = max(candidates, key=lambda pair: pair[0])
    return _upgrade_banner_url(best[1])


def channel_banner_url(listing: dict[str, Any], about_data: dict[str, Any] | None = None) -> str:
    """Best channel banner URL from About data or yt-dlp channel thumbnails."""
    if isinstance(about_data, dict) and about_data.get("banner_url"):
        return str(about_data["banner_url"])
    best_url = ""
    best_width = 0
    for thumb in (listing or {}).get("thumbnails") or []:
        if not isinstance(thumb, dict):
            continue
        url = str(thumb.get("url") or "")
        width = int(thumb.get("width") or 0)
        height = int(thumb.get("height") or 0)
        is_banner = (
            "banner" in url.lower()
            or width >= 1000
            or (height and width / max(height, 1) > 2)
        )
        if is_banner and width >= best_width and url:
            best_url, best_width = url, width
    return _upgrade_banner_url(best_url)


def download_channel_banner(url: str, dest: Path) -> Path | None:
    """Download a channel banner as a real .jpg."""
    if not url:
        return None
    try:
        response = requests.get(
            url,
            timeout=HTTP_TIMEOUT,
            headers={"User-Agent": USER_AGENT, "Accept": "image/jpeg,image/png,image/*"},
        )
    except requests.RequestException as exc:
        logger.warning("Banner download failed: %s", exc)
        return None
    if not response.ok or len(response.content) < 3000:
        return None
    dest.parent.mkdir(parents=True, exist_ok=True)
    if response.content[:3] == b"\xff\xd8\xff":
        dest.write_bytes(response.content)
        return dest
    try:
        from PIL import Image

        with Image.open(io.BytesIO(response.content)) as image:
            image.convert("RGB").save(dest, "JPEG", quality=92)
        return dest
    except Exception:
        try:
            dest.write_bytes(response.content)
            return dest
        except OSError:
            return None


def download_avatar(url: str, temp_dir: Path) -> Path | None:
    """Download the channel avatar (highest resolution) as a Telegram-friendly image."""
    if not url:
        return None
    try:
        response = requests.get(
            url,
            timeout=HTTP_TIMEOUT,
            headers={"User-Agent": USER_AGENT, "Accept": "image/jpeg,image/png,image/*"},
        )
    except requests.RequestException as exc:
        logger.warning("Avatar download failed: %s", exc)
        return None
    content = response.content
    if not response.ok or len(content) < 1_000:
        logger.warning("Avatar download returned no usable image (%s)", response.status_code)
        return None
    dest = temp_dir / "channel_avatar.jpg"
    if content[:3] == b"\xff\xd8\xff":  # already JPEG
        dest.write_bytes(content)
        return dest
    try:
        from PIL import Image
    except ImportError:  # pragma: no cover - Pillow ships with most installs
        suffix = ".png" if content[:8] == b"\x89PNG\r\n\x1a\n" else ".webp"
        fallback = temp_dir / f"channel_avatar{suffix}"
        fallback.write_bytes(content)
        return fallback
    try:
        with Image.open(io.BytesIO(content)) as image:
            image.convert("RGB").save(dest, "JPEG", quality=92)
        return dest
    except Exception as exc:
        logger.warning("Could not convert avatar to JPEG: %s", exc)
        return None


# ---- channel content ------------------------------------------------------ #
def fetch_channel_listing(
    channel_base_url: str, tab: str, limit: int
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Fetch a channel tab (videos/shorts) as a flat list of entries."""
    options = dict(CHANNEL_LIST_OPTS)
    if limit:
        options["playlistend"] = limit
    with yt_dlp.YoutubeDL(options) as ydl:
        info = ydl.extract_info(f"{channel_base_url}/{tab}", download=False)
    if not isinstance(info, dict):
        raise RuntimeError(f"yt-dlp returned no data for the {tab} tab.")
    entries = [
        entry
        for entry in (info.get("entries") or [])
        if isinstance(entry, dict) and entry.get("id")
    ]
    if limit:
        entries = entries[:limit]
    return info, entries


def thumbnail_url_for(video_id: str) -> str:
    """Best-effort thumbnail URL for a video id (caller should fall back on 404)."""
    return f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg"


def build_channel_item(
    source: dict[str, Any],
    kind: str,
    *,
    own_handles: tuple[str, ...] = (),
    collaboration_source: dict[str, Any] | None = None,
    video_id: str | None = None,
    captions: dict[str, Any] | None = None,
) -> ChannelItem:
    """Build one ChannelItem row from yt-dlp info (or a flat playlist entry)."""
    resolved_id = str(video_id or source.get("id") or "")
    title = str(source.get("title") or "Untitled")
    description = str(source.get("description") or "")
    url = (
        f"https://www.youtube.com/shorts/{resolved_id}"
        if kind == "Short"
        else str(source.get("webpage_url") or f"https://www.youtube.com/watch?v={resolved_id}")
    )
    return ChannelItem(
        kind=kind,
        title=title,
        url=url,
        upload_date=_iso_upload_date(source),
        views=_as_int(source.get("view_count")),
        likes=_as_int(source.get("like_count")),
        comment_count=_as_int(source.get("comment_count")),
        duration_sec=_as_int(source.get("duration")),
        category=", ".join(source.get("categories") or []) if kind == "Video" else "",
        tags=[str(tag) for tag in (source.get("tags") or [])],
        collaborations=detect_collaborations(
            collaboration_source,
            title=title,
            description=description,
            own_handles=own_handles,
        ),
        description=description,
        captions=captions or {},
        thumbnail_url=str(source.get("thumbnail") or thumbnail_url_for(resolved_id)),
    )


def fetch_items_with_details(
    entries: list[dict[str, Any]],
    kind: str,
    *,
    own_handles: tuple[str, ...],
    progress: Callable[[str], None] | None = None,
) -> list[ChannelItem]:
    """Extract full per-item metadata (likes, comments, tags, description...)."""
    items: list[ChannelItem] = []

    def note(message: str) -> None:
        if progress is not None:
            progress(message)
    with yt_dlp.YoutubeDL(FULL_ITEM_OPTS) as ydl:
        for index, entry in enumerate(entries, start=1):
            video_id = str(entry["id"])
            info: dict[str, Any] | None = None
            try:
                info = ydl.extract_info(
                    f"https://www.youtube.com/watch?v={video_id}", download=False
                )
            except Exception as exc:
                logger.warning("Could not fetch details for %s: %s", video_id, exc)
            source = info if isinstance(info, dict) else entry
            items.append(
                build_channel_item(
                    source,
                    kind,
                    own_handles=own_handles,
                    collaboration_source=source if info else None,
                    video_id=video_id,
                    captions=_caption_tracks(source if info else None),
                )
            )
            label_title = str(source.get("title") or "Untitled")
            label = label_title if len(label_title) <= 45 else label_title[:44] + "…"
            note(f"Inspected {kind.lower()} {index}/{len(entries)}: {label}")
    return items


def _attachment_type(attachment: dict[str, Any]) -> str:
    keys = set(attachment)
    if "pollRenderer" in keys:
        return "Poll"
    if keys & {"postMultiImageRenderer", "backstageImageRenderer", "imageRenderer"}:
        return "Image"
    if keys & {"videoRenderer", "backstageVideoRenderer"}:
        return "Video"
    if "playlistRenderer" in keys:
        return "Playlist"
    return "Text"


def _parse_community_post(
    renderer: dict[str, Any], own_handles: tuple[str, ...]
) -> CommunityPost | None:
    post_id = str(renderer.get("postId") or "").strip()
    if not post_id:
        return None
    published = _text_of(renderer.get("publishedTimeText")).strip()
    text = _text_of(renderer.get("contentText")).strip()

    vote_node = renderer.get("voteCount")
    likes = _parse_compact_number(_text_of(vote_node))
    if likes is None and isinstance(vote_node, dict):
        label = str(
            ((vote_node.get("accessibility") or {}).get("accessibilityData") or {}).get("label")
            or ""
        )
        match = re.search(r"along with ([0-9,]+) other people", label)
        if match:
            likes = _parse_compact_number(match.group(1))

    buttons = (renderer.get("actionButtons") or {}).get("commentActionButtonsRenderer") or {}
    reply = (buttons.get("replyButton") or {}).get("buttonRenderer") or {}
    comment_count = _parse_compact_number(_text_of(reply.get("text")))

    attachment = renderer.get("backstageAttachment") or {}
    return CommunityPost(
        post_id=post_id,
        published_time=published,
        text=text,
        likes=likes,
        comment_count=comment_count,
        attachment_type=_attachment_type(attachment if isinstance(attachment, dict) else {}),
        collaborations=detect_collaborations(None, title=text, own_handles=own_handles),
        url=f"https://www.youtube.com/post/{post_id}",
    )


def fetch_community_posts(
    channel_base_url: str, own_handles: tuple[str, ...]
) -> list[CommunityPost]:
    """Scrape recent community posts from the /posts tab via ytInitialData."""
    try:
        page = _fetch_youtube_page(f"{channel_base_url}/posts")
        data = _initial_data_from_html(page)
    except Exception as exc:
        logger.warning("Community posts unavailable: %s", exc)
        return []

    posts: list[CommunityPost] = []
    seen: set[str] = set()
    for key in ("backstagePostRenderer", "postRenderer"):
        for renderer in _walk_values(data, key):
            if not isinstance(renderer, dict):
                continue
            post = _parse_community_post(renderer, own_handles)
            if post and post.post_id not in seen:
                seen.add(post.post_id)
                posts.append(post)
    return posts


def top_channel_tags(items: list[ChannelItem], limit: int = 25) -> list[tuple[str, int]]:
    counts: dict[str, int] = {}
    display: dict[str, str] = {}
    for item in items:
        for tag in item.tags:
            tag = str(tag).strip()
            if not tag:
                continue
            key = tag.lower()
            counts[key] = counts.get(key, 0) + 1
            display.setdefault(key, tag)
    ranked = sorted(counts.items(), key=lambda pair: (-pair[1], pair[0]))
    return [(display[key], count) for key, count in ranked[:limit]]


def collect_collaborators(items: list[ChannelItem], posts: list[CommunityPost]) -> list[str]:
    seen: set[str] = set()
    collaborators: list[str] = []
    for blob in [item.collaborations for item in items] + [
        post.collaborations for post in posts
    ]:
        for name in blob.split(", "):
            name = name.strip()
            if name and name.lower() not in seen:
                seen.add(name.lower())
                collaborators.append(name)
    return collaborators


# ---- Excel workbook ------------------------------------------------------- #
def _style_header(worksheet: Any) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(bold=True, color="FFFFFF", size=11)
    for cell in worksheet[1]:
        cell.font = font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _autofit_columns(
    worksheet: Any,
    headers: list[str],
    rows: list[list[Any]],
    caps: dict[int, int] | None = None,
    minimum: int = 10,
) -> None:
    caps = caps or {}
    for index, header in enumerate(headers, start=1):
        longest = len(str(header))
        for row in rows:
            if len(row) >= index:
                value = row[index - 1]
                longest = max(longest, len("" if value is None else str(value)))
        cap = caps.get(index, 60)
        worksheet.column_dimensions[get_column_letter(index)].width = min(
            max(minimum, longest + 2), cap
        )


def _write_table(
    worksheet: Any,
    headers: list[str],
    rows: list[list[Any]],
    *,
    freeze: bool = True,
    wrap_columns: tuple[int, ...] = (),
    caps: dict[int, int] | None = None,
) -> None:
    worksheet.append(headers)
    for row in rows:
        worksheet.append(["" if value is None else value for value in row])
    _style_header(worksheet)
    if freeze:
        worksheet.freeze_panes = "A2"
    wrapped = Alignment(wrap_text=True, vertical="top")
    plain = Alignment(vertical="top")
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = wrapped if cell.column in wrap_columns else plain
    _autofit_columns(worksheet, headers, rows, caps=caps)
    worksheet.auto_filter.ref = worksheet.dimensions


def _overview_rows(
    profile: ChannelProfile,
    items: list[ChannelItem],
    posts: list[CommunityPost],
    collaborators: list[str],
) -> list[list[Any]]:
    videos = [item for item in items if item.kind == "Video"]
    shorts = [item for item in items if item.kind == "Short"]
    average_video_views = _average([item.views for item in videos])
    average_short_views = _average([item.views for item in shorts])
    tags = top_channel_tags(items, 20)

    rows: list[list[Any]] = [
        ["Channel", "Channel Name", profile.name],
        ["Channel", "Handle", profile.handle or "N/A"],
        ["Channel", "Canonical URL", profile.url],
        ["Channel", "Channel ID", profile.channel_id or "N/A"],
        ["Channel", "Subscribers", profile.subscribers_text or _fmt_int(profile.subscribers)],
        ["Channel", "Total Channel Views", profile.total_views_text or _fmt_int(profile.total_views)],
        ["Channel", "Total Videos (per YouTube)", profile.video_count_text or _fmt_int(profile.video_count)],
        ["Channel", "Videos Analysed", len(videos)],
        ["Channel", "Shorts Analysed", len(shorts)],
        ["Channel", "Community Posts Analysed", len(posts)],
        ["Channel", "Joined Date", profile.joined_date or "N/A"],
        ["Channel", "Country", profile.country or "N/A"],
        ["Channel", "Public Email", profile.email],
        ["Channel", "Channel Keywords", ", ".join(profile.keywords) if profile.keywords else "N/A"],
        ["Channel", "Description", profile.description or "N/A"],
    ]
    for link in profile.links:
        rows.append(["External Links", link.title, link.url])
    rows.append(
        [
            "SEO Summary",
            "Avg Views / Video",
            int(average_video_views) if average_video_views is not None else "N/A",
        ]
    )
    rows.append(
        [
            "SEO Summary",
            "Avg Views / Short",
            int(average_short_views) if average_short_views is not None else "N/A",
        ]
    )
    if tags:
        for rank, (tag, count) in enumerate(tags, start=1):
            rows.append(["SEO Summary", f"Top Tag #{rank}", f"{tag} ({count} uses)"])
    else:
        rows.append(["SEO Summary", "Top Tags", "N/A"])
    if collaborators:
        rows.append(["Collaborations", "Collaborators Detected", len(collaborators)])
        for rank, name in enumerate(collaborators, start=1):
            rows.append(["Collaborations", f"Collaborator #{rank}", name])
    else:
        rows.append(["Collaborations", "Collaborators Detected", "None detected"])
    return rows


def write_channel_workbook(
    profile: ChannelProfile,
    items: list[ChannelItem],
    posts: list[CommunityPost],
    dest: Path,
    *,
    banner_file: str = "",
    banner_url: str = "",
) -> Path:
    """Write the 4-tab Channel Intelligence workbook."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for channel audits. Run: pip install openpyxl")

    videos = [item for item in items if item.kind == "Video"]
    shorts = [item for item in items if item.kind == "Short"]
    collaborators = collect_collaborators(items, posts)
    banner_rows = [
        ["Channel", "Banner URL", banner_url or "N/A"],
        ["Channel", "Banner File", banner_file or "N/A"],
    ]

    workbook = Workbook()

    overview = workbook.active
    overview.title = "Channel_Overview"
    _write_table(
        overview,
        ["Section", "Metric", "Value"],
        banner_rows + _overview_rows(profile, items, posts, collaborators),
        freeze=False,
        wrap_columns=(3,),
        caps={1: 24, 2: 34, 3: 90},
    )

    video_sheet = workbook.create_sheet("Videos")
    video_headers = [
        "Title",
        "Video_URL",
        "Upload_Date",
        "Views",
        "Likes",
        "Comment_Count",
        "Duration_Sec",
        "Category",
        "Tags",
        "Collaborations_Detected",
        "Description",
    ]
    video_rows = [
        [
            item.title,
            item.url,
            item.upload_date,
            item.views,
            item.likes,
            item.comment_count,
            item.duration_sec,
            item.category,
            ", ".join(item.tags),
            item.collaborations,
            item.description[:30000],
        ]
        for item in videos
    ]
    _write_table(
        video_sheet,
        video_headers,
        video_rows,
        wrap_columns=(1, 9, 11),
        caps={1: 60, 2: 50, 9: 50, 10: 40, 11: 80},
    )

    short_sheet = workbook.create_sheet("Shorts")
    short_headers = [
        "Title",
        "Shorts_URL",
        "Upload_Date",
        "Views",
        "Likes",
        "Comment_Count",
        "Duration_Sec",
        "Tags",
        "Collaborations_Detected",
        "Description",
    ]
    short_rows = [
        [
            item.title,
            item.url,
            item.upload_date,
            item.views,
            item.likes,
            item.comment_count,
            item.duration_sec,
            ", ".join(item.tags),
            item.collaborations,
            item.description[:30000],
        ]
        for item in shorts
    ]
    _write_table(
        short_sheet,
        short_headers,
        short_rows,
        wrap_columns=(1, 8, 10),
        caps={1: 60, 2: 50, 8: 50, 9: 40, 10: 80},
    )

    post_sheet = workbook.create_sheet("Community_Posts")
    post_headers = [
        "Post_ID",
        "Published_Time",
        "Post_Text",
        "Likes",
        "Comment_Count",
        "Attachment_Type",
        "Collaborations_Detected",
    ]
    post_rows = [
        [
            post.post_id,
            post.published_time,
            post.text[:30000],
            post.likes,
            post.comment_count,
            post.attachment_type,
            post.collaborations,
        ]
        for post in posts
    ]
    if not post_rows:
        post_rows = [["No community posts found (this channel may not have a Posts tab).", "", "", "", "", "", ""]]
    _write_table(
        post_sheet,
        post_headers,
        post_rows,
        wrap_columns=(3,),
        caps={1: 36, 2: 18, 3: 80, 7: 40},
    )

    dest.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(dest)
    return dest


# ---- Markdown report ------------------------------------------------------ #
def _md_cell(value: Any) -> str:
    text = "" if value is None else str(value)
    return text.replace("|", "\\|").replace("\n", " ").strip()


def _md_link(label: str, url: str) -> str:
    return f"[{_md_cell(label)}]({url})"


def build_channel_markdown(
    profile: ChannelProfile,
    items: list[ChannelItem],
    posts: list[CommunityPost],
    *,
    banner_file: str = "",
    banner_url: str = "",
) -> str:
    videos = [item for item in items if item.kind == "Video"]
    shorts = [item for item in items if item.kind == "Short"]
    collaborators = collect_collaborators(items, posts)
    tags = top_channel_tags(items, 25)

    lines: list[str] = [f"# {profile.name}", ""]
    if profile.avatar_url:
        lines += [f"![Channel Avatar]({profile.avatar_url})", ""]
    if banner_file:
        lines += [f"![Channel Banner](./{banner_file})", ""]
    elif banner_url:
        lines += [f"![Channel Banner]({banner_url})", ""]

    lines += [
        "## 📊 Channel Profile",
        "",
        "| Field | Value |",
        "| --- | --- |",
    ]
    profile_rows = [
        ("Handle", profile.handle or "N/A"),
        ("Channel URL", profile.url),
        ("Channel ID", profile.channel_id or "N/A"),
        ("Banner URL", banner_url or "N/A"),
        ("Banner File", banner_file or "N/A"),
        ("Subscribers", profile.subscribers_text or _fmt_int(profile.subscribers)),
        ("Total Channel Views", profile.total_views_text or _fmt_int(profile.total_views)),
        ("Total Videos (per YouTube)", profile.video_count_text or _fmt_int(profile.video_count)),
        ("Videos Analysed", len(videos)),
        ("Shorts Analysed", len(shorts)),
        ("Community Posts Analysed", len(posts)),
        ("Joined Date", profile.joined_date or "N/A"),
        ("Country", profile.country or "N/A"),
        ("Public Email", profile.email),
        ("Channel Keywords", ", ".join(profile.keywords) if profile.keywords else "N/A"),
    ]
    for field, value in profile_rows:
        lines.append(f"| {_md_cell(field)} | {_md_cell(value)} |")

    lines += ["", "### 🔗 External Links", ""]
    if profile.links:
        for link in profile.links:
            lines.append(f"- **{link.title}:** {link.url}")
    else:
        lines.append("_No external links listed._")

    lines += ["", "### 📝 Channel Description", "", profile.description or "_No description provided._", ""]

    lines += ["## 📈 SEO Analytics", "", "### Top 10 Long-Form Videos by Views", ""]
    top_videos = sorted(videos, key=lambda item: item.views if item.views is not None else -1, reverse=True)[:10]
    if top_videos:
        for rank, item in enumerate(top_videos, start=1):
            lines.append(
                f"{rank}. {_md_link(item.title, item.url)} — {_fmt_int(item.views)} views · "
                f"{_fmt_int(item.likes)} likes · {item.upload_date}"
            )
    else:
        lines.append("_No long-form videos found._")

    lines += ["", "### Top 10 Shorts by Views", ""]
    top_shorts = sorted(shorts, key=lambda item: item.views if item.views is not None else -1, reverse=True)[:10]
    if top_shorts:
        for rank, item in enumerate(top_shorts, start=1):
            lines.append(
                f"{rank}. {_md_link(item.title, item.url)} — {_fmt_int(item.views)} views · "
                f"{_fmt_int(item.likes)} likes · {item.upload_date}"
            )
    else:
        lines.append("_No Shorts found._")

    lines += ["", "### Top 25 Channel Tags / Keywords", ""]
    if tags:
        lines += ["| # | Tag | Uses |", "| --- | --- | --- |"]
        for rank, (tag, count) in enumerate(tags, start=1):
            lines.append(f"| {rank} | {_md_cell(tag)} | {count} |")
    elif profile.keywords:
        lines.append(
            "_No per-video tags are exposed by YouTube for this channel; "
            f"channel keywords: {', '.join(profile.keywords)}._"
        )
    else:
        lines.append("_No tags found on this channel's videos._")

    lines += ["", "## 🤝 Collaborations Summary", ""]
    collaboration_rows: list[list[Any]] = []
    for item in items:
        if item.collaborations:
            collaboration_rows.append(
                [
                    item.kind,
                    _md_link(item.title, item.url),
                    item.collaborations,
                    f"{_fmt_int(item.views)} views · {_fmt_int(item.likes)} likes",
                    item.upload_date,
                    item.url,
                ]
            )
    for post in posts:
        if post.collaborations:
            collaboration_rows.append(
                [
                    "Community Post",
                    _md_cell(post.text[:120]) or post.post_id,
                    post.collaborations,
                    f"{_fmt_int(post.likes)} likes",
                    post.published_time,
                    post.url,
                ]
            )
    if collaboration_rows:
        lines += [
            f"**{len(collaborators)} collaborator(s) detected across {len(collaboration_rows)} item(s).**",
            "",
            "| Item Type | Title / Post | Collaborator(s) | Views / Likes | Date | URL |",
            "| --- | --- | --- | --- | --- | --- |",
        ]
        for row in collaboration_rows:
            lines.append(
                "| " + " | ".join(_md_cell(value) for value in row[:5]) + f" | {row[5]} |"
            )
    else:
        lines.append("_No collaborations detected on this channel's recent content._")

    lines += ["", "## 💬 Recent Community Posts", ""]
    if posts:
        for index, post in enumerate(posts, start=1):
            lines.append(
                f"### Post {index} — {post.published_time or 'unknown date'} · "
                f"👍 {_fmt_int(post.likes)} · 💬 {_fmt_int(post.comment_count)} · {post.attachment_type}"
            )
            if post.text:
                for text_line in post.text.splitlines() or [post.text]:
                    lines.append(f"> {text_line}")
            else:
                lines.append("> _No text._")
            if post.collaborations:
                lines.append(f"- Collaborations: {post.collaborations}")
            lines.append(f"- URL: {post.url}")
            lines.append("")
    else:
        lines.append("_No community posts available (the channel may not have a Posts tab)._")

    lines += [
        "",
        "---",
        "",
        "## 📁 Saved Files",
        "",
        f"- Excel workbook: `{safe_stem(profile.name, 'YouTube_Channel')}_Channel_Intel.xlsx`",
        f"- Markdown report: `{safe_stem(profile.name, 'YouTube_Channel')}_Channel_Intel.md`",
        "",
    ]
    return "\n".join(lines)


def audit_channel(
    channel_url: str,
    temp_dir: Path,
    progress: Callable[[str], None] | None = None,
) -> ChannelAuditResult:
    """Run the complete Channel Intelligence & SEO Audit (blocking)."""

    def note(message: str) -> None:
        if progress is not None:
            progress(message)

    base_url = channel_url.rstrip("/")
    note("Reading the channel page and About metadata")
    try:
        listing, video_entries = fetch_channel_listing(base_url, "videos", MAX_CHANNEL_VIDEOS)
    except Exception as first_exc:
        # Shorts-only channels sometimes have no /videos tab: read the channel home
        # page for its metadata and treat the long-form list as empty.
        try:
            listing, _ = fetch_channel_listing(base_url, "", 1)
            video_entries = []
            logger.warning("No /videos tab for %s: %s", base_url, first_exc)
            note("No /videos tab found (shorts-only channel)")
        except Exception:
            raise RuntimeError(
                f"Could not load that channel: {str(first_exc)[:300]}"
            ) from first_exc
    resolved_url = _resolve_channel_base(base_url, listing)
    about_data = fetch_channel_about(resolved_url)
    profile = build_channel_profile(resolved_url, listing, about_data)
    banner_url = channel_banner_url(listing, about_data)
    own_handles = tuple(
        handle for handle in (profile.handle, profile.channel_id) if handle
    )
    note(
        f"Channel resolved: {profile.name} ({profile.handle or 'no handle'}) · "
        f"{profile.subscribers_text} · {profile.country or 'country n/a'}"
    )
    note(f"Long-form videos queued for inspection: {len(video_entries)}")
    logger.info(
        "Channel %s (%s): %d videos listed, %d shorts limit",
        profile.name,
        profile.handle,
        len(video_entries),
        MAX_CHANNEL_SHORTS,
    )

    note(f"Loading the Shorts tab (limit {MAX_CHANNEL_SHORTS or 'all'})")
    try:
        _, short_entries = fetch_channel_listing(resolved_url, "shorts", MAX_CHANNEL_SHORTS)
    except Exception as exc:
        logger.warning("Shorts tab unavailable: %s", exc)
        note("Shorts tab unavailable")
        short_entries = []

    videos = fetch_items_with_details(
        video_entries, "Video", own_handles=own_handles, progress=note
    )
    note(f"Videos scraped: {len(videos)}")
    shorts = fetch_items_with_details(
        short_entries, "Short", own_handles=own_handles, progress=note
    )
    note(f"Shorts scraped: {len(shorts)}")
    note("Scraping community posts")
    posts = fetch_community_posts(resolved_url, own_handles)
    note(f"Community posts found: {len(posts)}")
    note("Downloading the channel avatar and banner")
    avatar_path = download_avatar(profile.avatar_url, temp_dir)
    note("Resolving collaborator channel names")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    channel_stem = safe_stem(profile.name, "YouTube_Channel")
    banner_file = ""
    banner_path = download_channel_banner(
        banner_url, OUTPUT_DIR / f"{channel_stem}_banner.jpg"
    )
    if banner_path is not None:
        banner_file = banner_path.name
        note(f"Banner saved: {banner_file}")
    stem = f"{channel_stem}_Channel_Intel"
    xlsx_path = OUTPUT_DIR / f"{stem}.xlsx"
    markdown_path = OUTPUT_DIR / f"{stem}.md"
    items = videos + shorts
    resolve_collaborator_ids(items, posts, own_channel_id=profile.channel_id)
    note("Writing the Excel workbook and Markdown report")
    write_channel_workbook(
        profile,
        items,
        posts,
        xlsx_path,
        banner_file=banner_file,
        banner_url=banner_url,
    )
    markdown_path.write_text(
        build_channel_markdown(
            profile,
            items,
            posts,
            banner_file=banner_file,
            banner_url=banner_url,
        ),
        encoding="utf-8",
    )
    return ChannelAuditResult(
        profile=profile,
        videos=videos,
        shorts=shorts,
        posts=posts,
        collaborators=collect_collaborators(items, posts),
        avatar_path=avatar_path,
        xlsx_path=xlsx_path,
        markdown_path=markdown_path,
    )


@dataclass
class ScrapeResult:
    video_id: str
    title: str
    channel: str
    upload_date: str
    views: str
    likes: str
    comments: list[dict[str, Any]]
    thumbnail_path: Path | None
    thumbnail_url: str
    markdown_path: Path


def scrape_video(
    video_id: str,
    temp_dir: Path,
    progress: Callable[[str], None] | None = None,
) -> ScrapeResult:
    """Blocking scrape of one video; the .md is written permanently into OUTPUT_DIR."""

    def note(message: str) -> None:
        if progress is not None:
            progress(message)

    url = f"https://www.youtube.com/watch?v={video_id}"
    note("Fetching video metadata with yt-dlp")
    info = fetch_metadata(url)
    title = str(info.get("title") or "YouTube Video")
    note(f"Metadata ready: {title[:60]} ({_fmt_duration(info.get('duration'))} · {_fmt_int(info.get('view_count'))} views)")
    note("Downloading the highest-resolution thumbnail")
    thumbnail_path, thumbnail_url = download_thumbnail(
        video_id, info, temp_dir / f"{video_id}.jpg"
    )

    try:
        note("Fetching transcript (hi / en / en-IN / en-US / en-GB, yt-dlp fallback)")
        transcript = fetch_transcript(video_id, _caption_tracks(info))
        note(f"Transcript lines: {len([line for line in transcript.splitlines() if line])}")
    except Exception as exc:  # transcripts are best-effort, never fatal
        logger.warning("Transcript unavailable for %s: %s", video_id, exc)
        note(f"Transcript unavailable ({str(exc)[:80]})")
        transcript = ""

    try:
        note("Fetching comments (popular first, replies included)")
        comments = fetch_comments(url)
        note(f"Comments fetched: {len(comments):,}")
    except Exception as exc:  # comments are best-effort, never fatal
        logger.warning("Comments unavailable for %s: %s", video_id, exc)
        note(f"Comments unavailable ({str(exc)[:80]})")
        comments = []

    note("Building the Markdown dossier")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    markdown_path = OUTPUT_DIR / safe_filename(title, video_id)
    markdown_path.write_text(
        build_markdown(info, video_id, thumbnail_url, transcript, comments),
        encoding="utf-8",
    )

    return ScrapeResult(
        video_id=video_id,
        title=title,
        channel=str(info.get("channel") or info.get("uploader") or "Unknown channel"),
        upload_date=_fmt_upload_date(info),
        views=_fmt_int(info.get("view_count")),
        likes=_fmt_int(info.get("like_count")),
        comments=comments,
        thumbnail_path=thumbnail_path,
        thumbnail_url=thumbnail_url,
        markdown_path=markdown_path,
    )


# --------------------------------------------------------------------------- #
# Telegram handlers
# --------------------------------------------------------------------------- #
def build_caption(result: ScrapeResult) -> str:
    title = result.title if len(result.title) <= 300 else result.title[:297] + "…"
    return (
        f"<b>{html.escape(title)}</b>\n"
        f"👤 {html.escape(result.channel)}\n"
        f"📅 {html.escape(result.upload_date)}\n"
        f"👁 {html.escape(result.views)} views · 👍 {html.escape(result.likes)} likes\n"
        f"💬 {len(result.comments):,} comments fetched"
    )[:MAX_CAPTION]


LIVE_EDIT_INTERVAL = 3.0  # never edit the live card more than once every 3 seconds


class LiveStatus:
    """Throttled in-place updater for the expandable thinking card."""

    def __init__(self, message: Message, card: ProgressCard):
        self.message = message
        self.card = card
        self._last_edit = 0.0

    async def edit(self, force: bool = False) -> None:
        if not force and time.monotonic() - self._last_edit < LIVE_EDIT_INTERVAL:
            return
        try:
            await self.message.edit_text(
                self.card.render(),
                parse_mode="HTML",
                disable_web_page_preview=True,
            )
            self._last_edit = time.monotonic()
        except TelegramError:
            pass  # "message is not modified", flood limits, deleted message…

    async def typing(self, action: ChatAction = ChatAction.TYPING) -> None:
        try:
            await self.message.chat.send_action(action)
        except TelegramError:
            pass

    def mark_failed(self, error: object) -> None:
        self.card.set_state(f"❌ Failed after {self.card.elapsed_text()}")
        for index, step in enumerate(self.card.steps):
            if step.status == "🔄":
                self.card.fail_step(index)
        self.card.add_trace(f"❌ {str(error)[:300]}")
        self.card.set_state(f"❌ Failed after {self.card.elapsed_text()}")


async def _run_with_live_progress(
    work: Callable[[Callable[[str], None]], Any],
    live: LiveStatus,
    *,
    action: ChatAction = ChatAction.TYPING,
    action_interval: float = 4.0,
) -> Any:
    """Run blocking ``work(progress)`` in a thread while the card updates live."""
    events: "queue.Queue[str]" = queue.Queue()

    def progress(message: str) -> None:
        events.put_nowait(str(message))

    task = asyncio.create_task(asyncio.to_thread(work, progress))
    last_action = 0.0
    while not task.done():
        await asyncio.sleep(0.4)
        while not events.empty():
            live.card.add_trace(events.get_nowait())
        await live.edit()
        if time.monotonic() - last_action >= action_interval:
            await live.typing(action)
            last_action = time.monotonic()
    while not events.empty():
        live.card.add_trace(events.get_nowait())
    await live.edit(force=True)
    return await task


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    await message.reply_text(
        "👋 <b>YT-Bot</b> — YouTube intelligence with a live 🧠 thinking card.\n\n"
        "<b>Commands</b>\n"
        "• <code>/video &lt;url&gt;</code> — video + channel DP/banner + Gemini hook &amp; "
        "audience + top 5 viral + 5 similar channels\n"
        "• <code>/channel &lt;url&gt;</code> — full channel audit (Excel + Markdown)\n"
        "• <code>/batch &lt;urls…&gt;</code> or 2+ links — cloud batch (or Render-only)\n"
        "• <code>/viral &lt;url&gt;</code> — top 5 most-viewed videos on that channel\n"
        "• <code>/similar &lt;url&gt;</code> — 5-8 peer channels in the same niche\n"
        "• <code>/mode</code> — switch engine: ⚡ Auto / 🖥️ Render / ☁️ GitHub / 💻 Local\n"
        "• <code>/help</code> — this card\n\n"
        "A single YouTube link runs the full smart video pipeline automatically; "
        "inline buttons let you extract viral videos, the full channel, or switch engine.",
        parse_mode="HTML",
        reply_markup=mode_keyboard(),
    )


HINT_TEXT = (
    "That doesn't look like a YouTube video, Shorts or channel link. Send one of these:\n"
    "• https://www.youtube.com/watch?v=dQw4w9WgXcQ\n"
    "• https://youtu.be/dQw4w9WgXcQ\n"
    "• https://www.youtube.com/shorts/dQw4w9WgXcQ\n"
    "• https://www.youtube.com/@MrBeast\n"
    "• https://www.youtube.com/channel/UCX6OQ3DkcsbYNE6H8uQQuVA\n\n"
    "Cloud batch: /batch followed by any mix of channel + video links, or paste 2+ links."
)


def extract_channel_urls(text: str) -> list[str]:
    """Every distinct channel URL in a message (order preserved, no duplicates)."""
    urls: list[str] = []
    seen: set[str] = set()
    for match in CHANNEL_URL_RE.finditer(text or ""):
        base = extract_channel_url(match.group(0))
        if not base:
            continue
        key = base.lower().rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        urls.append(base)
    return urls


def extract_batch_links(text: str) -> tuple[list[str], list[str]]:
    """Split a message into (channel URLs, standalone video URLs), deduplicated."""
    channels = extract_channel_urls(text)
    videos: list[str] = []
    seen_ids: set[str] = set()
    for token in (text or "").split():
        video_id = extract_video_id(token)
        if not video_id or video_id in seen_ids:
            continue
        seen_ids.add(video_id)
        videos.append(f"https://www.youtube.com/watch?v={video_id}")
    return channels, videos


def channel_label(url: str) -> str:
    match = re.search(r"/@([^/?#]+)", url)
    if match:
        return "@" + match.group(1)
    match = re.search(r"/channel/([A-Za-z0-9_-]{22})", url)
    if match:
        return "channel " + match.group(1)[:8] + "…"
    match = re.search(r"/(?:c|user)/([^/?#]+)", url)
    if match:
        return match.group(1)
    return url


async def handle_youtube_link(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None or not message.text:
        return

    channels, videos = extract_batch_links(message.text)
    total = len(channels) + len(videos)
    if total >= 2:
        await _route_batch(message, channels, videos)
        return
    if videos:
        await _route_single_video(message, extract_video_id(videos[0]) or "")
        return
    if channels:
        await _route_channel(message, channels[0])
        return

    await message.reply_text(HINT_TEXT, reply_markup=mode_keyboard())


async def process_video(message: Message, video_id: str) -> None:
    card = ProgressCard(
        title=f"🧠 Video Intelligence — {video_id}",
        steps=[
            "Link Received",
            "Fetching metadata, thumbnail, transcript & comments",
            "Building Markdown dossier",
            "Delivering to Telegram",
        ],
    )
    card.finish_step(0)
    card.start_step(1)
    card.set_state("🔍 Scraping video…")
    status = await message.reply_text(
        card.render(), parse_mode="HTML", disable_web_page_preview=True
    )
    live = LiveStatus(status, card)

    temp_dir: Path | None = None
    result: ScrapeResult | None = None
    try:
        temp_dir = Path(tempfile.mkdtemp(prefix="yt_bot_"))
        result = await _run_with_live_progress(
            lambda progress: scrape_video(video_id, temp_dir, progress=progress), live
        )
        logger.info("Markdown saved to %s", result.markdown_path)
        card.finish_step(1)
        card.finish_step(2)
        card.start_step(3)
        card.set_state("📤 Delivering thumbnail and dossier…")
        await live.edit(force=True)

        if result.thumbnail_path is not None:
            await live.typing(ChatAction.UPLOAD_PHOTO)
            await _send_photo_safe(message, result.thumbnail_path, build_caption(result))

        document_caption = (
            f"📄 Metadata, description, transcript and {len(result.comments):,} comments.\n"
            f"💾 Saved on your PC: <code>{html.escape(str(result.markdown_path))}</code>"
        )
        await live.typing(ChatAction.UPLOAD_DOCUMENT)
        await _send_document_with_retry(message, result.markdown_path, document_caption)
        card.finish_step(3)
        card.set_state(f"✅ Completed in {card.elapsed_text()}")
        card.add_trace(f"Delivered: {result.markdown_path.name}")
        await live.edit(force=True)
    except Exception as exc:
        logger.exception("Processing failed for video %s", video_id)
        live.mark_failed(exc)
        if result is not None:
            card.add_trace(f"💾 Saved locally: {result.markdown_path.name}")
        await live.edit(force=True)
    finally:
        # The .md lives in OUTPUT_DIR permanently; only the temporary thumbnail goes.
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)


async def _send_photo_safe(
    message: Message,
    path: Path,
    caption: str,
    reply_markup: Any = None,
) -> bool:
    """Send a photo; a timeout is logged but never aborts the handler."""
    try:
        await message.reply_photo(
            photo=path,
            caption=caption[:MAX_CAPTION],
            parse_mode="HTML",
            reply_markup=reply_markup,
            read_timeout=TELEGRAM_READ_TIMEOUT,
            write_timeout=TELEGRAM_WRITE_TIMEOUT,
        )
        return True
    except TimedOut:
        logger.warning(
            "Timed out waiting for the photo upload; Telegram usually still delivers it. "
            "Continuing with the documents."
        )
        return False


async def _send_document_with_retry(
    message: Message,
    path: Path,
    caption: str,
    reply_markup: Any = None,
) -> None:
    """Send a document, retrying once after TimedOut; raises if both attempts fail."""
    for attempt in range(1, DOCUMENT_SEND_ATTEMPTS + 1):
        try:
            await message.reply_document(
                document=path,
                filename=path.name,
                caption=caption[:MAX_CAPTION],
                parse_mode="HTML",
                reply_markup=reply_markup,
                read_timeout=TELEGRAM_READ_TIMEOUT,
                write_timeout=TELEGRAM_WRITE_TIMEOUT,
            )
            return
        except TimedOut:
            if attempt == DOCUMENT_SEND_ATTEMPTS:
                raise
            logger.warning(
                "Timed out sending %s (attempt %d/%d); retrying in %.0fs…",
                path.name,
                attempt,
                DOCUMENT_SEND_ATTEMPTS,
                DOCUMENT_RETRY_DELAY,
            )
            await asyncio.sleep(DOCUMENT_RETRY_DELAY)


def build_channel_caption(result: ChannelAuditResult) -> str:
    profile = result.profile
    title = profile.name if len(profile.name) <= 300 else profile.name[:297] + "…"
    subscribers = profile.subscribers_text or _fmt_int(profile.subscribers)
    total_views = profile.total_views_text or _fmt_int(profile.total_views)
    caption = (
        f"<b>{html.escape(title)}</b>\n"
        f"📛 {html.escape(profile.handle or 'N/A')}\n"
        f"👥 {html.escape(subscribers)} · 👁 {html.escape(total_views)}\n"
        f"🎬 {len(result.videos):,} videos · 📱 {len(result.shorts):,} shorts · "
        f"💬 {len(result.posts):,} posts\n"
        f"🤝 {len(result.collaborators):,} collaborators detected"
    )
    return caption[:MAX_CAPTION]


async def process_channel(message: Message, channel_url: str) -> None:
    video_scope = f"{MAX_CHANNEL_VIDEOS:,} videos" if MAX_CHANNEL_VIDEOS else "all videos"
    short_scope = f"{MAX_CHANNEL_SHORTS:,} shorts" if MAX_CHANNEL_SHORTS else "all shorts"
    card = ProgressCard(
        title=f"🧠 Channel Intelligence — {channel_label(channel_url)}",
        steps=[
            "Link Received",
            "Scraping Channel Data (profile, videos, shorts, posts)",
            "Categorizing & Building Excel + Report",
            "Delivering to Telegram",
        ],
    )
    card.finish_step(0)
    card.start_step(1)
    card.set_state("🔍 Auditing channel…")
    card.add_trace(f"Queued {channel_url} (up to {video_scope} and {short_scope})")
    status = await message.reply_text(
        card.render(), parse_mode="HTML", disable_web_page_preview=True
    )
    live = LiveStatus(status, card)

    temp_dir: Path | None = None
    result: ChannelAuditResult | None = None
    try:
        temp_dir = Path(tempfile.mkdtemp(prefix="yt_channel_"))
        result = await _run_with_live_progress(
            lambda progress: audit_channel(channel_url, temp_dir, progress=progress), live
        )
        logger.info(
            "Channel intel saved: %s | %s", result.xlsx_path, result.markdown_path
        )
        card.finish_step(1)
        card.start_step(2)
        card.set_state("🧠 Categorizing channel niche…")
        try:
            from categorizer import build_snapshot, categorize_channel

            category = categorize_channel(
                build_snapshot(result.profile, result.videos, result.shorts, result.posts)
            )
            card.add_trace(
                f"🧠 Niche: {category.category} / {category.sub_niche} "
                f"(score {category.score}, {category.source})"
            )
            for reason in category.reasoning[:1]:
                card.add_trace(f"↳ {reason}")
        except Exception as exc:  # categorization is a bonus, never fatal
            logger.warning("Categorizer unavailable: %s", exc)
        card.finish_step(2)
        card.start_step(3)
        card.set_state("📤 Delivering avatar, workbook and report…")
        await live.edit(force=True)

        if result.avatar_path is not None:
            await live.typing(ChatAction.UPLOAD_PHOTO)
            await _send_photo_safe(message, result.avatar_path, build_channel_caption(result))

        workbook_caption = (
            f"📊 Channel intelligence workbook — {len(result.videos):,} videos, "
            f"{len(result.shorts):,} shorts, {len(result.posts):,} posts.\n"
            f"💾 Saved on your PC: <code>{html.escape(str(result.xlsx_path))}</code>"
        )
        await live.typing(ChatAction.UPLOAD_DOCUMENT)
        await _send_document_with_retry(message, result.xlsx_path, workbook_caption)

        report_caption = (
            "📄 Full channel report (profile, SEO analytics, collaborations, community posts).\n"
            f"💾 Saved on your PC: <code>{html.escape(str(result.markdown_path))}</code>"
        )
        await live.typing(ChatAction.UPLOAD_DOCUMENT)
        await _send_document_with_retry(message, result.markdown_path, report_caption)
        card.finish_step(3)
        card.set_state(f"✅ Completed in {card.elapsed_text()}")
        card.add_trace(f"Workbook: {result.xlsx_path.name}")
        card.add_trace(f"Report: {result.markdown_path.name}")
        await live.edit(force=True)
    except Exception as exc:
        logger.exception("Channel audit failed for %s", channel_url)
        live.mark_failed(exc)
        if result is not None:
            card.add_trace(f"💾 Workbook: {result.xlsx_path.name}")
            card.add_trace(f"💾 Report: {result.markdown_path.name}")
        await live.edit(force=True)
    finally:
        # The .xlsx and .md live in OUTPUT_DIR; only the temporary avatar goes.
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)


# --------------------------------------------------------------------------- #
# Cloud batch: /batch + multi-channel messages -> GitHub Actions dispatch
# --------------------------------------------------------------------------- #
CLOUD_WORKFLOW_FILE = "scrape_channels.yml"


def cloud_batch_config() -> tuple[str, str, str]:
    pat = (os.getenv("GITHUB_PAT") or "").strip()
    repo = (os.getenv("GITHUB_REPO") or "").strip()
    branch = (os.getenv("GITHUB_BRANCH") or "main").strip() or "main"
    return pat, repo, branch


def dispatch_github_batch(links: list[str], chat_id: int | str) -> tuple[bool, str]:
    """Trigger the cloud workflow via the GitHub Actions dispatch API."""
    pat, repo, branch = cloud_batch_config()
    if not pat or not repo:
        return False, (
            "Cloud batch is not configured yet. Add GITHUB_PAT and GITHUB_REPO to .env "
            "(GITHUB_PAT comes from `gh auth token`) and restart the bot."
        )
    url = (
        f"https://api.github.com/repos/{repo}/actions/workflows/"
        f"{CLOUD_WORKFLOW_FILE}/dispatches"
    )
    payload = {
        "ref": branch,
        "inputs": {
            "extra_channels": "\n".join(links),
            "chat_id": str(chat_id),
        },
    }
    headers = {
        "Authorization": f"Bearer {pat}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "YT-Bot/1.0",
    }
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=30)
    except requests.RequestException as exc:
        return False, f"GitHub API unreachable: {str(exc)[:200]}"
    if response.status_code in (201, 204):
        return True, (
            f"Workflow dispatched to {repo} ({branch}) for {len(links)} link(s)"
        )
    detail = ""
    try:
        detail = str((response.json() or {}).get("message") or "")
    except ValueError:
        detail = response.text
    return False, f"GitHub API returned {response.status_code}: {detail[:200]}"


async def _handle_cloud_batch(
    message: Message, channels: list[str], videos: list[str]
) -> None:
    links = channels + videos
    card = ProgressCard(
        title="🚀 Cloud Batch — GitHub Actions",
        steps=[
            "Message Received",
            "Validating Links",
            "Dispatching to GitHub Actions",
            "Handoff Complete",
        ],
    )
    card.finish_step(0)
    card.start_step(1)
    card.set_state("🚀 Preparing cloud batch…")
    card.add_trace(
        f"Received {len(channels)} channel link(s) and {len(videos)} standalone video(s)"
    )
    status = await message.reply_text(
        card.render(), parse_mode="HTML", disable_web_page_preview=True
    )
    live = LiveStatus(status, card)

    if not links:
        card.fail_step(1)
        card.set_state("⚠️ Nothing to dispatch")
        card.add_trace(
            "No links found. Send /batch followed by channel and/or video URLs, or paste "
            "2+ links in one message."
        )
        await live.edit(force=True)
        return

    labels = [channel_label(url) for url in channels[:6]]
    labels += [f"video {extract_video_id(url) or url}" for url in videos[:6]]
    card.add_trace("Validated: " + ", ".join(labels))
    card.finish_step(1)
    card.start_step(2)
    card.set_state("☁️ Dispatching to GitHub Actions…")
    await live.edit(force=True)

    ok, detail = await asyncio.to_thread(
        dispatch_github_batch,
        links,
        message.chat_id or telegram_chat_id(),
    )
    card.add_trace(detail)
    if ok:
        card.finish_step(2)
        card.finish_step(3)
        card.set_state(
            f"✅ Cloud batch dispatched · {len(channels)} channels + "
            f"{len(videos)} videos · {card.elapsed_text()}"
        )
        card.footer = (
            f"🚀 Cloud Batch Dispatched to GitHub Actions ({len(channels)} channels, "
            f"{len(videos)} standalone videos)! You can safely shut down your laptop now — "
            "live Thinking updates & final categorized Excel/Markdown/ZIP will arrive right "
            "here on your phone."
        )
        card.add_trace(
            "Cloud runner will post the live thinking card + master workbook + ZIP."
        )
    else:
        card.fail_step(2)
        card.set_state("❌ Cloud dispatch failed")
        card.add_trace("Check GITHUB_PAT scope (repo + workflow) and GITHUB_REPO in .env.")
    await live.edit(force=True)


async def cmd_batch(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    text = message.text or ""
    if text.lower().startswith("/batch"):
        text = text[len("/batch"):]
    channels, videos = extract_batch_links(text)
    if not channels and not videos and context.args:
        channels, videos = extract_batch_links(" ".join(context.args))
    await _route_batch(message, channels, videos)


# --------------------------------------------------------------------------- #
# Execution engine selection (/mode) + inline buttons
# --------------------------------------------------------------------------- #
def mode_keyboard() -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton(execution_mode.mode_label(mode), callback_data=f"engine:{mode}")]
        for mode in execution_mode.MODE_ORDER
    ]
    return InlineKeyboardMarkup(rows)


def reply_buttons(video_id: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "🔥 Extract Top 5 Viral Videos", callback_data=f"viral:{video_id}"
                ),
                InlineKeyboardButton(
                    "📊 Extract Full Channel", callback_data=f"chan:{video_id}"
                ),
            ],
            [InlineKeyboardButton("⚙️ Switch Engine", callback_data="engine:menu")],
        ]
    )


def _mode_card_text() -> str:
    mode = execution_mode.current_mode()
    return (
        f"⚙️ <b>Execution engine:</b> {execution_mode.mode_label(mode)}\n"
        f"{execution_mode.describe_mode(mode)}\n\n"
        "• 🖥️ <b>Render Only</b> — 100% of jobs run on this Render instance\n"
        "• ☁️ <b>GitHub Only</b> — 100% dispatched to GitHub Actions\n"
        "• 💻 <b>Local PC Only</b> — Render standby; your PC runs everything\n"
        "• ⚡ <b>Auto (Never Fail)</b> — fast jobs local, heavy /channel &amp; /batch "
        "on Actions, 5-min failover poller as safety net\n\n"
        "Pick an engine:"
    )


async def cmd_mode(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    await message.reply_text(
        _mode_card_text(), parse_mode="HTML", reply_markup=mode_keyboard()
    )


async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if query is None:
        return
    await query.answer()
    data = str(query.data or "")
    message = query.message
    if data == "engine:menu":
        await query.edit_message_text(
            _mode_card_text(), parse_mode="HTML", reply_markup=mode_keyboard()
        )
        return
    if data.startswith("engine:"):
        mode = execution_mode.normalize_mode(data.split(":", 1)[1])
        ok, detail = await asyncio.to_thread(execution_mode.sync_mode, mode)
        text = (
            f"✅ <b>Engine switched to {execution_mode.mode_label(mode)}</b>\n"
            f"{execution_mode.describe_mode(mode)}\n"
            f"GitHub sync: {'✅ ' + detail if ok else '⚠️ ' + detail}"
        )
        try:
            await query.edit_message_text(
                text, parse_mode="HTML", reply_markup=mode_keyboard()
            )
        except TelegramError:
            await message.reply_text(text, parse_mode="HTML", reply_markup=mode_keyboard())
        return
    if data.startswith("viral:"):
        video_id = data.split(":", 1)[1]
        await _send_viral_message(message, video_id)
        return
    if data.startswith("chan:"):
        video_id = data.split(":", 1)[1]
        watch_url = f"https://www.youtube.com/watch?v={video_id}"
        try:
            from recommender import resolve_channel

            channel = await asyncio.to_thread(resolve_channel, watch_url)
        except Exception as exc:
            await message.reply_text(f"❌ Could not resolve the channel: {exc}")
            return
        await _route_channel(message, channel.get("channel_url") or "")


# --------------------------------------------------------------------------- #
# Commands
# --------------------------------------------------------------------------- #
def _command_payload(message: Message, command: str) -> str:
    text = message.text or ""
    lowered = text.lower()
    if lowered.startswith(command):
        text = text[len(command):]
    return text.strip()


async def cmd_video(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    payload_text = _command_payload(message, "/video") or " ".join(context.args or [])
    channels, videos = extract_batch_links(payload_text)
    video_id = extract_video_id(videos[0]) if videos else None
    if not video_id:
        await message.reply_text("Usage: <code>/video https://youtu.be/VIDEO_ID</code>", parse_mode="HTML")
        return
    await _route_single_video(message, video_id)


async def cmd_channel_links(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    payload_text = _command_payload(message, "/channel") or " ".join(context.args or [])
    channels, _videos = extract_batch_links(payload_text)
    channel_url = channels[0] if channels else extract_channel_url(payload_text)
    if not channel_url:
        await message.reply_text(
            "Usage: <code>/channel https://www.youtube.com/@handle</code>", parse_mode="HTML"
        )
        return
    await _route_channel(message, channel_url)


async def cmd_viral(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    payload_text = _command_payload(message, "/viral") or " ".join(context.args or [])
    video_id = extract_video_id(payload_text)
    channel_url = extract_channel_url(payload_text)
    if not video_id and not channel_url:
        await message.reply_text("Usage: <code>/viral https://youtu.be/VIDEO_ID</code>", parse_mode="HTML")
        return
    await _send_viral_message(message, video_id, channel_url)


async def cmd_similar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if message is None:
        return
    payload_text = _command_payload(message, "/similar") or " ".join(context.args or [])
    video_id = extract_video_id(payload_text)
    channel_url = extract_channel_url(payload_text)
    if not video_id and not channel_url:
        await message.reply_text("Usage: <code>/similar https://youtu.be/VIDEO_ID</code>", parse_mode="HTML")
        return
    await _send_similar_message(message, video_id, channel_url)


async def _send_viral_message(
    message: Message, video_id: str | None, channel_url: str | None = None
) -> None:
    from recommender import viral_videos

    target = channel_url or f"https://www.youtube.com/watch?v={video_id}"
    status = await message.reply_text("🔥 Scanning the channel for its top 5 viral videos…")
    try:
        result = await asyncio.to_thread(viral_videos, target, 5)
    except Exception as exc:
        await status.edit_text(f"❌ Viral scan failed: {exc}")
        return
    videos = result.get("videos") or []
    if not videos:
        await status.edit_text("❌ No long-form videos found on that channel.")
        return
    lines = [f"🔥 <b>Top {len(videos)} viral videos</b> — {html.escape(str(result['channel'].get('name') or ''))}", ""]
    for index, video in enumerate(videos, 1):
        lines.append(
            f"{index}. <a href=\"{video['url']}\">{html.escape(str(video['title'])[:90])}</a> — "
            f"{_fmt_int(video.get('views'))} views"
        )
    lines += ["", "Use the buttons below or /channel to go deeper."]
    await status.edit_text(
        "\n".join(lines),
        parse_mode="HTML",
        disable_web_page_preview=True,
        reply_markup=reply_buttons(video_id) if video_id else mode_keyboard(),
    )


async def _send_similar_message(
    message: Message, video_id: str | None, channel_url: str | None = None
) -> None:
    from recommender import similar_channels

    target = channel_url or f"https://www.youtube.com/watch?v={video_id}"
    status = await message.reply_text("🔎 Finding peer channels in the same niche…")
    try:
        result = await asyncio.to_thread(similar_channels, target, 6)
    except Exception as exc:
        await status.edit_text(f"❌ Similar-channel search failed: {exc}")
        return
    peers = result.get("channels") or []
    if not peers:
        await status.edit_text("❌ No similar channels found.")
        return
    lines = [
        f"🧭 <b>Similar channels</b> (query: {html.escape(str(result.get('query') or ''))})",
        "",
    ]
    for index, peer in enumerate(peers, 1):
        link = peer.get("channel_url") or ""
        label = html.escape(str(peer.get("name") or "Unknown"))
        handle = html.escape(str(peer.get("handle") or ""))
        if link:
            lines.append(f"{index}. <a href=\"{link}\">{label}</a> {handle}")
        else:
            lines.append(f"{index}. {label} {handle}")
    await status.edit_text(
        "\n".join(lines),
        parse_mode="HTML",
        disable_web_page_preview=True,
        reply_markup=reply_buttons(video_id) if video_id else mode_keyboard(),
    )


# --------------------------------------------------------------------------- #
# Mode-aware routing
# --------------------------------------------------------------------------- #
async def _route_single_video(message: Message, video_id: str) -> None:
    if not video_id:
        await message.reply_text("Usage: <code>/video https://youtu.be/VIDEO_ID</code>", parse_mode="HTML")
        return
    if execution_mode.route_job("single") == "cloud":
        await _handle_cloud_batch(
            message, [], [f"https://www.youtube.com/watch?v={video_id}"]
        )
    else:
        await process_smart_video(message, video_id)


async def _route_channel(message: Message, channel_url: str) -> None:
    if not channel_url:
        await message.reply_text("Usage: <code>/channel https://www.youtube.com/@handle</code>", parse_mode="HTML")
        return
    if execution_mode.route_job("channel") == "cloud":
        await _handle_cloud_batch(message, [channel_url], [])
    else:
        await process_channel(message, channel_url)


async def _route_batch(message: Message, channels: list[str], videos: list[str]) -> None:
    if not channels and not videos:
        await message.reply_text(
            "Send <code>/batch</code> followed by channel/video links, or paste 2+ links.",
            parse_mode="HTML",
            reply_markup=mode_keyboard(),
        )
        return
    if execution_mode.route_job("batch") == "cloud":
        await _handle_cloud_batch(message, channels, videos)
    else:
        await _run_render_batch(message, channels, videos)


async def _run_render_batch(message: Message, channels: list[str], videos: list[str]) -> None:
    """RENDER mode: execute the whole batch on this instance."""
    status = await message.reply_text(
        "🖥️ <b>Render mode</b> — running the batch on this instance "
        "(live Thinking card follows).",
        parse_mode="HTML",
    )
    token = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip()

    def work():
        import batch_runner

        client = batch_runner.TelegramClient(
            token, str(message.chat_id or telegram_chat_id())
        )
        ui = batch_runner.LiveBatchUI(client, len(channels), len(videos))
        ui.start()
        return batch_runner.run_batch(ui, channels, videos)

    try:
        summary = await asyncio.to_thread(work)
        await status.edit_text(
            f"✅ Batch complete on Render: {len(summary.channels)} channel(s), "
            f"{len(summary.standalone)} standalone video(s), "
            f"{summary.total_videos} videos."
        )
    except Exception as exc:
        logger.exception("Render batch failed")
        await status.edit_text(f"❌ Render batch failed: {str(exc)[:300]}")


# --------------------------------------------------------------------------- #
# Smart single-video pipeline (metadata + channel DP/banner + AI + viral + similar)
# --------------------------------------------------------------------------- #
def _video_thumbnail_url(info: dict[str, Any], video_id: str) -> str:
    return str(info.get("thumbnail") or thumbnail_url_for(video_id))


def build_smart_video_bundle(
    video_id: str, progress: Callable[[str], None] | None = None
) -> dict[str, Any]:
    """Blocking smart single-video pipeline (runs in a worker thread)."""
    from recommender import resolve_channel, similar_channels, viral_videos

    def note(message: str) -> None:
        if progress is not None:
            progress(message)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    watch_url = f"https://www.youtube.com/watch?v={video_id}"
    note("Fetching video metadata with yt-dlp")
    info = fetch_metadata(watch_url)
    item = build_channel_item(
        info,
        "Video",
        own_handles=(),
        collaboration_source=info,
        video_id=video_id,
        captions=_caption_tracks(info),
    )
    video_stem = safe_stem(item.title, "Video")
    note("Downloading HD thumbnail")
    thumb_path, _thumb_url = download_thumbnail(
        video_id, info, OUTPUT_DIR / f"{video_stem}_{video_id}_thumb.jpg"
    )
    note("Fetching transcript (API → yt-dlp captions → Gemini audio backup)")
    transcript = fetch_transcript(video_id, _caption_tracks(info))
    audio_used = False
    if not transcript.strip() and ai_analyzer.available():
        note("No captions found — Gemini audio transcription backup")
        transcript = ai_analyzer.transcribe_video_audio(video_id)
        audio_used = bool(transcript)
    item.transcript = transcript
    note("Fetching all comments")
    item.comments = fetch_comments(watch_url)

    note("Parent channel overview, DP and banner")
    profile = None
    avatar_file = ""
    banner_file = ""
    channel_url = ""
    try:
        channel = resolve_channel(watch_url)
        channel_url = str(channel.get("channel_url") or "")
        if channel_url:
            listing, _ = fetch_channel_listing(channel_url, "videos", 1)
            about_data = fetch_channel_about(channel_url)
            profile = build_channel_profile(channel_url, listing, about_data)
            channel_stem = safe_stem(profile.name, "Channel")
            avatar_src = download_avatar(profile.avatar_url, OUTPUT_DIR)
            if avatar_src:
                avatar_target = OUTPUT_DIR / f"{channel_stem}_avatar.jpg"
                shutil.copyfile(avatar_src, avatar_target)
                avatar_file = avatar_target.name
            banner_url = channel_banner_url(listing, about_data)
            banner_target = download_channel_banner(
                banner_url, OUTPUT_DIR / f"{channel_stem}_banner.jpg"
            )
            if banner_target is not None:
                banner_file = banner_target.name
    except Exception as exc:
        logger.warning("Channel enrichment failed: %s", exc)

    note("Gemini Hook Analysis + Audience Pulse (no summary)")
    hook = (
        ai_analyzer.analyze_hook_and_pulse(transcript, item.comments)
        if ai_analyzer.available()
        else {"skipped": True, "reason": "no AI key"}
    )
    note("Finding the top 5 viral videos")
    try:
        viral = viral_videos(watch_url, 5)
    except Exception as exc:
        viral = {"channel": {}, "videos": [], "error": str(exc)[:160]}
    note("Finding similar channels")
    try:
        similar = similar_channels(watch_url, 6)
    except Exception as exc:
        similar = {"channels": [], "error": str(exc)[:160]}

    markdown_path = OUTPUT_DIR / f"{video_stem}_{video_id}_Intel.md"
    markdown_path.write_text(
        build_smart_video_markdown(
            item,
            profile,
            hook,
            viral,
            similar,
            thumb_path.name if thumb_path else "",
            avatar_file,
            banner_file,
            audio_used,
        ),
        encoding="utf-8",
    )
    return {
        "item": item,
        "profile": profile,
        "channel_url": channel_url,
        "hook": hook,
        "viral": viral,
        "similar": similar,
        "thumb_path": thumb_path,
        "avatar_path": OUTPUT_DIR / avatar_file if avatar_file else None,
        "banner_path": OUTPUT_DIR / banner_file if banner_file else None,
        "markdown_path": markdown_path,
        "audio_used": audio_used,
    }


def build_smart_video_markdown(
    item: ChannelItem,
    profile: ChannelProfile | None,
    hook: dict[str, Any],
    viral: dict[str, Any],
    similar: dict[str, Any],
    thumb_file: str,
    avatar_file: str,
    banner_file: str,
    audio_used: bool,
) -> str:
    lines = [f"# {item.title}", ""]
    if thumb_file:
        lines += [f"![Thumbnail](./{thumb_file})", ""]
    lines += [
        "## 📊 Video Details",
        "",
        "| Field | Value |",
        "| --- | --- |",
        f"| URL | {item.url} |",
        f"| Uploaded | {item.upload_date} |",
        f"| Views / Likes | {_fmt_int(item.views)} / {_fmt_int(item.likes)} |",
        f"| Comments | {_fmt_int(item.comment_count)} (all fetched) |",
        f"| Duration | {_fmt_duration(item.duration_sec)} |",
        f"| Category | {item.category or 'N/A'} |",
        f"| Tags | {', '.join(item.tags) if item.tags else 'N/A'} |",
        f"| Collaborations | {item.collaborations or 'None detected'} |",
        f"| Transcript source | {'Gemini audio transcription' if audio_used else 'captions'} |",
        "",
    ]
    if profile is not None:
        lines += ["## 📺 Parent Channel", ""]
        if avatar_file:
            lines += [f"![Channel DP](./{avatar_file})", ""]
        if banner_file:
            lines += [f"![Channel Banner](./{banner_file})", ""]
        lines += [
            "| Field | Value |",
            "| --- | --- |",
            f"| Name | {profile.name} |",
            f"| Handle | {profile.handle or 'N/A'} |",
            f"| URL | {profile.url} |",
            f"| Subscribers | {profile.subscribers_text} |",
            f"| Total Views | {profile.total_views_text} |",
            f"| Joined / Country | {profile.joined_date or 'N/A'} / {profile.country or 'N/A'} |",
            "",
        ]
    lines += ["## 🎯 AI Hook Analysis", ""]
    hook_data = hook.get("hook") if isinstance(hook, dict) else None
    audience = hook.get("audience") if isinstance(hook, dict) else None
    if isinstance(hook_data, dict):
        lines += [
            f"- **Hook type:** {hook_data.get('type', 'N/A')}",
            f"- **Strength:** {hook_data.get('strength', 'N/A')}/10",
            f"- **Analysis:** {hook_data.get('analysis', 'N/A')}",
        ]
    elif hook:
        lines.append(f"- _{hook.get('error') or hook.get('reason') or 'AI analysis unavailable'}_")
    lines += ["", "## 👥 Audience Pulse", ""]
    if isinstance(audience, dict):
        lines += [
            f"- **Sentiment:** {audience.get('sentiment', 'N/A')}",
            f"- **Top themes:** {', '.join(audience.get('top_themes') or []) or 'N/A'}",
            f"- **Insight:** {audience.get('notable_insight', 'N/A')}",
        ]
    else:
        lines.append("_No audience pulse available._")
    lines += ["", "## 🗣️ Full Transcript", "", item.transcript or "_Transcript unavailable._", ""]
    lines += [f"## 💬 Comments ({len(item.comments):,})", "", format_comments(item.comments) or "_No comments._", ""]
    lines += ["## 🔥 Top 5 Viral Videos (this channel)", ""]
    viral_videos_list = viral.get("videos") or []
    if viral_videos_list:
        for rank, video in enumerate(viral_videos_list, 1):
            lines.append(
                f"{rank}. [{video.get('title', '')}]({video.get('url', '')}) — "
                f"{_fmt_int(video.get('views'))} views"
            )
    else:
        lines.append(f"_No viral data ({viral.get('error', 'unavailable')})._")
    lines += ["", "## 🧭 Similar Channels", ""]
    peers = similar.get("channels") or []
    if peers:
        for rank, peer in enumerate(peers, 1):
            handle = peer.get("handle") or ""
            channel_link = peer.get("channel_url") or ""
            lines.append(f"{rank}. [{peer.get('name', 'Unknown')}]({channel_link}) {handle}")
    else:
        lines.append(f"_No similar channels ({similar.get('error', 'unavailable')})._")
    lines.append("")
    return "\n".join(lines)


async def process_smart_video(message: Message, video_id: str) -> None:
    card = ProgressCard(
        title=f"🧠 Smart Video Intelligence — {video_id}",
        steps=[
            "Link Received",
            "Video metadata, HD thumbnail, transcript & comments",
            "Parent channel (DP + banner)",
            "AI hook/pulse + viral + similar",
            "Delivering to Telegram",
        ],
    )
    card.finish_step(0)
    card.start_step(1)
    card.set_state("🔍 Running the smart video pipeline…")
    status = await message.reply_text(
        card.render(), parse_mode="HTML", disable_web_page_preview=True
    )
    live = LiveStatus(status, card)
    try:
        bundle = await _run_with_live_progress(
            lambda progress: build_smart_video_bundle(video_id, progress), live
        )
        card.finish_step(1)
        card.finish_step(2)
        card.finish_step(3)
        card.start_step(4)
        card.set_state("📤 Delivering photos, dossier and buttons…")
        await live.edit(force=True)

        item: ChannelItem = bundle["item"]
        profile = bundle.get("profile")
        caption = (
            f"<b>{html.escape(item.title[:300])}</b>\n"
            f"👁 {_fmt_int(item.views)} views · 👍 {_fmt_int(item.likes)} likes · "
            f"💬 {_fmt_int(item.comment_count)} comments\n"
            f"🗣 Transcript: {len(item.transcript):,} chars\n"
            f"📺 {html.escape(profile.name if profile else '')}"
        )
        if bundle.get("thumb_path"):
            await _send_photo_safe(
                message, bundle["thumb_path"], caption, reply_buttons(video_id)
            )
        if profile is not None:
            channel_caption = (
                f"📺 <b>{html.escape(profile.name)}</b>\n"
                f"👥 {html.escape(profile.subscribers_text)} · "
                f"👁 {html.escape(profile.total_views_text)}"
            )
            if bundle.get("avatar_path"):
                await _send_photo_safe(message, bundle["avatar_path"], channel_caption)
            if bundle.get("banner_path"):
                await _send_photo_safe(message, bundle["banner_path"], "🖼 Channel banner")

        doc_caption = (
            f"📄 Smart dossier for <b>{html.escape(item.title[:120])}</b>\n"
            f"🎯 Hook + audience pulse · 🔥 {len(bundle['viral'].get('videos') or [])} viral videos · "
            f"🧭 {len(bundle['similar'].get('channels') or [])} similar channels\n"
            f"💾 Saved: <code>{html.escape(str(bundle['markdown_path']))}</code>"
        )
        await _send_document_with_retry(
            message, bundle["markdown_path"], doc_caption, reply_buttons(video_id)
        )
        card.finish_step(4)
        card.set_state(f"✅ Completed in {card.elapsed_text()}")
        card.add_trace("Delivered: HD thumbnail, DP, banner, dossier + inline buttons")
        await live.edit(force=True)
    except Exception as exc:
        logger.exception("Smart video pipeline failed")
        live.mark_failed(exc)
        await live.edit(force=True)


# --------------------------------------------------------------------------- #
# Render.com health endpoint + plugin API
# --------------------------------------------------------------------------- #
def health_port() -> int:
    try:
        return int(os.getenv("PORT") or "10000")
    except ValueError:
        return 10000


class HealthHandler(BaseHTTPRequestHandler):
    """Health endpoint (/health) + OpenAPI plugin endpoints (/openapi.json, /api/*)."""

    server_version = "YTBotHealth/2.0"

    def _json(self, status: int, payload: dict[str, Any], send_body: bool = True) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if send_body:
            self.wfile.write(body)

    def _ok(self, send_body: bool = True) -> None:
        self._json(200, {"status": "alive"}, send_body)

    def _plugin(self, method: str) -> bool:
        try:
            import plugin_api  # lazy import avoids the bot<->plugin import cycle

            path = self.path.split("?")[0]
            if not plugin_api.is_plugin_path(path):
                return False
            length = int(self.headers.get("Content-Length") or 0)
            body = self.rfile.read(length) if length else b""
            status, payload = plugin_api.handle_request(method, path, self.headers, body)
            self._json(status, payload)
            return True
        except Exception as exc:  # never break the health server
            logger.exception("Plugin endpoint failed: %s", exc)
            self._json(500, {"error": str(exc)[:300]})
            return True

    def do_GET(self) -> None:  # noqa: N802 - http.server API
        if self.path.split("?")[0] in ("/", "/health", "/ping"):
            self._ok()
            return
        if self._plugin("GET"):
            return
        self.send_error(404)

    def do_POST(self) -> None:  # noqa: N802 - http.server API
        if self._plugin("POST"):
            return
        self.send_error(404)

    def do_HEAD(self) -> None:  # noqa: N802 - http.server API
        if self.path.split("?")[0] in ("/", "/health", "/ping"):
            self._ok(send_body=False)
        else:
            self.send_error(404)

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        logger.debug("health: " + format, *args)


def start_health_server(port: int | None = None) -> ThreadingHTTPServer:
    """Start the health endpoint in a daemon thread (Render injects $PORT)."""
    server = ThreadingHTTPServer(("0.0.0.0", port or health_port()), HealthHandler)
    thread = threading.Thread(target=server.serve_forever, name="health-server", daemon=True)
    thread.start()
    return server


def cloud_single_links_enabled() -> bool:
    """Route single links to GitHub Actions too (default: on when hosted on Render)."""
    raw = (os.getenv("CLOUD_SINGLE_LINKS") or "").strip().lower()
    if raw:
        return raw in {"1", "true", "yes", "on"}
    return bool(
        os.getenv("RENDER")
        or os.getenv("RENDER_SERVICE_ID")
        or os.getenv("RENDER_EXTERNAL_URL")
    )


def telegram_chat_id() -> str:
    return (os.getenv("TELEGRAM_CHAT_ID") or "").strip()


# --------------------------------------------------------------------------- #
# Startup
# --------------------------------------------------------------------------- #
def load_token() -> str:
    load_dotenv(ENV_PATH)
    token = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip().strip('"').strip("'")

    if not token or token == TOKEN_PLACEHOLDER:
        print(
            "\n[YT-Bot] No bot token found.\n"
            f"  1. Open: {ENV_PATH}\n"
            "  2. Replace PASTE_YOUR_BOTFATHER_TOKEN_HERE with the token from @BotFather\n"
            "  3. Save the file and run start_bot.bat again\n"
        )
        sys.exit(1)

    if not TOKEN_RE.match(token):
        print(
            "\n[YT-Bot] TELEGRAM_BOT_TOKEN looks malformed.\n"
            "  A bot token looks like 123456789:AA... (numbers, a colon, then ~35 characters).\n"
            f"  Fix it in: {ENV_PATH}\n"
        )
        sys.exit(1)

    return token


def _build_application(token: str) -> Application:
    """Build the PTB application with all commands + the LOCAL-standby guard."""

    async def post_init(application: Application) -> None:
        async def guard_loop() -> None:
            while True:
                await asyncio.sleep(30)
                try:
                    remote = await asyncio.to_thread(execution_mode.fetch_remote_mode)
                except Exception:
                    remote = None
                if remote:
                    execution_mode.save_local_mode(remote)
                active = remote or execution_mode.current_mode()
                if execution_mode.standby_required(active):
                    logger.info("Execution mode is LOCAL: stopping listener for standby")
                    await application.stop()
                    return

        asyncio.get_running_loop().create_task(guard_loop())

    application = (
        Application.builder()
        .token(token)
        .connect_timeout(TELEGRAM_CONNECT_TIMEOUT)
        .read_timeout(TELEGRAM_READ_TIMEOUT)
        .write_timeout(TELEGRAM_WRITE_TIMEOUT)
        # PTB uses a separate media timeout for uploads (photo/document files);
        # keep it in sync with write_timeout instead of the 20s default.
        .media_write_timeout(TELEGRAM_WRITE_TIMEOUT)
        .pool_timeout(TELEGRAM_POOL_TIMEOUT)
        .post_init(post_init)
        .build()
    )
    application.add_handler(CommandHandler(["start", "help"], cmd_start))
    application.add_handler(CommandHandler("mode", cmd_mode))
    application.add_handler(CommandHandler("batch", cmd_batch))
    application.add_handler(CommandHandler("video", cmd_video))
    application.add_handler(CommandHandler("channel", cmd_channel_links))
    application.add_handler(CommandHandler("viral", cmd_viral))
    application.add_handler(CommandHandler("similar", cmd_similar))
    application.add_handler(CallbackQueryHandler(on_callback))
    application.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_youtube_link)
    )
    return application


def _wait_for_mode_change() -> None:
    """Block in polite standby until /mode is switched away from LOCAL."""
    while True:
        time.sleep(30)
        try:
            remote = execution_mode.fetch_remote_mode()
        except Exception:
            remote = None
        if remote:
            execution_mode.save_local_mode(remote)
        local = execution_mode.load_local_mode()
        if local and not execution_mode.standby_required(local):
            logger.info("Execution mode changed to %s: waking up", local)
            return


def main() -> None:
    token = load_token()
    mode = execution_mode.resolve_mode()
    execution_mode.save_local_mode(mode)

    port = health_port()
    health_server: ThreadingHTTPServer | None = None
    try:
        health_server = start_health_server(port)
        print(
            f"[YT-Bot] Health server listening on 0.0.0.0:{port} "
            "(/, /health, /ping, /openapi.json, /api/*)"
        )
    except OSError as exc:  # port busy locally: polling should still work
        logger.warning("Health server could not bind port %s: %s", port, exc)

    try:
        while True:
            mode = execution_mode.current_mode()
            if execution_mode.standby_required(mode):
                print(
                    "[YT-Bot] Execution mode LOCAL: polite polling standby. "
                    "Your local PC owns the bot; waiting for /mode change…"
                )
                _wait_for_mode_change()
                continue
            print(
                f"[YT-Bot] Running · engine {execution_mode.mode_label(mode)} "
                f"({execution_mode.describe_mode(mode)}). Ctrl+C to stop."
            )
            application = _build_application(token)
            try:
                application.run_polling(allowed_updates=Update.ALL_TYPES)
            except KeyboardInterrupt:
                raise
            if not execution_mode.standby_required(execution_mode.current_mode()):
                break
    finally:
        if health_server is not None:
            health_server.shutdown()


if __name__ == "__main__":
    main()
