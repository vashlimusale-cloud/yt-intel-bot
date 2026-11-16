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
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse

import requests
import yt_dlp
from dotenv import load_dotenv
from telegram import Message, Update
from telegram.constants import ChatAction
from telegram.error import TelegramError, TimedOut
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
from youtube_comment_downloader import SORT_BY_POPULAR, YoutubeCommentDownloader
from youtube_transcript_api import NoTranscriptFound, YouTubeTranscriptApi

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


def fetch_transcript(video_id: str) -> str:
    """Return the full timestamped transcript, preferring hi / en / en-IN."""
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
    for snippet in snippets:
        if isinstance(snippet, dict):
            text, start = snippet.get("text", ""), snippet.get("start", 0)
        else:
            text, start = snippet.text, snippet.start
        # Captions often wrap mid-sentence; keep one transcript line per snippet.
        text = re.sub(r"\s+", " ", text or "").strip()
        if text:
            lines.append(f"[{_fmt_timestamp(start)}] {text}")
    return "\n".join(lines)


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
    return {
        "about": about if isinstance(about, dict) else {},
        "metadata": metadata if isinstance(metadata, dict) else {},
    }


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


def build_channel_item(
    source: dict[str, Any],
    kind: str,
    *,
    own_handles: tuple[str, ...] = (),
    collaboration_source: dict[str, Any] | None = None,
    video_id: str | None = None,
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
) -> Path:
    """Write the 4-tab Channel Intelligence workbook."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for channel audits. Run: pip install openpyxl")

    videos = [item for item in items if item.kind == "Video"]
    shorts = [item for item in items if item.kind == "Short"]
    collaborators = collect_collaborators(items, posts)

    workbook = Workbook()

    overview = workbook.active
    overview.title = "Channel_Overview"
    _write_table(
        overview,
        ["Section", "Metric", "Value"],
        _overview_rows(profile, items, posts, collaborators),
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
    profile: ChannelProfile, items: list[ChannelItem], posts: list[CommunityPost]
) -> str:
    videos = [item for item in items if item.kind == "Video"]
    shorts = [item for item in items if item.kind == "Short"]
    collaborators = collect_collaborators(items, posts)
    tags = top_channel_tags(items, 25)

    lines: list[str] = [f"# {profile.name}", ""]
    if profile.avatar_url:
        lines += [f"![Channel Avatar]({profile.avatar_url})", ""]

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
    profile = build_channel_profile(resolved_url, listing, fetch_channel_about(resolved_url))
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
    note("Downloading the channel avatar")
    avatar_path = download_avatar(profile.avatar_url, temp_dir)
    note("Resolving collaborator channel names")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stem = f"{safe_stem(profile.name, 'YouTube_Channel')}_Channel_Intel"
    xlsx_path = OUTPUT_DIR / f"{stem}.xlsx"
    markdown_path = OUTPUT_DIR / f"{stem}.md"
    items = videos + shorts
    resolve_collaborator_ids(items, posts, own_channel_id=profile.channel_id)
    note("Writing the Excel workbook and Markdown report")
    write_channel_workbook(profile, items, posts, xlsx_path)
    markdown_path.write_text(
        build_channel_markdown(profile, items, posts), encoding="utf-8"
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
        note("Fetching transcript (hi / en / en-IN preferred)")
        transcript = fetch_transcript(video_id)
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
        "👋 Send me a YouTube video, Shorts or channel link — I keep a live, "
        "expandable 🧠 progress card updated while I work.\n\n"
        "• Video / Shorts → thumbnail + Markdown file with metadata, description, "
        "transcript and every comment.\n"
        "• Channel → avatar + Excel Channel Intelligence workbook + Markdown SEO audit.\n"
        "• /batch + ANY mix of channel and video links (or paste 2+ links) → runs the "
        "whole batch in the cloud on GitHub Actions, then shut your laptop down: live "
        "Thinking updates + categorized Excel/Markdown/ZIP arrive right here.\n\n"
        "Examples:\n"
        "• https://www.youtube.com/watch?v=dQw4w9WgXcQ\n"
        "• https://youtu.be/dQw4w9WgXcQ\n"
        "• https://www.youtube.com/shorts/dQw4w9WgXcQ\n"
        "• https://www.youtube.com/@MrBeast\n"
        "• https://www.youtube.com/channel/UCX6OQ3DkcsbYNE6H8uQQuVA\n"
        "• /batch https://www.youtube.com/@MrBeast https://youtu.be/dQw4w9WgXcQ"
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
    if len(channels) + len(videos) >= 2:
        await _handle_cloud_batch(message, channels, videos)
        return

    video_id = extract_video_id(videos[0]) if videos else None
    if video_id:
        await process_video(message, video_id)
        return
    if channels:
        await process_channel(message, channels[0])
        return

    await message.reply_text(HINT_TEXT)


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


async def _send_photo_safe(message: Message, path: Path, caption: str) -> bool:
    """Send a photo; a timeout is logged but never aborts the handler."""
    try:
        await message.reply_photo(
            photo=path,
            caption=caption[:MAX_CAPTION],
            parse_mode="HTML",
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


async def _send_document_with_retry(message: Message, path: Path, caption: str) -> None:
    """Send a document, retrying once after TimedOut; raises if both attempts fail."""
    for attempt in range(1, DOCUMENT_SEND_ATTEMPTS + 1):
        try:
            await message.reply_document(
                document=path,
                filename=path.name,
                caption=caption[:MAX_CAPTION],
                parse_mode="HTML",
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

    ok, detail = await asyncio.to_thread(dispatch_github_batch, links, message.chat_id)
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
    await _handle_cloud_batch(message, channels, videos)


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


def main() -> None:
    token = load_token()

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
        .build()
    )
    application.add_handler(CommandHandler(["start", "help"], cmd_start))
    application.add_handler(CommandHandler("batch", cmd_batch))
    application.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_youtube_link)
    )

    print(
        "[YT-Bot] Running. Send a YouTube link, or /batch + channel links for the "
        "cloud batch. Ctrl+C to stop."
    )
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
