"""Cloud batch runner for YT-Bot: mixed channels + standalone videos.

Tailored for the real workload (5-7 channels with ~6-70 videos each, plus
individual video links in the same batch), running unattended on GitHub
Actions:

* every channel -> profile/about, all videos (<= MAX_CHANNEL_VIDEOS), all Shorts
  (<= MAX_CHANNEL_SHORTS), community posts, collaborations, a full timestamped
  transcript and the top MAX_COMMENTS_PER_VIDEO comments per item;
* every standalone video -> metadata, thumbnail, full transcript, ALL comments;
* per-channel ``<Channel>_Full_Intel.xlsx`` / ``.md`` inside category folders,
  standalone dossiers inside ``Standalone_Videos/``;
* ``Master_Channels_Comparison.xlsx`` (overview + per-category tabs +
  ``Standalone_Videos`` + ``Top_Videos_Combined``), everything zipped as
  ``All_Channels_Categorized.zip`` and delivered to Telegram;
* ONE Telegram status message, edited in place (max 1 edit / 3 s) with the
  collapsible "Agent Thinking & Live Trace" block.

Polite pacing: 1-2 s between per-video transcript/comment calls and 6-10 s
between channels, so ~150-250 videos finish without YouTube rate limits.
"""

from __future__ import annotations

import io
import logging
import math
import os
import random
import re
import shutil
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

import requests
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import bot as yt_bot  # noqa: E402  (reuses the scraping pipeline helpers)
import categorizer  # noqa: E402
from progress import ProgressCard  # noqa: E402

logger = logging.getLogger("yt-bot.batch")

OUTPUT_ROOT = BASE_DIR / "output"
BATCH_DIR = OUTPUT_ROOT / "batch_result"
STANDALONE_DIR = BATCH_DIR / "Standalone_Videos"
MASTER_XLSX = OUTPUT_ROOT / "Master_Channels_Comparison.xlsx"
ZIP_PATH = OUTPUT_ROOT / "All_Channels_Categorized.zip"
CHANNELS_FILE = BASE_DIR / "channels.txt"

# Workload limits / pacing
MAX_CHANNEL_VIDEOS = 100
MAX_CHANNEL_SHORTS = 100
MAX_COMMENTS_PER_VIDEO = 25  # top popular comments for channel items
MAX_COMMENTS_STANDALONE = 0  # 0 = unlimited for standalone videos
VIDEO_DELAY_MIN, VIDEO_DELAY_MAX = 1.0, 2.0
CHANNEL_DELAY_MIN, CHANNEL_DELAY_MAX = 6.0, 10.0

TELEGRAM_REQUEST_TIMEOUT = 30
TELEGRAM_UPLOAD_TIMEOUT = 120
EDIT_INTERVAL = 3.0
CHANNEL_RETRIES = 2
RETRY_BACKOFF = 5.0

STEP_QUEUE = "Queue Initialized"
STEP_SCRAPE = "Scraping Channels & Standalone Videos"
STEP_MASTER = "Categorizing & Building Master Excel"
STEP_PACK = "Packaging ZIP & Delivering"

STANDALONE_FOLDER = "Standalone_Videos"


# --------------------------------------------------------------------------- #
# Telegram HTTP client
# --------------------------------------------------------------------------- #
class TelegramClient:
    def __init__(self, token: str, chat_id: str):
        self.base_url = f"https://api.telegram.org/bot{token}"
        self.chat_id = str(chat_id)
        self.message_id: int | None = None
        self._last_edit = 0.0

    def _api(self, method: str, payload: dict[str, Any], timeout: float) -> dict[str, Any]:
        response = requests.post(
            f"{self.base_url}/{method}", json=payload, timeout=timeout
        )
        try:
            data = response.json()
        except ValueError:
            data = {}
        if not response.ok or not data.get("ok"):
            raise RuntimeError(
                f"Telegram {method} failed: {response.status_code} "
                f"{data.get('description') or response.text[:200]}"
            )
        return data

    def send_message(self, html_text: str) -> int:
        data = self._api(
            "sendMessage",
            {
                "chat_id": self.chat_id,
                "text": html_text,
                "parse_mode": "HTML",
                "disable_web_page_preview": True,
            },
            TELEGRAM_REQUEST_TIMEOUT,
        )
        self.message_id = int(data["result"]["message_id"])
        self._last_edit = time.monotonic()
        return self.message_id

    def edit_message(self, html_text: str, force: bool = False) -> bool:
        if self.message_id is None:
            return False
        if not force and time.monotonic() - self._last_edit < EDIT_INTERVAL:
            return False
        try:
            data = requests.post(
                f"{self.base_url}/editMessageText",
                json={
                    "chat_id": self.chat_id,
                    "message_id": self.message_id,
                    "text": html_text,
                    "parse_mode": "HTML",
                    "disable_web_page_preview": True,
                },
                timeout=TELEGRAM_REQUEST_TIMEOUT,
            ).json()
            if data.get("ok"):
                self._last_edit = time.monotonic()
                return True
            description = str(data.get("description") or "")
            if "message is not modified" in description.lower():
                self._last_edit = time.monotonic()
                return True
            logger.warning("editMessageText rejected: %s", description)
        except Exception as exc:
            logger.warning("editMessageText failed: %s", exc)
        return False

    def send_document(self, path: Path, caption: str) -> bool:
        for attempt in range(1, 3):
            try:
                with path.open("rb") as handle:
                    response = requests.post(
                        f"{self.base_url}/sendDocument",
                        data={
                            "chat_id": self.chat_id,
                            "caption": caption[:1024],
                            "parse_mode": "HTML",
                        },
                        files={"document": (path.name, handle)},
                        timeout=TELEGRAM_UPLOAD_TIMEOUT,
                    )
                if response.ok and response.json().get("ok"):
                    return True
                logger.warning(
                    "sendDocument %s rejected (%s): %s",
                    path.name,
                    response.status_code,
                    response.text[:200],
                )
            except requests.Timeout:
                logger.warning(
                    "sendDocument %s timed out (attempt %d/2)", path.name, attempt
                )
            except requests.RequestException as exc:
                logger.warning("sendDocument %s failed: %s", path.name, exc)
            if attempt == 1:
                time.sleep(2)
        return False

    def send_photo(self, path: Path, caption: str = "") -> bool:
        """Send one avatar/thumbnail photo (120s timeout, 1 retry)."""
        for attempt in range(1, 3):
            try:
                with path.open("rb") as handle:
                    response = requests.post(
                        f"{self.base_url}/sendPhoto",
                        data={
                            "chat_id": self.chat_id,
                            "caption": caption[:1024],
                            "parse_mode": "HTML",
                        },
                        files={"photo": (path.name, handle)},
                        timeout=TELEGRAM_UPLOAD_TIMEOUT,
                    )
                if response.ok and response.json().get("ok"):
                    return True
                logger.warning(
                    "sendPhoto %s rejected (%s): %s",
                    path.name,
                    response.status_code,
                    response.text[:200],
                )
            except requests.Timeout:
                logger.warning("sendPhoto %s timed out (attempt %d/2)", path.name, attempt)
            except requests.RequestException as exc:
                logger.warning("sendPhoto %s failed: %s", path.name, exc)
            if attempt == 1:
                time.sleep(2)
        return False


# --------------------------------------------------------------------------- #
# Live progress UI
# --------------------------------------------------------------------------- #
class LiveBatchUI:
    def __init__(self, client: TelegramClient, channels: int, videos: int):
        self.client = client
        self.channels_total = channels
        self.videos_total = videos
        self.items_done = 0
        self.card = ProgressCard(
            title="🧠 YouTube Intelligence — Cloud Batch",
            steps=[STEP_QUEUE, STEP_SCRAPE, STEP_MASTER, STEP_PACK],
        )

    def start(self) -> None:
        self.card.finish_step(0)
        self.card.start_step(1, "Preparing the batch")
        self.card.set_state(
            f"🚄 Starting · {self.channels_total} channels + {self.videos_total} videos · 00:00"
        )
        self.card.add_trace(
            f"Queue initialized: {self.channels_total} channel(s) + "
            f"{self.videos_total} standalone video(s)"
        )
        self.client.send_message(self.card.render())

    def trace(self, message: object, icon: str = "•") -> None:
        self.card.add_trace(message, icon=icon)

    def update(self, force: bool = False) -> None:
        self.client.edit_message(self.card.render(), force=force)

    def tick(
        self,
        channels_done: int,
        videos_done: int,
        items_done: int | None = None,
        items_total: int | None = None,
    ) -> None:
        if items_done is not None:
            self.items_done = items_done
        items_part = (
            f" · {items_done}/{items_total} items"
            if items_total is not None
            else ""
        )
        self.card.set_state(
            f"🔄 Running · {channels_done}/{self.channels_total} channels · "
            f"{videos_done}/{self.videos_total} videos{items_part} · "
            f"{self.card.elapsed_text()} elapsed"
        )

    def fail(self, error: object) -> None:
        self.card.set_state(f"❌ Batch failed · {self.card.elapsed_text()} elapsed")
        for index, step in enumerate(self.card.steps):
            if step.status == "🔄":
                self.card.fail_step(index)
        self.card.add_trace(f"❌ {str(error)[:300]}", icon="❌")
        self.update(force=True)


# --------------------------------------------------------------------------- #
# Link queue
# --------------------------------------------------------------------------- #
def load_links() -> tuple[list[str], list[str]]:
    """Read channel + standalone video URLs from EXTRA_CHANNELS or channels.txt."""
    raw = (os.getenv("EXTRA_CHANNELS") or "").strip()
    source = "EXTRA_CHANNELS"
    if not raw and CHANNELS_FILE.exists():
        raw = CHANNELS_FILE.read_text(encoding="utf-8", errors="ignore")
        source = CHANNELS_FILE.name

    channels: list[str] = []
    videos: list[str] = []
    seen_channels: set[str] = set()
    seen_videos: set[str] = set()
    for token in re.split(r"[\s,]+", raw):
        candidate = token.strip().strip('"').strip("'")
        if not candidate or candidate.startswith("#"):
            continue
        channel_url = yt_bot.extract_channel_url(candidate)
        if channel_url:
            key = channel_url.lower().rstrip("/")
            if key not in seen_channels:
                seen_channels.add(key)
                channels.append(channel_url)
            continue
        video_id = yt_bot.extract_video_id(candidate)
        if video_id and video_id not in seen_videos:
            seen_videos.add(video_id)
            videos.append(f"https://www.youtube.com/watch?v={video_id}")

    logger.info(
        "Loaded %d channel(s) and %d standalone video(s) from %s",
        len(channels),
        len(videos),
        source,
    )
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


def category_folder(category: str) -> Path:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", category).strip().strip(".") or "General"
    return BATCH_DIR / safe


# --------------------------------------------------------------------------- #
# Image downloads (channel avatar + video/Short thumbnails, as real .jpg files)
# --------------------------------------------------------------------------- #
def _save_jpeg(content: bytes, dest: Path, minimum: int = 2000) -> Path | None:
    """Write JPEG bytes (convert other formats); reject tiny 404 placeholders."""
    if len(content) < minimum:
        return None
    dest.parent.mkdir(parents=True, exist_ok=True)
    if content[:3] == b"\xff\xd8\xff":
        dest.write_bytes(content)
        return dest
    try:
        from PIL import Image  # local import: Pillow is optional at runtime

        with Image.open(io.BytesIO(content)) as image:
            image.convert("RGB").save(dest, "JPEG", quality=92)
        return dest
    except Exception:
        try:
            dest.write_bytes(content)
            return dest
        except OSError as exc:
            logger.warning("Could not save image %s: %s", dest, exc)
            return None


def download_image(url: str, dest: Path, minimum: int = 2000) -> Path | None:
    if not url:
        return None
    try:
        response = requests.get(
            url,
            headers={
                "User-Agent": yt_bot.USER_AGENT,
                "Accept": "image/jpeg,image/png,image/*",
            },
            timeout=yt_bot.HTTP_TIMEOUT,
        )
    except requests.RequestException as exc:
        logger.warning("Image download failed (%s): %s", str(url)[:90], exc)
        return None
    if not response.ok:
        return None
    return _save_jpeg(response.content, dest, minimum=minimum)


def download_channel_avatar(profile: Any, dest: Path) -> Path | None:
    """Download the channel profile picture as a real .jpg (highest resolution)."""
    url = str(getattr(profile, "avatar_url", "") or "").strip()
    if not url:
        return None
    candidates: list[str] = []
    upgraded = yt_bot._upgrade_avatar_url(url)
    if upgraded:
        candidates.append(upgraded)
    if url not in candidates:
        candidates.append(url)
    for candidate in candidates:
        path = download_image(candidate, dest)
        if path:
            return path
    return None


def thumbnail_filename(index: int, title: str, video_id: str) -> str:
    stem = yt_bot.safe_stem(title, "Video")[:60]
    return f"{index:03d}_{stem}_{video_id}.jpg"


def download_video_thumbnail(
    video_id: str, info_thumbnail: Any, dest: Path
) -> Path | None:
    """Download the best video/Short thumbnail, rejecting <2KB 404 placeholders."""
    candidates = [
        f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg",
        f"https://i.ytimg.com/vi/{video_id}/hqdefault.jpg",
    ]
    fallback = str(info_thumbnail or "").strip()
    if fallback and fallback not in candidates:
        candidates.append(fallback)
    for candidate in candidates:
        path = download_image(candidate, dest)
        if path:
            return path
    return None


def retry_call(label: str, work: Callable[[], Any], ui: LiveBatchUI) -> tuple[Any | None, str]:
    """Run `work` with one retry; returns (result, error)."""
    last_error = ""
    for attempt in range(1, CHANNEL_RETRIES + 1):
        try:
            return work(), ""
        except Exception as exc:  # noqa: BLE001 - one bad item must not kill the batch
            last_error = str(exc)[:250]
            logger.warning("%s failed (attempt %d/%d): %s", label, attempt, CHANNEL_RETRIES, exc)
            ui.trace(f"⚠️ {label} attempt {attempt}/{CHANNEL_RETRIES} failed: {last_error}")
            if attempt < CHANNEL_RETRIES:
                ui.trace(f"↻ Retrying {label} in {int(RETRY_BACKOFF)}s")
                ui.update()
                time.sleep(RETRY_BACKOFF)
    return None, last_error


def polite_delay(ui: LiveBatchUI, minimum: float, maximum: float, next_label: str) -> None:
    """Randomised cooldown with a live countdown in the thinking trace."""
    remaining = random.uniform(minimum, maximum)
    while remaining > 0:
        ui.trace(f"😴 Polite delay: {math.ceil(remaining)}s before {next_label}")
        ui.update()
        step = min(3.0, remaining)
        time.sleep(step)
        remaining -= step


# --------------------------------------------------------------------------- #
# Data structures
# --------------------------------------------------------------------------- #
@dataclass
class FullChannel:
    url: str
    profile: Any
    videos: list[Any]
    shorts: list[Any]
    posts: list[Any]
    collaborators: list[str]
    category: categorizer.CategoryResult
    xlsx_path: Path
    md_path: Path
    avatar_path: Path | None = None
    elapsed: float = 0.0

    @property
    def items(self) -> list[Any]:
        return self.videos + self.shorts


@dataclass
class StandaloneVideo:
    video_id: str
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
    thumbnail_url: str
    transcript: str
    comments: list[dict[str, Any]]
    md_path: Path
    thumbnail_file: str = ""
    elapsed: float = 0.0


@dataclass
class BatchSummary:
    channels: list[FullChannel] = field(default_factory=list)
    standalone: list[StandaloneVideo] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)
    started_at: float = field(default_factory=time.monotonic)

    @property
    def elapsed(self) -> float:
        return time.monotonic() - self.started_at

    @property
    def total_videos(self) -> int:
        return sum(len(channel.videos) for channel in self.channels) + len(self.standalone)

    @property
    def total_shorts(self) -> int:
        return sum(len(channel.shorts) for channel in self.channels)

    @property
    def transcripts(self) -> int:
        count = sum(1 for channel in self.channels for item in channel.items if item.transcript)
        return count + sum(1 for video in self.standalone if video.transcript)

    @property
    def comments(self) -> int:
        count = sum(len(item.comments) for channel in self.channels for item in channel.items)
        return count + sum(len(video.comments) for video in self.standalone)

    @property
    def categories(self) -> list[str]:
        return sorted({channel.category.category for channel in self.channels})


# --------------------------------------------------------------------------- #
# Channel intel: metadata + transcript + top comments
# --------------------------------------------------------------------------- #
def _enrich_item(
    item: Any,
    ui: LiveBatchUI,
    index: int,
    total: int,
    kind: str,
    thumbnails_dir: Path | None = None,
) -> None:
    """Attach transcript, top comments and a real thumbnail .jpg to one item."""
    video_id = yt_bot.extract_video_id(item.url)
    transcript = ""
    comments: list[dict[str, Any]] = []
    if video_id:
        watch_url = f"https://www.youtube.com/watch?v={video_id}"
        try:
            transcript = yt_bot.fetch_transcript(
                video_id, getattr(item, "captions", None)
            )
        except Exception as exc:  # transcripts are best-effort
            logger.warning("Transcript unavailable for %s: %s", video_id, exc)
            ui.trace(f"⚠️ {kind} {index}/{total}: transcript unavailable ({str(exc)[:60]})")
        has_comments = item.comment_count != 0
        if has_comments:
            try:
                comments = yt_bot.fetch_comments(watch_url, limit=MAX_COMMENTS_PER_VIDEO)
            except Exception as exc:
                logger.warning("Comments unavailable for %s: %s", video_id, exc)
                ui.trace(f"⚠️ {kind} {index}/{total}: comments unavailable ({str(exc)[:60]})")
        item.transcript = transcript
        item.comments = comments
        if thumbnails_dir is not None:
            thumb_path = download_video_thumbnail(
                video_id,
                getattr(item, "thumbnail_url", ""),
                thumbnails_dir / thumbnail_filename(index, item.title, video_id),
            )
            if thumb_path:
                item.thumbnail_file = thumb_path.name
    label = item.title if len(item.title) <= 42 else item.title[:41] + "…"
    ui.trace(
        f"🎬 {kind} {index}/{total}: {label} · transcript {len(transcript):,} chars · "
        f"{len(comments)} comments · thumb {'✓' if getattr(item, 'thumbnail_file', '') else '✗'}"
    )
    ui.update()


def audit_channel_full(url: str, ui: LiveBatchUI) -> FullChannel:
    started = time.monotonic()
    base_url = url.rstrip("/")
    listing, video_entries = yt_bot.fetch_channel_listing(
        base_url, "videos", MAX_CHANNEL_VIDEOS
    )
    resolved_url = yt_bot._resolve_channel_base(base_url, listing)
    profile = yt_bot.build_channel_profile(
        resolved_url, listing, yt_bot.fetch_channel_about(resolved_url)
    )
    own_handles = tuple(
        handle for handle in (profile.handle, profile.channel_id) if handle
    )
    ui.trace(
        f"👤 {profile.name} ({profile.handle or 'no handle'}) · "
        f"{profile.subscribers_text} · {profile.total_views_text}"
    )
    ui.update()

    try:
        _, short_entries = yt_bot.fetch_channel_listing(
            resolved_url, "shorts", MAX_CHANNEL_SHORTS
        )
    except Exception as exc:
        logger.warning("Shorts unavailable for %s: %s", resolved_url, exc)
        ui.trace(f"⚠️ No Shorts tab for {channel_label(url)}")
        short_entries = []

    ui.trace(
        f"📥 {len(video_entries)} videos + {len(short_entries)} Shorts queued for "
        f"{profile.name} — fetching metadata"
    )
    ui.update()
    videos = yt_bot.fetch_items_with_details(
        video_entries, "Video", own_handles=own_handles, progress=ui.trace
    )
    shorts = yt_bot.fetch_items_with_details(
        short_entries, "Short", own_handles=own_handles, progress=ui.trace
    )

    category = categorizer.categorize_channel(
        categorizer.build_snapshot(profile, videos, shorts, [])
    )
    channel_stem = yt_bot.safe_stem(profile.name, "YouTube_Channel")
    folder = category_folder(category.category) / channel_stem
    thumbnails_dir = folder / "thumbnails"
    thumbnails_dir.mkdir(parents=True, exist_ok=True)
    avatar_name = f"{channel_stem}_avatar.jpg"
    avatar_path = download_channel_avatar(profile, folder / avatar_name)
    if avatar_path:
        ui.trace(f"🖼 Avatar saved: {category.category}/{channel_stem}/{avatar_name}")
    else:
        ui.trace(f"⚠️ No avatar could be downloaded for {profile.name}")
    ui.update()

    ui.trace(
        f"🗣 Transcripts + top {MAX_COMMENTS_PER_VIDEO} comments + real thumbnails "
        f"for {len(videos)} videos and {len(shorts)} Shorts"
    )
    ui.update()
    for index, item in enumerate(videos, start=1):
        _enrich_item(item, ui, index, len(videos), "video", thumbnails_dir)
        if index < len(videos):
            time.sleep(random.uniform(VIDEO_DELAY_MIN, VIDEO_DELAY_MAX))
    for index, item in enumerate(shorts, start=1):
        _enrich_item(item, ui, index, len(shorts), "short", thumbnails_dir)
        if index < len(shorts):
            time.sleep(random.uniform(VIDEO_DELAY_MIN, VIDEO_DELAY_MAX))

    ui.trace(f"💬 Scraping community posts for {profile.name}")
    ui.update()
    posts = yt_bot.fetch_community_posts(resolved_url, own_handles)
    yt_bot.resolve_collaborator_ids(videos + shorts, posts, own_channel_id=profile.channel_id)
    collaborators = yt_bot.collect_collaborators(videos + shorts, posts)

    ui.trace(
        f"🧠 Niche: {category.category} / {category.sub_niche} "
        f"(score {category.score}, {category.source}, confidence {category.confidence})"
    )
    for reason in category.reasoning[:2]:
        ui.trace(f"↳ {reason}")
    if collaborators:
        ui.trace(f"🤝 Collaborators detected: {', '.join(collaborators[:8])}")
    ui.update()

    xlsx_path = folder / f"{channel_stem}_Full_Intel.xlsx"
    md_path = folder / f"{channel_stem}_Full_Intel.md"
    write_full_channel_workbook(
        profile,
        videos,
        shorts,
        posts,
        category,
        xlsx_path,
        avatar_file=avatar_name if avatar_path else "",
    )
    md_path.write_text(
        build_full_channel_markdown(
            profile,
            videos,
            shorts,
            posts,
            category,
            avatar_file=avatar_name if avatar_path else "",
        ),
        encoding="utf-8",
    )
    downloaded = sum(1 for item in videos + shorts if getattr(item, "thumbnail_file", ""))
    ui.trace(
        f"📁 Saved {xlsx_path.name} + {md_path.name} + {downloaded} thumbnail(s) in "
        f"{category.category}/{channel_stem}/"
    )
    ui.update()
    return FullChannel(
        url=url,
        profile=profile,
        videos=videos,
        shorts=shorts,
        posts=posts,
        collaborators=collaborators,
        category=category,
        xlsx_path=xlsx_path,
        md_path=md_path,
        avatar_path=avatar_path,
        elapsed=time.monotonic() - started,
    )


# --------------------------------------------------------------------------- #
# Standalone video intel
# --------------------------------------------------------------------------- #
def process_standalone_video(url: str, ui: LiveBatchUI, index: int, total: int) -> StandaloneVideo:
    started = time.monotonic()
    video_id = yt_bot.extract_video_id(url) or ""
    if not video_id:
        raise RuntimeError(f"Not a video URL: {url}")
    watch_url = f"https://www.youtube.com/watch?v={video_id}"
    info = yt_bot.fetch_metadata(watch_url)
    item = yt_bot.build_channel_item(
        info,
        "Video",
        own_handles=(),
        collaboration_source=info,
        video_id=video_id,
        captions=yt_bot._caption_tracks(info),
    )
    thumbnail_url = str(info.get("thumbnail") or yt_bot.thumbnail_url_for(video_id))

    STANDALONE_DIR.mkdir(parents=True, exist_ok=True)
    thumbnails_dir = STANDALONE_DIR / "thumbnails"
    thumbnails_dir.mkdir(parents=True, exist_ok=True)
    thumb_path = download_video_thumbnail(
        video_id,
        info.get("thumbnail"),
        thumbnails_dir / thumbnail_filename(index, item.title, video_id),
    )
    ui.trace(
        f"🎬 Standalone {index}/{total}: {item.title[:48]} — thumbnail "
        f"{'✓' if thumb_path else '✗'}, fetching transcript"
    )
    ui.update()

    transcript = ""
    try:
        transcript = yt_bot.fetch_transcript(video_id, yt_bot._caption_tracks(info))
    except Exception as exc:
        logger.warning("Transcript unavailable for %s: %s", video_id, exc)
        ui.trace(f"⚠️ Transcript unavailable for {video_id}: {str(exc)[:80]}")

    ui.trace(
        f"🗣 Transcript {len(transcript):,} chars — fetching "
        f"{'ALL' if not MAX_COMMENTS_STANDALONE else MAX_COMMENTS_STANDALONE} comments"
    )
    ui.update()
    comments: list[dict[str, Any]] = []
    try:
        comments = yt_bot.fetch_comments(
            watch_url, limit=MAX_COMMENTS_STANDALONE or None
        )
    except Exception as exc:
        logger.warning("Comments unavailable for %s: %s", video_id, exc)
        ui.trace(f"⚠️ Comments unavailable for {video_id}: {str(exc)[:80]}")

    md_path = STANDALONE_DIR / (
        f"{yt_bot.safe_stem(item.title, 'Video')}_{video_id}_Full_Intel.md"
    )
    video = StandaloneVideo(
        video_id=video_id,
        title=item.title,
        url=watch_url,
        upload_date=item.upload_date,
        views=item.views,
        likes=item.likes,
        comment_count=item.comment_count,
        duration_sec=item.duration_sec,
        category=item.category,
        tags=item.tags,
        collaborations=item.collaborations,
        description=item.description,
        thumbnail_url=thumbnail_url,
        transcript=transcript,
        comments=comments,
        md_path=md_path,
        thumbnail_file=thumb_path.name if thumb_path else "",
        elapsed=time.monotonic() - started,
    )
    md_path.write_text(build_standalone_markdown(video), encoding="utf-8")
    ui.trace(
        f"💾 Standalone dossier saved: {md_path.name} · {len(comments)} comments · "
        f"thumb {'✓' if thumb_path else '✗'}"
    )
    ui.update()
    time.sleep(random.uniform(VIDEO_DELAY_MIN, VIDEO_DELAY_MAX))
    return video


# --------------------------------------------------------------------------- #
# Formatting helpers for Excel / Markdown
# --------------------------------------------------------------------------- #
MAX_CELL_CHARS = 30000


def _cell(text: Any, limit: int = MAX_CELL_CHARS) -> str:
    value = str(text if text is not None else "")
    return value if len(value) <= limit else value[: limit - 20] + "\n… (truncated)"


def format_comments_cell(comments: list[dict[str, Any]], limit: int = MAX_CELL_CHARS) -> str:
    lines: list[str] = []
    for index, comment in enumerate(comments, start=1):
        text = " ".join(str(comment.get("text") or "").split())
        author = str(comment.get("author") or "Unknown")
        when = str(comment.get("time") or "")
        votes = str(comment.get("votes") or "0")
        flag = " (reply)" if comment.get("reply") else ""
        lines.append(f"{index}. {author}{flag} · {when} · 👍 {votes}\n{text}")
    return _cell("\n".join(lines), limit)


def format_comments_markdown(comments: list[dict[str, Any]]) -> str:
    if not comments:
        return "_No comments fetched._"
    blocks: list[str] = []
    for index, comment in enumerate(comments, start=1):
        author = str(comment.get("author") or "Unknown")
        when = str(comment.get("time") or "")
        votes = str(comment.get("votes") or "0")
        flag = " · ↳ reply" if comment.get("reply") else ""
        text = (comment.get("text") or "").strip() or "_empty_"
        quoted = "\n".join(f"   > {line}" for line in text.splitlines())
        blocks.append(f"{index}. **{author}** — {when} · 👍 {votes}{flag}\n{quoted}")
    return "\n\n".join(blocks)


def _thumbnail_for(url: str) -> str:
    video_id = yt_bot.extract_video_id(url)
    if not video_id:
        return ""
    return f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg"


def _md_table_row(cells: list[Any]) -> str:
    return "| " + " | ".join(
        yt_bot._md_cell(value) if hasattr(yt_bot, "_md_cell") else str(value)
        for value in cells
    ) + " |"


# --------------------------------------------------------------------------- #
# Full channel workbook / markdown
# --------------------------------------------------------------------------- #
def write_full_channel_workbook(
    profile: Any,
    videos: list[Any],
    shorts: list[Any],
    posts: list[Any],
    category: categorizer.CategoryResult,
    dest: Path,
    *,
    avatar_file: str = "",
) -> Path:
    """`<Channel>_Full_Intel.xlsx` with overview, videos, shorts, posts tabs."""
    if not yt_bot.OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for the channel workbook")

    collaborators = yt_bot.collect_collaborators(videos + shorts, posts)
    overview_rows = [
        ["Channel", "Detected Category", category.category],
        ["Channel", "Detected Sub-Niche", category.sub_niche],
        ["Channel", "Category Confidence", category.confidence],
        ["Channel", "Channel Avatar File", avatar_file or profile.avatar_url or "N/A"],
    ] + yt_bot._overview_rows(profile, videos + shorts, posts, collaborators)
    overview_rows.append(["SEO Summary", "Transcripts Collected", sum(1 for i in videos + shorts if i.transcript)])
    overview_rows.append(["SEO Summary", "Comments Collected", sum(len(i.comments) for i in videos + shorts)])
    overview_rows.append(["SEO Summary", "Thumbnails Downloaded", sum(1 for i in videos + shorts if getattr(i, "thumbnail_file", ""))])

    workbook = yt_bot.Workbook()
    overview = workbook.active
    overview.title = "Channel_Overview"
    yt_bot._write_table(
        overview,
        ["Section", "Metric", "Value"],
        overview_rows,
        freeze=False,
        wrap_columns=(3,),
        caps={1: 24, 2: 34, 3: 90},
    )

    video_sheet = workbook.create_sheet("Videos")
    video_headers = [
        "Title", "Video_URL", "Upload_Date", "Views", "Likes", "Comment_Count",
        "Duration_Sec", "Category", "Tags", "Collaborations_Detected",
        "Thumbnail_URL", "Thumbnail_File", "Description", "Full_Transcript", "Top_Comments",
    ]
    video_rows = [
        [
            item.title, item.url, item.upload_date, item.views, item.likes,
            item.comment_count, item.duration_sec, item.category, ", ".join(item.tags),
            item.collaborations, getattr(item, "thumbnail_url", ""),
            getattr(item, "thumbnail_file", ""), _cell(item.description),
            _cell(item.transcript), format_comments_cell(item.comments),
        ]
        for item in videos
    ]
    yt_bot._write_table(
        video_sheet, video_headers, video_rows,
        wrap_columns=(1, 9, 13, 14, 15),
        caps={1: 60, 2: 50, 9: 40, 10: 34, 13: 60, 14: 70, 15: 70},
    )

    short_sheet = workbook.create_sheet("Shorts")
    short_headers = [
        "Title", "Shorts_URL", "Upload_Date", "Views", "Likes", "Comment_Count",
        "Duration_Sec", "Tags", "Collaborations_Detected",
        "Thumbnail_URL", "Thumbnail_File", "Description", "Full_Transcript", "Top_Comments",
    ]
    short_rows = [
        [
            item.title, item.url, item.upload_date, item.views, item.likes,
            item.comment_count, item.duration_sec, ", ".join(item.tags),
            item.collaborations, getattr(item, "thumbnail_url", ""),
            getattr(item, "thumbnail_file", ""), _cell(item.description),
            _cell(item.transcript), format_comments_cell(item.comments),
        ]
        for item in shorts
    ]
    yt_bot._write_table(
        short_sheet, short_headers, short_rows,
        wrap_columns=(1, 8, 12, 13, 14),
        caps={1: 60, 2: 50, 8: 40, 9: 34, 12: 60, 13: 70, 14: 70},
    )

    post_sheet = workbook.create_sheet("Community_Posts")
    post_headers = [
        "Post_ID", "Published_Time", "Post_Text", "Likes",
        "Comment_Count", "Attachment_Type", "Collaborations_Detected",
    ]
    post_rows = [
        [
            post.post_id, post.published_time, _cell(post.text), post.likes,
            post.comment_count, post.attachment_type, post.collaborations,
        ]
        for post in posts
    ]
    if not post_rows:
        post_rows = [["No community posts found (this channel may not have a Posts tab).", "", "", "", "", "", ""]]
    yt_bot._write_table(
        post_sheet, post_headers, post_rows,
        wrap_columns=(3,), caps={1: 36, 2: 18, 3: 80, 7: 40},
    )

    dest.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(dest)
    return dest


def build_full_channel_markdown(
    profile: Any,
    videos: list[Any],
    shorts: list[Any],
    posts: list[Any],
    category: categorizer.CategoryResult,
    *,
    avatar_file: str = "",
) -> str:
    collaborators = yt_bot.collect_collaborators(videos + shorts, posts)
    lines: list[str] = [f"# {profile.name}", ""]
    avatar_link = f"./{avatar_file}" if avatar_file else profile.avatar_url
    if avatar_link:
        lines += [f"![Channel Avatar]({avatar_link})", ""]
    lines += [
        f"**Niche:** {category.category} / {category.sub_niche} "
        f"(score {category.score}, {category.source}, confidence {category.confidence})",
        "",
        "## 📊 Channel Profile",
        "",
        "| Field | Value |",
        "| --- | --- |",
    ]
    profile_rows = [
        ("Handle", profile.handle or "N/A"),
        ("Channel URL", profile.url),
        ("Channel ID", profile.channel_id or "N/A"),
        ("Subscribers", profile.subscribers_text),
        ("Total Channel Views", profile.total_views_text),
        ("Total Videos (per YouTube)", profile.video_count_text),
        ("Joined Date", profile.joined_date or "N/A"),
        ("Country", profile.country or "N/A"),
        ("Public Email", profile.email),
        ("Videos Analysed", len(videos)),
        ("Shorts Analysed", len(shorts)),
        ("Community Posts", len(posts)),
        ("Collaborators", ", ".join(collaborators) if collaborators else "None detected"),
        ("Channel Keywords", ", ".join(profile.keywords) if profile.keywords else "N/A"),
    ]
    for field, value in profile_rows:
        lines.append(_md_table_row([field, value]))
    if profile.links:
        lines += ["", "### 🔗 External Links", ""]
        for link in profile.links:
            lines.append(f"- **{link.title}:** {link.url}")
    lines += ["", "## 📝 Description", "", profile.description or "_No description._", ""]

    lines += ["## 🎬 Long-Form Videos (" + str(len(videos)) + ")", ""]
    for index, item in enumerate(videos, start=1):
        lines += [
            f"### {index}. {item.title}",
            "",
            f"- **URL:** {item.url}",
            f"- **Uploaded:** {item.upload_date} · **Views:** {yt_bot._fmt_int(item.views)} · "
            f"**Likes:** {yt_bot._fmt_int(item.likes)} · **Comments:** {yt_bot._fmt_int(item.comment_count)}",
            f"- **Duration:** {yt_bot._fmt_duration(item.duration_sec)} · "
            f"**Category:** {item.category or 'N/A'}",
            f"- **Tags:** {', '.join(item.tags) if item.tags else 'N/A'}",
            f"- **Collaborations:** {item.collaborations or 'None detected'}",
            "",
        ]
        thumbnail = (
            f"./thumbnails/{item.thumbnail_file}"
            if getattr(item, "thumbnail_file", "")
            else _thumbnail_for(item.url)
        )
        if thumbnail:
            lines += [f"![Thumbnail]({thumbnail})", ""]
        lines += ["**Transcript**", "", item.transcript or "_Transcript unavailable._", ""]
        lines += [f"**Top {len(item.comments)} Comments**", ""]
        lines += [format_comments_markdown(item.comments), ""]

    if shorts:
        lines += ["## 📱 Shorts (" + str(len(shorts)) + ")", ""]
        for index, item in enumerate(shorts, start=1):
            lines += [
                f"### {index}. {item.title}",
                "",
                f"- **URL:** {item.url}",
                f"- **Uploaded:** {item.upload_date} · **Views:** {yt_bot._fmt_int(item.views)} · "
                f"**Likes:** {yt_bot._fmt_int(item.likes)} · **Comments:** {yt_bot._fmt_int(item.comment_count)}",
                f"- **Duration:** {yt_bot._fmt_duration(item.duration_sec)}",
                f"- **Tags:** {', '.join(item.tags) if item.tags else 'N/A'}",
                f"- **Collaborations:** {item.collaborations or 'None detected'}",
                "",
            ]
            thumbnail = _thumbnail_for(item.url)
            if thumbnail:
                lines += [f"![Thumbnail]({thumbnail})", ""]
            lines += ["**Transcript**", "", item.transcript or "_Transcript unavailable._", ""]
            lines += [f"**Top {len(item.comments)} Comments**", ""]
            lines += [format_comments_markdown(item.comments), ""]

    lines += ["## 💬 Community Posts (" + str(len(posts)) + ")", ""]
    if posts:
        for index, post in enumerate(posts, start=1):
            lines += [
                f"### Post {index} — {post.published_time} · 👍 {yt_bot._fmt_int(post.likes)} · "
                f"💬 {yt_bot._fmt_int(post.comment_count)} · {post.attachment_type}",
                "",
                f"> {post.text}" if post.text else "> _No text._",
                "",
                f"- Collaborations: {post.collaborations or 'None detected'}",
                f"- URL: {post.url}",
                "",
            ]
    else:
        lines += ["_No community posts available._", ""]
    return "\n".join(lines)


def build_standalone_markdown(video: StandaloneVideo) -> str:
    thumbnail_link = (
        f"./thumbnails/{video.thumbnail_file}" if video.thumbnail_file else video.thumbnail_url
    )
    lines = [
        f"# {video.title}",
        "",
        f"![Thumbnail]({thumbnail_link})",
        "",
        "## 📊 Video Details",
        "",
        "| Field | Value |",
        "| --- | --- |",
    ]
    rows = [
        ("Video ID", video.video_id),
        ("URL", video.url),
        ("Uploaded", video.upload_date),
        ("Views", yt_bot._fmt_int(video.views)),
        ("Likes", yt_bot._fmt_int(video.likes)),
        ("Comments (reported)", yt_bot._fmt_int(video.comment_count)),
        ("Comments fetched", f"{len(video.comments):,}"),
        ("Duration", yt_bot._fmt_duration(video.duration_sec)),
        ("Category", video.category or "N/A"),
        ("Tags", ", ".join(video.tags) if video.tags else "N/A"),
        ("Collaborations", video.collaborations or "None detected"),
    ]
    for field, value in rows:
        lines.append(_md_table_row([field, value]))
    lines += [
        "",
        "## 📝 Description",
        "",
        video.description or "_No description._",
        "",
        "## 🗣️ Full Transcript",
        "",
        video.transcript or "_Transcript unavailable._",
        "",
        f"## 💬 Comments ({len(video.comments):,} fetched)",
        "",
        format_comments_markdown(video.comments),
        "",
    ]
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Master workbook
# --------------------------------------------------------------------------- #
OVERVIEW_HEADERS = [
    "Channel", "Handle", "Category", "Sub_Niche", "Category_Source", "Category_Score",
    "Subscribers", "Total_Views", "Videos_Fetched", "Shorts_Fetched", "Posts_Fetched",
    "Transcripts", "Comments_Collected", "Collaborations_Count", "Collaborators",
    "Avg_Views_Video", "Avg_Views_Short", "Joined_Date", "Country", "Public_Email",
    "Channel_URL", "Avatar_File", "Excel_File", "Markdown_File",
]


def _relative_path(path: Path) -> str:
    """Path relative to output/batch_result (used inside the master workbook)."""
    try:
        return path.relative_to(BATCH_DIR).as_posix()
    except ValueError:
        return path.name


def _average_views(items: list[Any]) -> int | None:
    values = [item.views for item in items if item.views is not None]
    return round(sum(values) / len(values)) if values else None


def overview_row(channel: FullChannel) -> list[Any]:
    profile = channel.profile
    return [
        profile.name,
        profile.handle,
        channel.category.category,
        channel.category.sub_niche,
        channel.category.source,
        channel.category.score,
        profile.subscribers,
        profile.total_views,
        len(channel.videos),
        len(channel.shorts),
        len(channel.posts),
        sum(1 for item in channel.items if item.transcript),
        sum(len(item.comments) for item in channel.items),
        len(channel.collaborators),
        ", ".join(channel.collaborators),
        _average_views(channel.videos),
        _average_views(channel.shorts),
        profile.joined_date,
        profile.country,
        profile.email,
        profile.url,
        channel.avatar_path.name if channel.avatar_path else "",
        _relative_path(channel.xlsx_path),
        _relative_path(channel.md_path),
    ]


def build_master_workbook(
    channels: list[FullChannel], standalone: list[StandaloneVideo], dest: Path
) -> Path:
    if not yt_bot.OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for the master workbook")

    workbook = yt_bot.Workbook()
    overview = workbook.active
    overview.title = "All_Channels_Overview"
    rows = [overview_row(channel) for channel in channels]
    yt_bot._write_table(
        overview, OVERVIEW_HEADERS, rows,
        wrap_columns=(1, 15, 23), caps={1: 34, 2: 24, 15: 46, 23: 60},
    )

    for category in sorted({channel.category.category for channel in channels}):
        sheet = workbook.create_sheet(title=category[:31])
        subset = [overview_row(c) for c in channels if c.category.category == category]
        yt_bot._write_table(
            sheet, OVERVIEW_HEADERS, subset,
            wrap_columns=(1, 15, 23), caps={1: 34, 2: 24, 15: 46, 23: 60},
        )

    if standalone:
        sheet = workbook.create_sheet("Standalone_Videos")
        headers = [
            "Video_ID", "Title", "Video_URL", "Upload_Date", "Views", "Likes",
            "Comment_Count", "Comments_Fetched", "Duration_Sec", "Category", "Tags",
            "Collaborations_Detected", "Thumbnail_URL", "Thumbnail_File",
            "Full_Transcript", "All_Comments", "Markdown_File",
        ]
        rows = [
            [
                video.video_id, video.title, video.url, video.upload_date, video.views,
                video.likes, video.comment_count, len(video.comments), video.duration_sec,
                video.category, ", ".join(video.tags), video.collaborations,
                video.thumbnail_url, video.thumbnail_file,
                _cell(video.transcript), format_comments_cell(video.comments),
                _relative_path(video.md_path),
            ]
            for video in standalone
        ]
        yt_bot._write_table(
            sheet, headers, rows,
            wrap_columns=(2, 16, 17), caps={2: 50, 3: 50, 12: 34, 16: 70, 17: 70},
        )

    top_sheet = workbook.create_sheet("Top_Videos_Combined")
    top_headers = [
        "Rank", "Source", "Channel", "Category", "Title", "Video_URL", "Upload_Date",
        "Views", "Likes", "Comment_Count", "Duration_Sec", "Collaborations",
    ]
    entries: list[tuple[str, str, str, Any]] = []
    for channel in channels:
        for item in channel.videos:
            entries.append(("Long-form", channel.profile.name, channel.category.category, item))
    for video in standalone:
        entries.append(("Standalone", video.title, "Standalone Video", video))
    entries.sort(
        key=lambda entry: entry[3].views if entry[3].views is not None else -1,
        reverse=True,
    )
    top_rows = [
        [
            rank, source, channel_name, category, getattr(item, "title", ""),
            getattr(item, "url", ""), getattr(item, "upload_date", ""),
            getattr(item, "views", None), getattr(item, "likes", None),
            getattr(item, "comment_count", None), getattr(item, "duration_sec", None),
            getattr(item, "collaborations", ""),
        ]
        for rank, (source, channel_name, category, item) in enumerate(entries[:50], start=1)
    ]
    yt_bot._write_table(
        top_sheet, top_headers, top_rows,
        wrap_columns=(5, 12), caps={5: 60, 6: 50, 12: 40},
    )

    dest.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(dest)
    return dest


def package_zip(batch_dir: Path, zip_path: Path) -> Path:
    if zip_path.exists():
        zip_path.unlink()
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(batch_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(batch_dir).as_posix())
    return zip_path


# --------------------------------------------------------------------------- #
# Batch orchestration
# --------------------------------------------------------------------------- #
def run_batch(
    ui: LiveBatchUI, channels: list[str], videos: list[str]
) -> BatchSummary:
    summary = BatchSummary()
    if BATCH_DIR.exists():
        shutil.rmtree(BATCH_DIR, ignore_errors=True)
    BATCH_DIR.mkdir(parents=True, exist_ok=True)

    # -- channels ---------------------------------------------------------- #
    for index, url in enumerate(channels, start=1):
        label = channel_label(url)
        ui.card.start_step(1, f"Scraping Channel {index}/{len(channels)}: {label}")
        ui.tick(index - 1, 0)
        ui.trace(f"▶ Channel {index}/{len(channels)}: {label}")
        ui.update()

        result, error = retry_call(
            f"channel {label}", lambda url=url: audit_channel_full(url, ui), ui
        )
        if result is None:
            summary.skipped.append((url, error))
            ui.trace(f"❌ Skipped {label}: {error}")
            ui.tick(index, 0)
            ui.update()
            if index < len(channels):
                polite_delay(
                    ui, CHANNEL_DELAY_MIN, CHANNEL_DELAY_MAX,
                    f"channel {index + 1}/{len(channels)}",
                )
            continue

        summary.channels.append(result)
        ui.trace(
            f"✅ {label}: {len(result.videos)} videos · {len(result.shorts)} shorts · "
            f"{len(result.posts)} posts · {len(result.collaborators)} collaborators · "
            f"{result.category.category} · {result.elapsed:.0f}s"
        )
        if result.avatar_path is not None:
            avatar_caption = (
                f"<b>{result.profile.name}</b>\n"
                f"🧠 {result.category.category} / {result.category.sub_niche}\n"
                f"🎬 {len(result.videos)} videos · 📱 {len(result.shorts)} shorts · "
                f"💬 {len(result.posts)} posts"
            )
            if ui.client.send_photo(result.avatar_path, avatar_caption):
                ui.trace("🖼 Channel avatar sent to Telegram")
            else:
                ui.trace("⚠️ Channel avatar photo upload failed (file still in ZIP)")
        ui.tick(index, 0)
        ui.update()
        if index < len(channels):
            polite_delay(
                ui, CHANNEL_DELAY_MIN, CHANNEL_DELAY_MAX,
                f"channel {index + 1}/{len(channels)}",
            )

    # -- standalone videos ------------------------------------------------- #
    for index, url in enumerate(videos, start=1):
        video_id = yt_bot.extract_video_id(url) or url
        ui.card.set_state(
            f"🎬 Standalone video {index}/{len(videos)} · "
            f"{ui.card.elapsed_text()} elapsed"
        )
        ui.trace(f"▶ Standalone {index}/{len(videos)}: {video_id}")
        ui.update()

        result, error = retry_call(
            f"video {video_id}",
            lambda url=url, index=index: process_standalone_video(
                url, ui, index, len(videos)
            ),
            ui,
        )
        if result is None:
            summary.skipped.append((url, error))
            ui.trace(f"❌ Skipped standalone {video_id}: {error}")
            continue
        summary.standalone.append(result)

    if not summary.channels and not summary.standalone:
        raise RuntimeError("No channel or video could be scraped — nothing to build")

    # -- master workbook ---------------------------------------------------- #
    ui.card.finish_step(1)
    ui.card.start_step(2)
    ui.card.set_state(f"🧠 Building master workbook · {ui.card.elapsed_text()} elapsed")
    ui.trace(
        f"📚 Aggregating {len(summary.channels)} channel(s) + "
        f"{len(summary.standalone)} standalone video(s)"
    )
    ui.update()
    build_master_workbook(summary.channels, summary.standalone, MASTER_XLSX)
    if summary.categories:
        ui.trace(f"🗂 Category tabs: {', '.join(summary.categories)}")
    if summary.standalone:
        ui.trace(f"🎬 Standalone_Videos sheet: {len(summary.standalone)} video(s)")
    ui.card.finish_step(2)

    # -- ZIP + delivery ------------------------------------------------------ #
    ui.card.start_step(3)
    ui.card.set_state(f"📦 Packaging & delivering · {ui.card.elapsed_text()} elapsed")
    package_zip(BATCH_DIR, ZIP_PATH)
    file_count = sum(1 for path in BATCH_DIR.rglob("*") if path.is_file())
    ui.trace(f"🗜 Packaged {file_count} file(s) into All_Channels_Categorized.zip")
    ui.trace("📤 Uploading Master_Channels_Comparison.xlsx (120s timeout, 1 retry)")
    ui.update()

    master_caption = (
        f"📊 <b>Master Channel Comparison</b>\n"
        f"{len(summary.channels)} channel(s) · {len(summary.standalone)} standalone video(s) · "
        f"{len(summary.categories)} categories\n"
        f"🎬 {summary.total_videos} videos · 📱 {summary.total_shorts} shorts · "
        f"🗣 {summary.transcripts} transcripts · 💬 {summary.comments} comments"
    )
    if ui.client.send_document(MASTER_XLSX, master_caption):
        ui.trace("✅ Master workbook delivered")
    else:
        ui.trace("⚠️ Master workbook upload failed (file kept in output/)")

    ui.trace("📤 Uploading All_Channels_Categorized.zip (120s timeout, 1 retry)")
    zip_caption = (
        f"🗂 <b>All Channels Categorized</b>\n"
        f"{file_count} files · folders: "
        + ", ".join(summary.categories + ([STANDALONE_FOLDER] if summary.standalone else []))
    )
    if ui.client.send_document(ZIP_PATH, zip_caption):
        ui.trace("✅ ZIP delivered")
    else:
        ui.trace("⚠️ ZIP upload failed (file kept in output/)")
    ui.card.finish_step(3)

    ui.card.set_state(
        f"✅ Completed · {len(summary.channels)} channels + {len(summary.standalone)} videos · "
        f"{summary.total_videos} videos, {summary.total_shorts} shorts, "
        f"{summary.transcripts} transcripts, {summary.comments} comments · "
        f"{ui.card.elapsed_text()}"
    )
    ui.trace("🎉 Batch complete — full intel delivered above")
    ui.update(force=True)
    return summary


def main() -> int:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s | %(levelname)-8s | %(message)s"
    )
    load_dotenv(BASE_DIR / ".env")

    token = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip()
    chat_id = (os.getenv("TELEGRAM_CHAT_ID") or os.getenv("CHAT_ID") or "").strip()
    if not token or not chat_id:
        print(
            "[batch] TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID are required "
            "(env vars or .env).",
            file=sys.stderr,
        )
        return 2

    channels, videos = load_links()
    if not channels and not videos:
        print(
            "[batch] No YouTube links found. Set EXTRA_CHANNELS or fill channels.txt.",
            file=sys.stderr,
        )
        return 2

    client = TelegramClient(token, chat_id)
    ui = LiveBatchUI(client, len(channels), len(videos))
    try:
        ui.start()
        summary = run_batch(ui, channels, videos)
    except Exception as exc:  # noqa: BLE001 - report failure into the live card
        logger.exception("Batch run failed")
        ui.fail(exc)
        return 1

    print(
        f"[batch] completed: {len(summary.channels)} channel(s), "
        f"{len(summary.standalone)} standalone video(s), "
        f"{len(summary.skipped)} skipped, {summary.elapsed:.0f}s"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
