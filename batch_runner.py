"""Cloud batch runner for YT-Bot channel intelligence.

Runs locally or on GitHub Actions (triggered by Telegram ``/batch``). It reads
channel URLs from ``EXTRA_CHANNELS`` (newline/space separated) or
``channels.txt``, audits every channel with the same pipeline as ``bot.py``,
files the per-channel outputs into category-named folders, builds
``Master_Channels_Comparison.xlsx``, zips everything and delivers it to Telegram
while editing ONE live status message (expandable "Agent Thinking" trace).
"""

from __future__ import annotations

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
from typing import Any

import requests
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import bot as yt_bot  # noqa: E402  (reuses the channel audit pipeline)
import categorizer  # noqa: E402
from progress import ProgressCard  # noqa: E402

logger = logging.getLogger("yt-bot.batch")

OUTPUT_ROOT = BASE_DIR / "output"
BATCH_DIR = OUTPUT_ROOT / "batch_result"
MASTER_XLSX = OUTPUT_ROOT / "Master_Channels_Comparison.xlsx"
ZIP_PATH = OUTPUT_ROOT / "All_Channels_Categorized.zip"
CHANNELS_FILE = BASE_DIR / "channels.txt"

TELEGRAM_REQUEST_TIMEOUT = 30
TELEGRAM_UPLOAD_TIMEOUT = 120
EDIT_INTERVAL = 3.0  # Telegram edits are throttled to one per 3 seconds
MIN_DELAY = 8.0  # polite randomised cooldown between channels (YouTube rate limits)
MAX_DELAY = 15.0
CHANNEL_RETRIES = 2
RETRY_BACKOFF = 5.0

STEP_QUEUE = "Queue Initialized"
STEP_SCRAPE = "Scraping Channels"
STEP_MASTER = "Categorizing & Building Master Excel"
STEP_PACK = "Packaging ZIP & Delivering"


# --------------------------------------------------------------------------- #
# Telegram HTTP client (works from any machine / CI, no python-telegram-bot)
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


# --------------------------------------------------------------------------- #
# Live progress UI
# --------------------------------------------------------------------------- #
class LiveBatchUI:
    """One pinned status message, edited in place (throttled to 1 edit / 3s)."""

    def __init__(self, client: TelegramClient, channels: list[str]):
        self.client = client
        self.total = len(channels)
        self.card = ProgressCard(
            title="🧠 YouTube Channel Intelligence — Cloud Batch",
            steps=[STEP_QUEUE, STEP_SCRAPE, STEP_MASTER, STEP_PACK],
        )

    def start(self) -> None:
        self.card.finish_step(0)
        self.card.start_step(1, f"Preparing {self.total} channel(s)")
        self.card.set_state(f"🚄 Starting · 0/{self.total} channels · 00:00")
        self.card.add_trace(f"Queue initialized with {self.total} channel URL(s)")
        self.client.send_message(self.card.render())

    def trace(self, message: object, icon: str = "•") -> None:
        self.card.add_trace(message, icon=icon)

    def update(self, force: bool = False) -> None:
        self.client.edit_message(self.card.render(), force=force)

    def tick(self, done: int) -> None:
        self.card.set_state(
            f"🔄 Running · {done}/{self.total} channels · {self.card.elapsed_text()} elapsed"
        )

    def fail(self, error: object) -> None:
        self.card.set_state(f"❌ Batch failed · {self.card.elapsed_text()} elapsed")
        for index, step in enumerate(self.card.steps):
            if step.status == "🔄":
                self.card.fail_step(index)
        self.card.add_trace(f"❌ {str(error)[:300]}", icon="❌")
        self.update(force=True)


# --------------------------------------------------------------------------- #
# Channel queue + audit
# --------------------------------------------------------------------------- #
def load_channel_urls() -> list[str]:
    """Read channel URLs from EXTRA_CHANNELS env or channels.txt."""
    raw = (os.getenv("EXTRA_CHANNELS") or "").strip()
    source = "EXTRA_CHANNELS"
    if not raw and CHANNELS_FILE.exists():
        raw = CHANNELS_FILE.read_text(encoding="utf-8", errors="ignore")
        source = CHANNELS_FILE.name
    urls: list[str] = []
    seen: set[str] = set()
    for token in re.split(r"[\s,]+", raw):
        candidate = token.strip().strip('"').strip("'")
        if not candidate or candidate.startswith("#"):
            continue
        url = yt_bot.extract_channel_url(candidate)
        if not url:
            continue
        key = url.lower().rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        urls.append(url)
    logger.info("Loaded %d channel URL(s) from %s", len(urls), source)
    return urls


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


def audit_with_retries(url: str, ui: LiveBatchUI) -> tuple[Any | None, str]:
    """Audit one channel, retrying once; returns (result, error)."""
    label = channel_label(url)
    last_error = ""
    for attempt in range(1, CHANNEL_RETRIES + 1):
        temp_dir = Path(tempfile.mkdtemp(prefix="yt_batch_"))
        try:
            result = yt_bot.audit_channel(url, temp_dir, progress=ui.trace)
            return result, ""
        except Exception as exc:  # noqa: BLE001 - one bad channel must not kill the batch
            last_error = str(exc)[:250]
            logger.warning("Channel %s failed (attempt %d/%d): %s", url, attempt, CHANNEL_RETRIES, exc)
            ui.trace(f"⚠️ {label} attempt {attempt}/{CHANNEL_RETRIES} failed: {last_error}")
            if attempt < CHANNEL_RETRIES:
                ui.trace(f"↻ Retrying {label} in {int(RETRY_BACKOFF)}s")
                ui.update()
                time.sleep(RETRY_BACKOFF)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    return None, last_error


def polite_delay(ui: LiveBatchUI, seconds: float, next_label: str) -> None:
    """Randomised cooldown with live countdown in the thinking trace."""
    remaining = seconds
    while remaining > 0:
        ui.trace(f"😴 Politeness delay: {math.ceil(remaining)}s before {next_label}")
        ui.update()
        step = min(3.0, remaining)
        time.sleep(step)
        remaining -= step


def category_folder(category: str) -> Path:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", category).strip().strip(".") or "General"
    return BATCH_DIR / safe


def store_channel_files(result: Any, category: str) -> tuple[Path, Path]:
    """Move a channel's .xlsx/.md into its category folder inside batch_result/."""
    folder = category_folder(category)
    folder.mkdir(parents=True, exist_ok=True)
    xlsx_path = folder / result.xlsx_path.name
    md_path = folder / result.markdown_path.name
    shutil.move(str(result.xlsx_path), xlsx_path)
    shutil.move(str(result.markdown_path), md_path)
    return xlsx_path, md_path


# --------------------------------------------------------------------------- #
# Aggregate workbook + ZIP
# --------------------------------------------------------------------------- #
def _average_views(items: list[Any]) -> int | None:
    values = [item.views for item in items if item.views is not None]
    return round(sum(values) / len(values)) if values else None


@dataclass
class ChannelRun:
    url: str
    result: Any
    category: categorizer.CategoryResult
    xlsx_path: Path
    md_path: Path


@dataclass
class BatchSummary:
    runs: list[ChannelRun] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)
    started_at: float = field(default_factory=time.monotonic)

    @property
    def elapsed(self) -> float:
        return time.monotonic() - self.started_at

    @property
    def videos(self) -> int:
        return sum(len(run.result.videos) for run in self.runs)

    @property
    def shorts(self) -> int:
        return sum(len(run.result.shorts) for run in self.runs)

    @property
    def posts(self) -> int:
        return sum(len(run.result.posts) for run in self.runs)

    @property
    def collaborators(self) -> int:
        return sum(len(run.result.collaborators) for run in self.runs)


OVERVIEW_HEADERS = [
    "Channel",
    "Handle",
    "Category",
    "Sub_Niche",
    "Category_Source",
    "Category_Score",
    "Subscribers",
    "Total_Views",
    "Videos_Fetched",
    "Shorts_Fetched",
    "Posts_Fetched",
    "Collaborations_Count",
    "Collaborators",
    "Avg_Views_Video",
    "Avg_Views_Short",
    "Joined_Date",
    "Country",
    "Public_Email",
    "Channel_URL",
    "Excel_File",
    "Markdown_File",
]


def overview_row(run: ChannelRun) -> list[Any]:
    profile = run.result.profile
    return [
        profile.name,
        profile.handle,
        run.category.category,
        run.category.sub_niche,
        run.category.source,
        run.category.score,
        profile.subscribers,
        profile.total_views,
        len(run.result.videos),
        len(run.result.shorts),
        len(run.result.posts),
        len(run.result.collaborators),
        ", ".join(run.result.collaborators),
        _average_views(run.result.videos),
        _average_views(run.result.shorts),
        profile.joined_date,
        profile.country,
        profile.email,
        profile.url,
        run.xlsx_path.name,
        run.md_path.name,
    ]


def build_master_workbook(runs: list[ChannelRun], dest: Path) -> Path:
    """Master_Channels_Comparison.xlsx: overview + one tab per category + top videos."""
    if not yt_bot.OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for the master workbook")

    workbook = yt_bot.Workbook()
    overview = workbook.active
    overview.title = "All_Channels_Overview"
    rows = [overview_row(run) for run in runs]
    yt_bot._write_table(
        overview,
        OVERVIEW_HEADERS,
        rows,
        wrap_columns=(1, 13, 21),
        caps={1: 34, 2: 24, 13: 46, 19: 40},
    )

    categories = sorted({run.category.category for run in runs})
    for category in categories:
        sheet = workbook.create_sheet(title=category[:31])
        subset = [overview_row(run) for run in runs if run.category.category == category]
        yt_bot._write_table(
            sheet,
            OVERVIEW_HEADERS,
            subset,
            wrap_columns=(1, 13, 21),
            caps={1: 34, 2: 24, 13: 46, 19: 40},
        )

    top_sheet = workbook.create_sheet("Top_Videos_Combined")
    video_rows = []
    for run in runs:
        for item in run.result.videos:
            video_rows.append((run, item))
    video_rows.sort(
        key=lambda pair: pair[1].views if pair[1].views is not None else -1, reverse=True
    )
    top_headers = [
        "Rank",
        "Channel",
        "Category",
        "Title",
        "Video_URL",
        "Upload_Date",
        "Views",
        "Likes",
        "Comment_Count",
        "Duration_Sec",
        "Collaborations",
    ]
    top_rows = [
        [
            rank,
            run.result.profile.name,
            run.category.category,
            item.title,
            item.url,
            item.upload_date,
            item.views,
            item.likes,
            item.comment_count,
            item.duration_sec,
            item.collaborations,
        ]
        for rank, (run, item) in enumerate(video_rows[:50], start=1)
    ]
    yt_bot._write_table(
        top_sheet,
        top_headers,
        top_rows,
        wrap_columns=(4, 11),
        caps={4: 60, 5: 50, 11: 40},
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
def run_batch(ui: LiveBatchUI, channels: list[str]) -> BatchSummary:
    summary = BatchSummary()
    if BATCH_DIR.exists():
        shutil.rmtree(BATCH_DIR, ignore_errors=True)
    BATCH_DIR.mkdir(parents=True, exist_ok=True)
    total = len(channels)

    for index, url in enumerate(channels, start=1):
        label = channel_label(url)
        ui.card.start_step(1, f"Scraping Channel {index}/{total}: {label}")
        ui.tick(index - 1)
        ui.trace(f"▶ Channel {index}/{total}: {label}")

        result, error = audit_with_retries(url, ui)
        if result is None:
            summary.skipped.append((url, error))
            ui.trace(f"❌ Skipped {label}: {error}")
        else:
            snapshot = categorizer.build_snapshot(
                result.profile, result.videos, result.shorts, result.posts
            )
            category = categorizer.categorize_channel(snapshot)
            ui.trace(
                f"🧠 Niche for {label}: {category.category} / {category.sub_niche} "
                f"(score {category.score}, {category.source}, confidence {category.confidence})"
            )
            for reason in category.reasoning[:2]:
                ui.trace(f"↳ {reason}")
            xlsx_path, md_path = store_channel_files(result, category.category)
            summary.runs.append(
                ChannelRun(url=url, result=result, category=category, xlsx_path=xlsx_path,
                           md_path=md_path)
            )
            ui.trace(
                f"📊 {label}: {len(result.videos)} videos · {len(result.shorts)} shorts · "
                f"{len(result.posts)} posts · {len(result.collaborators)} collaborators"
            )
            ui.trace(f"📁 Filed under {category.category}/")
        ui.tick(index)
        ui.update()
        if index < total:
            polite_delay(
                ui,
                random.uniform(MIN_DELAY, MAX_DELAY),
                f"channel {index + 1}/{total}",
            )

    ui.card.finish_step(1)
    if not summary.runs:
        raise RuntimeError("No channel could be scraped — nothing to build")

    # 3. Master workbook --------------------------------------------------- #
    ui.card.start_step(2)
    ui.card.set_state(
        f"🔄 Building master workbook · {len(summary.runs)}/{total} channels · "
        f"{ui.card.elapsed_text()} elapsed"
    )
    ui.trace(f"📚 Aggregating {len(summary.runs)} channel(s) into Master_Channels_Comparison.xlsx")
    build_master_workbook(summary.runs, MASTER_XLSX)
    categories = sorted({run.category.category for run in summary.runs})
    ui.trace(f"🗂 Category tabs: {', '.join(categories)}")
    ui.card.finish_step(2)

    # 4. ZIP + delivery ---------------------------------------------------- #
    ui.card.start_step(3)
    ui.card.set_state(f"🔄 Packaging & delivering · {ui.card.elapsed_text()} elapsed")
    package_zip(BATCH_DIR, ZIP_PATH)
    file_count = sum(1 for path in BATCH_DIR.rglob("*") if path.is_file())
    ui.trace(f"🗜 Packaged {file_count} file(s) into All_Channels_Categorized.zip")
    ui.trace("📤 Uploading Master_Channels_Comparison.xlsx (timeout 120s, 1 retry)")
    ui.update()

    master_caption = (
        f"📊 <b>Master Channel Comparison</b>\n"
        f"{len(summary.runs)} channel(s) · {len(categories)} categories · "
        f"{summary.videos} videos · {summary.shorts} shorts analysed\n"
        f"🤝 {summary.collaborators} collaborators detected"
    )
    if ui.client.send_document(MASTER_XLSX, master_caption):
        ui.trace("✅ Master workbook delivered")
    else:
        ui.trace("⚠️ Master workbook upload failed (file kept in output/)")

    ui.trace("📤 Uploading All_Channels_Categorized.zip (timeout 120s, 1 retry)")
    zip_caption = (
        f"🗂 <b>All Channels Categorized</b>\n"
        f"{len(summary.runs)} channel(s) · {file_count} files · "
        f"folders: {', '.join(categories)}"
    )
    if ui.client.send_document(ZIP_PATH, zip_caption):
        ui.trace("✅ ZIP delivered")
    else:
        ui.trace("⚠️ ZIP upload failed (file kept in output/)")
    ui.card.finish_step(3)

    summary_line = (
        f"✅ Completed · {len(summary.runs)} channel(s), {len(summary.skipped)} skipped · "
        f"{summary.videos} videos · {summary.shorts} shorts · {summary.posts} posts · "
        f"{summary.collaborators} collaborators · {ui.card.elapsed_text()} elapsed"
    )
    ui.card.set_state(summary_line)
    ui.trace("🎉 Batch complete — files delivered above")
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

    channels = load_channel_urls()
    if not channels:
        print(
            "[batch] No channel URLs found. Set EXTRA_CHANNELS or fill channels.txt.",
            file=sys.stderr,
        )
        return 2

    client = TelegramClient(token, chat_id)
    ui = LiveBatchUI(client, channels)
    try:
        ui.start()
        summary = run_batch(ui, channels)
    except Exception as exc:  # noqa: BLE001 - report failure into the live card
        logger.exception("Batch run failed")
        ui.fail(exc)
        return 1

    print(
        f"[batch] completed: {len(summary.runs)} channel(s), "
        f"{len(summary.skipped)} skipped, {summary.elapsed:.0f}s"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
