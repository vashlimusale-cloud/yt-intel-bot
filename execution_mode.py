"""Strict 4-mode execution router shared by bot.py and cloud_poller.py.

Modes:
  AUTO   - fast single jobs on the active listener, heavy channel/batch jobs on
           GitHub Actions, plus the 5-min Actions cron failover poller.
  RENDER - 100% of jobs run on this Render instance (no cloud dispatch).
  CLOUD  - 100% of jobs are dispatched to GitHub Actions.
  LOCAL  - Render goes into polite polling standby; the user's local PC runs
           everything (no Telegram 409 conflicts).

The chosen mode is persisted in ``.mode_state.json`` and mirrored to the
GitHub repository variable ``BOT_EXECUTION_MODE`` so Render and Actions agree.
"""

from __future__ import annotations

import json
import logging
import os
import time
from pathlib import Path

import requests

logger = logging.getLogger("yt-bot.mode")

BASE_DIR = Path(__file__).resolve().parent
STATE_PATH = BASE_DIR / ".mode_state.json"
REMOTE_VARIABLE = "BOT_EXECUTION_MODE"

MODES = {
    "AUTO": "⚡ Auto (Never Fail)",
    "RENDER": "🖥️ Render Only",
    "CLOUD": "☁️ GitHub Only",
    "LOCAL": "💻 Local PC Only",
}
MODE_ORDER = ["AUTO", "RENDER", "CLOUD", "LOCAL"]
DEFAULT_MODE = "AUTO"
HEAVY_JOBS = {"channel", "batch"}
LIGHT_JOBS = {"single", "viral", "similar"}
API_TIMEOUT = 20


def normalize_mode(raw: object) -> str:
    value = str(raw or "").strip().upper()
    aliases = {
        "RENDER": "RENDER",
        "RENDER_ONLY": "RENDER",
        "CLOUD": "CLOUD",
        "GITHUB": "CLOUD",
        "GITHUB_ONLY": "CLOUD",
        "LOCAL": "LOCAL",
        "LOCAL_PC": "LOCAL",
        "AUTO": "AUTO",
        "AUTO_NEVER_FAIL": "AUTO",
    }
    return aliases.get(value, DEFAULT_MODE)


def mode_label(mode: str) -> str:
    return MODES.get(normalize_mode(mode), MODES[DEFAULT_MODE])


def _pat_repo() -> tuple[str, str]:
    return (
        (os.getenv("GITHUB_PAT") or "").strip(),
        (os.getenv("GITHUB_REPO") or "").strip(),
    )


def _headers(pat: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {pat}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "YT-Bot/1.0",
    }


def load_local_mode() -> str | None:
    try:
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    mode = normalize_mode(data.get("mode")) if isinstance(data, dict) else None
    return mode if data else None


def save_local_mode(mode: str) -> None:
    payload = {
        "mode": normalize_mode(mode),
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    }
    try:
        STATE_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    except OSError as exc:
        logger.warning("Could not persist mode state: %s", exc)


def fetch_remote_mode(pat: str | None = None, repo: str | None = None) -> str | None:
    pat = pat or _pat_repo()[0]
    repo = repo or _pat_repo()[1]
    if not pat or not repo:
        return None
    try:
        response = requests.get(
            f"https://api.github.com/repos/{repo}/actions/variables/{REMOTE_VARIABLE}",
            headers=_headers(pat),
            timeout=API_TIMEOUT,
        )
        if response.ok:
            return normalize_mode((response.json() or {}).get("value"))
        if response.status_code != 404:
            logger.warning("Remote mode read failed: %s %s", response.status_code, response.text[:120])
    except requests.RequestException as exc:
        logger.warning("Remote mode read error: %s", exc)
    return None


def push_remote_mode(mode: str) -> tuple[bool, str]:
    pat, repo = _pat_repo()
    mode = normalize_mode(mode)
    if not pat or not repo:
        return False, "GitHub PAT/repo not configured"
    base = f"https://api.github.com/repos/{repo}/actions/variables"
    try:
        response = requests.patch(
            f"{base}/{REMOTE_VARIABLE}",
            headers=_headers(pat),
            json={"name": REMOTE_VARIABLE, "value": mode},
            timeout=API_TIMEOUT,
        )
        if response.status_code in (200, 204):
            return True, "Repository variable updated"
        if response.status_code == 404:
            response = requests.post(
                base,
                headers=_headers(pat),
                json={"name": REMOTE_VARIABLE, "value": mode},
                timeout=API_TIMEOUT,
            )
            if response.status_code in (201, 204):
                return True, "Repository variable created"
        return False, f"GitHub API {response.status_code}: {response.text[:120]}"
    except requests.RequestException as exc:
        return False, f"GitHub API unreachable: {exc}"


def resolve_mode() -> str:
    """Mode to use at startup: local state first, then GitHub variable, then AUTO."""
    local = load_local_mode()
    if local:
        return local
    remote = fetch_remote_mode()
    if remote:
        save_local_mode(remote)
        return remote
    return normalize_mode(os.getenv("BOT_EXECUTION_MODE") or os.getenv("EXECUTION_MODE"))


def current_mode() -> str:
    return load_local_mode() or normalize_mode(os.getenv("BOT_EXECUTION_MODE"))


def mode_changed(new_mode: str) -> bool:
    return normalize_mode(new_mode) != current_mode()


def sync_mode(mode: str) -> tuple[bool, str]:
    """Persist locally and mirror to the GitHub repository variable."""
    mode = normalize_mode(mode)
    save_local_mode(mode)
    return push_remote_mode(mode)


def is_render() -> bool:
    return bool(
        os.getenv("RENDER")
        or os.getenv("RENDER_SERVICE_ID")
        or os.getenv("RENDER_EXTERNAL_URL")
    )


def standby_required(mode: str | None = None) -> bool:
    """Render must not poll Telegram while the user's local PC owns the bot."""
    return is_render() and normalize_mode(mode or current_mode()) == "LOCAL"


def route_job(job: str, mode: str | None = None) -> str:
    """Return "local" or "cloud" for a job kind (single/viral/similar/channel/batch)."""
    mode = normalize_mode(mode or current_mode())
    if mode == "CLOUD":
        return "cloud"
    if mode in ("RENDER", "LOCAL"):
        return "local"
    # AUTO: heavy channel/batch work goes to GitHub Actions, light work stays local.
    return "cloud" if job in HEAVY_JOBS else "local"


def describe_mode(mode: str | None = None) -> str:
    mode = normalize_mode(mode or current_mode())
    if mode == "RENDER":
        return "everything runs on this Render instance (no GitHub dispatch)"
    if mode == "CLOUD":
        return "every job is dispatched to GitHub Actions"
    if mode == "LOCAL":
        return "this Render instance is on standby; your local PC runs every job"
    return "single/viral/similar on the active listener, channel/batch on GitHub Actions"
