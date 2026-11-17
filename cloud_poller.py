"""Telegram failover poller for GitHub Actions (runs every 5 minutes).

Safety net for AUTO mode: if the Render listener sleeps/dies, messages older
than 90 s are picked up here, dispatched to the scrape workflow (or processed by
the same dispatch path), acknowledged via getUpdates offset, and the user gets a
short failure-over note.

If the active listener is running, getUpdates returns 409 and this poller exits
without touching anything.
"""

from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import bot  # noqa: E402
import execution_mode  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)-8s | %(message)s")
logger = logging.getLogger("yt-bot.poller")

STALE_SECONDS = 90
TELEGRAM_TIMEOUT = 30
DISPATCH_TIMEOUT = 30


def _telegram(token: str, method: str, payload: dict[str, Any]) -> requests.Response:
    return requests.post(
        f"https://api.telegram.org/bot{token}/{method}", json=payload, timeout=TELEGRAM_TIMEOUT
    )


def _dispatch(links: list[str], chat_id: int | str, token: str, repo: str) -> tuple[bool, str]:
    if not links:
        return False, "no links"
    url = f"https://api.github.com/repos/{repo}/actions/workflows/scrape_channels.yml/dispatches"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "YT-Bot-Poller/1.0",
    }
    payload = {
        "ref": os.getenv("GITHUB_BRANCH") or "main",
        "inputs": {"extra_channels": "\n".join(links), "chat_id": str(chat_id)},
    }
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=DISPATCH_TIMEOUT)
    except requests.RequestException as exc:
        return False, str(exc)[:200]
    if response.status_code in (201, 204):
        return True, "dispatched"
    return False, f"GitHub API {response.status_code}: {response.text[:160]}"


def _notify(token: str, chat_id: int | str, text: str) -> None:
    try:
        _telegram(
            token,
            "sendMessage",
            {
                "chat_id": chat_id,
                "text": text,
                "parse_mode": "HTML",
                "disable_web_page_preview": True,
            },
        )
    except requests.RequestException as exc:
        logger.warning("Failover notification failed: %s", exc)


def main() -> int:
    load_dotenv(BASE_DIR / ".env")
    mode = execution_mode.normalize_mode(
        os.getenv("BOT_EXECUTION_MODE") or execution_mode.current_mode()
    )
    if mode == "LOCAL":
        logger.info("Execution mode LOCAL: Render/local listener owns jobs, poller exiting")
        return 0

    token = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip()
    repo = (os.getenv("GITHUB_REPO") or os.getenv("GITHUB_REPOSITORY") or "").strip()
    dispatch_token = (os.getenv("GH_PAT") or os.getenv("GITHUB_PAT") or os.getenv("GITHUB_TOKEN") or "").strip()
    if not token or not repo or not dispatch_token:
        logger.error("TELEGRAM_BOT_TOKEN, GITHUB_REPO and a GitHub token are required")
        return 2

    try:
        response = requests.get(
            f"https://api.telegram.org/bot{token}/getUpdates",
            params={"timeout": 0, "limit": 100, "allowed_updates": json.dumps(["message"])},
            timeout=TELEGRAM_TIMEOUT,
        )
    except requests.RequestException as exc:
        logger.error("getUpdates failed: %s", exc)
        return 1

    data = response.json()
    if not response.ok or not data.get("ok"):
        description = str(data.get("description") or '')
        if "Conflict" in description or response.status_code == 409:
            logger.info("Active listener is polling (409) - failover not needed")
            return 0
        logger.error("getUpdates rejected: %s", description or response.text[:200])
        return 1

    updates = data.get("result") or []
    if not updates:
        logger.info("No pending updates")
        return 0

    now = time.time()
    pending: list[tuple[dict[str, Any], list[str]]] = []
    for update in updates:
        message = update.get("message") or {}
        text = str(message.get("text") or "")
        date = int(message.get("date") or 0)
        if now - date < STALE_SECONDS:
            continue
        channels, videos = bot.extract_batch_links(text)
        links = channels + videos
        if links:
            pending.append((message, links))

    if not pending:
        logger.info("No stale unhandled messages older than %ds", STALE_SECONDS)
        return 0

    failures = 0
    for message, links in pending:
        chat_id = (message.get("chat") or {}).get("id")
        if chat_id is None:
            continue
        ok, detail = _dispatch(links, chat_id, dispatch_token, repo)
        logger.info("Failover dispatch for chat %s: %s (%s)", chat_id, ok, detail)
        if ok:
            _notify(
                token,
                chat_id,
                "🛟 <b>Failover active</b> — I picked up your message after the main "
                "listener went quiet. The job is now running on GitHub Actions; live "
                "Thinking updates and results will arrive in this chat.",
            )
        else:
            failures += 1
            _notify(token, chat_id, f"⚠️ Failover could not dispatch your message: {detail}")

    last_update = max(int(update["update_id"]) for update in updates)
    try:
        _telegram(token, "getUpdates", {"offset": last_update + 1, "timeout": 0, "limit": 1})
    except requests.RequestException as exc:
        logger.warning("Could not acknowledge updates: %s", exc)

    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
