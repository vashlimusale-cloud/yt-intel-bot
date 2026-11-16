"""Shared Telegram "agent-style" live progress cards for YT-Bot.

Both the local bot (single links) and the cloud batch runner render the same
expandable thinking UI. This module is pure rendering + state: the callers do
the actual Telegram send/edit calls, so it works from python-telegram-bot and
from plain HTTP alike.

The card looks like this:

    🧠 YouTube Channel Intelligence - Cloud Batch
    Status: 🔄 Running · 4/15 channels · 03:52

    <blockquote expandable>🧠 Agent Thinking & Live Trace
    • [00:01] Queue initialized with 15 channel URLs
    • [00:07] Detected niche: Tech & AI (score 18.0, rules)
    </blockquote>

    ✅ 1. Queue Initialized
    🔄 2. Scraping Channel 4/15: @MrBeast
    ⏳ 3. Categorizing & Building Master Excel
    ⏳ 4. Packaging ZIP & Delivering
"""

from __future__ import annotations

import html
import time
from dataclasses import dataclass, field

MAX_MESSAGE_LENGTH = 4096  # Telegram hard limit
MAX_TRACE_LINES = 18
MAX_TRACE_LINE_LENGTH = 240
MAX_STATE_LENGTH = 180
MAX_FOOTER_LENGTH = 400
MAX_STEP_LABEL_LENGTH = 120

PENDING = "⏳"
ACTIVE = "🔄"
DONE = "✅"
FAILED = "❌"


def escape(value: object) -> str:
    """Escape text for Telegram HTML parse mode."""
    return html.escape(str(value if value is not None else ""), quote=False)


def clip(value: object, limit: int) -> str:
    text = " ".join(str(value if value is not None else "").split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def format_elapsed(seconds: float) -> str:
    total = max(0, int(seconds))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


@dataclass
class Step:
    label: str
    status: str = PENDING

    def render(self, index: int) -> str:
        return f"{self.status} {index}. {escape(clip(self.label, MAX_STEP_LABEL_LENGTH))}"


@dataclass
class ProgressCard:
    """State holder for one live Telegram message."""

    title: str
    steps: list[Step] = field(default_factory=list)
    trace: list[str] = field(default_factory=list)
    state: str = "🔄 Working…"
    footer: str = ""
    started_at: float = field(default_factory=time.monotonic)

    def __post_init__(self) -> None:
        self.steps = [
            step if isinstance(step, Step) else Step(str(step)) for step in self.steps
        ]

    # -- state ------------------------------------------------------------ #
    def set_step(self, index: int, status: str, label: str | None = None) -> None:
        if 0 <= index < len(self.steps):
            self.steps[index].status = status
            if label is not None:
                self.steps[index].label = label

    def start_step(self, index: int, label: str | None = None) -> None:
        self.set_step(index, ACTIVE, label)

    def finish_step(self, index: int, label: str | None = None) -> None:
        self.set_step(index, DONE, label)

    def fail_step(self, index: int, label: str | None = None) -> None:
        self.set_step(index, FAILED, label)

    def set_state(self, state: str) -> None:
        self.state = state

    # -- trace ------------------------------------------------------------ #
    def add_trace(self, message: object, icon: str = "•") -> None:
        line = f"{icon} [{self.elapsed_text()}] {clip(message, MAX_TRACE_LINE_LENGTH)}"
        self.trace.append(line)
        # Keep memory bounded even on very long runs.
        if len(self.trace) > 300:
            del self.trace[: len(self.trace) - 300]

    def elapsed_text(self) -> str:
        return format_elapsed(time.monotonic() - self.started_at)

    # -- rendering -------------------------------------------------------- #
    def render(self) -> str:
        """Render the HTML message, shrinking the trace until it fits."""
        for keep in (MAX_TRACE_LINES, 12, 8, 5, 3, 1, 0):
            text = self._render_with(keep)
            if len(text) <= MAX_MESSAGE_LENGTH:
                return text
        return text[: MAX_MESSAGE_LENGTH - 1] + "…"

    def _render_with(self, keep: int) -> str:
        trace = self.trace[-keep:] if keep else []
        omitted = max(0, len(self.trace) - len(trace))
        trace_block = ["… (earlier entries omitted)"] if omitted else []
        trace_block += trace
        if not trace_block:
            trace_block = ["(waiting for the first event)"]

        body: list[str] = [
            f"<b>{escape(clip(self.title, 200))}</b>",
            escape(clip(self.state, MAX_STATE_LENGTH)),
            "",
            "<blockquote expandable><b>🧠 Agent Thinking &amp; Live Trace</b>",
            *[escape(line) for line in trace_block],
            "</blockquote>",
            "",
            *[step.render(index) for index, step in enumerate(self.steps, 1)],
        ]
        if self.footer:
            body += ["", escape(clip(self.footer, MAX_FOOTER_LENGTH))]
        return "\n".join(body)
