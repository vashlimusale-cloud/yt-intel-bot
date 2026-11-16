# YT-Bot — YouTube Channel Intelligence

Telegram bot that turns YouTube links into deep research dossiers with an
agent-style live progress card (expandable 🧠 thinking trace). It handles
**channels and individual videos in the same batch** and can offload the whole
workload to GitHub Actions so you can close your laptop.

## What it does

| You send | You get |
| --- | --- |
| Video / Shorts URL | Thumbnail + Markdown dossier (metadata, description, timestamped transcript, every comment) |
| One channel URL | Channel avatar + Excel `Channel_Intel` workbook + Markdown SEO audit (local run) |
| `/batch` + any mix of channel/video links, or 2+ links in one message | Full cloud run on GitHub Actions: live Thinking card, `Master_Channels_Comparison.xlsx` + `All_Channels_Categorized.zip` delivered right here |
| `/start` | Help card |

Every action uses one status card edited in place (max 1 edit / 3 s) with a
tappable `<blockquote expandable>` **🧠 Agent Thinking & Live Trace** block:
current item, detected niche + reasoning, `24/64` video progress, transcript /
comment status, collaborations and elapsed time. The card stays visible with the
final ✅/❌ state.

## Workload it is tuned for

5–7 channels with ~6–70 videos each, **plus individual video URLs mixed into the
same batch** (~150–250 items total) → roughly 10–18 minutes on GitHub Actions.
Polite pacing is built in: 1–2 s between per-video transcript/comment calls and
6–10 s between channels.

### Mixed-link router (`batch_runner.py`)

- **Channel URLs** (`@handle`, `/channel/`, `/c/`, `/user/`) → full profile/about
  (avatar, subscribers, total views, joined date, country, public email,
  description, external links, keywords), **all videos up to
  `MAX_CHANNEL_VIDEOS = 100`**, **all Shorts up to `MAX_CHANNEL_SHORTS = 100`**,
  community posts, collaborations, and for every item a **full timestamped
  transcript** (`[MM:SS]`) plus the **top `MAX_COMMENTS_PER_VIDEO = 25`**
  comments.
- **Standalone video / Shorts URLs** (`watch?v=`, `youtu.be/`, `/shorts/`) → full
  metadata, thumbnail URL, full timestamped transcript and **ALL comments**
  (`MAX_COMMENTS_STANDALONE = 0` = unlimited), saved in
  `output/batch_result/Standalone_Videos/` and in a `Standalone_Videos` master tab.

### Output layout

```
output/
  Master_Channels_Comparison.xlsx      # overview + one tab per detected category
                                       # + Standalone_Videos + Top_Videos_Combined
  All_Channels_Categorized.zip         # everything below, zipped
  batch_result/
    <Category>/                        # e.g. "Tech & AI", "Gaming", …
      <Channel_Name>_Full_Intel.xlsx   # Channel_Overview, Videos, Shorts, Community_Posts
                                       # Videos/Shorts have Full_Transcript + Top_Comments columns
      <Channel_Name>_Full_Intel.md     # profile + every video's metadata, thumbnail,
                                       # transcript and top comments
    Standalone_Videos/
      <Title>_<video_id>_Full_Intel.md # metadata, thumbnail, full transcript, ALL comments
```

## Smart auto-niche categorization

`categorizer.py` classifies every channel into `Documentary & Storytelling`,
`Tech & AI`, `Finance & Business`, `Podcasts & Interviews`,
`Education & Explainers`, `Gaming`, `Lifestyle & Vlogs` or `General` using a free
weighted scoring model over bio, channel tags, YouTube categories, recent
titles, posts and links — and explains itself in the Thinking trace (e.g.
*"Matched 7 tech & ai signals …: ai, chatgpt, llm"*). Set `LLM_API_KEY` (OpenRouter
/ DeepSeek / OpenAI-compatible, optional `LLM_BASE_URL` / `LLM_MODEL`) to refine
the sub-niche with an LLM; any failure falls back to the rules.

## Setup

`.env` (already configured by the installer):

```
TELEGRAM_BOT_TOKEN=...      # from @BotFather
GITHUB_REPO=owner/yt-intel-bot
GITHUB_PAT=...              # from `gh auth token` (repo + workflow scopes)
GITHUB_BRANCH=main
LLM_API_KEY=                # optional
```

GitHub side: private repo with `scrape_channels.yml` active and the
`TELEGRAM_BOT_TOKEN` secret set.

## Run

Double-click **`start_bot.bat`**, then from Telegram:

- One link at a time → runs locally with the live Thinking card.
- `/batch https://www.youtube.com/@ChannelA https://youtu.be/VIDEO_ID …` or paste
  2+ links → the bot replies *"🚀 Cloud Batch Dispatched to GitHub Actions
  (X channels, Y standalone videos)! You can safely shut down your laptop now…"*
  and triggers the workflow with your chat id. The cloud runner then posts the
  live Thinking card, `Master_Channels_Comparison.xlsx` and
  `All_Channels_Categorized.zip` to the same chat.

The bot must be running when you send `/batch`; the cloud job survives your
laptop shutting down. The workbook + ZIP are also kept as GitHub Actions
artifacts.

## Repo layout

| File | Purpose |
| --- | --- |
| `bot.py` | Local Telegram bot: video/Shorts/channel scrapes + mixed `/batch` dispatch |
| `batch_runner.py` | Cloud runner: mixed-link router, transcripts/comments, category folders, master workbook, ZIP, live card |
| `categorizer.py` | Free weighted niche categorizer + optional LLM refinement |
| `progress.py` | Shared expandable thinking-card renderer |
| `.github/workflows/scrape_channels.yml` | `workflow_dispatch` batch (`extra_channels`, `chat_id`) + artifact backup |
| `channels.txt` | Optional local queue for `batch_runner.py` |
| `.env.example` | Template of the environment variables |
| `start_bot.bat` | Double-click launcher |
| `requirements.txt` | Python dependencies |

## Troubleshooting

- **Cloud dispatch failed** — `GITHUB_PAT` needs `repo` + `workflow` scopes and
  `GITHUB_REPO` must be `<owner>/yt-intel-bot`. Regenerate with `gh auth token`.
- **Workflow not triggering** — check `gh workflow list` shows `Scrape Channels`
  active and the push went to `main` (`GITHUB_BRANCH`).
- **A video has no transcript** — it is recorded as "Transcript unavailable" and
  the batch continues.
- **Some videos skipped** — retried once, then listed in the Thinking trace; the
  batch never stops for one bad item.
- **No live updates** — edits are throttled to one per 3 s; tap the expandable
  block to re-read the trace.
