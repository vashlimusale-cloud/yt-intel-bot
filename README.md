# YT-Bot — YouTube Channel Intelligence

Telegram bot that turns YouTube links into research dossiers with an agent-style
live progress card (expandable 🧠 thinking trace), and can offload big channel
batches to GitHub Actions so you can close your laptop.

## What it does

| You send | You get |
| --- | --- |
| Video / Shorts URL | Thumbnail + Markdown dossier (metadata, description, timestamped transcript, every comment) |
| One channel URL | Channel avatar + Excel `Channel_Intel` workbook + Markdown SEO audit |
| `/batch` + channel links, or 2+ channel links in one message | Runs the whole batch in the cloud (GitHub Actions), then posts live progress + `Master_Channels_Comparison.xlsx` + `All_Channels_Categorized.zip` here |
| `/start` | Help card |

Every action shows the same **live status card** that is edited in place (max
one edit every 3 seconds) with a tappable `<blockquote expandable>` **🧠 Agent
Thinking & Live Trace** block: detections, counts, retries, cooldowns, elapsed
time. The card stays visible with the final ✅/❌ state after the run.

## Smart auto-niche categorization

Each channel is classified into one of:

`Documentary & Storytelling`, `Tech & AI`, `Finance & Business`,
`Podcasts & Interviews`, `Education & Explainers`, `Gaming`,
`Lifestyle & Vlogs`, `General`

The default engine is a free weighted scoring model over the channel bio,
channel tags, YouTube categories, recent video/short titles and community posts
(and it explains itself in the trace, e.g. *"Matched 6 tech & ai signals in bio,
keywords and titles: ai, chatgpt, llm, machine learning"*).

Optional AI refinement: set `LLM_API_KEY` (OpenRouter, DeepSeek or any
OpenAI-compatible endpoint; optionally `LLM_BASE_URL` / `LLM_MODEL`) and the
sub-niche label is refined by the model. Any LLM error falls back to the
rule-based result.

## Setup

1. `.env` already contains your `TELEGRAM_BOT_TOKEN`; the installer also writes
   `GITHUB_REPO`, `GITHUB_PAT` and `GITHUB_BRANCH`.
2. (Optional) add `LLM_API_KEY` for AI sub-niche refinement.
3. Double-click **`start_bot.bat`**.

## Cloud batch (GitHub Actions)

Requirements (already done by the setup): private repo `yt-intel-bot`,
`TELEGRAM_BOT_TOKEN` repo secret, workflow `scrape_channels.yml`.

Send:

```
/batch https://www.youtube.com/@MrBeast https://www.youtube.com/@mkbhd
```

The bot immediately replies *"🚀 Cloud Batch Dispatched to GitHub Actions for 2
channels! You can safely shut down your laptop now…"* and triggers the workflow
via the GitHub dispatch API with your channel list and chat id. The cloud runner
then:

1. audits every channel (profile, subscribers, total views, joined, country,
   email, external links, up to `MAX_CHANNEL_VIDEOS` videos and
   `MAX_CHANNEL_SHORTS` shorts, community posts, collaborations),
2. categorizes each channel and files its `.xlsx`/`.md` into
   `output/batch_result/<Category>/`,
3. builds `Master_Channels_Comparison.xlsx` (`All_Channels_Overview` +
   one tab per category + `Top_Videos_Combined` top 50),
4. zips everything into `All_Channels_Categorized.zip`, updates the live card to
   ✅ Completed and sends both files to your chat (120 s timeout + retry),
5. keeps the ZIP/workbook as a GitHub Actions artifact backup.

Channels are processed with 8–15 s randomized polite delays; failures are
retried once, then skipped and listed in the trace.

## Output

- Local single links: `output/<Channel>_Channel_Intel.xlsx` / `.md`
- Cloud batch: `output/batch_result/<Category>/…`,
  `output/Master_Channels_Comparison.xlsx`,
  `output/All_Channels_Categorized.zip`

Nothing in `output/` is committed (see `.gitignore`); `.env` never leaves your
machine, and the token only lives in the repo's GitHub Secret.

## Repo layout

| File | Purpose |
| --- | --- |
| `bot.py` | Local Telegram bot: video/Shorts/channel scrapes + `/batch` dispatch |
| `batch_runner.py` | Cloud runner: queue, categorized folders, master workbook, ZIP, live card |
| `categorizer.py` | Free weighted niche categorizer + optional LLM refinement |
| `progress.py` | Shared expandable `<blockquote>` thinking-card renderer |
| `.github/workflows/scrape_channels.yml` | `workflow_dispatch` cloud batch with `extra_channels` + `chat_id` |
| `channels.txt` | Optional local channel queue for `batch_runner.py` |
| `.env.example` | Template of the required environment variables |
| `start_bot.bat` | Double-click launcher |
| `requirements.txt` | Python dependencies |

## Troubleshooting

- **Cloud dispatch failed** — `GITHUB_PAT` needs `repo` + `workflow` scopes and
  `GITHUB_REPO` must be `<owner>/yt-intel-bot`. Regenerate with `gh auth token`.
- **Workflow not triggering** — check `gh workflow list` shows
  `Scrape Channels` active; the bot uses the `main` branch (`GITHUB_BRANCH`).
- **No live updates** — Telegram edits are throttled to one per 3 s; tap the
  expandable block to re-read the trace.
- **A channel fails** — it is retried once, then skipped; the trace lists the
  reason and the batch continues.
