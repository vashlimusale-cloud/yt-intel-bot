# YT-Bot — YouTube Channel Intelligence

Telegram bot that turns YouTube links into deep research dossiers with an
agent-style live progress card (expandable 🧠 thinking trace). It handles
**channels and individual videos in the same batch**, keeps real `.jpg`
avatars/thumbnails on disk, and runs 24/7 on Render.com while heavy scraping is
offloaded to GitHub Actions.

## What it does

| You send | You get |
| --- | --- |
| Video / Shorts URL | Thumbnail + Markdown dossier (metadata, description, timestamped transcript, every comment) |
| One channel URL | Channel avatar photo + Excel `Channel_Intel` workbook + Markdown SEO audit (local mode) |
| `/batch` + any mix of channel/video links, or 2+ links in one message | Full cloud run on GitHub Actions: live Thinking card, per-channel avatars sent to chat, `Master_Channels_Comparison.xlsx` + `All_Channels_Categorized.zip` |
| `/start` | Help card |

On Render (or with `CLOUD_SINGLE_LINKS=true`) even a single link is scraped on
GitHub Actions, so the 512 MB Render free instance only routes messages and the
heavy scraping always gets runner RAM/disk.

## Transcripts that survive datacenter IPs

`fetch_transcript()` tries **youtube-transcript-api** first. When it fails or
returns empty (common on GitHub/Render datacenter IPs), it automatically:

1. reads the caption tracks already present in the yt-dlp `extract_info` result
   (`subtitles` / `automatic_captions`), prioritizing `hi`, `en`, `en-IN`,
   `en-US`, `en-GB`, then any `hi`/`en` track, then any track with a `json3` or
   `vtt` format;
2. downloads that caption track and parses it into clean, deduplicated
   `[MM:SS] text` lines;
3. if `extract_info` had no usable subtitle URL, runs a targeted yt-dlp
   extraction with
   `extractor_args={"youtube": {"player_client": ["android", "ios", "tv_embedded", "web"]}}`
   and tries again.

## Cloud batch outputs (real files, zipped)

```
output/
  Master_Channels_Comparison.xlsx      # All_Channels_Overview + one tab per
                                       # category + Standalone_Videos + Top_Videos_Combined
  All_Channels_Categorized.zip         # everything below, zipped
  batch_result/
    <Category>/                        # "Tech & AI", "Gaming", …
      <Channel_Name>/
        <Channel_Name>_Full_Intel.xlsx # Channel_Overview, Videos, Shorts, Community_Posts
                                       #   Videos/Shorts have Thumbnail_URL, Thumbnail_File,
                                       #   Full_Transcript, Top_Comments columns
        <Channel_Name>_Full_Intel.md   # profile + every item's metadata, local thumbnail,
                                       #   transcript and top comments
        <Channel_Name>_avatar.jpg      # real channel profile picture
        thumbnails/
          <index>_<Safe_Title>_<video_id>.jpg
    Standalone_Videos/
      <Title>_<video_id>_Full_Intel.md # metadata, thumbnail, full transcript, ALL comments
      thumbnails/<index>_<Safe_Title>_<video_id>.jpg
```

Every channel avatar is also **sent as a Telegram photo** when that channel
finishes. Thumbnails use `maxresdefault.jpg` → `hqdefault.jpg` →
`info['thumbnail']`, rejecting YouTube's tiny (<2 KB) 404 placeholder.

Limits: `MAX_CHANNEL_VIDEOS = 100`, `MAX_CHANNEL_SHORTS = 100`,
`MAX_COMMENTS_PER_VIDEO = 25`, `MAX_COMMENTS_STANDALONE = 0` (unlimited). Polite
pacing: 1–2 s between per-video transcript/comment calls, 6–10 s between
channels.

## 24/7 hosting on Render.com

`render.yaml` defines a free Python **web** service:

- build: `pip install -r requirements.txt`
- start: `python bot.py`
- health check: `/health`

`bot.py` starts a tiny HTTP server bound to `$PORT` (default `10000`) answering
`200 OK` `{"status": "alive"}` on `/`, `/health` and `/ping` (GET + HEAD), then
runs Telegram polling in the main thread. Ping the URL with UptimeRobot to stop
Render's free instance from sleeping.

Render needs these environment variables (dashboard → Environment):

```
TELEGRAM_BOT_TOKEN
GITHUB_PAT
GITHUB_REPO
GITHUB_BRANCH=main
TELEGRAM_CHAT_ID
CLOUD_SINGLE_LINKS=true
```

Only one bot instance may poll a token: when the Render service is live, stop
the local `start_bot.bat`.

## GitHub Actions cloud batch

`/batch` (or 2+ links) calls the GitHub dispatch API with the links and your
chat id. The workflow (`workflow_dispatch` inputs `extra_channels`, `chat_id`)
runs `batch_runner.py` on `ubuntu-latest` / Python 3.11 with
`TELEGRAM_BOT_TOKEN` from repo secrets, then uploads the ZIP + workbook as
artifacts. Channels are processed with 6–10 s delays and every item with 1–2 s
delays, so ~150–250 videos finish without rate limits.

## Local setup

`.env` (auto-configured locally):

```
TELEGRAM_BOT_TOKEN=...
GITHUB_REPO=owner/yt-intel-bot
GITHUB_PAT=...            # from `gh auth token` (repo + workflow scopes)
GITHUB_BRANCH=main
TELEGRAM_CHAT_ID=...
CLOUD_SINGLE_LINKS=false  # local: scrape single links locally
```

Run: double-click `start_bot.bat`.

## Repo layout

| File | Purpose |
| --- | --- |
| `bot.py` | Telegram bot, caption/transcript fallback, health server, mixed `/batch` dispatch |
| `batch_runner.py` | Cloud runner: mixed-link router, transcripts/comments, avatars/thumbnails, category folders, master workbook, ZIP, live card |
| `categorizer.py` | Free weighted niche categorizer + optional LLM refinement |
| `progress.py` | Shared expandable thinking-card renderer |
| `render.yaml` | Render.com free web-service blueprint |
| `.github/workflows/scrape_channels.yml` | Cloud batch workflow + artifact backup |
| `channels.txt` | Optional local queue for `batch_runner.py` |
| `.env.example` | Template of the environment variables |
| `start_bot.bat` | Double-click launcher |

## Troubleshooting

- **Cloud dispatch failed** — `GITHUB_PAT` needs `repo` + `workflow` scopes,
  `GITHUB_REPO` must be `<owner>/yt-intel-bot`.
- **Workflow not triggering** — `gh workflow list` must show `Scrape Channels`
  active on `main`.
- **No transcript** — some videos have no captions at all; they are marked
  "Transcript unavailable" and the batch continues.
- **No live updates** — edits are throttled to one per 3 s; tap the expandable
  block to re-read the trace.
- **Bot silent on Render** — check the service logs, the `/health` URL, and that
  only one instance is polling the token.
