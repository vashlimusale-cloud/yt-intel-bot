# YT-Bot Plugin Setup — ChatGPT Custom GPT + Claude MCP

The bot exposes the same intelligence through:

- **HTTP / OpenAPI 3.1** — for ChatGPT Custom GPT *Actions* (and any REST client)
- **MCP over stdio** — for Claude Desktop, Claude Code and Cursor

Both return lightweight JSON (transcript excerpt ~900 chars, top 20 comments,
top videos) so responses stay well under ChatGPT's ~50KB action limit.

## 0. Deploy / URL

- Local: `http://127.0.0.1:10000` (or your `PORT`)
- Render: `https://<your-service>.onrender.com`
- Optional auth: set `PLUGIN_API_KEY` in the same environment as the bot, and
  send it as the `X-API-Key` header. The OpenAPI schema automatically advertises
  the `ApiKeyAuth` security scheme.
- Optional `PUBLIC_BASE_URL` — the public URL advertised in `/openapi.json`
  (on Render this is auto-detected from `RENDER_EXTERNAL_URL`).

## 1. ChatGPT Custom GPT Actions

1. Open <https://chatgpt.com/gpts/editor> → **Create / Configure**.
2. **Actions → Create new action**.
3. **Schema**: paste the contents of `https://<your-service>.onrender.com/openapi.json`
   (or import from URL).
4. **Authentication** (only if `PLUGIN_API_KEY` is set):
   - Auth Type: **API Key**
   - Header name: `X-API-Key`
   - Value: your `PLUGIN_API_KEY`
5. Save. The GPT can now call:
   - `smartQuery` — `POST /api/smart_query` with `{"url": "...", "request": "..."}`
   - `videoIntelligence` — `POST /api/video`
   - `channelOverview` — `POST /api/channel`
   - `similarChannels` — `POST /api/similar`
6. Example instructions for the GPT: *"Use smartQuery for any YouTube URL; pass
   the user's exact goal in `request` so the API only fetches what is needed."*

## 2. Claude Desktop

Add to `claude_desktop_config.json` (`%APPDATA%\Claude\` on Windows,
`~/Library/Application Support/Claude/` on macOS):

```json
{
  "mcpServers": {
    "yt-intel-bot": {
      "command": "python",
      "args": ["C:\\\\Users\\\\premm\\\\Desktop\\\\YT-Bot\\\\mcp_server.py"],
      "env": {
        "GEMINI_API_KEY": "your-key-optional"
      }
    }
  }
}
```

Restart Claude Desktop → the tools `smart_query`, `analyze_youtube_video`,
`analyze_youtube_channel`, `find_similar_channels` appear in the tools menu.

## 3. Claude Code

```bash
claude mcp add yt-intel-bot -- python "C:\Users\premm\Desktop\YT-Bot\mcp_server.py"
```

## 4. Cursor

`.cursor/mcp.json` in your project (or Settings → MCP → Add):

```json
{
  "mcpServers": {
    "yt-intel-bot": {
      "command": "python",
      "args": ["C:\\\\Users\\\\premm\\\\Desktop\\\\YT-Bot\\\\mcp_server.py"]
    }
  }
}
```

## 5. Quick manual tests

```bash
# health + schema
curl http://127.0.0.1:10000/health
curl http://127.0.0.1:10000/openapi.json

# smart query
curl -X POST http://127.0.0.1:10000/api/smart_query \
  -H "Content-Type: application/json" \
  -d '{"url":"https://youtu.be/dQw4w9WgXcQ","request":"hook analysis and top comments"}'

# MCP smoke test (JSON-RPC over stdio)
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | python mcp_server.py
```
