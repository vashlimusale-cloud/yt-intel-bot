"""Minimal MCP (Model Context Protocol) server over stdio for YT-Bot tools.

Works with Claude Desktop, Claude Code and Cursor: run ``python mcp_server.py``
and register it as a stdio MCP server (see PLUGIN_SETUP.md).

Tools:
  smart_query(url, request)        - any YouTube insight, lightweight modules
  analyze_youtube_video(url)       - metadata + transcript excerpt + comments + hook
  analyze_youtube_channel(url)     - profile, stats, avatar/banner, top videos
  find_similar_channels(url)       - 5-8 peer channels in the same niche
"""

from __future__ import annotations

import json
import logging
import sys
from typing import Any

import plugin_api

logging.basicConfig(level=logging.WARNING, format="%(asctime)s | %(levelname)-8s | %(message)s")

TOOLS = [
    {
        "name": "smart_query",
        "description": "Answer a natural-language request about a YouTube video/channel using only the lightweight data modules needed.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "YouTube video or channel URL"},
                "request": {"type": "string", "description": "What the user wants to know"},
            },
            "required": ["url"],
        },
    },
    {
        "name": "analyze_youtube_video",
        "description": "Video metadata, thumbnail, transcript excerpt, top comments, hook analysis and parent channel overview.",
        "inputSchema": {
            "type": "object",
            "properties": {"url": {"type": "string"}},
            "required": ["url"],
        },
    },
    {
        "name": "analyze_youtube_channel",
        "description": "Channel profile, subscribers, total views, avatar/banner URLs and top 5 videos.",
        "inputSchema": {
            "type": "object",
            "properties": {"url": {"type": "string"}},
            "required": ["url"],
        },
    },
    {
        "name": "find_similar_channels",
        "description": "Discover 5-8 peer channels in the same niche as a video/channel.",
        "inputSchema": {
            "type": "object",
            "properties": {"url": {"type": "string"}},
            "required": ["url"],
        },
    },
]


def _call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    url = str(arguments.get("url") or "").strip()
    if not url:
        raise ValueError("missing 'url' argument")
    if name == "smart_query":
        result = plugin_api.smart_query_payload(url, str(arguments.get("request") or "analyse this"))
    elif name == "analyze_youtube_video":
        result = plugin_api.video_payload(url)
    elif name == "analyze_youtube_channel":
        result = plugin_api.channel_payload(url)
    elif name == "find_similar_channels":
        result = plugin_api.similar_payload(url)
    else:
        raise ValueError(f"unknown tool: {name}")
    return {
        "content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False)}],
        "isError": False,
    }


def handle(request: dict[str, Any]) -> dict[str, Any] | None:
    method = request.get("method")
    request_id = request.get("id")
    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": "yt-intel-bot", "version": "1.0.0"},
            },
        }
    if method in ("notifications/initialized", "notifications/cancelled"):
        return None
    if method == "ping":
        return {"jsonrpc": "2.0", "id": request_id, "result": {}}
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": request_id, "result": {"tools": TOOLS}}
    if method == "tools/call":
        params = request.get("params") or {}
        name = str(params.get("name") or "")
        arguments = params.get("arguments") or {}
        try:
            payload = _call_tool(name, arguments)
            return {"jsonrpc": "2.0", "id": request_id, "result": payload}
        except Exception as exc:
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": {
                    "content": [{"type": "text", "text": f"Error: {str(exc)[:300]}"}],
                    "isError": True,
                },
            }
    return {
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def main() -> int:
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except ValueError:
            continue
        response = handle(request)
        if response is not None and request.get("id") is not None:
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
