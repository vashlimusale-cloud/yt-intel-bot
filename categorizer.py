"""Smart auto-niche channel categorization for YT-Bot.

Primary mode is a free, deterministic weighted keyword/metadata scoring engine
(bio, channel tags, YouTube category, recent titles, community posts and external
links). It returns the category, a refined sub-niche label, a confidence score
and human-readable reasoning for the live "Agent Thinking" trace.

Optional LLM enhancement: when ``LLM_API_KEY`` is set (OpenRouter, DeepSeek or
any OpenAI-compatible endpoint), the exact category/sub-niche is refined with an
API call and the code falls back cleanly to the rule-based result on any error.
"""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import urlparse

import requests

logger = logging.getLogger("yt-bot.categorizer")

CATEGORIES = (
    "Documentary & Storytelling",
    "Tech & AI",
    "Finance & Business",
    "Podcasts & Interviews",
    "Education & Explainers",
    "Gaming",
    "Lifestyle & Vlogs",
    "General",
)

# Weight of one keyword hit, by where it was found.
SOURCE_WEIGHTS = {
    "name": 1.5,
    "keywords": 3.0,
    "description": 2.0,
    "titles": 1.0,
    "posts": 0.75,
    "links": 2.0,
}

MIN_CATEGORY_SCORE = 3.5  # below this the channel is classified as General

CATEGORY_KEYWORDS: dict[str, tuple[str, ...]] = {
    "Documentary & Storytelling": (
        "documentary", "storytelling", "story", "stories", "true crime", "investigation",
        "investigative", "history", "historical", "biography", "archive", "archival",
        "unsolved", "mystery", "real life", "human stories", "cinematic", "narrative",
        "video essay", "deep dive", "case study", "profile", "journey", "witness",
    ),
    "Tech & AI": (
        "tech", "technology", "ai", "artificial intelligence", "machine learning",
        "chatgpt", "openai", "llm", "chatbot", "gadget", "gadgets", "smartphone",
        "iphone", "android", "laptop", "hardware", "software", "coding", "programming",
        "developer", "python", "javascript", "robotics", "robot", "cybersecurity",
        "cyber security", "data science", "automation", "gpu", "cpu", "nvidia",
        "benchmark", "unboxing", "review", "macbook", "gadget review",
    ),
    "Finance & Business": (
        "finance", "financial", "money", "investing", "investment", "investor", "stock",
        "stocks", "stock market", "crypto", "cryptocurrency", "bitcoin", "ethereum",
        "trading", "trader", "entrepreneur", "entrepreneurship", "startup", "startups",
        "business", "ecommerce", "marketing", "sales", "economy", "economics", "wealth",
        "passive income", "side hustle", "real estate", "portfolio", "dividend",
        "recession", "inflation", "billionaire", "founder",
    ),
    "Podcasts & Interviews": (
        "podcast", "podcasts", "interview", "interviews", "conversation",
        "conversations", "episode", "episodes", "guest", "guests", "talk show",
        "fireside", "roundtable", "q&a", "sit down", "hosted by", "dialogue",
    ),
    "Education & Explainers": (
        "education", "educational", "explainer", "explained", "explain", "tutorial",
        "tutorials", "lecture", "course", "courses", "lesson", "lessons", "learn",
        "learning", "teach", "teacher", "science", "physics", "chemistry", "biology",
        "mathematics", "math", "study", "academy", "university", "school", "guide",
        "how to", "tips", "knowledge", "classroom", "revision", "exam", "student",
    ),
    "Gaming": (
        "gaming", "game", "games", "gameplay", "gamer", "gamers", "playthrough",
        "lets play", "walkthrough", "speedrun", "minecraft", "fortnite", "roblox",
        "gta", "valorant", "league of legends", "counter strike", "call of duty",
        "esports", "fps", "rpg", "mmorpg", "nintendo", "playstation", "xbox", "steam",
        "twitch", "controller", "boss fight", "level up", "loot",
    ),
    "Lifestyle & Vlogs": (
        "vlog", "vlogs", "daily life", "day in the life", "lifestyle", "travel",
        "vlogger", "food", "cooking", "recipe", "recipes", "fitness", "workout", "gym",
        "beauty", "makeup", "fashion", "outfit", "family", "couple", "relationship",
        "routine", "morning routine", "home", "house tour", "decor", "diy", "wellness",
        "health", "self care", "motivation", "personal", "shopping", "haul", "pet",
        "dog", "cat", "wedding", "birthday",
    ),
}

SUB_NICHES: dict[str, dict[str, tuple[str, ...]]] = {
    "Documentary & Storytelling": {
        "True crime & investigations": (
            "true crime", "investigation", "investigative", "unsolved", "detective",
            "forensic", "mystery", "case",
        ),
        "History & biography": (
            "history", "historical", "biography", "archive", "ancient", "war", "decades",
        ),
        "Cinematic essays & deep dives": (
            "video essay", "essay", "cinematic", "narrative", "deep dive", "documentary",
        ),
        "Social & human stories": ("human stories", "real life", "society", "community"),
    },
    "Tech & AI": {
        "AI tools & models": (
            "ai", "artificial intelligence", "chatgpt", "openai", "llm", "machine learning",
            "prompt", "model",
        ),
        "Gadget reviews & unboxings": (
            "review", "gadget", "iphone", "smartphone", "laptop", "unboxing", "hardware",
            "benchmark",
        ),
        "Programming & dev": (
            "coding", "programming", "developer", "python", "javascript", "software",
            "github", "api",
        ),
        "Cybersecurity & privacy": (
            "security", "cyber", "hacking", "privacy", "vpn", "malware",
        ),
    },
    "Finance & Business": {
        "Investing & markets": (
            "investing", "investment", "stock", "stocks", "trading", "crypto", "bitcoin",
            "portfolio", "dividend",
        ),
        "Entrepreneurship & startups": (
            "entrepreneur", "startup", "founder", "business", "ecommerce", "marketing",
            "sales",
        ),
        "Personal finance": (
            "money", "savings", "budget", "passive income", "side hustle", "wealth",
            "financial",
        ),
    },
    "Podcasts & Interviews": {
        "Long-form interviews": ("interview", "conversation", "guest", "sit down", "talk"),
        "Podcast episodes": ("podcast", "episode", "episodes", "series", "hosted by"),
        "Roundtables & Q&A": ("roundtable", "panel", "fireside", "q&a", "discussion"),
    },
    "Education & Explainers": {
        "Science explainers": ("science", "physics", "chemistry", "biology", "space", "quantum"),
        "How-to & tutorials": ("tutorial", "how to", "guide", "tips", "step by step", "learn"),
        "Study & exam prep": ("exam", "revision", "study", "math", "mathematics", "lecture", "course"),
    },
    "Gaming": {
        "Gameplay & let's play": (
            "gameplay", "playthrough", "lets play", "walkthrough", "speedrun", "stream",
        ),
        "Game reviews & news": ("game review", "patch", "update", "tier list", "gaming news"),
        "Specific titles": (
            "minecraft", "fortnite", "roblox", "gta", "valorant", "call of duty",
            "league of legends",
        ),
    },
    "Lifestyle & Vlogs": {
        "Daily vlogs & routines": ("vlog", "daily", "routine", "day in the life", "morning"),
        "Travel & food": ("travel", "food", "cooking", "recipe", "street food", "trip"),
        "Fitness & wellness": ("fitness", "workout", "gym", "health", "wellness", "weight loss"),
        "Fashion & beauty": ("fashion", "beauty", "makeup", "outfit", "haul", "style"),
    },
    "General": {},
}

# YouTube's own category names, mapped to the strongest matching niche.
YOUTUBE_CATEGORY_MAP = {
    "education": "Education & Explainers",
    "science & technology": "Tech & AI",
    "gaming": "Gaming",
    "howto & style": "Lifestyle & Vlogs",
    "people & blogs": "Lifestyle & Vlogs",
    "travel & events": "Lifestyle & Vlogs",
    "pets & animals": "Lifestyle & Vlogs",
    "film & animation": "Documentary & Storytelling",
    "news & politics": "Documentary & Storytelling",
    "nonprofits & activism": "Documentary & Storytelling",
}

_KEYWORD_PATTERNS: dict[str, re.Pattern[str]] = {}


def _pattern(keyword: str) -> re.Pattern[str]:
    pattern = _KEYWORD_PATTERNS.get(keyword)
    if pattern is None:
        escaped = re.escape(keyword)
        if " " in keyword:
            pattern = re.compile(escaped, re.IGNORECASE)
        else:
            pattern = re.compile(rf"(?<![a-z0-9]){escaped}(?![a-z0-9])", re.IGNORECASE)
        _KEYWORD_PATTERNS[keyword] = pattern
    return pattern


def _hits(keyword: str, text: str) -> int:
    if not text:
        return 0
    return len(_pattern(keyword).findall(text))


def _attr(obj: Any, name: str, default: Any = "") -> Any:
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def _titles(items: list[Any]) -> list[str]:
    titles = []
    for item in items:
        title = str(_attr(item, "title", "") or "").strip()
        if title:
            titles.append(title)
    return titles


def _post_texts(posts: list[Any]) -> list[str]:
    texts = []
    for post in posts:
        text = str(_attr(post, "text", "") or "").strip()
        if text:
            texts.append(text)
    return texts


def _yt_categories(items: list[Any]) -> list[str]:
    values: list[str] = []
    for item in items:
        raw = str(_attr(item, "category", "") or "")
        values.extend(part.strip() for part in raw.split(",") if part.strip())
    return values


def _link_hosts(links: list[Any]) -> str:
    hosts = []
    for link in links:
        url = link if isinstance(link, str) else str(_attr(link, "url", "") or "")
        if not url:
            continue
        try:
            hosts.append(urlparse(url).netloc.lower())
        except ValueError:
            continue
    return " ".join(hosts)


@dataclass
class ChannelSnapshot:
    name: str = ""
    description: str = ""
    keywords: list[str] = field(default_factory=list)
    links: list[str] = field(default_factory=list)
    videos: list[Any] = field(default_factory=list)
    shorts: list[Any] = field(default_factory=list)
    posts: list[Any] = field(default_factory=list)

    @property
    def titles(self) -> list[str]:
        return _titles(self.videos) + _titles(self.shorts)

    @property
    def category_texts(self) -> list[str]:
        return _yt_categories(self.videos) + _yt_categories(self.shorts)


@dataclass
class CategoryResult:
    category: str
    sub_niche: str
    confidence: float
    score: float
    reasoning: list[str] = field(default_factory=list)
    signals: list[str] = field(default_factory=list)
    source: str = "rules"  # "rules" or "llm"
    rules_category: str = ""
    rules_score: float = 0.0

    def as_dict(self) -> dict[str, Any]:
        return {
            "category": self.category,
            "sub_niche": self.sub_niche,
            "confidence": self.confidence,
            "score": self.score,
            "reasoning": self.reasoning,
            "signals": self.signals,
            "source": self.source,
            "rules_category": self.rules_category,
            "rules_score": self.rules_score,
        }


def build_snapshot(
    profile: Any, videos: list[Any], shorts: list[Any], posts: list[Any]
) -> ChannelSnapshot:
    """Build a snapshot from bot.ChannelProfile/ChannelItem/CommunityPost objects."""
    links = [
        str(_attr(link, "url", "") or "")
        for link in (_attr(profile, "links", []) or [])
    ]
    return ChannelSnapshot(
        name=str(_attr(profile, "name", "") or ""),
        description=str(_attr(profile, "description", "") or ""),
        keywords=[str(word) for word in (_attr(profile, "keywords", []) or [])],
        links=[link for link in links if link],
        videos=list(videos),
        shorts=list(shorts),
        posts=list(posts),
    )


def _rule_result(snapshot: ChannelSnapshot) -> CategoryResult:
    scores = {category: 0.0 for category in CATEGORIES}
    matched: dict[str, list[str]] = {category: [] for category in CATEGORIES}

    name_text = snapshot.name
    keywords_text = " ".join(snapshot.keywords)
    description_text = snapshot.description
    titles = snapshot.titles
    posts = _post_texts(snapshot.posts)
    links_text = _link_hosts(snapshot.links)

    for category, words in CATEGORY_KEYWORDS.items():
        for word in words:
            found = False
            if _hits(word, name_text):
                scores[category] += SOURCE_WEIGHTS["name"]
                found = True
            if _hits(word, keywords_text):
                scores[category] += SOURCE_WEIGHTS["keywords"]
                found = True
            if _hits(word, description_text):
                scores[category] += SOURCE_WEIGHTS["description"]
                found = True
            if _hits(word, links_text):
                scores[category] += SOURCE_WEIGHTS["links"]
                found = True
            title_hits = sum(1 for title in titles if _hits(word, title))
            if title_hits:
                scores[category] += min(title_hits, 6) * SOURCE_WEIGHTS["titles"]
                found = True
            post_hits = sum(1 for text in posts if _hits(word, text))
            if post_hits:
                scores[category] += min(post_hits, 4) * SOURCE_WEIGHTS["posts"]
                found = True
            if found:
                matched[category].append(word)

    yt_counts: dict[str, int] = {}
    for raw in snapshot.category_texts:
        key = raw.lower()
        yt_counts[key] = yt_counts.get(key, 0) + 1
    for raw, count in sorted(yt_counts.items()):
        target = YOUTUBE_CATEGORY_MAP.get(raw)
        if target:
            scores[target] += min(count * 0.6, 9.0)
            matched[target].append(f"YouTube category '{raw.title()}' (x{count})")

    ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    top_category, top_score = ranked[0]
    second_category, second_score = ranked[1]

    if top_score < MIN_CATEGORY_SCORE:
        confidence = 0.35
        reasoning = [
            f"No strong niche signal found (best guess {top_category}, score {top_score:.1f}) "
            "— defaulting to General"
        ]
        return CategoryResult(
            category="General",
            sub_niche="General interest",
            confidence=confidence,
            score=round(top_score, 1),
            reasoning=reasoning,
            signals=[],
            source="rules",
            rules_category="General",
            rules_score=round(top_score, 1),
        )

    confidence = round(min(0.95, top_score / (top_score + second_score + 1)), 2)
    keyword_signals = [word for word in matched[top_category] if not word.startswith("YouTube category")]
    yt_signals = [word for word in matched[top_category] if word.startswith("YouTube category")]

    reasoning = []
    if keyword_signals:
        reasoning.append(
            f"Matched {len(keyword_signals)} {top_category.lower()} signals in bio, keywords "
            f"and titles: {', '.join(keyword_signals[:6])}"
        )
    if yt_signals:
        reasoning.append(f"YouTube metadata backs this up: {yt_signals[0]}")
    if second_score:
        reasoning.append(f"Runner-up: {second_category} (score {second_score:.1f})")

    sub_niche, sub_reason = _pick_sub_niche(top_category, snapshot)
    reasoning.extend(sub_reason)

    return CategoryResult(
        category=top_category,
        sub_niche=sub_niche,
        confidence=confidence,
        score=round(top_score, 1),
        reasoning=reasoning,
        signals=matched[top_category][:10],
        source="rules",
        rules_category=top_category,
        rules_score=round(top_score, 1),
    )


def _pick_sub_niche(category: str, snapshot: ChannelSnapshot) -> tuple[str, list[str]]:
    rules = SUB_NICHES.get(category) or {}
    if not rules:
        return "General interest", []

    corpus = " ".join(
        [
            snapshot.name,
            snapshot.description,
            " ".join(snapshot.keywords),
            *snapshot.titles,
            *_post_texts(snapshot.posts),
        ]
    )
    best_label = ""
    best_score = 0
    for label, words in rules.items():
        score = sum(_hits(word, corpus) for word in words)
        if score > best_score:
            best_label, best_score = label, score
    if best_label:
        return best_label, [f"Sub-niche '{best_label}' matched {best_score} signal(s)"]
    return f"General {category.lower()}", []


# --------------------------------------------------------------------------- #
# Optional LLM refinement
# --------------------------------------------------------------------------- #
def _parse_llm_json(content: str) -> dict[str, Any] | None:
    text = (content or "").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.IGNORECASE | re.MULTILINE).strip()
    start = text.find("{")
    if start == -1:
        return None
    depth = 0
    in_string = False
    escaped = False
    for index in range(start, len(text)):
        char = text[index]
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
        elif char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(text[start:index + 1])
                except json.JSONDecodeError:
                    return None
    return None


def _llm_refine(snapshot: ChannelSnapshot, rules: CategoryResult) -> CategoryResult | None:
    """Ask an OpenAI-compatible model to refine the category/sub-niche."""
    api_key = (os.getenv("LLM_API_KEY") or "").strip()
    if not api_key:
        return None
    base_url = (os.getenv("LLM_BASE_URL") or "https://openrouter.ai/api/v1").rstrip("/")
    model = (os.getenv("LLM_MODEL") or "openai/gpt-4o-mini").strip()
    titles = snapshot.titles[:24]

    prompt = (
        "Classify this YouTube channel. Reply with ONLY compact JSON: "
        '{"category": one of ' + json.dumps(list(CATEGORIES)) + ', '
        '"sub_niche": "max 5 words", "confidence": 0.0-1.0, "reason": "max 20 words"}.\n\n'
        f"Channel name: {snapshot.name}\n"
        f"Description: {snapshot.description[:600]}\n"
        f"Channel tags: {', '.join(snapshot.keywords[:20])}\n"
        f"Recent titles:\n- " + "\n- ".join(titles) +
        f"\n\nRule-based guess: {rules.category} / {rules.sub_niche} (score {rules.score})"
    )
    payload = {
        "model": model,
        "temperature": 0,
        "max_tokens": 250,
        "messages": [
            {"role": "system", "content": "You are a precise YouTube niche classifier."},
            {"role": "user", "content": prompt},
        ],
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    try:
        response = requests.post(
            f"{base_url}/chat/completions", json=payload, headers=headers, timeout=30
        )
        response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"]
    except Exception as exc:  # network/auth/parsing: fall back to rules
        logger.warning("LLM refinement unavailable, using rule-based category: %s", exc)
        return None

    parsed = _parse_llm_json(str(content))
    if not parsed:
        logger.warning("LLM returned unparseable classification; using rules")
        return None

    category = str(parsed.get("category") or "").strip()
    if category not in CATEGORIES:
        category = rules.category
    sub_niche = str(parsed.get("sub_niche") or rules.sub_niche).strip()[:80]
    try:
        confidence = float(parsed.get("confidence", rules.confidence))
    except (TypeError, ValueError):
        confidence = rules.confidence
    confidence = max(0.0, min(1.0, confidence))
    reason = str(parsed.get("reason") or "").strip()

    reasoning = list(rules.reasoning)
    reasoning.append(
        "LLM refinement: " + (reason or f"{category} / {sub_niche}")
    )
    return CategoryResult(
        category=category,
        sub_niche=sub_niche or rules.sub_niche,
        confidence=round(confidence, 2),
        score=rules.score,
        reasoning=reasoning,
        signals=rules.signals,
        source="llm",
        rules_category=rules.category,
        rules_score=rules.score,
    )


def categorize_channel(snapshot: ChannelSnapshot) -> CategoryResult:
    """Classify a channel into one of CATEGORIES (rule-based, LLM-refined if configured)."""
    rules = _rule_result(snapshot)
    if rules.category == "General" and not snapshot.titles and not snapshot.description:
        return rules
    refined = _llm_refine(snapshot, rules)
    return refined or rules
