#!/usr/bin/env python3
"""
Post a message to a Teams "Workflows" webhook as a styled Adaptive Card.
Stdlib only — nothing to install. Webhook URL comes from the environment.

Two input modes (auto-detected):

1. STRUCTURED (recommended for Newswatch): pass JSON describing topics.
   Each topic gets its own tinted container with a number badge, a bold
   title, and labelled "Warum jetzt" / "Storytelling-Form" rows.

   JSON schema:
   {
     "title": "Newswatch — Themenvorschläge",
     "subtitle": "22. Juni – 6. Juli 2026",
     "topics": [
       {"tag": "Politik", "title": "...", "why": "...", "form": "..."}
     ],
     "sources": ["https://...", "https://..."]
   }

2. PLAIN: any non-JSON text is posted as a simple card (backward compatible).

Usage (command shapes stay stable, so one allow rule covers them):
   python3 post_to_teams.py --json topics.json
   cat topics.json | python3 post_to_teams.py --json -
   python3 post_to_teams.py "a plain message"

Env:
   TEAMS_WEBHOOK_URL   (required)
"""

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from datetime import datetime
from zoneinfo import ZoneInfo

TZ = ZoneInfo("Europe/Zurich")

# A text that begins with "22. " is parsed by Teams as an ordered-list item and
# renumbered to "1.". Inserting a zero-width space between the number and the
# delimiter breaks that detection while staying visually identical.
_LEADING_LIST_MARKER = re.compile(r"^(\s*\d+)([.)])(\s)")


def md_safe(text):
    if not text:
        return text
    return _LEADING_LIST_MARKER.sub("\\1\u200b\\2\\3", text)


def _tb(text, **kw):
    block = {"type": "TextBlock", "text": text, "wrap": True}
    block.update(kw)
    return block


def header_blocks(title, subtitle):
    blocks = [_tb(md_safe(title), weight="Bolder", size="Large")]
    if subtitle:
        blocks.append(_tb(md_safe(subtitle), isSubtle=True, spacing="None"))
    return blocks


def topic_container(index, topic):
    items = []
    tag = topic.get("tag")
    if tag:
        items.append(_tb(tag.upper(), size="Small", weight="Bolder",
                         color="Accent", spacing="None"))
    items.append({
        "type": "ColumnSet",
        "spacing": "Small",
        "columns": [
            {"type": "Column", "width": "auto", "verticalContentAlignment": "Center",
             "items": [_tb(str(index), size="ExtraLarge", weight="Bolder",
                           color="Accent", spacing="None")]},
            {"type": "Column", "width": "stretch", "verticalContentAlignment": "Center",
             "items": [_tb(md_safe(topic.get("title", "")), weight="Bolder", size="Medium")]},
        ],
    })
    if topic.get("why"):
        items.append(_tb(f"**Warum jetzt:** {topic['why']}", spacing="Small"))
    if topic.get("form"):
        items.append(_tb(f"**Storytelling-Form:** {topic['form']}", spacing="Small"))
    return {
        "type": "Container",
        "style": "emphasis",
        "separator": True,
        "spacing": "Medium",
        "items": items,
    }


def sources_block(sources):
    links = " · ".join(f"[{i + 1}]({u})" for i, u in enumerate(sources))
    return _tb(f"**Quellen:** {links}", isSubtle=True, size="Small",
               separator=True, spacing="Medium")


def build_structured(data):
    today = datetime.now(TZ).strftime("%d.%m.%Y")
    body = header_blocks(data.get("title", "Update"), data.get("subtitle") or today)
    for i, topic in enumerate(data.get("topics", []), start=1):
        body.append(topic_container(i, topic))
    if data.get("sources"):
        body.append(sources_block(data["sources"]))
    return body


def build_plain(message, title):
    body = header_blocks(title, datetime.now(TZ).strftime("%d.%m.%Y"))
    for line in message.split("\n"):
        if line.strip():
            body.append(_tb(md_safe(line), spacing="Small"))
    return body


def wrap(body):
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.4",
                "body": body,
                "msteams": {"width": "Full"},
            },
        }],
    }


def post(payload, url):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status


def read_input(args):
    if args.json:
        raw = sys.stdin.read() if args.json == "-" else open(args.json, encoding="utf-8").read()
        return json.loads(raw)
    if args.message is not None:
        return args.message
    return sys.stdin.read()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("message", nargs="?", help="plain message (or pipe via stdin)")
    ap.add_argument("--json", help="path to a JSON topics file, or '-' for stdin")
    ap.add_argument("--title", default="Update")
    ap.add_argument("--dry-run", action="store_true", help="print payload, don't post")
    args = ap.parse_args()

    content = read_input(args)
    if isinstance(content, dict):
        body = build_structured(content)
    else:
        text = (content or "").strip()
        if not text:
            print("No content provided.", file=sys.stderr)
            return 1
        body = build_plain(text, args.title)

    payload = wrap(body)

    if args.dry_run:
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not url:
        print("ERROR: TEAMS_WEBHOOK_URL is not set.", file=sys.stderr)
        return 1
    try:
        status = post(payload, url)
    except urllib.error.HTTPError as e:
        print(f"FAILED — HTTP {e.code}: {e.read()[:500]!r}", file=sys.stderr)
        return 1
    except urllib.error.URLError as e:
        print(f"FAILED — {e.reason}", file=sys.stderr)
        return 1
    print(f"OK — HTTP {status}. Posted to Teams.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
