#!/usr/bin/env python3
"""
Post a message to a Teams "Workflows" webhook as an Adaptive Card.
Stdlib only — nothing to install. Reads the webhook URL from the environment,
so the secret never appears on the command line or in any settings file.

Usage (this is the *stable* command you allow once in Claude Code):
    python3 post_to_teams.py "Your message text here"
    python3 post_to_teams.py --title "Motivation des Tages" "Carpe diem."
    echo "piped message" | python3 post_to_teams.py     # reads stdin if no arg

Env:
    TEAMS_WEBHOOK_URL   (required)  the Workflows webhook URL
"""

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime
from zoneinfo import ZoneInfo

TZ = ZoneInfo("Europe/Zurich")


def build_payload(message: str, title: str) -> dict:
    today = datetime.now(TZ).strftime("%d.%m.%Y")
    body = [
        {"type": "TextBlock", "text": title, "weight": "Bolder",
         "size": "Large", "wrap": True},
        {"type": "TextBlock", "text": today, "isSubtle": True,
         "spacing": "None", "wrap": True},
    ]
    for line in message.split("\n"):
        if line.strip():
            body.append({"type": "TextBlock", "text": line, "wrap": True,
                         "spacing": "Small"})
    card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": body,
    }
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": card,
        }],
    }


def post(payload: dict, url: str) -> int:
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("message", nargs="?", help="message text (or pipe via stdin)")
    ap.add_argument("--title", default="Nachricht")
    args = ap.parse_args()

    message = args.message if args.message is not None else sys.stdin.read()
    message = message.strip()
    if not message:
        print("No message provided.", file=sys.stderr)
        return 1

    url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not url:
        print("ERROR: TEAMS_WEBHOOK_URL is not set.", file=sys.stderr)
        return 1

    try:
        status = post(build_payload(message, args.title), url)
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
