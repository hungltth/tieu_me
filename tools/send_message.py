#!/usr/bin/env python3
import sys
import os
import urllib.request
import urllib.parse
import json
from pathlib import Path


def load_env(path: str) -> dict:
    env = {}
    with open(path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, val = line.partition("=")
                env[key.strip()] = val.strip()
    return env


def send_telegram(bot_token: str, chat_id: str, text: str) -> None:
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    data = urllib.parse.urlencode({"chat_id": chat_id, "text": text}).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req) as resp:
        result = json.load(resp)
    if not result.get("ok"):
        raise RuntimeError(f"Telegram API error: {result}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python send_message.py \"Nội dung cần gửi\"")
        sys.exit(1)

    message = " ".join(sys.argv[1:])

    env_path = Path(__file__).parent.parent / ".env"
    env = load_env(env_path)

    bot_token = env.get("CLAUDE_BOT_API")
    chat_id = env.get("MY_USER_ID")

    if not bot_token or not chat_id:
        print("Lỗi: thiếu CLAUDE_BOT_API hoặc MY_USER_ID trong .env")
        sys.exit(1)

    send_telegram(bot_token, chat_id, message)
    print(f"Đã gửi: {message}")


if __name__ == "__main__":
    main()
