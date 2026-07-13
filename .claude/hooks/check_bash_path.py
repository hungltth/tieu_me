#!/usr/bin/env python3
"""
PreToolUse hook: chặn lệnh Bash truy cập ra ngoài thư mục dự án.
"""
import json
import re
import sys

PROJECT_DIR = "/Users/hung/Projects/private/tieu_me"
HOME_DIR = "/Users/hung"

# Các system path luôn được phép (chỉ đọc/thực thi, không ghi)
ALLOWED_PREFIXES = [
    "/usr/",
    "/bin/",
    "/sbin/",
    "/etc/",
    "/tmp/",
    "/var/",
    "/opt/",
    "/System/",
    "/Library/",
    "/Applications/",
    "/private/",
    "/Users/hung/.claude/channels/telegram/",  # Telegram inbox (đọc file đính kèm)
    PROJECT_DIR,
]


def expand_path(path: str) -> str:
    if path.startswith("~/"):
        return HOME_DIR + path[1:]
    if path == "~":
        return HOME_DIR
    return path


def is_allowed_path(path: str) -> bool:
    path = expand_path(path)
    return any(path.startswith(prefix) for prefix in ALLOWED_PREFIXES)


def find_violations(command: str) -> list[str]:
    violations = []

    # Tìm tất cả absolute paths trong lệnh
    abs_paths = re.findall(r'(?:^|[\s\'"`=])(\/[^\s\'"`;&|><()\[\]{}\\]+)', command)
    for path in abs_paths:
        path = path.rstrip("/.,")
        if path and not is_allowed_path(path):
            violations.append(path)

    # Kiểm tra cd đến thư mục ngoài project
    cd_targets = re.findall(r'\bcd\s+([^\s;&|]+)', command)
    for target in cd_targets:
        if target.startswith("/") or target.startswith("~"):
            if not is_allowed_path(target) and target not in violations:
                violations.append(target)

    return violations


def main():
    try:
        data = json.load(sys.stdin)
    except Exception:
        sys.exit(0)

    if data.get("tool_name") != "Bash":
        sys.exit(0)

    command = data.get("tool_input", {}).get("command", "")
    violations = find_violations(command)

    if violations:
        result = {
            "decision": "block",
            "reason": (
                f"Bị chặn: lệnh truy cập path ngoài thư mục dự án.\n"
                f"Path vi phạm: {', '.join(violations)}\n"
                f"Chỉ được phép thao tác trong: {PROJECT_DIR}"
            ),
        }
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(2)

    sys.exit(0)


if __name__ == "__main__":
    main()
