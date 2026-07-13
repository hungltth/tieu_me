#!/usr/bin/env python3
"""
Bot Server - Telegram bot thực thi lệnh từ người dùng được phép.
"""

import json
import logging
import shlex
import subprocess
import sys
from pathlib import Path

from telegram import Update, BotCommand
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"Config file not found: {CONFIG_FILE}")
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


config = load_config()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=getattr(logging, config.get("log_level", "INFO")),
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Access control
# ---------------------------------------------------------------------------


def normalize_ids(id_list: list) -> list[int]:
    """Chuyển list id (có thể string hoặc int) về list int."""
    result = []
    for x in id_list:
        try:
            result.append(int(x))
        except (ValueError, TypeError):
            pass
    return result


def is_allowed(update: Update) -> bool:
    """Kiểm tra user/group có trong allowed list không."""
    cfg = load_config()
    user_id = update.effective_user.id if update.effective_user else None
    chat_id = update.effective_chat.id if update.effective_chat else None
    chat_type = update.effective_chat.type if update.effective_chat else None

    admin_users = normalize_ids(cfg.get("admin_users", []))
    allowed_users = normalize_ids(cfg.get("allowed_users", []))
    allowed_groups = normalize_ids(cfg.get("allowed_groups", []))

    if user_id and user_id in admin_users:
        return True

    if chat_type in ("group", "supergroup"):
        return chat_id in allowed_groups

    return user_id in allowed_users


def is_admin(update: Update) -> bool:
    user_id = update.effective_user.id if update.effective_user else None
    cfg = load_config()
    return user_id in normalize_ids(cfg.get("admin_users", []))


async def deny(update: Update):
    await update.message.reply_text("⛔ Bạn không có quyền sử dụng bot này.")


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text(
        "👋 Bot server đang hoạt động!\nDùng /help để xem danh sách lệnh."
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return

    lines = [
        "📋 *Danh sách lệnh:*",
        "",
        "*/start* - Kiểm tra bot",
        "*/help* - Xem lệnh này",
        "*/id* - Lấy user ID và chat ID",
        "*/myid* - Lấy user ID và chat ID",
        "*/ping* - Kiểm tra kết nối",
        "*/baocao* - Liệt kê các file báo cáo",
        "*/check_log* - Chạy kiểm tra log hàng ngày",
        "*/primes <number>* - Tính số nguyên tố (ví dụ: /primes 100)",
        '*/tra_cuu <file> <keyword>* - Tra cứu trong file Excel (keyword có dấu cách dùng `"...")`',
        "",
        "_Admin only:_",
        "*/add_user <user_id>* - Thêm user vào allowed list",
        "*/remove_user <user_id>* - Xóa user khỏi allowed list",
        "*/add_group <chat_id>* - Thêm group vào allowed list",
        "*/remove_group <chat_id>* - Xóa group khỏi allowed list",
        "*/list_allowed* - Xem danh sách allowed",
        "*/reset_claude_bot* - Reset Claude bot",
    ]
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def cmd_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    user = update.effective_user
    chat = update.effective_chat
    await update.message.reply_text(
        f"👤 User ID: `{user.id}`\n"
        f"💬 Chat ID: `{chat.id}`\n"
        f"📦 Chat type: `{chat.type}`",
        parse_mode="Markdown",
    )


async def cmd_myid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await cmd_id(update, context)


async def cmd_baocao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    bao_cao_dir = BASE_DIR.parent / "bao_cao"
    if not bao_cao_dir.exists():
        await update.message.reply_text("📂 Thư mục bao_cao chưa tồn tại.")
        return
    files = sorted(bao_cao_dir.iterdir())
    if not files:
        await update.message.reply_text("📂 Thư mục bao_cao trống.")
        return
    lines = ["📋 *Danh sách báo cáo:*", ""]
    for f in files:
        lines.append(f"• `{f.name}`")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def cmd_ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text("🏓 Pong!")


# ---------------------------------------------------------------------------
# Admin: manage allowed lists
# ---------------------------------------------------------------------------


async def cmd_add_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Chỉ admin mới dùng được lệnh này.")
        return
    if not context.args:
        await update.message.reply_text("Dùng: /add_user <user_id>")
        return
    try:
        uid = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ user_id phải là số nguyên.")
        return

    cfg = load_config()
    if uid not in cfg["allowed_users"]:
        cfg["allowed_users"].append(uid)
        save_config(cfg)
        await update.message.reply_text(
            f"✅ Đã thêm user `{uid}` vào allowed list.", parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(
            f"ℹ️ User `{uid}` đã có trong list rồi.", parse_mode="Markdown"
        )


async def cmd_remove_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Chỉ admin mới dùng được lệnh này.")
        return
    if not context.args:
        await update.message.reply_text("Dùng: /remove_user <user_id>")
        return
    try:
        uid = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ user_id phải là số nguyên.")
        return

    cfg = load_config()
    if uid in cfg["allowed_users"]:
        cfg["allowed_users"].remove(uid)
        save_config(cfg)
        await update.message.reply_text(
            f"✅ Đã xóa user `{uid}` khỏi allowed list.", parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(
            f"ℹ️ User `{uid}` không có trong list.", parse_mode="Markdown"
        )


async def cmd_add_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Chỉ admin mới dùng được lệnh này.")
        return
    if not context.args:
        await update.message.reply_text("Dùng: /add_group <chat_id>")
        return
    try:
        gid = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ chat_id phải là số nguyên.")
        return

    cfg = load_config()
    if gid not in cfg["allowed_groups"]:
        cfg["allowed_groups"].append(gid)
        save_config(cfg)
        await update.message.reply_text(
            f"✅ Đã thêm group `{gid}` vào allowed list.", parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(
            f"ℹ️ Group `{gid}` đã có trong list rồi.", parse_mode="Markdown"
        )


async def cmd_remove_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Chỉ admin mới dùng được lệnh này.")
        return
    if not context.args:
        await update.message.reply_text("Dùng: /remove_group <chat_id>")
        return
    try:
        gid = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ chat_id phải là số nguyên.")
        return

    cfg = load_config()
    if gid in cfg["allowed_groups"]:
        cfg["allowed_groups"].remove(gid)
        save_config(cfg)
        await update.message.reply_text(
            f"✅ Đã xóa group `{gid}` khỏi allowed list.", parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(
            f"ℹ️ Group `{gid}` không có trong list.", parse_mode="Markdown"
        )


async def cmd_list_allowed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Chỉ admin mới dùng được lệnh này.")
        return
    cfg = load_config()
    lines = [
        "📋 *Danh sách được phép:*",
        "",
        f"👤 *Allowed users:* {cfg.get('allowed_users', [])}",
        f"👥 *Allowed groups:* {cfg.get('allowed_groups', [])}",
        f"🔑 *Admin users:* {cfg.get('admin_users', [])}",
    ]
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def cmd_check_log(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    script = Path("/Users/hung/Projects/private/vss_data/tools/daily_check_log.py")
    if not script.exists():
        await update.message.reply_text(
            f"❌ Không tìm thấy script: `{script}`", parse_mode="Markdown"
        )
        return
    await update.message.reply_text("⏳ Đang chạy check log...")
    cfg = load_config()
    timeout = cfg.get("command_timeout", 30)
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    output = (result.stdout + result.stderr).strip()
    if result.returncode == 0:
        msg = "✅ Check log hoàn thành."
        # if output:
        #    msg += f"\n```\n{output[:3000]}\n```"
    else:
        msg = f"❌ Check log thất bại (exit {result.returncode})."
        if output:
            msg += f"\n```\n{output[:3000]}\n```"
    await update.message.reply_text(msg, parse_mode="Markdown")


async def cmd_reset_claude_bot(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Chỉ admin mới dùng được lệnh này.")
        return
    kill_script = BASE_DIR.parent / "tools" / "kill_claude"
    if not kill_script.exists():
        await update.message.reply_text(
            f"❌ Không tìm thấy script: `{kill_script}`", parse_mode="Markdown"
        )
        return
    await update.message.reply_text("🔄 Đang reset Claude bot...")
    result = subprocess.run(
        ["bash", str(kill_script)],
        capture_output=True,
        text=True,
        timeout=30,
    )
    output = (result.stdout + result.stderr).strip()
    if result.returncode == 0:
        msg = "✅ Reset Claude bot thành công."
        if output:
            msg += f"\n```\n{output}\n```"
    else:
        msg = f"❌ Reset thất bại (exit {result.returncode})."
        if output:
            msg += f"\n```\n{output}\n```"
    await update.message.reply_text(msg, parse_mode="Markdown")


async def cmd_tra_cuu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return

    # Parse thủ công để hỗ trợ keyword có dấu cách (dùng ngoặc kép)
    text = update.message.text or ""
    try:
        parts = shlex.split(text)
    except ValueError:
        await update.message.reply_text('❌ Cú pháp sai. Ví dụ: `/tra_cuu qd3176 "từ khóa"`', parse_mode="Markdown")
        return

    # parts[0] là /tra_cuu
    if len(parts) < 3:
        await update.message.reply_text(
            '📖 Cách dùng: `/tra_cuu <tên_file> <từ_khóa>`\n'
            'Nếu từ khóa có dấu cách, dùng ngoặc kép:\n'
            '`/tra_cuu qd3176 "NGAY VAO"`',
            parse_mode="Markdown",
        )
        return

    file_pattern = parts[1]
    search_term = parts[2]

    await update.message.reply_text(
        f'🔍 Đang tra cứu `{search_term}` trong file `{file_pattern}`...', parse_mode="Markdown"
    )

    script = BASE_DIR.parent / "tools" / "tra_cuu.py"
    if not script.exists():
        await update.message.reply_text(
            f"❌ Không tìm thấy script: `{script}`", parse_mode="Markdown"
        )
        return

    cfg = load_config()
    timeout = cfg.get("command_timeout", 60)

    try:
        result = subprocess.run(
            [sys.executable, str(script), file_pattern, search_term],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        output = (result.stdout + result.stderr).strip()

        if not output:
            output = "Không có kết quả."

        # Gửi output, cắt nếu quá dài
        MAX_LEN = 3800
        if len(output) <= MAX_LEN:
            await update.message.reply_text(f"```\n{output}\n```", parse_mode="Markdown")
        else:
            # Gửi từng phần
            chunks = [output[i:i+MAX_LEN] for i in range(0, len(output), MAX_LEN)]
            for i, chunk in enumerate(chunks):
                header = f"📄 Phần {i+1}/{len(chunks)}:\n" if len(chunks) > 1 else ""
                await update.message.reply_text(f"{header}```\n{chunk}\n```", parse_mode="Markdown")

    except subprocess.TimeoutExpired:
        await update.message.reply_text(f"⏱️ Tra cứu quá thời gian ({timeout}s). Thử thu hẹp phạm vi tìm kiếm.")
    except Exception as e:
        await update.message.reply_text(f"❌ Có lỗi xảy ra: {str(e)}")


async def cmd_nguyento(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return

    try:
        # Import hàm tính số nguyên tố
        script_path = BASE_DIR.parent / "tools" / "primes.py"
        import importlib.util
        spec = importlib.util.spec_from_file_location("primes", script_path)
        primes_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(primes_module)

        # Lấy limit từ arguments, mặc định là 100
        limit = 100
        if context.args:
            try:
                limit = int(context.args[0])
                if limit < 2:
                    await update.message.reply_text("❌ Số phải >= 2")
                    return
            except ValueError:
                await update.message.reply_text("❌ Số phải là số nguyên")
                return

        # Tính số nguyên tố
        result = primes_module.primes_up_to(limit)

        # Tạo thông điệp
        lines = [
            f"🔢 Số nguyên tố từ 2 đến {limit}:",
            "",
        ]
        if len(result) > 0:
            # Chia thành các cặp nếu quá nhiều
            chunk_size = 50
            for i in range(0, len(result), chunk_size):
                chunk = result[i:i + chunk_size]
                lines.append(", ".join(map(str, chunk)))

            lines.extend([
                "",
                f"📊 Tổng cộng: {len(result)} số nguyên tố",
            ])
        else:
            lines.append("Không có số nguyên tố nào.")

        message = "\n".join(lines)
        await update.message.reply_text(message, parse_mode="Markdown")

    except Exception as e:
        await update.message.reply_text(f"❌ Có lỗi xảy ra: {str(e)}")


# ---------------------------------------------------------------------------
# Unknown message
# ---------------------------------------------------------------------------


async def handle_unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text(
        "❓ Lệnh không nhận ra. Dùng /help để xem danh sách lệnh."
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


async def post_init(app: Application):
    commands = [
        BotCommand("start", "Kiểm tra bot"),
        BotCommand("help", "Xem danh sách lệnh"),
        BotCommand("id", "Lấy user ID và chat ID"),
        BotCommand("myid", "Lấy user ID và chat ID"),
        BotCommand("ping", "Kiểm tra kết nối"),
        BotCommand("baocao", "Liệt kê các file báo cáo"),
        BotCommand("check_log", "Chạy kiểm tra log hàng ngày"),
        BotCommand("add_user", "Thêm user vào allowed list"),
        BotCommand("remove_user", "Xóa user khỏi allowed list"),
        BotCommand("add_group", "Thêm group vào allowed list"),
        BotCommand("remove_group", "Xóa group khỏi allowed list"),
        BotCommand("list_allowed", "Xem danh sách allowed"),
        BotCommand("reset_claude_bot", "Reset Claude bot (admin only)"),
        BotCommand("tra_cuu", "Tra cứu trong file Excel"),
    ]
    await app.bot.set_my_commands(commands)
    logger.info("Đã đăng ký %d commands với Telegram.", len(commands))


def main():
    cfg = load_config()
    token = cfg.get("bot_token", "")
    if not token or token == "YOUR_BOT_TOKEN_HERE":
        logger.error("Chưa cấu hình bot_token trong config.json")
        sys.exit(1)

    app = Application.builder().token(token).post_init(post_init).build()

    # Thêm handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("myid", cmd_myid))
    app.add_handler(CommandHandler("ping", cmd_ping))
    app.add_handler(CommandHandler("baocao", cmd_baocao))
    app.add_handler(CommandHandler("check_log", cmd_check_log))
    app.add_handler(CommandHandler("add_user", cmd_add_user))
    app.add_handler(CommandHandler("remove_user", cmd_remove_user))
    app.add_handler(CommandHandler("add_group", cmd_add_group))
    app.add_handler(CommandHandler("remove_group", cmd_remove_group))
    app.add_handler(CommandHandler("list_allowed", cmd_list_allowed))
    app.add_handler(CommandHandler("reset_claude_bot", cmd_reset_claude_bot))
    app.add_handler(CommandHandler("tra_cuu", cmd_tra_cuu))
    app.add_handler(MessageHandler(filters.COMMAND, handle_unknown))

    logger.info("Bot server đang chạy...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
