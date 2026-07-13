# Bot Server

Bot Telegram để thực thi lệnh từ người dùng được phép.

## Cài đặt

```bash
cd bot
pip install -r requirements.txt
```

## Cấu hình

Sửa `config.json`:

```json
{
  "bot_token": "TOKEN_CỦA_BỌN",
  "allowed_users": [123456789],
  "allowed_groups": [-1001234567890],
  "admin_users": [123456789],
  "command_timeout": 30,
  "log_level": "INFO"
}
```

- `bot_token`: Token từ @BotFather
- `allowed_users`: Danh sách user ID được phép dùng bot (chat riêng)
- `allowed_groups`: Danh sách group chat ID được phép
- `admin_users`: Danh sách admin (có thể quản lý allowed list)

> Lấy user ID / chat ID bằng lệnh `/id` sau khi cho phép user đó trước.

## Chạy

```bash
python bot_server.py
```

## Lệnh có sẵn

| Lệnh | Mô tả | Quyền |
|------|-------|-------|
| `/start` | Kiểm tra bot | Allowed |
| `/help` | Xem danh sách lệnh | Allowed |
| `/id` | Lấy user ID và chat ID | Allowed |
| `/ping` | Kiểm tra kết nối | Allowed |
| `/add_user <id>` | Thêm user vào allowed list | Admin |
| `/remove_user <id>` | Xóa user khỏi allowed list | Admin |
| `/add_group <id>` | Thêm group vào allowed list | Admin |
| `/remove_group <id>` | Xóa group khỏi allowed list | Admin |
| `/list_allowed` | Xem danh sách được phép | Admin |

## Thêm lệnh mới

Tạo handler function trong `bot_server.py` và đăng ký trong `main()`:

```python
async def cmd_my_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text("Kết quả lệnh...")

# Trong main():
app.add_handler(CommandHandler("my_command", cmd_my_command))
```
