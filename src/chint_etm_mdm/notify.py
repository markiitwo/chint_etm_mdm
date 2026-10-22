"""Short run summaries by email and Telegram.

After a partner file is filled the program can send the same lines it shows in
its log: to an email address and/or to a Telegram chat. The settings (mail
server, password, bot token) live next to the program's own settings in the
user's profile, never in the work folder or the repository.

Telegram needs a bot: create it with @BotFather, paste its token, write any
message to the bot, then press «Найти чат» — the program reads the chat id from
that message.
"""
from __future__ import annotations

import json
import smtplib
import ssl
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass, fields
from email.message import EmailMessage
from pathlib import Path
from typing import Callable

from .atomic_files import atomic_write_json
from .net import open_url

SETTINGS_NAME = "notify.json"
TELEGRAM_API = "https://api.telegram.org/bot{token}/{method}"
TELEGRAM_LIMIT = 4000
TIMEOUT = 20


@dataclass
class NotifySettings:
    send_after_run: bool = False
    email_to: str = ""
    smtp_host: str = ""
    smtp_port: int = 465
    smtp_user: str = ""
    smtp_password: str = ""
    telegram_token: str = ""
    telegram_chat_id: str = ""

    @property
    def email_ready(self) -> bool:
        return bool(self.email_to and self.smtp_host and self.smtp_user)

    @property
    def telegram_ready(self) -> bool:
        return bool(self.telegram_token and self.telegram_chat_id)


def settings_path() -> Path:
    from .config import user_config_path

    return user_config_path().with_name(SETTINGS_NAME)


def load_settings(path: Path | None = None) -> NotifySettings:
    path = path or settings_path()
    if not path.exists():
        return NotifySettings()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return NotifySettings()
    known = {item.name for item in fields(NotifySettings)}
    settings = NotifySettings(**{key: value for key, value in raw.items() if key in known})
    settings.smtp_port = int(settings.smtp_port or 465)
    return settings


def save_settings(settings: NotifySettings, path: Path | None = None) -> None:
    path = path or settings_path()
    atomic_write_json(path, asdict(settings))
    try:
        path.chmod(0o600)  # the password and the bot token stay readable by the user only
    except OSError:
        pass


def post_json(url: str, data: dict) -> dict:
    request = urllib.request.Request(url, data=urllib.parse.urlencode(data).encode(), method="POST")
    response = open_url(request, TIMEOUT)
    with response:
        return json.loads(response.read().decode("utf-8"))


def send_telegram(settings: NotifySettings, text: str, post: Callable[[str, dict], dict] = post_json) -> None:
    url = TELEGRAM_API.format(token=settings.telegram_token.strip(), method="sendMessage")
    for start in range(0, max(len(text), 1), TELEGRAM_LIMIT):
        answer = post(url, {"chat_id": settings.telegram_chat_id.strip(), "text": text[start : start + TELEGRAM_LIMIT]})
        if not answer.get("ok"):
            raise RuntimeError(f"Telegram ответил: {answer.get('description', answer)}")


def find_telegram_chat(token: str, post: Callable[[str, dict], dict] = post_json) -> tuple[str, str]:
    """(chat id, who) of the last message written to the bot."""
    answer = post(TELEGRAM_API.format(token=token.strip(), method="getUpdates"), {})
    if not answer.get("ok"):
        raise RuntimeError(f"Telegram ответил: {answer.get('description', answer)}")
    for update in reversed(answer.get("result", [])):
        chat = (update.get("message") or update.get("channel_post") or {}).get("chat")
        if chat:
            who = chat.get("title") or " ".join(filter(None, [chat.get("first_name"), chat.get("last_name")])) or chat.get("username", "")
            return str(chat["id"]), who
    raise RuntimeError("Боту ещё никто не писал. Напишите боту любое сообщение и нажмите кнопку ещё раз.")


def send_email(settings: NotifySettings, subject: str, text: str, smtp_factory=None) -> None:
    message = EmailMessage()
    message["From"] = settings.smtp_user
    message["To"] = settings.email_to
    message["Subject"] = subject
    message.set_content(text)
    port = int(settings.smtp_port or 465)
    if smtp_factory is None:
        smtp_factory = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    with smtp_factory(settings.smtp_host, port, timeout=TIMEOUT) as server:
        if port != 465 and hasattr(server, "starttls"):
            server.starttls(context=ssl.create_default_context())
        server.login(settings.smtp_user, settings.smtp_password)
        server.send_message(message)


def send_summary(
    settings: NotifySettings,
    title: str,
    lines: list[str],
    telegram=send_telegram,
    email=send_email,
) -> list[str]:
    """Send to every channel that is set up. Returns one line per channel: sent or why not."""
    text = "\n".join([title, ""] + lines)
    report = []
    if settings.telegram_ready:
        try:
            telegram(settings, text)
            report.append("Сводка отправлена в Telegram.")
        except Exception as exc:  # a summary must never break the run itself
            report.append(f"Telegram: не отправилось ({exc}).")
    if settings.email_ready:
        try:
            email(settings, title, text)
            report.append(f"Сводка отправлена на почту {settings.email_to}.")
        except Exception as exc:
            report.append(f"Почта: не отправилось ({exc}).")
    if not report:
        report.append("Сводку некуда отправить: настройте почту или Telegram (вкладка «База», кнопка «Сводка на почту и в Telegram»).")
    return report
