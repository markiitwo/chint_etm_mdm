import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.notify import (
    NotifySettings,
    find_telegram_chat,
    load_settings,
    save_settings,
    send_email,
    send_summary,
    send_telegram,
)


class FakeSMTP:
    sent = []

    def __init__(self, host, port, timeout):
        self.host, self.port = host, port

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def login(self, user, password):
        self.user = user

    def send_message(self, message):
        FakeSMTP.sent.append((self.host, self.port, self.user, message["To"], message["Subject"], message.get_content()))


class NotifyTests(unittest.TestCase):
    def test_settings_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "notify.json"
            self.assertEqual(NotifySettings(), load_settings(path))
            settings = NotifySettings(True, "me@x.ru", "smtp.x.ru", 465, "bot@x.ru", "secret", "123:abc", "42")
            save_settings(settings, path)
            self.assertEqual(settings, load_settings(path))
            self.assertEqual(0o600, path.stat().st_mode & 0o777)

    def test_telegram_message_and_chat(self) -> None:
        calls = []

        def post(url, data):
            calls.append((url, data))
            if url.endswith("getUpdates"):
                return {"ok": True, "result": [{"message": {"chat": {"id": 42, "first_name": "Марк"}}}]}
            return {"ok": True}

        self.assertEqual(("42", "Марк"), find_telegram_chat("123:abc", post))
        send_telegram(NotifySettings(telegram_token="123:abc", telegram_chat_id="42"), "x" * 4500, post)
        sends = [data for url, data in calls if url == "https://api.telegram.org/bot123:abc/sendMessage"]
        self.assertEqual([4000, 500], [len(data["text"]) for data in sends])
        with self.assertRaises(RuntimeError):
            find_telegram_chat("123:abc", lambda url, data: {"ok": True, "result": []})

    def test_email(self) -> None:
        FakeSMTP.sent.clear()
        settings = NotifySettings(email_to="me@x.ru", smtp_host="smtp.x.ru", smtp_user="bot@x.ru", smtp_password="p")
        send_email(settings, "РАЭК готов", "Заполнено 10", smtp_factory=FakeSMTP)
        self.assertEqual([("smtp.x.ru", 465, "bot@x.ru", "me@x.ru", "РАЭК готов", "Заполнено 10\n")], FakeSMTP.sent)

    def test_summary_reports_each_channel(self) -> None:
        def broken(*args):
            raise OSError("нет сети")

        settings = NotifySettings(email_to="me@x.ru", smtp_host="h", smtp_user="u", telegram_token="t", telegram_chat_id="1")
        report = send_summary(settings, "РАЭК", ["Заполнено 10"], telegram=lambda s, text: None, email=broken)
        self.assertEqual(["Сводка отправлена в Telegram.", "Почта: не отправилось (нет сети)."], report)
        self.assertIn("некуда отправить", send_summary(NotifySettings(), "РАЭК", [])[0])


if __name__ == "__main__":
    unittest.main()
