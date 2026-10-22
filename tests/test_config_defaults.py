import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.config import AppConfig, first_run_defaults


class FirstRunDefaultsTests(unittest.TestCase):
    def test_finds_database_in_downloads(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            home = Path(tmp)
            (home / "Downloads").mkdir()
            (home / "Downloads" / "chint_mdm.sqlite").write_bytes(b"")
            config = first_run_defaults(AppConfig(), home)
            self.assertEqual(str(home / "CHINT ETM MDM Work"), config.work_dir)
            self.assertEqual(str(home / "Downloads" / "chint_mdm.sqlite"), config.db_path)

    def test_keeps_chosen_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            config = first_run_defaults(AppConfig("w", "d.sqlite", "t", "c", "01.10.2026"), Path(tmp))
            self.assertEqual(AppConfig("w", "d.sqlite", "t", "c", "01.10.2026"), config)

    def test_no_database_found(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual("", first_run_defaults(AppConfig(), Path(tmp)).db_path)




class InboxFolderTests(unittest.TestCase):
    def test_inbox_first_then_downloads(self) -> None:
        from chint_etm_mdm.config import incoming_folders

        home = Path("/home/u")
        self.assertEqual([home / "Downloads", home / "Загрузки"], incoming_folders("", home))
        self.assertEqual([Path("/mnt/rdp"), home / "Downloads", home / "Загрузки"], incoming_folders("/mnt/rdp", home))

    def test_inbox_kept_by_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            config = first_run_defaults(AppConfig(inbox_dir="/mnt/rdp"), Path(tmp))
        self.assertEqual("/mnt/rdp", config.inbox_dir)


if __name__ == "__main__":
    unittest.main()


class BrokenSettingsTests(unittest.TestCase):
    """CFG-01: a damaged or odd settings file never stops the start and is kept for recovery."""

    def setUp(self) -> None:
        from unittest import mock

        from chint_etm_mdm import config

        self.config = config
        self.temp_dir = tempfile.TemporaryDirectory()
        self.home = Path(self.temp_dir.name)
        self.patch = mock.patch("pathlib.Path.home", return_value=self.home)
        self.patch.start()
        config.LOAD_PROBLEMS.clear()

    def tearDown(self) -> None:
        self.patch.stop()
        self.config.LOAD_PROBLEMS.clear()
        self.temp_dir.cleanup()

    def write(self, text: str) -> Path:
        path = self.config.user_config_path()
        path.write_text(text, encoding="utf-8")
        return path

    def test_wrong_kinds_of_json_give_defaults_and_a_kept_copy(self) -> None:
        for text in ("[]", "null", "42", '"text"', "{broken"):
            with self.subTest(text=text):
                self.config.LOAD_PROBLEMS.clear()
                path = self.write(text)
                self.assertEqual(self.config.AppConfig(), self.config.load_config())
                self.assertEqual(1, len(self.config.LOAD_PROBLEMS))
                copies = list(path.parent.glob("config.json.broken-*"))
                self.assertTrue(copies)
                self.assertEqual(text, copies[0].read_text(encoding="utf-8"))
                for copy in copies:
                    copy.unlink()

    def test_good_old_settings_still_read(self) -> None:
        self.write('{"work_dir": "W", "db_path": "D.sqlite", "inbox_dir": 5, "codes_file": null}')
        loaded = self.config.load_config()
        self.assertEqual(("W", "D.sqlite", "", ""), (loaded.work_dir, loaded.db_path, loaded.inbox_dir, loaded.codes_file))
        self.assertEqual([], self.config.LOAD_PROBLEMS)

    def test_settings_with_byte_order_mark_are_read(self) -> None:
        self.config.user_config_path().write_text('{"work_dir": "W"}', encoding="utf-8-sig")
        self.assertEqual("W", self.config.load_config().work_dir)

    def test_notify_settings_survive_odd_values(self) -> None:
        from chint_etm_mdm.notify import load_settings

        path = self.home / "notify.json"
        path.write_text(
            '{"send_after_run": "yes", "email_to": ["a"], "smtp_port": "abc", "telegram_chat_id": 12345}',
            encoding="utf-8",
        )
        settings = load_settings(path)
        self.assertFalse(settings.send_after_run)
        self.assertEqual("", settings.email_to)
        self.assertEqual(465, settings.smtp_port)
        self.assertEqual("12345", settings.telegram_chat_id)
        path.write_text("[1, 2]", encoding="utf-8")
        self.assertEqual(465, load_settings(path).smtp_port)
