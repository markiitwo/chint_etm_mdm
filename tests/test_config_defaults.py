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


if __name__ == "__main__":
    unittest.main()
