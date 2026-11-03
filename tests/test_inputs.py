import os
import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.inputs import bring_in, input_folders, kind_folder
from chint_etm_mdm.pm_requests import latest_assignment_file


class BringInTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.work = self.root / "work"
        self.work.mkdir()
        self.downloads = self.root / "Downloads"
        self.downloads.mkdir()

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def source(self, name: str, text: str, mtime: int = 1_700_000_000) -> Path:
        path = self.downloads / name
        path.write_text(text, encoding="utf-8")
        os.utime(path, (mtime, mtime))
        return path

    def test_copied_once_into_its_folder_and_original_stays(self) -> None:
        src = self.source("Закрепление ПМ.xlsx", "v1")
        copy = bring_in(src, self.work, "закрепление ПМ")
        self.assertEqual(kind_folder(self.work, "закрепление ПМ") / "Закрепление ПМ.xlsx", copy)
        self.assertEqual("v1", copy.read_text(encoding="utf-8"))
        self.assertTrue(src.exists())
        self.assertEqual(src.stat().st_mtime, copy.stat().st_mtime)  # "the newest file" keeps its meaning
        self.assertEqual(copy, bring_in(src, self.work, "закрепление ПМ"))
        self.assertEqual(1, len(list(copy.parent.iterdir())))  # no second copy, no temp file left

    def test_same_name_other_content_is_kept_beside(self) -> None:
        first = bring_in(self.source("ЧЗ.xlsx", "old"), self.work, "ЧЗ")
        second = bring_in(self.source("ЧЗ.xlsx", "new"), self.work, "ЧЗ")
        self.assertEqual("ЧЗ_1.xlsx", second.name)
        self.assertEqual(("old", "new"), (first.read_text(encoding="utf-8"), second.read_text(encoding="utf-8")))
        self.assertEqual(second, bring_in(self.downloads / "ЧЗ.xlsx", self.work, "ЧЗ"))

    def test_file_already_in_work_folder_and_empty_values_are_left_alone(self) -> None:
        inside = self.work / "input" / "a.xlsx"
        inside.parent.mkdir()
        inside.write_text("x")
        self.assertEqual(inside, bring_in(inside, self.work, "ЧЗ"))
        self.assertIsNone(bring_in("", self.work, "ЧЗ"))
        missing = self.downloads / "нет.xlsx"
        self.assertEqual(missing, bring_in(missing, self.work, "ЧЗ"))  # the reader reports a missing file as before
        self.assertEqual(self.downloads / "ЧЗ.xlsx", bring_in(self.source("ЧЗ.xlsx", "x"), "", "ЧЗ"))  # no work folder yet

    def test_newest_file_search_sees_the_copies(self) -> None:
        old = bring_in(self.source("Закрепление ПМ.xlsx", "old", 1_700_000_000), self.work, "закрепление ПМ")
        self.assertIn(old.parent, input_folders(self.work))
        self.assertEqual(old, latest_assignment_file(input_folders(self.work)))


if __name__ == "__main__":
    unittest.main()
