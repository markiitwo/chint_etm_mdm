import os
import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.data_checks import is_range, name_conflict, review_note
from chint_etm_mdm.etm_categories import latest_codes_file

ALL = frozenset({"name_numbers", "ranges"})


class NameConflictTests(unittest.TestCase):
    def test_a_voltage_the_name_does_not_have_is_held(self) -> None:
        self.assertIn("48", name_conflict("Номинальное напряжение вторичной обмотки, В", "48", "Трансформатор NDK-100 220/24*2"))
        self.assertEqual("", name_conflict("Номинальное напряжение вторичной обмотки, В", "24", "Трансформатор NDK-100 220/24*2"))

    def test_currents_and_curves(self) -> None:
        self.assertEqual("", name_conflict("Номинальный ток (А)", "16", "Выключатель NXB-63 1P C16 6кА"))
        self.assertIn("25", name_conflict("Номинальный ток (А)", "25", "Выключатель NXB-63 1P C16 6кА"))
        self.assertEqual("", name_conflict("Номинальный ток (А)", "2.5", "Блок питания 24В 2,5А"))

    def test_names_without_such_values_are_not_checked(self) -> None:
        self.assertEqual("", name_conflict("Номинальное напряжение, В", "400", "Выключатель NXB-63 1P C16"))
        self.assertEqual("", name_conflict("Световой поток", "100", "Лампа ND16 220В 5А"))
        self.assertEqual("", name_conflict("Дифференциальный ток (мА)", "30", "АВДТ NB1L 1P+N C16 30мА"))
        self.assertEqual("", name_conflict("Напряжение изоляции, В", "690", "Реле 24В"))

    def test_ranges(self) -> None:
        self.assertTrue(is_range("AC 110-230"))
        self.assertTrue(is_range("24…48"))
        self.assertFalse(is_range("1P+N"))
        self.assertFalse(is_range("NB1L-20"))
        self.assertIn("Диапазон", review_note(ALL, "Напряжение питания, В", "110-230", ""))
        self.assertEqual("", review_note(frozenset({"name_numbers"}), "Напряжение питания, В", "110-230", ""))
        self.assertEqual("", review_note(frozenset(), "Напряжение, В", "48", "Трансформатор 220/24"))


class CodesFileTests(unittest.TestCase):
    def test_newest_file_by_the_date_in_its_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            downloads = root / "Downloads"
            downloads.mkdir()
            old = root / "Список товаров под ЧЗ_01.09.2026.xlsx"
            new = downloads / "Список товаров под ЧЗ_19.10.2026.xlsx"
            other = downloads / "прайс_20.10.2026.xlsx"
            for path in (new, other, old):
                path.write_bytes(b"")
            os.utime(new, (0, 0))  # an old file date does not matter when the name has one
            self.assertEqual(new, latest_codes_file([downloads, root / "нет"], old))
            self.assertEqual(old, latest_codes_file([root / "нет"], old))
            self.assertIsNone(latest_codes_file([root / "нет"]))


if __name__ == "__main__":
    unittest.main()
