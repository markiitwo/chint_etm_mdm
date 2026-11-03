"""ETIM import: every sheet is read, and the base and the report never disagree."""

from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook

from chint_etm_mdm import etim_importer


def make_base(path: Path, articles: list[str]) -> None:
    with closing(sqlite3.connect(path)) as conn, conn:
        conn.execute("CREATE TABLE products (article TEXT PRIMARY KEY)")
        conn.executemany("INSERT INTO products VALUES (?)", [(article,) for article in articles])


def write_etim(path: Path, sheets: list[tuple[str, list[list[object]]]]) -> Path:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets:
        sheet = workbook.create_sheet(title)
        for row in rows:
            sheet.append(row)
    workbook.save(path)
    workbook.close()
    return path


SHORT = ("Кратко", [["Артикул", "Наименование"], ["A-1", "Выключатель"]])
DETAILED = ("Подробно", [["Артикул", "Длина, мм", "Ширина, мм", "Высота, мм", "Цвет"], ["A-1", 100, 50, 80, "белый"]])


def rows(path: Path, query: str) -> list[tuple]:
    with closing(sqlite3.connect(path)) as conn:
        return conn.execute(query).fetchall()


def has_table(path: Path, name: str) -> bool:
    return bool(rows(path, f"SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = '{name}'"))


class ScanAllSheetsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def scan(self, sheets, articles=("A-1",), notes=None):
        path = write_etim(self.root / "etim.xlsx", sheets)
        return etim_importer.scan_etim_workbook(path, list(articles), notes)

    def test_short_sheet_first_does_not_hide_the_detailed_one(self) -> None:
        found = self.scan([SHORT, DETAILED])["A-1"]
        self.assertEqual({"length_mm": 100.0, "width_mm": 50.0, "height_mm": 80.0}, found.dimensions)
        self.assertEqual("Подробно", found.dimension_sheets["length_mm"])
        self.assertEqual("Выключатель", found.attributes["Наименование"])
        self.assertEqual("белый", found.attributes["Цвет"])
        self.assertEqual("Кратко", found.attribute_sheets["Наименование"])
        self.assertEqual("Подробно", found.attribute_sheets["Цвет"])

    def test_sheet_order_does_not_change_the_result(self) -> None:
        first = self.scan([SHORT, DETAILED])["A-1"]
        second = self.scan([DETAILED, SHORT])["A-1"]
        self.assertEqual(first.dimensions, second.dimensions)
        self.assertEqual(first.attributes, second.attributes)

    def test_analog_is_not_taken_for_the_product(self) -> None:
        notes = etim_importer.ScanNotes()
        found = self.scan(
            [("Лист", [["Артикул", "Артикул аналога", "Длина, мм"], ["B-9", "A-1", 300]])], notes=notes
        )
        self.assertEqual({}, found)
        self.assertEqual({"B-9": "Лист"}, notes.unknown_articles)

    def test_different_sizes_in_the_file_are_not_chosen_silently(self) -> None:
        notes = etim_importer.ScanNotes()
        other = ("Ещё", [["Артикул", "Длина, мм", "Ширина, мм"], ["A-1", 120, 50]])
        found = self.scan([DETAILED, other], notes=notes)["A-1"]
        self.assertNotIn("length_mm", found.dimensions)
        self.assertEqual(50.0, found.dimensions["width_mm"])
        lengths = [conflict for conflict in notes.conflicts if conflict.field == "Длина, мм"]
        self.assertEqual(1, len(lengths))
        self.assertEqual({"100", "120"}, {value for value, _sheet, _row in lengths[0].values})
        self.assertEqual("", lengths[0].kept)

    def test_same_value_twice_is_not_a_conflict(self) -> None:
        notes = etim_importer.ScanNotes()
        found = self.scan([DETAILED, ("Копия", DETAILED[1])], notes=notes)["A-1"]
        self.assertEqual(100.0, found.dimensions["length_mm"])
        self.assertEqual([], notes.conflicts)

    def test_real_448129_conflicting_temperature_and_warranty_are_not_imported(self) -> None:
        # Unchanged cells from Лист8, rows 458 and 1472 of the user's ETIM export.
        fixture = json.loads(
            (Path(__file__).parent / "fixtures" / "etim_448129_conflicts.json").read_text(encoding="utf-8")
        )
        notes = etim_importer.ScanNotes()
        found = self.scan([(fixture["sheet"], fixture["rows"])], articles=("448129",), notes=notes)["448129"]
        for field in fixture["rows"][0][1:3]:
            with self.subTest(field=field):
                self.assertNotIn(field, found.attributes)
                self.assertNotIn(field, found.attribute_sheets)
                conflict = next(item for item in notes.conflicts if item.field == field)
                self.assertEqual("", conflict.kept)
                column = fixture["rows"][0].index(field)
                self.assertEqual(
                    {str(fixture["rows"][row][column]) for row in (1, 3)},
                    {value for value, _sheet, _row in conflict.values},
                )
        self.assertEqual(fixture["rows"][1][3], found.attributes["Наименование"])


class ImportSafetyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.db = self.root / "base" / "chint_mdm.sqlite"
        self.db.parent.mkdir()
        make_base(self.db, ["A-1", "A-2"])
        self.reports = self.root / "reports"

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def run_import(self, sheets, make_backup: bool = False):
        path = write_etim(self.root / "etim.xlsx", sheets)
        return etim_importer.import_etim_workbook(path, self.db, make_backup=make_backup, report_dir=self.reports)

    def test_import_reads_both_sheets_and_reports_extras(self) -> None:
        other = ("Ещё", [["Артикул", "Высота, мм"], ["A-1", 90], ["Z-404", 10]])
        result = self.run_import([SHORT, DETAILED, other])
        self.assertEqual(1, result.found_articles)
        self.assertEqual(1, result.unknown_articles)
        self.assertEqual(1, result.file_conflicts)
        self.assertEqual(
            [(100.0, 50.0, None)],
            rows(self.db, "SELECT length_mm, width_mm, height_mm FROM product_dimensions_resolved"),
        )
        self.assertEqual(
            [("length_mm", "Подробно"), ("width_mm", "Подробно")],
            rows(self.db, "SELECT dimension_kind, source_sheet FROM product_dimensions_raw ORDER BY dimension_kind"),
        )
        report = load_workbook(result.report_path)
        self.assertIn(["Z-404", "Ещё"], [list(row) for row in report["Нет в базе"].iter_rows(min_row=2, values_only=True)])
        heights = [row for row in report["Разное в файле"].iter_rows(min_row=2, values_only=True) if row[1] == "Высота, мм"]
        self.assertEqual(2, len(heights))
        self.assertEqual({"не записано"}, {row[5] for row in heights})
        self.assertEqual(result.import_id, etim_importer.report_import_id(report))
        report.close()
        self.assertEqual([], list(self.reports.glob("*.saving.xlsx")))
        lines = etim_importer.result_lines(result)
        self.assertTrue(any("Разные значения в самом файле: 1" in line for line in lines))
        self.assertTrue(any("нет в базе: 1" in line for line in lines))

    def test_nothing_found_is_said_plainly(self) -> None:
        result = self.run_import([("Лист", [["Артикул", "Длина, мм"], ["Z-1", 10]])])
        self.assertEqual(0, result.found_articles)
        self.assertTrue(etim_importer.result_lines(result)[0].startswith("Внимание: в файле не найдено"))

    def test_report_that_cannot_be_saved_leaves_the_base_unchanged(self) -> None:
        with mock.patch.object(etim_importer, "write_etim_report", side_effect=PermissionError("файл открыт")):
            with self.assertRaisesRegex(RuntimeError, "База не изменена"):
                self.run_import([DETAILED])
        self.assertEqual([], rows(self.db, "SELECT * FROM dimension_imports"))
        self.assertEqual([], rows(self.db, "SELECT * FROM product_dimensions_resolved"))
        self.assertEqual([], rows(self.db, "SELECT * FROM product_attribute_values"))
        self.assertEqual([], list(self.reports.glob("*.xlsx")))

    def test_report_name_failure_after_commit_still_gives_the_report(self) -> None:
        with mock.patch.object(etim_importer.os, "replace", side_effect=PermissionError("занято")):
            result = self.run_import([DETAILED])
        self.assertIn("временным именем", result.report_warning)
        self.assertTrue(result.report_path.name.endswith(".saving.xlsx"))
        self.assertTrue(result.report_path.exists())
        self.assertEqual(1, len(rows(self.db, "SELECT * FROM dimension_imports")))
        report = load_workbook(result.report_path)
        self.assertEqual(result.import_id, etim_importer.report_import_id(report))
        report.close()
        self.assertTrue(any("Внимание: Отчёт сохранён" in line for line in etim_importer.result_lines(result)))

    def test_same_file_twice_is_refused_unless_asked(self) -> None:
        path = write_etim(self.root / "etim.xlsx", [DETAILED])
        first = etim_importer.import_etim_workbook(path, self.db, make_backup=False, report_dir=self.reports)
        before = rows(self.db, "SELECT COUNT(*) FROM product_attribute_values")
        self.assertEqual((first.import_id, mock.ANY), etim_importer.etim_already_imported(path, self.db))
        with self.assertRaisesRegex(ValueError, "уже импортирован"):
            etim_importer.import_etim_workbook(path, self.db, make_backup=False, report_dir=self.reports)
        self.assertEqual(before, rows(self.db, "SELECT COUNT(*) FROM product_attribute_values"))
        self.assertEqual(1, len(rows(self.db, "SELECT * FROM dimension_imports")))
        again = etim_importer.import_etim_workbook(path, self.db, make_backup=False, report_dir=self.reports, allow_repeat=True)
        self.assertNotEqual(first.import_id, again.import_id)

    def attrs(self) -> list[tuple]:
        return rows(
            self.db,
            "SELECT article, attribute_name, value_raw, source_ref FROM product_attribute_values "
            "WHERE source_kind = 'etim_manual' ORDER BY article, attribute_name",
        )

    def import_named(self, name: str, table: list[list[object]], **kwargs):
        path = write_etim(self.root / name, [("Лист", table)])
        return etim_importer.import_etim_workbook(path, self.db, make_backup=False, report_dir=self.reports, **kwargs)

    def test_new_file_updates_values_in_place_and_keeps_missing_ones(self) -> None:
        self.import_named("a.xlsx", [["Артикул", "Цвет", "Материал"], ["A-1", "белый", "пластик"]])
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute(
                "INSERT INTO product_attribute_values (article, attribute_name, value_raw, source_kind) VALUES ('A-1', 'Цвет', 'X', 'catalog_manual')"
            )
        result = self.import_named("b.xlsx", [["Артикул", "Цвет"], ["A-1", "серый"]])
        self.assertEqual([("A-1", "Материал", "пластик", "a.xlsx"), ("A-1", "Цвет", "серый", "b.xlsx")], self.attrs())
        self.assertEqual((0, 1, 1), (result.attributes_written, result.attributes_changed, result.not_in_file))
        self.assertEqual(
            [("A-1", "Цвет", "X")],
            rows(self.db, "SELECT article, attribute_name, value_raw FROM product_attribute_values WHERE source_kind = 'catalog_manual'"),
        )
        report = load_workbook(result.report_path, read_only=True)
        missing = [list(row) for row in report["Нет в файле"].iter_rows(values_only=True)]
        report.close()
        self.assertEqual([["Артикул", "Характеристика", "Значение в базе", "Из файла"], ["A-1", "Материал", "пластик", "a.xlsx"]], missing)
        self.assertTrue(any("Есть в базе, нет в файле: 1" in line for line in etim_importer.result_lines(result)))

        before = self.attrs()
        again = self.import_named("b.xlsx", [["Артикул", "Цвет"], ["A-1", "серый"]], allow_repeat=True)
        self.assertEqual(before, self.attrs())
        self.assertEqual((0, 0), (again.attributes_written, again.attributes_changed))

    def test_failed_report_keeps_the_old_value(self) -> None:
        self.import_named("a.xlsx", [["Артикул", "Цвет"], ["A-1", "белый"]])
        with mock.patch.object(etim_importer, "write_etim_report", side_effect=PermissionError("файл открыт")):
            with self.assertRaises(RuntimeError):
                self.import_named("b.xlsx", [["Артикул", "Цвет"], ["A-1", "серый"]])
        self.assertEqual([("A-1", "Цвет", "белый", "a.xlsx")], self.attrs())

    def test_duplicates_of_older_versions_are_removed_keeping_the_newest(self) -> None:
        self.import_named("a.xlsx", [["Артикул", "Цвет"], ["A-1", "белый"]])
        with closing(sqlite3.connect(self.db)) as conn, conn:
            for value in ("красный", "чёрный"):  # what older versions left: the same field written again
                conn.execute(
                    "INSERT INTO product_attribute_values (article, attribute_name, value_raw, value_normalized, source_kind, source_ref) "
                    "VALUES ('A-1', 'Цвет', ?, ?, 'etim_manual', 'old.xlsx')",
                    (value, value),
                )
        result = self.import_named("b.xlsx", [["Артикул", "Материал"], ["A-2", "сталь"]])
        self.assertEqual(2, result.duplicates_removed)
        self.assertEqual([("A-1", "Цвет", "чёрный", "old.xlsx"), ("A-2", "Материал", "сталь", "b.xlsx")], self.attrs())
        self.assertEqual("ok", rows(self.db, "PRAGMA integrity_check")[0][0])
        plan = rows(self.db, "EXPLAIN QUERY PLAN SELECT * FROM product_attribute_values WHERE article = 'A-1'")
        self.assertIn("idx_product_attribute_values_article", str(plan))

    def test_broken_file_stops_before_backup_and_tables(self) -> None:
        broken = self.root / "etim.xlsx"
        broken.write_bytes(b"not a workbook")
        with self.assertRaises(Exception):
            etim_importer.import_etim_workbook(broken, self.db, make_backup=True, report_dir=self.reports)
        self.assertEqual(["chint_mdm.sqlite"], sorted(path.name for path in self.db.parent.iterdir()))
        self.assertFalse(has_table(self.db, "dimension_imports"))


if __name__ == "__main__":
    unittest.main()


class SizeColumnTests(unittest.TestCase):
    """A column that only mentions a length is not the product's size."""

    def test_size_headers(self) -> None:
        own = [
            "Длина, мм", "Ширина, см", "Высота, м", "Глубина, мм", "Высота корпуса, мм", "Габаритная длина, см",
            "Длина (L), мм", "name:EF000040:Высота корпуса, мм", "Width, mm", "Ширина изделия, мм",
        ]
        other = [
            "Макс. длина кабеля, м", "Ширина шины, мм", "Длина зачистки провода, мм", "Длина упаковки, мм",
            "Высота над уровнем моря, м", "Ширина в модулях", "Механическая износостойкость, циклов",
        ]
        for header in own:
            self.assertTrue(etim_importer.is_product_size_header(header), header)
        for header in other:
            self.assertFalse(etim_importer.is_product_size_header(header), header)

    def test_cable_length_is_not_taken_for_the_product_length(self) -> None:
        # The case from a real ETIM file: 1000 m of cable became a 1 000 000 mm product length.
        with tempfile.TemporaryDirectory() as tmp:
            path = write_etim(
                Path(tmp) / "etim.xlsx",
                [(
                    "Лист19",
                    [
                        ["Артикул", "Макс. длина кабеля, м", "Механическая износостойкость, циклов", "Высота, мм"],
                        ["311015", 1000, 1000000, 80],
                    ],
                )],
            )
            found = etim_importer.scan_etim_workbook(path, ["311015"])["311015"]
        self.assertEqual({"height_mm": 80.0}, found.dimensions)
        self.assertEqual("1000", found.attributes["Макс. длина кабеля, м"])

    def test_impossible_size_is_not_used(self) -> None:
        self.assertIsNone(etim_importer.parse_dimension_number(1000, "Длина, м"))
        self.assertEqual(2200.0, etim_importer.parse_dimension_number(2.2, "Высота, м"))
        self.assertEqual({}, etim_importer.parse_overall_dimensions("100x200x30000", "ДхШхВ, мм"))
