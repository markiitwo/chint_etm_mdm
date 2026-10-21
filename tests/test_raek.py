import datetime as dt
import json
import sqlite3
import tempfile
import unittest
import zipfile
from contextlib import closing
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

from chint_etm_mdm.raek import fill_raek_zip, is_raek_zip, latest_raek_zip, remember_name_answers, result_lines
from test_import_data_integrity import make_database

BASE = ["RAEC_ID", "Код поставщика", "Название от поставщика", "Бренд", "Статус поставщика", "ABC Приоритет"]
RELAY = ("3244400", "631929", "Промежуточное реле NJDC-17/4ZS(MD)L 4 конт. 6А 24В AC (R)")
DRIVE = ("3249224", "616173", "Преобразователь частоты NVF7-3.0T-S6, 690В")
UNKNOWN = ("3249999", "999999", "Чего-то нет в базе")


def plain_file(column, products):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(BASE + (column if isinstance(column, list) else [column]))
    for raec_id, article, name in products:
        sheet.append([raec_id, article, name, "ЧИНТ (CHINT)", "Z", "A"])
    return workbook


def features_file():
    workbook = Workbook()
    relay = workbook.active
    relay.title = "EC001437"
    relay.append([None] * 8 + ["EV000460 - AC (перемен.)\nEV000461 - DC (постоян.)", None, None, "EV006405 - IP20\nEV006415 - IP40", None, None])
    relay.append(
        [None] * 6
        + [
            "<обязательное>\nR - Номин. напряжение питания цепи управления Us AC 50 Гц, В",
            "<обязательное>\nR - Номин. напряжение питания цепи управления Us постоян. тока DC, В",
            "<обязательное>\nA - Тип напряжения управления",
            "<обязательное>\nN - Количество переключающих (перекидных) контактов",
            "<обязательное>\nL - В комплекте с розеткой (цоколем)",
            "<обязательное>\nA - Степень защиты (IP)",
            "<обязательное>\nN - Ширина, мм",
            "<обязательное>\nN - Номин. ток, А",
        ]
    )
    relay.append(BASE + ["EF003978", "EF003980", "EF008242", "EF008257", "EF003359", "EF005474", "EF000008", "EF000001"])
    relay.append(list(RELAY) + ["ЧИНТ (CHINT)", "Z", "A"])
    drive = workbook.create_sheet("EC001857")
    drive.append([None] * 7 + ["EV000599 - 1\nEV000069 - 3\nEV000154 - Прочее", None, None, "EV008013 - 50 Гц\nEV008017 - 60 Гц\nEV008015 - 50/60 Гц"])
    drive.append(
        [None] * 6
        + [
            "<обязательное>\nR - Сетевое напряжение, В",
            "<обязательное>\nA - Количество вход. фаз",
            "<обязательное>\nN - Макс. частота на выходе, Гц",
            "<обязательное>\nN - Высота, мм",
            "<обязательное>\nA - Частота сети",
        ]
    )
    drive.append(BASE + ["EF006780", "EF007117", "EF007350", "EF000040", "EF006783"])
    drive.append(list(DRIVE) + ["ЧИНТ (CHINT)", "Z", "A"])
    return workbook


def make_zip(path: Path) -> None:
    files = {
        "ChINT (CHINT)_COUNTRY_ID.xlsx": plain_file("Страна происхождения", [RELAY, DRIVE]),
        "ChINT (CHINT)_THVED.xlsx": plain_file("Код ТН ВЭД", [DRIVE, UNKNOWN]),
        "ChINT (CHINT)_EAN.xlsx": plain_file("Штрих-код EAN", [DRIVE, RELAY]),
        "ChINT (CHINT)_OKPD2_ID.xlsx": plain_file(["Код ОКПД2", "Название ОКПД2"], [DRIVE]),
        "ChINT (CHINT)_SERIES_ID.xlsx": plain_file("Серия", [DRIVE]),
        "ChINT (CHINT)_IMAGE.xlsx": plain_file("Изображение", [DRIVE]),
        "ChINT (CHINT)_certificate.xlsx": plain_file(
            ["Тип сертификата", "Номер сертификата", "Сертификат от", "Сертификат до", "Файл", "Ссылка на реестр"], [DRIVE, RELAY]
        ),
        "ChINT (CHINT)_required_features.xlsx": features_file(),
    }
    with zipfile.ZipFile(path, "w") as archive:
        for name, workbook in files.items():
            buffer = BytesIO()
            workbook.save(buffer)
            archive.writestr(name, buffer.getvalue())
        archive.writestr("readme.txt", "оставить как есть")


def make_codes(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Price-list"
    sheet.append(["PM"])
    sheet.append(["Part number", "ОКПД 2 \n(9 зн.)", "Серия", "ШК / GTIN", "Сертификат 1", "Ссылка на реестр 1", "Tariff code"])
    sheet.append([616173, "26.51.70.190", "NVF7", "6941498900011", "ЕАЭС N RU Д-CN.РА06.В.41573/24", "https://pub.fsa.gov.ru/x", 9032890000])
    workbook.save(path)


def cells(workbook_bytes: bytes, sheet: str | None = None) -> list[tuple]:
    workbook = load_workbook(BytesIO(workbook_bytes))
    target = workbook[sheet] if sheet else workbook.active
    return list(target.iter_rows(values_only=True))


class RaekTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.db = self.root / "mdm.sqlite"
        make_database(self.db)
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.executemany(
                "INSERT INTO products(article, name, series_name) VALUES (?, ?, ?)",
                [(RELAY[1], RELAY[2], "NJDC-17"), (DRIVE[1], DRIVE[2], "NVF7")],
            )
            conn.executemany(
                "INSERT INTO product_attribute_values(article, attribute_name, value_raw, confidence) VALUES (?, ?, ?, 1)",
                [
                    (RELAY[1], "Номинальное напряжение питания цепи управления Us перемен. тока АС при 50 Гц, В", "24"),
                    (RELAY[1], "Род тока включения", "Переменный ток (AC)"),
                    (RELAY[1], "Количество переключающих контактов", "4"),
                    (RELAY[1], "В комплекте с розеткой", "Нет"),
                    (RELAY[1], "Степень защиты (IP)", "IP40"),
                    (RELAY[1], "Габаритные и установочные размеры (ШхВхГ), мм", "27,8×21,8×44"),
                    (RELAY[1], "Номинальный ток, А", "6"),
                    (RELAY[1], "Номер сертификата", "ЕАЭС N RU Д-CN.PA03.B.02078/25"),
                    (RELAY[1], "Дата начала", "2025-03-21 00:00:00"),
                    (RELAY[1], "Дата окончания", "2030-03-20 00:00:00"),
                    (DRIVE[1], "Напряжение сети, В", "525-690"),
                    (DRIVE[1], "Количество входных фаз", "3"),
                    (DRIVE[1], "Диапазон входной частоты, Гц", "47-63"),
                    (DRIVE[1], "Диапазон выходной частоты, Гц", "0-500"),
                    (DRIVE[1], "Габаритные и установочные размеры (ШхВхГ), мм", "270х466х268.1"),
                    (DRIVE[1], "Номер сертификата", "ЕАЭС N RU Д-CN.РА06.В.41573/24"),
                    (DRIVE[1], "Дата начала", "2024-07-25 00:00:00"),
                    (DRIVE[1], "Дата окончания", "2029-07-25 00:00:00"),
                ],
            )
            conn.execute(
                "INSERT INTO material_data_items(article, gtin_inner) VALUES (?, '16941498941592')", (RELAY[1],)
            )
        self.zip = self.root / "CHINT.zip"
        make_zip(self.zip)
        self.codes = self.root / "Список товаров под ЧЗ_19.10.2026.xlsx"
        make_codes(self.codes)
        self.rules = self.root / "rules" / "attribute_mappings.json"
        self.rules.parent.mkdir()

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def fill(self):
        return fill_raek_zip(self.db, self.zip, self.root / "out", self.codes, self.rules)

    def member(self, result, name: str) -> bytes:
        with zipfile.ZipFile(result.output_path) as archive:
            return archive.read(name)

    def test_plain_columns(self) -> None:
        result = self.fill()
        self.assertEqual(f"CHINT_заполнено_{dt.date.today().isoformat()}.zip", result.output_path.name)
        with zipfile.ZipFile(result.output_path) as archive:
            self.assertEqual(
                [
                    "ChINT (CHINT)_COUNTRY_ID.xlsx", "ChINT (CHINT)_THVED.xlsx", "ChINT (CHINT)_EAN.xlsx",
                    "ChINT (CHINT)_OKPD2_ID.xlsx", "ChINT (CHINT)_SERIES_ID.xlsx", "ChINT (CHINT)_IMAGE.xlsx",
                    "ChINT (CHINT)_certificate.xlsx", "ChINT (CHINT)_required_features.xlsx", "readme.txt",
                ],
                archive.namelist(),
            )
            self.assertEqual("оставить как есть", archive.read("readme.txt").decode())
        self.assertEqual(["Китай", "Китай"], [row[6] for row in cells(self.member(result, "ChINT (CHINT)_COUNTRY_ID.xlsx"))[1:]])
        self.assertEqual(["9032890000", None], [row[6] for row in cells(self.member(result, "ChINT (CHINT)_THVED.xlsx"))[1:]])
        ean = cells(self.member(result, "ChINT (CHINT)_EAN.xlsx"))
        self.assertEqual("6941498900011", ean[1][6])
        self.assertEqual("6941498941595", ean[2][6])  # from the inner pack GTIN-14, flagged for a check
        self.assertEqual(("26.51.70.190", None), cells(self.member(result, "ChINT (CHINT)_OKPD2_ID.xlsx"))[1][6:8])
        self.assertEqual("NVF7", cells(self.member(result, "ChINT (CHINT)_SERIES_ID.xlsx"))[1][6])
        cert = cells(self.member(result, "ChINT (CHINT)_certificate.xlsx"))
        self.assertEqual(
            ("Декларация ЕАС", "ЕАЭС N RU Д-CN.РА06.В.41573/24", dt.datetime(2024, 7, 25), dt.datetime(2029, 7, 25), None, "https://pub.fsa.gov.ru/x"),
            cert[1][6:12],
        )
        self.assertEqual(("Декларация ЕАС", "ЕАЭС N RU Д-CN.PA03.B.02078/25", dt.datetime(2025, 3, 21)), cert[2][6:9])
        self.assertIsNone(cert[2][11])  # no registry link without the ЧЗ file

        report = load_workbook(result.report_path)
        empty = [row for row in report["Нечем заполнить"].iter_rows(min_row=2, values_only=True)]
        self.assertIn(("ChINT (CHINT)_THVED.xlsx", "999999", "Артикула нет в базе"), [(row[0], row[3], row[7]) for row in empty])
        self.assertTrue(any(row[5] == "Изображение" and "ensmas" in row[7] for row in empty))
        checks = [row for row in report["Проверить"].iter_rows(min_row=2, values_only=True)]
        self.assertTrue(any(row[5] == "Штрих-код EAN" and row[3] == "631929" for row in checks))

    def test_features(self) -> None:
        result = self.fill()
        relay = cells(self.member(result, "ChINT (CHINT)_required_features.xlsx"), "EC001437")
        self.assertEqual(("24...24", "-...-", "AC (перемен.)", "4", "false", "IP40", "27,8", "6"), relay[3][6:14])
        drive = cells(self.member(result, "ChINT (CHINT)_required_features.xlsx"), "EC001857")
        self.assertEqual(("525...690", "3", "500", "466", "50/60 Гц"), drive[3][6:11])
        report = load_workbook(result.report_path)
        checks = [row for row in report["Проверить"].iter_rows(min_row=2, values_only=True)]
        self.assertTrue(any(row[5].startswith("EF007350") and "верхний предел" in row[7] for row in checks))
        lines = "\n".join(result_lines(result))
        self.assertIn("Архив для РАЭК: CHINT_заполнено_", lines)
        self.assertIn("Список товаров под ЧЗ_19.10.2026.xlsx", lines)

    def test_filled_cells_stay(self) -> None:
        with zipfile.ZipFile(self.zip) as archive:
            data = archive.read("ChINT (CHINT)_required_features.xlsx")
        workbook = load_workbook(BytesIO(data))
        workbook["EC001437"]["G4"] = "220...220"
        buffer = BytesIO()
        workbook.save(buffer)
        with zipfile.ZipFile(self.zip, "w") as archive:
            archive.writestr("ChINT (CHINT)_required_features.xlsx", buffer.getvalue())
        result = self.fill()
        relay = cells(self.member(result, "ChINT (CHINT)_required_features.xlsx"), "EC001437")
        self.assertEqual("220...220", relay[3][6])
        self.assertEqual(1, result.kept)

    def test_similar_names_are_checked_and_answers_remembered(self) -> None:
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute(
                "INSERT INTO product_attribute_values(article, attribute_name, value_raw, confidence) VALUES (?, ?, ?, 1)",
                (DRIVE[1], "Высота корпуса, мм", "470"),
            )
            conn.execute("DELETE FROM product_attribute_values WHERE attribute_name LIKE 'Габарит%' AND article = ?", (DRIVE[1],))
        result = self.fill()
        drive = cells(self.member(result, "ChINT (CHINT)_required_features.xlsx"), "EC001857")
        self.assertEqual("470", drive[3][9])
        self.assertEqual([["EC001857", "EF000040", "Высота, мм", "Высота корпуса, мм", 1, "470", ""]], result.names)

        report = load_workbook(result.report_path)
        report["Названия"]["G2"] = "нет"
        report.save(result.report_path)
        self.assertEqual((0, 1), remember_name_answers(result.report_path, self.rules))
        learned = json.loads((self.rules.parent / "raek_learned.json").read_text(encoding="utf-8"))
        self.assertEqual({"EF000040": ["Высота корпуса, мм"]}, learned["name_rejects"])
        again = self.fill()
        self.assertIsNone(cells(self.member(again, "ChINT (CHINT)_required_features.xlsx"), "EC001857")[3][9])

    def test_finds_the_newest_raek_archive(self) -> None:
        other = self.root / "photos.zip"
        with zipfile.ZipFile(other, "w") as archive:
            archive.writestr("a.jpg", b"")
        self.assertTrue(is_raek_zip(self.zip))
        self.assertFalse(is_raek_zip(other))
        self.assertEqual(self.zip, latest_raek_zip([self.root]))

    def test_app_worker(self) -> None:
        from chint_etm_mdm.app import RaekWorker

        results = []
        worker = RaekWorker(self.db, self.zip, self.root / "out", self.codes, self.rules)
        worker.done.connect(results.append)
        worker.failed.connect(self.fail)
        worker.run()
        self.assertTrue(results[0].output_path.exists())


if __name__ == "__main__":
    unittest.main()
