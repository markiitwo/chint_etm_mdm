from __future__ import annotations

import datetime as dt
import re
import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook

from .db import ProductRecord, fetch_products


ARTICLE_HEADERS = ("Артикул", "Расширенный артикул")
CONFIDENT_STATIC_VALUES = {
    "Код производителя": "CHINT",
    "Страна": "CHN",
    "Название упаковки": "шт",
}


@dataclass(frozen=True)
class FillReportRow:
    row_number: int
    article: str
    column: str
    status: str
    value: str
    source: str
    note: str = ""


@dataclass(frozen=True)
class FillResult:
    output_path: Path
    report_path: Path
    total_rows: int
    found_articles: int
    filled_cells: int
    suggested_cells: int
    missing_articles: int


def normalize_header(value: object) -> str:
    return str(value or "").replace("\ufeff", "").strip()


def normalize_article(value: object) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def comparable_text(value: str) -> str:
    return re.sub(r"[^0-9a-zа-я]+", "", value.lower())


def short_wms_name(value: str, limit: int = 60) -> str:
    clean = " ".join(value.split())
    return clean[:limit].rstrip()


def format_decimal(value: float | int | None, digits: int = 6) -> str:
    if value is None:
        return ""
    text = f"{float(value):.{digits}f}".rstrip("0").rstrip(".")
    return text


def mm_to_m(value: float | int | None) -> str:
    if value is None:
        return ""
    return format_decimal(float(value) / 1000.0, 6)


def suggested_attribute(product: ProductRecord, attr_name: str) -> tuple[str, str]:
    attrs = product.attributes or {}
    wanted = comparable_text(attr_name)
    if not wanted:
        return "", ""

    for source_name, value in attrs.items():
        if comparable_text(source_name) == wanted:
            return value, source_name

    for source_name, value in attrs.items():
        source_key = comparable_text(source_name)
        if wanted in source_key or source_key in wanted:
            return value, source_name

    return "", ""


def product_value(product: ProductRecord, header: str) -> tuple[str, str, str]:
    if header in CONFIDENT_STATIC_VALUES:
        return CONFIDENT_STATIC_VALUES[header], "filled", "static"
    if header == "Расширенный артикул":
        return product.article, "filled", "products.article"
    if header == "81 класс":
        return product.class81_code, "filled", "ipro_goods.class81_code"
    if header == "Название":
        return product.name or product.full_name, "filled", "products/ipro_goods.name"
    if header == "Полное название":
        return product.full_name or product.name, "filled", "ipro_goods.full_name"
    if header == "Краткое имя WMS":
        return short_wms_name(product.name or product.full_name), "filled", "generated"
    if header == "Вес, кг":
        return format_decimal(product.weight_kg, 6), "filled", "product_dimensions_resolved.weight_kg"
    if header == "Длина, м":
        return mm_to_m(product.length_mm), "filled", "product_dimensions_resolved.length_mm"
    if header == "Ширина, м":
        return mm_to_m(product.width_mm), "filled", "product_dimensions_resolved.width_mm"
    if header == "Высота, м":
        return mm_to_m(product.height_mm), "filled", "product_dimensions_resolved.height_mm"
    if header == "Объем, м3":
        return format_decimal(product.volume_m3, 9), "filled", "product_dimensions_resolved.volume_m3"
    if header == "Код ТН ВЭД":
        return product.tnved_code, "filled", "product_tnved_codes.tnved_code"
    if header == "Код ОКПД2":
        return product.okpd2_code, "filled", "product_okpd2_codes.okpd2"
    if header.startswith("Конфиг:"):
        attr_name = header.removeprefix("Конфиг:").strip()
        value = (product.attributes or {}).get(attr_name, "")
        if value:
            return value, "filled", "product_attribute_values.attribute_name"
        suggestion, source_name = suggested_attribute(product, attr_name)
        if suggestion:
            return suggestion, "suggested", f"product_attribute_values.attribute_name:{source_name}"
        return "", "blank", ""
    return "", "blank", ""


def find_article_index(headers: list[str]) -> int | None:
    for candidate in ARTICLE_HEADERS:
        if candidate in headers:
            return headers.index(candidate)
    return None


def make_output_paths(template_path: Path, output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = output_dir / f"{template_path.stem}_filled_{stamp}{template_path.suffix}"
    report_path = output_dir / f"{template_path.stem}_report_{stamp}.xlsx"
    return output_path, report_path


def write_report(path: Path, rows: Iterable[FillReportRow]) -> None:
    rows = list(rows)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Итог"
    details = workbook.create_sheet("Детали")

    summary.append(
        [
            "Строка",
            "Артикул",
            "Статус",
            "Заполнено точно",
            "Заполнено по предложению",
            "Комментарий",
        ]
    )
    details.append(["Строка", "Артикул", "Колонка", "Статус", "Значение", "Источник", "Комментарий"])

    by_row: dict[tuple[int, str], dict[str, object]] = {}
    for item in rows:
        key = (item.row_number, item.article)
        bucket = by_row.setdefault(
            key,
            {
                "filled": 0,
                "suggested": 0,
                "status": "ok",
                "notes": [],
            },
        )
        if item.status == "filled":
            bucket["filled"] = int(bucket["filled"]) + 1
        elif item.status == "filled_suggested":
            bucket["suggested"] = int(bucket["suggested"]) + 1
            bucket["status"] = "needs_review"
        elif item.status == "not_found":
            bucket["status"] = "not_found"
        if item.note:
            notes = bucket["notes"]
            assert isinstance(notes, list)
            if item.note not in notes:
                notes.append(item.note)

        details.append(
            [
                item.row_number,
                item.article,
                item.column,
                item.status,
                item.value,
                item.source,
                item.note,
            ]
        )

    for (row_number, article), bucket in sorted(by_row.items()):
        notes = bucket["notes"]
        assert isinstance(notes, list)
        summary.append(
            [
                row_number,
                article,
                bucket["status"],
                bucket["filled"],
                bucket["suggested"],
                "; ".join(notes),
            ]
        )

    for sheet in (summary, details):
        sheet.freeze_panes = "A2"
        for column_cells in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 70)

    workbook.save(path)


def fill_template(db_path: Path, template_path: Path, output_dir: Path) -> FillResult:
    suffix = template_path.suffix.lower()
    if suffix == ".csv":
        return fill_csv_template(db_path, template_path, output_dir)
    if suffix in {".xlsx", ".xlsm"}:
        return fill_xlsx_template(db_path, template_path, output_dir)
    raise ValueError("Поддерживаются только CSV, XLSX и XLSM шаблоны.")


def fill_csv_template(db_path: Path, template_path: Path, output_dir: Path) -> FillResult:
    with template_path.open("r", newline="", encoding="utf-8-sig") as fh:
        sample = fh.read(4096)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        except csv.Error:
            dialect = csv.excel()
            dialect.delimiter = ";"
        reader = list(csv.reader(fh, dialect))

    if not reader:
        raise ValueError("Шаблон пустой.")

    headers = [normalize_header(h) for h in reader[0]]
    article_idx = find_article_index(headers)
    if article_idx is None:
        raise ValueError("Не найдена колонка Артикул или Расширенный артикул.")

    articles = [normalize_article(row[article_idx]) for row in reader[1:] if len(row) > article_idx]
    products = fetch_products(db_path, articles)
    report: list[FillReportRow] = []
    filled_cells = 0
    suggested_cells = 0

    for row_number, row in enumerate(reader[1:], start=2):
        while len(row) < len(headers):
            row.append("")
        article = normalize_article(row[article_idx])
        if not article:
            continue
        product = products.get(article)
        if not product:
            report.append(FillReportRow(row_number, article, "", "not_found", "", "", "Артикул не найден в базе"))
            continue
        for col_idx, header in enumerate(headers):
            if header == "Артикул":
                continue
            value, status, source = product_value(product, header)
            if status == "filled" and value:
                row[col_idx] = value
                filled_cells += 1
                report.append(FillReportRow(row_number, article, header, status, value, source))
            elif status == "suggested" and value:
                row[col_idx] = value
                filled_cells += 1
                suggested_cells += 1
                report.append(
                    FillReportRow(
                        row_number,
                        article,
                        header,
                        "filled_suggested",
                        value,
                        source,
                        "Заполнено по близкому совпадению; желательно проверить",
                    )
                )

    output_path, report_path = make_output_paths(template_path, output_dir)
    with output_path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh, delimiter=getattr(dialect, "delimiter", ";"))
        writer.writerows(reader)
    write_report(report_path, report)

    unique_articles = {a for a in articles if a}
    found = len(unique_articles.intersection(products.keys()))
    return FillResult(
        output_path,
        report_path,
        len(reader) - 1,
        found,
        filled_cells,
        suggested_cells,
        len(unique_articles) - found,
    )


def fill_xlsx_template(db_path: Path, template_path: Path, output_dir: Path) -> FillResult:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    headers = [normalize_header(cell.value) for cell in sheet[1]]
    article_idx = find_article_index(headers)
    if article_idx is None:
        raise ValueError("Не найдена колонка Артикул или Расширенный артикул.")

    articles: list[str] = []
    for row in sheet.iter_rows(min_row=2):
        if article_idx < len(row):
            articles.append(normalize_article(row[article_idx].value))

    products = fetch_products(db_path, articles)
    report: list[FillReportRow] = []
    filled_cells = 0
    suggested_cells = 0

    for row_number, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        article = normalize_article(row[article_idx].value if article_idx < len(row) else "")
        if not article:
            continue
        product = products.get(article)
        if not product:
            report.append(FillReportRow(row_number, article, "", "not_found", "", "", "Артикул не найден в базе"))
            continue
        for col_idx, header in enumerate(headers):
            if header == "Артикул" or col_idx >= len(row):
                continue
            value, status, source = product_value(product, header)
            if status == "filled" and value:
                row[col_idx].value = value
                filled_cells += 1
                report.append(FillReportRow(row_number, article, header, status, value, source))
            elif status == "suggested" and value:
                row[col_idx].value = value
                filled_cells += 1
                suggested_cells += 1
                report.append(
                    FillReportRow(
                        row_number,
                        article,
                        header,
                        "filled_suggested",
                        value,
                        source,
                        "Заполнено по близкому совпадению; желательно проверить",
                    )
                )

    output_path, report_path = make_output_paths(template_path, output_dir)
    workbook.save(output_path)
    write_report(report_path, report)

    unique_articles = {a for a in articles if a}
    found = len(unique_articles.intersection(products.keys()))
    return FillResult(
        output_path=output_path,
        report_path=report_path,
        total_rows=sheet.max_row - 1,
        found_articles=found,
        filled_cells=filled_cells,
        suggested_cells=suggested_cells,
        missing_articles=len(unique_articles) - found,
    )
