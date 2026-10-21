"""ETM report of missing characteristics ("ТОВАРЫ НЕДОСТАЮЩИЕ ХАРАКТЕРИСТИКИ").

ETM generates this CSV (Win-1251, ";") with one row per product and missing
characteristic; the value goes into the last column and the same file is sent
back. The program fills that column from the product's characteristics, matched
to ETM's allowed values, and keeps every other byte of the file as it was.
"""
from __future__ import annotations

import csv
import io
from dataclasses import dataclass, replace
from pathlib import Path

from .db import fetch_products
from .export_templates import PartnerTemplate, load_partner_template
from .filler import FillReportRow, FillResult, make_output_paths, normalize_article, write_report

REPORT_TITLE = "ТОВАРЫ НЕДОСТАЮЩИЕ ХАРАКТЕРИСТИКИ"
ENCODING = "cp1251"
CLASS_COLUMN = "Класс81"
ARTICLE_COLUMN = "Артикул"
NAME_COLUMN = "наименование характеристики"
VALUE_COLUMN = "Поле для заполнения характеристики"
CONFIG_PREFIX = "Конфиг:"


@dataclass
class ReportFile:
    lines: list[str]
    newline: str
    header_index: int
    columns: dict[str, int]


def is_characteristics_report(path: Path) -> bool:
    try:
        with path.open("rb") as fh:
            start = fh.read(200)
    except OSError:
        return False
    return start.decode(ENCODING, errors="ignore").lstrip().upper().startswith(REPORT_TITLE)


def read_report(path: Path) -> ReportFile:
    text = path.read_bytes().decode(ENCODING)
    newline = "\r\n" if "\r\n" in text else "\n"
    lines = text.split(newline)
    for idx, line in enumerate(lines):
        if line.startswith(CLASS_COLUMN):
            header = [cell.strip() for cell in line.split(";")]
            columns = {name: header.index(name) for name in (CLASS_COLUMN, ARTICLE_COLUMN, NAME_COLUMN, VALUE_COLUMN) if name in header}
            missing = {CLASS_COLUMN, ARTICLE_COLUMN, NAME_COLUMN, VALUE_COLUMN} - set(columns)
            if missing:
                raise ValueError(f"В отчете ЭТМ нет колонок: {', '.join(sorted(missing))}.")
            return ReportFile(lines, newline, idx, columns)
    raise ValueError("Не найдена строка заголовка отчета ЭТМ (Класс81;...).")


def split_row(line: str) -> list[str]:
    return next(csv.reader([line], delimiter=";"))


def join_row(cells: list[str]) -> str:
    buffer = io.StringIO()
    csv.writer(buffer, delimiter=";", lineterminator="").writerow(cells)
    return buffer.getvalue()


def clean_code(value: str) -> str:
    return normalize_article(value.strip().lstrip("'"))


def report_rows(report: ReportFile):
    for idx in range(report.header_index + 1, len(report.lines)):
        line = report.lines[idx]
        if not line.strip():
            continue
        cells = split_row(line)
        if len(cells) <= max(report.columns.values()):
            continue
        yield idx, cells


def fits_encoding(value: str) -> bool:
    try:
        value.encode(ENCODING)
    except UnicodeEncodeError:
        return False
    return True


def fill_characteristics_report(
    db_path: Path,
    report_path: Path,
    output_dir: Path,
    rules_path: Path | None = None,
    manual_values_path: Path | None = None,
    template: PartnerTemplate | None = None,
) -> FillResult:
    from .manual_values import load_manual_values

    template = template or load_partner_template("etm")
    columns_spec = template.attribute_columns_for(CONFIG_PREFIX)
    if columns_spec is None:
        raise ValueError(f"В шаблоне {template.name} нет колонок характеристик.")
    report = read_report(report_path)
    cols = report.columns
    rows = list(report_rows(report))
    articles = [clean_code(cells[cols[ARTICLE_COLUMN]]) for _, cells in rows]
    products = fetch_products(db_path, articles)
    manual_values = load_manual_values(manual_values_path)

    details: list[FillReportRow] = []
    filled = suggested = missing = 0
    for line_idx, cells in rows:
        row_number = line_idx + 1
        article = clean_code(cells[cols[ARTICLE_COLUMN]])
        name = cells[cols[NAME_COLUMN]].strip()
        header = f"{CONFIG_PREFIX}{name}"
        if cells[cols[VALUE_COLUMN]].strip():
            details.append(
                FillReportRow(row_number, article, header, "existing_value_preserved", cells[cols[VALUE_COLUMN]].strip(), "template_existing")
            )
            continue
        product = products.get(article)
        if product is None:
            missing += 1
            details.append(FillReportRow(row_number, article, header, "not_found", "", "", "Артикул не найден в базе"))
            continue
        etm_class = clean_code(cells[cols[CLASS_COLUMN]])
        if etm_class:
            product = replace(product, **{template.class_field: etm_class})
        manual_value = manual_values.get((article, header), "")
        if manual_value:
            value, status, source = manual_value, "filled", "manual_completion"
        else:
            value, status, source = template.attribute_value(product, header, columns_spec, rules_path)
        if status in {"filled", "suggested"} and value and not fits_encoding(value):
            missing += 1
            details.append(
                FillReportRow(row_number, article, header, "missing_value", value, source, "Значение содержит символы, которые нельзя сохранить в кодировке файла ЭТМ.")
            )
            continue
        if status == "filled" and value:
            cells[cols[VALUE_COLUMN]] = value
            filled += 1
            details.append(FillReportRow(row_number, article, header, "filled", value, source))
        elif status == "suggested" and value:
            cells[cols[VALUE_COLUMN]] = value
            filled += 1
            suggested += 1
            details.append(
                FillReportRow(
                    row_number,
                    article,
                    header,
                    "filled_suggested",
                    value,
                    source,
                    "Подобрано по похожему значению из справочника ЭТМ; желательно проверить",
                    raw=template.raw_value(product, header, rules_path),
                )
            )
        elif status == "value_not_allowed":
            missing += 1
            details.append(
                FillReportRow(
                    row_number,
                    article,
                    header,
                    status,
                    value,
                    source,
                    template.not_allowed_note(header, value, rules_path),
                    raw=value,
                    candidates=template.candidates_for(header, value, rules_path),
                )
            )
        else:
            missing += 1
            details.append(FillReportRow(row_number, article, header, "missing_value", "", "", template.missing_note(header)))
        report.lines[line_idx] = join_row(cells)

    output_path, details_path = make_output_paths(report_path, output_dir)
    output_path.write_bytes(report.newline.join(report.lines).encode(ENCODING))
    write_report(details_path, details)
    unique = {article for article in articles if article}
    found = len(unique.intersection(products))
    return FillResult(
        output_path=output_path,
        report_path=details_path,
        total_rows=len(rows),
        found_articles=found,
        filled_cells=filled,
        suggested_cells=suggested,
        missing_cells=missing,
        missing_articles=len(unique) - found,
    )


def read_report_values(path: Path) -> list[tuple[str, str, str, str]]:
    """(article, etm class, characteristic, value) for every filled row of a report."""
    report = read_report(path)
    cols = report.columns
    result = []
    for _, cells in report_rows(report):
        value = cells[cols[VALUE_COLUMN]].strip()
        if value:
            result.append(
                (
                    clean_code(cells[cols[ARTICLE_COLUMN]]),
                    clean_code(cells[cols[CLASS_COLUMN]]),
                    cells[cols[NAME_COLUMN]].strip(),
                    value,
                )
            )
    return result
