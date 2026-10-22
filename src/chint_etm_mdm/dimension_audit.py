"""Check the product sizes already in the base: what earlier ETIM imports wrote from wrong columns.

Before 0.5.7 a column such as "Макс. длина кабеля, м" was taken for the product length, so a
cable length could land in the base as the product's size. Every ETIM value keeps the column it
came from, so such values can be found, listed and, after the user agrees, cleared.
"""

from __future__ import annotations

import datetime as dt
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from .db_utils import connect_database
from .etim_importer import (
    MAX_PRODUCT_SIZE_MM,
    SOURCE_KIND,
    chain_order_from_header,
    dimension_label,
    find_overall_dimension_column,
    is_product_size_header,
    format_mm,
    is_size_source,
)
from .number_parsing import parse_number
from .price_importer import backup_database

SIZE_COLUMNS = ("length_mm", "width_mm", "height_mm")
CLEAR_ACTION = "Очистить (значение из ETIM, неверная колонка)"
CHECK_ACTION = "Проверить вручную"


@dataclass(frozen=True)
class AuditFinding:
    article: str
    column: str
    value: str
    problem: str
    source: str
    fixable: bool

    @property
    def field(self) -> str:
        return dimension_label(self.column)


@dataclass(frozen=True)
class AuditResult:
    db_path: Path
    report_path: Path
    products_checked: int
    values_checked: int
    findings: tuple[AuditFinding, ...]

    @property
    def fixable(self) -> tuple[AuditFinding, ...]:
        return tuple(finding for finding in self.findings if finding.fixable)


@dataclass(frozen=True)
class ClearResult:
    db_path: Path
    backup_path: Path
    cleared: int
    changed_since_check: int


def etim_sources(conn) -> dict[tuple[str, str], list[tuple[float, str, str]]]:
    """(article, column) -> every ETIM value written for it, with its sheet and column."""
    sources: dict[tuple[str, str], list[tuple[float, str, str]]] = {}
    rows = conn.execute(
        "SELECT article, dimension_kind, length_mm, width_mm, height_mm, source_sheet, source_field "
        "FROM product_dimensions_raw WHERE source_kind = ?",
        (SOURCE_KIND,),
    )
    for row in rows:
        column = str(row["dimension_kind"] or "")
        if column not in SIZE_COLUMNS:
            continue
        value = parse_number(row[column])
        if value is None:
            continue
        sources.setdefault((str(row["article"]), column), []).append(
            (value, str(row["source_sheet"] or ""), str(row["source_field"] or ""))
        )
    return sources


def is_guessed_chain(field: str) -> bool:
    """A "100x200x300" column whose header does not say which number is which."""
    return (
        not is_product_size_header(field)
        and find_overall_dimension_column([field]) is not None
        and chain_order_from_header(field) is None
    )


def audit_dimensions(db_path: Path, report_dir: Path) -> AuditResult:
    findings: list[AuditFinding] = []
    values_checked = 0
    with closing(connect_database(db_path)) as conn:
        tables = {str(row[0]) for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
        # Only reads: a base that never had sizes or ETIM imports is not changed by the check.
        sources = etim_sources(conn) if "product_dimensions_raw" in tables else {}
        rows = (
            conn.execute(
                "SELECT article, length_mm, width_mm, height_mm FROM product_dimensions_resolved ORDER BY article"
            ).fetchall()
            if "product_dimensions_resolved" in tables
            else []
        )
    for row in rows:
        article = str(row["article"])
        for column in SIZE_COLUMNS:
            raw = row[column]
            if raw is None or str(raw).strip() == "":
                continue
            values_checked += 1
            value = parse_number(raw)
            if value is None:
                findings.append(AuditFinding(article, column, str(raw), "Не число", "", False))
                continue
            matching = [item for item in sources.get((article, column), []) if abs(item[0] - value) <= 0.001]
            known = [item for item in matching if item[2]]
            where = "; ".join(dict.fromkeys(f"ETIM, лист «{sheet}», колонка «{field}»" for _v, sheet, field in known))
            if known and len(known) == len(matching) and not any(is_size_source(field) for _v, _s, field in known):
                findings.append(
                    AuditFinding(article, column, format_mm(value), "Взято из колонки, которая не размер товара", where, True)
                )
            elif value > MAX_PRODUCT_SIZE_MM:
                findings.append(
                    AuditFinding(
                        article,
                        column,
                        format_mm(value),
                        "Больше 10 м",
                        where or "не из ETIM",
                        bool(matching),
                    )
                )
            elif value <= 0:
                findings.append(AuditFinding(article, column, format_mm(value), "Не больше нуля", where, False))
            elif value < 1:
                findings.append(
                    AuditFinding(article, column, format_mm(value), "Меньше 1 мм: возможно, перепутаны единицы", where, False)
                )
            elif known and all(is_guessed_chain(field) for _v, _s, field in known):
                findings.append(
                    AuditFinding(
                        article,
                        column,
                        format_mm(value),
                        "Из общей колонки размеров без указанного порядка: порядок взят как Ш×В×Д, проверьте",
                        where,
                        False,
                    )
                )
    result = AuditResult(
        db_path=db_path,
        report_path=report_dir / f"dimension_check_{dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
        products_checked=len(rows),
        values_checked=values_checked,
        findings=tuple(findings),
    )
    write_audit_report(result)
    return result


def write_audit_report(result: AuditResult) -> None:
    result.report_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Итог"
    for line in audit_lines(result):
        summary.append([line])
    summary.column_dimensions["A"].width = 110

    sheet = workbook.create_sheet("Проверить")
    sheet.append(["Артикул", "Поле", "В базе, мм", "Проблема", "Откуда значение", "Что сделает исправление"])
    for finding in result.findings:
        sheet.append(
            [
                finding.article,
                finding.field,
                finding.value,
                finding.problem,
                finding.source,
                CLEAR_ACTION if finding.fixable else CHECK_ACTION,
            ]
        )
    header_fill = PatternFill(fill_type="solid", fgColor="FFD9EAF7")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column_cells in sheet.columns:
        width = max(len(str(cell.value or "")) for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 80)
    workbook.save(result.report_path)


def clear_wrong_dimensions(db_path: Path, findings: tuple[AuditFinding, ...]) -> ClearResult:
    """Empty the sizes the check found ETIM wrote from wrong columns; anything changed since stays."""
    backup_path = backup_database(db_path)
    cleared = 0
    changed = 0
    with closing(connect_database(db_path)) as conn, conn:
        for finding in findings:
            if not finding.fixable or finding.column not in SIZE_COLUMNS:
                continue
            value = parse_number(finding.value)
            cursor = conn.execute(
                f"UPDATE product_dimensions_resolved SET {finding.column} = NULL, last_resolved_at = CURRENT_TIMESTAMP "
                f"WHERE article = ? AND ABS({finding.column} - ?) <= 0.001",
                (finding.article, value),
            )
            if cursor.rowcount:
                cleared += 1
            else:
                changed += 1
        conn.commit()
    return ClearResult(db_path, backup_path, cleared, changed)


def audit_lines(result: AuditResult) -> list[str]:
    lines = [
        f"База: {result.db_path}",
        f"Товаров с размерами: {result.products_checked}",
        f"Значений проверено: {result.values_checked}",
        f"Найдено сомнительных: {len(result.findings)}",
        f"Из них можно очистить автоматически (пришли из ETIM из неподходящей колонки): {len(result.fixable)}",
    ]
    by_source: dict[str, int] = {}
    for finding in result.fixable:
        for part in finding.source.split("; "):
            column = part.rsplit("колонка ", 1)[-1]
            by_source[column] = by_source.get(column, 0) + 1
    for column, count in sorted(by_source.items(), key=lambda item: -item[1])[:10]:
        lines.append(f"  {column}: {count}")
    lines.append(f"Отчёт: {result.report_path}")
    return lines


def clear_lines(result: ClearResult) -> list[str]:
    lines = [
        f"Очищено значений: {result.cleared}",
        f"Бэкап базы до очистки: {result.backup_path}",
        "Чтобы подтянуть правильные размеры, импортируйте ETIM-файл ещё раз: пустые размеры заполнятся "
        "из правильных колонок.",
    ]
    if result.changed_since_check:
        lines.insert(1, f"Не тронуто (изменились после проверки): {result.changed_since_check}")
    return lines
