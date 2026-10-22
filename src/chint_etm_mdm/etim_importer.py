from __future__ import annotations

import datetime as dt
import re
import sqlite3
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill

from .db_utils import CHINT_TABLES, connect_database, restore_sqlite_backup
from .number_parsing import parse_number
from .price_importer import backup_database


ARTICLE_HINT = "артикул"
DIMENSION_HINTS = {
    "length_mm": ("глубин", "длин"),
    "width_mm": ("ширин",),
    "height_mm": ("высот",),
}
DIMENSION_EXCLUDE_HINTS = {
    "length_mm": (),
    "width_mm": (),
    "height_mm": ("уровнем моря", "над уровнем", "altitude"),
}
OVERALL_DIMENSION_HINT = "габарит"
OVERALL_DIMENSION_SECONDARY_HINT = "размер"
SOURCE_KIND = "etim_manual"
KEEP_DB_DECISION = "Оставить базу"
ACCEPT_ETIM_DECISION = "Принять ETIM"
FIELD_TO_COLUMN = {
    "Длина, мм": "length_mm",
    "Ширина, мм": "width_mm",
    "Высота, мм": "height_mm",
}


@dataclass(frozen=True)
class EtimArticleData:
    article: str
    sheet_name: str
    row_number: int
    dimensions: dict[str, float]
    dimension_sources: dict[str, str]
    attributes: dict[str, str]


@dataclass(frozen=True)
class EtimImportResult:
    etim_file: Path
    db_path: Path
    report_path: Path
    backup_path: Path | None
    import_id: int
    scanned_articles: int
    found_articles: int
    dimension_values_found: int
    dimension_values_written: int
    dimension_conflicts: int
    attributes_written: int


@dataclass(frozen=True)
class EtimDecisionApplyResult:
    report_path: Path
    db_path: Path
    backup_path: Path
    accepted_rows: int
    skipped_rows: int
    already_applied_rows: int = 0
    rejected: tuple[str, ...] = ()


@dataclass(frozen=True)
class DimensionReportRow:
    row_type: str
    article: str
    field: str
    current_value: str
    etim_value: str
    source_sheet: str
    source_field: str
    action: str


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return " ".join(text.split())


UNIT_FACTORS_MM = {"мм": 1.0, "mm": 1.0, "см": 10.0, "cm": 10.0, "м": 1000.0, "m": 1000.0}
UNIT_RE = re.compile(r"(?<![^\W\d_])(мм|mm|см|cm|м|m)(?![^\W\d_])", re.IGNORECASE)
CHAIN_SEPARATOR = r"\s*[xXхХ×*]\s*"
ORDER_LETTERS = {
    "Д": "length_mm", "Г": "length_mm", "L": "length_mm", "D": "length_mm",
    "Ш": "width_mm", "W": "width_mm",
    "В": "height_mm", "H": "height_mm",
}
ORDER_WORDS = {
    "длин": "length_mm", "глубин": "length_mm", "length": "length_mm", "depth": "length_mm",
    "ширин": "width_mm", "width": "width_mm",
    "высот": "height_mm", "height": "height_mm",
}
# Historic files without an order in the header were read as Ш×В×Д; keep that only for 3 numbers.
LEGACY_CHAIN_ORDER = ("width_mm", "height_mm", "length_mm")


def unit_factor(text: str) -> float | None:
    units = {match.group(1).casefold() for match in UNIT_RE.finditer(text)}
    factors = {UNIT_FACTORS_MM[unit] for unit in units}
    if len(factors) != 1:
        return None
    return factors.pop()


def parse_dimension_number(value: object, header: object) -> float | None:
    """A positive size in mm, or None when the cell is not one unambiguous number.

    Units come from the value (``12 см``) or the header (``Длина, см``); with
    neither, millimetres are assumed. Ranges, limits and words are not guessed.
    """
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number, value_unit = parse_number(value), None
    else:
        text = normalize_text(value)
        match = re.fullmatch(r"(.*?)\s*(мм|mm|см|cm|м|m)\.?", text, flags=re.IGNORECASE)
        number_text, value_unit = (match.group(1), match.group(2).casefold()) if match else (text, None)
        number = parse_number(number_text)
    if number is None or number <= 0:
        return None
    if value_unit is not None:
        return number * UNIT_FACTORS_MM[value_unit]
    header_factor = unit_factor(normalize_text(header))
    if header_factor is None and UNIT_RE.search(normalize_text(header)):
        return None
    return number * (header_factor or 1.0)


def chain_order_from_header(header: object) -> tuple[str, ...] | None:
    text = normalize_text(header)
    letters = "".join(ORDER_LETTERS)
    match = re.search(
        rf"(?<![^\W\d_])([{letters}](?:{CHAIN_SEPARATOR}[{letters}])+)(?![^\W\d_])", text
    )
    if match:
        keys = [ORDER_LETTERS[char] for char in re.findall(rf"[{letters}]", match.group(1))]
    else:
        words = "|".join(ORDER_WORDS)
        match = re.search(
            rf"((?:{words})\w*(?:{CHAIN_SEPARATOR}(?:{words})\w*)+)", text, flags=re.IGNORECASE
        )
        if not match:
            return None
        keys = [
            next(key for word, key in ORDER_WORDS.items() if part.casefold().startswith(word))
            for part in re.split(CHAIN_SEPARATOR, match.group(1))
        ]
    return tuple(keys) if len(set(keys)) == len(keys) else None


def parse_overall_dimensions(value: object, header: object = "") -> dict[str, float]:
    """Split a ``100x200x300`` cell using the order named in the header.

    Returns {} for anything ambiguous: no recognisable chain, several chains,
    a count that does not match the header order, non-positive numbers or
    conflicting units.
    """
    text = normalize_text(value)
    if not text:
        return {}
    number = r"\d+(?:[.,]\d+)?"
    chains = re.findall(rf"{number}(?:{CHAIN_SEPARATOR}{number})+", text)
    if len(chains) != 1:
        return {}
    remainder = text.replace(chains[0], " ", 1).strip(" ,.;")
    if remainder and not UNIT_RE.fullmatch(remainder.rstrip(".")):
        return {}
    values = [parse_number(item) for item in re.split(CHAIN_SEPARATOR, chains[0])]
    if any(item is None or item <= 0 for item in values):
        return {}

    order = chain_order_from_header(header)
    if order is None and len(values) == len(LEGACY_CHAIN_ORDER):
        order = LEGACY_CHAIN_ORDER
    if order is None or len(order) != len(values):
        return {}

    header_text = normalize_text(header)
    if remainder:
        factor = UNIT_FACTORS_MM[remainder.rstrip(".").casefold()]
    elif UNIT_RE.search(header_text):
        factor = unit_factor(header_text)
        if factor is None:
            return {}
    else:
        factor = 1.0
    return {key: float(item) * factor for key, item in zip(order, values)}


def parse_dimension_cell(value: object, header: object, dimension_key: str) -> float | None:
    header_text = normalize_text(header).casefold()
    if any(hint in header_text for hint in DIMENSION_EXCLUDE_HINTS.get(dimension_key, ())):
        return None
    return parse_dimension_number(value, header)


def find_dimension_column(
    headers: list[object],
    hints: Iterable[str],
    exclude_hints: Iterable[str],
    skip_columns: Iterable[int] = (),
) -> int | None:
    normalized = [normalize_text(header).casefold() for header in headers]
    skipped = set(skip_columns)
    candidates = [
        index
        for index, header in enumerate(normalized)
        if index not in skipped
        and any(hint in header for hint in hints)
        and not any(exclude in header for exclude in exclude_hints)
        and chain_order_from_header(headers[index]) is None
    ]
    for index in candidates:
        if UNIT_RE.search(normalized[index]):
            return index
    for index in candidates:
        if any(marker in normalized[index] for marker in ("габарит", "размер", "dimension")):
            return index
    return None


def find_overall_dimension_column(headers: list[object]) -> int | None:
    normalized = [normalize_text(header).casefold() for header in headers]
    for index, header in enumerate(normalized):
        if OVERALL_DIMENSION_HINT in header and chain_order_from_header(headers[index]) is not None:
            return index
    for index, header in enumerate(normalized):
        if OVERALL_DIMENSION_HINT in header and OVERALL_DIMENSION_SECONDARY_HINT in header:
            return index
    # "Габаритная длина" is a single dimension, not an overall chain.
    single_hints = [hint for hints in DIMENSION_HINTS.values() for hint in hints]
    for index, header in enumerate(normalized):
        if OVERALL_DIMENSION_HINT in header and not any(hint in header for hint in single_hints):
            return index
    return None


def find_header_row(rows: list[tuple[object, ...]], row_index: int, column_index: int) -> int | None:
    for index in range(row_index - 1, -1, -1):
        if column_index < len(rows[index]) and ARTICLE_HINT in normalize_text(rows[index][column_index]).casefold():
            return index
    return None


def collect_attributes(headers: list[object], row: tuple[object, ...], article_column: int) -> dict[str, str]:
    attrs: dict[str, str] = {}
    for index, header_value in enumerate(headers):
        if index == article_column:
            continue
        header = normalize_text(header_value)
        if not header or ARTICLE_HINT == header.casefold():
            continue
        value = normalize_text(row[index] if index < len(row) else "")
        if value:
            attrs.setdefault(header, value)
    return attrs


def scan_etim_workbook(etim_path: Path, articles: Iterable[str]) -> dict[str, EtimArticleData]:
    if etim_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("ETIM-файл должен быть XLSX/XLSM.")

    articles_to_find = {normalize_text(article) for article in articles if normalize_text(article)}
    if not articles_to_find:
        return {}

    workbook = load_workbook(etim_path, read_only=True, data_only=True)
    found: dict[str, EtimArticleData] = {}
    try:
        for sheet in workbook.worksheets:
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                continue
            for row_index, row in enumerate(rows):
                if len(found) == len(articles_to_find):
                    return found
                for column_index, cell_value in enumerate(row):
                    article = normalize_text(cell_value)
                    if article not in articles_to_find or article in found:
                        continue
                    header_row_index = find_header_row(rows, row_index, column_index)
                    if header_row_index is None:
                        continue
                    headers = list(rows[header_row_index])
                    dimensions: dict[str, float] = {}
                    dimension_sources: dict[str, str] = {}

                    overall_column = find_overall_dimension_column(headers)

                    # Separately labelled columns are exact; the overall chain only fills gaps.
                    for key, hints in DIMENSION_HINTS.items():
                        column = find_dimension_column(
                            headers,
                            hints,
                            DIMENSION_EXCLUDE_HINTS.get(key, ()),
                            skip_columns=() if overall_column is None else (overall_column,),
                        )
                        if column is None or column >= len(row):
                            continue
                        parsed_value = parse_dimension_cell(row[column], headers[column], key)
                        if parsed_value is not None:
                            dimensions[key] = parsed_value
                            dimension_sources[key] = normalize_text(headers[column])

                    if overall_column is not None and overall_column < len(row):
                        parsed = parse_overall_dimensions(row[overall_column], headers[overall_column])
                        for key, value in parsed.items():
                            if key in dimensions:
                                continue
                            dimensions[key] = value
                            dimension_sources[key] = normalize_text(headers[overall_column])

                    found[article] = EtimArticleData(
                        article=article,
                        sheet_name=sheet.title,
                        row_number=row_index + 1,
                        dimensions=dimensions,
                        dimension_sources=dimension_sources,
                        attributes=collect_attributes(headers, row, column_index),
                    )
    finally:
        workbook.close()
    return found


def ensure_etim_import_tables(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS dimension_imports (
            id INTEGER PRIMARY KEY,
            source_kind TEXT NOT NULL,
            local_file TEXT,
            workbook_name TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS product_dimensions_raw (
            id INTEGER PRIMARY KEY,
            import_id INTEGER,
            article TEXT NOT NULL,
            source_kind TEXT NOT NULL,
            dimension_kind TEXT NOT NULL,
            length_mm REAL,
            width_mm REAL,
            height_mm REAL,
            weight_kg REAL,
            volume_m3 REAL,
            source_sheet TEXT,
            source_field TEXT,
            value_raw TEXT,
            confidence REAL DEFAULT 1.0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(import_id) REFERENCES dimension_imports(id)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS product_dimensions_resolved (
            article TEXT PRIMARY KEY,
            length_mm REAL,
            width_mm REAL,
            height_mm REAL,
            weight_kg REAL,
            volume_m3 REAL,
            dimensions_source TEXT,
            weight_source TEXT,
            volume_source TEXT,
            last_resolved_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS product_attribute_values (
            id INTEGER PRIMARY KEY,
            article TEXT NOT NULL,
            attribute_code TEXT,
            attribute_name TEXT,
            sheet_name TEXT,
            etim_class_code TEXT,
            value_raw TEXT,
            value_normalized TEXT,
            attribute_group TEXT,
            source_kind TEXT,
            source_ref TEXT,
            confidence REAL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )


def fetch_all_articles(conn: sqlite3.Connection) -> list[str]:
    return [
        str(row[0])
        for row in conn.execute(
            "SELECT article FROM products WHERE COALESCE(article, '') <> '' ORDER BY article"
        ).fetchall()
    ]


def dimension_label(key: str) -> str:
    labels = {
        "length_mm": "Длина, мм",
        "width_mm": "Ширина, мм",
        "height_mm": "Высота, мм",
    }
    return labels.get(key, key)


def format_mm(value: object) -> str:
    if value is None or value == "":
        return ""
    number = float(value)
    if number.is_integer():
        return str(int(number))
    return f"{number:.3f}".rstrip("0").rstrip(".")


def build_dimension_report_rows(
    conn: sqlite3.Connection, article_data: dict[str, EtimArticleData]
) -> tuple[int, int, list[DimensionReportRow], list[DimensionReportRow]]:
    written = 0
    conflicts = 0
    new_rows: list[DimensionReportRow] = []
    conflict_rows: list[DimensionReportRow] = []
    for data in article_data.values():
        if not data.dimensions:
            continue
        row = conn.execute(
            """
            SELECT length_mm, width_mm, height_mm
            FROM product_dimensions_resolved
            WHERE article = ?
            """,
            (data.article,),
        ).fetchone()
        existing = dict(row) if row else {}
        for key, value in data.dimensions.items():
            current = existing.get(key)
            source_field = data.dimension_sources.get(key, "")
            if current is None:
                written += 1
                new_rows.append(
                    DimensionReportRow(
                        row_type="new",
                        article=data.article,
                        field=dimension_label(key),
                        current_value="",
                        etim_value=format_mm(value),
                        source_sheet=data.sheet_name,
                        source_field=source_field,
                        action="Будет записано в базу",
                    )
                )
            elif abs(float(current) - float(value)) > 0.001:
                conflicts += 1
                conflict_rows.append(
                    DimensionReportRow(
                        row_type="conflict",
                        article=data.article,
                        field=dimension_label(key),
                        current_value=format_mm(current),
                        etim_value=format_mm(value),
                        source_sheet=data.sheet_name,
                        source_field=source_field,
                        action="Проверить вручную: база не перезаписана",
                    )
                )
    return written, conflicts, new_rows, conflict_rows


def write_etim_report(
    report_path: Path,
    result: EtimImportResult,
    new_rows: list[DimensionReportRow],
    conflict_rows: list[DimensionReportRow],
) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Итог"
    summary.append(["Показатель", "Значение"])
    summary.append(["ETIM-файл", str(result.etim_file)])
    summary.append(["База", str(result.db_path)])
    summary.append(["Импорт id", result.import_id])
    summary.append(["Артикулов в базе для поиска", result.scanned_articles])
    summary.append(["Артикулов найдено в ETIM", result.found_articles])
    summary.append(["Габаритных значений найдено", result.dimension_values_found])
    summary.append(["Новых габаритных значений записано", result.dimension_values_written])
    summary.append(["Конфликтов с уже заполненными габаритами", result.dimension_conflicts])
    summary.append(["Характеристик записано", result.attributes_written])
    summary.append(["Бэкап базы", str(result.backup_path or "")])

    instruction = workbook.create_sheet("Как разобрать")
    instruction.append(["Что делать с конфликтами"])
    instruction.append(["1. Откройте лист 'Конфликты'."])
    instruction.append(["2. Сравните 'Сейчас в базе' и 'Пришло из ETIM'."])
    instruction.append(["3. В колонке 'Решение' выберите один вариант из списка."])
    instruction.append([f"4. '{KEEP_DB_DECISION}' означает: ничего не менять."])
    instruction.append([f"5. '{ACCEPT_ETIM_DECISION}' означает: перезаписать это поле в базе значением из ETIM."])
    instruction.append(["6. Сохраните отчет и примените решения в программе."])

    new_sheet = workbook.create_sheet("Новые габариты")
    new_sheet.append(["Артикул", "Поле", "Записано из ETIM", "Лист ETIM", "Колонка ETIM", "Действие"])
    for row in new_rows:
        new_sheet.append([row.article, row.field, row.etim_value, row.source_sheet, row.source_field, row.action])

    conflicts = workbook.create_sheet("Конфликты")
    conflicts.append(
        [
            "Решение",
            "Артикул",
            "Поле",
            "Сейчас в базе",
            "Пришло из ETIM",
            "Лист ETIM",
            "Колонка ETIM",
            "Что делать",
        ]
    )
    for row in conflict_rows:
        conflicts.append(
            [
                KEEP_DB_DECISION,
                row.article,
                row.field,
                row.current_value,
                row.etim_value,
                row.source_sheet,
                row.source_field,
                row.action,
            ]
        )

    decision_validation = DataValidation(
        type="list",
        formula1=f'"{KEEP_DB_DECISION},{ACCEPT_ETIM_DECISION}"',
        allow_blank=False,
    )
    conflicts.add_data_validation(decision_validation)
    if conflict_rows:
        decision_validation.add(f"A2:A{len(conflict_rows) + 1}")

    header_fill = PatternFill(fill_type="solid", fgColor="FFD9EAF7")
    keep_fill = PatternFill(fill_type="solid", fgColor="FFFFF2CC")
    accept_fill = PatternFill(fill_type="solid", fgColor="FFD9EAD3")
    for cell in conflicts[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
    for row_index in range(2, len(conflict_rows) + 2):
        conflicts.cell(row_index, 1).fill = keep_fill
        conflicts.cell(row_index, 5).fill = accept_fill

    for sheet in (summary, instruction, new_sheet, conflicts):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 70)

    workbook.save(report_path)


def insert_etim_data(
    conn: sqlite3.Connection,
    import_id: int,
    article_data: dict[str, EtimArticleData],
    source_ref: str,
) -> tuple[int, int]:
    dimension_values_found = 0
    attributes_written = 0
    for data in article_data.values():
        for key, value in data.dimensions.items():
            dimension_values_found += 1
            length_mm = value if key == "length_mm" else None
            width_mm = value if key == "width_mm" else None
            height_mm = value if key == "height_mm" else None
            source_field = data.dimension_sources.get(key, "")
            conn.execute(
                """
                INSERT INTO product_dimensions_raw (
                    import_id, article, source_kind, dimension_kind,
                    length_mm, width_mm, height_mm, source_sheet, source_field,
                    value_raw, confidence
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    import_id,
                    data.article,
                    SOURCE_KIND,
                    key,
                    length_mm,
                    width_mm,
                    height_mm,
                    data.sheet_name,
                    source_field,
                    str(value),
                    0.95,
                ),
            )
        if data.dimensions:
            conn.execute(
                """
                INSERT INTO product_dimensions_resolved (
                    article, length_mm, width_mm, height_mm, dimensions_source, last_resolved_at
                ) VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(article) DO UPDATE SET
                    length_mm=COALESCE(product_dimensions_resolved.length_mm, excluded.length_mm),
                    width_mm=COALESCE(product_dimensions_resolved.width_mm, excluded.width_mm),
                    height_mm=COALESCE(product_dimensions_resolved.height_mm, excluded.height_mm),
                    dimensions_source=CASE
                        WHEN product_dimensions_resolved.dimensions_source IS NULL
                             OR product_dimensions_resolved.dimensions_source = ''
                        THEN excluded.dimensions_source
                        ELSE product_dimensions_resolved.dimensions_source
                    END,
                    last_resolved_at=CURRENT_TIMESTAMP
                """,
                (
                    data.article,
                    data.dimensions.get("length_mm"),
                    data.dimensions.get("width_mm"),
                    data.dimensions.get("height_mm"),
                    SOURCE_KIND,
                ),
            )

        for attribute_name, value in data.attributes.items():
            conn.execute(
                """
                INSERT INTO product_attribute_values (
                    article, attribute_name, sheet_name, value_raw, value_normalized,
                    source_kind, source_ref, confidence
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data.article,
                    attribute_name,
                    data.sheet_name,
                    value,
                    value,
                    SOURCE_KIND,
                    source_ref,
                    0.8,
                ),
            )
            attributes_written += 1
    return dimension_values_found, attributes_written


def import_etim_workbook(
    etim_path: Path,
    db_path: Path,
    make_backup: bool = True,
    report_dir: Path | None = None,
) -> EtimImportResult:
    if not etim_path.exists():
        raise FileNotFoundError(f"ETIM-файл не найден: {etim_path}")
    if not db_path.exists():
        raise FileNotFoundError(f"База не найдена: {db_path}")

    backup_path = backup_database(db_path) if make_backup else None
    with closing(connect_database(db_path)) as conn, conn:
        ensure_etim_import_tables(conn)
        articles = fetch_all_articles(conn)
        scanned_articles = len(articles)

    article_data = scan_etim_workbook(etim_path, articles)

    with closing(connect_database(db_path)) as conn, conn:
        ensure_etim_import_tables(conn)
        (
            dimension_values_written,
            dimension_conflicts,
            new_dimension_rows,
            conflict_rows,
        ) = build_dimension_report_rows(conn, article_data)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO dimension_imports (source_kind, local_file, workbook_name, notes)
            VALUES (?, ?, ?, ?)
            """,
            (
                SOURCE_KIND,
                str(etim_path),
                etim_path.name,
                f"ETIM import {dt.datetime.now().isoformat(timespec='seconds')}",
            ),
        )
        import_id = int(cur.lastrowid)
        dimension_values_found, attributes_written = insert_etim_data(
            conn, import_id, article_data, etim_path.name
        )
        conn.commit()

    report_dir = report_dir or db_path.parent / "reports"
    report_path = report_dir / f"etim_import_report_{dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    result = EtimImportResult(
        etim_file=etim_path,
        db_path=db_path,
        report_path=report_path,
        backup_path=backup_path,
        import_id=import_id,
        scanned_articles=scanned_articles,
        found_articles=len(article_data),
        dimension_values_found=dimension_values_found,
        dimension_values_written=dimension_values_written,
        dimension_conflicts=dimension_conflicts,
        attributes_written=attributes_written,
    )
    write_etim_report(report_path, result, new_dimension_rows, conflict_rows)
    return result


def parse_float(value: object) -> float | None:
    return parse_number(normalize_text(value))


def report_import_id(workbook) -> int | None:
    if "Итог" not in workbook.sheetnames:
        return None
    for row in workbook["Итог"].iter_rows(values_only=True):
        if row and normalize_text(row[0]) == "Импорт id" and len(row) > 1:
            number = parse_number(row[1])
            return int(number) if number is not None and number.is_integer() else None
    return None


def same_value(left: float | None, right: float | None) -> bool:
    if left is None or right is None:
        return left is None and right is None
    return abs(float(left) - float(right)) <= 0.001


def apply_etim_conflict_decisions(report_path: Path, db_path: Path) -> EtimDecisionApplyResult:
    """Apply 'Принять ETIM' rows only where the database still holds the reported value.

    A row is rejected (and the field left untouched) when the database changed
    since the report, the product or the ETIM value is unknown to this
    database/import, the value is not a positive finite number, or the same
    field has contradicting decisions.
    """
    if not report_path.exists():
        raise FileNotFoundError(f"Отчет ETIM не найден: {report_path}")
    if not db_path.exists():
        raise FileNotFoundError(f"База не найдена: {db_path}")

    workbook = load_workbook(report_path, read_only=True, data_only=True)
    try:
        if "Конфликты" not in workbook.sheetnames:
            raise ValueError("В отчете нет листа 'Конфликты'.")
        import_id = report_import_id(workbook)
        sheet = workbook["Конфликты"]
        rows = list(sheet.iter_rows(values_only=True))
    finally:
        workbook.close()

    if not rows:
        raise ValueError("Лист 'Конфликты' пустой.")
    headers = [normalize_text(value) for value in rows[0]]
    required = ["Решение", "Артикул", "Поле", "Сейчас в базе", "Пришло из ETIM"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError(
            f"В листе 'Конфликты' нет колонок: {', '.join(missing)}. "
            "Без них нельзя проверить, что отчет относится к текущей базе; пересоздайте отчет импортом ETIM."
        )
    index = {name: headers.index(name) for name in required}

    def cell(row: tuple[object, ...], name: str) -> object:
        return row[index[name]] if index[name] < len(row) else None

    candidates: dict[tuple[str, str], list[tuple[int, float | None, float]]] = {}
    rejected: list[str] = []
    skipped_rows = 0
    for row_number, row in enumerate(rows[1:], start=2):
        if not any(value not in (None, "") for value in row):
            continue
        decision = normalize_text(cell(row, "Решение"))
        if decision != ACCEPT_ETIM_DECISION:
            skipped_rows += 1
            continue
        article = normalize_text(cell(row, "Артикул"))
        field = normalize_text(cell(row, "Поле"))
        column = FIELD_TO_COLUMN.get(field)
        reported_current_text = normalize_text(cell(row, "Сейчас в базе"))
        reported_current = parse_float(reported_current_text)
        value = parse_float(cell(row, "Пришло из ETIM"))
        if not article or not column:
            rejected.append(f"строка {row_number}: не указан артикул или неизвестное поле '{field}'")
        elif value is None or value <= 0:
            rejected.append(f"строка {row_number}, {article}: значение ETIM не является положительным числом")
        elif reported_current_text and reported_current is None:
            rejected.append(f"строка {row_number}, {article}: не читается значение 'Сейчас в базе'")
        else:
            candidates.setdefault((article, column), []).append((row_number, reported_current, value))

    accepted = 0
    already_applied = 0
    backup_path = backup_database(db_path)
    with closing(connect_database(db_path)) as conn, conn:
        ensure_etim_import_tables(conn)
        for (article, column), decisions in candidates.items():
            row_number = decisions[0][0]
            label = f"строка {row_number}, {article}, {dimension_label(column)}"
            if len({(current, round(value, 3)) for _row, current, value in decisions}) > 1:
                rows_text = ", ".join(str(item[0]) for item in decisions)
                rejected.append(f"строки {rows_text}, {article}: противоречивые решения по одному полю")
                continue
            _row, reported_current, value = decisions[0]
            if conn.execute("SELECT 1 FROM products WHERE article = ?", (article,)).fetchone() is None:
                rejected.append(f"{label}: товара нет в этой базе")
                continue
            source_query = (
                f"SELECT 1 FROM product_dimensions_raw WHERE article = ? AND source_kind = ? "
                f"AND dimension_kind = ? AND ABS({column} - ?) <= 0.001"
            )
            source_params: list[object] = [article, SOURCE_KIND, column, value]
            if import_id is not None:
                source_query += " AND import_id = ?"
                source_params.append(import_id)
            if conn.execute(source_query, source_params).fetchone() is None:
                rejected.append(f"{label}: значение {format_mm(value)} не найдено в импорте ETIM этой базы")
                continue
            current_row = conn.execute(
                f"SELECT {column} FROM product_dimensions_resolved WHERE article = ?", (article,)
            ).fetchone()
            current = parse_number(current_row[0]) if current_row else None
            if same_value(current, value):
                already_applied += 1
                continue
            if not same_value(current, reported_current):
                rejected.append(
                    f"{label}: в базе уже {format_mm(current) or 'пусто'}, а в отчете "
                    f"{format_mm(reported_current) or 'пусто'}; база изменилась после отчета"
                )
                continue
            # Compare-and-swap: write only if the field still holds the reported value.
            cursor = conn.execute(
                f"""
                UPDATE product_dimensions_resolved
                SET {column} = ?, dimensions_source = ?, last_resolved_at = CURRENT_TIMESTAMP
                WHERE article = ? AND (
                    ({column} IS NULL AND ? IS NULL) OR ABS({column} - ?) <= 0.001
                )
                """,
                (value, SOURCE_KIND, article, reported_current, reported_current),
            )
            if cursor.rowcount == 0 and current_row is None:
                conn.execute(
                    f"""
                    INSERT INTO product_dimensions_resolved (article, {column}, dimensions_source, last_resolved_at)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                    """,
                    (article, value, SOURCE_KIND),
                )
            elif cursor.rowcount == 0:
                rejected.append(f"{label}: база изменилась во время применения")
                continue
            accepted += 1
        conn.commit()

    return EtimDecisionApplyResult(
        report_path=report_path,
        db_path=db_path,
        backup_path=backup_path,
        accepted_rows=accepted,
        skipped_rows=skipped_rows,
        already_applied_rows=already_applied,
        rejected=tuple(rejected),
    )


def restore_database_backup(db_path: Path, backup_path: Path) -> Path:
    return restore_sqlite_backup(db_path, backup_path, CHINT_TABLES).path


def result_lines(result: EtimImportResult) -> list[str]:
    lines = [
        f"ETIM-файл: {result.etim_file}",
        f"База: {result.db_path}",
        f"Импорт id: {result.import_id}",
        f"Артикулов в базе для поиска: {result.scanned_articles}",
        f"Артикулов найдено в ETIM: {result.found_articles}",
        f"Габаритных значений найдено: {result.dimension_values_found}",
        f"Новых габаритных значений записано: {result.dimension_values_written}",
        f"Конфликтов с уже заполненными габаритами: {result.dimension_conflicts}",
        f"Характеристик записано: {result.attributes_written}",
        f"Отчет ETIM: {result.report_path}",
    ]
    if result.backup_path:
        lines.append(f"Бэкап базы: {result.backup_path}")
    return lines


def decision_result_lines(result: EtimDecisionApplyResult) -> list[str]:
    return [
        f"Отчет ETIM: {result.report_path}",
        f"База: {result.db_path}",
        f"Принято ETIM-значений: {result.accepted_rows}",
        f"Оставлено без изменений: {result.skipped_rows}",
        f"Уже было применено раньше: {result.already_applied_rows}",
        f"Отклонено строк: {len(result.rejected)}",
        *(f"  - {reason}" for reason in result.rejected),
        f"Бэкап базы: {result.backup_path}",
    ]
