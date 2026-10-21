from __future__ import annotations

import datetime as dt
import re
import csv
import shutil
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Font, PatternFill

from .db import ProductRecord, fetch_products
from .export_templates import (
    DEFAULT_PARTNER,
    PartnerTemplate,
    load_partner_template,
)


MISSING_CELL_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
MISSING_CELL_FONT = Font(color="FF9C0006")
DATA_VALIDATIONS_RE = re.compile(br"<dataValidations\b.*?</dataValidations>", re.DOTALL)
EXT_LST_RE = re.compile(br"<extLst\b.*?</extLst>", re.DOTALL)
WORKSHEET_ROOT_RE = re.compile(br"<worksheet\b[^>]*>")
XMLNS_ATTR_RE = re.compile(br'\s+xmlns(?::[A-Za-z_][\w.-]*)?="[^"]*"')
MC_IGNORABLE_RE = re.compile(br'\s+mc:Ignorable="([^"]*)"')


@dataclass(frozen=True)
class FillReportRow:
    row_number: int
    article: str
    column: str
    status: str
    value: str
    source: str
    note: str = ""
    raw: str = ""
    candidates: tuple[str, ...] = ()


MATCHES_SHEET = "Сопоставления"
MATCHES_HEADERS = [
    "Поле",
    "Значение в базе",
    "Предложено программой",
    "Товаров",
    "Похожие допустимые значения",
    "Принять значение",
]


@dataclass(frozen=True)
class FillResult:
    output_path: Path
    report_path: Path
    total_rows: int
    found_articles: int
    filled_cells: int
    suggested_cells: int
    missing_cells: int
    missing_articles: int


def normalize_header(value: object) -> str:
    return str(value or "").replace("\ufeff", "").strip()


def is_yellow_header_cell(cell: Cell) -> bool:
    fill = cell.fill
    if fill.fill_type != "solid":
        return False
    rgb = str(fill.fgColor.rgb or "").upper()
    return rgb in {"FFFFFF00", "FFFF00"} or rgb.endswith("FFFF00")


def normalize_article(value: object) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def comparable_text(value: str) -> str:
    return re.sub(r"[^0-9a-zа-я]+", "", value.lower())


def partner_template(template: PartnerTemplate | None = None) -> PartnerTemplate:
    return template or load_partner_template(DEFAULT_PARTNER)


def product_value(
    product: ProductRecord,
    header: str,
    rules_path: Path | None = None,
    template: PartnerTemplate | None = None,
) -> tuple[str, str, str]:
    return partner_template(template).product_value(product, header, rules_path)


def is_reportable_header(header: str, template: PartnerTemplate | None = None) -> bool:
    return partner_template(template).is_reportable(header)


def missing_note(header: str, template: PartnerTemplate | None = None) -> str:
    return partner_template(template).missing_note(header)


def report_status_label(status: str) -> str:
    labels = {
        "filled": "Заполнено",
        "filled_suggested": "Заполнено, нужно проверить",
        "existing_value_preserved": "Существующее значение сохранено",
        "missing_value": "Не заполнено",
        "value_not_allowed": "Значения нет в списке партнёра",
        "not_found": "Артикул не найден",
        "ok": "Готово",
        "needs_review": "Нужно проверить",
        "incomplete": "Заполнено не полностью",
    }
    return labels.get(status, status)


def source_label(source: str) -> str:
    if not source:
        return ""
    if source == "static":
        return "Значение по умолчанию"
    if source == "generated":
        return "Сформировано программой"
    if source.startswith("approved_class_rule:"):
        return f"Выбранный источник: {source.split(':', 1)[1]}"
    if source == "manual_completion":
        return "Ручное дозаполнение"
    if source == "template_existing":
        return "Уже было в шаблоне"
    if "attribute" in source:
        return "Характеристика товара в базе"
    if "article" in source:
        return "Артикул из шаблона"
    if "class81" in source:
        return "Категория из базы"
    if "name" in source:
        return "Название из базы"
    if "tnved" in source.lower() or "okpd2" in source.lower():
        return "Код из базы"
    if "weight" in source or "length" in source or "width" in source or "height" in source or "volume" in source:
        return "Вес или габариты из базы"
    if "shipment" in source or "gtin" in source or "transport" in source:
        return "Данные упаковки из базы"
    return "База"


def mark_missing_cell(cell: Cell) -> None:
    cell.fill = MISSING_CELL_FILL
    cell.font = MISSING_CELL_FONT


def preserve_worksheet_dropdowns(template_path: Path, output_path: Path) -> None:
    """Restore worksheet validation XML that openpyxl may simplify or drop."""
    if template_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return

    with ZipFile(template_path, "r") as source_zip, ZipFile(output_path, "r") as output_zip:
        source_names = set(source_zip.namelist())
        output_names = set(output_zip.namelist())
        worksheet_names = sorted(
            name
            for name in source_names.intersection(output_names)
            if name.startswith("xl/worksheets/") and name.endswith(".xml")
        )
        replacements: dict[str, bytes] = {}
        for name in worksheet_names:
            source_xml = source_zip.read(name)
            output_xml = output_zip.read(name)
            source_validations = DATA_VALIDATIONS_RE.search(source_xml)
            source_extensions = EXT_LST_RE.search(source_xml)
            if not source_validations and not source_extensions:
                continue

            updated_xml = DATA_VALIDATIONS_RE.sub(b"", output_xml)
            if source_validations:
                updated_xml = insert_after_xml_tag(
                    updated_xml, b"</sheetData>", source_validations.group(0)
                )

            if source_extensions and b"dataValidations" in source_extensions.group(0):
                updated_xml = EXT_LST_RE.sub(b"", updated_xml)
                updated_xml = updated_xml.replace(
                    b"</worksheet>", source_extensions.group(0) + b"</worksheet>", 1
                )

            updated_xml = copy_worksheet_root_namespaces(source_xml, updated_xml)

            if updated_xml != output_xml:
                replacements[name] = updated_xml

        if not replacements:
            return

        with tempfile.NamedTemporaryFile(delete=False, suffix=output_path.suffix) as tmp:
            temp_path = Path(tmp.name)

        try:
            with ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as new_zip:
                for info in output_zip.infolist():
                    data = replacements.get(info.filename)
                    if data is None:
                        data = output_zip.read(info.filename)
                    new_zip.writestr(info, data)
            shutil.move(str(temp_path), output_path)
        finally:
            if temp_path.exists():
                temp_path.unlink()


def insert_after_xml_tag(xml: bytes, tag: bytes, payload: bytes) -> bytes:
    index = xml.find(tag)
    if index == -1:
        return xml.replace(b"</worksheet>", payload + b"</worksheet>", 1)
    insert_at = index + len(tag)
    return xml[:insert_at] + payload + xml[insert_at:]


def copy_worksheet_root_namespaces(source_xml: bytes, output_xml: bytes) -> bytes:
    source_root = WORKSHEET_ROOT_RE.search(source_xml)
    output_root = WORKSHEET_ROOT_RE.search(output_xml)
    if not source_root or not output_root:
        return output_xml

    source_tag = source_root.group(0)
    output_tag = output_root.group(0)
    updated_tag = output_tag

    existing_namespace_names = {
        declaration.split(b"=", 1)[0].strip()
        for declaration in XMLNS_ATTR_RE.findall(output_tag)
    }
    for declaration in XMLNS_ATTR_RE.findall(source_tag):
        namespace_name = declaration.split(b"=", 1)[0].strip()
        if namespace_name not in existing_namespace_names:
            updated_tag = insert_before_tag_end(updated_tag, declaration)
            existing_namespace_names.add(namespace_name)

    source_ignorable = MC_IGNORABLE_RE.search(source_tag)
    output_ignorable = MC_IGNORABLE_RE.search(updated_tag)
    if source_ignorable and not output_ignorable:
        updated_tag = insert_before_tag_end(updated_tag, source_ignorable.group(0))
    elif source_ignorable and output_ignorable:
        values = dict.fromkeys(
            output_ignorable.group(1).split() + source_ignorable.group(1).split()
        )
        merged = b' mc:Ignorable="' + b" ".join(values) + b'"'
        updated_tag = (
            updated_tag[: output_ignorable.start()]
            + merged
            + updated_tag[output_ignorable.end() :]
        )

    if updated_tag == output_tag:
        return output_xml
    return output_xml[: output_root.start()] + updated_tag + output_xml[output_root.end() :]


def insert_before_tag_end(tag: bytes, attribute: bytes) -> bytes:
    end = tag.rfind(b">")
    if end == -1:
        return tag
    return tag[:end] + attribute + tag[end:]


def find_article_index(headers: list[str], template: PartnerTemplate | None = None) -> int | None:
    return partner_template(template).find_article_index(headers)


def class81_index(headers: list[str], template: PartnerTemplate | None = None) -> int | None:
    return partner_template(template).class_index(headers)


def product_with_template_class(
    product: ProductRecord, template_class: str, template: PartnerTemplate | None = None
) -> ProductRecord:
    return partner_template(template).with_template_class(product, template_class)


def cell_text(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").split())


def has_vba_project(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except OSError:
        return False


def make_output_paths(template_path: Path, output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = output_dir / f"{template_path.stem}_filled_{stamp}{template_path.suffix}"
    report_path = output_dir / f"{template_path.stem}_report_{stamp}.xlsx"
    return output_path, report_path


def value_matches_to_review(rows: list[FillReportRow]) -> list[list[object]]:
    """One line per distinct (field, value in DB, suggestion), most frequent first.

    The user confirms or corrects "Принять значение" once per line instead of
    once per product; importing the report back remembers those choices.
    """
    counts: dict[tuple[str, str, str], int] = {}
    candidates: dict[tuple[str, str, str], tuple[str, ...]] = {}
    for item in rows:
        if not item.raw or item.status not in {"filled_suggested", "value_not_allowed"}:
            continue
        suggestion = item.value if item.status == "filled_suggested" else ""
        key = (item.column, item.raw, suggestion)
        counts[key] = counts.get(key, 0) + 1
        candidates.setdefault(key, item.candidates)
    ordered = sorted(counts, key=lambda key: (-counts[key], key))
    return [
        [field, raw, suggestion, counts[(field, raw, suggestion)], "; ".join(candidates[(field, raw, suggestion)]), suggestion]
        for field, raw, suggestion in ordered
    ]


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
            "Состояние",
            "Заполнено",
            "Заполнено после проверки",
            "Не заполнено",
            "Что проверить",
        ]
    )
    details.append(["Строка", "Артикул", "Поле", "Состояние", "Значение", "Откуда взяли", "Комментарий"])

    by_row: dict[tuple[int, str], dict[str, object]] = {}
    for item in rows:
        key = (item.row_number, item.article)
        bucket = by_row.setdefault(
            key,
            {
                "filled": 0,
                "suggested": 0,
                "missing": 0,
                "status": "ok",
                "notes": [],
            },
        )
        if item.status == "filled":
            bucket["filled"] = int(bucket["filled"]) + 1
        elif item.status == "filled_suggested":
            bucket["suggested"] = int(bucket["suggested"]) + 1
            bucket["status"] = "needs_review"
        elif item.status in {"missing_value", "value_not_allowed"}:
            bucket["missing"] = int(bucket["missing"]) + 1
            bucket["status"] = "incomplete"
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
                report_status_label(item.status),
                item.value,
                source_label(item.source),
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
                report_status_label(str(bucket["status"])),
                bucket["filled"],
                bucket["suggested"],
                bucket["missing"],
                "; ".join(notes),
            ]
        )

    sheets = [summary, details]
    matches = value_matches_to_review(rows)
    if matches:
        review = workbook.create_sheet(MATCHES_SHEET)
        review.append(MATCHES_HEADERS)
        for item in matches:
            review.append(item)
        sheets.append(review)

    for sheet in sheets:
        sheet.freeze_panes = "A2"
        for column_cells in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 70)

    workbook.save(path)


def fill_template(
    db_path: Path,
    template_path: Path,
    output_dir: Path,
    rules_path: Path | None = None,
    manual_values_path: Path | None = None,
    partner: str = DEFAULT_PARTNER,
) -> FillResult:
    template = load_partner_template(partner)
    if template.code == "etm" and template_path.suffix.lower() == ".csv":
        from .etm_characteristics import fill_characteristics_report, is_characteristics_report

        if is_characteristics_report(template_path):
            return fill_characteristics_report(
                db_path, template_path, output_dir, rules_path, manual_values_path, template
            )
    if template.writer.get("kind") != "fill_template":
        raise ValueError(f"Шаблон партнёра {template.name} не заполняется поверх файла партнёра.")
    suffix = template_path.suffix.lower()
    formats = {f".{name}" for name in template.writer.get("formats", [])}
    if suffix not in formats:
        raise ValueError(f"Для {template.name} поддерживаются файлы: {', '.join(sorted(formats))}.")
    if suffix == ".csv":
        return fill_csv_template(
            db_path, template_path, output_dir, rules_path, manual_values_path, template
        )
    if suffix in {".xlsx", ".xlsm"}:
        return fill_xlsx_template(
            db_path, template_path, output_dir, rules_path, manual_values_path, template
        )
    raise ValueError("Поддерживаются только CSV, XLSX и XLSM шаблоны.")


def fill_csv_template(
    db_path: Path,
    template_path: Path,
    output_dir: Path,
    rules_path: Path | None = None,
    manual_values_path: Path | None = None,
    template: PartnerTemplate | None = None,
) -> FillResult:
    from .manual_values import load_manual_values

    template = partner_template(template)

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
    article_idx = template.find_article_index(headers)
    if article_idx is None:
        raise ValueError("Не найдена колонка Артикул или Расширенный артикул.")
    class_idx = template.class_index(headers)

    articles = [normalize_article(row[article_idx]) for row in reader[1:] if len(row) > article_idx]
    products = fetch_products(db_path, articles)
    manual_values = load_manual_values(manual_values_path)
    report: list[FillReportRow] = []
    filled_cells = 0
    suggested_cells = 0
    missing_cells = 0

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
        if class_idx is not None and class_idx < len(row):
            product = template.with_template_class(product, normalize_header(row[class_idx]))
        for col_idx, header in enumerate(headers):
            if header in template.skip_headers:
                continue
            existing_value = cell_text(row[col_idx])
            if existing_value:
                if template.is_reportable(header):
                    report.append(
                        FillReportRow(
                            row_number,
                            article,
                            header,
                            "existing_value_preserved",
                            existing_value,
                            "template_existing",
                            "Ячейка уже была заполнена; режим EMPTY_ONLY не перезаписывает ее.",
                        )
                    )
                continue
            manual_value = manual_values.get((article, header), "")
            if manual_value:
                value, status, source = manual_value, "filled", "manual_completion"
            else:
                value, status, source = template.product_value(product, header, rules_path)
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
                        raw=template.raw_value(product, header, rules_path),
                    )
                )
            elif status == "value_not_allowed":
                missing_cells += 1
                report.append(
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
            elif status in {"blank", "filled"} and template.is_reportable(header):
                missing_cells += 1
                report.append(
                    FillReportRow(
                        row_number,
                        article,
                        header,
                        "missing_value",
                        "",
                        source,
                        template.missing_note(header),
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
        missing_cells,
        len(unique_articles) - found,
    )


def fill_xlsx_template(
    db_path: Path,
    template_path: Path,
    output_dir: Path,
    rules_path: Path | None = None,
    manual_values_path: Path | None = None,
    template: PartnerTemplate | None = None,
) -> FillResult:
    from .manual_values import load_manual_values

    template = partner_template(template)

    keep_vba = template_path.suffix.lower() == ".xlsm"
    input_has_vba = has_vba_project(template_path) if keep_vba else False
    workbook = load_workbook(template_path, keep_vba=keep_vba)
    sheet = workbook.active
    header_cells = list(sheet[1])
    headers = [normalize_header(cell.value) for cell in header_cells]
    fillable_columns = {
        idx
        for idx, cell in enumerate(header_cells)
        if is_yellow_header_cell(cell) or not template.fills_only_yellow_headers
    }
    article_idx = template.find_article_index(headers)
    if article_idx is None:
        raise ValueError("Не найдена колонка Артикул или Расширенный артикул.")
    class_idx = template.class_index(headers)

    articles: list[str] = []
    for row in sheet.iter_rows(min_row=2):
        if article_idx < len(row):
            articles.append(normalize_article(row[article_idx].value))

    products = fetch_products(db_path, articles)
    manual_values = load_manual_values(manual_values_path)
    report: list[FillReportRow] = []
    filled_cells = 0
    suggested_cells = 0
    missing_cells = 0

    for row_number, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        article = normalize_article(row[article_idx].value if article_idx < len(row) else "")
        if not article:
            continue
        product = products.get(article)
        if not product:
            report.append(FillReportRow(row_number, article, "", "not_found", "", "", "Артикул не найден в базе"))
            for col_idx in fillable_columns:
                if col_idx != article_idx and col_idx < len(row):
                    mark_missing_cell(row[col_idx])
                    if template.is_reportable(headers[col_idx]):
                        missing_cells += 1
            continue
        if class_idx is not None and class_idx < len(row):
            product = template.with_template_class(
                product, normalize_header(row[class_idx].value)
            )
        for col_idx, header in enumerate(headers):
            if header in template.skip_headers or col_idx >= len(row):
                continue
            if col_idx not in fillable_columns:
                continue
            existing_value = cell_text(row[col_idx].value)
            if existing_value:
                if template.is_reportable(header):
                    report.append(
                        FillReportRow(
                            row_number,
                            article,
                            header,
                            "existing_value_preserved",
                            existing_value,
                            "template_existing",
                            "Ячейка уже была заполнена; режим EMPTY_ONLY не перезаписывает ее.",
                        )
                    )
                continue
            manual_value = manual_values.get((article, header), "")
            if manual_value:
                value, status, source = manual_value, "filled", "manual_completion"
            else:
                value, status, source = template.product_value(product, header, rules_path)
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
                        raw=template.raw_value(product, header, rules_path),
                    )
                )
            elif status == "value_not_allowed":
                missing_cells += 1
                mark_missing_cell(row[col_idx])
                report.append(
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
            elif status in {"blank", "filled"} and template.is_reportable(header):
                missing_cells += 1
                mark_missing_cell(row[col_idx])
                report.append(
                    FillReportRow(
                        row_number,
                        article,
                        header,
                        "missing_value",
                        "",
                        source,
                        template.missing_note(header),
                    )
                )

    output_path, report_path = make_output_paths(template_path, output_dir)
    workbook.save(output_path)
    preserve_worksheet_dropdowns(template_path, output_path)
    if input_has_vba and not has_vba_project(output_path):
        raise ValueError("После сохранения XLSM потерян xl/vbaProject.bin.")
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
        missing_cells=missing_cells,
        missing_articles=len(unique_articles) - found,
    )
