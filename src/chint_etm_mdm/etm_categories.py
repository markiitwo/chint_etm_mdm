"""New product cards for ETM through category templates.

ETM's OneTeam cabinet gives one upload template per product class ("81 класс"):
the usual card columns plus the class's required characteristics ("Конфиг:..."),
most of them with a drop-down list on the "Списки" sheet. The user keeps the
downloaded templates in one folder. For new products the program picks each
product's class, fills a copy of that class's template and writes the CSV that
ETM takes. Classes without a template are listed so the user knows what to
download.
"""
from __future__ import annotations

import csv
import datetime as dt
import re
import sqlite3
from collections import Counter
from contextlib import closing
from copy import copy
from dataclasses import dataclass, field, replace
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .db import ProductRecord, connect, fetch_products
from .export_templates import PartnerTemplate, clean_product_name, load_partner_template
from .filler import FillResult, fill_xlsx_template, normalize_article, normalize_header
from .value_matching import canonical

TEMPLATE_SHEET = "Шаблон"
LISTS_SHEET = "Списки"
CLASS_HEADER = "81 класс"
CONFIG_PREFIX = "Конфиг:"
FILE_CLASS_RE = re.compile(r"upload_goods_category_template_([0-9A-Za-z]+)", re.IGNORECASE)
CODE_HEADERS = {"part number": "article", "tariff code": "Код ТН ВЭД", "окпд 2 (9 зн.)": "Код ОКПД2"}
CSV_ENCODING = "cp1251"


@dataclass(frozen=True)
class CategoryTemplate:
    class_code: str
    path: Path
    dropdowns: dict[str, tuple[str, ...]]


@dataclass(frozen=True)
class ClassChoice:
    code: str
    basis: str
    sure: bool


@dataclass
class CategoryRun:
    output_dir: Path
    summary_path: Path
    classes: dict[str, ClassChoice] = field(default_factory=dict)
    results: dict[str, tuple[FillResult, Path]] = field(default_factory=dict)
    missing_templates: dict[str, list[str]] = field(default_factory=dict)
    unclassified: list[str] = field(default_factory=list)
    not_found: list[str] = field(default_factory=list)


def cell_text(value: object) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return normalize_header(value)


def read_category_template(path: Path) -> CategoryTemplate:
    workbook = load_workbook(path)
    if TEMPLATE_SHEET not in workbook.sheetnames:
        raise ValueError(f"В файле {path.name} нет листа «{TEMPLATE_SHEET}».")
    sheet = workbook[TEMPLATE_SHEET]
    headers = {cell.column_letter: normalize_header(cell.value) for cell in sheet[1]}
    class_letter = next((letter for letter, name in headers.items() if name == CLASS_HEADER), None)
    class_code = cell_text(sheet[f"{class_letter}2"].value) if class_letter else ""
    if not class_code:
        found = FILE_CLASS_RE.search(path.stem)
        class_code = found.group(1) if found else ""
    if not class_code:
        raise ValueError(f"В шаблоне {path.name} не указан 81 класс.")
    lists = workbook[LISTS_SHEET] if LISTS_SHEET in workbook.sheetnames else None
    dropdowns: dict[str, tuple[str, ...]] = {}
    for validation in sheet.data_validations.dataValidation:
        if validation.type != "list" or not validation.formula1 or lists is None:
            continue
        sheet_name, _, area = validation.formula1.partition("!")
        if sheet_name.strip("'") != LISTS_SHEET:
            continue
        values = tuple(
            text for row in lists[area.replace("$", "")] for cell in row if (text := cell_text(cell.value))
        )
        for cell_range in validation.sqref.ranges:
            header = headers.get(re.sub(r"\d", "", str(cell_range).split(":")[0]), "")
            if header.startswith(CONFIG_PREFIX) and values:
                dropdowns[canonical(header.removeprefix(CONFIG_PREFIX))] = values
    return CategoryTemplate(class_code, path, dropdowns)


def category_library(folder: Path) -> dict[str, Path]:
    """Class code -> template file, for every category template in the folder."""
    library: dict[str, Path] = {}
    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        found = FILE_CLASS_RE.search(path.stem)
        if found:
            library[found.group(1)] = path
            continue
        try:
            library[read_category_template(path).class_code] = path
        except (ValueError, KeyError, OSError):
            continue
    return library


def new_articles_since(db_path: Path, since: dt.date) -> list[str]:
    """Articles whose price-list introduction date ("Введено в ассортимент") is after `since`."""
    with closing(connect(db_path)) as conn:
        rows = conn.execute("SELECT article, introduced_at FROM new_products").fetchall()
    result = []
    for article, introduced in rows:
        try:
            date = dt.datetime.strptime(str(introduced or "").strip()[:10], "%d.%m.%Y").date()
        except ValueError:
            continue
        if date > since:
            result.append(str(article))
    return sorted(result)


def series_family(series: str) -> str:
    return re.split(r"[-\s/(]", series.strip(), maxsplit=1)[0].upper()


def suggest_classes(
    db_path: Path, products: dict[str, ProductRecord], with_template: frozenset[str] = frozenset()
) -> dict[str, ClassChoice]:
    """The class each product already has, else the usual class of its series in ETM's catalog.

    When the series is split between classes, a class whose template the user has wins over a more
    frequent one without a template; such a choice is marked for a check.
    """
    choices = {
        article: ClassChoice(product.class81_code, "уже есть в базе", True)
        for article, product in products.items()
        if product.class81_code
    }
    pending = [article for article in products if article not in choices]
    if not pending:
        return choices
    try:
        with closing(connect(db_path)) as conn:
            known = conn.execute(
                """
                SELECT DISTINCT p.article, COALESCE(p.series_name, ''), COALESCE(p.equipment_group, ''), ig.class81_code
                FROM products p JOIN ipro_goods ig ON ig.article = p.article
                WHERE COALESCE(ig.class81_code, '') <> ''
                """
            ).fetchall()
            placeholders = ",".join("?" for _ in pending)
            own = conn.execute(
                f"SELECT article, COALESCE(series_name, ''), COALESCE(equipment_group, '') FROM products WHERE article IN ({placeholders})",
                pending,
            ).fetchall()
    except sqlite3.OperationalError:
        return choices
    by_series: dict[str, Counter[str]] = {}
    by_family: dict[str, Counter[str]] = {}
    by_group: dict[str, Counter[str]] = {}
    for _, series, group, code in known:
        code = cell_text(code)
        if series:
            by_series.setdefault(series.strip().upper(), Counter())[code] += 1
            by_family.setdefault(series_family(series), Counter())[code] += 1
        if group:
            by_group.setdefault(group.strip(), Counter())[code] += 1
    for article, series, group in own:
        for counts, basis, sure in (
            (by_series.get(series.strip().upper()), f"как у серии {series}", True),
            (by_family.get(series_family(series)) if series else None, f"как у серий {series_family(series)}", True),
            (by_group.get(group.strip()) if group else None, f"как у группы «{group}»", False),
        ):
            if counts:
                total = sum(counts.values())
                code, votes = counts.most_common(1)[0]
                if code not in with_template:
                    usable = [(item, count) for item, count in counts.most_common() if item in with_template]
                    if usable and usable[0][1] / total >= 0.2:
                        code, votes = usable[0]
                        sure = False
                choices[str(article)] = ClassChoice(code, f"{basis} ({votes} из {total})", sure and votes / total >= 0.8)
                break
    return choices


def load_codes_file(path: Path) -> dict[tuple[str, str], str]:
    """ТН ВЭД and ОКПД2 by article from the "Честный знак" product list (sheet "Price-list")."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook["Price-list"] if "Price-list" in workbook.sheetnames else workbook.worksheets[0]
        rows = sheet.iter_rows(values_only=True)
        columns: dict[str, int] = {}
        for row in rows:
            names = [" ".join(str(cell or "").split()).casefold() for cell in row]
            if "part number" in names:
                columns = {CODE_HEADERS[name]: idx for idx, name in enumerate(names) if name in CODE_HEADERS}
                break
        if "article" not in columns:
            raise ValueError("В файле кодов нет колонки «Part number».")
        values: dict[tuple[str, str], str] = {}
        for row in rows:
            article = normalize_article(row[columns["article"]] if columns["article"] < len(row) else "")
            if not article:
                continue
            for header, idx in columns.items():
                value = cell_text(row[idx]) if header != "article" and idx < len(row) else ""
                if value and value not in {"0", "-"}:
                    values[(article, header)] = value
        return values
    finally:
        workbook.close()


def prepare_input(category: CategoryTemplate, articles: list[str], target: Path) -> None:
    """A copy of the category template with one row per article (style of ETM's first row)."""
    workbook = load_workbook(category.path)
    sheet = workbook[TEMPLATE_SHEET]
    workbook.active = workbook.sheetnames.index(TEMPLATE_SHEET)
    headers = [normalize_header(cell.value) for cell in sheet[1]]
    columns = {name: idx + 1 for idx, name in enumerate(headers)}
    class_value = sheet.cell(2, columns[CLASS_HEADER]).value if CLASS_HEADER in columns else category.class_code
    for row in range(2, sheet.max_row + 1):
        for col in range(1, len(headers) + 1):
            sheet.cell(row, col).value = None
    for offset, article in enumerate(articles):
        row = offset + 2
        for col in range(1, len(headers) + 1):
            cell = sheet.cell(row, col)
            if row > 6:
                cell._style = copy(sheet.cell(2, col)._style)
        for header in ("Артикул", "Расширенный артикул"):
            if header in columns:
                sheet.cell(row, columns[header]).value = article
        if CLASS_HEADER in columns:
            sheet.cell(row, columns[CLASS_HEADER]).value = class_value
    workbook.save(target)


def csv_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "ДА" if value else ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else f"{value:.9f}".rstrip("0").rstrip(".")
    return str(value)


def export_upload_csv(xlsx_path: Path, csv_path: Path) -> int:
    """The "Шаблон" sheet as ETM takes it: Win-1251, ";", a dot in numbers; only rows with an article."""
    workbook = load_workbook(xlsx_path, data_only=True)
    sheet = workbook[TEMPLATE_SHEET]
    rows = list(sheet.iter_rows(values_only=True))
    headers = [normalize_header(value) for value in rows[0]]
    width = max(idx for idx, name in enumerate(headers) if name) + 1
    article_idx = headers.index("Артикул")
    written = 0
    with csv_path.open("w", encoding=CSV_ENCODING, errors="replace", newline="") as handle:
        writer = csv.writer(handle, delimiter=";", lineterminator="\r\n")
        writer.writerow(headers[:width])
        for row in rows[1:]:
            if normalize_article(row[article_idx]):
                writer.writerow([csv_text(value) for value in row[:width]])
                written += 1
    return written


def write_summary(run: CategoryRun, products: dict[str, ProductRecord], library: dict[str, Path]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Новинки"
    sheet.append(["Артикул", "Название", "81 класс", "Как выбран класс", "Проверить класс", "Шаблон", "Файл для загрузки"])
    for article in sorted(products):
        product = products[article]
        choice = run.classes.get(article)
        code = choice.code if choice else ""
        made = run.results.get(code)
        sheet.append(
            [
                article,
                clean_product_name(product.full_name or product.name),
                code,
                choice.basis if choice else "класс не найден, укажите вручную",
                "" if choice and choice.sure else "да",
                library[code].name if code in library else ("нет, скачайте из личного кабинета ЭТМ" if code else ""),
                made[1].name if made else "",
            ]
        )
    for article in run.not_found:
        sheet.append([article, "", "", "артикула нет в базе", "", "", ""])
    widths = {"A": 14, "B": 70, "C": 12, "D": 40, "E": 16, "F": 48, "G": 48}
    for letter, width in widths.items():
        sheet.column_dimensions[letter].width = width
    workbook.save(run.summary_path)


def fill_new_products(
    db_path: Path,
    articles: list[str],
    templates_dir: Path,
    output_dir: Path,
    rules_path: Path | None = None,
    manual_values_path: Path | None = None,
    codes_path: Path | None = None,
    template: PartnerTemplate | None = None,
) -> CategoryRun:
    template = template or load_partner_template("etm")
    products = fetch_products(db_path, articles)
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    run = CategoryRun(output_dir=output_dir, summary_path=output_dir / f"новинки_сводка_{stamp}.xlsx")
    run.not_found = sorted(set(articles) - set(products))
    library = category_library(templates_dir)
    run.classes = suggest_classes(db_path, products, frozenset(library))
    run.unclassified = sorted(article for article in products if article not in run.classes)
    fallback = load_codes_file(codes_path) if codes_path else {}

    by_class: dict[str, list[str]] = {}
    for article, choice in run.classes.items():
        by_class.setdefault(choice.code, []).append(article)
    for code, members in sorted(by_class.items()):
        members.sort()
        if code not in library:
            run.missing_templates[code] = members
            continue
        category = read_category_template(library[code])
        input_path = output_dir / f"новинки_{code}.xlsx"
        prepare_input(category, members, input_path)
        result = fill_xlsx_template(
            db_path,
            input_path,
            output_dir,
            rules_path,
            manual_values_path,
            template=replace(template, dropdowns=category.dropdowns),
            fallback_values=fallback,
        )
        csv_path = result.output_path.with_suffix(".csv")
        export_upload_csv(result.output_path, csv_path)
        input_path.unlink()
        run.results[code] = (result, csv_path)
    write_summary(run, products, library)
    return run
