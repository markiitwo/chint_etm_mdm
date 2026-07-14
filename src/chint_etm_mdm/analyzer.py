from __future__ import annotations

import datetime as dt
import json
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .db import ProductRecord, fetch_products
from .filler import (
    comparable_text,
    find_article_index,
    is_reportable_header,
    is_yellow_header_cell,
    normalize_article,
    normalize_header,
    product_value,
)
from .mapping_rules import rejected_sources_for, rules_for, source_attributes_for


@dataclass(frozen=True)
class CandidateDetail:
    source: str
    count: int
    examples: tuple[str, ...]
    articles: tuple[str, ...]


@dataclass(frozen=True)
class FieldCoverage:
    class81_code: str
    field: str
    products_count: int
    exact_count: int
    approved_rule_count: int
    candidate_count: int
    filled_direct_count: int
    missing_count: int
    approved_sources: tuple[str, ...]
    candidate_sources: tuple[str, ...]
    candidate_details: tuple[CandidateDetail, ...]
    note: str


def analyze_template_mapping(
    db_path: Path, template_path: Path, output_dir: Path, rules_path: Path | None = None
) -> Path:
    suffix = template_path.suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError("Анализ маппинга сейчас поддерживает только XLSX/XLSM, потому что нужен цвет заголовков.")

    workbook = load_workbook(template_path, read_only=False, data_only=True)
    sheet = workbook.active
    header_cells = list(sheet[1])
    headers = [normalize_header(cell.value) for cell in header_cells]
    fillable_columns = {
        idx
        for idx, cell in enumerate(header_cells)
        if is_yellow_header_cell(cell)
    }
    article_idx = find_article_index(headers)
    if article_idx is None:
        raise ValueError("Не найдена колонка Артикул или Расширенный артикул.")

    articles: list[str] = []
    for row in sheet.iter_rows(min_row=2):
        if article_idx < len(row):
            articles.append(normalize_article(row[article_idx].value))

    products = fetch_products(db_path, articles)
    by_class: dict[str, list[ProductRecord]] = defaultdict(list)
    for article in articles:
        if not article:
            continue
        product = products.get(article)
        if product:
            by_class[product.class81_code or ""].append(product)

    coverages: list[FieldCoverage] = []
    examples: list[list[str | int]] = []

    yellow_headers = [
        header
        for idx, header in enumerate(headers)
        if idx in fillable_columns and header and header not in {"Артикул", "Расширенный артикул"}
    ]

    for class81_code, class_products in sorted(by_class.items()):
        for header in yellow_headers:
            if header.startswith("Конфиг:"):
                coverage, field_examples = analyze_config_field(
                    class81_code, header, class_products, rules_path
                )
            elif is_reportable_header(header):
                coverage, field_examples = analyze_direct_field(
                    class81_code, header, class_products, rules_path
                )
            else:
                coverage = FieldCoverage(
                    class81_code=class81_code,
                    field=header,
                    products_count=len(class_products),
                    exact_count=0,
                    approved_rule_count=0,
                    candidate_count=0,
                    filled_direct_count=0,
                    missing_count=len(class_products),
                    approved_sources=(),
                    candidate_sources=(),
                    candidate_details=(),
                    note="Поле пока не поддерживается логикой заполнения",
                )
                field_examples = []
            coverages.append(coverage)
            examples.extend(field_examples)

    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_path = output_dir / f"{template_path.stem}_mapping_review_{stamp}.xlsx"
    write_mapping_report(report_path, coverages, examples)
    return report_path


def analyze_config_field(
    class81_code: str,
    field: str,
    products: list[ProductRecord],
    rules_path: Path | None = None,
) -> tuple[FieldCoverage, list[list[str | int]]]:
    attr_name = field.removeprefix("Конфиг:").strip()
    approved_sources = source_attributes_for(class81_code, field, rules_path)
    approved_count = 0
    exact_count = 0
    candidate_count = 0
    missing_count = 0
    candidate_counter: Counter[str] = Counter()
    candidate_values: dict[str, list[str]] = defaultdict(list)
    candidate_articles: dict[str, list[str]] = defaultdict(list)
    example_rows: list[list[str | int]] = []

    wanted = comparable_text(attr_name)
    rejected_sources = set(rejected_sources_for(class81_code, field, rules_path))
    for product in products:
        attrs = product.attributes or {}
        if attrs.get(attr_name):
            exact_count += 1
            example_rows.append([class81_code, field, attr_name, product.article, attrs[attr_name], "exact"])
            continue

        approved_value = ""
        approved_source = ""
        for source in approved_sources:
            if attrs.get(source):
                approved_value = attrs[source]
                approved_source = source
                break
        if approved_value:
            approved_count += 1
            example_rows.append(
                [class81_code, field, approved_source, product.article, approved_value, "approved_class_rule"]
            )
            continue

        candidates = [
            candidate
            for candidate in find_attribute_candidates(attr_name, attrs)
            if candidate[0] not in rejected_sources
        ]
        if candidates:
            candidate_count += 1
            source_name, value = candidates[0]
            candidate_counter[source_name] += 1
            if value not in candidate_values[source_name]:
                candidate_values[source_name].append(value)
            if product.article not in candidate_articles[source_name]:
                candidate_articles[source_name].append(product.article)
            example_rows.append([class81_code, field, source_name, product.article, value, "candidate"])
            continue

        missing_count += 1

    known_rules = rules_for(class81_code, field, rules_path)
    note = "Есть утвержденное правило для класса" if known_rules else "Требуется подтверждение правила для класса"
    if not wanted:
        note = "Пустое имя Конфиг-поля"

    candidate_details = tuple(
        CandidateDetail(
            source=name,
            count=count,
            examples=tuple(candidate_values[name][:5]),
            articles=tuple(candidate_articles[name][:5]),
        )
        for name, count in candidate_counter.most_common(8)
    )

    return (
        FieldCoverage(
            class81_code=class81_code,
            field=field,
            products_count=len(products),
            exact_count=exact_count,
            approved_rule_count=approved_count,
            candidate_count=candidate_count,
            filled_direct_count=0,
            missing_count=missing_count,
            approved_sources=approved_sources,
            candidate_sources=tuple(item.source for item in candidate_details),
            candidate_details=candidate_details,
            note=note,
        ),
        example_rows[:50],
    )


def analyze_direct_field(
    class81_code: str,
    field: str,
    products: list[ProductRecord],
    rules_path: Path | None = None,
) -> tuple[FieldCoverage, list[list[str | int]]]:
    filled_count = 0
    missing_count = 0
    sources: Counter[str] = Counter()
    example_rows: list[list[str | int]] = []

    for product in products:
        value, status, source = product_value(product, field, rules_path)
        if status == "filled" and value:
            filled_count += 1
            sources[source] += 1
            example_rows.append([class81_code, field, source, product.article, value, "filled"])
        else:
            missing_count += 1

    return (
        FieldCoverage(
            class81_code=class81_code,
            field=field,
            products_count=len(products),
            exact_count=0,
            approved_rule_count=0,
            candidate_count=0,
            filled_direct_count=filled_count,
            missing_count=missing_count,
            approved_sources=tuple(name for name, _ in sources.most_common(5)),
            candidate_sources=(),
            candidate_details=(),
            note="Прямое поле базы, не ETIM-маппинг",
        ),
        example_rows[:50],
    )


def find_attribute_candidates(attr_name: str, attrs: dict[str, str]) -> list[tuple[str, str]]:
    wanted = comparable_text(attr_name)
    if not wanted:
        return []

    candidates: list[tuple[int, str, str]] = []
    for source_name, value in attrs.items():
        source_key = comparable_text(source_name)
        if not source_key or not value:
            continue
        if wanted == source_key:
            score = 100
        elif wanted in source_key or source_key in wanted:
            score = 70
        else:
            wanted_tokens = set(attr_name.lower().replace(",", " ").split())
            source_tokens = set(source_name.lower().replace(",", " ").split())
            score = len(wanted_tokens.intersection(source_tokens)) * 10
        if score > 0:
            candidates.append((score, source_name, value))

    candidates.sort(key=lambda item: (-item[0], item[1]))
    return [(source_name, value) for _, source_name, value in candidates[:5]]


def write_mapping_report(
    path: Path, coverages: list[FieldCoverage], examples: list[list[str | int]]
) -> None:
    workbook = Workbook()
    coverage_sheet = workbook.active
    coverage_sheet.title = "Покрытие"
    coverage_sheet.append(
        [
            "81 класс",
            "Поле шаблона",
            "Товаров",
            "Прямо заполнится",
            "Точное имя атрибута",
            "Утвержденное правило",
            "Есть кандидаты",
            "Не найдено",
            "Утвержденные источники",
            "Кандидаты",
            "Комментарий",
        ]
    )
    for item in coverages:
        coverage_sheet.append(
            [
                item.class81_code,
                item.field,
                item.products_count,
                item.filled_direct_count,
                item.exact_count,
                item.approved_rule_count,
                item.candidate_count,
                item.missing_count,
                "; ".join(item.approved_sources),
                "; ".join(item.candidate_sources),
                item.note,
            ]
        )

    examples_sheet = workbook.create_sheet("Примеры")
    examples_sheet.append(["81 класс", "Поле шаблона", "Источник/кандидат", "Артикул", "Значение", "Статус"])
    for row in examples:
        examples_sheet.append(row)

    rules_sheet = workbook.create_sheet("Правила")
    rules_sheet.append(
        [
            "81 класс",
            "Поле шаблона",
            "Кандидат источника",
            "Покрытие кандидата",
            "Примеры значений",
            "Примеры артикулов",
            "Действие",
            "JSON для rules",
        ]
    )
    for item in coverages:
        if not item.field.startswith("Конфиг:"):
            continue
        if item.approved_sources:
            rules_sheet.append(
                [
                    item.class81_code,
                    item.field,
                    "; ".join(item.approved_sources),
                    item.approved_rule_count,
                    "",
                    "",
                    "уже утверждено",
                    "",
                ]
            )
            continue
        for candidate in item.candidate_details:
            snippet = {
                "class81_code": item.class81_code,
                "template_field": item.field,
                "source_attributes": [candidate.source],
                "confidence": "approved_class_rule",
                "note": "Confirm manually before use.",
            }
            rules_sheet.append(
                [
                    item.class81_code,
                    item.field,
                    candidate.source,
                    candidate.count,
                    "; ".join(candidate.examples),
                    "; ".join(candidate.articles),
                    "проверить и при подтверждении добавить в rules",
                    json.dumps(snippet, ensure_ascii=False),
                ]
            )

    for sheet in (coverage_sheet, examples_sheet, rules_sheet):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 80)

    workbook.save(path)
