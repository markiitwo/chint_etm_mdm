"""Partner export templates.

The product card (ProductRecord) knows nothing about partners. Each partner is
described by a JSON file in the ``partners`` package: which column gets which
product field, how the value is formatted, and what to say when it is missing.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, replace
from functools import lru_cache
from importlib import resources
from pathlib import Path
from typing import Any

from . import partners
from .db import ProductRecord
from .mapping_rules import source_attributes_for
from .value_matching import (
    Characteristic,
    MatchResult,
    ValueLists,
    canonical,
    current_kind_from_text,
    current_kind_from_value,
    format_number,
    match_value,
    only_number,
    product_type_from_name,
    same_name_attributes,
)

DEFAULT_PARTNER = "etm"
NAME_ALIAS_SOURCE_PREFIX = "name_alias:"
SAME_NAME_SOURCE_PREFIX = "same_name:"
TEMPLATE_ALIAS_SOURCE_PREFIX = "template_alias:"
CHECK_ALIAS_SOURCE_PREFIX = "template_alias_check:"
PRODUCT_TYPE_KEY = "тип изделия"
PRODUCT_TYPE_SOURCE = "product_name_type"
CURRENT_KIND_ATTRIBUTES = ("Род тока", "Род тока включения", "Род тока катушки управления", "Тип напряжения (тока)")


PRODUCT_MARK_RE = re.compile(r"\s*\(R\)")


def clean_product_name(value: str) -> str:
    """Drop the price list's "(R)" marker, which is not part of the product name."""
    text = PRODUCT_MARK_RE.sub("", value or "")
    return " ".join(text.split()).strip(" ,")


def short_wms_name(value: str, limit: int = 60) -> str:
    clean = clean_product_name(value)
    return clean[:limit].rstrip()


def format_decimal(value: float | int | None, digits: int = 6) -> str:
    if value is None:
        return ""
    text = f"{float(value):.{digits}f}".rstrip("0").rstrip(".")
    if text in {"0", "-0"} and float(value) != 0 and digits < 6:
        # A light part (0.0004 kg) must not round to zero.
        return format_decimal(value, 6)
    return text


def mm_to_m(value: float | int | None) -> str:
    if value is None:
        return ""
    return format_decimal(float(value) / 1000.0, 6)


def format_count(value: float | int | None) -> str:
    if value is None:
        return ""
    number = float(value)
    if number.is_integer():
        return str(int(number))
    return format_decimal(number, 6)


def apply_transform(value: Any, spec: "FieldSpec") -> str:
    if spec.transform == "":
        return "" if value is None else value
    if spec.transform in NUMBER_TRANSFORMS:
        return NUMBER_TRANSFORMS[spec.transform](value, spec)
    if spec.transform == "name":
        return clean_product_name(value or "")
    if spec.transform == "truncate":
        return short_wms_name(value or "", spec.limit)
    raise ValueError(f"Неизвестное преобразование в шаблоне партнёра: {spec.transform}")


NUMBER_TRANSFORMS = {
    "decimal": lambda value, spec: format_decimal(value, spec.digits),
    "mm_to_m": lambda value, spec: mm_to_m(value),
    "count": lambda value, spec: format_count(value),
}


def range_max(value: str) -> str:
    """"525-690" / "525...690" -> "690"; a single value stays as it is."""
    numbers = re.findall(r"\d+(?:[.,]\d+)?", value)
    if len(numbers) == 2 and re.fullmatch(r"\s*\d+(?:[.,]\d+)?\s*(?:-|–|…|\.\.\.)\s*\d+(?:[.,]\d+)?\s*\D{0,4}", value):
        return numbers[1].replace(",", ".")
    return value


def scaled(value: str, factor: float) -> str:
    """"0.03" with factor 1000 -> "30"; a value that is not one number is dropped."""
    number = only_number(value)
    return "" if number is None else format_number(number * factor)


def is_present(value: Any) -> bool:
    return value is not None and value != ""


@dataclass(frozen=True)
class FieldSpec:
    header: str
    fields: tuple[str, ...] = ()
    const: str | None = None
    when_present: str = ""
    transform: str = ""
    digits: int = 6
    limit: int = 60
    source: str = ""
    note: str = ""

    def raw_value(self, product: ProductRecord) -> Any:
        values = [getattr(product, name) for name in self.fields]
        for value in values:
            if value:
                return value
        return values[-1]

    def value(self, product: ProductRecord) -> tuple[str, str, str]:
        if self.when_present and not is_present(getattr(product, self.when_present)):
            return "", "blank", ""
        if self.const is not None:
            return self.const, "filled", self.source
        if not self.fields:
            return "", "blank", ""
        return apply_transform(self.raw_value(product), self), "filled", self.source


def _file_stamp(path: Path) -> tuple[int, int]:
    try:
        stat = path.stat()
    except FileNotFoundError:
        return (0, 0)
    return (stat.st_mtime_ns, stat.st_size)


_VALUE_LISTS_CACHE: dict[tuple[Path, Path], tuple[tuple[object, ...], ValueLists]] = {}


def cached_value_lists(reference: Path, learned: Path) -> ValueLists:
    key = (reference, learned)
    stamp = (_file_stamp(reference), _file_stamp(learned))
    cached = _VALUE_LISTS_CACHE.get(key)
    if cached is None or cached[0] != stamp:
        cached = (stamp, ValueLists(reference, learned))
        _VALUE_LISTS_CACHE[key] = cached
    return cached[1]


@dataclass(frozen=True)
class AttributeColumns:
    prefix: str
    source: str
    rule_source_prefix: str
    note: str = ""


@dataclass(frozen=True)
class NameAlias:
    """A characteristic of ours that means the partner's one, checked by hand and shipped with the template."""

    attributes: tuple[str, ...]
    values: dict[str, str]
    check: bool = False
    # Only values listed in `values` carry over; any other value of the attribute is skipped.
    only_listed: bool = False
    # Unit conversion, e.g. 1000 for А -> мА.
    scale: float | None = None
    # "max": a range such as "525-690" gives its upper value, for lists of single values.
    range_pick: str = ""
    # ETM classes (81 класс) the pair holds for; empty means every class.
    classes: frozenset[str] = frozenset()


@dataclass(frozen=True)
class PartnerTemplate:
    code: str
    name: str
    writer: dict[str, Any]
    article_headers: tuple[str, ...]
    skip_headers: frozenset[str]
    class_header: str
    class_field: str
    notes: dict[str, str]
    note_prefixes: tuple[tuple[str, str], ...]
    fields: dict[str, FieldSpec]
    attribute_columns: tuple[AttributeColumns, ...]
    value_list_files: tuple[str, str] | None = None
    name_aliases: dict[str, tuple[NameAlias, ...]] | None = None
    # Pairs checked and found wrong: never offered on the "Названия" sheet.
    name_rejects: dict[str, frozenset[str]] | None = None
    # Drop-down lists of an ETM category template, by canonical characteristic name.
    dropdowns: dict[str, tuple[str, ...]] | None = None
    # Checks that hold a filled value back for review (see data_checks).
    checks: frozenset[str] = frozenset()

    def spreadsheet_value(self, header: str, value: str) -> str | int | float:
        """Weights, sizes and counts go into xlsx as real numbers, so no decimal separator can be misread."""
        spec = self.fields.get(header)
        if spec is None or spec.transform not in NUMBER_TRANSFORMS:
            return value
        try:
            number = float(value.replace(",", "."))
        except ValueError:
            return value
        return int(number) if number.is_integer() else number

    def attribute_columns_for(self, header: str) -> AttributeColumns | None:
        for columns in self.attribute_columns:
            if header.startswith(columns.prefix):
                return columns
        return None

    def product_value(
        self, product: ProductRecord, header: str, rules_path: Path | None = None
    ) -> tuple[str, str, str]:
        spec = self.fields.get(header)
        if spec is not None:
            return spec.value(product)
        columns = self.attribute_columns_for(header)
        if columns is not None:
            return self.attribute_value(product, header, columns, rules_path)
        return "", "blank", ""

    def attribute_raw(
        self,
        product: ProductRecord,
        header: str,
        columns: AttributeColumns,
        rules_path: Path | None,
    ) -> tuple[str, str]:
        attrs = product.attributes or {}
        attr_name = header.removeprefix(columns.prefix).strip()
        value = attrs.get(attr_name, "")
        if value:
            return value, columns.source
        class_code = getattr(product, self.class_field)
        for alias in source_attributes_for(class_code, f"{columns.prefix}{attr_name}", rules_path):
            value = attrs.get(alias, "")
            if value:
                return value, f"{columns.rule_source_prefix}{alias}"
        lists = self.value_lists(rules_path)
        for alias in lists.aliases_for(attr_name) if lists else []:
            value = attrs.get(alias, "")
            if value:
                return value, f"{NAME_ALIAS_SOURCE_PREFIX}{alias}"
        for alias in (self.name_aliases or {}).get(canonical(attr_name), ()):
            if alias.classes and class_code not in alias.classes:
                continue
            for attribute in alias.attributes:
                value = attrs.get(attribute, "")
                if alias.only_listed and value not in alias.values:
                    continue
                if value and alias.range_pick == "max":
                    value = range_max(value)
                if value and alias.scale:
                    value = scaled(value, alias.scale)
                if value:
                    prefix = CHECK_ALIAS_SOURCE_PREFIX if alias.check else TEMPLATE_ALIAS_SOURCE_PREFIX
                    return alias.values.get(value, value), f"{prefix}{attribute}"
        if lists:
            for attribute in same_name_attributes(attr_name, attrs):
                return attrs[attribute], f"{SAME_NAME_SOURCE_PREFIX}{attribute}"
        return "", ""

    def attribute_value(
        self,
        product: ProductRecord,
        header: str,
        columns: AttributeColumns,
        rules_path: Path | None,
    ) -> tuple[str, str, str]:
        raw, source = self.attribute_raw(product, header, columns, rules_path)
        attr_name = header.removeprefix(columns.prefix).strip()
        dropdown = self.dropdown_for(attr_name)
        if not raw and dropdown and canonical(attr_name) == PRODUCT_TYPE_KEY:
            raw, source = product_type_from_name(product.full_name or product.name, dropdown), PRODUCT_TYPE_SOURCE
        if not raw:
            return "", "blank", ""
        lists = self.value_lists(rules_path)
        if lists is None and not dropdown:
            return raw, "filled", source
        kind = ""
        if lists is not None and not dropdown and lists.wants_current_kind(attr_name):
            raw, kind, source = self.with_current_kind(product, raw, source)
        match = self.match_allowed(attr_name, raw, rules_path, kind)
        if match.confident and not source.startswith((SAME_NAME_SOURCE_PREFIX, CHECK_ALIAS_SOURCE_PREFIX)):
            return match.value, "filled", source
        if match.confident or match.kind == "words":
            return match.value, "suggested", source
        return raw, "value_not_allowed", source

    def dropdown_for(self, attr_name: str) -> tuple[str, ...]:
        return (self.dropdowns or {}).get(canonical(attr_name), ())

    def match_allowed(self, attr_name: str, raw: str, rules_path: Path | None, kind: str = "") -> MatchResult:
        """Against the category template's drop-down when it has one, else ETM's reference."""
        dropdown = self.dropdown_for(attr_name)
        lists = self.value_lists(rules_path)
        if lists is not None:
            return lists.match(attr_name, raw, kind, dropdown)
        return match_value(raw, Characteristic(name=attr_name, values=list(dropdown)))

    def with_current_kind(self, product: ProductRecord, raw: str, source: str) -> tuple[str, str, str]:
        """For values like "AC 220": the product name ("... AC/DC230В") wins, else the "Род тока" attributes."""
        # The short name is cut at 100 characters ("... AC23"), so the full name goes first.
        for text in (product.full_name, product.name):
            kind, number = current_kind_from_text(text)
            if kind:
                return number, kind, "product_name_current"
        attrs = product.attributes or {}
        for name in CURRENT_KIND_ATTRIBUTES:
            kind = current_kind_from_value(attrs.get(name, ""))
            if kind:
                return raw, kind, source
        return raw, "", source

    def value_lists(self, rules_path: Path | None) -> ValueLists | None:
        if self.value_list_files is None or rules_path is None:
            return None
        reference = rules_path.parent / self.value_list_files[0]
        if not reference.exists():
            return None
        return cached_value_lists(reference, rules_path.parent / self.value_list_files[1])

    def learned_matches_path(self, rules_path: Path | None) -> Path | None:
        if self.value_list_files is None or rules_path is None:
            return None
        return rules_path.parent / self.value_list_files[1]

    def allowed_values_for(self, header: str, rules_path: Path | None) -> list[str]:
        columns = self.attribute_columns_for(header)
        lists = self.value_lists(rules_path)
        if columns is None or lists is None:
            return []
        characteristic = lists.characteristic(header.removeprefix(columns.prefix).strip())
        return list(characteristic.values) if characteristic else []

    def raw_value(self, product: ProductRecord, header: str, rules_path: Path | None) -> str:
        columns = self.attribute_columns_for(header)
        if columns is None:
            return ""
        return self.attribute_raw(product, header, columns, rules_path)[0]

    def candidates_for(self, header: str, raw: str, rules_path: Path | None) -> tuple[str, ...]:
        columns = self.attribute_columns_for(header)
        if columns is None:
            return ()
        attr_name = header.removeprefix(columns.prefix).strip()
        if self.value_lists(rules_path) is None and not self.dropdown_for(attr_name):
            return ()
        return self.match_allowed(attr_name, raw, rules_path).candidates

    def not_allowed_note(self, header: str, raw: str, rules_path: Path | None) -> str:
        candidates = self.candidates_for(header, raw, rules_path)
        note = f"В базе есть значение «{raw}», но его нет в списке допустимых значений {self.name}."
        if candidates:
            note += " Похожие допустимые: " + "; ".join(candidates) + "."
        return note + " Выберите значение на листе «Сопоставления» или впишите вручную, программа его запомнит."

    def is_reportable(self, header: str) -> bool:
        return header in self.fields or self.attribute_columns_for(header) is not None

    def missing_note(self, header: str) -> str:
        spec = self.fields.get(header)
        if spec is not None and spec.note:
            return self.notes[spec.note]
        for prefix, note in self.note_prefixes:
            if header.startswith(prefix):
                return self.notes[note]
        return self.notes["default"]

    def find_article_index(self, headers: list[str]) -> int | None:
        for candidate in self.article_headers:
            if candidate in headers:
                return headers.index(candidate)
        return None

    def class_index(self, headers: list[str]) -> int | None:
        return headers.index(self.class_header) if self.class_header in headers else None

    def with_template_class(self, product: ProductRecord, template_class: str) -> ProductRecord:
        if template_class and not getattr(product, self.class_field):
            return replace(product, **{self.class_field: template_class})
        return product

    @property
    def fills_only_yellow_headers(self) -> bool:
        return self.writer.get("fillable_columns") == "yellow_header"


def _as_tuple(value: Any) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        return (value,)
    return tuple(value)


def parse_partner_template(data: dict[str, Any]) -> PartnerTemplate:
    notes = dict(data.get("notes") or {})
    notes.setdefault("default", "В базе нет значения для этого поля.")
    fields: dict[str, FieldSpec] = {}
    for item in data.get("fields", []):
        spec = FieldSpec(
            header=item["header"],
            fields=_as_tuple(item.get("field")),
            const=item.get("const"),
            when_present=item.get("when_present", ""),
            transform=item.get("transform", ""),
            digits=int(item.get("digits", 6)),
            limit=int(item.get("limit", 60)),
            source=item.get("source", ""),
            note=item.get("note", ""),
        )
        fields[spec.header] = spec
    classification = data.get("classification") or {}
    value_lists = data.get("value_lists") or {}
    template = PartnerTemplate(
        code=data["code"],
        name=data.get("name", data["code"]),
        writer=dict(data.get("writer") or {}),
        article_headers=_as_tuple(data.get("article_headers")),
        skip_headers=frozenset(_as_tuple(data.get("skip_headers"))),
        class_header=classification.get("header", ""),
        class_field=classification.get("field", "class81_code"),
        notes=notes,
        note_prefixes=tuple((item["prefix"], item["note"]) for item in data.get("note_prefixes", [])),
        fields=fields,
        attribute_columns=tuple(
            AttributeColumns(
                prefix=item["prefix"],
                source=item.get("source", ""),
                rule_source_prefix=item.get("rule_source_prefix", ""),
                note=item.get("note", ""),
            )
            for item in data.get("attribute_columns", [])
        ),
        value_list_files=(
            (value_lists["reference"], value_lists["learned"]) if value_lists else None
        ),
        name_aliases=parse_name_aliases(data.get("name_aliases", [])),
        name_rejects={
            canonical(item["characteristic"]): frozenset(_as_tuple(item.get("attributes")))
            for item in data.get("name_rejects", [])
        },
        checks=frozenset(name for name, enabled in (data.get("checks") or {}).items() if enabled),
    )
    validate_partner_template(template)
    return template


def parse_name_aliases(items: list[dict[str, Any]]) -> dict[str, tuple[NameAlias, ...]]:
    """Pairs by characteristic, in file order: the first pair that gives a value wins."""
    result: dict[str, tuple[NameAlias, ...]] = {}
    for item in items:
        alias = NameAlias(
            attributes=_as_tuple(item.get("attributes")),
            values=dict(item.get("values") or {}),
            check=bool(item.get("check", False)),
            only_listed=bool(item.get("only_listed", False)),
            scale=float(item["scale"]) if item.get("scale") else None,
            range_pick=item.get("range", ""),
            classes=frozenset(str(code) for code in _as_tuple(item.get("classes"))),
        )
        key = canonical(item["characteristic"])
        result[key] = (*result.get(key, ()), alias)
    return result


def validate_partner_template(template: PartnerTemplate) -> None:
    product_fields = set(ProductRecord.__dataclass_fields__)
    used_notes = {spec.note for spec in template.fields.values() if spec.note}
    used_notes.update(note for _, note in template.note_prefixes)
    unknown_notes = used_notes - set(template.notes)
    if unknown_notes:
        raise ValueError(f"Шаблон {template.code}: нет текста для заметок {sorted(unknown_notes)}")
    for spec in template.fields.values():
        unknown = [name for name in (*spec.fields, spec.when_present) if name and name not in product_fields]
        if unknown:
            raise ValueError(f"Шаблон {template.code}, колонка {spec.header}: нет полей товара {unknown}")
    if template.class_field not in product_fields:
        raise ValueError(f"Шаблон {template.code}: нет поля товара {template.class_field}")


def partner_template_text(code: str) -> str:
    return resources.files(partners).joinpath(f"{code}.json").read_text(encoding="utf-8")


@lru_cache(maxsize=None)
def load_partner_template(code: str = DEFAULT_PARTNER) -> PartnerTemplate:
    try:
        raw = partner_template_text(code)
    except FileNotFoundError as exc:
        raise ValueError(f"Нет шаблона партнёра: {code}") from exc
    return parse_partner_template(json.loads(raw))


def available_partners() -> list[str]:
    folder = resources.files(partners)
    return sorted(item.name.removesuffix(".json") for item in folder.iterdir() if item.name.endswith(".json"))


def etm_reference_path(rules_path: Path) -> Path:
    """Where the imported ETM reference lives: next to the mapping rules file."""
    template = load_partner_template("etm")
    assert template.value_list_files is not None
    return rules_path.parent / template.value_list_files[0]
