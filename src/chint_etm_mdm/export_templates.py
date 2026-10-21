"""Partner export templates.

The product card (ProductRecord) knows nothing about partners. Each partner is
described by a JSON file in the ``partners`` package: which column gets which
product field, how the value is formatted, and what to say when it is missing.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, replace
from functools import lru_cache
from importlib import resources
from pathlib import Path
from typing import Any

from . import partners
from .db import ProductRecord
from .mapping_rules import source_attributes_for
from .value_matching import ValueLists, same_name_attributes

DEFAULT_PARTNER = "etm"
NAME_ALIAS_SOURCE_PREFIX = "name_alias:"
SAME_NAME_SOURCE_PREFIX = "same_name:"


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
    if spec.transform == "decimal":
        return format_decimal(value, spec.digits)
    if spec.transform == "mm_to_m":
        return mm_to_m(value)
    if spec.transform == "count":
        return format_count(value)
    if spec.transform == "truncate":
        return short_wms_name(value or "", spec.limit)
    raise ValueError(f"Неизвестное преобразование в шаблоне партнёра: {spec.transform}")


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
        if not raw:
            return "", "blank", ""
        lists = self.value_lists(rules_path)
        match = lists.match(header.removeprefix(columns.prefix).strip(), raw) if lists else None
        if match is None:
            return raw, "filled", source
        if match.confident and not source.startswith(SAME_NAME_SOURCE_PREFIX):
            return match.value, "filled", source
        if match.confident or match.kind == "words":
            return match.value, "suggested", source
        return raw, "value_not_allowed", source

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
        lists = self.value_lists(rules_path)
        if columns is None or lists is None:
            return ()
        return lists.match(header.removeprefix(columns.prefix).strip(), raw).candidates

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
    )
    validate_partner_template(template)
    return template


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
