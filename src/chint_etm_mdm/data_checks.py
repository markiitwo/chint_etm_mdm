"""Checks that hold back values which may be wrong before they go to a partner.

A value the program filled is "safe" unless one of these checks fires; then the
product (new cards) or the row (characteristics report) goes to a separate
"to check" file, and the report says why. Which checks a partner uses is set in
its template ("checks"):

- "name_numbers": the product name contradicts the value. A transformer named
  "220/24*2" with a secondary voltage of 48 in the database is a data error: the
  name is what the product is sold under, so a voltage or a current the name
  does not mention is held back.
- "ranges": the value is a range ("AC 110-230"). Not yet known whether ETM
  accepts ranges, so they are uploaded separately until it is confirmed.
"""
from __future__ import annotations

import re

from .value_matching import canonical

NUMBER_RE = re.compile(r"\d+(?:[.,]\d+)?")
RANGE_RE = re.compile(r"\d(?:[.,]\d+)?\s*(?:-|–|…|\.\.\.)\s*\d")
UNIT_MARKS = {
    "в": re.compile(r"\d\s*(?:в|v|vac|vdc)(?![a-zа-я])|\d+/\d+", re.IGNORECASE),
    "а": re.compile(r"\d\s*(?:а|a)(?![a-zа-я])|(?<![a-zа-я])[bcdk]\d+", re.IGNORECASE),
}
# Characteristics a name rarely carries: their values are not checked against it.
SKIP_WORDS = ("изоляц", "импульс", "испыт", "максимал", "минимал", "отключающ", "кратковрем", "пиков")
CURRENT_WORD_RE = re.compile(r"(?<![а-я])ток")  # not "поток"


def characteristic_unit(characteristic: str) -> str:
    name = canonical(characteristic)
    if any(word in name for word in SKIP_WORDS):
        return ""
    if "напряжен" in name:
        return "в"
    if CURRENT_WORD_RE.search(name.split("(")[0]) and "утечк" not in name and "диф" not in name:
        return "а"
    return ""


def numbers(text: str) -> set[float]:
    return {float(item.replace(",", ".")) for item in NUMBER_RE.findall(text)}


def name_conflict(characteristic: str, value: str, product_name: str) -> str:
    """A note when the name gives this kind of value and the value is not among its numbers."""
    unit = characteristic_unit(characteristic)
    if not unit or not product_name or not UNIT_MARKS[unit].search(product_name):
        return ""
    wanted = numbers(value)
    if not wanted:
        return ""
    missing = wanted - numbers(product_name)
    if not missing:
        return ""
    shown = ", ".join(f"{item:g}" for item in sorted(missing))
    return f"Противоречие: в названии товара нет числа {shown}; проверьте значение в базе."


def is_range(value: str) -> bool:
    return bool(RANGE_RE.search(value))


def review_note(checks: frozenset[str], characteristic: str, value: str, product_name: str) -> str:
    """Why a filled value should be checked before upload; empty when it is safe."""
    if not checks or not value:
        return ""
    if "name_numbers" in checks:
        note = name_conflict(characteristic, value, product_name)
        if note:
            return note
    if "ranges" in checks and is_range(value):
        return "Диапазон значений: ЭТМ может его не принять, грузите отдельно."
    return ""
