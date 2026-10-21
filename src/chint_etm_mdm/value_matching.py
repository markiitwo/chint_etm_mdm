"""Matching product characteristic values to a partner's allowed values.

ETM publishes a reference ("Справочник") with the allowed values of each
characteristic. A value outside that list is rejected by ETM, so before a value
goes into a closed-list column it is matched to one of the allowed values:

- exact or after harmless normalization (case, ё/е, spaces, decimal comma);
- by a match the user confirmed earlier (learned matches);
- yes/no wording for logical characteristics;
- the same number ("230 В" -> 230);
- by words, when exactly one allowed value has the same words in another order
  or spelling ("AC (перемен.)" -> "Переменный (AC)"); such a value needs a check.
"""
from __future__ import annotations

import difflib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from .atomic_files import atomic_write_json

LOGICAL_YES = {"да", "yes", "true", "1", "есть", "имеется", "наличие", "y"}
LOGICAL_NO = {"нет", "no", "false", "0", "отсутствует", "без", "n", "-"}
WORD_SYNONYMS = {
    "перемен": "переменный",
    "переменного": "переменный",
    "переменное": "переменный",
    "пост": "постоянный",
    "постоян": "постоянный",
    "постоянного": "постоянный",
    "постоянное": "постоянный",
    "ac": "ac",
    "dc": "dc",
}
NUMBER_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
WORD_RE = re.compile(r"[0-9a-zа-я]+")


def canonical(value: str) -> str:
    text = str(value).replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е").casefold()
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"\s*([/,;:()+\-])\s*", r"\1", text)
    return " ".join(text.split()).strip(" .")


def words(value: str) -> set[str]:
    return {WORD_SYNONYMS.get(word, word) for word in WORD_RE.findall(canonical(value))}


def only_number(value: str) -> float | None:
    """Return the number if the value is a single number, optionally with a unit."""
    text = canonical(value)
    numbers = NUMBER_RE.findall(text)
    if len(numbers) != 1:
        return None
    rest = NUMBER_RE.sub("", text, count=1)
    if re.search(r"\d", rest) or len(rest.replace(" ", "")) > 6:
        return None
    return float(numbers[0].replace(",", "."))


def format_allowed(value: object) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


@dataclass
class Characteristic:
    name: str
    code: str = ""
    formats: list[str] = field(default_factory=list)
    values: list[str] = field(default_factory=list)
    comment: str = ""

    @property
    def closed(self) -> bool:
        return bool(self.values)


@dataclass(frozen=True)
class MatchResult:
    value: str
    kind: str
    candidates: tuple[str, ...] = ()

    @property
    def confident(self) -> bool:
        return self.kind in {"exact", "normalized", "learned", "logical", "number"}


def match_value(
    raw: str,
    characteristic: Characteristic,
    learned: dict[str, str] | None = None,
) -> MatchResult:
    allowed = characteristic.values
    if not allowed:
        return MatchResult(raw, "open")
    if raw in allowed:
        return MatchResult(raw, "exact")
    by_canonical = {canonical(value): value for value in allowed}
    key = canonical(raw)
    if key in by_canonical:
        return MatchResult(by_canonical[key], "normalized")
    learned_value = (learned or {}).get(key)
    if learned_value and learned_value in allowed:
        return MatchResult(learned_value, "learned")
    if set(by_canonical) <= {"да", "нет"}:
        if key in LOGICAL_YES and "да" in by_canonical:
            return MatchResult(by_canonical["да"], "logical")
        if key in LOGICAL_NO and "нет" in by_canonical:
            return MatchResult(by_canonical["нет"], "logical")
    number = only_number(raw)
    numeric = [(value, only_number(value)) for value in allowed]
    numeric = [(value, item) for value, item in numeric if item is not None]
    if number is not None:
        same = [value for value, item in numeric if item == number]
        if len(same) == 1:
            return MatchResult(same[0], "number")
    raw_words = words(raw)
    if raw_words:
        same_words = [value for value in allowed if words(value) == raw_words]
        if len(same_words) == 1:
            return MatchResult(same_words[0], "words")
    if number is not None and numeric:
        nearest = sorted(numeric, key=lambda pair: abs(pair[1] - number))[:3]
        return MatchResult("", "none", tuple(value for value, _ in nearest))
    close = difflib.get_close_matches(key, list(by_canonical), n=3, cutoff=0.5)
    if not close and raw_words:
        close = [canonical(value) for value in allowed if raw_words & words(value)][:3]
    return MatchResult("", "none", tuple(by_canonical[item] for item in close))


class ValueLists:
    """Allowed values of one partner plus the matches the user confirmed."""

    def __init__(self, reference_path: Path, learned_path: Path) -> None:
        self.reference_path = reference_path
        self.learned_path = learned_path
        self.characteristics = load_reference(reference_path)
        self.learned = load_learned(learned_path)

    def characteristic(self, name: str) -> Characteristic | None:
        return self.characteristics.get(canonical(name))

    def match(self, name: str, raw: str) -> MatchResult | None:
        characteristic = self.characteristic(name)
        if characteristic is None or not characteristic.closed:
            return None
        return match_value(raw, characteristic, self.learned.get(canonical(name)))


def load_reference(path: Path) -> dict[str, Characteristic]:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    result: dict[str, Characteristic] = {}
    for item in data.get("characteristics", []):
        characteristic = Characteristic(
            name=item["name"],
            code=str(item.get("code", "")),
            formats=list(item.get("formats", [])),
            values=[str(value) for value in item.get("values", [])],
            comment=item.get("comment", ""),
        )
        result[canonical(characteristic.name)] = characteristic
    return result


def load_learned(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    learned: dict[str, dict[str, str]] = {}
    for item in data.get("matches", []):
        name = canonical(item.get("characteristic", ""))
        raw = canonical(item.get("raw", ""))
        value = str(item.get("value", "")).strip()
        if name and raw and value:
            learned.setdefault(name, {})[raw] = value
    return learned


def remember_matches(path: Path, matches: list[tuple[str, str, str]]) -> int:
    """Save (characteristic, raw value, allowed value) triples; return how many are new."""
    data = {"version": 1, "matches": []}
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
        data.setdefault("matches", [])
    by_key = {
        (canonical(item.get("characteristic", "")), canonical(item.get("raw", ""))): item
        for item in data["matches"]
    }
    added = 0
    for characteristic, raw, value in matches:
        key = (canonical(characteristic), canonical(raw))
        if not key[0] or not key[1]:
            continue
        item = by_key.get(key)
        if item is None:
            item = {"characteristic": characteristic, "raw": raw}
            data["matches"].append(item)
            by_key[key] = item
            added += 1
        item["value"] = value
    path.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_json(path, data)
    return added


def read_etm_reference_workbook(path: Path) -> list[Characteristic]:
    """Read the ETM OneTeam "Справочник" export: one row per allowed value."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook["Справочник"] if "Справочник" in workbook.sheetnames else workbook.worksheets[0]
        rows = sheet.iter_rows(values_only=True)
        header = [str(cell or "").strip().casefold() for cell in next(rows, ())]

        def column(*names: str) -> int:
            for idx, title in enumerate(header):
                if any(title.startswith(name) for name in names):
                    return idx
            raise ValueError(f"В справочнике нет колонки «{names[0]}».")

        code_idx = column("код характеристики")
        name_idx = column("название характеристики", "наименование характеристики")
        format_idx = column("формат")
        value_idx = column("возможные значения")
        comment_idx = column("комментарий")
        result: dict[str, Characteristic] = {}
        for row in rows:
            if not row or row[name_idx] in (None, ""):
                continue
            name = str(row[name_idx]).strip()
            characteristic = result.setdefault(canonical(name), Characteristic(name=name))
            code = format_allowed(row[code_idx]) if row[code_idx] is not None else ""
            if code and not characteristic.code:
                characteristic.code = code
            kind = str(row[format_idx] or "").strip()
            if kind and kind not in characteristic.formats:
                characteristic.formats.append(kind)
            value = row[value_idx]
            if value not in (None, ""):
                text = format_allowed(value)
                if text not in characteristic.values:
                    characteristic.values.append(text)
            comment = str(row[comment_idx] or "").strip() if comment_idx < len(row) else ""
            if comment and not characteristic.comment:
                characteristic.comment = comment
        return list(result.values())
    finally:
        workbook.close()


def import_etm_reference(source: Path, target: Path) -> tuple[int, int]:
    """Convert the ETM reference XLSX into the program's JSON; return (characteristics, values)."""
    characteristics = read_etm_reference_workbook(source)
    if not characteristics:
        raise ValueError("В справочнике не найдено ни одной характеристики.")
    target.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_json(
        target,
        {
            "version": 1,
            "source_file": source.name,
            "characteristics": [
                {
                    "name": item.name,
                    "code": item.code,
                    "formats": item.formats,
                    "values": item.values,
                    "comment": item.comment,
                }
                for item in characteristics
            ],
        },
    )
    return len(characteristics), sum(len(item.values) for item in characteristics)
