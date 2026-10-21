from __future__ import annotations

import math
import re


_SPACES = ("\xa0", " ", " ", " ")
_PLAIN_RE = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")


def _grouped(text: str, group: str, decimal: str | None) -> str | None:
    pattern = rf"([+-]?\d{{1,3}}(?:{re.escape(group)}\d{{3}})+)"
    if decimal is not None:
        pattern += rf"(?:{re.escape(decimal)}(\d+))?"
    match = re.fullmatch(pattern, text)
    if not match:
        return None
    integer = match.group(1).replace(group, "")
    fraction = match.group(2) if decimal is not None else None
    return f"{integer}.{fraction}" if fraction else integer


def parse_number(value: object) -> float | None:
    """Strictly parse a finite number from a cell value or Russian/English text.

    Accepts a decimal comma or dot and space/dot/comma thousands separators
    (``0,125``, ``1 234,5``, ``1.234,5``, ``1,234.5``). Words, Excel errors,
    ranges and NaN/Inf return ``None`` instead of a guessed number.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None

    text = str(value)
    for space in _SPACES:
        text = text.replace(space, " ")
    text = " ".join(text.split())
    if not text:
        return None

    commas, dots = text.count(","), text.count(".")
    if " " in text:
        decimal = "," if commas else "."
        if commas and dots:
            return None
        normalized = _grouped(text, " ", decimal)
    elif commas and dots:
        decimal = "," if text.rindex(",") > text.rindex(".") else "."
        group = "." if decimal == "," else ","
        normalized = _grouped(text, group, decimal) if text.count(decimal) == 1 else None
    elif commas > 1 or dots > 1:
        normalized = _grouped(text, "," if commas else ".", None)
    else:
        normalized = text.replace(",", ".")

    if normalized is None or not _PLAIN_RE.fullmatch(normalized):
        return None
    number = float(normalized)
    return number if math.isfinite(number) else None
