from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from importlib import resources


@dataclass(frozen=True)
class AttributeRule:
    class81_code: str
    template_field: str
    source_attributes: tuple[str, ...]
    confidence: str
    note: str = ""


@lru_cache(maxsize=1)
def load_attribute_rules() -> tuple[AttributeRule, ...]:
    raw = resources.files("chint_etm_mdm.rules").joinpath("attribute_mappings.json").read_text(
        encoding="utf-8"
    )
    data = json.loads(raw)
    rules = []
    for item in data.get("class_rules", []):
        rules.append(
            AttributeRule(
                class81_code=str(item.get("class81_code") or ""),
                template_field=str(item.get("template_field") or ""),
                source_attributes=tuple(str(v) for v in item.get("source_attributes", [])),
                confidence=str(item.get("confidence") or "candidate"),
                note=str(item.get("note") or ""),
            )
        )
    return tuple(rules)


def source_attributes_for(class81_code: str, template_field: str) -> tuple[str, ...]:
    matches: list[str] = []
    for rule in load_attribute_rules():
        if rule.class81_code != class81_code:
            continue
        if rule.template_field != template_field:
            continue
        if rule.confidence != "approved_class_rule":
            continue
        matches.extend(rule.source_attributes)
    return tuple(dict.fromkeys(matches))


def rules_for(class81_code: str, template_field: str) -> tuple[AttributeRule, ...]:
    return tuple(
        rule
        for rule in load_attribute_rules()
        if rule.class81_code == class81_code and rule.template_field == template_field
    )
