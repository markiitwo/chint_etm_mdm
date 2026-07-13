from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from importlib import resources
from pathlib import Path


@dataclass(frozen=True)
class AttributeRule:
    class81_code: str
    template_field: str
    source_attributes: tuple[str, ...]
    confidence: str
    note: str = ""


DEFAULT_RULES_RESOURCE = "attribute_mappings.json"


def default_rules_text() -> str:
    return resources.files("chint_etm_mdm.rules").joinpath(DEFAULT_RULES_RESOURCE).read_text(
        encoding="utf-8"
    )


def workdir_rules_path(work_dir: Path) -> Path:
    return work_dir / "rules" / DEFAULT_RULES_RESOURCE


def ensure_default_rules(work_dir: Path) -> Path:
    path = workdir_rules_path(work_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        path.write_text(default_rules_text(), encoding="utf-8")
    return path


@lru_cache(maxsize=32)
def load_attribute_rules(rules_path_text: str = "") -> tuple[AttributeRule, ...]:
    rules_path = Path(rules_path_text) if rules_path_text else None
    if rules_path and rules_path.exists():
        raw = rules_path.read_text(encoding="utf-8")
    else:
        raw = default_rules_text()
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


def source_attributes_for(
    class81_code: str, template_field: str, rules_path: Path | None = None
) -> tuple[str, ...]:
    matches: list[str] = []
    for rule in load_attribute_rules(str(rules_path or "")):
        if rule.class81_code != class81_code:
            continue
        if rule.template_field != template_field:
            continue
        if rule.confidence != "approved_class_rule":
            continue
        matches.extend(rule.source_attributes)
    return tuple(dict.fromkeys(matches))


def rules_for(
    class81_code: str, template_field: str, rules_path: Path | None = None
) -> tuple[AttributeRule, ...]:
    return tuple(
        rule
        for rule in load_attribute_rules(str(rules_path or ""))
        if rule.class81_code == class81_code and rule.template_field == template_field
    )
