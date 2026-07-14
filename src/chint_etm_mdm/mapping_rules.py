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
DEFAULT_RULES_FALLBACK = {
    "version": 1,
    "class_rules": [
        {
            "class81_code": "50301005",
            "template_field": "Конфиг:Напряжение, В",
            "source_attributes": ["Напряжение лампы, В", "Напряжение лампы"],
            "confidence": "approved_class_rule",
            "note": "For signal lamps this field means lamp voltage.",
        },
        {
            "class81_code": "50301005",
            "template_field": "Конфиг:Цвет свечения",
            "source_attributes": ["Цвет"],
            "confidence": "approved_class_rule",
            "note": "For signal lamps the source attribute is generic color.",
        },
    ],
}


def default_rules_text() -> str:
    try:
        return resources.files("chint_etm_mdm.rules").joinpath(DEFAULT_RULES_RESOURCE).read_text(
            encoding="utf-8"
        )
    except (FileNotFoundError, ModuleNotFoundError):
        return json.dumps(DEFAULT_RULES_FALLBACK, ensure_ascii=False, indent=2)


def workdir_rules_path(work_dir: Path) -> Path:
    return work_dir / "rules" / DEFAULT_RULES_RESOURCE


def ensure_default_rules(work_dir: Path) -> Path:
    path = workdir_rules_path(work_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        path.write_text(default_rules_text(), encoding="utf-8")
    return path


def read_rules_document(rules_path: Path) -> dict:
    if rules_path.exists():
        raw = rules_path.read_text(encoding="utf-8")
    else:
        raw = default_rules_text()
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("Rules file must contain a JSON object.")
    data.setdefault("version", 1)
    data.setdefault("class_rules", [])
    if not isinstance(data["class_rules"], list):
        raise ValueError("Rules file field class_rules must be a list.")
    return data


def write_rules_document(rules_path: Path, data: dict) -> None:
    rules_path.parent.mkdir(parents=True, exist_ok=True)
    rules_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    load_attribute_rules.cache_clear()


def add_approved_class_rule(
    rules_path: Path,
    class81_code: str,
    template_field: str,
    source_attribute: str,
    note: str = "Confirmed in GUI mapping review.",
) -> bool:
    class81_code = str(class81_code or "").strip()
    template_field = str(template_field or "").strip()
    source_attribute = str(source_attribute or "").strip()
    if not class81_code or not template_field or not source_attribute:
        raise ValueError("81 class, template field and source attribute are required.")

    data = read_rules_document(rules_path)
    new_rule = {
        "class81_code": class81_code,
        "template_field": template_field,
        "source_attributes": [source_attribute],
        "confidence": "approved_class_rule",
        "note": note,
    }

    for item in data["class_rules"]:
        if not isinstance(item, dict):
            continue
        if str(item.get("class81_code") or "").strip() != class81_code:
            continue
        if str(item.get("template_field") or "").strip() != template_field:
            continue
        if str(item.get("confidence") or "").strip() != "approved_class_rule":
            continue
        sources = [str(value) for value in item.get("source_attributes", []) if str(value).strip()]
        if source_attribute in sources:
            return False
        sources.append(source_attribute)
        item["source_attributes"] = list(dict.fromkeys(sources))
        if not str(item.get("note") or "").strip():
            item["note"] = note
        write_rules_document(rules_path, data)
        return True

    data["class_rules"].append(new_rule)
    write_rules_document(rules_path, data)
    return True


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
