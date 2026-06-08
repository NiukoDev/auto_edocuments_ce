from __future__ import annotations

import re
from pathlib import Path

from models import DataRecord


def build_file_index(root: str | Path) -> list[str]:
    root_path = Path(root)
    if not root_path.exists() or not root_path.is_dir():
        return []

    indexed_names: list[str] = []
    for path in root_path.rglob("*"):
        if path.is_file() and not path.name.startswith("~$"):
            indexed_names.append(str(path.relative_to(root_path)).lower())
    return indexed_names


def has_document(record: DataRecord, indexed_names: list[str], pattern: str, extension: str) -> bool:
    compact_key = record.llave.replace("-", "").lower()
    pedimento = record.pedimento.lower()

    for name in indexed_names:
        if pattern not in name:
            continue
        if extension and not name.endswith(extension.lower()):
            continue
        if compact_key in name or pedimento in name:
            return True
    return False


def has_document_by_llave(record: DataRecord, indexed_names: list[str], extension: str) -> bool:
    target = record.llave.lower() + extension.lower()
    for name in indexed_names:
        filename = name.split("/")[-1]
        if filename == target:
            return True
    return False


def has_document_by_prefix(record: DataRecord, indexed_names: list[str], prefix: str, extension: str) -> bool:
    compact_key = record.llave.replace("-", "").lower()
    pedimento = record.pedimento.lower()
    prefix_lower = prefix.lower()
    ext_lower = extension.lower()

    for name in indexed_names:
        filename = name.split("/")[-1]
        if not filename.startswith(prefix_lower):
            continue
        if not filename.endswith(ext_lower):
            continue
        if compact_key in name or pedimento in name:
            return True
    return False


def _has_payment_validation_file(record: DataRecord, indexed_names: list[str], rule: dict) -> bool:
    compact_key = record.llave.replace("-", "").lower()
    pedimento = record.pedimento.lower()
    pattern = rule.get("pattern")
    extension = rule.get("extension", "")
    file_regex = rule.get("regex")

    for name in indexed_names:
        if compact_key not in name and pedimento not in name:
            continue
        filename = name.split("/")[-1]

        if file_regex is not None:
            if re.fullmatch(file_regex, filename, re.IGNORECASE):
                return True
        elif pattern is not None:
            if pattern not in name:
                continue
            if extension == "numeric3":
                dot_pos = filename.rfind(".")
                if dot_pos != -1:
                    ext = filename[dot_pos + 1:]
                    if len(ext) == 3 and ext.isdigit():
                        return True
            elif extension and filename.endswith(extension.lower()):
                return True
    return False


def find_document_path(record: DataRecord, indexed_names: list[str], pattern: str, extension: str) -> str:
    compact_key = record.llave.replace("-", "").lower()
    pedimento = record.pedimento.lower()

    for name in indexed_names:
        if pattern not in name:
            continue
        if extension and not name.endswith(extension.lower()):
            continue
        if compact_key in name or pedimento in name:
            return name
    return ""


def missing_payment_validation_files(
    record: DataRecord,
    indexed_names: list[str],
    payment_rules: list[dict],
) -> list[str]:
    return [
        rule["label"]
        for rule in payment_rules
        if not _has_payment_validation_file(record, indexed_names, rule)
    ]
