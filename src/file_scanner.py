from __future__ import annotations

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


def has_document(record: DataRecord, indexed_names: list[str], pattern: str) -> bool:
    compact_key = record.llave.replace("-", "").lower()
    pedimento = record.pedimento.lower()

    for name in indexed_names:
        if pattern not in name:
            continue
        if compact_key in name or pedimento in name:
            return True
    return False
