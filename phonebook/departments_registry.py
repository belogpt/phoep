"""Работа с реестром сокращений отделов."""
import json
import os
from typing import Dict

from . import repository

ALIASES_FILENAME = 'departments_aliases.json'


def _get_aliases_path() -> str:
    remote_dir, _ = repository.get_paths()
    os.makedirs(remote_dir, exist_ok=True)
    return os.path.join(remote_dir, ALIASES_FILENAME)


def load_department_aliases() -> Dict[str, str]:
    """Читает JSON с сокращениями отделов."""

    path = _get_aliases_path()
    if not os.path.exists(path):
        return {}
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_department_aliases(mapping: Dict[str, str]) -> None:
    """Сохраняет JSON с сокращениями отделов."""

    path = _get_aliases_path()
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)


def suggest_alias(full_name: str) -> str:
    """Простая эвристика сокращения: одно слово или аббревиатура из первых букв."""

    parts = [p for p in full_name.split() if p]
    if not parts:
        return ''
    if len(parts) == 1:
        return parts[0]
    return ''.join(p[0].upper() for p in parts)

