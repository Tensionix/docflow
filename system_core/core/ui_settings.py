from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .config import load_yaml_or_json


SUPPORTED_LANGUAGES = {"en", "ru"}


@dataclass
class UiSettings:
    language: str = "ru"
    theme: str = "code_dark"
    emoji: bool = False
    allow_runtime_switching: bool = True
    advanced_open: bool = False


def _safe_language(value: Any) -> str:
    text = str(value or "ru").strip().lower()
    return text if text in SUPPORTED_LANGUAGES else "ru"


def _safe_theme(value: Any) -> str:
    text = str(value or "code_dark").strip().lower()
    cleaned = "".join(char for char in text if char.isalnum() or char in {"_", "-"})
    return cleaned or "code_dark"


def load_ui_settings(path: Path) -> UiSettings:
    data = load_yaml_or_json(path) if path.exists() else {}
    ui_data = data.get("gui", data) if isinstance(data, dict) else {}
    if not isinstance(ui_data, dict):
        ui_data = {}
    return UiSettings(
        language=_safe_language(ui_data.get("language", "ru")),
        theme=_safe_theme(ui_data.get("theme", "code_dark")),
        emoji=bool(ui_data.get("emoji", False)),
        allow_runtime_switching=bool(ui_data.get("allow_runtime_switching", True)),
        advanced_open=bool(ui_data.get("advanced_open", False)),
    )


def save_ui_settings(path: Path, settings: UiSettings) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = (
        "gui:\n"
        "  # Change to \"en\" for public GitHub builds.\n"
        f"  language: \"{_safe_language(settings.language)}\"\n"
        f"  theme: \"{_safe_theme(settings.theme)}\"\n"
        f"  emoji: {str(bool(settings.emoji)).lower()}\n"
        f"  allow_runtime_switching: {str(bool(settings.allow_runtime_switching)).lower()}\n"
        f"  advanced_open: {str(bool(settings.advanced_open)).lower()}\n"
    )
    path.write_text(text, encoding="utf-8", newline="\n")
