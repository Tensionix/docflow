from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any
import atexit
import argparse
import ctypes
import ipaddress
import json
import logging
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from ctypes import wintypes

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from nicegui import app as nicegui_app, run, ui  # type: ignore

AUDION_CANONICAL_TOOLTIP_DELAY_MS = 1500
AUDION_CANONICAL_TOOLTIP_HIDE_DELAY_MS = 100
AUDION_CANONICAL_TOOLTIP_TRANSITION_MS = 100


def install_audion_canonical_tooltip_defaults() -> None:
    try:
        from nicegui.elements.tooltip import Tooltip as NiceGuiTooltip  # type: ignore
    except Exception:
        return
    if getattr(NiceGuiTooltip, "_audion_canonical_tooltip_defaults", False):
        return
    original_init = NiceGuiTooltip.__init__

    def audion_tooltip_init(self: Any, text: str = "") -> None:
        original_init(self, text)
        self.props["delay"] = AUDION_CANONICAL_TOOLTIP_DELAY_MS
        self.props["hide-delay"] = AUDION_CANONICAL_TOOLTIP_HIDE_DELAY_MS
        self.props["transition-duration"] = AUDION_CANONICAL_TOOLTIP_TRANSITION_MS
        self.classes("audion-tooltip")

    NiceGuiTooltip.__init__ = audion_tooltip_init  # type: ignore[method-assign]
    NiceGuiTooltip._audion_canonical_tooltip_defaults = True  # type: ignore[attr-defined]


install_audion_canonical_tooltip_defaults()


AUDION_CANONICAL_UI_CSS = """
<style id="audion-canonical-tooltip-icon-style">
  html body .q-tooltip,
  html body .audion-tooltip {
    background: rgb(23, 33, 43) !important;
    background-color: rgb(23, 33, 43) !important;
    color: #f4f8fb !important;
    border: 1px solid rgba(88, 166, 255, 0.24) !important;
    border-radius: 8px !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, 0.34) !important;
  }
  html body .q-icon.material-icons,
  html body .q-icon.material-symbols-outlined,
  html body .q-icon.material-symbols-rounded,
  html body i.material-icons,
  html body i.material-symbols-outlined,
  html body i.material-symbols-rounded,
  html body .q-btn .q-icon,
  html body .q-btn .material-icons,
  html body .q-btn .material-symbols-outlined,
  html body .q-btn .material-symbols-rounded,
  html body .q-field .q-field__append .q-icon,
  html body .q-field .q-field__prepend .q-icon,
  html body .q-item .q-icon,
  html body .q-menu .q-icon,
  html body .audion-label-icon,
  html body .audion-path-option-pin,
  html body .audion-select-option-pin {
    font-size: 14px !important;
    width: 14px !important;
    min-width: 14px !important;
    height: 14px !important;
    line-height: 14px !important;
  }
  html body .material-icons,
  html body .q-icon.material-icons {
    font-family: "Material Icons" !important;
  }
  html body .material-symbols-outlined,
  html body .q-icon.material-symbols-outlined {
    font-family: "Material Symbols Outlined" !important;
  }
  html body .material-symbols-rounded,
  html body .q-icon.material-symbols-rounded {
    font-family: "Material Symbols Rounded" !important;
  }
</style>
"""


def add_audion_canonical_ui_styles() -> None:
    ui.add_head_html(AUDION_CANONICAL_UI_CSS)



def audion_tooltip_path_text(path_value: Any) -> str:
    raw = str(path_value or "").strip()
    if not raw:
        return ""
    try:
        path = Path(raw).expanduser()
        if not path.is_absolute():
            path = ROOT / path
        return str(path)
    except Exception:
        return raw


def audion_folder_button_tooltip(folder_id: str, path_value: Any) -> str:
    key = str(folder_id or "folder").strip().lower()
    path_text = audion_tooltip_path_text(path_value)
    if getattr(settings, "language", "ru") == "ru":
        descriptions = {
            "logs": "папку логов запусков и вывода терминала",
            "report": "папку отчётов и результатов операций",
            "reports": "папку отчётов и результатов операций",
            "config": "папку конфигурации проекта: manifest, GUI-настройки и кэши",
            "state": "папку рабочего состояния GUI",
            "project": "корневую папку проекта",
            "root": "корневую папку проекта",
            "data": "папку данных проекта",
            "pipeline": "папку pipeline-артефактов и промежуточных результатов",
            "github": "папку GitHub-артефактов проекта",
            "install": "папку install/runtime-артефактов проекта",
        }
        description = descriptions.get(key, f"папку {folder_id}")
        return f"Открыть {description}: {path_text}" if path_text else f"Открыть {description}."
    descriptions = {
        "logs": "the logs folder with run and terminal output",
        "report": "the reports/results folder",
        "reports": "the reports/results folder",
        "config": "the project config folder with manifest, GUI settings, and caches",
        "state": "the GUI state folder",
        "project": "the project root folder",
        "root": "the project root folder",
        "data": "the project data folder",
        "pipeline": "the pipeline artifacts and intermediate results folder",
        "github": "the project GitHub artifacts folder",
        "install": "the project install/runtime artifacts folder",
    }
    description = descriptions.get(key, f"the {folder_id} folder")
    return f"Open {description}: {path_text}" if path_text else f"Open {description}."


def audion_terminal_action_tooltip(action: str) -> str:
    key = str(action or "").strip().lower()
    if getattr(settings, "language", "ru") == "ru":
        tips = {
            "clear_terminal_window": "Очистить только видимое окно терминала. Файлы логов, отчёты и результаты операций не удаляются.",
            "expand": "Открыть терминал в большом окне, чтобы читать длинный вывод без тесной панели.",
            "expand_log": "Открыть терминал в большом окне, чтобы читать длинный вывод без тесной панели.",
            "pin_command": "Закрепить текущую команду в истории терминала для быстрого повторного запуска.",
            "unpin_command": "Открепить текущую команду от верхней части истории терминала.",
            "clear_history": "Очистить историю команд терминала. Закреплённые команды и файлы логов не удаляются.",
            "terminal_shell": "Выбрать оболочку, в которой будут запускаться команды терминала.",
            "terminal_history": "Выбрать ранее сохранённую или закреплённую команду терминала.",
            "terminal_command": "Команда, которая будет выполнена из выбранной рабочей папки.",
            "terminal_cwd": "Рабочая папка терминала. Команда будет запущена именно отсюда.",
            "pick_folder": "Выбрать рабочую папку терминала через системный диалог.",
            "terminal_run": "Запустить введённую команду в выбранной оболочке и рабочей папке.",
            "latest_report": "Открыть последний созданный отчёт, если он уже есть.",
            "command_preview": "Показать команду, которая будет запущена с текущими параметрами, без выполнения операции.",
            "report_view": "Открыть встроенный список отчётов без перехода в проводник.",
            "close": "Закрыть большое окно терминала и вернуться к основной панели.",
        }
    else:
        tips = {
            "clear_terminal_window": "Clear only the visible terminal window. Log files, reports, and operation results are not deleted.",
            "expand": "Open the terminal in a large window for reading long output comfortably.",
            "expand_log": "Open the terminal in a large window for reading long output comfortably.",
            "pin_command": "Pin the current terminal command for quick reuse.",
            "unpin_command": "Remove the current command from the pinned command list.",
            "clear_history": "Clear terminal command history. Pinned commands and log files are not deleted.",
            "terminal_shell": "Choose the shell used to run terminal commands.",
            "terminal_history": "Pick a saved or pinned terminal command.",
            "terminal_command": "Command to run from the selected working folder.",
            "terminal_cwd": "Terminal working folder. Commands are started from here.",
            "pick_folder": "Choose the terminal working folder with the system dialog.",
            "terminal_run": "Run the entered command in the selected shell and working folder.",
            "latest_report": "Open the latest generated report, if one exists.",
            "command_preview": "Show the command that would run with the current settings, without executing it.",
            "report_view": "Open the built-in reports list without switching to the file explorer.",
            "close": "Close the large terminal window and return to the main panel.",
        }
    return tips.get(key, key.replace("_", " ").strip())


from system_core.core.ansi_terminal import StatefulAnsiHtmlRenderer, terminal_html as _render_terminal_html
from system_core.core.config import load_yaml_or_json
from system_core.core.jobs import execute_operation
from system_core.core.manifest import CommandNode, Operation, load_manifest
from system_core.core.paths import ensure_project_dirs, get_project_paths, open_folder
from system_core.core.ui_settings import load_ui_settings, save_ui_settings
from system_core.ui_nicegui.workbench import (
    WORKBENCH_FEEDBACK_CSS,
    WORKBENCH_LAYOUT_CSS,
    WORKBENCH_OVERRIDE_CSS,
    WorkbenchAdapter,
    WorkbenchConfig,
    WorkbenchHandlers,
    WorkbenchRenderer,
    WorkbenchRole,
    canonical_role,
)


paths = get_project_paths(ROOT)
ensure_project_dirs(paths)
manifest = load_manifest(paths.config / "tool_manifest.yaml")
settings_path = paths.config / "gui_settings.yaml"
rules_folder = paths.config / "rules"
settings = load_ui_settings(settings_path)
tool_info: dict[str, Any] = manifest.raw.get("tool", {})
ui_info: dict[str, Any] = manifest.raw.get("ui", {})
TERMINAL_HISTORY_LIMIT = 1500
PATH_HISTORY_LIMIT = 100


def _startup_workspace_path(role: str, configured: str, legacy: str, default_path: Path) -> str:
    return str(default_path)


def load_workspace_route_settings() -> tuple[str, str]:
    return (
        _startup_workspace_path("source", "", "", paths.input),
        _startup_workspace_path("target", "", "", paths.output),
    )


def _yaml_string(value: Any) -> str:
    return "'" + str(value or "").replace("'", "''") + "'"


def _workspace_setting_for_disk(role: str, path_value: Any, default_path: Path) -> str:
    return ""


def display_path(path_value: Any) -> str:
    text = str(path_value or "").strip()
    if not text:
        return ""
    path = Path(text).expanduser()
    if not path.is_absolute():
        path = ROOT / path
    try:
        resolved = path.resolve()
        relative = resolved.relative_to(ROOT)
    except (OSError, ValueError):
        return str(path)
    return str(relative) or "."

def save_app_settings() -> None:
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    source_path = _workspace_setting_for_disk("source", getattr(settings, "source_path", ""), paths.input)
    destination_path = _workspace_setting_for_disk("target", getattr(settings, "destination_path", ""), paths.output)
    text = (
        "gui:\n"
        "  # Change to \"en\" for public GitHub builds.\n"
        f"  language: \"{settings.language if settings.language in {'en', 'ru'} else 'ru'}\"\n"
        f"  theme: \"{normalize_theme_id(settings.theme)}\"\n"
        f"  emoji: {str(bool(getattr(settings, 'emoji', False))).lower()}\n"
        f"  allow_runtime_switching: {str(bool(getattr(settings, 'allow_runtime_switching', True))).lower()}\n"
        f"  advanced_open: {str(bool(getattr(settings, 'advanced_open', False))).lower()}\n"
        f"  source_path: {_yaml_string(source_path)}\n"
        f"  destination_path: {_yaml_string(destination_path)}\n"
    )
    settings_path.write_text(text, encoding="utf-8", newline="\n")

settings.source_path, settings.destination_path = load_workspace_route_settings()

DEFAULT_THEME_ID = "code_dark"
THEME_ALIASES = {"dark": "code_dark", "light": "code_light"}


def _string_map(value: Any) -> dict[str, str]:
    return {str(key).strip(): str(item).strip() for key, item in dict(value).items() if str(key).strip()} if isinstance(value, dict) else {}


def render_terminal_html(lines: list[str] | tuple[str, ...]) -> str:
    return _render_terminal_html(lines).replace("\n", "")


def load_ui_colors(path: Path) -> dict[str, Any]:
    data = load_yaml_or_json(path) if path.exists() else {}
    if not isinstance(data, dict):
        data = {}
    themes: dict[str, dict[str, Any]] = {}
    themes_raw = data.get("themes", {})
    if not isinstance(themes_raw, dict):
        themes_raw = {}
    for theme_id, theme_data in themes_raw.items():
        if not isinstance(theme_data, dict):
            continue
        normalized_id = str(theme_id).strip().lower()
        if not normalized_id:
            continue
        themes[normalized_id] = {
            "label": str(theme_data.get("label") or normalized_id).strip(),
            "label_ru": str(theme_data.get("label_ru") or theme_data.get("label") or normalized_id).strip(),
            "mode": "dark" if str(theme_data.get("mode", "dark")).lower() == "dark" else "light",
            "tokens": _string_map(theme_data.get("tokens", {})),
        }
    if DEFAULT_THEME_ID not in themes:
        themes[DEFAULT_THEME_ID] = {
            "label": "Code Dark",
            "label_ru": "Code Темная",
            "mode": "dark",
            "tokens": {
                "color-background-primary": "#141413",
                "color-background-secondary": "#1f1e1a",
                "color-background-tertiary": "#0f0f0e",
                "color-text-primary": "#faf9f5",
                "color-text-secondary": "#e8e6dc",
                "color-text-tertiary": "#b0aea5",
                "color-border-tertiary": "rgba(250, 249, 245, 0.15)",
                "color-border-secondary": "rgba(250, 249, 245, 0.3)",
                "color-border-primary": "rgba(250, 249, 245, 0.4)",
                "color-accent-primary": "#d97757",
            },
        }
    return {
        "ramps": data.get("ramps", {}) if isinstance(data.get("ramps", {}), dict) else {},
        "tokens": _string_map(data.get("tokens", {})),
        "themes": themes,
    }


ui_colors = load_ui_colors(paths.config / "ui_colors.yaml")


def tolerate_missing_process_pool() -> None:
    """Keep NiceGUI alive when multiprocessing is blocked by the environment."""
    try:
        import nicegui.run as nicegui_run  # type: ignore
    except Exception:
        return

    original_setup = getattr(nicegui_run, "setup", None)
    if not callable(original_setup):
        return

    def safe_setup() -> None:
        try:
            original_setup()
        except (OSError, PermissionError) as exc:
            logging.warning("NiceGUI process pool disabled: %s", exc)
            nicegui_run.process_pool = None

    nicegui_run.setup = safe_setup


tolerate_missing_process_pool()

LABELS = {
    "ru": {
        "workspace": "Рабочие папки",
        "operations": "Операции",
        "maintenance": "Обслуживание",
        "status": "Статус",
        "log": "Журнал операции",
        "idle": "Ожидание",
        "running": "Выполняется",
        "done": "Готово",
        "error": "Ошибка",
        "cancel": "Отменить",
        "another_running": "Другая операция уже выполняется.",
        "confirm_title": "Подтвердите действие",
        "confirm_note": "Действие может изменить управляемую рабочую область.",
        "run": "Запустить",
        "back": "Назад",
        "selected_operation": "Выбрана команда",
        "open_menu": "Открыть",
        "parameters": "Параметры",
        "close": "Закрыть",
        "logs": "LOGS",
        "report": "REPORT",
        "config": "CONFIG",
        "rules": "ПРАВИЛА",
        "expand": "Развернуть",
        "clear_terminal_window": "Очистить окно терминала",
        "add_file_short": "Добавить файл...",
        "file_list": "File List",
        "file_list_button": "Список",
        "file_list_empty": "INPUT has no files.",
        "file_list_missing": "INPUT was not found: {path}",
        "file_list_ready": "File list generated: {count}.",
        "terminal_file": "File",
        "source_folder_missing": "Источник не найден: {path}",
        "picker_cancelled": "Выбор отменен.",
        "choose_file": "Выбрать файл",
        "operation_done": "Операция завершена.",
        "operation_failed": "Операция завершилась с кодом {code}.",
        "source_folder": "Источник",
        "target_folder": "Назначение",
        "clear_io_short": "Сбросить",
        "delete_io_short": "Удалить",
        "source_selected": "Источник выбран.",
        "target_selected": "Назначение выбрано.",
        "path_required": "Выберите путь.",
        "path_pinned": "Путь закреплен.",
        "path_unpinned": "Закрепление снято.",
        "select_required": "Выберите хотя бы один пункт: {field}",
        "theme": "Тема",
        "theme_saved": "Тема сохранена. Перезагружаю интерфейс.",
        "lang_switch": "EN",
        "resize_panels": "Изменить ширину панелей",
    },
    "en": {
        "workspace": "Workspace folders",
        "operations": "Operations",
        "maintenance": "Maintenance",
        "status": "Status",
        "log": "Operation log",
        "idle": "Idle",
        "running": "Running",
        "done": "Done",
        "error": "Error",
        "cancel": "Cancel",
        "another_running": "Another operation is already running.",
        "confirm_title": "Confirm action",
        "confirm_note": "This action may change the managed workspace.",
        "run": "Run",
        "back": "Back",
        "selected_operation": "Selected command",
        "open_menu": "Open",
        "parameters": "Parameters",
        "close": "Close",
        "logs": "Logs",
        "report": "Report",
        "config": "CONFIG",
        "rules": "RULES",
        "expand": "Expand",
        "clear_terminal_window": "Clear terminal window",
        "add_file_short": "Add file...",
        "file_list": "File List",
        "file_list_button": "List",
        "file_list_empty": "INPUT has no files.",
        "file_list_missing": "INPUT was not found: {path}",
        "file_list_ready": "File list generated: {count}.",
        "terminal_file": "File",
        "source_folder_missing": "Source was not found: {path}",
        "picker_cancelled": "Selection cancelled.",
        "choose_file": "Choose file",
        "operation_done": "Operation finished.",
        "operation_failed": "Operation finished with exit code {code}.",
        "source_folder": "Source",
        "target_folder": "Target",
        "clear_io_short": "Reset",
        "delete_io_short": "Delete",
        "source_selected": "Source selected.",
        "target_selected": "Target selected.",
        "path_required": "Choose a path.",
        "path_pinned": "Path pinned.",
        "path_unpinned": "Path unpinned.",
        "select_required": "Select at least one item: {field}",
        "theme": "Theme",
        "theme_saved": "Theme saved. Reloading UI.",
        "lang_switch": "RU",
        "resize_panels": "Resize panels",
    },
}

PICKER_BOOTSTRAP = r"""
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
try {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class AudionDpiAwareness {
  [DllImport("user32.dll")]
  public static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);
  [DllImport("shcore.dll")]
  public static extern int SetProcessDpiAwareness(int value);
}
"@
  try { [AudionDpiAwareness]::SetProcessDpiAwarenessContext([IntPtr](-4)) | Out-Null }
  catch { [AudionDpiAwareness]::SetProcessDpiAwareness(2) | Out-Null }
} catch {}
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
"""

state: dict[str, Any] = {
    "running": False,
    "cancel": False,
    "progress": 0.0,
    "status": "",
    "lines": [],
    "line_html": [],
    "line_seq": 0,
    "line_base_seq": 1,
    "log_version": 0,
    "terminal_scroll_top_seq": 0,
    "terminal_epoch": 0,
    "ansi_renderer": StatefulAnsiHtmlRenderer(),
    "exit_code": None,
    "command_path": [],
    "pending_command": None,
    "field_values": {},
    "processor_tab": "hygiene",
    "processor_modes": {"hygiene": "scan", "anomalies": "scan", "styles": "scan"},
    "source_path": str(getattr(settings, "source_path", "") or ""),
    "destination_path": str(getattr(settings, "destination_path", "") or ""),
    "workspace_feedback": {},
}


def tr(key: str, **kwargs: Any) -> str:
    lang = settings.language if settings.language in LABELS else "en"
    text = LABELS.get(lang, LABELS["en"]).get(key, key)
    return text.format(**kwargs) if kwargs else text


def em(key: str) -> str:
    if not bool(getattr(settings, "emoji", False)):
        return ""
    return {
        "workspace": "📁 ",
        "operations": "⚙ ",
        "maintenance": "🧰 ",
        "status": "● ",
        "log": "🖥 ",
    }.get(key, "")


def app_title() -> str:
    title = str(ui_info.get("title") or tool_info.get("name") or "Audion GUI Tool")
    return title[:-3] if title.endswith(" UI") else title


def normalize_theme_id(theme_id: Any) -> str:
    text = str(theme_id or DEFAULT_THEME_ID).strip().lower()
    return THEME_ALIASES.get(text, text)


def active_theme() -> str:
    theme_id = normalize_theme_id(settings.theme)
    themes = ui_colors["themes"]
    if theme_id in themes:
        return theme_id
    return DEFAULT_THEME_ID if DEFAULT_THEME_ID in themes else next(iter(themes))


def active_theme_data() -> dict[str, Any]:
    return dict(ui_colors["themes"][active_theme()])


def active_theme_mode() -> str:
    return str(active_theme_data().get("mode", "dark"))


def theme_label(theme_id: str) -> str:
    theme_data = ui_colors["themes"].get(theme_id, {})
    label_key = "label_ru" if settings.language == "ru" else "label"
    return str(theme_data.get(label_key) or theme_data.get("label") or theme_id)


def theme_options() -> dict[str, str]:
    return {theme_id: theme_label(theme_id) for theme_id in ui_colors["themes"]}


def set_theme(theme_id: Any) -> None:
    selected = normalize_theme_id(theme_id)
    if selected not in ui_colors["themes"]:
        return
    settings.theme = selected
    save_app_settings()
    safe_notify(tr("theme_saved"), "positive")
    reload_ui()


def theme_change_handler(event: Any) -> None:
    set_theme(getattr(event, "value", None))


def reload_ui() -> None:
    ui.run_javascript(
        """
        (() => {
          try {
            if ('scrollRestoration' in window.history) {
              window.history.scrollRestoration = 'manual';
            }
            window.sessionStorage.setItem('audion_force_scroll_top', '1');
            window.scrollTo(0, 0);
            document.documentElement.scrollTop = 0;
            document.body.scrollTop = 0;
          } catch (error) {}
          window.location.reload();
        })();
        """
    )


def toggle_language() -> None:
    settings.language = "en" if settings.language == "ru" else "ru"
    save_app_settings()
    reload_ui()


def theme_variables() -> dict[str, str]:
    variables: dict[str, str] = {}
    for ramp_name, stops in ui_colors["ramps"].items():
        if not isinstance(stops, dict):
            continue
        for stop, color in stops.items():
            variables[f"color-{ramp_name}-{stop}"] = str(color).strip()
    variables.update(ui_colors["tokens"])
    variables.update(_string_map(active_theme_data().get("tokens", {})))
    variables.setdefault("color-background-primary", "#141413")
    variables.setdefault("color-background-secondary", "#1f1e1a")
    variables.setdefault("color-background-tertiary", "#0f0f0e")
    variables.setdefault("color-text-primary", "#faf9f5")
    variables.setdefault("color-text-secondary", "#e8e6dc")
    variables.setdefault("color-text-tertiary", "#b0aea5")
    variables.setdefault("color-border-tertiary", "rgba(250, 249, 245, 0.15)")
    variables.setdefault("color-border-secondary", "rgba(250, 249, 245, 0.3)")
    variables.setdefault("color-border-primary", "rgba(250, 249, 245, 0.4)")
    variables.setdefault("color-background-warning", "#412402")
    variables.setdefault("color-text-warning", "#FAC775")
    variables.setdefault("color-border-warning", "#BA7517")
    variables.setdefault("color-accent-primary", "#d97757")
    variables.setdefault("font-sans", "Inter, Segoe UI, Arial, sans-serif")
    variables.setdefault("font-mono", "Cascadia Mono, Consolas, monospace")
    variables.setdefault("border-radius-md", "8px")
    variables.setdefault("border-radius-lg", "12px")
    return variables


def reset_terminal_log() -> None:
    state["lines"] = []
    state["line_html"] = []
    state["line_seq"] = 0
    state["line_base_seq"] = 1
    state["terminal_epoch"] = int(state["terminal_epoch"]) + 1
    state["ansi_renderer"] = StatefulAnsiHtmlRenderer()
    state["log_version"] = int(state["log_version"]) + 1


def clear_terminal_log() -> None:
    reset_terminal_log()


def add_log(message: str) -> None:
    line = str(message).rstrip("\r\n")
    renderer = state.get("ansi_renderer")
    if not isinstance(renderer, StatefulAnsiHtmlRenderer):
        renderer = StatefulAnsiHtmlRenderer()
        state["ansi_renderer"] = renderer
    safe_line_html = renderer.render(line, final=True)
    state["line_seq"] = int(state["line_seq"]) + 1
    state["lines"].append(line)
    state["line_html"].append(safe_line_html)
    if len(state["lines"]) > TERMINAL_HISTORY_LIMIT:
        overflow = len(state["lines"]) - TERMINAL_HISTORY_LIMIT
        del state["lines"][:overflow]
        del state["line_html"][:overflow]
    state["line_base_seq"] = int(state["line_seq"]) - len(state["lines"]) + 1
    state["log_version"] = int(state["log_version"]) + 1


def terminal_cached_lines_html(lines: list[str]) -> str:
    return "".join(f'<span class="audion-terminal-line">{line}</span>' for line in lines)


def terminal_html(lines: list[str] | None = None) -> str:
    if lines is not None:
        return render_terminal_html(lines)
    return f'<pre class="audion-terminal-pre">{terminal_cached_lines_html(state["line_html"])}</pre>'


def progress_text() -> str:
    return f"{round(max(0.0, min(1.0, float(state['progress']))) * 100):.0f}%"


def safe_notify(message: str, kind: str = "info", **notify_kwargs: Any) -> None:
    notify_type = str(notify_kwargs.pop("type", kind))
    options = {"message": str(message), "type": notify_type, **notify_kwargs}
    delivered = False
    for client in list(nicegui_app.clients()):
        if getattr(client, "_deleted", False) or not client.has_socket_connection:
            continue
        try:
            client.outbox.enqueue_message("notify", options, client.id)
            delivered = True
        except Exception as exc:
            logging.warning("NiceGUI notification delivery failed for client %s: %s", getattr(client, "id", "?"), exc)
    if delivered:
        return

    try:
        ui.notify(message, type=notify_type, **notify_kwargs)
    except RuntimeError as exc:
        message_text = str(exc)
        if "slot belongs to has been deleted" not in message_text and "current slot cannot be determined" not in message_text:
            raise
        logging.warning("NiceGUI notification skipped because no live client slot was available: %s", message)


RUN_STATE_LABELS = {
    "idle": ("idle", "audion-status-idle"),
    "running": ("running", "audion-status-running"),
    "done": ("done", "audion-status-done"),
    "error": ("error", "audion-status-error"),
}


def run_state() -> str:
    """Which of the four states the panel is showing.

    Colour carries this everywhere it appears, so it is decided once.
    """
    if bool(state["running"]):
        return "running"
    exit_code = state.get("exit_code")
    if exit_code is None:
        return "idle"
    return "done" if int(exit_code or 0) == 0 else "error"


def status_row_classes() -> str:
    return f"audion-status-row {RUN_STATE_LABELS[run_state()][1]}"


def status_state_text() -> str:
    return tr(RUN_STATE_LABELS[run_state()][0]).upper()


def elapsed_text(seconds: float | None) -> str:
    """A run's own clock, mm:ss, or an em dash before anything has run.

    The start is noticed by the refresh timer rather than written by the code that
    starts a run: there are several such places, and none of them has to know
    about the panel.
    """
    if seconds is None:
        return "—"
    total = max(0, int(seconds))
    return f"{total // 60:02d}:{total % 60:02d}"


def status_dot_classes() -> str:
    base = "audion-status-dot text-lg leading-none"
    if bool(state["running"]):
        return f"{base} text-sky-400 animate-pulse"
    if state.get("exit_code") is None:
        return f"{base} text-gray-500"
    if int(state.get("exit_code") or 0) == 0:
        return f"{base} text-green-400"
    return f"{base} text-red-400"


def set_progress(value: float) -> None:
    state["progress"] = max(0.0, min(1.0, float(value)))


def cancel_requested() -> bool:
    return bool(state["cancel"])


def hidden_subprocess_flags() -> int:
    if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        return int(subprocess.CREATE_NO_WINDOW)
    return 0


def hidden_subprocess_startupinfo() -> subprocess.STARTUPINFO | None:
    if os.name != "nt" or not hasattr(subprocess, "STARTUPINFO"):
        return None
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    return startupinfo


def hidden_subprocess_kwargs() -> dict[str, Any]:
    return {
        "creationflags": hidden_subprocess_flags(),
        "startupinfo": hidden_subprocess_startupinfo(),
    }


def resolve_dialog_powershell() -> list[str]:
    candidates = [
        [str(paths.system_core / "powershell" / "pwsh.exe"), "-NoLogo", "-NoProfile", "-STA", "-Command"],
        ["pwsh.exe", "-NoLogo", "-NoProfile", "-STA", "-Command"],
        ["powershell.exe", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command"],
    ]
    for candidate in candidates:
        exe = candidate[0]
        if Path(exe).exists() or shutil.which(exe):
            return candidate
    raise RuntimeError("PowerShell was not found for Windows picker.")


def parse_picker_paths(text: str) -> list[Path]:
    import json

    payload = text.strip()
    if not payload:
        return []
    data = json.loads(payload)
    if isinstance(data, str):
        data = [data]
    return [Path(str(item)).resolve() for item in data if str(item).strip()]


_PICKER_RUN_LOCK = threading.Lock()
_PICKER_JOB_LOCK = threading.Lock()
_PICKER_SHUTDOWN = threading.Event()
_PICKER_JOB_HANDLE: int | None = None


class _JobObjectBasicLimitInformation(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", ctypes.c_int64),
        ("PerJobUserTimeLimit", ctypes.c_int64),
        ("LimitFlags", wintypes.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wintypes.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wintypes.DWORD),
        ("SchedulingClass", wintypes.DWORD),
    ]


class _IoCounters(ctypes.Structure):
    _fields_ = [
        ("ReadOperationCount", ctypes.c_uint64),
        ("WriteOperationCount", ctypes.c_uint64),
        ("OtherOperationCount", ctypes.c_uint64),
        ("ReadTransferCount", ctypes.c_uint64),
        ("WriteTransferCount", ctypes.c_uint64),
        ("OtherTransferCount", ctypes.c_uint64),
    ]


class _JobObjectExtendedLimitInformation(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", _JobObjectBasicLimitInformation),
        ("IoInfo", _IoCounters),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryUsed", ctypes.c_size_t),
        ("PeakJobMemoryUsed", ctypes.c_size_t),
    ]


def close_picker_job() -> None:
    global _PICKER_JOB_HANDLE
    _PICKER_SHUTDOWN.set()
    with _PICKER_JOB_LOCK:
        handle = _PICKER_JOB_HANDLE
        _PICKER_JOB_HANDLE = None
    if os.name == "nt" and handle:
        ctypes.WinDLL("kernel32", use_last_error=True).CloseHandle(wintypes.HANDLE(handle))


def _picker_job_handle() -> int | None:
    global _PICKER_JOB_HANDLE
    if os.name != "nt" or _PICKER_SHUTDOWN.is_set():
        return None
    with _PICKER_JOB_LOCK:
        if _PICKER_SHUTDOWN.is_set():
            return None
        if _PICKER_JOB_HANDLE:
            return _PICKER_JOB_HANDLE
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateJobObjectW.restype = wintypes.HANDLE
        job = kernel32.CreateJobObjectW(None, None)
        if not job:
            logging.warning("Could not create the Windows picker job: %s", ctypes.get_last_error())
            return None
        info = _JobObjectExtendedLimitInformation()
        info.BasicLimitInformation.LimitFlags = 0x00002000
        if not kernel32.SetInformationJobObject(wintypes.HANDLE(job), 9, ctypes.byref(info), ctypes.sizeof(info)):
            error = ctypes.get_last_error()
            kernel32.CloseHandle(wintypes.HANDLE(job))
            logging.warning("Could not configure the Windows picker job: %s", error)
            return None
        _PICKER_JOB_HANDLE = int(job)
        return _PICKER_JOB_HANDLE


def _assign_picker_to_job(process: subprocess.Popen[str]) -> None:
    handle = _picker_job_handle()
    if os.name != "nt" or not handle:
        if _PICKER_SHUTDOWN.is_set() and process.poll() is None:
            process.kill()
        return
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    if not kernel32.AssignProcessToJobObject(
        wintypes.HANDLE(handle),
        wintypes.HANDLE(int(process._handle)),  # type: ignore[attr-defined]
    ):
        logging.warning("Could not attach picker PID %s to its Windows job: %s", process.pid, ctypes.get_last_error())


def run_picker_script(script: str, error_message: str) -> list[Path]:
    if not _PICKER_RUN_LOCK.acquire(blocking=False):
        raise RuntimeError("A Windows picker is already open.")
    process: subprocess.Popen[str] | None = None
    try:
        if _PICKER_SHUTDOWN.is_set():
            raise RuntimeError("Windows picker supervisor is shutting down.")
        _picker_job_handle()
        process = subprocess.Popen(
            [*resolve_dialog_powershell(), script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            **hidden_subprocess_kwargs(),
        )
        _assign_picker_to_job(process)
        if _PICKER_SHUTDOWN.is_set():
            if process.poll() is None:
                process.kill()
            raise RuntimeError("Windows picker supervisor is shutting down.")
        try:
            stdout, stderr = process.communicate(timeout=3600)
        except subprocess.TimeoutExpired as exc:
            process.kill()
            process.communicate()
            raise RuntimeError("Windows picker timed out.") from exc
        if process.returncode != 0:
            raise RuntimeError(stderr.strip() or error_message)
        return parse_picker_paths(stdout)
    finally:
        if process is not None and process.poll() is None:
            process.kill()
        _PICKER_RUN_LOCK.release()


atexit.register(close_picker_job)
nicegui_app.on_shutdown(close_picker_job)


def ps_single_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def field_file_extensions(field: dict[str, Any]) -> list[str]:
    raw = field.get("extensions") or field.get("extension") or field.get("file_extensions") or ""
    if isinstance(raw, str):
        parts = raw.replace(";", ",").split(",")
    elif isinstance(raw, list):
        parts = [str(item) for item in raw]
    else:
        parts = []

    extensions: list[str] = []
    for part in parts:
        value = str(part).strip().lower()
        if not value or value == "*.*":
            continue
        value = value.lstrip("*")
        if not value.startswith("."):
            value = f".{value}"
        if value not in extensions:
            extensions.append(value)
    return extensions


def field_file_filter(field: dict[str, Any]) -> str:
    extensions = field_file_extensions(field)
    if not extensions:
        return "All files|*.*"
    patterns = ";".join(f"*{extension}" for extension in extensions)
    label = "/".join(extension.lstrip(".").upper() for extension in extensions)
    return f"{label} files|{patterns}|All files|*.*"


def pick_single_file(field: dict[str, Any] | None = None) -> list[Path]:
    field = field or {}
    title = field_label(field) or tr("choose_file")
    source = current_source_path().resolve(strict=False)
    input_dir = source.parent if source.is_file() else source
    script = PICKER_BOOTSTRAP + f"""
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = {ps_single_quote(title)}
$dialog.InitialDirectory = {ps_single_quote(str(input_dir))}
$dialog.Multiselect = $false
$dialog.Filter = {ps_single_quote(field_file_filter(field))}
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
  @($dialog.FileName) | ConvertTo-Json -Compress
}}
"""
    return run_picker_script(script, "File picker failed.")


def pick_folder() -> list[Path]:
    script = PICKER_BOOTSTRAP + r"""
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = 'Select source or target folder'
$dialog.ShowNewFolderButton = $false
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  @($dialog.SelectedPath) | ConvertTo-Json -Compress
}
"""
    return run_picker_script(script, "Folder picker failed.")


def select_field_file_value(field: dict[str, Any]) -> str:
    sources = pick_single_file(field)
    if not sources:
        return ""
    return str(sources[0].resolve())


async def pick_field_file(key: str, field: dict[str, Any], control: Any) -> None:
    if state["running"]:
        safe_notify(tr("another_running"), "warning")
        return
    try:
        value = await run.io_bound(select_field_file_value, field)
    except Exception as exc:
        add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
        safe_notify(str(exc), "negative")
        return
    if not value:
        safe_notify(tr("picker_cancelled"), "info")
        return
    set_field_value(key, value)
    control.set_value(value)
    safe_notify(tr("source_selected"), "positive")


def input_file_list_lines(source: Path) -> list[str]:
    if not source.exists():
        return [tr("file_list_missing", path=source)]
    if source.is_file():
        return ["No.  Size  List", f"001  {source.stat().st_size}  {source.name}"]
    if not source.is_dir():
        return [f"Unsupported source path: {source}"]

    names = sorted((path.name for path in source.rglob("*") if path.is_file()), key=lambda item: item.casefold())
    if not names:
        return [tr("file_list_empty")]

    number_width = max(3, len(str(len(names))))
    lines = [
        f"{'No.':>{number_width}}  List",
        f"{'-' * number_width}  ----",
    ]
    lines.extend(f"{index:0{number_width}d}. {name}" for index, name in enumerate(names, start=1))
    return lines


async def show_input_file_list() -> None:
    if state["running"]:
        safe_notify(tr("another_running"), "warning")
        return

    reset_terminal_log()
    title = tr("file_list")
    state.update(
        {
            "running": True,
            "cancel": False,
            "progress": 0.02,
            "status": f"{tr('running')}: {title}",
            "exit_code": None,
        }
    )
    try:
        lines = await run.io_bound(input_file_list_lines, current_source_path())
        for line in lines:
            add_log(line)
        count = max(0, len(lines) - 2)
        state["terminal_scroll_top_seq"] = int(state["log_version"])
        state["exit_code"] = 0
        state["progress"] = 1.0
        state["status"] = f"{tr('done')}: {title} [{count}]"
        safe_notify(tr("file_list_ready", count=count), "positive")
    except Exception as exc:
        state["exit_code"] = 1
        state["progress"] = max(float(state["progress"]), 0.98)
        state["status"] = f"{tr('error')}: {exc}"
        add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
        safe_notify(str(exc), "negative")
    finally:
        state["running"] = False


def current_source_path() -> Path:
    return Path(str(state.get("source_path") or getattr(settings, "source_path", "") or paths.input)).expanduser()


def current_target_path() -> Path:
    return Path(str(state.get("destination_path") or getattr(settings, "destination_path", "") or paths.output)).expanduser()


def active_project_paths():
    source = current_source_path().resolve(strict=False)
    target = current_target_path().resolve(strict=False)
    return replace(paths, input=source, output=target)


async def start_operation(operation: Operation) -> None:
    if state["running"]:
        safe_notify(tr("another_running"), "warning")
        return

    if operation.kind == "dangerous":
        with ui.dialog() as dialog, ui.card().classes("audion-dialog audion-confirm-card rounded-lg"):
            ui.label(tr("confirm_title")).classes("text-base font-semibold")
            ui.label(operation.display_title(settings.language)).classes("text-sm font-semibold")
            description = operation.display_description(settings.language)
            if description:
                ui.label(description).classes("text-sm text-gray-400")
            ui.label(tr("confirm_note")).classes("audion-confirm-warning")
            with ui.row().classes("w-full items-center justify-end gap-2"):
                ui.button(tr("cancel"), on_click=dialog.close).props("dense flat").classes("audion-action rounded-lg")
                ui.button(tr("run"), on_click=lambda: dialog.submit(True)).props("dense flat no-wrap").classes("audion-action rounded-lg")
        if not await dialog:
            return

    source = current_source_path().resolve(strict=False)
    if not source.exists():
        safe_notify(tr("source_folder_missing", path=source), "negative")
        return

    reset_terminal_log()
    state.update(
        {
            "running": True,
            "cancel": False,
            "progress": 0.02,
            "status": f"{tr('running')}: {operation.display_title(settings.language)}",
            "exit_code": None,
        }
    )
    started = time.perf_counter()
    try:
        result = await run.io_bound(
            execute_operation,
            active_project_paths(),
            operation,
            add_log,
            set_progress,
            cancel_requested,
        )
        elapsed = time.perf_counter() - started
        state["exit_code"] = 0 if result.ok else 1
        state["progress"] = 1.0 if result.ok else max(float(state["progress"]), 0.98)
        state["status"] = (
            f"{tr('done') if result.ok else tr('error')}: "
            f"{operation.display_title(settings.language)} [{state['exit_code']}] {elapsed:.1f}s"
        )
        safe_notify(result.message, "positive" if result.ok else "negative")
    except Exception as exc:
        state["exit_code"] = 1
        state["progress"] = max(float(state["progress"]), 0.98)
        state["status"] = f"{tr('error')}: {exc}"
        add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
        safe_notify(str(exc), "negative")
    finally:
        state["running"] = False


def save_workspace_path(kind: str, value: Any) -> None:
    text = str(value or "").strip()
    if kind == "source":
        settings.source_path = text
        state["source_path"] = text
    else:
        settings.destination_path = text
        state["destination_path"] = text
    save_app_settings()


def open_workspace_folder(role: str) -> None:
    folder = current_target_path() if role == "target" else current_source_path()
    if role != "target" and not folder.exists():
        raise FileNotFoundError(tr("source_folder_missing", path=folder))
    if folder.is_file():
        if os.name == "nt":
            subprocess.Popen(["explorer.exe", f"/select,{folder}"], **hidden_subprocess_kwargs())
        else:
            open_folder(folder.parent)
        return
    if role == "target":
        folder.mkdir(parents=True, exist_ok=True)
    open_folder(folder)


def mark_workspace_feedback(role: str, action: str) -> None:
    state["workspace_feedback"] = {"role": canonical_role(role), "action": str(action or "path")}


def _save_workbench_path(role: WorkbenchRole, value: Any) -> None:
    save_workspace_path("target" if role == "target" else "source", value)


def normalized_absolute_path(path_value: Any) -> Path:
    candidate = Path(str(path_value or "")).expanduser()
    if not candidate.is_absolute():
        candidate = ROOT / candidate
    return candidate.resolve(strict=False)


def paths_equal(left: Any, right: Any) -> bool:
    return os.path.normcase(str(normalized_absolute_path(left))) == os.path.normcase(str(normalized_absolute_path(right)))


def remove_path_tree(path: Path) -> int:
    is_junction = bool(getattr(os.path, "isjunction", lambda _path: False)(path))
    if path.is_symlink() or is_junction:
        path.rmdir() if path.is_dir() else path.unlink()
        return 1
    if path.is_file():
        path.unlink()
        return 1
    if path.is_dir():
        shutil.rmtree(path)
        return 1
    return 0


def validate_workspace_delete_target(path_value: Any) -> Path:
    target = normalized_absolute_path(path_value)
    if target.parent == target:
        raise RuntimeError(f"Refusing to delete a filesystem root: {target}")
    if paths_equal(target, ROOT):
        raise RuntimeError(f"Refusing to delete the project root: {target}")
    return target


def delete_workspace_path_contents(path_value: Any) -> dict[str, Any]:
    target = validate_workspace_delete_target(path_value)
    if not target.exists() and not target.is_symlink():
        return {"path": str(target), "kind": "missing", "removed": 0}
    is_junction = bool(getattr(os.path, "isjunction", lambda _path: False)(target))
    if target.is_file() or target.is_symlink() or is_junction:
        return {"path": str(target), "kind": "file", "removed": remove_path_tree(target)}
    if not target.is_dir():
        raise RuntimeError(f"Unsupported workspace path: {target}")
    removed = 0
    for child in sorted(target.iterdir(), key=lambda item: item.name.casefold()):
        # .gitkeep is not spared: input and output must be genuinely empty after
        # a clear, so nobody has to wonder what the leftover file is or whether it
        # is safe to delete. The folders come from install/init_folders.cmd.
        removed += remove_path_tree(child)
    return {"path": str(target), "kind": "folder", "removed": removed}


def delete_workspace_io_contents(source: Path, target: Path) -> dict[str, Any]:
    source_result = delete_workspace_path_contents(source)
    target_result = (
        {"path": str(normalized_absolute_path(target)), "kind": "same", "removed": 0}
        if paths_equal(source, target)
        else delete_workspace_path_contents(target)
    )
    return {"source": source_result, "target": target_result}


WORKBENCH_CONFIG = WorkbenchConfig(
    root=ROOT,
    input_path=paths.input,
    output_path=paths.output,
    history_path=paths.config / "path_history.json",
    history_limit=PATH_HISTORY_LIMIT,
)
WORKBENCH_ADAPTER = WorkbenchAdapter(
    config=WORKBENCH_CONFIG,
    current_path_callback=lambda role: current_target_path() if role == "target" else current_source_path(),
    save_path_callback=_save_workbench_path,
    language_callback=lambda: str(settings.language),
    translate_callback=tr,
    log_callback=add_log,
    notify_callback=safe_notify,
    reload_callback=lambda _delay=0: reload_ui(),
    busy_callback=lambda: bool(state.get("running")),
    feedback_callback=lambda: dict(state.get("workspace_feedback") or {}),
    set_feedback_callback=mark_workspace_feedback,
    clear_feedback_callback=lambda: state.update({"workspace_feedback": {}}),
)
WORKBENCH_ADAPTER.validate()
WORKBENCH_ADAPTER.ensure_initial_history()
if not WORKBENCH_ADAPTER.history_entries("source"):
    WORKBENCH_ADAPTER.remember_path("source", str(paths.input))
if not WORKBENCH_ADAPTER.history_entries("target"):
    WORKBENCH_ADAPTER.remember_path("target", str(paths.output))


def canonical_workspace_pin_click_handler(role: str, pinned: bool):
    async def handler() -> None:
        selected = current_target_path() if role == "target" else current_source_path()
        try:
            await run.io_bound(WORKBENCH_ADAPTER.set_path_pinned, role, str(selected), pinned)
            mark_workspace_feedback(role, "pin" if pinned else "unpin")
            add_log(f"{'Pinned' if pinned else 'Unpinned'} {role} path: {selected}")
            reload_ui()
        except Exception as exc:
            add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
            safe_notify(str(exc), "negative")

    return handler


def _confirm_external_source_delete(selected: Path):
    with ui.dialog() as dialog, ui.card().classes("rounded-lg"):
        title = "Удалить исходный файл?" if selected.is_file() else "Очистить внешний ИСТОЧНИК?"
        warning = "Будет удалён исходный файл. Другой копии может не существовать." if selected.is_file() else "Будут безвозвратно удалены все файлы и вложенные папки."
        if settings.language != "ru":
            title = "Delete the source file?" if selected.is_file() else "Clear the external SOURCE?"
            warning = "The source file will be deleted. Another copy may not exist." if selected.is_file() else "All files and nested folders will be permanently deleted."
        ui.label(title).classes("text-base font-semibold")
        ui.label(warning).classes("text-sm text-gray-300")
        ui.label(str(normalized_absolute_path(selected))).classes("max-w-3xl break-all font-mono text-xs text-gray-400")
        with ui.row().classes("gap-2"):
            ui.button(tr("cancel"), on_click=dialog.close).props("dense flat")
            ui.button(tr("delete_io_short"), on_click=lambda: dialog.submit(True)).props("dense color=negative")
    return dialog


def canonical_workspace_delete_path_click_handler(role: str):
    async def handler() -> None:
        if state["running"]:
            safe_notify(tr("another_running"), "warning")
            return
        selected = current_target_path() if role == "target" else current_source_path()
        if role == "source" and not paths_equal(selected, paths.input):
            if not await _confirm_external_source_delete(selected):
                return
        try:
            result = await run.io_bound(delete_workspace_path_contents, selected)
            if result.get("kind") == "file":
                await run.io_bound(WORKBENCH_ADAPTER.delete_path_history, role, str(selected))
                save_workspace_path("target" if role == "target" else "source", "")
            mark_workspace_feedback(role, "delete")
            add_log(f"Cleared {role.upper()}: {result.get('path')} [kind={result.get('kind')}, removed={result.get('removed', 0)}]")
            reload_ui()
        except Exception as exc:
            add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
            safe_notify(str(exc), "negative")

    return handler


def canonical_workspace_path_select_handler(role: str):
    async def handler(event: Any) -> None:
        path_value = str(getattr(event, "value", "") or "").strip()
        if not path_value:
            return
        save_workspace_path("target" if role == "target" else "source", path_value)
        await run.io_bound(WORKBENCH_ADAPTER.remember_path, role, path_value)
        mark_workspace_feedback(role, "path")
        add_log(f"{'TARGET' if role == 'target' else 'SOURCE'} -> {path_value}")
        reload_ui()

    return handler


def canonical_workspace_pick_click_handler(role: str):
    async def handler() -> None:
        if state["running"]:
            safe_notify(tr("another_running"), "warning")
            return
        try:
            selected = await run.io_bound(pick_folder)
        except Exception as exc:
            add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
            safe_notify(str(exc), "negative")
            return
        if not selected:
            add_log(tr("picker_cancelled"))
            return
        path_value = str(selected[0])
        save_workspace_path("target" if role == "target" else "source", path_value)
        await run.io_bound(WORKBENCH_ADAPTER.remember_path, role, path_value)
        mark_workspace_feedback(role, "path")
        reload_ui()

    return handler


def canonical_workspace_open_click_handler(role: str):
    async def handler() -> None:
        try:
            await run.io_bound(open_workspace_folder, role)
        except Exception as exc:
            add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
            safe_notify(str(exc), "negative")

    return handler


def canonical_workspace_single_file_click_handler():
    async def handler() -> None:
        if state["running"]:
            safe_notify(tr("another_running"), "warning")
            return
        try:
            selected = await run.io_bound(pick_single_file)
        except Exception as exc:
            add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
            safe_notify(str(exc), "negative")
            return
        if not selected:
            add_log(tr("picker_cancelled"))
            return
        path_value = str(selected[0])
        save_workspace_path("source", path_value)
        await run.io_bound(WORKBENCH_ADAPTER.remember_path, "source", path_value)
        mark_workspace_feedback("source", "path")
        reload_ui()

    return handler


def canonical_reset_workspace_paths_click_handler():
    async def handler() -> None:
        result = await run.io_bound(WORKBENCH_ADAPTER.clear_path_history_cache_keep_pins)
        save_workspace_path("source", "")
        save_workspace_path("target", "")
        add_log(f"Workspace reset: SOURCE -> {paths.input}; TARGET -> {paths.output}; pins kept={result.get('kept_pins', 0)}")
        safe_notify(tr("operation_done"), "positive")
        reload_ui()

    return handler


def canonical_workspace_delete_both_click_handler():
    async def handler() -> None:
        if state["running"]:
            safe_notify(tr("another_running"), "warning")
            return
        source, target = current_source_path(), current_target_path()
        with ui.dialog() as dialog, ui.card().classes("rounded-lg"):
            ui.label("Удалить содержимое I/O?" if settings.language == "ru" else "Delete I/O contents?").classes("text-base font-semibold")
            ui.label(f"SOURCE: {normalized_absolute_path(source)}").classes("max-w-3xl break-all font-mono text-xs text-gray-400")
            ui.label(f"TARGET: {normalized_absolute_path(target)}").classes("max-w-3xl break-all font-mono text-xs text-gray-400")
            with ui.row().classes("gap-2"):
                ui.button(tr("cancel"), on_click=dialog.close).props("dense flat")
                ui.button(tr("delete_io_short"), on_click=lambda: dialog.submit(True)).props("dense color=negative")
        if not await dialog:
            return
        state["running"] = True
        try:
            result = await run.io_bound(delete_workspace_io_contents, source, target)
            for role, selected in (("source", source), ("target", target)):
                if result[role].get("kind") == "file":
                    await run.io_bound(WORKBENCH_ADAPTER.delete_path_history, role, str(selected))
                    save_workspace_path(role, "")
            add_log(f"Cleared SOURCE: {result['source']}")
            add_log(f"Cleared TARGET: {result['target']}")
            mark_workspace_feedback("source", "delete")
            reload_ui()
        except Exception as exc:
            add_log(f"ERROR: {exc.__class__.__name__}: {exc}")
            safe_notify(str(exc), "negative")
        finally:
            state["running"] = False

    return handler


WORKBENCH_RENDERER = WorkbenchRenderer(
    adapter=WORKBENCH_ADAPTER,
    handlers=WorkbenchHandlers(
        delete_path=canonical_workspace_delete_path_click_handler,
        pin_path=canonical_workspace_pin_click_handler,
        select_path=canonical_workspace_path_select_handler,
        pick_path=canonical_workspace_pick_click_handler,
        open_path=canonical_workspace_open_click_handler,
        add_file=canonical_workspace_single_file_click_handler,
        reset_paths=canonical_reset_workspace_paths_click_handler,
        delete_io=canonical_workspace_delete_both_click_handler,
        list_files=show_input_file_list,
    ),
    display_path_callback=display_path,
)


def operation_button(operation: Operation) -> None:
    tooltip = command_tooltip(operation)
    with ui.element("div").classes("audion-operation-row"):
        button = ui.button(
            operation.display_title(settings.language),
            on_click=operation_click_handler(operation),
        ).props("dense flat no-wrap").classes("audion-action audion-operation-button rounded-lg")
        attach_tooltip(button, tooltip)
        attach_tooltip(ui.label(operation.display_description(settings.language)).classes("audion-operation-description"), tooltip)


def operation_click_handler(operation: Operation):
    async def handler() -> None:
        await start_operation(operation)

    return handler


def operation_to_command_node(operation: Operation) -> CommandNode:
    return CommandNode(
        id=operation.id,
        title=operation.title,
        description=operation.description,
        service=operation.service,
        kind=operation.kind,
        title_ru=operation.title_ru,
        description_ru=operation.description_ru,
        parameters=dict(operation.parameters),
        fields=operation.fields,
    )


def root_command_nodes() -> list[CommandNode]:
    if manifest.operation_groups:
        return manifest.operation_groups
    return [operation_to_command_node(operation) for operation in manifest.operations]


def command_node_by_id(node_id: str, nodes: list[CommandNode] | tuple[CommandNode, ...] | None = None) -> CommandNode | None:
    search_nodes = root_command_nodes() if nodes is None else list(nodes)
    for node in search_nodes:
        if node.id == node_id:
            return node
        found = command_node_by_id(node_id, node.children)
        if found is not None:
            return found
    return None


def child_command_node(parent: CommandNode, node_id: str) -> CommandNode | None:
    return next((child for child in parent.children if child.id == node_id), None)


def current_command_level() -> tuple[list[CommandNode], list[CommandNode]]:
    trail: list[CommandNode] = []
    nodes = root_command_nodes()
    for node_id in list(state.get("command_path", [])):
        node = next((candidate for candidate in nodes if candidate.id == node_id), None)
        if node is None:
            state["command_path"] = []
            state["pending_command"] = None
            return [], root_command_nodes()
        trail.append(node)
        nodes = list(node.children)
    return trail, nodes


def enter_command_node(node: CommandNode) -> None:
    state["pending_command"] = None
    state["command_path"] = [*state.get("command_path", []), node.id]
    command_tree.refresh()


def select_command_node(node: CommandNode) -> None:
    state["pending_command"] = node
    command_tree.refresh()


async def activate_command_node(node: CommandNode) -> None:
    if node.children:
        enter_command_node(node)
        return
    if is_direct_action_node(node):
        await run_command_node(node)
        return
    select_command_node(node)


def command_click_handler(node: CommandNode):
    async def handler() -> None:
        await activate_command_node(node)

    return handler


def go_back_command() -> None:
    if state.get("pending_command") is not None:
        state["pending_command"] = None
    else:
        path = list(state.get("command_path", []))
        if path:
            path.pop()
        state["command_path"] = path
    command_tree.refresh()


def field_id(field: dict[str, Any]) -> str:
    return str(field.get("id") or field.get("name") or "").strip()


def field_label(field: dict[str, Any]) -> str:
    language = settings.language
    if language == "ru" and field.get("label_ru"):
        return str(field["label_ru"])
    return str(field.get("label") or field.get("title") or field_id(field))


def field_hint(field: dict[str, Any]) -> str:
    language = settings.language
    if language == "ru" and field.get("hint_ru"):
        return str(field["hint_ru"])
    return str(field.get("hint") or "")


def clean_tooltip(text: object) -> str:
    return " ".join(str(text or "").split())


def attach_tooltip(element: Any, text: object) -> Any:
    tooltip = clean_tooltip(text)
    if tooltip:
        element.tooltip(tooltip)
    return element


def command_tooltip(node: CommandNode | Operation) -> str:
    return clean_tooltip(node.display_description(settings.language) or node.display_title(settings.language))


def field_tooltip(field: dict[str, Any]) -> str:
    hint = field_hint(field)
    if hint:
        return clean_tooltip(hint)
    label = field_label(field)
    if settings.language == "ru":
        return clean_tooltip(f"Параметр: {label}")
    return clean_tooltip(f"Parameter: {label}")


def render_field_hint(field: dict[str, Any], hint: str) -> None:
    if hint and bool(field.get("show_hint", False)):
        attach_tooltip(ui.label(hint).classes("audion-field-hint"), hint)


def field_default(field: dict[str, Any]) -> Any:
    if "default" in field:
        return field["default"]
    kind = str(field.get("type", field.get("kind", "text"))).lower()
    options = field.get("options", [])
    if kind in {"checkboxes", "multi_checkbox", "multicheckbox", "multi-select", "multiselect"}:
        if not isinstance(options, list):
            return []
        selected: list[Any] = []
        for option in options:
            if isinstance(option, dict) and option.get("default", False):
                selected.append(option.get("value", option.get("id", option.get("label"))))
        return selected
    if isinstance(options, list) and options:
        first = options[0]
        if isinstance(first, dict):
            return first.get("value", first.get("id", ""))
        return first
    return ""


def current_field_value(field: dict[str, Any]) -> Any:
    key = field_id(field)
    values = state.setdefault("field_values", {})
    if key not in values:
        values[key] = field_default(field)
    return values[key]


def set_field_value(key: str, value: Any) -> None:
    state.setdefault("field_values", {})[key] = value


def adjusted_number_value(field: dict[str, Any], current: Any, direction: int) -> int | float:
    step_raw = field.get("step", 1)
    try:
        step = float(step_raw)
    except (TypeError, ValueError):
        step = 1.0

    seed = current
    if seed is None or seed == "":
        seed = field_default(field) or 0
    try:
        value = float(seed)
    except (TypeError, ValueError):
        value = 0.0

    value += step * (1 if direction > 0 else -1)
    for bound_key, clamp in (("min", max), ("max", min)):
        bound = field.get(bound_key)
        if bound is None or bound == "":
            continue
        try:
            value = clamp(value, float(bound))
        except (TypeError, ValueError):
            continue

    kind = str(field.get("type", field.get("kind", "number"))).lower()
    integer_like = kind in {"number", "int", "integer"} and float(step).is_integer()
    return int(round(value)) if integer_like else round(value, 6)


def spin_number_field(key: str, field: dict[str, Any], control: Any, direction: int) -> None:
    value = adjusted_number_value(field, state.setdefault("field_values", {}).get(key), direction)
    set_field_value(key, value)
    control.set_value(value)


def select_options(field: dict[str, Any]) -> dict[Any, str] | list[Any]:
    options = field.get("options", [])
    if not isinstance(options, list):
        return []
    if all(isinstance(option, dict) for option in options):
        result: dict[Any, str] = {}
        for option in options:
            value = option.get("value", option.get("id", ""))
            if settings.language == "ru" and option.get("label_ru"):
                label = str(option["label_ru"])
            else:
                label = str(option.get("label") or option.get("title") or value)
            result[value] = label
        return result
    return options


def option_value(option: Any) -> Any:
    if isinstance(option, dict):
        return option.get("value", option.get("id", option.get("label", "")))
    return option


def option_label(option: Any) -> str:
    if not isinstance(option, dict):
        return str(option)
    language = settings.language
    if language == "ru" and option.get("label_ru"):
        return str(option["label_ru"])
    return str(option.get("label") or option.get("title") or option_value(option))


def checkbox_options(field: dict[str, Any]) -> list[tuple[Any, str]]:
    options = field.get("options", [])
    if not isinstance(options, list):
        return []
    return [(option_value(option), option_label(option)) for option in options]


def is_checkbox_group(field: dict[str, Any]) -> bool:
    kind = str(field.get("type", field.get("kind", "text"))).lower()
    return kind in {"checkboxes", "multi_checkbox", "multicheckbox", "multi-select", "multiselect"}


def is_boolean_field(field: dict[str, Any]) -> bool:
    kind = str(field.get("type", field.get("kind", "text"))).lower()
    return kind in {"checkbox", "bool", "boolean", "toggle"} or is_checkbox_group(field)


def is_workbench_route_field(field: dict[str, Any]) -> bool:
    key = field_id(field).lower()
    kind = str(field.get("type", field.get("kind", ""))).lower()
    label_text = " ".join(
        str(field.get(item) or "").lower()
        for item in ("label", "label_ru", "title", "hint", "hint_ru", "placeholder")
    )
    haystack = f"{key} {label_text}"
    route_markers = (
        "input",
        "output",
        "source",
        "destination",
        "target",
        "src",
        "dst",
        "исход",
        "источник",
        "вход",
        "выход",
        "цель",
        "результ",
        "приём",
        "прием",
    )
    path_markers = ("path", "folder", "directory", "dir", "пап", "каталог", "путь")
    return kind in {"path", "folder", "directory"} or (
        any(marker in haystack for marker in route_markers)
        and any(marker in haystack for marker in path_markers)
    )


def command_visible_fields(fields: tuple[dict[str, Any], ...] | list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [field for field in fields if not is_workbench_route_field(field)]


def workbench_value_for_field(field: dict[str, Any]) -> str:
    key = field_id(field).lower()
    if any(part in key for part in ("target", "output", "destination", "dst")):
        return str(current_target_path())
    return str(current_source_path())


def field_container_classes(field: dict[str, Any]) -> str:
    span = str(field.get("span") or field.get("width") or "").lower()
    field_key = str(field.get("id") or field.get("name") or "").lower()
    if span in {"full", "wide", "100%", "1/-1"}:
        return "audion-field audion-field-wide"
    if span in {"margin", "margins"} or field_key.startswith("margin_"):
        classes = "audion-field audion-field-narrow audion-field-margin"
        if field_key in {"margin_left_mm", "left_margin_mm", "margin_left"}:
            classes += " audion-field-margin-start"
        return classes
    if span in {"compact", "short"}:
        return "audion-field audion-field-compact"
    if span in {"narrow", "tiny", "small"}:
        return "audion-field audion-field-narrow"
    kind = str(field.get("type", field.get("kind", "text"))).lower()
    if kind in {"radio", "radiobuttons", "radio-buttons"} or is_checkbox_group(field):
        return "audion-field audion-field-wide audion-control-field"
    if kind in {"checkbox", "bool", "boolean", "toggle"}:
        return "audion-field audion-control-field"
    if kind in {"select", "choice", "format"}:
        return "audion-field audion-field-select"
    if kind in {"textarea", "multiline", "path", "file", "folder"}:
        return "audion-field audion-field-wide"
    return "audion-field"


def field_group_id(field: dict[str, Any]) -> str:
    return str(field.get("field_group") or field.get("group") or "").strip()


def field_group_title(field: dict[str, Any], fallback: str) -> str:
    if settings.language == "ru" and field.get("group_title_ru"):
        return str(field["group_title_ru"])
    return str(field.get("group_title") or fallback)


def field_group_description(field: dict[str, Any]) -> str:
    if settings.language == "ru" and field.get("group_description_ru"):
        return str(field["group_description_ru"])
    return str(field.get("group_description") or "")


def field_group_grid_classes(group_fields: list[dict[str, Any]]) -> str:
    classes = "audion-field-group-grid"
    boolean_count = sum(1 for field in group_fields if is_boolean_field(field))
    if boolean_count >= 3 and boolean_count >= max(1, len(group_fields) - 1):
        classes += " audion-field-group-grid-balanced"
    return classes


def render_fields(fields: list[dict[str, Any]]) -> None:
    if not any(field_group_id(field) for field in fields):
        with ui.element("div").classes("audion-fields-grid"):
            for field in fields:
                render_field(field)
        return

    grouped: list[dict[str, Any]] = []
    group_map: dict[str, dict[str, Any]] = {}
    loose_fields: list[dict[str, Any]] = []
    for field in fields:
        group_key = field_group_id(field)
        if not group_key:
            loose_fields.append(field)
            continue
        if group_key not in group_map:
            group_map[group_key] = {
                "key": group_key,
                "title": field_group_title(field, group_key),
                "description": field_group_description(field),
                "show_description": bool(field.get("show_group_description", False)),
                "fields": [],
            }
            grouped.append(group_map[group_key])
        elif bool(field.get("show_group_description", False)):
            group_map[group_key]["show_description"] = True
        group_map[group_key]["fields"].append(field)

    with ui.column().classes("audion-fields-stack w-full gap-3"):
        for group in grouped:
            title = str(group["title"])
            description = str(group["description"])
            tooltip = clean_tooltip(description or title)
            with ui.element("div").classes("audion-field-group"):
                attach_tooltip(ui.label(title).classes("audion-field-group-title"), tooltip)
                if description and bool(group.get("show_description", False)):
                    attach_tooltip(ui.label(description).classes("audion-field-group-description"), tooltip)
                with ui.element("div").classes(field_group_grid_classes(group["fields"])):
                    for field in group["fields"]:
                        render_field(field)
        if loose_fields:
            with ui.element("div").classes("audion-fields-grid"):
                for field in loose_fields:
                    render_field(field)


def field_file_picker_click_handler(key: str, field: dict[str, Any], control: Any):
    async def handler() -> None:
        await pick_field_file(key, field, control)

    return handler


def render_field(field: dict[str, Any]) -> None:
    key = field_id(field)
    if not key:
        return
    kind = str(field.get("type", field.get("kind", "text"))).lower()
    label = field_label(field)
    value = current_field_value(field)
    hint = field_hint(field)
    tooltip = field_tooltip(field)

    with ui.element("div").classes(field_container_classes(field)):
        if kind in {"select", "choice", "format"}:
            select_control = ui.select(
                options=select_options(field),
                label=label,
                value=value,
                on_change=lambda event, item_key=key: set_field_value(item_key, event.value),
            ).props("dense outlined popup-content-class=audion-select-popup").classes("audion-select w-full")
            attach_tooltip(select_control, tooltip)
            render_field_hint(field, hint)
            return

        if kind in {"radio", "radiobuttons", "radio-buttons"}:
            attach_tooltip(ui.label(label).classes("audion-field-label"), tooltip)
            radio_control = ui.radio(
                options=select_options(field),
                value=value,
                on_change=lambda event, item_key=key: set_field_value(item_key, event.value),
            ).props("dense inline").classes("audion-choice-row")
            attach_tooltip(radio_control, tooltip)
            render_field_hint(field, hint)
            return

        if kind in {"number", "int", "integer", "float"}:
            number_control = ui.number(
                label=label,
                value=value if value != "" else None,
                min=field.get("min"),
                max=field.get("max"),
                step=field.get("step", 1),
                on_change=lambda event, item_key=key: set_field_value(item_key, event.value),
            ).props("dense outlined").classes("audion-number w-full")
            with number_control.add_slot("append"):
                with ui.element("div").classes("audion-number-spinner"):
                    up_button = ui.button(
                        icon="keyboard_arrow_up",
                        on_click=lambda item_key=key, item_field=field, control=number_control: spin_number_field(item_key, item_field, control, 1),
                    ).props("dense flat round tabindex=-1").classes("audion-number-spin-button")
                    attach_tooltip(up_button, tooltip)
                    down_button = ui.button(
                        icon="keyboard_arrow_down",
                        on_click=lambda item_key=key, item_field=field, control=number_control: spin_number_field(item_key, item_field, control, -1),
                    ).props("dense flat round tabindex=-1").classes("audion-number-spin-button")
                    attach_tooltip(down_button, tooltip)
            attach_tooltip(number_control, tooltip)
            render_field_hint(field, hint)
            return

        if kind in {"checkbox", "bool", "boolean", "toggle"}:
            with ui.element("div").classes("audion-control-chip audion-control-chip-check"):
                checkbox_control = ui.checkbox(
                    label,
                    value=bool(value),
                    on_change=lambda event, item_key=key: set_field_value(item_key, bool(event.value)),
                ).props("dense").classes("audion-single-checkbox")
            attach_tooltip(checkbox_control, tooltip)
            render_field_hint(field, hint)
            return

        if is_checkbox_group(field):
            selected = set(value if isinstance(value, list) else [])
            controls: dict[Any, Any] = {}

            def sync_checkboxes(item_key: str = key) -> None:
                set_field_value(
                    item_key,
                    [option_key for option_key, checkbox in controls.items() if bool(checkbox.value)],
                )

            attach_tooltip(ui.label(label).classes("audion-field-label"), tooltip)
            with ui.row().classes("audion-choice-row"):
                for option_key, option_text in checkbox_options(field):
                    checkbox = ui.checkbox(
                        option_text,
                        value=option_key in selected,
                        on_change=lambda _event: sync_checkboxes(),
                    ).props("dense")
                    attach_tooltip(checkbox, tooltip)
                    controls[option_key] = checkbox
            render_field_hint(field, hint)
            sync_checkboxes()
            return

        if kind == "file":
            file_control = ui.input(
                label=label,
                value=str(value) if value is not None else "",
                placeholder=str(field.get("placeholder", "")),
                on_change=lambda event, item_key=key: set_field_value(item_key, event.value),
            ).props("dense outlined").classes("audion-file-input w-full")
            with file_control.add_slot("append"):
                pick_button = ui.button(
                    icon="folder_open",
                    on_click=field_file_picker_click_handler(key, field, file_control),
                ).props("dense flat round tabindex=-1").classes("audion-field-picker-button")
                attach_tooltip(pick_button, tr("choose_file"))
            attach_tooltip(file_control, tooltip)
            render_field_hint(field, hint)
            return

        input_control = ui.input(
            label=label,
            value=str(value) if value is not None else "",
            placeholder=str(field.get("placeholder", "")),
            on_change=lambda event, item_key=key: set_field_value(item_key, event.value),
        ).props("dense outlined").classes("w-full")
        attach_tooltip(input_control, tooltip)
        render_field_hint(field, hint)


def operation_from_command_node(node: CommandNode, extra_parameters: dict[str, Any] | None = None) -> Operation:
    parameters = dict(node.parameters)
    values = state.setdefault("field_values", {})
    for field in node.fields:
        key = field_id(field)
        if key and is_workbench_route_field(field):
            parameters[key] = workbench_value_for_field(field)
        elif key:
            parameters[key] = values.get(key, field_default(field))
    if extra_parameters:
        parameters.update(extra_parameters)
    return node.to_operation(parameters)


def operation_from_pending_command(node: CommandNode) -> Operation:
    return operation_from_command_node(node)


def validate_pending_fields(node: CommandNode) -> bool:
    values = state.setdefault("field_values", {})
    for field in command_visible_fields(node.fields):
        if not is_checkbox_group(field):
            continue
        min_selected = int(field.get("min_selected", 0) or 0)
        if min_selected <= 0:
            continue
        key = field_id(field)
        selected = values.get(key, field_default(field))
        if not isinstance(selected, list) or len(selected) < min_selected:
            safe_notify(tr("select_required", field=field_label(field)), "warning")
            return False
    return True


async def run_pending_command(node: CommandNode) -> None:
    if validate_pending_fields(node):
        await start_operation(operation_from_pending_command(node))


async def run_command_node(node: CommandNode, extra_parameters: dict[str, Any] | None = None) -> None:
    if validate_pending_fields(node):
        await start_operation(operation_from_command_node(node, extra_parameters))


def run_pending_click_handler(node: CommandNode):
    async def handler() -> None:
        await run_pending_command(node)

    return handler


def run_command_node_click_handler(node: CommandNode, extra_parameters: dict[str, Any] | None = None):
    async def handler() -> None:
        await run_command_node(node, extra_parameters)

    return handler


def is_direct_action_node(node: CommandNode) -> bool:
    """Leaf commands without user-editable controls run from their section."""
    return bool(node.service) and not node.children and not command_visible_fields(node.fields)


def direct_action_button(node: CommandNode) -> None:
    label = node.display_title(settings.language)
    description = node.display_description(settings.language)
    tooltip = command_tooltip(node)
    with ui.element("div").classes("audion-direct-action-item"):
        with ui.button(on_click=run_command_node_click_handler(node)).props("dense flat no-wrap").classes(
            "audion-direct-action rounded-lg"
        ) as button:
            ui.label(label).classes("audion-direct-action-title")
        attach_tooltip(button, tooltip)
        if description:
            attach_tooltip(ui.label(description).classes("audion-direct-action-description"), tooltip)
        else:
            ui.element("div").classes("audion-direct-action-description")


def command_node_button(node: CommandNode) -> None:
    if is_direct_action_node(node):
        direct_action_button(node)
        return
    has_children = bool(node.children)
    label = node.display_title(settings.language)
    description = node.display_description(settings.language)
    if has_children and not description:
        description = tr("open_menu")
    tooltip = command_tooltip(node)

    with ui.element("div").classes("audion-operation-row"):
        button = ui.button(
            label,
            on_click=command_click_handler(node),
        ).props("dense flat no-wrap").classes("audion-action audion-operation-button rounded-lg")
        attach_tooltip(button, tooltip)
        attach_tooltip(ui.label(description).classes("audion-operation-description"), tooltip)


def folder_button(label: str, folder: Path) -> None:
    button = ui.button(
        label,
        icon="folder_open",
        on_click=lambda: open_folder(folder),
    ).props("dense flat no-wrap").classes("audion-action audion-context-folder-button rounded-lg")
    attach_tooltip(button, audion_folder_button_tooltip("rules", folder))


def command_nav_row(trail: list[CommandNode], pending: CommandNode | None) -> None:
    can_go_back = pending is not None or bool(trail)
    if pending is not None:
        title = pending.display_title(settings.language)
    elif trail:
        title = " / ".join(node.display_title(settings.language) for node in trail)
    else:
        title = ""

    with ui.row().classes("audion-command-nav w-full items-center gap-2"):
        if can_go_back:
            ui.button(
                tr("back"),
                on_click=go_back_command,
            ).props("dense flat no-wrap").classes("audion-action w-28 rounded-lg")
        ui.label(title).classes("min-w-0 flex-1 truncate text-sm text-gray-400")
        if pending is not None:
            ui.button(
                tr("run"),
                on_click=run_pending_click_handler(pending),
            ).props("dense flat no-wrap").classes("audion-action w-28 rounded-lg")
        elif len(trail) == 1 and trail[0].id == "processors":
            active_tab = str(state.get("processor_tab") or "hygiene")
            if active_tab not in processor_tab_ids():
                active_tab = "hygiene"
            node = processor_action_node(trail[0], active_tab)
            if node is not None:
                mode = processor_mode(active_tab)
                extra_parameters = {"deep_hygiene_mode": mode} if active_tab == "hygiene" else None
                run_button = ui.button(
                    processor_run_label(mode),
                    on_click=run_command_node_click_handler(node, extra_parameters),
                ).props("dense flat no-wrap").classes(f"audion-workbench-run audion-workbench-run-{mode} audion-command-nav-run")
                attach_tooltip(run_button, command_tooltip(node))


def is_audit_rules_context(trail: list[CommandNode], pending: CommandNode | None) -> bool:
    ids = {node.id for node in trail}
    if pending is not None:
        ids.add(pending.id)
    return bool(ids & {"audit_rules", "docx_audit_processor", "docx_audit_strip_anchors"})


def processor_tab_ids() -> tuple[str, str, str]:
    return ("hygiene", "anomalies", "styles")


def processor_tab_title(tab_id: str) -> str:
    titles = {
        "ru": {
            "hygiene": "ГИГИЕНА ТЕКСТА",
            "anomalies": "АНОМАЛИИ ДОКУМЕНТА",
            "styles": "СТИЛИ ДОКУМЕНТА",
        },
        "en": {
            "hygiene": "TEXT HYGIENE",
            "anomalies": "DOCUMENT ANOMALIES",
            "styles": "DOCUMENT STYLES",
        },
    }
    return titles.get(settings.language, titles["en"]).get(tab_id, tab_id)


def processor_tab_tooltip(tab_id: str) -> str:
    tooltips = {
        "ru": {
            "hygiene": "Гигиена текста: пробелы, пунктуация, мягкие переносы и предметные правила аудита. Может подключить аномалии документа вторым проходом.",
            "anomalies": "Поиск и безопасная корректировка форматных сбоев: секции, поля, ориентация, таблицы, подписи, нумерация, колонтитулы и разрывы.",
            "styles": "Редкий, но важный слой: проверка и консервативное назначение существующих стилей Word для заголовков, подписей, оглавления и списков.",
        },
        "en": {
            "hygiene": "Existing text hygiene: spacing, punctuation, soft hyphens, and deterministic audit rules. Can attach document anomalies as a second pass.",
            "anomalies": "Scan and safe correction for formatting issues: sections, margins, orientation, tables, captions, numbering, headers, footers, and breaks.",
            "styles": "Less frequent but important layer: scan and conservative assignment of existing Word styles for headings, captions, TOC, and lists.",
        },
    }
    return tooltips.get(settings.language, tooltips["en"]).get(tab_id, tab_id)


def processor_mode(section: str) -> str:
    modes = state.setdefault("processor_modes", {})
    value = str(modes.get(section) or "scan").lower()
    return "fix" if value == "fix" else "scan"


def set_processor_tab(tab_id: str) -> None:
    if tab_id not in processor_tab_ids():
        return
    state["processor_tab"] = tab_id
    state["pending_command"] = None
    command_tree.refresh()


def set_processor_mode(section: str, mode: str) -> None:
    normalized = "fix" if str(mode).lower() == "fix" else "scan"
    state.setdefault("processor_modes", {})[section] = normalized
    if section == "hygiene":
        set_field_value("deep_hygiene_mode", normalized)
    command_tree.refresh()


def processor_mode_title(mode: str) -> str:
    if settings.language == "ru":
        return "Корректировка" if mode == "fix" else "Проверка"
    return "Fix" if mode == "fix" else "Scan"


def processor_mode_tooltip(section: str, mode: str) -> str:
    if settings.language == "ru":
        if mode == "fix":
            return "Корректировка применяет только безопасные правки и пишет копии DOCX, не трогая исходные документы."
        return "Проверка строит отчёты и не меняет документы."
    if mode == "fix":
        return "Fix mode applies safe corrections to copied DOCX files only."
    return "Scan mode is read-only and only writes reports."


def render_processor_mode_buttons(section: str) -> None:
    selected = processor_mode(section)
    with ui.element("div").classes("audion-mode-row"):
        label = "Режим" if settings.language == "ru" else "Mode"
        attach_tooltip(ui.label(label).classes("audion-mode-label"), processor_mode_tooltip(section, selected))
        for mode in ("scan", "fix"):
            active = selected == mode
            classes = f"audion-mode-button audion-mode-{mode}"
            if active:
                classes += " audion-mode-button-active"
            button = ui.button(
                processor_mode_title(mode),
                on_click=lambda item_mode=mode, item_section=section: set_processor_mode(item_section, item_mode),
            ).props("dense flat no-wrap").classes(classes)
            attach_tooltip(button, processor_mode_tooltip(section, mode))


def render_processor_tabs(active_tab: str) -> None:
    with ui.element("div").classes("audion-workbench-tabs"):
        for tab_id in processor_tab_ids():
            classes = f"audion-workbench-tab audion-workbench-tab-{tab_id}"
            if tab_id == active_tab:
                classes += " audion-workbench-tab-active"
            button = ui.button(
                processor_tab_title(tab_id),
                on_click=lambda item_tab=tab_id: set_processor_tab(item_tab),
            ).props("dense flat no-wrap").classes(classes)
            attach_tooltip(button, processor_tab_tooltip(tab_id))


def processor_action_node(processors_node: CommandNode, tab_id: str) -> CommandNode | None:
    if tab_id == "hygiene":
        return child_command_node(processors_node, "docx_deep_hygiene")
    if tab_id == "anomalies":
        return child_command_node(processors_node, "docx_anomaly_correct" if processor_mode("anomalies") == "fix" else "docx_anomaly_scan")
    if tab_id == "styles":
        return child_command_node(processors_node, "docx_style_fix" if processor_mode("styles") == "fix" else "docx_style_scan")
    return None


def processor_run_label(mode: str) -> str:
    if settings.language == "ru":
        return "Запустить корректировку" if mode == "fix" else "Запустить проверку"
    return "Run fix" if mode == "fix" else "Run scan"


def render_processor_action_card(
    tab_id: str,
    section: str,
    node: CommandNode,
    *,
    exclude_fields: set[str] | None = None,
) -> None:
    fields = [
        field
        for field in command_visible_fields(node.fields)
        if field_id(field) not in (exclude_fields or set())
    ]
    with ui.element("div").classes(f"audion-workbench-card audion-workbench-card-{tab_id}"):
        with ui.row().classes("audion-workbench-card-header w-full items-center gap-2"):
            attach_tooltip(ui.label(processor_tab_title(tab_id)).classes("audion-workbench-title"), processor_tab_tooltip(tab_id))
        if fields:
            render_fields(fields)


def render_processor_extra_operation(node: CommandNode) -> None:
    tooltip = command_tooltip(node)
    with ui.element("div").classes("audion-workbench-extra"):
        attach_tooltip(ui.label(node.display_title(settings.language)).classes("audion-workbench-extra-title"), tooltip)
        button = ui.button(
            "Запустить" if settings.language == "ru" else "Run",
            on_click=run_command_node_click_handler(node),
        ).props("dense flat no-wrap").classes("audion-workbench-extra-run")
        attach_tooltip(button, tooltip)


def render_processors_workbench(processors_node: CommandNode) -> None:
    active_tab = str(state.get("processor_tab") or "hygiene")
    if active_tab not in processor_tab_ids():
        active_tab = "hygiene"
        state["processor_tab"] = active_tab

    with ui.element("div").classes("audion-workbench"):
        render_processor_tabs(active_tab)
        node = processor_action_node(processors_node, active_tab)
        if node is None:
            ui.label("Команда не найдена." if settings.language == "ru" else "Command not found.").classes("text-sm text-gray-400")
            return

        if active_tab == "hygiene":
            render_processor_mode_buttons("hygiene")
            render_processor_action_card(
                "hygiene",
                "hygiene",
                node,
                exclude_fields={"deep_hygiene_mode"},
            )
            extra = child_command_node(processors_node, "docx_nonprinting_clean")
            if extra is not None:
                render_processor_extra_operation(extra)
            return

        if active_tab == "anomalies":
            render_processor_mode_buttons("anomalies")
            render_processor_action_card("anomalies", "anomalies", node)
            return

        render_processor_mode_buttons("styles")
        render_processor_action_card("styles", "styles", node)


def audit_run_label(node_id: str) -> str:
    if settings.language == "ru":
        return "Запустить аудит" if node_id == "docx_audit_processor" else "Снять якоря"
    return "Run audit" if node_id == "docx_audit_processor" else "Remove anchors"


def render_audit_operation_card(node: CommandNode, *, accent: str) -> None:
    tooltip = command_tooltip(node)
    visible_fields = command_visible_fields(node.fields)
    with ui.element("section").classes(f"audion-audit-card audion-audit-card-{accent}"):
        with ui.row().classes("audion-audit-card-header w-full items-start gap-3"):
            with ui.column().classes("min-w-0 flex-1 gap-1"):
                attach_tooltip(ui.label(node.display_title(settings.language)).classes("audion-audit-card-title"), tooltip)
                description = node.display_description(settings.language)
                if description:
                    attach_tooltip(ui.label(description).classes("audion-audit-card-description"), tooltip)
            run_button = ui.button(
                audit_run_label(node.id),
                on_click=run_command_node_click_handler(node),
            ).props("dense flat no-wrap").classes("audion-audit-run rounded-lg")
            attach_tooltip(run_button, tooltip)
        if visible_fields:
            with ui.element("div").classes("audion-parameters-block audion-audit-parameters"):
                render_fields(visible_fields)


def render_audit_rules_workbench(audit_node: CommandNode) -> None:
    processor = child_command_node(audit_node, "docx_audit_processor")
    strip_anchors = child_command_node(audit_node, "docx_audit_strip_anchors")
    with ui.element("div").classes("audion-audit-workbench"):
        if processor is not None:
            render_audit_operation_card(processor, accent="processor")
        if strip_anchors is not None:
            render_audit_operation_card(strip_anchors, accent="anchors")


@ui.refreshable
def command_tree() -> None:
    trail, nodes = current_command_level()
    pending = state.get("pending_command")
    command_nav_row(trail, pending)

    if is_audit_rules_context(trail, pending):
        folder_button(tr("rules"), rules_folder)

    if pending is not None:
        ui.label(tr("selected_operation")).classes("text-sm font-semibold text-gray-300")
        tooltip = command_tooltip(pending)
        attach_tooltip(ui.label(pending.display_title(settings.language)).classes("audion-operation-description font-semibold"), tooltip)
        description = pending.display_description(settings.language)
        if description:
            attach_tooltip(ui.label(description).classes("text-sm text-gray-400 audion-operation-description"), tooltip)
        visible_fields = command_visible_fields(pending.fields)
        if visible_fields:
            ui.label(tr("parameters")).classes("text-sm font-semibold text-gray-300")
            with ui.element("div").classes("audion-parameters-block"):
                render_fields(visible_fields)
        return

    if len(trail) == 1 and trail[0].id == "processors":
        render_processors_workbench(trail[0])
        return

    if len(trail) == 1 and trail[0].id == "audit_rules":
        render_audit_rules_workbench(trail[0])
        return

    direct_nodes = [node for node in nodes if is_direct_action_node(node)]
    regular_nodes = [node for node in nodes if not is_direct_action_node(node)]
    if direct_nodes:
        with ui.element("div").classes("audion-direct-action-grid"):
            for node in direct_nodes:
                direct_action_button(node)

    for node in regular_nodes:
        command_node_button(node)


def operation_by_id(operation_id: str) -> Operation | None:
    for operation in [*manifest.operations, *manifest.maintenance_operations]:
        if operation.id == operation_id:
            return operation
    return None


_application_css_cache: dict[str, str] = {}


def application_css(name: str) -> str:
    """A stylesheet that lives next to this module rather than inside it."""
    if name not in _application_css_cache:
        path = Path(__file__).resolve().with_name(name)
        _application_css_cache[name] = path.read_text(encoding="utf-8")
    return _application_css_cache[name]


def add_styles() -> None:
    add_audion_canonical_ui_styles()
    variables_css = "\n".join(
        f"            --{key}: {value};"
        for key, value in sorted(theme_variables().items())
    )
    ui.add_head_html(
        "<style>\n"
        ":root {\n"
        f"{variables_css}\n"
        "}\n"
        + application_css("tokens.css")
        + application_css("theme.css")
        + "\n</style>\n"
    )


def build_ui() -> None:
    ensure_project_dirs(paths)
    if not state["status"]:
        state["status"] = tr("idle")
    if active_theme_mode() == "dark":
        ui.dark_mode().enable()
    else:
        ui.dark_mode().disable()
    add_styles()
    ui.add_head_html(f"<style>{WORKBENCH_LAYOUT_CSS}\n{WORKBENCH_OVERRIDE_CSS}</style>")
    ui.add_head_html(WORKBENCH_FEEDBACK_CSS)

    with ui.header().classes("audion-header h-[42px] items-center justify-between px-4"):
        ui.label(app_title()).classes("audion-header-title text-lg font-bold")
        with ui.row().classes("audion-header-controls items-center gap-2"):
            ui.icon("palette").classes("text-lg")
            ui.select(
                options=theme_options(),
                value=active_theme(),
                on_change=theme_change_handler,
            ).props("dense outlined options-dense").classes("audion-theme-select")
            ui.button(tr("lang_switch"), on_click=toggle_language).props("dense flat").classes("audion-action rounded-lg")
            cancel_button = ui.button(tr("cancel"), on_click=lambda: state.update({"cancel": True})).props("dense flat color=negative")
            cancel_button.visible = False

    with ui.element("div").classes("audion-shell"):
        with ui.column().classes("audion-pane audion-scroll gap-3"):
            with ui.column().classes("audion-panel audion-workspace-panel w-full gap-2 p-2"):
                WORKBENCH_RENDERER.render_address_rows()
                WORKBENCH_RENDERER.render_action_bar()

            ui.label(f"{em('operations')}{tr('operations')}").classes("audion-section-heading")
            command_tree()

            visible_maintenance_operations = [
                operation
                for operation in manifest.maintenance_operations
                if operation.id != "cleanup_input_output"
            ]
            if visible_maintenance_operations:
                ui.label(f"{em('maintenance')}{tr('maintenance')}").classes("audion-section-heading")
                for operation in visible_maintenance_operations:
                    operation_button(operation)

        ui.element("div").classes("audion-splitter").props(f'title="{tr("resize_panels")}"')

        with ui.element("div").classes("audion-pane audion-right gap-2 pt-3"):
            with ui.column().classes("audion-panel w-full gap-2 p-3"):
                        with ui.element("div").classes(status_row_classes()) as status_row:
                            status_dot_main = ui.element("span").classes("audion-status-dot-mark")
                            status_state_label = ui.label(status_state_text()).classes("audion-status-state")
                            status_label = ui.label(str(state["status"])).classes("audion-status-message")
                            status_clock = ui.label(elapsed_text(None)).classes("audion-status-clock")
                            with ui.element("div").classes("audion-status-bar"):
                                status_bar_fill = ui.element("i").style("width: 0%")
                            status_percent = ui.label(progress_text()).classes("audion-status-percent")

            with ui.column().classes("audion-terminal-panel w-full gap-2 p-3"):
                with ui.row().classes("audion-log-toolbar w-full items-center gap-2"):
                    ui.label(f"{em('log')}{tr('log')}").classes("text-base font-semibold")
                    ui.space()
                    ui.button(tr("logs"), on_click=lambda: open_folder(paths.logs)).props("dense flat").classes("audion-action rounded-lg").tooltip(audion_folder_button_tooltip("logs", paths.logs))
                    ui.button(tr("report"), on_click=lambda: open_folder(paths.report)).props("dense flat").classes("audion-action rounded-lg").tooltip(audion_folder_button_tooltip("report", paths.report))
                    ui.button(tr("rules"), on_click=lambda: open_folder(rules_folder)).props("dense flat").classes("audion-action rounded-lg")
                    ui.button(tr("config"), on_click=lambda: open_folder(paths.config)).props("dense flat").classes("audion-action rounded-lg").tooltip(audion_folder_button_tooltip("config", paths.config))
                    clear_log_button = ui.button(icon="delete_sweep", on_click=clear_terminal_log).props("dense flat round").classes("audion-action audion-log-icon-button")
                    clear_log_button.tooltip(audion_terminal_action_tooltip("clear_terminal_window"))
                    expand_log_button = ui.button(icon="open_in_full", on_click=lambda: log_dialog.open()).props("dense flat round").classes("audion-action audion-log-icon-button")
                    expand_log_button.tooltip(audion_terminal_action_tooltip("expand"))
                log_view = ui.html(terminal_html(), sanitize=False).classes("audion-terminal w-full min-h-[66vh]")
                with ui.row().classes("audion-terminal-footer w-full items-center gap-2 px-1 pt-1"):
                    status_dot = ui.label("●").classes(status_dot_classes())
                    terminal_status_label = ui.label(str(state["status"])).classes("min-w-0 flex-1 truncate text-xs")

    with ui.dialog() as log_dialog:
        with ui.card().classes("audion-dialog h-[92vh] w-[92vw] rounded-lg p-3"):
            with ui.row().classes("w-full items-center gap-2"):
                ui.label(f"{em('log')}{tr('log')}").classes("text-base font-semibold")
                ui.space()
                ui.button(tr("rules"), on_click=lambda: open_folder(rules_folder)).props("dense flat").classes("audion-action rounded-lg")
                ui.button(tr("config"), on_click=lambda: open_folder(paths.config)).props("dense flat").classes("audion-action rounded-lg").tooltip(audion_folder_button_tooltip("config", paths.config))
                clear_expanded_log_button = ui.button(icon="delete_sweep", on_click=clear_terminal_log).props("dense flat round").classes("audion-action audion-log-icon-button")
                clear_expanded_log_button.tooltip(audion_terminal_action_tooltip("clear_terminal_window"))
                ui.button(tr("close"), on_click=log_dialog.close).props("dense flat").classes("audion-action rounded-lg").tooltip(audion_terminal_action_tooltip("close"))
            expanded_log_view = ui.html(terminal_html(), sanitize=False).classes("audion-terminal audion-terminal-expanded w-full")

    ui.run_javascript(
        """
        (() => {
          const storageKey = 'audion_gui_terminal_width_px';
          const defaultWidth = 666;
          const minLeft = 460;
          const minRight = 460;

          const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

          const applyWidth = (width) => {
            const shell = document.querySelector('.audion-shell');
            if (!shell) return;
            const rect = shell.getBoundingClientRect();
            const maxRight = Math.max(minRight, rect.width - minLeft - 40);
            const next = clamp(Number(width) || defaultWidth, minRight, maxRight);
            shell.style.setProperty('--audion-terminal-width', `${Math.round(next)}px`);
            localStorage.setItem(storageKey, String(Math.round(next)));
          };

          const setup = () => {
            const shell = document.querySelector('.audion-shell');
            const splitter = document.querySelector('.audion-splitter');
            if (!shell || !splitter) {
              setTimeout(setup, 80);
              return;
            }
            if (splitter.dataset.audionReady === '1') return;
            splitter.dataset.audionReady = '1';

            applyWidth(localStorage.getItem(storageKey) || defaultWidth);

            let dragging = false;
            const updateFromEvent = (event) => {
              if (!dragging) return;
              const rect = shell.getBoundingClientRect();
              const rightWidth = rect.right - event.clientX - 10;
              applyWidth(rightWidth);
            };

            splitter.addEventListener('pointerdown', (event) => {
              dragging = true;
              splitter.setPointerCapture?.(event.pointerId);
              document.body.classList.add('audion-resizing');
              event.preventDefault();
            });
            splitter.addEventListener('pointermove', updateFromEvent);
            splitter.addEventListener('pointerup', (event) => {
              dragging = false;
              splitter.releasePointerCapture?.(event.pointerId);
              document.body.classList.remove('audion-resizing');
            });
            splitter.addEventListener('pointercancel', () => {
              dragging = false;
              document.body.classList.remove('audion-resizing');
            });
            window.addEventListener('resize', () => applyWidth(localStorage.getItem(storageKey) || defaultWidth));
          };

          setup();
        })();
        """
    )

    last_log_state = {
        "version": int(state["log_version"]),
        "epoch": int(state["terminal_epoch"]),
        "seq": int(state["line_seq"]),
    }

    refresh_timer: Any | None = None

    # Every one of these used to be written twice a second whether or not it had
    # changed, so an idle window still sent ten element updates a second. Holding
    # the last value makes an idle panel cost nothing and pays for the clock.
    shown = {"status": None, "state": None, "row": None, "clock": None, "percent": None, "fill": None}
    run_clock: dict[str, float | None] = {"started": None, "frozen": None}

    def refresh() -> None:
        nonlocal refresh_timer
        try:
            running = bool(state["running"])
            if running and run_clock["started"] is None:
                run_clock["started"] = time.monotonic()
                run_clock["frozen"] = None
            elif not running and run_clock["started"] is not None:
                run_clock["frozen"] = time.monotonic() - run_clock["started"]
                run_clock["started"] = None
            seconds = (
                time.monotonic() - run_clock["started"]
                if run_clock["started"] is not None
                else run_clock["frozen"]
            )

            def show(key: str, value: Any, assign: Any) -> None:
                if shown[key] != value:
                    shown[key] = value
                    assign(value)

            message = str(state["status"])
            show("status", message, lambda value: (
                setattr(status_label, "text", value),
                setattr(terminal_status_label, "text", value),
            ))
            show("state", status_state_text(), lambda value: setattr(status_state_label, "text", value))
            show("row", status_row_classes(), lambda value: (
                status_row.classes(replace=value),
                status_dot.classes(replace=status_dot_classes()),
            ))
            show("clock", elapsed_text(seconds), lambda value: setattr(status_clock, "text", value))
            show("percent", progress_text(), lambda value: setattr(status_percent, "text", value))
            show("fill", f"{float(state['progress']) * 100:.1f}%",
                lambda value: status_bar_fill.style(f"width: {value}"))
            log_version = int(state["log_version"])
            terminal_epoch = int(state["terminal_epoch"])
            line_seq = int(state["line_seq"])
            line_base_seq = int(state["line_base_seq"])
            if log_version != last_log_state["version"]:
                previous_seq = int(last_log_state["seq"])
                reset_required = terminal_epoch != int(last_log_state["epoch"]) or previous_seq < line_base_seq - 1
                if reset_required:
                    rendered_log = terminal_html()
                    log_view.set_content(rendered_log)
                    expanded_log_view.set_content(rendered_log)
                    ui.run_javascript(
                        """
                        requestAnimationFrame(() => {
                          document.querySelectorAll('.audion-terminal').forEach((el) => {
                            el.scrollTop = el.scrollHeight;
                          });
                        });
                        """
                    )
                elif line_seq > previous_seq:
                    start_index = max(0, previous_seq - line_base_seq + 1)
                    fragment = terminal_cached_lines_html(state["line_html"][start_index:])
                    fragment_json = json.dumps(fragment, ensure_ascii=False)
                    ui.run_javascript(
                        f"""
                        (() => {{
                          const fragment = {fragment_json};
                          const maxLines = {TERMINAL_HISTORY_LIMIT};
                          requestAnimationFrame(() => {{
                            document.querySelectorAll('.audion-terminal').forEach((container) => {{
                              const pre = container.querySelector('.audion-terminal-pre');
                              if (!pre) return;
                              const selection = window.getSelection ? window.getSelection() : null;
                              const selectingInside = Boolean(
                                selection && !selection.isCollapsed &&
                                (container.contains(selection.anchorNode) || container.contains(selection.focusNode))
                              );
                              const distanceFromBottom = container.scrollHeight - container.scrollTop - container.clientHeight;
                              const shouldStickToBottom = !selectingInside && distanceFromBottom <= 8;
                              pre.insertAdjacentHTML('beforeend', fragment);
                              const lines = pre.querySelectorAll(':scope > .audion-terminal-line');
                              const excess = lines.length - maxLines;
                              for (let index = 0; index < excess; index += 1) {{
                                lines[index].remove();
                              }}
                              if (shouldStickToBottom) {{
                                container.scrollTop = container.scrollHeight;
                              }}
                            }});
                          }});
                        }})();
                        """
                    )
                last_log_state["version"] = log_version
                last_log_state["epoch"] = terminal_epoch
                last_log_state["seq"] = line_seq
            cancel_button.visible = bool(state["running"])
        except RuntimeError as exc:
            message = str(exc)
            if "slot belongs to has been deleted" not in message and "current slot cannot be determined" not in message:
                raise
            logging.warning("NiceGUI refresh timer stopped because the client slot was deleted.")
            if refresh_timer is not None:
                refresh_timer.deactivate()

    refresh_timer = ui.timer(0.5, refresh)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audion NiceGUI shell.")
    parser.add_argument("--host", default=str(ui_info.get("host", "127.0.0.1")))
    parser.add_argument("--port", type=int, default=int(ui_info.get("port", 8080)))
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--smoke", action="store_true")
    return parser.parse_args()


def port_is_open(host: str, port: int) -> bool:
    family = socket.AF_INET6 if ":" in str(host or "") else socket.AF_INET
    try:
        with socket.socket(family, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.3)
            return sock.connect_ex((host, port)) == 0
    except OSError:
        return False


def assert_gui_host_allowed(host: str) -> None:
    normalized = str(host or "").strip().lower().strip("[]")
    try:
        is_loopback = normalized == "localhost" or ipaddress.ip_address(normalized).is_loopback
    except ValueError:
        is_loopback = normalized == "localhost"
    remote_allowed = str(os.environ.get("AUDION_ALLOW_REMOTE_GUI", "")).strip().lower() in {"1", "true", "yes", "on"}
    if is_loopback or remote_allowed:
        return
    raise SystemExit(
        "Refusing non-loopback host for a GUI with process execution. "
        "Use 127.0.0.1/localhost/::1, or set AUDION_ALLOW_REMOTE_GUI=1 explicitly."
    )


def build_ui_once() -> dict[str, int]:
    """Build the whole page once, headlessly, and report what came of it.

    `--smoke` used to print a line and return, so an app could ship a `build_ui`
    that raised on its first statement and still pass — twice in this fleet it did.
    Here the page is actually built: no browser and no HTTP request, so whatever
    the app defers until a client attaches is skipped, but every widget is
    constructed and the stylesheet has to arrive.
    """
    import asyncio
    import logging
    import re

    from nicegui import core
    from nicegui.client import Client
    from nicegui.page import page as page_definition

    async def build() -> tuple[int, str]:
        core.loop = asyncio.get_running_loop()
        # Work deferred to a connected browser fails here and says nothing about
        # the build. An exception raised by build_ui itself still propagates.
        core.loop.set_exception_handler(lambda _loop, _context: None)
        logging.getLogger("nicegui").setLevel(logging.CRITICAL)
        client = Client(page_definition("/__smoke__"))
        with client:
            build_ui()
        report = len(client.elements), client.shared_head_html + client.head_html
        # The page starts work that waits for a browser to attach. Nothing will
        # attach, so stop it deliberately instead of letting the loop close on it.
        pending = asyncio.all_tasks(core.loop) - {asyncio.current_task()}
        for task in pending:
            task.cancel()
        if pending:
            await asyncio.gather(*pending, return_exceptions=True)
        return report

    element_count, head = asyncio.run(build())
    if element_count < 2:
        raise RuntimeError("build_ui produced no widgets")
    # Token prefixes differ between apps, so look for any custom property rather
    # than for one project's naming.
    if not re.search(r"--[\w-]+\s*:", head):
        raise RuntimeError("the stylesheet never reached the page")
    return {"elements": element_count, "stylesheet_bytes": len(head)}


def main() -> int:
    args = parse_args()
    ensure_project_dirs(paths)
    assert_gui_host_allowed(args.host)
    if args.smoke:
        try:
            report = build_ui_once()
        except Exception as error:  # noqa: BLE001
            print(f"FAIL nicegui shell: {ROOT}: {error}")
            return 1
        print(
            f"OK nicegui shell: {ROOT}"
            f" | widgets={report['elements']}"
            f" | stylesheet={report['stylesheet_bytes']} bytes"
        )
        return 0

    if port_is_open(args.host, args.port):
        url = f"http://{args.host}:{args.port}/"
        print(f"GUI already appears to be running: {url}")
        if not args.no_browser:
            webbrowser.open(url)
        return 0

    ui.run(
        root=build_ui,
        title=app_title(),
        host=args.host,
        port=args.port,
        reload=False,
        native=False,
        show=not args.no_browser,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
