from __future__ import annotations

from pathlib import Path
import os
import shutil
import subprocess
import sys

from system_core.core.jobs import JobContext, iter_subprocess_lines, popen_gui_command, unbuffer_python_command
from system_core.core.paths import open_folder
from system_core.services.office_service import SCRIPT_RUNNER


def _is_inside(child: Path, parent: Path) -> bool:
    try:
        return os.path.commonpath([str(child.resolve()), str(parent.resolve())]) == str(parent.resolve())
    except (OSError, ValueError):
        return False


def _run(context: JobContext, command: list[str]) -> dict[str, object]:
    command = unbuffer_python_command(command)
    context.log("> " + " ".join(f'"{part}"' if " " in part else part for part in command))
    process = popen_gui_command(command, cwd=context.paths.root)
    try:
        for line in iter_subprocess_lines(process, command):
            context.log(line)
            if context.cancelled():
                process.terminate()
                break
    finally:
        if process.stdout:
            process.stdout.close()
    return_code = process.wait()
    context.progress(1.0)
    if context.cancelled() and return_code != 0:
        raise RuntimeError("Команда отменена")
    if return_code != 0:
        raise RuntimeError(f"Команда завершилась с кодом {return_code}")
    return {"exit_code": return_code}


def _python_script(context: JobContext, script: Path, *args: str | Path) -> list[str]:
    if not script.exists():
        raise FileNotFoundError(f"Не найден скрипт: {script}")
    return [sys.executable, "-u", "-c", SCRIPT_RUNNER, str(script), *[str(arg) for arg in args]]


def _cmd_script(context: JobContext, script: Path) -> list[str]:
    if not script.exists():
        raise FileNotFoundError(f"Не найден скрипт: {script}")
    return ["cmd.exe", "/c", str(script)]


def _open(context: JobContext, folder: Path) -> dict[str, object]:
    if not _is_inside(folder, context.paths.root):
        raise RuntimeError("Папка находится вне корня проекта.")
    open_folder(folder)
    context.log(f"Открыта папка: {folder}")
    context.progress(1.0)
    return {"folder": str(folder)}


def env_doctor(context: JobContext) -> dict[str, object]:
    return _run(context, _python_script(context, context.paths.system_core / "doctor.py"))


def collect_licenses_python(context: JobContext) -> dict[str, object]:
    script = context.paths.system_core / "license" / "collect_third_party_licenses.py"
    return _run(context, _python_script(context, script))


def prune_stale_licenses(context: JobContext) -> dict[str, object]:
    script = context.paths.system_core / "license" / "prune_stale_collected_licenses.py"
    return _run(context, _python_script(context, script))


def deduplicate_licenses(context: JobContext) -> dict[str, object]:
    script = context.paths.system_core / "license" / "deduplicate_collected_licenses.py"
    return _run(context, _python_script(context, script))


def make_release_archive(context: JobContext) -> dict[str, object]:
    return _run(context, _cmd_script(context, context.paths.root / "install" / "make_release_archive.cmd"))


def open_tools_launcher(context: JobContext) -> dict[str, object]:
    launcher = context.paths.root / "launcher_tools.cmd"
    if not launcher.exists():
        raise FileNotFoundError(f"Не найден launcher_tools.cmd: {launcher}")
    subprocess.Popen(["cmd.exe", "/c", "start", "", str(launcher)], cwd=context.paths.root)
    context.log(f"Открыто консольное меню обслуживания: {launcher}")
    context.progress(1.0)
    return {"launcher": str(launcher)}


def open_licenses(context: JobContext) -> dict[str, object]:
    return _open(context, context.paths.root / "licenses")


def open_release(context: JobContext) -> dict[str, object]:
    return _open(context, context.paths.release)


def open_system_core(context: JobContext) -> dict[str, object]:
    return _open(context, context.paths.system_core)


def open_install(context: JobContext) -> dict[str, object]:
    return _open(context, context.paths.root / "install")


def open_wheelhouse(context: JobContext) -> dict[str, object]:
    return _open(context, context.paths.root / "wheelhouse")
