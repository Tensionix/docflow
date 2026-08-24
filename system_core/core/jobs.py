from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
import codecs
import importlib
import json
import locale
import os
import subprocess
import traceback
import unicodedata

from .logging_utils import append_log, timestamp
from .manifest import Operation
from .output_decode import decode_process_bytes
from .paths import ProjectPaths


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float], None]
CancelCallback = Callable[[], bool]


@dataclass
class JobContext:
    paths: ProjectPaths
    operation: Operation
    log_file: Path
    report_dir: Path
    log_callback: LogCallback | None = None
    progress_callback: ProgressCallback | None = None
    cancel_callback: CancelCallback | None = None

    def log(self, message: str) -> None:
        append_log(self.log_file, message)
        if self.log_callback:
            self.log_callback(message)

    def progress(self, value: float) -> None:
        if self.progress_callback:
            self.progress_callback(max(0.0, min(1.0, float(value))))

    def cancelled(self) -> bool:
        return bool(self.cancel_callback and self.cancel_callback())


@dataclass
class JobResult:
    ok: bool
    message: str
    data: dict[str, Any]


def utf8_subprocess_env(extra: dict[str, str] | None = None) -> dict[str, str]:
    env = os.environ.copy()
    if extra:
        env.update(extra)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    env.pop("NO_COLOR", None)
    env["CLICOLOR"] = "1"
    env["CLICOLOR_FORCE"] = "1"
    env["FORCE_COLOR"] = "1"
    env["AUDION_GUI_TERMINAL"] = "1"
    return env


def unbuffer_python_command(command: list[Any]) -> list[Any]:
    if not command:
        return command
    if not is_python_command(command):
        return command
    if len(command) > 1 and command[1] == "-u":
        return command
    return [command[0], "-u", *command[1:]]


def is_python_command(command: list[Any]) -> bool:
    if not command:
        return False
    executable = Path(str(command[0])).name.lower()
    return executable in {"python", "python.exe", "pythonw.exe"}


def _decode_score(text: str) -> int:
    score = 0
    suspicious = set("�¤ҐЎўЋЌЊЉЏ╬╪╨╤╥╫╘╙╒╓╔╗╝╚")
    common_cyrillic = set("оеаинтсрвлкмдпуяызьгзбчйхжюшцщэфъёОЕАИНТСРВЛКМДПУЯЫЗЬГЗБЧЙХЖЮШЦЩЭФЪЁ")
    for char in text:
        if char in "\r\n\t":
            continue
        if char in suspicious:
            score -= 12
            continue
        codepoint = ord(char)
        category = unicodedata.category(char)
        if char in common_cyrillic:
            score += 4
        elif "CYRILLIC" in unicodedata.name(char, ""):
            score += 2
        elif char.isascii() and (char.isprintable() or char == " "):
            score += 1
        elif 0x2500 <= codepoint <= 0x257F:
            score -= 10
        elif category.startswith("C"):
            score -= 10
        elif category.startswith(("L", "N", "P", "Z", "S")):
            score += 1
    return score


def _decode_with_encoding(data: bytes, encoding: str) -> str | None:
    try:
        return data.decode(encoding, errors="strict")
    except (LookupError, UnicodeDecodeError):
        return None


def _utf16_candidate(data: bytes) -> str | None:
    sample = data[:200]
    if len(sample) < 4 or b"\x00" not in sample:
        return None
    even_nulls = sample[0::2].count(0)
    odd_nulls = sample[1::2].count(0)
    if odd_nulls >= 2 and odd_nulls > even_nulls * 2:
        return "utf-16-le"
    if even_nulls >= 2 and even_nulls > odd_nulls * 2:
        return "utf-16-be"
    return None


def decode_subprocess_bytes(data: bytes) -> str:
    return decode_process_bytes(data)


SPINNER_FRAME_CHARS = set("-\\|/ \t")


def _is_spinner_only_line(line: str) -> bool:
    stripped = line.strip()
    return bool(stripped) and all(char in SPINNER_FRAME_CHARS for char in line)


def decoded_process_lines(raw_line: bytes | str) -> list[str]:
    text = str(raw_line) if isinstance(raw_line, str) else decode_process_bytes(raw_line)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines: list[str] = []
    for part in text.split("\n"):
        line = part.rstrip()
        if not line or _is_spinner_only_line(line):
            continue
        lines.append(line)
    return lines


def popen_gui_command(command: list[Any], *, cwd: Path) -> subprocess.Popen[Any]:
    common: dict[str, Any] = {
        "cwd": cwd,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.STDOUT,
        "env": utf8_subprocess_env(),
        **hidden_subprocess_kwargs(),
    }
    if is_python_command(command):
        return subprocess.Popen(
            command,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            **common,
        )
    return subprocess.Popen(
        command,
        text=False,
        bufsize=0,
        **common,
    )


def iter_subprocess_lines(process: subprocess.Popen[Any], command: list[Any]):
    stream = process.stdout
    if stream is None:
        return
    if is_python_command(command):
        for line in stream:
            for decoded_line in decoded_process_lines(str(line)):
                yield decoded_line
        return
    for raw_line in stream:
        yield from decoded_process_lines(raw_line if isinstance(raw_line, str) else bytes(raw_line))


def hidden_subprocess_startupinfo() -> subprocess.STARTUPINFO | None:
    if os.name != "nt" or not hasattr(subprocess, "STARTUPINFO"):
        return None
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    return startupinfo


def hidden_subprocess_creationflags() -> int:
    if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        return int(subprocess.CREATE_NO_WINDOW)
    return 0


def hidden_subprocess_kwargs() -> dict[str, Any]:
    return {
        "startupinfo": hidden_subprocess_startupinfo(),
        "creationflags": hidden_subprocess_creationflags(),
    }


def _load_callable(service: str) -> Callable[[JobContext], Any]:
    module_name, function_name = service.split(":", 1)
    module = importlib.import_module(module_name)
    return getattr(module, function_name)


def execute_operation(
    paths: ProjectPaths,
    operation: Operation,
    log_callback: LogCallback | None = None,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
) -> JobResult:
    run_stamp = timestamp().replace(":", "-")
    log_file = paths.logs / f"{run_stamp}_{operation.id}.log"
    report_dir = paths.report / f"{run_stamp}_{operation.id}"
    report_dir.mkdir(parents=True, exist_ok=True)
    context = JobContext(paths, operation, log_file, report_dir, log_callback, progress_callback, cancel_callback)

    try:
        context.log(f"Starting operation: {operation.id}")
        if operation.parameters:
            context.log(f"Parameters: {json.dumps(operation.parameters, ensure_ascii=False, sort_keys=True)}")
        context.progress(0.0)
        result = _load_callable(operation.service)(context)
        context.progress(1.0)
        context.log(f"Finished operation: {operation.id}")

        if isinstance(result, dict):
            return JobResult(True, "Operation finished.", result)
        return JobResult(True, str(result or "Operation finished."), {})

    except Exception as exc:
        context.log(traceback.format_exc())
        return JobResult(False, f"{exc.__class__.__name__}: {exc}", {})
