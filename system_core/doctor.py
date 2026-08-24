from __future__ import annotations

from pathlib import Path
import importlib
import platform
import sys

REQUIRED_MODULES = [
    ("requests", "requests"),
    ("tqdm", "tqdm"),
    ("pydantic", "pydantic"),
    ("rich", "rich"),
    ("nicegui", "nicegui"),
    ("webview", "pywebview"),
    ("yaml", "pyyaml"),
    ("docx", "python-docx"),
    ("lxml", "lxml"),
    ("markdown_it", "markdown-it-py"),
]

OPTIONAL_MODULES = [
    ("openai", "openai"),
    ("google.genai", "google-genai"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("fitz", "pymupdf"),
    ("pptx", "python-pptx"),
]


def check_module(import_name: str) -> tuple[bool, str]:
    try:
        mod = importlib.import_module(import_name)
        version = getattr(mod, "__version__", "unknown")
        return True, str(version)
    except Exception as exc:
        return False, exc.__class__.__name__


def detect_python_mode(root: Path) -> str:
    if (root / "runtime" / "python.exe").exists():
        return "portable-runtime"
    if (root / "runtime" / "python" / "python.exe").exists():
        return "portable-runtime"
    return "system-python"


def check_cmd_encoding(root: Path) -> tuple[bool, list[tuple[str, bool, str]]]:
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

    try:
        from system_core.core.cmd_encoding import check_cmd_files
    except Exception as exc:
        return False, [("(loader)", False, f"{exc.__class__.__name__}: {exc}")]

    rows: list[tuple[str, bool, str]] = []
    all_ok = True
    for result in check_cmd_files(root):
        try:
            relative = str(result.path.resolve().relative_to(root.resolve()))
        except ValueError:
            relative = str(result.path)

        detail = result.summary()
        if result.error:
            detail = f"{detail} {result.error}"
        rows.append((relative, result.ok, detail))
        if not result.ok:
            all_ok = False

    return all_ok, rows


def main() -> int:
    root = Path(__file__).resolve().parents[1]

    print("======================================================================")
    print("AUDION DOCFLOW - DOCTOR")
    print("======================================================================")
    print(f"Project root : {root}")
    print(f"Executable   : {sys.executable}")
    print(f"Python       : {sys.version.split()[0]}")
    print(f"Python mode  : {detect_python_mode(root)}")
    print(f"Platform     : {platform.platform()}")
    print()

    failed = False

    print("[Required modules]")
    for import_name, package_name in REQUIRED_MODULES:
        ok, detail = check_module(import_name)
        status = "OK" if ok else "FAIL"
        print(f"  - {package_name:<18} : {status:<4} {detail}")
        if not ok:
            failed = True

    print()
    print("[Optional modules]")
    for import_name, package_name in OPTIONAL_MODULES:
        ok, detail = check_module(import_name)
        status = "OK" if ok else "MISS"
        print(f"  - {package_name:<18} : {status:<4} {detail}")

    print()
    print("[CMD encoding]")
    cmd_ok, cmd_rows = check_cmd_encoding(root)
    if not cmd_rows:
        print("  (no CMD files found)")
    else:
        for relative, result_ok, detail in cmd_rows:
            status = "OK" if result_ok else "FAIL"
            print(f"  - {relative:<58} : {status:<4} {detail}")
    if not cmd_ok:
        failed = True

    print()
    if failed:
        print("[RESULT] One or more required modules are missing.")
        return 1

    print("[RESULT] Required environment looks good.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
