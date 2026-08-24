#!/usr/bin/env python3
"""
Audion DocFlow - Common helpers (no network).
All user-facing output is English by default.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Tuple, Dict, Any

WS_RE = re.compile(r"\s+")
NBSP = "\u00A0"

def norm_space(s: str) -> str:
    s = (s or "").replace(NBSP, " ")
    s = WS_RE.sub(" ", s)
    return s.strip()

def norm_key(s: str) -> str:
    s = norm_space(s)
    s = s.replace("Ё", "Е").replace("ё", "е")
    return s.lower()

def safe_mkdir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def write_json_file(path: Path, payload: Any) -> None:
    safe_mkdir(path.parent)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

def find_docx_files(root: Path) -> List[Path]:
    if root.is_file():
        return [root] if root.suffix.lower() == ".docx" and not root.name.startswith("~$") else []
    return sorted(
        p for p in root.rglob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
    )

def rel_posix(path: Path, root: Path) -> str:
    base = root.parent if root.is_file() else root
    return path.relative_to(base).as_posix()

def mirrored_output_path(path: Path, input_root: Path, output_root: Path) -> Path:
    base = input_root.parent if input_root.is_file() else input_root
    return output_root / path.relative_to(base)

def md_escape(s: str) -> str:
    # Minimal escaping for Markdown tables
    s = str(s) if s is not None else ""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("|", "\\|")
    s = s.replace("\n", "\\n")
    return s

def truncate(s: str, n: int = 200) -> str:
    s = s if isinstance(s, str) else str(s)
    return s if len(s) <= n else s[:n-3] + "..."

@dataclass
class DiffField:
    field: str
    a: Any
    b: Any
