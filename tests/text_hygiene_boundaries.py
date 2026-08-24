from __future__ import annotations

from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
SYSTEM_CORE = ROOT / "system_core"
if str(SYSTEM_CORE) not in sys.path:
    sys.path.insert(0, str(SYSTEM_CORE))

from docx_text_hygiene_fix import find_text_issues as fix_find_text_issues, fix_text
from docx_text_hygiene_scan import find_text_issues as scan_find_text_issues


def classes(findings: list[dict[str, object]]) -> list[str]:
    return [str(item.get("class_id") or "") for item in findings]


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def test_text_hygiene_owns_mechanical_ranges() -> None:
    text = "Это  текст ,а также мягкий\u00ADперенос."
    findings = scan_find_text_issues("word/document.xml", 7, text, check_dot=True)
    found = classes(findings)
    assert_true("double_space" in found, "double_space должен принадлежать text hygiene")
    assert_true("space_before_punct" in found, "space_before_punct должен принадлежать text hygiene")
    assert_true("missing_after_punct" in found, "missing_after_punct должен принадлежать text hygiene")
    assert_true("soft_hyphen" in found, "soft_hyphen должен принадлежать text hygiene")
    assert_true(all("start" in item and "end" in item and "text_node_index" in item for item in findings), "каждая находка должна иметь диапазон")
    assert_true(fix_text("Это  текст ,а", fix_dot=True) == "Это текст, а", "fix_text должен исправлять механическую гигиену")


def test_text_hygiene_does_not_own_audit_dot_cases() -> None:
    text = "12кв.м, 5куб.м, 100руб. г.Тюмень ул.Ленина табл.1 рис.2"
    scan_findings = scan_find_text_issues("word/document.xml", 3, text, check_dot=True)
    fix_findings = fix_find_text_issues("word/document.xml", 3, text, fix_dot=True)
    assert_true("missing_after_dot" not in classes(scan_findings), "scan не должен забирать audit/морфологические dot-cases")
    assert_true("missing_after_dot" not in classes(fix_findings), "fix не должен забирать audit/морфологические dot-cases")
    assert_true(fix_text(text, fix_dot=True) == text, "fix_text не должен менять audit/морфологические dot-cases")


def main() -> int:
    test_text_hygiene_owns_mechanical_ranges()
    test_text_hygiene_does_not_own_audit_dot_cases()
    print("[OK] text hygiene boundary tests passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
