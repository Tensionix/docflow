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
    text = "12кв.м, 5куб.м, 100руб. г.Энск ул.Садовая табл.1 рис.2"
    scan_findings = scan_find_text_issues("word/document.xml", 3, text, check_dot=True)
    fix_findings = fix_find_text_issues("word/document.xml", 3, text, fix_dot=True)
    assert_true("missing_after_dot" not in classes(scan_findings), "scan не должен забирать audit/морфологические dot-cases")
    assert_true("missing_after_dot" not in classes(fix_findings), "fix не должен забирать audit/морфологические dot-cases")
    assert_true(fix_text(text, fix_dot=True) == text, "fix_text не должен менять audit/морфологические dot-cases")


def test_text_hygiene_keeps_numbers_addresses_and_nbsp() -> None:
    nbsp = chr(0xA0)
    kept = [
        f"за последние 5{nbsp}лет, №{nbsp}80 от{nbsp}05.11.2014",
        "обеспеченностью 0,98 и 0,92",
        "ЗУ 72:17:1313004:13058, в 10:30, масштаб 1:500",
        "см. https://example.org и C:" + chr(92) + "Data",
        "Правда?! Да!",
        f"слово{nbsp},",
    ]
    for text in kept:
        assert_true(fix_text(text, fix_dot=False) == text, f"fix_text не должен менять: {text!r}")
        found = classes(scan_find_text_issues("word/document.xml", 1, text, check_dot=False))
        assert_true("missing_after_punct" not in found and "space_before_punct" not in found,
                    f"scan не должен находить ошибку в: {text!r}")
    assert_true(fix_text("текст,а итог:100", fix_dot=False) == "текст, а итог: 100",
                "после запятой и двоеточия перед словом или числом пробел нужен")


def test_dot_spacing_keeps_codes_and_registry_abbreviations() -> None:
    kept = [
        "Горелка P61M-PR.S.RU.A.8.32.E.A.",
        "Горелка P630M-MG.PR.SR.RU.A.8.50.EC",
        "кг.у.т./Гкал",
        "расход топлива за 2024 год, т.у.т/год",
        "составляет 3 077,22 м.п.",
        "деятельности, у.е.",
        "e.g. text",
    ]
    for text in kept:
        assert_true(fix_text(text, fix_dot=True) == text, f"fix_text не должен менять: {text!r}")
        found = classes(scan_find_text_issues("word/document.xml", 1, text, check_dot=True))
        assert_true("missing_after_dot" not in found, f"scan не должен находить пробел после точки в: {text!r}")
    assert_true(fix_text("т.е.", fix_dot=True) == "т. е.", "составное сокращение вне реестра по-прежнему получает пробел")
    assert_true(fix_text("Итого.Далее", fix_dot=True) == "Итого. Далее", "пробел после точки перед словом по-прежнему нужен")


def test_hygiene_keeps_initials_split_numbers_and_quotes() -> None:
    kept_with_dot = [
        "имени В.К. Арсеньева",
        "Ф.И.О.",
        "директор - Киселев И.И.",
    ]
    for text in kept_with_dot:
        assert_true(fix_text(text, fix_dot=True) == text, f"fix_text не должен менять: {text!r}")
    assert_true(fix_text("И.И.Иванов", fix_dot=True) == "И.И. Иванов", "пробел ставится только перед фамилией")
    kept = [
        ",4 коек на",
        ",1 тыс. ",
        "«Музей для всех!», проводят",
        "бухта Прогулочная, , пляж",
    ]
    for text in kept:
        assert_true(fix_text(text, fix_dot=False) == text, f"fix_text не должен менять: {text!r}")
    assert_true(fix_text("ул. Первая,6)", fix_dot=False) == "ул. Первая, 6)", "пробел после запятой перед номером дома нужен")
    assert_true(fix_text("номеров) , центр", fix_dot=False) == "номеров), центр", "пробел перед запятой убирается")


def main() -> int:
    test_text_hygiene_owns_mechanical_ranges()
    test_text_hygiene_does_not_own_audit_dot_cases()
    test_text_hygiene_keeps_numbers_addresses_and_nbsp()
    test_dot_spacing_keeps_codes_and_registry_abbreviations()
    test_hygiene_keeps_initials_split_numbers_and_quotes()
    print("[OK] text hygiene boundary tests passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
