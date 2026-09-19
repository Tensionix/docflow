#!/usr/bin/env python3
"""
Audit rules kept as data.

The rules of docx_audit_processor live in config/rules, one YAML file per topic:
units.yaml, house.yaml, abbreviations.yaml, acronyms.yaml, addresses.yaml. A rule
is a regular expression with named groups and a replacement template, so a new
unit or abbreviation is a line in a file, not a change of code.

A rule carries its own tests. Every "fix" pair must come out exactly as written,
every "keep" text must stay unchanged and unreported. A rule whose examples fail,
or a fix rule with no examples at all, still runs, but only reports: an untested
rule never edits a document.

Check every example in both norms: python audit_rule_files.py
"""

from __future__ import annotations

# >>> audion CLI bootstrap >>>
import os as _os, sys as _sys
_HERE = _os.path.dirname(_os.path.abspath(__file__))
if _HERE not in _sys.path:
    _sys.path.insert(0, _HERE)
try:
    _sys.stdout.reconfigure(encoding="utf-8")
    _sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass
# <<< audion CLI bootstrap <<<

import argparse
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Match, Pattern

import yaml


RULES_DIR = Path(__file__).resolve().parents[1] / "config" / "rules"
# Files apply in this order whatever order they are asked in: the units go first,
# so the rules after them see "10 кв. м" and not "10кв.м".
RULE_SETS = ("units", "house", "abbreviations", "acronyms", "addresses")
DEFAULT_RULE_SETS = ("units", "house", "abbreviations", "acronyms")
NORMS = ("rules", "gost")
DEFAULT_NORM = "rules"
ACTIONS = ("fix", "suggest")
CONTEXT_WINDOW = 70

NBSP = "\u00A0"

# Building blocks a pattern names in braces: {NUM}, {END}, {SP}, {NBSP}, {list:name}.
NUM_RE = r"(?P<num>(?<![\wА-Яа-яЁё])[-+]?\d+(?:[ \u00A0]?\d{3})*(?:[,.]\d+)?)"
AFTER_UNIT_RE = r"(?=$|[\s\u00A0,.;:!?()\[\]{}<>«»\"'“”‘’/-])"
MACROS = {
    "NUM": NUM_RE,
    "END": AFTER_UNIT_RE,
    "SP": r"[ \u00A0]",
    "NBSP": r"\u00A0",
}
MACRO_RE = re.compile(r"\{(NUM|END|SP|NBSP|list:[A-Za-z0-9_]+|address:[a-z_]+)\}")
LEFTOVER_MACRO_RE = re.compile(r"\{[A-Za-z_:][A-Za-z0-9_:]*\}")
# A template writes {group}, {group|canon}, {group|address}, {group?text} or {NBSP}.
TEMPLATE_RE = re.compile(r"\{(?:(NBSP)|([A-Za-z_]\w*)(?:\|(\w+)|\?([^{}]*))?)\}")
TEMPLATE_FILTERS = ("canon", "address")

# Words the address engine may write another way: types of settlements, streets and
# house parts. Every other word, number and house letter must come out as written.
ADDRESS_TYPE_WORDS = frozenset({
    "г", "гор", "город", "с", "село", "д", "дер", "деревня", "п", "пос", "поселок", "пгт", "городского", "типа",
    "рп", "р", "х", "хутор", "ст", "ца", "станица", "станция", "сл", "слобода",
    "ул", "улица", "пер", "переулок", "пр", "кт", "т", "просп", "проспект", "проезд", "б", "бульв", "бульвар",
    "наб", "набережная", "пл", "площадь", "ш", "шоссе", "туп", "тупик", "мкр", "микрорайон", "о", "остров",
    "проул", "тер", "территория", "км",
    "дом", "зд", "здание", "стр", "строение", "к", "кор", "корп", "корпус", "лит", "литера", "вл", "влд", "владение",
})
HOUSE_LETTER_RE = re.compile(rf"(\d+)[ {NBSP}]?([А-Яа-яЁё])(?![А-Яа-яЁё\d])")
# The slash counts too: "д. 3/1" read as "д. 3, к. 1" is a new meaning, not a new spelling.
ADDRESS_TOKEN_RE = re.compile(r"\d+[А-Яа-яЁё]?(?![А-Яа-яЁё\d])|[А-Яа-яЁёA-Za-z]+|\d+|/")


class RuleError(ValueError):
    """A rule file or a rule that cannot be used."""


@dataclass(frozen=True)
class RuleCase:
    pattern: Pattern[str]
    template: str


@dataclass
class AuditRule:
    code: str
    title: str
    severity: str
    description: str
    cases: tuple[RuleCase, ...]
    fixable: bool = True
    rule_set: str = ""
    status: str = "implemented"
    title_en: str = ""
    norms: tuple[str, ...] = NORMS
    canon: dict[str, str] = field(default_factory=dict)
    fix_examples: dict[str, str] = field(default_factory=dict)
    keep_examples: tuple[str, ...] = ()
    sources: tuple[str, ...] = ()
    broken: str = ""

    @property
    def edits(self) -> bool:
        """Only a fix rule that passes its own examples may change a document."""
        return self.fixable and not self.broken

    @property
    def mode(self) -> str:
        return "FIX" if self.edits else "SUGGEST"

    def replacement_for(self, case: RuleCase, match: Match[str]) -> str:
        return carry_spaces(match.group(0), render(case.template, match, self.canon))


@dataclass
class RuleHit:
    rule: AuditRule
    action: str
    before: str
    after: str
    context: str
    start: int
    end: int


@dataclass
class RuleBook:
    norm: str
    sets: tuple[str, ...]
    rules: list[AuditRule] = field(default_factory=list)
    files: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def canon_key(value: str) -> str:
    return re.sub(r"[ \u00A0]+", "", value).lower().replace("ё", "е")


def expand_pattern(pattern: str, lists: dict[str, list[str]]) -> str:
    def one_list(match: Match[str]) -> str:
        name = match.group(1)
        if not name.startswith("list:"):
            return match.group(0)
        key = name[len("list:"):]
        if key not in lists:
            raise RuleError(f"нет списка {key!r} в разделе lists")
        return "(?:" + "|".join(lists[key]) + ")"

    def one_macro(match: Match[str]) -> str:
        name = match.group(1)
        if name.startswith("address:"):
            return address_types(name[len("address:"):])
        return MACROS.get(name, match.group(0))

    # Lists first: an item may use {SP} itself.
    expanded = MACRO_RE.sub(one_list, pattern)
    expanded = MACRO_RE.sub(one_macro, expanded)
    leftover = LEFTOVER_MACRO_RE.search(expanded)
    if leftover:
        raise RuleError(f"неизвестная вставка {leftover.group(0)}; есть {{NUM}}, {{END}}, {{SP}}, {{NBSP}}, {{list:имя}}")
    return expanded


def render(template: str, match: Match[str], canon: dict[str, str]) -> str:
    def one(part: Match[str]) -> str:
        if part.group(1):
            return NBSP
        value = match.group(part.group(2)) or ""
        if part.group(4) is not None:
            return part.group(4) if value else ""
        if part.group(3) == "canon":
            return canon.get(canon_key(value), value)
        if part.group(3) == "address":
            return address_writing(value)
        return value

    return TEMPLATE_RE.sub(one, template)


def _address_engine():
    """The address naming modules of Audion Address Processor, copied to address_engine."""
    try:
        from address_engine import address_slots, audion_address_core
    except ImportError as exc:
        raise RuleError(f"адресный движок address_engine не загрузился: {exc}") from exc
    return audion_address_core, address_slots


def address_types(name: str) -> str:
    """{address:settlement} and {address:street}: the type words the address engine knows.

    Lower case only, as an address writes them: "Г.Петров" and "Д. Иванов" are
    initials, "Ул." opens a sentence - neither is an address type.
    """
    core, _ = _address_engine()
    patterns = {
        "settlement": core.SETTLEMENT_TYPE_PATTERN,
        "street": f"{core.STREET_TYPE_PATTERN}|{core.TERRITORY_TYPE_PATTERN}",
    }
    if name not in patterns:
        raise RuleError(f"неизвестная вставка {{address:{name}}}; есть {{address:settlement}} и {{address:street}}")
    return f"(?:{patterns[name]})"


def _address_names_and_numbers(text: str) -> list[str]:
    glued = HOUSE_LETTER_RE.sub(r"\1\2", text)
    return [token for token in ADDRESS_TOKEN_RE.findall(glued) if token.lower().replace("ё", "е") not in ADDRESS_TYPE_WORDS]


def address_writing(text: str) -> str:
    """The address written the way the Audion Address Processor engine writes it.

    Only the writing may change: abbreviations of settlement, street and house
    types, dots, commas and spaces. Names, numbers and their order must come out
    as written, and a house letter keeps its case; if the engine would change any
    of them, the text stays as it is.
    """
    _, slots = _address_engine()
    source = text.replace(NBSP, " ")
    rendered = slots.render_address_core(slots.build_address_parts(source))
    if not rendered:
        return text
    for number, letter in HOUSE_LETTER_RE.findall(source):
        rendered = re.sub(rf"(?<!\d){number}([ {NBSP}]?){letter.swapcase()}(?![А-Яа-яЁё\d])", rf"{number}\g<1>{letter}", rendered)
    if _address_names_and_numbers(source) != _address_names_and_numbers(rendered):
        return text
    return rendered


def _space_gaps(text: str) -> dict[int, str]:
    """The spaces of a text, keyed by how many other characters stand before them."""
    gaps: dict[int, str] = {}
    seen = 0
    for char in text:
        if char in (" ", NBSP):
            gaps[seen] = gaps.get(seen, "") + char
        else:
            seen += 1
    return gaps


def _common_prefix(left: str, right: str) -> int:
    size = 0
    for a, b in zip(left, right):
        if a != b:
            break
        size += 1
    return size


def carry_spaces(before: str, after: str) -> str:
    """Keep a non-breaking space where the text had one.

    A template writes ordinary spaces. Where the original held a non-breaking
    space at the same place - between the same characters, counted from the
    start or from the end - it stays: authors put them there on purpose. A space
    the template adds where there was none stays ordinary; only {NBSP} writes a
    non-breaking one.
    """
    if " " not in after or NBSP not in before:
        return after
    gaps = _space_gaps(before)
    old = before.replace(" ", "").replace(NBSP, "")
    new = after.replace(" ", "").replace(NBSP, "")
    prefix = _common_prefix(old, new)
    suffix = min(_common_prefix(old[::-1], new[::-1]), min(len(old), len(new)) - prefix)
    result: list[str] = []
    seen = 0
    for char in after:
        if char == " ":
            from_end = len(new) - seen
            if seen <= prefix and NBSP in gaps.get(seen, ""):
                char = NBSP
            elif from_end <= suffix and NBSP in gaps.get(len(old) - from_end, ""):
                char = NBSP
        elif char != NBSP:
            seen += 1
        result.append(char)
    return "".join(result)


def _context(text: str, start: int, end: int, window: int) -> str:
    return text[max(0, start - window):min(len(text), end + window)]


def apply_rules(
    rules: Iterable[AuditRule],
    text: str,
    *,
    fix: bool,
    force: bool = False,
    window: int = CONTEXT_WINDOW,
) -> tuple[str, list[RuleHit]]:
    """Run the rules over one piece of text, in order.

    A rule that edits sees the text the rules before it left; with fix off it only
    plans and the text stays as it was. force makes every rule edit: an example
    checks a suggestion by what it would write.
    """
    current = text
    hits: list[RuleHit] = []
    for rule in rules:
        edits = force or rule.edits
        for case in rule.cases:
            if edits:

                def replace(match: Match[str], audit_rule: AuditRule = rule, rule_case: RuleCase = case) -> str:
                    before = match.group(0)
                    after = audit_rule.replacement_for(rule_case, match)
                    if after != before:
                        hits.append(
                            RuleHit(
                                rule=audit_rule,
                                action="FIX" if fix else "PLAN",
                                before=before,
                                after=after,
                                context=_context(current, match.start(), match.end(), window),
                                start=match.start(),
                                end=match.end(),
                            )
                        )
                    return after if fix else before

                current = case.pattern.sub(replace, current)
                continue

            for match in case.pattern.finditer(current):
                before = match.group(0)
                after = rule.replacement_for(case, match)
                if after != before:
                    hits.append(
                        RuleHit(
                            rule=rule,
                            action="SUGGEST",
                            before=before,
                            after=after,
                            context=_context(current, match.start(), match.end(), window),
                            start=match.start(),
                            end=match.end(),
                        )
                    )
    return current, hits


def example_failures(rule: AuditRule) -> list[str]:
    problems: list[str] = []
    if rule.fixable and not rule.fix_examples:
        problems.append("нет примеров fix")
    for before, expected in rule.fix_examples.items():
        got, _ = apply_rules([rule], before, fix=True, force=True)
        if got != expected:
            problems.append(f"fix {before!r} -> {got!r}, ожидалось {expected!r}")
    for text in rule.keep_examples:
        got, hits = apply_rules([rule], text, fix=True, force=True)
        if got != text or hits:
            problems.append(f"keep {text!r} -> {got!r}")
    return problems


def _string(value: object, where: str) -> str:
    if not isinstance(value, str):
        raise RuleError(f"{where}: ожидается строка")
    return value


def _compile_rule(raw: object, *, rule_set: str, lists: dict[str, list[str]]) -> AuditRule:
    if not isinstance(raw, dict):
        raise RuleError("правило должно быть словарём")
    code = str(raw.get("code") or "").strip()
    if not code:
        raise RuleError("правило без code")
    action = str(raw.get("action") or "fix").strip()
    if action not in ACTIONS:
        raise RuleError(f"{code}: action {action!r}, ожидается fix или suggest")
    norms = tuple(str(norm) for norm in (raw.get("norms") or NORMS))
    unknown_norms = [norm for norm in norms if norm not in NORMS]
    if unknown_norms:
        raise RuleError(f"{code}: неизвестная норма {', '.join(unknown_norms)}; есть {', '.join(NORMS)}")

    flags = re.IGNORECASE if raw.get("ignore_case") else 0
    raw_cases = raw.get("cases") or [{"match": raw.get("match"), "replace": raw.get("replace")}]
    cases: list[RuleCase] = []
    for index, item in enumerate(raw_cases, start=1):
        where = f"{code}, вариант {index}" if raw.get("cases") else code
        if not isinstance(item, dict):
            raise RuleError(f"{where}: ожидаются match и replace")
        source = _string(item.get("match"), f"{where} match")
        template = _string(item.get("replace"), f"{where} replace")
        try:
            pattern = re.compile(expand_pattern(source, lists), flags)
        except re.error as exc:
            raise RuleError(f"{where}: регулярное выражение не собирается: {exc}") from exc
        except RuleError as exc:
            raise RuleError(f"{where}: {exc}") from exc
        for part in TEMPLATE_RE.finditer(template):
            name, template_filter = part.group(2), part.group(3)
            if name and name not in pattern.groupindex:
                raise RuleError(f"{where}: в replace стоит {{{name}}}, а в match нет группы {name}")
            if template_filter and template_filter not in TEMPLATE_FILTERS:
                raise RuleError(f"{where}: неизвестный фильтр |{template_filter}")
            if template_filter == "address":
                try:
                    _address_engine()
                except RuleError as exc:
                    raise RuleError(f"{where}: {exc}") from exc
        cases.append(RuleCase(pattern=pattern, template=template))

    return AuditRule(
        code=code,
        title=str(raw.get("title_ru") or raw.get("title_en") or code),
        severity=str(raw.get("severity") or "minor"),
        description=str(raw.get("description_ru") or ""),
        cases=tuple(cases),
        fixable=action == "fix",
        rule_set=rule_set,
        status=str(raw.get("status") or "implemented"),
        title_en=str(raw.get("title_en") or ""),
        norms=norms,
        canon={canon_key(str(key)): str(value) for key, value in (raw.get("canon") or {}).items()},
        fix_examples={str(key): str(value) for key, value in (raw.get("fix") or {}).items()},
        keep_examples=tuple(str(value) for value in (raw.get("keep") or [])),
        sources=tuple(str(value) for value in (raw.get("sources") or [])),
    )


def parse_rule_sets(value: str | Iterable[str] | None) -> tuple[str, ...]:
    if value is None:
        return DEFAULT_RULE_SETS
    items = value.split(",") if isinstance(value, str) else list(value)
    names = {item.strip() for item in items if item and item.strip()}
    unknown = sorted(names - set(RULE_SETS))
    if unknown:
        raise RuleError(f"неизвестные файлы правил: {', '.join(unknown)}; есть {', '.join(RULE_SETS)}")
    return tuple(name for name in RULE_SETS if name in names)


def load_rule_book(
    sets: Iterable[str] = DEFAULT_RULE_SETS,
    norm: str = DEFAULT_NORM,
    *,
    rules_dir: Path = RULES_DIR,
    check: bool = True,
) -> RuleBook:
    """Read the chosen rule files for one norm.

    A file or a rule that cannot be read is skipped with a warning, and the rest
    still run: one typo in a file must not switch the whole audit off.
    """
    if norm not in NORMS:
        raise RuleError(f"неизвестная норма {norm!r}; есть {', '.join(NORMS)}")
    chosen = parse_rule_sets(list(sets))
    book = RuleBook(norm=norm, sets=chosen)
    seen: dict[str, str] = {}
    for name in chosen:
        path = Path(rules_dir) / f"{name}.yaml"
        try:
            data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        except (OSError, yaml.YAMLError) as exc:
            book.warnings.append(f"{path.name}: файл не прочитан: {exc}")
            continue
        if not isinstance(data, dict):
            book.warnings.append(f"{path.name}: ожидается словарь с разделом rules")
            continue
        book.files.append(path)
        lists = {str(key): [str(item) for item in values or []] for key, values in (data.get("lists") or {}).items()}
        for raw in data.get("rules") or []:
            try:
                rule = _compile_rule(raw, rule_set=name, lists=lists)
            except RuleError as exc:
                book.warnings.append(f"{path.name}: {exc}; правило пропущено")
                continue
            if norm not in rule.norms:
                continue
            if rule.code in seen:
                book.warnings.append(f"{path.name}: {rule.code} уже есть в {seen[rule.code]} для нормы {norm}; повтор пропущен")
                continue
            seen[rule.code] = path.name
            if check:
                problems = example_failures(rule)
                if problems:
                    rule.broken = "; ".join(problems)
                    if rule.fixable:
                        book.warnings.append(f"{path.name}: {rule.code} работает только в отчёт: {rule.broken}")
                    else:
                        book.warnings.append(f"{path.name}: {rule.code}: {rule.broken}")
            book.rules.append(rule)
    return book


def main() -> int:
    parser = argparse.ArgumentParser(description="Check the fix and keep examples of every audit rule file in both norms.")
    parser.add_argument("--rules-dir", default=str(RULES_DIR), help="Folder with the rule files")
    args = parser.parse_args()
    problems = 0
    for norm in NORMS:
        book = load_rule_book(RULE_SETS, norm, rules_dir=Path(args.rules_dir))
        examples = sum(len(rule.fix_examples) + len(rule.keep_examples) for rule in book.rules)
        print(f"[{norm}] файлов: {len(book.files)}, правил: {len(book.rules)}, примеров: {examples}")
        for warning in book.warnings:
            print(f"[WARN] {warning}")
        problems += len(book.warnings)
    print("[OK] Все примеры прошли." if not problems else f"[FAIL] Замечаний: {problems}")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
