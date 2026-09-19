from __future__ import annotations

from collections import defaultdict
from difflib import SequenceMatcher
from typing import Iterable
import re

try:
    from rapidfuzz import fuzz, process
except Exception:  # pragma: no cover - optional acceleration
    fuzz = None
    process = None


CITY_NAME = ""

STREET_WORD_ALIASES = {
    "совесткая": "советская",
    "советска": "советская",
    "совсетов": "советов",
    "украинска": "украинская",
}

STREET_TYPE_PATTERN = (
    r"улица|ул\.?|переулок|пер\.?|пр-кт\.?|пр-т\.?|проспект|просп\.?|"
    r"пр-д\.?|проезд|пр\.?|"
    r"бульвар|б-р\.?|бульв\.?|тракт|шоссе|ш\.?|тупик|туп\.?|"
    r"аллея|линия|набережная|наб\.?|площадь|пл\.?|"
    r"квартал|кв-л\.?|микрорайон|мкр\.?|км\.?|автодорог\w*|"
    r"проул\.?"
)
TERRITORY_TYPE_PATTERN = r"тер\.?|территория|снт|сосн|днт|тсн|остров|о(?=\.)"
RAILWAY_ABBR_PATTERN = r"ж\s*[./-]?\s*-?\s*д\.?"
STANITSA_ABBR_PATTERN = r"ст\s*[\.-]?\s*ца\.?"
P_STATION_ABBR_PATTERN = r"п\.?\s*(?:/|\.)?\s*ст\.?"
RZD_ABBR_PATTERN = r"р\.?\s*зд\.?"
SETTLEMENT_COMPACT_PATTERN = (
    r"г\.?\s*п\.?|р\.?\s*п\.?|к\.?\s*п\.?|д\.?\s*п\.?|"
    r"п\.?\s*г\.?\s*т\.?|н\.?\s*п\.?"
)
SETTLEMENT_TYPE_PATTERN = (
    rf"п\s+{RAILWAY_ABBR_PATTERN}\s+ст\.?|п\s+{RAILWAY_ABBR_PATTERN}\s+{RZD_ABBR_PATTERN}|"
    rf"{RAILWAY_ABBR_PATTERN}\s+остановочн\w*\s+пункт|{RAILWAY_ABBR_PATTERN}\s+блокпост|"
    rf"{RAILWAY_ABBR_PATTERN}\s+будка|{RAILWAY_ABBR_PATTERN}\s+ветка|{RAILWAY_ABBR_PATTERN}\s+казарма|"
    rf"{RAILWAY_ABBR_PATTERN}\s+платформа|{RAILWAY_ABBR_PATTERN}\s+площадка|"
    rf"{RAILWAY_ABBR_PATTERN}\s+путевой\s+пост|{RAILWAY_ABBR_PATTERN}\s+{RZD_ABBR_PATTERN}|"
    rf"{RAILWAY_ABBR_PATTERN}\s+ст\.?|"
    rf"{P_STATION_ABBR_PATTERN}|п\s+станци[ия]|пос[её]лок\s+при\s+станции|пос[её]лок\s+станции|"
    r"городской\s+пос[её]лок|рабочий\s+пос[её]лок|курортный\s+пос[её]лок|"
    r"дачный\s+пос[её]лок|пос[её]лок\s+городского\s+типа|насел[её]нный\s+пункт|"
    rf"{SETTLEMENT_COMPACT_PATTERN}|{RZD_ABBR_PATTERN}|разъезд|"
    r"город|г\.?|пос[её]лок|пос\.?|п\.?|деревня|д\.?|село|с\.?|"
    rf"слобода|сл\.?|станица|{STANITSA_ABBR_PATTERN}|станция|ст\.?|хутор|х\.?|"
    r"улус|у\.?|местечко|м\.?|кишлак|к\.?|"
    r"аул|аал|арбан|починок|выселок|выселки|заимка|кордон|маяк|"
    r"погост|слободка|усадьба|лесоучасток|метеостанция"
)
STREET_OR_TERRITORY_RE = re.compile(
    rf"\b(?:{STREET_TYPE_PATTERN}|{TERRITORY_TYPE_PATTERN})\b",
    re.IGNORECASE,
)
MICRODISTRICT_RE = re.compile(r"\b(?:микрорайон|мкр\.?|квартал|кв-л\.?)\b", re.IGNORECASE)
SETTLEMENT_SEGMENT_RE = re.compile(
    rf"^\s*(?:{SETTLEMENT_TYPE_PATTERN})\s+(?P<name>[^,]+?)\s*$",
    re.IGNORECASE,
)
SETTLEMENT_LEADING_RE = re.compile(
    rf"^\s*(?:{SETTLEMENT_TYPE_PATTERN})\s+"
    rf"(?P<name>.+?)"
    rf"(?=\s+(?:{STREET_TYPE_PATTERN}|{TERRITORY_TYPE_PATTERN})\b|,|$)",
    re.IGNORECASE,
)
SETTLEMENT_ADMIN_RE = re.compile(
    rf"\b(?:"
    rf"городской\s+пос[её]лок|рабочий\s+пос[её]лок|курортный\s+пос[её]лок|"
    rf"дачный\s+пос[её]лок|пос[её]лок\s+городского\s+типа|насел[её]нный\s+пункт|"
    rf"пос[её]лок|деревня|село|слобода|станица|станция|хутор|улус|местечко|кишлак|"
    rf"{SETTLEMENT_COMPACT_PATTERN}|{RZD_ABBR_PATTERN}|{STANITSA_ABBR_PATTERN}"
    rf")\b",
    re.IGNORECASE,
)

OBJECT_WORD_RE = re.compile(
    r"(?:школа|сош|оош|ноо|лицей|гимназия|детск\w*\s+сад|садик|доу|"
    r"колледж|техникум|училищ\w*|университет|академия|институт|"
    r"общежит\w*|кампус|корпус|учрежден\w*|организац\w*|"
    r"school|kindergarten|college|university)",
    re.IGNORECASE,
)
PREMISE_MARKER_PATTERN = r"помещ\.?|помещение|пом\.?|кв\.?|квартира|офис|оф\.?"

HOUSE_TAIL_RE = re.compile(
    r"(?P<base>\d+)(?P<tail>\s*(?:"
    r"[а-яa-z](?!\d)|"
    r"[/\-]\s*\d+[а-яa-z]?|"
    r"[/\-]\s*[а-яa-z]|"
    r"[,;]?\s*(?:к\.?|кор\.?|корп\.?|корпус)\s*\d+[а-яa-z]?|"
    r"[,;]?\s*(?:с\.?|стр\.?|строение)\s*\d+[а-яa-z]?|"
    r"[,;]?\s*(?:лит\.?а?|литера)\s*[а-яa-z]"
    r")*)\s*$",
    re.IGNORECASE,
)


def normalize_house_mods(tail: str | None) -> str:
    text = str(tail or "").lower().replace("ё", "е")
    text = re.sub(r"(?:корпус|корп\.?|кор\.?|к\.)\s*", "к", text)
    text = re.sub(r"(?:строение|стр\.?|с\.)\s*", "с", text)
    text = re.sub(r"(?:литера|лит\.?)\s*", "лит", text)
    text = re.sub(r"[\s,;]+", "", text)

    mods = ""
    while text:
        if text[0] in "/-":
            token = _read_house_number_token(text[1:])
            if token:
                number, end = token
                mods += "к" + number
                text = text[end + 1 :]
                continue
            match = re.match(r"[/\-]([а-яa-z])", text)
            if match:
                mods += match.group(1)
                text = text[match.end() :]
                continue
            break
        if text.startswith("лит"):
            match = re.match(r"лит([а-яa-z])", text)
            if not match:
                break
            mods += "лит" + match.group(1)
            text = text[match.end() :]
            continue
        if text.startswith("к"):
            token = _read_house_number_token(text[1:])
            if token:
                number, end = token
                mods += "к" + number
                text = text[end + 1 :]
                continue
        if text.startswith("с"):
            token = _read_house_number_token(text[1:])
            if token:
                number, end = token
                mods += "с" + number
                text = text[end + 1 :]
                continue
        match = re.match(r"([а-яa-z])", text)
        if match:
            mods += match.group(1)
            text = text[match.end() :]
            continue
        break
    return _canonical_house_mods(mods)


def _canonical_house_mods(value: str) -> str:
    text = str(value or "")
    tokens: list[str] = []
    index = 0
    while index < len(text):
        if text.startswith("лит", index):
            match = re.match(r"лит([а-яa-z])", text[index:], flags=re.IGNORECASE)
            if match:
                tokens.append("лит" + match.group(1).lower())
                index += match.end()
                continue
        marker = text[index]
        if marker in {"к", "с"}:
            token = _read_house_number_token(text[index + 1 :])
            if token:
                number, end = token
                tokens.append(marker + number.lower())
                index += end + 1
                continue
        tokens.append(marker.lower())
        index += 1

    priority = {"к": 1, "с": 2}

    def sort_key(token: str) -> tuple[int, str]:
        if token.startswith("лит"):
            return (3, token)
        return (priority.get(token[:1], 0), token)

    return "".join(sorted(tokens, key=sort_key))


def _read_house_number_token(text: str) -> tuple[str, int] | None:
    match = re.match(r"(\d+)([а-яa-z])?", text, flags=re.IGNORECASE)
    if not match:
        return None
    number = match.group(1)
    suffix = match.group(2) or ""
    digits_end = len(number)
    if suffix and suffix.lower() == "л" and text[digits_end:].lower().startswith("лит"):
        return number, digits_end
    if suffix and suffix.lower() in {"к", "с"} and len(text) > digits_end + 1 and text[digits_end + 1].isdigit():
        return number, digits_end
    return number + suffix, match.end()


def compact_house_modifier_tokens(value: str | None) -> list[str]:
    text = str(value or "").lower().replace("ё", "е")
    tokens: list[str] = []
    index = 0
    while index < len(text):
        if text.startswith("лит", index):
            match = re.match(r"лит([а-яa-z])", text[index:], flags=re.IGNORECASE)
            if match:
                tokens.append("лит" + match.group(1).lower())
                index += match.end()
                continue
        marker = text[index]
        if marker in {"к", "с"}:
            token = _read_house_number_token(text[index + 1 :])
            if token:
                number, end = token
                tokens.append(marker + number.lower())
                index += end + 1
                continue
        index += 1
    return tokens


def extract_mods(after: str) -> str:
    tail = str(after or "").strip()
    if OBJECT_WORD_RE.match(tail):
        return ""
    if re.match(rf"(?i)^[\s,;]*(?:{PREMISE_MARKER_PATTERN})\b", tail):
        return ""
    return normalize_house_mods(tail)


def find_trailing_house(text: str) -> tuple[int, str, str] | None:
    match = HOUSE_TAIL_RE.search(text)
    if not match:
        return None
    prefix = text[: match.start()]
    if _candidate_is_numbered_street_name(prefix, text[match.start() :]):
        return None
    if not re.search(r"[а-яa-z]", prefix):
        return None
    previous = text[match.start() - 1] if match.start() > 0 else ""
    if previous and not (
        previous.isspace()
        or previous in ",.;:"
        or re.match(r"[а-яa-z]", previous, re.IGNORECASE)
    ):
        return None
    return match.start(), match.group("base"), normalize_house_mods(match.group("tail"))


def _candidate_is_numbered_street_name(prefix: str, candidate: str) -> bool:
    if not re.search(
        rf"(?:^|[\s,;])(?:{STREET_TYPE_PATTERN})\s*(?:(?:имени|им\.?)\s*)?$",
        str(prefix or ""),
        flags=re.IGNORECASE,
    ):
        return False
    tail = re.sub(r"^\s*\d+\s*[-–]?\s*", "", str(candidate or "").strip(), flags=re.IGNORECASE)
    if not tail:
        return False
    if re.match(r"(?i)^(?:я|й|ая|ой)\s+[а-яa-z]", tail):
        return True
    return bool(re.search(r"(?i)\b[а-яa-z]{2,}\b", tail))


def find_pre_street_house(text: str) -> tuple[int, int, str, str, int] | None:
    street_marker = re.search(rf"\b(?:{STREET_TYPE_PATTERN})\b", str(text or ""), flags=re.IGNORECASE)
    if not street_marker:
        return None
    prefix = str(text or "")[: street_marker.start()]
    match = re.search(
        r"(?:^|[,;]\s*)(?P<base>\d+[а-яa-z]?)(?P<tail>\s*(?:"
        r"[/\-]\s*\d+[а-яa-z]?|"
        r"[,;]?\s*(?:к\.?|корп\.?|корпус)\s*\d+[а-яa-z]?|"
        r"[,;]?\s*(?:с\.?|стр\.?|строение)\s*\d+[а-яa-z]?|"
        r"[,;]?\s*(?:лит\.?а?|литера)\s*[а-яa-z]"
        r")*)\s*[,;]?\s*$",
        prefix,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    return match.start(), match.end(), match.group("base"), normalize_house_mods(match.group("tail")), street_marker.start()


def find_construction_with_base(text: str) -> tuple[int, str, str] | None:
    candidate: tuple[int, str, str] | None = None
    pattern = re.compile(
        r"(?P<base>\d+[а-яa-z]?)\s*[,;]?\s*"
        r"(?:с\.?|стр\.?|строение)\s*(?P<construction>\d+[а-яa-z]?)\b",
        flags=re.IGNORECASE,
    )
    for match in pattern.finditer(str(text or "")):
        prefix = str(text or "")[: match.start("base")]
        has_explicit_street = bool(re.search(rf"\b(?:{STREET_TYPE_PATTERN})\b", prefix, flags=re.IGNORECASE))
        if not has_explicit_street and not _prefix_has_implicit_street_before_house(prefix):
            continue
        if re.search(
            r"(?:^|[\s,;])(?:д\.?|дом|зд\.?|здание|к\.?|кор\.?|корп\.?|корпус|стр\.?|строение)\s*$",
            prefix,
            flags=re.IGNORECASE,
        ):
            continue
        if _candidate_is_numbered_street_name(prefix, match.group("base")):
            continue
        base_match = re.fullmatch(r"(\d+)([а-яa-z]?)", match.group("base"), flags=re.IGNORECASE)
        base = base_match.group(1) if base_match else match.group("base")
        base_suffix = base_match.group(2) if base_match else ""
        candidate = (
            match.start("base"),
            base,
            normalize_house_mods(f"{base_suffix} стр. {match.group('construction')}"),
        )
    return candidate


def find_house_with_premise(text: str) -> tuple[int, str, str] | None:
    candidate: tuple[int, str, str] | None = None
    pattern = re.compile(
        rf"(?P<marker>(?<![-а-яa-z])\b(?:д\.?\s*|дом\s+|зд\.?\s*|здание\s+))?"
        rf"(?P<base>\d+[а-яa-z]?)\s*[,;]?\s*"
        rf"(?:{PREMISE_MARKER_PATTERN})\s*\d+[а-яa-z]?\b",
        flags=re.IGNORECASE,
    )
    for match in pattern.finditer(str(text or "")):
        start = match.start("marker") if match.group("marker") else match.start("base")
        prefix = str(text or "")[:start]
        has_explicit_street = bool(re.search(rf"\b(?:{STREET_TYPE_PATTERN})\b", prefix, flags=re.IGNORECASE))
        if not has_explicit_street and not _prefix_has_implicit_street_before_house(prefix):
            continue
        if re.search(
            r"(?:^|[\s,;])(?:к\.?|кор\.?|корп\.?|корпус|стр\.?|строение)\s*$",
            prefix,
            flags=re.IGNORECASE,
        ):
            continue
        base_match = re.fullmatch(r"(\d+)([а-яa-z]?)", match.group("base"), flags=re.IGNORECASE)
        base = base_match.group(1) if base_match else match.group("base")
        base_suffix = base_match.group(2) if base_match else ""
        candidate = (start, base, normalize_house_mods(base_suffix))
    return candidate


def _prefix_has_implicit_street_before_house(prefix: str) -> bool:
    parts = [part.strip() for part in re.split(r"[,;]", str(prefix or "")) if part.strip()]
    if len(parts) < 2:
        return False
    street = parts[-1]
    if looks_like_settlement_segment(street) or looks_like_address_text(street):
        return False
    if _looks_like_admin_context_segment(street) or OBJECT_WORD_RE.search(street):
        return False
    words = re.findall(r"[а-яa-z0-9]+", street.lower().replace("ё", "е"), flags=re.IGNORECASE)
    if not (1 <= len(words) <= 5) or not any(re.search(r"[а-яa-z]", word, flags=re.IGNORECASE) for word in words):
        return False
    return any(looks_like_settlement_segment(part) or _looks_like_admin_context_segment(part) for part in parts[:-1])


def _drop_repeated_house_alias_from_location(location: str, house_base: str) -> str:
    if not house_base:
        return location
    escaped = re.escape(house_base)
    return re.sub(rf"([,;]\s*){escaped}\s*[,;]\s*$", r"\1", location, flags=re.IGNORECASE)


def strip_object_noise(text: str) -> str:
    text = re.sub(r"\b(?:школа|сош|лицей|гимназия|детск\w*\s+сад|садик)\s*(?:№|#)?\s*\d+[а-яa-z]?\b", " ", text)
    text = re.sub(r"\b(?:мбоу|мбдоу|мкдоу|маоу|мкоу|ано|аоу|сош|доу|оош|нош)\b", " ", text)
    text = re.sub(r"\b(?:школа|лицей|гимназия|колледж|техникум|университет|академия|институт|общежит\w*|кампус|корпус|учрежден\w*)\b", " ", text)
    text = re.sub(r"\b(?:school|kindergarten|college|university)\b", " ", text)
    text = re.sub(r"(?:№|#)\s*\d+[а-яa-z]?", " ", text)
    text = re.sub(r"\b(?:имени|им\.?|филиал|отделение|корпус)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def looks_like_settlement_segment(value: str | None) -> bool:
    segment = str(value or "").strip()
    match = SETTLEMENT_SEGMENT_RE.match(segment)
    return bool(match and _valid_settlement_hint_segment(segment, match.group("name")))


def looks_like_address_text(value: str | None) -> bool:
    return bool(STREET_OR_TERRITORY_RE.search(str(value or "").lower().replace("ё", "е")))


def explicit_settlement_hint(value: str | None) -> str | None:
    if not value:
        return None
    text = re.sub(r"\s+", " ", str(value).replace("ё", "е")).strip()

    for segment in reversed([part.strip() for part in re.split(r"[,;]", text) if part.strip()]):
        if looks_like_address_text(segment):
            continue
        match = SETTLEMENT_SEGMENT_RE.match(segment)
        if match and _valid_settlement_hint_segment(segment, match.group("name")):
            return segment

    boundary = r"(?:^|[,;]\s*)"
    patterns = (
        (re.compile(boundary + r"(?:г\.\s*|г\s+|город\s+)([А-ЯA-Z][А-ЯA-Zа-яa-z0-9.-]+)", re.IGNORECASE), "г"),
        (re.compile(boundary + r"(?:с\.\s*|с\s+|село\s+)([А-ЯA-Z][А-ЯA-Zа-яa-z0-9.-]+)", re.IGNORECASE), "с"),
        (re.compile(boundary + r"(?:д\.\s*|д\s+|деревня\s+)([А-ЯA-Z][А-ЯA-Zа-яa-z0-9.-]+)", re.IGNORECASE), "д"),
        (re.compile(boundary + r"пгт\s*([А-ЯA-Z][А-ЯA-Zа-яa-z0-9.-]+)", re.IGNORECASE), "пгт"),
        (re.compile(boundary + r"(?:п\.\s*|п\s+|пос\.\s*|поселок\s+)([А-ЯA-Z][А-ЯA-Zа-яa-z0-9.-]+)", re.IGNORECASE), "п"),
    )
    for pattern, canonical_prefix in patterns:
        match = pattern.search(text)
        if not match:
            continue
        if _previous_hint_segment_looks_like_street(text, match.start()):
            continue
        name = match.group(1).strip(" .,;")
        if name and _valid_settlement_hint_segment(match.group(0), name):
            return f"{canonical_prefix} {name}"
    return None


def _valid_settlement_hint_segment(segment: str, name: str | None) -> bool:
    value = str(segment or "").strip()
    clean_name = str(name or "").strip(" .,;")
    if not clean_name:
        return False
    if re.match(r"(?i)^(?:д\.?|дом|зд\.?|здание|стр\.?|строение|пом\.?|помещение)\s*\d", value):
        return False
    if re.fullmatch(r"(?i)(?:к\.?|корп\.?|корпус)\s*(?:\d+[а-яa-z]?|[а-яa-z])", value):
        return False
    # The same marker followed by a number, whatever trails it: "к. 2 (86:19:...)"
    # is a building corpus carrying a cadastral number, not a кишлак named "2".
    if re.match(r"(?i)^(?:к\.?|корп\.?|корпус)\s*\d", value):
        return False
    if re.fullmatch(r"(?i)(?:лит\.?|литера)\s*[а-яa-z]", value):
        return False
    if re.fullmatch(r"\d+[а-яa-z]?(?:\s*(?:/|-|к\.?|корп\.?|стр\.?|с\.?)\s*\d+[а-яa-z]?)?", clean_name, flags=re.IGNORECASE):
        return False
    return True


def _previous_hint_segment_looks_like_street(text: str, match_start: int) -> bool:
    previous = str(text[:match_start]).rsplit(",", 1)[-1]
    return looks_like_address_text(previous)


def _normalize_named_words(value: str | None) -> str:
    text = str(value or "").lower().replace("ё", "е")
    text = re.sub(r"[\"'«»\u201c\u201d\u2018\u2019]", "", text)
    text = re.sub(r"[^\w\s]", " ", text)
    words = [word for word in text.split() if len(word) > 1 or word.isdigit()]
    return " ".join(sorted(words))


def _settlement_territory_from_name(name: str | None) -> str:
    territory = _normalize_named_words(name)
    if not territory:
        return ""
    if CITY_NAME and territory == _normalize_named_words(CITY_NAME):
        return ""
    return territory


def _settlement_territory_from_segment(segment: str | None) -> str:
    value = str(segment or "").strip()
    match = SETTLEMENT_SEGMENT_RE.match(value)
    if not match or not _valid_settlement_hint_segment(value, match.group("name")):
        return ""
    return _settlement_territory_from_name(match.group("name"))


def _strip_leading_settlement_prefix(text: str | None) -> tuple[str, str]:
    value = str(text or "").strip()
    match = SETTLEMENT_LEADING_RE.match(value)
    if not match:
        return "", value
    territory = _settlement_territory_from_name(match.group("name"))
    rest = value[match.end() :].lstrip(" ,")
    return territory, rest or value


def _merge_territory_parts(*parts: str) -> str:
    words: list[str] = []
    seen: set[str] = set()
    for part in parts:
        for word in str(part or "").split():
            if word and word not in seen:
                seen.add(word)
                words.append(word)
    return " ".join(sorted(words))


def parse_address_components(address: str | None) -> tuple[str, str, frozenset[str], str, str]:
    if not address:
        return ("", "", frozenset(), "", "")

    text = str(address)
    text = re.sub(r"[\"'«»\u201c\u201d\u2018\u2019]", "", text)
    text = text.lower().replace("ё", "е").replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    text = _normalize_common_street_typos_for_parse(text)
    text = _repair_repeated_settlement_before_trailing_street_type(text)
    text = re.sub(r"\b\d{6}\b", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(
        r"(?P<house>\d+[а-яa-z]?(?:\s*(?:[/\-]|к\.?|корп\.?|корпус|с\.?|стр\.?|строение)\s*\d+[а-яa-z]?)?)\s*\([^)]{1,80}\)\s*$",
        r"\g<house>",
        text,
        flags=re.IGNORECASE,
    )

    house_base = ""
    house_mods = ""
    location = text

    construction = None
    house_match = None
    premise_house = find_house_with_premise(text)
    if premise_house:
        house_start, house_base, house_mods = premise_house
        location = text[:house_start]
    else:
        construction = find_construction_with_base(text)
    if not premise_house and construction:
        house_start, house_base, house_mods = construction
        location = text[:house_start]
    else:
        house_match = re.search(
            r"(?<![-а-яa-z])\b(?:д\.?\s*|дом\s+|зд\.?\s*|здание\s+|стр\.?\s*|строение\s+)(\d+)",
            text,
        )
    if not premise_house and not construction and house_match:
        trailing = find_trailing_house(text)
        marker = house_match.group(0).lower().replace("ё", "е")
        if marker.startswith(("стр", "строение")) and trailing and trailing[0] < house_match.start():
            house_start, house_base, house_mods = trailing
            location = text[:house_start]
        else:
            house_base = house_match.group(1)
            location = text[: house_match.start()]
            house_mods = extract_mods(text[house_match.end() :])
    elif not premise_house and not construction:
        street_marker = re.search(
            rf"\b(?:{STREET_TYPE_PATTERN})\b",
            text,
        )
        if street_marker:
            trailing = find_trailing_house(text)
            if trailing:
                house_start, house_base, house_mods = trailing
                location = text[:house_start]
            else:
                pre_street = find_pre_street_house(text)
                if pre_street:
                    house_start, house_end, house_base, house_mods, street_start = pre_street
                    left = text[:house_start].rstrip(" ,;")
                    right = text[street_start:].lstrip(" ,;")
                    location = ", ".join(part for part in (left, right) if part)
        else:
            trailing = find_trailing_house(text)
            if trailing:
                house_start, house_base, house_mods = trailing
                location = text[:house_start]

    location = _drop_repeated_house_alias_from_location(location, house_base)

    segments = location.split(",")
    start_index = 0
    found_address_segment = False
    for index, segment in enumerate(segments):
        if STREET_OR_TERRITORY_RE.search(segment) and not _looks_like_admin_territory_context_segment(segment):
            start_index = index
            found_address_segment = True
            break
    if not found_address_segment:
        for index, segment in enumerate(segments):
            if _looks_like_admin_context_segment(segment) or _settlement_territory_from_segment(segment):
                start_index = index + 1
    settlement_territory = ""
    for segment in reversed(segments[:start_index]):
        settlement_territory = _settlement_territory_from_segment(segment)
        if settlement_territory:
            break

    meaningful_segments = segments[start_index:]
    if any(STREET_OR_TERRITORY_RE.search(segment) and not MICRODISTRICT_RE.search(segment) for segment in meaningful_segments):
        meaningful_segments = [segment for segment in meaningful_segments if not MICRODISTRICT_RE.search(segment)]
    meaningful = ",".join(meaningful_segments).strip()
    inline_settlement_territory, stripped_meaningful = _strip_leading_settlement_prefix(meaningful)
    if inline_settlement_territory:
        settlement_territory = settlement_territory or inline_settlement_territory
        meaningful = stripped_meaningful

    territory = ""
    territory_match = re.search(
        r"(?:тер\.?\s*(?:оно\s+)?)?(?:сосн|снт|днт|тсн)\s+([\w][\w\s\-]*?)(?:\s*,|\s*$)",
        meaningful,
    )
    if territory_match:
        territory_words = re.sub(r"[^\w\s]", " ", territory_match.group(1).strip()).split()
        territory = " ".join(sorted(word for word in territory_words if len(word) > 1 or word.isdigit()))
    else:
        territory_match = re.search(r"(?:сосн|снт|днт|тсн)\s+([\w][\w\s\-]+?)(?:\s*,|\s*$)", meaningful)
        if territory_match:
            territory_words = re.sub(r"[^\w\s]", " ", territory_match.group(1).strip()).split()
            territory = " ".join(sorted(word for word in territory_words if len(word) > 1 or word.isdigit()))
    territory = _merge_territory_parts(settlement_territory, territory)

    street = strip_object_noise(meaningful)
    street = re.sub(r"\b(?:территория|тер\.?|тсн|снт|днт|сосн|дачн\w*|оно)\b", " ", street)
    street = SETTLEMENT_ADMIN_RE.sub(" ", street)
    street = re.sub(rf"^\s*(?:{SETTLEMENT_TYPE_PATTERN})\s+", " ", street, flags=re.IGNORECASE)
    street = re.sub(r"\b(?:[дспх]|рп|пос|ст|ст-ца|рзд)\.\s", " ", street)
    street = re.sub(
        rf"\b(?:{STREET_TYPE_PATTERN})\b",
        " ",
        street,
    )
    street = re.sub(r"\b(?:город|г\.?|район|р-н\.?|область|обл\.?|край|округ)\b", " ", street)
    street = re.sub(rf"\b(?:участок|уч\.?|{PREMISE_MARKER_PATTERN}|автодорог\w*)\b", " ", street)
    if CITY_NAME:
        street = re.sub(r"\b" + re.escape(CITY_NAME) + r"\b", " ", street)
    street = re.sub(r"\b(\d+)\s*-?\s*(?:я|й|го|ой|ая|ий|ых|ому)\b", r"\1", street)
    street = re.sub(r"[^\w\s]", " ", street)

    text_words: list[str] = []
    numbers: set[str] = set()
    for word in street.split():
        word = _canonical_street_word(word)
        if re.fullmatch(r"\d+", word):
            numbers.add(word)
        elif len(word) > 1:
            text_words.append(word)
    if not text_words:
        text_words = _fallback_street_name_words(meaningful)

    return (
        territory,
        " ".join(sorted(text_words)),
        frozenset(numbers),
        house_base,
        house_mods,
    )


def _canonical_street_word(word: str | None) -> str:
    text = str(word or "").lower().replace("ё", "е")
    return STREET_WORD_ALIASES.get(text, text)


def _normalize_common_street_typos_for_parse(value: str) -> str:
    # "прт" is a mistyped "пр-т"; the avenue name after it is left as written.
    text = re.sub(
        r"(?i)\b(?:ул\.?|улица)\s+прт\s+(?=[А-ЯЁA-Zа-яёa-z])",
        "пр-кт ",
        str(value or ""),
    )
    text = re.sub(r"(?i)\bпрт\s+(?=[А-ЯЁA-Zа-яёa-z])", "пр-кт ", text)
    return text


def _repair_repeated_settlement_before_trailing_street_type(value: str) -> str:
    parts = [part.strip() for part in str(value or "").split(",")]
    if len(parts) < 2:
        return value

    repaired: list[str] = []
    street_type = r"тракт|шоссе|проезд|бульвар|переулок|проспект|набережная|площадь"
    for part in parts:
        match = re.match(
            rf"(?i)^(?P<settlement>(?:г\.?|город)\s+[а-яa-z0-9 .'-]+?)\s+"
            rf"(?P<street>[а-яa-z0-9 .'-]+?\s+(?:{street_type}))$",
            part,
        )
        if match and repaired and _settlement_segment_key(repaired[-1]) == _settlement_segment_key(match.group("settlement")):
            repaired.append(match.group("street").strip(" ,"))
            continue
        repaired.append(part)
    return ", ".join(repaired)


def _settlement_segment_key(value: str | None) -> str:
    text = str(value or "").lower().replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    text = re.sub(r"\b(?:г|город|с|село|д|деревня|п|пос|поселок|пгт|рп)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _looks_like_admin_context_segment(segment: str | None) -> bool:
    text = str(segment or "").lower().replace("ё", "е").strip()
    if not text:
        return False
    return any(
        marker in text
        for marker in (
            "область",
            "край",
            "республика",
            "автономный округ",
            "муниципаль",
            "городской округ",
            "муниципальный округ",
            "район",
            "сельсовет",
            "поссовет",
        )
    )


def _looks_like_admin_territory_context_segment(segment: str | None) -> bool:
    text = str(segment or "").lower().replace("ё", "е")
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return False
    if re.search(r"\b(?:снт|сосн|днт|тсн|оно)\b", text):
        return False
    return any(
        marker in text
        for marker in (
            "межселенная территория",
            "межселенная тер",
            "на территории которой",
            "земли следующих населенных пунктов",
            "земли следующих населённых пунктов",
        )
    )


def _fallback_street_name_words(value: str | None) -> list[str]:
    segments = [part.strip() for part in str(value or "").lower().replace("ё", "е").split(",") if part.strip()]
    for segment in reversed(segments):
        match = re.search(rf"\b(?:{STREET_TYPE_PATTERN})\b\.?\s+(?P<name>[а-яa-z0-9\s.-]+)$", segment, re.IGNORECASE)
        if match:
            name = re.sub(r"[^\w\s]", " ", match.group("name"))
            words = [
                _canonical_street_word(word)
                for word in name.split()
                if len(word) > 1 and not word.isdigit()
            ]
            if words:
                return words
        if re.fullmatch(r"(?:набережная|наб\.?)", segment, re.IGNORECASE):
            return ["набережная"]
    return []


def extract_locality(text: str) -> str:
    matches = re.findall(
        r"\b(?:ж\.\s*-?\s*д\.\s*(?:ст|рзд)|гп|дп|кп|г|город|с|село|п|пос|поселок|посёлок|"
        r"пгт|рп|рабочий\s+поселок|д|деревня|ст-ца|ст|станица|х|хутор|рзд|разъезд)\.?\s+"
        r"([а-яa-z0-9][а-яa-z0-9\-\s]*?)(?=,|$)",
        text.lower().replace("ё", "е"),
    )
    if not matches:
        return ""
    words = re.sub(r"[^\w\s]", " ", matches[-1]).split()
    return " ".join(sorted(word for word in words if len(word) > 1 or word.isdigit()))


def strict_house_check(left_house_base: str, left_house_mods: str, right_house_base: str, right_house_mods: str) -> bool:
    if left_house_base and right_house_base:
        return left_house_base == right_house_base and left_house_mods == right_house_mods
    if left_house_base and not right_house_base:
        return False
    if not left_house_base and right_house_base:
        return False
    return True


def relaxed_house_check(
    left_house_base: str,
    left_house_mods: str,
    right_house_base: str,
    right_house_mods: str,
) -> bool:
    """Relax street matching, but keep house modifiers protected."""
    if left_house_base and right_house_base:
        return left_house_base == right_house_base and (left_house_mods or "") == (right_house_mods or "")
    if left_house_base and not right_house_base:
        return False
    if not left_house_base and right_house_base:
        return False
    return True


def similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    if fuzz is not None:
        return float(fuzz.token_set_ratio(left, right)) / 100.0
    return SequenceMatcher(a=left, b=right).ratio()


def fuzzy_candidates(query: str, choices: dict[object, str], score_cutoff: int, limit: int) -> list[tuple[object, float]]:
    if not query or not choices:
        return []
    if process is not None:
        matches = process.extract(query, choices, scorer=fuzz.token_set_ratio, score_cutoff=score_cutoff, limit=limit)
        return [(key, float(score) / 100.0) for _, score, key in matches]

    ranked: list[tuple[object, float]] = []
    for key, text in choices.items():
        score = similarity(query, text)
        if score * 100 >= score_cutoff:
            ranked.append((key, score))
    ranked.sort(key=lambda item: item[1], reverse=True)
    return ranked[:limit]


def group_by_territory(items: Iterable[tuple[object, str]]) -> dict[str, dict[object, str]]:
    grouped: dict[str, dict[object, str]] = defaultdict(dict)
    for key, territory in items:
        grouped[territory][key] = territory
    return grouped
