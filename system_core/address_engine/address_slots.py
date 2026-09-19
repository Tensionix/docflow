from __future__ import annotations

import re
from typing import Any

from .audion_address_core import (
    compact_house_modifier_tokens,
    find_trailing_house,
    looks_like_address_text,
    looks_like_settlement_segment,
    normalize_house_mods,
    parse_address_components,
)
from .models import AddressParts


BARE_PR_PROEZD_STREET_NAMES = {
    "геологоразведчиков",
    "заречный",
}


def build_address_parts(
    address_text: str | None,
    *,
    postal_index: str | None = None,
    parent_subject: str | None = None,
    subject: str | None = None,
    federal_district: str | None = None,
    autonomous_okrug: str | None = None,
    municipality: str | None = None,
    locality: str | None = None,
    oktmo_name: str | None = None,
) -> AddressParts:
    """Split a normalized address into card slots without replacing address core."""
    address = _clean(address_text)
    segments = _segments(address)
    territory, _street_words, street_numbers, house_base, house_mods = parse_address_components(address)

    subject_value = _clean(subject) or _first_subject_segment(segments)
    municipality_value = _clean(municipality) or _first_municipality_segment(segments)
    official_locality = _format_oktmo_locality_for_address(oktmo_name)
    explicit_locality = _format_oktmo_locality_for_address(locality)
    locality_value = (
        _official_locality_display_if_same(explicit_locality, official_locality)
        or explicit_locality
        or official_locality
        or _format_oktmo_locality_for_address(_last_locality_segment(segments))
    )
    territory_display = _extract_territory_segment(segments)
    microdistrict_value = _extract_microdistrict_segment(segments)
    street_value = _extract_street_segment(segments, locality_value)
    territory = _repair_territory_overconsumed_street(territory, locality_value, street_value)
    territory_value = territory_display or (territory or None if not locality_value else None)
    microdistrict_tail_only = _microdistrict_tail_is_only_locator(segments, street_value)
    if microdistrict_tail_only:
        house_base = ""
        house_mods = ""
    else:
        house_mods = _merge_house_mods(house_mods, _extract_house_modifier_mods(segments))
    house_value = _extract_house_segment(segments, street_value, house_mods, house_base)
    if not house_value and not microdistrict_tail_only:
        house_value = _format_house(house_base, house_mods)
    premise_value = _extract_premise_segment(segments)

    base = AddressParts(
        postal_index=_clean(postal_index),
        parent_subject=_clean(parent_subject),
        subject=subject_value,
        federal_district=_clean(federal_district),
        autonomous_okrug=_clean(autonomous_okrug),
        municipality=municipality_value,
        locality=locality_value,
        territory=territory_value,
        microdistrict=microdistrict_value,
        street=street_value,
        house=house_value,
        house_kind=_house_kind_from_segment(house_value),
        house_base=house_base or None,
        house_mods=house_mods or None,
        house_modifiers=tuple(_house_modifier_parts(house_mods)),
        premise=premise_value,
        street_numbers=tuple(sorted(street_numbers)),
    )
    rendered = render_address_parts(base, include_subject=False) or address
    rendered_with_subject = render_address_parts(base, include_subject=True) or rendered
    return AddressParts(
        postal_index=base.postal_index,
        parent_subject=base.parent_subject,
        subject=base.subject,
        federal_district=base.federal_district,
        autonomous_okrug=base.autonomous_okrug,
        municipality=base.municipality,
        locality=base.locality,
        territory=base.territory,
        microdistrict=base.microdistrict,
        street=base.street,
        house=base.house,
        house_kind=base.house_kind,
        house_base=base.house_base,
        house_mods=base.house_mods,
        house_modifiers=base.house_modifiers,
        premise=base.premise,
        street_numbers=base.street_numbers,
        rendered=rendered,
        rendered_with_subject=rendered_with_subject,
    )


def render_address_parts(parts: AddressParts | dict[str, Any] | None, *, include_subject: bool = False) -> str | None:
    return render_address_core(parts, include_subject=include_subject)


def render_address_core(parts: AddressParts | dict[str, Any] | None, *, include_subject: bool = False) -> str | None:
    """Render only final-table address core in the configured administrative-to-house order."""
    if not parts:
        return None
    data = parts.as_dict() if isinstance(parts, AddressParts) else parts
    values: list[str] = []
    if include_subject:
        _append_unique_level(values, _clean(data.get("federal_district")))
        _append_unique_level(values, _clean(data.get("parent_subject")))
        _append_unique_level(values, _clean(data.get("subject")))
        _append_unique_level(values, _clean(data.get("autonomous_okrug")))
    _append_unique_level(values, _clean(data.get("municipality")))
    _append_unique_level(values, _clean(data.get("locality")))
    _append_unique_level(values, _clean(data.get("territory")))
    _append_unique_level(values, _clean(data.get("microdistrict")))
    _append_unique_level(values, _clean(data.get("street")))
    _append_unique_level(values, _clean(data.get("house")))
    _append_unique_level(values, _clean(data.get("premise")))
    return ", ".join(values) or None


def _segments(value: str | None) -> list[str]:
    text = _normalize_compact_address_spacing(str(value or ""))
    return [part.strip() for part in text.split(",") if part.strip()]


def _normalize_compact_address_spacing(value: str) -> str:
    text = str(value or "").replace("\xa0", " ")
    text = re.sub(
        r"\b(г|с|д|п|х|у|ул|пер|пр|пр-т|пр-кт|просп|пр-д|б-р|наб|пл)\s+\.",
        r"\1.",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\b(г|с|д|п|х|у)\.\s*(?=[А-ЯЁA-Z])", r"\1. ", text, flags=re.IGNORECASE)
    text = re.sub(
        r"\b(ул|пер|пр|пр-т|пр-кт|просп|пр-д|б-р|наб|пл)\.\s*(?=[А-ЯЁA-Z0-9])",
        r"\1. ",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"(?i)(\d+)\s*[\"'«»]\s*([а-яa-z])\s*[\"'«»]", r"\1\2", text)
    locator = (
        r"улица|ул\.?|переулок|пер\.?|проспект|пр-т\.?|просп\.?|пр\.?|пр-кт\.?|"
        r"проезд|пр-д\.?|бульвар|б-р\.?|набережная|наб\.?|площадь|пл\.?|"
        r"остров|о\."
    )
    text = re.sub(
        r"(?i)\b(?P<locality>(?:г\.?|город|с\.?|село|д\.?|деревня|п\.?|пос\.?|поселок|посёлок|пгт|рп)\s+"
        r"[А-ЯЁA-Zа-яёa-z0-9 .'-]+?)\s+"
        rf"(?=(?:{locator})(?:\s|$))",
        r"\g<locality>, ",
        text,
    )
    return text


def _first_subject_segment(segments: list[str]) -> str | None:
    for segment in segments:
        if _looks_like_subject_segment(segment):
            return segment
    return None


def _first_municipality_segment(segments: list[str]) -> str | None:
    for segment in segments:
        if _looks_like_municipality_segment(segment):
            return segment
    subject_index = next((index for index, segment in enumerate(segments) if _looks_like_subject_segment(segment)), -1)
    locality_index = next((index for index, segment in enumerate(segments) if looks_like_settlement_segment(segment)), -1)
    if locality_index <= subject_index + 1:
        return None
    for segment in segments[subject_index + 1 : locality_index]:
        if _looks_like_implicit_municipality_segment(segment):
            return segment
    return None


def _last_locality_segment(segments: list[str]) -> str | None:
    for segment in reversed(segments):
        if (
            _looks_like_house_segment(segment)
            or _looks_like_house_modifier_segment(segment)
            or _looks_like_premise_segment(segment)
            or _looks_like_microdistrict_segment(segment)
            or _looks_like_standalone_house_number_segment(segment)
            or looks_like_address_text(segment)
        ):
            continue
        if looks_like_settlement_segment(segment):
            return segment
    return None


def _extract_microdistrict_segment(segments: list[str]) -> str | None:
    for segment in segments:
        if _looks_like_microdistrict_segment(segment):
            return _normalize_microdistrict_segment(segment)
    return None


def _extract_territory_segment(segments: list[str]) -> str | None:
    for segment in segments:
        if _looks_like_territory_segment(segment):
            return _normalize_territory_segment(segment)
    return None


def _extract_street_segment(segments: list[str], locality: str | None = None) -> str | None:
    for index, segment in enumerate(segments):
        if _looks_like_object_context_segment(segment):
            continue
        if _looks_like_municipality_segment(segment) or _looks_like_parent_territory(segment):
            continue
        if _looks_like_territory_segment(segment):
            continue
        if _looks_like_microdistrict_segment(segment):
            continue
        if _looks_like_house_modifier_segment(segment):
            continue
        if _looks_like_premise_segment(segment):
            continue
        if _looks_like_standalone_house_number_segment(segment):
            continue
        if not looks_like_address_text(segment):
            continue
        if _looks_like_house_segment(segment):
            continue
        street_segment = _strip_repeated_locality_prefix_from_street_segment(_strip_house_tail(segment), locality)
        street = _normalize_street_segment(street_segment)
        if _looks_like_empty_street_locator_segment(street):
            continue
        street = _extend_dangling_street_range(street, segments[index + 1 :])
        return _clean(street)
    return _extract_implicit_street_segment(segments)


def _extract_implicit_street_segment(segments: list[str]) -> str | None:
    for index, segment in enumerate(segments):
        if _looks_like_non_street_context_segment(segment):
            continue
        if not _looks_like_implicit_numbered_street_segment(segment):
            continue
        if not _later_segments_have_house(segments[index + 1 :]):
            continue
        return _normalize_implicit_street_segment(segment)
    for index, segment in enumerate(segments):
        if _looks_like_non_street_context_segment(segment):
            continue
        if not _looks_like_implicit_bare_street_segment(segment):
            continue
        if not _earlier_segments_have_locality_context(segments[:index]):
            continue
        if not _later_segments_have_house(segments[index + 1 :]):
            continue
        return _normalize_street_segment(f"ул. {segment}")
    return None


def _extract_house_segment(
    segments: list[str],
    street: str | None,
    house_mods: str | None = None,
    house_base: str | None = None,
) -> str | None:
    for segment in segments:
        if _modifier_segment_belongs_to_parsed_house(segment, house_base, house_mods):
            continue
        if _looks_like_house_segment(segment):
            return _append_house_modifiers(_normalize_house_segment(segment), house_mods)

    street_seen = False
    for segment in segments:
        if _looks_like_microdistrict_segment(segment):
            continue
        if _looks_like_house_modifier_segment(segment) or _looks_like_premise_segment(segment):
            continue
        if street and _same_street_segment(_strip_house_tail(segment), street):
            street_seen = True
            continue
        if street_seen and _looks_like_standalone_house_number_segment(segment):
            return _append_house_modifiers(_normalize_standalone_house_segment(segment), house_mods)

    for segment in segments:
        if _looks_like_microdistrict_segment(segment):
            continue
        house_segment = _strip_premise_tail(segment)
        if not looks_like_address_text(house_segment):
            continue
        trailing = find_trailing_house(house_segment)
        if not trailing:
            continue
        start, _base, _mods = trailing
        prefix = _clean(house_segment[:start])
        if street and prefix and _norm(prefix) != _norm(street):
            continue
        tail = _clean(house_segment[start:])
        if tail:
            return _append_house_modifiers(_normalize_house_segment(f"д. {tail}"), house_mods)
    return None


def _modifier_segment_belongs_to_parsed_house(
    segment: str | None,
    house_base: str | None,
    house_mods: str | None,
) -> bool:
    return bool(house_base and house_mods and _looks_like_house_modifier_segment(segment))


def _extract_premise_segment(segments: list[str]) -> str | None:
    for segment in segments:
        if _looks_like_premise_segment(segment):
            return _normalize_premise_segment(segment)
    return None


def _extract_house_modifier_mods(segments: list[str]) -> str | None:
    mods = ""
    street_seen = False
    house_seen = False
    for segment in segments:
        if _looks_like_microdistrict_segment(segment) or _looks_like_premise_segment(segment):
            continue
        if _looks_like_house_segment(segment):
            house_seen = True
            continue
        if street_seen and _looks_like_standalone_house_number_segment(segment):
            house_seen = True
            continue
        if looks_like_address_text(segment):
            street_seen = True
            trailing = find_trailing_house(segment)
            if trailing:
                house_seen = True
            continue
        if not _looks_like_house_modifier_segment(segment):
            continue
        if not house_seen:
            continue
        compact = normalize_house_mods(segment)
        for token in _compact_house_modifier_tokens(compact):
            if token not in _compact_house_modifier_tokens(mods):
                mods += token
    return mods or None


def _strip_house_tail(segment: str) -> str:
    text = str(segment or "")
    if _looks_like_microdistrict_segment(text):
        return text.strip(" ,")
    explicit = re.search(r"\b(?:д\.?|дом|зд\.?|здание|стр\.?|строение)\s*\d", text, flags=re.IGNORECASE)
    if explicit:
        return text[: explicit.start()].strip(" ,")
    trailing = find_trailing_house(text)
    if trailing:
        return text[: trailing[0]].strip(" ,")
    return text.strip(" ,")


def _normalize_house_segment(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip(" ,")
    text = re.sub(r"^(?:дом|д)\.?\s*", "д. ", text, flags=re.IGNORECASE)
    text = re.sub(r"^(?:здание|зд)\.?\s*", "зд. ", text, flags=re.IGNORECASE)
    text = re.sub(r"^(?:строение|стр)\.?\s*", "стр. ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bкорпус\s*", "к. ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bстроение\s*", "стр. ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text)
    normalized = text.strip(" ,")
    normalized = _strip_premise_tail(normalized)
    match = re.match(r"(?i)^(?P<kind>д|зд|стр)\.\s*(?P<base>\d+[а-яa-z]?)(?P<mods>.*)$", normalized)
    if not match:
        return normalized
    kind = match.group("kind").lower()
    base = match.group("base")
    mods = normalize_house_mods(match.group("mods"))
    if not mods:
        return f"{kind}. {base}"
    return _format_house_with_kind(kind, base, mods) or f"{kind}. {base}"


def _normalize_street_segment(value: str | None) -> str | None:
    text = _clean(value)
    if not text:
        return None
    text = _normalize_compact_address_spacing(text)
    text = re.sub(
        r"\b(ул|пер|пр|пр-т|пр-кт|просп|пр-д|б-р|наб|пл)\.\.+",
        r"\1.",
        text,
        flags=re.IGNORECASE,
    )
    text = _strip_leading_locality_prefix_before_locator(text)
    bare_pr_as_proezd = _bare_pr_should_be_proezd(text)
    replacements = (
        (r"^улица\s+", "ул. "),
        (r"^переулок\s+", "пер. "),
        # An abbreviation written without its dot gets one: "ул Ленина" is "ул. Ленина".
        (r"^ул\.?\s+", "ул. "),
        (r"^пер\.?\s+", "пер. "),
        (r"^наб\.?\s+", "наб. "),
        (r"^пл\.?\s+", "пл. "),
        (r"^туп\.?\s+", "туп. "),
        (r"^ш\.?\s+", "ш. "),
        (r"^проул\.?\s+", "проул. "),
        (r"^бульв\.?\s+", "б-р "),
        (r"^проспект\s+", "пр-кт "),
        (r"^просп\.?\s+", "пр-кт "),
        (r"^пр-т\.?\s+", "пр-кт "),
        (r"^пр\.\s+", "пр-д " if bare_pr_as_proezd else "пр-кт "),
        (r"^пр-кт\.?\s+", "пр-кт "),
        (r"^проезд\s+", "пр-д "),
        (r"^бульвар\s+", "б-р "),
        (r"^набережная\s+", "наб. "),
        (r"^площадь\s+", "пл. "),
        (r"^остров\s+", "о. "),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    # "прт" is a mistyped "пр-т"; the avenue name after it is left as written.
    text = re.sub(r"(?i)\b(?:ул\.?|улица)\s+прт\s+(?=[А-ЯЁA-Zа-яёa-z])", "пр-кт ", text)
    text = _strip_trailing_locality_suffix_from_street(text)
    text = re.sub(r"\bул\.\s*(?=\w)", "ул. ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bпер\.\s*(?=\w)", "пер. ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bпр\.\s*(?=\w)", "пр-д " if bare_pr_as_proezd else "пр-кт ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bпр-кт\.?\s*(?=\w)", "пр-кт ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bо\.\s*(?=\w)", "о. ", text, flags=re.IGNORECASE)
    text = re.sub(r"(?i)\bсовсетов\b", "Советов", text)
    text = re.sub(r"(?i)\bпрт\s+(?=[А-ЯЁA-Zа-яёa-z])", "пр-кт ", text)
    text = _normalize_street_name_typos(text)
    text = _collapse_repeated_street_type_markers(text)
    text = _strip_residential_complex_tail_from_street(text)
    text = _normalize_street_range_dashes(text)
    return re.sub(r"\s+", " ", text).strip(" ,")


def _strip_repeated_locality_prefix_from_street_segment(value: str | None, locality: str | None) -> str | None:
    text = _clean(value)
    if not text or not locality:
        return text
    locality_words = re.findall(r"[0-9A-Za-zА-ЯЁа-яё]+", locality)
    if not locality_words:
        return text
    locality_pattern = r"\W+".join(re.escape(word) for word in locality_words)
    match = re.match(rf"(?i)^\s*{locality_pattern}\W+(?P<tail>.+)$", text)
    if not match:
        return text
    tail = match.group("tail").strip(" ,")
    if _looks_like_trailing_type_street(tail) or looks_like_address_text(tail):
        return tail
    return text


def _looks_like_trailing_type_street(value: str | None) -> bool:
    return bool(
        re.search(
            r"(?i)\b[А-ЯЁA-Zа-яёa-z0-9 .'-]+?\s+"
            r"(?:тракт|шоссе|проезд|бульвар|переулок|проспект|набережная|площадь)\b",
            str(value or "").strip(" ,"),
        )
    )


def _repair_territory_overconsumed_street(territory: str, locality: str | None, street: str | None) -> str:
    if not territory or not locality or not street:
        return territory
    territory_words = set(_norm(territory).split())
    locality_words = set(_norm_locality_without_type(locality).split())
    street_words = set(_norm_street_for_compare(street).split()) - {
        "тракт",
        "шоссе",
        "проезд",
        "бульвар",
        "переулок",
        "проспект",
        "набережная",
        "площадь",
    }
    if locality_words and street_words and locality_words <= territory_words and territory_words.intersection(street_words):
        return " ".join(sorted(locality_words))
    return territory


def _norm_locality_without_type(value: str | None) -> str:
    text = _norm(value)
    text = re.sub(
        r"^(?:г|город|с|село|д|деревня|п|пос|поселок|пгт|р\s*п|рп|х|хутор|ст|станица)\s+",
        "",
        text,
    )
    return text


def _normalize_street_name_typos(value: str) -> str:
    replacements = {
        "совесткая": "Советская",
        "советска": "Советская",
        "украинска": "Украинская",
    }
    text = value
    for source, replacement in replacements.items():
        text = re.sub(rf"(?i)\b{source}\b", replacement, text)
    return text


def _extend_dangling_street_range(street: str | None, later_segments: list[str]) -> str | None:
    current = _clean(street)
    if not current or not re.search(r"[-–—]\s*$", current):
        return current
    for segment in later_segments:
        if _looks_like_empty_street_locator_segment(segment):
            continue
        if (
            _looks_like_subject_segment(segment)
            or _looks_like_municipality_segment(segment)
            or looks_like_settlement_segment(segment)
            or _looks_like_microdistrict_segment(segment)
            or _looks_like_house_segment(segment)
            or _looks_like_house_modifier_segment(segment)
            or _looks_like_premise_segment(segment)
            or _looks_like_standalone_house_number_segment(segment)
        ):
            continue
        if not looks_like_address_text(segment):
            continue
        tail = _normalize_street_segment(_strip_house_tail(segment))
        if not tail or _looks_like_empty_street_locator_segment(tail):
            continue
        return f"{current.rstrip()} {tail}".strip(" ,")
    return current


def _bare_pr_should_be_proezd(text: str) -> bool:
    match = re.match(r"(?i)^\s*пр\.\s+(.+)$", str(text or "").strip())
    if not match:
        return False
    name = _norm(match.group(1))
    return name in BARE_PR_PROEZD_STREET_NAMES


def _collapse_repeated_street_type_markers(value: str) -> str:
    text = str(value or "")
    marker_groups = (
        (r"(?:ул\.?|улица)", "ул. "),
        (r"(?:пер\.?|переулок)", "пер. "),
        (r"(?:пр-кт\.?|пр-т\.?|просп\.?|проспект)", "пр-кт "),
        (r"(?:пр-д\.?|проезд)", "пр-д "),
        (r"(?:б-р\.?|бульвар)", "б-р "),
        (r"(?:наб\.?|набережная)", "наб. "),
        (r"(?:пл\.?|площадь)", "пл. "),
        (r"(?:о\.?|остров)", "о. "),
    )
    for marker_pattern, replacement in marker_groups:
        pattern = rf"(?i)^\s*{marker_pattern}\s+{marker_pattern}\s+"
        previous = None
        while previous != text:
            previous = text
            text = re.sub(pattern, replacement, text)
    return re.sub(r"\.\s+\.", ".", text)


def _strip_residential_complex_tail_from_street(value: str) -> str:
    text = str(value or "").strip()
    if not re.match(
        r"(?i)^\s*(?:ул\.?|улица|пер\.?|переулок|проспект|пр-т\.?|пр-кт\.?|просп\.?|пр-д\.?|проезд|"
        r"б-р\.?|бульвар|наб\.?|набережная|пл\.?|площадь)\s+",
        text,
    ):
        return text
    stripped = re.sub(
        r"(?i)\s+жк\s+(?:[\"«][^\"»]+[\"»]?|[А-ЯЁA-Zа-яёa-z0-9 .'-]{2,60})\s*$",
        "",
        text,
    ).strip(" ,")
    if stripped and not _looks_like_empty_street_locator_segment(stripped):
        return stripped
    return text


def _normalize_street_range_dashes(value: str) -> str:
    text = str(value or "")
    text = re.sub(r"(?i)\bСалтыкова\s*[-–—]\s*Щедрина\b", "Салтыкова-Щедрина", text)
    text = re.sub(r"\s+[-–—]\s+", " – ", text)
    return text


def _strip_trailing_locality_suffix_from_street(value: str) -> str:
    text = str(value or "").strip()
    if not re.match(
        r"(?i)^\s*(?:ул\.?|улица|пер\.?|переулок|проспект|пр-т\.?|пр-кт\.?|просп\.?|пр-д\.?|проезд|"
        r"б-р\.?|бульвар|наб\.?|набережная|пл\.?|площадь)\s+",
        text,
    ):
        return text
    return re.sub(
        r"(?i)\s+(?:г\.?|город)\s+[А-ЯЁA-Zа-яёa-z0-9 .'-]+$",
        "",
        text,
    ).strip(" ,")


def _normalize_territory_segment(value: str | None) -> str | None:
    text = _clean(value)
    if not text:
        return None
    text = re.sub(r"^(?:территория|тер)\.?\s*", "тер. ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(тсн|снт|днт|сосн)\b", lambda match: match.group(1).upper(), text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip(" ,")


def _normalize_microdistrict_segment(value: str | None) -> str | None:
    text = _clean(value)
    if not text:
        return None
    replacements = (
        (r"^микрорайон\s+", "мкр. "),
        (r"^мкр\.?\s*", "мкр. "),
        (r"^квартал\s+", "кв-л "),
        (r"^кв-л\.?\s*", "кв-л "),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip(" ,")


def _strip_leading_locality_prefix_before_locator(value: str) -> str:
    locator = (
        r"улица|ул\.?|переулок|пер\.?|проспект|пр-т\.?|просп\.?|пр\.?|пр-кт\.?|"
        r"проезд|пр-д\.?|бульвар|б-р\.?|набережная|наб\.?|площадь|пл\.?|"
        r"остров|о\."
    )
    pattern = (
        r"(?i)^\s*(?:г\.?|город|с\.?|село|д\.?|деревня|п\.?|пос\.?|поселок|посёлок|пгт|рп)\s+"
        r"[А-ЯЁA-Zа-яёa-z0-9 .'-]+?(?:\s*,\s*|\s+)"
        rf"(?=(?:{locator})(?:\s|$))"
    )
    return re.sub(pattern, "", str(value or "")).strip(" ,")


def _normalize_premise_segment(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip(" ,")
    text = re.sub(r"^(?:помещение|пом)\.?\s*", "пом. ", text, flags=re.IGNORECASE)
    return text.strip(" ,")


def _normalize_standalone_house_segment(value: str) -> str | None:
    text = re.sub(r"\s+", " ", str(value or "")).strip(" ,.")
    text = _strip_premise_tail(text)
    match = re.match(r"(?i)^(?P<base>\d{1,4}[а-яa-z]?)(?P<mods>.*)$", text)
    if not match:
        return None
    return _format_house(match.group("base"), normalize_house_mods(match.group("mods")))


def _strip_premise_tail(value: str) -> str:
    return re.sub(
        r"(?i)\s*[,;]?\s*(?:помещ\.?|помещение|пом\.?|офис|оф\.?)\s*\d+[а-яa-z]?\b.*$",
        "",
        str(value or ""),
    ).strip(" ,")


def _normalize_implicit_street_segment(value: str) -> str | None:
    text = _clean(value)
    if not text:
        return None
    match = re.match(r"(?i)^(?P<number>\d+)\s*[-–]?\s*(?P<suffix>я|й|ая|ой)?\s+(?P<name>.+)$", text)
    if match:
        suffix = match.group("suffix") or "я"
        number = match.group("number") + (f"-{suffix}" if suffix else "")
        return _normalize_street_segment(f"ул. {number} {match.group('name').strip()}")
    return _normalize_street_segment(text)


def _append_house_modifiers(house: str | None, house_mods: str | None) -> str | None:
    house = _clean(house)
    if not house:
        return None
    parts = [house]
    existing = _norm(house)
    for modifier in _house_modifier_parts(house_mods):
        if _norm(modifier) not in existing:
            parts.append(modifier)
    return ", ".join(parts)


def _house_modifier_parts(house_mods: str | None) -> list[str]:
    text = str(house_mods or "")
    parts: list[str] = []
    seen: set[str] = set()
    for marker, number in re.findall(
        r"(?:^|[\s,;])(?:"
        r"(к\.|корп\.?|корпус|с\.|стр\.?|строение|к\s+|с\s+)\s*(\d+[а-яa-z]?)"
        r")",
        text,
        flags=re.IGNORECASE,
    ):
        marker_key = "к" if marker.lower().startswith(("к", "корп")) else "с"
        token = marker_key + number.lower()
        if token in seen:
            continue
        seen.add(token)
        parts.append(("к. " if marker_key == "к" else "стр. ") + number)
    compact_tokens = _compact_house_modifier_tokens(text)
    index = 0
    while index < len(compact_tokens):
        token = compact_tokens[index]
        if token.startswith("лит"):
            letter = token[3:4]
            lit_token = "лит" + letter.lower()
            if letter and lit_token not in seen:
                seen.add(lit_token)
                parts.append("лит. " + letter.upper())
            index += 1
            continue
        marker = token[0]
        number = token[1:]
        if index + 1 < len(compact_tokens) and compact_tokens[index + 1] == token:
            pair_token = token + token
            if pair_token not in seen:
                seen.add(token)
                seen.add(pair_token)
                parts.append(("к. " if marker == "к" else "стр. ") + f"{number}/{number}")
            index += 2
            continue
        if token in seen:
            index += 1
            continue
        seen.add(token)
        parts.append(("к. " if marker == "к" else "стр. ") + number)
        index += 1
    for letter in re.findall(r"лит([а-яa-z])", text, flags=re.IGNORECASE):
        token = "лит" + letter.lower()
        if token in seen:
            continue
        seen.add(token)
        parts.append("лит. " + letter.upper())
    return parts


def _merge_house_mods(*values: str | None) -> str:
    merged = ""
    seen: set[str] = set()
    for value in values:
        raw = str(value or "")
        if not raw:
            continue
        compact = normalize_house_mods(raw)
        prefix = _house_suffix_prefix(compact)
        if prefix and not merged:
            merged += prefix
        tokens = _compact_house_modifier_tokens(compact)
        for index, token in enumerate(tokens):
            preserve_duplicate = index > 0 and token == tokens[index - 1]
            if token in seen and not preserve_duplicate:
                continue
            if not preserve_duplicate:
                seen.add(token)
            merged += token
    return merged


def _house_suffix_prefix(value: str) -> str:
    match = re.match(r"(?i)^([а-яa-z]+)(?=к\d|с\d|лит[а-яa-z]|$)", str(value or ""))
    if not match:
        return ""
    prefix = match.group(1)
    if prefix.startswith("лит"):
        return ""
    return prefix


def _compact_house_modifier_tokens(value: str | None) -> list[str]:
    return compact_house_modifier_tokens(value)


def _house_kind_from_segment(value: str | None) -> str | None:
    text = _clean(value)
    if not text:
        return None
    match = re.match(r"(?i)^(д|дом|зд|здание|стр|строение)\.?\b", text)
    if not match:
        return None
    raw = match.group(1).lower()
    if raw in {"зд", "здание"}:
        return "зд"
    if raw in {"стр", "строение"}:
        return "стр"
    return "д"


def _format_house(house_base: str | None, house_mods: str | None) -> str | None:
    return _format_house_with_kind("д", house_base, house_mods)


def _format_house_with_kind(kind: str, house_base: str | None, house_mods: str | None) -> str | None:
    base = _clean(house_base)
    if not base:
        return None
    normalized_kind = "зд" if kind == "зд" else "стр" if kind == "стр" else "д"
    mods = str(house_mods or "")
    suffix = ""
    if mods and re.match(r"[а-яa-z]", mods[0], flags=re.IGNORECASE) and not re.match(
        r"(?i)^(?:[кс]\d|лит[а-яa-z])",
        mods,
    ):
        suffix = mods[0].upper()
        mods = mods[1:]
    parts = [f"{normalized_kind}. {base}{suffix}"]
    parts.extend(_house_modifier_parts(mods))
    return ", ".join(parts)


def _format_oktmo_locality_for_address(value: str | None) -> str | None:
    text = _clean(value)
    if not text or _looks_like_parent_territory(text):
        return None
    replacements = (
        (r"^город\s+", "г. "),
        (r"^село\s+", "с. "),
        (r"^деревня\s+", "д. "),
        (r"^пос[её]лок\s+городского\s+типа\s+", "пгт "),
        (r"^пос[её]лок\s+", "п. "),
        (r"^пос\s+", "пос. "),
        (r"^хутор\s+", "х. "),
        (r"^станица\s+", "ст-ца "),
        (r"^станция\s+", "ст. "),
        (r"^рп\s+", "р. п. "),
        (r"^р\.?\s*п\.?\s+", "р. п. "),
        (r"^пгт\s+", "пгт "),
        (r"^г\.?\s+", "г. "),
        (r"^с\.?\s+", "с. "),
        (r"^д\.?\s+", "д. "),
        (r"^п\.?\s+", "п. "),
        (r"^х\.?\s+", "х. "),
        (r"^ст\.?\s+", "ст. "),
        (r"^у\.?\s+", "у. "),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text


def _official_locality_display_if_same(locality: str | None, official_locality: str | None) -> str | None:
    if not locality or not official_locality:
        return None
    if _locality_display_key(locality) == _locality_display_key(official_locality):
        return official_locality
    return None


def _locality_display_key(value: str | None) -> str:
    text = _norm(value)
    text = re.sub(r"\bуссурйиск\b", "уссурийск", text)
    text = re.sub(
        r"^(?:г|город|с|село|д|деревня|п|пос|поселок|пгт|р\s*п|рп|х|хутор|ст|станица)\s+",
        "",
        text,
    )
    key = re.sub(r"\s+", "", text)
    aliases = {
        "алексееникольск": "алексейникольское",
        "алексееникольское": "алексейникольское",
    }
    return aliases.get(key, key)


def _append_unique_level(values: list[str], value: str | None) -> None:
    if not value:
        return
    value_key = _norm(value)
    value_level_key = _level_norm(value)
    if any(_norm(existing) == value_key or _level_norm(existing) == value_level_key for existing in values):
        return
    values.append(value)


def _looks_like_subject_segment(value: str) -> bool:
    text = value.lower().replace("ё", "е")
    return any(marker in text for marker in ("область", "край", "республика", "автономный округ"))


def _looks_like_municipality_segment(value: str) -> bool:
    text = value.lower().replace("ё", "е")
    return any(
        marker in text
        for marker in (
            "муниципаль",
            "городской округ",
            "муниципальный округ",
            "район",
            "межселен",
            "поселение",
            "сельсовет",
            "поссовет",
        )
    ) or text.startswith("город ")


def _looks_like_implicit_municipality_segment(value: str) -> bool:
    text = str(value or "").strip()
    if not text:
        return False
    if _looks_like_subject_segment(text) or looks_like_settlement_segment(text):
        return False
    if (
        _looks_like_house_segment(text)
        or _looks_like_house_modifier_segment(text)
        or _looks_like_standalone_house_number_segment(text)
        or _looks_like_microdistrict_segment(text)
        or looks_like_address_text(text)
    ):
        return False
    return bool(re.fullmatch(r"[А-ЯЁA-Z][А-ЯЁа-яёA-Za-z -]+(?:ское|цкое|ское|ский|ской|ская|ое)", text))


def _looks_like_parent_territory(value: str) -> bool:
    text = value.lower().replace("ё", "е")
    return _looks_like_municipality_segment(text) or any(marker in text for marker in ("поселение", "межселен"))


def _looks_like_house_segment(value: str) -> bool:
    return bool(re.match(r"(?i)^\s*(?:д\.?|дом|зд\.?|здание|стр\.?|строение)\s*\d", str(value or "")))


def _looks_like_house_modifier_segment(value: str) -> bool:
    text = str(value or "").strip()
    return bool(
        re.fullmatch(r"(?i)(?:к\.?|корп\.?|корпус)\s*(?:\d+[а-яa-z]?|[а-яa-z])", text)
        or re.fullmatch(r"(?i)(?:с\.?|стр\.?|строение)\s*\d+[а-яa-z]?", text)
        or re.fullmatch(r"(?i)(?:лит\.?|литера)\s*[а-яa-z]", text)
    )


def _looks_like_standalone_house_number_segment(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"(?i)\s*\d{1,4}[а-яa-z]?(?:"
            r"\s*(?:/|-)\s*(?:\d+[а-яa-z]?|[а-яa-z])|"
            r"\s*(?:к\.?|корп\.?|корпус)\s*\d+[а-яa-z]?|"
            r"\s*(?:с\.?|стр\.?|строение)\s*\d+[а-яa-z]?|"
            r"\s*(?:лит\.?а?|литера)\s*[а-яa-z]"
            r")*\s*\.?\s*",
            str(value or ""),
        )
    )


def _looks_like_implicit_numbered_street_segment(value: str) -> bool:
    text = str(value or "").strip()
    return bool(re.match(r"(?i)^\d+\s*[-–]?\s*(?:я|й|ая|ой)?\s+[а-яa-z][а-яa-z0-9 .'-]+$", text))


def _looks_like_implicit_bare_street_segment(value: str) -> bool:
    text = str(value or "").strip(" ,")
    if not text or len(text) > 80 or re.search(r"\d", text):
        return False
    words = re.findall(r"[а-яa-z][а-яa-z.'-]*", text.lower().replace("ё", "е"), flags=re.IGNORECASE)
    if not (1 <= len(words) <= 5):
        return False
    blocked = (
        r"сад|школ|сош|доу|лицей|гимназ|колледж|техникум|университет|"
        r"корпус|филиал|общежит|спортзал|учебн|центр|учрежден|организац|"
        r"район|область|город|муниципаль"
    )
    return not bool(re.search(rf"(?i)\b(?:{blocked})\b", " ".join(words)))


def _earlier_segments_have_locality_context(segments: list[str]) -> bool:
    return any(
        _looks_like_subject_segment(segment)
        or _looks_like_municipality_segment(segment)
        or looks_like_settlement_segment(segment)
        for segment in segments
    )


def _later_segments_have_house(segments: list[str]) -> bool:
    for segment in segments:
        if (
            _looks_like_house_segment(segment)
            or _looks_like_standalone_house_number_segment(segment)
            or bool(find_trailing_house(segment))
        ):
            return True
    return False


def _looks_like_non_street_context_segment(value: str) -> bool:
    segment = str(value or "").strip()
    return (
        not segment
        or _looks_like_empty_street_locator_segment(segment)
        or _looks_like_object_context_segment(segment)
        or _looks_like_subject_segment(segment)
        or _looks_like_municipality_segment(segment)
        or looks_like_settlement_segment(segment)
        or _looks_like_parent_territory(segment)
        or _looks_like_territory_segment(segment)
        or _looks_like_microdistrict_segment(segment)
        or _looks_like_house_segment(segment)
        or _looks_like_house_modifier_segment(segment)
        or _looks_like_premise_segment(segment)
        or _looks_like_standalone_house_number_segment(segment)
        or looks_like_address_text(segment)
    )


def _looks_like_object_context_segment(value: str) -> bool:
    text = str(value or "").lower().replace("ё", "е")
    if not text:
        return False
    return bool(
        re.search(
            r"\b(?:маоу|мадоу|мбоу|мбдоу|моу|гбоу|сош|оош|нош|лицей|гимназия|"
            r"школа|детск\w*\s+сад|дошкольн\w+|образовательн\w+|учрежден\w+|"
            r"колледж|техникум|университет|корпус)\b",
            text,
        )
    )


def _same_street_segment(left: str | None, right: str | None) -> bool:
    return _norm_street_for_compare(left) == _norm_street_for_compare(right)


def _norm_street_for_compare(value: str | None) -> str:
    text = _norm(value)
    text = re.sub(
        r"^(?:ул|улица|пер|переулок|пр кт|проспект|пр д|проезд|пл|площадь|наб|набережная|о|остров)\s+",
        "",
        text,
    )
    return text


def _looks_like_premise_segment(value: str) -> bool:
    return bool(re.search(r"(?i)\b(?:помещ\.?|помещение|пом\.?|офис|оф\.?)\s*\d", str(value or "")))


def _looks_like_empty_street_locator_segment(value: str | None) -> bool:
    return bool(
        re.fullmatch(
            r"(?i)\s*(?:ул\.?|улица|пер\.?|переулок|проспект|пр-т\.?|пр-кт\.?|просп\.?|пр-д\.?|проезд|"
            r"б-р\.?|бульвар|наб\.?|набережная|пл\.?|площадь)\s*",
            str(value or "").strip(" ,"),
        )
    )


def _looks_like_microdistrict_segment(value: str) -> bool:
    text = str(value or "").lower().replace("ё", "е")
    return bool(re.search(r"\b(?:мкр\.?|микрорайон|квартал|кв-л\.?)\b", text))


def _looks_like_territory_segment(value: str) -> bool:
    text = str(value or "").lower().replace("ё", "е")
    return bool(re.search(r"\b(?:тер\.?|территория|тсн|снт|днт|сосн)\b", text))


def _microdistrict_tail_is_only_locator(segments: list[str], street: str | None) -> bool:
    return bool(street and _looks_like_microdistrict_segment(street) and segments and _norm(segments[-1]) == _norm(street))


def _clean(value: Any) -> str | None:
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip(" ,")
    if not text or text.lower() == "nan":
        return None
    return text


def _norm(value: str | None) -> str:
    text = str(value or "").lower().replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _level_norm(value: str | None) -> str:
    text = _norm(value)
    text = re.sub(
        r"\b(?:город|г|село|с|деревня|д|поселок|пос|п|р\s*п|пгт|хутор|х|станица|ст\s*ца)\b",
        " ",
        text,
    )
    return re.sub(r"\s+", " ", text).strip()
