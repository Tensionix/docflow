from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class SourceAnchor:
    source_path: Path
    source_kind: str
    sheet_name: str | None = None
    row_number: int | None = None
    fragment_index: int | None = None


@dataclass
class RawObservation:
    anchor: SourceAnchor
    payload: dict[str, Any]
    text: str | None = None


@dataclass(frozen=True)
class AddressFingerprint:
    territory: str
    street_words: str
    street_numbers: tuple[str, ...]
    house_base: str
    house_mods: str

    def compact_key(self) -> str:
        number_key = ",".join(self.street_numbers)
        return "|".join(
            [
                self.territory,
                self.street_words,
                number_key,
                self.house_base,
                self.house_mods,
            ]
        ).strip("|")


@dataclass(frozen=True)
class AddressParts:
    postal_index: str | None = None
    parent_subject: str | None = None
    subject: str | None = None
    federal_district: str | None = None
    autonomous_okrug: str | None = None
    municipality: str | None = None
    locality: str | None = None
    territory: str | None = None
    microdistrict: str | None = None
    street: str | None = None
    house: str | None = None
    house_kind: str | None = None
    house_base: str | None = None
    house_mods: str | None = None
    house_modifiers: tuple[str, ...] = ()
    premise: str | None = None
    street_numbers: tuple[str, ...] = ()
    rendered: str | None = None
    rendered_with_subject: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "postal_index": self.postal_index,
            "parent_subject": self.parent_subject,
            "subject": self.subject,
            "federal_district": self.federal_district,
            "autonomous_okrug": self.autonomous_okrug,
            "municipality": self.municipality,
            "locality": self.locality,
            "territory": self.territory,
            "microdistrict": self.microdistrict,
            "street": self.street,
            "house": self.house,
            "house_kind": self.house_kind,
            "house_base": self.house_base,
            "house_mods": self.house_mods,
            "house_modifiers": list(self.house_modifiers),
            "premise": self.premise,
            "street_numbers": list(self.street_numbers),
            "rendered": self.rendered,
            "rendered_with_subject": self.rendered_with_subject,
        }


@dataclass
class NormalizedObservation:
    anchor: SourceAnchor
    raw_payload: dict[str, Any]
    raw_text: str
    display_name: str | None
    address_text: str | None
    oktmo_code: str | None
    oktmo_name: str | None
    cadastral_number: str | None
    land_plot_number: str | None
    identity_key: str | None
    address_fingerprint: AddressFingerprint
    address_parts: AddressParts
    search_text: str


@dataclass
class ClassificationDecision:
    class_code: str | None
    class_name: str | None
    confidence: float
    rationale: str
    needs_review: bool = False
    basis: str = "unresolved"
    candidate_class_ids: list[str] = field(default_factory=list)


@dataclass
class ObjectCluster:
    cluster_key: str
    observations: list[NormalizedObservation] = field(default_factory=list)
    classification: ClassificationDecision | None = None


@dataclass
class QAFinding:
    level: str
    object_id: str
    message: str
