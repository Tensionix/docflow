#!/usr/bin/env python3
"""
DOCX low-level helpers using ZIP + lxml.

We use lxml only for XML manipulation inside DOCX parts.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from lxml import etree

# Namespaces used in WordprocessingML
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

def _etree_from_bytes(data: bytes) -> etree._ElementTree:
    parser = etree.XMLParser(remove_blank_text=False, recover=True, huge_tree=True)
    return etree.fromstring(data, parser=parser).getroottree()

def _etree_to_bytes(tree: etree._ElementTree) -> bytes:
    return etree.tostring(tree, encoding="UTF-8", xml_declaration=True, standalone=False)

def read_zip_map(docx_path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return {n: zf.read(n) for n in zf.namelist()}

def write_zip_map(out_path: Path, files: Dict[str, bytes]) -> None:
    # Preserve compression
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)

def list_xml_parts(files: Dict[str, bytes]) -> List[str]:
    # document + headers + footers are the typical places to scrub
    parts = []
    for name in files.keys():
        if not name.startswith("word/"):
            continue
        if name.endswith(".xml") and (
            name == "word/document.xml" or
            name.startswith("word/header") or
            name.startswith("word/footer") or
            name in ("word/footnotes.xml", "word/endnotes.xml")
        ):
            parts.append(name)
    return sorted(parts)


def list_style_parts(files: Dict[str, bytes]) -> List[str]:
    """Return DOCX parts that define styles and defaults.

    Some documents enforce formatting through styles (not direct run properties).
    These parts can carry colors/highlights/shading even when the visible text
    appears mostly clean.
    """
    parts: List[str] = []
    for name in ("word/styles.xml", "word/stylesWithEffects.xml", "word/numbering.xml"):
        if name in files:
            parts.append(name)
    return parts

def has_part(files: Dict[str, bytes], prefix: str) -> bool:
    return any(k.startswith(prefix) for k in files.keys())

def find_tracked_changes(files: Dict[str, bytes]) -> Dict[str, int]:
    tags = ["<w:ins", "<w:del", "<w:moveFrom", "<w:moveTo"]
    counts = {t: 0 for t in tags}
    for part in list_xml_parts(files) + ["word/document.xml"]:
        if part not in files:
            continue
        xml = files[part].decode("utf-8", errors="ignore")
        for t in tags:
            counts[t] += xml.count(t)
    return counts

def find_comments(files: Dict[str, bytes]) -> List[str]:
    return sorted([k for k in files.keys() if k.startswith("word/comments") and k.endswith(".xml")])

def remove_comment_relationships(rels_xml: bytes) -> bytes:
    tree = _etree_from_bytes(rels_xml)
    root = tree.getroot()
    # Relationships namespace is "rel"
    removed = 0
    for rel in list(root):
        # Type attribute varies, check substring
        typ = rel.get("Type", "")
        if "comments" in typ.lower():
            root.remove(rel)
            removed += 1
    return _etree_to_bytes(tree)

def strip_comment_markers(xml_bytes: bytes) -> bytes:
    tree = _etree_from_bytes(xml_bytes)
    root = tree.getroot()

    # Remove comment range start/end and reference
    for tag in ["commentRangeStart", "commentRangeEnd", "commentReference"]:
        for el in root.xpath(f".//w:{tag}", namespaces=NS):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    return _etree_to_bytes(tree)

def strip_shading_and_highlight(xml_bytes: bytes) -> bytes:
    tree = _etree_from_bytes(xml_bytes)
    root = tree.getroot()

    # Remove shading elements (background fills)
    for el in root.xpath(".//w:shd", namespaces=NS):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    # Remove highlights
    for el in root.xpath(".//w:highlight", namespaces=NS):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    return _etree_to_bytes(tree)

def force_black_color(xml_bytes: bytes) -> bytes:
    tree = _etree_from_bytes(xml_bytes)
    root = tree.getroot()

    # w:color is commonly found in run properties (w:rPr), but styles also use w:rPr.
    for el in root.xpath(".//w:rPr/w:color", namespaces=NS):
        # Treat themeColor as an explicit non-black signal for our purposes.
        theme_color = el.get(f"{{{NS['w']}}}themeColor")

        val = el.get(f"{{{NS['w']}}}val") or el.get("val")  # be tolerant
        if val is None:
            # If there's no val but themeColor is set, force black.
            if theme_color:
                el.set(f"{{{NS['w']}}}val", "000000")
            continue

        v = val.strip().lower()
        if v not in ("000000", "auto") or theme_color:
            # set to black
            el.set(f"{{{NS['w']}}}val", "000000")

        # Remove theme-based attributes to avoid re-coloring.
        for attr in ("themeColor", "themeShade", "themeTint"):
            key = f"{{{NS['w']}}}{attr}"
            if el.get(key) is not None:
                del el.attrib[key]
    return _etree_to_bytes(tree)


def _w_val(el: etree._Element, name: str = "val") -> Optional[str]:
    return el.get(f"{{{NS['w']}}}{name}") or el.get(name)


def _toggle_is_on(el: etree._Element) -> bool:
    """Interpret WordprocessingML toggle values; a missing value means enabled."""
    val = _w_val(el)
    if val is None:
        return True
    return val.strip().lower() not in ("0", "false", "off", "none")


def _has_enabled_strike(rpr: Optional[etree._Element]) -> bool:
    if rpr is None:
        return False
    for tag in ("strike", "dstrike"):
        for el in rpr.xpath(f"./w:{tag}", namespaces=NS):
            if _toggle_is_on(el):
                return True
    return False


def find_strikethrough_style_ids(files: Dict[str, bytes]) -> set[str]:
    """Find run style ids that apply strike/double-strike formatting."""
    style_ids: set[str] = set()
    for part in ("word/styles.xml", "word/stylesWithEffects.xml"):
        if part not in files:
            continue
        tree = _etree_from_bytes(files[part])
        root = tree.getroot()
        for style in root.xpath(".//w:style", namespaces=NS):
            style_id = _w_val(style, "styleId")
            if not style_id:
                continue
            rpr = style.find("w:rPr", namespaces=NS)
            if _has_enabled_strike(rpr):
                style_ids.add(style_id)
    return style_ids


def strip_strikethrough_marks(xml_bytes: bytes) -> bytes:
    """Remove strike/double-strike formatting tags without touching text."""
    tree = _etree_from_bytes(xml_bytes)
    root = tree.getroot()
    for tag in ("strike", "dstrike"):
        for el in root.xpath(f".//w:rPr/w:{tag}", namespaces=NS):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
    return _etree_to_bytes(tree)


def _blank_visible_text(text: str) -> str:
    # Keep the same character count to minimize Word reflow after cleanup.
    return "".join(ch if ch.isspace() else " " for ch in text)


def remove_strikethrough_text(
    xml_bytes: bytes,
    *,
    mode: str = "preserve-layout",
    strike_style_ids: Iterable[str] = (),
) -> bytes:
    """Remove visible text from strikethrough runs.

    mode="preserve-layout" replaces stricken characters with spaces of the
    same length. mode="delete" clears those text nodes.
    """
    strike_styles = set(strike_style_ids)
    tree = _etree_from_bytes(xml_bytes)
    root = tree.getroot()
    modified = False

    for run in root.xpath(".//w:r", namespaces=NS):
        rpr = run.find("w:rPr", namespaces=NS)
        is_struck = _has_enabled_strike(rpr)
        if not is_struck and rpr is not None and strike_styles:
            rstyle = rpr.find("w:rStyle", namespaces=NS)
            if rstyle is not None and (_w_val(rstyle) or "") in strike_styles:
                is_struck = True
        if not is_struck:
            continue

        for t in run.xpath("./w:t", namespaces=NS):
            if t.text is None:
                continue
            new_text = "" if mode == "delete" else _blank_visible_text(t.text)
            if new_text != t.text:
                t.text = new_text
                if new_text.startswith(" ") or new_text.endswith(" "):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                modified = True

        if rpr is not None:
            for tag in ("strike", "dstrike"):
                for el in rpr.xpath(f"./w:{tag}", namespaces=NS):
                    rpr.remove(el)
                    modified = True

    if not modified:
        return xml_bytes
    return _etree_to_bytes(tree)

def accept_changes_simple(xml_bytes: bytes) -> bytes:
    """
    Accept insertions, reject deletions.
    - Unwrap <w:ins> keeping its children
    - Remove <w:del> entirely
    - Remove moveFrom, unwrap moveTo
    """
    tree = _etree_from_bytes(xml_bytes)
    root = tree.getroot()

    # Remove deletions
    for el in root.xpath(".//w:del", namespaces=NS):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    # Unwrap insertions
    for el in root.xpath(".//w:ins", namespaces=NS):
        parent = el.getparent()
        if parent is None:
            continue
        idx = parent.index(el)
        for child in list(el):
            parent.insert(idx, child)
            idx += 1
        parent.remove(el)

    # Moves: remove moveFrom, unwrap moveTo
    for el in root.xpath(".//w:moveFrom", namespaces=NS):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    for el in root.xpath(".//w:moveTo", namespaces=NS):
        parent = el.getparent()
        if parent is None:
            continue
        idx = parent.index(el)
        for child in list(el):
            parent.insert(idx, child)
            idx += 1
        parent.remove(el)

    return _etree_to_bytes(tree)

def disable_track_revisions(settings_xml: bytes) -> bytes:
    tree = _etree_from_bytes(settings_xml)
    root = tree.getroot()
    # Remove <w:trackRevisions/> if present
    for el in root.xpath(".//w:trackRevisions", namespaces=NS):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    return _etree_to_bytes(tree)
