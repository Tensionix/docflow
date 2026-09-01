# Audion DocFlow

<!-- audion:release -->
<p align="center">
  <a href="https://audion.dev/downloads/docflow"><img alt="Windows" src="https://img.shields.io/badge/Windows-10%20%7C%2011-0b6db8?style=flat-square&logo=windows&logoColor=white"></a>
  <a href="https://github.com/Tensionix/docflow/releases/latest"><img alt="Release" src="https://img.shields.io/github/v/release/Tensionix/docflow?style=flat-square&label=release&color=e08a63"></a>
  <a href="https://github.com/Tensionix/docflow/releases"><img alt="Downloads" src="https://img.shields.io/github/downloads/Tensionix/docflow/total?style=flat-square&label=downloads&color=5fd08a"></a>
  <a href="https://github.com/Tensionix/docflow/blob/main/LICENSE"><img alt="License" src="https://img.shields.io/github/license/Tensionix/docflow?style=flat-square&color=5fd08a&logo=apache&logoColor=white&cacheSeconds=3600"></a>
</p>

**Version 2.0.2** · 2026-09-02 · 212.0 MB

- [Direct download](https://audion.dev/get/docflow/2.0.2/Audion_DocFlow_v2.0.2_Full.zip) — unmetered, no rate limits
- [Project page](https://audion.dev/downloads/docflow) — every version and how to install

<p align="center"><img src="docs/screenshot.png" alt="The program window" width="560"></p>

`SHA-256: b96477466b7db6bcf982eb5f37e3d8ee23682f2ef5565a7ee12d0ef454db89b0`

---

An **Audion** tool, published by [Tensionix](https://github.com/Tensionix).
<!-- /audion:release -->


[Русский](README_RU.md) · [User Guide](USER_GUIDE_EN.md)

Predictable cleaning and checking of Word documents, spreadsheets, CSV, and
markup. Works entirely offline.

## Why It Exists

A document that has passed through several hands carries the marks of that work:
coloured reviewer notes, highlighting, comments, struck-through text, unaccepted
revisions. Before submission all of it has to go — and go without taking anything
needed with it.

Manual cleaning is unreliable twice over. Something gets missed, and something
gets deleted along with what mattered: accept the revisions and a page break goes
with them; remove the struck-through text and the paragraph collapses.

## Principles

**Checking is separate from fixing.** First the document is checked and produces a
list of findings: where the font is not black, where there is highlighting, fill,
strikethrough, comments, unaccepted revisions. You look, then you decide.

For automated pipelines there is a strict check that returns a non-zero code — it
can be made a build condition.

**Removing struck-through text leaves the structure alone.** Only the text goes.
Paragraphs, tables, page and section breaks stay — otherwise the document falls
apart after cleaning.

**Invisible-character cleaning is conservative.** Only soft hyphens are removed.
Spaces, tabs, breaks, visible hyphens, and line-break control characters are left
alone: in a legal document a non-breaking space is there on purpose.

**A dry run shows the plan.** Before anything changes you can see exactly what
will be removed and what will be written to the output.

**The report shows the fragment, not coordinates.** Every hit is quoted with a
piece of text — already as it would look after the edit. Whether the replacement
is right can be judged without opening the document.

## What It Can Do

| area | about |
|---|---|
| Quality check | font colour, highlighting, fill, strikethrough, comments, revisions |
| Finalisation | one-step cleanup with a final check |
| Spacing and punctuation | checking and fixing |
| Find and replace | including morphological search: by the base form of a word, with case agreement |
| Styles | headings, appendices, lists, table and figure captions, contents, sections |
| Anomaly inspector | checking without modifying the document |

## About Morphological Search

Ordinary find-and-replace is useless in Russian: a word appears in six cases. Here
the search runs on the base form, and the replacement phrase agrees in case where
needed.

Two modes: replace what was found, or add beside it. The second is for when the
old wording may not be removed but the new one must appear.

The report quotes every replacement as a **whole sentence** — full stop to full
stop.

## Next

* [User Guide](USER_GUIDE_EN.md) — step by step.

---

## Technical Reference

### Layout

```
output\      processed files
report\      reports
._runtime\   temporary
```

User launchers sit in the root; the service and build layer is kept apart from the
user menu.

### Offline

Nothing is sent anywhere. Every check and edit runs locally.
