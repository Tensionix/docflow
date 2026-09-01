# Audion DocFlow

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
