#!/usr/bin/env python3
"""Survey DOCX files before a style transfer: what is in the reference, what survived.

The reading code comes from the docx-restyle-by-template skill unchanged. The
entry point is this project's: a folder in, a report and per-document dumps out.

Per document it writes a flat dump - one line per block, with its style and any
surviving markers (_Toc bookmarks, SEQ/STYLEREF fields). That dump is the raw
material for the heading map the transfer needs. With a reference document it
also lists the reference styles by how often they are actually applied.
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

from pathlib import Path

from _office_common import find_docx_files, safe_mkdir, write_json_file

import argparse, collections, io, os, re, sys, zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return '{%s}%s' % (W, t)


def read_part(docx, name):
    with zipfile.ZipFile(docx) as z:
        try:
            return z.read(name)
        except KeyError:
            return None


def parse_part(docx, name):
    raw = read_part(docx, name)
    return etree.fromstring(raw) if raw else None


def ptext(el):
    return ''.join(t.text or '' for t in el.iter(q('t')))


def norm(s):
    return re.sub(r'\s+', ' ', s).strip()


def attrs(el):
    return ','.join('%s=%s' % (etree.QName(k).localname, v) for k, v in el.attrib.items())


# ------------------------------------------------------------------ стили
def style_map(styles_root):
    out = {}
    if styles_root is None:
        return out
    for st in styles_root.findall(q('style')):
        name = st.find(q('name'))
        out[st.get(q('styleId'))] = name.get(q('val')) if name is not None else '?'
    return out


def describe_styles(styles_root, usage):
    """стили образца: id, имя, наследование, ключевые свойства, частота"""
    lines = []
    if styles_root is None:
        return 'styles.xml отсутствует\n'
    for st in styles_root.findall(q('style')):
        sid = st.get(q('styleId'))
        used = usage.get(sid, 0)
        name = st.find(q('name'))
        name = name.get(q('val')) if name is not None else '?'
        based = st.find(q('basedOn'))
        head = '%-10s %-9s %-32s' % (sid, st.get(q('type')), name[:32])
        if based is not None:
            head += ' <- %s' % based.get(q('val'))
        head = '%6d  %s' % (used, head)
        props = []
        for part in ('pPr', 'rPr'):
            el = st.find(q(part))
            if el is None:
                continue
            for ch in el:
                tag = etree.QName(ch).localname
                if tag in ('rFonts', 'sz', 'b', 'i', 'jc', 'ind', 'spacing', 'numPr',
                           'keepNext', 'keepLines', 'pageBreakBefore', 'outlineLvl', 'color'):
                    a = attrs(ch)
                    if tag == 'numPr':
                        nid = ch.find(q('numId')); ilvl = ch.find(q('ilvl'))
                        a = 'numId=%s,ilvl=%s' % (nid.get(q('val')) if nid is not None else '-',
                                                  ilvl.get(q('val')) if ilvl is not None else '0')
                    props.append('%s(%s)' % (tag, a) if a else tag)
        lines.append(head + ('\n            ' + ' '.join(props) if props else ''))
    lines.sort(key=lambda s: -int(s.split()[0]))
    return '\n'.join(lines) + '\n'


# ------------------------------------------------------------------ дампы
def marks(p):
    """уцелевшие зацепки абзаца: закладки оглавления, поля, стиль, секция"""
    m = []
    for b in p.iter(q('bookmarkStart')):
        n = b.get(q('name')) or ''
        if n.startswith('_Toc'):
            m.append('TOC')
        elif n.startswith('_Ref'):
            m.append('REF')
    instr = ''.join(t.text or '' for t in p.iter(q('instrText')))
    if 'SEQ' in instr:
        m.append('SEQ')
    if 'STYLEREF' in instr:
        m.append('STYREF')
    if re.search(r'REF\s+_Ref', instr):
        m.append('XREF')
    if p.find('.//' + q('drawing')) is not None or p.find('.//' + q('pict')) is not None:
        m.append('IMG')
    pPr = p.find(q('pPr'))
    if pPr is not None:
        for tag, label in (('numPr', 'NUM'), ('sectPr', 'SECT'), ('keepNext', 'keep')):
            if pPr.find(q(tag)) is not None:
                m.append(label)
        jc = pPr.find(q('jc'))
        if jc is not None:
            m.append('jc:' + jc.get(q('val')))
        ind = pPr.find(q('ind'))
        if ind is not None:
            m.append('ind(%s)' % attrs(ind))
    r = p.find(q('r'))
    if r is not None:
        rPr = r.find(q('rPr'))
        if rPr is not None:
            f = rPr.find(q('rFonts')); sz = rPr.find(q('sz'))
            bits = []
            if f is not None:
                bits.append(f.get(q('ascii')) or '?')
            if sz is not None:
                bits.append('sz' + sz.get(q('val')))
            if rPr.find(q('b')) is not None:
                bits.append('b')
            if bits:
                m.append('run[%s]' % ','.join(bits))
    return ','.join(m)


def flat_dump(doc_root, smap, out_path):
    body = doc_root.find(q('body'))
    w = io.open(out_path, 'w', encoding='utf-8')
    usage = collections.Counter()
    i = 0
    for ch in body:
        i += 1
        if ch.tag == q('p'):
            pPr = ch.find(q('pPr'))
            ps = pPr.find(q('pStyle')) if pPr is not None else None
            sid = ps.get(q('val')) if ps is not None else None
            usage[sid or '(без стиля)'] += 1
            label = smap.get(sid, sid) if sid else ''
            t = norm(ptext(ch))
            if t:
                w.write('%04d P  [%s] %s | %s\n' % (i, label, marks(ch), t[:170]))
            else:
                w.write('%04d ·  [%s] %s\n' % (i, label, marks(ch)))
        elif ch.tag == q('tbl'):
            rows = ch.findall(q('tr'))
            ncol = len(rows[0].findall(q('tc'))) if rows else 0
            head = ' ¦ '.join(norm(ptext(tc))[:24] for tc in (rows[0].findall(q('tc'))[:5] if rows else []))
            pr = ch.find(q('tblPr'))
            lay = pr.find(q('tblLayout')) if pr is not None else None
            st = pr.find(q('tblStyle')) if pr is not None else None
            w.write('%04d TBL %dx%d layout=%s style=%s | %s\n' % (
                i, len(rows), ncol,
                lay.get(q('type')) if lay is not None else 'auto',
                smap.get(st.get(q('val')), st.get(q('val'))) if st is not None else '-', head))
        elif ch.tag == q('sectPr'):
            w.write('%04d ==== последний разрыв раздела ====\n' % i)
    w.close()
    return usage


# ------------------------------------------------------------------ сводка
def summarize(tag, docx, doc_root):
    out = ['=' * 70, '%s  %s' % (tag, os.path.basename(docx))]
    raw = read_part(docx, 'word/document.xml') or b''
    out.append('  document.xml: %.2f МБ, файл целиком %.0f КБ'
               % (len(raw) / 1048576, os.path.getsize(docx) / 1024))

    body = doc_root.find(q('body'))
    paras = doc_root.findall('.//' + q('p'))
    tbls = doc_root.findall('.//' + q('tbl'))
    out.append('  абзацев всего %d (из них верхнего уровня %d), таблиц %d, строк в таблицах %d'
               % (len(paras), len([c for c in body if c.tag == q('p')]),
                  len(tbls), len(doc_root.findall('.//' + q('tr')))))

    rsid = len(re.findall(r'w:rsid[A-Za-z]*="', raw.decode('utf-8', 'ignore')))
    out.append('  атрибутов ревизий rsid*: %d %s' % (rsid, '<- мусор, вычищается' if rsid > 1000 else ''))

    styles_used = collections.Counter(
        el.get(q('val')) for el in doc_root.iter(q('pStyle')))
    out.append('  стилей абзацев применено: %d' % len(styles_used))
    if not styles_used:
        out.append('    ВНИМАНИЕ: ни одного pStyle — форматирование снесено под ноль')

    direct = collections.Counter()
    for tag_ in ('sz', 'rFonts', 'ind', 'jc', 'numPr', 'b'):
        direct[tag_] = len(doc_root.findall('.//' + q(tag_)))
    out.append('  прямое форматирование: ' + ', '.join('%s=%d' % kv for kv in direct.items()))

    lay = collections.Counter()
    for t in tbls:
        pr = t.find(q('tblPr'))
        el = pr.find(q('tblLayout')) if pr is not None else None
        lay[el.get(q('type')) if el is not None else 'auto (медленно)'] += 1
    out.append('  раскладка таблиц: %s' % dict(lay))
    big = sorted((len(t.findall(q('tr'))) for t in tbls), reverse=True)[:5]
    out.append('  самые большие таблицы (строк): %s' % big)

    secs = doc_root.findall('.//' + q('sectPr'))
    orient = collections.Counter()
    for s in secs:
        pg = s.find(q('pgSz'))
        orient[(pg.get(q('orient')) or 'portrait') if pg is not None else '?'] += 1
    out.append('  разделов %d %s' % (len(secs), dict(orient)))

    instr = collections.Counter(
        norm(t.text or '')[:40] for t in doc_root.iter(q('instrText')))
    if instr:
        out.append('  поля: %s' % instr.most_common(6))
    bm = [b.get(q('name')) for b in doc_root.iter(q('bookmarkStart'))]
    out.append('  закладок %d, из них оглавления (_Toc) %d'
               % (len(bm), len([b for b in bm if (b or '').startswith('_Toc')])))
    return '\n'.join(out) + '\n'

# ------------------------------------------------------- batch entry point


def probe_one(source: Path, out_dir: Path, template: Path | None) -> dict:
    """Dump one document; returns the summary text and the files written."""
    safe_mkdir(out_dir)
    tgt_doc = parse_part(str(source), 'word/document.xml')
    smap_tgt = style_map(parse_part(str(source), 'word/styles.xml'))
    flat = out_dir / 'flat.txt'
    flat_dump(tgt_doc, smap_tgt, str(flat))

    written = [flat]
    summary = summarize('ПРИЁМНИК', str(source), tgt_doc)

    if template is not None:
        tpl_doc = parse_part(str(template), 'word/document.xml')
        tpl_styles = parse_part(str(template), 'word/styles.xml')
        usage = collections.Counter(el.get(q('val')) for el in tpl_doc.iter(q('pStyle')))
        usage.update(el.get(q('val')) for el in tpl_doc.iter(q('tblStyle')))
        styles_path = out_dir / 'reference_styles.txt'
        with io.open(str(styles_path), 'w', encoding='utf-8') as w:
            w.write('частота  id         тип       имя                              наследование\n')
            w.write(describe_styles(tpl_styles, usage))
        written.append(styles_path)
        summary = summarize('ЭТАЛОН  ', str(template), tpl_doc) + summary

    summary_path = out_dir / 'summary.txt'
    io.open(str(summary_path), 'w', encoding='utf-8').write(summary)
    written.append(summary_path)
    return {"summary": summary, "files": written}


def write_report(path: Path, input_root: Path, template: Path | None, results: list) -> None:
    safe_mkdir(path.parent)
    lines = ["# Разведка стилей DOCX", ""]
    lines.append("Папка: `%s`" % input_root)
    lines.append("Эталон: %s" % ("`%s`" % template if template else "не используется"))
    lines.append("")
    lines.append("Плоский дамп читается целиком - он единственный показывает структуру:")
    lines.append("где главы, где подписи таблиц, где перечни, а где мусор.")
    lines.append("")
    for item in results:
        lines.append("## %s" % item["source"].name)
        lines.append("")
        if item.get("error"):
            lines.append("**Ошибка:** %s" % item["error"])
            lines.append("")
            continue
        lines.append("```")
        lines.extend(item["summary"].rstrip().splitlines())
        lines.append("```")
        lines.append("")
        for path_written in item["files"]:
            lines.append("- `%s`" % path_written)
        lines.append("")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Survey DOCX files before a style transfer.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--input", help="Input folder with DOCX files")
    group.add_argument("--file", help="Single DOCX file")
    parser.add_argument("--outdir", default="output/style_probe", help="Folder for the dumps")
    parser.add_argument("--report", default="report/docx_style_probe.md", help="Markdown report path")
    parser.add_argument("--json-out", default="", help="Optional JSON report path")
    parser.add_argument("--template", default="", help="Optional reference DOCX to describe alongside")
    args = parser.parse_args()

    report_path = Path(args.report).resolve()
    out_dir = Path(args.outdir).resolve()

    if args.file:
        source_file = Path(args.file).resolve()
        input_root = source_file.parent
        docx_files = [source_file] if source_file.is_file() else []
        if not docx_files:
            print("[ERROR] DOCX file not found: %s" % source_file)
            return 2
    else:
        input_root = Path(args.input).resolve()
        if not input_root.exists():
            print("[ERROR] Input folder does not exist: %s" % input_root)
            return 2
        docx_files = find_docx_files(input_root)

    template: Path | None = None
    if args.template:
        template = Path(args.template).resolve()
        if not template.is_file():
            print("[ERROR] Reference DOCX not found: %s" % template)
            return 2

    if not docx_files:
        safe_mkdir(report_path.parent)
        report_path.write_text("# Разведка стилей DOCX\n\nФайлы DOCX не найдены.\n", encoding="utf-8")
        print("[WARN] No DOCX files found: %s" % input_root)
        print("[OK] Report: %s" % report_path)
        return 0

    results = []
    failed = False
    for source in docx_files:
        if template is not None and source.resolve() == template:
            print("[SKIP] reference document: %s" % source)
            continue
        target_dir = out_dir / source.stem
        item = {"source": source, "files": [], "summary": "", "error": ""}
        try:
            item.update(probe_one(source, target_dir, template))
            print("[OK] %s -> %s" % (source, target_dir))
        except Exception as exc:
            item["error"] = str(exc)
            failed = True
            print("[FAILED] %s: %s" % (source, exc))
        results.append(item)

    write_report(report_path, input_root, template, results)
    if args.json_out:
        json_path = Path(args.json_out).resolve()
        write_json_file(json_path, {
            "input": str(input_root),
            "template": str(template) if template else "",
            "documents": [
                {
                    "source": str(item["source"]),
                    "error": item["error"],
                    "files": [str(p) for p in item["files"]],
                }
                for item in results
            ],
        })
        print("[OK] JSON: %s" % json_path)
    print("[OK] Report: %s" % report_path)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
