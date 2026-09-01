#!/usr/bin/env python3
"""Move a reference DOCX style base onto a document, and clean the XML.

The engine below comes from the docx-restyle-by-template skill unchanged: it is
the part proven on real volumes. Only the entry point differs - this project
processes a folder and writes a Markdown report, while the skill worked on one
file driven by a config.

Four ways to run, set by two switches:
    --template + --config   full transfer: style base, headings, captions, title
    --template alone        style base and cleanup, nothing re-marked
    --config alone          re-mark with the document's own styles, plus cleanup
    neither                 cleanup only (same as --clean-only)
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

from dataclasses import dataclass, field
from pathlib import Path

from _office_common import find_docx_files, mirrored_output_path, safe_mkdir, write_json_file

import argparse, io, json, os, re, shutil, sys, tempfile, zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
def q(t): return '{%s}%s' % (W, t)

# порядок дочерних элементов w:pPr по схеме — Word отвергает файл при нарушении
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
             'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
             'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
             'rPr', 'sectPr', 'pPrChange']
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
               'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
               'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']

XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
report = []
def say(msg): report.append(msg)


# --------------------------------------------------------------- мелкие утилиты
def ptext(el):
    return ''.join(t.text or '' for t in el.iter(q('t')))

def norm(s):
    return re.sub(r'\s+', ' ', s).strip()

def is_caps(s, min_letters=8):
    letters = [c for c in s if c.isalpha()]
    if len(letters) < min_letters:
        return False
    return sum(1 for c in letters if c.isupper()) / len(letters) >= 0.85

def ordered_insert(parent, el, order):
    name = etree.QName(el).localname
    idx = order.index(name) if name in order else 999
    for child in parent:
        if not isinstance(child.tag, str):
            continue
        cname = etree.QName(child).localname
        cidx = order.index(cname) if cname in order else 999
        if cidx > idx:
            child.addprevious(el)
            return
    parent.append(el)

def get_pPr(p):
    pPr = p.find(q('pPr'))
    if pPr is None:
        pPr = etree.Element(q('pPr'))
        p.insert(0, pPr)
    return pPr

def set_style(p, style_id):
    if not style_id:
        return
    pPr = get_pPr(p)
    old = pPr.find(q('pStyle'))
    if old is not None:
        pPr.remove(old)
    el = etree.Element(q('pStyle')); el.set(q('val'), style_id)
    ordered_insert(pPr, el, PPR_ORDER)

def set_numbering(p, num_id, ilvl=None):
    """переопределить нумерацию абзаца; num_id='0' — снять её совсем"""
    pPr = get_pPr(p)
    old = pPr.find(q('numPr'))
    if old is not None:
        pPr.remove(old)
    numPr = etree.Element(q('numPr'))
    if ilvl is not None:
        e = etree.SubElement(numPr, q('ilvl')); e.set(q('val'), str(ilvl))
    e = etree.SubElement(numPr, q('numId')); e.set(q('val'), str(num_id))
    ordered_insert(pPr, numPr, PPR_ORDER)

def append_text(target_p, text=' '):
    last = None
    for t in target_p.iter(q('t')):
        last = t
    if last is not None and not (last.text or '').endswith(' '):
        last.text = (last.text or '') + text
        last.set(XMLSPACE, 'preserve')


# ------------------------------------------------------------------- 1. ЧИСТКА
def strip_rsids(root):
    """атрибуты ревизий не влияют на вид, но занимают половину веса файла"""
    n = 0
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        for k in list(el.attrib):
            ln = etree.QName(k).localname
            if ln.startswith('rsid') or ln in ('paraId', 'textId'):
                del el.attrib[k]; n += 1
    return n

def rpr_key(r):
    rPr = r.find(q('rPr'))
    return etree.tostring(rPr, method='c14n') if rPr is not None else b''

SPECIAL_RUN = ('fldChar', 'instrText', 'drawing', 'pict', 'br', 'tab',
               'footnoteReference', 'endnoteReference', 'object')

def merge_runs(root):
    """склеить соседние раны с одинаковым форматированием (Word дробит их по правкам)"""
    merged = 0
    for p in root.iter(q('p')):
        prev = None
        for child in list(p):
            if child.tag != q('r'):
                prev = None; continue
            if any(child.find(q(t)) is not None for t in SPECIAL_RUN):
                prev = None; continue
            if prev is not None and rpr_key(prev) == rpr_key(child):
                pt, ct = prev.findall(q('t')), child.findall(q('t'))
                if pt and ct:
                    pt[-1].text = (pt[-1].text or '') + ''.join(t.text or '' for t in ct)
                    pt[-1].set(XMLSPACE, 'preserve')
                    p.remove(child); merged += 1
                    continue
            prev = child
    return merged

def drop_empty_props(root):
    n = 0
    for el in list(root.iter(q('rPr'))) + list(root.iter(q('pPr'))):
        if len(el) == 0 and not el.attrib:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el); n += 1
    return n

END_OK = '.;:!?»)"…'

def join_broken(body):
    """
    Собрать абзацы, разорванные при потере форматирования: там, где раньше стоял
    мягкий перенос, конвертер оставил два абзаца. Признак надёжный — предыдущий
    кусок не кончается знаком конца, а следующий начинается со строчной буквы.
    """
    joins = []
    children = [c for c in body if c.tag == q('p')]
    i = 0
    while i < len(children) - 1:
        cur, nxt = children[i], children[i + 1]
        if cur.getnext() is not nxt:
            i += 1; continue
        pPr = cur.find(q('pPr'))
        if pPr is not None and pPr.find(q('sectPr')) is not None:
            i += 1; continue
        ct, nt = norm(ptext(cur)), norm(ptext(nxt))
        ok = (ct and nt and ct[-1] not in END_OK and not ct.endswith(('-', '–'))
              and nt[0].isalpha() and nt[0].islower())
        if ok:
            append_text(cur)
            for r in [x for x in nxt if x.tag in (q('r'), q('hyperlink'), q('bookmarkStart'),
                                                  q('bookmarkEnd'))]:
                cur.append(r)
            body.remove(nxt)
            joins.append((ct[-45:], nt[:45]))
            children.pop(i + 1)
            continue
        i += 1
    return joins


# --------------------------------------------------------------- 2. РАЗМЕТКА
class Rules(object):
    def __init__(self, cfg):
        self.styles = cfg.get('styles', {})
        self.headings = sorted(cfg.get('headings', []),
                               key=lambda r: 0 if r.get('under') else 1)
        caps = cfg.get('captions', {})
        self.cap_re = re.compile(caps.get('pattern', r'^Таблица\s*\d+(\.\d+)*\s*$'))
        self.cap_title_caps = caps.get('title_caps', True)
        self.lists = cfg.get('lists', {'detect': True})

    def heading_style(self, level):
        return (self.styles.get('heading') or {}).get(str(level))

    def match_heading(self, text, parents):
        """вернуть (уровень, правило) или (None, None)"""
        for rule in self.headings:
            under = rule.get('under')
            if under and under not in parents:
                continue
            kind = rule.get('match', 'exact')
            if kind == 'exact' and text in rule.get('values', []):
                return rule.get('level'), rule
            if kind == 'prefix' and any(text.startswith(v) for v in rule.get('values', [])):
                return rule.get('level'), rule
            if kind == 'regex' and re.match(rule['pattern'], text):
                return rule.get('level'), rule
        return None, None


def classify(body, rules, stats):
    """расставить стили абзацам верхнего уровня; вернуть дерево заголовков"""
    st = rules.styles
    heads, parents = [], {}
    pending_title = False     # прошли номер таблицы, ждём её название
    last_title = None         # последний абзац-название (для склейки продолжений)
    in_list = prev_item = False

    def bump(k): stats[k] = stats.get(k, 0) + 1

    for c in [x for x in body if x.tag in (q('p'), q('tbl'))]:
        if c.tag == q('tbl'):
            pending_title = False; last_title = None
            continue
        t = norm(ptext(c))

        if not t:
            set_style(c, st.get('text')); bump('пустой'); continue

        # заголовки проверяем раньше подписей: и те и другие бывают капсом
        level, rule = rules.match_heading(t, parents.values())
        if level:
            set_style(c, rules.heading_style(level))
            num = rule.get('numbering', 'auto')
            if num == 'none':
                set_numbering(c, '0')
            elif isinstance(num, dict):
                set_numbering(c, num['numId'], num.get('ilvl'))
            parents[level] = t
            for deeper in [k for k in parents if k > level]:
                parents.pop(deeper)
            heads.append((level, t))
            bump('заголовок %d' % level)
            pending_title = in_list = prev_item = False; last_title = None
            continue

        # номер таблицы отдельным абзацем перед названием
        if st.get('caption_number') and rules.cap_re.match(t):
            set_style(c, st['caption_number'])
            pending_title = True; last_title = None; in_list = prev_item = False
            bump('подпись: номер'); continue

        # название таблицы капсом сразу за номером; продолжение бывает коротким
        if pending_title and st.get('caption_title') and \
                (is_caps(t, 3 if last_title is not None else 8) or not rules.cap_title_caps):
            if last_title is not None and last_title.getnext() is c:
                append_text(last_title)
                for r in [x for x in c if x.tag in (q('r'), q('hyperlink'))]:
                    last_title.append(r)
                body.remove(c); bump('склейка названия'); continue
            set_style(c, st['caption_title'])
            last_title = c; in_list = prev_item = False
            bump('подпись: название'); continue
        last_title = None

        # перечни: маркер вручную, строчная буква, либо блок после двоеточия
        item = False
        first, last = t[0], t[-1]
        if rules.lists.get('detect', True) and st.get('list_bullet'):
            if first in '-–—•':
                for tt in c.iter(q('t')):
                    if tt.text and re.match(r'^\s*[-–—•]', tt.text):
                        tt.text = re.sub(r'^\s*[-–—•]\s*', '', tt.text)
                        break
                item = True
            elif first.isalpha() and first.islower():
                item = True
            elif in_list and last == ';':
                item = True                      # пункт с заглавной или цифры — тоже пункт
            elif in_list and prev_item and last == '.' and not is_caps(t):
                item = True                      # последний пункт перечня
        if item:
            set_style(c, st['list_bullet']); bump('перечень')
            prev_item = True
            in_list = in_list or last in ';:'
            if last == '.':
                in_list = prev_item = False
            continue

        set_style(c, st.get('text')); bump('текст')
        in_list = (last == ':'); prev_item = False

    return heads


# ---------------------------------------------- 3. НУМЕРАЦИЯ ПОДПИСЕЙ ПОЛЯМИ
def scan_fields(p):
    """разобрать абзац на поля Word: [{instr, runs, result}]"""
    fields, stack = [], []
    for child in list(p):
        if child.tag != q('r'):
            continue
        fc = child.find(q('fldChar'))
        if fc is not None:
            kind = fc.get(q('fldCharType'))
            if kind == 'begin':
                stack.append({'instr': '', 'runs': [child], 'result': [], 'sep': False})
                continue
            if stack:
                cur = stack[-1]; cur['runs'].append(child)
                if kind == 'separate':
                    cur['sep'] = True
                elif kind == 'end':
                    done = stack.pop()
                    if stack:
                        stack[-1]['runs'].append(child)
                    fields.append(done)
            continue
        if stack:
            cur = stack[-1]; cur['runs'].append(child)
            it = child.find(q('instrText'))
            if it is not None:
                cur['instr'] += it.text or ''
            elif cur['sep']:
                cur['result'].append(child)
    return fields

def set_field_result(field, text):
    """подменить закешированный результат поля, чтобы номера были верны и до F9"""
    first = True
    for r in field['result']:
        for t in r.findall(q('t')):
            if first:
                t.text = text; t.set(XMLSPACE, 'preserve'); first = False
            else:
                t.text = ''
    return not first

def renumber_captions(body, cfg, caption_style):
    """
    «Таблица 2.1» -> сквозная «Таблица 1, 2, 3...».
    Номер главы даёт поле STYLEREF, счётчик — SEQ с ключом «\\s 1» (сброс на главе).
    Убираем первое вместе с точкой-разделителем, у второго снимаем сброс.
    """
    mode = (cfg.get('captions') or {}).get('renumber', 'keep')
    if mode == 'keep' or not caption_style:
        return 0, 0
    label = (cfg.get('captions') or {}).get('label', 'Таблица')
    counter, bookmark2num = 0, {}
    for p in body.iter(q('p')):
        pPr = p.find(q('pPr'))
        ps = pPr.find(q('pStyle')) if pPr is not None else None
        if ps is None or ps.get(q('val')) != caption_style:
            continue
        fields = scan_fields(p)
        seq = [f for f in fields if 'SEQ' in f['instr']]
        if not seq:
            continue
        counter += 1
        if mode == 'sequential':
            for f in [x for x in fields if 'STYLEREF' in x['instr']]:
                nxt = f['runs'][-1].getnext()
                for r in f['runs']:
                    p.remove(r)
                if nxt is not None and nxt.tag == q('r') and \
                        norm(''.join(t.text or '' for t in nxt.findall(q('t')))) == '.':
                    p.remove(nxt)
            for f in seq:
                for r in f['runs']:
                    it = r.find(q('instrText'))
                    if it is not None and 'SEQ' in (it.text or ''):
                        it.text = re.sub(r'\s*\\s\s*\d+\s*', ' ', it.text)
                set_field_result(f, str(counter))
        for bm in p.findall(q('bookmarkStart')):
            bookmark2num[bm.get(q('name'))] = counter

    refs = 0
    for p in body.iter(q('p')):
        for f in scan_fields(p):
            m = re.search(r'REF\s+(\S+)', f['instr'])
            if not m or 'STYLEREF' in f['instr']:
                continue
            num = bookmark2num.get(m.group(1))
            if num and set_field_result(f, '%s %d' % (label, num)):
                refs += 1
    return counter, refs


# ------------------------------------------------------------- 4. ТАБЛИЦЫ
def set_fixed_layout(tbl):
    """
    Фиксированная раскладка — главный ускоритель. При auto Word пересчитывает
    ширины по содержимому каждой ячейки; на таблице в 1500 строк это заметно.
    """
    pr = tbl.find(q('tblPr'))
    if pr is None:
        return
    lay = pr.find(q('tblLayout'))
    if lay is None:
        lay = etree.Element(q('tblLayout'))
        ordered_insert(pr, lay, TBLPR_ORDER)
    lay.set(q('type'), 'fixed')

def format_cells(body, cfg, skip_tables):
    st = cfg.get('styles', {})
    hdr, ctr, left = st.get('table_header'), st.get('table_cell_center'), st.get('table_cell_left')
    limit = st.get('cell_left_threshold', 60)
    if not (hdr or ctr or left):
        return 0
    n = 0
    for ti, tbl in enumerate(body.iter(q('tbl'))):
        if ti in skip_tables:
            continue
        rows = tbl.findall(q('tr'))
        for ri, tr in enumerate(rows):
            trPr = tr.find(q('trPr'))
            is_head = (ri == 0) or (trPr is not None and trPr.find(q('tblHeader')) is not None)
            for tc in tr.findall(q('tc')):
                for p in tc.findall(q('p')):
                    t = norm(ptext(p))
                    style = hdr if is_head else (left if len(t) > limit else ctr)
                    set_style(p, style); n += 1
    return n

def clone_pPr_rPr(src_p, dst_p):
    old = dst_p.find(q('pPr'))
    keep_sect = old.find(q('sectPr')) if old is not None else None
    if old is not None:
        dst_p.remove(old)
    src_pPr = src_p.find(q('pPr'))
    new = etree.fromstring(etree.tostring(src_pPr)) if src_pPr is not None else etree.Element(q('pPr'))
    for junk in new.findall(q('sectPr')):
        new.remove(junk)
    if keep_sect is not None:
        new.append(keep_sect)
    dst_p.insert(0, new)
    src_r = src_p.find(q('r'))
    src_rPr = src_r.find(q('rPr')) if src_r is not None else None
    if src_rPr is not None:
        for r in dst_p.findall(q('r')):
            old_r = r.find(q('rPr'))
            if old_r is not None:
                r.remove(old_r)
            r.insert(0, etree.fromstring(etree.tostring(src_rPr)))

def clone_title(tgt_body, tpl_body, cfg, first_heading_style):
    """
    Титул устроен одинаково в документах одной серии, поэтому его дешевле
    клонировать по позициям, чем описывать правилами: те же таблицы с гербом
    и реквизитами, тот же набор пустых абзацев.
    """
    tcfg = cfg.get('title') or {}
    skip = set()
    n_par = n_cell = 0

    def before_first_heading(body):
        out = []
        for c in body:
            if c.tag != q('p'):
                continue
            pPr = c.find(q('pPr'))
            ps = pPr.find(q('pStyle')) if pPr is not None else None
            if ps is not None and ps.get(q('val')) == first_heading_style:
                break
            out.append(c)
        return out

    if tcfg.get('paragraphs'):
        tgt_ps, tpl_ps = before_first_heading(tgt_body), before_first_heading(tpl_body)
        for i, kp in enumerate(tgt_ps):
            if not tpl_ps:
                break
            clone_pPr_rPr(tpl_ps[min(i, len(tpl_ps) - 1)], kp); n_par += 1

    count = int(tcfg.get('tables', 0))
    if count:
        tgt_t = [c for c in tgt_body if c.tag == q('tbl')]
        tpl_t = [c for c in tpl_body if c.tag == q('tbl')]
        for idx in range(min(count, len(tgt_t), len(tpl_t))):
            skip.add(idx)
            tgt_rows, tpl_rows = tgt_t[idx].findall(q('tr')), tpl_t[idx].findall(q('tr'))
            for ri, tr in enumerate(tgt_rows):
                src_tr = tpl_rows[min(ri, len(tpl_rows) - 1)]
                src_cells = src_tr.findall(q('tc'))
                for ci, tc in enumerate(tr.findall(q('tc'))):
                    if not src_cells:
                        continue
                    src_tc = src_cells[min(ci, len(src_cells) - 1)]
                    src_ps = src_tc.findall(q('p'))
                    for pi, p in enumerate(tc.findall(q('p'))):
                        if not src_ps:
                            continue
                        clone_pPr_rPr(src_ps[min(pi, len(src_ps) - 1)], p); n_cell += 1

    extra = tcfg.get('plain_table')          # один стиль на все ячейки указанной таблицы
    if extra:
        idx = int(extra.get('index'))
        tgt_t = [c for c in tgt_body if c.tag == q('tbl')]
        if idx < len(tgt_t):
            skip.add(idx)
            for tr in tgt_t[idx].findall(q('tr')):
                for tc in tr.findall(q('tc')):
                    for p in tc.findall(q('p')):
                        set_style(p, extra['style']); n_cell += 1
    return n_par, n_cell, skip


# --------------------------------------------------- 5. ОТСЕВ ЛИШНИХ СТИЛЕЙ
def prune_styles(word_dir, hide_latent=True):
    """
    Оставить только применённые стили и связанные с ними по basedOn/next/link.
    Эталон обычно тащит сотни стилей, накопленных за годы; в галерее Word они
    выглядят как мусор и мешают человеку выбрать нужный.
    """
    spath = os.path.join(word_dir, 'styles.xml')
    npath = os.path.join(word_dir, 'numbering.xml')
    if not os.path.exists(spath):
        return None

    used = set()
    for fn in sorted(os.listdir(word_dir)):
        if not fn.endswith('.xml') or fn in ('styles.xml', 'numbering.xml'):
            continue
        root = etree.parse(os.path.join(word_dir, fn)).getroot()
        for tag in ('pStyle', 'rStyle', 'tblStyle'):
            for el in root.iter(q(tag)):
                used.add(el.get(q('val')))
    used.discard(None)

    stree = etree.parse(spath); sroot = stree.getroot()
    styles = {s.get(q('styleId')): s for s in sroot.findall(q('style'))}
    before = len(styles)
    keep = {sid for sid, s in styles.items() if s.get(q('default')) == '1'} | (used & set(styles))
    frontier = set(keep)
    while frontier:
        nxt = set()
        for sid in frontier:
            s = styles.get(sid)
            if s is None:
                continue
            for tag in ('basedOn', 'next', 'link'):
                el = s.find(q(tag))
                if el is not None and el.get(q('val')) in styles and el.get(q('val')) not in keep:
                    nxt.add(el.get(q('val')))
        keep |= nxt; frontier = nxt
    for sid, s in styles.items():
        if sid not in keep:
            sroot.remove(s)

    lat = sroot.find(q('latentStyles'))
    if lat is not None and hide_latent:
        lat.set(q('defSemiHidden'), '1')
        lat.set(q('defUnhideWhenUsed'), '1')
        lat.set(q('defQFormat'), '0')
        for e in lat.findall(q('lsdException')):
            lat.remove(e)

    n_before = n_after = a_before = a_after = 0
    if os.path.exists(npath):
        ntree = etree.parse(npath); nroot = ntree.getroot()
        nums = {n.get(q('numId')): n for n in nroot.findall(q('num'))}
        abstracts = {a.get(q('abstractNumId')): a for a in nroot.findall(q('abstractNum'))}
        n_before, a_before = len(nums), len(abstracts)
        want_num = set()
        droot = etree.parse(os.path.join(word_dir, 'document.xml')).getroot()
        for el in droot.iter(q('numId')):
            want_num.add(el.get(q('val')))
        for sid in keep:
            for el in styles[sid].iter(q('numId')):
                want_num.add(el.get(q('val')))
        want_num.discard('0')
        want_abs = set()
        for nid in list(want_num):
            n = nums.get(nid)
            if n is None:
                continue
            aid = n.find(q('abstractNumId'))
            if aid is not None:
                want_abs.add(aid.get(q('val')))
        for nid, n in nums.items():
            if nid not in want_num:
                nroot.remove(n)
        for aid, a in abstracts.items():
            if aid not in want_abs:
                nroot.remove(a)
        for a in nroot.findall(q('abstractNum')):
            for lvl in a.findall(q('lvl')):
                ps = lvl.find(q('pStyle'))
                if ps is not None and ps.get(q('val')) not in keep:
                    lvl.remove(ps)
        ntree.write(npath, xml_declaration=True, encoding='UTF-8', standalone=True)
        n_after, a_after = len(want_num), len(want_abs)

    stree.write(spath, xml_declaration=True, encoding='UTF-8', standalone=True)
    return {'before': before, 'after': len(keep), 'num': (n_before, n_after),
            'abs': (a_before, a_after),
            'kept': sorted((styles[s].find(q('name')).get(q('val')), s) for s in keep)}


# ------------------------------------------------------ 6. СЛУЖЕБНЫЕ ЧАСТИ
def patch_table_styles(spath, ids):
    """
    Табличные стили эталона обычно наследуют базовый абзацный стиль вместе с
    красной строкой — в ячейке она не нужна.
    """
    if not ids or not os.path.exists(spath):
        return
    tree = etree.parse(spath); root = tree.getroot()
    for st in root.findall(q('style')):
        if st.get(q('styleId')) not in ids:
            continue
        pPr = st.find(q('pPr'))
        if pPr is None:
            pPr = etree.Element(q('pPr'))
            rPr = st.find(q('rPr'))          # по схеме pPr идёт перед rPr
            (rPr.addprevious(pPr) if rPr is not None else st.append(pPr))
        ind = pPr.find(q('ind'))
        if ind is None:
            ind = etree.Element(q('ind')); ordered_insert(pPr, ind, PPR_ORDER)
        ind.set(q('firstLine'), '0'); ind.set(q('left'), '0')
        sp = pPr.find(q('spacing'))
        if sp is None:
            sp = etree.Element(q('spacing')); ordered_insert(pPr, sp, PPR_ORDER)
        sp.set(q('before'), '0'); sp.set(q('after'), '0')
    tree.write(spath, xml_declaration=True, encoding='UTF-8', standalone=True)

SETTINGS_AFTER = ['hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids',
                  'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'shapeDefaults',
                  'decimalSymbol', 'listSeparator']

def patch_settings(path, update_fields=True):
    if not os.path.exists(path):
        return
    tree = etree.parse(path); root = tree.getroot()
    for el in root.findall(q('rsids')):
        root.remove(el)
    if update_fields and root.find(q('updateFields')) is None:
        el = etree.Element(q('updateFields')); el.set(q('val'), 'true')
        placed = False
        for child in root:
            if isinstance(child.tag, str) and etree.QName(child).localname in SETTINGS_AFTER:
                child.addprevious(el); placed = True; break
        if not placed:
            root.append(el)
    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)

def ensure_part_registered(base, part, content_type, rel_type, rel_target):
    ct_path = os.path.join(base, '[Content_Types].xml')
    tree = etree.parse(ct_path); root = tree.getroot()
    if part not in {o.get('PartName') for o in root.findall('{%s}Override' % CT)}:
        el = etree.SubElement(root, '{%s}Override' % CT)
        el.set('PartName', part); el.set('ContentType', content_type)
        tree.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    rels_path = os.path.join(base, 'word', '_rels', 'document.xml.rels')
    tree = etree.parse(rels_path); root = tree.getroot()
    rels = root.findall('{%s}Relationship' % PKG)
    if rel_target not in {r.get('Target') for r in rels}:
        ids = {r.get('Id') for r in rels}
        i = 1
        while ('rId%d' % i) in ids:
            i += 1
        el = etree.SubElement(root, '{%s}Relationship' % PKG)
        el.set('Id', 'rId%d' % i); el.set('Type', rel_type); el.set('Target', rel_target)
        tree.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)


def unpack(docx, dest):
    with zipfile.ZipFile(docx) as z:
        for info in z.infolist():
            if info.filename.startswith(('/', '..')) or '..' in info.filename.split('/'):
                continue
            z.extract(info, dest)

def pack(src, dst):
    if os.path.exists(dst):
        os.remove(dst)
    zf = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
    first = '[Content_Types].xml'
    zf.write(os.path.join(src, first), first)
    for root, dirs, files in os.walk(src):
        for f in sorted(files):
            full = os.path.join(root, f)
            rel = os.path.relpath(full, src).replace(os.sep, '/')
            if rel != first:
                zf.write(full, rel)
    zf.close()

def sanity_check(docx):
    """каждая часть должна остаться разбираемым XML — грубые поломки видно сразу"""
    bad = []
    with zipfile.ZipFile(docx) as z:
        for info in z.infolist():
            if info.filename.endswith('.xml') or info.filename.endswith('.rels'):
                try:
                    etree.fromstring(z.read(info.filename))
                except Exception as e:
                    bad.append('%s: %s' % (info.filename, e))
    return bad

# ------------------------------------------------------- batch entry point

DEFAULT_COPY_PARTS = ["word/styles.xml", "word/numbering.xml",
                      "word/fontTable.xml", "word/theme/theme1.xml"]


@dataclass
class RestyleResult:
    source: Path
    output: Path
    status: str = "OK"
    error: str = ""
    size_before: int = 0
    size_after: int = 0
    lines: list[str] = field(default_factory=list)


def restyle_document(source: Path, output: Path, template: Path | None,
                     cfg: dict, use_template: bool, markup: bool) -> list[str]:
    """Run one document through the skill pipeline; returns the report lines.

    Stage order is the skill's and must not be reshuffled: cleanup, paragraph
    joins, styling, caption numbering, tables and title, style pruning.
    """
    del report[:]
    clean = cfg.get("clean", {})
    work = tempfile.mkdtemp(prefix="restyle_")
    try:
        tgt_dir = os.path.join(work, "target")
        unpack(str(source), tgt_dir)

        if use_template and template is not None:
            with zipfile.ZipFile(str(template)) as z:
                names = set(z.namelist())
                for part in cfg.get("copy_parts", DEFAULT_COPY_PARTS):
                    if part not in names:
                        say("  нет в эталоне, пропущено: %s" % part)
                        continue
                    dest = os.path.join(tgt_dir, part.replace("/", os.sep))
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                    with io.open(dest, "wb") as f:
                        f.write(z.read(part))
            if os.path.exists(os.path.join(tgt_dir, "word", "numbering.xml")):
                ensure_part_registered(
                    tgt_dir, "/word/numbering.xml",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                    "numbering.xml")
            patch_table_styles(os.path.join(tgt_dir, "word", "styles.xml"),
                               set(cfg.get("flatten_indent_styles", [])))

        doc_path = os.path.join(tgt_dir, "word", "document.xml")
        size_before = os.path.getsize(doc_path)
        tree = etree.parse(doc_path)
        root = tree.getroot()
        body = root.find(q("body"))

        say("ЧИСТКА")
        if clean.get("rsids", True):
            say("  атрибутов ревизий удалено: %d" % strip_rsids(root))
        if clean.get("merge_runs", True):
            say("  ранов слито: %d" % merge_runs(root))
        if clean.get("drop_empty_props", True):
            say("  пустых pPr/rPr убрано: %d" % drop_empty_props(root))
        if clean.get("join_broken_paragraphs", True):
            joins = join_broken(body)
            say("  склеено разорванных абзацев: %d" % len(joins))
            for x, y in joins:
                say("     ...%s + %s..." % (x, y))

        stats, heads = {}, []
        if markup:
            rules = Rules(cfg)
            heads = classify(body, rules, stats)
            n_caps, n_refs = renumber_captions(body, cfg, (cfg.get("styles") or {}).get("caption_number"))
            n_par = n_cell = 0
            skip = set()
            if use_template and template is not None:
                tpl_root = etree.fromstring(read_template_document(str(template)))
                n_par, n_cell, skip = clone_title(body, tpl_root.find(q("body")), cfg,
                                                  rules.heading_style(1))
            n_cells = format_cells(body, cfg, skip)
            say("")
            say("РАЗМЕТКА")
            for k, v in sorted(stats.items(), key=lambda x: -x[1]):
                say("  %-24s %d" % (k, v))
            say("  ячеек таблиц оформлено: %d" % n_cells)
            if use_template:
                say("  титул клонирован: абзацев %d, ячеек %d" % (n_par, n_cell))
            if n_caps:
                say("  подписей перенумеровано: %d, ссылок обновлено: %d" % (n_caps, n_refs))

        if clean.get("fixed_table_layout", True):
            n = 0
            for tbl in body.iter(q("tbl")):
                set_fixed_layout(tbl)
                n += 1
            say("  таблиц переведено на фиксированную раскладку: %d" % n)

        tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        patch_settings(os.path.join(tgt_dir, "word", "settings.xml"),
                       clean.get("update_fields", True))

        if clean.get("prune_styles", True):
            pruned = prune_styles(os.path.join(tgt_dir, "word"), clean.get("hide_latent_styles", True))
            if pruned:
                say("")
                say("ОТСЕВ СТИЛЕЙ")
                say("  определений: %d -> %d" % (pruned["before"], pruned["after"]))
                say("  нумераций: %d -> %d, abstractNum: %d -> %d"
                    % (pruned["num"][0], pruned["num"][1], pruned["abs"][0], pruned["abs"][1]))
                for name, sid in pruned["kept"]:
                    say("     %-28s (id=%s)" % (name, sid))

        safe_mkdir(output.parent)
        pack(tgt_dir, str(output))
        bad = sanity_check(str(output))
        say("")
        say("ИТОГ")
        say("  document.xml: %.2f МБ -> %.2f МБ"
            % (size_before / 1048576, os.path.getsize(doc_path) / 1048576))
        say("  файл: %.0f КБ -> %.0f КБ" % (os.path.getsize(str(source)) / 1024,
                                            os.path.getsize(str(output)) / 1024))
        say("  XML-проверка: %s" % ("ОШИБКИ: %s" % bad if bad else "все части разбираются"))
        if heads:
            say("")
            say("ДЕРЕВО ЗАГОЛОВКОВ (%d)" % len(heads))
            for lvl, t in heads:
                say("%s%s" % ("    " * (lvl - 1), t[:110]))
        if bad:
            raise RuntimeError("XML damaged after restyle: %s" % "; ".join(bad))
        return list(report)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def write_report(path: Path, input_root: Path, template: Path | None,
                 config_path: Path | None, markup: bool, results: list) -> None:
    safe_mkdir(path.parent)
    ok = [r for r in results if r.status == "OK"]
    lines = ["# Перенос стилей по эталону", ""]
    lines.append("Папка: `%s`" % input_root)
    lines.append("")
    lines.append("Эталон: %s" % ("`%s`" % template if template else "не используется"))
    lines.append("Карта заголовков: %s" % ("`%s`" % config_path if markup else "не задана, разметка не выполнялась"))
    lines.append("")
    lines.append("Обработано: %d из %d" % (len(ok), len(results)))
    lines.append("")
    for result in results:
        lines.append("## %s" % result.source.name)
        lines.append("")
        if result.status != "OK":
            lines.append("**Ошибка:** %s" % result.error)
            lines.append("")
            continue
        lines.append("Результат: `%s`" % result.output)
        lines.append("")
        lines.append("```")
        lines.extend(result.lines)
        lines.append("```")
        lines.append("")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_json(input_root: Path, results: list) -> dict:
    return {
        "input": str(input_root),
        "documents": [
            {
                "source": str(r.source),
                "output": str(r.output),
                "status": r.status,
                "error": r.error,
                "size_before": r.size_before,
                "size_after": r.size_after,
            }
            for r in results
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Restyle DOCX files against a reference document.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--input", help="Input folder with DOCX files")
    group.add_argument("--file", help="Single DOCX file")
    parser.add_argument("--outdir", default="output/restyled", help="Output folder for batch mode")
    parser.add_argument("--out", default="", help="Output DOCX path for single-file mode")
    parser.add_argument("--report", default="report/docx_restyle.md", help="Markdown report path")
    parser.add_argument("--json-out", default="", help="Optional JSON report path")
    parser.add_argument("--template", default="", help="Reference DOCX whose style base is copied over")
    parser.add_argument("--config", default="", help="JSON map of headings, captions and styles")
    parser.add_argument("--clean-only", action="store_true", help="Cleanup only: no style base, no markup")
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
    if args.template and not args.clean_only:
        template = Path(args.template).resolve()
        if not template.is_file():
            print("[ERROR] Reference DOCX not found: %s" % template)
            return 2

    cfg: dict = {}
    config_path: Path | None = None
    if args.config and not args.clean_only:
        config_path = Path(args.config).resolve()
        if not config_path.is_file():
            print("[ERROR] Config not found: %s" % config_path)
            return 2
        cfg = json.load(io.open(str(config_path), encoding="utf-8"))

    use_template = template is not None
    markup = bool(cfg) and not args.clean_only

    if not docx_files:
        safe_mkdir(report_path.parent)
        report_path.write_text("# Перенос стилей по эталону\n\nФайлы DOCX не найдены.\n", encoding="utf-8")
        print("[WARN] No DOCX files found: %s" % input_root)
        print("[OK] Report: %s" % report_path)
        return 0

    results: list[RestyleResult] = []
    for source in docx_files:
        if template is not None and source.resolve() == template:
            print("[SKIP] reference document: %s" % source)
            continue
        output = Path(args.out).resolve() if args.file and args.out else mirrored_output_path(source, input_root, out_dir)
        result = RestyleResult(source=source, output=output, size_before=source.stat().st_size)
        try:
            result.lines = restyle_document(source, output, template, cfg, use_template, markup)
            result.size_after = output.stat().st_size
            print("[OK] %s -> %s" % (source, output))
        except Exception as exc:
            result.status = "FAILED"
            result.error = str(exc)
            result.lines = list(report)
            print("[FAILED] %s: %s" % (source, exc))
        results.append(result)

    write_report(report_path, input_root, template, config_path, markup, results)
    if args.json_out:
        json_path = Path(args.json_out).resolve()
        write_json_file(json_path, build_json(input_root, results))
        print("[OK] JSON: %s" % json_path)
    print("[OK] Report: %s" % report_path)
    return 0 if all(r.status == "OK" for r in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
