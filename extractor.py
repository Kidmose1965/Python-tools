#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor - Python-udgave af "Revised Extractor"
=================================================
Udtrækker kommentarer, trackchanges, markerede inputfelter og kravmatrix
direkte fra .docx-filer - uden at Word åbnes.

Kræver kun: openpyxl   (pip install openpyxl)

Brug:
  python extractor.py kommentarer  FIL.docx [FLERE.docx ...] [-o output.xlsx]
  python extractor.py trackchanges FIL.docx [FLERE.docx ...] [-o output.xlsx]
  python extractor.py inputfelter --farve groen FIL.docx [...] [-o output.xlsx]
  python extractor.py inputfelter --farve gul   FIL.docx [...] [-o output.xlsx]
  python extractor.py kravmatrix   KRAVSPEC.docx            [-o output.xlsx]
  python extractor.py kravmatrix-kommentarer KRAVSPEC.docx  [-o output.xlsx]
  python extractor.py styles       FIL.docx        (viser dokumentets typografier)
"""

import argparse
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
BAND = PatternFill("solid", start_color="DCE6F1")
THIN = Border(*[Side(style="thin")] * 4)


def q(tag):
    return W + tag


# ---------------------------------------------------------------------------
# Nummereringsmotor: genskaber Words "ListString" ud fra numbering.xml
# ---------------------------------------------------------------------------
def _fmt_number(n, numfmt):
    if numfmt == "decimal":
        return str(n)
    if numfmt in ("lowerLetter", "upperLetter"):
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(97 + r) + s
        return s.upper() if numfmt == "upperLetter" else s
    if numfmt in ("lowerRoman", "upperRoman"):
        vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
                (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
                (5, "v"), (4, "iv"), (1, "i")]
        s = ""
        for v, sym in vals:
            while n >= v:
                s += sym
                n -= v
        return s.upper() if numfmt == "upperRoman" else s
    if numfmt == "none":
        return ""
    return None  # bullets m.v. -> ingen numerisk nummerering


class NumberingEngine:
    def __init__(self, numbering_root):
        self.levels = {}   # numId -> {ilvl: (start, numfmt, lvlText)}
        self.counters = {}  # numId -> {ilvl: aktuel værdi}
        if numbering_root is None:
            return
        abstract = {}
        for an in numbering_root.findall(q("abstractNum")):
            aid = an.get(q("abstractNumId"))
            lvls = {}
            for lvl in an.findall(q("lvl")):
                ilvl = int(lvl.get(q("ilvl")))
                start_el = lvl.find(q("start"))
                fmt_el = lvl.find(q("numFmt"))
                txt_el = lvl.find(q("lvlText"))
                lvls[ilvl] = (
                    int(start_el.get(q("val"))) if start_el is not None else 1,
                    fmt_el.get(q("val")) if fmt_el is not None else "decimal",
                    txt_el.get(q("val")) if txt_el is not None else "%1",
                )
            abstract[aid] = lvls
        for num in numbering_root.findall(q("num")):
            nid = num.get(q("numId"))
            ref = num.find(q("abstractNumId"))
            if ref is not None and ref.get(q("val")) in abstract:
                self.levels[nid] = abstract[ref.get(q("val"))]

    def next_label(self, num_id, ilvl):
        """Tæl op og returnér label, fx '3.2.1' - eller None for punkttegn."""
        lvls = self.levels.get(num_id)
        if not lvls or ilvl not in lvls:
            return None
        c = self.counters.setdefault(num_id, {})
        for lv in range(ilvl):                       # init overliggende niveauer
            if lv in lvls and lv not in c:
                c[lv] = lvls[lv][0]
        c[ilvl] = c.get(ilvl, lvls[ilvl][0] - 1) + 1  # tæl dette niveau op
        for lv in list(c):                            # nulstil dybere niveauer
            if lv > ilvl:
                del c[lv]
        text = lvls[ilvl][2]
        for lv in range(ilvl + 1):
            if lv in lvls:
                part = _fmt_number(c.get(lv, lvls[lv][0]), lvls[lv][1])
                if part is None:
                    return None
                text = text.replace(f"%{lv + 1}", part)
        return text.strip()


# ---------------------------------------------------------------------------
# Docx-læser
# ---------------------------------------------------------------------------
class Docx:
    def __init__(self, path):
        # .resolve() retter Windows' korte 8.3-filnavne (fx "03UDBU~1.DOC" fra
        # tkinters filvalgsdialog) tilbage til det fulde filnavn - uden det
        # fejler bilag-genkendelsen i krydstjek, som matcher på filnavnet.
        self.path = Path(path).resolve()
        with zipfile.ZipFile(path) as z:
            def load(name):
                try:
                    return ET.fromstring(z.read(name))
                except KeyError:
                    return None

            self.doc = load("word/document.xml")
            self.numbering = NumberingEngine(load("word/numbering.xml"))
            self.comments_xml = load("word/comments.xml")

            # styleId -> (visningsnavn, numPr-fra-style, basedOn)
            self.styles = {}
            styles_root = load("word/styles.xml")
            if styles_root is not None:
                for st in styles_root.findall(q("style")):
                    sid = st.get(q("styleId"))
                    name_el = st.find(q("name"))
                    numpr = st.find(f"{q('pPr')}/{q('numPr')}")
                    based = st.find(q("basedOn"))
                    self.styles[sid] = (
                        name_el.get(q("val")) if name_el is not None else sid,
                        numpr,
                        based.get(q("val")) if based is not None else None,
                    )

        self._scan()

    # -- styles ------------------------------------------------------------
    def style_name(self, style_id):
        return self.styles.get(style_id, (style_id or "Normal",))[0]

    def _style_numpr(self, style_id, depth=0):
        if style_id not in self.styles or depth > 10:
            return None
        _, numpr, based = self.styles[style_id]
        return numpr if numpr is not None else self._style_numpr(based, depth + 1)

    # -- gennemløb af dokumentet i læserækkefølge ---------------------------
    def _scan(self):
        """Én tur gennem brødteksten: nummerér alle afsnit og notér,
        om de står i en tabel."""
        self.paras = []          # [(element, label|None, tekst, in_table)]
        self.index_of = {}       # id(element) -> indeks i self.paras
        body = self.doc.find(q("body"))
        for el, in_tbl in self._iter_block(body, False):
            label = self._para_label(el)
            text = para_text(el)
            self.index_of[id(el)] = len(self.paras)
            self.paras.append((el, label, text, in_tbl))

    def _iter_block(self, parent, in_table):
        for child in parent:
            if child.tag == q("p"):
                yield child, in_table
            elif child.tag == q("tbl"):
                for tr in child.findall(q("tr")):
                    for tc in tr.findall(q("tc")):
                        yield from self._iter_block(tc, True)

    def _para_label(self, p):
        ppr = p.find(q("pPr"))
        numpr = ppr.find(q("numPr")) if ppr is not None else None
        if numpr is None and ppr is not None:
            pstyle = ppr.find(q("pStyle"))
            if pstyle is not None:
                numpr = self._style_numpr(pstyle.get(q("val")))
        if numpr is None:
            return None
        nid_el = numpr.find(q("numId"))
        ilvl_el = numpr.find(q("ilvl"))
        if nid_el is None:
            return None
        nid = nid_el.get(q("val"))
        ilvl = int(ilvl_el.get(q("val"))) if ilvl_el is not None else 0
        if nid == "0":
            return None
        return self.numbering.next_label(nid, ilvl)

    # -- "nærmeste nummererede overskrift" (som i VBA-værktøjet) ------------
    def section_for(self, p_element):
        i = self.index_of.get(id(p_element))
        if i is None:
            return "Ukendt placering"
        while i >= 0:
            _, label, text, _ = self.paras[i]
            if label and label[:1].isdigit():
                return f"{label} - {text}".strip(" -")
            i -= 1
        return "Før nummereret sektion"

    def para_style_name(self, p):
        ppr = p.find(q("pPr"))
        if ppr is not None:
            ps = ppr.find(q("pStyle"))
            if ps is not None:
                return self.style_name(ps.get(q("val")))
        return self.style_name(None)


def para_text(p, include_del=False):
    parts = []
    for node in p.iter():
        if node.tag == q("t") and node.text:
            parts.append(node.text)
        elif include_del and node.tag == q("delText") and node.text:
            parts.append(node.text)
        elif node.tag == q("tab"):
            parts.append("\t")
    return "".join(parts).strip()


# ---------------------------------------------------------------------------
# Udtræk 1: Kommentarer
# ---------------------------------------------------------------------------
def extract_comments(docx):
    out = []
    if docx.comments_xml is None:
        return out
    meta = {}
    for c in docx.comments_xml.findall(q("comment")):
        cid = c.get(q("id"))
        text = "\n".join(filter(None, (para_text(p) for p in c.findall(q("p")))))
        meta[cid] = (c.get(q("initials")) or c.get(q("author")) or "", text)

    # find scope-tekst og ankerafsnit for hver kommentar
    scope_text, anchor, active = {}, {}, set()
    for p_el, _, _, _ in docx.paras:
        for node in p_el.iter():
            if node.tag == q("commentRangeStart"):
                cid = node.get(q("id"))
                active.add(cid)
                scope_text.setdefault(cid, [])
                anchor.setdefault(cid, p_el)
            elif node.tag == q("commentRangeEnd"):
                active.discard(node.get(q("id")))
            elif node.tag == q("t") and node.text and active:
                for cid in active:
                    scope_text[cid].append(node.text)
            elif node.tag == q("commentReference"):
                anchor.setdefault(node.get(q("id")), p_el)

    for n, (cid, (initials, text)) in enumerate(sorted(meta.items(), key=lambda kv: int(kv[0])), 1):
        sec = docx.section_for(anchor[cid]) if cid in anchor else "Ukendt placering"
        out.append({
            "Dokument": docx.path.name, "Nummer": n, "Kommentar": text,
            "Markeret tekst": "".join(scope_text.get(cid, [])).strip(),
            "Initialer": initials, "Nummereret sektion": sec,
        })
    return out


# ---------------------------------------------------------------------------
# Udtræk 2: Trackchanges
# ---------------------------------------------------------------------------
def extract_trackchanges(docx):
    out = []
    for p_el, _, _, _ in docx.paras:
        for node in p_el:
            if node.tag == q("ins"):
                txt = para_text(node)
                if txt:
                    out.append({"Dokument": docx.path.name,
                                "Sektion": docx.section_for(p_el),
                                "Ændring": txt, "Type": "Indsat",
                                "Forfatter": node.get(q("author")) or ""})
            elif node.tag == q("del"):
                txt = para_text(node, include_del=True)
                if txt:
                    out.append({"Dokument": docx.path.name,
                                "Sektion": docx.section_for(p_el),
                                "Ændring": txt, "Type": "Slettet",
                                "Forfatter": node.get(q("author")) or ""})
    return out


# ---------------------------------------------------------------------------
# Udtræk 3: Markerede inputfelter (highlight)
# ---------------------------------------------------------------------------
HIGHLIGHT = {"groen": "green", "grøn": "green", "gul": "yellow"}


def extract_highlights(docx, colour):
    out = []
    want = HIGHLIGHT[colour]
    for p_el, _, _, _ in docx.paras:
        current = []
        for r in p_el.iter(q("r")):
            rpr = r.find(q("rPr"))
            hl = rpr.find(q("highlight")) if rpr is not None else None
            if hl is not None and hl.get(q("val")) == want:
                current.append(para_text(r))
            elif current:
                out.append({"Dokument": docx.path.name,
                            "Sektion": docx.section_for(p_el),
                            "Inputfelt": "".join(current)})
                current = []
        if current:
            out.append({"Dokument": docx.path.name,
                        "Sektion": docx.section_for(p_el),
                        "Inputfelt": "".join(current)})
    return out


# ---------------------------------------------------------------------------
# Udtræk 4: Kravmatrix
# ---------------------------------------------------------------------------
def _comments_by_anchor_paragraph(docx):
    """id(afsnit) -> liste af kommentartekster der starter i det afsnit."""
    if docx.comments_xml is None:
        return {}
    meta = {}
    for c in docx.comments_xml.findall(q("comment")):
        cid = c.get(q("id"))
        text = "\n".join(filter(None, (para_text(p) for p in c.findall(q("p")))))
        meta[cid] = text

    anchor = {}
    for p_el, _, _, _ in docx.paras:
        for node in p_el.iter():
            if node.tag in (q("commentRangeStart"), q("commentReference")):
                anchor.setdefault(node.get(q("id")), p_el)

    by_para = {}
    for cid, p_el in anchor.items():
        if cid in meta:
            by_para.setdefault(id(p_el), []).append(meta[cid])
    return by_para


def list_styles(docx):
    seen, ordered = set(), []
    for p_el, _, _, _ in docx.paras:
        nm = docx.para_style_name(p_el)
        if nm not in seen:
            seen.add(nm)
            ordered.append(nm)
    return ordered


def _pick(prompt, styles, allow_all):
    print(prompt)
    extra = "  0: [ALLE]\n" if allow_all else "  0: [INGEN/færdig]\n"
    print(extra + "\n".join(f"  {i}: {s}" for i, s in enumerate(styles, 1)))
    raw = input("Angiv numre adskilt af komma (fx 2,5): ").strip()
    chosen = []
    for tok in raw.split(","):
        tok = tok.strip()
        if tok == "0" and allow_all:
            return "ALLE"
        if tok.isdigit() and 1 <= int(tok) <= len(styles):
            chosen.append(styles[int(tok) - 1])
    return chosen


def extract_kravmatrix(docx, table_styles=None, heading_styles=None, with_comments=False):
    if table_styles is None:
        styles = list_styles(docx)
        table_styles = _pick("\nStyles tilhørende TABEL-elementer (kravene):", styles, True)
        heading_styles = _pick("\nStyles tilhørende OVERSKRIFTER (separatorer):", styles, False) or []
    get_all = table_styles == "ALLE"
    comments_by_para = _comments_by_anchor_paragraph(docx) if with_comments else {}

    rows = []
    body = docx.doc.find(q("body"))
    for child in body:
        if child.tag == q("tbl"):
            for tr in child.findall(q("tr")):
                tcs = tr.findall(q("tc"))
                if not tcs:
                    continue
                p1 = tcs[0].find(q("p"))
                style = docx.para_style_name(p1) if p1 is not None else ""
                if not (get_all or style in table_styles):
                    continue
                c1 = cell_text(docx, tcs[0])
                c2 = cell_text(docx, tcs[1]) if len(tcs) > 1 else ""
                if not (c1 or c2):
                    continue
                if with_comments:
                    kommentarer = [txt for tc in tcs for p in tc.findall(q("p"))
                                   for txt in comments_by_para.get(id(p), [])]
                    rows.append(("krav", c1, c2, "\n".join(kommentarer)))
                else:
                    rows.append(("krav", c1, c2))
        elif child.tag == q("p"):
            style = docx.para_style_name(child)
            if style in (heading_styles or []):
                i = docx.index_of.get(id(child))
                label = docx.paras[i][1] if i is not None else None
                text = para_text(child)
                if text:
                    rows.append(("overskrift", label or "", text, "") if with_comments
                               else ("overskrift", label or "", text))
    return rows


def cell_text(docx, tc):
    parts = []
    for p in tc.findall(q("p")):
        i = docx.index_of.get(id(p))
        label = docx.paras[i][1] if i is not None else None
        txt = para_text(p)
        parts.append(f"{label} {txt}".strip() if label else txt)
    return "\n".join(filter(None, parts)).strip()


# ---------------------------------------------------------------------------
# Excel-output
# ---------------------------------------------------------------------------
def write_simple(records, headers, outfile, widths):
    wb = Workbook()
    ws = wb.active
    for col, (h, w) in enumerate(zip(headers, widths), 2):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
        ws.column_dimensions[c.column_letter].width = w
    for r, rec in enumerate(records, 2):
        for col, h in enumerate(headers, 2):
            c = ws.cell(row=r, column=col, value=rec.get(h, ""))
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=1, max_row=max(1, len(records) + 1),
                            min_col=2, max_col=len(headers) + 1):
        for c in row:
            c.border = THIN
    wb.save(outfile)


def write_kravmatrix(rows, outfile, tilbudsgiver="[…]"):
    wb = Workbook()
    ws = wb.active
    ws["B1"] = "Bilag XX, Appendiks XX - Kravmatrix"
    ws["B1"].font = Font(bold=True, size=12)
    ws["B3"] = f"Navn på tilbudsgiver: {tilbudsgiver}"
    headers = [("B5", "Krav nr."), ("C5", "Kravbeskrivelse"), ("D5", "Krav opfyldt"),
               ("G5", "Standard-programmel"), ("H5", "Tilpasning/opsætning"),
               ("I5", "Redegørelse"), ("D6", "J"), ("E6", "N"), ("F6", "D")]
    for ref, val in headers:
        ws[ref] = val
        ws[ref].font = Font(bold=True)
    ws.merge_cells("D5:F5")
    for col, w in zip("BCDEFGHI", (21, 85, 6, 6, 6, 20, 22, 22)):
        ws.column_dimensions[col].width = w

    r = 6
    for kind, a, b in rows:
        r += 1
        ws.cell(row=r, column=2, value=a)
        ws.cell(row=r, column=3, value=b)
        if kind == "overskrift":
            ws.cell(row=r, column=2).font = Font(bold=True)
            ws.cell(row=r, column=3).font = Font(bold=True)
    if r > 6:
        for i in range(7, r + 1):
            for col in range(2, 10):
                c = ws.cell(row=i, column=col)
                c.border = THIN
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if i % 2 == 0:
                    c.fill = BAND
    wb.save(outfile)


def write_kravmatrix_kommentarer(rows, outfile):
    wb = Workbook()
    ws = wb.active
    headers = [("B1", "Krav nr."), ("C1", "Kravbeskrivelse"), ("D1", "Kommentarer")]
    for ref, val in headers:
        ws[ref] = val
        ws[ref].font = Font(bold=True)
    for col, w in zip("BCD", (21, 85, 60)):
        ws.column_dimensions[col].width = w

    r = 1
    for kind, a, b, kommentar in rows:
        r += 1
        ws.cell(row=r, column=2, value=a)
        ws.cell(row=r, column=3, value=b)
        ws.cell(row=r, column=4, value=kommentar)
        if kind == "overskrift":
            ws.cell(row=r, column=2).font = Font(bold=True)
            ws.cell(row=r, column=3).font = Font(bold=True)
    if r > 1:
        for i in range(2, r + 1):
            for col in range(2, 5):
                c = ws.cell(row=i, column=col)
                c.border = THIN
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if i % 2 == 0:
                    c.fill = BAND
    wb.save(outfile)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Udtræk fra Word-kravspecifikationer (.docx)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for navn in ("kommentarer", "trackchanges", "kravmatrix", "kravmatrix-kommentarer", "styles"):
        p = sub.add_parser(navn)
        p.add_argument("filer", nargs="+")
        p.add_argument("-o", "--output")
    p = sub.add_parser("inputfelter")
    p.add_argument("filer", nargs="+")
    p.add_argument("--farve", choices=["groen", "grøn", "gul"], required=True)
    p.add_argument("-o", "--output")
    args = ap.parse_args()

    docs = [Docx(f) for f in args.filer]

    if args.cmd == "styles":
        for d in docs:
            print(f"\n{d.path.name}:")
            for s in list_styles(d):
                print(f"  - {s}")
        return

    if args.cmd == "kommentarer":
        recs = [r for d in docs for r in (extract_comments(d) or
                [{"Dokument": d.path.name, "Kommentar": "Ingen kommentarer"}])]
        out = args.output or "kommentarer.xlsx"
        write_simple(recs, ["Dokument", "Nummer", "Kommentar", "Markeret tekst",
                            "Initialer", "Nummereret sektion"], out,
                     [35, 9, 50, 50, 12, 50])

    elif args.cmd == "trackchanges":
        recs = [r for d in docs for r in (extract_trackchanges(d) or
                [{"Dokument": d.path.name, "Ændring": "Ingen trackchanges"}])]
        out = args.output or "trackchanges.xlsx"
        write_simple(recs, ["Dokument", "Sektion", "Ændring", "Type", "Forfatter"],
                     out, [35, 50, 60, 10, 20])

    elif args.cmd == "inputfelter":
        recs = [r for d in docs for r in (extract_highlights(d, args.farve) or
                [{"Dokument": d.path.name, "Inputfelt": "Ingen markerede felter"}])]
        out = args.output or f"inputfelter_{args.farve}.xlsx"
        write_simple(recs, ["Dokument", "Sektion", "Inputfelt"], out, [35, 50, 60])

    elif args.cmd == "kravmatrix":
        rows = []
        for d in docs:
            rows += extract_kravmatrix(d)
        out = args.output or "kravmatrix.xlsx"
        write_kravmatrix(rows, out)

    elif args.cmd == "kravmatrix-kommentarer":
        rows = []
        for d in docs:
            rows += extract_kravmatrix(d, with_comments=True)
        out = args.output or "kravmatrix_kommentarer.xlsx"
        write_kravmatrix_kommentarer(rows, out)

    er_kravmatrix = args.cmd in ("kravmatrix", "kravmatrix-kommentarer")
    print(f"Udtræk gennemført -> {out}  ({len(rows) if er_kravmatrix else len(recs)} rækker)")


if __name__ == "__main__":
    sys.exit(main())
