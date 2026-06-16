#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Krydstjek - validering af krydshenvisninger i kontrakter og bilag
==================================================================
1. Kortlægger alle definitioner (nummererede afsnit/punkter + bilag/appendiks)
2. Finder alle krydshenvisninger (jf. afsnit 4.2, se Bilag 3, iht. punkt 2.1 ...)
3. Validerer hver henvisning mod dokumentsamlingen

Resultat: Excel-rapport med ❌ UGYLDIGE, ⚠️ USIKRE og ✅ GYLDIGE henvisninger
samt en opsummering.

Kan køres alene:
    python krydstjek.py kontrakt.docx "Bilag 1 Tidsplan.docx" [...] -o rapport.xlsx
eller via knappen "Validér krydshenvisninger" i extractor_gui.py.

Kræver: extractor.py i samme mappe + openpyxl.
"""

import argparse
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import extractor as ex

THIN = Border(*[Side(style="thin")] * 4)
FILLS = {"ugyldig": PatternFill("solid", start_color="F8CBAD"),
         "usikker": PatternFill("solid", start_color="FFE699"),
         "gyldig": PatternFill("solid", start_color="C6E0B4"),
         "stoej": PatternFill("solid", start_color="E0E0E0")}
SYMBOL = {"ugyldig": "❌ UGYLDIG", "usikker": "⚠️ USIKKER",
          "gyldig": "✅ GYLDIG", "stoej": "➖ STØJ"}

# -- mønstre ----------------------------------------------------------------
SEKTIONSORD = r"(?:afsnit|punkt|pkt\.?|kapitel|sektion|klausul|§)"
BILAGSORD = r"(?:underbilag|bilag|appendiks|appendix)"
NUM = r"\d+(?:\.\d+)*"

# Bilagsnummer i LØBENDE TEKST: kræver mellemrum efter ordet ("se Bilag 3").
# Bogstav må kun klæbe DIREKTE til tallet (3A), ikke efter mellemrum (16 i = "16").
BILAG_NR = r"(\d+[A-Za-z]?|[A-ZÆØÅ])"
RE_BILAG_SEKTION = re.compile(
    rf"\b({BILAGSORD})\s+{BILAG_NR}\b[,\s]+\s*({SEKTIONSORD})\s*({NUM})",
    re.IGNORECASE)
RE_SEKTION = re.compile(rf"\b({SEKTIONSORD})\s*({NUM})(?:\s*[-–]\s*({NUM}))?", re.IGNORECASE)
RE_BILAG = re.compile(rf"\b({BILAGSORD})\s+{BILAG_NR}\b", re.IGNORECASE)

RE_MANUELT_NR = re.compile(r"^(\d+(?:\.\d+)+)[\.\)]?\s+\S")   # "4.2 Betaling"

# Enkeltbogstavs-"numre" der i virkeligheden er danske småord, ikke bilagsnumre.
# "bilag i det omfang", "bilag og kontrakt", "bilag e-mail" osv.
BILAG_STOPORD = {"I", "O"}
# Hele ord der aldrig er et bilagsnummer selvom de står efter "bilag"
EFTERORD_STOP = {"i", "og", "samt", "eller", "til", "for", "med", "som",
                 "der", "er", "kan", "skal", "jf", "nr", "ovenfor", "nedenfor",
                 "heri", "herunder", "m", "mv", "etc"}

# Bilagsnummer i FILNAVN / overskrift: tolerant - "Bilag" må stå midt i navnet,
# og _ . - eller mellemrum må skille ord og nummer. Fanger "xx. Bilag 01_...",
# "05. Bilag 03", "Udbudsbetingelser_Bilag A_...", "Bilag 03A_...", "Appendiks 2".
RE_BILAG_DEF = re.compile(
    rf"({BILAGSORD})[\s._-]*(\d+\s*[A-Za-z]?|[A-ZÆØÅ])(?![A-Za-z0-9])",
    re.IGNORECASE)


def er_bilagsnummer(raw):
    """Falsk hvis 'nummeret' i virkeligheden er et dansk småord (i, og, til...)
    eller et enkelt vildledende bogstav (I/O). Tal og A-H, J-Z er ok."""
    s = raw.strip().lower().rstrip(".")
    if s in EFTERORD_STOP:
        return False
    if len(s) == 1 and s.upper() in BILAG_STOPORD:
        return False
    return True


def norm_nr(s):
    """Normaliser så samme bilag/afsnit altid får samme nøgle.
    '03' -> '3', '4.20.' -> '4.20', '03A' -> '3A', 'a' -> 'A'.
    Tal og efterstillet bogstav holdes sammen ('3A' bliver IKKE '3.A')."""
    s = s.strip().rstrip(".")
    out = []
    for niveau in s.split("."):
        niveau = niveau.strip().replace(" ", "")
        m = re.match(r"^(\d+)([A-Za-z]?)$", niveau)
        if m:
            out.append(str(int(m.group(1))) + m.group(2).upper())
        elif niveau:
            out.append(niveau.upper())
    return ".".join(out)


# ---------------------------------------------------------------------------
# Trin 1: Kortlæg definitioner
# ---------------------------------------------------------------------------
def kortlaeg(docs):
    """Returnerer (sektioner_pr_dok, bilag_def).
    sektioner_pr_dok: {doknavn: set af normaliserede numre}
    bilag_def: {(type, nr): definitionssted}  fx ('BILAG','3'): 'Bilag 3 Tidsplan.docx'
    """
    sektioner = {}
    bilag = {}
    for d in docs:
        navn = d.path.name
        numre = set()
        for p_el, label, tekst, _ in d.paras:
            if label and label[:1].isdigit():
                numre.add(norm_nr(label))
            m = RE_MANUELT_NR.match(tekst)
            if m and len(tekst) <= 150:          # ligner en overskrift
                numre.add(norm_nr(m.group(1)))
            if tekst and len(tekst) <= 150:
                mb = RE_BILAG_DEF.search(tekst)
                if mb and mb.start() <= 4:   # bilagsordet står forrest i overskriften
                    bilag.setdefault((_btype(mb.group(1)), norm_nr(mb.group(2))),
                                     f"{navn} (overskrift)")
        sektioner[navn] = numre
        # bilagsidentitet fra filnavnet, fx "xx. Bilag 03_Tidsplan.docx".
        # Filnavnet vægter højest: overskriv evt. overskrifts-gæt.
        m = RE_BILAG_DEF.search(d.path.stem)
        if m:
            bilag[(_btype(m.group(1)), norm_nr(m.group(2)))] = navn
    return sektioner, bilag


def _btype(ord_):
    """Normaliser bilagstype: appendix/appendiks -> APPENDIKS, ellers ordet i caps."""
    o = ord_.upper()
    return "APPENDIKS" if o.startswith("APPEND") else o


def bilagstype_match(reftype, deftype):
    """'BILAG' matcher 'BILAG', 'APPENDIKS' matcher 'APPENDIKS'/'APPENDIX'."""
    return _btype(reftype) == _btype(deftype)


def find_bilag(bilag_def, reftype, nr):
    """Returnerer (status, sted).
    'gyldig'  ved PRÆCIST match (3 matcher kun 3, ikke 3A).
    'usikker' kun ved ægte skrivevariant af SAMME bilag (fx '3.0' vs '3'),
              ALDRIG mellem 3 og 3A - de er forskellige bilag.
    'ugyldig' ellers."""
    nr = norm_nr(nr)
    for (dtype, dnr), sted in bilag_def.items():
        if bilagstype_match(reftype, dtype) and dnr == nr:
            return "gyldig", sted
    # ægte nær-match: kun hvis cifrene er ens og forskellen er et trailing ".0"
    kerne = nr.rstrip(".0") or nr
    for (dtype, dnr), sted in bilag_def.items():
        if bilagstype_match(reftype, dtype) and (dnr.rstrip(".0") or dnr) == kerne and dnr != nr:
            return "usikker", f"nærmeste match: {dtype.title()} {dnr} ({sted}) - tjek skrivemåde"
    return "ugyldig", None


# ---------------------------------------------------------------------------
# Trin 2+3: Find og validér henvisninger
# ---------------------------------------------------------------------------
def saetning(tekst, start, slut):
    """Klip den omgivende sætning ud som kontekst."""
    a = max(tekst.rfind(". ", 0, start), tekst.rfind("\n", 0, start)) + 1
    b = tekst.find(". ", slut)
    b = len(tekst) if b == -1 else b + 1
    s = tekst[a:b].strip()
    return (s[:300] + " ...") if len(s) > 300 else s


def er_definition(tekst, m):
    """Spring selve definitionen over (overskriften 'Bilag 3 - Tidsplan'
    eller '4.2 Betaling' er ikke en henvisning)."""
    return m.start() == 0 and len(tekst) <= 150


def validér(docs):
    sektioner, bilag_def = kortlaeg(docs)
    fund = []

    def tjek_sektion(nr, doknavn):
        nr = norm_nr(nr)
        if nr in sektioner[doknavn]:
            return "gyldig", f"intern henvisning til afsnit {nr} - findes i samme dokument"
        andre = [n for n, s in sektioner.items() if nr in s and n != doknavn]
        if andre:
            return "usikker", (f"afsnit {nr} findes ikke i dette dokument, "
                               f"men i {', '.join(andre)} - hvis det er en intern henvisning "
                               f"er den forkert; ellers bør dokumentet angives")
        # findes hverken her eller andre steder = ugyldig INTERN henvisning
        forael = ".".join(nr.split(".")[:-1])
        if forael and forael in sektioner[doknavn]:
            return "ugyldig", (f"ugyldig intern henvisning: afsnit {nr} findes ikke - "
                               f"{forael} findes, men har intet underpunkt {nr.split('.')[-1]}")
        return "ugyldig", (f"ugyldig intern henvisning: afsnit {nr} findes ikke "
                           f"i dokumentet (og heller ikke i de øvrige medtagne dokumenter)")

    for d in docs:
        navn = d.path.name
        for p_el, _, tekst, _ in d.paras:
            if not tekst:
                continue
            optaget = []   # tegnpositioner der allerede er matchet

            def ledig(m):
                return not any(a < m.end() and m.start() < b for a, b in optaget)

            # 1) kombineret: "Bilag 3, punkt 2.1"
            for m in RE_BILAG_SEKTION.finditer(tekst):
                if er_definition(tekst, m) or not er_bilagsnummer(m.group(2)):
                    continue
                optaget.append((m.start(), m.end()))
                btype, bnr = m.group(1), norm_nr(m.group(2))
                snr = norm_nr(m.group(4))
                bstatus, sted = find_bilag(bilag_def, btype, bnr)
                if bstatus == "ugyldig":
                    status, forkl = "ugyldig", f"{btype.title()} {bnr} findes ikke i samlingen"
                elif bstatus == "usikker":
                    status, forkl = "usikker", sted
                else:
                    if sted.endswith("(overskrift)") :
                        status, forkl = "stoej", (f"{btype.title()} {bnr} kun nævnt som "
                                                  f"overskrift i {sted.split(' (')[0]} - "
                                                  f"ikke eget dokument i samlingen")
                    elif snr in sektioner.get(sted, set()):
                        status, forkl = "gyldig", f"afsnit {snr} findes i {sted}"
                    else:
                        status, forkl = "ugyldig", f"afsnit {snr} findes ikke i {sted}"
                fund.append((status, navn, m.group(0), saetning(tekst, m.start(), m.end()), forkl))

            # 2) bilagshenvisninger: "se Bilag 3"
            for m in RE_BILAG.finditer(tekst):
                if not ledig(m) or er_definition(tekst, m) or not er_bilagsnummer(m.group(2)):
                    continue
                optaget.append((m.start(), m.end()))
                btype, bnr = m.group(1), norm_nr(m.group(2))
                status, sted = find_bilag(bilag_def, btype, bnr)
                if status == "gyldig":
                    forkl = f"defineret: {sted}"
                    if sted.endswith("(overskrift)"):
                        status = "stoej"
                        forkl = (f"kun nævnt som overskrift i {sted.split(' (')[0]} - "
                                 f"ikke eget dokument i samlingen")
                elif status == "usikker":
                    forkl = sted
                else:
                    forkl = f"{btype.title()} {bnr} findes ikke i dokumentsamlingen"
                fund.append((status, navn, m.group(0), saetning(tekst, m.start(), m.end()), forkl))

            # 3) sektionshenvisninger: "jf. afsnit 4.2" / "punkt 2.1-2.3"
            for m in RE_SEKTION.finditer(tekst):
                if not ledig(m) or er_definition(tekst, m):
                    continue
                optaget.append((m.start(), m.end()))
                status1, forkl = tjek_sektion(m.group(2), navn)
                if m.group(3):                       # interval, fx 2.1-2.3
                    status2, forkl2 = tjek_sektion(m.group(3), navn)
                    par = sorted([status1, status2], key=["gyldig", "usikker", "ugyldig"].index)
                    if par[0] != par[1]:
                        status1, forkl = "usikker", f"interval delvist gyldigt: {forkl} / {forkl2}"
                    else:
                        status1 = par[0]
                        forkl = f"{forkl}; {forkl2}"
                fund.append((status1, navn, m.group(0), saetning(tekst, m.start(), m.end()), forkl))

    return fund, sektioner, bilag_def


# ---------------------------------------------------------------------------
# Rapport
# ---------------------------------------------------------------------------
def _skriv_ark(wb, titel, raekker, intro=None):
    ws = wb.create_sheet(titel)
    rr = 1
    if intro:
        ws.cell(row=1, column=1, value=intro).font = Font(italic=True, color="666666")
        rr = 2
    hoved = ["Status", "Dokument", "Henvisning", "Sætning (kontekst)", "Vurdering"]
    bredder = [16, 35, 22, 75, 55]
    for c, (h, b) in enumerate(zip(hoved, bredder), 1):
        cell = ws.cell(row=rr, column=c, value=h)
        cell.font = Font(bold=True)
        ws.column_dimensions[cell.column_letter].width = b
    orden = {"ugyldig": 0, "usikker": 1, "stoej": 2, "gyldig": 3}
    for status, dok, ref, ktx, forkl in sorted(raekker, key=lambda f: orden[f[0]]):
        rr += 1
        ws.cell(row=rr, column=1, value=SYMBOL[status]).fill = FILLS[status]
        for c, v in enumerate((dok, ref, ktx, forkl), 2):
            ws.cell(row=rr, column=c, value=v)
        for c in range(1, 6):
            ws.cell(row=rr, column=c).border = THIN
            ws.cell(row=rr, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    if rr == (2 if intro else 1):
        ws.cell(row=rr + 1, column=2, value="(ingen)")


def skriv_rapport(fund, sektioner, bilag_def, outfile):
    wb = Workbook()
    wb.remove(wb.active)

    n = {s: sum(1 for f in fund if f[0] == s)
         for s in ("gyldig", "ugyldig", "usikker", "stoej")}

    # Fane 1: Opsummering
    ops = wb.create_sheet("Opsummering")
    ops["B1"] = "Opsummering af krydshenvisningstjek"
    ops["B1"].font = Font(bold=True, size=12)
    linjer = [("ugyldig", "Ugyldige henvisninger (reelle fejl)"),
              ("usikker", "Usikre henvisninger (bør tjekkes)"),
              ("gyldig", "Gyldige henvisninger"),
              ("stoej", "Støj (interne appendiks-overskrifter m.m.)")]
    for i, (s, tekst) in enumerate(linjer, 3):
        ops.cell(row=i, column=2, value=SYMBOL[s]).fill = FILLS[s]
        ops.cell(row=i, column=3, value=tekst)
        ops.cell(row=i, column=4, value=n[s]).font = Font(bold=True)
    ops["B8"] = f"Henvisninger i alt: {len(fund)}"
    ops["B9"] = "Til gennemgang (ugyldige + usikre): " + str(n["ugyldig"] + n["usikker"])
    ops["B9"].font = Font(bold=True)
    ops["B11"] = "Definerede bilag/appendikser i samlingen:"
    ops["B11"].font = Font(bold=True)
    r = 11
    for (btype, bnr), sted in sorted(bilag_def.items()):
        r += 1
        ops.cell(row=r, column=2, value=f"{btype.title()} {bnr}")
        ops.cell(row=r, column=3, value=sted)
    r += 2
    ops.cell(row=r, column=2, value="Dokumenter og antal fundne afsnitsnumre:").font = Font(bold=True)
    for navn, numre in sektioner.items():
        r += 1
        ops.cell(row=r, column=2, value=navn)
        ops.cell(row=r, column=3, value=len(numre))
    for col, b in zip("BCD", (30, 60, 10)):
        ops.column_dimensions[col].width = b

    # Fane 2: Til gennemgang (kun det vigtige) - ugyldige + usikre
    _skriv_ark(wb, "Til gennemgang",
               [f for f in fund if f[0] in ("ugyldig", "usikker")],
               intro="De henvisninger der kræver menneskelig vurdering - ugyldige øverst.")
    # Fane 3: Gyldige
    _skriv_ark(wb, "Gyldige", [f for f in fund if f[0] == "gyldig"])
    # Fane 4: Støj
    _skriv_ark(wb, "Støj (kan ignoreres)",
               [f for f in fund if f[0] == "stoej"],
               intro="Henvisninger til appendikser/bilag der kun findes som overskrifter "
                     "inde i andre dokumenter - normalt ikke fejl.")

    wb.save(outfile)
    return n


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Validér krydshenvisninger i kontrakt + bilag")
    ap.add_argument("filer", nargs="+", help="Kontrakt og alle bilag (.docx)")
    ap.add_argument("-o", "--output", default="krydstjek.xlsx")
    args = ap.parse_args()

    docs = [ex.Docx(f) for f in args.filer]
    fund, sektioner, bilag_def = validér(docs)
    n = skriv_rapport(fund, sektioner, bilag_def, args.output)
    print(f"Krydstjek gennemfort -> {args.output}")
    print(f"  Ugyldige: {n['ugyldig']}   Usikre: {n['usikker']}   "
          f"Gyldige: {n['gyldig']}   Stoj: {n['stoej']}")
    print(f"  -> Til gennemgang: {n['ugyldig'] + n['usikker']}")


if __name__ == "__main__":
    sys.exit(main())
