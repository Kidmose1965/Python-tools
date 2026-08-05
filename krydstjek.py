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

De dokumenter du angiver direkte bliver TJEKKET (deres henvisninger valideres
og rapporteres). Med --kontekst kan du derudover angive filer/mapper der kun
bruges til OPSLAG - fx selve udbudsbetingelserne - så en henvisning som
"punkt 14.1 i udbudsbetingelserne" ikke fejlagtigt markeres ugyldig, bare
fordi udbudsbetingelserne ikke selv er en af de dokumenter du vil have tjekket:

    python krydstjek.py kontrakt.docx tidsplan.docx --kontekst "01. Udbudsbetingelser.docx" -o rapport.xlsx

--kontekst kan også pege på en mappe - så bruges alle .docx-filer direkte i
den mappe (ikke undermapper som fx ARKIV) som baggrundskontekst:

    python krydstjek.py kontrakt.docx tidsplan.docx --kontekst "C:/.../Udbudsmateriale" -o rapport.xlsx

Kræver: extractor.py i samme mappe + openpyxl.
"""

import argparse
import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

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
SEKTIONSORD = r"(?:afsnit|punkt|pkt\.?|kapitel|sektion|klausul|§|underpunkt|del)"
BILAGSORD = r"(?:underbilag|kontraktbilag|bilag|appendiks|appendix)"
NUM = r"\d+(?:\.\d+)*"

# Bilagsnummer i LØBENDE TEKST: kræver mellemrum efter ordet ("se Bilag 3").
# Bogstav må kun klæbe DIREKTE til tallet (3A), ikke efter mellemrum (16 i = "16").
BILAG_NR = r"(\d+[A-Za-z]?|[A-ZÆØÅ])"
RE_BILAG_SEKTION = re.compile(
    rf"\b({BILAGSORD})\s+{BILAG_NR}\b[,\s]+\s*({SEKTIONSORD})\s*({NUM})",
    re.IGNORECASE)
RE_SEKTION = re.compile(
    rf"(?:"
    rf"(kontraktens|nærværende\s+(?:\w+\s+)?)\s*({SEKTIONSORD})\s*({NUM})(?:\s*[-–]\s*({NUM}))?"
    rf"|"
    # (?<!\w) i stedet for \b: \b matcher ikke foran et symbol som "§" (begge
    # sider ikke-ordtegn), så "§ 4.2" faldt hidtil helt igennem nettet.
    rf"(?<!\w)({SEKTIONSORD})\s*({NUM})(?:\s*[-–]\s*({NUM}))?"
    rf")",
    re.IGNORECASE)
RE_BILAG = re.compile(rf"\b({BILAGSORD})\s+{BILAG_NR}\b", re.IGNORECASE)
RE_BILAG_OG = re.compile(
    rf"\b({BILAGSORD})\s+(\d+[A-Za-z]?)\s+og\s+(\d+[A-Za-z]?)\b",
    re.IGNORECASE)

RE_MANUELT_NR = re.compile(r"^(\d+(?:\.\d+)+)[\.\)]?\s+\S")   # "4.2 Betaling"

# Pladsholder-henvisninger: skabelontekst hvor nummeret ikke er udfyldt endnu,
# fx "jf. punkt x.x", "se Kapitel ?", "Bilag ??" eller "afsnit [nummer]". De
# matcher ALDRIG mønstrene ovenfor (som kræver et rigtigt tal/bogstav), og blev
# derfor tidligere slet ikke opdaget - i stedet for at blive markeret ugyldig
# forsvandt de bare. En henvisning uden udfyldt nummer er pr. definition en
# fejl, så den flages altid som ugyldig, uanset om et "rigtigt" nummer med
# samme værdi findes et sted i samlingen.
# - x/X (evt. gentaget og/eller med punktum: x, xx, x.x, X.X.X ...) er kun
#   entydigt en pladsholder for afsnit/punkt/kapitel/§ osv., hvor et rigtigt
#   nummer ALDRIG er et bogstav. For bilag kan et enkelt bogstav (fx "Bilag A")
#   være en ægte betegnelse, så der bruges her kun symbol-pladsholdere.
PLADSHOLDER_ORD = r"(?:N/A|TBD|TBC|\[[^\]\n]{1,20}\]|_{2,}|\.{3,}|…)"
PLADSHOLDER_SEKTION_NR = rf"(?:[xX?]+(?:[.\s][xX?]+)*|{PLADSHOLDER_ORD})"
PLADSHOLDER_BILAG_NR = rf"(?:\?+(?:[.\s]\?+)*|{PLADSHOLDER_ORD})"
RE_PLADSHOLDER_SEKTION = re.compile(
    rf"(?<!\w)({SEKTIONSORD})\s*({PLADSHOLDER_SEKTION_NR})(?!\w)",
    re.IGNORECASE)
RE_PLADSHOLDER_BILAG = re.compile(
    rf"\b({BILAGSORD})\s+({PLADSHOLDER_BILAG_NR})(?!\w)",
    re.IGNORECASE)

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
def validér_samlinger(samlinger, kontekst_docs=None):
    """
    samlinger: liste af (navn, [Docx])
    kontekst_docs: valgfri liste af Docx der KUN bruges til opslag af afsnits-
        og bilagsnumre (fx udbudsbetingelserne) - de bliver ikke selv scannet
        for henvisninger og indgår ikke i rapportens fund.
    Returnerer (fund, alle_sektioner, alle_bilag)
    fund-tupler: (samling, status, doknavn, ref, kontekst, forklaring)
    """
    EKSTERN_ORD = re.compile(
        r"\b(forordning|direktiv|loven?|bekendtgørelse|EU|GDPR|"
        r"databeskyttelsesloven|udbudslov|persondatalov)\b",
        re.IGNORECASE)

    # Byg indeks per samling
    indeks = {}
    for navn, docs in samlinger:
        sektioner, bilag_def = kortlaeg(docs)
        indeks[navn] = (sektioner, bilag_def)

    # Kontekstdokumenters afsnits-/bilagsnumre lægges oveni HVER samlings eget
    # indeks (kun til opslag under validering) - men holdes ude af "indeks",
    # som cross-samling-logikken herunder bruger uændret.
    kctx_sektioner, kctx_bilag = kortlaeg(kontekst_docs) if kontekst_docs else ({}, {})

    alle_fund = []

    for samling_navn, docs in samlinger:
        sektioner, bilag_def = indeks[samling_navn]

        sektioner_m_kontekst = {**sektioner, **kctx_sektioner}
        bilag_m_kontekst = dict(bilag_def)
        for k, v in kctx_bilag.items():
            bilag_m_kontekst.setdefault(k, v)

        # Byg samlet indeks over alle ANDRE samlinger
        andre_bilag = {}
        andre_sektioner = {}
        for andet_navn, (a_sek, a_bil) in indeks.items():
            if andet_navn != samling_navn:
                for k, v in a_bil.items():
                    andre_bilag[k] = (andet_navn, v)
                for dok, numre in a_sek.items():
                    for nr in numre:
                        andre_sektioner.setdefault(nr, []).append(andet_navn)

        fund_lokal, _, _ = validér(
            docs, _bilag_override=bilag_m_kontekst, _sektioner_override=sektioner_m_kontekst)

        for status, dok, ref, ctx, forkl in fund_lokal:

            # Ekstern lovhenvisning -> støj
            if EKSTERN_ORD.search(ref) or EKSTERN_ORD.search(ctx[:80]):
                alle_fund.append((samling_navn, "stoej", dok, ref, ctx,
                                  "Ekstern lovhenvisning – ikke valideret"))
                continue

            # Ugyldig internt -> tjek om den findes i en anden samling
            if status == "ugyldig":
                bm = re.search(
                    r"\b(?:bilag|appendiks|kontraktbilag)\s+(\w+)", ref, re.IGNORECASE)
                pm = re.search(
                    r"\b(?:punkt|afsnit|underpunkt)\s+([\d.]+)", ref, re.IGNORECASE)

                if bm:
                    nr = norm_nr(bm.group(1))
                    btype = re.match(r"\w+", ref).group(0)
                    for (dtype, dnr), (andet_navn, sted) in andre_bilag.items():
                        if dnr == nr:
                            alle_fund.append((
                                samling_navn, "usikker", dok, ref, ctx,
                                f"Ikke fundet i '{samling_navn}' – men findes i "
                                f"'{andet_navn}' ({sted}). Krydssamlingsreference?"))
                            break
                    else:
                        alle_fund.append((samling_navn, status, dok, ref, ctx, forkl))
                    continue

                if pm:
                    nr = norm_nr(pm.group(1))
                    andre = andre_sektioner.get(nr, [])
                    if andre:
                        alle_fund.append((
                            samling_navn, "usikker", dok, ref, ctx,
                            f"Afsnit {nr} ikke fundet i '{samling_navn}' – "
                            f"men findes i '{', '.join(andre)}'. "
                            f"Krydssamlingsreference?"))
                    else:
                        alle_fund.append((samling_navn, status, dok, ref, ctx, forkl))
                    continue

            alle_fund.append((samling_navn, status, dok, ref, ctx, forkl))

    alle_sektioner = {}
    alle_bilag = {}
    for navn, (sek, bil) in indeks.items():
        alle_sektioner.update(sek)
        alle_bilag.update(bil)
    # Kontekstdokumenter mærkes tydeligt i opsummeringen, så det er synligt
    # at de kun er brugt til opslag - ikke tjekket for henvisninger.
    for navn, numre in kctx_sektioner.items():
        alle_sektioner[f"{navn} (kontekst - ikke tjekket)"] = numre
    for k, v in kctx_bilag.items():
        alle_bilag.setdefault(k, v)

    return alle_fund, alle_sektioner, alle_bilag


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


def validér(docs, _bilag_override=None, _sektioner_override=None):
    if _bilag_override is not None and _sektioner_override is not None:
        sektioner, bilag_def = _sektioner_override, _bilag_override
    else:
        sektioner, bilag_def = kortlaeg(docs)
    fund = []

    def tjek_sektion(nr, doknavn, tvunget_dok=None):
        nr = norm_nr(nr)
        if tvunget_dok:
            # henvisningen navngiver eksplicit "Kontraktens" - slå kun op der
            if nr in sektioner.get(tvunget_dok, set()):
                return "gyldig", f"henvisning til Kontraktens afsnit {nr} - findes i {tvunget_dok}"
            return "ugyldig", (f"ugyldig henvisning til Kontraktens afsnit {nr} - "
                               f"findes ikke i {tvunget_dok}")
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

    def find_kontrakt_dok():
        """Find dokumentet der repræsenterer selve kontrakten (ikke et bilag) -
        bruges når en henvisning eksplicit siger "Kontraktens punkt X"."""
        kandidater = [navn for navn in sektioner
                     if re.search(r"\bkontrakt", navn, re.IGNORECASE)
                     and not RE_BILAG_DEF.search(navn)]
        return kandidater[0] if len(kandidater) == 1 else None

    for d in docs:
        navn = d.path.name
        for p_el, _, tekst, _ in d.paras:
            if not tekst:
                continue
            optaget = []   # tegnpositioner der allerede er matchet

            def ledig(m):
                return not any(a < m.end() and m.start() < b for a, b in optaget)

            # 0) pladsholder-henvisninger: "punkt x.x", "Kapitel ?", "Bilag ??"
            # - skabelontekst hvor nummeret ikke er udfyldt. Altid ugyldig.
            for m in RE_PLADSHOLDER_SEKTION.finditer(tekst):
                if not ledig(m) or er_definition(tekst, m):
                    continue
                optaget.append((m.start(), m.end()))
                fund.append(("ugyldig", navn, m.group(0),
                             saetning(tekst, m.start(), m.end()),
                             f"henvisningen \"{m.group(0)}\" bruger en pladsholder "
                             f"i stedet for et rigtigt afsnitsnummer - mangler at "
                             f"blive udfyldt"))
            for m in RE_PLADSHOLDER_BILAG.finditer(tekst):
                if not ledig(m) or er_definition(tekst, m):
                    continue
                optaget.append((m.start(), m.end()))
                fund.append(("ugyldig", navn, m.group(0),
                             saetning(tekst, m.start(), m.end()),
                             f"henvisningen \"{m.group(0)}\" bruger en pladsholder "
                             f"i stedet for et rigtigt bilagsnummer - mangler at "
                             f"blive udfyldt"))

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

            # 2a) "bilag 3 og 4" – to bilag i ét udtryk
            for m in RE_BILAG_OG.finditer(tekst):
                if not ledig(m) or er_definition(tekst, m):
                    continue
                optaget.append((m.start(), m.end()))
                btype = m.group(1)
                for bnr_raw in (m.group(2), m.group(3)):
                    if not er_bilagsnummer(bnr_raw):
                        continue
                    bnr = norm_nr(bnr_raw)
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
                    fund.append((status, navn, f"{btype} {bnr_raw}",
                                 saetning(tekst, m.start(), m.end()), forkl))

            # 2b) bilagshenvisninger: "se Bilag 3", "Kontraktbilag 2"
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
                # RE_SEKTION har 7 grupper: g1 = præfiks-ord ("kontraktens"/
                # "nærværende ..."), (g2,g3,g4) = præfiks-variant, (g5,g6,g7) = standard
                praefiks = (m.group(1) or "").lower()
                nr1 = m.group(3) or m.group(6)
                nr2 = m.group(4) or m.group(7)
                if not nr1:
                    continue
                tvunget_dok = None
                if praefiks.startswith("kontrakt"):
                    tvunget_dok = find_kontrakt_dok()
                    if tvunget_dok == navn:
                        tvunget_dok = None  # selv-henvisning - ingen grund til at tvinge
                status1, forkl = tjek_sektion(nr1, navn, tvunget_dok)
                if nr2:                              # interval, fx 2.1-2.3
                    status2, forkl2 = tjek_sektion(nr2, navn, tvunget_dok)
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
def _normaliser_fund(fund):
    """Tillader bagudkompatible 5-tupler (status, dok, ref, ctx, forkl) fra
    validér() - de mappes til 6-tupler med samling=None, ligesom 6-tuplerne
    (samling, status, dok, ref, ctx, forkl) fra validér_samlinger()."""
    return [f if len(f) == 6 else (None,) + tuple(f) for f in fund]


SEM_FILLS = {"OK": PatternFill("solid", start_color="C6E0B4"),
             "ADVARSEL": PatternFill("solid", start_color="FFE699"),
             "FEJL": PatternFill("solid", start_color="F8CBAD")}


def _skriv_ark(wb, titel, raekker, vis_samling=False, intro=None, semantik_resultater=None):
    ws = wb.create_sheet(titel)
    rr = 1
    if intro:
        ws.cell(row=1, column=1, value=intro).font = Font(italic=True, color="666666")
        rr = 2
    hoved = ["Status"]
    bredder = [16]
    if vis_samling:
        hoved.append("Samling")
        bredder.append(18)
    hoved += ["Dokument", "Henvisning", "Sætning (kontekst)", "Strukturel vurdering"]
    bredder += [35, 22, 75, 55]
    if semantik_resultater:
        hoved += ["Semantisk vurdering", "Forslag"]
        bredder += [40, 50]
    for c, (h, b) in enumerate(zip(hoved, bredder), 1):
        cell = ws.cell(row=rr, column=c, value=h)
        cell.font = Font(bold=True)
        ws.column_dimensions[cell.column_letter].width = b
    sidste_kol = len(hoved)
    orden = {"ugyldig": 0, "usikker": 1, "stoej": 2, "gyldig": 3}
    for samling, status, dok, ref, ktx, forkl in sorted(raekker, key=lambda f: orden[f[1]]):
        rr += 1
        ws.cell(row=rr, column=1, value=SYMBOL[status]).fill = FILLS[status]
        col = 2
        if vis_samling:
            ws.cell(row=rr, column=col, value=samling)
            col += 1
        for v in (dok, ref, ktx, forkl):
            ws.cell(row=rr, column=col, value=v)
            col += 1
        if semantik_resultater:
            sem = semantik_resultater.get((dok, ref, ktx), {})
            sem_vurd = sem.get("vurdering", "")
            sem_forkl = sem.get("forklaring", "")
            sem_forslag = sem.get("forslag", "")
            sem_celle = ws.cell(row=rr, column=col,
                                value=f"{sem_vurd}\n{sem_forkl}" if sem_forkl else sem_vurd)
            sem_fill = SEM_FILLS.get(sem_vurd)
            if sem_fill:
                sem_celle.fill = sem_fill
            col += 1
            ws.cell(row=rr, column=col, value=sem_forslag)
            col += 1
        for c in range(1, sidste_kol + 1):
            ws.cell(row=rr, column=c).border = THIN
            ws.cell(row=rr, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    if rr == (2 if intro else 1):
        ws.cell(row=rr + 1, column=2, value="(ingen)")


def skriv_rapport(fund, sektioner, bilag_def, outfile, semantik_resultater=None):
    fund = _normaliser_fund(fund)
    vis_samling = any(s is not None for s, *_ in fund)
    wb = Workbook()
    wb.remove(wb.active)

    n = {s: sum(1 for f in fund if f[1] == s)
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
               [f for f in fund if f[1] in ("ugyldig", "usikker")], vis_samling,
               intro="De henvisninger der kræver menneskelig vurdering - ugyldige øverst.",
               semantik_resultater=semantik_resultater)
    # Fane 3: Gyldige
    _skriv_ark(wb, "Gyldige", [f for f in fund if f[1] == "gyldig"], vis_samling,
               semantik_resultater=semantik_resultater)
    # Fane 4: Støj
    _skriv_ark(wb, "Støj (kan ignoreres)",
               [f for f in fund if f[1] == "stoej"], vis_samling,
               intro="Henvisninger til appendikser/bilag der kun findes som overskrifter "
                     "inde i andre dokumenter - normalt ikke fejl.",
               semantik_resultater=semantik_resultater)

    wb.save(outfile)
    return n


def _dok_tekst(d):
    """Fuldt tekstindhold af ét dokument til semantisk analyse.
    Nummererede afsnit (Word-nummerering ELLER manuelt tastet '4.2 Betaling')
    får et '# nr'-præfiks foran teksten, så semantik.py's opslag af
    'punkt X.Y' kan finde det rigtige sted i dokumentet - ellers går
    afsnitsnummeret tabt, fordi det normalt ligger i label, ikke i teksten."""
    linjer = []
    for _, label, t, _ in d.paras:
        if not t:
            continue
        nr = None
        if label and label[:1].isdigit():
            nr = label
        else:
            m = RE_MANUELT_NR.match(t)
            if m and len(t) <= 150:
                nr = m.group(1)
        linjer.append(f"# {nr} {t}" if nr else t)
    return "\n".join(linjer)


def byg_dokument_indhold(samlinger):
    """Returnerer {doknavn: fuldt tekstindhold} for alle dokumenter."""
    indhold = {}
    for samling_navn, docs in samlinger:
        for d in docs:
            indhold[d.path.name] = _dok_tekst(d)
    return indhold


def byg_kontekst_docs(kilder, maal_docs):
    """Indlæser kontekst-dokumenter til opslag (afsnits-/bilagsnumre) uden at
    de selv skal tjekkes for henvisninger.

    kilder: liste af filstier og/eller mappestier (str/Path). En mappe
        udvides til alle .docx-filer direkte i mappen (ikke undermapper);
        Words åbne-låsefiler ('~$...') springes automatisk over.
    maal_docs: Docx-objekter der allerede indgår som mål (tjekkes for
        henvisninger) - udelades her, så de ikke indlæses/tælles dobbelt.

    Returnerer liste af Docx, klar til validér_samlinger(..., kontekst_docs=...).
    """
    if not kilder:
        return []
    maal_stier = {d.path for d in maal_docs}
    kontekst_stier = {}   # dict bruges som ordnet sæt (bevarer rækkefølge)
    for sti in kilder:
        p = Path(sti)
        if p.is_dir():
            for f in sorted(p.glob("*.docx")):
                if f.name.startswith("~$"):   # Words åbne-låsefil - spring over
                    continue
                kontekst_stier[f.resolve()] = None
        else:
            kontekst_stier[p.resolve()] = None
    return [ex.Docx(rp) for rp in kontekst_stier if rp not in maal_stier]


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Validér krydshenvisninger på tværs af udbudspakke")
    ap.add_argument(
        "--samling", nargs="+", action="append", metavar=("NAVN", "FIL"),
        help="--samling NAVN fil1.docx fil2.docx ...")
    ap.add_argument("-o", "--output", default="krydstjek.xlsx")
    ap.add_argument(
        "--semantik", action="store_true",
        help="Kør semantisk AI-analyse af ugyldige/usikre henvisninger")
    ap.add_argument(
        "--kontekst", nargs="+", metavar="FIL_ELLER_MAPPE",
        help="Filer og/eller mapper der KUN bruges til opslag af afsnits- og "
             "bilagsnumre (fx udbudsbetingelserne) - bruges ikke til at finde "
             "henvisninger og optræder ikke som fund i rapporten. En mappe "
             "udvides til alle .docx-filer direkte i mappen (ikke undermapper).")
    ap.add_argument("filer_pos", nargs="*", metavar="FIL",
                    help="Filer uden samling (bagudkompatibelt)")
    args = ap.parse_args()

    # Byg samlinger: liste af (samlingsnavn, [Docx-objekter])
    samlinger = []
    if args.samling:
        for gruppe in args.samling:
            navn = gruppe[0]
            filer = gruppe[1:]
            samlinger.append((navn, [ex.Docx(f) for f in filer]))
    if args.filer_pos and not args.samling:
        samlinger.append(("Standard", [ex.Docx(f) for f in args.filer_pos]))
    if not samlinger:
        ap.error("Angiv mindst én --samling eller angiv filer direkte")

    # Kontekstdokumenter: filer/mapper der kun bruges til opslag.
    maal_docs = [d for _, docs in samlinger for d in docs]
    kontekst_docs = byg_kontekst_docs(args.kontekst, maal_docs)
    if kontekst_docs:
        print(f"Kontekst: {len(kontekst_docs)} dokument(er) brugt til opslag "
              f"(tjekkes ikke selv for henvisninger)")

    fund, alle_sektioner, alle_bilag = validér_samlinger(samlinger, kontekst_docs=kontekst_docs)

    semantik_resultater = {}
    if args.semantik:
        try:
            import semantik
            dok_indhold = byg_dokument_indhold(samlinger)
            for d in kontekst_docs:
                dok_indhold[d.path.name] = _dok_tekst(d)
            semantik_resultater = semantik.analysér_batch(fund, dok_indhold)
        except ImportError:
            print("ADVARSEL: semantik.py ikke fundet — springer over")
        except Exception as e:
            print(f"ADVARSEL: Semantisk analyse fejlede: {e}")

    n = skriv_rapport(fund, alle_sektioner, alle_bilag,
                      args.output, semantik_resultater)
    print(f"Krydstjek gennemfort -> {args.output}")
    print(f"  Ugyldige: {n['ugyldig']}   Usikre: {n['usikker']}   "
          f"Gyldige: {n['gyldig']}   Stoj: {n['stoej']}")
    print(f"  -> Til gennemgang: {n['ugyldig'] + n['usikker']}")


if __name__ == "__main__":
    sys.exit(main())
