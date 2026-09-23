"""Læser .docx til en flad liste af blokke med beregnede punktnumre.

Overskrifterne i paradigmerne er autonummererede, så "33.2" står ikke i
teksten. Numrene beregnes ud fra overskriftsniveauerne (Heading 1-3).
Alt før første Heading 1 (vejledning, indholdsfortegnelse, bilagsliste)
markeres som forside og indgår ikke i udtrækket af retsvirkninger.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

HEADING_RE = re.compile(r"^Heading (\d)$")
BILAG_ID_RE = re.compile(r"Bilag[_ ]0*(\d+[A-Z]?)", re.IGNORECASE)
APPENDIKS_RE = re.compile(r"Appendiks[_ ]([A-Z])\b")


@dataclass
class Blok:
    dokument: str            # "Kontrakt", "Bilag 2", "Bilag 1A" ...
    punkt: str               # beregnet punktnummer, fx "33.2" ("" på forsiden)
    overskrift: str          # nærmeste overskrift
    type: str                # "overskrift" | "afsnit" | "tabel"
    tekst: str = ""
    stil: str = ""           # Word-typografi, fx "List Bullet"
    raekker: list[list[str]] = field(default_factory=list)  # kun for tabeller
    forside: bool = False    # True før første Heading 1


def dokument_id(filnavn: str) -> str | None:
    """'11._Bilag_02_Hovedtidsplan.docx' -> 'Bilag 2';
    '12. Bilag 03A_..._Appendiks D_...' -> 'Bilag 3A, Appendiks D';
    kontrakten -> 'Kontrakt'; øvrige (erklæringer mv.) -> None."""
    m = BILAG_ID_RE.search(filnavn)
    if m:
        a = APPENDIKS_RE.search(filnavn)
        return f"Bilag {m.group(1).upper()}" + (f", Appendiks {a.group(1)}" if a else "")
    return "Kontrakt" if "kontrakt" in filnavn.lower() else None


def _blokke_i_body(d):
    for el in d.element.body.iterchildren():
        if el.tag.endswith("}p"):
            yield Paragraph(el, d)
        elif el.tag.endswith("}tbl"):
            yield Table(el, d)


def _tabel_raekker(t: Table) -> list[list[str]]:
    raekker = []
    for r in t.rows:
        celler: list[str] = []
        for c in r.cells:
            x = c.text.strip().replace("\n", " / ")
            if not celler or celler[-1] != x:   # flettede celler gentages
                celler.append(x)
        raekker.append(celler)
    return raekker


def laes_docx(sti: str | Path) -> list[Blok]:
    sti = Path(sti)
    dok = dokument_id(sti.name)
    d = docx.Document(str(sti))
    tael = [0, 0, 0]
    punkt, overskrift, forside = "", "", True
    blokke: list[Blok] = []
    for b in _blokke_i_body(d):
        if isinstance(b, Paragraph):
            tekst = b.text.strip()
            if not tekst:
                continue
            m = HEADING_RE.match(b.style.name or "")
            if m and int(m.group(1)) <= 3:
                niv = int(m.group(1)) - 1
                tael[niv] += 1
                for i in range(niv + 1, 3):
                    tael[i] = 0
                punkt = ".".join(str(x) for x in tael[: niv + 1])
                overskrift, forside = tekst, False
                blokke.append(Blok(dok, punkt, overskrift, "overskrift", tekst))
            elif (b.style.name or "").lower().startswith("toc"):
                continue
            else:
                blokke.append(Blok(dok, punkt, overskrift, "afsnit", tekst,
                                   stil=b.style.name or "", forside=forside))
        else:
            r = _tabel_raekker(b)
            blokke.append(Blok(dok, punkt, overskrift, "tabel",
                               " | ".join(" | ".join(x) for x in r), raekker=r,
                               forside=forside))
    return blokke


def laes_mappe(mappe: str | Path) -> dict[str, list[Blok]]:
    """Kontrakt og bilag i mappen, nøgle = dokument-id ('Kontrakt', 'Bilag 2' ...).
    Filer, der hverken er kontrakt eller bilag (erklæringer mv.), springes over."""
    res = {}
    for p in sorted(Path(mappe).glob("*.docx")):
        did = dokument_id(p.name)
        if p.name.startswith("~$") or did is None:
            continue
        res[did] = laes_docx(p)
    return res


def punkter(blokke: list[Blok]) -> dict[str, str]:
    """Punktnummer -> overskrift, til opslag af henvisninger."""
    return {b.punkt: b.tekst for b in blokke if b.type == "overskrift"}
