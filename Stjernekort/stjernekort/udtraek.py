"""Udtræk af milepæle, retsvirkninger og henvisninger fra kontrakt + bilag.

Tre kilder i faldende sikkerhed:
  1. Tabeller (bodstabel, betalingsplan, hovedtidsplan)          -> "høj"
  2. Forsinkelsesdefinitionen (punktopstilling efter indledning)  -> "høj"
  3. Sætninger med nøgleord + milepælssynonym                     -> "middel"
Alt markeres med kilde (dokument + punkt), så brugeren kan efterprøve.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from .docx_reader import Blok, laes_mappe, punkter
from .model import Blueprint, Henvisning, Kilde, Milepael, Retsvirkning

PCT_RE = re.compile(r"(\d+(?:,\d+)?)\s*%")
SAETNING_RE = re.compile(r"(?<=[.;:])\s+(?=[A-ZÆØÅ])")


def laes_ordbog(sti: str | Path) -> dict:
    return json.loads(Path(sti).read_text(encoding="utf-8"))


def find_milepael(tekst: str, ordbog: dict) -> list[str]:
    """Milepæls-id'er hvis synonymer optræder i teksten (længste synonym vinder)."""
    t = tekst.lower()
    fund = []
    for m in ordbog["milepaele"]:
        if any(s in t for s in m["synonymer"]):
            fund.append(m["id"])
    return fund


def _match_tabel(b: Blok, ordbog: dict) -> str | None:
    if b.type != "tabel" or not b.raekker:
        return None
    hoved = " ".join(b.raekker[0]).lower()
    for navn, spec in ordbog["tabeller"].items():
        if all(o in hoved for o in spec["overskrift_indeholder"]):
            return spec["type"]
    return None


def _kort(tekst: str, n: int = 140) -> str:
    return tekst if len(tekst) <= n else tekst[: n - 1] + "…"


def udtraek(mappe: str | Path, ordbog: dict) -> Blueprint:
    docs = laes_mappe(mappe)
    bp = Blueprint(navn=f"Udtræk: {Path(mappe).name}")
    set_rv: set[tuple] = set()

    def tilfoej(mid: str, typ: str, detalje: str, kilde: Kilde, sikkerhed: str):
        noegle = (mid, typ, kilde.dokument, kilde.punkt)
        if noegle in set_rv:
            return
        set_rv.add(noegle)
        rv = Retsvirkning(mid, typ, detalje, kilde, oprindelse=f"udtræk ({sikkerhed})")
        bp.retsvirkninger.append(rv)

    # Milepæle: alle fra ordbogen, med de steder de nævnes
    for i, m in enumerate(ordbog["milepaele"], 1):
        bp.milepaele.append(Milepael(m["id"], m["navn"], m["type"], i, m.get("option", False),
                                    etiket=m.get("etiket", "")))
    for dok, blokke in docs.items():
        for b in blokke:
            if b.forside or b.type == "overskrift":
                continue
            for mid in find_milepael(b.tekst, ordbog):
                ms = bp.milepael(mid)
                if len(ms.kilder) < 25:
                    ms.kilder.append(Kilde(dok, b.punkt, _kort(b.tekst)))

    for dok, blokke in docs.items():
        for idx, b in enumerate(blokke):
            if b.forside:
                continue
            kilde = Kilde(dok, b.punkt, _kort(b.tekst))
            # 1. Tabeller
            ttype = _match_tabel(b, ordbog)
            if ttype in ("bod", "betaling"):
                for r in b.raekker[1:]:
                    for mid in find_milepael(r[0], ordbog):
                        detalje = " ".join(r[1:]).replace(" /  / ", "; ")
                        if ttype == "betaling":
                            p = PCT_RE.search(detalje)
                            detalje = f"{p.group(1)} %" if p else detalje
                        tilfoej(mid, ttype, _kort(detalje, 90),
                                Kilde(dok, b.punkt, _kort(" | ".join(r))), "høj")
            elif ttype == "tidsplan":
                for r in b.raekker[1:]:
                    for mid in find_milepael(r[0], ordbog):
                        bp.milepael(mid).kilder.insert(0, Kilde(dok, b.punkt, "Hovedtidsplan: " + r[0]))
            # 2. Forsinkelsesdefinitionen
            fs = ordbog.get("forsinkelse")
            if fs and b.type == "afsnit" and fs["indledning"] in b.tekst.lower():
                for nb in blokke[idx + 1:]:
                    if "bullet" not in nb.stil.lower():
                        break
                    for mid in find_milepael(nb.tekst, ordbog):
                        for typ in fs["retsvirkninger"]:
                            tilfoej(mid, typ, "ved forsinkelse, jf. forsinkelsesbestemmelsen",
                                    kilde, "høj")
            # 3. Sætninger: nøgleord og milepæl skal stå i samme sætning
            if b.type == "afsnit":
                for saetning in SAETNING_RE.split(b.tekst):
                    mids = find_milepael(saetning, ordbog)
                    if not mids:
                        continue
                    t = saetning.lower()
                    for typ, moenstre in ordbog["retsvirkninger"].items():
                        if any(re.search(mo, t) for mo in moenstre):
                            for mid in mids:
                                tilfoej(mid, typ, "", Kilde(dok, b.punkt, _kort(saetning)),
                                        "middel")

    bp.henvisninger = tjek_henvisninger(docs)
    return bp


# --- Henvisninger ----------------------------------------------------------

REF_RE = re.compile(
    r"(?P<ctx>(?:Kontraktens\s+|(?:bilag|Bilag)\s+(?P<bilag>\d+[A-Z]?)(?:,\s*Appendiks\s+\w+)?,?\s+))?"
    r"(?:punkt|pkt\.)\s+(?P<punkt>\d+(?:\.\d+)*)"
    r"(?P<efter>\s+i\s+udbudsbetingelserne)?")


def tjek_henvisninger(docs: dict[str, list[Blok]]) -> list[Henvisning]:
    opslag = {dok: punkter(bl) for dok, bl in docs.items()}
    res: list[Henvisning] = []
    for dok, blokke in docs.items():
        for b in blokke:
            if b.forside or b.type == "overskrift":
                continue
            for m in REF_RE.finditer(b.tekst):
                if m.group("efter") or "udbudsbetingelser" in b.tekst[m.end():m.end() + 30]:
                    continue                                   # ekstern henvisning
                ctx = m.group("ctx") or ""
                if "Appendiks" in ctx:
                    continue                                   # appendiks = regneark mv.
                if ctx.startswith("Kontraktens"):
                    til = "Kontrakt"
                elif m.group("bilag"):
                    til = f"Bilag {m.group('bilag').upper()}"
                else:
                    til = dok
                fundet = til in opslag and m.group("punkt") in opslag[til]
                if til not in opslag:
                    fundet = None  # dokumentet er ikke med i mappen
                res.append(Henvisning(Kilde(dok, b.punkt, _kort(b.tekst)), til,
                                      m.group("punkt"), fundet))
    return res
