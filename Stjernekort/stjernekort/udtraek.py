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
BELOEB_RE = re.compile(r"(\d{1,3}(?:\.\d{3})+|\d+)\s*(?:DKK|kr\.)\s*pr\.\s*Arbejdsdag", re.IGNORECASE)
MAKS_RE = re.compile(r"(?:ikke overstige|maks\.?)\s*(\d{1,3}(?:\.\d{3})+|\d+)\s*(?:DKK|kr\.)", re.IGNORECASE)


def _bod_detalje(tekster: list[str]) -> str:
    t = " ".join(tekster)
    b, m = BELOEB_RE.search(t), MAKS_RE.search(t)
    if not b:
        return ""
    return f"{b.group(1)} kr./AD" + (f", maks. {m.group(1)} kr." if m else "")
SAETNING_RE = re.compile(r"(?<=[.;:])\s+(?=[A-ZÆØÅ])")


def laes_ordbog(sti: str | Path) -> dict:
    """Læser en ordbog. Har den "bygger_paa", lægges den oven på grundordbogen:
    milepæle med samme ID erstattes (eller fjernes med "fjern": true), nye tilføjes;
    søgeord, tabeller og regler lægges til."""
    sti = Path(sti)
    if sti.suffix.lower() == ".xlsx":
        from .ordbog_excel import laes_ordbog_excel
        o = laes_ordbog_excel(sti)
    else:
        o = json.loads(sti.read_text(encoding="utf-8"))
    if not o.get("bygger_paa"):
        return o
    g = laes_ordbog(sti.parent / o["bygger_paa"])
    ms = {m["id"]: m for m in g["milepaele"]}
    for m in o.get("milepaele", []):
        if m.get("fjern"):
            ms.pop(m["id"], None)
        else:
            ms[m["id"]] = {**ms.get(m["id"], {}), **m}
    g["milepaele"] = sorted(ms.values(), key=lambda m: m.get("raekkefoelge", 999))
    for typ, ord_ in o.get("retsvirkninger", {}).items():
        g["retsvirkninger"].setdefault(typ, [])
        g["retsvirkninger"][typ] += [x for x in ord_ if x not in g["retsvirkninger"][typ]]
    g.setdefault("tabeller", {}).update(o.get("tabeller", {}))
    g["regler"] = g.get("regler", []) + o.get("regler", [])
    for rolle, navne in o.get("andre_formers_navne", {}).items():
        g.setdefault("andre_formers_navne", {}).setdefault(rolle, [])
        g["andre_formers_navne"][rolle] += [n for n in navne if n not in g["andre_formers_navne"][rolle]]
    g["paradigme"] = f'{g.get("paradigme", "")} + {o.get("paradigme", "")}'
    return g


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
                                    etiket=m.get("etiket", ""), relativ_tid=m.get("relativ_tid", "")))
    roller = {m["id"]: m.get("rolle", "") for m in ordbog["milepaele"]}
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
            # 2b. Regler: indledning efterfulgt af punktopstilling, hvor et punkt
            #     indeholder en bestemt vending -> retsvirkning for milepæle med givne roller
            for rg in ordbog.get("regler", []):
                if b.type == "afsnit" and rg["indledning"] in b.tekst.lower():
                    for nb in blokke[idx + 1:idx + 40]:
                        if nb.type != "afsnit" or nb.punkt != b.punkt:
                            break
                        if rg["punkt_indeholder"] in nb.tekst.lower():
                            for mid, rolle in roller.items():
                                if rolle in rg["roller"]:
                                    tilfoej(mid, rg["retsvirkning"], _kort(nb.tekst, 90),
                                            Kilde(dok, b.punkt, _kort(nb.tekst)), "høj")
            # 3. Sætninger: nøgleord og milepæl skal stå i samme sætning
            if b.type == "afsnit":
                for saetning in SAETNING_RE.split(b.tekst):
                    mids = find_milepael(saetning, ordbog)
                    if not mids:
                        continue
                    t = saetning.lower()
                    for typ, moenstre in ordbog["retsvirkninger"].items():
                        if any(re.search(mo, t) for mo in moenstre):
                            detalje = ""
                            if typ == "bod":   # beløb står ofte i næste afsnit i samme punkt
                                naeste = [x.tekst for x in blokke[idx:idx + 3] if x.punkt == b.punkt]
                                detalje = _bod_detalje(naeste)
                            for mid in mids:
                                tilfoej(mid, typ, detalje, Kilde(dok, b.punkt, _kort(saetning)),
                                        "middel")

    bp.henvisninger = tjek_henvisninger(docs)
    bp.advarsler = find_fremmede_begreber(docs, ordbog)
    return bp


def find_fremmede_begreber(docs, ordbog) -> list[tuple[str, str, str]]:
    """Steder hvor teksten bruger en anden kontraktforms navn for en rolle,
    fx "Overtagelsesprøve" i en servicekontrakt. Returnerer (begreb, kilde, citat)."""
    res = []
    for rolle, navne in ordbog.get("andre_formers_navne", {}).items():
        for dok, blokke in docs.items():
            for b in blokke:
                if b.forside:
                    continue
                t = b.tekst.lower()
                for n in navne:
                    if re.search(r"\b" + re.escape(n) + r"(?:s|en|ens|er|erne|ene)?\b", t):
                        res.append((f"{n} ({rolle})", f"{dok} pkt. {b.punkt}", _kort(b.tekst)))
    return res


# --- Henvisninger ----------------------------------------------------------

REF_RE = re.compile(
    r"(?P<ctx>(?:[Kk]ontraktens\s+|(?:bilag|Bilag)\s+(?P<bilag>\d+[A-Z]?)(?:,\s*Appendiks\s+\w+)?,?\s+))?"
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
                if ctx.lower().startswith("kontraktens"):
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
