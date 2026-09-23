"""Genererer et udkast til en udbudsordbog ud fra selve kontrakten og bilagene.

Kilder (i den rækkefølge de bruges):
  1. Hovedtidsplanens tabel  -> milepæle og deres rækkefølge
  2. Betalingsplanens tabel  -> milepæle (betalingsudløsende)
  3. Definitionsafsnittet    -> prøver og dage (fx "Overtagelsesprøve", "Ibrugtagningsdag")
Hver milepæl får en rolle, hvis navnet ligner en kendt rolle, og udkastet lægges
oven på den grundordbog (kontraktform), der passer bedst. Alt markeres "forslag",
så brugeren gennemser det i Excel.
"""
from __future__ import annotations

import re
from pathlib import Path

from .docx_reader import laes_mappe
from .udtraek import laes_ordbog

ROLLE_MOENSTRE = [  # (rolle, regex på milepælsnavn)
    ("KONTRAKTINDGÅELSE", r"kontrakt(indgåelse|underskrivelse)|ikrafttræden"
                          r"|afklaringsfase\w* (igangsæt|påbegynd)|påbegyndelse af afklaringsfase"),
    ("AFKLARING_SLUT", r"afklaringsfase\w* (godkend|afslut)|godkendelse af afklaringsfase|leverancebeskrivelse"),
    ("HOVEDPRØVE", r"overtagelsesprøve|ibrugtagningsprøve"),
    ("DRIFTSPRØVE_START", r"driftsprøve\w* (igangsæt|påbegynd)"),
    ("DRIFTSPRØVE", r"driftsprøve"),
    ("KONTRAKTOPHØR", r"kontraktophør|kontrakten ophører"),
]
TABEL_TIDSPLAN = re.compile(r"milepæl", re.I)
TABEL_BETALING = re.compile(r"%")
DEFINERET_PROEVE = re.compile(r"^(?:Ved\s+)?([A-ZÆØÅ][a-zæøå]+(?:prøve|prøven|sdag|sdagen))\b(?:\s+forstås)?")
IKKE_MILEPAEL = {"arbejdsdag", "arbejdsdagen"}
FORLED = re.compile(r"^(gennemførelse|godkendelse|påbegyndelse|afslutning|afholdelse) af (\d+ )?")


def _rolle(navn: str) -> str:
    n = navn.lower()
    for rolle, rx in ROLLE_MOENSTRE:
        if re.search(rx, n):
            return rolle
    return ""


def _synonymer(navn: str) -> list[str]:
    """Navnet i små bogstaver + bøjninger:
    'X godkendes' -> 'godkendelse af X'; 'Gennemførelse af 2 prøvekonverteringer' -> 'prøvekonvertering'."""
    n = re.sub(r"\s+", " ", navn.lower()).strip(" .")
    s = {n}
    m = re.match(r"(.+?) godkendes$", n)
    if m:
        s.add("godkendelse af " + m.group(1))
    rest = FORLED.sub("", n)
    if rest != n and " " not in rest:
        s.add(re.sub(r"(erne|er|en)$", "", rest))      # enkeltord: brug stammen
    return sorted(s)


def generer_udkast(mappe: str | Path, grundordboeger: list[Path]) -> dict:
    docs = laes_mappe(mappe)
    kand: dict[str, dict] = {}          # nøgle = normaliseret navn
    orden = 0

    def tilfoej(navn, kilde, typ):
        nonlocal orden
        navn = re.sub(r"\s+", " ", navn).strip(" .")
        if not navn or navn.lower() in ("i alt", "...", "…") or len(navn) > 90:
            return
        k = navn.lower()
        if k in IKKE_MILEPAEL:
            return
        for eksisterende in kand:                     # samme milepæl med kortere/længere navn
            if k.startswith(eksisterende) or eksisterende.startswith(k):
                kand[eksisterende]["synonymer"] = sorted(set(kand[eksisterende]["synonymer"]) | set(_synonymer(navn)))
                kand[eksisterende]["kilde"] += f"; {kilde}"
                return
        if k not in kand:
            orden += 10
            kand[k] = {"navn": navn, "rolle": _rolle(navn), "type": typ, "raekkefoelge": orden,
                       "synonymer": _synonymer(navn), "kilde": kilde}

    for dok, blokke in docs.items():                      # 1. hovedtidsplan
        for b in blokke:
            if b.type == "tabel" and b.raekker and not b.forside \
                    and TABEL_TIDSPLAN.search(" ".join(b.raekker[0])) and "dato" in " ".join(b.raekker[0]).lower():
                for r in b.raekker[1:]:
                    if len(r) > 1 or (r and r[0] and not r[0].lower().endswith("fasen")):
                        tilfoej(r[0], f"{dok} pkt. {b.punkt} (hovedtidsplan)", "milepæl")
    for dok, blokke in docs.items():                      # 2. betalingsplan
        for b in blokke:
            if b.type == "tabel" and b.raekker and not b.forside \
                    and TABEL_BETALING.search(" ".join(b.raekker[0])) \
                    and any(re.search(r"\d+\s*%", " ".join(r)) for r in b.raekker[1:]):
                for r in b.raekker[1:]:
                    tilfoej(r[0], f"{dok} pkt. {b.punkt} (betalingsplan)", "milepæl")
    kontrakt = docs.get("Kontrakt", [])                   # 3. definerede prøver/dage
    for b in kontrakt:
        if b.overskrift.lower().startswith("definition") and b.type == "afsnit":
            m = DEFINERET_PROEVE.match(b.tekst)
            if m and not any(m.group(1).lower() in k for k in kand):
                tilfoej(m.group(1), f"Kontrakt pkt. {b.punkt} (definition)", "frist")

    # Typer: prøver og faseafslutninger
    for c in kand.values():
        if c["rolle"] in ("HOVEDPRØVE", "DRIFTSPRØVE"):
            c["type"] = "prøve"
        elif c["rolle"] == "AFKLARING_SLUT":
            c["type"] = "faseafslutning"
        elif c["rolle"] == "KONTRAKTINDGÅELSE":
            c["type"] = "kontraktindgåelse"

    # Vælg grundordbog: flest roller/synonymer tilfælles med kandidaterne
    tekst = " ".join(b.tekst.lower() for b in kontrakt)
    bedst, score = None, -1
    for g in grundordboeger:
        o = laes_ordbog(g)
        s = sum(tekst.count(syn) for m in o["milepaele"] if m.get("rolle") for syn in m["synonymer"])
        if s > score:
            bedst, score, go = g, s, o
    grund_roller = {m.get("rolle") for m in go["milepaele"]}
    grund_syn = {s for m in go["milepaele"] for s in m["synonymer"]}

    ms, n = [], 0
    for c in sorted(kand.values(), key=lambda c: c["raekkefoelge"]):
        if c["rolle"] and c["rolle"] in grund_roller:
            continue                                   # dækket af kontraktformen
        if any(s in grund_syn for s in c["synonymer"]):
            continue
        n += 1
        ms.append({"id": f"U{n}", "rolle": c["rolle"], "navn": c["navn"], "etiket": "",
                   "type": c["type"], "raekkefoelge": _placering(c, kand, go),
                   "synonymer": c["synonymer"], "kilde": c["kilde"]})
    return {"paradigme": f"Udbud: {Path(mappe).name} (FORSLAG – gennemses)",
            "bygger_paa": Path(bedst).name, "milepaele": ms}


def _placering(c, kand, go) -> int:
    """Rækkefølge mellem grundordbogens milepæle: find nærmeste forgænger med rolle."""
    rolle_orden = {m.get("rolle"): m.get("raekkefoelge", 0) for m in go["milepaele"] if m.get("rolle")}
    foer = [x for x in kand.values() if x["raekkefoelge"] < c["raekkefoelge"] and x["rolle"] in rolle_orden]
    base = rolle_orden[foer[-1]["rolle"]] if foer else 10
    efter = sorted(v for v in rolle_orden.values() if v > base)
    loft = efter[0] if efter else base + 10
    trin = [x for x in kand.values() if x["raekkefoelge"] < c["raekkefoelge"]
            and (not foer or x["raekkefoelge"] > foer[-1]["raekkefoelge"])]
    return base + (loft - base) * (len(trin) + 1) // 10 or base + 1
