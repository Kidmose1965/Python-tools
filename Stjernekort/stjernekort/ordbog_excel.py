"""Ordbøger som Excel-filer, så de kan vedligeholdes uden at røre kode.

Faner:
  Info            Paradigme, Bygger på
  Milepæle        ID, Rolle, Navn, Etiket, Type, Rækkefølge, Option, Relativ tid, Synonymer, Fjern
  Retsvirkninger  Type, Søgeord
  Tabeller        Navn, Type, Overskrift indeholder
  Regler          Indledning, Punkt indeholder, Retsvirkning, Roller
  Andre former    Rolle, Navne i andre kontraktformer
Lister i en celle adskilles med semikolon.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

MS_KOL = ["ID", "Rolle", "Navn", "Etiket", "Type", "Rækkefølge", "Option", "Relativ tid",
          "Synonymer", "Fjern", "Kilde"]


def _liste(v) -> list[str]:
    return [x.strip() for x in str(v or "").split(";") if x.strip()]


def _ark(wb, navn, kol, bredder):
    ws = wb.create_sheet(navn)
    ws.append(kol)
    for c, b in zip(ws[1], bredder):
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
        ws.column_dimensions[c.column_letter].width = b
    ws.freeze_panes = "A2"
    return ws


def skriv_ordbog(o: dict, sti: str | Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    ws = _ark(wb, "Info", ["Felt", "Værdi"], [16, 60])
    ws.append(["Paradigme", o.get("paradigme", "")])
    ws.append(["Bygger på", (o.get("bygger_paa") or "").replace(".json", ".xlsx")])
    ws = _ark(wb, "Milepæle", MS_KOL, [6, 22, 40, 28, 16, 12, 8, 34, 70, 8, 50])
    for m in o.get("milepaele", []):
        ws.append([m["id"], m.get("rolle", ""), m.get("navn", ""), m.get("etiket", ""),
                   m.get("type", ""), m.get("raekkefoelge", ""), "ja" if m.get("option") else "",
                   m.get("relativ_tid", ""), "; ".join(m.get("synonymer", [])),
                   "ja" if m.get("fjern") else "", m.get("kilde", "")])
    ws = _ark(wb, "Retsvirkninger", ["Type", "Søgeord"], [20, 90])
    for t, ord_ in o.get("retsvirkninger", {}).items():
        ws.append([t, "; ".join(ord_)])
    ws = _ark(wb, "Tabeller", ["Navn", "Type", "Overskrift indeholder"], [14, 12, 60])
    for n, spec in o.get("tabeller", {}).items():
        ws.append([n, spec["type"], "; ".join(spec["overskrift_indeholder"])])
    ws = _ark(wb, "Regler", ["Indledning", "Punkt indeholder", "Retsvirkning", "Roller"],
              [45, 40, 14, 30])
    fs = o.get("forsinkelse")
    if fs:
        ws.append([fs["indledning"], "*", "; ".join(fs["retsvirkninger"]), "(milepæle i punktopstillingen)"])
    for r in o.get("regler", []):
        ws.append([r["indledning"], r["punkt_indeholder"], r["retsvirkning"], "; ".join(r["roller"])])
    ws = _ark(wb, "Andre former", ["Rolle", "Navne i andre kontraktformer"], [22, 70])
    for rolle, navne in o.get("andre_formers_navne", {}).items():
        ws.append([rolle, "; ".join(navne)])
    for w in wb.worksheets:
        for r in w.iter_rows(min_row=2):
            for c in r:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    sti = Path(sti)
    wb.save(sti)
    return sti


def laes_ordbog_excel(sti: str | Path) -> dict:
    wb = load_workbook(sti, data_only=True)
    info = {r[0]: r[1] for r in wb["Info"].iter_rows(min_row=2, values_only=True) if r[0]}
    o: dict = {"paradigme": info.get("Paradigme") or "",
               "bygger_paa": info.get("Bygger på") or None,
               "milepaele": [], "retsvirkninger": {}, "tabeller": {}, "regler": [],
               "andre_formers_navne": {}}
    for r in wb["Milepæle"].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        d = dict(zip(MS_KOL, list(r) + [None] * len(MS_KOL)))
        m = {"id": str(d["ID"]), "rolle": d["Rolle"] or "", "navn": d["Navn"] or "",
             "etiket": d["Etiket"] or "", "type": d["Type"] or "",
             "raekkefoelge": int(d["Rækkefølge"] or 999), "option": (d["Option"] or "") == "ja",
             "relativ_tid": d["Relativ tid"] or "", "synonymer": [s.lower() for s in _liste(d["Synonymer"])]}
        if (d["Fjern"] or "") == "ja":
            m["fjern"] = True
        o["milepaele"].append(m)
    for r in wb["Retsvirkninger"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            o["retsvirkninger"][r[0]] = _liste(r[1])
    for r in wb["Tabeller"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            o["tabeller"][r[0]] = {"type": r[1], "overskrift_indeholder": [x.lower() for x in _liste(r[2])]}
    for r in wb["Regler"].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        if r[1] == "*":
            o["forsinkelse"] = {"indledning": r[0].lower(), "retsvirkninger": _liste(r[2])}
        else:
            o["regler"].append({"indledning": r[0].lower(), "punkt_indeholder": (r[1] or "").lower(),
                                "retsvirkning": r[2], "roller": _liste(r[3])})
    if "Andre former" in wb.sheetnames:
        for r in wb["Andre former"].iter_rows(min_row=2, values_only=True):
            if r[0]:
                o["andre_formers_navne"][r[0]] = [x.lower() for x in _liste(r[1])]
    return o
