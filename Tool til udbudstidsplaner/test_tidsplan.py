#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automatiserede tests for tidsplan.py (dato-/varighedsparsing, farvesemantik,
indlaesning og aksebygning). Koeres med:  pytest test_tidsplan.py -v
(eller:  python -m unittest test_tidsplan -v)
Kraever kun det, der allerede er installeret (openpyxl + stdlib unittest,
evt. pytest som testrunner - se requirements.txt).
"""
import datetime as dt
import io
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color

import tidsplan as tp

REFERENCEFIL = Path(__file__).resolve().parent / "Testmateriale og baggrund" / "Excel udtræk tidsplan.xlsx"


def _workbook(rows):
    """Bygger et Excel-ark i hukommelsen der efterligner MS Project-udtraekket
    (header i raekke 2, data fra raekke 3). Hver row er en dict med noeglerne:
    navn, indrykket (bool), fed (bool), varighed, start, slut, fill (hex, valgfri),
    tema_farve (Color, valgfri - bruges til at simulere tema-/indekserede farver)."""
    wb = Workbook()
    ws = wb.active
    ws.cell(2, 2, "Aktivitet"); ws.cell(2, 3, "Længde"); ws.cell(2, 4, "Start"); ws.cell(2, 5, "Slut")
    for i, row in enumerate(rows, start=3):
        navn = ("  " + row["navn"]) if row.get("indrykket") else row["navn"]
        c = ws.cell(i, 2, navn)
        if row.get("fed"):
            c.font = Font(bold=True)
        if row.get("fill"):
            c.fill = PatternFill(start_color=row["fill"], end_color=row["fill"], fill_type="solid")
        if row.get("tema_farve") is not None:
            c.fill = PatternFill(fill_type="solid", fgColor=row["tema_farve"])
        ws.cell(i, 3, row.get("varighed", "5 dage"))
        ws.cell(i, 4, row.get("start", ""))
        ws.cell(i, 5, row.get("slut", ""))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class TestParseDato(unittest.TestCase):
    def test_date(self):
        self.assertEqual(tp.parse_dato(dt.date(2026, 9, 10)), dt.date(2026, 9, 10))

    def test_datetime(self):
        self.assertEqual(tp.parse_dato(dt.datetime(2026, 9, 10, 13, 30)), dt.date(2026, 9, 10))

    def test_iso_streng(self):
        self.assertEqual(tp.parse_dato("2026-09-10"), dt.date(2026, 9, 10))

    def test_dansk_streng(self):
        self.assertEqual(tp.parse_dato("10-09-2026"), dt.date(2026, 9, 10))

    def test_dansk_streng_2cifret_aar(self):
        self.assertEqual(tp.parse_dato("10-09-26"), dt.date(2026, 9, 10))

    def test_ms_project_ugedag_praefiks(self):
        self.assertEqual(tp.parse_dato("Tue 07-10-25"), dt.date(2025, 10, 7))

    def test_ms_project_ugedag_praefiks_4cifret_aar(self):
        self.assertEqual(tp.parse_dato("Mon 15-09-2026"), dt.date(2026, 9, 15))

    def test_ugyldig_tekst(self):
        with self.assertRaises(ValueError):
            tp.parse_dato("ikke en dato")

    def test_ugyldig_kalenderdato(self):
        with self.assertRaises(ValueError):
            tp.parse_dato("31-02-2026")   # 31. februar findes ikke

    def test_none(self):
        with self.assertRaises(ValueError):
            tp.parse_dato(None)


class TestParseVarighed(unittest.TestCase):
    def test_0_days(self):
        self.assertEqual(tp.parse_varighed("0 days"), 0)

    def test_1_day(self):
        self.assertEqual(tp.parse_varighed("1 day"), 1)

    def test_5_days(self):
        self.assertEqual(tp.parse_varighed("5 days"), 5)

    def test_0_dage(self):
        self.assertEqual(tp.parse_varighed("0 dage"), 0)

    def test_1_dag(self):
        self.assertEqual(tp.parse_varighed("1 dag"), 1)

    def test_5_dage(self):
        self.assertEqual(tp.parse_varighed("5 dage"), 5)

    def test_numerisk_int(self):
        self.assertEqual(tp.parse_varighed(5), 5)

    def test_numerisk_float(self):
        self.assertEqual(tp.parse_varighed(5.0), 5)

    def test_ugyldig_varighed(self):
        with self.assertRaises(ValueError):
            tp.parse_varighed("ukendt")

    def test_ugyldig_varighed_none(self):
        with self.assertRaises(ValueError):
            tp.parse_varighed(None)


class TestFarveSemantik(unittest.TestCase):
    """Farven i kolonne B (Task Name) er semantisk - se laes_plan()/_celle_farve()."""

    def _farve(self, fill=None, tema_farve=None, varighed="5 dage"):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed=varighed,
                     start="01-09-2026", slut="05-09-2026", fill=fill, tema_farve=tema_farve)]
        lanes, advarsler = tp.laes_plan(_workbook(rows))
        return lanes[0]["tasks"][0], advarsler

    def test_rod_spidsbelastning(self):
        task, _ = self._farve(fill="FF0000")
        self.assertEqual(task[3], tp.RED)

    def test_orange_belastning(self):
        task, _ = self._farve(fill="FFC000")
        self.assertEqual(task[3], tp.ORANGE)

    def test_groen_ekstern(self):
        task, _ = self._farve(fill="92D050")
        self.assertEqual(task[3], tp.GREEN)

    def test_lilla_godkendelse(self):
        task, _ = self._farve(fill="7030A0")
        self.assertEqual(task[3], tp.PURPLE)

    def test_blaa_officiel_milepael(self):
        task, _ = self._farve(fill="00B0F0", varighed="0 dage")
        self.assertIsNone(task[2])
        self.assertEqual(task[3], tp.MS_BLUE)

    def test_orange_moede(self):
        task, _ = self._farve(fill="ED7D31")
        self.assertEqual(task[3], tp.MOEDE)

    def test_hvid_neutral(self):
        task, _ = self._farve(fill="FFFFFF")
        self.assertEqual(task[3], tp.NEUTRAL)

    def test_ukendt_farve_neutral(self):
        task, _ = self._farve(fill=None)
        self.assertEqual(task[3], tp.NEUTRAL)

    def test_tema_farve_neutral_med_advarsel(self):
        task, advarsler = self._farve(tema_farve=Color(theme=4, tint=0.0))
        self.assertEqual(task[3], tp.NEUTRAL)
        self.assertTrue(advarsler)
        self.assertTrue(any("tema" in a.lower() for a in advarsler))

    def test_indekseret_farve_neutral_med_advarsel(self):
        task, advarsler = self._farve(tema_farve=Color(indexed=22))
        self.assertEqual(task[3], tp.NEUTRAL)
        self.assertTrue(advarsler)


class TestMoeder(unittest.TestCase):
    def test_moede_markoer_ved_0_dage(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="0 dage",
                     start="01-09-2026", slut="01-09-2026", fill="ED7D31")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        task = lanes[0]["tasks"][0]
        self.assertIsNone(task[2])
        self.assertEqual(task[3], tp.MOEDE)
        prims, _ = tp.byg(lanes, dt.date(2026, 9, 1))
        self.assertTrue(any(p["t"] == "diamond" and p["fill"] == tp.MOEDE for p in prims))

    def test_moede_bjaelke_ved_over_0_dage(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                     start="01-09-2026", slut="05-09-2026", fill="ED7D31")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        task = lanes[0]["tasks"][0]
        self.assertIsNotNone(task[2])
        self.assertEqual(task[3], tp.MOEDE)
        prims, _ = tp.byg(lanes, dt.date(2026, 9, 1))
        self.assertTrue(any(p["t"] == "rect" and p["fill"] == tp.MOEDE for p in prims))

    def test_moede_paavirker_ikke_belastningsbaand(self):
        # H_LOAD er den faste hoejde paa belastningsbaandets ugerektangler i byg() -
        # en moede-farvet rect med den hoejde ville betyde at moedet fejlagtigt har
        # faerget belastningsbaandet.
        H_LOAD = 0.13
        rows = [
            dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                 start="01-09-2026", slut="05-09-2026"),
            dict(navn="Møde", indrykket=True, varighed="5 dage",
                 start="01-09-2026", slut="05-09-2026", fill="ED7D31"),
        ]
        lanes, _ = tp.laes_plan(_workbook(rows))
        prims, _ = tp.byg(lanes, dt.date(2026, 9, 1))
        belastningsbaand_moede = [p for p in prims if p["t"] == "rect" and p["fill"] == tp.MOEDE
                                  and abs(p["h"] - H_LOAD) < 1e-6]
        self.assertEqual(belastningsbaand_moede, [])

    def test_moede_ikke_i_niveau(self):
        self.assertNotIn(tp.MOEDE, tp.NIVEAU)


class TestMilepaele(unittest.TestCase):
    def test_milepael_praefiks_0_dage_blaa(self):
        rows = [dict(navn="Milepæl: Aflevering", fed=True, indrykket=False, varighed="0 dage",
                     start="10-09-2026", slut="10-09-2026")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        task = lanes[0]["tasks"][0]
        self.assertEqual(task[0], "Aflevering")   # praefiks fjernet
        self.assertIsNone(task[2])
        self.assertEqual(task[3], tp.MS_BLUE)

    def test_00B0F0_0_dage_blaa(self):
        rows = [dict(navn="Aflevering", fed=True, indrykket=False, varighed="0 dage",
                     start="10-09-2026", slut="10-09-2026", fill="00B0F0")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        task = lanes[0]["tasks"][0]
        self.assertEqual(task[3], tp.MS_BLUE)

    def test_milepael_praefiks_over_0_dage_ikke_blaa(self):
        rows = [dict(navn="Milepæl: Godkendelsesperiode", fed=True, indrykket=False, varighed="5 dage",
                     start="01-09-2026", slut="05-09-2026")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        task = lanes[0]["tasks"][0]
        self.assertEqual(task[0], "Godkendelsesperiode")  # praefiks fjernet
        self.assertIsNotNone(task[2])                     # almindelig bjaelke, ikke markoer
        self.assertNotEqual(task[3], tp.MS_BLUE)

    def test_almindelig_0_dages_graa(self):
        rows = [dict(navn="Modtagelse", fed=True, indrykket=False, varighed="0 dage",
                     start="10-09-2026", slut="10-09-2026")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        task = lanes[0]["tasks"][0]
        self.assertIsNone(task[2])
        self.assertEqual(task[3], tp.MS_GREY)


class TestLaesPlanOgByg(unittest.TestCase):
    def _normale_raekker(self):
        return [
            dict(navn="Fase 1", fed=True, indrykket=False, varighed="0 dage",
                 start="01-09-2026", slut="01-09-2026"),
            dict(navn="Aktivitet A", indrykket=True, varighed="5 dage",
                 start="01-09-2026", slut="05-09-2026"),
            dict(navn="Milepæl: Aflevering", indrykket=True, varighed="0 dage",
                 start="10-09-2026", slut="10-09-2026"),
        ]

    def test_normal_aktiv_tidsplan(self):
        lanes, advarsler = tp.laes_plan(_workbook(self._normale_raekker()))
        self.assertEqual(len(lanes), 1)
        self.assertEqual(len(lanes[0]["tasks"]), 2)
        self.assertEqual(advarsler, [])
        prims, byg_advarsler = tp.byg(lanes, dt.date(2026, 9, 3))
        self.assertTrue(any(p["t"] == "rect" for p in prims))
        self.assertTrue(any(p["t"] == "diamond" for p in prims))

    def test_statusdato_efter_plan_afslutning(self):
        lanes, _ = tp.laes_plan(_workbook(self._normale_raekker()))
        # Planens sidste dato er 10-09-2026; statusdato er sat langt efter.
        prims, advarsler = tp.byg(lanes, dt.date(2027, 6, 1))
        self.assertTrue(len(prims) > 0)

    def test_byg_uden_aktiviteter_fejler_klart(self):
        with self.assertRaises(ValueError):
            tp.byg([], dt.date(2026, 9, 10))

    def test_manglende_varighed(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="",
                     start="01-09-2026", slut="05-09-2026")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("række 3", str(ctx.exception).lower())

    def test_ugyldig_varighed_i_plan(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="ukendt",
                     start="01-09-2026", slut="05-09-2026")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("række 3", str(ctx.exception).lower())

    def test_manglende_startdato(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                     start="", slut="05-09-2026")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("række 3", str(ctx.exception).lower())

    def test_ugyldig_startdato(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                     start="ikke en dato", slut="05-09-2026")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("række 3", str(ctx.exception).lower())

    def test_manglende_slutdato_almindelig_aktivitet(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                     start="01-09-2026", slut="")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("række 3", str(ctx.exception).lower())

    def test_slutdato_foer_startdato(self):
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                     start="10-09-2026", slut="01-09-2026")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("slutdato", str(ctx.exception).lower())

    def test_indrykket_uden_fase(self):
        rows = [dict(navn="Aktivitet uden fase", indrykket=True, varighed="5 dage",
                     start="01-09-2026", slut="05-09-2026")]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("række 3", str(ctx.exception).lower())

    def test_milepael_uden_gyldig_slutdato_fejler_ikke(self):
        # Slutdato er uden betydning for en milepael og skal ikke blokere indlaesningen.
        rows = [dict(navn="Fase 1", fed=True, indrykket=False, varighed="0 dage",
                     start="01-09-2026", slut="")]
        lanes, _ = tp.laes_plan(_workbook(rows))
        self.assertEqual(len(lanes), 1)

    def test_blank_raekke_ignoreres(self):
        rows = [
            dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                 start="01-09-2026", slut="05-09-2026"),
            dict(navn="", varighed="", start="", slut=""),
            dict(navn="Aktivitet A", indrykket=True, varighed="5 dage",
                 start="01-09-2026", slut="05-09-2026"),
        ]
        lanes, _ = tp.laes_plan(_workbook(rows))
        self.assertEqual(len(lanes[0]["tasks"]), 1)

    def test_manglende_task_name_med_data_fejler(self):
        rows = [
            dict(navn="Fase 1", fed=True, indrykket=False, varighed="5 dage",
                 start="01-09-2026", slut="05-09-2026"),
            dict(navn="", varighed="5 dage", start="01-09-2026", slut="05-09-2026"),
        ]
        with self.assertRaises(ValueError) as ctx:
            tp.laes_plan(_workbook(rows))
        self.assertIn("task name", str(ctx.exception).lower())


class TestReferenceFil(unittest.TestCase):
    """End-to-end paa det autoritative modeludtraek fra MS Project."""

    def setUp(self):
        if not REFERENCEFIL.exists():
            self.skipTest(f"Referencefilen findes ikke: {REFERENCEFIL}")

    def test_laeser_75_aktiviteter_i_6_faser(self):
        lanes, advarsler = tp.laes_plan(str(REFERENCEFIL))
        self.assertEqual(len(lanes), 6)
        self.assertEqual(sum(len(l["tasks"]) for l in lanes), 75)

    def test_generering_fejler_ikke_og_legend_har_moede(self):
        konfig = tp.Konfig(kunde="Modelkunde", udbud="Udbud med forhandling - 1 runde")
        lanes, laes_adv = tp.laes_plan(str(REFERENCEFIL), konfig)
        prims, byg_adv = tp.byg(lanes, dt.date(2026, 9, 17), konfig)
        self.assertTrue(len(prims) > 0)
        legend_tekster = [p["s"] for p in prims if p["t"] == "text" and isinstance(p["s"], str)]
        self.assertTrue(any("Møde" == t or "Møde" in t for t in legend_tekster))


if __name__ == "__main__":
    unittest.main()
