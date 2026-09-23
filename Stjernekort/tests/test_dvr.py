"""Accepttest mod DVR (K02-kontraktform). Kræver kontrakt og bilag i testdata/dvr.
Kør: python -m pytest -q"""
from functools import lru_cache
from pathlib import Path

import pytest

from stjernekort.udtraek import laes_ordbog, udtraek

ROD = Path(__file__).resolve().parents[1]
DOCS = ROD / "testdata" / "dvr"
ORDBOG = ROD / "stjernekort" / "ordboeger" / "ordbog_udbud_dvr.xlsx"

pytestmark = pytest.mark.skipif(not any(DOCS.glob("*.docx")), reason="testdata/dvr mangler")


@lru_cache(maxsize=1)
def _bp():
    return udtraek(DOCS, laes_ordbog(ORDBOG))


def _rv():
    return {(r.milepael, r.type): r for r in _bp().retsvirkninger}


def test_lagdelt_ordbog():
    o = laes_ordbog(ORDBOG)
    ids = [m["id"] for m in o["milepaele"]]
    assert ids[:5] == ["M1", "M2", "U1", "U2", "U3"]        # udbuddets milepæle flettet ind
    assert any(m["rolle"] == "HOVEDPRØVE" and m["navn"].startswith("Overtagelsesprøve")
               for m in o["milepaele"])


def test_betalingsplan():
    rv = _rv()
    assert rv[("M2", "betaling")].detalje == "15 %"
    assert rv[("U3", "betaling")].detalje == "10 %"
    assert rv[("M3", "betaling")].detalje == "50 %"
    assert rv[("M4", "betaling")].detalje == "25 %"


def test_bod_og_ophaevelse():
    b = [r for r in _bp().retsvirkninger if r.type == "bod" and r.kilde.punkt == "17.2"]
    assert {r.milepael for r in b} == {"M3", "M4"}
    assert all("2.000 kr./AD, maks. 100.000 kr." == r.detalje for r in b)
    rv = _rv()
    assert "50 Arbejdsdage" in rv[("M3", "ophævelse")].detalje
    assert "50 Arbejdsdage" in rv[("M4", "ophævelse")].detalje


def test_overtagelse_og_udtraedelse():
    rv = _rv()
    assert ("M3", "overtagelse") in rv and ("M3", "ibrugtagning") in rv
    assert rv[("M2", "udtrædelsesadgang")].kilde.punkt == "5.3.2"


def test_kendte_fejl_i_henvisninger():
    brudte = {(str(h.fra), h.til_dokument, h.til_punkt) for h in _bp().henvisninger if h.fundet is False}
    assert ("Bilag 5 pkt. 1", "Kontrakt", "24.2") in brudte      # skal være 23.2
    assert ("Bilag 12 pkt. 5.3", "Bilag 5", "8.4") in brudte     # punkt 8.4 findes ikke


def test_fremmed_begreb_fundet():
    assert any(a[1] == "Bilag 14 pkt. 3.2" for a in _bp().advarsler)   # "brugertests"
