"""Accepttest mod ESDH-paradigmet (testdata/esdh). Kør: python -m pytest -q"""
from functools import lru_cache
from pathlib import Path

from stjernekort.docx_reader import laes_mappe, punkter
from stjernekort.udtraek import laes_ordbog, udtraek

ROD = Path(__file__).resolve().parents[1]
DOCS = ROD / "testdata" / "esdh"
ORDBOG = laes_ordbog(ROD / "stjernekort" / "ordbog_esdh.json")


@lru_cache(maxsize=1)
def _bp():
    return udtraek(DOCS, ORDBOG)


def _rv(bp):
    return {(r.milepael, r.type): r for r in bp.retsvirkninger}


def test_punktnumre_beregnes():
    k = punkter(laes_mappe(DOCS)["Kontrakt"])
    assert k["6.3"] == "Ret til udtrædelse"
    assert k["33.2"] == "Bod ved forsinkelse"
    assert k["34.6"] == "Kundens ret til ophævelse"
    assert k["49.2"] == "Exit-plan"


def test_betalingsplan():
    rv = _rv(_bp())
    assert rv[("M2", "betaling")].detalje == "15 %"
    assert rv[("M5", "betaling")].detalje == "60 %"
    assert rv[("M6", "betaling")].detalje == "25 %"


def test_bod_og_ophaevelse_ved_proever():
    rv = _rv(_bp())
    for m in ("M5", "M6"):
        assert ("2.500 kr." in rv[(m, "bod")].detalje)
        assert rv[(m, "bod")].kilde.punkt == "33.2"
        assert (m, "ophævelse") in rv


def test_oevrige_retsvirkninger():
    rv = _rv(_bp())
    assert rv[("M2", "udtrædelsesadgang")].kilde.punkt == "6.3"
    assert rv[("M5", "overtagelse")].kilde.punkt == "11"
    assert ("M5", "ibrugtagning") in rv


def test_ingen_brudte_henvisninger():
    bp = _bp()
    assert len(bp.henvisninger) >= 30
    assert [h for h in bp.henvisninger if h.fundet is False] == []
