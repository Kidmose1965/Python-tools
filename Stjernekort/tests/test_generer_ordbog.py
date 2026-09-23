"""Et genereret udkast skal give samme udtræk som den håndlavede udbudsordbog."""
from pathlib import Path

import pytest

from stjernekort.generer_ordbog import generer_udkast
from stjernekort.ordbog_excel import skriv_ordbog
from stjernekort.udtraek import laes_ordbog, udtraek

ROD = Path(__file__).resolve().parents[1]
ORDB = ROD / "stjernekort" / "ordboeger"
GRUND = sorted(ORDB.glob("ordbog_kontraktform_*.xlsx"))


@pytest.mark.parametrize("udbud,forventet_grund,forventet_navne", [
    ("esdh", "ordbog_kontraktform_service.xlsx", {"brugertest", "prøvekonvertering"}),
    ("dvr", "ordbog_kontraktform_k02.xlsx", {"poc gennemført", "verificeringsvindue afsluttes"}),
])
def test_udkast_svarer_til_haandlavet(tmp_path, udbud, forventet_grund, forventet_navne):
    docs = ROD / "testdata" / udbud
    if not any(docs.glob("*.docx")):
        pytest.skip(f"testdata/{udbud} mangler")
    u = generer_udkast(docs, GRUND)
    assert u["bygger_paa"] == forventet_grund
    syn = {s for m in u["milepaele"] for s in m["synonymer"]}
    assert forventet_navne <= syn
    sti = ORDB / f"_test_udkast_{udbud}.xlsx"            # skal ligge ved grundordbøgerne
    try:
        skriv_ordbog(u, sti)
        auto = udtraek(docs, laes_ordbog(sti))
    finally:
        sti.unlink(missing_ok=True)
    manuel = udtraek(docs, laes_ordbog(ORDB / f"ordbog_udbud_{udbud}.xlsx"))
    noegle = lambda bp: {(r.type, r.detalje, r.kilde.dokument, r.kilde.punkt) for r in bp.retsvirkninger}
    assert noegle(auto) == noegle(manuel)
