"""Kommandolinje:
  python -m stjernekort.cli udtraek  <mappe med docx> <ordbog.json> <ud.xlsx>
  python -m stjernekort.cli generer  <blueprint.xlsx> <template.pptx> <ud.pptx>
  python -m stjernekort.cli ny-ordbog <mappe med docx> <ud: ordbog_udbud_navn.xlsx>
"""
import sys

from . import excel_io
from .pptx_generator import generer
from .udtraek import laes_ordbog, udtraek


def main(argv=None):
    a = argv or sys.argv[1:]
    if not a:
        print(__doc__)
        return 1
    if a[0] == "udtraek":
        bp = udtraek(a[1], laes_ordbog(a[2]))
        print(f"{len(bp.milepaele)} milepæle, {len(bp.retsvirkninger)} retsvirkninger, "
              f"{len(bp.henvisninger)} henvisninger -> {excel_io.skriv(bp, a[3])}")
    elif a[0] == "generer":
        print("Skrevet:", generer(excel_io.laes(a[1]), a[2], a[3]))
    elif a[0] == "ny-ordbog":
        from pathlib import Path
        from .generer_ordbog import generer_udkast
        from .ordbog_excel import skriv_ordbog
        mappe_ordb = Path(__file__).parent / "ordboeger"
        grund = sorted(mappe_ordb.glob("ordbog_kontraktform_*.xlsx"))
        u = generer_udkast(a[1], grund)
        ud = Path(a[2])
        if ud.parent.resolve() != mappe_ordb.resolve():
            print("Bemærk: læg ordbogen i", mappe_ordb, "så 'Bygger på' kan findes.")
        skriv_ordbog(u, ud)
        print(f"Udkast bygger på {u['bygger_paa']} og tilføjer {len(u['milepaele'])} milepæle -> {ud}")
    else:
        print(__doc__)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
