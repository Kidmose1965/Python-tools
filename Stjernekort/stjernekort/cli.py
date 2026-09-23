"""Kommandolinje:
  python -m stjernekort.cli udtraek  <mappe med docx> <ordbog.json> <ud.xlsx>
  python -m stjernekort.cli generer  <blueprint.xlsx> <template.pptx> <ud.pptx>
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
    else:
        print(__doc__)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
