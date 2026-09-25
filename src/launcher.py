"""
Start_Extractor: lille launcher til Extractor.exe.

Formaal: kollegerne skal IKKE koere Extractor.exe direkte fra den delte
OneDrive-mappe, fordi OneDrive Files-on-Demand's cloud-placeholders og
synkroniseringstiming goer at det fejler intermitterende. I stedet:

1. Find Extractor.exe i samme mappe som denne launcher selv ligger i
   (dvs. OneDrive-mappen).
2. Kopier den til en lokal mappe (%LOCALAPPDATA%\Extractor\), men kun hvis
   kildefilen er nyere/anderledes end den lokale kopi (SHA256-sammenligning).
   Kopieringen sker via en midlertidig fil + atomisk omdoebning, saa en
   afbrudt OneDrive-hentning ikke efterlader en korrupt lokal exe.
3. Start den lokale kopi.
4. Hvis OneDrive lige nu er utilgaengelig, men der findes en lokal kopi i
   forvejen, bruges den (med en fejlbesked hvis der slet intet findes).
"""
import ctypes
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

APP_NAVN = "Extractor.exe"
LOKAL_MAPPE = Path(os.environ["LOCALAPPDATA"]) / "Extractor"
LOKAL_APP = LOKAL_MAPPE / APP_NAVN
FORSOEG = 3
FORSOEG_PAUSE_SEK = 2


def vis_fejl(besked: str) -> None:
    ctypes.windll.user32.MessageBoxW(0, besked, "Extractor - opstartsfejl", 0x10)


def launcher_mappe() -> Path:
    """Mappen launcheren selv koerer fra (dvs. OneDrive-mappen)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def sha256(sti: Path) -> str | None:
    try:
        h = hashlib.sha256()
        with open(sti, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return None


def er_opdateret(kilde: Path, maal: Path) -> bool:
    if not maal.exists():
        return False
    kilde_hash = sha256(kilde)
    maal_hash = sha256(maal)
    return kilde_hash is not None and kilde_hash == maal_hash


def kopier_atomisk(kilde: Path, maal: Path) -> bool:
    maal.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=maal.parent, delete=False, suffix=".tmp") as tmp:
        tmp_sti = Path(tmp.name)
    try:
        shutil.copy2(kilde, tmp_sti)
        if kilde.stat().st_size != tmp_sti.stat().st_size:
            raise OSError("Ufuldstaendig kopiering fra OneDrive (stoerrelse matcher ikke)")
        os.replace(tmp_sti, maal)
        return True
    except OSError:
        tmp_sti.unlink(missing_ok=True)
        return False


def sikr_lokal_kopi(kilde: Path, maal: Path) -> bool:
    """Returnerer True hvis maal-filen er klar til brug."""
    if not kilde.exists():
        # OneDrive utilgaengelig lige nu - brug evt. eksisterende lokal kopi
        return maal.exists()

    if er_opdateret(kilde, maal):
        return True

    for forsoeg_nr in range(1, FORSOEG + 1):
        if kopier_atomisk(kilde, maal):
            return True
        if forsoeg_nr < FORSOEG:
            time.sleep(FORSOEG_PAUSE_SEK)

    # Kopiering fejlede efter flere forsoeg - brug evt. gammel lokal kopi
    return maal.exists()


def main() -> int:
    kilde = launcher_mappe() / APP_NAVN
    ok = sikr_lokal_kopi(kilde, LOKAL_APP)

    if not ok or not LOKAL_APP.exists():
        vis_fejl(
            "Kunne ikke finde eller hente Extractor.\n\n"
            f"Forventet kildefil: {kilde}\n\n"
            "Tjek at OneDrive er logget ind og har synkroniseret mappen "
            "faerdig, og proev igen."
        )
        return 1

    subprocess.Popen([str(LOKAL_APP)], cwd=str(LOKAL_APP.parent))
    return 0


if __name__ == "__main__":
    sys.exit(main())
