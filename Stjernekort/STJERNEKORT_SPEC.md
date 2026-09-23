# Stjernekort-værktøj (kontraktens blueprint) – behovsnotat og kodeforslag til Claude Code

*23. september 2026 · Birger · Rambøll*

Dette notat er både behovsnotat og byggevejledning. Læg det i projektmappen som `STJERNEKORT_SPEC.md`, og bed Claude Code læse det først (se afsnit 8). Kodeforslagene i afsnit 7 er afprøvet mod ESDH-paradigmet: udtrækket finder alle retsvirkninger i tabellen i afsnit 3, de fem accepttests består, og den genererede PowerPoint validerer og ser ud som figuren i afsnit 6.

---

## 1. Formål

Værktøjet laver kontraktens blueprint (stjernekort) i Rambølls blueprint-template ud fra et Excel-ark, og det kan udlede blueprintet af et kontraktudkast med bilag.

Blueprintet bruges to gange i et projekt:

1. **Som dialogværktøj med kunden.** I møderne arbejdes der på whiteboard; værktøjet bruges mellem møderne til at lave en ny version. Forløbet er iterativt.
2. **Som endeligt kontraktbilag.** Når kontrakt og bilag er skrevet, udtrækker værktøjet milepæle, retsvirkninger og henvisninger fra teksten. Det endelige blueprint bygger på det, der faktisk står i kontrakten, og erstatter den forenklede forløbsfigur (Fig. 1) i bilag 2, punkt 1.

```
Paradigme ──► Excel (master) ◄── Whiteboard i mødet
                  │
                  ▼
        Blueprint v1, v2 … (PowerPoint)

Kontraktudkast + bilag ──► Udtræk ──► Afvigelser mod dialogversion
                                  └──► Endeligt blueprint (bilag 2)
```

## 2. Beslutninger og antagelser

| Emne | Beslutning |
| --- | --- |
| Master | Excel-arket er altid master. PowerPoint genereres altid på ny; manuelle rettelser i PowerPoint er sidste finish på den endelige version. |
| Udtræk | Regelbaseret i Python, styret af en ordbog pr. paradigme. Intet AI-trin i første version. |
| Paradigmer | Rambølls egne. ESDH-paradigmet er det første og er testgrundlag. |
| Template | `Kontraktens_blueprint_template.pptx`, dias 2 (symboler i banerne). |
| Tidsakse | Antaget: milepæle i rækkefølge, jævnt fordelt (som i templaten). Åben beslutning. |
| Miljø | Windows, Python i `C:\Python`, Claude Code i terminalen, git-repo `Python-tools`. Biblioteker: `python-docx`, `python-pptx`, `openpyxl`, `pytest`. |

## 3. Hvad ESDH-paradigmet indeholder

Tabellen er den forventede facitliste for udtrækket. "Punkt" uden bilag er kontraktens punkt; AD = Arbejdsdage.

| # | Milepæl | Retsvirkninger | Hjemmel |
| --- | --- | --- | --- |
| 1 | Kontraktindgåelse | Kontrakten træder i kraft; afklaringsfasen starter | Pkt. 48.1; bilag 2 pkt. 2 |
| 2 | Kontraktworkshop (aktivitet, efter Kundens ønske) | Ingen | Pkt. 6.2 |
| 3 | Godkendelse af afklaringsfasen | Betaling 15 %; udtrædelsesadgang indtil 20 AD efter afvisning, mod udtrædelsesvederlag | Pkt. 6.1, 6.3; bilag 8 pkt. 2-3 |
| 4 | Brugertest (afslutter etableringsfase del 1) | Ingen – ikke en formel prøve | Bilag 2 pkt. 1 og 6; bilag 4 pkt. 3 |
| 5 | 1. og 2. prøvekonvertering | Forudsætning for Ibrugtagningsprøven | Bilag 4 pkt. 4.2.4 |
| 6 | Brugerdokumentation godkendt | Frist: senest 10 Dage før Ibrugtagningsprøven | Pkt. 7 |
| 7 | Godkendelse af Ibrugtagningsprøven (= Ibrugtagningsdag) | Overtagelse; ibrugtagning; betaling 60 %; bod 2.500 kr./AD, maks. 100.000 kr.; ophævelse ved forsinkelse over 40 AD; fristen kan ikke udskydes | Pkt. 9.1, 10.1, 11, 33.1-33.3; bilag 8 pkt. 3 |
| 8 | Driftsfasen starter | Servicevederlag månedligt bagud; servicemål og bod efter bilag 6; ophævelse ved driftseffektivitet ≤ 90 % over 3 mdr. eller bod 4 måneder i træk | Pkt. 30.3, 34.6; bilag 6 pkt. 1 og 7 |
| 9 | Endelig konvertering | I forbindelse med Driftsprøven | Bilag 4 pkt. 4.2.4 |
| 10 | Godkendelse af Driftsprøven (mindst 20 AD i træk) | Betaling 25 %; bod 2.500 kr./AD, maks. 100.000 kr.; ophævelse ved forsinkelse over 40 AD | Pkt. 33.1-33.3; bilag 4 pkt. 5; bilag 8 pkt. 3 |
| 11 | Exit-plan | Frist: senest 3 mdr. efter Ibrugtagningsdag | Pkt. 49.2 |
| 12 | Tidligste opsigelse | 6 mdr. varsel, tidligst til 24 mdr. efter Ibrugtagningsdag | Pkt. 48.2 |
| 13 | Option: aflevering til Rigsarkivet (2033) | Option; pris efter bilag 8 | Pkt. 8; bilag 1A pkt. 3; bilag 8 pkt. 6 |
| 14 | Kontraktophør efter 8 år | Forlængelse 2 × 1 år; overflytning; fortsat varetagelse i op til 6 mdr. | Pkt. 48.1, 49.3, 49.4 |

Templatens "Overtagelsesprøve" hedder "Ibrugtagningsprøve" i dette paradigme. Symbolerne Bonus og Delleveranceprøve bruges ikke her.

## 4. Observationer der styrer udtrækket

1. **Punktnumre står ikke i teksten.** Overskrifterne er autonummererede. Numrene beregnes ud fra Heading 1-3 og stemmer med indholdsfortegnelserne i alle ni dokumenter.
2. **Samme milepæl har flere navne.** Afklaringsfasens afslutning hedder "Godkendelse af afklaringsfasen" (bilag 8), "Afslutning af afklaringsfasen" (bilag 2), "Godkendt Ydelses- og Servicebeskrivelse" (bilag 2, Fig. 1) og godkendelse af "revideret Leverancebeskrivelse" (pkt. 6.1). Derfor en ordbog med kanonisk navn og synonymer.
3. **Retsvirkninger står mest i tabeller.** Bodstabellen (pkt. 33.2), betalingsplanen (bilag 8 pkt. 3) og hovedtidsplanen (bilag 2 pkt. 2) giver høj sikkerhed. Resten findes via faste vendinger ("berettiget til at hæve", "ret til at udtræde", "overtaget af Kunden") i samme sætning som en milepæl – middel sikkerhed.
4. **Forsinkelsesdefinitionen** (pkt. 33.1) lister de milepæle, der kan udløse ophævelse, som punktopstilling efter "Forsinkelse foreligger".
5. **Henvisninger skal læses med kontekst.** "bilag 6, punkt 4.2" og "Kontraktens punkt 6.3" er én henvisning til et andet dokument; "punkt 14.1 i udbudsbetingelserne" er ekstern. Alle 38 interne henvisninger i paradigmet kan slås op.
6. **Datoer udfyldes af tilbudsgiver**, så udbudsversionen af blueprintet er relativ.
7. **Vejledninger skal springes over.** Alt før første Heading 1 (vejledning, indholdsfortegnelse, bilagsliste) er forside.
8. **Filnavne bærer bilagsnummeret** (`11._Bilag_02_…docx` → Bilag 2). Filer uden "Bilag_NN" tolkes som kontrakten.

**QA-fund i paradigmet og templaten:** Overskriften til pkt. 6 hedder "Aflaringsfase og udtrædelse". Fig. 1 i bilag 2 kalder milepælen "Godkendt Ydelses- og Servicebeskrivelse", mens pkt. 6.1 taler om revideret Leverancebeskrivelse. I templatens dias 2 står den blå stjerne i signaturforklaringen som "Betaling"; i dias 3 hedder den "Bonus".

## 5. Datamodel

Samme model bruges til paradigme, dialogversioner og udtræk. Excel-arket har disse faner:

| Fane | Én række pr. | Kolonner |
| --- | --- | --- |
| Milepæle | milepæl eller frist | ID, Navn, Etiket (kort tekst på blueprintet), Type, Rækkefølge, Option, Relativ tid, Dato, Kilder |
| Aktiviteter | pilefigur | Bane, Tekst, Fra, Til, Farvekategori, Række |
| Retsvirkninger | symbol | Milepæl, Type, Detalje, Hjemmel, Citat, Oprindelse (inkl. sikkerhed) |
| Henvisninger | henvisning | Fra, Til dokument, Til punkt, Fundet |

Ordbogen (`ordbog_esdh.json`) ligger ved siden af koden: milepæle med synonymer, nøgleord pr. retsvirkning og kendetegn for tabellerne.

## 6. Arkitektur og resultat

```
C:\Python\Stjernekort\
  STJERNEKORT_SPEC.md          <- dette notat
  stjernekort\
    __init__.py
    model.py                   datamodel
    docx_reader.py             .docx -> blokke med beregnede punktnumre
    udtraek.py                 milepæle, retsvirkninger, henvisninger
    excel_io.py                Excel er master (skriv/læs)
    pptx_generator.py          tegner i blueprint-templaten ("stempler" dens figurer)
    cli.py                     kommandolinje
    ordbog_esdh.json           ordbog for ESDH-paradigmet
  tests\test_esdh.py           accepttest mod paradigmet
  testdata\esdh\*.docx         paradigmet (kontrakt + bilag)
  testdata\Kontraktens_blueprint_template.pptx
```

Kørsel:

```
python -m stjernekort.cli udtraek  testdata\esdh stjernekort\ordbog_esdh.json ud\esdh.xlsx
python -m stjernekort.cli generer  ud\esdh.xlsx testdata\Kontraktens_blueprint_template.pptx ud\esdh.pptx
python -m pytest -q
```

Resultat på ESDH-paradigmet: 11 milepæle, 11 retsvirkninger (alle i afsnit 3's tabel for M2, M5 og M6), 38 henvisninger, 0 brudte. Aktiviteterne blev lagt ind i Excel (de kommer fra dialogen), og PowerPoint blev genereret med templatens egne figurer, farver og symboler.

**Generatorens princip:** Templaten har en palette over diaset (pilefigurer i farvekategorier), milepælsgrupper (linje + Mx-boks + navneboks) og symboler i signaturforklaringen. Generatoren kloner disse figurer ("stempler") og placerer dem. Derfor ser resultatet ud som det, I er vant til, og kan finjusteres i hånden. Figurerne findes i dag via shape-id; se forbedring 1 i afsnit 8.

## 7. Kodeforslag (afprøvet)

### 7.1 Datamodel – `stjernekort/model.py`

```python
"""Datamodel for kontraktens blueprint (stjernekort).

Samme model bruges til paradigme, dialogversioner og udtræk, så en
sammenligning mellem dem blot er en forskel mellem to Blueprint-objekter.
"""
from __future__ import annotations

from dataclasses import dataclass, field

BANER = ["Etableringsprojekt", "Løbende ydelser", "Optioner", "Governance", "Kundens ydelser"]
RETSVIRKNINGER = ["betaling", "bod", "ophævelse", "overtagelse", "ibrugtagning",
                  "udtrædelsesadgang", "bonus"]
MILEPAELSTYPER = ["kontraktindgåelse", "faseafslutning", "prøve", "frist", "ophør"]


@dataclass
class Kilde:
    dokument: str        # "Kontrakt", "Bilag 8" ...
    punkt: str           # "33.2"
    citat: str = ""      # kort uddrag af teksten, til QA

    def __str__(self) -> str:
        return f"{self.dokument} pkt. {self.punkt}" if self.punkt else self.dokument


@dataclass
class Milepael:
    id: str                          # "M1" ...
    navn: str                        # kanonisk navn fra ordbogen
    type: str                        # se MILEPAELSTYPER
    raekkefoelge: int
    option: bool = False
    etiket: str = ""                 # kort tekst på blueprintet, fx "Kontrakt-indgåelse"
    relativ_tid: str = ""            # fx "Ibrugtagningsdag + 3 mdr."
    dato: str = ""                   # udfyldes efter kontraktindgåelse
    kilder: list[Kilde] = field(default_factory=list)


@dataclass
class Aktivitet:
    bane: str                        # se BANER
    tekst: str
    fra: str                         # milepæl-id
    til: str                         # milepæl-id
    farvekategori: str = "fase1"     # nøgle i template-konfigurationen
    raekke: int = 0                  # lodret placering i banen


@dataclass
class Retsvirkning:
    milepael: str                    # milepæl-id
    type: str                        # se RETSVIRKNINGER
    detalje: str = ""                # "60 %", "2.500 kr./AD, maks. 100.000 kr."
    kilde: Kilde | None = None
    oprindelse: str = "udtræk"       # "paradigme" | "dialog" | "udtræk"


@dataclass
class Henvisning:
    fra: Kilde
    til_dokument: str
    til_punkt: str
    fundet: bool


@dataclass
class Blueprint:
    navn: str
    milepaele: list[Milepael] = field(default_factory=list)
    aktiviteter: list[Aktivitet] = field(default_factory=list)
    retsvirkninger: list[Retsvirkning] = field(default_factory=list)
    henvisninger: list[Henvisning] = field(default_factory=list)

    def milepael(self, mid: str) -> Milepael | None:
        return next((m for m in self.milepaele if m.id == mid), None)
```

### 7.2 Læsning af .docx med beregnede punktnumre – `stjernekort/docx_reader.py`

```python
"""Læser .docx til en flad liste af blokke med beregnede punktnumre.

Overskrifterne i paradigmerne er autonummererede, så "33.2" står ikke i
teksten. Numrene beregnes ud fra overskriftsniveauerne (Heading 1-3).
Alt før første Heading 1 (vejledning, indholdsfortegnelse, bilagsliste)
markeres som forside og indgår ikke i udtrækket af retsvirkninger.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

HEADING_RE = re.compile(r"^Heading (\d)$")
BILAG_ID_RE = re.compile(r"Bilag[_ ]0*(\d+[A-Z]?)", re.IGNORECASE)


@dataclass
class Blok:
    dokument: str            # "Kontrakt", "Bilag 2", "Bilag 1A" ...
    punkt: str               # beregnet punktnummer, fx "33.2" ("" på forsiden)
    overskrift: str          # nærmeste overskrift
    type: str                # "overskrift" | "afsnit" | "tabel"
    tekst: str = ""
    stil: str = ""           # Word-typografi, fx "List Bullet"
    raekker: list[list[str]] = field(default_factory=list)  # kun for tabeller
    forside: bool = False    # True før første Heading 1


def dokument_id(filnavn: str) -> str:
    """'11._Bilag_02_Hovedtidsplan.docx' -> 'Bilag 2'; kontrakten -> 'Kontrakt'."""
    m = BILAG_ID_RE.search(filnavn)
    return f"Bilag {m.group(1).upper()}" if m else "Kontrakt"


def _blokke_i_body(d):
    for el in d.element.body.iterchildren():
        if el.tag.endswith("}p"):
            yield Paragraph(el, d)
        elif el.tag.endswith("}tbl"):
            yield Table(el, d)


def _tabel_raekker(t: Table) -> list[list[str]]:
    raekker = []
    for r in t.rows:
        celler: list[str] = []
        for c in r.cells:
            x = c.text.strip().replace("\n", " / ")
            if not celler or celler[-1] != x:   # flettede celler gentages
                celler.append(x)
        raekker.append(celler)
    return raekker


def laes_docx(sti: str | Path) -> list[Blok]:
    sti = Path(sti)
    dok = dokument_id(sti.name)
    d = docx.Document(str(sti))
    tael = [0, 0, 0]
    punkt, overskrift, forside = "", "", True
    blokke: list[Blok] = []
    for b in _blokke_i_body(d):
        if isinstance(b, Paragraph):
            tekst = b.text.strip()
            if not tekst:
                continue
            m = HEADING_RE.match(b.style.name or "")
            if m and int(m.group(1)) <= 3:
                niv = int(m.group(1)) - 1
                tael[niv] += 1
                for i in range(niv + 1, 3):
                    tael[i] = 0
                punkt = ".".join(str(x) for x in tael[: niv + 1])
                overskrift, forside = tekst, False
                blokke.append(Blok(dok, punkt, overskrift, "overskrift", tekst))
            elif (b.style.name or "").lower().startswith("toc"):
                continue
            else:
                blokke.append(Blok(dok, punkt, overskrift, "afsnit", tekst,
                                   stil=b.style.name or "", forside=forside))
        else:
            r = _tabel_raekker(b)
            blokke.append(Blok(dok, punkt, overskrift, "tabel",
                               " | ".join(" | ".join(x) for x in r), raekker=r,
                               forside=forside))
    return blokke


def laes_mappe(mappe: str | Path) -> dict[str, list[Blok]]:
    """Alle .docx i mappen, nøgle = dokument-id ('Kontrakt', 'Bilag 2' ...)."""
    return {dokument_id(p.name): laes_docx(p)
            for p in sorted(Path(mappe).glob("*.docx")) if not p.name.startswith("~$")}


def punkter(blokke: list[Blok]) -> dict[str, str]:
    """Punktnummer -> overskrift, til opslag af henvisninger."""
    return {b.punkt: b.tekst for b in blokke if b.type == "overskrift"}
```

### 7.3 Ordbog for ESDH-paradigmet – `stjernekort/ordbog_esdh.json`

```json
{
  "paradigme": "ESDH-system som en Service",
  "milepaele": [
    {
      "id": "M1",
      "navn": "Kontraktindgåelse",
      "etiket": "Kontrakt-indgåelse",
      "type": "kontraktindgåelse",
      "synonymer": [
        "kontraktindgåelse",
        "kontraktunderskrivelse",
        "underskrevet af både"
      ]
    },
    {
      "id": "M2",
      "navn": "Godkendelse af afklaringsfasen",
      "etiket": "Afklaringsfase godkendt",
      "type": "faseafslutning",
      "synonymer": [
        "godkendelse af afklaringsfasen",
        "afslutning af afklaringsfasen",
        "afklaringsfasen afsluttes",
        "revideret leverancebeskrivelse",
        "godkendt ydelses- og servicebeskrivelse"
      ]
    },
    {
      "id": "M3",
      "navn": "Brugertest",
      "etiket": "Brugertest",
      "type": "prøve",
      "synonymer": [
        "brugertest"
      ]
    },
    {
      "id": "M4",
      "navn": "Prøvekonverteringer",
      "etiket": "Prøve-konverteringer",
      "type": "prøve",
      "synonymer": [
        "prøvekonvertering"
      ]
    },
    {
      "id": "F1",
      "navn": "Brugerdokumentation godkendt",
      "type": "frist",
      "synonymer": [
        "brugerdokumentation skal være godkendt"
      ]
    },
    {
      "id": "M5",
      "navn": "Godkendelse af Ibrugtagningsprøven",
      "etiket": "Ibrugtagnings-prøve godkendt",
      "type": "prøve",
      "synonymer": [
        "ibrugtagningsprøve",
        "ibrugtagningsdag"
      ]
    },
    {
      "id": "M6",
      "navn": "Godkendelse af Driftsprøven",
      "etiket": "Driftsprøve godkendt",
      "type": "prøve",
      "synonymer": [
        "driftsprøve"
      ]
    },
    {
      "id": "F2",
      "navn": "Exit-plan",
      "type": "frist",
      "synonymer": [
        "exit-plan"
      ]
    },
    {
      "id": "F3",
      "navn": "Tidligste opsigelse",
      "type": "frist",
      "synonymer": [
        "tidligst til effekt"
      ]
    },
    {
      "id": "O1",
      "navn": "Option: aflevering til Rigsarkivet",
      "type": "frist",
      "option": true,
      "synonymer": [
        "aflevering til rigsarkivet"
      ]
    },
    {
      "id": "M7",
      "navn": "Kontraktophør",
      "etiket": "Kontrakt-ophør",
      "type": "ophør",
      "synonymer": [
        "kontraktophør",
        "kontrakten løber i"
      ]
    }
  ],
  "retsvirkninger": {
    "betaling": [
      "forfalder til betaling",
      "% af det samlede etableringsvederlag"
    ],
    "bod": [
      "\\bbod\\b",
      "\\bdagbod\\b"
    ],
    "ophævelse": [
      "hæve kontrakten",
      "berettiget til at hæve"
    ],
    "overtagelse": [
      "overtaget af kunden"
    ],
    "ibrugtagning": [
      "tages i fuld brug",
      "ibrugtager kunden"
    ],
    "udtrædelsesadgang": [
      "ret til at udtræde",
      "udtræde af kontrakten"
    ]
  },
  "tabeller": {
    "bod": {
      "overskrift_indeholder": [
        "bodsudløsende"
      ],
      "type": "bod"
    },
    "betaling": {
      "overskrift_indeholder": [
        "milepæl",
        "%"
      ],
      "type": "betaling"
    },
    "tidsplan": {
      "overskrift_indeholder": [
        "milepæle og vigtigste aktiviteter"
      ],
      "type": "tidsplan"
    }
  },
  "forsinkelse": {
    "indledning": "forsinkelse foreligger",
    "retsvirkninger": [
      "ophævelse"
    ],
    "note": "Milepælene i punktopstillingen efter indledningen får retsvirkningen ophævelse (forsinkelse over grænsen i punktet)."
  }
}
```

### 7.4 Udtræk – `stjernekort/udtraek.py`

```python
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
SAETNING_RE = re.compile(r"(?<=[.;:])\s+(?=[A-ZÆØÅ])")


def laes_ordbog(sti: str | Path) -> dict:
    return json.loads(Path(sti).read_text(encoding="utf-8"))


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
                                    etiket=m.get("etiket", "")))
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
            # 3. Sætninger: nøgleord og milepæl skal stå i samme sætning
            if b.type == "afsnit":
                for saetning in SAETNING_RE.split(b.tekst):
                    mids = find_milepael(saetning, ordbog)
                    if not mids:
                        continue
                    t = saetning.lower()
                    for typ, moenstre in ordbog["retsvirkninger"].items():
                        if any(re.search(mo, t) for mo in moenstre):
                            for mid in mids:
                                tilfoej(mid, typ, "", Kilde(dok, b.punkt, _kort(saetning)),
                                        "middel")

    bp.henvisninger = tjek_henvisninger(docs)
    return bp


# --- Henvisninger ----------------------------------------------------------

REF_RE = re.compile(
    r"(?P<ctx>(?:Kontraktens\s+|(?:bilag|Bilag)\s+(?P<bilag>\d+[A-Z]?)(?:,\s*Appendiks\s+\w+)?,?\s+))?"
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
                if ctx.startswith("Kontraktens"):
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
```

### 7.5 Excel som master – `stjernekort/excel_io.py`

```python
"""Excel-arket er master: skriv et Blueprint til .xlsx og læs det tilbage."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from .model import BANER, RETSVIRKNINGER, Aktivitet, Blueprint, Kilde, Milepael, Retsvirkning

FANER = {
    "Milepæle": ["ID", "Navn", "Etiket", "Type", "Rækkefølge", "Option", "Relativ tid", "Dato", "Kilder"],
    "Aktiviteter": ["Bane", "Tekst", "Fra", "Til", "Farvekategori", "Række"],
    "Retsvirkninger": ["Milepæl", "Type", "Detalje", "Hjemmel", "Citat", "Oprindelse"],
    "Henvisninger": ["Fra", "Til dokument", "Til punkt", "Fundet"],
}


def skriv(bp: Blueprint, sti: str | Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    for navn, kol in FANER.items():
        ws = wb.create_sheet(navn)
        ws.append(kol)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E79")
        ws.freeze_panes = "A2"
    ws = wb["Milepæle"]
    for m in bp.milepaele:
        ws.append([m.id, m.navn, m.etiket, m.type, m.raekkefoelge, "ja" if m.option else "nej",
                   m.relativ_tid, m.dato, "; ".join(str(k) for k in m.kilder[:6])])
    ws = wb["Aktiviteter"]
    for a in bp.aktiviteter:
        ws.append([a.bane, a.tekst, a.fra, a.til, a.farvekategori, a.raekke])
    dv = DataValidation(type="list", formula1='"' + ",".join(BANER) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("A2:A500")
    ws = wb["Retsvirkninger"]
    for r in bp.retsvirkninger:
        ws.append([r.milepael, r.type, r.detalje, str(r.kilde) if r.kilde else "",
                   r.kilde.citat if r.kilde else "", r.oprindelse])
    dv2 = DataValidation(type="list", formula1='"' + ",".join(RETSVIRKNINGER) + '"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add("B2:B500")
    ws = wb["Henvisninger"]
    for h in bp.henvisninger:
        ws.append([str(h.fra), h.til_dokument, h.til_punkt,
                   {True: "ja", False: "NEJ", None: "dokument mangler"}[h.fundet]])
    for ws in wb.worksheets:
        for kol in ws.columns:
            ws.column_dimensions[kol[0].column_letter].width = min(
                60, max(10, *(len(str(c.value or "")) for c in kol[:50])) + 2)
    sti = Path(sti)
    wb.save(sti)
    return sti


def _kilde(tekst: str) -> Kilde | None:
    if not tekst:
        return None
    dok, _, punkt = tekst.partition(" pkt. ")
    return Kilde(dok, punkt)


def laes(sti: str | Path) -> Blueprint:
    wb = load_workbook(sti, data_only=True)
    bp = Blueprint(navn=Path(sti).stem)
    for r in wb["Milepæle"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            bp.milepaele.append(Milepael(
                id=str(r[0]), navn=r[1] or "", etiket=r[2] or "", type=r[3] or "",
                raekkefoelge=int(r[4] or 0), option=(r[5] or "").lower() == "ja",
                relativ_tid=r[6] or "", dato=str(r[7] or "")))
    for r in wb["Aktiviteter"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            bp.aktiviteter.append(Aktivitet(r[0], r[1] or "", str(r[2]), str(r[3]),
                                            r[4] or "fase1", int(r[5] or 0)))
    for r in wb["Retsvirkninger"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            bp.retsvirkninger.append(Retsvirkning(str(r[0]), r[1] or "", r[2] or "",
                                                  _kilde(r[3] or ""), r[5] or ""))
    return bp
```

### 7.6 PowerPoint-generator – `stjernekort/pptx_generator.py`

```python
"""Tegner blueprintet i Rambølls blueprint-template ved at "stemple" templatens
egne figurer: paletten over diaset (pilefigurer i farvekategorier), en
milepælsgruppe (linje + Mx-boks + navneboks) og symbolerne i signaturforklaringen.

Figurerne findes via deres shape-id i templaten (se TEMPLATE_KONFIG). Anbefaling:
navngiv stempel-figurerne i PowerPoints markeringsrude, fx "STEMPEL_fase1",
og slå op på navn i stedet for id, så templaten kan redigeres frit.
"""
from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Inches

from .model import BANER, Blueprint

TEMPLATE_KONFIG = {
    "dias": 1,                       # 0-baseret: dias 2 = symboler i banerne
    "milepael_gruppe": 16411,        # Group 16410: linje + "Delleverance-prøve" + "Mx"
    "aktivitet": {                   # farvekategori -> pentagon i paletten
        "fase1": 78, "fase1_2l": 83, "fase2": 89, "fase2_2l": 90,
        "drift": 140, "drift_2l": 86, "videreudvikling": 87, "videreudvikling_2l": 88,
        "governance": 91, "ophoer": 107, "uddannelse": 98,
    },
    "symbol": {                      # retsvirkning -> figur i signaturforklaringen
        "betaling": 16397, "ophævelse": 16398, "bod": 16399, "ibrugtagning": 16400,
        "udtrædelsesadgang": 117, "overtagelse": 121, "bonus": 111,
    },
    "x_start": 1.72, "x_slut": 12.14,  # tommer: kontraktindgåelse .. kontraktophør
    "baner_top": {"Etableringsprojekt": 0.46, "Løbende ydelser": 3.53, "Optioner": 5.14,
                  "Governance": 5.46, "Kundens ydelser": 5.78},
    "symbol_top": 2.2,               # første symbol under aktiviteterne i Etableringsprojekt
    "raekkehoejde": 0.36,
}

def _figur(slide, shape_id):
    for sh in slide.shapes:
        if sh.shape_id == shape_id:
            return sh
    raise KeyError(f"Figur med id {shape_id} findes ikke i templaten")


def _klon(slide, kilde, left=None, top=None, width=None, tekst=None):
    el = copy.deepcopy(kilde._element)
    slide.shapes._spTree.insert_element_before(el, "p:extLst")
    ny = slide.shapes[-1]
    if left is not None:
        ny.left = Emu(int(left))
    if top is not None:
        ny.top = Emu(int(top))
    if width is not None:
        ny.width = Emu(int(width))
    if tekst is not None and ny.has_text_frame:
        _saet_tekst(ny, tekst)
    return ny


def _saet_tekst(shape, tekst: str):
    """Bevarer formateringen fra første run og skriver ny tekst."""
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    runs = p0.runs
    if runs:
        runs[0].text = tekst
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.text = tekst


def _ryd_laerred(slide):
    """Fjerner eksempelindholdet på lærredet; palette og signaturforklaring bevares."""
    y_min, y_max = Inches(0.1), Inches(6.8)
    slet = []
    for sh in slide.shapes:
        tekst = sh.text_frame.text if sh.has_text_frame else ""
        er_milepael = sh.shape_type == 6 and any(
            c.has_text_frame and c.text_frame.text in ("Mx", "Kontrakt-indgåelse")
            for c in sh.shapes)
        er_indhold = sh.shape_type == 1 and y_min < sh.top < y_max and sh.left > Inches(1.1) \
            and tekst not in ("Milepæle",)
        if er_milepael or er_indhold:
            slet.append(sh)
    for sh in slet:
        sh._element.getparent().remove(sh._element)


def _behold_kun_dias(prs, indeks: int):
    sld_ids = prs.slides._sldIdLst
    for i, sld in reversed(list(enumerate(list(sld_ids)))):
        if i != indeks:
            prs.part.drop_rel(sld.rId)
            sld_ids.remove(sld)


def generer(bp: Blueprint, template: str | Path, ud: str | Path, konfig: dict = TEMPLATE_KONFIG,
            vis_typer: tuple[str, ...] = ("kontraktindgåelse", "faseafslutning", "prøve", "ophør")):
    """Tegner blueprintet. Frister vises kun, hvis "frist" er med i vis_typer."""
    prs = Presentation(str(template))
    slide = prs.slides[konfig["dias"]]
    stempler = {
        "milepael": _figur(slide, konfig["milepael_gruppe"]),
        "aktivitet": {k: _figur(slide, v) for k, v in konfig["aktivitet"].items()},
        "symbol": {k: _figur(slide, v) for k, v in konfig["symbol"].items()},
    }
    # Klon stemplerne fra lærredet før det ryddes
    mp_stempel = copy.deepcopy(stempler["milepael"]._element)
    _ryd_laerred(slide)
    slide.shapes._spTree.insert_element_before(mp_stempel, "p:extLst")
    mp_kilde = slide.shapes[-1]

    ms = sorted((m for m in bp.milepaele if m.type in vis_typer), key=lambda m: m.raekkefoelge)
    x0, x1 = Inches(konfig["x_start"]), Inches(konfig["x_slut"])
    skridt = (x1 - x0) / max(len(ms) - 1, 1)
    x_for = {m.id: x0 + i * skridt for i, m in enumerate(ms)}

    # Milepæle: gruppens linje ligger 0,59" inde i gruppen
    linje_offset = Inches(0.59)
    for m in ms:
        g = _klon(slide, mp_kilde, left=x_for[m.id] - linje_offset)
        for c in g.shapes:
            if c.has_text_frame and c.text_frame.text == "Mx":
                _saet_tekst(c, m.id)
            elif c.has_text_frame and c.text_frame.text:
                _saet_tekst(c, m.etiket or m.navn)
    mp_kilde._element.getparent().remove(mp_kilde._element)

    # Aktiviteter
    for a in bp.aktiviteter:
        if a.fra not in x_for or a.til not in x_for:
            continue
        kilde = stempler["aktivitet"].get(a.farvekategori) or stempler["aktivitet"]["fase1"]
        venstre, hoejre = x_for[a.fra], x_for[a.til]
        top = Inches(konfig["baner_top"][a.bane] + a.raekke * konfig["raekkehoejde"])
        _klon(slide, kilde, left=venstre, top=top, width=max(hoejre - venstre, Inches(0.5)),
              tekst=a.tekst)

    # Retsvirkninger: symboler stables under hinanden ved milepælens linje
    stak: dict[str, int] = {}
    tegnet: set[tuple[str, str]] = set()
    for r in bp.retsvirkninger:
        kilde = stempler["symbol"].get(r.type)
        if kilde is None or r.milepael not in x_for or (r.milepael, r.type) in tegnet:
            continue
        tegnet.add((r.milepael, r.type))
        n = stak.get(r.milepael, 0)
        stak[r.milepael] = n + 1
        kol, rk = divmod(n, 4)          # maks. 4 symboler pr. kolonne
        _klon(slide, kilde, left=x_for[r.milepael] - kilde.width // 2 + Inches(0.28) * kol,
              top=Inches(konfig["symbol_top"] + rk * 0.25))

    _behold_kun_dias(prs, konfig["dias"])
    prs.save(str(ud))
    return ud
```

### 7.7 Kommandolinje – `stjernekort/cli.py`

```python
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
```

### 7.8 Accepttest – `tests/test_esdh.py`

```python
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
```

`stjernekort/__init__.py` er en tom fil.

## 8. Byggeplan med prompts til Claude Code

Kør trinnene i rækkefølge, og start hvert trin i plan mode, så du kan godkende planen, før der skrives kode. Commit efter hvert trin, der virker.

### Trin 0 – Opsætning (PowerShell)

```powershell
cd C:\Python
mkdir Stjernekort
cd Stjernekort
mkdir testdata\esdh, ud
pip install python-docx python-pptx openpyxl pytest
```

Kopiér derefter dette notat til `C:\Python\Stjernekort\STJERNEKORT_SPEC.md`, paradigmets .docx-filer til `testdata\esdh\` og templaten til `testdata\`. Filnavnene skal indeholde `Bilag_NN` som i paradigmet. Start Claude Code i mappen.

### Trin 1 – Opret projektet ud fra notatet

```text
Læs STJERNEKORT_SPEC.md. Opret mappestrukturen fra afsnit 6 og filerne fra
afsnit 7 præcis som de står (plus en tom stjernekort/__init__.py). Rediger ikke
koden endnu. Kør derefter:
  python -m pytest -q
  python -m stjernekort.cli udtraek testdata\esdh stjernekort\ordbog_esdh.json ud\esdh.xlsx
og rapportér resultatet. Forventet: 5 tests består; 11 milepæle,
11 retsvirkninger, 38 henvisninger. Tilføj ud/ og __pycache__/ til .gitignore.
```

### Trin 2 – Stempler slås op på navn i stedet for id

```text
Stemplerne i pptx_generator.py findes i dag via shape-id, som ændrer sig, hvis
templaten redigeres. Lav:
1. en kommando "python -m stjernekort.cli figurer <template.pptx>", der lister alle
   figurer på dias 2 med id, navn, tekst, position og fyldfarve, så jeg kan
   navngive stemplerne i PowerPoints markeringsrude (fx STEMPEL_fase1,
   STEMPEL_milepael, STEMPEL_bod);
2. opslag på navn med shape-id som fallback, styret af TEMPLATE_KONFIG.
Flyt TEMPLATE_KONFIG til en JSON-fil ved siden af templaten. Behold testene grønne.
```

### Trin 3 – Standardaktiviteter i ordbogen og "ny"-kommando

```text
Aktiviteterne (pilefigurerne) kommer fra paradigmet og dialogen, ikke fra
udtrækket. Tilføj en nøgle "standardaktiviteter" i ordbog_esdh.json med:
Afklaringsfase M1-M2, Kontraktworkshop M1-M2, Etablering del 1 M2-M3,
Etablering del 2 M3-M5, Driftsprøve M5-M6 (bane Etableringsprojekt);
Drift, support og vedligeholdelse M5-M7, Servicevederlag M5-M7 (Løbende ydelser);
Samarbejdsorganisation M1-M7 (Governance); Kundens ydelser M1-M5 (Kundens ydelser).
Lav kommandoen "python -m stjernekort.cli ny <ordbog.json> <ud.xlsx>", der skriver et
Excel-ark med paradigmets milepæle og standardaktiviteter – startpunktet for dialogen.
Lav også "udtraek" om, så standardaktiviteterne kommer med i Excel.
```

### Trin 4 – Afvigelsesrapport (dialog mod kontrakt)

```text
Lav stjernekort/sammenlign.py og kommandoen
"python -m stjernekort.cli sammenlign <dialog.xlsx> <udtraek.xlsx> <rapport.xlsx>".
Rapporten skal vise:
- milepæle der findes i den ene men ikke den anden (match på ID),
- retsvirkninger (milepæl + type) i dialogen, som kontrakten ikke hjemler,
- retsvirkninger i kontrakten, som ikke var i dialogen,
- ændret detalje (fx 60 % -> 50 %),
- henvisninger med Fundet = NEJ.
Én fane pr. kategori, med hjemmel og citat fra udtrækket. Skriv tests med to
små Blueprint-objekter bygget i koden.
```

### Trin 5 – Versioner og ændringsoversigt

```text
Når "generer" køres, skal den gemme ud\<projekt>_v<NN>_<ÅÅÅÅ-MM-DD>.pptx og en kopi
af Excel-arket med samme navn, hvor NN tælles op automatisk. Tilføj et
ekstra dias efter blueprintet med "Ændret siden v<NN-1>" (nye, fjernede og
flyttede milepæle, aktiviteter og retsvirkninger), baseret på sammenlign.py.
```

### Trin 6 – Pænere layout

```text
Forbedr pptx_generator.py:
- symboler: maks. 4 pr. kolonne ved en milepæl (findes), men undgå overlap med
  aktiviteter i samme højde;
- aktiviteter i samme bane og række må ikke overlappe; flyt automatisk til
  næste række;
- vælg 1- eller 2-linjers stempel efter tekstlængden;
- mulighed for variant dias 3 (symboler i milepælsboksen nederst) via konfig.
Render til PDF med LibreOffice, hvis den findes, ellers spring over.
```

### Trin 7 – Brugerflade og .exe (valgfrit)

```text
Lav en enkel customtkinter-brugerflade (samme stil som mit Extractor-værktøj)
med fire knapper: Nyt blueprint fra paradigme, Udtræk fra kontraktudkast,
Sammenlign, Generér PowerPoint. Vælg filer og mapper med dialogbokse, vis
resultatet i et tekstfelt. Pak med PyInstaller til én .exe, hvor ordbøger og
template ligger i en mappe ved siden af .exe-filen.
```

### Tilføj et nyt paradigme

Kopiér `ordbog_esdh.json` til fx `ordbog_drift.json`, ret milepæle, synonymer og tabelkendetegn, læg paradigmets filer i `testdata\drift\`, og kopiér `tests/test_esdh.py` med paradigmets facitliste. Prompt:

```text
Jeg har lagt et nyt paradigme i testdata\drift. Kør udtrækket med ordbog_esdh.json
og vis mig, hvilke milepæle og retsvirkninger der ikke findes, og hvilke
formuleringer i paradigmet der i stedet bruges. Foreslå ordbog_drift.json.
```

## 9. Kendte begrænsninger

- Sætningsreglerne kan give falske fund. Kolonnen "Oprindelse" viser sikkerheden (høj = tabel eller forsinkelsesdefinition; middel = sætning); middel-fund skal gennemses.
- Nummereringen forudsætter typografierne Heading 1-3. Et paradigme med andre overskriftstypografier kræver en indstilling.
- Med mere end ca. 8 milepæle bliver navneboksene trange; brug kortere etiketter i ordbogen.
- Frister (exit-plan, brugerdokumentation, tidligste opsigelse) udtrækkes, men tegnes ikke som standard (se `vis_typer` i generatoren).

## 10. Åbne beslutninger

- [ ] Tidsakse: kun rækkefølge (antaget) eller tidsskala?
- [ ] Symboler i banerne (dias 2, valgt nu) eller i milepælsboksen (dias 3), eller begge som valg?
- [ ] Skal frister uden retsvirkning stå på blueprintet?
- [ ] Skal banen "Kundens ydelser" fyldes fra bilag 3 eller kun fra dialogen?
- [ ] Selvstændigt værktøj, eller del af Extractor eller tidsplanværktøjet?
- [ ] AI-trin senere, og hvad siger Rambølls retningslinjer om kontraktudkast før offentliggørelse?
