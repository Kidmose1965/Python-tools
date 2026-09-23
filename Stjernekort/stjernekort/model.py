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
    advarsler: list[tuple[str, str, str]] = field(default_factory=list)  # (begreb, kilde, citat)

    def milepael(self, mid: str) -> Milepael | None:
        return next((m for m in self.milepaele if m.id == mid), None)
