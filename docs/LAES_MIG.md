# Rettelser til extractor.py — sådan kører du dem

Pakken indeholder to opgavebeskrivelser til Claude Code, en referenceudgave af
`extractor.py` og et forventet resultat til sammenligning.

**Filerne skal pakkes ud i den mappe hvor din `extractor.py` ligger** (fx `C:\Python`).

---

## Trin 1 — forbered mappen

Efter udpakning skal mappen se sådan ud:

```
C:\Python\
    extractor.py                          <- din egen, den der bliver rettet
    extractor_gui.py
    krydstjek.py
    semantik.py
    intern_tjek.py

    LAES_MIG.md                           <- denne fil
    OPGAVE_1_sektionsfix.md
    OPGAVE_2_kommentarer.md
    extractor_reference.py                <- KUN til sammenligning
    forventet_resultat_kommentarer.xlsx   <- KUN til sammenligning

    XXBILA_2.docx                         <- testfil, se nedenfor
```

**Omdøb testfilen.** `XXBILA_2.DOC` er reelt et Word 2007+-dokument, men
`extractor.py` afviser den på filendelsen. Omdøb til `XXBILA_2.docx`.

**Overskriv IKKE din `extractor.py` med `extractor_reference.py`.**
Referenceudgaven har mistet filens CRLF-linjeskift undervejs, så en `diff`
mellem de to viser hele filen som ændret. Den er kun til at slå op i, hvis
Claude Code er i tvivl.

---

## Trin 2 — tag en kopi

```bash
cd C:\Python
git init
git add .
git commit -m "foer rettelser"
```

Bruger du ikke git:

```bash
copy extractor.py extractor_backup.py
```

Du skal kunne komme tilbage.

---

## Trin 3 — kør Claude Code

```bash
cd C:\Python
claude
```

Indsæt denne prompt:

```
Læs OPGAVE_1_sektionsfix.md og implementér ændring 1a-1d i extractor.py.
Bevar filens CRLF-linjeskift. Kør accepttesten mod XXBILA_2.docx til sidst.
```

Når den er færdig, så kontrollér selv:

```bash
python extractor.py kommentarer XXBILA_2.docx -o test1.xlsx
```

Kolonnen **Nummereret sektion** skal nu være udfyldt i alle rækker
(fx `3.3 - Vet-forhandlere`). Står der `Før nummereret sektion`, er ændring 1
ikke slået igennem — gå ikke videre.

---

## Trin 4 — anden omgang

Samme session, ny prompt:

```
Læs OPGAVE_2_kommentarer.md og implementér ændring 2-4 i extractor.py og
extractor_gui.py. Bevar CRLF. Kør accepttesten til sidst.
```

Kontrollér:

```bash
python extractor.py kommentarer XXBILA_2.docx -o test2.xlsx
```

Hold `test2.xlsx` op mod `forventet_resultat_kommentarer.xlsx`. De skal være
identiske. Er de ikke, har Claude Code afveget fra beskrivelsen — og det er
nemmere at se på et regneark end i en diff.

---

## Hvorfor to omgange

Ændring 1 rører `Docx`-klassen og påvirker **alle fire** udtrækstyper
(kommentarer, trackchanges, inputfelter, kravmatrix). Ændring 2 rører kun
`extract_comments`. Kører du dem samlet og noget opfører sig mærkeligt
bagefter, ved du ikke hvilken af dem der gjorde det.

---

## Hvad der ændrer sig i outputtet

Kommentarudtrækket får nye kolonner:

| Før | Efter |
|---|---|
| Dokument, Nummer, **Tråd**, Svar på, Kommentar, Markeret tekst, Initialer, Nummereret sektion | Dokument, Nummer, **Tråd**, **Type**, Svar på, **Dato**, Kommentar, Markeret tekst, Initialer, Nummereret sektion |

**Bemærk at `Tråd` skifter betydning.** Før indeholdt den `Ny tråd`/`Svar`.
Nu er den et løbenummer pr. samtale, og den gamle værdi flytter til `Type`.
Har du regneark, pivottabeller eller makroer der læser den kolonne, skal de
rettes.

Målt på XXBILA_2 (207 kommentarer, heraf 100 svar):

| Måling | Før | Efter |
|---|---|---|
| Svar adskilt fra sin tråd | 8 | 0 |
| Kommentarer placeret forkert ift. dokumentet | 196 | 0 |
| Tom "Markeret tekst" | 60 | 26 |
| Celletekst løbet sammen | 30 | 0 |
| "Nummereret sektion" udfyldt | 0 af 207 | 207 af 207 |

---

## Hvad der IKKE er med i denne omgang

Følgende er diagnosticeret, men bevidst udeladt:

- `find_bilag`s `rstrip(".0")` — Bilag 10 rapporteres som skrivemådefejl for
  Bilag 1 (`krydstjek.py`)
- `semantik.py` gemmer API-fejl som vurdering `"FEJL"`, hvilket i skemaet er en
  substansdom
- `semantik.py`s emoji-`print()` inde i `try` — på en cp1252-konsol (dansk
  Windows) overskriver den et **gennemført** kald med samme `"FEJL"`
- `§39` og bare §-henvisninger som falske positive i `krydstjek.py`
- Manglende deduplikering af identiske fund

De rører andre filer og hører til en senere omgang.
