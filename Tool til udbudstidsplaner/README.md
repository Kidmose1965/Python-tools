# Tool til udbudstidsplaner

Genererer en visuel PowerPoint-tidsplan (kundeversion) ud fra et Excel-udtræk
kopieret direkte fra Microsoft Project.

Microsoft Project er masterplanen. Excel er kun transportformat. PowerPoint
er den visuelle kundeversion. Der er ingen direkte .mpp-integration - Excel-
udtrækket er den eneste vej ind i værktøjet.

## Sådan bruges værktøjet

1. Lav/vedligehold tidsplanen i Microsoft Project.
2. Kopiér Project-tabellen (fx Gantt-visningen) til Excel via copy/paste.
3. Bevar farverne fra Project/Excel - de bruges til at fortolke belastning,
   møder og milepæle (se farveskema nedenfor).
4. Gem Excel-arket som `.xlsx`.
5. Kør enten den grafiske brugerflade (`python tidsplan_gui.py`) eller CLI'en
   (`python tidsplan.py ...`, se nedenfor).
6. PowerPoint-tidsplanen (`.pptx`) genereres automatisk.

## Forventet Excel-struktur

- Række 2 indeholder headers.
- Data starter i række 3.
- Kolonne A kan være tom.
- Kolonne B: Task Name (aktivitetsnavn).
- Kolonne C: Duration (fx "5 days" eller "5 dage").
- Kolonne D: Start.
- Kolonne E: Finish.
- Kolonne F/G (Predecessors/Successors) læses ikke af værktøjet og behøver
  ikke være udfyldt.
- Faser/summary tasks: fed skrift, ikke indrykket navn.
- Aktiviteter under en fase: indrykket (mellemrum/tab foran navnet).
- Cellefarven i Task Name-kolonnen (kolonne B) er semantisk - se nedenfor.

## Farveskema (cellefarve i Task Name-kolonnen)

| Farve (hex) | Betydning              | Vises som                                   |
|-------------|-------------------------|----------------------------------------------|
| `FF0000`    | Spidsbelastning          | Rød bjælke/belastningsbånd                    |
| `FFC000`    | Belastning                | Gul/orange bjælke/belastningsbånd            |
| `92D050`    | Bolden hos tilbudsgiver / ekstern | Grøn bjælke/belastningsbånd          |
| `7030A0`    | Godkendelse (STG/DB)      | Lilla bjælke/belastningsbånd                 |
| `00B0F0`    | Officiel milepæl (0 dage) | Blå diamant                                  |
| `ED7D31`    | Møde                       | Orange markør (0 dage) eller orange bjælke (>0 dage) |
| Hvid / ingen relevant fill | Neutral      | Grå bjælke                                    |
| Tema-/indekseret farve (ingen RGB-værdi) | Kan ikke fortolkes semantisk | Behandles som neutral + advarsel |

Yderligere regler:

- **Milepæle:** en aktivitet med Duration = 0 er en punktaktivitet. Hvis
  navnet starter med `Milepæl:`, eller cellen er fyldt `00B0F0`, vises den
  som en officiel blå milepælsmarkør (præfikset `Milepæl:` fjernes fra
  visningsnavnet). Andre 0-dages aktiviteter uden særlig semantik vises som
  en grå punktmarkør. Hvis `Milepæl:`-præfikset står på en aktivitet med
  Duration > 0, fjernes præfikset stadig fra visningsnavnet, men aktiviteten
  vises som en almindelig bjælke i sin fills farve - ikke som en blå markør
  (blå markør er forbeholdt punktaktiviteter).
- **Møder (`ED7D31`):** en selvstændig kategori, uafhængig af belastning og
  milepæle. Duration = 0 -> orange mødemarkør. Duration > 0 -> orange
  aktivitetsbjælke. Møder påvirker aldrig belastningsbåndet pr. uge, uanset
  varighed.
- **Belastningsbånd pr. uge:** viser den højeste relevante belastning blandt
  ugens aktiviteter, med prioritet Spidsbelastning > Belastning >
  Godkendelse > Neutral > Ekstern. Møder og milepæle indgår ikke. Hvis en uge
  udelukkende indeholder møder/milepæle (eller ingen aktiviteter), vises
  ugen uden belastningsfarve.

## Kør via GUI

```
python tidsplan_gui.py
```

Vælg Excel-fil, statusdato, kundenavn, udbudsnavn, en evt. gemt
konfiguration (avanceret) og outputfil, og klik "Lav tidsplan".

## Kør via CLI

```
python tidsplan.py "Excel udtræk tidsplan.xlsx" ^
  --idag 2026-09-17 ^
  --kunde "Modelkunde" ^
  --udbud "Udbud med forhandling - 1 runde" ^
  -o "Test - modeltidsplan.pptx"
```

Valgfrit: `--config konfigurationer/mit_udbud.json` for fase-omdøbninger,
korte aktivitetsnavne, ekstra milepæle og undtagelser fra belastningsbåndet
(se `Konfig`/`load_konfig()` i `tidsplan.py`).

## Krav

Se `requirements.txt` i repo-roden: `openpyxl`, `Pillow`, `python-pptx`,
`pytest`. Installér med:

```
pip install -r requirements.txt
```

## Tests

```
pytest test_tidsplan.py -v
```

## Arkitektur

```
Excel -> laes_plan() -> byg() -> grafiske primitiver -> pptx_render.render() -> PowerPoint
```

`tidsplan.py` står for datalæsning og layout, `pptx_render.py` for selve
PowerPoint-renderingen (python-pptx). Der er ikke direkte .mpp-integration,
og værktøjet ændrer aldrig Project-data eller planens faglige indhold - det
tegner den plan, der allerede står i Excel-udtrækket.
