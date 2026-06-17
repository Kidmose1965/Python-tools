EXTRACTOR - PYTHON-UDGAVE
=========================
Python-version af "Revised Extractor". Læser .docx-filer direkte
(som XML) - Word åbnes aldrig. Ingen skjulte Word-processer, ingen
fillåse, ingen risiko for at lukke dokumenter du selv har åbne.
Virker også selvom Word slet ikke er installeret.

INSTALLATION (engangsopgave)
  1. Installer Python fra jeres software store.
  2. Åbn en kommandoprompt (tryk Windows-tasten, skriv "cmd", Enter)
     og kør:
        pip install openpyxl
     Det er den eneste pakke, værktøjet kræver.
  3. Læg extractor.py i en mappe efter eget valg, fx sammen med
     de kravspecifikationer du arbejder med.

BRUG
  Åbn en kommandoprompt i mappen med extractor.py (i Stifinder:
  skriv "cmd" i adresselinjen og tryk Enter). Kør derefter:

  Se dokumentets typografier (godt at starte med):
    python extractor.py styles kravspec.docx

  Udtræk kommentarer (en eller flere filer ad gangen):
    python extractor.py kommentarer kravspec.docx bilag2.docx

  Udtræk trackchanges:
    python extractor.py trackchanges kravspec.docx

  Udtræk kundens inputfelter (grøn markering):
    python extractor.py inputfelter --farve groen kravspec.docx

  Udtræk tilbudsgivers inputfelter (gul markering):
    python extractor.py inputfelter --farve gul kravspec.docx

  Dan kravmatrix (interaktiv - du vælger Styles ud fra en nummereret
  liste, præcis som i Excel-værktøjets formularer):
    python extractor.py kravmatrix kravspec.docx

  Dan kravmatrix med kommentarer (samme krav-rækker som ovenfor, men
  med en ekstra kolonne der viser alle Word-kommentarer der er
  knyttet til den pågældende kravrække):
    python extractor.py kravmatrix-kommentarer kravspec.docx

  Alle kommandoer accepterer -o til at vælge output-filnavn:
    python extractor.py kommentarer kravspec.docx -o mine_kommentarer.xlsx

  Uden -o gemmes resultatet i mappen som kommentarer.xlsx,
  trackchanges.xlsx, inputfelter_groen.xlsx, inputfelter_gul.xlsx
  eller kravmatrix.xlsx.

  Tip: du kan bruge jokertegn til at tage en hel mappe ad gangen:
    python extractor.py kommentarer *.docx

OUTPUT
  Kolonnerne svarer til Excel-værktøjet:
  - Kommentarer:   Dokument, Nummer, Kommentar, Markeret tekst,
                   Initialer, Nummereret sektion
  - Trackchanges:  Dokument, Sektion, Ændring, Type (Indsat/Slettet),
                   Forfatter
  - Inputfelter:   Dokument, Sektion, Inputfelt
  - Kravmatrix:    Skabelonen med Krav nr. / Kravbeskrivelse /
                   Krav opfyldt (J/N/D) / Standard-programmel /
                   Tilpasning/opsætning / Redegørelse, med overskrifter
                   i fed og zebrastriber som i originalen.
  - Kravmatrix med kommentarer: Krav nr. / Kravbeskrivelse / Kommentarer
                   (samler alle Word-kommentarer der er knyttet til
                   den enkelte kravrække i en kolonne).

  "Sektion" er som i Excel-værktøjet den nærmeste foranstående
  NUMMEREREDE overskrift (fx "2.1 - Funktionelle krav").

FORSKELLE FRA EXCEL-VÆRKTØJET
  + Markant hurtigere (sekunder, ikke minutter)
  + Kan tage mange filer / hele mapper i én kørsel
  + Trackchanges viser nu også forfatter
  + Kræver ikke Word, ActiveX eller makro-tilladelser
  - Kun .docx (ikke gamle .doc-filer - gem dem som .docx først)
  - Meget eksotiske nummereringsopsætninger i Word (fx manuelle
    "genstart nummerering"-overstyringer) kan i sjældne tilfælde
    give en anelse anden nummerering end Word viser. Standard
    Heading-nummerering håndteres korrekt.

FEJLSØGNING
  "python blev ikke fundet"  -> Python er ikke installeret endnu,
                                eller prøv kommandoen "py" i stedet.
  "No module named openpyxl" -> kør: pip install openpyxl
  Forkerte/manglende sektioner -> kør "styles"-kommandoen og tjek,
                                at dokumentet bruger nummererede
                                overskrifts-typografier.
