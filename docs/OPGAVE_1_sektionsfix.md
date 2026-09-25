# Opgave: ret sektionsgenkendelse og berig trackchanges-udtrækket i `extractor.py`

Filen bruger CRLF-linjeskift. Bevar dem. Rør ikke ved `NumberingEngine`,
`_find_hoveddel`, `extract_kravmatrix`, `write_kravmatrix*` eller Excel-outputtet
ud over de kolonner der er nævnt eksplicit nedenfor.

---

## Baggrund (så du ikke "retter" det forkerte sted)

`Docx._para_label()` returnerer `None`, medmindre afsnittet har `w:numPr` — direkte
eller arvet via typografi. `section_for()` går derfor baglæns gennem hele
dokumentet, finder aldrig et label, og returnerer `"Før nummereret sektion"`.

Testdokumentet (Bilag 3A, Appendiks D) har **0 `numPr` på 163 overskrifter**, og
heller ingen på typografierne Heading 1–6. Numrene "3.1 -", "6.1.1 -" er tastet
manuelt ind i overskriftsteksten. `NumberingEngine` er korrekt — der er bare intet
at læse. Alle 349 rækker havner derfor uden sektion.

Bemærk: `styleId` er `Overskrift1` i det ene dokument og `Heading1` i det andet,
mens `w:name` er `heading 1` i begge. Match derfor på **visningsnavnet**
(`para_style_name`), ikke på styleId.

---

## Ændring 1 — tekstbaseret fallback for sektionsnummer (den vigtigste)

I klassen `Docx`:

### 1a. Gem style-elementerne, så `outlineLvl` kan slås op

I `__init__`, hvor `self.styles = {}` sættes, tilføj `self._style_elements = {}`,
og gem `st` i løkken:

```python
self.styles = {}
self._style_elements = {}
styles_root = load_i_mappe("styles.xml")
if styles_root is not None:
    for st in styles_root.findall(q("style")):
        sid = st.get(q("styleId"))
        ...
        self._style_elements[sid] = st
        self.styles[sid] = (...)
```

### 1b. Nye metoder på `Docx`

```python
    # Overskrift med manuelt indtastet nummer, fx "3.1 - Dyrlæger" eller
    # "6.1.1\tRecept". Kræver at der står tekst efter nummeret.
    _MANUELT_NR = re.compile(r"^\s*(\d+(?:\.\d+)*)\s*(?:[-\u2013\u2014.)\t ]\s*)(\S.*)$")

    def _style_element(self, sid):
        return self._style_elements.get(sid)

    def _er_overskrift(self, p):
        """True hvis afsnittet er en overskrift - enten via de indbyggede
        typografinavne ("heading 1"..."heading 9") eller via outlineLvl,
        som brugerdefinerede overskriftstypografier bruger."""
        if re.match(r"heading [1-9]$", self.para_style_name(p) or "", re.I):
            return True
        ppr = p.find(q("pPr"))
        if ppr is not None:
            if ppr.find(q("outlineLvl")) is not None:
                return True
            ps = ppr.find(q("pStyle"))
            if ps is not None:
                sid = ps.get(q("val"))
                for _ in range(10):          # følg basedOn-kæden
                    if sid not in self.styles:
                        break
                    st_el = self._style_element(sid)
                    if st_el is not None and st_el.find(
                            f"{q('pPr')}/{q('outlineLvl')}") is not None:
                        return True
                    sid = self.styles[sid][2]
        return False

    def _tekst_label(self, p, text):
        """Fallback for dokumenter uden automatisk nummerering: læs nummeret
        ud af selve overskriftsteksten."""
        if not text or not self._er_overskrift(p):
            return None
        m = self._MANUELT_NR.match(text)
        return m.group(1) if m else None
```

### 1c. Brug fallback i `_scan()`

```python
        for el, in_tbl in self._iter_block(body, False):
            label = self._para_label(el)
            text = para_text(el, include_del=True)
            if label is None:
                label = self._tekst_label(el, text)
            self.index_of[id(el)] = len(self.paras)
            self.paras.append((el, label, text, in_tbl))
```

`include_del=True` er nødvendigt: i en compare-fil kan dele af en overskrift
ligge i `<w:delText>`, og uden det bliver overskriftsteksten afkortet.

### 1d. Undgå dobbelt nummer i `section_for()`

Når labelet kom fra teksten, står nummeret allerede i `text` — ellers får du
"3.1 - 3.1 - Dyrlæger":

```python
            if label and label[:1].isdigit():
                if text.lstrip().startswith(label):
                    return text.strip()
                return f"{label} - {text}".strip(" -")
```

---

## Ændring 2 — ny kolonne `Placering` (tabel/række/kolonne)

I et datamodelbilag ligger 182 af 349 ændringer inde i tabeller. Uden koordinat
er de umulige at finde igen.

I `_scan()`, sammen med de øvrige felter:

```python
        self.celle_ref = {}      # id(afsnit) -> "Tabel 4, række 7, kolonne 2"
        self._tabel_nr = 0
```

Udvid `_iter_block` med en `ref`-parameter:

```python
    def _iter_block(self, parent, in_table, ref=None):
        for child in parent:
            if child.tag == q("p"):
                if ref:
                    self.celle_ref[id(child)] = ref
                yield child, in_table
            elif child.tag == q("tbl"):
                self._tabel_nr += 1
                tnr = self._tabel_nr
                for ri, tr in enumerate(child.findall(q("tr")), 1):
                    for ci, tc in enumerate(tr.findall(q("tc")), 1):
                        yield from self._iter_block(
                            tc, True, f"Tabel {tnr}, række {ri}, kolonne {ci}")
```

---

## Ændring 3 — ny kolonne `Kontekst` + sammenlægning af naboer

Dette er det egentlige problem: medianlængden på `Ændring`-kolonnen er 14 tegn.
En række med `en` → `et` er ubrugelig uden sætningen omkring.

Erstat hele `extract_trackchanges` og tilføj to hjælpefunktioner over den:

```python
def _kontekst(p_el):
    """Hele afsnittet med {-slettet-} / {+indsat+} markeret."""
    ud = []

    def gaa(el, tilstand):
        for barn in el:
            t = barn.tag
            ny = tilstand
            if t == q("ins"):
                ny = "ins"
            elif t == q("del"):
                ny = "del"
            elif t == q("pPr"):
                continue          # afsnitsmærke-ændringer er ikke tekst
            if t == q("t") and barn.text:
                ud.append("{+%s+}" % barn.text if tilstand == "ins"
                          else "{-%s-}" % barn.text if tilstand == "del"
                          else barn.text)
            elif t == q("delText") and barn.text:
                ud.append("{-%s-}" % barn.text)
            else:
                gaa(barn, ny)

    gaa(p_el, None)
    return "".join(ud).strip()


def _placering(docx, p_el):
    return docx.celle_ref.get(id(p_el), "brødtekst")


def extract_trackchanges(docx, saml=True):
    out = []
    for p_el, _, _, _ in docx.paras:
        sektion = docx.section_for(p_el)
        kontekst = _kontekst(p_el)
        placering = _placering(docx, p_el)
        fund = []
        for node in p_el.iter():          # .iter(), ikke direkte børn - se note
            if node.tag == q("ins"):
                txt = para_text(node)
                if txt:
                    fund.append(("Indsat", txt, node.get(q("author")) or ""))
            elif node.tag == q("del"):
                txt = para_text(node, include_del=True)
                if txt:
                    fund.append(("Slettet", txt, node.get(q("author")) or ""))
        if saml:
            # Word splitter ofte én redigering op i mange <w:ins>/<w:del>
            samlet = []
            for typ, txt, forf in fund:
                if samlet and samlet[-1][0] == typ and samlet[-1][2] == forf:
                    samlet[-1][1] += txt
                else:
                    samlet.append([typ, txt, forf])
            fund = [tuple(x) for x in samlet]
        for typ, txt, forf in fund:
            out.append({"Dokument": docx.path.name,
                        "Sektion": sektion,
                        "Placering": placering,
                        "Ændring": txt, "Type": typ,
                        "Kontekst": kontekst,
                        "Forfatter": forf})
    return out
```

**Note om `p_el.iter()`:** den nuværende `for node in p_el` ser kun direkte børn
og taber `ins`/`del` der ligger inde i `<w:hyperlink>` eller `<w:smartTag>`. Det
sker ikke i testdokumentet, men det sker i andre — og det fejler lydløst.

### CLI-kolonner

I `main()`, grenen `elif args.cmd == "trackchanges":`

```python
        write_simple(recs, ["Dokument", "Sektion", "Placering", "Ændring",
                            "Type", "Kontekst", "Forfatter"],
                     out, [35, 45, 26, 55, 10, 90, 20])
```

---

## Ændring 4 — advar når alle ændringer har samme forfatter

`node.get(q("author"))` er korrekt kode. Men i et **Word Compare-output** stempler
Word alle forskelle med den bruger der kørte sammenligningen — ikke med den der
lavede ændringerne. Alle 349 rækker stod derfor som samme navn, hvilket er
misvisende.

Skriv en advarsel til stdout når `trackchanges` har fundet ændringer og samtlige
rækker har samme ikke-tomme `Forfatter`:

```
ADVARSEL: alle N ændringer er tilskrevet "<navn>". Er filen et Word
Compare-output, er forfatteren den der kørte sammenligningen - ikke den
der lavede ændringerne.
```

---

## Valgfri ændring 5 — strukturændringer uden tekst

Af 473 `ins`/`del`-elementer i testfilen ligger:

| Placering | Antal | Betydning |
|---|---|---|
| `w:p` (tekst) | 351 | fanges i dag |
| `pPr/rPr` | 93 | slettede/indsatte afsnitsmærker |
| `tr/trPr` | 29 | **hele tabelrækker slettet (15) / indsat (14)** |

Rækkemarkørerne er de interessante: celleteksten fanges nok, men man kan ikke
skelne "et ord er slettet inde i en række" fra "hele rækken er væk" — og det er
forskellen mellem en sproglig rettelse og en fjernet attribut i en datamodel.

Læs `<w:del>`/`<w:ins>` i `<w:trPr>` med i `_iter_block`, og udsend dem som
`Type = "Række slettet"` / `"Række indsat"` med rækkens celletekst i `Ændring`.

---

## Accepttest

Kør mod compare-filen for Bilag 3A, Appendiks D:

```bash
python extractor.py trackchanges <compare>.docx -o test.xlsx
python extractor.py kommentarer  <compare>.docx -o test_kom.xlsx
python extractor.py trackchanges <original>.docx -o test_orig.xlsx
```

Forventet:

- `trackchanges`: **341 rækker** (349 før sammenlægning), **46 unikke sektioner**,
  **0 rækker** med "Før nummereret sektion"
- Sektionsværdier ser ud som `3.1 - Dyrlæger`, ikke `3.1 - 3.1 - Dyrlæger`
- `Placering` er "brødtekst" eller `Tabel 14, række 60, kolonne 2`
- `Kontekst` indeholder hele sætningen med `{-...-}`/`{+...+}`
- `kommentarer`: 28 rækker, alle med udfyldt "Nummereret sektion"
  (fx `3.3 - Vet-forhandlere`, `6.1.1 - Recept`)
- Originalfilen giver 1 række ("Ingen trackchanges") — ingen crash

Kør også `kravmatrix` og `inputfelter` på et vilkårligt bilag for at bekræfte at
de stadig virker: begge bruger samme `section_for()` og skal nu også få rigtige
sektionsnumre.
