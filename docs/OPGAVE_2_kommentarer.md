# Opgave: ret rækkefølge og markeret tekst i `extract_comments` (extractor.py)

Filen bruger CRLF-linjeskift. **Bevar dem.** Rør ikke ved `extract_trackchanges`,
`extract_highlights`, `extract_kravmatrix`, `write_kravmatrix*` eller
`NumberingEngine`.

Testfil: et Word-dokument med >=100 kommentarer og mindst én svartråd.
I det følgende hedder den `TESTFIL.docx`.

---

## Diagnose (så du ikke retter det forkerte)

`extract_comments` sorterer kommentarer med:

```python
ordered = sorted(meta.items(), key=lambda kv: int(kv[0]))
```

`kv[0]` er `w:id` fra `comments.xml`. **Word tildeler ikke kommentar-id'er i
dokumentrækkefølge.** Målt på testfilen (207 kommentarer):

- 8 kommentarer har id'er som `304033493`, `1214843503`, `2016482722` — Word
  genererer tilfældige 32-bit id'er i visse tilfælde. De sorterer bagerst.
- **196 af 207 kommentarer havner et andet sted, end de står i dokumentet.**

Konsekvensen er synlig for svar: 8 svar landede 140 rækker fra deres tråd.
Selve trådkoblingen via `commentsExtended.xml` er korrekt — 100 af 100 svar
kobles til rigtig forælder. Det er kun rækkefølgen der er gal.

Der er desuden to fejl i `Markeret tekst`:

- Tekst fra to tabelceller løber sammen uden skilletegn
  (`PakningsstørrelseTekstF.eks. "12 x 500 ml"`), fordi `scope_text[cid]`
  akkumulerer rå `w:t`-strenge på tværs af afsnit.
- 60 af 207 rækker (29 %) står tomme. Det er ikke et tab: alle 207 har korrekte
  `commentRangeStart`/`commentRangeEnd`, men 60 er **punktkommentarer** — markøren
  stod i teksten uden at noget var markeret. Uden fallback er de ulæselige.

---

## Ændring 1 — forudsætning: sektionsgenkendelse

`section_for()` kræver rettelsen fra `OPGAVE_extractor_sektionsfix.md` (ændring 1a-1d).
Testdokumentet har manuelt indtastede overskriftsnumre, så uden den står alle 207
rækker som `"Før nummereret sektion"`.

**Er ændring 1a-1d ikke lavet endnu, så lav den først.** Resten af denne opgave
virker uden, men outputtet er ubrugeligt.

---

## Ændring 2 — erstat `extract_comments` helt

Erstat hele funktionen (fra `def extract_comments(docx):` til og med den
afsluttende `return out`) med:

```python
def extract_comments(docx):
    out = []
    if docx.comments_xml is None:
        return out
    meta = {}
    para_to_cid = {}
    for c in docx.comments_xml.findall(q("comment")):
        cid = c.get(q("id"))
        ps = c.findall(q("p"))
        text = "\n".join(filter(None, (para_text(p) for p in ps)))
        meta[cid] = (c.get(q("initials")) or c.get(q("author")) or "",
                     text, c.get(q("date")) or "")
        if ps:
            pid = ps[-1].get(W14 + "paraId")
            if pid:
                para_to_cid[pid] = cid

    parent_of = {}
    if docx.comments_ext_xml is not None:
        for cex in docx.comments_ext_xml.iter(W15 + "commentEx"):
            pid = cex.get(W15 + "paraId")
            parent_pid = cex.get(W15 + "paraIdParent")
            if pid in para_to_cid and parent_pid in para_to_cid:
                parent_of[para_to_cid[pid]] = para_to_cid[parent_pid]

    # scope-tekst, ankerafsnit OG dokumentposition.
    # Word tildeler IKKE kommentar-id'er i dokumentrækkefølge (nogle id'er er
    # tilfældige 32-bit tal), så positionen skal aflæses af dokumentet selv.
    scope_dele, anchor, anchor_pos, active = {}, {}, {}, set()
    for p_i, (p_el, _, _, _) in enumerate(docx.paras):
        loebende = {cid: [] for cid in active}
        for n_i, node in enumerate(p_el.iter()):
            if node.tag == q("commentRangeStart"):
                cid = node.get(q("id"))
                active.add(cid)
                loebende.setdefault(cid, [])
                scope_dele.setdefault(cid, [])
                anchor.setdefault(cid, p_el)
                anchor_pos.setdefault(cid, (p_i, n_i))
            elif node.tag == q("commentRangeEnd"):
                active.discard(node.get(q("id")))
            elif node.tag == q("t") and node.text and active:
                for cid in active:
                    loebende.setdefault(cid, []).append(node.text)
            elif node.tag == q("commentReference"):
                cid = node.get(q("id"))
                anchor.setdefault(cid, p_el)
                anchor_pos.setdefault(cid, (p_i, n_i))
        # ét afsnit/én tabelcelle = ét stykke, så tekst fra to celler ikke
        # løber sammen ("PakningsstoerrelseTekstF.eks. ...")
        for cid, dele in loebende.items():
            t = "".join(dele).strip()
            if t:
                scope_dele.setdefault(cid, []).append(t)

    SIDST = (10 ** 9, 0)

    def pos(cid):
        return anchor_pos.get(cid, SIDST)

    # Traade: rod foerst, derefter svar sorteret efter dato (id-raekkefoelge
    # er ubrugelig - se ovenfor). Roden bestemmer traadens plads i arket.
    def rod(cid):
        r, hop = cid, 0
        while r in parent_of and hop < 50:
            r = parent_of[r]
            hop += 1
        return r

    svar_til = {}
    for cid in meta:
        r = rod(cid)
        if r != cid:
            svar_til.setdefault(r, []).append(cid)

    raekkefoelge = []
    for r in sorted((c for c in meta if rod(c) == c), key=pos):
        raekkefoelge.append(r)
        for s_cid in sorted(svar_til.get(r, []), key=lambda c: (meta[c][2], pos(c))):
            raekkefoelge.append(s_cid)
    for cid in meta:                      # sikkerhedsnet
        if cid not in raekkefoelge:
            raekkefoelge.append(cid)

    num_of = {cid: n for n, cid in enumerate(raekkefoelge, 1)}
    traad_nr, t = {}, 0
    for cid in raekkefoelge:
        r = rod(cid)
        if r not in traad_nr:
            t += 1
            traad_nr[r] = t

    for cid in raekkefoelge:
        initials, text, dato = meta[cid]
        a_el = anchor.get(cid)
        sec = docx.section_for(a_el) if a_el is not None else "Ukendt placering"
        markeret = " | ".join(scope_dele.get(cid, [])).strip()
        if not markeret and a_el is not None:
            # punktkommentar: markoeren stod i teksten uden at der var
            # markeret noget. Vis afsnittet den haenger paa, ellers staar
            # kolonnen tom og kommentaren er ulaeselig ude af kontekst.
            afsnit = para_text(a_el)
            if afsnit:
                markeret = f"(punktkommentar) {afsnit[:300]}"
        r = rod(cid)
        out.append({
            "Dokument": docx.path.name, "Nummer": num_of[cid],
            "Tråd": traad_nr[r],
            "Type": "Ny tråd" if r == cid else "Svar",
            "Svar på": "" if r == cid else f"Nr. {num_of[parent_of[cid]]}",
            "Dato": dato[:10], "Kommentar": text,
            "Markeret tekst": markeret,
            "Initialer": initials, "Nummereret sektion": sec,
        })
    return out
```

Bemærk:

- `meta[cid]` er nu en 3-tuple `(initials, text, date)` — `w:date` findes på alle
  kommentarer og bruges til at sortere svar inden for en tråd.
- `anchor_pos[cid]` er `(afsnitsindeks, nodeindeks)` og aflæses af dokumentet
  under det gennemløb der alligevel foretages. Det er den nye sorteringsnøgle.
- `scope_dele[cid]` er en **liste af stykker** — ét pr. afsnit/tabelcelle —
  der samles med `" | "`. Det er rettelsen af sammenløbet.
- `SIDST = (10 ** 9, 0)` er fallback for en kommentar uden fundet anker, så den
  havner sidst i stedet for at kaste.
- Sikkerhedsnettet nederst (`for cid in meta: if cid not in raekkefoelge`) sikrer
  at ingen kommentar tabes, hvis `parent_of` skulle indeholde en cyklus.

---

## Ændring 3 — kolonner i CLI

I `main()`, grenen `if args.cmd == "kommentarer":`

```python
        write_simple(recs, ["Dokument", "Nummer", "Tråd", "Type", "Svar på",
                            "Dato", "Kommentar", "Markeret tekst", "Initialer",
                            "Nummereret sektion"], out,
                     [35, 8, 7, 9, 9, 11, 50, 50, 10, 40])
```

Den gamle `"Tråd"`-kolonne indeholdt `"Ny tråd"`/`"Svar"`. Den hedder nu `"Type"`,
og `"Tråd"` er et løbenummer pr. samtale, så man kan filtrere på én tråd ad gangen.

---

## Ændring 4 — GUI'en

`extractor_gui.py` har to knapper der bruger kommentarudtrækket
(`kommentarer` og `kommentarer_filtreret`). Tjek om de bygger deres egen
kolonneliste i stedet for at kalde `write_simple` med samme liste som CLI'en —
hvis de gør, skal listen opdateres begge steder. Filtreringen ("kun dokumenter
med fund") skal virke uændret; den ser kun på om `extract_comments` returnerer
noget.

---

## Accepttest

```bash
python extractor.py kommentarer TESTFIL.docx -o test_kom.xlsx
```

Forventet på en fil med 207 kommentarer, heraf 100 svar:

| Måling | Før | Efter |
|---|---|---|
| Rækker i alt | 207 | **207** (ingen tabt) |
| Svar koblet til rigtig forælder | 100 | **100** (uændret) |
| Svar adskilt fra sin tråd | 8 | **0** |
| Spring baglæns i sektionsrækkefølge | 196 reelt | **0** |
| Tom `Markeret tekst` | 60 | **26** (34 udfyldt med fallback) |
| Celletekst løbet sammen | 30 | **0** |
| `Nummereret sektion` udfyldt | 0 | **207** |

Kontrollér desuden manuelt:

- Hvert `Svar` står umiddelbart efter sin forælder, og `Svar på` peger på det
  `Nummer` der står lige ovenfor eller få rækker oppe i samme `Tråd`.
- `Tråd`-numre er stigende gennem arket og springer aldrig tilbage.
- Punktkommentarer har `Markeret tekst` der begynder med `(punktkommentar)`.
- `Dato` er en tekststreng på formen `ÅÅÅÅ-MM-DD` i alle rækker (nogle
  `w:date`-værdier ender på `Z`, andre ikke — begge skal give samme format).

Kør til sidst de øvrige underkommandoer mod samme fil for at bekræfte at intet
er brudt:

```bash
python extractor.py trackchanges TESTFIL.docx -o test_tc.xlsx
python extractor.py styles       TESTFIL.docx
```
