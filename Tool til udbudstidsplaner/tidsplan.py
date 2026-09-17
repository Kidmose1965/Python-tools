"""
Visuel udbudstidsplan - generisk skabelon
Brug:  python tidsplan.py <plan.xlsx> --idag 2026-09-10 --kunde "Min Kunde" --udbud "Mit udbud" [-o Tidsplan_visuel.pptx]
       python tidsplan.py <plan.xlsx> --idag 2026-09-10 --config konfigurationer/mit_udbud.json

Input: Excel-udtraek fra MS Project (kolonner B-E: Task Name, Duration, Start, Finish; header i
raekke 2, data fra raekke 3). Kolonne F/G (Predecessors/Successors) laeses ikke.
  - Fase-raekker: fed, ikke-indrykket navn. Aktiviteter: indrykket med mellemrum/tab.
  - Farve i kolonne B (Task Name) er semantisk - se FARVER nedenfor. Hvid/ukendt = neutral (graa bjaelke).
  - 0 dage = punktaktivitet. "Milepael:"-praefiks eller blaa (00B0F0) udfyldning -> blaa officiel
    milepael (praefikset fjernes fra visningsnavnet). ED7D31 -> moede (orange markoer ved 0 dage,
    orange bjaelke ved >0 dage; paavirker aldrig belastningsbaandet). Andre 0-dages aktiviteter -> graa
    punktmarkoer. "Milepael:"-praefiks ved >0 dage vises som almindelig bjaelke, ikke blaa markoer.

FARVER (Birgers koder, cellefarve i kolonne B):
  FF0000 Spidsbelastning   FFC000 Belastning   92D050 Ekstern/tilbudsgiver   7030A0 Godkendelse
  00B0F0 Officiel milepael  ED7D31 Moede         hvid/ukendt/tema-/indekseret Neutral

Kunde/udbud-navn og evt. projektspecifikke tilfoejelser (fase-omdoebninger, ekstra
milepaele) angives enten som --kunde/--udbud direkte, eller via en JSON-konfigurationsfil
(--config), se load_konfig(). --kunde/--udbud overskriver konfigurationsfilens vaerdier,
hvis begge angives. Uden nogen af delene tegnes en helt generisk tidsplan direkte fra
Excel-filen.

Kraever: Python (openpyxl, Pillow, python-pptx) - se pptx_render.py i samme mappe.
"""
import argparse, json, math, os, re, sys, datetime as dt
from dataclasses import dataclass, field
from openpyxl import load_workbook
from PIL import ImageFont

D = dt.date


@dataclass
class Konfig:
    """Kunde-/udbudsspecifik konfiguration. Alle felter er valgfrie - uden dem
    tegnes en generisk tidsplan direkte fra Excel-filens egne navne og datoer."""
    kunde: str = ""
    udbud: str = ""
    # Visningsnavne for faser (Excel-navn -> lane-label). Ukendte faser bruger Excel-navnet.
    fase_navne: dict = field(default_factory=dict)
    # Forkortede aktivitetsnavne (Excel-navn uden "Milepæl:" -> visningsnavn)
    korte_navne: dict = field(default_factory=dict)
    # Ekstra milepaele der ikke staar i MS Project-planen: (fase i Excel, navn, dato, blaa?)
    ekstra_milepaele: list = field(default_factory=list)
    # Aktiviteter der ikke taeller med i belastningsbaandet
    ikke_i_belastning: set = field(default_factory=set)

    @property
    def titel(self):
        return f"Udbud af {self.udbud} – tidsplan" if self.udbud else "Udbudstidsplan"

    @property
    def hos_kunde(self):
        """' hos <kunde>' hvis kunde er angivet, ellers tom streng."""
        return f" hos {self.kunde}" if self.kunde else ""

    @property
    def suffix_kunde(self):
        """' <kunde>' hvis kunde er angivet, ellers tom streng (til fx 'Spidsbelastning <kunde>')."""
        return f" {self.kunde}" if self.kunde else ""

    @property
    def label_belastning(self):
        return f"{self.kunde}-belastning pr. uge" if self.kunde else "Belastning pr. uge"


def load_konfig(sti):
    with open(sti, encoding="utf-8") as f:
        raw = json.load(f)
    k = Konfig(
        kunde=raw.get("kunde", ""),
        udbud=raw.get("udbud", ""),
        fase_navne=dict(raw.get("fase_navne", {})),
        korte_navne=dict(raw.get("korte_navne", {})),
        ikke_i_belastning=set(raw.get("ikke_i_belastning", [])),
    )
    k.ekstra_milepaele = [
        (m["fase"], m["navn"], D.fromisoformat(m["dato"]), bool(m.get("blaa", False)))
        for m in raw.get("ekstra_milepaele", [])
    ]
    return k


def save_konfig(konfig, sti):
    raw = dict(
        kunde=konfig.kunde,
        udbud=konfig.udbud,
        fase_navne=konfig.fase_navne,
        korte_navne=konfig.korte_navne,
        ekstra_milepaele=[dict(fase=f, navn=n, dato=d.isoformat(), blaa=b)
                          for f, n, d, b in konfig.ekstra_milepaele],
        ikke_i_belastning=sorted(konfig.ikke_i_belastning),
    )
    with open(sti, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)


# Farver (Birgers koder) -> belastningsniveau - fast visuelt sprog for alle tidsplaner
RED, ORANGE, GREEN, PURPLE = "FF0000", "FFC000", "92D050", "7030A0"
NEUTRAL, MS_BLUE, MS_GREY, MOEDE = "A6A6A6", "00B0F0", "7F7F7F", "ED7D31"
# Belastningsbaand-prioritet (hoejeste vinder). MOEDE og MS_BLUE/MS_GREY indgaar
# bevidst ikke her - moeder og milepaele paavirker aldrig belastningsbaandet.
NIVEAU = {RED: 5, ORANGE: 4, PURPLE: 3, NEUTRAL: 2, GREEN: 1}
LANE_FARVER = ["44546A", "2F4B6E", "3C5A80"]  # skiftes; foerste lane bruger [0]


def ferie(aar):
    """Ferie/helligdage der skraveres (svarer til projektkalenderen)."""
    a = aar                                            # paaskedag (Meeus/Butcher)
    g, c = a % 19, a // 100
    h = (c - c // 4 - (8 * c + 13) // 25 + 19 * g + 15) % 30
    i = h - (h // 28) * (1 - (h // 28) * (29 // (h + 1)) * ((21 - g) // 11))
    j = (a + a // 4 + i + 2 - c + c // 4) % 7
    l = i - j; mnd = 3 + (l + 40) // 44
    paaske = D(a, mnd, l + 28 - 31 * (mnd // 4))
    td = dt.timedelta
    uge = lambda w: (D.fromisocalendar(a, w, 1), D.fromisocalendar(a, w, 5))
    return [
        uge(7),                                        # vinterferie
        (paaske - td(3), paaske + td(1)),              # skaertorsdag - 2. paaskedag
        (paaske + td(39), paaske + td(39)),            # Kr. Himmelfart
        (paaske + td(50), paaske + td(50)),            # 2. pinsedag
        (uge(27)[0], uge(30)[1]),                      # sommerferie uge 27-30
        uge(42),                                       # efteraarsferie
        (D(a, 12, 21), D(a + 1, 1, 1)),                # jul/nytaar
    ]


FONT_STIER = {False: ["/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", "C:/Windows/Fonts/calibri.ttf"],
              True: ["/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf", "C:/Windows/Fonts/calibrib.ttf"]}
_fc = {}
def tw(text, pt, bold=False):
    """Tekstbredde i tommer (Calibri-metrik)."""
    k = (pt, bold)
    if k not in _fc:
        sti = next(p for p in FONT_STIER[bold] if os.path.exists(p))
        _fc[k] = ImageFont.truetype(sti, pt * 20)
    return _fc[k].getlength(text) / (72 * 20)


def linjer_i_bredde(tekst, bredde, pt, bold=False):
    """Estimeret antal linjer teksten fylder ved auto-wrap i en given bredde (tommer)."""
    if not tekst:
        return 0
    total = 0
    for del_ in str(tekst).split("\n"):
        w = tw(del_, pt, bold)
        total += max(1, math.ceil(w / bredde)) if bredde > 0 else 1
    return total


def parse_dato(s):
    """Parser en dato fra en Excel-celle. Accepterer datetime.datetime,
    datetime.date, ISO-format ('YYYY-MM-DD') og dansk/MS Project-format
    ('DD-MM-YYYY', ogsaa med 2-cifret aar). MS Project-udtraek indeholder ofte
    en engelsk ugedag foran datoen (fx 'Tue 07-10-25') - ugedagen ignoreres,
    og tallene efter den fortolkes altid som dag-maaned-aar (aldrig
    maaned-dag-aar). Kaster ValueError med den oprindelige vaerdi hvis ingen
    af formaterne matcher, eller tallene ikke udgoer en gyldig dato
    (fx 31. februar)."""
    if isinstance(s, dt.datetime):
        return s.date()
    if isinstance(s, dt.date):
        return s
    if s is None:
        raise ValueError("Ugyldig dato: None - forventede YYYY-MM-DD eller DD-MM-YYYY.")
    tekst = str(s).strip()
    iso = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", tekst)
    if iso:
        y, mth, d = (int(g) for g in iso.groups())
    else:
        # Dansk/MS Project-format, evt. med engelsk ugedagspraefiks (fx "Tue 07-10-25"),
        # som ignoreres - re.search finder dato-moensteret uanset hvad der staar foran det.
        dansk = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2,4})", tekst)
        if not dansk:
            raise ValueError(f"Ugyldig dato: {s!r} - forventede YYYY-MM-DD eller DD-MM-YYYY "
                             "(evt. med ugedag foran, fx 'Tue 07-10-25').")
        d, mth, y = (int(g) for g in dansk.groups())
        y = y + 2000 if y < 100 else y
    try:
        return D(y, mth, d)
    except ValueError as ex:
        raise ValueError(f"Ugyldig dato: {s!r} ({ex}).") from ex


def parse_varighed(v):
    """Parser Duration fra en Excel-celle til et helt antal dage. Accepterer
    tal (int/float, som Excel kan levere hvis kolonnen ikke er tekstformateret)
    og tekststrenge som '0 days', '1 day', '5 days' samt de danske varianter
    '0 dage', '1 dag', '5 dage'. Kaster ValueError med den oprindelige vaerdi
    hvis der ikke kan udledes et heltal."""
    if isinstance(v, bool):
        raise ValueError(f"Ugyldig varighed: {v!r} - forventede et heltal, evt. med enhed "
                         "(fx '5 days' eller '5 dage').")
    if isinstance(v, (int, float)):
        return int(v)
    tekst = "" if v is None else str(v).strip()
    m = re.match(r"(\d+)", tekst)
    if not m:
        raise ValueError(f"Ugyldig varighed: {v!r} - forventede et heltal, evt. med enhed "
                         "(fx '5 days' eller '5 dage').")
    return int(m.group(1))


def _celle_farve(cell, raekke, navn, advarsler):
    """Returnerer cellens semantiske fyldfarve som 6-cifret hex (fx 'FF0000'),
    eller 'FFFFFF' hvis cellen ikke har nogen (fortolkelig) fyldfarve. Haandterer
    baade 6- og 8-cifrede (ARGB) RGB-vaerdier fra openpyxl. Tema- eller
    indekserede farver (celle.fill.fgColor.type == 'theme'/'indexed') har ingen
    brugbar RGB-streng og kan derfor ikke fortolkes semantisk - de behandles som
    neutrale, og der tilfoejes en advarsel til advarselslisten, saa det er
    synligt at farven ikke kunne laeses."""
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return "FFFFFF"
    fg = fill.fgColor
    if fg is None:
        return "FFFFFF"
    if getattr(fg, "type", None) == "rgb":
        rgb = fg.rgb
        if isinstance(rgb, str) and len(rgb) >= 6:
            return rgb[-6:].upper()
        return "FFFFFF"
    advarsler.append(
        f"Excel-række {raekke} (\"{navn}\"): cellefarven er en tema-/indekseret farve uden "
        "fortolkelig RGB-vaerdi - behandlet som neutral.")
    return "FFFFFF"


def laes_plan(sti, konfig=None):
    """Laeser et MS Project Excel-udtraek (header i raekke 2, data fra raekke 3;
    kolonne B-E = Task Name/Duration/Start/Finish er paakraevede, F/G laeses
    ikke) og returnerer (lanes, advarsler). Kaster ValueError med Excel-
    raekkenummer, kolonne/felt og den oprindelige vaerdi ved ugyldigt input."""
    konfig = konfig or Konfig()
    ws = load_workbook(sti).active
    if ws.max_column < 5:
        raise ValueError("Excel-arket mangler en eller flere af de påkrævede kolonner B-E "
                         "(Task Name, Duration, Start, Finish).")
    advarsler = []
    lanes, cur = [], None
    for r in range(3, ws.max_row + 1):
        c = ws.cell(r, 2)
        navn_raw = c.value
        c3, c4, c5 = ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value
        navn_tom = navn_raw is None or str(navn_raw).strip() == ""
        if navn_tom and all(v is None or str(v).strip() == "" for v in (c3, c4, c5)):
            continue                                     # helt blank raekke - ignoreres
        if navn_tom:
            raise ValueError(f"Excel-række {r}: Task Name (kolonne B) mangler.")
        raw = str(navn_raw)
        navn = raw.strip()
        indrykket = bool(re.match(r"^\s", raw))
        fed = bool(c.font.b)

        varighed_raw = c3
        try:
            varighed = parse_varighed(varighed_raw)
        except ValueError as ex:
            raise ValueError(f"Excel-række {r} (\"{navn}\"): kolonne C (Duration) - {ex}") from ex
        milepael = varighed == 0

        start_raw = c4
        if start_raw is None or str(start_raw).strip() == "":
            raise ValueError(f"Excel-række {r} (\"{navn}\"): kolonne D (Start) mangler en startdato.")
        try:
            s = parse_dato(start_raw)
        except ValueError as ex:
            raise ValueError(f"Excel-række {r} (\"{navn}\"): kolonne D (Start) - {ex}") from ex

        e_raw = c5
        if milepael:
            try:
                e = parse_dato(e_raw)
            except ValueError:
                e = None   # slutdato bruges ikke for en milepael, saa en ugyldig/manglende vaerdi er uden betydning
        else:
            if e_raw is None or str(e_raw).strip() == "":
                raise ValueError(f"Excel-række {r} (\"{navn}\"): kolonne E (Finish) mangler en slutdato.")
            try:
                e = parse_dato(e_raw)
            except ValueError as ex:
                raise ValueError(f"Excel-række {r} (\"{navn}\"): kolonne E (Finish) - {ex}") from ex
            if e < s:
                raise ValueError(f"Excel-række {r} (\"{navn}\"): slutdato ({e.isoformat()}) ligger "
                                 f"før startdato ({s.isoformat()}).")

        fill = _celle_farve(c, r, navn, advarsler)
        er_moede = fill == MOEDE
        blaa = (not er_moede) and (navn.startswith("Milepæl:") or fill == MS_BLUE)
        vis = re.sub(r"^Milepæl:\s*", "", navn)
        vis = konfig.korte_navne.get(vis, vis)
        if milepael:
            farve = MOEDE if er_moede else (MS_BLUE if blaa else MS_GREY)
            task = (vis, s, None, farve, blaa)
        else:
            farve = fill if fill in NIVEAU else (MOEDE if er_moede else NEUTRAL)
            task = (vis, s, e, farve, False)
        if not indrykket and fed:                       # fase-raekke
            cur = dict(key=navn, name=konfig.fase_navne.get(navn, navn), s=s, e=e, tasks=[], summary=task)
            lanes.append(cur)
        elif indrykket:
            if cur is None:
                raise ValueError(f"Excel-række {r}: aktiviteten \"{navn}\" er indrykket, men der er "
                                 "endnu ikke defineret nogen fase at knytte den til. Tjek at "
                                 "faserækken ovenfor er fed skrift og ikke selv indrykket.")
            cur["tasks"].append(task)
        else:                                           # ikke-indrykket enkeltraekke
            if cur is None or cur.get("loose") is None:
                cur = dict(key=navn, name=konfig.fase_navne.get(navn, navn), s=s, e=e or s, tasks=[], loose=True)
                lanes.append(cur)
            cur["tasks"].append(task)
            cur["e"] = max(cur["e"], e or s)
    for l in lanes:   # fase uden underaktiviteter -> vis fasen selv som bjaelke
        if not l["tasks"] and "summary" in l:
            l["tasks"].append(l["summary"])
        for fase, nm, d, blaa in konfig.ekstra_milepaele:
            if l["key"] == fase:
                l["tasks"].append((nm, d, None, MS_BLUE if blaa else MS_GREY, blaa))
    return lanes, advarsler


def byg(lanes, idag, konfig=None):
    konfig = konfig or Konfig()
    alle = [t for l in lanes for t in l["tasks"]]
    if not alle:
        raise ValueError("Tidsplanen indeholder ingen aktiviteter at tegne - tjek input-filen.")
    sidste = max((t[2] or t[1]) for t in alle)
    # max(sidste, idag) sikrer at "i dag"-linjen altid ligger inden for tidsaksen,
    # ogsaa hvis statusdatoen ligger efter planens sidste relevante aktivitet.
    seneste_relevante = max(sidste, idag)
    A0 = D(idag.year, idag.month, 1)
    m = D(seneste_relevante.year + (seneste_relevante.month == 12), seneste_relevante.month % 12 + 1, 1)
    A1 = D(m.year + (m.month == 12), m.month % 12 + 1, 1)      # en maaned luft efter sidste relevante dato

    W, H = 13.333, 7.5
    LX, LW, TX0, TX1 = 0.25, 1.30, 1.62, 13.10
    dage_i_aksen = (A1 - A0).days
    if dage_i_aksen <= 0:
        raise ValueError("Tidsaksen kunne ikke beregnes (ugyldigt datointerval mellem statusdato og plan).")
    PX = (TX1 - TX0) / dage_i_aksen
    X = lambda d: TX0 + (max(A0, min(A1, d)) - A0).days * PX
    Y_YEAR, H_YEAR = 0.62, 0.20
    Y_MON, H_MON = Y_YEAR + H_YEAR, 0.22
    Y_LOAD, H_LOAD = Y_MON + H_MON + 0.05, 0.13
    Y_LANES, Y_BUND, GAP_L = Y_LOAD + H_LOAD + 0.08, 7.36, 0.04
    nlanes = len(lanes)
    nrows = sum(len(l["tasks"]) for l in lanes)
    ROW = (Y_BUND - 0.08 - Y_LANES - GAP_L * (nlanes - 1)) / nrows
    if ROW <= 0:
        raise ValueError(f"Tidsplanen har for mange aktiviteter ({nrows}) til at få plads på én "
                         "slide - der kan ikke beregnes en positiv rækkehøjde. Planens indhold "
                         "bevares ikke automatisk kortere; layoutet skal udvides i stedet.")
    PT_NAME, PT_DATE, BAR_H, GAP = 8, 7, ROW * 0.62, 0.05
    advarsler = []
    if ROW < 0.15:
        advarsler.append(f"Rækkehøjde {ROW:.3f}\" er meget lav - planen er meget tæt med {nrows} "
                         f"aktiviteter i {nlanes} faser. Alle aktiviteter vises stadig, men overvej "
                         "et bredere layout eller kortere fasenavne, hvis det er svært at læse.")

    dm = lambda d: f"{d.day}/{d.month}"
    out = []
    def rect(x, y, w, h, fill, **k): out.append(dict(t="rect", x=x, y=y, w=w, h=h, fill=fill, **k))
    def text(x, y, w, h, s, **k): out.append(dict(t="text", x=x, y=y, w=w, h=h, s=s, **k))
    def diamond(cx, cy, sz, fill): out.append(dict(t="diamond", x=cx - sz/2, y=cy - sz/2, w=sz, h=sz, fill=fill))
    def line(x1, y1, x2, y2, color, width=0.75, dash=None):
        out.append(dict(t="line", x=x1, y=y1, w=x2 - x1, h=y2 - y1, color=color, width=width, dash=dash))

    # Titel
    text(LX, 0.14, 7.5, 0.32, konfig.titel, size=18, bold=True, color="1F2A3A")
    text(LX, 0.43, 7.5, 0.18,
         f"Kilde: projektplan (MS Project) pr. {dm(idag)}-{str(idag.year)[2:]} · Farver angiver belastning{konfig.hos_kunde}",
         size=9, color="595959")

    # Aar- og maanedsbaand
    MON = ["jan","feb","mar","apr","maj","jun","jul","aug","sep","okt","nov","dec"]
    for y in range(A0.year, A1.year + 1):
        y0, y1 = max(A0, D(y, 1, 1)), min(A1, D(y + 1, 1, 1))
        if y1 <= y0: continue
        rect(X(y0), Y_YEAR, X(y1) - X(y0) - 0.01, H_YEAR, "44546A")
        text(X(y0) + 0.05, Y_YEAR, 1, H_YEAR, str(y), size=8, bold=True, color="FFFFFF", valign="middle")
    months, m = [], A0
    while m < A1:
        nm = D(m.year + (m.month == 12), m.month % 12 + 1, 1); months.append(m)
        rect(X(m), Y_MON, X(nm) - X(m) - 0.01, H_MON, "5B6F87")
        text(X(m), Y_MON, X(nm) - X(m), H_MON, MON[m.month - 1], size=8, color="FFFFFF", align="center", valign="middle")
        m = nm

    # Belastningsbaand - moeder (ED7D31) og milepaele (e is None) paavirker aldrig baandet
    bars = [t for t in alle
            if t[2] is not None and t[3] != MOEDE and t[0] not in konfig.ikke_i_belastning]
    wk = A0 - dt.timedelta(days=A0.weekday())
    while wk < A1:
        we = wk + dt.timedelta(days=4)
        act = [t[3] for t in bars if t[1] <= we and t[2] >= wk]
        if act:
            x0, x1 = X(wk), X(wk + dt.timedelta(days=7))
            if x1 > x0: rect(x0, Y_LOAD, x1 - x0 - 0.008, H_LOAD, max(act, key=lambda c: NIVEAU[c]))
        wk += dt.timedelta(days=7)
    text(LX, Y_LOAD - 0.035, LW, H_LOAD + 0.07, konfig.label_belastning, size=7, italic=True, color="404040",
         align="right", valign="middle")

    # Lanes
    geo, y = [], Y_LANES
    lane_def = []
    for i, l in enumerate(lanes):
        sub = (f"{l['s'].day}/{l['s'].month}-{str(l['s'].year)[2:]}–{dm(l['e'])}" if l["s"] < A0
               else f"{dm(l['s'])}–{dm(l['e'])}")
        col = LANE_FARVER[0] if (i == 0 or l.get("loose")) else LANE_FARVER[1 + (i - 1) % 2]
        lane_def.append(("std", l["name"], sub, col, len(l["tasks"]), l["tasks"]))
    for i, (kind, name, sub, col, n, tasks) in enumerate(lane_def):
        h = n * ROW
        rect(TX0, y, TX1 - TX0, h, "F2F2F2" if i % 2 == 0 else "E7EBF0")
        rect(LX, y, LW, h, col)
        if kind == "std":
            linjer = linjer_i_bredde(name, LW - 0.12, 9, True) + linjer_i_bredde(sub, LW - 0.12, 7)
            if linjer * 0.145 > h + 0.02:
                advarsler.append(f"Fasenavnet \"{name.replace(chr(10), ' ')}\" fylder muligvis for meget i venstre "
                                 "kolonne - overvej et kortere navn via 'fase_navne' i konfigurationen.")
        geo.append((y, h, kind, name, sub, tasks)); y += h + GAP_L
    Y_END = y - GAP_L

    for aar in range(A0.year - 1, A1.year + 1):
        for s, e in ferie(aar):
            if e < A0 or s >= A1: continue
            rect(X(s), Y_LANES, max(X(e + dt.timedelta(days=1)) - X(s), 0.03), Y_END - Y_LANES, "C5CFDB", transparency=35)
    for m in months[1:]:
        line(X(m), Y_LANES, X(m), Y_END, "FFFFFF", 1.0)
    for (ly, h, kind, name, sub, tasks) in geo:
        text(LX + 0.06, ly, LW - 0.12, h, [(name, dict(bold=True, size=9, color="FFFFFF")),
                                            (("\n" + sub) if sub else "", dict(size=7, color="E0E6EE"))],
             valign="middle", wrap=True)

    def labels(ry, bx0, bx1, name, dates, bold=False):
        nw, dw = tw(name, PT_NAME, bold), tw(dates, PT_DATE)
        if bx1 + GAP + nw <= TX1 and bx0 - GAP - dw >= TX0:
            text(bx0 - GAP - dw - 0.02, ry, dw + 0.02, ROW, dates, size=PT_DATE, color="595959", align="right", valign="middle")
            text(bx1 + GAP, ry, nw + 0.05, ROW, name, size=PT_NAME, color="1A1A1A", bold=bold, valign="middle")
        elif bx1 + GAP + nw <= TX1:
            text(bx1 + GAP, ry, nw + tw("   ", PT_NAME) + dw + 0.08, ROW,
                 [(name + "   ", dict(size=PT_NAME, color="1A1A1A", bold=bold)), (dates, dict(size=PT_DATE, color="595959"))],
                 valign="middle")
        else:
            if bx0 - GAP - nw < TX0:
                advarsler.append(f"Navn for langt - tilfoej til korte_navne i konfigurationen: {name}")
            text(bx0 - GAP - nw - 0.03, ry, nw + 0.03, ROW, name, size=PT_NAME, color="1A1A1A", bold=bold, align="right", valign="middle")
            text(bx1 + GAP, ry, dw + 0.05, ROW, dates, size=PT_DATE, color="595959", valign="middle")

    for (ly, h, kind, name, sub, tasks) in geo:
        for j, (nm, s, e, col, blaa) in enumerate(tasks):
            ry = ly + j * ROW; cy = ry + ROW / 2
            if e is None:
                sz = ROW * 0.72; cx = X(s) + PX / 2
                diamond(cx, cy, sz, col)
                labels(ry, cx - sz / 2, cx + sz / 2, nm, dm(s), bold=blaa)
            else:
                x0, x1 = X(s), X(e + dt.timedelta(days=1))
                x1 = max(x1, x0 + 0.05)
                rect(x0, cy - BAR_H / 2, x1 - x0, BAR_H, col)
                dates = f"{s.day}/{s.month}-{str(s.year)[2:]}–{dm(e)}" if s < A0 else f"{dm(s)}–{dm(e)}"
                labels(ry, x0, x1, nm, dates)

    xt = X(idag) + PX / 2
    line(xt, Y_MON, xt, Y_END, "C00000", 1.25, dash="dash")
    text(xt + 0.05, Y_LOAD - 0.02, 1.0, H_LOAD + 0.04, "I dag", size=7, bold=True, color="C00000", valign="middle")

    COLS = [[("rect", RED, f"Spidsbelastning{konfig.suffix_kunde}"), ("rect", ORANGE, "Belastning"), ("rect", GREEN, "Bolden hos tilbudsgivere/eksterne")],
            [("rect", PURPLE, "Godkendelse (STG/DB)"), ("rect", NEUTRAL, "Ingen belastningsmarkering"), ("rect", MOEDE, "Møde")],
            [("dia", MS_BLUE, "Milepæl"), ("dia", MS_GREY, "Frist/modtagelse"), ("hol", "C5CFDB", "Ferie/helligdag")],
            [("today", None, "I dag")]]
    LPT, SW, CG = 7.5, 0.2, 0.18
    cw = [SW + 0.06 + max(tw(i[2], LPT) for i in c) + 0.03 for c in COLS]
    lx = TX1 - sum(cw) - CG * (len(COLS) - 1)
    for c, w_ in zip(COLS, cw):
        for r, (kind, col, lbl) in enumerate(c):
            ly = 0.1 + r * 0.155
            if kind == "rect": rect(lx, ly + 0.035, SW, 0.08, col)
            elif kind == "dia": diamond(lx + SW / 2, ly + 0.075, 0.11, col)
            elif kind == "hol": rect(lx, ly + 0.005, SW, 0.14, col)
            else: line(lx + SW / 2, ly, lx + SW / 2, ly + 0.15, "C00000", 1.25, dash="dash")
            text(lx + (SW / 2 + 0.08 if kind == "today" else SW + 0.06), ly, w_, 0.15, lbl, size=LPT, color="404040", valign="middle")
        lx += w_ + CG
    return out, advarsler


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("plan")
    ap.add_argument("--idag", required=True, help="Statusdato YYYY-MM-DD")
    ap.add_argument("-o", "--out", default="Tidsplan_visuel.pptx")
    ap.add_argument("--config", help="Valgfri JSON-konfigurationsfil (kunde, udbud, fase-navne mv.)")
    ap.add_argument("--kunde", help="Kundenavn - overskriver evt. vaerdi fra --config")
    ap.add_argument("--udbud", help="Udbuddets navn - overskriver evt. vaerdi fra --config")
    a = ap.parse_args()

    konfig = load_konfig(a.config) if a.config else Konfig()
    if a.kunde: konfig.kunde = a.kunde
    if a.udbud: konfig.udbud = a.udbud

    lanes, laes_adv = laes_plan(a.plan, konfig)
    prims, byg_adv = byg(lanes, D.fromisoformat(a.idag), konfig)
    from pptx_render import render
    render(prims, a.out)
    print(f"Skrevet: {a.out}  ({sum(len(l['tasks']) for l in lanes)} aktiviteter i {len(lanes)} faser)")
    for w in laes_adv + byg_adv: print("ADVARSEL:", w)
