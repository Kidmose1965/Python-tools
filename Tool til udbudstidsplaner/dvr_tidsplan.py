"""
DVR visuel tidsplan - template
Brug:  python dvr_tidsplan.py <plan.xlsx> --idag 2026-09-10 [-o DVR_tidsplan_visuel.pptx]

Input: Excel-udtraek fra MS Project (kolonner B-E: Aktivitet, Laengde, Start, Slut; header i raekke 2).
  - Fase-raekker: fed, ikke-indrykket navn. Aktiviteter: indrykket med mellemrum.
  - Farve i kolonne B = LMST-belastning (se FARVER). Hvid = ingen markering (graa bjaelke).
  - 0 dage = milepael. "Milepael:"-praefiks eller blaa udfyldning -> blaa markoer, ellers graa.
Kraever: Python (openpyxl, Pillow) + Node med pptxgenjs (draw.js i samme mappe).
"""
import argparse, json, os, re, subprocess, sys, datetime as dt
from openpyxl import load_workbook
from PIL import ImageFont

D = dt.date
HERE = os.path.dirname(os.path.abspath(__file__))

# =====================================================================
# KONFIGURATION - ret her ved revisioner
# =====================================================================
TITEL = "Udbud af Den Veterinære Receptløsning (DVR) – tidsplan"

# Visningsnavne for faser (Excel-navn -> lane-label). Ukendte faser bruger Excel-navnet.
FASE_NAVNE = {
    "Udformning af udbudsmateriale": "Udbudsmateriale",
    "Prækvalifikationsfase": "Prækvalifikation",
    "Tilbudsfasen (indledende tilbud)": "Tilbudsfasen\n(indledende tilbud)",
    "Forhandlingsrunde 1": "Forhandlingsrunde 1",
    "Tilbudsfasen (endeligt tilbud)": "Tilbudsfasen\n(endeligt tilbud)",
    "Standstill": "Standstill",
}
# Forkortede aktivitetsnavne (Excel-navn uden "Milepæl:" -> visningsnavn)
KORTE_NAVNE = {
    "Udarbejdelse af afslagsbreve + opfordring til afgivelse af indledende tilbud":
        "Afslagsbreve + opfordring til afgivelse af indledende tilbud",
    "Forventet tidspunkt for fremsendelse af diskussionsliste og agenda mv. til Tilbudsgiver 1, 2 og 3":
        "Fremsendelse af diskussionsliste og agenda mv. til tilbudsgiver 1, 2 og 3",
    "Forventet tidspunkt for meddelelse om første reviderede tilbud samt udsendelse af evt. revideret udbudsmateriale":
        "Meddelelse om første reviderede tilbud samt udsendelse af evt. revideret udbudsmateriale",
    "Tilbudsgivere udarbejder tilbud (frist vurderes konkret I relation til karakteren af ændringer I udbudsmaterialet)":
        "Tilbudsgivere udarbejder endeligt tilbud (frist vurderes konkret ift. ændringer i udbudsmaterialet)",
    "Evaluering af tilbud og udarbejdelse af spørgsmål til forhandlingsrunden (beslutning om forhandling)":
        "Evaluering af tilbud og udarbejdelse af spørgsmål til forhandlingsrunden (beslutning om forhandling)",
}
# Ekstra milepaele der ikke staar i MS Project-planen: (fase i Excel, navn, dato, blaa?)
EKSTRA_MILEPAELE = [
    ("Udformning af udbudsmateriale", "Offentliggørelse af udbud", D(2026, 10, 2), True),
]
# Aktiviteter der ikke taeller med i belastningsbaandet
IKKE_I_BELASTNING = {"Udformning af udbudsmateriale"}

# Lovgivningsspor (kilde: Evas plan 04-09-26). Saet LOV_VIS = False for at udelade.
LOV_VIS = True
LOV_BEKRAEFTET = False
LOV_BARS = [
    ("LMST har lovhjemmel på arbejdsprogram (1/10–31/12)", D(2026, 10, 1), D(2026, 12, 31)),
    ("LMST og DEP forbereder lovhjemmel til lovprogrammet (1/1–10/9)", D(2027, 1, 1), D(2027, 9, 10)),
]
LOV_MS = [("DEP har givet tilsagn om lovhjemmel på arbejdsprogrammet", D(2027, 3, 10)),
          ("DEP har godkendt lovhjemmel til lovprogrammet", D(2027, 9, 10))]
LOV_HALE = ("Forhandling og vedtagelse i Folketinget (10/9-27 – 1/6-28) · Lovhjemmel godkendt i FT 1/1-28",
            D(2027, 9, 10))   # bjaelke fra dato til hoejre kant med pil

# Farver (Birgers koder) -> belastningsniveau
RED, ORANGE, GREEN, PURPLE = "FF0000", "FFC000", "92D050", "7030A0"
NEUTRAL, MS_BLUE, MS_GREY, LOV = "A6A6A6", "00B0F0", "7F7F7F", "B4C7E7"
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
# =====================================================================

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


def parse_dato(s):
    if isinstance(s, dt.datetime): return s.date()
    m = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2,4})", str(s))
    d, mth, y = map(int, m.groups())
    return D(y + 2000 if y < 100 else y, mth, d)


def laes_plan(sti):
    ws = load_workbook(sti).active
    lanes, cur = [], None
    for r in range(3, ws.max_row + 1):
        c = ws.cell(r, 2)
        if c.value is None or str(c.value).strip() == "": continue
        raw = str(c.value)
        navn = raw.strip()
        indrykket = raw.startswith(" ")
        fed = bool(c.font.b)
        varighed = int(re.match(r"\s*(\d+)", str(ws.cell(r, 3).value)).group(1))
        s, e = parse_dato(ws.cell(r, 4).value), parse_dato(ws.cell(r, 5).value)
        fill = (c.fill.fgColor.rgb or "FFFFFFFF")[-6:] if c.fill.fill_type == "solid" else "FFFFFF"
        milepael = varighed == 0
        blaa = navn.startswith("Milepæl:") or fill == MS_BLUE
        vis = re.sub(r"^Milepæl:\s*", "", navn)
        vis = KORTE_NAVNE.get(vis, vis)
        if milepael:
            task = (vis, s, None, MS_BLUE if blaa else MS_GREY, blaa)
        else:
            task = (vis, s, e, fill if fill in NIVEAU else NEUTRAL, False)
        if not indrykket and fed:                       # fase-raekke
            cur = dict(key=navn, name=FASE_NAVNE.get(navn, navn), s=s, e=e, tasks=[], summary=task)
            lanes.append(cur)
        elif indrykket and cur is not None:
            cur["tasks"].append(task)
        else:                                           # ikke-indrykket enkeltraekke
            if cur is None or cur.get("loose") is None:
                cur = dict(key=navn, name=FASE_NAVNE.get(navn, navn), s=s, e=e or s, tasks=[], loose=True)
                lanes.append(cur)
            cur["tasks"].append(task)
            cur["e"] = max(cur["e"], e or s)
    for l in lanes:   # fase uden underaktiviteter -> vis fasen selv som bjaelke
        if not l["tasks"] and "summary" in l:
            l["tasks"].append(l["summary"])
        for fase, nm, d, blaa in EKSTRA_MILEPAELE:
            if l["key"] == fase:
                l["tasks"].append((nm, d, None, MS_BLUE if blaa else MS_GREY, blaa))
    return lanes


def byg(lanes, idag):
    alle = [t for l in lanes for t in l["tasks"]]
    sidste = max((t[2] or t[1]) for t in alle)
    A0 = D(idag.year, idag.month, 1)
    m = D(sidste.year + (sidste.month == 12), sidste.month % 12 + 1, 1)
    A1 = D(m.year + (m.month == 12), m.month % 12 + 1, 1)      # en maaned luft efter sidste dato

    W, H = 13.333, 7.5
    LX, LW, TX0, TX1 = 0.25, 1.30, 1.62, 13.10
    PX = (TX1 - TX0) / (A1 - A0).days
    X = lambda d: TX0 + (max(A0, min(A1, d)) - A0).days * PX
    Y_YEAR, H_YEAR = 0.62, 0.20
    Y_MON, H_MON = Y_YEAR + H_YEAR, 0.22
    Y_LOAD, H_LOAD = Y_MON + H_MON + 0.05, 0.13
    Y_LANES, Y_BUND, GAP_L = Y_LOAD + H_LOAD + 0.08, 7.36, 0.04
    lov = LOV_VIS
    nlanes = len(lanes) + lov
    nrows = 3 * lov + sum(len(l["tasks"]) for l in lanes)
    ROW = (Y_BUND - 0.08 - Y_LANES - GAP_L * (nlanes - 1)) / nrows
    PT_NAME, PT_DATE, BAR_H, GAP = 8, 7, ROW * 0.62, 0.05
    advarsler = []
    if ROW < 0.15: advarsler.append(f"Raekkehoejde {ROW:.3f}\" er lav - overvej at forkorte planen.")

    dm = lambda d: f"{d.day}/{d.month}"
    out = []
    def rect(x, y, w, h, fill, **k): out.append(dict(t="rect", x=x, y=y, w=w, h=h, fill=fill, **k))
    def text(x, y, w, h, s, **k): out.append(dict(t="text", x=x, y=y, w=w, h=h, s=s, **k))
    def diamond(cx, cy, sz, fill): out.append(dict(t="diamond", x=cx - sz/2, y=cy - sz/2, w=sz, h=sz, fill=fill))
    def line(x1, y1, x2, y2, color, width=0.75, dash=None):
        out.append(dict(t="line", x=x1, y=y1, w=x2 - x1, h=y2 - y1, color=color, width=width, dash=dash))

    # Titel
    text(LX, 0.14, 7.5, 0.32, TITEL, size=18, bold=True, color="1F2A3A")
    text(LX, 0.43, 7.5, 0.18, f"Kilde: projektplan (MS Project) pr. {dm(idag)}-{str(idag.year)[2:]} · Farver angiver LMST's belastning",
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

    # Belastningsbaand
    bars = [t for t in alle if t[2] is not None and t[0] not in IKKE_I_BELASTNING]
    wk = A0 - dt.timedelta(days=A0.weekday())
    while wk < A1:
        we = wk + dt.timedelta(days=4)
        act = [t[3] for t in bars if t[1] <= we and t[2] >= wk]
        if act:
            x0, x1 = X(wk), X(wk + dt.timedelta(days=7))
            if x1 > x0: rect(x0, Y_LOAD, x1 - x0 - 0.008, H_LOAD, max(act, key=lambda c: NIVEAU[c]))
        wk += dt.timedelta(days=7)
    text(LX, Y_LOAD - 0.035, LW, H_LOAD + 0.07, "LMST-belastning pr. uge", size=7, italic=True, color="404040",
         align="right", valign="middle")

    # Lanes
    geo, y = [], Y_LANES
    lane_def = ([("lov", "Lovgivning", ("Datoer fra Evas plan\n– ikke bekræftet" if not LOV_BEKRAEFTET else ""), "8497B0", 3, [])]
                if lov else [])
    for i, l in enumerate(lanes):
        sub = (f"{l['s'].day}/{l['s'].month}-{str(l['s'].year)[2:]}–{dm(l['e'])}" if l["s"] < A0
               else f"{dm(l['s'])}–{dm(l['e'])}")
        col = LANE_FARVER[0] if (i == 0 or l.get("loose")) else LANE_FARVER[1 + (i - 1) % 2]
        lane_def.append(("std", l["name"], sub, col, len(l["tasks"]), l["tasks"]))
    for i, (kind, name, sub, col, n, tasks) in enumerate(lane_def):
        h = n * ROW
        rect(TX0, y, TX1 - TX0, h, "F2F2F2" if i % 2 == 0 else "E7EBF0")
        rect(LX, y, LW, h, col)
        geo.append((y, h, kind, name, sub, tasks)); y += h + GAP_L
    Y_END = y - GAP_L

    for aar in range(A0.year - 1, A1.year + 1):
        for s, e in ferie(aar):
            if e < A0 or s >= A1: continue
            rect(X(s), Y_LANES, max(X(e + dt.timedelta(days=1)) - X(s), 0.03), Y_END - Y_LANES, "C5CFDB", transparency=35)
    for m in months[1:]:
        line(X(m), Y_LANES, X(m), Y_END, "FFFFFF", 1.0)
    for (ly, h, kind, name, sub, tasks) in geo:
        text(LX + 0.06, ly, LW - 0.12, h, [(name, dict(bold=True, size=8.5 if kind == "lov" else 9, color="FFFFFF")),
                                            (("\n" + sub) if sub else "", dict(size=7, color="E0E6EE", italic=(kind == "lov")))],
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
                advarsler.append(f"Navn for langt - tilfoej til KORTE_NAVNE: {name}")
            text(bx0 - GAP - nw - 0.03, ry, nw + 0.03, ROW, name, size=PT_NAME, color="1A1A1A", bold=bold, align="right", valign="middle")
            text(bx1 + GAP, ry, dw + 0.05, ROW, dates, size=PT_DATE, color="595959", valign="middle")

    for (ly, h, kind, name, sub, tasks) in geo:
        if kind == "lov":
            r0, r1, r2 = ly, ly + ROW, ly + 2 * ROW
            for lbl, s, e in LOV_BARS:
                x0, x1 = X(s), X(e + dt.timedelta(days=1))
                rect(x0, r0 + (ROW - BAR_H * 1.25) / 2, x1 - x0, BAR_H * 1.25, LOV)
                text(x0 + 0.05, r0, x1 - x0 - 0.1, ROW, lbl, size=7, color="1F3864", valign="middle")
            for i2, (lbl, d) in enumerate(LOV_MS):
                cx, cy = X(d), r1 + ROW / 2
                out.append(dict(t="tri", x=cx - 0.05, y=cy - 0.06, w=0.1, h=0.12, fill="5B6F87"))
                nw_, dw_ = tw(lbl, 7.5) + 0.05, tw(dm(d), 7) + 0.03
                if i2 < len(LOV_MS) - 1:
                    text(cx - 0.09 - dw_, r1, dw_, ROW, dm(d), size=7, color="595959", align="right", valign="middle")
                    text(cx + 0.09, r1, nw_, ROW, lbl, size=7.5, color="1F3864", valign="middle")
                else:
                    text(cx - 0.09 - nw_, r1, nw_, ROW, lbl, size=7.5, color="1F3864", align="right", valign="middle")
                    text(cx + 0.09, r1, dw_, ROW, dm(d), size=7, color="595959", valign="middle")
            lbl, d = LOV_HALE
            x0 = X(d)
            rect(x0, r2 + (ROW - BAR_H * 1.25) / 2, TX1 - x0 - 0.02, BAR_H * 1.25, LOV)
            out.append(dict(t="arrowhead", x=TX1 - 0.02, y=r2 + (ROW - BAR_H * 1.9) / 2, w=0.12, h=BAR_H * 1.9, fill=LOV))
            wdt = tw(lbl + "  →", 7.5) + 0.05
            text(x0 - 0.06 - wdt, r2, wdt, ROW, lbl + "  →", size=7.5, color="1F3864", align="right", valign="middle")
            continue
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

    COLS = [[("rect", RED, "Spidsbelastning LMST"), ("rect", ORANGE, "Belastning"), ("rect", GREEN, "Bolden hos tilbudsgivere/eksterne")],
            [("rect", PURPLE, "Godkendelse (STG/DB)"), ("rect", NEUTRAL, "Ingen belastningsmarkering")]
            + ([("rect", LOV, "Lovgivning" + ("" if LOV_BEKRAEFTET else " (ikke bekræftet)"))] if lov else []),
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
    ap.add_argument("-o", "--out", default="DVR_tidsplan_visuel.pptx")
    a = ap.parse_args()
    lanes = laes_plan(a.plan)
    prims, adv = byg(lanes, D.fromisoformat(a.idag))
    pj = os.path.join(os.path.dirname(os.path.abspath(a.out)), "_prims.json")
    json.dump(prims, open(pj, "w", encoding="utf-8"), ensure_ascii=False)
    subprocess.run(["node", os.path.join(HERE, "draw.js"), pj, a.out], check=True)
    os.remove(pj)
    print(f"Skrevet: {a.out}  ({sum(len(l['tasks']) for l in lanes)} aktiviteter i {len(lanes)} faser)")
    for w in adv: print("ADVARSEL:", w)
