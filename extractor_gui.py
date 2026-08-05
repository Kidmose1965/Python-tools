#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor GUI - knap-interface til extractor.py
================================================
Læg denne fil i SAMME mappe som extractor.py og start den med:
    python extractor_gui.py
(eller dobbeltklik på filen, hvis .py er knyttet til Python)

Kræver kun Python + openpyxl - tkinter følger med Python.
"""

import os
import sys
import traceback
from pathlib import Path
from tkinter import (Tk, Toplevel, Frame, Label, Button, Listbox, Scrollbar,
                     filedialog, messagebox, BooleanVar, Checkbutton,
                     MULTIPLE, EXTENDED, END, BOTH, LEFT, RIGHT, X, Y)

try:
    import extractor as ex
except ImportError:
    print("FEJL: extractor.py skal ligge i samme mappe som denne fil.")
    sys.exit(1)

BG = "#F0F4F8"
KOMMENTAR_TYPER = [
    ("Alle understøttede dokumenter", "*.docx *.xlsx"),
    ("Word-dokumenter", "*.docx"),
    ("Excel-filer", "*.xlsx"),
]

# Rambøll-inspireret farvepalet
OXFORD = "#2D3748"        # mørk header/tekst
CYAN = "#009DE0"          # accent
WHITE = "#FFFFFF"         # baggrund
GRAA = "#718096"          # sekundær tekst/linjer
HOVER = "#E6F6FD"         # lys cyan ved hover
HEADER_UNDERTEKST = "#CBD5E0"
HEADER_VERSION = "#A0AEC0"
STATUS_BG = "#EDF2F7"


class App:
    def __init__(self, root):
        self.root = root
        root.title("Extractor - udtræk fra kravspecifikationer")
        root.configure(bg=WHITE)
        root.resizable(False, False)
        try:
            root.iconbitmap(str(Path(__file__).resolve().parent / "icon.ico"))
        except Exception:
            pass

        # ---- header -----------------------------------------------------
        header = Frame(root, bg=OXFORD)
        header.pack(fill=X)
        header_venstre = Frame(header, bg=OXFORD)
        header_venstre.pack(side=LEFT, anchor="w", padx=20, pady=(16, 14))
        Label(header_venstre, text="Extractor", bg=OXFORD, fg=WHITE,
              font=("Segoe UI", 18, "bold")).pack(anchor="w")
        Label(header_venstre, text="Kvalitetssikring af udbudsdokumenter",
              bg=OXFORD, fg=HEADER_UNDERTEKST,
              font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))
        Label(header, text="v2.0 – juni 2026", bg=OXFORD, fg=HEADER_VERSION,
              font=("Segoe UI", 8)).pack(side=RIGHT, anchor="n", padx=16, pady=(14, 0))

        # ---- statuslinje og kreditering (bund - pakkes først så de "bunder") --
        self.status = Label(root, text="Klar.", bg=STATUS_BG, fg=OXFORD,
                            font=("Segoe UI", 9), anchor="w", padx=14, pady=6)
        self.status.pack(fill=X, side="bottom")

        Label(root, text="Udviklet af Birger Kidmose", bg=WHITE, fg=GRAA,
              font=("Segoe UI", 8), anchor="e").pack(
              fill=X, padx=18, pady=(4, 8), side="bottom")

        # ---- knapgrupper --------------------------------------------------
        krop = Frame(root, bg=WHITE)
        krop.pack(fill=BOTH, expand=True)

        self._gruppe(krop, "UDTRÆK", [
            ("Udtræk kommentarer", self.kommentarer),
            ("Udtræk kommentarer (kun dokumenter med fund)", self.kommentarer_filtreret),
            ("Udtræk trackchanges", self.trackchanges),
            ("Udtræk kundens inputfelter (grøn)", lambda: self.inputfelter("groen")),
            ("Udtræk tilbudsgivers inputfelter (gul)", lambda: self.inputfelter("gul")),
        ])
        self._gruppe(krop, "ANALYSE", [
            ("Skab kravmatrix", self.kravmatrix),
            ("Skab kravmatrix med kommentarer", self.kravmatrix_kommentarer),
            ("Krydstjek (uden semantisk kontrol)", self.krydstjek),
            ("Krydstjek med semantisk kontrol", self.krydstjek_semantik),
            ("Tjek interne henvisninger", self.intern_tjek),
        ])
        self._gruppe(krop, "HJÆLP", [
            ("Vis dokumentets typografier (styles)", self.styles),
        ])

        # Sørg for en luftig, ca. 420px bred rude uden at klippe indhold
        root.update_idletasks()
        bredde = max(420, root.winfo_reqwidth())
        højde = root.winfo_reqheight()
        root.geometry(f"{bredde}x{højde}")

    # ------------------------------------------------------------ design ---
    def _gruppe(self, parent, titel, knapper):
        """Tegner en gruppe knapper under en lille grå overskrift."""
        Label(parent, text=titel, bg=WHITE, fg=GRAA,
              font=("Segoe UI", 9, "bold"), anchor="w").pack(
              fill=X, padx=20, pady=(16, 6))
        for tekst, cmd in knapper:
            self._knap(parent, tekst, cmd)

    def _knap(self, parent, tekst, cmd):
        """Flad knap med tynd cyan kant og lys cyan hover (plain tkinter -
        ttk-knapper respekterer ikke altid custom baggrundsfarver på Windows)."""
        kant = Frame(parent, bg=CYAN)
        kant.pack(fill=X, padx=20, pady=4)
        knap = Button(kant, text=tekst, command=self.beskyt(cmd),
                     bg=WHITE, fg=OXFORD, activebackground=HOVER,
                     activeforeground=OXFORD, relief="flat", borderwidth=0,
                     highlightthickness=0, anchor="w", padx=14, pady=10,
                     font=("Segoe UI", 10), cursor="hand2")
        knap.pack(fill=X, padx=1, pady=1)
        knap.bind("<Enter>", lambda e: knap.config(bg=HOVER))
        knap.bind("<Leave>", lambda e: knap.config(bg=WHITE))
        return knap

    # ------------------------------------------------------------------ utils
    def beskyt(self, fn):
        """Fang alle fejl og vis dem pænt i stedet for at lukke vinduet."""
        def wrapper():
            try:
                fn()
            except Exception:
                messagebox.showerror("Fejl", "Der opstod en uventet fejl:\n\n"
                                     + traceback.format_exc())
                self.sæt_status("Fejl - se fejlbesked.")
        return wrapper

    def sæt_status(self, tekst):
        self.status.config(text=tekst)
        self.root.update_idletasks()

    def vælg_filer(self, flere=True, filetypes=None):
        if filetypes is None:
            filetypes = [("Word-dokumenter", "*.docx")]
        fn = filedialog.askopenfilenames if flere else filedialog.askopenfilename
        valg = fn(title="Vælg kravspecifikation(er)", filetypes=filetypes)
        if not valg:
            self.sæt_status("Ingen dokumenter blev valgt.")
            return None
        return list(valg) if flere else [valg]

    def gem_som(self, forslag):
        sti = filedialog.asksaveasfilename(
            title="Gem resultat som", initialfile=forslag,
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not sti:
            self.sæt_status("Annulleret - intet gemt.")
        return sti

    def færdig(self, sti, antal):
        self.sæt_status(f"Udtræk gennemført: {antal} rækker -> {Path(sti).name}")
        if messagebox.askyesno("Udtræk gennemført",
                               f"{antal} rækker skrevet til:\n{sti}\n\nÅbne filen nu?"):
            try:
                os.startfile(sti)            # Windows
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", sti])

    def læs(self, filer):
        """Indlæser hver fil som ex.Docx. Filer der ikke kan læses (fx
        gamle .doc-filer, krypterede eller beskadigede .docx-filer)
        springes over, og brugeren informeres samlet til sidst. Fejler
        ALLE filer, vises en fejlbesked og None returneres, så det
        kaldende sted kan afbryde."""
        docs = []
        sprunget_over = []
        for f in filer:
            self.sæt_status(f"Læser {Path(f).name} ...")
            try:
                docs.append(ex.Docx(f))
            except ex.DocxFejl as fejl:
                sprunget_over.append(str(fejl))
        if sprunget_over:
            if not docs:
                messagebox.showerror(
                    "Ingen dokumenter kunne læses",
                    "Ingen af de valgte dokumenter kunne læses:\n\n"
                    + "\n\n".join(sprunget_over))
                return None
            messagebox.showwarning(
                "Nogle dokumenter blev sprunget over",
                f"{len(sprunget_over)} dokument(er) kunne ikke læses og "
                "blev sprunget over:\n\n" + "\n\n".join(sprunget_over))
        return docs

    def vælg_kontekst_mappe(self):
        """Spørger om der skal angives en baggrundsmappe (fx udbudsbetingelserne)
        hvis opslag bruges til at validere henvisninger korrekt, uden selv at
        blive tjekket for henvisninger. Returnerer valgt mappesti eller None."""
        if not messagebox.askyesno(
                "Baggrundsdokumenter?",
                "Vil du angive en mappe med baggrundsdokumenter (fx selve "
                "udbudsbetingelserne)?\n\nSå bliver henvisninger som "
                "\"punkt 14.1 i udbudsbetingelserne\" ikke fejlagtigt markeret "
                "som ugyldige, bare fordi udbudsbetingelserne ikke selv er "
                "blandt de dokumenter du vil have tjekket.\n\n"
                "Dokumenterne i mappen bruges KUN til opslag - de tjekkes ikke "
                "selv for henvisninger."):
            return None
        mappe = filedialog.askdirectory(title="Vælg mappe med baggrundsdokumenter")
        return mappe or None

    # ------------------------------------------------------------ udtrækkene
    def _udtræk_kommentarer_fra_fil(self, f):
        if Path(f).suffix.lower() == ".xlsx":
            return ex.extract_excel_comments(f)
        return ex.extract_comments(ex.Docx(f))

    def kommentarer(self):
        filer = self.vælg_filer(filetypes=KOMMENTAR_TYPER)
        if not filer:
            return
        recs = []
        sprunget_over = []
        for f in filer:
            self.sæt_status(f"Læser {Path(f).name} ...")
            try:
                fund = self._udtræk_kommentarer_fra_fil(f)
            except ex.DocxFejl as fejl:
                sprunget_over.append(str(fejl))
                continue
            recs += fund or [{"Dokument": Path(f).name,
                              "Kommentar": "Ingen kommentarer",
                              "__ingen_fund__": True}]
        if len(sprunget_over) == len(filer):
            messagebox.showerror(
                "Ingen dokumenter kunne læses",
                "Ingen af de valgte dokumenter kunne læses:\n\n"
                + "\n\n".join(sprunget_over))
            return
        if sprunget_over:
            messagebox.showwarning(
                "Nogle dokumenter blev sprunget over",
                f"{len(sprunget_over)} dokument(er) kunne ikke læses og "
                "blev sprunget over:\n\n" + "\n\n".join(sprunget_over))
        sti = self.gem_som("kommentarer.xlsx")
        if not sti:
            return
        ex.write_simple(recs, ["Dokument", "Nummer", "Tråd", "Svar på", "Kommentar",
                               "Markeret tekst", "Initialer", "Nummereret sektion"], sti,
                        [35, 9, 10, 10, 50, 50, 12, 50])
        self.færdig(sti, len(recs))

    def kommentarer_filtreret(self):
        filer = self.vælg_filer(filetypes=KOMMENTAR_TYPER)
        if not filer:
            return
        recs = []
        antal_uden = 0
        sprunget_over = []
        for f in filer:
            self.sæt_status(f"Læser {Path(f).name} ...")
            try:
                fund = self._udtræk_kommentarer_fra_fil(f)
            except ex.DocxFejl as fejl:
                sprunget_over.append(str(fejl))
                continue
            if fund:
                recs += fund
            else:
                antal_uden += 1

        if len(sprunget_over) == len(filer):
            messagebox.showerror(
                "Ingen dokumenter kunne læses",
                "Ingen af de valgte dokumenter kunne læses:\n\n"
                + "\n\n".join(sprunget_over))
            return
        if sprunget_over:
            messagebox.showwarning(
                "Nogle dokumenter blev sprunget over",
                f"{len(sprunget_over)} dokument(er) kunne ikke læses og "
                "blev sprunget over:\n\n" + "\n\n".join(sprunget_over))

        if not recs:
            messagebox.showinfo("Ingen kommentarer",
                                "Ingen af de valgte dokumenter indeholder kommentarer.")
            self.sæt_status("Ingen kommentarer fundet i nogen dokumenter.")
            return
        sti = self.gem_som("kommentarer_filtreret.xlsx")
        if not sti:
            return
        ex.write_simple(recs, ["Dokument", "Nummer", "Tråd", "Svar på", "Kommentar",
                               "Markeret tekst", "Initialer", "Nummereret sektion"], sti,
                        [35, 9, 10, 10, 50, 50, 12, 50])
        self.sæt_status(f"Udtræk gennemført: {len(recs)} rækker -> {Path(sti).name}")
        besked = f"{len(recs)} rækker skrevet til:\n{sti}"
        if antal_uden:
            besked += f"\n\n({antal_uden} dokument(er) uden kommentarer er udeladt)"
        if messagebox.askyesno("Udtræk gennemført", besked + "\n\nÅbne filen nu?"):
            try:
                os.startfile(sti)
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", sti])

    def trackchanges(self):
        filer = self.vælg_filer()
        if not filer:
            return
        docs = self.læs(filer)
        if not docs:
            return
        recs = []
        for d in docs:
            recs += ex.extract_trackchanges(d) or [{"Dokument": d.path.name,
                                                    "Ændring": "Ingen trackchanges"}]
        sti = self.gem_som("trackchanges.xlsx")
        if not sti:
            return
        ex.write_simple(recs, ["Dokument", "Sektion", "Ændring", "Type", "Forfatter"],
                        sti, [35, 50, 60, 10, 20])
        self.færdig(sti, len(recs))

    def inputfelter(self, farve):
        filer = self.vælg_filer()
        if not filer:
            return
        docs = self.læs(filer)
        if not docs:
            return
        recs = []
        for d in docs:
            recs += ex.extract_highlights(d, farve) or [{"Dokument": d.path.name,
                                                         "Inputfelt": "Ingen markerede felter"}]
        sti = self.gem_som(f"inputfelter_{farve}.xlsx")
        if not sti:
            return
        ex.write_simple(recs, ["Dokument", "Sektion", "Inputfelt"], sti, [35, 50, 60])
        self.færdig(sti, len(recs))

    def krydstjek(self):
        try:
            import krydstjek as kt
        except ImportError:
            messagebox.showerror("Mangler modul",
                                 "krydstjek.py skal ligge i samme mappe som denne fil.")
            return
        filer = filedialog.askopenfilenames(
            title="Vælg kontrakten OG alle bilag (markér flere med Ctrl)",
            filetypes=[("Word-dokumenter", "*.docx")])
        if not filer:
            self.sæt_status("Ingen dokumenter blev valgt.")
            return
        if len(filer) == 1:
            if not messagebox.askyesno(
                    "Kun ét dokument valgt",
                    "Du har kun valgt ét dokument. Henvisninger til bilag kan så "
                    "ikke valideres mod selve bilagsdokumenterne.\n\nFortsæt alligevel?"):
                return
        docs = self.læs(list(filer))
        if not docs:
            return
        kontekst_mappe = self.vælg_kontekst_mappe()
        kontekst_docs = kt.byg_kontekst_docs([kontekst_mappe], docs) if kontekst_mappe else []
        if kontekst_docs:
            self.sæt_status(f"Kontekst: {len(kontekst_docs)} baggrundsdokument(er) indlæst ...")
        self.sæt_status("Kortlægger definitioner og validerer henvisninger ...")
        fund, sektioner, bilag_def = kt.validér_samlinger(
            [("Standard", docs)], kontekst_docs=kontekst_docs)
        if not fund:
            messagebox.showinfo("Ingen henvisninger",
                                "Der blev ikke fundet krydshenvisninger i dokumenterne.")
            self.sæt_status("Ingen henvisninger fundet.")
            return
        sti = self.gem_som("krydstjek.xlsx")
        if not sti:
            return
        n = kt.skriv_rapport(fund, sektioner, bilag_def, sti)
        self.sæt_status(f"Krydstjek: ❌ {n['ugyldig']}  ⚠️ {n['usikker']}  ✅ {n['gyldig']}")
        if messagebox.askyesno(
                "Krydstjek gennemført",
                f"❌ Ugyldige: {n['ugyldig']}\n⚠️ Usikre: {n['usikker']}\n"
                f"✅ Gyldige: {n['gyldig']}\n\nRapport gemt:\n{sti}\n\nÅbne rapporten nu?"):
            try:
                os.startfile(sti)
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", sti])

    def krydstjek_semantik(self):
        try:
            import krydstjek as kt
        except ImportError:
            messagebox.showerror("Mangler modul",
                                 "krydstjek.py skal ligge i samme mappe som denne fil.")
            return
        try:
            import semantik
        except ImportError:
            messagebox.showerror("Mangler modul",
                                 "semantik.py skal ligge i samme mappe som denne fil.\n"
                                 "Kør 'pip install anthropic python-dotenv' og opret en "
                                 ".env-fil med ANTHROPIC_API_KEY.")
            return
        filer = filedialog.askopenfilenames(
            title="Vælg kontrakten OG alle bilag (markér flere med Ctrl)",
            filetypes=[("Word-dokumenter", "*.docx")])
        if not filer:
            self.sæt_status("Ingen dokumenter blev valgt.")
            return
        if len(filer) == 1:
            if not messagebox.askyesno(
                    "Kun ét dokument valgt",
                    "Du har kun valgt ét dokument. Henvisninger til bilag kan så "
                    "ikke valideres mod selve bilagsdokumenterne.\n\nFortsæt alligevel?"):
                return
        docs = self.læs(list(filer))
        if not docs:
            return
        kontekst_mappe = self.vælg_kontekst_mappe()
        kontekst_docs = kt.byg_kontekst_docs([kontekst_mappe], docs) if kontekst_mappe else []
        if kontekst_docs:
            self.sæt_status(f"Kontekst: {len(kontekst_docs)} baggrundsdokument(er) indlæst ...")
        samlinger = [("Standard", docs)]
        self.sæt_status("Kortlægger definitioner og validerer henvisninger ...")
        fund, sektioner, bilag_def = kt.validér_samlinger(samlinger, kontekst_docs=kontekst_docs)
        if not fund:
            messagebox.showinfo("Ingen henvisninger",
                                "Der blev ikke fundet krydshenvisninger i dokumenterne.")
            self.sæt_status("Ingen henvisninger fundet.")
            return
        self.sæt_status("Kører semantisk AI-analyse af ugyldige/usikre henvisninger ...")
        try:
            dok_indhold = kt.byg_dokument_indhold(samlinger)
            for d in kontekst_docs:
                dok_indhold[d.path.name] = kt._dok_tekst(d)
            semantik_resultater = semantik.analysér_batch(fund, dok_indhold)
        except Exception:
            messagebox.showwarning("Semantisk analyse fejlede",
                                   "Semantisk analyse kunne ikke gennemføres:\n\n"
                                   + traceback.format_exc(limit=3)
                                   + "\nRapporten gemmes uden semantisk vurdering.")
            semantik_resultater = {}
        sti = self.gem_som("krydstjek_semantik.xlsx")
        if not sti:
            return
        n = kt.skriv_rapport(fund, sektioner, bilag_def, sti, semantik_resultater)
        self.sæt_status(f"Krydstjek: ❌ {n['ugyldig']}  ⚠️ {n['usikker']}  ✅ {n['gyldig']}")
        if messagebox.askyesno(
                "Krydstjek gennemført",
                f"❌ Ugyldige: {n['ugyldig']}\n⚠️ Usikre: {n['usikker']}\n"
                f"✅ Gyldige: {n['gyldig']}\n\nRapport gemt:\n{sti}\n\nÅbne rapporten nu?"):
            try:
                os.startfile(sti)
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", sti])

    def intern_tjek(self):
        try:
            import intern_tjek as it
        except ImportError:
            messagebox.showerror("Mangler modul",
                                 "intern_tjek.py skal ligge i samme mappe som denne fil.")
            return
        filer = self.vælg_filer(flere=False)
        if not filer:
            return
        docs = self.læs(filer)
        if not docs:
            return
        d = docs[0]
        self.sæt_status("Validerer interne afsnits-/punkthenvisninger ...")
        fund, numre = it.validér_internt(d)
        if not fund:
            messagebox.showinfo("Ingen henvisninger",
                                "Der blev ikke fundet interne afsnits-/punkthenvisninger "
                                "i dokumentet.")
            self.sæt_status("Ingen henvisninger fundet.")
            return
        sti = self.gem_som("intern_tjek.xlsx")
        if not sti:
            return
        n = it.skriv_rapport(fund, sti)
        self.sæt_status(f"Intern tjek: ❌ {n['ugyldig']}  ⚠️ {n['usikker']}  ✅ {n['gyldig']}")
        if messagebox.askyesno(
                "Intern tjek gennemført",
                f"❌ Ugyldige: {n['ugyldig']}\n⚠️ Usikre: {n['usikker']}\n"
                f"✅ Gyldige: {n['gyldig']}\n\nRapport gemt:\n{sti}\n\nÅbne rapporten nu?"):
            try:
                os.startfile(sti)
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", sti])

    def styles(self):
        filer = self.vælg_filer()
        if not filer:
            return
        docs = self.læs(filer)
        if not docs:
            return
        tekst = ""
        for d in docs:
            tekst += f"{d.path.name}:\n" + "\n".join(
                f"   - {s}" for s in ex.list_styles(d)) + "\n\n"
        self.sæt_status("Typografier fundet.")
        messagebox.showinfo("Dokumentets typografier", tekst.strip())

    def kravmatrix(self):
        filer = self.vælg_filer(flere=False)
        if not filer:
            return
        docs = self.læs(filer)
        if not docs:
            return
        d = docs[0]
        valg = StyleDialog(self.root, ex.list_styles(d)).resultat
        if valg is None:
            self.sæt_status("Annulleret.")
            return
        table_styles, heading_styles = valg
        self.sæt_status("Danner kravmatrix ...")
        rows = ex.extract_kravmatrix(d, table_styles, heading_styles)
        if not rows:
            messagebox.showwarning("Tomt resultat",
                                   "Ingen rækker matchede de valgte typografier.\n"
                                   "Tjek valgene via 'Vis dokumentets typografier'.")
            self.sæt_status("Intet udtrukket.")
            return
        sti = self.gem_som("kravmatrix.xlsx")
        if not sti:
            return
        ex.write_kravmatrix(rows, sti)
        self.færdig(sti, len(rows))

    def kravmatrix_kommentarer(self):
        filer = self.vælg_filer(flere=False)
        if not filer:
            return
        docs = self.læs(filer)
        if not docs:
            return
        d = docs[0]
        valg = StyleDialog(self.root, ex.list_styles(d)).resultat
        if valg is None:
            self.sæt_status("Annulleret.")
            return
        table_styles, heading_styles = valg
        self.sæt_status("Danner kravmatrix med kommentarer ...")
        rows = ex.extract_kravmatrix(d, table_styles, heading_styles, with_comments=True)
        if not rows:
            messagebox.showwarning("Tomt resultat",
                                   "Ingen rækker matchede de valgte typografier.\n"
                                   "Tjek valgene via 'Vis dokumentets typografier'.")
            self.sæt_status("Intet udtrukket.")
            return
        sti = self.gem_som("kravmatrix_kommentarer.xlsx")
        if not sti:
            return
        ex.write_kravmatrix_kommentarer(rows, sti)
        self.færdig(sti, len(rows))


class StyleDialog:
    """Erstatter Excel-værktøjets UserForm6 + UserForm4 i ét vindue."""

    def __init__(self, parent, styles):
        self.resultat = None
        top = self.top = Toplevel(parent)
        top.title("Vælg typografier til kravmatrix")
        top.configure(bg=BG)
        top.grab_set()

        venstre = Frame(top, bg=BG)
        venstre.pack(side=LEFT, padx=12, pady=10, fill=BOTH)
        Label(venstre, text="Styles tilhørende TABEL-elementer\n(kravrækkerne)",
              bg=BG, font=("Segoe UI", 9, "bold")).pack()
        self.alle = BooleanVar()
        Checkbutton(venstre, text="[ALLE] - tag alle tabelrækker med",
                    variable=self.alle, bg=BG).pack(anchor="w")
        self.lb_tabel = self._listbox(venstre, styles)

        højre = Frame(top, bg=BG)
        højre.pack(side=LEFT, padx=12, pady=10, fill=BOTH)
        Label(højre, text="Styles tilhørende OVERSKRIFTER\n(separatorer - valgfrit)",
              bg=BG, font=("Segoe UI", 9, "bold")).pack()
        Label(højre, text="(typisk Overskrift 1-3 / Heading 1-3)",
              bg=BG, fg="#666", font=("Segoe UI", 8)).pack(anchor="w")
        self.lb_overskrift = self._listbox(højre, styles)

        knapper = Frame(top, bg=BG)
        knapper.pack(side=RIGHT, padx=12, pady=10, anchor="s")
        Button(knapper, text="OK", width=12, command=self.ok).pack(pady=4)
        Button(knapper, text="Annuller", width=12, command=top.destroy).pack(pady=4)

        top.wait_window()

    def _listbox(self, parent, styles):
        ramme = Frame(parent)
        ramme.pack(fill=BOTH, expand=True, pady=4)
        sb = Scrollbar(ramme)
        sb.pack(side=RIGHT, fill=Y)
        lb = Listbox(ramme, selectmode=EXTENDED, width=34, height=12,
                     yscrollcommand=sb.set, exportselection=False)
        for s in styles:
            lb.insert(END, s)
        lb.pack(side=LEFT, fill=BOTH, expand=True)
        sb.config(command=lb.yview)
        return lb

    def ok(self):
        if self.alle.get():
            tabel = "ALLE"
        else:
            tabel = [self.lb_tabel.get(i) for i in self.lb_tabel.curselection()]
            if not tabel:
                messagebox.showwarning("Manglende valg",
                                       "Vælg mindst én tabel-style, eller sæt kryds i [ALLE].",
                                       parent=self.top)
                return
        overskrifter = [self.lb_overskrift.get(i)
                        for i in self.lb_overskrift.curselection()]
        self.resultat = (tabel, overskrifter)
        self.top.destroy()


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()
