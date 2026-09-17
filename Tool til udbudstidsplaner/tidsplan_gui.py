#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tidsplan GUI - knap-interface til tidsplan.py
==============================================
Læg denne fil i SAMME mappe som tidsplan.py og pptx_render.py og start den med:
    python tidsplan_gui.py
(eller dobbeltklik på filen, hvis .py er knyttet til Python)

Kræver kun Python + openpyxl + Pillow + python-pptx - tkinter følger med Python.
"""

import datetime as dt
import os
import traceback
from pathlib import Path
from tkinter import Tk, Frame, Label, Entry, Button, filedialog, messagebox, StringVar, X, LEFT, RIGHT
from tkinter import ttk

try:
    import tidsplan as tp
    from pptx_render import render
except ImportError:
    messagebox.showerror("Mangler modul",
                         "tidsplan.py og pptx_render.py skal ligge i samme mappe som denne fil.")
    raise SystemExit(1)

# Faste konfigurationer (pr. udbud) gemmes her, saa de kan vaelges fra en liste
# i stedet for at skulle findes med en filvaelger.
KONFIG_MAPPE = Path(__file__).resolve().parent / "konfigurationer"
KONFIG_MAPPE.mkdir(exist_ok=True)
INGEN_KONFIG = "(ingen - generisk)"

# Rambøll-inspireret farvepalet (samme som extractor_gui.py)
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
        root.title("Tidsplan - visuel udbudstidsplan")
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
        Label(header_venstre, text="Tidsplan", bg=OXFORD, fg=WHITE,
              font=("Segoe UI", 18, "bold")).pack(anchor="w")
        Label(header_venstre, text="Visuel udbudstidsplan fra MS Project-udtræk",
              bg=OXFORD, fg=HEADER_UNDERTEKST,
              font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))
        Label(header, text="v1.0 - sep 2026", bg=OXFORD, fg=HEADER_VERSION,
              font=("Segoe UI", 8)).pack(side=RIGHT, anchor="n", padx=16, pady=(14, 0))

        # ---- statuslinje og kreditering (bund - pakkes først så de "bunder") --
        self.status = Label(root, text="Klar.", bg=STATUS_BG, fg=OXFORD,
                            font=("Segoe UI", 9), anchor="w", padx=14, pady=6)
        self.status.pack(fill=X, side="bottom")

        Label(root, text="Udviklet af Birger Kidmose", bg=WHITE, fg=GRAA,
              font=("Segoe UI", 8), anchor="e").pack(
              fill=X, padx=18, pady=(4, 8), side="bottom")

        # ---- felter -------------------------------------------------------
        krop = Frame(root, bg=WHITE)
        krop.pack(fill=X)

        self.plan_var = StringVar()
        self.idag_var = StringVar(value=dt.date.today().isoformat())
        self.kunde_var = StringVar()
        self.udbud_var = StringVar()
        self.config_navn_var = StringVar(value=INGEN_KONFIG)
        self.out_var = StringVar(value=str(Path.cwd() / "Tidsplan_visuel.pptx"))
        self._avanceret = tp.Konfig()   # fase_navne/korte_navne/ekstra_milepaele fra evt. valgt konfig
        self._config_stier = {}          # visningsnavn -> Path, udfyldes af _konfig_liste()

        self._felt(krop, "Excel-fil (MS Project-udtræk)", self.plan_var,
                  self.vælg_plan, "Gennemse ...")
        self._felt(krop, "Statusdato (I dag) - YYYY-MM-DD", self.idag_var)
        self._felt(krop, "Kundenavn (fx \"Kommune Nord\")", self.kunde_var)
        self._felt(krop, "Udbuddets navn (fx \"Nyt ESDH-system\")", self.udbud_var)
        self._konfig_felt(krop)
        self._felt(krop, "Output (.pptx)", self.out_var,
                  self.vælg_output, "Gem som ...")

        # ---- kør-knap -------------------------------------------------------
        kant = Frame(krop, bg=CYAN)
        kant.pack(fill=X, padx=20, pady=(16, 6))
        knap = Button(kant, text="Lav tidsplan", command=self.beskyt(self.kør),
                     bg=WHITE, fg=OXFORD, activebackground=HOVER,
                     activeforeground=OXFORD, relief="flat", borderwidth=0,
                     highlightthickness=0, padx=14, pady=12,
                     font=("Segoe UI", 11, "bold"), cursor="hand2")
        knap.pack(fill=X, padx=1, pady=1)
        knap.bind("<Enter>", lambda e: knap.config(bg=HOVER))
        knap.bind("<Leave>", lambda e: knap.config(bg=WHITE))

        Frame(krop, bg=WHITE, height=8).pack()

        # Sørg for en luftig, passende bred rude uden at klippe indhold
        root.update_idletasks()
        bredde = max(560, root.winfo_reqwidth())
        højde = root.winfo_reqheight()
        root.geometry(f"{bredde}x{højde}")

    # ------------------------------------------------------------ design ---
    def _felt(self, parent, label, var, browse_cmd=None, browse_tekst=None):
        Label(parent, text=label, bg=WHITE, fg=GRAA,
              font=("Segoe UI", 9, "bold"), anchor="w").pack(
              fill=X, padx=20, pady=(16, 4))
        række = Frame(parent, bg=WHITE)
        række.pack(fill=X, padx=20)
        Entry(række, textvariable=var, font=("Segoe UI", 10),
              relief="solid", borderwidth=1).pack(side=LEFT, fill=X, expand=True, ipady=4)
        if browse_cmd is not None:
            Button(række, text=browse_tekst, command=browse_cmd,
                  bg=WHITE, fg=OXFORD, relief="solid", borderwidth=1,
                  font=("Segoe UI", 9), cursor="hand2", padx=10).pack(side=LEFT, padx=(8, 0))

    def _konfig_felt(self, parent):
        Label(parent, text="Konfiguration (valgfri - fase-omdøbninger, ekstra milepæle mv., "
                           "vælges fra listen - intet at lede efter)",
              bg=WHITE, fg=GRAA, font=("Segoe UI", 9, "bold"), anchor="w",
              wraplength=520, justify=LEFT).pack(fill=X, padx=20, pady=(16, 4))
        række = Frame(parent, bg=WHITE)
        række.pack(fill=X, padx=20)
        self.config_combo = ttk.Combobox(række, textvariable=self.config_navn_var,
                                         values=self._konfig_liste(), state="readonly",
                                         font=("Segoe UI", 10))
        self.config_combo.pack(side=LEFT, fill=X, expand=True, ipady=2)
        self.config_combo.bind("<<ComboboxSelected>>", self._on_konfig_valgt)
        Button(række, text="Gennemse ...", command=self.vælg_config,
              bg=WHITE, fg=OXFORD, relief="solid", borderwidth=1,
              font=("Segoe UI", 9), cursor="hand2", padx=10).pack(side=LEFT, padx=(8, 0))
        Button(række, text="Gem som ...", command=self.gem_config,
              bg=WHITE, fg=OXFORD, relief="solid", borderwidth=1,
              font=("Segoe UI", 9), cursor="hand2", padx=10).pack(side=LEFT, padx=(8, 0))
        Label(parent, text=f"Gemte konfigurationer ligger i: {KONFIG_MAPPE}", bg=WHITE, fg=GRAA,
              font=("Segoe UI", 8), anchor="w").pack(fill=X, padx=20, pady=(3, 0))

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

    def vælg_plan(self):
        sti = filedialog.askopenfilename(
            title="Vælg Excel-udtræk fra MS Project",
            filetypes=[("Excel-filer", "*.xlsx"), ("Alle filer", "*.*")])
        if sti:
            self.plan_var.set(sti)
            self.sæt_status(f"Valgt: {Path(sti).name}")

    def vælg_output(self):
        sti = filedialog.asksaveasfilename(
            title="Gem tidsplan som", initialfile="Tidsplan_visuel.pptx",
            defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")])
        if sti:
            self.out_var.set(sti)

    def _konfig_liste(self):
        """Scanner KONFIG_MAPPE for gemte konfigurationer og bygger navn->sti-listen."""
        filer = sorted(KONFIG_MAPPE.glob("*.json"), key=lambda p: p.stem.lower())
        self._config_stier = {p.stem: p for p in filer}
        return [INGEN_KONFIG] + list(self._config_stier.keys())

    def _tilføj_til_liste(self, navn, sti):
        """Føjer en (evt. eksternt valgt) konfiguration til dropdown-listen og vælger den."""
        self._config_stier[navn] = Path(sti)
        værdier = [INGEN_KONFIG] + sorted(self._config_stier.keys(), key=str.lower)
        self.config_combo["values"] = værdier
        self.config_navn_var.set(navn)

    def _indlæs_konfig(self, sti):
        try:
            konfig = tp.load_konfig(sti)
        except Exception:
            messagebox.showerror("Kan ikke læse konfigurationsfil",
                                 "Filen kunne ikke læses som en gyldig konfiguration:\n\n"
                                 + traceback.format_exc(limit=2))
            return None
        self._avanceret = konfig
        self.kunde_var.set(konfig.kunde)
        self.udbud_var.set(konfig.udbud)
        return konfig

    def _on_konfig_valgt(self, event=None):
        navn = self.config_navn_var.get()
        if navn == INGEN_KONFIG:
            self._avanceret = tp.Konfig()
            self.kunde_var.set("")
            self.udbud_var.set("")
            self.sæt_status("Ingen konfiguration valgt - generisk tidsplan.")
            return
        sti = self._config_stier.get(navn)
        if sti is not None and self._indlæs_konfig(sti) is not None:
            self.sæt_status(f"Konfiguration valgt: {navn}")

    def vælg_config(self):
        sti = filedialog.askopenfilename(
            title="Vælg konfigurationsfil", initialdir=KONFIG_MAPPE,
            filetypes=[("JSON", "*.json"), ("Alle filer", "*.*")])
        if not sti:
            return
        if self._indlæs_konfig(sti) is None:
            return
        self._tilføj_til_liste(Path(sti).stem, sti)
        self.sæt_status(f"Konfiguration indlæst: {Path(sti).name}")

    def gem_config(self):
        nuværende = self.config_navn_var.get()
        forslag = f"{nuværende}.json" if nuværende != INGEN_KONFIG else "nyt_udbud.json"
        sti = filedialog.asksaveasfilename(
            title="Gem konfiguration som", initialdir=KONFIG_MAPPE, initialfile=forslag,
            defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not sti:
            return
        self._avanceret.kunde = self.kunde_var.get().strip()
        self._avanceret.udbud = self.udbud_var.get().strip()
        tp.save_konfig(self._avanceret, sti)
        self._tilføj_til_liste(Path(sti).stem, sti)
        self.sæt_status(f"Konfiguration gemt: {Path(sti).name}")

    # ------------------------------------------------------------------ kør
    def kør(self):
        plan = self.plan_var.get().strip()
        if not plan:
            messagebox.showwarning("Manglende input", "Vælg en Excel-fil med tidsplanen.")
            return
        if not Path(plan).exists():
            messagebox.showerror("Filen findes ikke", f"Kan ikke finde:\n{plan}")
            return

        idag_tekst = self.idag_var.get().strip()
        try:
            idag = dt.date.fromisoformat(idag_tekst)
        except ValueError:
            messagebox.showerror("Ugyldig dato",
                                 "Statusdato skal skrives som YYYY-MM-DD, fx 2026-09-10.")
            return

        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning("Manglende input", "Angiv et output-filnavn (.pptx).")
            return

        self._avanceret.kunde = self.kunde_var.get().strip()
        self._avanceret.udbud = self.udbud_var.get().strip()

        self.sæt_status(f"Læser {Path(plan).name} ...")
        lanes, laes_advarsler = tp.laes_plan(plan, self._avanceret)
        self.sæt_status("Bygger tidsplan ...")
        prims, byg_advarsler = tp.byg(lanes, idag, self._avanceret)
        advarsler = laes_advarsler + byg_advarsler
        self.sæt_status(f"Skriver {Path(out).name} ...")
        render(prims, out)

        antal = sum(len(l["tasks"]) for l in lanes)
        besked = f"{antal} aktiviteter i {len(lanes)} faser skrevet til:\n{out}"
        if advarsler:
            besked += "\n\nAdvarsler:\n" + "\n".join(f"- {w}" for w in advarsler)
        self.sæt_status(f"Færdig: {antal} aktiviteter i {len(lanes)} faser -> {Path(out).name}"
                        + (f"  ({len(advarsler)} advarsel/-ler)" if advarsler else ""))

        if messagebox.askyesno("Tidsplan gennemført", besked + "\n\nÅbne filen nu?"):
            try:
                os.startfile(out)
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", out])


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()
