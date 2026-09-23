#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Stjernekort GUI - knap-interface til stjernekort.cli
=====================================================
To trin, samme som CLI'en (se stjernekort/cli.py):
  1. Udtræk:  mappe med docx + ordbog.json  -> blueprint.xlsx (masteren - kan rettes i Excel)
  2. Generer: blueprint.xlsx + template.pptx -> blueprint.pptx

Læg denne fil i SAMME mappe som "stjernekort"-pakken og start den med:
    python stjernekort_gui.py
(eller dobbeltklik på stjernekort_gui.pyw for at slippe for det sorte konsolvindue)
"""

import os
import traceback
from pathlib import Path
from tkinter import Tk, Frame, Label, Entry, Button, filedialog, messagebox, StringVar, X, LEFT, RIGHT

try:
    from stjernekort import excel_io
    from stjernekort.pptx_generator import generer as generer_pptx
    from stjernekort.udtraek import laes_ordbog, udtraek
except ImportError:
    from tkinter import Tk as _Tk, messagebox as _mb
    _mb.showerror("Mangler modul",
                  "\"stjernekort\"-pakken skal ligge i samme mappe som denne fil.")
    raise SystemExit(1)

# Rambøll-inspireret farvepalet (samme som tidsplan_gui.py)
OXFORD = "#2D3748"
CYAN = "#009DE0"
WHITE = "#FFFFFF"
GRAA = "#718096"
HOVER = "#E6F6FD"
HEADER_UNDERTEKST = "#CBD5E0"
HEADER_VERSION = "#A0AEC0"
STATUS_BG = "#EDF2F7"

HER = Path(__file__).resolve().parent
FORVALGT_ORDBOG = HER / "stjernekort" / "ordbog_esdh.json"
FORVALGT_TEMPLATE = HER / "testdata" / "Kontraktens_blueprint_template.pptx"


class App:
    def __init__(self, root):
        self.root = root
        root.title("Stjernekort - kontrakt-blueprint")
        root.configure(bg=WHITE)
        root.resizable(False, False)
        try:
            root.iconbitmap(str(HER / "icon.ico"))
        except Exception:
            pass

        # ---- header -----------------------------------------------------
        header = Frame(root, bg=OXFORD)
        header.pack(fill=X)
        header_venstre = Frame(header, bg=OXFORD)
        header_venstre.pack(side=LEFT, anchor="w", padx=20, pady=(16, 14))
        Label(header_venstre, text="Stjernekort", bg=OXFORD, fg=WHITE,
              font=("Segoe UI", 18, "bold")).pack(anchor="w")
        Label(header_venstre, text="Kontrakt-blueprint fra docx til PowerPoint via Excel",
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

        self.docx_mappe_var = StringVar()
        self.ordbog_var = StringVar(value=str(FORVALGT_ORDBOG) if FORVALGT_ORDBOG.exists() else "")
        self.udtraek_xlsx_var = StringVar(value=str(HER / "ud" / "blueprint.xlsx"))

        self.blueprint_xlsx_var = StringVar()
        self.template_var = StringVar(value=str(FORVALGT_TEMPLATE) if FORVALGT_TEMPLATE.exists() else "")
        self.ud_pptx_var = StringVar(value=str(HER / "ud" / "blueprint.pptx"))
        self.kunde_var = StringVar()
        self.udbud_var = StringVar()

        self._sektion_overskrift(krop, "1. Udtræk - kontrakt + bilag → Excel (masteren)")
        self._felt(krop, "Mappe med docx (kontrakt + bilag)", self.docx_mappe_var,
                  self.vælg_docx_mappe, "Gennemse ...")
        self._felt(krop, "Ordbog (.json)", self.ordbog_var,
                  self.vælg_ordbog, "Gennemse ...")
        self._felt(krop, "Output Excel (.xlsx)", self.udtraek_xlsx_var,
                  self.vælg_udtraek_output, "Gem som ...")

        kant1 = Frame(krop, bg=CYAN)
        kant1.pack(fill=X, padx=20, pady=(10, 6))
        knap1 = self._knap(kant1, "Lav udtræk", self.beskyt(self.kør_udtraek))

        self._sektion_overskrift(krop, "2. Generer - Excel → PowerPoint-blueprint")
        self._felt(krop, "Blueprint-Excel (.xlsx)", self.blueprint_xlsx_var,
                  self.vælg_blueprint_xlsx, "Gennemse ...")
        self._felt(krop, "Template (.pptx)", self.template_var,
                  self.vælg_template, "Gennemse ...")
        self._felt(krop, "Kundenavn (valgfri - vises i dias-titlen)", self.kunde_var)
        self._felt(krop, "Navn på udbud/projekt (valgfri - vises i dias-titlen)", self.udbud_var)
        self._felt(krop, "Output PowerPoint (.pptx)", self.ud_pptx_var,
                  self.vælg_pptx_output, "Gem som ...")

        kant2 = Frame(krop, bg=CYAN)
        kant2.pack(fill=X, padx=20, pady=(10, 6))
        knap2 = self._knap(kant2, "Lav blueprint", self.beskyt(self.kør_generer))

        Frame(krop, bg=WHITE, height=8).pack()

        # Sørg for en luftig, passende bred rude uden at klippe indhold
        root.update_idletasks()
        bredde = max(600, root.winfo_reqwidth())
        højde = root.winfo_reqheight()
        root.geometry(f"{bredde}x{højde}")

    # ------------------------------------------------------------ design ---
    def _sektion_overskrift(self, parent, tekst):
        Label(parent, text=tekst, bg=WHITE, fg=OXFORD,
              font=("Segoe UI", 11, "bold"), anchor="w").pack(
              fill=X, padx=20, pady=(18, 2))

    def _felt(self, parent, label, var, browse_cmd=None, browse_tekst=None):
        Label(parent, text=label, bg=WHITE, fg=GRAA,
              font=("Segoe UI", 9, "bold"), anchor="w").pack(
              fill=X, padx=20, pady=(10, 4))
        række = Frame(parent, bg=WHITE)
        række.pack(fill=X, padx=20)
        Entry(række, textvariable=var, font=("Segoe UI", 10),
              relief="solid", borderwidth=1).pack(side=LEFT, fill=X, expand=True, ipady=4)
        if browse_cmd is not None:
            Button(række, text=browse_tekst, command=browse_cmd,
                  bg=WHITE, fg=OXFORD, relief="solid", borderwidth=1,
                  font=("Segoe UI", 9), cursor="hand2", padx=10).pack(side=LEFT, padx=(8, 0))

    def _knap(self, kant, tekst, command):
        knap = Button(kant, text=tekst, command=command,
                     bg=WHITE, fg=OXFORD, activebackground=HOVER,
                     activeforeground=OXFORD, relief="flat", borderwidth=0,
                     highlightthickness=0, padx=14, pady=12,
                     font=("Segoe UI", 11, "bold"), cursor="hand2")
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

    def _åbn_fil(self, sti):
        try:
            os.startfile(sti)
        except AttributeError:
            import subprocess
            subprocess.Popen(["xdg-open", sti])

    # ------------------------------------------------------------ filvalg ---
    def vælg_docx_mappe(self):
        sti = filedialog.askdirectory(title="Vælg mappe med kontrakt + bilag (docx)")
        if sti:
            self.docx_mappe_var.set(sti)
            self.sæt_status(f"Valgt mappe: {Path(sti).name}")

    def vælg_ordbog(self):
        sti = filedialog.askopenfilename(
            title="Vælg ordbog", initialdir=str(HER / "stjernekort"),
            filetypes=[("JSON", "*.json"), ("Alle filer", "*.*")])
        if sti:
            self.ordbog_var.set(sti)

    def vælg_udtraek_output(self):
        sti = filedialog.asksaveasfilename(
            title="Gem Excel-udtræk som", initialfile="blueprint.xlsx",
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if sti:
            self.udtraek_xlsx_var.set(sti)

    def vælg_blueprint_xlsx(self):
        sti = filedialog.askopenfilename(
            title="Vælg blueprint-Excel", filetypes=[("Excel", "*.xlsx"), ("Alle filer", "*.*")])
        if sti:
            self.blueprint_xlsx_var.set(sti)

    def vælg_template(self):
        sti = filedialog.askopenfilename(
            title="Vælg PowerPoint-template", initialdir=str(HER / "testdata"),
            filetypes=[("PowerPoint", "*.pptx"), ("Alle filer", "*.*")])
        if sti:
            self.template_var.set(sti)

    def vælg_pptx_output(self):
        sti = filedialog.asksaveasfilename(
            title="Gem blueprint som", initialfile="blueprint.pptx",
            defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")])
        if sti:
            self.ud_pptx_var.set(sti)

    # ------------------------------------------------------------------ kør
    def kør_udtraek(self):
        mappe = self.docx_mappe_var.get().strip()
        ordbog_sti = self.ordbog_var.get().strip()
        ud = self.udtraek_xlsx_var.get().strip()

        if not mappe or not Path(mappe).is_dir():
            messagebox.showwarning("Manglende input", "Vælg en gyldig mappe med docx-filer.")
            return
        if not ordbog_sti or not Path(ordbog_sti).exists():
            messagebox.showwarning("Manglende input", "Vælg en gyldig ordbog (.json).")
            return
        if not ud:
            messagebox.showwarning("Manglende input", "Angiv et output-filnavn (.xlsx).")
            return

        Path(ud).parent.mkdir(parents=True, exist_ok=True)
        self.sæt_status(f"Læser docx i {Path(mappe).name} ...")
        ordbog = laes_ordbog(ordbog_sti)
        bp = udtraek(mappe, ordbog)
        self.sæt_status("Skriver Excel-udtræk ...")
        excel_io.skriv(bp, ud)

        self.blueprint_xlsx_var.set(ud)
        besked = (f"{len(bp.milepaele)} milepæle, {len(bp.retsvirkninger)} retsvirkninger, "
                  f"{len(bp.henvisninger)} henvisninger skrevet til:\n{ud}")
        self.sæt_status(f"Udtræk færdigt -> {Path(ud).name}")
        if messagebox.askyesno("Udtræk gennemført", besked + "\n\nÅbne Excel-arket nu?"):
            self._åbn_fil(ud)

    def kør_generer(self):
        xlsx = self.blueprint_xlsx_var.get().strip()
        template = self.template_var.get().strip()
        ud = self.ud_pptx_var.get().strip()

        if not xlsx or not Path(xlsx).exists():
            messagebox.showwarning("Manglende input", "Vælg en gyldig blueprint-Excel (.xlsx).")
            return
        if not template or not Path(template).exists():
            messagebox.showwarning("Manglende input", "Vælg en gyldig PowerPoint-template.")
            return
        if not ud:
            messagebox.showwarning("Manglende input", "Angiv et output-filnavn (.pptx).")
            return

        Path(ud).parent.mkdir(parents=True, exist_ok=True)
        self.sæt_status(f"Læser {Path(xlsx).name} ...")
        bp = excel_io.laes(xlsx)
        self.sæt_status("Tegner blueprint ...")
        generer_pptx(bp, template, ud, kunde=self.kunde_var.get().strip(),
                    udbud=self.udbud_var.get().strip())

        self.sæt_status(f"Færdig -> {Path(ud).name}")
        if messagebox.askyesno("Blueprint gennemført", f"Skrevet til:\n{ud}\n\nÅbne filen nu?"):
            self._åbn_fil(ud)


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()
