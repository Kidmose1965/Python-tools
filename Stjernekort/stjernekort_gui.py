#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Stjernekort GUI - knap-interface til stjernekort.cli
=====================================================
To trin, samme som CLI'en (se stjernekort/cli.py):
  1. Udtræk:  mappe med docx + valgt ordbog -> blueprint.xlsx (masteren - kan rettes i Excel)
  2. Generer: blueprint.xlsx + template.pptx -> blueprint.pptx

Ordbøgerne er lagdelte Excel-filer i stjernekort/ordboeger/: en grundordbog pr.
kontraktform ("ordbog_kontraktform_*.xlsx") og en udbudsspecifik ordbog pr. udbud
("ordbog_udbud_*.xlsx"), der bygger oven på grundordbogen. GUI'en kan generere et
udkast til en ny udbudsordbog ud fra selve kontrakten (se "Generér ordbog fra
kontrakt ..." nedenfor).

Læg denne fil i SAMME mappe som "stjernekort"-pakken og start den med:
    python stjernekort_gui.py
(eller dobbeltklik på stjernekort_gui.pyw for at slippe for det sorte konsolvindue)
"""

import json
import os
import re
import traceback
from pathlib import Path
from tkinter import (Tk, Frame, Label, Entry, Button, filedialog, messagebox, simpledialog,
                     StringVar, X, LEFT, RIGHT)
from tkinter import ttk

try:
    from stjernekort import excel_io
    from stjernekort.generer_ordbog import generer_udkast
    from stjernekort.ordbog_excel import laes_ordbog_excel, skriv_ordbog
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
ORDBOEGER_MAPPE = HER / "stjernekort" / "ordboeger"
FORVALGT_TEMPLATE = HER / "testdata" / "Kontraktens_blueprint_template.pptx"
STATE_STI = HER / ".stjernekort_gui_state.json"


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
        self.ordbog_navn_var = StringVar()
        self.ordbog_info_var = StringVar()
        self._ordbog_stier = {}          # visningsnavn -> Path, udfyldes af _ordbog_liste()
        self.udtraek_xlsx_var = StringVar(value=str(HER / "ud" / "blueprint.xlsx"))

        self.blueprint_xlsx_var = StringVar()
        self.template_var = StringVar(value=str(FORVALGT_TEMPLATE) if FORVALGT_TEMPLATE.exists() else "")
        self.ud_pptx_var = StringVar(value=str(HER / "ud" / "blueprint.pptx"))
        self.kunde_var = StringVar()
        self.udbud_var = StringVar()

        self._sektion_overskrift(krop, "1. Udtræk - kontrakt + bilag → Excel (masteren)")
        self._felt(krop, "Mappe med docx (kontrakt + bilag)", self.docx_mappe_var,
                  self.vælg_docx_mappe, "Gennemse ...")
        self._ordbog_felt(krop)
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

    def _ordbog_felt(self, parent):
        Label(parent, text="Ordbog (udbudsspecifik - bygger på en grundordbog for kontraktformen)",
              bg=WHITE, fg=GRAA, font=("Segoe UI", 9, "bold"), anchor="w",
              wraplength=520, justify=LEFT).pack(fill=X, padx=20, pady=(10, 4))
        række = Frame(parent, bg=WHITE)
        række.pack(fill=X, padx=20)
        self.ordbog_combo = ttk.Combobox(række, textvariable=self.ordbog_navn_var,
                                         values=self._ordbog_liste(), state="readonly",
                                         font=("Segoe UI", 10))
        self.ordbog_combo.pack(side=LEFT, fill=X, expand=True, ipady=2)
        self.ordbog_combo.bind("<<ComboboxSelected>>", self._on_ordbog_valgt)
        Button(række, text="Åbn i Excel", command=self.beskyt(self.aabn_ordbog_i_excel),
              bg=WHITE, fg=OXFORD, relief="solid", borderwidth=1,
              font=("Segoe UI", 9), cursor="hand2", padx=10).pack(side=LEFT, padx=(8, 0))
        Label(parent, textvariable=self.ordbog_info_var, bg=WHITE, fg=GRAA,
              font=("Segoe UI", 8), anchor="w").pack(fill=X, padx=20, pady=(3, 0))

        kant = Frame(parent, bg=CYAN)
        kant.pack(fill=X, padx=20, pady=(8, 0))
        self._knap(kant, "Generér ordbog fra kontrakt ...", self.beskyt(self.generer_ordbog_fra_kontrakt))

        self._genindlæs_sidste_ordbog()

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

    # -------------------------------------------------------------- ordbog ---
    def _ordbog_liste(self):
        """Scanner ordboeger-mappen for ordbog_udbud_*.xlsx og bygger navn->sti-listen.
        Navnet vises som Paradigme-feltet fra ordbogens Info-fane."""
        filer = sorted(ORDBOEGER_MAPPE.glob("ordbog_udbud_*.xlsx"), key=lambda p: p.stem.lower())
        self._ordbog_stier = {}
        for p in filer:
            try:
                navn = laes_ordbog_excel(p).get("paradigme") or p.stem
            except Exception:
                navn = p.stem
            if navn in self._ordbog_stier:                     # undgå navnekollision
                navn = f"{navn} ({p.stem})"
            self._ordbog_stier[navn] = p
        return list(self._ordbog_stier.keys())

    def _bygger_paa_tekst(self, sti):
        try:
            o = laes_ordbog_excel(sti)
        except Exception:
            return ""
        grund_navn = o.get("bygger_paa") or ""
        if not grund_navn:
            return "Grundordbog (kontraktform) - bygger ikke på andre ordbøger."
        try:
            grund = laes_ordbog_excel(Path(sti).parent / grund_navn)
            return f"Bygger på: {grund.get('paradigme') or grund_navn}"
        except Exception:
            return f"Bygger på: {grund_navn}"

    def _indlæs_state(self):
        try:
            return json.loads(STATE_STI.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _gem_state(self, **felter):
        tilstand = self._indlæs_state()
        tilstand.update(felter)
        try:
            STATE_STI.write_text(json.dumps(tilstand, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _genindlæs_sidste_ordbog(self):
        sidste = self._indlæs_state().get("sidste_ordbog")
        if not sidste:
            return
        navn = next((n for n, p in self._ordbog_stier.items() if p.name == sidste), None)
        if navn:
            self.ordbog_navn_var.set(navn)
            self.ordbog_info_var.set(self._bygger_paa_tekst(self._ordbog_stier[navn]))

    def _on_ordbog_valgt(self, event=None):
        navn = self.ordbog_navn_var.get()
        sti = self._ordbog_stier.get(navn)
        if sti is None:
            return
        self.ordbog_info_var.set(self._bygger_paa_tekst(sti))
        self._gem_state(sidste_ordbog=sti.name)
        self.sæt_status(f"Ordbog valgt: {navn}")

    def aabn_ordbog_i_excel(self):
        sti = self._ordbog_stier.get(self.ordbog_navn_var.get())
        if sti is None:
            messagebox.showwarning("Ingen ordbog valgt", "Vælg en ordbog i rullelisten først.")
            return
        self._åbn_fil(sti)

    def generer_ordbog_fra_kontrakt(self):
        mappe = self.docx_mappe_var.get().strip()
        if not mappe or not Path(mappe).is_dir():
            messagebox.showwarning("Manglende input",
                                   "Vælg mappen med docx (kontrakt + bilag) først.")
            return
        navn = simpledialog.askstring(
            "Navn på udbud", "Navn på udbud/projekt (bruges i filnavnet):", parent=self.root)
        if not navn or not navn.strip():
            return
        slug = re.sub(r"[^a-z0-9]+", "_", navn.strip().lower()).strip("_")
        if not slug:
            messagebox.showwarning("Ugyldigt navn", "Navnet gav intet brugbart filnavn.")
            return

        grund = sorted(ORDBOEGER_MAPPE.glob("ordbog_kontraktform_*.xlsx"))
        if not grund:
            messagebox.showerror("Ingen grundordbøger",
                                 f"Fandt ingen ordbog_kontraktform_*.xlsx i {ORDBOEGER_MAPPE}.")
            return

        self.sæt_status(f"Genererer ordbogsudkast fra {Path(mappe).name} ...")
        u = generer_udkast(mappe, grund)
        ORDBOEGER_MAPPE.mkdir(parents=True, exist_ok=True)
        ud = ORDBOEGER_MAPPE / f"ordbog_udbud_{slug}.xlsx"
        skriv_ordbog(u, ud)

        self.ordbog_combo["values"] = self._ordbog_liste()
        ny_navn = next((n for n, p in self._ordbog_stier.items() if p.name == ud.name), ud.stem)
        self.ordbog_navn_var.set(ny_navn)
        self.ordbog_info_var.set(self._bygger_paa_tekst(ud))
        self._gem_state(sidste_ordbog=ud.name)

        self.sæt_status(f"Ordbog genereret -> {ud.name}")
        messagebox.showinfo(
            "Ordbog genereret",
            f"Bygger på: {u['bygger_paa']}\nTilføjede milepæle: {len(u['milepaele'])}\n\n{ud}")
        self._åbn_fil(ud)

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
        ordbog_sti = self._ordbog_stier.get(self.ordbog_navn_var.get())
        ud = self.udtraek_xlsx_var.get().strip()

        if not mappe or not Path(mappe).is_dir():
            messagebox.showwarning("Manglende input", "Vælg en gyldig mappe med docx-filer.")
            return
        if ordbog_sti is None:
            messagebox.showwarning("Manglende input",
                                   "Vælg en ordbog i rullelisten (eller generér en fra kontrakten).")
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
