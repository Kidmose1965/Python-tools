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
BTN = {"width": 34, "pady": 6, "bg": "#DCE6F1", "activebackground": "#C5D9F1",
       "relief": "groove", "anchor": "w", "padx": 12}


class App:
    def __init__(self, root):
        self.root = root
        root.title("Extractor - udtræk fra kravspecifikationer")
        root.configure(bg=BG)
        root.resizable(False, False)

        Label(root, text="Vælg udtræk", bg=BG,
              font=("Segoe UI", 13, "bold")).pack(pady=(14, 8))

        knapper = [
            ("Udtræk kommentarer", self.kommentarer),
            ("Udtræk trackchanges", self.trackchanges),
            ("Udtræk kundens inputfelter (grøn)", lambda: self.inputfelter("groen")),
            ("Udtræk tilbudsgivers inputfelter (gul)", lambda: self.inputfelter("gul")),
            ("Skab kravmatrix", self.kravmatrix),
            ("Validér krydshenvisninger (kontrakt + bilag)", self.krydstjek),
            ("Vis dokumentets typografier (styles)", self.styles),
        ]
        for tekst, cmd in knapper:
            Button(root, text=tekst, command=self.beskyt(cmd), **BTN).pack(padx=16, pady=3)

        self.status = Label(root, text="Klar.", bg=BG, fg="#444",
                            font=("Segoe UI", 9), anchor="w")
        self.status.pack(fill=X, padx=16, pady=(10, 12))

    # ------------------------------------------------------------------ utils
    def beskyt(self, fn):
        """Fang alle fejl og vis dem pænt i stedet for at lukke vinduet."""
        def wrapper():
            try:
                fn()
            except Exception:
                messagebox.showerror("Fejl", "Der opstod en uventet fejl:\n\n"
                                     + traceback.format_exc(limit=3))
                self.sæt_status("Fejl - se fejlbesked.")
        return wrapper

    def sæt_status(self, tekst):
        self.status.config(text=tekst)
        self.root.update_idletasks()

    def vælg_filer(self, flere=True):
        fn = filedialog.askopenfilenames if flere else filedialog.askopenfilename
        valg = fn(title="Vælg kravspecifikation(er)",
                  filetypes=[("Word-dokumenter", "*.docx")])
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
        docs = []
        for f in filer:
            self.sæt_status(f"Læser {Path(f).name} ...")
            docs.append(ex.Docx(f))
        return docs

    # ------------------------------------------------------------ udtrækkene
    def kommentarer(self):
        filer = self.vælg_filer()
        if not filer:
            return
        recs = []
        for d in self.læs(filer):
            recs += ex.extract_comments(d) or [{"Dokument": d.path.name,
                                                "Kommentar": "Ingen kommentarer"}]
        sti = self.gem_som("kommentarer.xlsx")
        if not sti:
            return
        ex.write_simple(recs, ["Dokument", "Nummer", "Kommentar", "Markeret tekst",
                               "Initialer", "Nummereret sektion"], sti,
                        [35, 9, 50, 50, 12, 50])
        self.færdig(sti, len(recs))

    def trackchanges(self):
        filer = self.vælg_filer()
        if not filer:
            return
        recs = []
        for d in self.læs(filer):
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
        recs = []
        for d in self.læs(filer):
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
        self.sæt_status("Kortlægger definitioner og validerer henvisninger ...")
        fund, sektioner, bilag_def = kt.validér(docs)
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

    def styles(self):
        filer = self.vælg_filer()
        if not filer:
            return
        tekst = ""
        for d in self.læs(filer):
            tekst += f"{d.path.name}:\n" + "\n".join(
                f"   - {s}" for s in ex.list_styles(d)) + "\n\n"
        self.sæt_status("Typografier fundet.")
        messagebox.showinfo("Dokumentets typografier", tekst.strip())

    def kravmatrix(self):
        filer = self.vælg_filer(flere=False)
        if not filer:
            return
        d = self.læs(filer)[0]
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
