#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dobbeltklik-starter til tidsplan_gui.py - starter GUI'en uden det sorte
konsolvindue (Windows knytter .pyw til pythonw.exe, som ikke åbner et
konsolvindue - i modsætning til .py, der knyttes til py.exe).
"""
import runpy
from pathlib import Path

runpy.run_path(str(Path(__file__).resolve().parent / "tidsplan_gui.py"), run_name="__main__")
