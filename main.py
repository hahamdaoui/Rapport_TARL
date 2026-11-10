#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rapport TARL : point d'entrée.
  python main.py                       -> interface graphique
  python main.py f1.xlsx [f2.xlsx] -o rapport.docx --lang ar|fr|en --prof "Nom" --directeur "Nom"
"""
import argparse
import os
import sys

if getattr(sys, "frozen", False):
    DOSSIER = os.path.dirname(sys.executable)
else:
    DOSSIER = os.path.dirname(os.path.abspath(__file__))
    if DOSSIER not in sys.path:
        sys.path.insert(0, DOSSIER)


def cli(argv=None):
    from rapport_tarl import ErreurRapport
    from rapport_tarl.pipeline import generate_reports, open_file
    ap = argparse.ArgumentParser(description="Génère un rapport Word (FR / AR / EN) à partir d'un ou deux fichiers Excel.")
    ap.add_argument("fichiers", nargs="+")
    ap.add_argument("-o", "--output", default="rapport_tarl.docx")
    ap.add_argument("--lang", default="fr", choices=["fr", "ar", "en"], help="langue du rapport")
    ap.add_argument("--prof", default=""); ap.add_argument("--directeur", default="")
    a = ap.parse_args(argv)
    try:
        states = [open_file(p) for p in a.fichiers]
        for s in states:
            print(f"{s.name}: feuille « {s.current.sheet.name} », {s.current.n_students} élèves")
        res = generate_reports(states, a.output, a.lang, a.prof, a.directeur, log=print)
        print("Rapport(s) créé(s) :", ", ".join(res))
    except ErreurRapport as e:
        sys.exit(f"Erreur : {e}")


def gui():
    from rapport_tarl.app_ui import main
    main()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        cli()
    else:
        gui()
