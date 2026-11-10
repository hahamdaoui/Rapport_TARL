# -*- coding: utf-8 -*-
"""Génération du rapport Word (DOCX) en français, arabe (RTL) ou anglais, à partir des résultats du moteur.
Les sections qui ne peuvent pas être calculées avec les données disponibles ne sont jamais affichées."""
import datetime
import os
import re
import shutil
import tempfile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import ErreurRapport, charts
from . import docx_helpers as H
from .docx_helpers import (AMB, BLU, CENTER, END, GRN, GRY, NAVY, ORG, RED, START, TEAL, MUTED, TXT, LARGEUR,
                           callout, cell_borders, cell_margins, fixer_largeurs, image, kpis, legende, no_split,
                           nouvelle_page, note, para, puces, row_height, run, section_bar, set_cell, shade, spacer,
                           sous_titre, table_borders, tableau, tint)
from .report_grade import build_grade
from .report_skill import build_skill
from .report_translator import Translator, subject_name
from .text_utils import normalize

SEUIL_ACQUIS, SEUIL_COURS = 60, 35


def _level(T, p):
    if p is None or p != p:
        return T("na"), GRY
    if p >= SEUIL_ACQUIS:
        return T("lvl_ach"), GRN
    if p >= SEUIL_COURS:
        return T("lvl_prog"), AMB
    return T("lvl_weak"), RED


def _names(lst):
    return " ، ".join(lst)


# =============================================================================================== cadre commun
def _new_doc(T, kind, meta, subject):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.5); s.top_margin = Cm(1.8); s.bottom_margin = Cm(1.5)
    s.header_distance = Cm(0.8); s.footer_distance = Cm(0.7)
    st = doc.styles["Normal"]; st.font.name = H.FONT; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:cs"), H.FONT)
    group = meta.get("groupe", "")
    H.en_tete_pied(doc, T("hdr_" + kind, s=subject, g=group), T("ftr", y=meta.get("annee", "")), T("page_of", x="{x}", y="{y}"))
    return doc


def _cover(doc, T, kind, meta, subject, info, kpi_items):
    strip = doc.add_table(rows=1, cols=3); strip.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(strip, [6, 6, 6])
    for c, col in zip(strip.rows[0].cells, (ORG, BLU, GRN)):
        shade(c, col); cell_borders(c); c.paragraphs[0].paragraph_format.space_after = Pt(0)
        c.paragraphs[0].paragraph_format.line_spacing = Pt(8)
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [18])
    c = t.rows[0].cells[0]; row_height(t.rows[0], 10.5)
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    shade(c, NAVY); cell_borders(c); cell_margins(c, 300, 300, 500, 500)
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(10)
    run(p, T("kicker_" + kind), 10, True, "#8FD0D9")
    para(c, T("title_" + kind), 36 if not T.rtl else 32, True, "#FFFFFF", after=2)
    para(c, subject, 24, False, "#FFFFFF", after=14)
    para(c, T("subtitle_" + kind), 12, False, "#C9D6E6", after=26)
    para(c, meta.get("etablissement", ""), 14, True, "#FFFFFF", after=2)
    para(c, T("year_group", y=meta.get("annee", ""), g=meta.get("groupe", "")), 11, False, "#C9D6E6", after=0)
    spacer(doc, 14)
    sous_titre(doc, T("ident"), NAVY)
    ti = doc.add_table(rows=0, cols=4); ti.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(ti)
    for i in range(0, len(info), 2):
        cells = ti.add_row().cells
        for j in range(2):
            lab, val = info[i + j]
            set_cell(cells[2 * j], lab, 9.5, True, NAVY, tint(NAVY, .92), START, (70, 70, 110, 80))
            set_cell(cells[2 * j + 1], val, 9.5, False, None, "#FFFFFF", START, (70, 70, 110, 80))
    fixer_largeurs(ti, [3.6, 5.1, 3.6, 5.1])
    spacer(doc, 12)
    sous_titre(doc, T("glance"), NAVY)
    kpis(doc, kpi_items)
    para(doc, T("issued", d=datetime.date.today().strftime("%d/%m/%Y")), 9, False, MUTED, italic=True, align=END, before=6)


def _signatures(doc, T, prof, director):
    spacer(doc, 30)
    sg = doc.add_table(rows=1, cols=2); sg.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(sg, [8.7, 8.7])
    row_height(sg.rows[0], 3.6)
    for c, (role, nom) in zip(sg.rows[0].cells, ((T("sign_teacher"), prof), (T("sign_director"), director))):
        cell_borders(c, top=(6, TEAL), left=(4, "FFFFFF"), right=(4, "FFFFFF")); cell_margins(c, 100, 100, 140, 140)
        c.text = ""
        p = c.paragraphs[0]; p.alignment = CENTER; p.paragraph_format.space_after = Pt(2)
        run(p, role, 11, True, NAVY)
        p2 = c.add_paragraph(); p2.alignment = CENTER
        run(p2, nom or T("dots"), 10, False, TXT)
        p3 = c.add_paragraph(); p3.alignment = CENTER
        run(p3, T("sign_stamp"), 8.5, False, "#94A3B8", italic=True)


def _info_rows(T, meta, subject, prof, director, extra):
    base = [(T("academy"), meta.get("academie", "")), (T("direction"), meta.get("direction", "")),
            (T("school"), meta.get("etablissement", "")), (T("year"), meta.get("annee", "")),
            (T("subject"), subject), (T("group"), meta.get("groupe", "")),
            (T("teacher"), prof or meta.get("enseignant", "") or T("dots")), (T("director"), director or T("dots"))]
    keep = {T("teacher"), T("director")}
    rows = [(a, b) for a, b in base if b or a in keep] + extra
    if len(rows) % 2:
        rows.append(("", ""))
    return rows


# =============================================================================================== point d'entrée
def generate(result, lang, sortie, prof="", director="", log=print):
    tmp = tempfile.mkdtemp()
    try:
        builder = build_skill if result["mode"] == "skill" else build_grade
        doc, T = builder(result, lang, sortie, prof, director, log, tmp)
        H.finaliser(doc, T.rtl)
        try:
            doc.save(sortie)
        except PermissionError:
            raise ErreurRapport("Impossible d'enregistrer le rapport : le fichier est ouvert dans Word ?\nFermez-le puis recommencez.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        H.set_rtl(False)
    log(f"Rapport enregistré : {sortie}")
