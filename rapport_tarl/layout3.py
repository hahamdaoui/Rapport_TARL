# -*- coding: utf-8 -*-
"""Composants de mise en page du rapport compact « 3 pages » (inspiré du modèle fourni : bandeau marine, cartes de chiffres,
figures côte à côte, tableaux à en-tête coloré, encadrés). Fonctionne en LTR (FR / EN) et en RTL (AR)."""
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import docx_helpers as H
from .docx_helpers import (CENTER, END, START, cell_borders, cell_margins, fixer_largeurs, no_split, para, row_height, run, set_cell, shade,
                           spacer, table_borders, tint)

W = 17.6
NAVY, ORG, BLU, GRN, RED, AMB, TEAL = "#1B3A5C", "#E07B39", "#1F5FA8", "#2E9E5B", "#C0392B", "#D99A1E", "#1B7F8C"
TXT, MUTED, PANEL, LINE = "#27323D", "#6E7C88", "#F6F9FC", "#D9E1EA"


def new_doc3():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.7)
    s.top_margin = Cm(1.2); s.bottom_margin = Cm(1.5)
    s.header_distance = Cm(0.5); s.footer_distance = Cm(0.55)
    st = doc.styles["Normal"]; st.font.name = H.FONT; st.font.size = Pt(8)
    st.element.rPr.rFonts.set(qn("w:cs"), H.FONT)
    return doc


def footer3(doc, left_text, page_fmt):
    """Pied de page : titre à un bord, « page x sur y » à l'autre, filet au-dessus."""
    s = doc.sections[0]
    s.different_first_page_header_footer = False
    f = s.footer
    p0 = f.paragraphs[0]
    p0.paragraph_format.space_after = Pt(0); p0.paragraph_format.line_spacing = Pt(1)
    t = f.add_table(1, 2, Cm(W))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [W - 4.2, 4.2])
    a, b = t.rows[0].cells
    for c in (a, b):
        cell_borders(c, top=(6, LINE)); cell_margins(c, 40, 0, 0, 0)
    set_cell(a, left_text, 7.5, False, MUTED, None, START, (40, 0, 0, 0))
    cell_borders(a, top=(6, LINE))
    b.text = ""
    p = b.paragraphs[0]; p.alignment = END
    pre, post = page_fmt.split("{x}")
    mid, end = post.split("{y}")
    run(p, pre, 7.5, False, MUTED); H.champ(p, "PAGE", 7.5); run(p, mid, 7.5, False, MUTED); H.champ(p, "NUMPAGES", 7.5)
    if end:
        run(p, end, 7.5, False, MUTED)
    cell_borders(b, top=(6, LINE)); cell_margins(b, 40, 0, 0, 0)
    p2 = f.add_paragraph(); p2.paragraph_format.space_after = Pt(0); p2.paragraph_format.line_spacing = Pt(1)


def banner(doc, kicker, subtitle, line3):
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [W])
    c = t.rows[0].cells[0]
    shade(c, NAVY); cell_borders(c); cell_margins(c, 150, 150, 200, 200)
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c.paragraphs[0]; p.alignment = CENTER; p.paragraph_format.space_after = Pt(3)
    run(p, kicker, 11, False, "#9DB9D9")
    para(c, subtitle, 12.5, True, "#FFFFFF", align=CENTER, after=3)
    para(c, line3, 10, False, "#C6D6E8", align=CENTER, after=0)
    spacer(doc, 5)


def sec_header(doc, num, title, right, color):
    t = doc.add_table(rows=1, cols=3); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [1.1, W - 1.1 - 5.0, 5.0])
    for i, c in enumerate(t.rows[0].cells):
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_borders(c, bottom=(8, color)); cell_margins(c, 60, 50, 40, 40)
    c0, c1, c2 = t.rows[0].cells
    for c, txt, sz, bold, col, al in ((c0, num, 13, True, color, START), (c1, title, 11, True, NAVY, START), (c2, right, 8.5, False, MUTED, END)):
        c.text = ""
        p = c.paragraphs[0]; p.alignment = al; p.paragraph_format.space_after = Pt(0)
        run(p, txt, sz, bold, col)
    spacer(doc, 5)


def kpi_cards(doc, items, color):
    """items = [(valeur, libellé)] : cartes grises, valeur colorée en grand, libellé gris dessous."""
    n = len(items)
    t = doc.add_table(rows=1, cols=n); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [W / n] * n)
    for c, (v, lab) in zip(t.rows[0].cells, items):
        shade(c, PANEL)
        cell_borders(c, left=(24, "FFFFFF"), right=(24, "FFFFFF")); cell_margins(c, 90, 90, 120, 80)
        c.text = ""
        p = c.paragraphs[0]; p.alignment = START; p.paragraph_format.space_after = Pt(0)
        run(p, v, 14, True, color)
        p2 = c.add_paragraph(); p2.alignment = START; p2.paragraph_format.space_after = Pt(0)
        run(p2, lab, 8, False, MUTED)
    spacer(doc, 5)


def fig_row(doc, items, widths):
    """items = [(chemin image, légende, largeur image cm)] dans une carte à fin contour."""
    n = len(items)
    t = doc.add_table(rows=1, cols=n); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(t, LINE, 4, inside=False)
    fixer_largeurs(t, widths)
    for c, (img, cap, wcm) in zip(t.rows[0].cells, items):
        cell_margins(c, 90, 60, 100, 100)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        c.text = ""
        p = c.paragraphs[0]; p.alignment = CENTER; p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(img, width=Cm(wcm))
        para(c, cap, 8, False, MUTED, align=CENTER, after=0)
    spacer(doc, 5)


def titre(doc, text, color=NAVY, size=9.5):
    return para(doc, text, size, True, color, before=2, after=2, keep=True)


def tri_bullets(doc, items, color=ORG):
    for it in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.5); pf.first_line_indent = Cm(-0.4); pf.space_after = Pt(1.5)
        run(p, "▸  ", 8, True, color)
        run(p, it, 8, False, TXT)


def box(doc, titre_txt, text, color, fill=PANEL):
    """Encadré avec barre colorée au bord de départ."""
    t = doc.add_table(rows=1, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [0.12, W - 0.12])
    a, c = t.rows[0].cells
    shade(a, color); cell_borders(a)
    shade(c, fill); cell_borders(c); cell_margins(c, 80, 80, 160, 140)
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(1)
    run(p, titre_txt, 8.5, True, color)
    if text:
        p2 = c.add_paragraph(); p2.paragraph_format.space_after = Pt(0)
        run(p2, text, 8, False, TXT)
    spacer(doc, 5)
    return c


def two_boxes(doc, boxes):
    """Deux encadrés côte à côte : [(titre, texte, couleur, fond)]."""
    t = doc.add_table(rows=1, cols=5); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    wc = (W - 0.4) / 2
    fixer_largeurs(t, [0.12, wc - 0.12, 0.4, 0.12, wc - 0.12])
    cells = t.rows[0].cells
    for (ti, tx, col, fill), (ca, cb) in zip(boxes, ((cells[0], cells[1]), (cells[3], cells[4]))):
        shade(ca, col); cell_borders(ca)
        shade(cb, fill); cell_borders(cb); cell_margins(cb, 80, 80, 160, 140)
        cb.text = ""
        p = cb.paragraphs[0]; p.paragraph_format.space_after = Pt(1)
        run(p, ti, 8.5, True, col)
        p2 = cb.add_paragraph(); p2.paragraph_format.space_after = Pt(0)
        run(p2, tx, 8, False, TXT)
    cell_borders(cells[2])
    spacer(doc, 5)


def table(doc, heads, rows, widths, color, align_cols=None, size=8):
    t = H.tableau(doc, heads, rows, widths, fill_head=color, size=size, align_cols=align_cols, margins=(35, 35, 80, 80))
    spacer(doc, 5)
    return t


def legend_inline(doc, items):
    return H.legende(doc, items)
