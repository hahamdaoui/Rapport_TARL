# -*- coding: utf-8 -*-
"""Briques de mise en page Word (python-docx) : titres de section, cartes de chiffres clés, encadrés,
tableaux, en-têtes / pieds de page. Compatibles LTR (français, anglais) et RTL (arabe)."""
import re

from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .arabic_rtl import mark_run_rtl, rtlize, set_par_bidi, split_script
from .text_utils import has_arabic

# ----------------------------------------------------------------------------- charte graphique (design system)
NAVY = "#1E3A5F"     # primaire
TEAL = "#1B7F8C"     # accent
ORG = "#E07B39"      # pré-test
BLU = "#2563EB"      # post-test
GRN = "#059669"      # progression
RED = "#DC2626"      # régression / faible
AMB = "#D97706"      # en cours
GRY = "#64748B"      # neutre
TXT = "#0F172A"
MUTED = "#475569"
FONT = "Segoe UI"
LARGEUR = 17.4
START, CENTER, END = WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT   # LEFT = « début » (bidi)


class Ctx:
    rtl = False


def set_rtl(flag):
    Ctx.rtl = bool(flag)


def hx(c):
    return c.lstrip("#")


def tint(c, f=.88):
    c = hx(c)
    r, g, b = (int(c[i:i + 2], 16) for i in (0, 2, 4))
    r, g, b = (int(v + (255 - v) * f) for v in (r, g, b))
    return f"#{r:02X}{g:02X}{b:02X}"


# ----------------------------------------------------------------------------- XML bas niveau
def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hx(color))
    tcPr.append(sh)


def cell_borders(cell, **edges):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}")
        v = edges.get(side)
        if v is None:
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(v[0])); e.set(qn("w:color"), hx(v[1])); e.set(qn("w:space"), "0")
        b.append(e)
    tcPr.append(b)


def cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    m = OxmlElement("w:tcMar")
    for side, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{side}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    tcPr.append(m)


def table_borders(table, color="D5DDE8", sz=4, inside=True):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right") + (("insideH", "insideV") if inside else ()):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:color"), hx(color)); e.set(qn("w:space"), "0")
        b.append(e)
    tblPr.append(b)


def fixer_largeurs(t, largeurs_cm):
    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for i, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(largeurs_cm[i] * 567)))
    for row in t.rows:
        for i, w in enumerate(largeurs_cm):
            row.cells[i].width = Cm(w)


def row_height(row, cm, rule="atLeast"):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight"); h.set(qn("w:val"), str(int(cm * 567))); h.set(qn("w:hRule"), rule); trPr.append(h)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    c = OxmlElement("w:cantSplit"); c.set(qn("w:val"), "true"); trPr.append(c)


# ----------------------------------------------------------------------------- texte
def run(par, text, size=10, bold=False, color=None, italic=False):
    """Ajoute du texte en séparant les portions arabes (RTL) des portions latines / chiffres."""
    last = None
    for seg, ar in ([(str(text), False)] if not has_arabic(text) else split_script(str(text))):
        if seg == "":
            continue
        r = par.add_run(seg)
        r.font.name = FONT
        r._r.get_or_add_rPr().rFonts.set(qn("w:cs"), FONT)
        r.font.size = Pt(size); r.bold = bold; r.italic = italic
        r.font.color.rgb = RGBColor.from_string(hx(color or TXT))
        if ar or Ctx.rtl:
            mark_run_rtl(r, size, bold) if ar else _cs_size(r, size, bold)
        last = r
        if ar and not Ctx.rtl:
            # texte arabe dans un document LTR : marque gauche-à-droite pour que les chiffres / codes / séparateurs
            # qui suivent ne soient pas aspirés dans le sens arabe (ex. « … 30 يونيو · 3APIC-1 »)
            m = par.add_run("\u200e")
            m.font.size = Pt(size)
    return last


def _cs_size(r, size, bold):
    rPr = r._r.get_or_add_rPr()
    s = OxmlElement("w:szCs"); s.set(qn("w:val"), str(int(size * 2))); rPr.append(s)
    if bold:
        rPr.append(OxmlElement("w:bCs"))


def para(parent, text="", size=10, bold=False, color=None, italic=False, align=None, before=0, after=4, keep=False):
    p = parent.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    if keep:
        pf.keep_with_next = True
    if align is not None:
        p.alignment = align
    if text != "":
        run(p, text, size, bold, color, italic)
    return p


def spacer(doc, pt=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(pt)
    return p


def nouvelle_page(doc):
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    return p


def puces(doc, items, size=10, color=None):
    """Liste à puces construite à la main (fonctionne identiquement en LTR et en RTL)."""
    for it in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.7); pf.first_line_indent = Cm(-0.45)
        pf.space_after = Pt(3)
        run(p, "•  ", size, True, TEAL)
        run(p, it, size, False, color)


def sous_titre(doc, text, color=NAVY):
    p = para(doc, text, 11.5, True, color, before=6, after=3, keep=True)
    pPr = p._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr"); b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "6"); b.set(qn("w:space"), "1"); b.set(qn("w:color"), hx(tint(color, .55)))
    bd.append(b); pPr.append(bd)
    return p


def note(doc, text):
    return para(doc, text, 8.5, color=MUTED, italic=True, before=2, after=4)


def legende(doc, items):
    """Légende de graphique en texte Word : [(couleur, libellé)] -> ■ libellé   ■ libellé."""
    p = para(doc, "", after=4, align=CENTER)
    for i, (col, lab) in enumerate(items):
        run(p, "■ ", 10, True, col)
        run(p, lab + ("      " if i < len(items) - 1 else ""), 9, False, MUTED)
    return p


def image(doc, fn, width_cm=16):
    p = doc.add_paragraph(); p.alignment = CENTER
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(fn, width=Cm(width_cm))


# ----------------------------------------------------------------------------- composants
def section_bar(doc, num, titre, color=NAVY):
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [1.2, LARGEUR - 1.2])
    c0, c1 = t.rows[0].cells
    for c in (c0, c1):
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_borders(c); cell_margins(c, 90, 90, 100, 140)
    shade(c0, color); shade(c1, tint(color, .9))
    cell_borders(c1, bottom=(12, color))
    p0 = c0.paragraphs[0]; p0.alignment = CENTER; p0.paragraph_format.space_after = Pt(0)
    run(p0, num, 16, True, "#FFFFFF")
    p1 = c1.paragraphs[0]; p1.paragraph_format.space_after = Pt(0)
    run(p1, titre, 15, True, NAVY)
    spacer(doc, 8)


def set_cell(cell, text, size=9, bold=False, color=None, fill=None, align=CENTER, margins=(45, 45, 80, 80)):
    """align : CENTER | START | 'name' (noms d'élèves : bord de lecture, sens RTL si le nom est arabe)."""
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    if align == "name":
        if has_arabic(text) or Ctx.rtl:
            set_par_bidi(p._p)
        p.alignment = START
    else:
        p.alignment = align
    run(p, text, size, bold, color)
    if fill:
        shade(cell, fill)
    cell_margins(cell, *margins)


def tableau(doc, entetes, lignes, largeurs_cm, fill_head=NAVY, size=9, align_cols=None, margins=(45, 45, 80, 80)):
    """lignes : listes de cellules ; cellule = texte ou (texte, {bold, color, fill, size})."""
    t = doc.add_table(rows=1, cols=len(entetes))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(t)
    for i, h in enumerate(entetes):
        set_cell(t.rows[0].cells[i], h, size, True, "#FFFFFF", fill_head, margins=(60, 60, 80, 80))
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for k, l in enumerate(lignes):
        row = t.add_row(); no_split(row)
        for i, v in enumerate(l):
            txt, o = v if isinstance(v, tuple) else (v, {})
            fill = o.get("fill", "#F4F7FB" if k % 2 == 0 else "#FFFFFF")
            al = (align_cols or {}).get(i, CENTER)
            set_cell(row.cells[i], txt, o.get("size", size), o.get("bold", False), o.get("color"), fill, al, margins)
    fixer_largeurs(t, largeurs_cm)
    return t


def kpis(doc, items, color=NAVY):
    n = len(items)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    w = LARGEUR / n
    fixer_largeurs(t, [w] * n)
    for i, it in enumerate(items):
        v, lab = it[0], it[1]
        col = it[2] if len(it) > 2 else color
        c = t.rows[0].cells[i]
        shade(c, tint(col, .92))
        cell_borders(c, top=(24, col), left=(4, "FFFFFF"), right=(4, "FFFFFF"), bottom=(4, tint(col, .6)))
        cell_margins(c, 90, 90, 60, 60)
        c.text = ""
        p = c.paragraphs[0]; p.alignment = CENTER; p.paragraph_format.space_after = Pt(0)
        run(p, v, 19, True, col)
        p2 = c.add_paragraph(); p2.alignment = CENTER; p2.paragraph_format.space_after = Pt(0)
        run(p2, lab, 8.5, False, MUTED)
    spacer(doc, 8)


def callout(doc, titre, texte=None, color=TEAL):
    """Encadré avec barre de couleur au bord de départ (gauche en LTR, droite en RTL). Retourne la cellule de contenu."""
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixer_largeurs(t, [0.28, LARGEUR - 0.28])
    a, c = t.rows[0].cells
    shade(a, color); cell_borders(a)
    a.paragraphs[0].paragraph_format.space_after = Pt(0)
    shade(c, tint(color, .92)); cell_borders(c); cell_margins(c, 90, 90, 160, 140)
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(2)
    run(p, titre, 10.5, True, color)
    if texte:
        p2 = c.add_paragraph(); p2.paragraph_format.space_after = Pt(0)
        run(p2, texte, 10, False, TXT)
    spacer(doc, 6)
    return c


def champ(par, instr, size=8, color=MUTED):
    for typ in ("begin", None, "separate", "txt", "end"):
        r = par.add_run()
        r.font.size = Pt(size); r.font.name = FONT; r.font.color.rgb = RGBColor.from_string(hx(color))
        if typ in ("begin", "separate", "end"):
            f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), typ); r._r.append(f)
        elif typ is None:
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "; r._r.append(it)
        else:
            t = OxmlElement("w:t"); t.text = "1"; r._r.append(t)


def en_tete_pied(doc, titre_court, pied_gauche, page_of_fmt):
    """En-tête (titre court + filet) et pied de page centré (texte + « Page X / Y »)."""
    s = doc.sections[0]
    s.different_first_page_header_footer = True
    h = s.header.paragraphs[0]
    h.paragraph_format.space_after = Pt(0)
    run(h, titre_court, 8.5, True, NAVY)
    pPr = h._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr"); b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "8"); b.set(qn("w:space"), "3"); b.set(qn("w:color"), hx(TEAL))
    bd.append(b); pPr.append(bd)
    f = s.footer.paragraphs[0]
    f.alignment = CENTER
    run(f, pied_gauche + "   ·   ", 8.5, False, MUTED)
    pre, post = page_of_fmt.split("{x}")
    mid, end = post.split("{y}")
    run(f, pre, 8.5, False, MUTED); champ(f, "PAGE", 8.5); run(f, mid, 8.5, False, MUTED); champ(f, "NUMPAGES", 8.5)
    if end:
        run(f, end, 8.5, False, MUTED)


_ORD = {
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
}


def normaliser(root):
    """Remet les éléments XML dans l'ordre exigé par Word (évite les messages « document endommagé »)."""
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        order = _ORD.get(tag)
        if not order:
            continue
        kids = list(el)
        key = lambda ch: order.index(ch.tag.split("}")[-1]) if ch.tag.split("}")[-1] in order else 999
        srt = sorted(kids, key=key)
        if srt != kids:
            for ch in kids:
                el.remove(ch)
            for ch in srt:
                el.append(ch)


def finaliser(doc, rtl):
    """RTL global (si demandé), ordre XML canonique, zoom valide."""
    parts = [doc.element] + [p._element for sec in doc.sections for p in
                             (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer)]
    for z in doc.settings.element.findall(qn("w:zoom")):
        if z.get(qn("w:percent")) is None:
            z.set(qn("w:percent"), "100")
    for root in parts:
        if rtl:
            rtlize(root)
        normaliser(root)
