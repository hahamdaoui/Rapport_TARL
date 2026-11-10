# -*- coding: utf-8 -*-
"""Mise en page de droite à gauche (RTL) pour les rapports arabes, et gestion des textes mixtes.

Principes appliqués :
  * tous les paragraphes sont marqués « bidi » (alignement et retraits « début / fin » inversés) ;
  * tous les tableaux sont en « bidiVisual » : la première colonne est à DROITE ;
  * seules les portions de texte arabe reçoivent l'attribut RTL (avec taille et police « complex script »),
    ce qui évite les lettres inversées ou séparées dans les textes mixtes (arabe + codes latins + chiffres) ;
  * police Arial (présente sur Windows 10/11) pour le latin ET l'arabe ;
  * aucun texte arabe n'est écrit dans les images des graphiques (pas de problème de lettres séparées) :
    les titres et légendes des graphiques sont du texte Word."""
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ARABIC_SEG = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF](?:[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF\s،؛؟\-\u2013/]*[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF])?")


def split_script(text):
    """Découpe un texte en [(segment, est_arabe)] en conservant l'ordre."""
    out, pos = [], 0
    for m in ARABIC_SEG.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(0), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out or [(text, False)]


def mark_run_rtl(run, size_pt, bold):
    """Attributs « complex script » d'un run arabe : sens RTL, taille et gras."""
    rPr = run._r.get_or_add_rPr()
    for tag in ("w:rtl", "w:szCs", "w:bCs", "w:lang"):
        for old in rPr.findall(qn(tag)):
            rPr.remove(old)
    szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), str(int(size_pt * 2))); rPr.append(szcs)
    if bold:
        rPr.append(OxmlElement("w:bCs"))
    rPr.append(OxmlElement("w:rtl"))
    lang = OxmlElement("w:lang"); lang.set(qn("w:bidi"), "ar-MA"); rPr.append(lang)


def set_par_bidi(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None:
        pPr.append(OxmlElement("w:bidi"))


def set_table_rtl(tbl_el):
    tblPr = tbl_el.tblPr
    if tblPr.find(qn("w:bidiVisual")) is None:
        tblPr.append(OxmlElement("w:bidiVisual"))


def rtlize(root):
    """Applique la mise en page RTL à tout un arbre XML (corps, en-tête, pied de page)."""
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        if tag == "p":
            set_par_bidi(el)
        elif tag == "tbl":
            tblPr = el.find(qn("w:tblPr"))
            if tblPr is not None and tblPr.find(qn("w:bidiVisual")) is None:
                tblPr.append(OxmlElement("w:bidiVisual"))
