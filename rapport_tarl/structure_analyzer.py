# -*- coding: utf-8 -*-
"""Analyse automatique de la structure d'une feuille Excel : zone des élèves, en-têtes (même sur plusieurs
lignes / cellules fusionnées), métadonnées (matière, classe, établissement, type de test…), langue."""
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from . import ErreurRapport
from .column_detector import best_header_role, detect_columns, summary
from .excel_reader import SheetModel, WorkbookModel, _empty
from .language_detector import detect_language
from .text_utils import has_arabic, normalize, to_number

SUMMARY_WORDS = {"المجموع", "المعدل", "المعدل العام", "المتوسط", "المجموع العام", "total", "moyenne", "moyenne generale", "somme",
                 "mean", "average", "effectif", "عدد", "النسبه", "نسبه", "pourcentage", "max", "min", "ecart type", "mediane", "median", "sum"}

META_KEYS = {
    "academie": ["الاكاديميه", "academie", "academy", "aref", "الاكاديمية الجهوية"],
    "direction": ["المديريه الاقليميه", "المديريه", "delegation", "direction provinciale", "direction", "directorate", "province", "مديريه"],
    "etablissement": ["المؤسسه", "المؤسسة", "etablissement", "ecole", "college", "lycee", "school", "institution", "الثانويه", "مؤسسه"],
    "annee": ["السنه الدراسيه", "السنه الدراسية", "annee scolaire", "annee", "school year", "academic year", "year"],
    "matiere": ["الماده", "matiere", "subject", "discipline", "course"],
    "groupe": ["مجموعه الاختبار", "القسم", "الفوج", "المجموعه", "classe", "groupe", "class", "group", "section"],
    "type_test": ["نوع الاختبار", "type de test", "type d evaluation", "test type", "type d examen", "type"],
    "enseignant": ["الاستاذ", "الاستاذه", "المدرس", "professeur", "enseignant", "teacher", "prof"],
    "periode": ["الدوره", "الاسدس", "الفصل", "semestre", "trimestre", "periode", "term", "semester", "period"],
}
_META_NORM = {k: [normalize(x) for x in v] for k, v in META_KEYS.items()}


@dataclass
class ColumnInfo:
    idx: int
    letter: str
    header: str = ""            # intitulé principal (ligne la plus proche des données)
    group: str = ""             # intitulé de groupe (ligne supérieure : période, partie…)
    header_full: str = ""
    item_code: str = ""         # code éventuel entre parenthèses : (LTC)
    values: List[Any] = field(default_factory=list)
    role: str = "ignore"
    confidence: float = 0
    status: str = "ignored"     # recognized | verify | ignored
    reason: str = ""
    invert: bool = False        # colonne « présence » (oui = présent)
    profile: Dict[str, Any] = field(default_factory=dict)

    @property
    def display(self):
        h = f"{self.group} — {self.header}" if self.group and self.header and self.group != self.header else (self.header or self.group)
        return h or f"(colonne {self.letter} sans en-tête)"


@dataclass
class SheetAnalysis:
    sheet: SheetModel
    ok: bool = False
    meta: Dict[str, str] = field(default_factory=dict)
    header_rows: List[int] = field(default_factory=list)
    student_rows: List[int] = field(default_factory=list)
    summary_rows: List[int] = field(default_factory=list)
    name_col: Optional[int] = None
    columns: List[ColumnInfo] = field(default_factory=list)
    language: str = "unknown"
    language_detail: Dict[str, Any] = field(default_factory=dict)
    data_language: str = "unknown"
    test_type: str = "?"
    problem: str = ""

    @property
    def n_students(self):
        return len(self.student_rows)


def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


# ------------------------------------------------------------------------------------------ zone des élèves
def _name_like(v):
    if not isinstance(v, str):
        return False
    s = v.strip()
    if len(s) < 5 or len(s) > 70:
        return False
    n = normalize(s)
    toks = n.split()
    if len(toks) < 2:
        return False
    letters = sum(ch.isalpha() for ch in n)
    digits = sum(ch.isdigit() for ch in n)
    return letters >= 5 and digits <= 0.15 * len(n.replace(" ", "")) and to_number(s) is None


def _find_student_block(sh: SheetModel):
    best = (0, None, [])
    for c in range(sh.ncols):
        rows = [r for r in range(sh.nrows) if _name_like(sh.cell(r, c)) and (r, c) not in sh.filled_by_merge]
        if len(rows) < 2:
            continue
        # découpage en séries : écart maximal toléré de 3 lignes
        runs, cur = [], [rows[0]]
        for r in rows[1:]:
            if r - cur[-1] <= 3:
                cur.append(r)
            else:
                runs.append(cur); cur = [r]
        runs.append(cur)
        run = max(runs, key=len)
        if len(run) > best[0]:
            best = (len(run), c, run)
    return best


def _header_rows(sh: SheetModel, first_row, name_col):
    rows = []
    r = first_row - 1
    while r >= 0 and len(rows) < 4:
        cells = [sh.cell(r, c) for c in range(sh.ncols)]
        filled = [v for v in cells if not _empty(v)]
        if len(filled) < 2:
            break
        rows.append(r)
        r -= 1
    return sorted(rows)


def _parts(sh, c, header_rows):
    parts = []
    for r in header_rows:
        v = sh.cell(r, c)
        if _empty(v):
            continue
        t = str(v).strip()
        if to_number(t) is not None and re.fullmatch(r"[\d\s.,٠-٩]+", t):
            continue                      # ligne de codes numériques : pas un intitulé
        if not parts or parts[-1] != t:
            parts.append(t)
    return parts


# ------------------------------------------------------------------------------------------ métadonnées
def _extract_meta(sh: SheetModel, upto_row):
    meta = {}
    for r in range(0, min(upto_row, sh.nrows)):
        for c in range(sh.ncols):
            v = sh.cell(r, c)
            if not isinstance(v, str) or (r, c) in sh.filled_by_merge:
                continue
            n = normalize(v)
            if not n or len(n) > 40:
                continue
            for key, syns in _META_NORM.items():
                if key in meta:
                    continue
                if n in syns:
                    # valeur = première cellule non vide à droite qui n'est pas elle-même un libellé connu
                    for cc in range(c + 1, min(sh.ncols, c + 8)):
                        w = sh.cell(r, cc)
                        if _empty(w) or (r, cc) in sh.filled_by_merge:
                            continue
                        wn = normalize(w)
                        if any(wn == s for ss in _META_NORM.values() for s in ss):
                            break
                        meta[key] = str(w).strip()
                        break
    if "direction" in meta:
        meta["direction"] = re.sub(r"^(إقليم|اقليم|province)\s*:?\s*", "", meta["direction"], flags=re.I).strip()
    return meta


def detect_test_type(*texts):
    """'pre' / 'post' / '?' d'après des textes (type de test, nom de feuille, nom de fichier)."""
    for t in texts:
        n = normalize(t)
        if not n:
            continue
        if re.search(r"قبلي|\bpre\b|pretest|pre test|initial|diagnostic|entree|debut|baseline|before|\bt0\b", n):
            return "pre"
        if re.search(r"بعدي|\bpost\b|posttest|post test|final|sortie|fin d|after|endline|\bt1\b", n):
            return "post"
    return "?"


# ------------------------------------------------------------------------------------------ analyse
def analyze_sheet(sh: SheetModel, filename="") -> SheetAnalysis:
    an = SheetAnalysis(sheet=sh)
    if sh.nrows < 2 or sh.ncols < 2:
        an.problem = "feuille vide ou trop petite"
        return an
    n_found, name_col, run = _find_student_block(sh)
    if n_found < 2:
        an.problem = "aucune liste d'élèves reconnue (au moins 2 noms sur plusieurs mots sont nécessaires)"
        return an
    # la ligne d'en-tête (« الإسم والنسب », « Nom et prénom »…) peut ressembler à un nom : on l'écarte
    def _looks_like_header(r):
        role, sc, _ = best_header_role(sh.cell(r, name_col))
        if role == "name" and sc >= 78:
            return True
        # texte au-dessus de nombres sur plusieurs colonnes = ligne d'en-têtes
        flips = 0
        for c in range(sh.ncols):
            a, b = sh.cell(r, c), sh.cell(r + 1, c)
            if isinstance(a, str) and to_number(a) is None and to_number(b) is not None:
                flips += 1
        return flips >= 2
    while len(run) > 2 and _looks_like_header(run[0]):
        run = run[1:]
    first, last = run[0], run[-1]
    # prolonger la liste vers le bas : les lignes suivantes qui contiennent un texte dans la colonne des noms
    # (noms à un seul mot, noms avec chiffre ou abréviation…) font aussi partie des élèves, sauf lignes de totaux
    gap, r = 0, last + 1
    while r < sh.nrows and gap <= 3:
        v = sh.cell(r, name_col)
        if isinstance(v, str) and v.strip() and normalize(v) not in {normalize(w) for w in SUMMARY_WORDS} \
                and any(not _empty(sh.cell(r, c)) for c in range(sh.ncols) if c != name_col):
            last, gap = r, 0
        else:
            gap += 1
        r += 1
    an.name_col = name_col
    an.header_rows = _header_rows(sh, first, name_col)
    hdr_top = an.header_rows[0] if an.header_rows else first
    an.meta = _extract_meta(sh, hdr_top)
    # lignes d'élèves (on exclut les lignes de totaux et les lignes totalement vides)
    for r in range(first, last + 1):
        nm = sh.cell(r, name_col)
        if _empty(nm):
            continue
        if normalize(nm) in {normalize(w) for w in SUMMARY_WORDS}:
            an.summary_rows.append(r)
            continue
        if _name_like(nm) or any(not _empty(sh.cell(r, c)) for c in range(sh.ncols) if c != name_col):
            an.student_rows.append(r)
    # colonnes
    for c in range(sh.ncols):
        parts = _parts(sh, c, an.header_rows)
        header = parts[-1] if parts else ""
        group = parts[0] if len(parts) > 1 else ""
        m = re.search(r"\(([A-Za-z0-9_\-]{2,8})\)", header)
        code = m.group(1) if m else (header if re.fullmatch(r"[A-Za-z]{1,4}\d{0,2}", header) else "")
        col = ColumnInfo(idx=c, letter=col_letter(c), header=header, group=group, header_full=" | ".join(parts),
                         item_code=code, values=[sh.cell(r, c) for r in an.student_rows])
        an.columns.append(col)
    detect_columns(an)
    # langues
    an.language, an.language_detail = detect_language([c.header_full for c in an.columns] + list(an.meta.values()))
    an.data_language, _ = detect_language([sh.cell(r, name_col) for r in an.student_rows])
    an.test_type = detect_test_type(an.meta.get("type_test", ""), sh.name, filename)
    an.ok = True
    return an


def analyze_workbook(wb: WorkbookModel):
    """Analyse toutes les feuilles ; retourne la liste triée (la meilleure feuille en premier)."""
    import os
    fn = os.path.basename(wb.path)
    res = [analyze_sheet(s, fn) for s in wb.sheets]
    res.sort(key=lambda a: (not a.ok, -a.n_students))
    return res


def best_analysis(wb: WorkbookModel):
    res = analyze_workbook(wb)
    if not res or not res[0].ok:
        raise ErreurRapport("Aucune liste d'élèves n'a été reconnue dans ce fichier.\n"
                            "Vérifiez qu'il contient bien une colonne avec les noms des élèves (prénom et nom) "
                            "et que la feuille n'est pas vide.")
    return res[0], res


def describe(an: SheetAnalysis):
    s = summary(an.columns)
    return (f"feuille « {an.sheet.name} » · {an.n_students} élèves · {len(an.columns)} colonnes "
            f"({s['recognized']} reconnues, {s['verify']} à vérifier, {s['ignored']} ignorées)")
