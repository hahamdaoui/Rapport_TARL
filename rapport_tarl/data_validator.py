# -*- coding: utf-8 -*-
"""Contrôles avant traitement. Ne supprime et ne modifie JAMAIS les données : chaque anomalie est signalée
avec une explication claire et une proposition de correction."""
from dataclasses import dataclass, field
from statistics import mean, pstdev
from typing import List

from .data_normalizer import Dataset
from .text_utils import normalize, to_number


@dataclass
class Issue:
    level: str          # 'error' (bloquant) | 'warning' | 'info'
    title: str
    detail: str = ""
    suggestion: str = ""
    rows: List[int] = field(default_factory=list)


def _rows_txt(rows, n=6):
    rows = sorted(set(rows))
    s = ", ".join(str(r) for r in rows[:n])
    return s + (f" … (+{len(rows) - n})" if len(rows) > n else "")


def validate(ds: Dataset, an=None):
    iss: List[Issue] = []
    names_ok = any(r == "name" for r in ds.mapping.values())
    if not names_ok:
        iss.append(Issue("error", "Aucune colonne ne permet d'identifier les élèves",
                         "Aucune colonne n'est associée à « Nom de l'élève ».",
                         "Choisissez la colonne des noms dans la liste « Fonction détectée »."))
    if not ds.items:
        iss.append(Issue("error", "Aucune colonne de résultats",
                         "Aucune colonne n'est associée à « Note d'évaluation » ou « Compétence validée (0/1) ».",
                         "Indiquez au moins une colonne de notes (ou de compétences 0/1)."))
    if not ds.students:
        iss.append(Issue("error", "Aucun élève trouvé", "La liste des élèves est vide.", "Vérifiez la feuille choisie."))
        return iss
    if names_ok and not ds.has_id:
        iss.append(Issue("info", "Pas d'identifiant d'élève",
                         "Les élèves seront reconnus uniquement par leur nom (utile si deux fichiers doivent être comparés).",
                         "Si une colonne contient le numéro Massar / matricule, associez-la à « Identifiant »."))
    # doublons
    seen, dup = {}, []
    for s in ds.students:
        k = s.key
        if k in seen:
            dup.append(s.row); dup.append(seen[k])
        else:
            seen[k] = s.row
    if dup:
        iss.append(Issue("warning", "Doublons possibles", f"Même identifiant ou même nom aux lignes : {_rows_txt(dup)}.",
                         "Vérifiez ces lignes dans Excel : un élève est peut-être saisi deux fois.", dup))
    # cellules vides / illisibles / hors barème
    for it in ds.items:
        empty_rows, bad_rows, out_rows, bad_vals = [], [], [], []
        vals = []
        for s in ds.students:
            raw = s.raw.get(it.key)
            blank = raw is None or str(raw).strip() == ""
            if s.absent:
                continue
            if blank:
                empty_rows.append(s.row)
                continue
            v = s.scores.get(it.key)
            if v is None:
                bad_rows.append(s.row); bad_vals.append(str(raw))
                continue
            vals.append(v)
            if it.kind == "skill" and v not in (0.0, 1.0):
                out_rows.append(s.row)
            if it.kind == "grade" and (v < 0 or (it.scale and v > it.scale)):
                out_rows.append(s.row)
        lab = it.short
        if bad_rows:
            iss.append(Issue("warning", f"Valeurs non numériques dans « {lab} »",
                             f"{len(bad_rows)} cellule(s) illisible(s) (lignes {_rows_txt(bad_rows)}), par ex. « {bad_vals[0]} ». "
                             "Elles seront ignorées dans les calculs.",
                             "Corrigez la cellule dans Excel (nombre uniquement) ou laissez-la vide.", bad_rows))
        if out_rows:
            lim = "0 ou 1" if it.kind == "skill" else f"0 à {it.scale:g}"
            iss.append(Issue("warning", f"Valeurs hors barème dans « {lab} »",
                             f"{len(out_rows)} valeur(s) hors de l'intervalle autorisé ({lim}) : lignes {_rows_txt(out_rows)}.",
                             "Vérifiez la saisie, ou choisissez le bon barème dans « Barème des notes ».", out_rows))
        if empty_rows and len(empty_rows) < len(ds.students) - sum(1 for s in ds.students if s.absent):
            iss.append(Issue("info", f"Cellules vides dans « {lab} »",
                             f"{len(empty_rows)} élève(s) sans résultat (lignes {_rows_txt(empty_rows)}). Aucune valeur n'est inventée : ils sont exclus de ce calcul.",
                             "Complétez le fichier si ces résultats existent.", empty_rows))
        elif empty_rows and it.kind == "skill":
            iss.append(Issue("info", f"« {lab} » non évalué", "Cette colonne est vide : elle sera marquée « N/E » (non évalué).", ""))
        if it.kind == "grade" and len(vals) >= 8:
            m, sd = mean(vals), pstdev(vals)
            if sd > 0:
                far = [s.row for s in ds.students if s.scores.get(it.key) is not None and abs(s.scores[it.key] - m) > 3 * sd]
                if far:
                    iss.append(Issue("info", f"Valeur(s) très éloignée(s) dans « {lab} »",
                                     f"Lignes {_rows_txt(far)} : valeur très différente du reste de la classe.",
                                     "Simple vérification : est-ce une faute de frappe ?", far))
    # cohérence des moyennes
    if ds.kind == "grade" and ds.has_average:
        g = [it for it in ds.items if it.kind == "grade"]
        diff = []
        for s in ds.students:
            vs = [s.scores[i.key] for i in g if s.scores.get(i.key) is not None]
            if vs and s.average_given is not None and abs(mean(vs) - s.average_given) > 0.5 * (ds.scale or 20) / 20:
                diff.append(s.row)
        if diff and len(diff) >= 0.2 * len(ds.students):
            iss.append(Issue("warning", "Moyennes différentes de la moyenne simple des notes",
                             f"{len(diff)} élève(s) ont une moyenne qui diffère de la moyenne simple des notes affichées (lignes {_rows_txt(diff)}).",
                             "C'est normal si des coefficients ou d'autres évaluations entrent dans le calcul. La moyenne du fichier est utilisée telle quelle."
                             , diff))
    return iss


def has_blocking(issues):
    return any(i.level == "error" for i in issues)
