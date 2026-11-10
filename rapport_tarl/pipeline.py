# -*- coding: utf-8 -*-
"""Chaîne complète : Excel -> analyse -> colonnes (validation manuelle possible) -> contrôles -> traitement -> DOCX.
C'est le point d'entrée utilisé par l'interface graphique et par la ligne de commande."""
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import ErreurRapport
from .data_normalizer import Dataset, build_dataset
from .data_processor import analyze_grades, analyze_skill_pair
from .data_validator import Issue, has_blocking, validate
from .excel_reader import WorkbookModel, read_workbook
from .report_generator import generate
from .structure_analyzer import SheetAnalysis, analyze_workbook, best_analysis


@dataclass
class FileState:
    """Tout ce qu'il faut savoir sur un fichier importé (modifiable via l'interface de vérification)."""
    path: str
    workbook: WorkbookModel
    analyses: List[SheetAnalysis]
    current: SheetAnalysis
    overrides: Dict[int, str] = field(default_factory=dict)     # {index colonne: rôle choisi par l'enseignant}
    scale: Optional[float] = None
    dataset: Optional[Dataset] = None
    issues: List[Issue] = field(default_factory=list)

    @property
    def name(self):
        return os.path.basename(self.path)

    def role_of(self, col):
        return self.overrides.get(col.idx, col.role)

    def rebuild(self):
        """Recalcule le dataset standardisé et les contrôles après une correction manuelle."""
        mapping = {c.idx: self.role_of(c) for c in self.current.columns}
        self.dataset = build_dataset(self.current, mapping, self.scale, self.path)
        self.issues = validate(self.dataset, self.current)
        return self

    def use_sheet(self, sheet_name):
        for a in self.analyses:
            if a.sheet.name == sheet_name and a.ok:
                self.current = a
                self.overrides = {}
                return self.rebuild()
        raise ErreurRapport(f"La feuille « {sheet_name} » ne contient pas de liste d'élèves reconnue.")


def open_file(path) -> FileState:
    """Étapes 1 à 3 : importer, analyser la structure, reconnaître les colonnes."""
    wb = read_workbook(path)
    best, allr = best_analysis(wb)
    return FileState(path=path, workbook=wb, analyses=allr, current=best).rebuild()


def _kinds(states):
    return {s.dataset.kind for s in states}


def generate_reports(states: List[FileState], sortie, lang="fr", prof="", director="", log=print, progress=lambda x: None):
    """Étapes 5 à 7 : contrôles bloquants, traitement statistique, génération du DOCX. Retourne la liste des fichiers créés."""
    if not states:
        raise ErreurRapport("Aucun fichier à traiter.")
    for s in states:
        s.rebuild()
        if has_blocking(s.issues):
            e = next(i for i in s.issues if i.level == "error")
            raise ErreurRapport(f"{s.name} : {e.title}.\n{e.detail}\n{e.suggestion}")
    kinds = _kinds(states)
    progress(25)
    if len(kinds) > 1:
        raise ErreurRapport("Les fichiers importés ne sont pas de même nature (compétences 0/1 et notes chiffrées mélangées).\n"
                            "Importez ensemble soit un pré-test et un post-test, soit des fichiers de notes.")
    kind = kinds.pop()
    out = []
    if kind == "skill":
        if len(states) != 2:
            raise ErreurRapport("Pour un rapport pré-test / post-test, importez exactement DEUX fichiers :\nle pré-test et le post-test.")
        a, b = states[0].dataset, states[1].dataset
        if a.test_type == "post" or b.test_type == "pre":
            a, b = b, a
        if not (a.test_type == "pre" and b.test_type == "post"):
            raise ErreurRapport("Impossible de reconnaître le pré-test et le post-test.\n"
                                "Le fichier doit indiquer le type de test (ex. « رائز الموضعة القبلي » / « البعدي », « pré-test » / « post-test »).")
        log(f"Pré-test et post-test identifiés · {len(a.students)} élèves")
        progress(45)
        res = analyze_skill_pair(a, b, log)
        progress(65)
        generate(res, lang, sortie, prof, director, log)
        out.append(sortie)
    elif kind == "grade":
        for i, s in enumerate(states):
            res = analyze_grades(s.dataset, log)
            progress(40 + int(50 * (i + 1) / len(states)))
            target = sortie if i == 0 else os.path.splitext(sortie)[0] + f"_{i + 1}" + os.path.splitext(sortie)[1]
            generate(res, lang, target, prof, director, log)
            out.append(target)
    else:
        raise ErreurRapport("Aucune colonne de résultats n'a été retenue. Vérifiez les fonctions des colonnes.")
    progress(100)
    return out
