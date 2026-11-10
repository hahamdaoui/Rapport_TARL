# -*- coding: utf-8 -*-
"""Lecture de fichiers Excel (.xlsx, .xlsm, .xls) -> modèle de feuilles indépendant du format.
Les cellules fusionnées sont propagées (la valeur de la cellule d'origine est recopiée) pour que les
en-têtes sur plusieurs lignes restent exploitables ; la liste des fusions est conservée."""
import os
from dataclasses import dataclass, field
from typing import Any, List, Set, Tuple

from . import ErreurRapport


@dataclass
class SheetModel:
    name: str
    grid: List[List[Any]]                 # valeurs (None = vide), rectangle nrows x ncols
    merged: List[Tuple[int, int, int, int]] = field(default_factory=list)   # (r1, c1, r2, c2) base 0, inclusif
    filled_by_merge: Set[Tuple[int, int]] = field(default_factory=set)       # cellules remplies par propagation
    declared_dims: str = ""

    @property
    def nrows(self):
        return len(self.grid)

    @property
    def ncols(self):
        return len(self.grid[0]) if self.grid else 0

    def cell(self, r, c):
        if 0 <= r < self.nrows and 0 <= c < self.ncols:
            return self.grid[r][c]
        return None

    def empty_rows(self):
        return [r for r, row in enumerate(self.grid) if all(_empty(v) for v in row)]

    def empty_cols(self):
        return [c for c in range(self.ncols) if all(_empty(self.grid[r][c]) for r in range(self.nrows))]


def _empty(v):
    return v is None or (isinstance(v, str) and not v.strip())


@dataclass
class WorkbookModel:
    path: str
    sheets: List[SheetModel]

    @property
    def sheet_names(self):
        return [s.name for s in self.sheets]


def _trim(grid):
    """Supprime les lignes / colonnes vides en bas et à droite (Excel déclare souvent des centaines de lignes vides)."""
    last_r = -1
    last_c = -1
    for r, row in enumerate(grid):
        for c, v in enumerate(row):
            if not _empty(v):
                last_r = max(last_r, r)
                last_c = max(last_c, c)
    if last_r < 0:
        return []
    return [row[:last_c + 1] + [None] * (last_c + 1 - len(row[:last_c + 1])) for row in grid[:last_r + 1]]


def _read_xlsx(path):
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:  # noqa: BLE001
        raise ErreurRapport(f"Impossible d'ouvrir le fichier Excel :\n{path}\n\nVérifiez qu'il n'est pas protégé par mot de passe ni ouvert dans un autre programme.\n({e})")
    sheets = []
    for ws in wb.worksheets:
        grid = [list(row) for row in ws.iter_rows(values_only=True)]
        merged = []
        for mr in ws.merged_cells.ranges:
            merged.append((mr.min_row - 1, mr.min_col - 1, mr.max_row - 1, mr.max_col - 1))
        sheets.append(_finish(ws.title, grid, merged, ws.dimensions))
    return sheets


def _read_xls(path):
    try:
        import xlrd  # type: ignore
    except ImportError:
        raise ErreurRapport("Les anciens fichiers .xls nécessitent la bibliothèque « xlrd ».\n"
                            "Solution simple : ouvrez le fichier dans Excel, faites « Enregistrer sous » au format .xlsx, "
                            "puis importez le nouveau fichier.")
    try:
        book = xlrd.open_workbook(path, formatting_info=False)
    except Exception as e:  # noqa: BLE001
        raise ErreurRapport(f"Impossible d'ouvrir le fichier .xls :\n{path}\n({e})")
    sheets = []
    for sh in book.sheets():
        grid = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                v = sh.cell_value(r, c)
                row.append(None if v == "" else v)
            grid.append(row)
        merged = [(rlo, clo, rhi - 1, chi - 1) for (rlo, rhi, clo, chi) in getattr(sh, "merged_cells", [])]
        sheets.append(_finish(sh.name, grid, merged, f"{sh.nrows}x{sh.ncols}"))
    return sheets


def _finish(name, grid, merged, dims):
    grid = _trim([list(r) for r in grid])
    filled = set()
    nr = len(grid)
    nc = len(grid[0]) if grid else 0
    for (r1, c1, r2, c2) in merged:
        if r1 >= nr or c1 >= nc:
            continue
        v = grid[r1][c1]
        if _empty(v):
            continue
        for r in range(r1, min(r2, nr - 1) + 1):
            for c in range(c1, min(c2, nc - 1) + 1):
                if (r, c) != (r1, c1) and _empty(grid[r][c]):
                    grid[r][c] = v
                    filled.add((r, c))
    return SheetModel(name=name, grid=grid, merged=merged, filled_by_merge=filled, declared_dims=dims)


def read_workbook(path):
    """Lit un classeur Excel complet. Lève ErreurRapport avec un message compréhensible."""
    if not path or not os.path.exists(path):
        raise ErreurRapport(f"Fichier introuvable :\n{path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        sheets = _read_xlsx(path)
    elif ext == ".xls":
        sheets = _read_xls(path)
    else:
        raise ErreurRapport(f"Format non pris en charge : « {ext} ».\nUtilisez un fichier Excel (.xlsx ou .xls).")
    if not sheets:
        raise ErreurRapport("Le fichier Excel ne contient aucune feuille.")
    return WorkbookModel(path=path, sheets=sheets)
