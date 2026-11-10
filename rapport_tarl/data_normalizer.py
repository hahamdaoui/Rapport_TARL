# -*- coding: utf-8 -*-
"""Conversion des données Excel en structure interne standardisée, indépendante de la langue du fichier,
des noms de colonnes, de leur ordre, de leur nombre et de la matière.
Les données originales ne sont JAMAIS modifiées : les valeurs illisibles deviennent None (et sont signalées
par data_validator), jamais remplacées par une valeur inventée."""
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from .structure_analyzer import SheetAnalysis
from .text_utils import clean_str, normalize, numeric_part_with_scale, to_number, yes_no


@dataclass
class Item:
    key: str
    col_idx: int
    label: str
    short: str
    group: str
    kind: str                       # 'skill' | 'grade'
    scale: Optional[float] = None
    coefficient: Optional[float] = None
    code_auto: bool = False
    title: str = ""


@dataclass
class Student:
    row: int                        # numéro de ligne Excel (base 1) pour les messages
    name: str
    sid: str = ""
    absent: bool = False
    absence_count: Optional[float] = None
    scores: Dict[str, Optional[float]] = field(default_factory=dict)
    raw: Dict[str, Any] = field(default_factory=dict)
    average_given: Optional[float] = None
    participation: Any = None
    behavior: Any = None
    observation: str = ""

    @property
    def key(self):
        return normalize(self.sid) if self.sid else normalize(self.name)


@dataclass
class Dataset:
    path: str
    sheet: str
    meta: Dict[str, str]
    kind: str                       # 'skill' | 'grade' | 'unknown'
    test_type: str
    items: List[Item]
    students: List[Student]
    scale: Optional[float] = None
    scale_source: str = ""          # 'en-tête' | 'valeurs' | 'choix utilisateur'
    has_id: bool = False
    has_average: bool = False
    has_absence: bool = False
    has_participation: bool = False
    has_behavior: bool = False
    has_observation: bool = False
    coefficient: Optional[float] = None
    file_language: str = "unknown"
    mapping: Dict[int, str] = field(default_factory=dict)

    def to_standard(self):
        """Structure interne standardisée (voir cahier des charges) : un dictionnaire par élève."""
        out = []
        for s in self.students:
            d = {"student_name": s.name, "student_id": s.sid or None, "grades": {}, "general_average": s.average_given,
                 "attendance": ("absent" if s.absent else "présent") if self.has_absence else None,
                 "participation": s.participation, "behavior": s.behavior}
            for i, it in enumerate(self.items, 1):
                d["grades"][f"assessment_{i}" if it.kind == "grade" else f"skill_{i}"] = s.scores.get(it.key)
            out.append(d)
        return out


_SCALE_RE = [re.compile(r"/\s*(\d{1,3})\b"), re.compile(r"\bsur\s*(\d{1,3})\b", re.I), re.compile(r"\(\s*(\d{1,3})\s*\)"),
             re.compile(r"\bout of\s*(\d{1,3})\b", re.I), re.compile(r"على\s*(\d{1,3})")]


def scale_from_header(text):
    t = (text or "").translate({ord(c): str(i) for i, c in enumerate("٠١٢٣٤٥٦٧٨٩")})
    for rx in _SCALE_RE:
        m = rx.search(t)
        if m:
            v = float(m.group(1))
            if v in (5, 10, 20, 40, 60, 100):
                return v
    return None


def _auto_scale(values):
    vals = [v for v in values if v is not None]
    if not vals:
        return 20.0
    mx = max(vals)
    for s in (20, 40, 60, 100):
        if mx <= s:
            return float(s)
    return float(max(vals))


def default_mapping(an: SheetAnalysis):
    return {c.idx: c.role for c in an.columns}


def build_dataset(an: SheetAnalysis, mapping=None, scale=None, path="") -> Dataset:
    """Construit le Dataset standardisé. `mapping` = {index de colonne: rôle} (validation manuelle de l'enseignant)."""
    mp = dict(default_mapping(an))
    if mapping:
        mp.update(mapping)
    sh = an.sheet
    cols = {c.idx: c for c in an.columns}
    roles = {}
    for idx, r in mp.items():
        roles.setdefault(r, []).append(idx)
    name_col = (roles.get("name") or [None])[0]
    id_col = (roles.get("id") or [None])[0]
    abs_col = (roles.get("absence") or [None])[0]
    avg_col = (roles.get("average") or [None])[0]
    par_col = (roles.get("participation") or [None])[0]
    beh_col = (roles.get("behavior") or [None])[0]
    obs_col = (roles.get("observation") or [None])[0]
    coef_col = (roles.get("coefficient") or [None])[0]

    items: List[Item] = []
    for role in ("skill", "grade"):
        for idx in sorted(roles.get(role, [])):
            c = cols[idx]
            n = len(items) + 1
            short = c.item_code or (f"S{len([i for i in items if i.kind == 'skill']) + 1}" if role == "skill"
                                    else f"E{len([i for i in items if i.kind == 'grade']) + 1}")
            label = c.display
            items.append(Item(key=f"i{idx}", col_idx=idx, label=label, short=short, group=c.group, kind=role,
                              scale=scale_from_header(c.header_full) if role == "grade" else None,
                              code_auto=not c.item_code,
                              title=re.sub(r"\s*\([A-Za-z0-9_\-]{2,8}\)\s*$", "", c.header).strip() or c.display))
    items.sort(key=lambda i: i.col_idx)

    students: List[Student] = []
    for r in an.student_rows:
        nm = clean_str(sh.cell(r, name_col)) if name_col is not None else f"Élève {len(students) + 1}"
        if not nm:
            continue
        st = Student(row=r + 1, name=nm)
        if id_col is not None:
            st.sid = clean_str(sh.cell(r, id_col))
        if abs_col is not None:
            v = sh.cell(r, abs_col)
            yn = yes_no(v)
            if yn is not None:
                st.absent = (not yn) if cols[abs_col].invert else yn
            else:
                n_ = to_number(v)
                if n_ is not None:
                    st.absence_count = n_
        for it in items:
            raw = sh.cell(r, it.col_idx)
            st.raw[it.key] = raw
            if it.kind == "skill":
                v = to_number(raw)
                st.scores[it.key] = v if v in (0.0, 1.0) else None
            else:
                v, sc = numeric_part_with_scale(raw)
                if sc and not it.scale:
                    it.scale = sc
                st.scores[it.key] = v
        if avg_col is not None:
            st.average_given = to_number(sh.cell(r, avg_col))
        if par_col is not None:
            st.participation = sh.cell(r, par_col)
        if beh_col is not None:
            st.behavior = sh.cell(r, beh_col)
        if obs_col is not None:
            st.observation = clean_str(sh.cell(r, obs_col))
        students.append(st)

    n_skill = sum(1 for i in items if i.kind == "skill")
    n_grade = len(items) - n_skill
    kind = "skill" if n_skill > n_grade else ("grade" if n_grade else "unknown")
    ds = Dataset(path=path, sheet=sh.name, meta=dict(an.meta), kind=kind, test_type=an.test_type, items=items,
                 students=students, has_id=id_col is not None, has_average=avg_col is not None,
                 has_absence=abs_col is not None, has_participation=par_col is not None,
                 has_behavior=beh_col is not None, has_observation=obs_col is not None, file_language=an.language,
                 mapping=mp)
    if kind == "grade":
        allv = [v for s in students for it in items for v in [s.scores.get(it.key)] if it.kind == "grade"]
        if scale:
            ds.scale, ds.scale_source = float(scale), "choix de l'enseignant"
            for it in items:
                it.scale = it.scale or float(scale)
        else:
            known = {it.scale for it in items if it.scale}
            if known and len(known) == 1:
                ds.scale, ds.scale_source = known.pop(), "en-tête"
            else:
                ds.scale, ds.scale_source = _auto_scale(allv), "valeurs"
            for it in items:
                if it.kind == "grade":
                    it.scale = it.scale or ds.scale
    if coef_col is not None:
        cs = {to_number(sh.cell(r, coef_col)) for r in an.student_rows} - {None}
        if len(cs) == 1:
            ds.coefficient = cs.pop()
    return ds
