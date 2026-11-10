# -*- coding: utf-8 -*-
"""Reconnaissance multilingue (FR / AR / EN) de la fonction des colonnes.

Approche combinée :
  1. dictionnaire de synonymes (extensible : voir SYNONYMS ou le fichier synonymes_utilisateur.json),
  2. normalisation des textes (accents, variantes arabes, chiffres, casse, caractères spéciaux),
  3. correspondance approximative (rapidfuzz si installé, sinon difflib),
  4. analyse du TYPE et du CONTENU des données de la colonne,
  5. règles de cohérence (une seule colonne « nom », valeurs compatibles avec le rôle, blocs de colonnes voisines).
On ne se base jamais uniquement sur le nom de la colonne."""
import json
import os
import re
from difflib import SequenceMatcher

from .text_utils import normalize, to_number, yes_no

try:  # facultatif
    from rapidfuzz import fuzz as _fuzz  # type: ignore
except Exception:  # noqa: BLE001
    _fuzz = None

ROLES = ["name", "id", "grade", "skill", "average", "coefficient", "absence", "participation", "behavior",
         "observation", "ignore"]
ROLE_LABELS_FR = {
    "name": "Nom de l'élève", "id": "Identifiant", "grade": "Note d'évaluation", "skill": "Compétence validée (0/1)",
    "average": "Moyenne", "coefficient": "Coefficient", "absence": "Absence", "participation": "Participation",
    "behavior": "Comportement", "observation": "Observation", "ignore": "Autre / Ignorer",
}

# ----------------------------------------------------------------------------- dictionnaire
SYNONYMS = {
    "name": [
        # arabe
        "الاسم الكامل", "الاسم والنسب", "الاسم الشخصي والعائلي", "الاسم العائلي والشخصي", "الاسم", "اسم التلميذ", "اسم الطالب",
        "اسم التلميذة", "التلميذ", "التلميذة", "الطالب", "الطالبة", "المتعلم", "المتعلمة", "النسب والاسم", "اسم ولقب", "الاسم واللقب",
        # français
        "nom et prenom", "nom prenom", "nom complet", "nom de l eleve", "nom eleve", "nom et prenoms", "eleve", "eleves", "nom", "prenom",
        "apprenant", "etudiant",
        # anglais
        "student name", "full name", "name", "student", "pupil", "learner", "first name last name"],
    "id": [
        "رقم التلميذ", "رقم مسار", "رقم مسار التلميذ", "مسار", "الرقم الوطني", "الرقم الترتيبي", "رمز التلميذ", "رقم التسجيل", "رقم الترتيب",
        "رقم", "كود", "الرمز", "ر ت",
        "identifiant", "identifiant eleve", "id eleve", "matricule", "massar", "code massar", "code eleve", "code", "numero", "numero d ordre",
        "n d ordre", "n ordre", "n°", "no", "num", "numero eleve",
        "student id", "id", "identifier", "number", "student number", "registration number"],
    "grade": [
        "الفرض الاول", "الفرض الثاني", "الفرض الثالث", "الفرض الرابع", "فرض محروس", "الفرض المحروس", "فرض كتابي", "الفرض", "فرض",
        "المراقبة المستمرة", "الاختبار", "اختبار", "التقويم", "تقويم", "النقطة", "نقطة", "النقط", "نقط", "الامتحان", "امتحان", "نقطة الفرض",
        "الانجاز", "الاداء",
        "devoir", "devoir 1", "devoir 2", "devoir surveille", "controle", "controle continu", "note", "notes", "note 1", "note 2",
        "evaluation", "examen", "interrogation", "composition", "ds", "dm", "test", "quiz", "epreuve",
        "grade", "grades", "mark", "marks", "score", "exam", "assignment", "homework", "test 1", "test 2", "first test", "second test",
        "continuous assessment", "assessment"],
    "average": [
        "المعدل العام", "المعدل", "معدل", "المعدل السنوي", "معدل الدورة", "المعدل الاجمالي", "المعدل الفصلي", "المتوسط", "المتوسط الحسابي", "المجموع العام",
        "moyenne generale", "moyenne", "moy", "moy gen", "moyenne annuelle", "moyenne du semestre", "moyenne semestre",
        "average", "mean", "overall", "overall average", "general average", "gpa", "final grade", "avg"],
    "coefficient": ["المعامل", "معامل", "coefficient", "coef", "coeff", "weight", "weighting", "factor", "poids"],
    "absence": [
        "الغياب", "غياب", "الغيابات", "عدد الغيابات", "عدد الغياب", "غائب", "التغيب", "absence", "absences", "abs", "nb abs", "nombre d absences",
        "absent", "attendance", "absent days", "days absent"],
    "presence_inverse": ["الحضور", "حضور", "presence", "present"],
    "participation": ["المشاركة", "مشاركة", "الانخراط", "التفاعل", "participation", "participation en classe", "implication", "engagement",
                      "class participation", "involvement"],
    "behavior": ["السلوك", "سلوك", "الانضباط", "الاخلاق", "حسن السلوك", "comportement", "conduite", "discipline", "attitude", "behavior",
                 "behaviour", "conduct"],
    "observation": ["الملاحظات", "ملاحظات", "الملاحظة", "ملاحظة", "التقدير", "تقدير", "ملاحظات الاستاذ", "observation", "observations", "remarque", "remarques",
                    "appreciation", "commentaire", "commentaires", "avis", "comment", "comments", "remark", "remarks", "teacher comment"],
    "other_known": ["الجنس", "النوع", "sexe", "genre", "gender", "sex", "تاريخ الازدياد", "date de naissance", "date naissance", "birth date",
                    "date of birth", "القسم", "classe", "class", "المستوى", "niveau", "level", "الاستاذ", "professeur", "teacher", "المادة", "matiere", "subject",
                    "المعلومات", "معلومات التلميذ", "informations", "info", "الرتبة", "rang", "rank", "الترتيب", "classement"],
    # indices (ne décident jamais seuls) pour les colonnes de compétences validées / non validées
    "skill_hint": ["المرحلة", "المهارة", "المهارات", "الكفاية", "الكفايات", "الطلاقة", "الفهم", "competence", "competences", "domaine", "etape",
                   "skill", "skills", "domain", "stage", "fluency", "comprehension", "validation", "valide", "acquis"],
}

import sys as _sys
_BASE = os.path.dirname(_sys.executable) if getattr(_sys, "frozen", False) else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_EXTRA = os.path.join(_BASE, "synonymes_utilisateur.json")


def charger_synonymes_utilisateur(path=None):
    """Ajoute les synonymes personnalisés : fichier JSON {"rôle": ["intitulé 1", ...]} (voir LISEZMOI)."""
    p = path or _EXTRA
    if not os.path.exists(p):
        return 0
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:  # noqa: BLE001
        return 0
    n = 0
    for role, words in data.items():
        if role in SYNONYMS and isinstance(words, list):
            for w in words:
                SYNONYMS[role].append(str(w)); n += 1
    _build_index()
    return n


_IDX = {}


def _build_index():
    _IDX.clear()
    for role, words in SYNONYMS.items():
        _IDX[role] = sorted({normalize(w) for w in words if normalize(w)}, key=lambda s: -len(s))


_build_index()


def ajouter_synonyme(role, mot):
    SYNONYMS.setdefault(role, []).append(mot)
    _build_index()


# ----------------------------------------------------------------------------- scores d'intitulés
def _tokens_contain(h_tokens, s_tokens):
    n = len(s_tokens)
    if n == 0 or n > len(h_tokens):
        return False
    return any(h_tokens[i:i + n] == s_tokens for i in range(len(h_tokens) - n + 1))


def score_header(text):
    """Retourne {rôle: score 0-100} pour un intitulé."""
    h = normalize(text)
    out = {}
    if not h:
        return out
    ht = h.split()
    for role, syns in _IDX.items():
        best = 0
        for s in syns:
            if h == s:
                sc = 100
            else:
                st = s.split()
                if _tokens_contain(ht, st):
                    if len(st) == 1 and len(ht) > 2:
                        sc = 62          # mot isolé dans un long intitulé (ex. « النص المتوسط ») : indice faible
                    else:
                        sc = 78 + min(18, 6 * len(st)) if len(s) > 2 else 70
                else:
                    sc = 0
                    if len(s) >= 4 and len(h) >= 4:
                        r = (_fuzz.token_set_ratio(h, s) if _fuzz else SequenceMatcher(None, h, s).ratio() * 100)
                        if r >= 86:
                            sc = r * 0.92
            best = max(best, sc)
        if best:
            out[role] = best
    return out


def best_header_role(*texts):
    """Meilleur rôle selon un ou plusieurs textes d'en-tête (ligne basse puis lignes hautes)."""
    best, bs = None, 0
    for i, t in enumerate(texts):
        for role, sc in score_header(t).items():
            if role in ("skill_hint",):
                continue
            sc2 = sc - (4 * i)          # l'intitulé le plus proche des données prime
            if sc2 > bs:
                best, bs = role, sc2
    if best == "presence_inverse":
        return "absence", bs, True
    if best == "other_known":
        return "ignore", bs, False
    return best, bs, False


def has_skill_hint(*texts):
    return any(score_header(t).get("skill_hint", 0) >= 70 for t in texts)


# ----------------------------------------------------------------------------- profil de contenu
CODE_RE = re.compile(r"^[A-Za-z]{0,3}\d{5,}[A-Za-z0-9]*$")


def profile(values):
    vals = [v for v in values if v is not None and str(v).strip() != ""]
    n = len(values)
    filled = len(vals)
    p = {"n": n, "filled": filled, "fill_rate": (filled / n) if n else 0}
    if not filled:
        p.update(numeric_frac=0, binary_frac=0, yn_frac=0, unique_ratio=0, avg_words=0, code_frac=0, text_frac=0,
                 maxv=None, minv=None, distinct=0)
        return p
    nums = [to_number(v) for v in vals]
    ok = [x for x in nums if x is not None]
    p["numeric_frac"] = len(ok) / filled
    p["binary_frac"] = sum(1 for x in ok if x in (0.0, 1.0)) / filled
    textual = [v for v in vals if isinstance(v, str) and to_number(v) is None]
    p["yn_frac"] = sum(1 for v in textual if yes_no(v) is not None) / filled
    norm = [normalize(v) for v in vals]
    p["unique_ratio"] = len(set(norm)) / filled
    p["avg_words"] = sum(len(x.split()) for x in norm) / filled
    p["code_frac"] = sum(1 for v in vals if CODE_RE.match(str(v).strip())) / filled
    p["text_frac"] = len(textual) / filled
    p["maxv"] = max(ok) if ok else None
    p["minv"] = min(ok) if ok else None
    p["distinct"] = len(set(norm))
    return p


def _compatible(role, p):
    if p["filled"] == 0:
        return role in ("observation", "ignore", "skill", "grade")
    if role == "name":
        return p["text_frac"] >= .8 and p["avg_words"] >= 1.2
    if role == "id":
        return p["unique_ratio"] >= .85
    if role in ("grade", "average"):
        return p["numeric_frac"] >= .7 and p["binary_frac"] < .95      # que des 0/1 : ce sont des compétences
    if role == "coefficient":
        return p["numeric_frac"] >= .7
    if role == "skill":
        return p["numeric_frac"] >= .9 and p["binary_frac"] >= .9
    if role == "absence":
        return p["yn_frac"] >= .7 or p["numeric_frac"] >= .7
    if role in ("participation", "behavior"):
        return True
    return True


# ----------------------------------------------------------------------------- détection principale
def detect_columns(analysis, skip_idx=()):
    """Remplit pour chaque ColumnInfo : role, confidence (0-100), status, reason. Ne modifie jamais les données."""
    cols = analysis.columns
    for c in cols:
        c.profile = profile(c.values)
        c.role, c.confidence, c.status, c.reason, c.invert = "ignore", 0, "ignored", "", False
        p = c.profile
        hrole, hscore, inv = best_header_role(c.header, c.group)
        c.invert = inv
        hint = has_skill_hint(c.header, c.group)
        role, conf, reason = None, 0, ""

        if p["filled"] == 0 and not hint and not c.header:
            c.reason = "colonne vide"
            continue
        if hrole and hscore >= 78:
            if _compatible(hrole, p):
                role, conf, reason = hrole, hscore, "intitulé reconnu"
            else:
                # intitulé reconnu mais contenu inattendu : on cherche d'après le contenu
                role, conf, reason = None, 0, "intitulé reconnu mais contenu inattendu"
        if role is None:
            # décision d'après le contenu
            if p["filled"] and p["numeric_frac"] >= .9 and p["binary_frac"] >= .9:
                role, conf, reason = "skill", 72 + (10 if hint else 0), "valeurs 0/1"
            elif p["filled"] and p["yn_frac"] >= .8 and p["text_frac"] >= .8:
                role, conf, reason = "absence", 70, "valeurs oui/non"
            elif p["filled"] and p["numeric_frac"] >= .8:
                if p["minv"] is not None and p["minv"] >= 0 and p["maxv"] <= 100:
                    role, conf, reason = "grade", 66 + (6 if (hrole == "grade") else 0), "valeurs numériques"
                else:
                    role, conf, reason = "ignore", 40, "valeurs numériques hors barème usuel"
            elif p["filled"] and p["text_frac"] >= .8 and p["unique_ratio"] >= .9 and p["avg_words"] >= 1.5:
                role, conf, reason = "name", 70, "textes uniques de plusieurs mots"
            elif p["filled"] and p["code_frac"] >= .6 and p["unique_ratio"] >= .9:
                role, conf, reason = "id", 70, "codes uniques"
            elif p["filled"] and p["text_frac"] >= .6:
                role, conf, reason = "ignore", 35, "texte non reconnu"
            else:
                role, conf, reason = "ignore", 30, "contenu non reconnu"
            if hrole and role != "ignore":
                reason += f" ; intitulé évoquant « {ROLE_LABELS_FR.get(hrole, hrole)} »"
        c.role, c.confidence, c.reason = role, min(100, conf), reason
        if c.idx in skip_idx:
            c.role, c.confidence, c.status, c.reason = "ignore", 100, "ignored", "ignorée"

    _coherence(cols)
    for c in cols:
        if c.status == "ignored" and c.reason == "ignorée":
            continue
        if c.role == "ignore":
            c.status = "ignored" if (c.profile["filled"] == 0 or not c.header or c.confidence <= 40 and not c.header) else "verify"
            if c.profile["filled"] and c.header and c.reason in ("texte non reconnu", "contenu non reconnu"):
                c.status = "verify"
        else:
            c.status = "recognized" if c.confidence >= 80 else "verify"
    return cols


def _coherence(cols):
    """Règles de cohérence : une seule colonne « nom » et « identifiant », blocs de colonnes voisines."""
    for role in ("name", "id"):
        cand = [c for c in cols if c.role == role]
        if len(cand) > 1:
            cand.sort(key=lambda c: -c.confidence)
            for c in cand[1:]:
                c.reason = f"doublon de « {ROLE_LABELS_FR[role]} » (ignoré)"
                c.role = "ignore" if (c.confidence < 90 or not c.header) else "ignore"
                c.confidence = 45
    # blocs : une colonne peu remplie voisine de colonnes « compétence » / « note » hérite du rôle du bloc
    for role in ("skill", "grade"):
        solid = {c.idx for c in cols if c.role == role and c.confidence >= 70}
        for c in cols:
            if c.role in ("ignore",) and c.idx not in solid and c.profile["filled"] <= max(2, c.profile["n"] * 0.1):
                if (c.idx - 1 in solid or c.idx + 1 in solid) and c.header:
                    ok_content = c.profile["filled"] == 0 or (role == "skill" and c.profile["binary_frac"] >= .9) or \
                                 (role == "grade" and c.profile["numeric_frac"] >= .9)
                    if ok_content:
                        c.role, c.confidence, c.reason = role, 78, "colonne voisine d'un bloc de même nature (peu remplie)"
    for role in ("skill", "grade"):
        run = []
        for c in cols + [None]:
            if c is not None and c.role == role:
                run.append(c)
            else:
                if len(run) >= 3:
                    for r in run:
                        if r.confidence < 84 and (role == "skill" or "intitulé reconnu" in r.reason):
                            r.confidence = 84
                            r.reason += " ; fait partie d'un bloc de colonnes similaires"
                run = []
    # un identifiant sans en-tête est « interne » : à ignorer si une autre colonne id existe
    ids = [c for c in cols if c.role == "id"]
    if len(ids) == 1 and not ids[0].header and any(c.role == "name" for c in cols):
        pass


def summary(cols):
    """Compte des colonnes par statut."""
    return {s: sum(1 for c in cols if c.status == s) for s in ("recognized", "verify", "ignored")}


charger_synonymes_utilisateur()      # synonymes personnalisés (facultatif), à côté du programme
