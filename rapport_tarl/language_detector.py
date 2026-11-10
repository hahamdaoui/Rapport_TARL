# -*- coding: utf-8 -*-
"""Détection hors-ligne de la langue d'un ensemble de textes (en-têtes ou données).
Méthode : analyse des alphabets + mots-clés (français / anglais). `langdetect` n'est pas nécessaire."""
import re
from .text_utils import has_arabic, normalize

_FR = {"le", "la", "les", "des", "de", "du", "nom", "prenom", "eleve", "eleves", "note", "notes", "moyenne", "classe",
       "devoir", "controle", "absence", "observation", "coefficient", "comportement", "matiere", "semestre", "examen",
       "resultat", "resultats", "numero", "date", "naissance", "participation", "etablissement", "et", "sur"}
_EN = {"the", "of", "name", "student", "students", "grade", "grades", "mark", "marks", "score", "average", "class",
       "test", "quiz", "exam", "absence", "absent", "remarks", "comment", "behavior", "behaviour", "attendance",
       "first", "last", "term", "total", "mean", "id", "number", "and"}
LANG_NAMES = {"fr": "Français", "ar": "Arabe", "en": "Anglais", "mixed": "Mixte (plusieurs langues)", "unknown": "Indéterminée"}


def detect_language(texts):
    """Retourne (code, détail) avec code ∈ {'ar','fr','en','mixed','unknown'}."""
    ar = lat = 0
    fr_hits = en_hits = 0
    accents = 0
    for t in texts:
        t = str(t or "")
        if not t.strip():
            continue
        a = len(re.findall(r"[\u0600-\u06FF]", t))
        l = len(re.findall(r"[A-Za-zÀ-ÿ]", t))
        ar += a
        lat += l
        if l:
            accents += len(re.findall(r"[àâçéèêëîïôùûüœ]", t.lower()))
            words = set(normalize(t).split())
            fr_hits += len(words & _FR)
            en_hits += len(words & _EN)
    total = ar + lat
    if total == 0:
        return "unknown", {"arabic": 0, "latin": 0}
    share_ar = ar / total
    detail = {"arabic": ar, "latin": lat, "fr_hits": fr_hits, "en_hits": en_hits}
    if share_ar >= 0.85:
        return "ar", detail
    if share_ar > 0.15:
        return "mixed", detail
    if fr_hits > en_hits or (fr_hits == en_hits and accents > 0):
        return "fr", detail
    if en_hits > fr_hits:
        return "en", detail
    return "fr", detail   # latin sans indice : français par défaut (contexte marocain)
