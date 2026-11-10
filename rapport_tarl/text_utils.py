# -*- coding: utf-8 -*-
"""Normalisation de textes (français / arabe / anglais) et conversion de nombres."""
import re
import unicodedata

_AR_DIGITS = {ord(c): str(i) for i, c in enumerate("٠١٢٣٤٥٦٧٨٩")}
_AR_DIGITS.update({ord(c): str(i) for i, c in enumerate("۰۱۲۳۴۵۶۷۸۹")})
_AR_MAP = {ord(c): "ا" for c in "إأآٱ"}
_AR_MAP.update({ord("ى"): "ي", ord("ة"): "ه", ord("ؤ"): "و", ord("ئ"): "ي", ord("ک"): "ك", ord("ی"): "ي"})
_AR_DIAC = re.compile("[\u064B-\u065F\u0670\u0640\u06D6-\u06ED]")
_ARABIC = re.compile(r"[\u0600-\u06FF\u0750-\u077F]")
_LATIN = re.compile(r"[A-Za-zÀ-ÿ]")


def has_arabic(s):
    return bool(_ARABIC.search(str(s or "")))


def normalize(s):
    """Minuscule, sans accents ni diacritiques, alef/ya/ta marbuta unifiés, chiffres arabes -> 0-9,
    caractères spéciaux et espaces multiples supprimés."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.translate(_AR_DIGITS)
    s = _AR_DIAC.sub("", s)
    s = s.translate(_AR_MAP)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.lower()
    s = re.sub(r"[^\w\u0600-\u06FF]+|_", " ", s)
    return re.sub(r"\s+", " ", s).strip()


MISSING = {"", "-", "--", "—", "–", "/", "n/a", "na", "nan", "none", "null", "abs", "absent", "ab", "غ", "غائب", "nd", "n.d", "ne"}


def clean_str(v):
    if v is None:
        return ""
    return str(v).strip()


def to_number(v):
    """Convertit une cellule en nombre (float) ou None.
    Gère : 12, 12.5, '12,5', '١٢', '12/20', ' 14 ', '14.00 pts'. Ne transforme jamais un texte libre en nombre."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if v != v else float(v)
    s = unicodedata.normalize("NFKC", str(v)).translate(_AR_DIGITS).strip()
    if s.lower() in MISSING:
        return None
    s = s.replace("٫", ".").replace("،", ".").replace(",", ".")
    m = re.fullmatch(r"([-+]?\d+(?:\.\d+)?)\s*(?:/\s*\d+(?:\.\d+)?)?\s*(?:pts?|points?|ن|نقطة)?", s, flags=re.I)
    return float(m.group(1)) if m else None


def numeric_part_with_scale(v):
    """'12/20' -> (12.0, 20.0) ; sinon (nombre, None)."""
    if v is None or isinstance(v, bool):
        return None, None
    s = unicodedata.normalize("NFKC", str(v)).translate(_AR_DIGITS).strip().replace(",", ".")
    m = re.fullmatch(r"([-+]?\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", s)
    if m:
        return float(m.group(1)), float(m.group(2))
    return to_number(v), None


YES = {"نعم", "oui", "yes", "y", "o", "ن", "true", "vrai", "1", "present", "présent", "حاضر"}
NO = {"لا", "non", "no", "n", "false", "faux", "0", "ل"}


def yes_no(v):
    """True / False / None pour une cellule de type oui/non."""
    s = normalize(v)
    if s in {normalize(x) for x in YES}:
        return True
    if s in {normalize(x) for x in NO}:
        return False
    return None
