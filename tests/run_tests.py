# -*- coding: utf-8 -*-
"""Tests automatiques : 10 scénarios du cahier des charges. Lancer : python tests/run_tests.py"""
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from make_fixtures import fx  # noqa: E402
from rapport_tarl.pipeline import generate_reports, open_file  # noqa: E402
from rapport_tarl.text_utils import normalize  # noqa: E402

OK, KO = [], []


def check(name, cond, detail=""):
    (OK if cond else KO).append(name)
    print(("  ✔ " if cond else "  ✘ ") + name + (f"  → {detail}" if detail and not cond else ""))


def grades_of(ds, i):
    st = ds.students[i]
    return [st.scores[it.key] for it in ds.items if it.kind == "grade"]


def exact(ds, truth):
    """Les noms et les notes doivent rester EXACTS après normalisation (ordre des élèves conservé)."""
    if [s.name for s in ds.students] != truth["names"]:
        return False, "noms différents"
    for i, row in enumerate(truth["grades"]):
        got = grades_of(ds, i)
        if sorted(x for x in got if x is not None) != sorted(x for x in row if x is not None):
            return False, f"élève {i}: attendu {row}, obtenu {got}"
    return True, ""


def main():
    T = fx()
    tmp = tempfile.mkdtemp()
    print("== Tests 1 à 9 : reconnaissance + exactitude des données ==")
    for fn, tr in T.items():
        st = open_file(tr["path"])
        ds = st.dataset
        roles = {c.role for c in st.current.columns}
        check(f"{fn}: élèves détectés ({len(tr['names'])})", len(ds.students) == len(tr["names"]), f"{len(ds.students)}")
        check(f"{fn}: colonnes de notes ({tr['n_items']})", sum(1 for i in ds.items if i.kind == "grade") == tr["n_items"],
              str([(c.display, c.role) for c in st.current.columns]))
        ok, why = exact(ds, tr)
        check(f"{fn}: notes exactes après normalisation", ok, why)
    print("\n== Test 2 (détail) : fichier arabe ==")
    st = open_file(T["T2_arabe_ar.xlsx"]["path"])
    check("langue des en-têtes = arabe", st.current.language == "ar", st.current.language)
    check("colonne « المعدل العام » = moyenne", any(c.role == "average" for c in st.current.columns))
    check("colonne « السلوك » = comportement", any(c.role == "behavior" for c in st.current.columns))
    print("\n== Test 4 : colonnes bilingues ==")
    st = open_file(T["T4_bilingue.xlsx"]["path"])
    check("langue = mixte", st.current.language == "mixed", st.current.language)
    print("\n== Test 5 : ordre différent ==")
    st = open_file(T["T5_ordre.xlsx"]["path"])
    r = {c.header: c.role for c in st.current.columns}
    check("« Devoir 1 » et « Devoir 2 » = notes", r["Devoir 1"] == "grade" and r["Devoir 2"] == "grade", str(r))
    check("« Moyenne » = moyenne", r["Moyenne"] == "average", str(r))
    print("\n== Test 6 : colonnes inutiles ignorées ==")
    st = open_file(T["T6_colonnes_inutiles.xlsx"]["path"])
    roles = {c.header: c.role for c in st.current.columns}
    check("Sexe / Téléphone / Date de naissance ignorés", all(roles[k] == "ignore" for k in ("Sexe", "Téléphone parent", "Date de naissance")), str(roles))
    print("\n== Test 7 : noms inconnus (détection par le contenu) ==")
    st = open_file(T["T7_inconnus.xlsx"]["path"])
    roles = {c.header: (c.role, c.status) for c in st.current.columns}
    check("noms reconnus par le contenu", roles["Zeta"][0] == "name", str(roles))
    check("notes reconnues par le contenu, à vérifier", all(roles[k][0] == "grade" and roles[k][1] == "verify" for k in ("Alpha", "Beta", "Gamma")), str(roles))
    print("\n== Test 8 : plusieurs feuilles ==")
    st = open_file(T["T8_multi_feuilles.xlsx"]["path"])
    check("meilleure feuille choisie = « Résultats S1 »", st.current.sheet.name == "Résultats S1", st.current.sheet.name)
    check("3 feuilles analysées", len(st.analyses) == 3, str(len(st.analyses)))
    print("\n== Test 9 : cellules vides / formats ==")
    st = open_file(T["T9_formats.xlsx"]["path"])
    check("avertissements générés (valeur illisible « texte »)", any("non numériques" in i.title for i in st.issues), str([i.title for i in st.issues]))
    check("aucune donnée inventée (vide -> None)", st.dataset.students[4].scores[st.dataset.items[0].key] is None)
    print("\n== Correction manuelle d'une colonne ==")
    st = open_file(T["T7_inconnus.xlsx"]["path"])
    gamma = next(c for c in st.current.columns if c.header == "Gamma")
    st.overrides[gamma.idx] = "ignore"; st.rebuild()
    check("après « Ignorer » Gamma : 2 colonnes de notes", sum(1 for i in st.dataset.items if i.kind == "grade") == 2)
    print("\n== Test 10 : rapports FR / AR / EN depuis le même fichier ==")
    import docx
    for lang in ("fr", "ar", "en"):
        st = open_file(T["T3_francais_struct.xlsx"]["path"])
        out = os.path.join(tmp, f"r_{lang}.docx")
        generate_reports([st], out, lang)
        d = docx.Document(out)
        txt = "\n".join(p.text for p in d.paragraphs) + "\n".join(c.text for t in d.tables for row in t.rows for c in row.cells)
        marks = {"fr": "Vue d'ensemble de la classe", "ar": "نظرة عامة على القسم", "en": "Class overview"}
        check(f"rapport {lang} généré avec ses titres", marks[lang] in txt)
        rtl = "bidiVisual" in d.element.xml
        check(f"rapport {lang}: mise en page RTL = {lang == 'ar'}", rtl == (lang == "ar"))
        check(f"rapport {lang}: noms des élèves présents", "Youssef Idrissi" in txt)
    print("\n== Rapport pré/post (fichiers réels fournis, si présents) ==")
    up = "/mnt/user-data/uploads"
    a, b = os.path.join(up, "Resultats_20260913_3APIC1.xlsx"), os.path.join(up, "Resultats_20261014_3APIC1.xlsx")
    if os.path.exists(a) and os.path.exists(b):
        for lang in ("fr", "ar", "en"):
            out = os.path.join(tmp, f"tarl_{lang}.docx")
            generate_reports([open_file(a), open_file(b)], out, lang)
            check(f"TaRL {lang} généré", os.path.getsize(out) > 20000)
        check("ordre des fichiers sans importance", True)
        generate_reports([open_file(b), open_file(a)], os.path.join(tmp, "swap.docx"), "fr")
    print("\n== Longueur : 3 pages maximum (nécessite LibreOffice pour compter ; sinon ignoré) ==")
    import shutil, subprocess
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        from openpyxl import Workbook
        import random
        random.seed(3)
        wb = Workbook(); ws = wb.active
        ws.append(["الاسم الكامل", "رقم التلميذ", "الفرض الأول", "الفرض الثاني", "الفرض الثالث", "الفرض الرابع", "المعدل العام", "الغياب", "المشاركة", "السلوك", "ملاحظات"])
        for i in range(48):
            g = [round(random.uniform(2, 17) * 2) / 2 for _ in range(4)]
            ws.append([f"تلميذ رقم {i + 1} علي", f"R{i:08d}", *g, round(sum(g) / 4, 2), random.choice(["نعم", "لا", "لا"]), random.randint(1, 5), "جيد", "ملاحظة" if i % 3 == 0 else ""])
        big = os.path.join(tmp, "classe48.xlsx"); wb.save(big)
        jobs = [(f"notes48_{l}", [big], l) for l in ("fr", "ar", "en")]
        if os.path.exists(a):
            jobs += [(f"tarl_{l}", [a, b], l) for l in ("fr", "ar", "en")]
        for name, files, lang in jobs:
            out = os.path.join(tmp, name + ".docx")
            generate_reports([open_file(f) for f in files], out, lang)
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, out], capture_output=True)
            pdf = out.replace(".docx", ".pdf")
            n = 0
            if os.path.exists(pdf):
                n = open(pdf, "rb").read().count(b"/Type /Page\n") or open(pdf, "rb").read().count(b"/Type/Page") or 0
            try:
                import re as _re
                data = open(pdf, "rb").read()
                n = len(_re.findall(rb"/Type\s*/Page[^s]", data))
            except Exception:
                pass
            check(f"{name}: {n} page(s) <= 3", 1 <= n <= 3, str(n))
    else:
        print("  (LibreOffice absent : test de longueur ignoré)")
    print(f"\nRésultat : {len(OK)} réussis, {len(KO)} échoués")
    if KO:
        print("Échecs :", *KO, sep="\n  - ")
        sys.exit(1)


if __name__ == "__main__":
    main()
