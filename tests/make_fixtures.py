# -*- coding: utf-8 -*-
"""Génère des fichiers Excel de test aux structures très différentes + la vérité terrain (valeurs exactes attendues)."""
import os
import random
from openpyxl import Workbook

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")
os.makedirs(HERE, exist_ok=True)
random.seed(7)

NAMES_AR = ["أحمد بنعلي", "فاطمة الزهراء العلوي", "يوسف الإدريسي", "مريم بنجلون", "عمر الحسني", "سلمى التازي", "ياسين الفاسي",
            "هاجر المرابط", "إلياس الكتاني", "خديجة الودغيري", "رشيد الأزرق", "نادية الشرقاوي"]
NAMES_FR = ["Ahmed Benali", "Fatima Zahra Alaoui", "Youssef Idrissi", "Meryem Benjelloun", "Omar Hassani", "Salma Tazi",
            "Yassine Fassi", "Hajar Mrabet", "Ilyas Kettani", "Khadija Ouazzani", "Rachid Lazrak", "Nadia Cherkaoui"]


def grades(n, k, lo=4, hi=19):
    return [[round(random.uniform(lo, hi) * 2) / 2 for _ in range(k)] for _ in range(n)]


def save(wb, name):
    path = os.path.join(HERE, name)
    wb.save(path)
    return path


def fx():
    truth = {}
    G = grades(12, 3)

    # T1 : mathématiques, colonnes en français
    wb = Workbook(); ws = wb.active; ws.title = "Maths"
    ws.append(["Nom et prénom", "Massar", "Devoir 1", "Devoir 2", "Devoir 3", "Moyenne", "Absence", "Observation"])
    for i, n in enumerate(NAMES_FR):
        ws.append([n, f"R13{i:07d}", *G[i], round(sum(G[i]) / 3, 2), "non" if i != 4 else "oui", "Bien" if sum(G[i]) / 3 > 12 else ""])
    truth["T1_maths_fr.xlsx"] = dict(path=save(wb, "T1_maths_fr.xlsx"), names=NAMES_FR, grades=G, n_items=3)

    # T2 : arabe, colonnes entièrement en arabe
    G2 = grades(12, 2)
    wb = Workbook(); ws = wb.active; ws.title = "اللغة العربية"
    ws.append(["الاسم الكامل", "الفرض الأول", "الفرض الثاني", "المعدل العام", "السلوك"])
    for i, n in enumerate(NAMES_AR):
        ws.append([n, *G2[i], round(sum(G2[i]) / 2, 2), random.choice(["جيد", "حسن", "متوسط"])])
    truth["T2_arabe_ar.xlsx"] = dict(path=save(wb, "T2_arabe_ar.xlsx"), names=NAMES_AR, grades=G2, n_items=2)

    # T3 : français, structure différente (métadonnées au-dessus, en-têtes en ligne 6, barème dans l'intitulé)
    G3 = grades(12, 2, 3, 20)
    wb = Workbook(); ws = wb.active; ws.title = "Classe 3A"
    ws["A1"] = "Matière"; ws["B1"] = "Français"; ws["A2"] = "Classe"; ws["B2"] = "3APIC-2"; ws["A3"] = "Année scolaire"; ws["B3"] = "2026/2027"
    ws.append([]); ws.append([])
    ws.append(["N°", "Élève", "Contrôle 1 /20", "Contrôle 2 /20", "Participation", "Comportement"])
    for i, n in enumerate(NAMES_FR):
        ws.append([i + 1, n, *G3[i], random.choice([1, 2, 3, 4, 5]), random.choice(["Bon", "Moyen", "Excellent"])])
    truth["T3_francais_struct.xlsx"] = dict(path=save(wb, "T3_francais_struct.xlsx"), names=NAMES_FR, grades=G3, n_items=2)

    # T4 : colonnes dans les deux langues
    G4 = grades(12, 2)
    wb = Workbook(); ws = wb.active
    ws.append(["Nom / الاسم", "Note 1 / الفرض 1", "Note 2 / الفرض 2", "Moyenne / المعدل"])
    for i, n in enumerate(NAMES_FR):
        ws.append([n, *G4[i], round(sum(G4[i]) / 2, 2)])
    truth["T4_bilingue.xlsx"] = dict(path=save(wb, "T4_bilingue.xlsx"), names=NAMES_FR, grades=G4, n_items=2)

    # T5 : ordre différent des colonnes
    wb = Workbook(); ws = wb.active
    ws.append(["Observation", "Moyenne", "Devoir 2", "Devoir 1", "Absence", "Nom", "Prénom-Nom complet"])
    for i, n in enumerate(NAMES_FR):
        ws.append(["", round((G[i][0] + G[i][1]) / 2, 2), G[i][1], G[i][0], "non", n, n])
    truth["T5_ordre.xlsx"] = dict(path=save(wb, "T5_ordre.xlsx"), names=NAMES_FR, grades=[[g[0], g[1]] for g in G], n_items=2)

    # T6 : colonnes supplémentaires et inutiles
    wb = Workbook(); ws = wb.active
    ws.append(["Rang", "Sexe", "Date de naissance", "Élève", "", "Devoir 1", "Devoir 2", "Téléphone parent", "Moyenne", "Remarque"])
    for i, n in enumerate(NAMES_FR):
        ws.append([i + 1, "M" if i % 2 else "F", f"2011-0{(i % 9) + 1}-1{i % 9}", n, None, G[i][0], G[i][1], f"06{i:08d}", round((G[i][0] + G[i][1]) / 2, 2), "RAS"])
    truth["T6_colonnes_inutiles.xlsx"] = dict(path=save(wb, "T6_colonnes_inutiles.xlsx"), names=NAMES_FR, grades=[[g[0], g[1]] for g in G], n_items=2)

    # T7 : noms de colonnes inconnus
    wb = Workbook(); ws = wb.active
    ws.append(["Zeta", "Alpha", "Beta", "Gamma"])
    for i, n in enumerate(NAMES_FR):
        ws.append([n, G[i][0], G[i][1], G[i][2]])
    truth["T7_inconnus.xlsx"] = dict(path=save(wb, "T7_inconnus.xlsx"), names=NAMES_FR, grades=G, n_items=3)

    # T8 : plusieurs feuilles (couverture, données, feuille vide)
    wb = Workbook(); ws0 = wb.active; ws0.title = "Lisez-moi"
    ws0["A1"] = "Fichier de suivi"; ws0["A3"] = "Merci de ne pas modifier"
    ws = wb.create_sheet("Résultats S1")
    ws.append(["Nom et prénom", "Devoir 1", "Devoir 2", "Devoir 3"])
    for i, n in enumerate(NAMES_FR):
        ws.append([n, *G[i]])
    wb.create_sheet("Vide")
    truth["T8_multi_feuilles.xlsx"] = dict(path=save(wb, "T8_multi_feuilles.xlsx"), names=NAMES_FR, grades=G, n_items=3)

    # T9 : cellules vides et formats de notes variés
    wb = Workbook(); ws = wb.active
    ws.append(["Nom et prénom", "Devoir 1", "Devoir 2"])
    raw = [("Ahmed Benali", "12,5", 14), ("Fatima Zahra Alaoui", "١٥", "13/20"), ("Youssef Idrissi", " 9.75 ", None), ("Meryem Benjelloun", "abs", 11),
           ("Omar Hassani", None, None), ("Salma Tazi", 16, "-"), ("Yassine Fassi", "texte", 8), ("Hajar Mrabet", 10, 10)]
    for r in raw:
        ws.append(list(r))
    expected = [[12.5, 14.0], [15.0, 13.0], [9.75, None], [None, 11.0], [None, None], [16.0, None], [None, 8.0], [10.0, 10.0]]
    truth["T9_formats.xlsx"] = dict(path=save(wb, "T9_formats.xlsx"), names=[r[0] for r in raw], grades=expected, n_items=2, partial=True)
    return truth


if __name__ == "__main__":
    t = fx()
    print(f"{len(t)} fichiers de test créés dans {HERE}")
