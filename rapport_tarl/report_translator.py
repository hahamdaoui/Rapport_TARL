# -*- coding: utf-8 -*-
"""Textes du rapport en trois langues : français (fr), arabe (ar), anglais (en).
La langue du rapport est INDÉPENDANTE de la langue du fichier Excel.
Ajouter une langue = ajouter une colonne dans _S ; ajouter un texte = ajouter une ligne."""

LANGS = {"fr": "Français", "ar": "العربية", "en": "English"}
RTL = {"ar"}

_S = {}


def S(key, fr, ar, en):
    _S[key] = {"fr": fr, "ar": ar, "en": en}


# ------------------------------------------------------------------------------ couverture / général
S("kicker_skill", "ÉVALUATION DIAGNOSTIQUE ET FINALE", "التقويم التشخيصي والنهائي", "DIAGNOSTIC AND FINAL ASSESSMENT")
S("kicker_grade", "ANALYSE DES RÉSULTATS", "تحليل النتائج", "RESULTS ANALYSIS")
S("title_skill", "Rapport TaRL", "تقرير رائز الموضعة (TaRL)", "TaRL Report")
S("title_grade", "Rapport d'analyse des résultats", "تقرير تحليل النتائج", "Results Analysis Report")
S("subtitle_skill", "Résultats du pré-test, du post-test et progression des élèves",
  "نتائج الرائز القبلي والبعدي وتقدم التلاميذ", "Pre-test and post-test results and student progress")
S("subtitle_grade", "Résultats, statistiques et recommandations pédagogiques",
  "النتائج والإحصائيات والتوصيات التربوية", "Results, statistics and teaching recommendations")
S("year_group", "Année scolaire {y}   ·   Groupe {g}", "السنة الدراسية {y}   ·   القسم {g}", "School year {y}   ·   Group {g}")
S("ident", "Fiche d'identification", "بطاقة التعريف", "Identification")
S("glance", "En un coup d'œil", "نظرة سريعة", "At a glance")
S("issued", "Document établi le {d}", "حُرّر هذا التقرير بتاريخ {d}", "Report issued on {d}")
S("hdr_skill", "RAPPORT D'ÉVALUATION TaRL  ·  {s}  ·  Groupe {g}", "تقرير تقويم TaRL  ·  {s}  ·  القسم {g}", "TaRL ASSESSMENT REPORT  ·  {s}  ·  Group {g}")
S("hdr_grade", "RAPPORT D'ANALYSE DES RÉSULTATS  ·  {s}  ·  Groupe {g}", "تقرير تحليل النتائج  ·  {s}  ·  القسم {g}", "RESULTS ANALYSIS REPORT  ·  {s}  ·  Group {g}")
S("ftr", "Rapport TaRL · {y}", "تقرير TaRL · {y}", "TaRL Report · {y}")
S("page_of", "Page {x} / {y}", "الصفحة {x} / {y}", "Page {x} / {y}")
S("academy", "Académie", "الأكاديمية", "Academy")
S("direction", "Direction provinciale", "المديرية الإقليمية", "Provincial directorate")
S("school", "Établissement", "المؤسسة", "School")
S("year", "Année scolaire", "السنة الدراسية", "School year")
S("subject", "Matière", "المادة", "Subject")
S("group", "Groupe / classe", "القسم / المجموعة", "Group / class")
S("teacher", "Professeur", "الأستاذ(ة)", "Teacher")
S("director", "Directeur", "المدير(ة)", "Principal")
S("enrolled", "Effectif inscrit", "عدد التلاميذ المسجلين", "Students enrolled")
S("present_pre", "Présents au pré-test", "الحاضرون في الرائز القبلي", "Present at pre-test")
S("present_post", "Présents au post-test", "الحاضرون في الرائز البعدي", "Present at post-test")
S("present_both", "Présents aux 2 tests", "الحاضرون في الرائزين", "Present at both tests")
S("present_n", "Élèves évalués", "التلاميذ المقيَّمون", "Students assessed")
S("students_n", "{n} élèves", "{n} تلميذ(ة)", "{n} students")
S("dots", "………………………………", "………………………………", "………………………………")
S("na", "N/E", "غ.م", "N/A")

# ------------------------------------------------------------------------------ niveaux
S("lvl_ach", "Acquis", "متحكَّم فيه", "Achieved")
S("lvl_prog", "En cours", "في طور التحقق", "In progress")
S("lvl_weak", "Faible", "ضعيف", "Weak")
S("sense_up", "▲ hausse", "▲ ارتفاع", "▲ increase")
S("sense_down", "▼ baisse", "▼ انخفاض", "▼ decrease")
S("sense_same", "= stable", "= مستقر", "= stable")
S("not_comparable", "non comparable", "غير قابل للمقارنة", "not comparable")
S("note_levels", "Valide = 1, non valide = 0. N/E : domaine non évalué. Niveau : Acquis ≥ {a} %, En cours {c}–{b} %, Faible < {c} %.",
  "متحكَّم = 1، غير متحكَّم = 0. غ.م: مجال غير مقيَّم. المستوى: متحكَّم فيه ≥ {a} %، في طور التحقق {c}–{b} %، ضعيف < {c} %.",
  "Achieved = 1, not achieved = 0. N/A: domain not assessed. Level: Achieved ≥ {a} %, In progress {c}–{b} %, Weak < {c} %.")

# ------------------------------------------------------------------------------ sections (compétences)
S("s_synth", "Synthèse générale", "خلاصة عامة", "Overall summary")
S("s_pre", "Résultats du pré-test (évaluation initiale)", "نتائج الرائز القبلي (التقويم الأولي)", "Pre-test results (initial assessment)")
S("s_post", "Résultats du post-test (évaluation finale)", "نتائج الرائز البعدي (التقويم النهائي)", "Post-test results (final assessment)")
S("s_diff", "Différence et progression par domaine", "الفرق والتقدم حسب المجال", "Difference and progress by domain")
S("s_prof", "Profils des élèves", "فئات التلاميذ", "Student profiles")
S("s_detail", "Détail des résultats par élève", "تفاصيل النتائج حسب التلميذ", "Detailed results by student")
S("s_concl", "Conclusion et recommandations", "الخلاصة والتوصيات", "Conclusion and recommendations")
S("kpi_gain", "gain moyen ({a} → {b} / {n})", "متوسط المكتسب ({a} ← {b} / {n})", "average gain ({a} → {b} / {n})")
S("kpi_up", "élèves en progression", "تلاميذ في تقدم", "students improving")
S("kpi_same", "élèves stables", "تلاميذ مستقرون", "students unchanged")
S("kpi_down", "élève(s) en baisse", "تلاميذ في تراجع", "students declining")
S("kpi_ge", "élèves ≥ {t} domaines / {n}", "تلاميذ ≥ {t} مجالات / {n}", "students ≥ {t} domains / {n}")
S("kpi_present", "élèves présents", "تلاميذ حاضرون", "students present")
S("kpi_absent", "absent(s)", "غائبون", "absent")
S("kpi_rate", "taux moyen de validation", "متوسط نسبة التحكم", "average success rate")
S("kpi_gain_pupil", "gain moyen par élève ({a} → {b} sur {n})", "متوسط المكتسب لكل تلميذ ({a} ← {b} من {n})", "average gain per student ({a} → {b} of {n})")
S("kpi_up_of", "élèves en progression", "تلاميذ في تقدم", "students improving")
S("chart_synth", "Taux de validation par domaine ({n} élèves présents aux deux tests)",
  "نسبة التحكم حسب المجال ({n} تلميذا حاضرا في الرائزين)", "Success rate by domain ({n} students present at both tests)")
S("chart_pre", "Pré-test : taux de validation par domaine", "الرائز القبلي: نسبة التحكم حسب المجال", "Pre-test: success rate by domain")
S("chart_post", "Post-test : taux de validation par domaine", "الرائز البعدي: نسبة التحكم حسب المجال", "Post-test: success rate by domain")
S("chart_diff", "Comparaison pré-test / post-test par domaine", "مقارنة الرائز القبلي بالبعدي حسب المجال", "Pre-test / post-test comparison by domain")
S("chart_trans", "Évolution des élèves par domaine", "تطور التلاميذ حسب المجال", "Student evolution by domain")
S("chart_gains", "Répartition des élèves selon leur gain", "توزيع التلاميذ حسب المكتسب", "Distribution of students by gain")
S("gain_axis", "Gain : domaines validés en plus entre le pré-test et le post-test",
  "المكتسب: عدد المجالات المحققة زيادة بين الرائزين", "Gain: extra domains achieved between pre-test and post-test")
S("threshold_note", "Ligne verte : seuil d'acquisition ({a} %).", "الخط الأخضر: عتبة التحكم ({a} %).", "Green line: achievement threshold ({a} %).")
S("legend_pre", "Pré-test", "الرائز القبلي", "Pre-test")
S("legend_post", "Post-test", "الرائز البعدي", "Post-test")
S("legend_both", "Acquis aux 2 tests", "متحكَّم فيه في الرائزين", "Achieved in both")
S("legend_prog", "Progression (0→1)", "تقدم (0←1)", "Progress (0→1)")
S("legend_reg", "Régression (1→0)", "تراجع (1←0)", "Regression (1→0)")
S("legend_non", "Non acquis aux 2 tests", "غير متحكَّم فيه في الرائزين", "Not achieved in both")
S("key_points", "Points clés", "أهم النقاط", "Key points")
S("kp_level", "Le niveau moyen passe de {a} à {b} domaines validés sur {n} ({g} par élève).",
  "يرتفع المستوى المتوسط من {a} إلى {b} مجالات محققة من أصل {n} ({g} لكل تلميذ).",
  "The average level rises from {a} to {b} domains achieved out of {n} ({g} per student).")
S("kp_counts", "{u} élèves sur {n} ont progressé ; {s} sont restés stables ; {d} ont régressé.",
  "تقدّم {u} تلميذا من أصل {n}؛ وبقي {s} مستقرين؛ وتراجع {d}.",
  "{u} of {n} students improved; {s} stayed the same; {d} declined.")
S("kp_best", "Plus forte progression : {b} ({bv} points). Plus faible progression : {w} ({wv} points).",
  "أكبر تقدم: {b} ({bv} نقطة). أضعف تقدم: {w} ({wv} نقطة).",
  "Largest improvement: {b} ({bv} points). Smallest improvement: {w} ({wv} points).")
S("kp_ach", "Domaines acquis au post-test (≥ {a} %) : {l}.", "المجالات المتحكَّم فيها في الرائز البعدي (≥ {a} %): {l}.", "Domains achieved at post-test (≥ {a} %): {l}.")
S("kp_weak", "Domaines encore faibles au post-test (< {c} %) : {l}.", "المجالات التي ما زالت ضعيفة في الرائز البعدي (< {c} %): {l}.", "Domains still weak at post-test (< {c} %): {l}.")
S("none_word", "aucun", "لا شيء", "none")
S("absences", "Absences", "الغيابات", "Absences")
S("abs_pre", "Absents au pré-test : {l}", "الغائبون في الرائز القبلي: {l}", "Absent at pre-test: {l}")
S("abs_post", "Absents au post-test : {l}", "الغائبون في الرائز البعدي: {l}", "Absent at post-test: {l}")
S("reading", "Lecture des résultats", "قراءة النتائج", "Reading the results")
S("reading_test", "{t} : {n} élèves présents. Le domaine le mieux réussi est {b} ({bv} %) ; le plus faible est {w} ({wv} %). {g} élève(s) sur {n} valident au moins {t2} domaines sur {c}.",
  "{t}: حضر {n} تلميذا. أفضل مجال هو {b} ({bv} %) وأضعفها {w} ({wv} %). يحقق {g} تلميذا من أصل {n} على الأقل {t2} مجالات من أصل {c}.",
  "{t}: {n} students present. The best-achieved domain is {b} ({bv} %); the weakest is {w} ({wv} %). {g} of {n} students achieve at least {t2} of {c} domains.")
S("test_pre", "Pré-test", "الرائز القبلي", "Pre-test")
S("test_post", "Post-test", "الرائز البعدي", "Post-test")
S("th_domain", "Domaine", "المجال", "Domain")
S("th_parcours", "Parcours", "المسار", "Pathway")
S("parcours", "Parcours {n}", "المسار {n}", "Pathway {n}")
S("th_valid", "Élèves ayant validé", "التلاميذ المتحكمون", "Students achieving")
S("th_rate", "Taux de validation", "نسبة التحكم", "Success rate")
S("th_level", "Niveau", "المستوى", "Level")
S("dist_title", "Répartition des élèves selon le nombre de domaines validés", "توزيع التلاميذ حسب عدد المجالات المحققة", "Students by number of domains achieved")
S("th_nbdom", "Nombre de domaines validés", "عدد المجالات المحققة", "Number of domains achieved")
S("th_students", "Élèves", "التلاميذ", "Students")
S("th_pct_eff", "% de l'effectif", "% من العدد", "% of class")
S("n_dom", "{k} domaine(s)", "{k} مجال(ات)", "{k} domain(s)")
S("th_pre", "Pré-test", "القبلي", "Pre-test")
S("th_post", "Post-test", "البعدي", "Post-test")
S("th_diff", "Différence (pts)", "الفرق (نقطة)", "Difference (pts)")
S("th_mult", "Progression", "التقدم", "Progress")
S("th_sense", "Sens", "الاتجاه", "Trend")
S("note_pair", "Comparaison sur les {n} élèves présents aux deux tests.", "المقارنة تخص {n} تلميذا حاضرا في الرائزين.", "Comparison based on the {n} students present at both tests.")
S("th_both", "Acquis aux 2 tests", "متحكَّم في الرائزين", "Achieved in both")
S("th_prog01", "Progression (0→1)", "تقدم (0←1)", "Progress (0→1)")
S("th_reg10", "Régression (1→0)", "تراجع (1←0)", "Regression (1→0)")
S("th_none", "Non acquis aux 2 tests", "غير متحكَّم في الرائزين", "Not achieved in both")
S("reading_diff", "La plus forte hausse concerne {b} ({bv} points) ; la plus faible concerne {w} ({wv} points).",
  "أكبر ارتفاع يخص {b} ({bv} نقطة) وأضعفه يخص {w} ({wv} نقطة).", "The largest increase concerns {b} ({bv} points); the smallest concerns {w} ({wv} points).")
S("legend_codes", "Légende des domaines", "دليل المجالات", "Domain key")
S("th_code", "Code", "الرمز", "Code")
S("th_domain_name", "Intitulé du domaine", "عنوان المجال", "Domain title")
S("prof_diff", "Élèves en difficulté persistante (moins de {t} domaines validés au post-test)",
  "تلاميذ في صعوبة مستمرة (أقل من {t} مجالات محققة في الرائز البعدي)", "Students with persistent difficulty (fewer than {t} domains achieved at post-test)")
S("prof_reg", "Élèves en régression", "تلاميذ في تراجع", "Students who declined")
S("prof_same", "Élèves sans progression (gain nul)", "تلاميذ بلا تقدم (مكتسب منعدم)", "Students with no progress (zero gain)")
S("prof_top", "Élèves ayant le plus progressé (gain ≥ {g})", "التلاميذ الأكثر تقدما (مكتسب ≥ {g})", "Students who improved most (gain ≥ {g})")
S("count_students", "{n} élève(s)", "{n} تلميذ(ة)", "{n} student(s)")
S("no_student", "Aucun élève.", "لا يوجد تلاميذ.", "No students.")
S("note_prof", "Les chiffres à côté de chaque nom indiquent le nombre de domaines validés au pré-test → post-test.",
  "الأرقام بجانب كل اسم تبيّن عدد المجالات المحققة في الرائز القبلي ← البعدي.",
  "Numbers next to each name show domains achieved at pre-test → post-test.")
S("excluded", "Élèves exclus de la comparaison (absents à l'un des tests) : {l}", "تلاميذ مستبعدون من المقارنة (غائبون عن أحد الرائزين): {l}",
  "Students excluded from the comparison (absent from one test): {l}")
S("th_no", "N°", "رقم", "No.")
S("th_name", "Nom et prénom", "الاسم الكامل", "Student name")
S("th_prepost", "{d}\npré / post", "{d}\nقبلي / بعدي", "{d}\npre / post")
S("th_pre_s", "Pré", "قبلي", "Pre")
S("th_post_s", "Post", "بعدي", "Post")
S("th_gain", "Gain", "المكتسب", "Gain")
S("abs_short", "abs.", "غ", "abs.")
S("note_detail", "✓ = validé · ✗ = non validé · – = non évalué / absent. Pré, Post : domaines validés sur {n}. Gain = Post – Pré.",
  "✓ = متحكَّم · ✗ = غير متحكَّم · – = غير مقيَّم / غائب. قبلي، بعدي: عدد المجالات المحققة من {n}. المكتسب = البعدي – القبلي.",
  "✓ = achieved · ✗ = not achieved · – = not assessed / absent. Pre, Post: domains achieved out of {n}. Gain = Post – Pre.")
S("balance", "Bilan", "الحصيلة", "Overall assessment")
S("balance_text", "Sur {n} élèves présents aux deux tests, {u} ({up} %) ont amélioré leur résultat. Le nombre d'élèves validant au moins {t} domaines sur {c} passe de {a} ({ap} %) à {b} ({bp} %). Le domaine le plus amélioré est {m} ({mv} points).",
  "من بين {n} تلميذا حاضرا في الرائزين، حسّن {u} تلميذا ({up} %) نتيجته. وانتقل عدد التلاميذ الذين يحققون على الأقل {t} مجالات من أصل {c} من {a} ({ap} %) إلى {b} ({bp} %). وأكثر المجالات تحسنا هو {m} ({mv} نقطة).",
  "Of {n} students present at both tests, {u} ({up} %) improved. The number of students achieving at least {t} of {c} domains rose from {a} ({ap} %) to {b} ({bp} %). The most improved domain is {m} ({mv} points).")
S("recos", "Recommandations", "التوصيات", "Recommendations")
S("reco_dom", "Poursuivre le renforcement des domaines non encore acquis : {l}.", "مواصلة دعم المجالات التي لم يتم التحكم فيها بعد: {l}.", "Keep reinforcing domains not yet achieved: {l}.")
S("reco_support", "Prévoir un soutien ciblé pour les {n} élèves qui valident moins de {t} domaines (voir section {s}).",
  "برمجة دعم موجَّه لفائدة {n} تلميذا يحققون أقل من {t} مجالات (انظر القسم {s}).", "Plan targeted support for the {n} students achieving fewer than {t} domains (see section {s}).")
S("reco_follow", "Suivre individuellement les {n} élèves sans progression ou en régression.", "تتبّع فردي لـ {n} تلميذا بلا تقدم أو في تراجع.", "Follow up individually the {n} students with no progress or who declined.")
S("reco_test", "Programmer un test intermédiaire pour mesurer l'effet des séances de remédiation.", "برمجة رائز وسيط لقياس أثر حصص المعالجة.", "Schedule an interim test to measure the effect of remediation sessions.")
S("sign_teacher", "Le professeur", "الأستاذ(ة)", "Teacher")
S("sign_director", "Le directeur", "المدير(ة)", "Principal")
S("sign_stamp", "Signature et cachet", "التوقيع والخاتم", "Signature and stamp")

# ------------------------------------------------------------------------------ rapport à partir de notes
S("g_s_synth", "Synthèse de la classe", "خلاصة القسم", "Class summary")
S("g_s_stats", "Statistiques des résultats", "إحصائيات النتائج", "Results statistics")
S("g_s_levels", "Répartition des élèves par niveau", "توزيع التلاميذ حسب المستوى", "Distribution of students by level")
S("g_s_best", "Meilleurs et plus faibles résultats", "أفضل النتائج وأضعفها", "Best and weakest results")
S("g_s_items", "Analyse des évaluations", "تحليل الفروض والتقويمات", "Analysis of assessments")
S("g_s_support", "Élèves nécessitant un accompagnement", "التلاميذ الذين يحتاجون إلى دعم", "Students needing support")
S("g_s_extra", "Assiduité, participation et comportement", "المواظبة والمشاركة والسلوك", "Attendance, participation and behaviour")
S("g_s_detail", "Détail des résultats par élève", "تفاصيل النتائج حسب التلميذ", "Detailed results by student")
S("g_s_concl", "Recommandations pédagogiques et conclusion", "التوصيات التربوية والخلاصة", "Teaching recommendations and conclusion")
S("g_kpi_students", "élèves évalués", "تلاميذ مقيَّمون", "students assessed")
S("g_kpi_mean", "moyenne de la classe (/{s})", "معدل القسم (/{s})", "class average (/{s})")
S("g_kpi_pass", "taux de réussite", "نسبة النجاح", "pass rate")
S("g_kpi_median", "médiane (/{s})", "الوسيط (/{s})", "median (/{s})")
S("g_chart_levels", "Répartition des élèves par niveau", "توزيع التلاميذ حسب المستوى", "Students by level")
S("g_chart_items", "Moyenne de la classe par évaluation", "معدل القسم حسب التقويم", "Class average by assessment")
S("g_chart_hist", "Distribution des moyennes", "توزيع المعدلات", "Distribution of averages")
S("g_stat_title", "Indicateurs statistiques de la moyenne", "المؤشرات الإحصائية للمعدل", "Statistical indicators of the average")
S("g_th_indicator", "Indicateur", "المؤشر", "Indicator")
S("g_th_value", "Valeur", "القيمة", "Value")
S("g_n", "Nombre d'élèves notés", "عدد التلاميذ المنقَّطين", "Students graded")
S("g_mean", "Moyenne de la classe", "معدل القسم", "Class average")
S("g_median", "Médiane", "الوسيط", "Median")
S("g_min", "Note minimale", "أدنى نقطة", "Lowest average")
S("g_max", "Note maximale", "أعلى نقطة", "Highest average")
S("g_std", "Écart-type", "الانحراف المعياري", "Standard deviation")
S("g_pass", "Taux de réussite (≥ {h}/{s})", "نسبة النجاح (≥ {h}/{s})", "Pass rate (≥ {h}/{s})")
S("g_passed", "Élèves ayant la moyenne", "التلاميذ الحاصلون على المعدل", "Students passing")
S("g_failed", "Élèves sous la moyenne", "التلاميذ دون المعدل", "Students below pass mark")
S("g_src_file", "Les moyennes sont celles du fichier Excel.", "المعدلات المعتمدة هي المعدلات الواردة في ملف Excel.", "Averages are those given in the Excel file.")
S("g_src_calc", "Moyennes calculées à partir des notes présentes dans le fichier (moyenne simple, sans coefficient).",
  "المعدلات محسوبة انطلاقا من النقط الموجودة في الملف (معدل بسيط دون معاملات).", "Averages calculated from the grades in the file (simple mean, no coefficients).")
S("lvl1", "Très faible", "ضعيف جدا", "Very weak")
S("lvl2", "Faible", "ضعيف", "Weak")
S("lvl3", "Passable", "متوسط", "Fair")
S("lvl4", "Bien", "جيد", "Good")
S("lvl5", "Très bien", "جيد جدا", "Very good")
S("g_th_level", "Niveau", "المستوى", "Level")
S("g_th_range", "Intervalle", "المجال", "Range")
S("g_th_count", "Élèves", "التلاميذ", "Students")
S("g_th_pct", "%", "%", "%")
S("g_best", "Les {n} meilleurs résultats", "أفضل {n} نتائج", "Top {n} results")
S("g_worst", "Les {n} résultats les plus faibles", "أضعف {n} نتائج", "Lowest {n} results")
S("g_th_rank", "Rang", "الرتبة", "Rank")
S("g_th_avg", "Moyenne", "المعدل", "Average")
S("g_th_item", "Évaluation", "التقويم", "Assessment")
S("g_th_scale", "Barème", "السلم", "Scale")
S("g_th_mean", "Moyenne", "المعدل", "Mean")
S("g_th_median", "Médiane", "الوسيط", "Median")
S("g_th_minmax", "Min – Max", "أدنى – أعلى", "Min – Max")
S("g_th_pass", "Réussite", "النجاح", "Pass")
S("g_th_n", "Notés", "المنقَّطون", "Graded")
S("g_item_key", "Légende des évaluations", "دليل التقويمات", "Assessment key")
S("g_evol_title", "Évolution entre la première et la dernière évaluation", "التطور بين أول تقويم وآخر تقويم", "Change between the first and last assessment")
S("g_evol_text", "Sur {n} élèves comparables : {i} ont progressé, {s} sont stables et {d} ont régressé. La moyenne de la classe passe de {a} à {b} (sur {sc}).",
  "من بين {n} تلميذا قابلا للمقارنة: تقدّم {i}، واستقر {s}، وتراجع {d}. ينتقل معدل القسم من {a} إلى {b} (على {sc}).",
  "Among {n} comparable students: {i} improved, {s} are stable and {d} declined. The class average moves from {a} to {b} (out of {sc}).")
S("g_support_intro", "{n} élève(s) ont une moyenne inférieure à {h}/{s} et nécessitent un accompagnement.",
  "{n} تلميذا لهم معدل أقل من {h}/{s} ويحتاجون إلى دعم.", "{n} student(s) have an average below {h}/{s} and need support.")
S("g_support_none", "Aucun élève n'a une moyenne inférieure à {h}/{s}.", "لا يوجد تلميذ معدله أقل من {h}/{s}.", "No student has an average below {h}/{s}.")
S("g_abs_text", "{a} élève(s) absent(s) sur {n}.", "{a} تلميذا غائبا من أصل {n}.", "{a} student(s) absent out of {n}.")
S("g_abs_count", "Nombre moyen d'absences par élève : {m} (total : {t}).", "متوسط الغيابات لكل تلميذ: {m} (المجموع: {t}).", "Average absences per student: {m} (total: {t}).")
S("g_part_num", "Participation : moyenne {m} (min {a}, max {b}) sur {n} élèves.", "المشاركة: المعدل {m} (أدنى {a}، أعلى {b}) لدى {n} تلميذا.", "Participation: mean {m} (min {a}, max {b}) over {n} students.")
S("g_beh_num", "Comportement : moyenne {m} (min {a}, max {b}) sur {n} élèves.", "السلوك: المعدل {m} (أدنى {a}، أعلى {b}) لدى {n} تلميذا.", "Behaviour: mean {m} (min {a}, max {b}) over {n} students.")
S("g_part_cat", "Participation : {l}.", "المشاركة: {l}.", "Participation: {l}.")
S("g_beh_cat", "Comportement : {l}.", "السلوك: {l}.", "Behaviour: {l}.")
S("g_obs", "{n} observation(s) saisie(s) par l'enseignant dans le fichier.", "{n} ملاحظة أدخلها الأستاذ في الملف.", "{n} teacher comment(s) entered in the file.")
S("g_th_absent", "Absent", "غائب", "Absent")
S("g_yes", "oui", "نعم", "yes")
S("g_note_detail", "– = pas de note. Moyenne : {src}", "– = لا توجد نقطة. المعدل: {src}", "– = no grade. Average: {src}")
S("g_balance", "Bilan", "الحصيلة", "Overall assessment")
S("g_balance_text", "La moyenne de la classe est de {m}/{s} ; {p} % des élèves ont la moyenne ({pc} sur {n}). Les résultats vont de {a} à {b}.",
  "معدل القسم هو {m}/{s}؛ ونسبة النجاح {p} % ({pc} تلميذا من أصل {n}). وتتراوح النتائج بين {a} و{b}.",
  "The class average is {m}/{s}; {p} % of students pass ({pc} of {n}). Results range from {a} to {b}.")
S("g_r_support", "Mettre en place un soutien ciblé pour les {n} élèves dont la moyenne est inférieure à {h}/{s}.", "برمجة دعم موجَّه لفائدة {n} تلميذا معدلهم أقل من {h}/{s}.", "Set up targeted support for the {n} students whose average is below {h}/{s}.")
S("g_r_hetero", "Les résultats sont très hétérogènes (écart-type élevé) : prévoir une différenciation pédagogique et des groupes de besoin.",
  "النتائج متباينة جدا (انحراف معياري مرتفع): ينبغي اعتماد بيداغوجيا فارقية وتكوين مجموعات حسب الحاجة.",
  "Results are highly varied (high standard deviation): plan differentiated teaching and needs-based groups.")
S("g_r_low", "Le taux de réussite est inférieur à 50 % : revoir les prérequis et proposer des activités de remédiation.",
  "نسبة النجاح أقل من 50 %: مراجعة المكتسبات السابقة واقتراح أنشطة للمعالجة.", "The pass rate is below 50 %: review prerequisites and offer remediation activities.")
S("g_r_good", "Le taux de réussite est satisfaisant : consolider les acquis et proposer des activités d'approfondissement aux meilleurs élèves.",
  "نسبة النجاح مرضية: ترسيخ المكتسبات واقتراح أنشطة تعميق لفائدة المتفوقين.", "The pass rate is satisfactory: consolidate achievements and offer enrichment to top students.")
S("g_r_item", "L'évaluation « {i} » a obtenu les résultats les plus faibles (moyenne {m}/{s}) : y revenir en classe.",
  "حصل التقويم «{i}» على أضعف النتائج (المعدل {m}/{s}): ينبغي العودة إليه داخل القسم.", "The assessment “{i}” had the weakest results (mean {m}/{s}): revisit it in class.")
S("g_r_decl", "{d} élève(s) ont régressé entre la première et la dernière évaluation : un suivi individuel est conseillé.",
  "تراجع {d} تلميذا بين أول تقويم وآخر تقويم: يُنصح بتتبع فردي.", "{d} student(s) declined between the first and last assessment: individual follow-up is advised.")
S("g_r_abs", "Un nombre notable d'absences est relevé : contacter les familles concernées.", "تم تسجيل عدد مهم من الغيابات: التواصل مع الأسر المعنية.", "A notable number of absences is recorded: contact the families concerned.")
S("g_r_next", "Programmer une nouvelle évaluation pour mesurer l'effet des mesures de remédiation.", "برمجة تقويم جديد لقياس أثر تدابير المعالجة.", "Schedule a new assessment to measure the effect of remediation measures.")
S("subjects_note", "", "", "")

# ------------------------------------------------------------------------------ matières
SUBJECTS = [
    (["الرياضيات", "رياضيات", "mathematiques", "mathematique", "maths", "math", "mathematics"], {"fr": "Mathématiques", "ar": "الرياضيات", "en": "Mathematics"}),
    (["اللغه العربيه", "العربيه", "arabe", "arabic", "langue arabe"], {"fr": "Langue arabe", "ar": "اللغة العربية", "en": "Arabic"}),
    (["اللغه الفرنسيه", "الفرنسيه", "francais", "french", "langue francaise"], {"fr": "Français", "ar": "اللغة الفرنسية", "en": "French"}),
    (["اللغه الانجليزيه", "الانجليزيه", "anglais", "english", "langue anglaise"], {"fr": "Anglais", "ar": "اللغة الإنجليزية", "en": "English"}),
    (["علوم الحياه والارض", "svt", "sciences de la vie et de la terre", "life and earth sciences", "biology"], {"fr": "SVT", "ar": "علوم الحياة والأرض", "en": "Life and Earth Sciences"}),
    (["الفيزياء والكيمياء", "physique chimie", "physique-chimie", "physics and chemistry", "pc"], {"fr": "Physique-Chimie", "ar": "الفيزياء والكيمياء", "en": "Physics and Chemistry"}),
    (["الاجتماعيات", "التاريخ والجغرافيا", "histoire geographie", "histoire-geographie", "history and geography", "social studies"], {"fr": "Histoire-Géographie", "ar": "الاجتماعيات", "en": "History and Geography"}),
    (["التربيه الاسلاميه", "education islamique", "islamic education"], {"fr": "Éducation islamique", "ar": "التربية الإسلامية", "en": "Islamic Education"}),
    (["المعلوميات", "informatique", "computer science", "ict"], {"fr": "Informatique", "ar": "المعلوميات", "en": "Computer Science"}),
]


def subject_name(raw, lang):
    """Traduit le nom d'une matière connue ; sinon renvoie le texte du fichier tel quel."""
    from .text_utils import normalize
    n = normalize(raw)
    if not n:
        return ""
    for keys, tr in SUBJECTS:
        for k in keys:
            if n == normalize(k) or (len(normalize(k)) > 4 and normalize(k) in n):
                return tr[lang]
    return str(raw).strip()


class Translator:
    def __init__(self, lang="fr"):
        if lang not in LANGS:
            lang = "fr"
        self.lang = lang
        self.rtl = lang in RTL

    def __call__(self, key, **kw):
        if key not in _S:
            raise KeyError(f"Texte de rapport manquant : {key}")
        txt = _S[key][self.lang]
        return txt.format(**kw) if kw else txt

    def num(self, x, d=1):
        """Nombre formaté (virgule décimale en français)."""
        if x is None or x != x:
            return self("na")
        s = f"{x:.{d}f}"
        return s.replace(".", ",") if self.lang == "fr" else s

    def pct(self, x, d=0):
        return self.num(x, d) + " %"

    def sgn(self, x, d=0):
        """Nombre signé (+3 / -2). En arabe, un marqueur gauche-à-droite garde le signe devant le nombre."""
        s = f"{x:+.{d}f}"
        s = s.replace(".", ",") if self.lang == "fr" else s
        return ("\u200e" + s) if self.rtl else s

    @property
    def arrow(self):
        return "←" if self.rtl else "→"


def missing_keys():
    return [k for k, v in _S.items() if k != "subjects_note" and any(not v[l] for l in LANGS)]


# ============================================================================== rapport compact « 3 pages »
S("k_kicker", "RAPPORT D'ANALYSE DES RÉSULTATS", "تقرير تحليل النتائج", "RESULTS ANALYSIS REPORT")
S("k_title", "Rapport d'analyse des résultats", "تقرير تحليل النتائج", "Results analysis report")
S("k_subtitle_skill", "Évaluation des acquis : pré-test et post-test", "تقييم المكتسبات: الاختبار القبلي والاختبار البعدي", "Learning assessment: pre-test and post-test")
S("k_subtitle_grade", "Résultats, statistiques et recommandations pédagogiques", "النتائج والإحصائيات والتوصيات التربوية", "Results, statistics and teaching recommendations")
S("k_year", "Année scolaire {y}", "السنة الدراسية {y}", "School year {y}")
S("k_page", "Page {x} / {y}", "صفحة {x} من {y}", "Page {x} of {y}")
S("k_intro", "Ce rapport compare les résultats d'un même test passé à deux moments successifs. Le test mesure {n} composantes ; chacune vaut 1 point (1 = acquise, 0 = non acquise), le total est donc sur {n}. {a} élèves sur {N} ont passé le pré-test et {b} le post-test ; la comparaison porte sur les {c} élèves présents aux deux tests.",
  "يقارن هذا التقرير بين نتيجتي الاختبار نفسه المطبَّق في مناسبتين متتاليتين. يقيس الاختبار {n} مكوّنات، ويُحتسب لكل مكوّن نقطة واحدة (1 = محصَّل، 0 = غير محصَّل)، فيكون المجموع من {n}. اجتاز الاختبار القبلي {a} تلميذًا من أصل {N}، والاختبار البعدي {b} تلميذًا؛ وتخص المقارنة {c} تلميذًا حضروا الاختبارين.",
  "This report compares the results of the same test taken on two successive occasions. The test measures {n} components; each is worth 1 point (1 = achieved, 0 = not achieved), so the total is out of {n}. {a} of {N} students took the pre-test and {b} took the post-test; the comparison covers the {c} students present at both.")
S("k_sec1", "Pré-test : situation de départ", "الاختبار القبلي: الوضعية الأولية", "Pre-test: starting point")
S("k_sec2", "Post-test : effet de l'accompagnement", "الاختبار البعدي: أثر المواكبة", "Post-test: effect of the support")
S("k_sec3", "Comparaison des deux tests", "مقارنة الاختبارين", "Comparison of the two tests")
S("k_n_assessed", "{n} élèves évalués", "{n} تلميذًا تم تقييمهم", "{n} students assessed")
S("k_n_both", "{n} élèves présents aux deux tests", "{n} تلميذًا حضروا الاختبارين", "{n} students present at both tests")
S("k_mean", "Moyenne générale", "المعدل العام", "Overall mean")
S("k_rate", "Taux de réussite", "نسبة النجاح", "Success rate")
S("k_median", "Médiane", "الوسيط", "Median")
S("k_zero", "élèves sans aucune composante acquise", "تلاميذ دون أي مكوّن", "students with no component achieved")
S("k_fig", "Figure {n} — {t}", "الشكل {n} — {t}", "Figure {n} — {t}")
S("k_tpre", "pré-test", "الاختبار القبلي", "pre-test")
S("k_tpost", "post-test", "الاختبار البعدي", "post-test")
S("k_fig_hist", "nombre de composantes acquises sur {n} ({t})", "عدد المكوّنات المحصَّلة من {n} ({t})", "components achieved out of {n} ({t})")
S("k_fig_rate", "taux de réussite par composante ({t})", "نسبة النجاح حسب المكوّن ({t})", "success rate by component ({t})")
S("k_fig_cmp", "évolution des taux de réussite entre les deux tests", "تطور نسب النجاح بين الاختبارين", "change in success rates between the two tests")
S("k_fig_gain", "répartition des élèves selon l'évolution de leur score", "توزيع التلاميذ حسب تطور النقطة", "students by change in score")
S("k_detail_comp", "Détail des composantes évaluées", "تفصيل المكوّنات المقيَّمة", "Assessed components in detail")
S("k_th_comp", "Composante", "المكوّن", "Component")
S("k_th_code", "Code", "الرمز", "Code")
S("k_th_ok", "Réussites", "الناجحون", "Passed")
S("k_th_pct", "Taux", "النسبة", "Rate")
S("k_th_lvl", "Niveau", "المستوى", "Level")
S("k_lv_sat", "Satisfaisant", "مرضٍ", "Satisfactory")
S("k_lv_avg", "Moyen", "متوسط", "Average")
S("k_lv_weak", "Faible", "ضعيف", "Weak")
S("k_r1", "Moyenne {m} sur {n} ({p} %), médiane {md} sur {n}, écart-type {sd}.", "المعدل {m} من {n} ({p} %)، والوسيط {md} من {n}، والانحراف المعياري {sd}.",
  "Mean {m} out of {n} ({p} %), median {md} out of {n}, standard deviation {sd}.")
S("k_r2", "Composante la mieux réussie : {b} ({bv} %). La plus fragile : {w} ({wv} %).", "أنجح مكوّن: {b} ({bv} %). وأكثرها هشاشة: {w} ({wv} %).",
  "Best-achieved component: {b} ({bv} %). Most fragile: {w} ({wv} %).")
S("k_r3", "{z} élève(s) sans aucune composante acquise ; meilleur résultat : {mx} sur {n}.", "{z} تلميذًا دون أي مكوّن محصَّل؛ وأعلى نتيجة مسجلة {mx} من {n}.",
  "{z} student(s) with no component achieved; highest score: {mx} out of {n}.")
S("k_concl", "Conclusion", "خلاصة", "Summary")
S("k_concl_pre", "Le groupe réussit surtout « {b} » ({bv} %) mais rencontre des difficultés en « {w} » ({wv} %) ; c'est sur cette composante que se jouera la progression.",
  "تنجح المجموعة خصوصًا في «{b}» ({bv} %) وتتعثر في «{w}» ({wv} %)؛ وعلى هذا المكوّن يتحدد التقدم.",
  "The group does best on “{b}” ({bv} %) and struggles with “{w}” ({wv} %); this is where progress will be decided.")
S("k_levels", "Niveaux de maîtrise", "مستويات التمكن", "Mastery levels")
S("k_th_level", "Niveau", "المستوى", "Level")
S("k_th_crit", "Critère (composantes acquises)", "المعيار (المكوّنات المحصَّلة)", "Criterion (components achieved)")
S("k_th_pre", "Pré-test", "قبلي", "Pre")
S("k_th_post", "Post-test", "بعدي", "Post")
S("k_th_evol", "Évolution", "التطور", "Change")
S("k_crit", "de {a} à {b}", "من {a} إلى {b}", "{a} to {b}")
S("k_strong", "Points forts", "نقاط القوة", "Strengths")
S("k_strong_t", "« {b} » est la composante la plus solide ({bv} %).", "«{b}» هو المكوّن الأكثر متانة ({bv} %).", "“{b}” is the strongest component ({bv} %).")
S("k_watch", "Points de vigilance", "نقاط تستدعي الانتباه", "Points of attention")
S("k_watch_t", "« {w} » reste la plus fragile ({wv} %) ; {z} élève(s) n'ont acquis aucune composante.", "«{w}» يبقى الأكثر هشاشة ({wv} %)؛ و{z} تلاميذ دون أي مكوّن محصَّل.",
  "“{w}” remains the most fragile ({wv} %); {z} student(s) achieved no component.")
S("k_cmp_ind", "Indicateur", "المؤشر", "Indicator")
S("k_cmp_diff", "Écart", "الفارق", "Change")
S("k_c_match", "Élèves comparés", "التلاميذ المطابقون", "Students compared")
S("k_c_mean", "Moyenne (sur {n})", "المعدل (من {n})", "Mean (out of {n})")
S("k_c_rate", "Taux de réussite", "نسبة النجاح", "Success rate")
S("k_c_median", "Médiane", "الوسيط", "Median")
S("k_c_sd", "Écart-type", "الانحراف المعياري", "Standard deviation")
S("k_c_zero", "Élèves à 0 / {n}", "تلاميذ بـ 0 / {n}", "Students at 0 / {n}")
S("k_c_note", "Lecture : un écart positif signifie que l'élève a gagné des composantes ; l'écart moyen du groupe est de {g} point(s) sur {n}.",
  "للقراءة: الفارق الموجب يعني أن التلميذ كسب مكوّنات؛ ومعدل الفارق لدى المجموعة {g} نقطة من {n}.",
  "Note: a positive change means the student gained components; the group's mean change is {g} point(s) out of {n}.")
S("k_pts", "pts", "نقطة", "pts")
S("k_ind", "Évolution individuelle de {n} élèves", "التطور الفردي لـ {n} تلميذًا", "Individual progress of {n} students")
S("k_ind1", "Écart moyen : {g} point(s) sur {n}, soit {p} points de pourcentage.", "الفارق المتوسط {g} نقطة من {n}، أي {p} نقطة مئوية.",
  "Mean change: {g} point(s) out of {n}, i.e. {p} percentage points.")
S("k_ind2", "{u} élèves ont progressé, {s} sont restés stables et {d} ont régressé (sur {t} comparés).", "{u} تلميذًا تقدّموا، و{s} استقرّوا، و{d} تراجعوا (من أصل {t} مطابقًا).",
  "{u} students improved, {s} stayed the same and {d} declined (out of {t} compared).")
S("k_ind3", "Plus forte hausse : « {b} » ({bv} pts) ; plus faible évolution : « {w} » ({wv} pts).", "أكبر ارتفاع: «{b}» ({bv} نقطة)؛ وأضعف تطور: «{w}» ({wv} نقطة).",
  "Largest rise: “{b}” ({bv} pts); smallest change: “{w}” ({wv} pts).")
S("k_lg_reg", "Régression", "تراجع", "Decline")
S("k_lg_same", "Stabilité", "استقرار", "No change")
S("k_lg_up", "Progression", "تقدم", "Improvement")
S("k_support", "{n} élèves à accompagner en priorité", "{n} تلميذًا يحتاجون إلى دعم مركَّز", "{n} students needing focused support")
S("k_support_sub", "dont {a} absent(s) · score du post-test ≤ {w} sur {c}", "منهم {a} غائبًا · نقطة الاختبار البعدي تساوي {w} من {c} أو أقل", "including {a} absent · post-test score {w} out of {c} or below")
S("k_th_prepost", "Pré → post", "قبلي ← بعدي", "Pre → post")
S("k_th_prio", "Priorité", "الأولوية", "Priority")
S("k_p_high", "Prioritaire", "أولوية", "Priority")
S("k_p_supp", "Soutien", "للدعم", "Support")
S("k_p_abs", "Absent", "غائب", "Absent")
S("k_more", "… et {n} autre(s) élève(s)", "… و{n} تلميذًا آخر", "… and {n} more")

# ------------------------------------------------------------------------------ rapport de notes en 3 pages
S("k_intro_grade", "Ce rapport analyse les résultats de {n} élèves à partir du fichier fourni ({k} évaluation(s), notes sur {s}). Les cellules vides sont exclues des calculs : aucune note n'est inventée.",
  "يحلل هذا التقرير نتائج {n} تلميذًا انطلاقًا من الملف المقدَّم ({k} تقويم(ات)، السلم من {s}). تُستبعد الخلايا الفارغة من الحسابات، ولا تُختلق أي نقطة.",
  "This report analyses the results of {n} students from the file provided ({k} assessment(s), marked out of {s}). Empty cells are excluded from calculations: no grade is invented.")
S("k_g_sec1", "Vue d'ensemble de la classe", "نظرة عامة على القسم", "Class overview")
S("k_g_sec2", "Analyse des évaluations et élèves à accompagner", "تحليل التقويمات والتلاميذ المحتاجون إلى دعم", "Assessments and students needing support")
S("k_g_sec3", "Résultats par élève et recommandations", "نتائج التلاميذ والتوصيات", "Student results and recommendations")
S("k_n_graded", "{n} élèves notés", "{n} تلميذًا منقَّطًا", "{n} students graded")
S("k_g_mean", "Moyenne de la classe", "معدل القسم", "Class average")
S("k_fig_levels", "répartition des élèves par niveau", "توزيع التلاميذ حسب المستوى", "students by level")
S("k_fig_items", "moyenne de la classe par évaluation", "معدل القسم حسب التقويم", "class average by assessment")
S("k_g_strong_t", "« {i} » est l'évaluation la mieux réussie (moyenne {m}/{s}).", "«{i}» هو التقويم الأكثر نجاحًا (المعدل {m}/{s}).", "“{i}” is the best-performed assessment (mean {m}/{s}).")
S("k_g_weak_t", "« {i} » est la moins bien réussie (moyenne {m}/{s}).", "«{i}» هو التقويم الأضعف نتائج (المعدل {m}/{s}).", "“{i}” is the weakest assessment (mean {m}/{s}).")
S("k_g_best", "Meilleurs résultats", "أفضل النتائج", "Top results")
S("k_g_worst", "Résultats les plus faibles", "أضعف النتائج", "Lowest results")
S("k_g_support", "{n} élèves à accompagner (moyenne < {h}/{s})", "{n} تلميذًا يحتاجون إلى دعم (المعدل < {h}/{s})", "{n} students needing support (average < {h}/{s})")
S("k_g_extra", "Assiduité, participation et comportement", "المواظبة والمشاركة والسلوك", "Attendance, participation and behaviour")
S("k_g_results", "Résultats par élève", "نتائج التلاميذ", "Results by student")
