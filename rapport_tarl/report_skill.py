# -*- coding: utf-8 -*-
"""Rapport pré-test / post-test en TROIS pages (A4), mêmes blocs que le modèle fourni :
page 1 pré-test · page 2 post-test · page 3 comparaison et élèves à accompagner. FR / AR (RTL) / EN."""
import math
import os
from statistics import mean, median, pstdev

from . import charts
from . import docx_helpers as H
from . import layout3 as L
from .docx_helpers import CENTER, END, START
from .layout3 import AMB, BLU, GRN, MUTED, NAVY, ORG, RED, TXT, W
from .report_translator import Translator, subject_name
from .text_utils import has_arabic

SAT, AVG = 50, 30          # seuils par composante (% de réussite) : satisfaisant ≥ 50, moyen ≥ 30


def _lvl(T, p):
    if p is None or p != p:
        return T("na"), MUTED
    if p >= SAT:
        return T("k_lv_sat"), GRN
    if p >= AVG:
        return T("k_lv_avg"), AMB
    return T("k_lv_weak"), RED


def _stats(xs):
    return {"mean": mean(xs), "median": median(xs), "sd": pstdev(xs) if len(xs) > 1 else 0.0, "zero": sum(1 for x in xs if x == 0), "max": max(xs)}


def build_skill(r, lang, sortie, prof, director, log, tmp):
    T = Translator(lang); H.set_rtl(T.rtl)
    P = lambda n: os.path.join(tmp, n)
    meta = r["meta"]
    subject = subject_name(meta.get("matiere", ""), lang)
    dom, nC, communs = r["dom"], r["nC"], r["communs"]
    stu = r["students"]
    year, group = meta.get("annee", ""), meta.get("groupe", "")
    doc = L.new_doc3()
    L.footer3(doc, "  ·  ".join(x for x in (subject, year, T("k_title"), group) if x), T("k_page", x="{x}", y="{y}"))

    w_max, m_max = nC // 3, math.ceil(2 * nC / 3)         # faible : 0..w · moyen : w+1..m · satisfaisant : m+1..nC
    pre_s = [s["tot_pre"] for s in stu if s["tot_pre"] is not None]
    post_s = [s["tot_post"] for s in stu if s["tot_post"] is not None]
    sp, sq = _stats(pre_s), _stats(post_s)
    codes = [dom[j]["code"] for j in communs]
    allcodes = [d["code"] for d in dom]
    tpre, tpost = T("k_tpre"), T("k_tpost")

    def nm(d):
        """Nom du composant utilisé DANS LES PHRASES : le code si l'écriture du titre n'est pas celle de la langue du rapport
        (évite les textes arabe/latin mélangés et inversés) ; le tableau garde l'intitulé complet."""
        return d["title"] if (has_arabic(d["title"]) == T.rtl) else d["code"]
    dist = lambda xs: [sum(1 for x in xs if x == k) for k in range(nC + 1)]

    def best_worst(which):
        ok = [(dom[j]["taux_" + which], j) for j in communs if dom[j]["taux_" + which] == dom[j]["taux_" + which]]
        return max(ok), min(ok)

    # ============================================================ PAGE 1 : en-tête + pré-test
    L.banner(doc, T("k_kicker"), T("k_subtitle_skill"), "  ·  ".join(x for x in (subject, meta.get("etablissement", ""), group, T("k_year", y=year) if year else "") if x))
    H.para(doc, T("k_intro", n=nC, a=r["n_pre"], N=r["n"], b=r["n_post"], c=r["n_pair"]), 8.5, color=TXT, before=0, after=4)

    def test_page(num, sec_key, color, which, n_present, st, stu_scores, figs):
        L.sec_header(doc, num, T(sec_key), T("k_n_assessed", n=n_present), color)
        rates = [dom[j]["taux_" + which] for j in range(len(dom))]
        valid = [x for x in rates if x == x]
        L.kpi_cards(doc, [(f"{T.num(st['mean'], 2)}/{nC}", T("k_mean")), (T.pct(st["mean"] / nC * 100, 1), T("k_rate")),
                          (f"{st['median']:g}/{nC}", T("k_median")), (str(st["zero"]), T("k_zero"))], color)
        charts.hbars(rates, allcodes, color, P(f"h_{which}.png"))
        charts.score_hist(dist(stu_scores), nC, w_max, m_max, P(f"s_{which}.png"))
        n_h, n_r = figs
        cell_h = (P(f"s_{which}.png"), T("k_fig", n=n_h, t=T("k_fig_hist", n=nC, t=tpre if which == "pre" else tpost)), 7.6)
        cell_r = (P(f"h_{which}.png"), T("k_fig", n=n_r, t=T("k_fig_rate", t=tpre if which == "pre" else tpost)), 7.6)
        L.fig_row(doc, [cell_r, cell_h] if T.rtl else [cell_h, cell_r], [W / 2, W / 2])
        L.titre(doc, T("k_detail_comp"), color)
        rows = []
        for d in dom:
            ev = d["eval_" + which]
            lv, col = _lvl(T, d["taux_" + which]) if ev else (T("na"), MUTED)
            rows.append([(d["title"], {}), (d["code"], {"bold": True, "size": 7.5}),
                         f"{d['nb_' + which]}/{d['cnt_' + which]}" if ev else "N/A",
                         (T.pct(d["taux_" + which], 1) if ev else T("na"), {"bold": True}), ("●  " + lv, {"bold": True, "color": col})])
        L.table(doc, [T("k_th_comp"), T("k_th_code"), T("k_th_ok"), T("k_th_pct"), T("k_th_lvl")], rows, [7.0, 1.9, 2.6, 2.6, 3.5], color,
                align_cols={0: "name"})
        (bv, bj), (wv, wj) = best_worst(which)
        L.titre(doc, T("reading"), color)
        L.tri_bullets(doc, [T("k_r1", m=T.num(st["mean"], 2), n=nC, p=T.num(st["mean"] / nC * 100, 1), md=f"{st['median']:g}", sd=T.num(st["sd"], 2)),
                            T("k_r2", b=nm(dom[bj]), bv=T.num(bv, 1), w=nm(dom[wj]), wv=T.num(wv, 1)),
                            T("k_r3", z=st["zero"], mx=f"{st['max']:g}", n=nC)], color)
        return (bv, bj), (wv, wj)

    (b1, bj1), (w1, wj1) = test_page("01", "k_sec1", ORG, "pre", r["n_pre"], sp, pre_s, (1, 2))
    L.box(doc, T("k_concl"), T("k_concl_pre", b=nm(dom[bj1]), bv=T.num(b1, 1), w=nm(dom[wj1]), wv=T.num(w1, 1)), BLU)
    log("Page 1 (pré-test) ✓")

    # ============================================================ PAGE 2 : post-test
    H.nouvelle_page(doc)
    (b2, bj2), (w2, wj2) = test_page("02", "k_sec2", BLU, "post", r["n_post"], sq, post_s, (3, 4))
    # niveaux de maîtrise
    L.titre(doc, T("k_levels"), BLU)
    bands = []
    if w_max >= 0:
        bands.append((T("k_lv_weak"), RED, 0, w_max))
    if m_max > w_max:
        bands.append((T("k_lv_avg"), AMB, w_max + 1, m_max))
    if nC > m_max:
        bands.append((T("k_lv_sat"), GRN, m_max + 1, nC))
    rows = []
    for lab, col, lo, hi in bands:
        a = sum(1 for x in pre_s if lo <= x <= hi); b = sum(1 for x in post_s if lo <= x <= hi)
        d = b - a
        rows.append([("●  " + lab, {"bold": True, "color": col}), T("k_crit", a=lo, b=hi),
                     f"{a}  ({T.pct(a / max(1, len(pre_s)) * 100, 1)})", f"{b}  ({T.pct(b / max(1, len(post_s)) * 100, 1)})",
                     (T.sgn(d) if d else "=", {"bold": True, "color": (GRN if (d > 0) == (lab != T("k_lv_weak")) else RED) if d else MUTED})])
    L.table(doc, [T("k_th_level"), T("k_th_crit"), T("k_th_pre"), T("k_th_post"), T("k_th_evol")], rows, [3.6, 4.4, 3.6, 3.6, 2.4], BLU)
    L.two_boxes(doc, ([(T("k_watch"), T("k_watch_t", w=nm(dom[wj2]), wv=T.num(w2, 1), z=sq["zero"]), AMB, "#FDF5EA"),
                       (T("k_strong"), T("k_strong_t", b=nm(dom[bj2]), bv=T.num(b2, 1)), GRN, "#F1F9F4")] if T.rtl else
                      [(T("k_strong"), T("k_strong_t", b=nm(dom[bj2]), bv=T.num(b2, 1)), GRN, "#F1F9F4"),
                       (T("k_watch"), T("k_watch_t", w=nm(dom[wj2]), wv=T.num(w2, 1), z=sq["zero"]), AMB, "#FDF5EA")]))
    log("Page 2 (post-test) ✓")

    # ============================================================ PAGE 3 : comparaison + soutien
    H.nouvelle_page(doc)
    L.sec_header(doc, "03", T("k_sec3"), T("k_n_both", n=r["n_pair"]), GRN)
    pa = [s["sp"] for s in stu if s["both"]]; pb = [s["sq"] for s in stu if s["both"]]
    A, B = _stats(pa), _stats(pb)
    g = lambda d, dec=2: T.sgn(d, dec) if abs(d) > 1e-9 else "="
    good = lambda d: GRN if d > 1e-9 else (RED if d < -1e-9 else MUTED)
    rate_a, rate_b = A["mean"] / nC * 100, B["mean"] / nC * 100
    rows = [[T("k_c_match"), str(r["n_pair"]), str(r["n_pair"]), "="],
            [T("k_c_mean", n=nC), T.num(A["mean"], 2), T.num(B["mean"], 2), (g(B["mean"] - A["mean"]), {"bold": True, "color": good(B["mean"] - A["mean"])})],
            [T("k_c_rate"), T.pct(rate_a, 1), T.pct(rate_b, 1), (g(rate_b - rate_a, 1) + " " + T("k_pts"), {"bold": True, "color": good(rate_b - rate_a)})],
            [T("k_c_median"), f"{A['median']:g}", f"{B['median']:g}", (g(B["median"] - A["median"], 0), {"bold": True, "color": good(B["median"] - A["median"])})],
            [T("k_c_sd"), T.num(A["sd"], 2), T.num(B["sd"], 2), (g(B["sd"] - A["sd"]), {"color": MUTED})],
            [T("k_c_zero", n=nC), str(A["zero"]), str(B["zero"]), (g(B["zero"] - A["zero"], 0), {"bold": True, "color": good(A["zero"] - B["zero"])})]]
    L.table(doc, [T("k_cmp_ind"), T("k_th_pre"), T("k_th_post"), T("k_cmp_diff")], rows, [6.8, 3.6, 3.6, 3.6], GRN, align_cols={0: START})
    mg = B["mean"] - A["mean"]
    H.note(doc, T("k_c_note", g=T.num(mg, 2), n=nC))
    rp = [dom[j]["taux_pre"] for j in range(len(dom))]; rq = [dom[j]["taux_post"] for j in range(len(dom))]
    charts.cmp_bars(rp, rq, allcodes, P("cmp.png"))
    L.fig_row(doc, [(P("cmp.png"), T("k_fig", n=5, t=T("k_fig_cmp")), 16.6)], [W])
    H.legende(doc, [(ORG, T("legend_pre")), (BLU, T("legend_post"))]) if False else None
    # évolution individuelle
    gains = [s["gain"] for s in stu if s["both"]]
    reg, same, up = sum(g_ < 0 for g_ in gains), sum(g_ == 0 for g_ in gains), sum(g_ > 0 for g_ in gains)
    charts.seg_bar([reg, same, up], P("seg.png"))
    dl = {j: dom[j]["taux_post"] - dom[j]["taux_pre"] for j in communs}
    jb, jw = max(dl, key=dl.get), min(dl, key=dl.get)
    t = doc.add_table(rows=1, cols=2); t.alignment = L.WD_TABLE_ALIGNMENT.CENTER
    L.table_borders(t, L.LINE, 4, inside=False)
    L.fixer_largeurs(t, [W * 0.5, W * 0.5])
    c_txt, c_bar = t.rows[0].cells
    for c in (c_txt, c_bar):
        L.cell_margins(c, 90, 80, 140, 140); c.vertical_alignment = L.WD_ALIGN_VERTICAL.CENTER
    c_txt.text = ""
    p = c_txt.paragraphs[0]; p.paragraph_format.space_after = Pt_(3)
    H.run(p, T("k_ind", n=r["n_pair"]), 9, True, NAVY)
    for line in (T("k_ind1", g=T.num(mg, 2), n=nC, p=T.num(rate_b - rate_a, 1)), T("k_ind2", u=up, s=same, d=reg, t=r["n_pair"]),
                 T("k_ind3", b=nm(dom[jb]), bv=T.sgn(dl[jb], 0), w=nm(dom[jw]), wv=T.sgn(dl[jw], 0))):
        q = c_txt.add_paragraph(); q.paragraph_format.space_after = Pt_(2); q.paragraph_format.left_indent = L.Cm(0.4); q.paragraph_format.first_line_indent = -L.Cm(0.35)
        H.run(q, "▸  ", 8, True, GRN); H.run(q, line, 8, False, TXT)
    c_bar.text = ""
    pb_ = c_bar.paragraphs[0]; pb_.alignment = CENTER; pb_.paragraph_format.space_after = Pt_(2)
    pb_.add_run().add_picture(P("seg.png"), width=L.Cm(7.6))
    H.para(c_bar, T("k_fig", n=6, t=T("k_fig_gain")), 8, False, MUTED, align=CENTER, after=2)
    H.legende(c_bar, [(RED, T("k_lg_reg")), ("#98A1AB", T("k_lg_same")), (GRN, T("k_lg_up"))]) if False else None
    lg = c_bar.add_paragraph(); lg.alignment = CENTER; lg.paragraph_format.space_after = Pt_(0)
    for i, (col, lab) in enumerate(((RED, T("k_lg_reg")), ("#98A1AB", T("k_lg_same")), (GRN, T("k_lg_up")))):
        H.run(lg, "■ ", 9, True, col); H.run(lg, lab + ("     " if i < 2 else ""), 8, False, MUTED)
    L.spacer(doc, 6)

    # élèves à accompagner
    sup = []
    for s in stu:
        if s["abs_post"]:
            sup.append((2, 0, s))
        elif s["tot_post"] is not None and s["tot_post"] <= w_max:
            sup.append((0, s["tot_post"], s))
    sup.sort(key=lambda x: (x[0], x[1], (x[2]["tot_pre"] if x[2]["tot_pre"] is not None else 99), x[2]["name"]))
    n_abs = sum(1 for x in sup if x[0] == 2)
    H.sous_titre(doc, T("k_support", n=len(sup)), NAVY)
    H.para(doc, T("k_support_sub", a=n_abs, w=w_max, c=nC), 8, False, MUTED, after=3)
    shown = sup[:16]
    half = max(1, math.ceil(len(shown) / 2))
    if shown:
        t = doc.add_table(rows=1, cols=9); t.alignment = L.WD_TABLE_ALIGNMENT.CENTER
        L.table_borders(t, "D5DCE4", 4)
        bw = [0.8, 3.9, 1.9, 2.1]
        widths = bw + [0.2] + bw
        heads = [T("th_no"), T("th_name"), T("k_th_prepost"), T("k_th_prio")]
        for i, h in enumerate(heads + [""] + heads):
            if i == 4:
                L.cell_borders(t.rows[0].cells[i]); continue
            L.set_cell(t.rows[0].cells[i], h, 7.5, True, "#FFFFFF", NAVY, CENTER, (35, 35, 40, 40))
        for k in range(half):
            row = t.add_row(); L.no_split(row)
            for blk in (0, 1):
                idx = k + blk * half
                base = blk * 5
                if idx >= len(shown):
                    for i in range(4):
                        L.set_cell(row.cells[base + i], "", 7.5, False, None, "#FFFFFF", CENTER, (30, 30, 40, 40))
                    continue
                kind, _, s = shown[idx]
                fill = "#F4F7FB" if k % 2 == 0 else "#FFFFFF"
                if kind == 2:
                    lab, col = T("k_p_abs"), MUTED
                elif s["tot_post"] == 0 or (s["gain"] is not None and s["gain"] < 0):
                    lab, col = T("k_p_high"), RED
                else:
                    lab, col = T("k_p_supp"), AMB
                prepost = (f"{s['tot_pre'] if s['tot_pre'] is not None else '–'} {T.arrow} {s['tot_post'] if s['tot_post'] is not None else '–'}")
                L.set_cell(row.cells[base], str(idx + 1), 7, False, MUTED, fill, CENTER, (30, 30, 40, 40))
                L.set_cell(row.cells[base + 1], s["name"], 7.5, False, None, fill, "name", (30, 30, 60, 40))
                L.set_cell(row.cells[base + 2], prepost, 7.5, True, None, fill, CENTER, (30, 30, 40, 40))
                L.set_cell(row.cells[base + 3], lab, 7.5, True, col, fill, CENTER, (30, 30, 40, 40))
            L.cell_borders(row.cells[4])
        for row in t.rows:
            L.cell_borders(row.cells[4])
        L.fixer_largeurs(t, widths)
        if len(sup) > len(shown):
            H.note(doc, T("k_more", n=len(sup) - len(shown)))
    log("Page 3 (comparaison) ✓")
    return doc, T


from docx.shared import Pt as Pt_  # noqa: E402
