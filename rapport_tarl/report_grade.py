# -*- coding: utf-8 -*-
"""Rapport d'analyse de NOTES en TROIS pages (même charte que le rapport pré/post). FR / AR (RTL) / EN.
Les blocs qui ne peuvent pas être calculés avec les données du fichier ne sont pas affichés."""
import math
import os

from . import charts
from . import docx_helpers as H
from . import layout3 as L
from .docx_helpers import CENTER, START
from .layout3 import AMB, BLU, GRN, MUTED, NAVY, ORG, RED, TEAL, TXT, W
from .report_translator import Translator, subject_name
from .text_utils import has_arabic


def _blocks(doc, rows, heads, bw, color, size=7.5):
    """Tableau en deux blocs côte à côte (compact) ; rows = liste de listes de cellules (str ou (str, opts))."""
    half = max(1, math.ceil(len(rows) / 2))
    nc = len(heads)
    t = doc.add_table(rows=1, cols=2 * nc + 1); t.alignment = L.WD_TABLE_ALIGNMENT.CENTER
    L.table_borders(t, "D5DCE4", 4)
    widths = bw + [0.2] + bw
    for i, h in enumerate(heads + [""] + heads):
        if i == nc:
            L.cell_borders(t.rows[0].cells[i]); continue
        L.set_cell(t.rows[0].cells[i], h, size, True, "#FFFFFF", color, CENTER, (30, 30, 40, 40))
    for k in range(half):
        row = t.add_row(); L.no_split(row)
        for blk in (0, 1):
            idx = k + blk * half
            base = blk * (nc + 1)
            for i in range(nc):
                if idx >= len(rows):
                    L.set_cell(row.cells[base + i], "", size, False, None, "#FFFFFF", CENTER, (25, 25, 40, 40)); continue
                v = rows[idx][i]
                txt, o = v if isinstance(v, tuple) else (v, {})
                fill = o.get("fill", "#F4F7FB" if k % 2 == 0 else "#FFFFFF")
                L.set_cell(row.cells[base + i], txt, o.get("size", size), o.get("bold", False), o.get("color"), fill,
                           "name" if o.get("name") else CENTER, (25, 25, 40, 40))
    for row in t.rows:
        L.cell_borders(row.cells[nc])
    L.fixer_largeurs(t, widths)
    L.spacer(doc, 5)


def build_grade(r, lang, sortie, prof, director, log, tmp):
    T = Translator(lang); H.set_rtl(T.rtl)
    P = lambda n: os.path.join(tmp, n)
    meta = r["meta"]
    subject = subject_name(meta.get("matiere", ""), lang)
    year, group = meta.get("annee", ""), meta.get("groupe", "")
    sc = r["scale"]; half = sc / 2; st = r["stats"]
    sn = lambda v: T.num(v, 2 if sc <= 20 else 1)
    items = [it for it in r["items"] if it["stats"]]
    doc = L.new_doc3()
    L.footer3(doc, "  ·  ".join(x for x in (subject, year, T("k_title"), group) if x), T("k_page", x="{x}", y="{y}"))
    pass_rate = r["pass_count"] / r["n_graded"] * 100
    nm = lambda it: it["label"] if (has_arabic(it["label"]) == T.rtl and len(it["label"]) <= 40) else it["short"]

    # ======================================================= PAGE 1 : vue d'ensemble
    L.banner(doc, T("k_kicker"), T("k_subtitle_grade"), "  ·  ".join(x for x in (subject, meta.get("etablissement", ""), group, T("k_year", y=year) if year else "") if x))
    H.para(doc, T("k_intro_grade", n=r["n_total"], k=len(items), s=f"{sc:g}"), 8.5, color=TXT, before=0, after=4)
    L.sec_header(doc, "01", T("k_g_sec1"), T("k_n_graded", n=r["n_graded"]), BLU)
    L.kpi_cards(doc, [(f"{sn(st['mean'])}/{sc:g}", T("k_g_mean")), (T.pct(pass_rate, 1), T("k_rate")),
                      (f"{sn(st['median'])}/{sc:g}", T("k_median")), (sn(st["std"]), T("k_c_sd"))], BLU)
    cnts = [b["count"] for b in r["bands"]]
    charts.levels(cnts, [f"{b['lo']:g}–{b['hi']:g}" for b in r["bands"]], P("lv.png"))
    figs = [(P("lv.png"), T("k_fig", n=1, t=T("k_fig_levels")), 7.8)]
    if len(items) >= 2:
        charts.items_means([it["stats"]["mean"] / it["scale"] * sc for it in items], [it["short"] for it in items], sc, P("it.png"))
        figs.append((P("it.png"), T("k_fig", n=2, t=T("k_fig_items")), 7.8))
    else:
        figs.append((P("lv.png"), "", 0.1)) if False else None
    if len(figs) == 2:
        L.fig_row(doc, figs[::-1] if T.rtl else figs, [W / 2, W / 2])
    else:
        L.fig_row(doc, [(P("lv.png"), T("k_fig", n=1, t=T("k_fig_levels")), 11)], [W])
    L.titre(doc, T("g_s_levels"), BLU)
    cols_l = [RED, "#F59E0B", "#D9B800", "#34A878", GRN]
    rows = [[("●  " + T(f"lvl{i + 1}"), {"bold": True, "color": cols_l[i]}), f"{b['lo']:g} – {b['hi']:g}", str(b["count"]), T.pct(b["pct"], 1)]
            for i, b in enumerate(r["bands"])]
    L.table(doc, [T("g_th_level"), T("g_th_range"), T("g_th_count"), T("g_th_pct")], rows, [5.6, 4.4, 3.8, 3.8], BLU)
    L.titre(doc, T("reading"), BLU)
    L.tri_bullets(doc, [T("g_balance_text", m=sn(st["mean"]), s=f"{sc:g}", p=T.num(pass_rate, 0), pc=r["pass_count"], n=r["n_graded"],
                          a=sn(st["min"]), b=sn(st["max"])),
                        T("k_r1", m=sn(st["mean"]), n=f"{sc:g}", p=T.num(st["mean"] / sc * 100, 1), md=sn(st["median"]), sd=sn(st["std"])),
                        T("g_src_file") if r["avg_source"] == "fichier" else T("g_src_calc")], BLU)
    log("Page 1 ✓")

    # ======================================================= PAGE 2 : évaluations + élèves à accompagner
    H.nouvelle_page(doc)
    L.sec_header(doc, "02", T("k_g_sec2"), T("k_n_graded", n=r["n_graded"]), TEAL)
    if items:
        rows = []
        for it in items:
            s_ = it["stats"]
            rows.append([(it["label"] if len(it["label"]) <= 38 else it["short"], {"name": True}), (it["short"], {"bold": True}), f"{it['scale']:g}", T.num(s_["mean"], 2),
                         T.num(s_["median"], 2), f"{T.num(s_['min'], 1)} – {T.num(s_['max'], 1)}", T.pct(s_["pass_rate"], 0), str(s_["n"])])
        L.titre(doc, T("g_s_items"), TEAL)
        L.table(doc, [T("g_th_item"), T("th_code") if False else T("k_th_code"), T("g_th_scale"), T("g_th_mean"), T("g_th_median"), T("g_th_minmax"),
                      T("g_th_pass"), T("g_th_n")], rows, [5.0, 1.5, 1.6, 1.9, 1.9, 2.6, 1.7, 1.4], TEAL, align_cols={0: "name"})
        means = [(it["stats"]["mean"] / it["scale"], it) for it in items]
        if len(items) >= 2:
            hi, lo = max(means, key=lambda x: x[0])[1], min(means, key=lambda x: x[0])[1]
            boxes = [(T("k_strong"), T("k_g_strong_t", i=nm(hi), m=T.num(hi["stats"]["mean"], 1), s=f"{hi['scale']:g}"), GRN, "#F1F9F4"),
                     (T("k_watch"), T("k_g_weak_t", i=nm(lo), m=T.num(lo["stats"]["mean"], 1), s=f"{lo['scale']:g}"), AMB, "#FDF5EA")]
            L.two_boxes(doc, boxes[::-1] if T.rtl else boxes)
        ev = r.get("evolution")
        if ev:
            L.box(doc, T("g_evol_title"), T("g_evol_text", n=ev["n"], i=ev["improved"], s=ev["stable"], d=ev["declined"], a=sn(ev["first_mean"]),
                                          b=sn(ev["last_mean"]), sc=f"{sc:g}"), BLU)
    k = min(5, len(r["top"]), len(r["bottom"]))
    if k:
        t_rows = [[str(i), (s["name"], {"name": True}), (f"{sn(s['avg'])}", {"bold": True, "color": GRN})] for i, s in enumerate(r["top"][:k], 1)]
        b_rows = [[str(i), (s["name"], {"name": True}), (f"{sn(s['avg'])}", {"bold": True, "color": RED})] for i, s in enumerate(r["bottom"][:k], 1)]
        H.para(doc, T("g_best", n=k) + "   |   " + T("g_worst", n=k), 9.5, True, NAVY, before=2, after=2, keep=True)
        merged = [t_rows[i] + [""] if False else None for i in range(0)]
        # deux petits tableaux fusionnés dans un seul à 2 blocs : meilleurs | plus faibles
        tt = doc.add_table(rows=1, cols=7); tt.alignment = L.WD_TABLE_ALIGNMENT.CENTER
        L.table_borders(tt, "D5DCE4", 4)
        heads = [T("g_th_rank"), T("th_name"), T("g_th_avg")]
        for i, h in enumerate(heads + [""] + heads):
            if i == 3:
                L.cell_borders(tt.rows[0].cells[i]); continue
            L.set_cell(tt.rows[0].cells[i], h, 7.5, True, "#FFFFFF", GRN if i < 3 else RED, CENTER, (30, 30, 40, 40))
        for j in range(k):
            row = tt.add_row(); L.no_split(row)
            for blk, rr in ((0, t_rows), (1, b_rows)):
                base = blk * 4
                fill = "#F4F7FB" if j % 2 == 0 else "#FFFFFF"
                for i in range(3):
                    v = rr[j][i]
                    txt, o = v if isinstance(v, tuple) else (v, {})
                    L.set_cell(row.cells[base + i], txt, 7.5, o.get("bold", False), o.get("color"), fill, "name" if o.get("name") else CENTER, (25, 25, 40, 40))
        for row in tt.rows:
            L.cell_borders(row.cells[3])
        L.fixer_largeurs(tt, [1.2, 5.3, 2.0, 0.2, 1.2, 5.3, 2.0] if False else [1.2, 5.1, 2.3, 0.2, 1.2, 5.1, 2.3])
        L.spacer(doc, 5)
    # élèves à accompagner
    sup = r["support"]
    L.titre(doc, T("k_g_support", n=len(sup), h=f"{half:g}", s=f"{sc:g}") if sup else T("g_s_support"), RED)
    if sup:
        shown = sup[:16]
        rows = [[str(i), (s["name"], {"name": True}), (sn(s["avg"]), {"bold": True, "color": RED})] for i, s in enumerate(shown, 1)]
        _blocks(doc, rows, [T("th_no"), T("th_name"), T("g_th_avg")], [0.9, 5.3, 2.4], RED)
        if len(sup) > len(shown):
            H.note(doc, T("k_more", n=len(sup) - len(shown)))
    else:
        L.box(doc, T("g_s_support"), T("g_support_none", h=f"{half:g}", s=f"{sc:g}"), GRN, "#F1F9F4")
    log("Page 2 ✓")

    # ======================================================= PAGE 3 : résultats par élève + recommandations
    H.nouvelle_page(doc)
    L.sec_header(doc, "03", T("k_g_sec3"), T("k_n_graded", n=r["n_graded"]), NAVY)
    stu = r["students"]
    L.titre(doc, T("k_g_results"), NAVY)
    few = [it for it in items][:2]
    idx = [i for i, it in enumerate(r["items"]) if it["stats"]][:2] if len(items) <= 2 else []
    rows = []
    for i, s in enumerate(stu, 1):
        row = [str(i), (s["name"], {"name": True})]
        for j in idx:
            v = s["scores"][j]
            row.append("–" if v is None else f"{v:g}")
        if s["avg"] is None:
            row.append((T("abs_short") if s["absent"] else "–", {"fill": "#ECEEF1"}))
        else:
            good = s["avg"] >= half
            row.append((sn(s["avg"]), {"bold": True, "color": GRN if good else RED, "fill": "#D5F0DF" if good else "#F6D5D1"}))
        rows.append(row)
    heads = [T("th_no"), T("th_name")] + [r["items"][j]["short"] for j in idx] + [T("g_th_avg")]
    bw = [0.8, 4.2] + [1.2] * len(idx) + [1.6]
    sc_ = 8.7 / sum(bw)
    _blocks(doc, rows, heads, [x * sc_ for x in bw], NAVY, size=7)
    # compléments (uniquement si présents dans le fichier)
    extra = []
    a = r.get("absence")
    if a:
        extra.append(T("g_abs_text", a=a["absent"], n=a["total"]))
        if a["count_mean"] is not None:
            extra.append(T("g_abs_count", m=T.num(a["count_mean"], 1), t=f"{a['count_total']:g}"))
    for key, kn, kc in (("participation", "g_part_num", "g_part_cat"), ("behavior", "g_beh_num", "g_beh_cat")):
        d = r.get(key)
        if d:
            if d["type"] == "num":
                extra.append(T(kn, m=T.num(d["mean"], 1), a=f"{d['min']:g}", b=f"{d['max']:g}", n=d["n"]))
            else:
                extra.append(T(kc, l=("، " if T.rtl else ", ").join(f"{k_} ({n_})" for k_, n_ in d["counts"])))
    if r.get("observations"):
        extra.append(T("g_obs", n=r["observations"]))
    if extra:
        L.titre(doc, T("k_g_extra"), TEAL)
        L.tri_bullets(doc, extra, TEAL)
    # recommandations
    reco = []
    if pass_rate < 50:
        reco.append(T("g_r_low"))
    elif pass_rate >= 75:
        reco.append(T("g_r_good"))
    if r["support"]:
        reco.append(T("g_r_support", n=len(r["support"]), h=f"{half:g}", s=f"{sc:g}"))
    if r["hetero"] > 0.2:
        reco.append(T("g_r_hetero"))
    if len(items) >= 2:
        wk = min(items, key=lambda it: it["stats"]["mean"] / it["scale"])
        reco.append(T("g_r_item", i=nm(wk), m=T.num(wk["stats"]["mean"], 1), s=f"{wk['scale']:g}"))
    ev = r.get("evolution")
    if ev and ev["declined"]:
        reco.append(T("g_r_decl", d=ev["declined"]))
    if a and a["absent"] >= 0.15 * a["total"]:
        reco.append(T("g_r_abs"))
    reco.append(T("g_r_next"))
    L.titre(doc, T("recos"), NAVY)
    L.tri_bullets(doc, reco, NAVY)
    # signatures (compactes)
    L.spacer(doc, 8)
    sg = doc.add_table(rows=1, cols=2); sg.alignment = L.WD_TABLE_ALIGNMENT.CENTER
    L.fixer_largeurs(sg, [W / 2, W / 2])
    L.row_height(sg.rows[0], 1.8)
    for c, (role, nom) in zip(sg.rows[0].cells, ((T("sign_teacher"), prof), (T("sign_director"), director))):
        L.cell_borders(c, top=(6, TEAL), left=(4, "FFFFFF"), right=(4, "FFFFFF")); L.cell_margins(c, 60, 60, 100, 100)
        c.text = ""
        p = c.paragraphs[0]; p.alignment = CENTER; p.paragraph_format.space_after = L.Pt(1)
        H.run(p, role, 9, True, NAVY)
        q = c.add_paragraph(); q.alignment = CENTER; q.paragraph_format.space_after = L.Pt(0)
        H.run(q, (nom or T("dots")) + "   ·   " + T("sign_stamp"), 7.5, False, MUTED)
    log("Page 3 ✓")
    return doc, T
