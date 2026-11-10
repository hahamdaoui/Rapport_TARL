# -*- coding: utf-8 -*-
"""Traitement statistique à partir de la structure standardisée (jamais du fichier Excel brut).
Deux familles d'analyses, choisies d'après les données réellement présentes :
  * compétences validées (0/1) en deux temps (pré-test / post-test)  -> analyze_skill_pair
  * notes chiffrées d'évaluations                                   -> analyze_grades
Aucune valeur absente n'est inventée : les cellules vides sont exclues des calculs concernés."""
import math
from collections import Counter
from statistics import mean, median, pstdev

import numpy as np

from . import ErreurRapport
from .data_normalizer import Dataset
from .text_utils import normalize, to_number

NAN = float("nan")


# =========================================================================== compétences 0/1 (pré / post)
def _match_items(pre: Dataset, post: Dataset):
    """Associe les colonnes du pré-test à celles du post-test (code, intitulé, puis position)."""
    pi = [i for i in pre.items if i.kind == "skill"]
    qi = [i for i in post.items if i.kind == "skill"]
    pairs, used = [], set()

    def find(it):
        for cand in qi:
            if cand.col_idx in used:
                continue
            if not it.code_auto and not cand.code_auto and cand.short == it.short:
                return cand
        for cand in qi:
            if cand.col_idx in used:
                continue
            if normalize(cand.label) == normalize(it.label):
                return cand
        return None
    for it in pi:
        c = find(it)
        if c:
            used.add(c.col_idx)
            pairs.append((it, c))
    if len(pairs) < min(len(pi), len(qi)) and len(pi) == len(qi):
        pairs = list(zip(pi, qi))          # même nombre de colonnes : association par position
    return pairs


def _align_students(pre: Dataset, post: Dataset):
    by_post = {s.key: s for s in post.students}
    hits = sum(1 for s in pre.students if s.key in by_post)
    if hits < 0.5 * max(1, len(pre.students)):               # peu de correspondances : on essaie par nom
        by_post = {normalize(s.name): s for s in post.students}
        keyf = lambda s: normalize(s.name)
    else:
        keyf = lambda s: s.key
    out, matched = [], set()
    for s in pre.students:
        q = by_post.get(keyf(s))
        if q is not None:
            matched.add(id(q))
        out.append((s, q))
    only_post = [q for q in post.students if id(q) not in matched]
    return out, only_post


def _rate(vals):
    vals = [v for v in vals if v is not None]
    return (sum(vals) / len(vals) * 100) if vals else NAN


def analyze_skill_pair(pre: Dataset, post: Dataset, log=print):
    pairs = _match_items(pre, post)
    if not pairs:
        raise ErreurRapport("Les colonnes du pré-test et du post-test ne correspondent pas.\n"
                            "Vérifiez que les deux fichiers portent sur les mêmes domaines / étapes.")
    if len(pairs) < len(pre.items):
        log(f"Attention : {len(pre.items) - len(pairs)} colonne(s) du pré-test sans équivalent dans le post-test (ignorées).")
    aligned, only_post = _align_students(pre, post)
    if only_post:
        log(f"Attention : {len(only_post)} élève(s) présent(s) seulement dans le post-test (ignorés).")
    if not aligned:
        raise ErreurRapport("Aucun élève dans le pré-test.")
    k = len(pairs)
    rows = []
    for s, q in aligned:
        pa = [s.scores.get(p.key) for p, _ in pairs]
        pb = [q.scores.get(c.key) if q else None for _, c in pairs]
        rows.append({"name": s.name, "id": s.sid, "abs_pre": s.absent, "abs_post": (q.absent if q else True),
                     "missing_post": q is None, "pre": pa, "post": pb})
    n = len(rows)
    pre_ok = [not r["abs_pre"] for r in rows]
    post_ok = [not r["abs_post"] for r in rows]
    both = [a and b for a, b in zip(pre_ok, post_ok)]
    n_pre, n_post, n_pair = sum(pre_ok), sum(post_ok), sum(both)
    if n_pair == 0:
        raise ErreurRapport("Aucun élève n'est présent aux deux tests : comparaison impossible.")

    def col(which, j, mask):
        return [rows[i][which][j] for i in range(n) if mask[i]]
    eval_pre, eval_post = [], []
    for j in range(k):
        eval_pre.append(sum(v is not None for v in col("pre", j, pre_ok)) >= 0.2 * max(1, n_pre))
        eval_post.append(sum(v is not None for v in col("post", j, post_ok)) >= 0.2 * max(1, n_post))
    communs = [j for j in range(k) if eval_pre[j] and eval_post[j]]
    if not communs:
        raise ErreurRapport("Aucun domaine n'est évalué aux deux tests.")
    nC = len(communs)
    dom = []
    for j, (pi, qi) in enumerate(pairs):
        d = {"short": pi.short if pi.short else qi.short, "label": qi.label or pi.label, "group": qi.group,
             "title": qi.title or pi.title or qi.label,
             "eval_pre": eval_pre[j], "eval_post": eval_post[j],
             "taux_pre": _rate(col("pre", j, pre_ok)) if eval_pre[j] else NAN,
             "taux_post": _rate(col("post", j, post_ok)) if eval_post[j] else NAN,
             "tp_pair": _rate(col("pre", j, both)) if eval_pre[j] else NAN,
             "tq_pair": _rate(col("post", j, both)) if eval_post[j] else NAN,
             "nb_pre": int(sum(v for v in col("pre", j, pre_ok) if v is not None)),
             "cnt_pre": sum(v is not None for v in col("pre", j, pre_ok)),
             "nb_post": int(sum(v for v in col("post", j, post_ok) if v is not None)),
             "cnt_post": sum(v is not None for v in col("post", j, post_ok))}
        dom.append(d)
    # codes uniques pour les graphiques (lettres latines / chiffres uniquement)
    used = set()
    for j, d in enumerate(dom):
        sh = d["short"] if all(ord(ch) < 128 for ch in d["short"]) and d["short"] else f"D{j + 1}"
        while sh in used:
            sh += "'"
        used.add(sh)
        d["code"] = sh
    thr = max(1, round(0.6 * nC))
    stu = []
    for i, r in enumerate(rows):
        sp = sq = None
        if both[i]:
            sp = int(sum(r["pre"][j] or 0 for j in communs))
            sq = int(sum(r["post"][j] or 0 for j in communs))
        tpre = None if r["abs_pre"] else int(sum(r["pre"][j] or 0 for j in communs))
        tpost = None if r["abs_post"] else int(sum(r["post"][j] or 0 for j in communs))
        stu.append(dict(r, both=both[i], sp=sp, sq=sq, gain=(sq - sp) if both[i] else None, tot_pre=tpre, tot_post=tpost))
    sa = [s["sp"] for s in stu if s["both"]]
    sb = [s["sq"] for s in stu if s["both"]]
    gains = [s["gain"] for s in stu if s["both"]]
    trans = {}
    for j in communs:
        t = {"maintenu": 0, "progres": 0, "regres": 0, "non": 0}
        for s in stu:
            if not s["both"]:
                continue
            a, b = s["pre"][j], s["post"][j]
            if a == 1 and b == 1: t["maintenu"] += 1
            elif a == 0 and b == 1: t["progres"] += 1
            elif a == 1 and b == 0: t["regres"] += 1
            elif a == 0 and b == 0: t["non"] += 1
        trans[j] = t
    gd = {j: dom[j]["tq_pair"] - dom[j]["tp_pair"] for j in communs}
    mx, mn = max(gd, key=gd.get), min(gd, key=gd.get)
    meta = dict(pre.meta)
    for kx, vx in post.meta.items():
        meta.setdefault(kx, vx)
    return {
        "mode": "skill", "meta": meta, "k": k, "n": n, "n_pre": n_pre, "n_post": n_post, "n_pair": n_pair,
        "dom": dom, "communs": communs, "nC": nC, "thr": thr, "students": stu, "trans": trans, "gains_dom": gd,
        "best_gain": mx, "worst_gain": mn,
        "mean_pre": mean(sa), "mean_post": mean(sb), "mean_gain": mean(gains),
        "up": sum(g > 0 for g in gains), "same": sum(g == 0 for g in gains), "down": sum(g < 0 for g in gains),
        "ge_pre": sum(x >= thr for x in sa), "ge_post": sum(x >= thr for x in sb),
        "gains": gains, "top_gain": 3 if nC >= 5 else 2,
        "abs_pre": [s["name"] for s in stu if s["abs_pre"]],
        "abs_post": [s["name"] for s in stu if s["abs_post"]],
        "excluded": [s["name"] for s in stu if not s["both"]],
        "ids_available": pre.has_id and post.has_id,
    }


# =========================================================================== notes chiffrées
BAND_EDGES = [0.0, 0.25, 0.5, 0.65, 0.8, 1.0000001]


def _stats(vals, scale):
    if not vals:
        return None
    sd = pstdev(vals) if len(vals) > 1 else 0.0
    return {"n": len(vals), "mean": mean(vals), "median": median(vals), "min": min(vals), "max": max(vals), "std": sd,
            "pass_rate": sum(v >= scale / 2 for v in vals) / len(vals) * 100, "scale": scale}


def analyze_grades(ds: Dataset, log=print):
    gi = [i for i in ds.items if i.kind == "grade"]
    if not gi:
        raise ErreurRapport("Aucune colonne de notes n'a été retenue.")
    scale = ds.scale or 20.0
    pres = [s for s in ds.students if not s.absent]
    items = []
    for it in gi:
        vals = [s.scores.get(it.key) for s in pres if s.scores.get(it.key) is not None]
        st = _stats(vals, it.scale or scale)
        items.append({"short": it.short, "label": it.label, "scale": it.scale or scale, "stats": st,
                      "n_missing": len(pres) - len(vals)})
    stu = []
    for s in ds.students:
        vals = [(s.scores[i.key], i.scale or scale) for i in gi if s.scores.get(i.key) is not None]
        avg, src = None, ""
        if s.absent and not vals:
            pass
        elif ds.has_average and s.average_given is not None:
            avg, src = s.average_given, "fichier"
        elif vals:
            avg, src = mean(v / sc for v, sc in vals) * scale, "calculée"
        stu.append({"name": s.name, "id": s.sid, "absent": s.absent, "avg": avg, "src": src, "n_notes": len(vals),
                    "scores": [s.scores.get(i.key) for i in gi], "absence_count": s.absence_count,
                    "participation": s.participation, "behavior": s.behavior, "observation": s.observation})
    with_avg = [s for s in stu if s["avg"] is not None]
    if not with_avg:
        raise ErreurRapport("Aucune note exploitable : impossible de calculer les résultats.")
    vals = [s["avg"] for s in with_avg]
    stats = _stats(vals, scale)
    counts = [0] * 5
    for v in vals:
        r = v / scale
        for b in range(5):
            if BAND_EDGES[b] <= r < BAND_EDGES[b + 1]:
                counts[b] += 1
                break
    bands = [{"lo": BAND_EDGES[b] * scale, "hi": min(1.0, BAND_EDGES[b + 1]) * scale, "count": counts[b],
              "pct": counts[b] / len(vals) * 100} for b in range(5)]
    ranked = sorted(with_avg, key=lambda s: -s["avg"])
    top = ranked[:5]
    bottom = sorted(ranked[-5:], key=lambda s: s["avg"])
    support = [s for s in sorted(with_avg, key=lambda s: s["avg"]) if s["avg"] < scale / 2]
    res = {"mode": "grade", "meta": dict(ds.meta), "scale": scale, "scale_source": ds.scale_source, "items": items,
           "students": stu, "n_total": len(ds.students), "n_present": len(pres), "n_graded": len(with_avg),
           "stats": stats, "bands": bands, "top": top, "bottom": bottom, "support": support,
           "avg_source": "fichier" if ds.has_average and any(s["src"] == "fichier" for s in with_avg) else "calculée",
           "pass_count": sum(v >= scale / 2 for v in vals), "fail_count": sum(v < scale / 2 for v in vals),
           "hetero": stats["std"] / scale if scale else 0, "coefficient": ds.coefficient}
    # évolution entre évaluations
    if len(gi) >= 2:
        tol = 0.025 * scale
        imp = dec = same = 0
        deltas = []
        for s in stu:
            sc = [(v, i.scale or scale) for v, i in zip(s["scores"], gi) if v is not None]
            if len(sc) >= 2:
                d = (sc[-1][0] / sc[-1][1] - sc[0][0] / sc[0][1]) * scale
                deltas.append((s["name"], d))
                if d > tol: imp += 1
                elif d < -tol: dec += 1
                else: same += 1
        if deltas:
            deltas.sort(key=lambda x: -x[1])
            first = [it["stats"]["mean"] / it["scale"] * scale for it in items if it["stats"]]
            res["evolution"] = {"improved": imp, "declined": dec, "stable": same, "n": len(deltas),
                                "best": deltas[:3], "worst": [d for d in deltas[-3:] if d[1] < -tol][::-1],
                                "first_mean": first[0] if first else None, "last_mean": first[-1] if first else None}
    # absences
    if ds.has_absence:
        n_abs = sum(s["absent"] for s in stu)
        counts_abs = [s["absence_count"] for s in stu if s["absence_count"] is not None]
        res["absence"] = {"absent": n_abs, "total": len(stu), "count_mean": mean(counts_abs) if counts_abs else None,
                          "count_total": sum(counts_abs) if counts_abs else None}
    # participation / comportement
    for key, has in (("participation", ds.has_participation), ("behavior", ds.has_behavior)):
        if not has:
            continue
        raw = [s[key] for s in stu if s[key] is not None and str(s[key]).strip() != ""]
        if not raw:
            continue
        nums = [to_number(v) for v in raw]
        if sum(x is not None for x in nums) >= 0.8 * len(raw):
            ok = [x for x in nums if x is not None]
            res[key] = {"type": "num", "n": len(ok), "mean": mean(ok), "min": min(ok), "max": max(ok)}
        else:
            cnt = Counter(str(v).strip() for v in raw)
            res[key] = {"type": "cat", "n": len(raw), "counts": cnt.most_common(8)}
    if ds.has_observation:
        res["observations"] = sum(1 for s in stu if s["observation"])
    return res
