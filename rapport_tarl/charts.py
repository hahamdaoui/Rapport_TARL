# -*- coding: utf-8 -*-
"""Graphiques (matplotlib). RÈGLE : aucune image ne contient de texte arabe -> pas de lettres séparées/inversées.
Les images n'affichent que des chiffres, des pourcentages et des codes latins ; les titres et légendes
sont écrits en texte Word (traduits et correctement orientés) par report_generator."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

NAVY, ORG, BLU, GRN, RED, GRY = "#1E3A5F", "#E07B39", "#2563EB", "#059669", "#DC2626", "#94A3B8"
SEUIL = 60


def _style():
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": "#CBD5E1",
                         "axes.spines.top": False, "axes.spines.right": False, "axes.labelcolor": "#475569",
                         "xtick.color": "#334155", "ytick.color": "#334155"})


def _fin(fig, fn):
    fig.tight_layout()
    fig.savefig(fn, dpi=200, facecolor="white")
    plt.close(fig)


def _grille(ax):
    ax.yaxis.grid(True, color="#E2E8F0"); ax.set_axisbelow(True)


def _val(v):
    return 0 if (v is None or v != v) else v


def bars(vals, labels, color, fn, ymax=110, threshold=None, fmt="{:.0f}%"):
    _style()
    fig, ax = plt.subplots(figsize=(7.2, 2.8))
    x = np.arange(len(labels))
    ax.bar(x, [_val(v) for v in vals], color=color, width=.58, zorder=3)
    if threshold is not None:
        ax.axhline(threshold, color=GRN, ls="--", lw=1, zorder=2)
    for i, v in enumerate(vals):
        ax.text(i, _val(v) + ymax * .02, "N/A" if (v is None or v != v) else fmt.format(v), ha="center", fontweight="bold", fontsize=9, color=NAVY)
    ax.set_xticks(x); ax.set_xticklabels(labels, fontweight="bold"); ax.set_ylim(0, ymax); _grille(ax)
    _fin(fig, fn)


def grouped(a, b, labels, fn):
    _style()
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    x = np.arange(len(labels)); w = .36
    ax.bar(x - w / 2, [_val(v) for v in a], w, color=ORG, zorder=3)
    ax.bar(x + w / 2, [_val(v) for v in b], w, color=BLU, zorder=3)
    for i in range(len(labels)):
        na = a[i] is None or a[i] != a[i]
        ax.text(i - w / 2, _val(a[i]) + 2, "N/A" if na else f"{a[i]:.0f}%", ha="center", fontsize=8, color="#334155")
        ax.text(i + w / 2, _val(b[i]) + 2, f"{b[i]:.0f}%", ha="center", fontsize=8, color="#334155")
        if not na:
            d = b[i] - a[i]
            ax.text(i, min(104, max(a[i], b[i]) + 12), f"{d:+.0f}", ha="center", color=GRN if d >= 0 else RED, fontweight="bold", fontsize=9)
    ax.set_xticks(x); ax.set_xticklabels(labels, fontweight="bold"); ax.set_ylim(0, 118); _grille(ax)
    _fin(fig, fn)


def stacked(trans, keys, labels, fn):
    _style()
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    cats = [("maintenu", GRN), ("progres", BLU), ("regres", RED), ("non", "#CBD5E1")]
    x = np.arange(len(keys)); bottom = np.zeros(len(keys))
    for k, col in cats:
        v = np.array([trans[j][k] for j in keys], float)
        ax.bar(x, v, bottom=bottom, color=col, width=.58, zorder=3)
        for i, val in enumerate(v):
            if val > 0:
                ax.text(i, bottom[i] + val / 2, int(val), ha="center", va="center", color="white" if col != "#CBD5E1" else "#1E293B", fontsize=8, fontweight="bold")
        bottom += v
    ax.set_xticks(x); ax.set_xticklabels(labels, fontweight="bold"); _grille(ax)
    _fin(fig, fn)


def gains_hist(g, fn):
    _style()
    fig, ax = plt.subplots(figsize=(7.2, 2.5))
    vals = list(range(int(min(min(g), 0)), int(max(g)) + 1))
    cnt = [int(sum(1 for x in g if x == v)) for v in vals]
    cols = [RED if v < 0 else ("#CBD5E1" if v == 0 else GRN) for v in vals]
    ax.bar([f"{v:+d}" if v else "0" for v in vals], cnt, color=cols, width=.58, zorder=3)
    for i, c in enumerate(cnt):
        ax.text(i, c + .2, c, ha="center", fontweight="bold", fontsize=9, color=NAVY)
    _grille(ax)
    _fin(fig, fn)


def levels(counts, labels, fn):
    _style()
    fig, ax = plt.subplots(figsize=(7.2, 2.7))
    cols = [RED, "#F59E0B", "#FACC15", "#34D399", GRN]
    ax.bar(range(len(counts)), counts, color=cols, width=.6, zorder=3)
    for i, c in enumerate(counts):
        ax.text(i, c + max(counts) * .03, int(c), ha="center", fontweight="bold", fontsize=9, color=NAVY)
    ax.set_xticks(range(len(counts))); ax.set_xticklabels(labels, fontsize=9); ax.set_ylim(0, max(counts) * 1.2 + 1); _grille(ax)
    _fin(fig, fn)


def items_means(means, labels, scale, fn):
    _style()
    fig, ax = plt.subplots(figsize=(7.2, 2.9))
    x = np.arange(len(labels))
    ax.bar(x, means, color=BLU, width=.55, zorder=3)
    ax.axhline(scale / 2, color=GRN, ls="--", lw=1, zorder=2)
    for i, v in enumerate(means):
        ax.text(i, v + scale * .02, f"{v:.1f}", ha="center", fontweight="bold", fontsize=9, color=NAVY)
    ax.set_xticks(x); ax.set_xticklabels(labels, fontweight="bold"); ax.set_ylim(0, scale * 1.12); _grille(ax)
    _fin(fig, fn)


# ======================================================================= graphiques du rapport « 3 pages »
from matplotlib.patches import FancyBboxPatch  # noqa: E402

C_ORG, C_BLU, C_GRN, C_RED, C_AMB = "#E07B39", "#1F5FA8", "#2E9E5B", "#C0392B", "#D99A1E"
C_TXT, C_MUT, C_TRACK = "#27323D", "#6E7C88", "#EEF2F7"


def hbars(rates, codes, color, fn, w=4.7):
    """Barres horizontales façon « jauge » : code du composant, piste claire, valeur en %."""
    _style()
    n = len(codes)
    fig, ax = plt.subplots(figsize=(w, 0.43 * n + 0.15))
    ax.set_xlim(0, 100); ax.set_ylim(-0.6, n - 0.4); ax.axis("off")
    for i, (c, v) in enumerate(zip(codes, rates)):
        y = n - 1 - i
        ax.add_patch(FancyBboxPatch((0, y - 0.13), 100, 0.26, boxstyle="round,pad=0,rounding_size=0.13", mutation_aspect=1 / 40,
                                    fc=C_TRACK, ec="none", transform=ax.transData))
        na = v is None or v != v
        if not na and v > 0:
            ax.add_patch(FancyBboxPatch((0, y - 0.13), max(v, 3), 0.26, boxstyle="round,pad=0,rounding_size=0.13", mutation_aspect=1 / 40,
                                        fc=color, ec="none"))
        ax.text(-3, y, c, ha="right", va="center", fontsize=10, fontweight="bold", color=C_TXT)
        ax.text(103, y, "N/A" if na else f"{v:.1f} %", ha="left", va="center", fontsize=10, color=C_TXT)
    ax.set_xlim(-22, 124)
    fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)
    fig.savefig(fn, dpi=200, facecolor="white"); plt.close(fig)


def score_hist(counts, nC, w, m, fn, size=(4.7, 2.5)):
    """Nombre d'élèves selon le nombre de composantes acquises ; couleur = niveau de maîtrise."""
    _style()
    fig, ax = plt.subplots(figsize=size)
    tot = max(1, sum(counts))
    cols = [C_RED if k <= w else (C_AMB if k <= m else C_GRN) for k in range(nC + 1)]
    ax.bar(range(nC + 1), counts, color=cols, width=.38, zorder=3)
    top = max(max(counts), 1)
    for k, c in enumerate(counts):
        ax.text(k, c + top * .03, f"{c}  ({c / tot * 100:.0f} %)" if c else "0", ha="center", va="bottom", fontsize=8.5, fontweight="bold", color="#1E293B")
    ax.set_xticks(range(nC + 1)); ax.set_ylim(0, top * 1.3 + .5); ax.set_xlim(-.6, nC + .6)
    ax.yaxis.grid(True, color="#E8EDF3"); ax.set_axisbelow(True)
    from matplotlib.ticker import MaxNLocator
    ax.yaxis.set_major_locator(MaxNLocator(integer=True, nbins=4))
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#D5DCE4")
    ax.tick_params(colors=C_MUT, labelsize=9, length=0)
    fig.tight_layout(pad=0.4)
    fig.savefig(fn, dpi=200, facecolor="white"); plt.close(fig)


def cmp_bars(a, b, codes, fn, size=(7.2, 2.15)):
    _style()
    fig, ax = plt.subplots(figsize=size)
    x = np.arange(len(codes)); w = .3
    ax.bar(x - w / 2 - .01, [0 if (v is None or v != v) else v for v in a], w, color=C_ORG, zorder=3)
    ax.bar(x + w / 2 + .01, b, w, color=C_BLU, zorder=3)
    for i in range(len(codes)):
        if not (a[i] is None or a[i] != a[i]):
            ax.text(i - w / 2, a[i] + 2, f"{a[i]:.1f}", ha="center", fontsize=7, fontweight="bold", color=C_TXT)
        ax.text(i + w / 2, b[i] + 2, f"{b[i]:.1f}", ha="center", fontsize=7, fontweight="bold", color=C_TXT)
    ax.set_xticks(x); ax.set_xticklabels(codes, fontweight="bold", fontsize=8); ax.set_ylim(0, 105)
    ax.set_yticks([0, 25, 50, 75, 100]); ax.set_yticklabels(["0%", "25%", "50%", "75%", "100%"], fontsize=6.5)
    ax.yaxis.grid(True, color="#E8EDF3"); ax.set_axisbelow(True)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#D5DCE4")
    ax.tick_params(colors=C_MUT, length=0)
    fig.tight_layout(pad=0.4)
    fig.savefig(fn, dpi=200, facecolor="white"); plt.close(fig)


def seg_bar(counts, fn, size=(4.8, 0.75)):
    """Bande tricolore : régression / stabilité / progression (effectif et %)."""
    _style()
    reg, same, up = counts
    tot = max(1, reg + same + up)
    fig, ax = plt.subplots(figsize=size)
    ax.axis("off"); ax.set_xlim(0, 100); ax.set_ylim(0, 1)
    x = 0
    for n_, col in ((reg, C_RED), (same, "#98A1AB"), (up, C_GRN)):
        wdt = n_ / tot * 100
        if wdt <= 0:
            continue
        ax.add_patch(plt.Rectangle((x, 0), wdt, 1, fc=col, ec="white", lw=1))
        if wdt > 6:
            ax.text(x + wdt / 2, .62, str(n_), ha="center", va="center", color="white", fontsize=13, fontweight="bold")
            ax.text(x + wdt / 2, .22, f"{n_ / tot * 100:.0f} %", ha="center", va="center", color="white", fontsize=8)
        x += wdt
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(fn, dpi=200, facecolor="white"); plt.close(fig)
