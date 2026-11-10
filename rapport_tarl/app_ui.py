# -*- coding: utf-8 -*-
"""Interface graphique de Rapport TARL (tkinter, hors connexion, mode clair / sombre).

Parcours de l'enseignant :
  1 Importer le(s) fichier(s) Excel  ->  2 Analyse automatique  ->  3 Vérifier / corriger les colonnes
  ->  4 Choisir la langue du rapport  ->  5 Générer et ouvrir le rapport Word.
Le design suit design-system/rapport-tarl/MASTER.md (style minimaliste « Swiss », bleu marine + accents)."""
import json
import math
import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import APP_NAME, VERSION, ErreurRapport
from .column_detector import ROLE_LABELS_FR, ROLES
from .data_validator import has_blocking
from .logo_data import LOGO_HEADER, LOGO_ICON
from .pipeline import FileState, generate_reports, open_file

CONFIG = os.path.join(os.path.expanduser("~"), ".rapport_tarl.json")

# ----------------------------------------------------------------------------- design system (voir MASTER.md)
LIGHT = dict(mode="light", bg="#F8FAFC", card="#FFFFFF", border="#E2E8F0", text="#0F172A", muted="#475569",
             input="#FFFFFF", header="#1E3A5F", header_sub="#CBD5E1", accent="#2563EB", badge="#1E3A5F", title="#1E3A5F",
             ok="#047857", warn="#B45309", err="#DC2626", info="#2563EB", logbg="#F1F5F9", logfg="#0F172A",
             b_primary="#059669", b_second="#1E3A5F", b_info="#2563EB", b_gray="#64748B", b_off="#CBD5E1", chip_bg="#F1F5F9",
             sel="#EFF6FF")
DARK = dict(mode="dark", bg="#0B1220", card="#111B2E", border="#243247", text="#E6EDF5", muted="#94A3B8",
            input="#0B1626", header="#0A1424", header_sub="#9FB3C8", accent="#60A5FA", badge="#60A5FA", title="#E6EDF5",
            ok="#34D399", warn="#FBBF24", err="#F87171", info="#60A5FA", logbg="#0A121F", logfg="#CBD5E1",
            b_primary="#059669", b_second="#2D4A73", b_info="#2563EB", b_gray="#475569", b_off="#2A3A52", chip_bg="#16233A",
            sel="#14233D")
ORG, BLU, GRN = "#E07B39", "#2563EB", "#059669"
F = "Segoe UI"


def eclaircir(hexcol, f=.14):
    c = hexcol.lstrip("#")
    r, g, b = (int(c[i:i + 2], 16) for i in (0, 2, 4))
    return "#%02X%02X%02X" % tuple(min(255, int(v + (255 - v) * f)) for v in (r, g, b))


def lire_config():
    try:
        with open(CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def ecrire_config(d):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
    except Exception:
        pass


# ----------------------------------------------------------------------------- petits composants
class ScrollFrame(tk.Frame):
    """Zone défilante (molette de souris) : l'interface reste utilisable sur un petit écran."""

    def __init__(self, parent, bg):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.win, width=e.width))
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _wheel(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")


def draw_flag(c, code, w=40, h=27):
    """Drapeaux dessinés (pas d'emoji : ils ne s'affichent pas sous Windows)."""
    c.delete("all")
    if code == "fr":
        for i, col in enumerate(("#0055A4", "#FFFFFF", "#EF4135")):
            c.create_rectangle(i * w / 3, 0, (i + 1) * w / 3, h, fill=col, outline=col)
    elif code == "en":
        c.create_rectangle(0, 0, w, h, fill="#012169", outline="#012169")
        c.create_line(0, 0, w, h, fill="#FFFFFF", width=5); c.create_line(0, h, w, 0, fill="#FFFFFF", width=5)
        c.create_line(0, 0, w, h, fill="#C8102E", width=2); c.create_line(0, h, w, 0, fill="#C8102E", width=2)
        c.create_rectangle(w / 2 - 4, 0, w / 2 + 4, h, fill="#FFFFFF", outline="#FFFFFF")
        c.create_rectangle(0, h / 2 - 4, w, h / 2 + 4, fill="#FFFFFF", outline="#FFFFFF")
        c.create_rectangle(w / 2 - 2, 0, w / 2 + 2, h, fill="#C8102E", outline="#C8102E")
        c.create_rectangle(0, h / 2 - 2, w, h / 2 + 2, fill="#C8102E", outline="#C8102E")
    else:  # Maroc : fond rouge, étoile verte à cinq branches
        c.create_rectangle(0, 0, w, h, fill="#C1272D", outline="#C1272D")
        cx, cy, R = w / 2, h / 2 + 1, h * .33
        pts = [(cx + R * math.sin(math.radians(144 * i)), cy - R * math.cos(math.radians(144 * i))) for i in range(5)]
        c.create_polygon([p for pt in pts for p in pt], outline="#006233", fill="", width=2)
    c.create_rectangle(0, 0, w - 1, h - 1, outline="#94A3B8")


# ============================================================================================== fenêtre principale
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        cfg = lire_config()
        self.theme = DARK if cfg.get("theme") == "dark" else LIGHT
        self.title(APP_NAME)
        h = min(900, max(660, self.winfo_screenheight() - 90))
        self.geometry(f"920x{h}")
        self.minsize(820, 600)
        self.img_header = tk.PhotoImage(data=LOGO_HEADER)
        self.img_icon = tk.PhotoImage(data=LOGO_ICON)
        try:
            self.iconphoto(True, self.img_icon)
        except Exception:
            pass
        self.states = []
        self.prof = tk.StringVar(value=cfg.get("prof", ""))
        self.directeur = tk.StringVar(value=cfg.get("directeur", ""))
        self.sortie = tk.StringVar()
        self.lang = tk.StringVar(value=cfg.get("lang", "fr"))
        self.q = queue.Queue()
        self.dernier = None
        self.busy = ""          # '' | 'analyse' | 'generation'
        self.anim = False
        self.log_lines = [("Prêt. Importez vos fichiers Excel pour commencer.", None)]
        self.protocol("WM_DELETE_WINDOW", self._fermer)
        self._build()
        self.after(100, self._poll)

    # ------------------------------------------------------------ construction
    def _build(self):
        for w in self.winfo_children():
            w.destroy()
        P = self.theme
        self.configure(bg=P["bg"])
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TProgressbar", troughcolor=P["border"], background=P["accent"], bordercolor=P["bg"],
                    lightcolor=P["accent"], darkcolor=P["accent"])
        s.configure("TScrollbar", troughcolor=P["bg"], background=P["border"], bordercolor=P["bg"], arrowcolor=P["muted"])
        self._btns = {}
        self._header()
        self._actions()          # barre fixe en bas (créée avant le contenu défilant pour rester visible)
        self.scroll = ScrollFrame(self, P["bg"])
        self.scroll.pack(fill="both", expand=True)
        self._corps(self.scroll.inner)
        self._maj()

    # ------------------------------------------------------------ en-tête
    def _header(self):
        P = self.theme
        hd = tk.Frame(self, bg=P["header"])
        hd.pack(fill="x", side="top")
        strip = tk.Frame(hd, bg=P["header"])
        strip.pack(fill="x")
        for col in (ORG, BLU, GRN):
            tk.Frame(strip, bg=col, height=4).pack(side="left", fill="x", expand=True)
        top = tk.Frame(hd, bg=P["header"])
        top.pack(fill="x", padx=18, pady=(8, 0))
        self.sw = tk.Canvas(top, width=62, height=30, bg=P["header"], highlightthickness=0, cursor="hand2")
        self.sw.pack(side="left")
        self._switch(1.0 if P["mode"] == "dark" else 0.0)
        self.sw.bind("<Button-1>", lambda e: self._basculer())
        lab = tk.Label(top, text="Mode sombre" if P["mode"] == "dark" else "Mode clair", bg=P["header"], fg=P["header_sub"],
                       font=(F, 9), cursor="hand2")
        lab.pack(side="left", padx=8)
        lab.bind("<Button-1>", lambda e: self._basculer())
        tk.Label(top, text=f"v{VERSION}", bg=P["header"], fg=P["header_sub"], font=(F, 9)).pack(side="right")
        brand = tk.Frame(hd, bg=P["header"])
        brand.pack(fill="x", padx=24, pady=(6, 8))
        tk.Label(brand, image=self.img_header, bg=P["header"]).pack(side="left")
        txt = tk.Frame(brand, bg=P["header"])
        txt.pack(side="left", padx=16)
        tk.Label(txt, text=APP_NAME, bg=P["header"], fg="white", font=(F, 24, "bold"), anchor="w").pack(anchor="w")
        tk.Label(txt, text="Analyse intelligente de vos fichiers Excel  ·  rapport Word professionnel en français, arabe ou anglais",
                 bg=P["header"], fg=P["header_sub"], font=(F, 10), anchor="w").pack(anchor="w")
        # fil d'étapes
        self.stepper = tk.Canvas(hd, height=44, bg=P["header"], highlightthickness=0)
        self.stepper.pack(fill="x", padx=18, pady=(0, 8))
        self.stepper.bind("<Configure>", lambda e: self._dessiner_etapes())

    STEPS = ["Importer", "Analyser", "Vérifier", "Langue", "Générer"]

    def _etape(self):
        if self.busy == "generation":
            return 5
        if self.busy == "analyse":
            return 2
        if not self.states:
            return 1
        if any(not getattr(s, "confirmed", False) and self._a_verifier(s) for s in self.states):
            return 3
        return 4

    def _dessiner_etapes(self):
        P = self.theme
        c = self.stepper
        c.delete("all")
        w = max(c.winfo_width(), 700)
        cur = self._etape()
        n = len(self.STEPS)
        step = (w - 60) / (n - 1)
        for i, name in enumerate(self.STEPS):
            x = 30 + i * step
            done, active = i + 1 < cur, i + 1 == cur
            if i < n - 1:
                c.create_line(x + 14, 14, x + step - 14, 14, fill="#34D399" if done else "#3B516E", width=2)
            col = "#34D399" if done else (P["accent"] if active and P["mode"] == "light" else ("#60A5FA" if active else "#3B516E"))
            c.create_oval(x - 12, 2, x + 12, 26, fill=col, outline=col)
            c.create_text(x, 14, text="✓" if done else str(i + 1), fill="white", font=(F, 9, "bold"))
            c.create_text(x, 36, text=name, fill="white" if active else P["header_sub"], font=(F, 9, "bold" if active else "normal"))

    def _switch(self, t):
        c = self.sw
        c.delete("all")
        col = "#2BA3B3" if t > .5 else "#5C7089"
        c.create_oval(1, 1, 29, 29, fill=col, outline=col)
        c.create_oval(33, 1, 61, 29, fill=col, outline=col)
        c.create_rectangle(15, 1, 47, 29, fill=col, outline=col)
        x = 4 + 32 * t
        c.create_oval(x, 4, x + 22, 26, fill="#FFFFFF", outline="#FFFFFF")
        c.create_text(x + 11, 15, text="☾" if t > .5 else "☀", fill="#1E3A5F" if t > .5 else "#D97706", font=("Segoe UI Symbol", 11, "bold"))

    def _basculer(self):
        if self.anim:
            return
        self.anim = True
        vers_sombre = self.theme["mode"] == "light"

        def etape(i):
            if i <= 8:
                t = i / 8
                self._switch(t if vers_sombre else 1 - t)
                self.after(12, lambda: etape(i + 1))
            else:
                self.theme = DARK if vers_sombre else LIGHT
                self.anim = False
                self._sauver()
                self._build()
        etape(1)

    # ------------------------------------------------------------ composants
    def _btn(self, parent, text, cmd, color, size=10, padx=14, pady=6):
        b = tk.Button(parent, text=text, command=cmd, bg=color, fg="white", activebackground=eclaircir(color),
                      activeforeground="white", relief="flat", bd=0, font=(F, size, "bold"), padx=padx, pady=pady,
                      cursor="hand2", disabledforeground="#E8EDF2")
        b.bind("<Enter>", lambda e, b=b: b.config(bg=eclaircir(self._btns[b])) if str(b["state"]) == "normal" else None)
        b.bind("<Leave>", lambda e, b=b: b.config(bg=self._btns[b]) if str(b["state"]) == "normal" else None)
        self._btns[b] = color
        return b

    def _activer(self, b, on):
        if on:
            b.config(state="normal", bg=self._btns[b], cursor="hand2")
        else:
            b.config(state="disabled", bg=self.theme["b_off"], cursor="arrow")

    def _carte(self, parent, num, titre, sous=None):
        P = self.theme
        outer = tk.Frame(parent, bg=P["bg"])
        outer.pack(fill="x", padx=22, pady=(0, 10))
        card = tk.Frame(outer, bg=P["card"], highlightbackground=P["border"], highlightthickness=1)
        card.pack(fill="x")
        tk.Frame(card, bg=P["accent"], width=4).pack(side="left", fill="y")
        inner = tk.Frame(card, bg=P["card"])
        inner.pack(side="left", fill="x", expand=True)
        head = tk.Frame(inner, bg=P["card"])
        head.pack(fill="x", padx=14, pady=(10, 2))
        tk.Label(head, text=num, bg=P["badge"], fg="white" if P["mode"] == "light" else "#0B1220", font=(F, 10, "bold"), width=3).pack(side="left")
        tk.Label(head, text="  " + titre, bg=P["card"], fg=P["title"], font=(F, 12, "bold")).pack(side="left")
        if sous:
            tk.Label(head, text=sous, bg=P["card"], fg=P["muted"], font=(F, 9, "italic")).pack(side="right")
        body = tk.Frame(inner, bg=P["card"])
        body.pack(fill="x", padx=14, pady=(2, 12))
        body.columnconfigure(1, weight=1)
        return body

    def _entry(self, parent, var, readonly=False):
        P = self.theme
        e = tk.Entry(parent, textvariable=var, bg=P["input"], fg=P["text"], insertbackground=P["text"], relief="flat", font=(F, 10),
                     highlightthickness=1, highlightbackground=P["border"], highlightcolor=P["accent"],
                     readonlybackground=P["input"], disabledforeground=P["text"])
        if readonly:
            e.config(state="readonly")
        return e

    # ------------------------------------------------------------ contenu
    def _corps(self, root):
        P = self.theme
        tk.Frame(root, bg=P["bg"], height=12).pack()
        # --- 1. fichiers
        b = self._carte(root, "1", "Fichiers Excel", "un ou plusieurs fichiers · .xlsx / .xls · n'importe quelle matière ou langue")
        self.files_box = tk.Frame(b, bg=P["card"])
        self.files_box.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.files_box.columnconfigure(0, weight=1)
        bar = tk.Frame(b, bg=P["card"])
        bar.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self._btn(bar, "＋  Ajouter des fichiers Excel…", self._ajouter, P["b_info"]).pack(side="left")
        tk.Label(bar, text="  Pré-test et post-test : choisissez les 2 fichiers (ordre libre). Notes chiffrées : 1 fichier.",
                 bg=P["card"], fg=P["muted"], font=(F, 9, "italic")).pack(side="left")
        self._rendre_fichiers()
        # --- 2. infos
        b = self._carte(root, "2", "Informations à afficher", "facultatif · mémorisées")
        for i, (var, lab) in enumerate(((self.prof, "Professeur"), (self.directeur, "Directeur"))):
            tk.Label(b, text=lab, bg=P["card"], fg=P["text"], font=(F, 10, "bold"), width=11, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
            self._entry(b, var).grid(row=i, column=1, columnspan=2, sticky="ew", padx=8, pady=4, ipady=5)
        # --- 3. langue du rapport (partie inférieure, comme demandé)
        b = self._carte(root, "3", "Langue du rapport  ·  Report Language  ·  لغة التقرير", "indépendante de la langue du fichier Excel")
        row = tk.Frame(b, bg=P["card"])
        row.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.lang_cards = {}
        for code, label in (("fr", "Français"), ("ar", "العربية"), ("en", "English")):
            card = tk.Frame(row, bg=P["card"], highlightthickness=2, highlightbackground=P["border"], cursor="hand2")
            card.pack(side="left", padx=(0, 10), ipadx=8, ipady=6)
            fl = tk.Canvas(card, width=40, height=27, bg=P["card"], highlightthickness=0, cursor="hand2")
            fl.pack(side="left", padx=(8, 8), pady=4)
            draw_flag(fl, "ma" if code == "ar" else code)
            tx = tk.Label(card, text=label, bg=P["card"], fg=P["text"], font=(F, 12, "bold"), cursor="hand2")
            tx.pack(side="left", padx=(0, 10))
            for wdg in (card, fl, tx):
                wdg.bind("<Button-1>", lambda e, c=code: self._choisir_langue(c))
            self.lang_cards[code] = (card, tx, fl)
        self._maj_langues()
        # --- 4. enregistrement
        b = self._carte(root, "4", "Enregistrement du rapport")
        tk.Label(b, text="Fichier Word", bg=P["card"], fg=P["text"], font=(F, 10, "bold"), width=11, anchor="w").grid(row=0, column=0, sticky="w")
        self._entry(b, self.sortie).grid(row=0, column=1, sticky="ew", padx=8, ipady=5)
        self._btn(b, "Choisir…", self._choisir_sortie, P["b_second"]).grid(row=0, column=2)
        # --- journal
        lf = tk.Frame(root, bg=P["bg"])
        lf.pack(fill="x", padx=22, pady=(0, 14))
        tk.Label(lf, text="Journal", bg=P["bg"], fg=P["title"], font=(F, 10, "bold")).pack(anchor="w")
        self.log = tk.Text(lf, height=7, bg=P["logbg"], fg=P["logfg"], font=("Consolas", 9), relief="flat", padx=10, pady=8,
                           highlightthickness=1, highlightbackground=P["border"], wrap="word")
        self.log.pack(fill="x")
        self.log.tag_config("ok", foreground=P["ok"]); self.log.tag_config("err", foreground=P["err"]); self.log.tag_config("warn", foreground=P["warn"])
        for msg, tag in self.log_lines:
            self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _actions(self):
        P = self.theme
        bar = tk.Frame(self, bg=P["card"], highlightbackground=P["border"], highlightthickness=1)
        bar.pack(fill="x", side="bottom")
        row = tk.Frame(bar, bg=P["card"])
        row.pack(fill="x", padx=22, pady=(10, 4))
        self.btn_gen = self._btn(row, "   Générer le rapport Word   ", self._generer, P["b_primary"], size=13, pady=10)
        self.btn_gen.pack(side="left")
        self.btn_open = self._btn(row, "Ouvrir le rapport", self._ouvrir, P["b_info"], pady=10)
        self.btn_dir = self._btn(row, "Ouvrir le dossier", self._ouvrir_dossier, P["b_gray"], pady=10)
        self.btn_open.pack(side="left", padx=(12, 0)); self.btn_dir.pack(side="left", padx=(8, 0))
        self.lbl_status = tk.Label(row, text="", bg=P["card"], fg=P["muted"], font=(F, 9), anchor="e")
        self.lbl_status.pack(side="right")
        self.pb = ttk.Progressbar(bar, mode="determinate", maximum=100)
        self.pb.pack(fill="x", padx=22, pady=(2, 10))

    # ------------------------------------------------------------ liste des fichiers
    @staticmethod
    def _a_verifier(st):
        return any(st.role_of(c) != "ignore" and c.status == "verify" and c.idx not in st.overrides for c in st.current.columns) \
            or any(c.status == "verify" and st.role_of(c) == "ignore" and c.profile["filled"] and c.header and c.idx not in st.overrides
                   for c in st.current.columns)

    def _rendre_fichiers(self):
        P = self.theme
        for w in self.files_box.winfo_children():
            w.destroy()
        if not self.states:
            tk.Label(self.files_box, text="Aucun fichier importé pour le moment.", bg=P["card"], fg=P["muted"], font=(F, 10, "italic"),
                     anchor="w").grid(row=0, column=0, sticky="w", pady=6)
            return
        for i, st in enumerate(self.states):
            ds, an = st.dataset, st.current
            row = tk.Frame(self.files_box, bg=P["chip_bg"], highlightbackground=P["border"], highlightthickness=1)
            row.grid(row=i, column=0, sticky="ew", pady=3)
            row.columnconfigure(1, weight=1)
            kind = {"skill": "Compétences 0/1", "grade": "Notes", "unknown": "?"}[ds.kind]
            test = {"pre": " · Pré-test", "post": " · Post-test"}.get(an.test_type, "") if ds.kind == "skill" else ""
            tk.Label(row, text=f"{i + 1}", bg=P["badge"], fg="white" if P["mode"] == "light" else "#0B1220", font=(F, 10, "bold"), width=3).grid(row=0, column=0, rowspan=2, sticky="ns")
            tk.Label(row, text=st.name, bg=P["chip_bg"], fg=P["text"], font=(F, 10, "bold"), anchor="w").grid(row=0, column=1, sticky="w", padx=10, pady=(6, 0))
            detail = f"Feuille « {an.sheet.name} » · {an.n_students} élèves · {kind}{test} · fichier {self._lang_nom(an.language)} · {len(ds.items)} colonne(s) de résultats"
            tk.Label(row, text=detail, bg=P["chip_bg"], fg=P["muted"], font=(F, 9), anchor="w").grid(row=1, column=1, sticky="w", padx=10, pady=(0, 6))
            bloq = has_blocking(st.issues)
            nv = sum(1 for c in an.columns if c.status == "verify" and c.idx not in st.overrides)
            if bloq:
                txt, col = "⛔ À corriger", P["err"]
            elif getattr(st, "confirmed", False) or nv == 0:
                txt, col = "✔ Prêt", P["ok"]
            else:
                txt, col = f"⚠ {nv} colonne(s) à vérifier", P["warn"]
            tk.Label(row, text=txt, bg=P["chip_bg"], fg=col, font=(F, 9, "bold")).grid(row=0, column=2, rowspan=2, padx=8)
            self._btn(row, "Vérifier les colonnes", lambda s=st: self._verifier(s), P["b_second"], size=9, padx=10, pady=4).grid(row=0, column=3, rowspan=2, padx=4, pady=6)
            self._btn(row, "✕", lambda s=st: self._retirer(s), P["b_gray"], size=9, padx=8, pady=4).grid(row=0, column=4, rowspan=2, padx=(0, 8))

    @staticmethod
    def _lang_nom(code):
        return {"ar": "en arabe", "fr": "en français", "en": "en anglais", "mixed": "bilingue", "unknown": "(langue ?)"}.get(code, "")

    def _maj_langues(self):
        P = self.theme
        for code, (card, tx, fl) in self.lang_cards.items():
            sel = self.lang.get() == code
            card.config(highlightbackground=P["accent"] if sel else P["border"], bg=P["sel"] if sel else P["card"])
            tx.config(bg=P["sel"] if sel else P["card"], fg=P["title"] if sel else P["text"])
            fl.config(bg=P["sel"] if sel else P["card"])

    def _choisir_langue(self, code):
        self.lang.set(code)
        self._maj_langues()
        self._sauver()
        if self.sortie.get() and self.sortie.get().endswith(".docx"):
            self._nom_sortie_auto()

    # ------------------------------------------------------------ actions
    def _log(self, msg, tag=None):
        self.log_lines.append((msg, tag))
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _maj(self):
        busy = bool(self.busy)
        self._activer(self.btn_gen, bool(self.states) and not busy)
        self.btn_gen.config(text="   Génération en cours…   " if self.busy == "generation" else "   Générer le rapport Word   ")
        pret = bool(self.dernier) and os.path.exists(self.dernier or "")
        self._activer(self.btn_open, pret and not busy)
        self._activer(self.btn_dir, pret and not busy)
        self.lbl_status.config(text={"analyse": "Analyse du fichier…", "generation": "Traitement en cours…"}.get(self.busy, ""))
        self._dessiner_etapes()

    def _ajouter(self):
        init = lire_config().get("dossier") or os.path.expanduser("~")
        chemins = filedialog.askopenfilenames(title="Choisir les fichiers Excel", initialdir=init,
                                              filetypes=[("Fichiers Excel", "*.xlsx *.xlsm *.xls"), ("Tous les fichiers", "*.*")])
        if not chemins:
            return
        self._sauver(dossier=os.path.dirname(chemins[0]))
        self.busy = "analyse"
        self.pb.config(mode="indeterminate"); self.pb.start(12)
        self._maj()
        for c in chemins:
            self._log(f"Analyse de {os.path.basename(c)}…")
        threading.Thread(target=self._analyser, args=(list(chemins),), daemon=True).start()

    def _analyser(self, chemins):
        for c in chemins:
            if any(s.path == c for s in self.states):
                self.q.put(("log", f"{os.path.basename(c)} est déjà importé.", "warn")); continue
            try:
                self.q.put(("opened", open_file(c)))
            except ErreurRapport as e:
                self.q.put(("log", f"✘ {os.path.basename(c)} : {str(e).splitlines()[0]}", "err"))
                self.q.put(("alerte", f"{os.path.basename(c)}\n\n{e}"))
            except Exception:  # noqa: BLE001
                self.q.put(("log", f"✘ {os.path.basename(c)} : erreur inattendue", "err"))
                self.q.put(("alerte", "Erreur inattendue :\n" + traceback.format_exc()))
        self.q.put(("fin_analyse", None))

    def _retirer(self, st):
        self.states = [s for s in self.states if s is not st]
        self._rendre_fichiers(); self._maj()

    def _nom_sortie_auto(self):
        if not self.states:
            return
        base = os.path.dirname(self.states[0].path)
        self.sortie.set(os.path.join(base, f"rapport_tarl_{self.lang.get()}.docx"))

    def _choisir_sortie(self):
        init = os.path.dirname(self.sortie.get()) or (os.path.dirname(self.states[0].path) if self.states else os.path.expanduser("~"))
        chemin = filedialog.asksaveasfilename(title="Enregistrer le rapport", defaultextension=".docx", initialdir=init,
                                              initialfile=f"rapport_tarl_{self.lang.get()}.docx", filetypes=[("Document Word", "*.docx")])
        if chemin:
            self.sortie.set(chemin)

    def _verifier(self, st):
        ColumnDialog(self, st)

    def _generer(self):
        if not self.states:
            messagebox.showwarning("Aucun fichier", "Importez d'abord un fichier Excel.")
            return
        for st in self.states:
            if not getattr(st, "confirmed", False) and self._a_verifier(st):
                if messagebox.askyesno("Colonnes à vérifier", f"Certaines colonnes de « {st.name} » n'ont pas été vérifiées.\n\n"
                                       "Voulez-vous les vérifier maintenant ? (recommandé)"):
                    self._verifier(st)
                    return
        if not self.sortie.get():
            self._nom_sortie_auto()
        self._sauver()
        self.busy = "generation"
        self.pb.stop(); self.pb.config(mode="determinate", value=5)
        self._maj()
        self._log("— Génération du rapport —")
        args = (list(self.states), self.sortie.get(), self.lang.get(), self.prof.get().strip(), self.directeur.get().strip())
        threading.Thread(target=self._travail, args=args, daemon=True).start()

    def _travail(self, states, sortie, lang, prof, directeur):
        try:
            res = generate_reports(states, sortie, lang, prof, directeur, log=lambda m: self.q.put(("log", "• " + m, None)),
                                   progress=lambda v: self.q.put(("progress", v)))
            self.q.put(("fin", res))
        except ErreurRapport as e:
            self.q.put(("err", str(e)))
        except Exception:  # noqa: BLE001
            self.q.put(("err", "Erreur inattendue :\n" + traceback.format_exc()))

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                k = m[0]
                if k == "log":
                    self._log(m[1], m[2] if len(m) > 2 else None)
                elif k == "progress":
                    self.pb.config(value=m[1])
                elif k == "opened":
                    st = m[1]
                    self.states.append(st)
                    an = st.current
                    self._log(f"✔ {st.name} : {an.n_students} élèves · fichier {self._lang_nom(an.language)} · "
                              f"{sum(1 for c in an.columns if c.status == 'recognized')} colonnes reconnues, "
                              f"{sum(1 for c in an.columns if c.status == 'verify')} à vérifier", "ok")
                    if not self.sortie.get():
                        self._nom_sortie_auto()
                elif k == "fin_analyse":
                    self.pb.stop(); self.pb.config(mode="determinate", value=0)
                    self.busy = ""
                    self._rendre_fichiers(); self._maj()
                elif k == "alerte":
                    messagebox.showwarning("Fichier non reconnu", m[1])
                elif k == "fin":
                    self.busy = ""
                    self.pb.config(value=100)
                    self.dernier = m[1][0]
                    self._maj()
                    self._log("✔ Terminé : " + ", ".join(os.path.basename(p) for p in m[1]), "ok")
                    if messagebox.askyesno("Rapport généré", "Le rapport a été créé avec succès.\n\nVoulez-vous l'ouvrir maintenant ?"):
                        self._ouvrir()
                elif k == "err":
                    self.busy = ""
                    self.pb.config(value=0)
                    self._maj()
                    self._log("✘ " + m[1].splitlines()[0], "err")
                    messagebox.showerror("Erreur", m[1])
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _sauver(self, dossier=None):
        cfg = lire_config()
        cfg.update(theme=self.theme["mode"], prof=self.prof.get().strip(), directeur=self.directeur.get().strip(), lang=self.lang.get())
        if dossier:
            cfg["dossier"] = dossier
        ecrire_config(cfg)

    def _fermer(self):
        self._sauver()
        self.destroy()

    def _ouvrir(self):
        if self.dernier and os.path.exists(self.dernier):
            self._start(self.dernier)

    def _ouvrir_dossier(self):
        if self.dernier:
            self._start(os.path.dirname(os.path.abspath(self.dernier)))

    @staticmethod
    def _start(chemin):
        try:
            if sys.platform.startswith("win"):
                os.startfile(chemin)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", chemin])
            else:
                subprocess.Popen(["xdg-open", chemin])
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erreur", f"Impossible d'ouvrir :\n{chemin}\n{e}")


# ============================================================================================== validation humaine des colonnes
class ColumnDialog(tk.Toplevel):
    """Étape « Vérifier / corriger les colonnes » : colonne d'origine, fonction détectée (liste déroulante modifiable), statut."""

    STATUS = {"recognized": ("Reconnu", "ok"), "verify": ("À vérifier", "warn"), "ignored": ("Ignoré", "muted"), "edited": ("Modifié par vous", "info")}

    def __init__(self, app: App, state: FileState):
        super().__init__(app)
        self.app, self.st, self.P = app, state, app.theme
        P = self.P
        self.title(f"Vérification des colonnes — {state.name}")
        self.configure(bg=P["bg"])
        self.geometry("1000x680")
        self.minsize(860, 520)
        self.transient(app)
        self.grab_set()
        self.labels = [ROLE_LABELS_FR[r] for r in ROLES]
        self.label_to_role = {ROLE_LABELS_FR[r]: r for r in ROLES}
        self._style()
        # --- haut : feuille, barème, résumé
        top = tk.Frame(self, bg=P["bg"])
        top.pack(fill="x", padx=18, pady=(14, 6))
        tk.Label(top, text="Vérifiez ce que le programme a reconnu", bg=P["bg"], fg=P["title"], font=(F, 15, "bold")).pack(anchor="w")
        tk.Label(top, text="Pour chaque colonne de votre fichier, la fonction détectée est affichée. Si elle est fausse, choisissez la bonne dans la liste. "
                           "Vos données Excel ne sont jamais modifiées.", bg=P["bg"], fg=P["muted"], font=(F, 9), wraplength=940, justify="left").pack(anchor="w", pady=(2, 6))
        opt = tk.Frame(top, bg=P["bg"])
        opt.pack(fill="x")
        names = [a.sheet.name for a in state.analyses if a.ok]
        self.sheet_var = tk.StringVar(value=state.current.sheet.name)
        if len(names) > 1:
            tk.Label(opt, text="Feuille Excel :", bg=P["bg"], fg=P["text"], font=(F, 10, "bold")).pack(side="left")
            cb = ttk.Combobox(opt, textvariable=self.sheet_var, values=names, state="readonly", width=26)
            cb.pack(side="left", padx=(6, 18))
            cb.bind("<<ComboboxSelected>>", lambda e: self._changer_feuille())
        tk.Label(opt, text="Barème des notes :", bg=P["bg"], fg=P["text"], font=(F, 10, "bold")).pack(side="left")
        self.scale_var = tk.StringVar(value="Automatique" if not state.scale else f"{state.scale:g}")
        sc = ttk.Combobox(opt, textvariable=self.scale_var, values=["Automatique", "5", "10", "20", "40", "100"], state="readonly", width=12)
        sc.pack(side="left", padx=6)
        sc.bind("<<ComboboxSelected>>", lambda e: self._changer_bareme())
        self.lbl_resume = tk.Label(opt, text="", bg=P["bg"], fg=P["muted"], font=(F, 9, "bold"))
        self.lbl_resume.pack(side="right")
        # --- onglets
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=18, pady=8)
        self.tab_cols = tk.Frame(self.nb, bg=P["card"]); self.tab_prev = tk.Frame(self.nb, bg=P["card"]); self.tab_chk = tk.Frame(self.nb, bg=P["card"])
        self.nb.add(self.tab_cols, text="  1 · Colonnes  ")
        self.nb.add(self.tab_prev, text="  2 · Aperçu des données  ")
        self.nb.add(self.tab_chk, text="  3 · Contrôles  ")
        # --- bas
        bas = tk.Frame(self, bg=P["bg"])
        bas.pack(fill="x", padx=18, pady=(0, 14))
        app._btn(bas, "Réinitialiser (détection automatique)", self._reset, P["b_gray"]).pack(side="left")
        app._btn(bas, "Valider et fermer", self._valider, P["b_primary"], size=11, pady=8).pack(side="right")
        self._rendre()

    def _style(self):
        P, s = self.P, ttk.Style(self)
        s.configure("TCombobox", fieldbackground=P["input"], background=P["card"], foreground=P["text"], arrowcolor=P["text"], bordercolor=P["border"])
        s.map("TCombobox", fieldbackground=[("readonly", P["input"])], foreground=[("readonly", P["text"])],
              selectbackground=[("readonly", P["input"])], selectforeground=[("readonly", P["text"])])
        self.option_add("*TCombobox*Listbox.background", P["input"]); self.option_add("*TCombobox*Listbox.foreground", P["text"])
        s.configure("TNotebook", background=P["bg"], bordercolor=P["border"])
        s.configure("TNotebook.Tab", background=P["chip_bg"], foreground=P["text"], padding=(10, 6), font=(F, 10, "bold"))
        s.map("TNotebook.Tab", background=[("selected", P["card"])], foreground=[("selected", P["accent"])])
        s.configure("Treeview", background=P["card"], fieldbackground=P["card"], foreground=P["text"], rowheight=24, bordercolor=P["border"])
        s.configure("Treeview.Heading", background=P["badge"], foreground="white" if P["mode"] == "light" else "#0B1220", font=(F, 9, "bold"))

    # --- actions
    def _changer_feuille(self):
        try:
            self.st.use_sheet(self.sheet_var.get())
        except ErreurRapport as e:
            messagebox.showwarning("Feuille", str(e), parent=self)
        self._rendre()

    def _changer_bareme(self):
        v = self.scale_var.get()
        self.st.scale = None if v == "Automatique" else float(v)
        self.st.rebuild()
        self._rendre()

    def _reset(self):
        self.st.overrides = {}
        self.st.confirmed = False
        self.st.rebuild()
        self._rendre()

    def _valider(self):
        self.st.confirmed = True
        self.st.rebuild()
        self.app._rendre_fichiers(); self.app._maj()
        self.app._log(f"✔ Colonnes de {self.st.name} validées", "ok")
        self.destroy()

    def _choisir(self, col, label):
        role = self.label_to_role[label]
        if role == col.role and col.idx in self.st.overrides:
            del self.st.overrides[col.idx]
        elif role != col.role:
            self.st.overrides[col.idx] = role
        else:
            self.st.overrides.pop(col.idx, None)
        self.st.rebuild()
        self._rendre_tout()

    # --- rendu
    def _rendre(self):
        self._rendre_tout()

    def _rendre_tout(self):
        self._rendre_colonnes(); self._rendre_apercu(); self._rendre_controles()

    def _vide(self, f):
        for w in f.winfo_children():
            w.destroy()

    def _statut(self, c):
        if c.idx in self.st.overrides:
            return "edited"
        return c.status

    def _rendre_colonnes(self):
        P = self.P
        self._vide(self.tab_cols)
        cols = [c for c in self.st.current.columns if c.profile["filled"] or c.header]
        n_vide = len(self.st.current.columns) - len(cols)
        cnt = {"recognized": 0, "verify": 0, "ignored": 0, "edited": 0}
        for c in cols:
            cnt[self._statut(c)] += 1
        self.lbl_resume.config(text=f"{cnt['recognized']} reconnues · {cnt['verify']} à vérifier · {cnt['ignored']} ignorées · {cnt['edited']} modifiées")
        sf = ScrollFrame(self.tab_cols, P["card"]); sf.pack(fill="both", expand=True)
        g = sf.inner
        heads = ("Colonne dans le fichier Excel", "Exemples de valeurs", "Fonction détectée (modifiable)", "Statut")
        widths = (30, 24, 28, 14)
        for j, (h, w) in enumerate(zip(heads, widths)):
            tk.Label(g, text=h, bg=P["badge"], fg="white" if P["mode"] == "light" else "#0B1220", font=(F, 9, "bold"), anchor="w",
                     padx=8, pady=6).grid(row=0, column=j, sticky="ew")
        g.columnconfigure(0, weight=3); g.columnconfigure(1, weight=2); g.columnconfigure(2, weight=2)
        for i, c in enumerate(cols, 1):
            bg = P["card"] if i % 2 else P["chip_bg"]
            tk.Label(g, text=f"{c.letter} · {c.display}", bg=bg, fg=P["text"], font=(F, 9), anchor="w", justify="left", wraplength=300, padx=8, pady=6).grid(row=i, column=0, sticky="nsew")
            ex = [str(v) for v in c.values if v is not None and str(v).strip()][:3]
            tk.Label(g, text="  ·  ".join(x[:14] for x in ex) or "(vide)", bg=bg, fg=P["muted"], font=(F, 9), anchor="w", padx=8).grid(row=i, column=1, sticky="nsew")
            cell = tk.Frame(g, bg=bg)
            cell.grid(row=i, column=2, sticky="nsew", padx=4, pady=3)
            cb = ttk.Combobox(cell, values=self.labels, state="readonly", width=28)
            cb.set(ROLE_LABELS_FR[self.st.role_of(c)])
            cb.pack(fill="x")
            cb.bind("<<ComboboxSelected>>", lambda e, c=c, cb=cb: self._choisir(c, cb.get()))
            txt, key = self.STATUS[self._statut(c)]
            colr = {"ok": P["ok"], "warn": P["warn"], "muted": P["muted"], "info": P["info"]}[key]
            tk.Label(g, text="●  " + txt, bg=bg, fg=colr, font=(F, 9, "bold"), anchor="w", padx=8).grid(row=i, column=3, sticky="nsew")
        if n_vide:
            tk.Label(g, text=f"{n_vide} colonne(s) totalement vide(s) ignorée(s) automatiquement.", bg=P["card"], fg=P["muted"], font=(F, 9, "italic"),
                     anchor="w", padx=8, pady=8).grid(row=len(cols) + 1, column=0, columnspan=4, sticky="w")

    def _rendre_apercu(self):
        P = self.P
        self._vide(self.tab_prev)
        ds = self.st.dataset
        tk.Label(self.tab_prev, text="Voici comment vos données seront comprises (12 premiers élèves).", bg=P["card"], fg=P["muted"],
                 font=(F, 9, "italic")).pack(anchor="w", padx=10, pady=(8, 4))
        cols = ["Nom"] + (["Identifiant"] if ds.has_id else []) + [it.short for it in ds.items] + (["Moyenne"] if ds.has_average else []) + \
               (["Absent"] if ds.has_absence else [])
        tv = ttk.Treeview(self.tab_prev, columns=cols, show="headings", height=14)
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=170 if c == "Nom" else 90, anchor="w" if c == "Nom" else "center")
        for s in ds.students[:12]:
            vals = [s.name] + ([s.sid] if ds.has_id else []) + ["" if s.scores.get(it.key) is None else f"{s.scores[it.key]:g}" for it in ds.items] + \
                   (["" if s.average_given is None else f"{s.average_given:g}"] if ds.has_average else []) + (["oui" if s.absent else ""] if ds.has_absence else [])
            tv.insert("", "end", values=vals)
        tv.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _rendre_controles(self):
        P = self.P
        self._vide(self.tab_chk)
        iss = self.st.issues
        t = tk.Text(self.tab_chk, bg=P["card"], fg=P["text"], font=(F, 10), relief="flat", wrap="word", padx=14, pady=10)
        t.pack(fill="both", expand=True)
        t.tag_config("err", foreground=P["err"], font=(F, 10, "bold")); t.tag_config("warn", foreground=P["warn"], font=(F, 10, "bold"))
        t.tag_config("info", foreground=P["info"], font=(F, 10, "bold")); t.tag_config("ok", foreground=P["ok"], font=(F, 11, "bold"))
        t.tag_config("sug", foreground=P["muted"], font=(F, 9, "italic"))
        if not iss:
            t.insert("end", "✔ Aucun problème détecté. Vos données sont prêtes pour le traitement.\n", "ok")
        icon = {"error": ("⛔ ", "err"), "warning": ("⚠ ", "warn"), "info": ("ℹ ", "info")}
        for lvl in ("error", "warning", "info"):
            for i in [x for x in iss if x.level == lvl]:
                ic, tag = icon[lvl]
                t.insert("end", ic + i.title + "\n", tag)
                if i.detail:
                    t.insert("end", "   " + i.detail + "\n")
                if i.suggestion:
                    t.insert("end", "   Proposition : " + i.suggestion + "\n", "sug")
                t.insert("end", "\n")
        t.insert("end", "Aucune donnée n'est supprimée ni modifiée silencieusement : les cellules vides ou illisibles sont simplement exclues des calculs concernés.", "sug")
        t.config(state="disabled")


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RapportTARL.App")
    except Exception:
        pass
    try:
        App().mainloop()
    except Exception:  # noqa: BLE001
        with open(os.path.join(os.path.expanduser("~"), "rapport_tarl_erreur.log"), "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        raise
