"""KPI-ORC v2.0 - Design moderne clair"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, datetime, math
from openpyxl import load_workbook

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")

EVENTS = [
    ("Pochon / Fibre",        "ratt_pochon",      "ratt"),
    ("Couture",               "ratt_couture",     "ratt"),
    ("Emballage",             "ratt_emb",         "ratt"),
    ("Presse Souder",         "ratt_presse_soud", "ratt"),
    ("Presse ZIP",            "ratt_presse_zip",  "ratt"),
    ("Chargeuse",             "pb_chargeuse",     "pb"),
    ("Carde",                 "pb_carde",         "pb"),
    ("Etaleur / Tour",        "pb_etaleur",       "pb"),
    ("Coupe / Circ.",         "pb_coupe",         "pb"),
    ("Tapis Bascule",         "pb_tapis1",        "pb"),
    ("Enrouleur Pochon",      "pb_enrouleur",     "pb"),
    ("Pesee / Tapis 2",       "pb_pesee",         "pb"),
    ("Deviation / Table",     "pb_deviation",     "pb"),
    ("Enfileur Pochon",       "pb_enfileur",      "pb"),
    ("Kinna / Stroebel",      "pb_kinna",         "pb"),
    ("Tapeuse",               "pb_tapeuse",       "pb"),
    ("Table Rot. / Twin",     "pb_table_rot",     "pb"),
    ("Enfileuse H100",        "pb_h100",          "pb"),
    ("Enfileuse Traversin",   "pb_traversin",     "pb"),
    ("Presse ORC",            "pb_presse_orc",    "pb"),
    ("Presse Housse ZIP",     "pb_presse_zip2",   "pb"),
    ("Cercleuse",             "pb_cercleuse",     "pb"),
    ("Enrouleuse Traversin",  "pb_enrouleuse",    "pb"),
]


def fmt(seconds):
    h, r = divmod(int(max(0, seconds)), 3600)
    m, s = divmod(r, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def load_cfg():
    try:
        with open(CONFIG_FILE) as f:
            return json.load(f)
    except Exception:
        return {"db_path": ""}


def save_cfg(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f)


# ─── Gauge TRS ───────────────────────────────────────────────────────────────
class Gauge(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._val = 0.0
        self._time = "--:--"
        self.bind("<Configure>", lambda e: self._draw())

    def update_gauge(self, value, time_str=""):
        self._val = max(0.0, min(100.0, float(value)))
        self._time = time_str
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 20 or h < 20:
            return
        cx, cy = w // 2, h - 18
        r = min(cx - 15, cy - 10)
        if r < 20:
            return
        self.create_arc(cx-r, cy-r, cx+r, cy+r,
                        start=0, extent=180, style="arc",
                        outline="#e2e8f0", width=18)
        ext = self._val * 180 / 100
        color = ("#ef4444" if self._val < 55
                 else "#f59e0b" if self._val < 75 else "#22c55e")
        if ext > 0:
            self.create_arc(cx-r, cy-r, cx+r, cy+r,
                            start=180, extent=-ext, style="arc",
                            outline=color, width=18)
        angle = math.radians(180 - self._val * 180 / 100)
        nx = cx + int((r - 6) * math.cos(angle))
        ny = cy - int((r - 6) * math.sin(angle))
        self.create_line(cx, cy, nx, ny, fill="#1e293b", width=2)
        self.create_oval(cx-5, cy-5, cx+5, cy+5, fill="#1e293b", outline="")
        self.create_text(cx, cy - r // 2,
                         text=f"TRS  {self._val:.0f}%",
                         font=("Arial", 14, "bold"), fill=color)
        self.create_text(cx, cy - 6,
                         text=f"a {self._time}",
                         font=("Arial", 9), fill="#64748b")


# ─── Bouton evenement (carre + chrono) ───────────────────────────────────────
class EventCell(tk.Frame):
    C_RATT  = "#f97316"
    C_PB    = "#3b82f6"
    C_RED   = "#dc2626"
    C_FG    = "white"

    def __init__(self, parent, label, key, cat, app, btn_w=88):
        bg = parent.cget("bg")
        super().__init__(parent, bg=bg, padx=3, pady=3)
        self.key  = key
        self.app  = app
        self.idle = self.C_RATT if cat == "ratt" else self.C_PB

        # Carre colore
        self.btn = tk.Frame(self, bg=self.idle,
                            width=btn_w, cursor="hand2")
        self.btn.pack(side="left", fill="y")
        self.btn.pack_propagate(False)

        self.lbl = tk.Label(self.btn, text=label,
                            bg=self.idle, fg=self.C_FG,
                            font=("Arial", 8, "bold"),
                            justify="center", wraplength=btn_w - 10)
        self.lbl.place(relx=0.5, rely=0.5, anchor="center")

        # Chrono a droite
        self.tmr = tk.Label(self, text="00:00:00",
                            bg=bg, fg="#94a3b8",
                            font=("Arial", 13, "bold"), width=8)
        self.tmr.pack(side="left", fill="both", expand=True, padx=(6, 0))

        for w in (self.btn, self.lbl):
            w.bind("<Button-1>", lambda e: self._toggle())

    def _toggle(self):
        if self.app._t_running(self.key):
            self.app._t_stop(self.key)
            self._color(self.idle)
        else:
            self.app._t_start(self.key)
            self._color(self.C_RED)

    def _color(self, c):
        self.btn.config(bg=c)
        self.lbl.config(bg=c)

    def refresh(self):
        s       = self.app._t_get(self.key)
        running = self.app._t_running(self.key)
        self.tmr.config(text=fmt(s))
        if running:
            self._color(self.C_RED)
            self.tmr.config(fg="#dc2626", font=("Arial", 15, "bold"))
        else:
            self._color(self.idle)
            if s > 0:
                self.tmr.config(fg="#f59e0b", font=("Arial", 13, "bold"))
            else:
                self.tmr.config(fg="#94a3b8", font=("Arial", 13, "bold"))


# ─── Application ─────────────────────────────────────────────────────────────
class App:
    BG    = "#f1f5f9"
    WHITE = "#ffffff"
    DARK  = "#0f172a"
    BLUE  = "#3b82f6"
    GREEN = "#16a34a"
    RED   = "#dc2626"
    GRAY  = "#64748b"
    LGRAY = "#e2e8f0"

    def __init__(self, root):
        self.root = root
        self.root.title("KPI-ORC | Ligne ORC1")
        self.root.configure(bg=self.BG)
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-fullscreen", True)

        self.cfg          = load_cfg()
        self.lists        = {}
        self._timers      = {}
        self._of_start    = None
        self._prod_active = False
        self._after_id    = None
        self._db_labels   = []
        self._cells       = []

        self._load_lists()
        self._show_main()

    # ── Helpers config ────────────────────────────────────────────────────────
    def _load_lists(self):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb["Listes"]
            headers = {}
            for cell in next(ws.iter_rows(max_row=1)):
                if cell.value:
                    headers[cell.column] = str(cell.value)
            self.lists = {h: [] for h in headers.values()}
            for row in ws.iter_rows(min_row=2, values_only=True):
                for ci, val in enumerate(row, 1):
                    if ci in headers and val is not None:
                        self.lists[headers[ci]].append(str(val))
            wb.close()
        except Exception:
            pass

    def _select_db(self):
        p = filedialog.askopenfilename(
            title="Selectionner la Base de Donnees",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tous", "*.*")])
        if p:
            self.cfg["db_path"] = p
            save_cfg(self.cfg)
            self._load_lists()
            name = os.path.basename(p)
            for lbl in self._db_labels:
                try:
                    lbl.config(text=f"DB: {name}")
                except Exception:
                    pass
            messagebox.showinfo("Succes", f"Connecte :\n{name}")

    def _db_widget(self, parent, bg):
        f = tk.Frame(parent, bg=bg)
        name = os.path.basename(self.cfg.get("db_path", "")) or "Non connectee"
        lbl = tk.Label(f, text=f"DB: {name}", bg=bg,
                       fg=self.GRAY, font=("Arial", 9))
        lbl.pack(side="left", padx=6)
        self._db_labels.append(lbl)
        tk.Button(f, text="Database", command=self._select_db,
                  bg=self.LGRAY, fg=self.DARK,
                  font=("Arial", 9, "bold"),
                  relief="flat", padx=10, pady=4,
                  cursor="hand2").pack(side="left")
        return f

    def _get_list(self, h):
        return self.lists.get(h, [])

    # ── Timers ────────────────────────────────────────────────────────────────
    def _t_start(self, key):
        t = self._timers.setdefault(
            key, {"start": None, "elapsed": 0.0, "running": False})
        if not t["running"]:
            t["start"]   = datetime.datetime.now()
            t["running"] = True

    def _t_stop(self, key):
        t = self._timers.get(key)
        if t and t["running"]:
            t["elapsed"] += (datetime.datetime.now() - t["start"]).total_seconds()
            t["running"]  = False
            t["start"]    = None

    def _t_get(self, key):
        t = self._timers.get(key, {"elapsed": 0.0, "running": False, "start": None})
        total = t["elapsed"]
        if t["running"]:
            total += (datetime.datetime.now() - t["start"]).total_seconds()
        return total

    def _t_running(self, key):
        return self._timers.get(key, {}).get("running", False)

    def _t_stop_all(self):
        for k in list(self._timers):
            self._t_stop(k)

    def _t_reset(self):
        self._timers.clear()

    def _clear(self):
        if self._after_id:
            self.root.after_cancel(self._after_id)
            self._after_id = None
        for w in self.root.winfo_children():
            w.destroy()

    # =========================================================================
    #  ECRAN PRINCIPAL
    # =========================================================================
    def _show_main(self):
        self._prod_active = False
        self._cells       = []
        self._clear()
        self._db_labels.clear()

        root = tk.Frame(self.root, bg=self.BG)
        root.pack(fill="both", expand=True)

        # ── En-tete
        hdr = tk.Frame(root, bg=self.WHITE, height=70)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="KPI-ORC", bg=self.WHITE, fg=self.BLUE,
                 font=("Arial", 24, "bold")).pack(side="left", padx=20)
        tk.Label(hdr, text="Ligne ORC1", bg=self.WHITE, fg=self.GRAY,
                 font=("Arial", 14)).pack(side="left")
        self._db_widget(hdr, self.WHITE).pack(side="right", padx=20)

        # Separateur
        tk.Frame(root, bg=self.LGRAY, height=2).pack(fill="x")

        # ── Contenu
        body = tk.Frame(root, bg=self.BG)
        body.pack(fill="both", expand=True, padx=24, pady=16)

        # KPI + Bouton cote a cote
        top = tk.Frame(body, bg=self.BG)
        top.pack(fill="x", pady=(0, 16))

        # Carte TRS
        kpi_card = tk.Frame(top, bg=self.WHITE,
                            relief="flat", bd=0,
                            highlightbackground=self.LGRAY,
                            highlightthickness=1)
        kpi_card.pack(side="left", padx=(0, 16))

        tk.Label(kpi_card, text="TRS",
                 bg=self.WHITE, fg=self.GRAY,
                 font=("Arial", 11, "bold")).pack(pady=(10, 0))
        tk.Label(kpi_card, text="Taux de Rendement Synthetique",
                 bg=self.WHITE, fg=self.GRAY,
                 font=("Arial", 9)).pack()
        self._main_gauge = Gauge(kpi_card, bg=self.WHITE,
                                  width=300, height=130,
                                  highlightthickness=0)
        self._main_gauge.pack(padx=20, pady=(0, 10))
        self._refresh_main_kpi()

        # Bouton demarrer
        btn_frame = tk.Frame(top, bg=self.BG)
        btn_frame.pack(side="left", fill="both", expand=True)

        start_btn = tk.Button(btn_frame,
                              text="▶\nDEMARRER\nUNE PRODUCTION",
                              command=self._start_production,
                              bg=self.GREEN, fg="white",
                              font=("Arial", 18, "bold"),
                              relief="flat", cursor="hand2",
                              activebackground="#15803d",
                              activeforeground="white")
        start_btn.pack(fill="both", expand=True, ipady=20)

        # ── Tableau recapitulatif
        tk.Label(body, text="15 Dernieres Declarations",
                 bg=self.BG, fg=self.DARK,
                 font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 6))

        tf = tk.Frame(body, bg=self.WHITE,
                      highlightbackground=self.LGRAY,
                      highlightthickness=1)
        tf.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Light.Treeview",
                        background=self.WHITE,
                        foreground=self.DARK,
                        fieldbackground=self.WHITE,
                        rowheight=28,
                        font=("Arial", 10))
        style.configure("Light.Treeview.Heading",
                        background=self.LGRAY,
                        foreground=self.DARK,
                        font=("Arial", 10, "bold"),
                        relief="flat")
        style.map("Light.Treeview",
                  background=[("selected", "#dbeafe")])

        cols = ("Date", "OF", "Pilote", "Poste",
                "Duree OF", "PB Techniques", "Rattrapages")
        tree = ttk.Treeview(tf, columns=cols, show="headings",
                            height=14, style="Light.Treeview")
        widths = {"Date": 100, "OF": 110, "Pilote": 180, "Poste": 110,
                  "Duree OF": 95, "PB Techniques": 120, "Rattrapages": 120}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 110), anchor="center")

        sb = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._load_table(tree)

    def _refresh_main_kpi(self):
        path      = self.cfg.get("db_path", "")
        last_time = "--:--"
        trs       = 0.0
        if path and os.path.exists(path):
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                ws = wb["Data"]
                rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                        if any(r)]
                wb.close()
                if rows:
                    last      = list(rows[-1]) + [None] * 55
                    last_time = str(last[18])[:5] if last[18] else "--:--"
            except Exception:
                pass
        self._main_gauge.update_gauge(trs, last_time)

    def _load_table(self, tree):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb["Data"]
            rows = [list(r) + [None]*55
                    for r in ws.iter_rows(min_row=2, values_only=True)
                    if any(r)]
            wb.close()
        except Exception:
            return

        def _sum_dur(row, indices):
            total = 0
            for i in indices:
                if i < len(row) and row[i]:
                    parts = str(row[i]).split(":")
                    try:
                        if len(parts) == 3:
                            total += int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
                    except Exception:
                        pass
            return total

        for row in list(reversed(rows))[:15]:
            tree.insert("", "end", values=(
                str(row[1])[:10] if row[1] else "",
                str(row[0])  if row[0] else "",
                str(row[3])  if row[3] else "",
                str(row[2])  if row[2] else "",
                str(row[16]) if row[16] else "",
                fmt(_sum_dur(row, range(31, 49))),
                fmt(_sum_dur(row, range(26, 31))),
            ))

    # =========================================================================
    #  ECRAN DE PRODUCTION
    # =========================================================================
    def _start_production(self):
        self._t_reset()
        self._of_start    = datetime.datetime.now()
        self._prod_active = True
        self._cells       = []
        self._show_production()

    def _show_production(self):
        self._clear()
        self._db_labels.clear()

        outer = tk.Frame(self.root, bg=self.BG)
        outer.pack(fill="both", expand=True)

        # ── En-tete sombre
        hdr = tk.Frame(outer, bg=self.DARK, height=72)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="ORC1 — PRODUCTION EN COURS",
                 bg=self.DARK, fg="white",
                 font=("Arial", 14, "bold")).pack(side="left", padx=20)

        # Heure debut
        t_start = self._of_start.strftime("%H:%M:%S")
        tk.Label(hdr, text=f"Debut : {t_start}",
                 bg=self.DARK, fg="#94a3b8",
                 font=("Arial", 12)).pack(side="left", padx=20)

        # Duree dynamique (gros compteur)
        self._of_clk = tk.Label(hdr, text="00:00:00",
                                 bg=self.DARK, fg="#4ade80",
                                 font=("Arial", 26, "bold"))
        self._of_clk.pack(side="left", padx=20)

        self._db_widget(hdr, self.DARK).pack(side="right", padx=16)

        # ── Bouton FIN (bas, toujours visible)
        end_bar = tk.Frame(outer, bg="#7f1d1d", height=54)
        end_bar.pack(fill="x", side="bottom")
        end_bar.pack_propagate(False)
        tk.Button(end_bar,
                  text="⏹   DECLARER LA FIN DE PRODUCTION",
                  command=self._end_production,
                  bg="#7f1d1d", fg="white",
                  font=("Arial", 14, "bold"),
                  relief="flat", cursor="hand2",
                  activebackground="#991b1b",
                  activeforeground="white"
                  ).pack(fill="both", expand=True)

        # ── Corps : formulaire gauche | evenements droite
        body = tk.Frame(outer, bg=self.BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=self.BG, width=480)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self._build_form(left)

        tk.Frame(body, bg=self.LGRAY, width=2).pack(side="left", fill="y")

        right = tk.Frame(body, bg=self.BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_events(right)

        self._tick()

    # ── Formulaire gauche ─────────────────────────────────────────────────────
    def _build_form(self, parent):
        cnv   = tk.Canvas(parent, bg=self.BG, highlightthickness=0)
        sb    = ttk.Scrollbar(parent, orient="vertical", command=cnv.yview)
        cnv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cnv.pack(fill="both", expand=True)

        frame  = tk.Frame(cnv, bg=self.BG)
        win_id = cnv.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>",
                   lambda e: cnv.configure(scrollregion=cnv.bbox("all")))
        cnv.bind("<Configure>",
                 lambda e: cnv.itemconfig(win_id, width=e.width))
        cnv.bind_all("<MouseWheel>",
                     lambda e: cnv.yview_scroll(-1*(e.delta//120), "units"))

        pad = {"padx": 12, "pady": 4}

        def section(text):
            tk.Label(frame, text=text, bg=self.BG, fg=self.BLUE,
                     font=("Arial", 10, "bold")).pack(anchor="w", **pad)

        def field(lbl, key, ftype, lh=None):
            row = tk.Frame(frame, bg=self.BG)
            row.pack(fill="x", padx=12, pady=2)
            tk.Label(row, text=lbl, bg=self.BG, fg=self.GRAY,
                     font=("Arial", 9), width=18,
                     anchor="w").pack(side="left")
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                w = tk.Entry(row, textvariable=var,
                             bg=self.WHITE, fg=self.DARK,
                             font=("Arial", 10), relief="solid",
                             bd=1, insertbackground=self.DARK)
                w.pack(side="left", fill="x", expand=True)
            else:
                vals = self._get_list(lh) if lh else []
                w = ttk.Combobox(row, textvariable=var, values=vals,
                                 font=("Arial", 10), state="readonly")
                w.pack(side="left", fill="x", expand=True)

        self.fv = {}

        # Titre section
        tk.Label(frame, text="DONNEES DE L'OF",
                 bg=self.BG, fg=self.DARK,
                 font=("Arial", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 4))

        section("Identification")
        field("N° OF *",          "of_num",     "entry")
        field("Pilote *",         "pilote",     "combo", "Pilotes")
        field("Co-Pilote",        "copilote",   "combo", "Co-Pilotes")
        field("Poste *",          "poste",      "combo", "Postes")
        field("Nb personnes",     "nb_pers",    "combo", "Nombre operateur")

        section("Produit")
        field("Taille produit",   "taille",     "combo", "Taille produit")
        field("Type produit",     "type_prod",  "combo", "Type produit")
        field("Code produit *",   "code_prod",  "entry")
        field("Poids garnissage", "poids",      "entry")
        field("Fibre",            "fibre",      "combo", "Fibre")

        section("Quantites")
        field("Qte fabriquee *",  "qte_fab",    "entry")
        field("Qte emballee",     "qte_emb",    "entry")

        section("Taie / Qualite")
        field("OF taie",          "of_taie",    "entry")
        field("Traca fibre",      "traca",      "entry")
        field("Ref. taie",        "ref_taie",   "entry")
        field("Nb taie 2nd choix","nb_taie2",   "entry")
        field("Nb defaut couture","nb_def_cout","entry")
        field("Manquant taie",    "mq_taie",    "entry")
        field("Manquant housse",  "mq_housse",  "entry")
        field("Manquant encart",  "mq_encart",  "entry")

        # Kit checkbox
        self._v_kit = tk.BooleanVar()
        kit_row = tk.Frame(frame, bg=self.BG)
        kit_row.pack(fill="x", padx=12, pady=6)
        tk.Checkbutton(kit_row, text="KIT de 2 pieces",
                       variable=self._v_kit,
                       bg=self.BG, fg=self.DARK,
                       selectcolor=self.WHITE,
                       activebackground=self.BG,
                       font=("Arial", 11, "bold"),
                       cursor="hand2").pack(side="left")

        # Commentaire (grand)
        section("Commentaire")
        self.fv["comment"] = tk.StringVar()
        txt_frame = tk.Frame(frame, bg=self.BG)
        txt_frame.pack(fill="x", padx=12, pady=2)
        self._comment_txt = tk.Text(txt_frame, height=5,
                                     bg=self.WHITE, fg=self.DARK,
                                     font=("Arial", 10), relief="solid",
                                     bd=1, wrap="word",
                                     insertbackground=self.DARK)
        self._comment_txt.pack(fill="x")

        tk.Frame(frame, bg=self.BG, height=20).pack()

    # ── Paneau evenements droite ───────────────────────────────────────────────
    def _build_events(self, parent):
        bg = self.BG

        # ─ Titre RATTRAPAGES
        r_hdr = tk.Frame(parent, bg="#fff7ed", height=34)
        r_hdr.pack(fill="x", padx=10, pady=(8, 2))
        r_hdr.pack_propagate(False)
        tk.Label(r_hdr, text="ARRETS RATTRAPAGE",
                 bg="#fff7ed", fg="#c2410c",
                 font=("Arial", 11, "bold")).pack(side="left", padx=10,
                                                   pady=4)

        # Grid rattrapages : 5 en 1 ligne
        ratt_grid = tk.Frame(parent, bg=bg)
        ratt_grid.pack(fill="x", padx=10, pady=(0, 6))
        for c in range(5):
            ratt_grid.columnconfigure(c, weight=1)
        ratt_grid.rowconfigure(0, weight=1)

        for i, (label, key, cat) in enumerate(EVENTS[:5]):
            cell = EventCell(ratt_grid, label, key, cat, self, btn_w=80)
            cell.grid(row=0, column=i, sticky="nsew", padx=3, pady=3)
            self._cells.append(cell)

        # ─ Titre PB TECHNIQUES
        pb_hdr = tk.Frame(parent, bg="#eff6ff", height=34)
        pb_hdr.pack(fill="x", padx=10, pady=(4, 2))
        pb_hdr.pack_propagate(False)
        tk.Label(pb_hdr, text="PROBLEMES TECHNIQUES",
                 bg="#eff6ff", fg="#1d4ed8",
                 font=("Arial", 11, "bold")).pack(side="left", padx=10,
                                                   pady=4)

        # Grid PB : 6 colonnes x 3 lignes = 18
        pb_grid = tk.Frame(parent, bg=bg)
        pb_grid.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        for c in range(6):
            pb_grid.columnconfigure(c, weight=1)
        for r in range(3):
            pb_grid.rowconfigure(r, weight=1)

        for i, (label, key, cat) in enumerate(EVENTS[5:]):
            row_i = i // 6
            col_i = i % 6
            cell = EventCell(pb_grid, label, key, cat, self, btn_w=74)
            cell.grid(row=row_i, column=col_i, sticky="nsew",
                      padx=3, pady=3)
            self._cells.append(cell)

    # ── Tick horloge ─────────────────────────────────────────────────────────
    def _tick(self):
        if not self._prod_active:
            return
        of_s = (datetime.datetime.now() - self._of_start).total_seconds()
        self._of_clk.config(text=fmt(of_s))
        for cell in self._cells:
            cell.refresh()
        self._after_id = self.root.after(1000, self._tick)

    # =========================================================================
    #  FIN DE PRODUCTION -> EXCEL
    # =========================================================================
    def _end_production(self):
        end_dt = datetime.datetime.now()
        self._t_stop_all()
        self._prod_active = False

        of_s = (end_dt - self._of_start).total_seconds()
        v    = {k: var.get().strip() for k, var in self.fv.items()}
        v["comment"] = self._comment_txt.get("1.0", "end").strip()

        if not v.get("of_num"):
            if not messagebox.askyesno(
                    "Attention", "N° OF non saisi. Continuer quand meme ?"):
                self._prod_active = True
                self._tick()
                return

        def _n(k):
            try:
                return int(v.get(k, 0) or 0)
            except Exception:
                return 0

        qte_fab = _n("qte_fab")
        nb_pers = max(1, _n("nb_pers") or 1)
        of_min  = of_s / 60
        of_hrs  = of_s / 3600
        c1      = round(qte_fab / of_min,  2) if of_min  > 0 else 0
        c2      = round(qte_fab / (nb_pers * of_hrs), 2) if of_hrs > 0 else 0
        equiv   = self._calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
        kit     = 2 if self._v_kit.get() else 1

        def _ts(key):
            return fmt(self._t_get(key))

        row = [
            v.get("of_num",""),                                  # A
            datetime.date.today().strftime("%d/%m/%Y"),          # B
            v.get("poste",""),                                   # C
            v.get("pilote",""),                                  # D
            v.get("copilote",""),                                # E
            v.get("nb_pers",""),                                 # F
            v.get("taille",""),                                  # G
            v.get("code_prod",""),                               # H
            v.get("type_prod",""),                               # I
            v.get("poids",""),                                   # J
            v.get("fibre",""),                                   # K
            v.get("of_taie",""),                                 # L
            v.get("traca",""),                                   # M
            qte_fab,                                             # N
            _n("qte_emb"),                                       # O
            equiv,                                               # P
            fmt(of_s),                                           # Q
            self._of_start.strftime("%H:%M:%S"),                # R
            end_dt.strftime("%H:%M:%S"),                        # S
            c1,                                                  # T
            c2,                                                  # U
            kit,                                                 # V
            v.get("ref_taie",""),                               # W
            _n("nb_def_cout"),                                   # X
            _n("mq_taie"),                                       # Y
            f"Housse:{v.get('mq_housse','')} Encart:{v.get('mq_encart','')}", # Z
            _ts("ratt_pochon"),    # AA
            _ts("ratt_couture"),   # AB
            _ts("ratt_emb"),       # AC
            _ts("ratt_presse_soud"),# AD
            _ts("ratt_presse_zip"),# AE
            _ts("pb_chargeuse"),   # AF
            _ts("pb_carde"),       # AG
            _ts("pb_etaleur"),     # AH
            _ts("pb_coupe"),       # AI
            _ts("pb_tapis1"),      # AJ
            _ts("pb_enrouleur"),   # AK
            _ts("pb_pesee"),       # AL
            _ts("pb_deviation"),   # AM
            _ts("pb_enfileur"),    # AN
            _ts("pb_kinna"),       # AO
            _ts("pb_tapeuse"),     # AP
            _ts("pb_table_rot"),   # AQ
            _ts("pb_h100"),        # AR
            _ts("pb_traversin"),   # AS
            _ts("pb_presse_orc"),  # AT
            _ts("pb_presse_zip2"), # AU
            _ts("pb_cercleuse"),   # AV
            _ts("pb_enrouleuse"),  # AW
            v.get("comment",""),   # AX
        ]

        ok = self._write_excel(row)
        if ok:
            messagebox.showinfo("Succes", "Production declaree !")
        else:
            messagebox.showerror("Erreur",
                                 "Impossible d'ecrire dans Excel.\n"
                                 "Verifiez que le fichier n'est pas ouvert.")
        self._show_main()

    def _calc_equiv(self, qte, taille, type_prod):
        for item in self._get_list("Equivalence"):
            try:
                return round(qte * float(str(item).replace(",", ".")), 2)
            except Exception:
                pass
        return qte

    def _write_excel(self, row):
        path = self.cfg.get("db_path", "")
        if not path:
            messagebox.showwarning("Attention", "Aucune base de donnees !")
            return False
        try:
            wb = load_workbook(path)
            wb["Data"].append(row)
            wb.save(path)
            wb.close()
            return True
        except PermissionError:
            return False
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur Excel :\n{e}")
            return False


if __name__ == "__main__":
    root = tk.Tk()
    root.configure(bg="#f1f5f9")
    try:
        root.state("zoomed")
    except Exception:
        root.attributes("-fullscreen", True)
    App(root)
    root.mainloop()
