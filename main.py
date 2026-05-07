"""KPI-ORC v1.0 - Systeme de declaration de production ORC1"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, datetime, math
from openpyxl import load_workbook

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")

EVENTS = [
    ("Rattrapage Pochon/fibre",               "ratt_pochon",      "ratt"),
    ("Rattrapage Couture",                     "ratt_couture",     "ratt"),
    ("Rattrapage Emballage",                   "ratt_emb",         "ratt"),
    ("Rattrapage Presse a souder",             "ratt_presse_soud", "ratt"),
    ("Rattrapage Presse a housse ZIP",         "ratt_presse_zip",  "ratt"),
    ("PB Technique : Chargeuse",               "pb_chargeuse",     "pb"),
    ("PB Technique : Carde",                   "pb_carde",         "pb"),
    ("PB Technique : Etaleur/Tour",            "pb_etaleur",       "pb"),
    ("PB Technique : Coupe/Coupe circulaire",  "pb_coupe",         "pb"),
    ("PB Technique : Tapis bascule/Tapis pese 1", "pb_tapis1",     "pb"),
    ("PB Technique : Enrouleur pochon",        "pb_enrouleur",     "pb"),
    ("PB Technique : Pesee/Tapis pesee 2",     "pb_pesee",         "pb"),
    ("PB Technique : Deviation/Table distribution", "pb_deviation","pb"),
    ("PB Technique : Enfileur pochon",         "pb_enfileur",      "pb"),
    ("PB Technique : Kinna/Stroebel",          "pb_kinna",         "pb"),
    ("PB Technique : Tapeuse",                 "pb_tapeuse",       "pb"),
    ("PB Technique : Table rotative - Twin pack","pb_table_rot",   "pb"),
    ("PB Technique : Enfileuse H100",          "pb_h100",          "pb"),
    ("PB Technique : Enfileuse traversin",     "pb_traversin",     "pb"),
    ("PB Technique : Presse ORC compressees et soudees","pb_presse_orc","pb"),
    ("PB Technique : Presse a housse ZIP",     "pb_presse_zip2",   "pb"),
    ("PB Technique : Cercleuse",               "pb_cercleuse",     "pb"),
    ("PB Technique : Enrouleuse traversin",    "pb_enrouleuse",    "pb"),
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
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 20 or h < 20:
            return
        cx = w // 2
        cy = h - 15
        r = min(cx - 15, cy - 10)
        if r < 20:
            return
        self.create_arc(cx-r, cy-r, cx+r, cy+r,
                        start=0, extent=180, style="arc",
                        outline="#2c3e50", width=18)
        ext = self._val * 180 / 100
        color = ("#e74c3c" if self._val < 55
                 else "#f39c12" if self._val < 75
                 else "#27ae60")
        if ext > 0:
            self.create_arc(cx-r, cy-r, cx+r, cy+r,
                            start=180, extent=-ext, style="arc",
                            outline=color, width=18)
        angle = math.radians(180 - self._val * 180 / 100)
        nx = cx + int((r - 6) * math.cos(angle))
        ny = cy - int((r - 6) * math.sin(angle))
        self.create_line(cx, cy, nx, ny, fill="white", width=2)
        self.create_oval(cx-4, cy-4, cx+4, cy+4, fill="white", outline="")
        self.create_text(cx, cy - r // 2,
                         text=f"TRS  {self._val:.0f}%",
                         font=("Arial", 11, "bold"), fill=color)
        self.create_text(cx, cy - 5,
                         text=self._time, font=("Arial", 9), fill="#aaa")


class App:
    BG     = "#16213e"
    HDR    = "#1a1a2e"
    CARD   = "#0f3460"
    GREEN  = "#27ae60"
    RED    = "#e74c3c"
    ORANGE = "#f39c12"
    FG     = "white"
    GRAY   = "#95a5a6"

    def __init__(self, root):
        self.root = root
        self.root.title("KPI-ORC | Ligne ORC1")
        self.root.configure(bg=self.BG)
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-fullscreen", True)
        self.cfg = load_cfg()
        self.lists = {}
        self._timers = {}
        self._of_start = None
        self._prod_active = False
        self._after_id = None
        self._db_labels = []
        self._load_lists()
        self._show_main()

    # ── Config & Excel ────────────────────────────────────────────────────────
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
            messagebox.showinfo("Succes", f"Base de donnees connectee :\n{name}")

    def _db_widget(self, parent):
        f = tk.Frame(parent, bg=self.HDR)
        name = os.path.basename(self.cfg.get("db_path", "")) or "Non connectee"
        lbl = tk.Label(f, text=f"DB: {name}", bg=self.HDR, fg=self.GRAY,
                       font=("Arial", 9))
        lbl.pack(side="left", padx=6)
        self._db_labels.append(lbl)
        tk.Button(f, text="Database", command=self._select_db,
                  bg="#2c3e50", fg=self.FG, font=("Arial", 10, "bold"),
                  relief="raised", padx=8, pady=4,
                  cursor="hand2").pack(side="left")
        return f

    def _get_list(self, h):
        return self.lists.get(h, [])

    # ── Timers ────────────────────────────────────────────────────────────────
    def _t_start(self, key):
        t = self._timers.setdefault(
            key, {"start": None, "elapsed": 0.0, "running": False})
        if not t["running"]:
            t["start"] = datetime.datetime.now()
            t["running"] = True

    def _t_stop(self, key):
        t = self._timers.get(key)
        if t and t["running"]:
            t["elapsed"] += (
                datetime.datetime.now() - t["start"]).total_seconds()
            t["running"] = False
            t["start"] = None

    def _t_get(self, key):
        t = self._timers.get(
            key, {"elapsed": 0.0, "running": False, "start": None})
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

    # ── Navigation ────────────────────────────────────────────────────────────
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
        self._clear()
        self._db_labels.clear()

        main = tk.Frame(self.root, bg=self.BG)
        main.pack(fill="both", expand=True)

        hdr = tk.Frame(main, bg=self.HDR, height=65)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="KPI-ORC  |  Ligne ORC1",
                 bg=self.HDR, fg=self.FG,
                 font=("Arial", 20, "bold")).pack(side="left", padx=20)
        self._db_widget(hdr).pack(side="right", padx=15)

        kpi = tk.Frame(main, bg=self.CARD, height=150)
        kpi.pack(fill="x", padx=20, pady=(12, 6))
        kpi.pack_propagate(False)
        tk.Label(kpi, text="TRS - Taux de Rendement Synthetique",
                 bg=self.CARD, fg=self.ORANGE,
                 font=("Arial", 13, "bold")).pack(pady=(8, 0))
        self._main_gauge = Gauge(kpi, bg=self.CARD, width=280, height=108,
                                  highlightthickness=0)
        self._main_gauge.pack()
        self._refresh_main_kpi()

        tk.Button(main, text="▶   DEMARRER UNE PRODUCTION",
                  command=self._start_production,
                  bg=self.GREEN, fg=self.FG, font=("Arial", 20, "bold"),
                  relief="raised", padx=40, pady=18,
                  cursor="hand2").pack(pady=18)

        tk.Label(main, text="15 Dernieres Declarations",
                 bg=self.BG, fg=self.FG,
                 font=("Arial", 12, "bold")).pack(padx=20, anchor="w")

        tf = tk.Frame(main, bg=self.BG)
        tf.pack(fill="both", expand=True, padx=20, pady=(4, 12))

        style = ttk.Style()
        style.configure("Dark.Treeview",
                        background=self.HDR, foreground=self.FG,
                        fieldbackground=self.HDR, rowheight=26,
                        font=("Arial", 10))
        style.configure("Dark.Treeview.Heading",
                        background=self.CARD, foreground=self.ORANGE,
                        font=("Arial", 10, "bold"))
        style.map("Dark.Treeview", background=[("selected", "#2980b9")])

        cols = ("Date", "OF", "Pilote", "Poste",
                "Duree OF", "PB Techniques", "Rattrapages")
        tree = ttk.Treeview(tf, columns=cols, show="headings",
                            height=14, style="Dark.Treeview")
        widths = {"Date": 100, "OF": 110, "Pilote": 160, "Poste": 110,
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
        path = self.cfg.get("db_path", "")
        last_time = "--:--"
        trs = 0.0
        if path and os.path.exists(path):
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                ws = wb["Data"]
                rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                        if any(r)]
                wb.close()
                if rows:
                    last = list(rows[-1]) + [None] * 55
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
                            total += (int(parts[0])*3600
                                      + int(parts[1])*60 + int(parts[2]))
                    except Exception:
                        pass
            return total

        for row in list(reversed(rows))[:15]:
            tree.insert("", "end", values=(
                str(row[1])[:10] if row[1] else "",
                str(row[0]) if row[0] else "",
                str(row[3]) if row[3] else "",
                str(row[2]) if row[2] else "",
                str(row[16]) if row[16] else "",
                fmt(_sum_dur(row, range(31, 49))),
                fmt(_sum_dur(row, range(26, 31))),
            ))

    # =========================================================================
    #  ECRAN DE PRODUCTION
    # =========================================================================
    def _start_production(self):
        self._t_reset()
        self._of_start = datetime.datetime.now()
        self._prod_active = True
        self._show_production()

    def _show_production(self):
        self._clear()
        self._db_labels.clear()

        outer = tk.Frame(self.root, bg=self.BG)
        outer.pack(fill="both", expand=True)

        hdr = tk.Frame(outer, bg=self.HDR, height=65)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="KPI-ORC  |  PRODUCTION EN COURS  ORC1",
                 bg=self.HDR, fg=self.ORANGE,
                 font=("Arial", 15, "bold")).pack(side="left", padx=20)
        self._of_clk = tk.Label(hdr, text="⏱  00:00:00",
                                 bg=self.HDR, fg="#2ecc71",
                                 font=("Arial", 16, "bold"))
        self._of_clk.pack(side="left", padx=20)
        self._prod_gauge = Gauge(hdr, bg=self.HDR, width=160, height=60,
                                  highlightthickness=0)
        self._prod_gauge.pack(side="right", padx=6)
        self._db_widget(hdr).pack(side="right", padx=12)

        end_bar = tk.Frame(outer, bg="#7b1414", height=58)
        end_bar.pack(fill="x", side="bottom")
        end_bar.pack_propagate(False)
        tk.Button(end_bar,
                  text="⏹   DECLARER LA FIN DE PRODUCTION",
                  command=self._end_production,
                  bg="#7b1414", fg=self.FG, font=("Arial", 15, "bold"),
                  relief="flat", cursor="hand2").pack(fill="both", expand=True)

        body = tk.Frame(outer, bg=self.BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=self.BG, width=560)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self._build_form(left)

        right = tk.Frame(body, bg=self.BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_events(right)

        self._tick()

    def _build_form(self, parent):
        cnv = tk.Canvas(parent, bg=self.BG, highlightthickness=0)
        sb  = ttk.Scrollbar(parent, orient="vertical", command=cnv.yview)
        cnv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cnv.pack(fill="both", expand=True)

        frame = tk.Frame(cnv, bg=self.BG)
        win_id = cnv.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>",
                   lambda e: cnv.configure(scrollregion=cnv.bbox("all")))
        cnv.bind("<Configure>",
                 lambda e: cnv.itemconfig(win_id, width=e.width))
        cnv.bind_all("<MouseWheel>",
                     lambda e: cnv.yview_scroll(-1*(e.delta//120), "units"))

        card = tk.Frame(frame, bg=self.CARD, relief="ridge", bd=2)
        card.pack(fill="x", padx=8, pady=8)

        tk.Label(card, text="DONNEES DE L'OF",
                 bg=self.CARD, fg=self.ORANGE,
                 font=("Arial", 13, "bold")).grid(
            row=0, column=0, columnspan=4, pady=(10, 6), padx=10)

        auto = tk.Frame(card, bg=self.CARD)
        auto.grid(row=1, column=0, columnspan=4,
                  padx=10, pady=4, sticky="ew")

        self._v_date  = tk.StringVar(
            value=datetime.date.today().strftime("%d/%m/%Y"))
        self._v_start = tk.StringVar(
            value=self._of_start.strftime("%H:%M:%S"))
        self._v_end   = tk.StringVar(value="--:--:--")
        self._v_duree = tk.StringVar(value="00:00:00")
        self._v_equiv = tk.StringVar(value="0")
        self._v_cad1  = tk.StringVar(value="0")
        self._v_cad2  = tk.StringVar(value="0")

        auto_rows = [
            ("Date",            self._v_date),
            ("Heure debut",     self._v_start),
            ("Heure fin",       self._v_end),
            ("Duree OF",        self._v_duree),
            ("Equivalence",     self._v_equiv),
            ("Cadence pcs/min", self._v_cad1),
            ("Cad. pcs/pers/h", self._v_cad2),
        ]
        for i, (lbl, var) in enumerate(auto_rows):
            c = (i % 2) * 2
            r = i // 2
            tk.Label(auto, text=f"{lbl}:", bg=self.CARD, fg=self.GRAY,
                     font=("Arial", 9)).grid(
                row=r, column=c, sticky="w", padx=4, pady=2)
            tk.Label(auto, textvariable=var,
                     bg="#0a2340", fg=self.ORANGE,
                     font=("Arial", 9, "bold"), width=13,
                     relief="sunken", anchor="center").grid(
                row=r, column=c+1, padx=4, pady=2)

        self.fv = {}
        MANUAL = [
            ("N° OF *",              "of_num",      "entry", None),
            ("Pilote *",             "pilote",       "combo", "Pilotes"),
            ("Co-Pilote",            "copilote",     "combo", "Co-Pilotes"),
            ("Taille produit",       "taille",       "combo", "Taille produit"),
            ("Type produit",         "type_prod",    "combo", "Type produit"),
            ("Code produit *",       "code_prod",    "entry", None),
            ("Poste *",              "poste",        "combo", "Postes"),
            ("Poids garnissage (g)", "poids",        "entry", None),
            ("Fibre",                "fibre",        "combo", "Fibre"),
            ("Qte fabriquee *",      "qte_fab",      "entry", None),
            ("Qte emballee",         "qte_emb",      "entry", None),
            ("Nb personnes",         "nb_pers",      "combo", "Nombre operateur"),
            ("OF taie",              "of_taie",      "entry", None),
            ("Traca fibre",          "traca",        "entry", None),
            ("Ref. taie",            "ref_taie",     "entry", None),
            ("Nb taie 2nd choix",    "nb_taie2",     "entry", None),
            ("Nb defaut couture",    "nb_def_cout",  "entry", None),
            ("Manquant taie",        "mq_taie",      "entry", None),
            ("Manquant housse",      "mq_housse",    "entry", None),
            ("Manquant encart",      "mq_encart",    "entry", None),
            ("Commentaire",          "comment",      "entry", None),
        ]
        mf = tk.Frame(card, bg=self.CARD)
        mf.grid(row=2, column=0, columnspan=4,
                padx=10, pady=4, sticky="ew")
        mf.columnconfigure(1, weight=1)

        for i, (lbl, key, ftype, lh) in enumerate(MANUAL):
            tk.Label(mf, text=lbl, bg=self.CARD, fg=self.FG,
                     font=("Arial", 10)).grid(
                row=i, column=0, sticky="w", padx=6, pady=3)
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                w = tk.Entry(mf, textvariable=var,
                             bg="#0a2340", fg=self.FG,
                             font=("Arial", 10), width=22,
                             insertbackground=self.FG, relief="sunken")
            else:
                vals = self._get_list(lh) if lh else []
                w = ttk.Combobox(mf, textvariable=var, values=vals,
                                 font=("Arial", 10), width=20,
                                 state="readonly")
            w.grid(row=i, column=1, padx=6, pady=3, sticky="ew")

        self._v_kit = tk.BooleanVar()
        tk.Checkbutton(mf, text="KIT de 2 pieces",
                       variable=self._v_kit,
                       bg=self.CARD, fg=self.FG,
                       selectcolor="#0a2340",
                       activebackground=self.CARD,
                       activeforeground=self.FG,
                       font=("Arial", 11, "bold")).grid(
            row=len(MANUAL), column=0, columnspan=2,
            pady=8, padx=6, sticky="w")

        def _upd_equiv(*_):
            try:
                qte = int(self.fv["qte_fab"].get() or 0)
            except Exception:
                qte = 0
            self._v_equiv.set(str(self._calc_equiv(
                qte, self.fv["taille"].get(), self.fv["type_prod"].get())))
        self.fv["qte_fab"].trace_add("write", _upd_equiv)
        self.fv["taille"].trace_add("write", _upd_equiv)
        self.fv["type_prod"].trace_add("write", _upd_equiv)

    def _build_events(self, parent):
        hdr = tk.Frame(parent, bg=self.CARD, height=36)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="EVENEMENTS", bg=self.CARD, fg=self.RED,
                 font=("Arial", 13, "bold")).pack(side="left", padx=14)

        cnv = tk.Canvas(parent, bg=self.BG, highlightthickness=0)
        sb  = ttk.Scrollbar(parent, orient="vertical", command=cnv.yview)
        cnv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cnv.pack(fill="both", expand=True)

        frame = tk.Frame(cnv, bg=self.BG)
        win_id = cnv.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>",
                   lambda e: cnv.configure(scrollregion=cnv.bbox("all")))
        cnv.bind("<Configure>",
                 lambda e: cnv.itemconfig(win_id, width=e.width))

        self._evt_btns = {}
        self._evt_lbls = {}

        ratt_lf = tk.LabelFrame(
            frame, text=" ARRETS RATTRAPAGE ",
            bg=self.BG, fg=self.ORANGE,
            font=("Arial", 11, "bold"), labelanchor="n")
        ratt_lf.pack(fill="x", padx=8, pady=(8, 4))

        pb_lf = tk.LabelFrame(
            frame, text=" PROBLEMES TECHNIQUES ",
            bg=self.BG, fg=self.RED,
            font=("Arial", 11, "bold"), labelanchor="n")
        pb_lf.pack(fill="x", padx=8, pady=(4, 8))

        for label, key, cat in EVENTS:
            container = ratt_lf if cat == "ratt" else pb_lf
            row_f = tk.Frame(container, bg=self.BG)
            row_f.pack(fill="x", padx=4, pady=2)

            t_lbl = tk.Label(row_f, text="00:00:00",
                             bg=self.BG, fg=self.GRAY,
                             font=("Arial", 10, "bold"), width=9)
            t_lbl.pack(side="right", padx=6)

            btn = tk.Button(row_f, text=f"  {label}",
                            command=lambda k=key: self._toggle_event(k),
                            bg="#2c3e50", fg=self.FG,
                            font=("Arial", 10), relief="raised",
                            anchor="w", padx=10, pady=5, cursor="hand2")
            btn.pack(side="left", fill="x", expand=True)

            self._evt_btns[key] = btn
            self._evt_lbls[key] = t_lbl

    def _toggle_event(self, key):
        if self._t_running(key):
            self._t_stop(key)
            self._evt_btns[key].config(bg="#27ae60")
        else:
            self._t_start(key)
            self._evt_btns[key].config(bg=self.RED)

    def _tick(self):
        if not self._prod_active:
            return
        now = datetime.datetime.now()
        of_s = (now - self._of_start).total_seconds()

        self._of_clk.config(text=f"⏱  {fmt(of_s)}")
        self._v_duree.set(fmt(of_s))

        try:
            qte = int(self.fv["qte_fab"].get() or 0)
            of_min = of_s / 60
            c1 = round(qte / of_min, 2) if of_min > 0 else 0
            try:
                np_ = int(self.fv["nb_pers"].get() or 1)
            except Exception:
                np_ = 1
            c2 = (round(qte / (np_ * of_s / 3600), 2)
                  if (np_ > 0 and of_s > 0) else 0)
            self._v_cad1.set(str(c1))
            self._v_cad2.set(str(c2))
        except Exception:
            pass

        for key, lbl in self._evt_lbls.items():
            s = self._t_get(key)
            lbl.config(text=fmt(s))
            if self._t_running(key):
                lbl.config(fg=self.RED)
            elif s > 0:
                lbl.config(fg=self.ORANGE)
            else:
                lbl.config(fg=self.GRAY)

        self._after_id = self.root.after(1000, self._tick)

    # =========================================================================
    #  FIN DE PRODUCTION -> ECRITURE EXCEL
    # =========================================================================
    def _end_production(self):
        end_dt = datetime.datetime.now()
        self._t_stop_all()
        self._prod_active = False

        of_s = (end_dt - self._of_start).total_seconds()
        self._v_end.set(end_dt.strftime("%H:%M:%S"))
        self._v_duree.set(fmt(of_s))

        v = {k: var.get().strip() for k, var in self.fv.items()}

        if not v.get("of_num"):
            if not messagebox.askyesno(
                    "Attention",
                    "N° OF non saisi. Continuer quand meme ?"):
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
        c1    = round(qte_fab / of_min, 2) if of_min > 0 else 0
        c2    = (round(qte_fab / (nb_pers * of_hrs), 2) if of_hrs > 0 else 0)
        equiv = self._calc_equiv(
            qte_fab, v.get("taille", ""), v.get("type_prod", ""))
        kit   = 2 if self._v_kit.get() else 1

        def _ts(key):
            return fmt(self._t_get(key))

        row = [
            v.get("of_num", ""),                                   # A
            datetime.date.today().strftime("%d/%m/%Y"),            # B
            v.get("poste", ""),                                    # C
            v.get("pilote", ""),                                   # D
            v.get("copilote", ""),                                 # E
            v.get("nb_pers", ""),                                  # F
            v.get("taille", ""),                                   # G
            v.get("code_prod", ""),                                # H
            v.get("type_prod", ""),                                # I
            v.get("poids", ""),                                    # J
            v.get("fibre", ""),                                    # K
            v.get("of_taie", ""),                                  # L
            v.get("traca", ""),                                    # M
            qte_fab,                                               # N
            _n("qte_emb"),                                         # O
            equiv,                                                 # P
            fmt(of_s),                                             # Q
            self._of_start.strftime("%H:%M:%S"),                  # R
            end_dt.strftime("%H:%M:%S"),                          # S
            c1,                                                    # T
            c2,                                                    # U
            kit,                                                   # V
            v.get("ref_taie", ""),                                # W
            _n("nb_def_cout"),                                     # X
            _n("mq_taie"),                                         # Y
            (f"Housse:{v.get('mq_housse','')} "
             f"Encart:{v.get('mq_encart','')}"),                  # Z
            _ts("ratt_pochon"),                                    # AA
            _ts("ratt_couture"),                                   # AB
            _ts("ratt_emb"),                                       # AC
            _ts("ratt_presse_soud"),                               # AD
            _ts("ratt_presse_zip"),                                # AE
            _ts("pb_chargeuse"),                                   # AF
            _ts("pb_carde"),                                       # AG
            _ts("pb_etaleur"),                                     # AH
            _ts("pb_coupe"),                                       # AI
            _ts("pb_tapis1"),                                      # AJ
            _ts("pb_enrouleur"),                                   # AK
            _ts("pb_pesee"),                                       # AL
            _ts("pb_deviation"),                                   # AM
            _ts("pb_enfileur"),                                    # AN
            _ts("pb_kinna"),                                       # AO
            _ts("pb_tapeuse"),                                     # AP
            _ts("pb_table_rot"),                                   # AQ
            _ts("pb_h100"),                                        # AR
            _ts("pb_traversin"),                                   # AS
            _ts("pb_presse_orc"),                                  # AT
            _ts("pb_presse_zip2"),                                 # AU
            _ts("pb_cercleuse"),                                   # AV
            _ts("pb_enrouleuse"),                                  # AW
            v.get("comment", ""),                                  # AX
        ]

        ok = self._write_excel(row)
        if ok:
            messagebox.showinfo("Succes", "Production declaree avec succes !")
        else:
            messagebox.showerror(
                "Erreur",
                "Impossible d'ecrire dans la base de donnees.\n"
                "Verifiez que le fichier Excel n'est pas ouvert.")
        self._show_main()

    def _calc_equiv(self, qte, taille, type_prod):
        equiv_list = self._get_list("Equivalence")
        if not equiv_list:
            return qte
        coef = 1.0
        for item in equiv_list:
            try:
                coef = float(str(item).replace(",", "."))
                break
            except Exception:
                pass
        return round(qte * coef, 2)

    def _write_excel(self, row):
        path = self.cfg.get("db_path", "")
        if not path:
            messagebox.showwarning("Attention",
                                   "Aucune base de donnees selectionnee !")
            return False
        try:
            wb = load_workbook(path)
            ws = wb["Data"]
            ws.append(row)
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
    root.configure(bg="#16213e")
    try:
        root.state("zoomed")
    except Exception:
        root.attributes("-fullscreen", True)
    App(root)
    root.mainloop()
