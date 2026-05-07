"""KPI-ORC v3.0 - Design dashboard moderne navy/orange"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, datetime, math
from openpyxl import load_workbook

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY   = "#1e2d4a"
NAVY_L = "#26395e"
ORANGE = "#f5a623"
WHITE  = "#ffffff"
BG     = "#eef2f7"
SHAD   = "#b8c4d4"
GREEN  = "#27ae60"
C_RED  = "#c0392b"
C_RATT = "#e67e22"
C_PB   = "#2980b9"
GRAY   = "#64748b"
LGRAY  = "#dde4ef"
DARK   = "#0f172a"

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


def _off(hx, d):
    """Eclaircit (+) ou assombrit (-) une couleur hex."""
    r = max(0, min(255, int(hx[1:3], 16) + d))
    g = max(0, min(255, int(hx[3:5], 16) + d))
    b = max(0, min(255, int(hx[5:7], 16) + d))
    return f"#{r:02x}{g:02x}{b:02x}"


def _rrect(cv, x1, y1, x2, y2, r, fill):
    """Rectangle a coins arrondis sur un Canvas."""
    cv.create_arc(x1,     y1,     x1+2*r, y1+2*r, start=90,  extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_arc(x2-2*r, y1,     x2,     y1+2*r, start=0,   extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_arc(x1,     y2-2*r, x1+2*r, y2,     start=180, extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_arc(x2-2*r, y2-2*r, x2,     y2,     start=270, extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_rectangle(x1+r, y1,   x2-r, y2,   fill=fill, outline=fill)
    cv.create_rectangle(x1,   y1+r, x2,   y2-r, fill=fill, outline=fill)


def shadow_frame(parent, bg=WHITE, shadow_px=4):
    """Renvoie un frame blanc avec ombre portee (bas-droite)."""
    wrap = tk.Frame(parent, bg=SHAD)
    inner = tk.Frame(wrap, bg=bg)
    inner.pack(fill="both", expand=True,
               padx=(0, shadow_px), pady=(0, shadow_px))
    return wrap, inner


# ── Widget Timeline ───────────────────────────────────────────────────────────
class Timeline(tk.Canvas):
    """Barre chronologique 8h : vert=prod, orange=rattrapage, rouge=PB."""
    BAR_Y = 16
    BAR_H = 18

    def __init__(self, parent, app, **kw):
        kw.setdefault("height", 48)
        kw.setdefault("bg", NAVY)
        kw.setdefault("highlightthickness", 0)
        super().__init__(parent, **kw)
        self.app = app
        self.bind("<Configure>", lambda e: self.redraw())

    def redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 20:
            return

        now = datetime.datetime.now()
        t0  = now - datetime.timedelta(hours=8)
        T   = 8 * 3600

        def px(dt):
            return max(0.0, min(float(w), (dt - t0).total_seconds() / T * w))

        BY, BH = self.BAR_Y, self.BAR_H

        # Fond barre (idle gris)
        self.create_rectangle(0, BY, w, BY + BH, fill="#2c4060", outline="")

        # Periodes OF (vert)
        for p in self.app._of_periods:
            x1 = px(p["start"])
            x2 = px(p.get("end") or now)
            if x2 > x1:
                self.create_rectangle(x1, BY, x2, BY + BH,
                                      fill=GREEN, outline="")

        # Rattrapages (orange) puis PB (rouge) par-dessus
        for cat_f, col in (("ratt", C_RATT), ("pb", C_RED)):
            for ev in self.app._tl_events:
                if ev["cat"] != cat_f:
                    continue
                x1 = px(ev["start"])
                x2 = px(ev.get("end") or now)
                if x2 > x1:
                    self.create_rectangle(x1, BY, x2, BY + BH,
                                          fill=col, outline="")

        # Marqueurs horaires
        for i in range(9):
            t  = t0 + datetime.timedelta(hours=i)
            x  = i / 8.0 * w
            self.create_line(x, BY - 4, x, BY + BH + 4,
                             fill="#3d5575", width=1)
            self.create_text(x, BY - 8, text=t.strftime("%H:%M"),
                             font=("Arial", 7), fill="#94a3b8", anchor="center")

        # Curseur "maintenant"
        self.create_line(w - 1, BY - 6, w - 1, BY + BH + 6,
                         fill=ORANGE, width=2)

        # Legende
        legend = [("  Prod.", GREEN), ("  Rattrapage", C_RATT), ("  PB Tech.", C_RED)]
        lx = 8
        for lbl, col in legend:
            self.create_rectangle(lx, BY + 4, lx + 10, BY + BH - 4,
                                  fill=col, outline="")
            self.create_text(lx + 13, BY + BH // 2,
                             text=lbl, anchor="w",
                             font=("Arial", 7, "bold"), fill="#94a3b8")
            lx += 90


# ── Widget Jauge TRS ─────────────────────────────────────────────────────────
class Gauge(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._val  = 0.0
        self._time = "--:--"
        self.bind("<Configure>", lambda e: self._draw())

    def update_gauge(self, value, time_str=""):
        self._val  = max(0.0, min(100.0, float(value)))
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
                        outline=LGRAY, width=16)
        ext   = self._val * 180 / 100
        color = (C_RED if self._val < 55
                 else ORANGE if self._val < 75 else GREEN)
        if ext > 0:
            self.create_arc(cx-r, cy-r, cx+r, cy+r,
                            start=180, extent=-ext, style="arc",
                            outline=color, width=16)
        angle = math.radians(180 - self._val * 180 / 100)
        nx = cx + int((r - 6) * math.cos(angle))
        ny = cy - int((r - 6) * math.sin(angle))
        self.create_line(cx, cy, nx, ny, fill=DARK, width=2)
        self.create_oval(cx-5, cy-5, cx+5, cy+5, fill=DARK, outline="")
        self.create_text(cx, cy - r // 2,
                         text=f"TRS  {self._val:.0f}%",
                         font=("Arial", 14, "bold"), fill=color)
        self.create_text(cx, cy - 6,
                         text=f"a {self._time}",
                         font=("Arial", 9), fill=GRAY)


# ── Bouton evenement 3D (Canvas) ─────────────────────────────────────────────
class EventCell(tk.Canvas):
    def __init__(self, parent, label, key, cat, app, **kw):
        super().__init__(parent, bg=parent.cget("bg"),
                         highlightthickness=0, cursor="hand2", **kw)
        self.key    = key
        self.app    = app
        self.label  = label
        self.accent = C_RATT if cat == "ratt" else C_PB
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>",  lambda e: self._toggle())

    def _toggle(self):
        if self.app._t_running(self.key):
            self.app._t_stop(self.key)
            self.app._tl_close(self.key)
        else:
            self.app._t_start(self.key)
            cat = "ratt" if self.accent == C_RATT else "pb"
            self.app._tl_open(self.key, cat)
        self._draw()

    def refresh(self):
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return

        running = self.app._t_running(self.key)
        elapsed = self.app._t_get(self.key)
        r       = 10

        if running:
            # Ombre
            _rrect(self, 4, 5, w - 1, h, r, fill=_off(C_RED, -40))
            # Corps rouge
            _rrect(self, 0, 0, w - 5, h - 5, r, fill=C_RED)
            # Reflet haut
            _rrect(self, 2, 2, w - 7, max(r * 2 + 2, h // 3), r,
                   fill=_off(C_RED, +30))
            # Texte label
            self.create_text(w // 2 - 2, h // 3,
                             text=self.label, fill=WHITE,
                             font=("Arial", 8, "bold"),
                             justify="center", width=w - 12)
            # Chrono gros
            self.create_text(w // 2 - 2, h * 2 // 3,
                             text=fmt(elapsed), fill=WHITE,
                             font=("Arial", 13, "bold"))
            # Dot rouge vif
            self.create_oval(w - 17, 7, w - 9, 15,
                             fill="#ff8080", outline=WHITE, width=1)
        else:
            # Ombre douce
            _rrect(self, 4, 5, w - 1, h, r, fill=SHAD)
            # Corps blanc
            _rrect(self, 0, 0, w - 5, h - 5, r, fill=WHITE)
            # Barre accent gauche
            self.create_rectangle(3, r + 2, 8, h - r - 7,
                                  fill=self.accent, outline="")
            # Label
            self.create_text(w // 2 + 2, h // 3,
                             text=self.label, fill=DARK,
                             font=("Arial", 8, "bold"),
                             justify="center", width=w - 20)
            # Chrono
            if elapsed > 0:
                self.create_text(w // 2 + 2, h * 2 // 3,
                                 text=fmt(elapsed), fill=ORANGE,
                                 font=("Arial", 11, "bold"))
            else:
                self.create_text(w // 2 + 2, h * 2 // 3,
                                 text="- - -", fill=LGRAY,
                                 font=("Arial", 9))


# ── Application ──────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("KPI-ORC | Ligne ORC1")
        self.root.configure(bg=BG)
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
        self._tl_widget   = None  # timeline canvas actuel

        # Historique timeline (persist entre les vues)
        self._of_periods  = []   # [{start, end|None}]
        self._tl_events   = []   # [{start, end|None, key, cat}]

        self._load_lists()
        self._show_main()

    # ── Config ────────────────────────────────────────────────────────────────
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
        f    = tk.Frame(parent, bg=bg)
        name = os.path.basename(self.cfg.get("db_path", "")) or "Non connectee"
        lbl  = tk.Label(f, text=f"DB: {name}", bg=bg,
                        fg="#7a99c0", font=("Arial", 9))
        lbl.pack(side="left", padx=6)
        self._db_labels.append(lbl)
        tk.Button(f, text="Database", command=self._select_db,
                  bg=NAVY_L, fg=WHITE,
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

    # ── Timeline history ──────────────────────────────────────────────────────
    def _tl_open(self, key, cat):
        self._tl_events.append({
            "key": key, "cat": cat,
            "start": datetime.datetime.now(), "end": None
        })

    def _tl_close(self, key):
        for ev in reversed(self._tl_events):
            if ev["key"] == key and ev["end"] is None:
                ev["end"] = datetime.datetime.now()
                break

    def _tl_close_all(self):
        now = datetime.datetime.now()
        for ev in self._tl_events:
            if ev["end"] is None:
                ev["end"] = now

    # ── Navigation ────────────────────────────────────────────────────────────
    def _clear(self):
        if self._after_id:
            self.root.after_cancel(self._after_id)
            self._after_id = None
        for w in self.root.winfo_children():
            w.destroy()
        self._tl_widget = None

    def _make_header(self, parent, title, subtitle=""):
        hdr = tk.Frame(parent, bg=NAVY, height=62)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=title, bg=NAVY, fg=WHITE,
                 font=("Arial", 18, "bold")).pack(side="left", padx=22)
        if subtitle:
            tk.Label(hdr, text=subtitle, bg=NAVY, fg="#7a99c0",
                     font=("Arial", 11)).pack(side="left", padx=4)
        self._db_widget(hdr, NAVY).pack(side="right", padx=16)
        return hdr

    def _make_timeline(self, parent):
        tl = Timeline(parent, self, bg=NAVY)
        tl.pack(fill="x")
        self._tl_widget = tl
        return tl

    def _tl_tick(self):
        if self._tl_widget:
            try:
                self._tl_widget.redraw()
            except Exception:
                pass
        self._after_id = self.root.after(10000, self._tl_tick)

    # =========================================================================
    #  ECRAN PRINCIPAL
    # =========================================================================
    def _show_main(self):
        self._prod_active = False
        self._cells       = []
        self._clear()
        self._db_labels.clear()

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)

        self._make_header(outer, "KPI-ORC", "Ligne ORC1")
        self._make_timeline(outer)
        tk.Frame(outer, bg=LGRAY, height=1).pack(fill="x")

        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=14)

        # ── Ligne haute : TRS + bouton demarrer
        top = tk.Frame(body, bg=BG)
        top.pack(fill="x", pady=(0, 14))

        # Carte TRS avec ombre
        trs_wrap, trs_inner = shadow_frame(top, bg=WHITE)
        trs_wrap.pack(side="left", padx=(0, 16))
        tk.Label(trs_inner, text="TRS",
                 bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold")).pack(pady=(10, 0), padx=16)
        tk.Label(trs_inner, text="Taux de Rendement Synthetique",
                 bg=WHITE, fg=LGRAY,
                 font=("Arial", 8)).pack()
        self._main_gauge = Gauge(trs_inner, bg=WHITE,
                                  width=310, height=130,
                                  highlightthickness=0)
        self._main_gauge.pack(padx=16, pady=(0, 10))
        self._refresh_main_kpi()

        # Bouton demarrer (orange 3D)
        btn_wrap = tk.Frame(top, bg=BG)
        btn_wrap.pack(side="left", fill="both", expand=True)

        btn_canvas = tk.Canvas(btn_wrap, bg=BG, highlightthickness=0)
        btn_canvas.pack(fill="both", expand=True)

        def _draw_start_btn(event=None):
            btn_canvas.delete("all")
            bw = btn_canvas.winfo_width()
            bh = btn_canvas.winfo_height()
            if bw < 10 or bh < 10:
                return
            r = 16
            _rrect(btn_canvas, 5, 7, bw - 1, bh, r, fill=_off(ORANGE, -40))
            _rrect(btn_canvas, 0, 0, bw - 6, bh - 7, r, fill=ORANGE)
            _rrect(btn_canvas, 3, 3, bw - 9, bh // 3, r, fill=_off(ORANGE, +40))
            btn_canvas.create_text(bw // 2 - 3, bh // 2 - 3,
                                   text="▶  DEMARRER UNE PRODUCTION",
                                   fill=WHITE, font=("Arial", 17, "bold"))
        btn_canvas.bind("<Configure>", _draw_start_btn)
        btn_canvas.bind("<Button-1>",  lambda e: self._start_production())
        btn_canvas.config(cursor="hand2")

        # ── Tableau recapitulatif
        lbl_frame = tk.Frame(body, bg=BG)
        lbl_frame.pack(fill="x", pady=(0, 6))
        tk.Label(lbl_frame, text="15 Dernieres Declarations",
                 bg=BG, fg=DARK, font=("Arial", 11, "bold")).pack(side="left")

        tbl_wrap, tbl_inner = shadow_frame(body, bg=WHITE)
        tbl_wrap.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("KPI.Treeview",
                        background=WHITE, foreground=DARK,
                        fieldbackground=WHITE, rowheight=27,
                        font=("Arial", 10))
        style.configure("KPI.Treeview.Heading",
                        background=LGRAY, foreground=DARK,
                        font=("Arial", 10, "bold"), relief="flat")
        style.map("KPI.Treeview",
                  background=[("selected", "#dbeafe")])

        cols = ("Date", "OF", "Pilote", "Poste",
                "Duree OF", "PB Techniques", "Rattrapages")
        tree = ttk.Treeview(tbl_inner, columns=cols, show="headings",
                            height=13, style="KPI.Treeview")
        widths = {"Date": 100, "OF": 110, "Pilote": 175, "Poste": 110,
                  "Duree OF": 95, "PB Techniques": 120, "Rattrapages": 120}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 110), anchor="center")
        sb = ttk.Scrollbar(tbl_inner, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._load_table(tree)

        self._after_id = self.root.after(10000, self._tl_tick)

    def _refresh_main_kpi(self):
        path = self.cfg.get("db_path", "")
        last_time = "--:--"
        trs = 0.0
        if path and os.path.exists(path):
            try:
                wb   = load_workbook(path, read_only=True, data_only=True)
                ws   = wb["Data"]
                rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
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
            wb   = load_workbook(path, read_only=True, data_only=True)
            ws   = wb["Data"]
            rows = [list(r) + [None]*55
                    for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
            wb.close()
        except Exception:
            return

        def _sd(row, idx):
            t = 0
            for i in idx:
                if i < len(row) and row[i]:
                    p = str(row[i]).split(":")
                    try:
                        if len(p) == 3:
                            t += int(p[0])*3600 + int(p[1])*60 + int(p[2])
                    except Exception:
                        pass
            return t

        for row in list(reversed(rows))[:15]:
            tree.insert("", "end", values=(
                str(row[1])[:10] if row[1]  else "",
                str(row[0])      if row[0]  else "",
                str(row[3])      if row[3]  else "",
                str(row[2])      if row[2]  else "",
                str(row[16])     if row[16] else "",
                fmt(_sd(row, range(31, 49))),
                fmt(_sd(row, range(26, 31))),
            ))

    # =========================================================================
    #  ECRAN DE PRODUCTION
    # =========================================================================
    def _start_production(self):
        self._t_reset()
        self._of_start    = datetime.datetime.now()
        self._prod_active = True
        self._cells       = []
        self._of_periods.append({"start": self._of_start, "end": None})
        self._show_production()

    def _show_production(self):
        self._clear()
        self._db_labels.clear()

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)

        # En-tete
        hdr = tk.Frame(outer, bg=NAVY, height=66)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="ORC1  —  PRODUCTION EN COURS",
                 bg=NAVY, fg=WHITE,
                 font=("Arial", 14, "bold")).pack(side="left", padx=20)
        tk.Label(hdr, text=f"Debut : {self._of_start.strftime('%H:%M:%S')}",
                 bg=NAVY, fg="#7a99c0",
                 font=("Arial", 11)).pack(side="left", padx=14)
        self._of_clk = tk.Label(hdr, text="00:00:00",
                                 bg=NAVY, fg="#4ade80",
                                 font=("Arial", 28, "bold"))
        self._of_clk.pack(side="left", padx=16)
        self._db_widget(hdr, NAVY).pack(side="right", padx=14)

        # Timeline
        self._make_timeline(outer)
        tk.Frame(outer, bg=LGRAY, height=1).pack(fill="x")

        # Bouton FIN
        end_bar = tk.Frame(outer, bg="#6b1c1c", height=50)
        end_bar.pack(fill="x", side="bottom")
        end_bar.pack_propagate(False)
        tk.Button(end_bar, text="⏹   DECLARER LA FIN DE PRODUCTION",
                  command=self._end_production,
                  bg="#6b1c1c", fg=WHITE,
                  font=("Arial", 14, "bold"),
                  relief="flat", cursor="hand2",
                  activebackground="#7f1d1d").pack(fill="both", expand=True)

        # Corps : panneau gauche (formulaire) + panneau droit (evenements)
        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True)

        # Panneau formulaire (navy)
        left = tk.Frame(body, bg=NAVY_L, width=460)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self._build_form(left)

        # Separateur
        tk.Frame(body, bg=_off(NAVY, +20), width=2).pack(side="left", fill="y")

        # Panneau evenements
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_events(right)

        self._tick()

    # ── Formulaire (dark) ─────────────────────────────────────────────────────
    def _build_form(self, parent):
        cnv   = tk.Canvas(parent, bg=NAVY_L, highlightthickness=0)
        sb    = ttk.Scrollbar(parent, orient="vertical", command=cnv.yview)
        cnv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cnv.pack(fill="both", expand=True)

        frame  = tk.Frame(cnv, bg=NAVY_L)
        win_id = cnv.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>",
                   lambda e: cnv.configure(scrollregion=cnv.bbox("all")))
        cnv.bind("<Configure>",
                 lambda e: cnv.itemconfig(win_id, width=e.width))
        cnv.bind_all("<MouseWheel>",
                     lambda e: cnv.yview_scroll(-1*(e.delta//120), "units"))

        def section(txt):
            tk.Label(frame, text=txt, bg=NAVY_L, fg=ORANGE,
                     font=("Arial", 9, "bold")).pack(
                anchor="w", padx=14, pady=(10, 2))

        def field(lbl, key, ftype, lh=None):
            row = tk.Frame(frame, bg=NAVY_L)
            row.pack(fill="x", padx=14, pady=2)
            tk.Label(row, text=lbl, bg=NAVY_L, fg="#94a3b8",
                     font=("Arial", 9), width=17,
                     anchor="w").pack(side="left")
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                w = tk.Entry(row, textvariable=var,
                             bg=_off(NAVY_L, +15), fg=WHITE,
                             font=("Arial", 10), relief="flat",
                             insertbackground=WHITE,
                             highlightbackground=_off(NAVY_L, +30),
                             highlightthickness=1)
                w.pack(side="left", fill="x", expand=True, ipady=3)
            else:
                style = ttk.Style()
                style.configure("Dark.TCombobox",
                                fieldbackground=_off(NAVY_L, +15),
                                foreground=WHITE,
                                background=NAVY_L)
                vals = self._get_list(lh) if lh else []
                w = ttk.Combobox(row, textvariable=var, values=vals,
                                 font=("Arial", 10), state="readonly",
                                 style="Dark.TCombobox")
                w.pack(side="left", fill="x", expand=True)

        self.fv = {}
        tk.Label(frame, text="DONNEES DE L'OF",
                 bg=NAVY_L, fg=WHITE,
                 font=("Arial", 12, "bold")).pack(anchor="w", padx=14, pady=(14, 4))

        section("Identification")
        field("N° OF *",           "of_num",      "entry")
        field("Pilote *",          "pilote",      "combo", "Pilotes")
        field("Co-Pilote",         "copilote",    "combo", "Co-Pilotes")
        field("Poste *",           "poste",       "combo", "Postes")
        field("Nb personnes",      "nb_pers",     "combo", "Nombre operateur")

        section("Produit")
        field("Taille produit",    "taille",      "combo", "Taille produit")
        field("Type produit",      "type_prod",   "combo", "Type produit")
        field("Code produit *",    "code_prod",   "entry")
        field("Poids garnissage",  "poids",       "entry")
        field("Fibre",             "fibre",       "combo", "Fibre")

        section("Quantites")
        field("Qte fabriquee *",   "qte_fab",     "entry")
        field("Qte emballee",      "qte_emb",     "entry")

        section("Taie / Qualite")
        field("OF taie",           "of_taie",     "entry")
        field("Traca fibre",       "traca",       "entry")
        field("Ref. taie",         "ref_taie",    "entry")
        field("Nb taie 2nd choix", "nb_taie2",    "entry")
        field("Nb defaut couture", "nb_def_cout", "entry")
        field("Manquant taie",     "mq_taie",     "entry")
        field("Manquant housse",   "mq_housse",   "entry")
        field("Manquant encart",   "mq_encart",   "entry")

        self._v_kit = tk.BooleanVar()
        kit_row = tk.Frame(frame, bg=NAVY_L)
        kit_row.pack(fill="x", padx=14, pady=8)
        tk.Checkbutton(kit_row, text="KIT de 2 pieces",
                       variable=self._v_kit,
                       bg=NAVY_L, fg=WHITE,
                       selectcolor=NAVY,
                       activebackground=NAVY_L,
                       font=("Arial", 10, "bold"),
                       cursor="hand2").pack(side="left")

        section("Commentaire")
        self._comment_txt = tk.Text(frame, height=5,
                                     bg=_off(NAVY_L, +15), fg=WHITE,
                                     font=("Arial", 10), relief="flat",
                                     wrap="word",
                                     insertbackground=WHITE,
                                     highlightbackground=_off(NAVY_L, +30),
                                     highlightthickness=1)
        self._comment_txt.pack(fill="x", padx=14, pady=2)
        tk.Frame(frame, bg=NAVY_L, height=20).pack()

    # ── Evenements ────────────────────────────────────────────────────────────
    def _build_events(self, parent):
        def sec_header(txt, bg_c, fg_c):
            f = tk.Frame(parent, bg=bg_c, height=30)
            f.pack(fill="x", padx=12, pady=(8, 2))
            f.pack_propagate(False)
            tk.Label(f, text=txt, bg=bg_c, fg=fg_c,
                     font=("Arial", 10, "bold")).pack(
                side="left", padx=10, pady=4)

        # RATTRAPAGES : 5 en 1 ligne, hauteur fixe
        sec_header("ARRETS RATTRAPAGE",
                   _off(NAVY, +10), _off(C_RATT, +30))

        ratt_wrap = tk.Frame(parent, bg=BG, height=108)
        ratt_wrap.pack(fill="x", padx=12, pady=(0, 4))
        ratt_wrap.pack_propagate(False)
        ratt_grid = tk.Frame(ratt_wrap, bg=BG)
        ratt_grid.pack(fill="both", expand=True)
        for c in range(5):
            ratt_grid.columnconfigure(c, weight=1)
        ratt_grid.rowconfigure(0, weight=1)

        for i, (label, key, cat) in enumerate(EVENTS[:5]):
            cell = EventCell(ratt_grid, label, key, cat, self)
            cell.grid(row=0, column=i, sticky="nsew", padx=4, pady=4)
            self._cells.append(cell)

        # PB TECHNIQUES : 6x3 = 18, remplit le reste
        sec_header("PROBLEMES TECHNIQUES",
                   _off(NAVY, +10), _off(C_PB, +40))

        pb_wrap = tk.Frame(parent, bg=BG)
        pb_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        pb_grid = tk.Frame(pb_wrap, bg=BG)
        pb_grid.pack(fill="both", expand=True)
        for c in range(6):
            pb_grid.columnconfigure(c, weight=1)
        for r in range(3):
            pb_grid.rowconfigure(r, weight=1)

        for i, (label, key, cat) in enumerate(EVENTS[5:]):
            cell = EventCell(pb_grid, label, key, cat, self)
            cell.grid(row=i // 6, column=i % 6,
                      sticky="nsew", padx=4, pady=4)
            self._cells.append(cell)

    # ── Tick ──────────────────────────────────────────────────────────────────
    def _tick(self):
        if not self._prod_active:
            return
        of_s = (datetime.datetime.now() - self._of_start).total_seconds()
        self._of_clk.config(text=fmt(of_s))
        for cell in self._cells:
            cell.refresh()
        if self._tl_widget:
            try:
                self._tl_widget.redraw()
            except Exception:
                pass
        self._after_id = self.root.after(1000, self._tick)

    # =========================================================================
    #  FIN DE PRODUCTION → EXCEL
    # =========================================================================
    def _end_production(self):
        end_dt = datetime.datetime.now()
        self._t_stop_all()
        self._tl_close_all()
        self._prod_active = False
        if self._of_periods:
            self._of_periods[-1]["end"] = end_dt

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
            _ts("ratt_pochon"),     # AA
            _ts("ratt_couture"),    # AB
            _ts("ratt_emb"),        # AC
            _ts("ratt_presse_soud"),# AD
            _ts("ratt_presse_zip"), # AE
            _ts("pb_chargeuse"),    # AF
            _ts("pb_carde"),        # AG
            _ts("pb_etaleur"),      # AH
            _ts("pb_coupe"),        # AI
            _ts("pb_tapis1"),       # AJ
            _ts("pb_enrouleur"),    # AK
            _ts("pb_pesee"),        # AL
            _ts("pb_deviation"),    # AM
            _ts("pb_enfileur"),     # AN
            _ts("pb_kinna"),        # AO
            _ts("pb_tapeuse"),      # AP
            _ts("pb_table_rot"),    # AQ
            _ts("pb_h100"),         # AR
            _ts("pb_traversin"),    # AS
            _ts("pb_presse_orc"),   # AT
            _ts("pb_presse_zip2"),  # AU
            _ts("pb_cercleuse"),    # AV
            _ts("pb_enrouleuse"),   # AW
            v.get("comment",""),    # AX
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
    root.configure(bg=BG)
    try:
        root.state("zoomed")
    except Exception:
        root.attributes("-fullscreen", True)
    App(root)
    root.mainloop()
