"""KPI-ORC v4.0 - Dashboard moderne navy/orange"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, datetime, math
from openpyxl import load_workbook

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")

NAVY    = "#1e2d4a"
NAVY_L  = "#26395e"
ORANGE  = "#f5a623"
WHITE   = "#ffffff"
BG      = "#eef2f7"
FORM_BG = "#e8f0f8"
SHAD    = "#b8c4d4"
GREEN   = "#27ae60"
C_RED   = "#c0392b"
C_RATT  = "#e67e22"
C_PB    = "#2980b9"
GRAY    = "#64748b"
LGRAY   = "#dde4ef"
DARK    = "#0f172a"

# Mettre 480 pour revenir a 8 heures
TIMELINE_WINDOW = 10  # minutes

EVENTS = [
    ("Pochon / Fibre",       "ratt_pochon",      "ratt"),
    ("Couture",              "ratt_couture",     "ratt"),
    ("Emballage",            "ratt_emb",         "ratt"),
    ("Presse Souder",        "ratt_presse_soud", "ratt"),
    ("Presse ZIP",           "ratt_presse_zip",  "ratt"),
    ("Chargeuse",            "pb_chargeuse",     "pb"),
    ("Carde",                "pb_carde",         "pb"),
    ("Etaleur / Tour",       "pb_etaleur",       "pb"),
    ("Coupe / Circ.",        "pb_coupe",         "pb"),
    ("Tapis Bascule",        "pb_tapis1",        "pb"),
    ("Enrouleur Pochon",     "pb_enrouleur",     "pb"),
    ("Pesee / Tapis 2",      "pb_pesee",         "pb"),
    ("Deviation / Table",    "pb_deviation",     "pb"),
    ("Enfileur Pochon",      "pb_enfileur",      "pb"),
    ("Kinna / Stroebel",     "pb_kinna",         "pb"),
    ("Tapeuse",              "pb_tapeuse",       "pb"),
    ("Table Rot. / Twin",    "pb_table_rot",     "pb"),
    ("Enfileuse H100",       "pb_h100",          "pb"),
    ("Enfileuse Traversin",  "pb_traversin",     "pb"),
    ("Presse ORC",           "pb_presse_orc",    "pb"),
    ("Presse Housse ZIP",    "pb_presse_zip2",   "pb"),
    ("Cercleuse",            "pb_cercleuse",     "pb"),
    ("Enrouleuse Traversin", "pb_enrouleuse",    "pb"),
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
    r = max(0, min(255, int(hx[1:3], 16) + d))
    g = max(0, min(255, int(hx[3:5], 16) + d))
    b = max(0, min(255, int(hx[5:7], 16) + d))
    return f"#{r:02x}{g:02x}{b:02x}"


def _rrect(cv, x1, y1, x2, y2, r, fill):
    cv.create_arc(x1,     y1,     x1+2*r, y1+2*r, start=90,  extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_arc(x2-2*r, y1,     x2,     y1+2*r, start=0,   extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_arc(x1,     y2-2*r, x1+2*r, y2,     start=180, extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_arc(x2-2*r, y2-2*r, x2,     y2,     start=270, extent=90,  style="pieslice", fill=fill, outline=fill)
    cv.create_rectangle(x1+r, y1,   x2-r, y2,   fill=fill, outline=fill)
    cv.create_rectangle(x1,   y1+r, x2,   y2-r, fill=fill, outline=fill)


def shadow_frame(parent, bg=WHITE, shadow_px=4):
    wrap  = tk.Frame(parent, bg=SHAD)
    inner = tk.Frame(wrap, bg=bg)
    inner.pack(fill="both", expand=True, padx=(0, shadow_px), pady=(0, shadow_px))
    return wrap, inner


# ─────────────────────────────────────────────────────────────────────────────
#  Timeline
# ─────────────────────────────────────────────────────────────────────────────
class Timeline(tk.Canvas):
    BAR_Y = 24
    BAR_H = 34

    def __init__(self, parent, app, **kw):
        kw.setdefault("height", 72)
        kw.setdefault("bg", WHITE)
        kw.setdefault("highlightthickness", 0)
        super().__init__(parent, **kw)
        self.app = app
        self.bind("<Configure>", lambda e: self.redraw())

    def redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 20:
            return

        now   = datetime.datetime.now()
        win_s = TIMELINE_WINDOW * 60
        t0    = now - datetime.timedelta(seconds=win_s)

        def px(dt):
            return max(0.0, min(float(w), (dt - t0).total_seconds() / win_s * w))

        BY, BH = self.BAR_Y, self.BAR_H

        # Fond barre
        self.create_rectangle(0, BY, w, BY + BH, fill="#cfdaeb", outline="")

        # Périodes OF (vert)
        for p in self.app._of_periods:
            x1 = px(p["start"])
            x2 = px(p.get("end") or now)
            if x2 > x1:
                self.create_rectangle(x1, BY, x2, BY + BH, fill=GREEN, outline="")

        # Rattrapages (orange) par-dessus
        for ev in self.app._tl_events:
            if ev["cat"] != "ratt":
                continue
            x1 = px(ev["start"])
            x2 = px(ev.get("end") or now)
            if x2 > x1:
                self.create_rectangle(x1, BY, x2, BY + BH, fill=C_RATT, outline="")

        # PB Techniques (rouge) par-dessus
        for ev in self.app._tl_events:
            if ev["cat"] != "pb":
                continue
            x1 = px(ev["start"])
            x2 = px(ev.get("end") or now)
            if x2 > x1:
                self.create_rectangle(x1, BY, x2, BY + BH, fill=C_RED, outline="")

        # Séparateurs changement d'OF
        for t_sep in self.app._of_changes:
            x = px(t_sep)
            if 2 < x < w - 2:
                self.create_line(x, BY - 6, x, BY + BH + 6, fill=ORANGE, width=3)
                self.create_polygon(x - 10, BY - 20, x + 10, BY - 20,
                                    x, BY - 6, fill=ORANGE, outline="")
                self.create_text(x, BY - 26, text="CHG OF",
                                 fill=ORANGE, font=("Arial", 7, "bold"), anchor="center")

        # Marqueurs minute
        for i in range(TIMELINE_WINDOW + 1):
            t  = t0 + datetime.timedelta(minutes=i)
            x  = i / TIMELINE_WINDOW * w
            col = "#aab8cc" if i % 5 == 0 else "#ccd6e4"
            self.create_line(x, BY - 2, x, BY + BH + 2, fill=col, width=1)
            if i % 2 == 0:
                self.create_text(x, BY - 10, text=t.strftime("%H:%M"),
                                 font=("Arial", 7), fill=GRAY, anchor="center")

        # Curseur maintenant
        self.create_line(w - 1, BY - 8, w - 1, BY + BH + 8, fill=ORANGE, width=2)

        # Légende
        legend = [("Prod.", GREEN), ("Rattrapage", C_RATT), ("PB Tech.", C_RED)]
        lx = 8
        for lbl, col in legend:
            self.create_rectangle(lx, BY + 9, lx + 12, BY + BH - 9,
                                  fill=col, outline="")
            self.create_text(lx + 15, BY + BH // 2, text=lbl, anchor="w",
                             font=("Arial", 8, "bold"), fill=GRAY)
            lx += 90


# ─────────────────────────────────────────────────────────────────────────────
#  Jauge TRS
# ─────────────────────────────────────────────────────────────────────────────
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
                        start=0, extent=180, style="arc", outline=LGRAY, width=16)
        ext   = self._val * 180 / 100
        color = C_RED if self._val < 55 else ORANGE if self._val < 75 else GREEN
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
        self.create_text(cx, cy - 6, text=f"a {self._time}",
                         font=("Arial", 9), fill=GRAY)


# ─────────────────────────────────────────────────────────────────────────────
#  Bouton evenement 3D
# ─────────────────────────────────────────────────────────────────────────────
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
        r = 10

        if running:
            _rrect(self, 4, 5, w - 1, h,       r, fill=_off(C_RED, -40))
            _rrect(self, 0, 0, w - 5, h - 5,   r, fill=C_RED)
            _rrect(self, 2, 2, w - 7, max(r*2+2, h//3), r, fill=_off(C_RED, +30))
            self.create_text(w//2 - 2, h*2//5,
                             text=self.label, fill=WHITE,
                             font=("Arial", 11, "bold"),
                             justify="center", width=w - 12)
            self.create_text(w//2 - 2, h*3//4,
                             text=fmt(elapsed), fill=WHITE,
                             font=("Arial", 15, "bold"))
            self.create_oval(w - 17, 7, w - 9, 15,
                             fill="#ff8080", outline=WHITE, width=1)
        else:
            _rrect(self, 4, 5, w - 1, h,     r, fill=SHAD)
            _rrect(self, 0, 0, w - 5, h - 5, r, fill=WHITE)
            self.create_rectangle(3, r + 2, 8, h - r - 7,
                                  fill=self.accent, outline="")
            self.create_text(w//2 + 2, h*2//5,
                             text=self.label, fill=DARK,
                             font=("Arial", 11, "bold"),
                             justify="center", width=w - 20)
            if elapsed > 0:
                self.create_text(w//2 + 2, h*3//4,
                                 text=fmt(elapsed), fill=ORANGE,
                                 font=("Arial", 13, "bold"))
            else:
                self.create_text(w//2 + 2, h*3//4,
                                 text="- - -", fill=LGRAY,
                                 font=("Arial", 10))


# ─────────────────────────────────────────────────────────────────────────────
#  Application
# ─────────────────────────────────────────────────────────────────────────────
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
        self._last_of_end = None
        self._prod_active = False
        self._after_id    = None
        self._db_labels   = []
        self._cells       = []
        self._tl_widget   = None
        self._cumul_lbl   = None

        self._of_periods  = []
        self._tl_events   = []
        self._of_changes  = []

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
                  bg=NAVY_L, fg=WHITE, font=("Arial", 9, "bold"),
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

    def _t_total_stops(self):
        return sum(self._t_get(k) for k in self._timers)

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
        self._cumul_lbl = None

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
        zone = tk.Frame(parent, bg=WHITE)
        zone.pack(fill="x")
        tl = Timeline(zone, self, bg=WHITE)
        tl.pack(fill="x", padx=6, pady=6)
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

        top = tk.Frame(body, bg=BG)
        top.pack(fill="x", pady=(0, 14))

        trs_wrap, trs_inner = shadow_frame(top, bg=WHITE)
        trs_wrap.pack(side="left", padx=(0, 16))
        tk.Label(trs_inner, text="TRS", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold")).pack(pady=(10, 0), padx=16)
        tk.Label(trs_inner, text="Taux de Rendement Synthetique",
                 bg=WHITE, fg=LGRAY, font=("Arial", 8)).pack()
        self._main_gauge = Gauge(trs_inner, bg=WHITE,
                                  width=310, height=130, highlightthickness=0)
        self._main_gauge.pack(padx=16, pady=(0, 10))
        self._refresh_main_kpi()

        btn_wrap   = tk.Frame(top, bg=BG)
        btn_wrap.pack(side="left", fill="both", expand=True)
        btn_canvas = tk.Canvas(btn_wrap, bg=BG, highlightthickness=0)
        btn_canvas.pack(fill="both", expand=True)

        def _draw_start(e=None):
            btn_canvas.delete("all")
            bw, bh = btn_canvas.winfo_width(), btn_canvas.winfo_height()
            if bw < 10 or bh < 10:
                return
            _rrect(btn_canvas, 5, 7, bw-1, bh, 16, fill=_off(ORANGE, -40))
            _rrect(btn_canvas, 0, 0, bw-6, bh-7, 16, fill=ORANGE)
            _rrect(btn_canvas, 3, 3, bw-9, bh//3, 16, fill=_off(ORANGE, +40))
            btn_canvas.create_text(bw//2 - 3, bh//2 - 3,
                                   text="▶  DEMARRER UNE PRODUCTION",
                                   fill=WHITE, font=("Arial", 17, "bold"))

        btn_canvas.bind("<Configure>", _draw_start)
        btn_canvas.bind("<Button-1>",  lambda e: self._start_production())
        btn_canvas.config(cursor="hand2")

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
        style.map("KPI.Treeview", background=[("selected", "#dbeafe")])

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
        last_time, trs = "--:--", 0.0
        if path and os.path.exists(path):
            try:
                wb   = load_workbook(path, read_only=True, data_only=True)
                rows = [r for r in wb["Data"].iter_rows(min_row=2, values_only=True)
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
            wb   = load_workbook(path, read_only=True, data_only=True)
            rows = [list(r) + [None]*55
                    for r in wb["Data"].iter_rows(min_row=2, values_only=True)
                    if any(r)]
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
        now = datetime.datetime.now()

        # Alerte changement d'OF si un OF précédent existe
        if self._last_of_end is not None:
            gap = (now - self._last_of_end).total_seconds()
            if gap > 30:
                h = int(gap // 3600)
                m = int((gap % 3600) // 60)
                s = int(gap % 60)
                time_str = f"{h}h {m:02d}min" if h > 0 else f"{m}min {s:02d}s"
                if messagebox.askyesno(
                        "Changement d'OF",
                        f"Le dernier OF a ete termine il y a {time_str}.\n\n"
                        "Voulez-vous declarer ce temps comme\n"
                        "'Changement d'OF' ?"):
                    self._write_changement_of_excel(self._last_of_end, now)
                    self._tl_events.append({
                        "key": "_changeof", "cat": "changeof",
                        "start": self._last_of_end, "end": now
                    })

        self._t_reset()
        self._of_start    = now
        self._prod_active = True
        self._cells       = []

        if self._of_periods:
            self._of_changes.append(now)
        self._of_periods.append({"start": now, "end": None})
        self._show_production()

    def _show_production(self):
        self._clear()
        self._db_labels.clear()

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)

        # En-tête
        hdr = tk.Frame(outer, bg=NAVY, height=66)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="ORC1  —  PRODUCTION EN COURS",
                 bg=NAVY, fg=WHITE,
                 font=("Arial", 14, "bold")).pack(side="left", padx=20)
        tk.Label(hdr,
                 text=f"Debut : {self._of_start.strftime('%H:%M:%S')}",
                 bg=NAVY, fg="#7a99c0",
                 font=("Arial", 11)).pack(side="left", padx=10)
        self._of_clk = tk.Label(hdr, text="00:00:00",
                                 bg=NAVY, fg="#4ade80",
                                 font=("Arial", 26, "bold"))
        self._of_clk.pack(side="left", padx=8)

        # Cumul des arrêts (à droite du chrono vert)
        self._cumul_lbl = tk.Label(hdr, text="Arrets: 00:00:00",
                                    bg=NAVY, fg=C_RATT,
                                    font=("Arial", 13, "bold"))
        self._cumul_lbl.pack(side="left", padx=18)

        self._db_widget(hdr, NAVY).pack(side="right", padx=14)

        # Timeline fond blanc
        self._make_timeline(outer)
        tk.Frame(outer, bg=LGRAY, height=1).pack(fill="x")

        # Bouton FIN (en bas)
        end_bar = tk.Frame(outer, bg="#6b1c1c", height=50)
        end_bar.pack(fill="x", side="bottom")
        end_bar.pack_propagate(False)
        tk.Button(end_bar,
                  text="⏹   DECLARER LA FIN DE PRODUCTION",
                  command=self._end_production,
                  bg="#6b1c1c", fg=WHITE,
                  font=("Arial", 14, "bold"),
                  relief="flat", cursor="hand2",
                  activebackground="#7f1d1d").pack(fill="both", expand=True)

        # Corps
        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=FORM_BG, width=480)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self._build_form(left)

        tk.Frame(body, bg=SHAD, width=2).pack(side="left", fill="y")

        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_events(right)

        self._tick()

    # ── Formulaire SANS SCROLL (grille 2 colonnes) ───────────────────────────
    def _build_form(self, parent):
        self.fv = {}

        tk.Label(parent, text="DONNEES DE L'OF",
                 bg=FORM_BG, fg=NAVY,
                 font=("Arial", 11, "bold")).pack(
            anchor="w", padx=12, pady=(8, 2))

        c = tk.Frame(parent, bg=FORM_BG)
        c.pack(fill="both", expand=True, padx=6, pady=2)
        c.columnconfigure(0, weight=1)
        c.columnconfigure(1, weight=1)

        ri = [0]  # row index mutable

        def sec(txt):
            tk.Label(c, text=txt, bg=FORM_BG, fg=NAVY_L,
                     font=("Arial", 8, "bold")).grid(
                row=ri[0], column=0, columnspan=2,
                sticky="w", padx=4, pady=(5, 0))
            ri[0] += 1

        def fld(lbl_txt, key, ftype, lh=None, col=0, adv=True):
            cell = tk.Frame(c, bg=FORM_BG)
            cell.grid(row=ri[0], column=col, sticky="ew", padx=3, pady=1)
            cell.columnconfigure(0, weight=1)
            tk.Label(cell, text=lbl_txt, bg=FORM_BG, fg=GRAY,
                     font=("Arial", 8), anchor="w").grid(
                row=0, column=0, sticky="w")
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                e = tk.Entry(cell, textvariable=var,
                             bg=WHITE, fg=DARK, font=("Arial", 9),
                             relief="solid", bd=1,
                             insertbackground=DARK)
                e.grid(row=1, column=0, sticky="ew", ipady=2)
            else:
                vals = self._get_list(lh) if lh else []
                cb = ttk.Combobox(cell, textvariable=var, values=vals,
                                  font=("Arial", 9), state="readonly")
                cb.grid(row=1, column=0, sticky="ew")
            if adv:
                ri[0] += 1

        def row2(l1, k1, t1, h1, l2, k2, t2, h2):
            fld(l1, k1, t1, h1, col=0, adv=False)
            fld(l2, k2, t2, h2, col=1, adv=True)

        sec("── Identification")
        row2("N° OF *",        "of_num",    "entry", None,
             "Poste *",        "poste",     "combo", "Postes")
        row2("Pilote *",       "pilote",    "combo", "Pilotes",
             "Co-Pilote",      "copilote",  "combo", "Co-Pilotes")
        fld("Nb personnes",    "nb_pers",   "combo", "Nombre operateur", col=0)

        sec("── Produit")
        row2("Taille produit", "taille",    "combo", "Taille produit",
             "Type produit",   "type_prod", "combo", "Type produit")
        row2("Code produit *", "code_prod", "entry", None,
             "Poids garnissage","poids",    "entry", None)
        fld("Fibre",           "fibre",     "combo", "Fibre", col=0)

        sec("── Quantites")
        row2("Qte fabriquee *","qte_fab",   "entry", None,
             "Qte emballee",   "qte_emb",   "entry", None)

        sec("── Taie / Qualite")
        row2("OF taie",        "of_taie",   "entry", None,
             "Traca fibre",    "traca",     "entry", None)
        row2("Ref. taie",      "ref_taie",  "entry", None,
             "Nb taie 2nd",    "nb_taie2",  "entry", None)
        row2("Nb def. couture","nb_def_cout","entry", None,
             "Mq. taie",       "mq_taie",   "entry", None)
        row2("Mq. housse",     "mq_housse", "entry", None,
             "Mq. encart",     "mq_encart", "entry", None)

        # KIT checkbox
        kit_row = tk.Frame(c, bg=FORM_BG)
        kit_row.grid(row=ri[0], column=0, columnspan=2,
                     sticky="w", padx=4, pady=(4, 0))
        ri[0] += 1
        self._v_kit = tk.BooleanVar()
        tk.Checkbutton(kit_row, text="KIT de 2 pieces",
                       variable=self._v_kit,
                       bg=FORM_BG, fg=DARK,
                       selectcolor=WHITE,
                       activebackground=FORM_BG,
                       font=("Arial", 9, "bold"),
                       cursor="hand2").pack(side="left")

        sec("── Commentaire")
        txt_cell = tk.Frame(c, bg=FORM_BG)
        txt_cell.grid(row=ri[0], column=0, columnspan=2,
                      sticky="ew", padx=3, pady=1)
        ri[0] += 1
        self._comment_txt = tk.Text(txt_cell, height=3,
                                     bg=WHITE, fg=DARK,
                                     font=("Arial", 9),
                                     relief="solid", bd=1,
                                     wrap="word",
                                     insertbackground=DARK)
        self._comment_txt.pack(fill="x")

    # ── Evenements ────────────────────────────────────────────────────────────
    def _build_events(self, parent):
        def sec_header(txt, color):
            f = tk.Frame(parent, bg=color, height=40)
            f.pack(fill="x", padx=10, pady=(8, 2))
            f.pack_propagate(False)
            tk.Label(f, text=txt, bg=color, fg=WHITE,
                     font=("Arial", 13, "bold")).pack(
                side="left", padx=12, pady=6)

        # RATTRAPAGES
        sec_header("▶  ARRETS RATTRAPAGE", C_RATT)

        ratt_wrap = tk.Frame(parent, bg=BG, height=115)
        ratt_wrap.pack(fill="x", padx=10, pady=(0, 4))
        ratt_wrap.pack_propagate(False)
        ratt_grid = tk.Frame(ratt_wrap, bg=BG)
        ratt_grid.pack(fill="both", expand=True)
        for col in range(5):
            ratt_grid.columnconfigure(col, weight=1)
        ratt_grid.rowconfigure(0, weight=1)

        for i, (label, key, cat) in enumerate(EVENTS[:5]):
            cell = EventCell(ratt_grid, label, key, cat, self)
            cell.grid(row=0, column=i, sticky="nsew", padx=4, pady=4)
            self._cells.append(cell)

        # Séparateur visuel
        sep_frame = tk.Frame(parent, bg=SHAD, height=3)
        sep_frame.pack(fill="x", padx=10, pady=(6, 2))

        # PB TECHNIQUES
        sec_header("⚠  PROBLEMES TECHNIQUES", C_PB)

        pb_wrap = tk.Frame(parent, bg=BG)
        pb_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        pb_grid = tk.Frame(pb_wrap, bg=BG)
        pb_grid.pack(fill="both", expand=True)
        for col in range(6):
            pb_grid.columnconfigure(col, weight=1)
        for row in range(3):
            pb_grid.rowconfigure(row, weight=1)

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
        if self._cumul_lbl:
            self._cumul_lbl.config(
                text=f"Arrets: {fmt(self._t_total_stops())}")
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
        # Alerte si arrêts actifs
        active = [k for k in self._timers if self._t_running(k)]
        if active:
            if not messagebox.askyesno(
                    "Attention",
                    f"Il y a {len(active)} arret(s) en cours.\n"
                    "Ils vont etre automatiquement arretes.\n\n"
                    "Continuer ?"):
                return

        end_dt = datetime.datetime.now()
        self._t_stop_all()
        self._tl_close_all()
        self._prod_active = False
        self._last_of_end = end_dt
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
        equiv   = self._calc_equiv(
            qte_fab, v.get("taille", ""), v.get("type_prod", ""))
        kit     = 2 if self._v_kit.get() else 1

        def _ts(key):
            return fmt(self._t_get(key))

        row = [
            v.get("of_num", ""),
            datetime.date.today().strftime("%d/%m/%Y"),
            v.get("poste", ""),
            v.get("pilote", ""),
            v.get("copilote", ""),
            v.get("nb_pers", ""),
            v.get("taille", ""),
            v.get("code_prod", ""),
            v.get("type_prod", ""),
            v.get("poids", ""),
            v.get("fibre", ""),
            v.get("of_taie", ""),
            v.get("traca", ""),
            qte_fab,
            _n("qte_emb"),
            equiv,
            fmt(of_s),
            self._of_start.strftime("%H:%M:%S"),
            end_dt.strftime("%H:%M:%S"),
            c1,
            c2,
            kit,
            v.get("ref_taie", ""),
            _n("nb_def_cout"),
            _n("mq_taie"),
            f"Housse:{v.get('mq_housse','')} Encart:{v.get('mq_encart','')}",
            _ts("ratt_pochon"),
            _ts("ratt_couture"),
            _ts("ratt_emb"),
            _ts("ratt_presse_soud"),
            _ts("ratt_presse_zip"),
            _ts("pb_chargeuse"),
            _ts("pb_carde"),
            _ts("pb_etaleur"),
            _ts("pb_coupe"),
            _ts("pb_tapis1"),
            _ts("pb_enrouleur"),
            _ts("pb_pesee"),
            _ts("pb_deviation"),
            _ts("pb_enfileur"),
            _ts("pb_kinna"),
            _ts("pb_tapeuse"),
            _ts("pb_table_rot"),
            _ts("pb_h100"),
            _ts("pb_traversin"),
            _ts("pb_presse_orc"),
            _ts("pb_presse_zip2"),
            _ts("pb_cercleuse"),
            _ts("pb_enrouleuse"),
            v.get("comment", ""),
        ]

        ok = self._write_excel(row, v)
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

    # ── Excel ─────────────────────────────────────────────────────────────────
    def _write_excel(self, row, v):
        path = self.cfg.get("db_path", "")
        if not path:
            messagebox.showwarning("Attention", "Aucune base de donnees !")
            return False
        try:
            wb = load_workbook(path)
            wb["Data"].append(row)
            self._write_events_to_wb(wb, v)
            wb.save(path)
            wb.close()
            return True
        except PermissionError:
            return False
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur Excel :\n{e}")
            return False

    def _ensure_events_sheet(self, wb):
        if "Evenements" not in wb.sheetnames:
            ws = wb.create_sheet("Evenements")
            ws.append([
                "Evenement", "OF", "Date", "Poste", "Pilote", "Co-Pilote",
                "Nb Personnes", "Taille", "Type Produit", "Code Produit",
                "Heure Debut", "Heure Fin", "Duree",
            ])
        return wb["Evenements"]

    def _write_events_to_wb(self, wb, v):
        ws = self._ensure_events_sheet(wb)
        of_start = self._of_start
        for ev in self._tl_events:
            if ev.get("cat") not in ("ratt", "pb"):
                continue
            if ev["start"] < of_start:
                continue
            cat_name = "Rattrapage" if ev["cat"] == "ratt" else "PB Technique"
            label = next(
                (e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
            start = ev["start"]
            end   = ev.get("end") or datetime.datetime.now()
            dur   = (end - start).total_seconds()
            ws.append([
                f"{cat_name}: {label}",
                v.get("of_num", ""),
                start.strftime("%d/%m/%Y"),
                v.get("poste", ""),
                v.get("pilote", ""),
                v.get("copilote", ""),
                v.get("nb_pers", ""),
                v.get("taille", ""),
                v.get("type_prod", ""),
                v.get("code_prod", ""),
                start.strftime("%H:%M:%S"),
                end.strftime("%H:%M:%S"),
                fmt(dur),
            ])

    def _write_changement_of_excel(self, start_dt, end_dt):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        try:
            wb  = load_workbook(path)
            ws  = self._ensure_events_sheet(wb)
            dur = (end_dt - start_dt).total_seconds()
            ws.append([
                "Changement d'OF",
                "", "", "", "", "", "", "", "", "",
                start_dt.strftime("%H:%M:%S"),
                end_dt.strftime("%H:%M:%S"),
                fmt(dur),
            ])
            wb.save(path)
            wb.close()
        except Exception:
            pass


if __name__ == "__main__":
    root = tk.Tk()
    root.configure(bg=BG)
    try:
        root.state("zoomed")
    except Exception:
        root.attributes("-fullscreen", True)
    App(root)
    root.mainloop()
