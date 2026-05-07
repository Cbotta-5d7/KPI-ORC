"""KPI-ORC v5.7 - Style Dodo (bleu marine #1a1f5e + rouge #e31e24)"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, sys, datetime, math
from openpyxl import load_workbook

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")
PASSWORD    = "0000"

NAVY    = "#1a1f5e"   # Dodo bleu marine
NAVY_L  = "#2d3490"   # Dodo bleu marine clair
ORANGE  = "#e31e24"   # Dodo rouge vif
WHITE   = "#ffffff"
BG      = "#f0f4fb"   # Fond general bleu tres clair
FORM_BG = "#e8eef8"   # Fond formulaire bleu Dodo
SHAD    = "#c2cce0"
GREEN   = "#1a8c4e"   # Vert production
C_RED   = "#e31e24"   # Rouge Dodo (pannes)
C_RATT  = "#d97706"   # Ambre rattrapages
C_PB    = "#1a1f5e"   # Bleu Dodo PB techniques
GRAY    = "#64748b"
LGRAY   = "#dde4ef"
DARK    = "#0f172a"

TIMELINE_WINDOW = 10  # minutes  ← mettre 480 pour 8h

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

DATA_HEADERS = [
    "OF", "Date", "Poste", "Pilote", "Co-Pilote", "Nb Personnes",
    "Taille", "Code Produit", "Type Produit", "Poids Garnissage", "Fibre",
    "OF Taie", "Traca Fibre", "Qte Fabriquee", "Qte Emballee", "Equivalence",
    "Duree OF", "Heure Debut", "Heure Fin", "Cadence/min", "Cadence/h/pers",
    "Kit", "Ref Taie", "Nb Defaut Couture", "Mq Taie", "Mq Housse/Encart",
    "Ratt Pochon/Fibre", "Ratt Couture", "Ratt Emballage",
    "Ratt Presse Souder", "Ratt Presse ZIP",
    "PB Chargeuse", "PB Carde", "PB Etaleur/Tour", "PB Coupe/Circ",
    "PB Tapis Bascule", "PB Enrouleur Pochon", "PB Pesee/Tapis 2",
    "PB Deviation/Table", "PB Enfileur Pochon", "PB Kinna/Stroebel",
    "PB Tapeuse", "PB Table Rot/Twin", "PB Enfileuse H100",
    "PB Enfileuse Traversin", "PB Presse ORC", "PB Presse Housse ZIP",
    "PB Cercleuse", "PB Enrouleuse Traversin", "Commentaire",
]

EVT_HEADERS = [
    "Evenement", "OF", "Date", "Poste", "Pilote", "Co-Pilote",
    "Nb Personnes", "Taille", "Type Produit", "Code Produit",
    "Heure Debut", "Heure Fin", "Duree", "Commentaire",
]

ALL_EVENT_TYPES = (
    [f"Rattrapage: {e[0]}" for e in EVENTS[:5]] +
    [f"PB Technique: {e[0]}" for e in EVENTS[5:]] +
    ["Changement d'OF"]
)


def _toast(root, msg, bg="#27ae60", duration=3000):
    """Notification flottante qui disparait automatiquement."""
    t = tk.Toplevel(root)
    t.overrideredirect(True)
    t.attributes("-topmost", True)
    t.configure(bg=bg)
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    w, h = 420, 64
    t.geometry(f"{w}x{h}+{(sw-w)//2}+{sh-120}")
    tk.Label(t, text=msg, bg=bg, fg="white",
             font=("Arial", 14, "bold"), padx=24, pady=16).pack(fill="both", expand=True)
    # Fondu progressif puis destruction
    def _fade(alpha=1.0):
        if alpha <= 0:
            t.destroy()
            return
        try:
            t.attributes("-alpha", alpha)
            t.after(50, _fade, alpha - 0.05)
        except Exception:
            pass
    t.after(duration, _fade)


def fmt(seconds):
    h, r = divmod(int(max(0, seconds)), 3600)
    m, s = divmod(r, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _hms_to_sec(s):
    try:
        p = str(s).split(":")
        return int(p[0]) * 3600 + int(p[1]) * 60 + int(p[2])
    except Exception:
        return 0


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
        self._tooltip_win = None
        self.bind("<Configure>", lambda e: self.redraw())
        self.bind("<Button-1>",  self._on_click)

    def _on_click(self, event):
        w = self.winfo_width()
        if w < 20:
            return
        now   = datetime.datetime.now()
        win_s = TIMELINE_WINDOW * 60
        t0    = now - datetime.timedelta(seconds=win_s)
        click_t = t0 + datetime.timedelta(seconds=event.x / w * win_s)

        # Chercher l'evenement le plus proche du clic
        found = None
        for ev in self.app._tl_events:
            if ev.get("cat") not in ("ratt", "pb"):
                continue
            start = ev["start"]
            end   = ev.get("end") or now
            if start <= click_t <= end:
                found = ev
                break
        if not found:
            for p in self.app._of_periods:
                if p["start"] <= click_t <= (p.get("end") or now):
                    found = {"cat": "prod", "start": p["start"],
                             "end": p.get("end"), "key": "prod",
                             "of_num": p.get("of_num", "")}
                    break
        if not found:
            return
        self._show_tooltip(event.x, found)

    def _show_tooltip(self, x, ev):
        try:
            if self._tooltip_win:
                self._tooltip_win.destroy()
        except Exception:
            pass

        cat = ev.get("cat", "")
        if cat == "prod":
            title = f"Production  OF: {ev.get('of_num','?') or 'en cours'}"
            col = GREEN
        elif cat == "ratt":
            label = next((e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
            title = f"Rattrapage : {label}"
            col = C_RATT
        else:
            label = next((e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
            title = f"PB Technique : {label}"
            col = C_RED

        now   = datetime.datetime.now()
        start = ev["start"]
        end   = ev.get("end") or now
        dur   = (end - start).total_seconds()
        body  = (f"Debut : {start.strftime('%H:%M:%S')}\n"
                 f"Fin   : {end.strftime('%H:%M:%S')}\n"
                 f"Duree : {fmt(dur)}")
        if ev.get("comment"):
            body += f"\n\n💬 {ev['comment']}"

        top = tk.Toplevel(self.winfo_toplevel())
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=col)
        self._tooltip_win = top

        tk.Label(top, text=title, bg=col, fg=WHITE,
                 font=("Arial", 13, "bold"), padx=14, pady=6).pack(fill="x")
        tk.Label(top, text=body, bg=_off(col, -30), fg=WHITE,
                 font=("Arial", 11), padx=14, pady=8,
                 justify="left").pack(fill="x")

        top.update_idletasks()
        rw, rh = top.winfo_width(), top.winfo_height()
        cx = self.winfo_rootx() + x - rw // 2
        cy = self.winfo_rooty() - rh - 6
        top.geometry(f"+{max(0, cx)}+{max(0, cy)}")

        def _fade(alpha=1.0, step=0):
            if step < 30:
                top.after(100, _fade, alpha, step + 1)
            elif alpha > 0:
                try:
                    top.attributes("-alpha", alpha)
                    top.after(60, _fade, alpha - 0.07, step)
                except Exception:
                    pass
            else:
                try:
                    top.destroy()
                except Exception:
                    pass
        top.after(2800, _fade)

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
        self.create_rectangle(0, BY, w, BY + BH, fill="#cfdaeb", outline="")

        for p in self.app._of_periods:
            x1, x2 = px(p["start"]), px(p.get("end") or now)
            if x2 > x1:
                self.create_rectangle(x1, BY, x2, BY + BH, fill=GREEN, outline="")

        for ev in self.app._tl_events:
            if ev["cat"] not in ("ratt", "pb"):
                continue
            col = C_RATT if ev["cat"] == "ratt" else C_RED
            x1, x2 = px(ev["start"]), px(ev.get("end") or now)
            if x2 > x1:
                self.create_rectangle(x1, BY, x2, BY + BH, fill=col, outline="")

        for t_sep in self.app._of_changes:
            x = px(t_sep)
            if 2 < x < w - 2:
                self.create_line(x, BY - 6, x, BY + BH + 6, fill=ORANGE, width=3)
                self.create_polygon(x - 10, BY - 20, x + 10, BY - 20,
                                    x, BY - 6, fill=ORANGE, outline="")
                self.create_text(x, BY - 26, text="CHG OF",
                                 fill=ORANGE, font=("Arial", 7, "bold"), anchor="center")

        for i in range(TIMELINE_WINDOW + 1):
            t = t0 + datetime.timedelta(minutes=i)
            x = i / TIMELINE_WINDOW * w
            col = "#aab8cc" if i % 5 == 0 else "#ccd6e4"
            self.create_line(x, BY - 2, x, BY + BH + 2, fill=col, width=1)
            if i % 2 == 0:
                self.create_text(x, BY - 10, text=t.strftime("%H:%M"),
                                 font=("Arial", 7), fill=GRAY, anchor="center")

        self.create_line(w - 1, BY - 8, w - 1, BY + BH + 8, fill=ORANGE, width=2)


# ─────────────────────────────────────────────────────────────────────────────
#  Barre OF (sous le chronogramme)
# ─────────────────────────────────────────────────────────────────────────────
class OFBar(tk.Canvas):
    """Mini chronogramme affichant les numeros d'OF."""
    BAR_Y = 6
    BAR_H = 22

    def __init__(self, parent, app, **kw):
        kw.setdefault("height", 38)
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
        # Fond
        self.create_rectangle(0, BY, w, BY + BH, fill="#e8eef8", outline="")
        # Etiquette gauche
        self.create_text(4, BY + BH // 2, text="N° OF",
                         font=("Arial", 7, "bold"), fill=GRAY, anchor="w")

        for p in self.app._of_periods:
            x1 = px(p["start"])
            x2 = px(p.get("end") or now)
            if x2 - x1 < 2:
                continue
            of_num = p.get("of_num", "")
            fill = NAVY_L if of_num else LGRAY
            self.create_rectangle(x1, BY, x2, BY + BH, fill=fill, outline=WHITE, width=1)
            label = of_num if of_num else "En cours"
            if x2 - x1 > 40:
                self.create_text((x1 + x2) / 2, BY + BH // 2,
                                 text=label, fill=WHITE,
                                 font=("Arial", 8, "bold"),
                                 anchor="center")


# ─────────────────────────────────────────────────────────────────────────────
#  Gauge TRS
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
            self._ask_stop_description()
        else:
            self.app._t_start(self.key)
            self.app._tl_open(self.key, "ratt" if self.accent == C_RATT else "pb")
            self._flash(14, 14)

    def _ask_stop_description(self):
        """Popup plein ecran pour saisir la description de l'arret."""
        top = tk.Toplevel(self.app.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=NAVY)
        sw = self.app.root.winfo_screenwidth()
        sh = self.app.root.winfo_screenheight()
        pw, ph = min(700, sw - 60), 360
        top.geometry(f"{pw}x{ph}+{(sw-pw)//2}+{(sh-ph)//2}")

        tk.Label(top, text=f"⚠  FIN D'ARRET  —  {self.label}",
                 bg=NAVY, fg=WHITE, font=("Arial", 18, "bold")).pack(pady=(28, 4))
        tk.Label(top, text="Décrivez la cause de l'arret :",
                 bg=NAVY, fg="#7a99c0", font=("Arial", 12)).pack()

        txt = tk.Text(top, height=5, font=("Arial", 13), bg=WHITE, fg=DARK,
                      insertbackground=DARK, relief="flat", padx=12, pady=8,
                      wrap="word")
        txt.pack(fill="x", padx=28, pady=12)
        txt.focus()

        def _valider():
            desc = txt.get("1.0", "end").strip()
            self.app._t_stop(self.key)
            self.app._tl_close(self.key, comment=desc)
            top.destroy()
            self._draw()

        tk.Button(top, text="✔   VALIDER", command=_valider,
                  bg=GREEN, fg=WHITE, font=("Arial", 16, "bold"),
                  relief="flat", pady=12, cursor="hand2").pack(
                  fill="x", padx=28, pady=(0, 20))
        top.bind("<Return>", lambda e: _valider())

    def _flash(self, n, max_n=14):
        if n <= 0:
            self._draw()
            return
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        progress = (max_n - n) / max_n  # 0→1

        # Croissance: margin diminue de 12 a 0
        margin = max(0, int(12 * (1 - progress)))
        bright = n % 2 == 0
        fill_col = "#ff2020" if bright else "#aa0000"
        halo_col = "#ff6060" if bright else "#cc2020"

        # Halos exterieurs
        for ring in range(3, 0, -1):
            rm = margin + ring * 5
            self.create_rectangle(
                max(0, rm), max(0, rm), max(0, w - rm), max(0, h - rm),
                fill=halo_col, outline="")

        # Corps
        _rrect(self, margin+3, margin+4, w-margin-1, h-margin,
               10, fill=_off(fill_col, -50))
        _rrect(self, margin, margin, w-margin-4, h-margin-4,
               10, fill=fill_col)
        _rrect(self, margin+2, margin+2, w-margin-6, max(margin+22, h//3),
               10, fill=_off(fill_col, +60))

        # Point clignotant (girofare)
        dot_col = "#ffff00" if bright else "#ff8800"
        self.create_oval(w-margin-20, margin+4, w-margin-8, margin+16,
                         fill=dot_col, outline=WHITE, width=1)

        self.create_text(w//2, h//2, text=self.label, fill=WHITE,
                         font=("Arial", 12, "bold"),
                         justify="center", width=w-12)
        self.after(100, self._flash, n - 1, max_n)

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
            blink = getattr(self.app, "_cell_blink", False)
            face  = "#ff2020" if blink else "#cc0000"
            shad  = _off(face, -50)
            hi    = _off(face, +60)
            _rrect(self, 4, 5, w-1, h,     r, fill=shad)
            _rrect(self, 0, 0, w-5, h-5,   r, fill=face)
            _rrect(self, 2, 2, w-7, max(r*2+2, h//3), r, fill=hi)
            self.create_text(w//2-2, h*2//5, text=self.label, fill=WHITE,
                             font=("Arial", 11, "bold"), justify="center", width=w-12)
            self.create_text(w//2-2, h*3//4, text=fmt(elapsed), fill=WHITE,
                             font=("Arial", 15, "bold"))
            # Dot clignotant
            dot_col = "#ffff00" if blink else "#ff8080"
            self.create_oval(w-17, 7, w-9, 15, fill=dot_col, outline=WHITE, width=1)
        else:
            _rrect(self, 4, 5, w-1, h,     r, fill=SHAD)
            _rrect(self, 0, 0, w-5, h-5,   r, fill=WHITE)
            self.create_rectangle(3, r+2, 8, h-r-7, fill=self.accent, outline="")
            self.create_text(w//2+2, h*2//5, text=self.label, fill=DARK,
                             font=("Arial", 11, "bold"), justify="center", width=w-20)
            if elapsed > 0:
                self.create_text(w//2+2, h*3//4, text=fmt(elapsed),
                                 fill=ORANGE, font=("Arial", 13, "bold"))
            else:
                self.create_text(w//2+2, h*3//4, text="- - -",
                                 fill=LGRAY, font=("Arial", 10))


def _app_dir():
    """Dossier de l'executable (ou du script en dev)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _load_logo_image(target_h=52):
    """Charge logo.jpg/jpeg/png/gif depuis le dossier de l'app.
    Utilise Pillow si disponible (JPEG), sinon PhotoImage natif (PNG/GIF).
    Retourne un tk.PhotoImage ou None."""
    folder = _app_dir()
    candidates = ["logo.jpg", "logo.jpeg", "logo.png", "logo.gif"]
    for name in candidates:
        path = os.path.join(folder, name)
        if not os.path.exists(path):
            continue
        # Essai avec Pillow (supporte JPEG + tous formats)
        try:
            from PIL import Image, ImageTk
            im = Image.open(path)
            # Redimensionner en conservant le ratio pour une hauteur cible
            iw, ih = im.size
            if ih > 0:
                new_w = int(iw * target_h / ih)
                im = im.resize((new_w, target_h), Image.LANCZOS)
            return ImageTk.PhotoImage(im)
        except ImportError:
            pass
        except Exception:
            continue
        # Fallback natif Tkinter (PNG/GIF uniquement)
        if name.endswith((".png", ".gif")):
            try:
                img = tk.PhotoImage(file=path)
                iw, ih = img.width(), img.height()
                if ih > 0 and ih > target_h:
                    s = max(1, ih // target_h)
                    img = img.subsample(s, s)
                return img
            except Exception:
                pass
    return None


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

        self._logo_img    = _load_logo_image()
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
        self._prod_indicator = None
        self._main_tree   = None
        self._mode        = "main"
        self._tick_count  = 0

        self._of_periods  = []
        self._tl_events   = []
        self._of_changes  = []
        self._blink_dot   = None
        self._blink_state = False
        self._outer_frame  = None
        self._alarm_blink  = False
        self._last_activity = datetime.datetime.now()
        self._elapsed_lbl  = None
        self._stops_lbl    = None
        self._of_bar_widget = None
        self._cell_blink  = False
        self._hdr_frame   = None
        self._hdr_widgets = []
        self._kpi_canvas  = None
        self._prod_state  = "prod"

        self._load_lists()
        self._load_history_from_excel()
        self._show_main()

    # ── Historique depuis Excel (reconstruit la timeline au demarrage) ────────
    def _load_history_from_excel(self):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return

        def _pdt(date_s, time_s):
            try:
                return datetime.datetime.strptime(
                    f"{str(date_s).strip()[:10]} {str(time_s).strip()}",
                    "%d/%m/%Y %H:%M:%S")
            except Exception:
                return None

        today = datetime.date.today().strftime("%d/%m/%Y")
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            # Periodes OF d'aujourd'hui
            if "Data" in wb.sheetnames:
                ws = wb["Data"]
                min_r = 2
                if str(ws.cell(1, 1).value or "").strip() == DATA_HEADERS[0]:
                    min_r = 2
                for row in ws.iter_rows(min_row=min_r, values_only=True):
                    if not any(row):
                        continue
                    r = list(row) + [None] * 55
                    date_val = str(r[1] or "").strip()[:10]
                    if date_val != today:
                        continue
                    start_dt = _pdt(date_val, r[17])
                    end_dt   = _pdt(date_val, r[18])
                    if start_dt:
                        self._of_periods.append({
                            "start": start_dt, "end": end_dt,
                            "of_num": str(r[0] or "").strip()
                        })
                # Separateurs d'OF (a partir du 2e OF de la journee)
                for p in self._of_periods[1:]:
                    self._of_changes.append(p["start"])
            # Evenements d'aujourd'hui
            if "Evenements" in wb.sheetnames:
                ws_e = wb["Evenements"]
                for row in ws_e.iter_rows(min_row=2, values_only=True):
                    if not row or not row[0]:
                        continue
                    evt_type = str(row[0] or "").lower()
                    date_val = str(row[2] or "").strip()[:10]
                    if date_val != today:
                        continue
                    start_dt = _pdt(date_val, str(row[10] or ""))
                    end_dt   = _pdt(date_val, str(row[11] or ""))
                    if start_dt is None:
                        continue
                    if "rattrapage" in evt_type:
                        cat = "ratt"
                    elif "pb" in evt_type or "technique" in evt_type:
                        cat = "pb"
                    else:
                        continue
                    self._tl_events.append({
                        "key": f"_hist_{len(self._tl_events)}",
                        "cat": cat,
                        "start": start_dt,
                        "end": end_dt or start_dt,
                    })
            wb.close()
        except Exception:
            pass

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
            self._ensure_excel_headers(p)
            self._load_history_from_excel()
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
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")
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

    def _tl_close(self, key, comment=""):
        for ev in reversed(self._tl_events):
            if ev["key"] == key and ev["end"] is None:
                ev["end"] = datetime.datetime.now()
                ev["comment"] = comment
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
        self._tl_widget      = None
        self._cumul_lbl      = None
        self._prod_indicator = None
        self._main_tree      = None
        self._blink_dot      = None
        self._outer_frame    = None
        self._elapsed_lbl    = None
        self._stops_lbl      = None
        self._of_bar_widget  = None
        self._hdr_frame   = None
        self._hdr_widgets = []
        self._kpi_canvas  = None

    def _make_header(self, parent, title, subtitle=""):
        hdr = tk.Frame(parent, bg=NAVY, height=68)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=title, bg=NAVY, fg=WHITE,
                 font=("Arial", 18, "bold")).pack(side="left", padx=22)
        if subtitle:
            tk.Label(hdr, text=subtitle, bg=NAVY, fg="#7a99c0",
                     font=("Arial", 11)).pack(side="left", padx=4)
        # Zone droite : logo sur fond blanc + DB
        right_bar = tk.Frame(hdr, bg=NAVY)
        right_bar.pack(side="right", padx=12)
        self._db_widget(right_bar, NAVY).pack(side="right", padx=4)
        # Logo dans un cadre blanc arrondi
        logo_frame = tk.Frame(right_bar, bg=WHITE, padx=8, pady=6)
        logo_frame.pack(side="right", padx=(0, 10))
        if self._logo_img:
            tk.Label(logo_frame, image=self._logo_img, bg=WHITE).pack()
        else:
            tk.Label(logo_frame, text="DODO", bg=WHITE, fg=NAVY,
                     font=("Arial", 15, "bold")).pack()
        return hdr

    def _make_timeline(self, parent):
        zone = tk.Frame(parent, bg=WHITE)
        zone.pack(fill="x")
        tl = Timeline(zone, self, bg=WHITE)
        tl.pack(fill="x", padx=6, pady=(6, 2))

        # Barre OF
        of_bar = OFBar(zone, self, bg=WHITE)
        of_bar.pack(fill="x", padx=6, pady=(0, 2))
        self._of_bar_widget = of_bar

        # Légende sous le chronogramme
        leg = tk.Frame(zone, bg=WHITE)
        leg.pack(anchor="w", padx=10, pady=(0, 5))
        for lbl, col in [("Production", GREEN), ("Rattrapage", C_RATT),
                         ("PB Technique", C_RED), ("Changement OF", ORANGE)]:
            tk.Frame(leg, bg=col, width=14, height=14).pack(side="left", padx=(0, 4))
            tk.Label(leg, text=lbl, bg=WHITE, fg=GRAY,
                     font=("Arial", 8, "bold")).pack(side="left", padx=(0, 16))
        self._tl_widget = tl
        return tl

    def _make_tabs(self, parent, active):
        bar = tk.Frame(parent, bg=NAVY_L)
        bar.pack(fill="x")

        def _bind_recursive(widget, event, cb):
            widget.bind(event, cb)
            try:
                widget.config(cursor="hand2")
            except Exception:
                pass
            for child in widget.winfo_children():
                _bind_recursive(child, event, cb)

        def _tab(label, mode, action, enabled):
            is_active = (mode == active)
            bg  = WHITE   if is_active else NAVY_L
            fg  = NAVY    if is_active else (WHITE if enabled else "#4a5578")
            top_col = (GREEN if mode == "main" else C_RED) if is_active else NAVY_L
            f = tk.Frame(bar, bg=bg)
            f.pack(side="left")
            tk.Frame(f, bg=top_col, height=4).pack(fill="x")
            row = tk.Frame(f, bg=bg)
            row.pack(padx=26, pady=11)
            lbl = tk.Label(row, text=label, bg=bg, fg=fg,
                           font=("Arial", 11, "bold"))
            lbl.pack(side="left")
            if enabled and not is_active:
                _bind_recursive(f, "<Button-1>",
                                lambda e, a=action: self._transition(a))
            return row

        _tab("📊  Tableau de bord", "main", self._show_main, True)

        # Onglet Production
        is_prod_tab = (active == "production")
        prod_bg = WHITE if is_prod_tab else NAVY_L
        prod_fg = NAVY  if is_prod_tab else (WHITE if self._prod_active else "#4a5578")
        prod_f = tk.Frame(bar, bg=prod_bg)
        prod_f.pack(side="left")
        tk.Frame(prod_f, bg=C_RED if is_prod_tab else NAVY_L, height=4).pack(fill="x")
        prod_row = tk.Frame(prod_f, bg=prod_bg)
        prod_row.pack(padx=26, pady=11)
        if self._prod_active and not is_prod_tab:
            self._blink_dot = tk.Label(prod_row, text="● ", bg=prod_bg,
                                       fg=C_RED, font=("Arial", 11, "bold"))
            self._blink_dot.pack(side="left")
        tk.Label(prod_row,
                 text="⚡  Production en cours" if self._prod_active else "⚡  Production",
                 bg=prod_bg, fg=prod_fg,
                 font=("Arial", 11, "bold")).pack(side="left")
        if self._prod_active and not is_prod_tab:
            _bind_recursive(prod_f, "<Button-1>",
                            lambda e: self._transition(self._nav_to_production))
        return bar

    # ── Tick unifie (main + production) ──────────────────────────────────────
    def _tick(self):
        self._tick_count += 1
        now = datetime.datetime.now()
        self._cell_blink = not self._cell_blink

        # ── Etat arret / production ───────────────────────────────────────────
        any_running = any(self._t_running(k) for k in self._timers)
        new_state = "stop" if any_running else "prod"

        if self._mode == "production" and self._prod_active:
            of_s   = (now - self._of_start).total_seconds()
            stop_s = self._t_total_stops()

            # Couleur header selon etat
            if new_state == "stop":
                hdr_col = C_RED if self._cell_blink else "#8b0000"
            else:
                hdr_col = GREEN
            self._set_header_color(hdr_col)

            # Chrono OF
            try:
                size_of = 18 if new_state == "stop" else 30
                self._of_clk.config(text=fmt(of_s),
                                    font=("Arial", size_of, "bold"),
                                    fg=WHITE if new_state == "stop" else "#4ade80")
            except Exception:
                pass

            # Cumul arrets — plus grand pendant un arret
            try:
                size_c = 34 if new_state == "stop" else 15
                fg_c   = "#ffff00" if new_state == "stop" else "#f59e0b"
                self._cumul_lbl.config(
                    text=f"⏸ ARRETS: {fmt(stop_s)}",
                    font=("Arial", size_c, "bold"), fg=fg_c)
            except Exception:
                pass

            # KPI % prod / arret
            try:
                if self._kpi_canvas and of_s > 0:
                    self._draw_kpi(of_s, stop_s)
            except Exception:
                pass

            # Indicateur arret flagrant dans le titre
            try:
                title = "🔴  EN ARRET  🔴" if new_state == "stop" else "ORC1  —  PRODUCTION EN COURS"
                self._hdr_title.config(text=title,
                    fg=WHITE if new_state == "stop" else WHITE,
                    font=("Arial", 14 if new_state == "stop" else 13, "bold"))
            except Exception:
                pass

            for cell in self._cells:
                try:
                    cell.refresh()
                except Exception:
                    pass

        # Mise a jour labels vue principale
        if self._mode == "main" and self._prod_active:
            of_s = (now - self._of_start).total_seconds()
            try:
                if self._elapsed_lbl:
                    self._elapsed_lbl.config(text=f"⏱  {fmt(of_s)}")
            except Exception:
                pass
            try:
                if self._stops_lbl:
                    n_a = sum(1 for k in self._timers if self._t_running(k))
                    col = C_RED if n_a > 0 else "#7a99c0"
                    self._stops_lbl.config(
                        text=f"⚠  Arrets actifs: {n_a}  |  Cumul: {fmt(self._t_total_stops())}",
                        fg=col)
            except Exception:
                pass

        # Auto-retour en production apres 30s d'inactivite
        if (self._mode == "main" and self._prod_active
                and (now - self._last_activity).total_seconds() > 30):
            self._last_activity = now
            self._transition(self._nav_to_production)
            return

        # Clignotement du point rouge sur l'onglet Production (vue principale)
        if self._mode == "main" and self._prod_active:
            self._blink_state = not self._blink_state
            try:
                self._blink_dot.config(
                    fg=C_RED if self._blink_state else NAVY_L)
            except Exception:
                pass

        if self._tl_widget:
            try:
                self._tl_widget.redraw()
            except Exception:
                pass
        try:
            if self._of_bar_widget:
                self._of_bar_widget.redraw()
        except Exception:
            pass

        self._after_id = self.root.after(1000, self._tick)

    def _set_header_color(self, col):
        """Met a jour la couleur de fond du header et de la tab bar."""
        try:
            if self._hdr_frame:
                self._hdr_frame.config(bg=col)
        except Exception:
            pass
        for w in self._hdr_widgets:
            try:
                w.config(bg=col)
            except Exception:
                pass

    def _draw_kpi(self, of_s, stop_s):
        """Redessine la barre KPI prod/arret dans le header."""
        cv = self._kpi_canvas
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 10 or h < 4:
            return
        prod_s = max(0, of_s - stop_s)
        ratio  = prod_s / of_s if of_s > 0 else 1.0
        xp = int(w * ratio)
        # Fond arret
        cv.create_rectangle(0, 0, w, h, fill="#8b0000", outline="")
        # Zone prod
        cv.create_rectangle(0, 0, xp, h, fill="#1a8c4e", outline="")
        # Labels
        pct_p = int(ratio * 100)
        pct_a = 100 - pct_p
        if xp > 60:
            cv.create_text(xp // 2, h // 2,
                           text=f"PROD {pct_p}%",
                           font=("Arial", 8, "bold"), fill=WHITE, anchor="center")
        if w - xp > 60:
            cv.create_text(xp + (w - xp) // 2, h // 2,
                           text=f"ARRET {pct_a}%",
                           font=("Arial", 8, "bold"), fill=WHITE, anchor="center")

    def _reset_activity(self, event=None):
        self._last_activity = datetime.datetime.now()

    def _transition(self, fn):
        """Bref flash de transition avant d'appeler fn."""
        overlay = tk.Frame(self.root, bg=NAVY_L)
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.root.update_idletasks()
        def _go():
            try:
                overlay.destroy()
            except Exception:
                pass
            fn()
        self.root.after(120, _go)

    # ── Mot de passe ─────────────────────────────────────────────────────────
    def _check_password(self, action=""):
        top = tk.Toplevel(self.root)
        top.title("Mot de passe")
        top.geometry("320x160")
        top.resizable(False, False)
        top.grab_set()
        # Centrer
        top.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width()  - 320) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 160) // 2
        top.geometry(f"320x160+{x}+{y}")

        result = tk.BooleanVar(value=False)
        tk.Label(top, text=action or "Entrez le mot de passe",
                 font=("Arial", 10, "bold"), fg=DARK).pack(pady=(18, 4))
        err_lbl = tk.Label(top, text="", fg=C_RED, font=("Arial", 9))
        err_lbl.pack()
        var = tk.StringVar()
        e = tk.Entry(top, textvariable=var, show="*",
                     font=("Arial", 18), width=10, justify="center",
                     relief="solid", bd=2)
        e.pack(pady=4)
        e.focus()

        def confirm(ev=None):
            if var.get() == PASSWORD:
                result.set(True)
                top.destroy()
            else:
                err_lbl.config(text="Mot de passe incorrect")
                var.set("")

        e.bind("<Return>", confirm)
        tk.Button(top, text="Valider", command=confirm,
                  bg=NAVY, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=16, pady=4,
                  cursor="hand2").pack(pady=6)
        top.wait_window()
        return result.get()

    # =========================================================================
    #  ECRAN PRINCIPAL
    # =========================================================================
    def _show_main(self):
        self._cells = []
        self._clear()
        self._db_labels.clear()
        self._mode = "main"
        self._last_activity = datetime.datetime.now()

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)
        self._outer_frame = outer
        # Suivi d'activite pour l'auto-retour
        outer.bind("<Motion>",  self._reset_activity)
        outer.bind("<Button-1>", self._reset_activity)

        self._make_header(outer, "KPI-ORC", "Ligne ORC1")
        self._make_tabs(outer, "main")
        self._make_timeline(outer)

        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=12)

        top = tk.Frame(body, bg=BG)
        top.pack(fill="x", pady=(0, 12))

        # TRS gauge (gauche, taille fixe)
        trs_wrap, trs_inner = shadow_frame(top, bg=WHITE)
        trs_wrap.pack(side="left", fill="y", padx=(0, 14))
        tk.Label(trs_inner, text="TRS", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold")).pack(pady=(10, 0), padx=16)
        tk.Label(trs_inner, text="Taux de Rendement Synthetique",
                 bg=WHITE, fg=LGRAY, font=("Arial", 8)).pack()
        self._main_gauge = Gauge(trs_inner, bg=WHITE,
                                  width=290, height=130, highlightthickness=0)
        self._main_gauge.pack(padx=16, pady=(0, 10))
        self._refresh_main_kpi()

        # Zone droite — remplit tout l'espace restant
        right_zone = tk.Frame(top, bg=BG)
        right_zone.pack(side="left", fill="both", expand=True)

        if self._prod_active:
            # Grand panneau "Production en cours"
            prod_wrap, prod_inner = shadow_frame(right_zone, bg=NAVY)
            prod_wrap.pack(fill="both", expand=True)

            # Ligne titre avec point clignotant
            dot_row = tk.Frame(prod_inner, bg=NAVY)
            dot_row.pack(fill="x", padx=20, pady=(16, 4))
            self._blink_dot = tk.Label(dot_row, text="●", bg=NAVY,
                                       fg=C_RED, font=("Arial", 20))
            self._blink_dot.pack(side="left", padx=(0, 8))
            tk.Label(dot_row, text="PRODUCTION EN COURS",
                     bg=NAVY, fg=WHITE,
                     font=("Arial", 16, "bold")).pack(side="left")

            # Heure de debut
            tk.Label(prod_inner,
                     text=f"Debut : {self._of_start.strftime('%H:%M:%S')}",
                     bg=NAVY, fg="#7a99c0",
                     font=("Arial", 12)).pack(anchor="w", padx=20, pady=(0, 4))

            # Chrono elapsed (mis a jour par _tick)
            of_s_now = (datetime.datetime.now() - self._of_start).total_seconds()
            self._elapsed_lbl = tk.Label(prod_inner,
                                          text=f"⏱  {fmt(of_s_now)}",
                                          bg=NAVY, fg="#4ade80",
                                          font=("Arial", 30, "bold"))
            self._elapsed_lbl.pack(anchor="center", pady=(8, 6))

            # Arrets actifs + cumul
            n_actifs = sum(1 for k in self._timers if self._t_running(k))
            col_arr = C_RED if n_actifs > 0 else "#7a99c0"
            self._stops_lbl = tk.Label(prod_inner,
                text=f"⚠  Arrets actifs: {n_actifs}  |  Cumul: {fmt(self._t_total_stops())}",
                bg=NAVY, fg=col_arr, font=("Arial", 12, "bold"))
            self._stops_lbl.pack(anchor="center", padx=20, pady=(0, 14))
        else:
            # Gros bouton vert "Démarrer"
            btn_canvas = tk.Canvas(right_zone, bg=BG, highlightthickness=0)
            btn_canvas.pack(fill="both", expand=True, padx=4, pady=8)

            def _draw_btn(e=None):
                btn_canvas.delete("all")
                bw, bh = btn_canvas.winfo_width(), btn_canvas.winfo_height()
                if bw < 10 or bh < 10:
                    return
                _rrect(btn_canvas, 5, 7, bw-1, bh, 16, fill=_off(GREEN, -40))
                _rrect(btn_canvas, 0, 0, bw-6, bh-7, 16, fill=GREEN)
                _rrect(btn_canvas, 2, 2, bw-8, bh//3, 16, fill=_off(GREEN, +40))
                btn_canvas.create_text(bw//2-3, bh//2-3,
                                       text="▶   DEMARRER UNE PRODUCTION",
                                       fill=WHITE, font=("Arial", 18, "bold"))

            btn_canvas.bind("<Configure>", _draw_btn)
            btn_canvas.bind("<Button-1>",  lambda e: self._start_production())
            btn_canvas.config(cursor="hand2")

        # Tableau recap
        lbl_frame = tk.Frame(body, bg=BG)
        lbl_frame.pack(fill="x", pady=(0, 4))
        tk.Label(lbl_frame, text="15 Dernieres Declarations",
                 bg=BG, fg=DARK, font=("Arial", 11, "bold")).pack(side="left")

        tbl_wrap, tbl_inner = shadow_frame(body, bg=WHITE)
        tbl_wrap.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("KPI.Treeview",
                        background=WHITE, foreground=DARK,
                        fieldbackground=WHITE, rowheight=28,
                        font=("Arial", 10))
        style.configure("KPI.Treeview.Heading",
                        background=LGRAY, foreground=DARK,
                        font=("Arial", 10, "bold"), relief="flat")
        style.map("KPI.Treeview", background=[("selected", "#dbeafe")])

        cols = ("Date", "OF", "Pilote", "Poste", "Qte Fab",
                "Duree OF", "PB Tech.", "Rattrap.", "✏", "🗑")
        tree = ttk.Treeview(tbl_inner, columns=cols, show="headings",
                            height=13, style="KPI.Treeview")
        widths = {"Date": 90, "OF": 100, "Pilote": 160, "Poste": 90,
                  "Qte Fab": 65, "Duree OF": 80,
                  "PB Tech.": 90, "Rattrap.": 90, "✏": 36, "🗑": 36}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 80), anchor="center",
                        stretch=(c not in ("✏", "🗑")))

        sb_v = ttk.Scrollbar(tbl_inner, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb_v.set)
        tree.pack(side="left", fill="both", expand=True)
        sb_v.pack(side="right", fill="y")

        self._main_tree = tree
        self._col_ids   = cols
        self._load_table(tree)

        tree.bind("<Button-1>", self._on_tree_click)

        self._after_id = self.root.after(1000, self._tick)

    def _nav_to_production(self):
        self._clear()
        self._db_labels.clear()
        self._show_production()

    def _refresh_table(self):
        if self._main_tree:
            for item in self._main_tree.get_children():
                self._main_tree.delete(item)
            self._load_table(self._main_tree)

    def _on_tree_click(self, event):
        region = self._main_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id  = self._main_tree.identify_column(event.x)
        item    = self._main_tree.identify_row(event.y)
        if not item:
            return
        col_num  = int(col_id.lstrip("#")) - 1
        col_name = self._col_ids[col_num] if col_num < len(self._col_ids) else ""
        excel_row = int(item)
        if col_name == "🗑":
            self._delete_declaration(excel_row)
        elif col_name == "✏":
            self._edit_declaration(excel_row)

    def _refresh_main_kpi(self):
        path = self.cfg.get("db_path", "")
        last_time = "--:--"
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
        self._main_gauge.update_gauge(0.0, last_time)

    def _load_table(self, tree):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb["Data"]
            # Sauter la ligne d'en-tete si presente
            min_r = 2
            first = ws.cell(1, 1).value
            if first and str(first).strip().upper() == "OF":
                min_r = 2
            rows_raw = list(ws.iter_rows(min_row=min_r, values_only=True))
            wb.close()
        except Exception:
            return

        # Conserver (excel_row, data)
        indexed = []
        for i, r in enumerate(rows_raw, start=min_r):
            if any(r):
                indexed.append((i, list(r) + [None] * 55))

        def _sd(row, idx):
            t = 0
            for i in idx:
                if i < len(row) and row[i]:
                    try:
                        t += _hms_to_sec(str(row[i]))
                    except Exception:
                        pass
            return t

        for excel_row, row in list(reversed(indexed))[:15]:
            tree.insert("", "end", iid=str(excel_row), values=(
                str(row[1])[:10] if row[1]  else "",
                str(row[0])      if row[0]  else "",
                str(row[3])      if row[3]  else "",
                str(row[2])      if row[2]  else "",
                str(row[13])     if row[13] else "0",
                str(row[16])     if row[16] else "",
                fmt(_sd(row, range(31, 49))),
                fmt(_sd(row, range(26, 31))),
                "✏", "🗑",
            ))

    # ── Suppression ──────────────────────────────────────────────────────────
    def _delete_declaration(self, excel_row):
        if not self._check_password("Supprimer la declaration"):
            return
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path)
            wb["Data"].delete_rows(excel_row)
            wb.save(path)
            wb.close()
            self._refresh_table()
            messagebox.showinfo("Supprime", "Declaration supprimee.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Suppression impossible:\n{e}")

    # ── Edition declaration ───────────────────────────────────────────────────
    def _edit_declaration(self, excel_row):
        if not self._check_password("Modifier la declaration"):
            return
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            messagebox.showwarning("Attention", "Base de donnees non connectee.")
            return
        try:
            wb       = load_workbook(path, data_only=True)
            ws       = wb["Data"]
            row_data = [ws.cell(row=excel_row, column=i).value
                        for i in range(1, len(DATA_HEADERS) + 1)]
            row_data += [None] * max(0, len(DATA_HEADERS) - len(row_data))
            of_num   = str(row_data[0] or "")
            of_date  = str(row_data[1] or "")
            evt_rows = []
            if "Evenements" in wb.sheetnames:
                for r in wb["Evenements"].iter_rows(min_row=2, values_only=True):
                    if r and str(r[1] or "") == of_num and str(r[2] or "") == of_date:
                        evt_rows.append(list(r))
            wb.close()
        except Exception as e:
            messagebox.showerror("Erreur", f"Lecture Excel:\n{e}")
            return

        self._open_edit_dialog(excel_row, row_data, evt_rows)

    def _open_edit_dialog(self, excel_row, row_data, evt_rows):
        of_lbl = str(row_data[0] or "?")
        top = tk.Toplevel(self.root)
        top.title(f"Modifier OF {of_lbl}")
        top.geometry("960x680")
        top.grab_set()
        top.resizable(True, True)
        top.update_idletasks()
        x = self.root.winfo_x() + max(0, (self.root.winfo_width()  - 960) // 2)
        y = self.root.winfo_y() + max(0, (self.root.winfo_height() - 680) // 2)
        top.geometry(f"960x680+{x}+{y}")

        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # ── Tab 1 : Donnees OF ────────────────────────────────────────────────
        tab1 = tk.Frame(nb, bg=WHITE)
        nb.add(tab1, text="  Donnees OF  ")

        canv1 = tk.Canvas(tab1, bg=WHITE, highlightthickness=0)
        sb1   = ttk.Scrollbar(tab1, orient="vertical", command=canv1.yview)
        canv1.configure(yscrollcommand=sb1.set)
        sb1.pack(side="right", fill="y")
        canv1.pack(fill="both", expand=True)
        inner1 = tk.Frame(canv1, bg=WHITE)
        win_id = canv1.create_window((0, 0), window=inner1, anchor="nw")
        inner1.bind("<Configure>",
                    lambda e: canv1.configure(scrollregion=canv1.bbox("all")))
        canv1.bind("<Configure>",
                   lambda e: canv1.itemconfig(win_id, width=e.width))
        canv1.bind_all("<MouseWheel>",
                       lambda e: canv1.yview_scroll(-1*(e.delta//120), "units"))

        field_vars = []
        for i, hdr in enumerate(DATA_HEADERS):
            val = row_data[i] if i < len(row_data) else None
            row_f = tk.Frame(inner1, bg=WHITE)
            row_f.pack(fill="x", padx=12, pady=2)
            row_f.columnconfigure(1, weight=1)
            tk.Label(row_f, text=hdr, bg=WHITE, fg=GRAY,
                     font=("Arial", 9), width=24, anchor="w").grid(
                row=0, column=0, sticky="w")
            var = tk.StringVar(value=str(val) if val is not None else "")
            e = tk.Entry(row_f, textvariable=var, bg=WHITE, fg=DARK,
                         font=("Arial", 10), relief="solid", bd=1)
            e.grid(row=0, column=1, sticky="ew", ipady=2, padx=(6, 0))
            field_vars.append(var)

        # ── Tab 2 : Evenements ────────────────────────────────────────────────
        tab2 = tk.Frame(nb, bg=WHITE)
        nb.add(tab2, text="  Evenements  ")

        evt_data = [list(r) for r in evt_rows]

        evt_cols = ("Type d'evenement", "H. Debut", "H. Fin", "Duree")
        evt_tree = ttk.Treeview(tab2, columns=evt_cols, show="headings",
                                height=18, style="KPI.Treeview")
        evt_tree.column("Type d'evenement", width=320, anchor="w")
        evt_tree.column("H. Debut",          width=90,  anchor="center")
        evt_tree.column("H. Fin",            width=90,  anchor="center")
        evt_tree.column("Duree",             width=90,  anchor="center")
        for c in evt_cols:
            evt_tree.heading(c, text=c)
        sb2 = ttk.Scrollbar(tab2, orient="vertical", command=evt_tree.yview)
        evt_tree.configure(yscrollcommand=sb2.set)
        evt_tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        sb2.pack(side="left", fill="y", pady=10)

        def _refresh_evt():
            for item in evt_tree.get_children():
                evt_tree.delete(item)
            for i, r in enumerate(evt_data):
                while len(r) < 13:
                    r.append("")
                evt_tree.insert("", "end", iid=str(i),
                                values=(r[0] or "", r[10] or "",
                                        r[11] or "", r[12] or ""))

        _refresh_evt()

        btn_bar = tk.Frame(tab2, bg=WHITE)
        btn_bar.pack(fill="y", side="right", padx=8, pady=10)

        def _add_evt():
            self._evt_edit_popup(top, evt_data, row_data, None, _refresh_evt)

        def _edit_evt():
            sel = evt_tree.selection()
            if not sel:
                return
            self._evt_edit_popup(top, evt_data, row_data, int(sel[0]), _refresh_evt)

        def _del_evt():
            sel = evt_tree.selection()
            if not sel:
                return
            if messagebox.askyesno("Confirmer",
                                   "Supprimer cet evenement ?", parent=top):
                evt_data.pop(int(sel[0]))
                _refresh_evt()

        for txt, cmd, col in [
            ("+ Ajouter",    _add_evt,  GREEN),
            ("✏ Modifier",   _edit_evt, ORANGE),
            ("🗑 Supprimer", _del_evt,  C_RED),
        ]:
            tk.Button(btn_bar, text=txt, command=cmd,
                      bg=col, fg=WHITE, font=("Arial", 10, "bold"),
                      relief="flat", padx=8, pady=6, cursor="hand2",
                      width=14).pack(pady=5)

        # ── Barre de sauvegarde ───────────────────────────────────────────────
        btm = tk.Frame(top, bg=BG)
        btm.pack(fill="x", padx=10, pady=(0, 8))

        def _save():
            path2 = self.cfg.get("db_path", "")
            if not path2:
                messagebox.showwarning("Attention", "Base de donnees non connectee.", parent=top)
                return
            try:
                wb2 = load_workbook(path2)
                ws2 = wb2["Data"]
                for col_i, var in enumerate(field_vars, start=1):
                    ws2.cell(row=excel_row, column=col_i).value = var.get()
                # Mettre a jour onglet Evenements
                of_num2 = str(row_data[0] or "")
                of_date2 = str(row_data[1] or "")
                ws_e = self._ensure_events_sheet(wb2)
                keep = []
                for r in ws_e.iter_rows(min_row=2, values_only=True):
                    if r and not (str(r[1] or "") == of_num2 and
                                  str(r[2] or "") == of_date2):
                        keep.append(list(r))
                for row_idx in range(ws_e.max_row, 1, -1):
                    ws_e.delete_rows(row_idx)
                for r in keep:
                    ws_e.append(r)
                for r in evt_data:
                    ws_e.append(r)
                wb2.save(path2)
                wb2.close()
                messagebox.showinfo("Succes", "Modifications sauvegardees !", parent=top)
                top.destroy()
                self._refresh_table()
            except Exception as e:
                messagebox.showerror("Erreur", f"Sauvegarde impossible:\n{e}", parent=top)

        tk.Button(btm, text="Annuler", command=top.destroy,
                  bg=SHAD, fg=DARK, font=("Arial", 11),
                  relief="flat", padx=14, pady=5, cursor="hand2").pack(side="right", padx=4)
        tk.Button(btm, text="💾  Sauvegarder", command=_save,
                  bg=NAVY, fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", padx=20, pady=6, cursor="hand2").pack(side="right", padx=4)

    def _evt_edit_popup(self, parent, evt_data, row_data, idx, refresh_cb):
        """Ajouter ou modifier un evenement dans la liste."""
        existing = evt_data[idx] if idx is not None else None
        dlg = tk.Toplevel(parent)
        dlg.title("Ajouter un evenement" if existing is None else "Modifier l'evenement")
        dlg.geometry("480x230")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width()  - 480) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 230) // 2
        dlg.geometry(f"480x230+{x}+{y}")

        frm = tk.Frame(dlg, bg=WHITE)
        frm.pack(fill="both", expand=True, padx=16, pady=12)
        frm.columnconfigure(1, weight=1)

        def lbl(txt, r):
            tk.Label(frm, text=txt, bg=WHITE, fg=GRAY,
                     font=("Arial", 10), anchor="w").grid(
                row=r, column=0, sticky="w", pady=4, padx=(0, 10))

        lbl("Type d'evenement :", 0)
        lbl("Heure debut (HH:MM:SS) :", 1)
        lbl("Heure fin   (HH:MM:SS) :", 2)

        type_var = tk.StringVar(value=existing[0] if existing else "")
        hd_var   = tk.StringVar(value=existing[10] if existing and len(existing) > 10 else "")
        hf_var   = tk.StringVar(value=existing[11] if existing and len(existing) > 11 else "")

        ttk.Combobox(frm, textvariable=type_var,
                     values=ALL_EVENT_TYPES, state="readonly",
                     width=36).grid(row=0, column=1, sticky="ew", pady=4)
        tk.Entry(frm, textvariable=hd_var, font=("Arial", 11),
                 relief="solid", bd=1).grid(row=1, column=1, sticky="ew", pady=4)
        tk.Entry(frm, textvariable=hf_var, font=("Arial", 11),
                 relief="solid", bd=1).grid(row=2, column=1, sticky="ew", pady=4)

        def confirm():
            if not type_var.get():
                messagebox.showwarning("Attention", "Choisissez un type.", parent=dlg)
                return
            try:
                dur = fmt(max(0, _hms_to_sec(hf_var.get()) -
                              _hms_to_sec(hd_var.get())))
            except Exception:
                dur = "00:00:00"
            new_row = [
                type_var.get(),
                str(row_data[0] or ""),
                str(row_data[1] or ""),
                str(row_data[2] or ""),
                str(row_data[3] or ""),
                str(row_data[4] or ""),
                str(row_data[5] or ""),
                str(row_data[6] or ""),
                str(row_data[8] or ""),
                str(row_data[7] or ""),
                hd_var.get(),
                hf_var.get(),
                dur,
            ]
            if idx is None:
                evt_data.append(new_row)
            else:
                evt_data[idx] = new_row
            refresh_cb()
            dlg.destroy()

        tk.Button(frm, text="Valider", command=confirm,
                  bg=NAVY, fg=WHITE, font=("Arial", 11, "bold"),
                  relief="flat", padx=18, pady=4, cursor="hand2").grid(
            row=3, column=0, columnspan=2, pady=(10, 0))

    # =========================================================================
    #  ECRAN DE PRODUCTION
    # =========================================================================
    def _start_production(self):
        now = datetime.datetime.now()
        if self._last_of_end is not None:
            gap = (now - self._last_of_end).total_seconds()
            if gap > 30:
                h = int(gap // 3600)
                m = int((gap % 3600) // 60)
                s = int(gap % 60)
                ts = f"{h}h {m:02d}min" if h > 0 else f"{m}min {s:02d}s"
                if messagebox.askyesno(
                        "Changement d'OF",
                        f"Le dernier OF a ete termine il y a {ts}.\n\n"
                        "Voulez-vous declarer ce temps comme\n"
                        "\"Changement d'OF\" ?"):
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
        self._of_periods.append({"start": now, "end": None, "of_num": ""})
        self._show_production()

    def _show_production(self):
        self._clear()
        self._db_labels.clear()
        self._mode = "production"
        self._prod_state = "prod"

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)
        self._outer_frame = outer

        # ── En-tete vert (devient rouge pendant un arret) ────────────────────
        hdr = tk.Frame(outer, bg=GREEN)
        hdr.pack(fill="x")
        self._hdr_frame = hdr
        self._hdr_widgets = []

        # Ligne 1 : titre + debut + DB
        line1 = tk.Frame(hdr, bg=GREEN)
        line1.pack(fill="x", padx=10, pady=(6, 0))
        self._hdr_widgets.append(line1)

        self._hdr_title = tk.Label(line1, text="ORC1  —  PRODUCTION EN COURS",
                                    bg=GREEN, fg=WHITE, font=("Arial", 13, "bold"))
        self._hdr_title.pack(side="left")
        self._hdr_widgets.append(self._hdr_title)

        debut_lbl = tk.Label(line1,
                              text=f"  |  Debut : {self._of_start.strftime('%H:%M:%S')}",
                              bg=GREEN, fg="#d4f5d4", font=("Arial", 11))
        debut_lbl.pack(side="left")
        self._hdr_widgets.append(debut_lbl)

        db_w = self._db_widget(hdr, GREEN)
        db_w.pack(side="right", padx=10, pady=(6, 0))
        self._hdr_widgets.append(db_w)
        for w in db_w.winfo_children():
            self._hdr_widgets.append(w)

        # Ligne 2 : chronos
        line2 = tk.Frame(hdr, bg=GREEN)
        line2.pack(fill="x", padx=10, pady=(2, 4))
        self._hdr_widgets.append(line2)

        of_lbl = tk.Label(line2, text="⏱ OF :", bg=GREEN, fg="#d4f5d4",
                          font=("Arial", 11))
        of_lbl.pack(side="left")
        self._hdr_widgets.append(of_lbl)

        self._of_clk = tk.Label(line2, text="00:00:00",
                                 bg=GREEN, fg="#4ade80", font=("Arial", 30, "bold"))
        self._of_clk.pack(side="left", padx=(2, 20))
        self._hdr_widgets.append(self._of_clk)

        stop_lbl = tk.Label(line2, text="⏸ :", bg=GREEN, fg="#d4f5d4",
                             font=("Arial", 11))
        stop_lbl.pack(side="left")
        self._hdr_widgets.append(stop_lbl)

        self._cumul_lbl = tk.Label(line2, text="⏸ ARRETS: 00:00:00",
                                    bg=GREEN, fg="#f59e0b", font=("Arial", 15, "bold"))
        self._cumul_lbl.pack(side="left", padx=(2, 16))
        self._hdr_widgets.append(self._cumul_lbl)

        # Barre KPI
        self._kpi_canvas = tk.Canvas(line2, bg=GREEN, height=22, width=220,
                                      highlightthickness=1, highlightbackground="#ffffff44")
        self._kpi_canvas.pack(side="left", padx=8)
        self._hdr_widgets.append(self._kpi_canvas)

        self._make_tabs(outer, "production")
        self._make_timeline(outer)

        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=FORM_BG, width=480)
        left.pack(side="left", fill="both")
        left.pack_propagate(False)
        self._build_form(left)

        tk.Frame(body, bg=SHAD, width=2).pack(side="left", fill="y")

        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        self._build_events(right)

        # Bouton FIN (vert, en bas de la zone droite) — packed AFTER
        # _build_events so it appears at the bottom of the events panel.
        end_cv = tk.Canvas(right, bg=BG, highlightthickness=0, height=68)
        end_cv.pack(fill="x", padx=10, pady=(4, 8))

        def _draw_end(e=None):
            end_cv.delete("all")
            bw, bh = end_cv.winfo_width(), end_cv.winfo_height()
            if bw < 10 or bh < 10:
                return
            _rrect(end_cv, 5, 7, bw-1, bh, 14, fill=_off(GREEN, -50))
            _rrect(end_cv, 0, 0, bw-6, bh-7, 14, fill=GREEN)
            _rrect(end_cv, 2, 2, bw-8, bh//3, 14, fill=_off(GREEN, +40))
            end_cv.create_text(bw//2-3, bh//2-4,
                               text="⏹   DECLARER LA FIN DE PRODUCTION",
                               fill=WHITE, font=("Arial", 16, "bold"))

        end_cv.bind("<Configure>", _draw_end)
        end_cv.bind("<Button-1>",  lambda e: self._end_production())
        end_cv.config(cursor="hand2")

        self._after_id = self.root.after(1000, self._tick)

    # ── Formulaire avec scroll ────────────────────────────────────────────────
    def _build_form(self, parent):
        self.fv = {}
        tk.Label(parent, text="DONNEES DE L'OF",
                 bg=FORM_BG, fg=NAVY,
                 font=("Arial", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 2))

        # Canvas scrollable
        canv = tk.Canvas(parent, bg=FORM_BG, highlightthickness=0)
        sb   = ttk.Scrollbar(parent, orient="vertical", command=canv.yview)
        canv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canv.pack(fill="both", expand=True)
        c = tk.Frame(canv, bg=FORM_BG)
        win_id = canv.create_window((0, 0), window=c, anchor="nw")
        c.bind("<Configure>",
               lambda e: canv.configure(scrollregion=canv.bbox("all")))
        canv.bind("<Configure>",
                  lambda e: canv.itemconfig(win_id, width=e.width))
        canv.bind_all("<MouseWheel>",
                      lambda e: canv.yview_scroll(-1*(e.delta//120), "units"))

        c.columnconfigure(0, weight=1)
        c.columnconfigure(1, weight=1)
        ri = [0]

        def sec(txt):
            f = tk.Frame(c, bg=NAVY_L)
            f.grid(row=ri[0], column=0, columnspan=2,
                   sticky="ew", padx=0, pady=(8, 2))
            tk.Label(f, text=txt, bg=NAVY_L, fg=WHITE,
                     font=("Arial", 9, "bold"), padx=8, pady=3).pack(anchor="w")
            ri[0] += 1

        def fld(lbl_txt, key, ftype, lh=None, col=0, adv=True):
            cell = tk.Frame(c, bg=FORM_BG)
            cell.grid(row=ri[0], column=col, sticky="ew", padx=3, pady=2)
            cell.columnconfigure(0, weight=1)
            tk.Label(cell, text=lbl_txt, bg=FORM_BG, fg=GRAY,
                     font=("Arial", 9), anchor="w").grid(row=0, column=0, sticky="w")
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                e = tk.Entry(cell, textvariable=var, bg="#f0f4fb", fg=DARK,
                             font=("Arial", 10), relief="groove", bd=1,
                             insertbackground=DARK)
                e.grid(row=1, column=0, sticky="ew", ipady=3)
            else:
                cb = ttk.Combobox(cell, textvariable=var,
                                  values=self._get_list(lh) if lh else [],
                                  font=("Arial", 10), state="readonly")
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
        # Bouton KIT en haut de la section Produit
        self._v_kit = tk.BooleanVar()
        kit_cell = tk.Frame(c, bg=FORM_BG)
        kit_cell.grid(row=ri[0], column=0, columnspan=2,
                      sticky="w", padx=4, pady=(4, 2))
        ri[0] += 1

        def _toggle_kit():
            self._v_kit.set(not self._v_kit.get())
            kit_btn.config(
                bg="#1a5e8c" if self._v_kit.get() else LGRAY,
                fg=WHITE if self._v_kit.get() else DARK,
                relief="sunken" if self._v_kit.get() else "flat",
                text="✔  KIT 2 pieces (ACTIF)" if self._v_kit.get() else "KIT 2 pieces")

        kit_btn = tk.Button(kit_cell, text="KIT 2 pieces",
                            command=_toggle_kit, bg=LGRAY, fg=DARK,
                            font=("Arial", 10, "bold"), relief="flat",
                            padx=14, pady=5, cursor="hand2")
        kit_btn.pack(side="left")

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

        sec("── Commentaire")
        txt_cell = tk.Frame(c, bg=FORM_BG)
        txt_cell.grid(row=ri[0], column=0, columnspan=2,
                      sticky="ew", padx=3, pady=2)
        ri[0] += 1
        self._comment_txt = tk.Text(txt_cell, height=5, bg=WHITE, fg=DARK,
                                     font=("Arial", 10), relief="solid", bd=1,
                                     wrap="word", insertbackground=DARK)
        self._comment_txt.pack(fill="x")

    # ── Evenements ────────────────────────────────────────────────────────────
    def _build_events(self, parent):
        def sec_header(txt, color):
            f = tk.Frame(parent, bg=color, height=40)
            f.pack(fill="x", padx=10, pady=(8, 2))
            f.pack_propagate(False)
            tk.Label(f, text=txt, bg=color, fg=WHITE,
                     font=("Arial", 13, "bold")).pack(side="left", padx=12, pady=6)

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

        tk.Frame(parent, bg=SHAD, height=3).pack(fill="x", padx=10, pady=(4, 2))

        sec_header("⚠  PROBLEMES TECHNIQUES", C_RED)
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

    # =========================================================================
    #  FIN DE PRODUCTION
    # =========================================================================
    def _end_production(self):
        active = [k for k in self._timers if self._t_running(k)]
        if active:
            if not messagebox.askyesno(
                    "Attention",
                    f"Il y a {len(active)} arret(s) en cours.\n"
                    "Ils vont etre automatiquement arretes.\n\nContinuer ?"):
                return

        end_dt = datetime.datetime.now()
        self._t_stop_all()
        self._tl_close_all()
        self._prod_active = False
        self._last_of_end = end_dt

        of_s = (end_dt - self._of_start).total_seconds()
        v    = {k: var.get().strip() for k, var in self.fv.items()}
        v["comment"] = self._comment_txt.get("1.0", "end").strip()

        if self._of_periods:
            self._of_periods[-1]["end"] = end_dt
            self._of_periods[-1]["of_num"] = v.get("of_num", "")

        if not v.get("of_num"):
            if not messagebox.askyesno(
                    "Attention", "N° OF non saisi. Continuer quand meme ?"):
                self._prod_active = True
                self._after_id = self.root.after(1000, self._tick)
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
        c1      = round(qte_fab / of_min, 2)  if of_min  > 0 else 0
        c2      = round(qte_fab / (nb_pers * of_hrs), 2) if of_hrs > 0 else 0
        equiv   = self._calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
        kit     = 2 if self._v_kit.get() else 1

        def _ts(key):
            return fmt(self._t_get(key))

        row = [
            v.get("of_num",""),
            datetime.date.today().strftime("%d/%m/%Y"),
            v.get("poste",""),
            v.get("pilote",""),
            v.get("copilote",""),
            v.get("nb_pers",""),
            v.get("taille",""),
            v.get("code_prod",""),
            v.get("type_prod",""),
            v.get("poids",""),
            v.get("fibre",""),
            v.get("of_taie",""),
            v.get("traca",""),
            qte_fab,
            _n("qte_emb"),
            equiv,
            fmt(of_s),
            self._of_start.strftime("%H:%M:%S"),
            end_dt.strftime("%H:%M:%S"),
            c1, c2, kit,
            v.get("ref_taie",""),
            _n("nb_def_cout"),
            _n("mq_taie"),
            f"Housse:{v.get('mq_housse','')} Encart:{v.get('mq_encart','')}",
            _ts("ratt_pochon"),     _ts("ratt_couture"),
            _ts("ratt_emb"),        _ts("ratt_presse_soud"),
            _ts("ratt_presse_zip"),
            _ts("pb_chargeuse"),    _ts("pb_carde"),
            _ts("pb_etaleur"),      _ts("pb_coupe"),
            _ts("pb_tapis1"),       _ts("pb_enrouleur"),
            _ts("pb_pesee"),        _ts("pb_deviation"),
            _ts("pb_enfileur"),     _ts("pb_kinna"),
            _ts("pb_tapeuse"),      _ts("pb_table_rot"),
            _ts("pb_h100"),         _ts("pb_traversin"),
            _ts("pb_presse_orc"),   _ts("pb_presse_zip2"),
            _ts("pb_cercleuse"),    _ts("pb_enrouleuse"),
            v.get("comment",""),
        ]

        ok = self._write_excel(row, v)
        self._show_main()
        if ok:
            _toast(self.root, "✔  Production declaree avec succes !", bg=GREEN)
        else:
            messagebox.showerror("Erreur",
                                 "Impossible d'ecrire dans Excel.\n"
                                 "Verifiez que le fichier n'est pas ouvert.")

    def _calc_equiv(self, qte, taille, type_prod):
        for item in self._get_list("Equivalence"):
            try:
                return round(qte * float(str(item).replace(",", ".")), 2)
            except Exception:
                pass
        return qte

    # ── Excel ─────────────────────────────────────────────────────────────────
    def _ensure_excel_headers(self, path):
        if not path or not os.path.exists(path):
            return
        try:
            wb      = load_workbook(path)
            changed = False
            # Data sheet
            if "Data" in wb.sheetnames:
                ws = wb["Data"]
                if ws.cell(1, 1).value != DATA_HEADERS[0]:
                    ws.insert_rows(1)
                    for i, h in enumerate(DATA_HEADERS, start=1):
                        ws.cell(1, i).value = h
                    changed = True
            # Evenements sheet
            self._ensure_events_sheet(wb)
            changed = True
            if changed:
                wb.save(path)
            wb.close()
        except Exception:
            pass

    def _write_excel(self, row, v):
        path = self.cfg.get("db_path", "")
        if not path:
            messagebox.showwarning("Attention", "Aucune base de donnees !")
            return False
        try:
            wb = load_workbook(path)
            # S'assurer que les en-tetes existent
            ws = wb["Data"]
            if ws.cell(1, 1).value != DATA_HEADERS[0]:
                ws.insert_rows(1)
                for i, h in enumerate(DATA_HEADERS, start=1):
                    ws.cell(1, i).value = h
            ws.append(row)
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
            ws.append(EVT_HEADERS)
        else:
            ws = wb["Evenements"]
            if ws.cell(1, 1).value != EVT_HEADERS[0]:
                ws.insert_rows(1)
                for i, h in enumerate(EVT_HEADERS, start=1):
                    ws.cell(1, i).value = h
        return wb["Evenements"]

    def _write_events_to_wb(self, wb, v):
        ws       = self._ensure_events_sheet(wb)
        of_start = self._of_start
        for ev in self._tl_events:
            if ev.get("cat") not in ("ratt", "pb"):
                continue
            if not ev.get("key") or ev["key"].startswith("_"):
                continue
            if ev["start"] < of_start:
                continue
            cat_name = "Rattrapage" if ev["cat"] == "ratt" else "PB Technique"
            label    = next((e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
            start    = ev["start"]
            end      = ev.get("end") or datetime.datetime.now()
            dur      = (end - start).total_seconds()
            ws.append([
                f"{cat_name}: {label}",
                v.get("of_num", ""),
                start.strftime("%d/%m/%Y"),
                v.get("poste", ""),    v.get("pilote", ""),
                v.get("copilote", ""), v.get("nb_pers", ""),
                v.get("taille", ""),   v.get("type_prod", ""),
                v.get("code_prod", ""),
                start.strftime("%H:%M:%S"),
                end.strftime("%H:%M:%S"),
                fmt(dur),
                ev.get("comment", ""),
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
                "Changement d'OF", "", "", "", "", "", "", "", "", "",
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
