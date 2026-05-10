"""KPI-ORC v5.7 - Style Dodo (bleu marine #1a1f5e + rouge #e31e24)"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, sys, datetime, math
from openpyxl import load_workbook

CONFIG_FILE  = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")
SESSION_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_session.json")
PASSWORD     = "0000"
DB_PASSWORD  = "4594"

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

TIMELINE_WINDOW = 480  # minutes (8h)

EVENTS = [
    ("Pochon / Fibre",       "ratt_pochon",      "ratt"),
    ("Couture",              "ratt_couture",     "ratt"),
    ("Emballage",            "ratt_emb",         "ratt"),
    ("Presse Souder",        "ratt_presse_soud", "ratt"),
    ("Presse ZIP",           "ratt_presse_zip",  "ratt"),
    ("Nettoyage",            "nettoyage",        "ratt"),
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
    "Duree OF", "Heure Debut", "Heure Fin", "Cadence/min", "Cadence par heure",
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
    "Nb Personnes", "Taille", "Type Produit", "Code Produit", "Fibre",
    "Poids Garnissage", "OF Taie", "Traca Fibre", "Ref Taie", "Kit",
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

        # Lignes de tick toutes les 30 min, label seulement heures pile et demies
        for i in range(TIMELINE_WINDOW + 1):
            if i % 30 != 0:
                continue
            t = t0 + datetime.timedelta(minutes=i)
            x = i / TIMELINE_WINDOW * w
            is_hour = (t.minute == 0)
            col = "#aab8cc" if is_hour else "#ccd6e4"
            lw  = 2 if is_hour else 1
            self.create_line(x, BY - 4, x, BY + BH + 4, fill=col, width=lw)
            self.create_text(x, BY - 10, text=t.strftime("%H:%M"),
                             font=("Arial", 8 if is_hour else 7, "bold" if is_hour else "normal"),
                             fill=GRAY if not is_hour else DARK, anchor="center")

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

        # Segments entre les OF
        periods_sorted = sorted(self.app._of_periods, key=lambda p: p["start"])
        for idx, p in enumerate(periods_sorted):
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
                                 font=("Arial", 8, "bold"), anchor="center")
            # Intervalle apres cet OF
            if idx + 1 < len(periods_sorted):
                gap_start = p.get("end") or now
                gap_end   = periods_sorted[idx + 1]["start"]
                gx1, gx2  = px(gap_start), px(gap_end)
                if gx2 - gx1 > 2:
                    self.create_rectangle(gx1, BY, gx2, BY + BH,
                                          fill="#ccd6e4", outline=WHITE, width=1)
                    if gx2 - gx1 > 40:
                        self.create_text((gx1 + gx2) / 2, BY + BH // 2,
                                         text="— Entre OF —", fill=GRAY,
                                         font=("Arial", 7), anchor="center")


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
        self._status_cv   = None
        self._stop_timer_lbls = {}
        self._active_stops_container = None
        self._main_prod_panel = None
        self._recap_panel     = None
        self._saved_form_data = {}   # Mémoire formulaire entre onglets

        self._load_lists()
        self._load_history_from_excel()

        # ── Vérifier si une session était en cours ──────────────────────────
        if self._try_restore_session():
            return  # Session restaurée, _show_production() déjà appelé
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
        # Demande le mot de passe DB
        top = tk.Toplevel(self.root)
        top.title("Acces base de donnees")
        top.geometry("320x160")
        top.resizable(False, False)
        top.grab_set()
        top.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width()  - 320) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 160) // 2
        top.geometry(f"320x160+{x}+{y}")
        allowed = tk.BooleanVar(value=False)
        tk.Label(top, text="Mot de passe base de donnees",
                 font=("Arial", 10, "bold"), fg=DARK).pack(pady=(18, 4))
        err2 = tk.Label(top, text="", fg=C_RED, font=("Arial", 9))
        err2.pack()
        pv = tk.StringVar()
        pe = tk.Entry(top, textvariable=pv, show="*",
                      font=("Arial", 18), width=10, justify="center",
                      relief="solid", bd=2)
        pe.pack(pady=4)
        pe.focus()
        def _confirm_db(ev=None):
            if pv.get() == DB_PASSWORD:
                allowed.set(True)
                top.destroy()
            else:
                err2.config(text="Mot de passe incorrect")
                pv.set("")
        pe.bind("<Return>", _confirm_db)
        tk.Button(top, text="Valider", command=_confirm_db,
                  bg=NAVY, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=16, pady=4,
                  cursor="hand2").pack(pady=6)
        top.wait_window()
        if not allowed.get():
            return
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
        tk.Button(f, text="⚙", command=self._select_db,
                  bg=NAVY_L, fg=WHITE, font=("Arial", 14),
                  relief="flat", padx=10, pady=2, cursor="hand2").pack(side="left")
        return f

    def _get_list(self, h):
        return self.lists.get(h, [])

    # ── Persistance session ───────────────────────────────────────────────────
    @staticmethod
    def _dt_str(dt):
        return dt.isoformat() if dt else None

    @staticmethod
    def _str_dt(s):
        return datetime.datetime.fromisoformat(s) if s else None

    def _save_session(self):
        """Sauvegarde l'état complet de la production en cours dans un fichier JSON."""
        if not self._prod_active:
            # Pas de prod active : supprimer le fichier si présent
            try:
                os.remove(SESSION_FILE)
            except FileNotFoundError:
                pass
            return
        try:
            # Timers : convertir les datetimes en string
            timers_serial = {}
            for k, t in self._timers.items():
                timers_serial[k] = {
                    "elapsed": t["elapsed"],
                    "running": t["running"],
                    "start":   self._dt_str(t.get("start")),
                }

            # Events : convertir les datetimes
            events_serial = []
            for ev in self._tl_events:
                events_serial.append({
                    "key":     ev["key"],
                    "cat":     ev["cat"],
                    "start":   self._dt_str(ev["start"]),
                    "end":     self._dt_str(ev.get("end")),
                    "comment": ev.get("comment", ""),
                })

            # Périodes OF
            periods_serial = []
            for p in self._of_periods:
                periods_serial.append({
                    "start":  self._dt_str(p["start"]),
                    "end":    self._dt_str(p.get("end")),
                    "of_num": p.get("of_num", ""),
                })

            # Changements OF
            changes_serial = [self._dt_str(d) for d in self._of_changes]

            data = {
                "of_start":       self._dt_str(self._of_start),
                "last_of_end":    self._dt_str(self._last_of_end),
                "timers":         timers_serial,
                "tl_events":      events_serial,
                "of_periods":     periods_serial,
                "of_changes":     changes_serial,
                "form_data":      self._saved_form_data,
            }
            with open(SESSION_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_session(self):
        """Recharge l'état sauvegardé. Retourne True si une session a été restaurée."""
        if not os.path.exists(SESSION_FILE):
            return False
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            if not data.get("of_start"):
                return False

            self._of_start    = self._str_dt(data["of_start"])
            self._last_of_end = self._str_dt(data.get("last_of_end"))
            self._prod_active = True

            # Timers
            self._timers = {}
            for k, t in data.get("timers", {}).items():
                self._timers[k] = {
                    "elapsed": float(t["elapsed"]),
                    "running": bool(t["running"]),
                    "start":   self._str_dt(t.get("start")),
                }

            # Events timeline
            self._tl_events = []
            for ev in data.get("tl_events", []):
                self._tl_events.append({
                    "key":     ev["key"],
                    "cat":     ev["cat"],
                    "start":   self._str_dt(ev["start"]),
                    "end":     self._str_dt(ev.get("end")),
                    "comment": ev.get("comment", ""),
                })

            # Périodes OF
            self._of_periods = []
            for p in data.get("of_periods", []):
                self._of_periods.append({
                    "start":  self._str_dt(p["start"]),
                    "end":    self._str_dt(p.get("end")),
                    "of_num": p.get("of_num", ""),
                })

            # Changements OF
            self._of_changes = [self._str_dt(d) for d in data.get("of_changes", [])
                                 if d]

            # Formulaire
            self._saved_form_data = data.get("form_data", {})

            return True
        except Exception:
            return False

    def _delete_session(self):
        try:
            os.remove(SESSION_FILE)
        except FileNotFoundError:
            pass

    def _try_restore_session(self):
        """Si un fichier session existe, propose de reprendre. Retourne True si restauré."""
        if not os.path.exists(SESSION_FILE):
            return False

        # Lire juste l'heure de début pour l'afficher dans la demande
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            of_start_str = data.get("of_start", "")
            if not of_start_str:
                return False
            dt = datetime.datetime.fromisoformat(of_start_str)
            debut_fmt = dt.strftime("%H:%M le %d/%m/%Y")
        except Exception:
            return False

        # Pop-up plein écran de reprise
        resume = [False]
        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=NAVY)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        pw, ph = min(560, sw - 60), 260
        top.geometry(f"{pw}x{ph}+{(sw-pw)//2}+{(sh-ph)//2}")

        tk.Label(top, text="⚠  SESSION INTERROMPUE",
                 bg=NAVY, fg=C_RED, font=("Arial", 20, "bold")).pack(pady=(30, 6))
        tk.Label(top,
                 text=f"L'application a été quittée pendant une production\n"
                      f"démarrée à {debut_fmt}.\n\n"
                      "Voulez-vous reprendre l'encours ?",
                 bg=NAVY, fg=WHITE, font=("Arial", 12),
                 justify="center").pack(pady=(0, 20))

        btn_f = tk.Frame(top, bg=NAVY)
        btn_f.pack()

        def _oui():
            resume[0] = True
            top.destroy()

        def _non():
            resume[0] = False
            top.destroy()

        tk.Button(btn_f, text="✔  OUI, REPRENDRE",
                  command=_oui, bg=GREEN, fg=WHITE,
                  font=("Arial", 13, "bold"), relief="flat",
                  padx=24, pady=10, cursor="hand2").pack(side="left", padx=8)
        tk.Button(btn_f, text="✕  NON, ABANDONNER",
                  command=_non, bg=C_RED, fg=WHITE,
                  font=("Arial", 13, "bold"), relief="flat",
                  padx=24, pady=10, cursor="hand2").pack(side="left", padx=8)

        top.grab_set()
        top.wait_window()

        if not resume[0]:
            self._delete_session()
            return False

        # Restaurer la session
        if self._load_session():
            self._show_production()
            _toast(self.root, "Session reprise !", bg=GREEN, duration=2500)
            return True

        self._delete_session()
        return False

    # ── Timers ────────────────────────────────────────────────────────────────
    def _t_start(self, key):
        t = self._timers.setdefault(
            key, {"start": None, "elapsed": 0.0, "running": False})
        if not t["running"]:
            t["start"]   = datetime.datetime.now()
            t["running"] = True
        self._save_session()

    def _t_stop(self, key):
        t = self._timers.get(key)
        if t and t["running"]:
            t["elapsed"] += (datetime.datetime.now() - t["start"]).total_seconds()
            t["running"]  = False
            t["start"]    = None
        self._save_session()

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

    def _t_wall_clock_stops(self):
        """Temps reel perdu (union des intervalles, sans double-comptage paralleles)."""
        if not self._of_start:
            return 0.0
        now = datetime.datetime.now()
        intervals = []
        for ev in self._tl_events:
            if ev.get("cat") not in ("ratt", "pb"):
                continue
            if ev["start"] < self._of_start:
                continue
            intervals.append((ev["start"], ev.get("end") or now))
        # Arrets courants pas encore dans tl_events (deja ouverts)
        if not intervals:
            return 0.0
        intervals.sort(key=lambda x: x[0])
        merged, cs, ce = [], intervals[0][0], intervals[0][1]
        for s, e in intervals[1:]:
            if s <= ce:
                ce = max(ce, e)
            else:
                merged.append((cs, ce))
                cs, ce = s, e
        merged.append((cs, ce))
        return sum((e - s).total_seconds() for s, e in merged)

    # ── Timeline history ──────────────────────────────────────────────────────
    def _tl_open(self, key, cat):
        self._tl_events.append({
            "key": key, "cat": cat,
            "start": datetime.datetime.now(), "end": None
        })
        self._save_session()
        self._refresh_stops_recap()

    def _tl_close(self, key, comment=""):
        for ev in reversed(self._tl_events):
            if ev["key"] == key and ev["end"] is None:
                ev["end"] = datetime.datetime.now()
                ev["comment"] = comment
                break
        self._save_session()
        self._refresh_stops_recap()

    def _tl_close_all(self):
        now = datetime.datetime.now()
        for ev in self._tl_events:
            if ev["end"] is None:
                ev["end"] = now
        self._save_session()
        self._refresh_stops_recap()

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
        self._status_cv   = None
        self._stop_timer_lbls = {}
        self._active_stops_container = None
        self._main_prod_panel = None

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
        return hdr

    def _make_timeline(self, parent):
        zone = tk.Frame(parent, bg=WHITE)
        zone.pack(fill="x")
        tl = Timeline(zone, self, bg=WHITE)
        tl.pack(fill="x", padx=6, pady=(4, 0))

        # Barre OF (collée au chronogramme principal)
        of_bar = OFBar(zone, self, bg=WHITE)
        of_bar.pack(fill="x", padx=6, pady=0)
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
        bar = tk.Frame(parent, bg=WHITE, bd=0)
        bar.pack(fill="x")
        tk.Frame(bar, bg=LGRAY, height=1).pack(fill="x", side="bottom")

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
            bg      = WHITE
            fg      = NAVY if is_active else (GRAY if enabled else LGRAY)
            accent  = (GREEN if mode == "main" else C_RED) if is_active else WHITE
            f = tk.Frame(bar, bg=bg)
            f.pack(side="left")
            tk.Frame(f, bg=accent, height=3).pack(fill="x")
            row = tk.Frame(f, bg=bg)
            row.pack(padx=22, pady=10)
            lbl = tk.Label(row, text=label, bg=bg, fg=fg,
                           font=("Arial", 11, "bold" if is_active else "normal"))
            lbl.pack(side="left")
            if enabled and not is_active:
                _bind_recursive(f, "<Button-1>",
                                lambda e, a=action: self._transition(a))
            return row

        _tab("📊  Tableau de bord", "main", self._show_main, True)

        # Onglet Production
        is_prod_tab = (active == "production")
        fg_prod = NAVY if is_prod_tab else (GRAY if self._prod_active else LGRAY)
        accent_prod = C_RED if is_prod_tab else WHITE
        prod_f = tk.Frame(bar, bg=WHITE)
        prod_f.pack(side="left")
        tk.Frame(prod_f, bg=accent_prod, height=3).pack(fill="x")
        prod_row = tk.Frame(prod_f, bg=WHITE)
        prod_row.pack(padx=22, pady=10)
        if self._prod_active and not is_prod_tab:
            self._blink_dot = tk.Label(prod_row, text="● ", bg=WHITE,
                                       fg=C_RED, font=("Arial", 11, "bold"))
            self._blink_dot.pack(side="left")
        tk.Label(prod_row,
                 text="⚡  Production en cours" if self._prod_active else "⚡  Production",
                 bg=WHITE, fg=fg_prod,
                 font=("Arial", 11, "bold" if is_prod_tab else "normal")).pack(side="left")
        if self._prod_active and not is_prod_tab:
            _bind_recursive(prod_f, "<Button-1>",
                            lambda e: self._transition(self._nav_to_production))
        return bar

    # ── Tick unifie ──────────────────────────────────────────────────────────
    def _tick(self):
        self._tick_count += 1
        now = datetime.datetime.now()
        self._cell_blink = not self._cell_blink
        any_running = any(self._t_running(k) for k in self._timers)

        # ── Vue production ────────────────────────────────────────────────────
        if self._mode == "production" and self._prod_active:
            of_s   = (now - self._of_start).total_seconds()
            stop_s = self._t_wall_clock_stops()
            try:
                self._redraw_status(of_s, stop_s, any_running)
            except Exception:
                pass
            # Timers des cartes d'arret actifs
            for key, lbl in list(self._stop_timer_lbls.items()):
                try:
                    lbl.config(text=fmt(self._t_get(key)))
                except Exception:
                    self._stop_timer_lbls.pop(key, None)

        # ── Vue principale ────────────────────────────────────────────────────
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
                        text=f"⚠  Arrets actifs: {n_a}  |  Cumul: {fmt(self._t_wall_clock_stops())}",
                        fg=col)
            except Exception:
                pass
            # Panneau prod clignote en rouge si arret actif
            try:
                def _set_bg_recursive(widget, col):
                    try:
                        widget.config(bg=col)
                    except Exception:
                        pass
                    for child in widget.winfo_children():
                        _set_bg_recursive(child, col)
                if self._main_prod_panel and any_running:
                    col = C_RED if self._cell_blink else "#7a0000"
                    _set_bg_recursive(self._main_prod_panel, col)
                elif self._main_prod_panel and not any_running:
                    _set_bg_recursive(self._main_prod_panel, NAVY)
            except Exception:
                pass

        # Auto-retour
        if (self._mode == "main" and self._prod_active
                and (now - self._last_activity).total_seconds() > 30):
            self._last_activity = now
            self._transition(self._nav_to_production)
            return

        # Blink dot onglet production
        if self._mode == "main" and self._prod_active:
            self._blink_state = not self._blink_state
            try:
                self._blink_dot.config(fg=C_RED if self._blink_state else NAVY_L)
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

    def _redraw_status(self, of_s=0, stop_s=0, any_running=False):
        """Redessine la barre de statut Canvas (chronos + KPI)."""
        cv = self._status_cv
        if not cv:
            return
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 20 or h < 10:
            return

        # Fond arrondi
        r = 14
        if any_running and self._cell_blink:
            face = "#cc0000"
        elif any_running:
            face = C_RED
        else:
            face = GREEN

        _rrect(cv, 0, 0, w, h, r, fill=_off(face, -30))
        _rrect(cv, 0, 0, w, h - 4, r, fill=face)
        # Reflet haut
        _rrect(cv, 2, 2, w - 2, h // 3, r, fill=_off(face, +35))

        if any_running:
            # Gros affichage EN ARRET
            cv.create_text(20, h // 2, text="⚠  EN ARRET",
                           font=("Arial", 20, "bold"), fill=WHITE, anchor="w")
            cv.create_text(w // 2, h // 2, text=fmt(stop_s),
                           font=("Arial", 32, "bold"), fill="#ffff44", anchor="center")
            # OF petit
            cv.create_text(w - 20, h // 2,
                           text=f"OF: {fmt(of_s)}",
                           font=("Arial", 13), fill="#ffcccc", anchor="e")
        else:
            # Affichage normal production
            cv.create_text(20, h // 2, text="⏱",
                           font=("Arial", 16), fill=WHITE, anchor="w")
            cv.create_text(50, h // 2, text=fmt(of_s),
                           font=("Arial", 30, "bold"), fill=WHITE, anchor="w")
            cv.create_text(260, h // 2, text=f"🔧  {fmt(stop_s)}",
                           font=("Arial", 14), fill="#d4f5d4", anchor="w")
            # KPI bar
            if of_s > 0:
                prod_r = max(0.0, (of_s - stop_s) / of_s)
                bx, bw2 = w - 280, 260
                cv.create_rectangle(bx, h // 2 - 10, bx + bw2, h // 2 + 10,
                                    fill=_off(GREEN, -60), outline="")
                cv.create_rectangle(bx, h // 2 - 10,
                                    bx + int(bw2 * prod_r), h // 2 + 10,
                                    fill="#4ade80", outline="")
                pct_p = int(prod_r * 100)
                cv.create_text(bx + bw2 // 2, h // 2,
                               text=f"PROD {pct_p}%  |  ARRET {100 - pct_p}%",
                               font=("Arial", 9, "bold"), fill=WHITE, anchor="center")

    def _reset_activity(self, event=None):
        self._last_activity = datetime.datetime.now()

    def _transition(self, fn):
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
        self._save_form_data()
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
            self._main_prod_panel = prod_wrap

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

        cols = ("Date", "OF", "Pilote", "Poste", "Qte Fab", "Qte Emb",
                "Equiv", "TRS %", "Duree OF", "Arrêts", "✏", "🗑")
        tree = ttk.Treeview(tbl_inner, columns=cols, show="headings",
                            height=13, style="KPI.Treeview")
        widths = {"Date": 85, "OF": 90, "Pilote": 130, "Poste": 80,
                  "Qte Fab": 60, "Qte Emb": 60, "Equiv": 60,
                  "TRS %": 65, "Duree OF": 75, "Arrêts": 75,
                  "✏": 36, "🗑": 36}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 70), anchor="center",
                        stretch=(c not in ("✏", "🗑")))
        style.configure("TRS.Treeview", font=("Arial", 10, "bold"))
        tree.tag_configure("trs_hi", background="#e6f7ee", foreground=GREEN)

        sb_v = ttk.Scrollbar(tbl_inner, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb_v.set)
        tree.pack(side="left", fill="both", expand=True)
        sb_v.pack(side="right", fill="y")

        self._main_tree = tree
        self._col_ids   = cols
        self._load_table(tree)

        tree.bind("<Button-1>", self._on_tree_click)

        self._after_id = self.root.after(1000, self._tick)

    def _save_form_data(self):
        """Sauvegarde les valeurs du formulaire de production avant de quitter la vue."""
        if hasattr(self, "fv") and self.fv:
            self._saved_form_data = {k: v.get() for k, v in self.fv.items()}
        if hasattr(self, "_comment_txt") and self._comment_txt:
            try:
                self._saved_form_data["_comment"] = self._comment_txt.get("1.0", "end").strip()
            except Exception:
                pass
        if hasattr(self, "_v_kit") and self._v_kit:
            try:
                self._saved_form_data["_kit"] = self._v_kit.get()
            except Exception:
                pass
        self._save_session()

    def _restore_form_data(self):
        """Restaure les valeurs du formulaire après retour en production."""
        if not self._saved_form_data:
            return
        for k, v in self._saved_form_data.items():
            if k.startswith("_"):
                continue
            if k in self.fv:
                try:
                    self.fv[k].set(v)
                except Exception:
                    pass
        if "_comment" in self._saved_form_data and self._comment_txt:
            try:
                self._comment_txt.delete("1.0", "end")
                self._comment_txt.insert("1.0", self._saved_form_data["_comment"])
            except Exception:
                pass
        if "_kit" in self._saved_form_data and self._v_kit:
            try:
                self._v_kit.set(self._saved_form_data["_kit"])
            except Exception:
                pass

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

    def _calc_trs_from_row(self, row):
        """TRS d'une ligne Data: (equiv produit) / (ref * temps_prod / 28800) * 100."""
        try:
            equiv = float(str(row[15]).replace(",", ".")) if row[15] else 0.0
            of_s  = _hms_to_sec(str(row[16])) if row[16] else 0
            if of_s <= 0 or equiv <= 0:
                return -1
            prod_ref = self._get_prod_ref()
            if prod_ref <= 0:
                return -1
            expected = prod_ref * of_s / 28800.0
            return min(200.0, equiv / expected * 100.0)
        except Exception:
            return -1

    def _get_prod_ref(self):
        """Lit 'Prod de reference (8h)' ligne 1 de la colonne dans Listes."""
        for key in self.lists:
            if "reference" in key.lower() or "ref" in key.lower():
                vals = self.lists[key]
                if vals:
                    try:
                        return float(str(vals[0]).replace(",", "."))
                    except Exception:
                        pass
        return 0.0

    def _refresh_main_kpi(self):
        path = self.cfg.get("db_path", "")
        last_time = "--:--"
        global_trs = 0.0
        if path and os.path.exists(path):
            try:
                wb   = load_workbook(path, read_only=True, data_only=True)
                rows = [list(r) + [None]*55
                        for r in wb["Data"].iter_rows(min_row=2, values_only=True)
                        if any(r)]
                wb.close()
                if rows:
                    last      = rows[-1]
                    last_time = str(last[18])[:5] if last[18] else "--:--"
                # TRS global : sum(equiv) / sum(expected) sur 12h
                cutoff = datetime.datetime.now() - datetime.timedelta(hours=12)
                total_equiv = 0.0
                total_of_s  = 0.0
                prod_ref    = self._get_prod_ref()
                for row in rows:
                    try:
                        date_s = str(row[1] or "")
                        time_s = str(row[17] or "")
                        dt = datetime.datetime.strptime(
                            f"{date_s} {time_s}", "%d/%m/%Y %H:%M:%S")
                        if dt < cutoff:
                            continue
                    except Exception:
                        pass
                    try:
                        total_equiv += float(str(row[15] or 0).replace(",", "."))
                        total_of_s  += _hms_to_sec(str(row[16] or "00:00:00"))
                    except Exception:
                        pass
                if prod_ref > 0 and total_of_s > 0:
                    expected   = prod_ref * total_of_s / 28800.0
                    global_trs = min(200.0, total_equiv / expected * 100.0) if expected > 0 else 0.0
            except Exception:
                pass
        self._main_gauge.update_gauge(global_trs, last_time)

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
            qte_fab = int(row[13]) if row[13] and str(row[13]).isdigit() else 0
            equiv   = float(str(row[15]).replace(",", ".")) if row[15] else 0.0
            of_s    = _hms_to_sec(str(row[16])) if row[16] else 0
            arr_s   = _sd(row, list(range(26, 31)) + list(range(31, 49)))
            prod_s  = max(0, of_s - arr_s)
            trs_str = ""
            if of_s > 0:
                trs_val = self._calc_trs_from_row(row)
                trs_str = f"{trs_val:.0f}%" if trs_val >= 0 else ""
            tags = ("trs_hi",) if trs_str else ()
            tree.insert("", "end", iid=str(excel_row), tags=tags, values=(
                str(row[1])[:10]      if row[1]  else "",
                str(row[0])           if row[0]  else "",
                str(row[3])           if row[3]  else "",
                str(row[2])           if row[2]  else "",
                str(row[13])          if row[13] else "0",
                str(row[14])          if row[14] else "0",
                f"{equiv:.1f}"        if equiv   else "",
                trs_str,
                str(row[16])          if row[16] else "",
                fmt(arr_s),
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
            self._refresh_main_kpi()
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
                # Recalculer equivalence et TRS avant sauvegarde
                vals = [v.get() for v in field_vars]
                try:
                    qte_f = int(str(vals[13] or 0))
                    taille_v  = str(vals[6] or "")
                    type_pv   = str(vals[8] or "")
                    new_equiv = self._calc_equiv(qte_f, taille_v, type_pv)
                    vals[15]  = str(new_equiv)
                    field_vars[15].set(str(new_equiv))
                except Exception:
                    pass
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
                self._refresh_main_kpi()
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
                do_changeof = False
                if gap < 300:  # Moins de 5 min → automatique
                    do_changeof = True
                else:
                    h = int(gap // 3600)
                    m = int((gap % 3600) // 60)
                    s = int(gap % 60)
                    ts = f"{h}h {m:02d}min" if h > 0 else f"{m}min {s:02d}s"
                    do_changeof = messagebox.askyesno(
                        "Changement d'OF",
                        f"Le dernier OF a ete termine il y a {ts}.\n\n"
                        "Voulez-vous declarer ce temps comme\n"
                        "\"Changement d'OF\" ?")
                if do_changeof:
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
        self._save_session()
        self._show_production()

    def _show_production(self):
        self._clear()
        self._db_labels.clear()
        self._mode = "production"
        self._prod_state = "prod"
        self._stop_timer_lbls = {}
        self._active_stops_container = None

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)
        self._outer_frame = outer

        # ── En-tete fixe (blanc, style Dodo) ─────────────────────────────────
        hdr = tk.Frame(outer, bg=WHITE, height=62)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        # Accent gauche coloré
        tk.Frame(hdr, bg=GREEN, width=6).pack(side="left", fill="y")
        tk.Label(hdr, text="KPI-ORC  |  ORC1",
                 bg=WHITE, fg=NAVY, font=("Arial", 15, "bold")).pack(
                 side="left", padx=16, pady=12)
        tk.Label(hdr, text=f"Debut : {self._of_start.strftime('%H:%M:%S')}",
                 bg=WHITE, fg=GRAY, font=("Arial", 11)).pack(side="left")
        right_bar = tk.Frame(hdr, bg=WHITE)
        right_bar.pack(side="right", padx=12)
        self._db_widget(right_bar, WHITE).pack(side="right", padx=4)

        # ── Barre de statut Canvas (chrono + KPI, change couleur) ────────────
        self._status_cv = tk.Canvas(outer, height=72, bg=BG, highlightthickness=0)
        self._status_cv.pack(fill="x", padx=8, pady=(4, 0))
        self._status_cv.bind("<Configure>",
                             lambda e: self._redraw_status(0, 0, False))

        self._make_tabs(outer, "production")
        self._make_timeline(outer)

        # ── Corps ─────────────────────────────────────────────────────────────
        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True, padx=6, pady=(4, 6))

        # Zone 1 : formulaire (40%)
        left = tk.Frame(body, bg=WHITE)
        left.pack(side="left", fill="both", expand=True)
        tk.Frame(left, bg=LGRAY, height=1).pack(fill="x")
        self._build_form(left)

        tk.Frame(body, bg=LGRAY, width=1).pack(side="left", fill="y")

        # Zone 2 : arrets actifs + boutons (35%)
        mid = tk.Frame(body, bg=BG, width=320)
        mid.pack(side="left", fill="both")
        mid.pack_propagate(False)
        self._build_right_panel(mid)

        tk.Frame(body, bg=LGRAY, width=1).pack(side="left", fill="y")

        # Zone 3 : récap arrêts de l'OF (25%)
        recap_panel = tk.Frame(body, bg=WHITE, width=220)
        recap_panel.pack(side="left", fill="both")
        recap_panel.pack_propagate(False)
        self._recap_panel = recap_panel
        self._build_stops_recap(recap_panel)

        self._after_id = self.root.after(1000, self._tick)

    # ── Formulaire compact 3 colonnes (pas de scroll) ────────────────────────
    def _build_form(self, parent):
        self.fv = {}
        c = tk.Frame(parent, bg=WHITE)
        c.pack(fill="both", expand=True, padx=4, pady=2)
        c.columnconfigure(0, weight=1)
        c.columnconfigure(1, weight=1)
        c.columnconfigure(2, weight=1)
        ri = [0]

        LFONT = ("Arial", 7)
        EFONT = ("Arial", 9)

        def sec(txt, color=NAVY, ncols=3):
            row = tk.Frame(c, bg=WHITE)
            row.grid(row=ri[0], column=0, columnspan=ncols, sticky="ew", pady=(5, 1))
            tk.Frame(row, bg=color, width=4).pack(side="left", fill="y")
            tk.Label(row, text=f"  {txt}", bg=WHITE, fg=color,
                     font=("Arial", 8, "bold"), pady=1).pack(side="left")
            ri[0] += 1

        def fld(lbl_txt, key, ftype, lh=None, col=0, adv=True, suffix=None):
            cell = tk.Frame(c, bg=WHITE)
            cell.grid(row=ri[0], column=col, sticky="ew", padx=2, pady=1)
            cell.columnconfigure(0, weight=1)
            tk.Label(cell, text=lbl_txt, bg=WHITE, fg=GRAY,
                     font=LFONT, anchor="w").grid(row=0, column=0, columnspan=2, sticky="w")
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                e = tk.Entry(cell, textvariable=var, bg=WHITE, fg=DARK,
                             font=EFONT, relief="solid", bd=1, insertbackground=DARK)
                e.grid(row=1, column=0, sticky="ew", ipady=2)
                if suffix:
                    tk.Label(cell, text=suffix, bg=WHITE, fg=GRAY,
                             font=LFONT).grid(row=1, column=1, sticky="w", padx=(2, 0))
            else:
                cb = ttk.Combobox(cell, textvariable=var,
                                  values=self._get_list(lh) if lh else [],
                                  font=EFONT, state="readonly")
                cb.grid(row=1, column=0, columnspan=2, sticky="ew")
            if adv:
                ri[0] += 1

        def row3(l1, k1, t1, h1, l2, k2, t2, h2, l3, k3, t3, h3,
                 s1=None, s2=None, s3=None):
            fld(l1, k1, t1, h1, col=0, adv=False, suffix=s1)
            fld(l2, k2, t2, h2, col=1, adv=False, suffix=s2)
            fld(l3, k3, t3, h3, col=2, adv=True,  suffix=s3)

        def row2(l1, k1, t1, h1, l2, k2, t2, h2, s1=None, s2=None):
            fld(l1, k1, t1, h1, col=0, adv=False, suffix=s1)
            fld(l2, k2, t2, h2, col=1, adv=True,  suffix=s2)

        # ── Identification (3 colonnes) ──
        sec("Identification", NAVY)
        row3("N° OF *",    "of_num",  "entry", None,
             "Poste *",    "poste",   "combo", "Postes",
             "Pilote *",   "pilote",  "combo", "Pilotes")
        row3("Co-Pilote",  "copilote","combo", "Co-pilotes",
             "Nb personnes","nb_pers","combo", "Nb personnes",
             "Fibre",      "fibre",   "combo", "Fibre")

        # ── Produit ──
        sec("Produit", NAVY_L)
        # KIT checkbox sur une ligne
        self._v_kit = tk.BooleanVar()
        kit_f = tk.Frame(c, bg=WHITE)
        kit_f.grid(row=ri[0], column=0, columnspan=3, sticky="w", padx=4, pady=1)
        ri[0] += 1
        tk.Checkbutton(kit_f, text="Kit de 2 pièces", variable=self._v_kit,
                       bg=WHITE, fg=DARK, font=EFONT,
                       activebackground=WHITE, selectcolor=WHITE).pack(side="left")

        row3("Taille",         "taille",    "combo", "Taille produit",
             "Type produit",   "type_prod", "combo", "Type produit",
             "Code produit *", "code_prod", "entry", None)
        row2("Poids garnissage","poids",    "entry", None,
             "OF taie",        "of_taie",   "entry", None,
             s1="gr")

        # ── Quantités / Qualité ──
        sec("Quantités & Qualité", GREEN)
        row3("Qte fabriquée *", "qte_fab",   "entry", None,
             "Qte emballée",    "qte_emb",   "entry", None,
             "Traca fibre",     "traca",      "entry", None)
        row3("Ref. taie",       "ref_taie",   "entry", None,
             "Nb déf. couture", "nb_def_cout","entry", None,
             "Mq. taie",        "mq_taie",    "entry", None)
        row3("Mq. housse (nb)", "mq_housse",  "entry", None,
             "Mq. encart (nb)", "mq_encart",  "entry", None,
             "Nb taie 2nd",     "nb_taie2",   "entry", None)

        # ── Commentaire ──
        sec("Commentaire", GRAY)
        txt_f = tk.Frame(c, bg=WHITE)
        txt_f.grid(row=ri[0], column=0, columnspan=3, sticky="ew", padx=2, pady=2)
        ri[0] += 1
        self._comment_txt = tk.Text(txt_f, height=2, bg=WHITE, fg=DARK,
                                     font=EFONT, relief="solid", bd=1,
                                     wrap="word", insertbackground=DARK)
        self._comment_txt.pack(fill="x")
        # Restaurer les valeurs sauvegardées si disponibles
        self.root.after(50, self._restore_form_data)

    # ── Panneau droit : arrets actifs + boutons ───────────────────────────────
    def _build_right_panel(self, parent):
        # Zone arrets actifs (prend tout l'espace disponible)
        stops_frame = tk.Frame(parent, bg=BG)
        stops_frame.pack(fill="both", expand=True, padx=6, pady=(6, 4))
        self._active_stops_container = stops_frame
        self._refresh_active_stops()

        BTN_H    = 66
        BTN_FONT = ("Arial", 14, "bold")

        def _make_cv_btn(text, color, cmd):
            cv = tk.Canvas(parent, height=BTN_H, highlightthickness=0, bg=BG)
            cv.pack(fill="x", padx=8, pady=(0, 4))
            pressed = [False]
            def _draw(e=None):
                cv.delete("all")
                bw, bh = cv.winfo_width(), cv.winfo_height()
                if bw < 10: return
                c = _off(color, -60) if pressed[0] else color
                _rrect(cv, 4, 5, bw-1, bh, 14, fill=_off(c, -40))
                _rrect(cv, 0, 0, bw-5, bh-5, 14, fill=c)
                _rrect(cv, 2, 2, bw-7, bh//3, 14, fill=_off(c, +45))
                cv.create_text(bw//2-2, bh//2-2, text=text,
                               fill=WHITE, font=BTN_FONT)
            def _press(e):
                pressed[0] = True
                _draw()
            def _release(e):
                pressed[0] = False
                _draw()
                cmd()
            cv.bind("<Configure>", _draw)
            cv.bind("<ButtonPress-1>",  _press)
            cv.bind("<ButtonRelease-1>", _release)
            cv.config(cursor="hand2")
            return cv

        _make_cv_btn("⚠   DÉCLARER UN ARRÊT / RATTRAPAGE", C_RATT,
                     self._show_stop_selector)
        _make_cv_btn("🧹  DÉCLARER UN ARRÊT NETTOYAGE", "#6b7280",
                     lambda: self._start_nettoyage())
        _make_cv_btn("⏹   DÉCLARER LA FIN DE PRODUCTION", GREEN,
                     self._end_production)

    # ── Arrets actifs ─────────────────────────────────────────────────────────
    def _refresh_active_stops(self):
        container = self._active_stops_container
        if not container:
            return
        for w in container.winfo_children():
            w.destroy()
        self._stop_timer_lbls = {}

        active_keys = [k for k in self._timers if self._t_running(k)]

        if not active_keys:
            ok_f = tk.Frame(container, bg=BG)
            ok_f.pack(fill="both", expand=True)
            tk.Label(ok_f, text="✅", bg=BG, fg=GREEN,
                     font=("Arial", 48)).pack(pady=(20, 6))
            tk.Label(ok_f, text="Aucun arrêt en cours",
                     bg=BG, fg=GREEN, font=("Arial", 14, "bold")).pack()
            return

        for key in active_keys:
            ev_info = next((e for e in EVENTS if e[1] == key), None)
            if not ev_info:
                continue
            label, _, cat = ev_info
            color    = C_RATT if cat == "ratt" else C_RED
            type_lbl = "Rattrapage" if cat == "ratt" else "Problème technique"

            # Carte arret
            card_shad = tk.Frame(container, bg=_off(color, -40))
            card_shad.pack(fill="x", pady=4)
            card = tk.Frame(card_shad, bg=color)
            card.pack(fill="both", padx=(0, 3), pady=(0, 3))

            top_row = tk.Frame(card, bg=color)
            top_row.pack(fill="x", padx=14, pady=(10, 2))
            name_col = tk.Frame(top_row, bg=color)
            name_col.pack(side="left")
            tk.Label(name_col, text=type_lbl.upper(), bg=color,
                     fg=_off(WHITE, -60), font=("Arial", 8, "bold")).pack(anchor="w")
            tk.Label(name_col, text=f"{'▶' if cat == 'ratt' else '⚠'}  {label}",
                     bg=color, fg=WHITE, font=("Arial", 13, "bold")).pack(anchor="w")

            def _stop(k=key):
                self._ask_stop_description(k)
            tk.Button(top_row, text="ARRÊTER  ✓", command=_stop,
                      bg=WHITE, fg=color, font=("Arial", 11, "bold"),
                      relief="flat", padx=14, pady=5, cursor="hand2").pack(side="right")

            elapsed = self._t_get(key)
            tlbl = tk.Label(card, text=fmt(elapsed), bg=color, fg=WHITE,
                            font=("Arial", 28, "bold"))
            tlbl.pack(pady=(2, 10))
            self._stop_timer_lbls[key] = tlbl

    # ── Récap arrêts de l'OF (zone droite) ───────────────────────────────────
    def _build_stops_recap(self, parent):
        tk.Frame(parent, bg=LGRAY, height=1).pack(fill="x")
        hdr_f = tk.Frame(parent, bg=WHITE)
        hdr_f.pack(fill="x", padx=8, pady=(6, 2))
        tk.Frame(hdr_f, bg=C_RED, width=4).pack(side="left", fill="y")
        tk.Label(hdr_f, text="  RÉCAP ARRÊTS OF",
                 bg=WHITE, fg=DARK, font=("Arial", 8, "bold")).pack(side="left")
        self._recap_inner = tk.Frame(parent, bg=WHITE)
        self._recap_inner.pack(fill="both", expand=True, padx=4, pady=4)
        self._refresh_stops_recap()

    def _refresh_stops_recap(self):
        inner = getattr(self, "_recap_inner", None)
        if not inner:
            return
        for w in inner.winfo_children():
            w.destroy()
        # Cumuler par clé d'arrêt
        cumuls = {}
        now = datetime.datetime.now()
        for ev in self._tl_events:
            if ev.get("cat") not in ("ratt", "pb"):
                continue
            if self._of_start and ev["start"] < self._of_start:
                continue
            key  = ev["key"]
            s    = ev["start"]
            e    = ev.get("end") or now
            dur  = (e - s).total_seconds()
            if key not in cumuls:
                cumuls[key] = {"dur": 0.0, "cat": ev["cat"], "n": 0}
            cumuls[key]["dur"] += dur
            cumuls[key]["n"]   += 1
        if not cumuls:
            tk.Label(inner, text="Aucun arrêt", bg=WHITE, fg=LGRAY,
                     font=("Arial", 9, "italic")).pack(pady=12)
            return
        for key, info in sorted(cumuls.items(), key=lambda x: -x[1]["dur"]):
            label = next((e[0] for e in EVENTS if e[1] == key), key)
            color = C_RATT if info["cat"] == "ratt" else C_RED
            mins  = int(info["dur"] // 60)
            secs  = int(info["dur"] % 60)
            dur_s = f"{mins}min {secs:02d}s" if mins > 0 else f"{secs}s"
            row_f = tk.Frame(inner, bg=WHITE)
            row_f.pack(fill="x", pady=1, padx=2)
            tk.Frame(row_f, bg=color, width=4).pack(side="left", fill="y")
            name_f = tk.Frame(row_f, bg=WHITE)
            name_f.pack(side="left", fill="both", expand=True, padx=(4, 0))
            tk.Label(name_f, text=label, bg=WHITE, fg=DARK,
                     font=("Arial", 8), anchor="w",
                     wraplength=120).pack(anchor="w")
            tk.Label(name_f, text=f"× {info['n']}  —  {dur_s}",
                     bg=WHITE, fg=color, font=("Arial", 8, "bold"),
                     anchor="w").pack(anchor="w")
        tk.Frame(inner, bg=LGRAY, height=1).pack(fill="x", pady=4)
        total = sum(v["dur"] for v in cumuls.values())
        tm = int(total // 60)
        ts = int(total % 60)
        tk.Label(inner, text=f"Total : {tm}min {ts:02d}s",
                 bg=WHITE, fg=DARK, font=("Arial", 8, "bold")).pack(anchor="w", padx=6)

    def _start_nettoyage(self):
        key = "nettoyage"
        if self._t_running(key):
            self._ask_stop_description(key)
        else:
            self._t_start(key)
            self._tl_open(key, "ratt")
            self._refresh_active_stops()

    # ── Selecteur d'arret ─────────────────────────────────────────────────────
    def _show_stop_selector(self):
        OV_BG = WHITE
        overlay = tk.Frame(self.root, bg=OV_BG)
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        overlay.lift()

        # Header blanc avec bande bleue
        hdr = tk.Frame(overlay, bg=OV_BG, height=68)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=NAVY, width=6).pack(side="left", fill="y")
        tk.Label(hdr, text="CHOISIR UN ARRÊT",
                 bg=OV_BG, fg=NAVY, font=("Arial", 18, "bold")).pack(
                 side="left", padx=16, pady=16)
        tk.Frame(hdr, bg=LGRAY, height=1).pack(side="bottom", fill="x")

        def _close():
            overlay.destroy()
            self._refresh_active_stops()

        close_cv = tk.Canvas(hdr, width=160, highlightthickness=0, bg=OV_BG)
        close_cv.pack(side="right", padx=16, pady=12)

        def _draw_close(e=None):
            close_cv.delete("all")
            w, h = close_cv.winfo_width(), close_cv.winfo_height()
            _rrect(close_cv, 2, 2, w-2, h-2, 10, fill=LGRAY)
            close_cv.create_text(w//2, h//2, text="✕  RETOUR",
                                 fill=DARK, font=("Arial", 12, "bold"))

        close_cv.bind("<Configure>", _draw_close)
        close_cv.bind("<Button-1>", lambda e: _close())
        close_cv.config(cursor="hand2")

        body = tk.Frame(overlay, bg=OV_BG)
        body.pack(fill="both", expand=True, padx=16, pady=8)

        # RATTRAPAGES (sans Nettoyage qui a son propre bouton)
        ratt_events = [(l, k, c) for l, k, c in EVENTS if c == "ratt" and k != "nettoyage"]
        rat_row = tk.Frame(body, bg=OV_BG)
        rat_row.pack(fill="x", pady=(0, 4))
        tk.Frame(rat_row, bg=C_RATT, width=5).pack(side="left", fill="y")
        tk.Label(rat_row, text="  ARRÊTS RATTRAPAGE", bg=OV_BG, fg=C_RATT,
                 font=("Arial", 12, "bold")).pack(side="left", pady=4)

        ratt_g = tk.Frame(body, bg=OV_BG)
        ratt_g.pack(fill="x", pady=(0, 10))
        ncols_r = 5
        for col in range(ncols_r):
            ratt_g.columnconfigure(col, weight=1)
        ratt_g.rowconfigure(0, weight=1)
        for i, (lbl, key, cat) in enumerate(ratt_events):
            self._make_selector_btn(ratt_g, lbl, key, cat, 0, i % ncols_r, _close)

        # PB TECHNIQUES
        pb_events = [(l, k, c) for l, k, c in EVENTS if c == "pb"]
        pb_row = tk.Frame(body, bg=OV_BG)
        pb_row.pack(fill="x", pady=(0, 4))
        tk.Frame(pb_row, bg=C_RED, width=5).pack(side="left", fill="y")
        tk.Label(pb_row, text="  PROBLÈMES TECHNIQUES", bg=OV_BG, fg=C_RED,
                 font=("Arial", 12, "bold")).pack(side="left", pady=4)

        pb_g = tk.Frame(body, bg=OV_BG)
        pb_g.pack(fill="both", expand=True)
        ncols_p = 6
        nrows_p = (len(pb_events) + ncols_p - 1) // ncols_p
        for col in range(ncols_p):
            pb_g.columnconfigure(col, weight=1)
        for row in range(nrows_p):
            pb_g.rowconfigure(row, weight=1)
        for i, (lbl, key, cat) in enumerate(pb_events):
            self._make_selector_btn(pb_g, lbl, key, cat, i // ncols_p, i % ncols_p, _close)

    def _make_selector_btn(self, parent, label, key, cat, row, col, close_fn):
        running = self._t_running(key)
        elapsed = self._t_get(key)
        color   = C_RATT if cat == "ratt" else C_RED
        face    = color if running else NAVY_L

        cv = tk.Canvas(parent, highlightthickness=0, bg=WHITE, cursor="hand2")
        cv.grid(row=row, column=col, sticky="nsew", padx=3, pady=3)

        def _draw(e=None):
            cv.delete("all")
            bw, bh = cv.winfo_width(), cv.winfo_height()
            if bw < 10 or bh < 10:
                return
            shad = _off(face, -50)
            _rrect(cv, 3, 4, bw-1, bh, 10, fill=shad)
            _rrect(cv, 0, 0, bw-4, bh-4, 10, fill=face)
            _rrect(cv, 1, 1, bw-5, bh//3, 10, fill=_off(face, +40))
            lines = [label]
            if running:
                lines.append(fmt(elapsed))
                lines.append("● EN COURS")
            txt = "\n".join(lines)
            cv.create_text(bw//2-2, bh//2-2, text=txt,
                           fill=WHITE, font=("Arial", 9, "bold"),
                           justify="center", width=bw-12)

        def _action(k=key, c=cat):
            if self._t_running(k):
                self._ask_stop_description(k)
                close_fn()
            else:
                self._t_start(k)
                self._tl_open(k, c)
                close_fn()

        cv.bind("<Configure>", _draw)
        cv.bind("<Button-1>",  lambda e: _action())

    # ── Popup description d'arret ─────────────────────────────────────────────
    def _ask_stop_description(self, key):
        ev_info = next((e for e in EVENTS if e[1] == key), None)
        label   = ev_info[0] if ev_info else key

        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=WHITE)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        pw, ph = min(640, sw - 60), 360
        top.geometry(f"{pw}x{ph}+{(sw-pw)//2}+{(sh-ph)//2}")

        # Header coloré
        hdr_col = C_RATT if (ev_info and ev_info[2] == "ratt") else C_RED
        hdr = tk.Frame(top, bg=hdr_col, height=62)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"Fin d'arrêt  —  {label}",
                 bg=hdr_col, fg=WHITE, font=("Arial", 15, "bold")).pack(
                 side="left", padx=20, pady=16)

        body = tk.Frame(top, bg=WHITE)
        body.pack(fill="both", expand=True, padx=24, pady=14)
        tk.Label(body, text="Cause de l'arrêt :", bg=WHITE, fg=DARK,
                 font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 8))
        txt = tk.Text(body, height=5, font=("Arial", 13), bg=WHITE, fg=DARK,
                      insertbackground=DARK, relief="solid", bd=1,
                      padx=10, pady=8, wrap="word")
        txt.pack(fill="x")
        txt.focus()

        def _valider(comment=None):
            desc = comment if comment is not None else txt.get("1.0", "end").strip()
            self._t_stop(key)
            self._tl_close(key, comment=desc)
            top.destroy()
            self._refresh_active_stops()

        btn_row = tk.Frame(top, bg=WHITE)
        btn_row.pack(fill="x", padx=24, pady=(8, 16))
        tk.Button(btn_row, text="Passer", command=lambda: _valider(""),
                  bg=LGRAY, fg=DARK, font=("Arial", 11), relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left")
        tk.Button(btn_row, text="✔   VALIDER",
                  command=lambda: _valider(), bg=GREEN, fg=WHITE,
                  font=("Arial", 15, "bold"), relief="flat",
                  padx=24, pady=10, cursor="hand2").pack(side="right")
        top.bind("<Return>", lambda e: _valider())

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

        v = {k: var.get().strip() for k, var in self.fv.items()}
        v["comment"] = self._comment_txt.get("1.0", "end").strip()

        # ── Validation des saisies ─────────────────────────────────────────────
        def _n(k):
            try:
                return int(v.get(k, 0) or 0)
            except Exception:
                return 0

        poids_str  = v.get("poids", "").strip()
        qte_str    = v.get("qte_fab", "").strip()
        if poids_str:
            try:
                poids_val = float(poids_str.replace(",", "."))
                if not (50 <= poids_val <= 4000):
                    messagebox.showwarning("Donnée incohérente",
                                          f"Poids de garnissage : {poids_val} gr\n"
                                          "Valeur attendue entre 50 et 4000 gr.")
                    return
            except Exception:
                messagebox.showwarning("Donnée incohérente",
                                       "Poids de garnissage : valeur non valide.")
                return
        if qte_str:
            try:
                qte_val = int(qte_str)
                if not (1 <= qte_val <= 10000):
                    messagebox.showwarning("Donnée incohérente",
                                           f"Quantité fabriquée : {qte_val}\n"
                                           "Valeur attendue entre 1 et 10 000.")
                    return
            except Exception:
                messagebox.showwarning("Donnée incohérente",
                                       "Quantité fabriquée : valeur non valide.")
                return

        end_dt = datetime.datetime.now()
        self._t_stop_all()
        self._tl_close_all()

        of_s    = (end_dt - self._of_start).total_seconds()
        stop_s  = self._t_wall_clock_stops()
        qte_fab = _n("qte_fab")
        nb_pers = max(1, _n("nb_pers") or 1)
        of_min  = of_s / 60
        of_hrs  = of_s / 3600
        c1      = round(qte_fab / of_min, 2)  if of_min  > 0 else 0
        c2      = round(qte_fab / (nb_pers * of_hrs), 2) if of_hrs > 0 else 0
        equiv   = self._calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
        kit     = 2 if self._v_kit.get() else 1

        prod_ref = self._get_prod_ref()
        trs_pct  = -1.0
        if prod_ref > 0 and of_s > 0:
            expected = prod_ref * of_s / 28800.0
            trs_pct  = min(200.0, equiv / expected * 100.0) if expected > 0 else -1.0

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
            _n("mq_housse") + _n("mq_encart"),
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

        # ── Popup recap avant confirmation ────────────────────────────────────
        confirmed = [False]
        modified  = [False]

        recap = tk.Toplevel(self.root)
        recap.overrideredirect(True)
        recap.attributes("-topmost", True)
        recap.configure(bg=WHITE)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        pw, ph = min(680, sw - 60), 440
        recap.geometry(f"{pw}x{ph}+{(sw-pw)//2}+{(sh-ph)//2}")

        # Header
        hdr_r = tk.Frame(recap, bg=NAVY, height=56)
        hdr_r.pack(fill="x")
        hdr_r.pack_propagate(False)
        tk.Label(hdr_r, text="RÉCAPITULATIF DE L'OF",
                 bg=NAVY, fg=WHITE, font=("Arial", 15, "bold")).pack(
                 side="left", padx=20, pady=14)

        # Corps en 2 colonnes : infos à gauche, jauge TRS à droite
        body_r = tk.Frame(recap, bg=WHITE)
        body_r.pack(fill="both", expand=True, padx=16, pady=8)
        left_r = tk.Frame(body_r, bg=WHITE)
        left_r.pack(side="left", fill="both", expand=True)
        right_r = tk.Frame(body_r, bg=WHITE, width=200)
        right_r.pack(side="right", fill="y")
        right_r.pack_propagate(False)

        def _row_info(lbl, val, color=DARK, bold=False):
            f = tk.Frame(left_r, bg=WHITE)
            f.pack(fill="x", pady=2)
            tk.Label(f, text=lbl, bg=WHITE, fg=GRAY,
                     font=("Arial", 10), width=22, anchor="w").pack(side="left")
            tk.Label(f, text=str(val), bg=WHITE, fg=color,
                     font=("Arial", 10, "bold" if bold else "normal")).pack(side="left")

        _row_info("N° OF",              v.get("of_num","—"))
        _row_info("Quantité fabriquée", f"{qte_fab}")
        _row_info("Equivalence",         f"{equiv:.2f}")
        _row_info("Durée production",    fmt(of_s))
        _row_info("Durée arrêts",         fmt(stop_s))

        # Jauge TRS graphique
        trs_col = GREEN if trs_pct >= 75 else C_RATT if trs_pct >= 55 else C_RED
        tk.Label(right_r, text="TRS cet OF", bg=WHITE, fg=GRAY,
                 font=("Arial", 9, "bold")).pack(pady=(12, 0))
        gauge_r = Gauge(right_r, bg=WHITE, width=180, height=120,
                        highlightthickness=0)
        gauge_r.pack(padx=8)
        trs_disp = max(0.0, trs_pct) if trs_pct >= 0 else 0.0
        trs_time  = f"{trs_disp:.1f}%" if trs_pct >= 0 else "—"
        gauge_r.update_gauge(trs_disp, trs_time)

        tk.Frame(left_r, bg=LGRAY, height=1).pack(fill="x", pady=8)

        btn_row_r = tk.Frame(recap, bg=WHITE)
        btn_row_r.pack(fill="x", padx=24, pady=(0, 18))

        def _modifier():
            confirmed[0] = False
            modified[0]  = True
            recap.destroy()

        def _confirmer():
            confirmed[0] = True
            recap.destroy()

        tk.Button(btn_row_r, text="✏  MODIFIER",
                  command=_modifier, bg=LGRAY, fg=DARK,
                  font=("Arial", 13, "bold"), relief="flat",
                  padx=20, pady=10, cursor="hand2").pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(btn_row_r, text="✔  CONFIRMER",
                  command=_confirmer, bg=GREEN, fg=WHITE,
                  font=("Arial", 13, "bold"), relief="flat",
                  padx=20, pady=10, cursor="hand2").pack(side="left", fill="x", expand=True)

        recap.wait_window()

        if modified[0]:
            # Retour a la vue production sans rien perdre
            self._prod_active = True
            self._after_id = self.root.after(1000, self._tick)
            return

        if not confirmed[0]:
            # Ferme sans confirmer → retour prod
            self._prod_active = True
            self._after_id = self.root.after(1000, self._tick)
            return

        # ── Ecriture Excel ────────────────────────────────────────────────────
        self._prod_active = False
        self._last_of_end = end_dt
        if self._of_periods:
            self._of_periods[-1]["end"]    = end_dt
            self._of_periods[-1]["of_num"] = v.get("of_num", "")

        ok = self._write_excel(row, v)
        if ok:
            self._delete_session()
            self._saved_form_data = {}   # Reset formulaire
            self._show_main()
            _toast(self.root, "✔  Production declaree avec succes !", bg=GREEN)
        else:
            # Fichier ouvert → message specifique, retour prod
            self._prod_active = True
            self._save_session()
            self._after_id = self.root.after(1000, self._tick)
            messagebox.showwarning(
                "Fichier Excel ouvert",
                "Veuillez fermer le fichier Excel,\ncontacter le Bureau méthodes.\n\n"
                "La production n'a PAS été annulée.")

    def _calc_equiv(self, qte, taille, type_prod):
        """Cherche le coef d'equivalence pour type_prod dans la colonne Equivalence coef."""
        types  = self._get_list("Type produit")
        # Accepte 'Equivalence coef' ou 'Equivalence' comme nom de colonne
        equivs = self._get_list("Equivalence coef") or self._get_list("Equivalence")
        if type_prod and types and equivs:
            for i, t in enumerate(types):
                if str(t).strip().lower() == str(type_prod).strip().lower():
                    if i < len(equivs):
                        try:
                            return round(qte * float(str(equivs[i]).replace(",", ".")), 2)
                        except Exception:
                            pass
        # Pas de correspondance
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
        kit_val  = "Oui" if self._v_kit.get() else "Non"
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
                v.get("fibre", ""),
                v.get("poids", ""),
                v.get("of_taie", ""),  v.get("traca", ""),
                v.get("ref_taie", ""), kit_val,
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
            # Même structure que _write_events_to_wb (EVT_HEADERS)
            # "Evenement","OF","Date","Poste","Pilote","Co-Pilote","Nb Personnes",
            # "Taille","Type Produit","Code Produit","Fibre","Poids Garnissage",
            # "OF Taie","Traca Fibre","Ref Taie","Kit",
            # "Heure Debut","Heure Fin","Duree","Commentaire"
            ws.append([
                "Changement d'OF",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                start_dt.strftime("%H:%M:%S"),
                end_dt.strftime("%H:%M:%S"),
                fmt(dur), "",
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
