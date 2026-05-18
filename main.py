"""KPI-ORC v5.49 - Style Dodo (bleu marine #1a1f5e + rouge #e31e24)"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, sys, datetime, math, threading
from openpyxl import load_workbook

# ── Serveur HTTP local pour édition OF depuis le dashboard HTML ───────────
EDIT_SERVER_PORT = 7892

def _start_edit_server(app):
    import http.server, socketserver, json as _jhttp
    class _H(http.server.BaseHTTPRequestHandler):
        def do_OPTIONS(self):
            self.send_response(200)
            for k, v in [('Access-Control-Allow-Origin','*'),
                         ('Access-Control-Allow-Methods','POST,OPTIONS'),
                         ('Access-Control-Allow-Headers','Content-Type')]:
                self.send_header(k, v)
            self.end_headers()
        def do_POST(self):
            if self.path != '/edit':
                self.send_response(404); self.end_headers(); return
            try:
                n = int(self.headers.get('Content-Length', 0))
                payload = _jhttp.loads(self.rfile.read(n))
                app.root.after(0, lambda p=payload: app._apply_html_edit(p))
                self.send_response(200)
                for k, v in [('Content-Type','application/json'),
                              ('Access-Control-Allow-Origin','*')]:
                    self.send_header(k, v)
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
            except Exception as e:
                self.send_response(500)
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(_jhttp.dumps({"error": str(e)}).encode())
        def log_message(self, *a): pass
    try:
        srv = socketserver.TCPServer(("localhost", EDIT_SERVER_PORT), _H)
        srv.daemon_threads = True
        srv.serve_forever()
    except Exception:
        pass
from openpyxl.styles import Alignment, Border, Side, PatternFill

CONFIG_FILE  = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")
SESSION_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_session.json")
PENDING_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_pending.json")
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
    "Duree OF", "Heure Debut", "Heure Fin", "Cadence/heure", "Cadence/h/pers",
    "Kit", "Ref Taie", "Qte Initiale Taie", "Nb Taie 2nd Choix",
    "Nb Defaut Couture", "Mq Taie", "Mq Housse/Encart",
    "Nb PP Cousue", "Changement de Serie",
    "Temps Arret Manquant MP", "Temps Arret Manquant Personnel/Reunion",
    "Nettoyage Fin de Poste",
    "Ratt Pochon/Fibre", "Ratt Couture", "Ratt Emballage",
    "Ratt Presse Souder", "Ratt Presse ZIP",
    "PB Chargeuse", "PB Carde", "PB Etaleur/Tour", "PB Coupe/Circ",
    "PB Tapis Bascule", "PB Enrouleur Pochon", "PB Pesee/Tapis 2",
    "PB Deviation/Table", "PB Enfileur Pochon", "PB Kinna/Stroebel",
    "PB Tapeuse", "PB Table Rot/Twin", "PB Enfileuse H1",
    "PB Enfileuse Traversin", "PB Presse ORC", "PB Presse Housse ZIP",
    "PB Cercleuse", "PB Enrouleuse Traversin",
    "Commentaire",
    "Temps Interposte",
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


def _row_date(v):
    """Normalise une valeur date Excel (datetime, date, ou string) en dd/mm/yyyy."""
    if v is None:
        return ""
    if hasattr(v, 'strftime'):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()[:10]
    if len(s) == 10 and s[4] == '-':   # ISO yyyy-mm-dd
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s


def _row_time(v):
    """Normalise une valeur time Excel (datetime.time ou string) en HH:MM:SS."""
    if v is None:
        return ""
    if hasattr(v, 'strftime'):
        return v.strftime("%H:%M:%S")
    return str(v).strip()


def _min_str_to_hms(val):
    """Convertit une valeur en minutes (saisie opérateur) vers HH:MM:SS pour Excel."""
    if not val or str(val).strip() == "":
        return ""
    try:
        total_s = int(float(str(val).replace(",", ".")) * 60)
        h, r = divmod(total_s, 3600)
        m, s = divmod(r, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"
    except Exception:
        return str(val)


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
    BAR_Y = 44
    BAR_H = 30

    def __init__(self, parent, app, **kw):
        kw.setdefault("height", 84)
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
            label = of_num if of_num else ""
            if x2 - x1 > 40 and label:
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
        self._val  = max(0.0, float(value))
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
        visual_val = min(100.0, self._val)  # arc capped at 100 visually
        ext   = visual_val * 180 / 100
        color = C_RED if self._val < 55 else ORANGE if self._val < 75 else GREEN
        if ext > 0:
            self.create_arc(cx-r, cy-r, cx+r, cy+r,
                            start=180, extent=-ext, style="arc",
                            outline=color, width=16)
        angle = math.radians(180 - visual_val * 180 / 100)
        nx = cx + int((r - 6) * math.cos(angle))
        ny = cy - int((r - 6) * math.sin(angle))
        self.create_line(cx, cy, nx, ny, fill=DARK, width=2)
        self.create_oval(cx-5, cy-5, cx+5, cy+5, fill=DARK, outline="")
        self.create_text(cx, cy - r // 2,
                         text=f"TRS  {self._val:.0f}%",
                         font=("Arial", 14, "bold"), fill=color)


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
        self.root.protocol("WM_DELETE_WINDOW", lambda: self._confirm_quit())

        self._px = lambda n: n   # pas de scaling

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
        self._saved_form_data  = {}   # Mémoire formulaire entre onglets
        self._reset_form_next  = False  # True = ne pas restaurer au prochain _show_production
        self._logged_in_pilot  = None   # Pilote actuellement connecté
        self._logged_in_poste  = None   # Poste choisi à la connexion
        self._login_time       = None   # Heure de connexion (pour postes de nuit)
        self._inter_of_s       = 0      # Durée inter-OF (changement de série)
        self._of_count_this_shift = 0   # Nb déclarations complétées ce poste
        self._data_rows_cache  = []
        self._events_cache     = []
        self._last_of_pilot    = ""
        self._last_of_num      = ""   # N° OF de la dernière déclaration (tous pilotes)
        self._interposte_s     = 0
        self._wb_cache         = None
        self._wb_path_cache    = ""
        self._wb_mtime_cache   = 0.0
        self._excel_lock       = threading.Lock()
        self._prod_ref_cached  = 0.0
        self._pilot_kpi_data   = {}
        self._pause_start      = None
        self._pause_total_s    = 0.0
        self._pause_periods    = []
        self._is_paused        = False
        self._pause_overlay    = None
        self._modal_open       = False
        self._trs_cache        = []
        self._loading_overlay  = None
        self._loading_canvas   = None
        self._loading_anim_id  = None
        self._loading_msg_lbl  = None
        self._review_full_data  = None   # Toutes les données Data
        self._review_full_evts  = None   # Tous les événements
        self._review_full_trs   = None   # Toutes les lignes TRS
        self._review_cache_ts   = 0.0    # Timestamp du dernier chargement complet

        self._load_lists()
        self._load_history_from_excel()
        # Démarrer le serveur d'édition HTML en arrière-plan
        import threading as _thr
        _thr.Thread(target=_start_edit_server, args=(self,), daemon=True).start()

        # ── Vérifier si une session était en cours ──────────────────────────
        if self._try_restore_session():
            return  # Session restaurée, _show_production() déjà appelé
        self.root.after(500, self._maybe_show_login)
        self._show_main()

    # ── Historique depuis Excel (reconstruit la timeline au demarrage) ────────
    def _apply_html_edit(self, payload):
        """Applique une modification d'OF depuis le dashboard HTML."""
        path = self.cfg.get("db_path", "")
        if not path:
            return
        data = payload.get("data", [])
        of_num_target = str(payload.get("of_num", "")).strip()
        date_target   = _row_date(str(payload.get("date", "")).strip())
        def _bg():
            try:
                with self._excel_lock:
                    wb = load_workbook(path)
                    ws = self._ensure_data_sheet(wb)
                    target_row = None
                    for row in ws.iter_rows(min_row=2):
                        if (str(row[0].value or "").strip() == of_num_target and
                                _row_date(str(row[1].value or "")) == date_target):
                            target_row = row[0].row
                            break
                    if target_row is None:
                        self.root.after(0, lambda: _toast(self.root,
                            "⚠  OF introuvable dans l'Excel", bg="#d97706"))
                        return
                    for col_i, val in enumerate(data[:len(DATA_HEADERS)], start=1):
                        ws.cell(row=target_row, column=col_i).value = val if val != "" else None
                    self._safe_excel_save(wb, path)
                    self._invalidate_wb_cache()
                self.root.after(0, self._reload_and_refresh)
                self.root.after(0, self._generate_dashboard_html)
                self.root.after(0, lambda: _toast(self.root,
                    "✔  Modifications sauvegardées dans Excel", bg="#16a34a", duration=3000))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda: _toast(self.root, f"✘  Erreur: {err}", bg="#dc2626"))
        import threading as _thr2
        _thr2.Thread(target=_bg, daemon=True).start()

    def _load_history_from_excel(self):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return

        def _pdt(date_s, time_s):
            try:
                return datetime.datetime.strptime(
                    f"{_row_date(date_s)} {_row_time(time_s)}",
                    "%d/%m/%Y %H:%M:%S")
            except Exception:
                return None

        today = datetime.date.today().strftime("%d/%m/%Y")
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            # Periodes OF d'aujourd'hui
            if "Data" in wb.sheetnames:
                ws = wb["Data"]
                min_r = 2 if str(ws.cell(1, 1).value or "").strip().upper() == "OF" else 1
                for row in ws.iter_rows(min_row=min_r, values_only=True):
                    if not any(row):
                        continue
                    r = list(row) + [None] * 55
                    date_val = _row_date(r[1])
                    if date_val != today:
                        continue
                    start_dt = _pdt(date_val, r[17])
                    end_dt   = _pdt(date_val, r[18])
                    if start_dt:
                        self._of_periods.append({
                            "start": start_dt, "end": end_dt,
                            "of_num": str(r[0] or "").strip()
                        })
                # Séparateurs d'OF — seulement si gap ≤ 8h
                for i, p in enumerate(self._of_periods[1:], start=1):
                    prev_end = self._of_periods[i - 1].get("end")
                    if prev_end is None:
                        continue
                    gap_s = (p["start"] - prev_end).total_seconds()
                    if gap_s <= 28800:
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
            rows_listes = list(ws.iter_rows(min_row=2, values_only=True))
            for row in rows_listes:
                for ci, val in enumerate(row, 1):
                    if ci in headers and val is not None:
                        self.lists[headers[ci]].append(str(val))
            # Prod de référence : même valeur pour tous les postes — première cellule non nulle de col I
            try:
                for row_ix in ws.iter_rows(min_row=2, min_col=9, max_col=9, values_only=True):
                    if row_ix and row_ix[0] is not None:
                        v_ref = str(row_ix[0]).replace(",", ".")
                        if v_ref.replace(".", "", 1).isdigit():
                            self._prod_ref_cached = float(v_ref)
                            break
            except Exception:
                pass
            wb.close()
        except Exception:
            pass

    def _select_db(self):
        # ── Étape 1 : mot de passe ─────────────────────────────────────────────
        self._modal_open = True
        top = tk.Toplevel(self.root)
        top.title("Acces base de donnees")
        top.resizable(False, False)
        top.grab_set()
        top.update_idletasks()
        self._center_on_root(top, 320, 180)
        top.bind("<Destroy>", lambda e: setattr(self, "_modal_open", False) if e.widget is top else None)
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
        btn_row_db = tk.Frame(top, bg=top.cget("bg"))
        btn_row_db.pack(pady=6)
        tk.Button(btn_row_db, text="Valider", command=_confirm_db,
                  bg=NAVY, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=16, pady=4,
                  cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row_db, text="Annuler", command=top.destroy,
                  bg=LGRAY, fg=DARK, font=("Arial", 10),
                  relief="flat", padx=12, pady=4,
                  cursor="hand2").pack(side="left")
        top.wait_window()
        if not allowed.get():
            return

        # ── Étape 2 : sélection du fichier ─────────────────────────────────────
        p = filedialog.askopenfilename(
            title="Selectionner la Base de Donnees",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tous", "*.*")])
        if not p:
            return

        self.cfg["db_path"] = p
        save_cfg(self.cfg)
        self._prod_ref_cached = 0.0
        self._invalidate_wb_cache()
        self._load_lists()
        self._load_history_from_excel()
        name = os.path.basename(p)
        for lbl in self._db_labels:
            try:
                lbl.config(text=f"DB: {name}")
            except Exception:
                pass

        # ── Étape 3 : saisie / confirmation de la référence de production ──────
        self._ask_prod_ref()


    def _ask_prod_ref(self):
        """Dialogue pour saisir/confirmer la référence production 8h (I2 Listes)."""
        current = float(self.cfg.get("prod_ref", self._prod_ref_cached) or 0)
        dlg = tk.Toplevel(self.root)
        dlg.title("Référence de production")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.update_idletasks()
        self._center_on_root(dlg, 380, 230)
        dlg.configure(bg=WHITE)

        tk.Frame(dlg, bg=NAVY, height=6).pack(fill="x")
        tk.Label(dlg, text="Référence de production (8h)",
                 bg=WHITE, fg=NAVY, font=("Arial", 13, "bold")).pack(pady=(16, 4))
        tk.Label(dlg,
                 text="Nombre de pièces attendues en 8h\n"
                      "(valeur de la cellule I2 de l'onglet Listes)",
                 bg=WHITE, fg=GRAY, font=("Arial", 10),
                 justify="center").pack()

        rv = tk.StringVar(value=str(int(current)) if current > 0 else "")
        err = tk.Label(dlg, text="", fg=C_RED, font=("Arial", 9), bg=WHITE)
        err.pack()
        e = tk.Entry(dlg, textvariable=rv, font=("Arial", 22, "bold"),
                     width=8, justify="center", relief="solid", bd=2,
                     fg=NAVY)
        e.pack(pady=4)
        e.focus()
        e.select_range(0, "end")

        def _save(ev=None):
            try:
                val = float(rv.get().replace(",", ".").strip())
                if val <= 0:
                    raise ValueError
            except ValueError:
                err.config(text="Entrez un nombre entier positif (ex: 850)")
                return
            self.cfg["prod_ref"] = val
            save_cfg(self.cfg)
            self._prod_ref_cached = val
            dlg.destroy()
            messagebox.showinfo("Configuré",
                                f"Référence production : {int(val)} pcs / 8h\n"
                                "Le TRS sera calculé sur cette base.")

        e.bind("<Return>", _save)
        tk.Button(dlg, text="✔  Enregistrer", command=_save,
                  bg=GREEN, fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", pady=8, cursor="hand2").pack(
                  fill="x", padx=40, pady=(4, 16))

    def _db_widget(self, parent, bg):
        f = tk.Frame(parent, bg=bg)
        tk.Button(f, text="⚙  Base", command=self._select_db,
                  bg=NAVY_L, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=12, pady=4, cursor="hand2").pack(side="left")
        return f

    def _get_list(self, h):
        # Aliases pour compatibilité avec les anciens et nouveaux noms de colonnes Excel
        _aliases = {
            "Mots de passe":    "Mots de passe pilote",
            "Co-Pilotes":       "Co-pilotes",
            "Type produit":     "Type de produit",
            "Fibre":            "Fibres",
            "Equivalence coef": "Equivalence Coef",
            "Equivalence":      "Equivalence Coef",
            "Nb Personnes":     "Nb personnes",
        }
        result = self.lists.get(h, [])
        if not result:
            result = self.lists.get(_aliases.get(h, h), [])
        return result

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
                "of_start":         self._dt_str(self._of_start),
                "last_of_end":      self._dt_str(self._last_of_end),
                "timers":           timers_serial,
                "tl_events":        events_serial,
                "of_periods":       periods_serial,
                "of_changes":       changes_serial,
                "form_data":        self._saved_form_data,
                "logged_in_pilot":  self._logged_in_pilot,
                "logged_in_poste":  self._logged_in_poste,
                "inter_of_s":       self._inter_of_s,
                "of_count_shift":   self._of_count_this_shift,
                "pause_total_s":    self._pause_total_s,
                "pause_start":      self._dt_str(self._pause_start) if self._pause_start else None,
                "is_paused":        self._is_paused,
                "pause_periods":    [[self._dt_str(a), self._dt_str(b)] for a, b in self._pause_periods],
            }
            tmp_sf = SESSION_FILE + ".tmp"
            with open(tmp_sf, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp_sf, SESSION_FILE)
        except Exception:
            pass

    def _safe_excel_save(self, wb, path):
        """Sauvegarde atomique Excel : tmp → backup → replace. À appeler sous _excel_lock."""
        import shutil
        tmp_path = path + ".tmp_kpi"
        bak_path = path + ".bak_kpi"
        wb.save(tmp_path)
        try:
            if os.path.exists(path):
                shutil.copy2(path, bak_path)
        except Exception:
            pass
        os.replace(tmp_path, path)

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

            # Pilote connecté
            pilot = data.get("logged_in_pilot")
            if pilot:
                self._logged_in_pilot = pilot
            poste = data.get("logged_in_poste")
            if poste:
                self._logged_in_poste = poste

            self._inter_of_s          = float(data.get("inter_of_s", 0))
            self._of_count_this_shift = int(data.get("of_count_shift", 0))

            # Pauses (restauration après crash)
            self._pause_total_s = float(data.get("pause_total_s", 0))
            self._is_paused     = bool(data.get("is_paused", False))
            self._pause_start   = self._str_dt(data.get("pause_start"))
            self._pause_periods = []
            for pair in data.get("pause_periods", []):
                if len(pair) == 2:
                    self._pause_periods.append(
                        (self._str_dt(pair[0]), self._str_dt(pair[1])))

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
        # Numéro d'OF courant au moment de l'ouverture de l'événement
        try:
            _cur_of = self.fv["of_num"].get() if hasattr(self, "fv") and "of_num" in self.fv else ""
        except Exception:
            _cur_of = ""
        self._tl_events.append({
            "key": key, "cat": cat,
            "start": datetime.datetime.now(), "end": None,
            "of_num": _cur_of,
            "pilot": self._logged_in_pilot or "",
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
        try:
            self._refresh_stops_recap()
        except Exception:
            pass

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

    def _show_login_overlay(self, on_success=None):
        """Overlay plein écran de connexion pilote."""
        ov = tk.Frame(self.root, bg=BG)
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        ov.lift()

        tk.Frame(ov, bg=GREEN, height=6).pack(fill="x")
        tk.Label(ov, text="KPI-ORC", bg=BG, fg=NAVY,
                 font=("Arial", 28, "bold")).pack(pady=(40, 4))
        tk.Label(ov, text="Connexion Pilote", bg=BG, fg=GRAY,
                 font=("Arial", 14)).pack(pady=(0, 32))

        card = tk.Frame(ov, bg=WHITE, bd=0)
        card.pack(padx=120, pady=0, fill="x")
        tk.Frame(card, bg=GREEN, height=4).pack(fill="x")
        inner = tk.Frame(card, bg=WHITE)
        inner.pack(padx=40, pady=30, fill="x")

        tk.Label(inner, text="Nom du pilote", bg=WHITE, fg=GRAY,
                 font=("Arial", 11)).pack(anchor="w")
        pilot_var = tk.StringVar()
        pilots = self._get_list("Pilotes")
        cb = ttk.Combobox(inner, textvariable=pilot_var,
                          values=pilots, font=("Arial", 14),
                          state="readonly", width=28)
        cb.pack(fill="x", pady=(4, 12))
        if pilots:
            cb.set(pilots[0])

        tk.Label(inner, text="Poste", bg=WHITE, fg=GRAY,
                 font=("Arial", 11)).pack(anchor="w")
        poste_var = tk.StringVar()
        postes_list = self._get_list("Postes")
        cb_poste = ttk.Combobox(inner, textvariable=poste_var,
                                values=postes_list, font=("Arial", 14),
                                state="readonly", width=28)
        cb_poste.pack(fill="x", pady=(4, 16))
        # Pré-sélectionner le dernier poste utilisé si disponible
        _last_poste = self._logged_in_poste or (postes_list[0] if postes_list else "")
        if _last_poste:
            cb_poste.set(_last_poste)

        passwords = self._get_list("Mots de passe pilote") or self._get_list("Mots de passe")
        need_pw = bool(passwords)

        if need_pw:
            tk.Label(inner, text="Mot de passe", bg=WHITE, fg=GRAY,
                     font=("Arial", 11)).pack(anchor="w")
            pw_var = tk.StringVar()
            pw_e = tk.Entry(inner, textvariable=pw_var, show="*",
                            font=("Arial", 14), width=28,
                            relief="solid", bd=1)
            pw_e.pack(fill="x", pady=(4, 16))
        else:
            pw_var = None

        err_lbl = tk.Label(inner, text="", bg=WHITE, fg=C_RED,
                           font=("Arial", 10))
        err_lbl.pack()

        def _connect(event=None):
            name = pilot_var.get().strip()
            if not name:
                err_lbl.config(text="Sélectionnez un pilote.")
                return
            poste_sel = poste_var.get().strip()
            if postes_list and not poste_sel:
                err_lbl.config(text="Sélectionnez un poste.")
                return
            if need_pw:
                pw = pw_var.get()
                try:
                    idx = pilots.index(name)
                    expected = str(passwords[idx]) if idx < len(passwords) else ""
                    if pw != expected:
                        err_lbl.config(text="Mot de passe incorrect.")
                        pw_var.set("")
                        return
                except (ValueError, IndexError):
                    err_lbl.config(text="Pilote introuvable.")
                    return
            self._logged_in_pilot = name
            self._logged_in_poste = poste_sel
            self._login_time      = datetime.datetime.now()
            self._of_count_this_shift = 0
            self._inter_of_s = 0
            ov.destroy()
            if on_success:
                on_success()
            else:
                self._show_main()

        tk.Button(inner, text="✔   SE CONNECTER", command=_connect,
                  bg=GREEN, fg=WHITE, font=("Arial", 14, "bold"),
                  relief="flat", pady=12, cursor="hand2").pack(fill="x", pady=(8, 0))
        if need_pw:
            pw_e.bind("<Return>", _connect)
        cb.bind("<Return>", _connect)
        cb.focus()

        # Bouton discret chargement fichier Excel
        def _pick_excel():
            p = filedialog.askopenfilename(
                title="Charger la base de données Excel",
                filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tous", "*.*")])
            if not p:
                return
            self.cfg["db_path"] = p
            save_cfg(self.cfg)
            self._prod_ref_cached = 0.0
            self._invalidate_wb_cache()
            self._load_lists()
            self._load_history_from_excel()
            # Rafraîchir la liste des pilotes
            pilots2 = self._get_list("Pilotes")
            cb.config(values=pilots2)
            if pilots2:
                cb.set(pilots2[0])
            name = os.path.basename(p)
            db_lbl.config(text=f"📂  {name}")

        db_current = self.cfg.get("db_path", "")
        db_name    = os.path.basename(db_current) if db_current else "Aucun fichier chargé"
        db_lbl = tk.Label(ov, text=f"📂  {db_name}", bg=BG, fg=GRAY,
                          font=("Arial", 9), cursor="hand2")
        db_lbl.pack(pady=(14, 0))
        db_lbl.bind("<Button-1>", lambda e: _pick_excel())
        tk.Label(ov, text="Cliquer pour changer de fichier",
                 bg=BG, fg=LGRAY, font=("Arial", 8)).pack()

    def _maybe_show_login(self):
        if not self._logged_in_pilot:
            self._show_login_overlay()

    def _do_logout_to_main(self):
        """Déconnecte le pilote courant et réaffiche l'écran principal (sans relogin auto)."""
        self._logged_in_pilot = None
        self._logged_in_poste = None
        self._login_time      = None
        try:
            self._save_session()
        except Exception:
            pass
        self._show_main()

    def _check_nettoyage_before_logout(self, on_confirmed):
        """Vérifie si le pilote courant a déclaré un nettoyage aujourd'hui.
        Si oui, appelle on_confirmed() directement.
        Sinon, affiche un overlay d'avertissement."""
        if not self._logged_in_pilot:
            on_confirmed()
            return
        if self._prod_active:
            ov_p = tk.Frame(self.root, bg=WHITE)
            ov_p.place(relx=0, rely=0, relwidth=1, relheight=1)
            ov_p.lift()
            hdr_p = tk.Frame(ov_p, bg=ORANGE, height=80)
            hdr_p.pack(fill="x")
            hdr_p.pack_propagate(False)
            tk.Frame(hdr_p, bg=DARK, width=6).pack(side="left", fill="y")
            tk.Label(hdr_p, text="⚠  Production en cours",
                     bg=ORANGE, fg=WHITE,
                     font=("Arial", 20, "bold")).pack(side="left", padx=20, pady=20)
            body_p = tk.Frame(ov_p, bg=WHITE)
            body_p.pack(fill="both", expand=True, padx=60, pady=40)
            tk.Label(body_p,
                     text="Une déclaration de production est en cours.\n"
                          "Veuillez la clôturer avant de vous déconnecter.",
                     bg=WHITE, fg=DARK, font=("Arial", 14),
                     justify="center").pack(pady=(0, 30))
            tk.Button(body_p, text="↩  Retour",
                      command=ov_p.destroy,
                      bg=NAVY, fg=WHITE,
                      font=("Arial", 13, "bold"), relief="flat",
                      padx=20, pady=12, cursor="hand2").pack()
            return
        today = datetime.date.today()
        has_nettoyage = any(
            ev.get("key") == "nettoyage"
            and ev.get("start") is not None
            and ev["start"].date() == today
            for ev in self._tl_events
        )
        if has_nettoyage:
            on_confirmed()
            return

        # Overlay d'avertissement nettoyage non déclaré
        ov = tk.Frame(self.root, bg=WHITE)
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        ov.lift()

        hdr_ov = tk.Frame(ov, bg=C_RED, height=80)
        hdr_ov.pack(fill="x")
        hdr_ov.pack_propagate(False)
        tk.Frame(hdr_ov, bg=DARK, width=6).pack(side="left", fill="y")
        tk.Label(hdr_ov, text="⚠  ATTENTION — Nettoyage non déclaré",
                 bg=C_RED, fg=WHITE,
                 font=("Arial", 20, "bold")).pack(side="left", padx=20, pady=20)

        body_ov = tk.Frame(ov, bg=WHITE)
        body_ov.pack(fill="both", expand=True, padx=60, pady=30)

        tk.Label(body_ov,
                 text="Vous n'avez pas encore déclaré d'arrêt nettoyage aujourd'hui.",
                 bg=WHITE, fg=DARK, font=("Arial", 14)).pack(pady=(0, 30))

        btn_row = tk.Frame(body_ov, bg=WHITE)
        btn_row.pack()

        def _ajouter_nettoyage():
            ov.destroy()
            # Demander la durée du nettoyage
            dur_ov = tk.Frame(self.root, bg=WHITE)
            dur_ov.place(relx=0, rely=0, relwidth=1, relheight=1)
            dur_ov.lift()

            dur_hdr = tk.Frame(dur_ov, bg=NAVY, height=70)
            dur_hdr.pack(fill="x")
            dur_hdr.pack_propagate(False)
            tk.Frame(dur_hdr, bg=GREEN, width=6).pack(side="left", fill="y")
            tk.Label(dur_hdr, text="Déclarer un arrêt Nettoyage",
                     bg=NAVY, fg=WHITE,
                     font=("Arial", 18, "bold")).pack(side="left", padx=20, pady=16)

            dur_body = tk.Frame(dur_ov, bg=WHITE)
            dur_body.pack(fill="both", expand=True, padx=80, pady=40)

            tk.Label(dur_body, text="Durée du nettoyage (minutes) :",
                     bg=WHITE, fg=DARK, font=("Arial", 14)).pack(pady=(0, 10))
            dur_var = tk.StringVar()
            dur_err = tk.Label(dur_body, text="", bg=WHITE, fg=C_RED, font=("Arial", 10))
            dur_err.pack()
            dur_entry = tk.Entry(dur_body, textvariable=dur_var,
                                 font=("Arial", 22, "bold"), width=8,
                                 justify="center", relief="solid", bd=2, fg=NAVY)
            dur_entry.pack(pady=8)
            dur_entry.focus()

            def _valider_dur(ev=None):
                try:
                    minutes = int(dur_var.get().strip())
                    if minutes <= 0:
                        raise ValueError
                except (ValueError, TypeError):
                    dur_err.config(text="Entrez un nombre de minutes valide (entier > 0)")
                    return
                now_ts = datetime.datetime.now()
                start_ts = now_ts - datetime.timedelta(minutes=minutes)
                self._tl_events.append({
                    "key": "nettoyage",
                    "cat": "ratt",
                    "start": start_ts,
                    "end": now_ts,
                    "comment": "Nettoyage déclaré au changement de poste",
                })
                self._save_session()
                dur_ov.destroy()
                on_confirmed()

            dur_entry.bind("<Return>", _valider_dur)
            tk.Button(dur_body, text="✔   Valider",
                      command=_valider_dur,
                      bg=GREEN, fg=WHITE, font=("Arial", 14, "bold"),
                      relief="flat", pady=12, cursor="hand2").pack(
                      fill="x", padx=80, pady=(8, 0))

        def _confirmer_sans():
            ov.destroy()
            on_confirmed()

        tk.Button(btn_row,
                  text="✔  Ajouter un nettoyage maintenant",
                  command=_ajouter_nettoyage,
                  bg=GREEN, fg=WHITE,
                  font=("Arial", 13, "bold"), relief="flat",
                  padx=20, pady=14, cursor="hand2").pack(pady=6, fill="x")
        tk.Button(btn_row,
                  text="✕  Confirmer sans nettoyage",
                  command=_confirmer_sans,
                  bg=LGRAY, fg=DARK,
                  font=("Arial", 13), relief="flat",
                  padx=20, pady=12, cursor="hand2").pack(pady=6, fill="x")
        tk.Button(btn_row,
                  text="↩  Annuler",
                  command=ov.destroy,
                  bg=NAVY, fg=WHITE,
                  font=("Arial", 13), relief="flat",
                  padx=20, pady=12, cursor="hand2").pack(pady=6, fill="x")

    def _pilot_badge(self, parent, bg):
        """Encart 'Pilote connecté: XXX' avec bouton Changer."""
        dark_bg = bg in (NAVY, NAVY_L, DARK)
        fg_sub  = "#aabbd0" if dark_bg else GRAY
        fg_main = WHITE     if dark_bg else NAVY
        f = tk.Frame(parent, bg=bg)
        name = self._logged_in_pilot or "Non connecté"
        col = GREEN if self._logged_in_pilot else C_RED
        tk.Label(f, text="Pilote :", bg=bg, fg=fg_sub,
                 font=("Arial", 9)).pack(side="left", padx=(0, 2))
        tk.Label(f, text=name, bg=bg, fg=fg_main,
                 font=("Arial", 11, "bold")).pack(side="left")
        tk.Button(f, text="⇄", bg=col, fg=WHITE,
                  font=("Arial", 10, "bold"), relief="flat",
                  padx=6, cursor="hand2",
                  command=lambda: self._check_nettoyage_before_logout(
                      lambda: self._show_login_overlay(on_success=self._show_main)
                  )).pack(side="left", padx=(6, 0))
        return f

    def _show_production_header_refresh(self):
        """Rafraîchit juste le badge pilote dans la vue production."""
        self._show_main() if not self._prod_active else None

    def _refresh_all(self):
        """Relit tout le fichier Excel en arrière-plan et met à jour tous les onglets."""
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            self._load_lists()
            _toast(self.root, "⚠  Fichier Excel introuvable", bg=C_RATT, duration=2500)
            return
        self._invalidate_wb_cache()
        self._load_lists()
        _toast(self.root, "⏳  Lecture du fichier Excel…", bg=NAVY_L, duration=1500)

        def _bg():
            rows, events, trs_data = None, None, None
            ok = False
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                rows = []
                if "Data" in wb.sheetnames:
                    ws = wb["Data"]
                    min_r = 2 if str(ws.cell(1, 1).value or "").strip().upper() == "OF" else 1
                    for i, r in enumerate(ws.iter_rows(min_row=min_r, values_only=True), start=min_r):
                        if any(r):
                            rows.append((i, list(r) + [None] * 60))
                events = []
                if "Evenements" in wb.sheetnames:
                    ws_e = wb["Evenements"]
                    for r in ws_e.iter_rows(min_row=2, values_only=True):
                        if r and any(r):
                            events.append(list(r))
                trs_data = []
                if "TRS" in wb.sheetnames:
                    ws_t = wb["TRS"]
                    for r in ws_t.iter_rows(min_row=2, values_only=True):
                        if r and any(r):
                            trs_data.append(list(r))
                wb.close()
                ok = True
            except Exception as e:
                self.root.after(0, lambda err=str(e): _toast(
                    self.root, f"⚠  Erreur lecture Excel : {err}", bg=C_RED, duration=3000))
            if ok:
                self.root.after(0, lambda: self._on_refresh_done(rows, events, trs_data))

        threading.Thread(target=_bg, daemon=True).start()

    def _on_refresh_done(self, rows, events, trs_data):
        """Appelé depuis le thread principal après relecture complète du fichier Excel."""
        # Mise à jour des caches
        if rows is not None:
            self._data_rows_cache = rows
        if events is not None:
            self._events_cache = events
        if trs_data is not None:
            self._trs_cache = trs_data
        # Refresh de tous les onglets visibles
        self._refresh_table()
        self._refresh_main_kpi()
        self._refresh_events_tab()
        self._refresh_postes_tab()
        if self._mode == "production":
            self._refresh_stops_recap()
        _toast(self.root, "✔  Données actualisées", bg=GREEN, duration=2000)

    def _confirm_quit(self):
        """Vérifie que tout est en ordre avant de quitter l'application."""
        issues = []

        # 1. Production active non terminée
        if self._prod_active and self._of_start is not None:
            issues.append(("prod", "Une production est en cours et n'a pas été déclarée."))

        # 2. Arrêts encore ouverts
        running_stops = [k for k in self._timers if self._t_running(k)]
        if running_stops:
            labels = [next((e[0] for e in EVENTS if e[1] == k), k) for k in running_stops]
            issues.append(("stops", f"{len(running_stops)} arrêt(s) encore en cours : {', '.join(labels)}"))

        # 3. Déclaration en attente (Excel verrouillé)
        if os.path.exists(PENDING_FILE):
            issues.append(("pending", "Une déclaration est en attente d'écriture (Excel verrouillé)."))

        # 4. Nettoyage non déclaré aujourd'hui
        if self._logged_in_pilot:
            today = datetime.date.today()
            has_activity_today = any(
                ev.get("start") and ev["start"].date() == today
                for ev in self._tl_events
            )
            if has_activity_today:
                has_nettoyage = any(
                    ev.get("key") == "nettoyage"
                    and ev.get("start") is not None
                    and ev["start"].date() == today
                    for ev in self._tl_events
                )
                if not has_nettoyage:
                    issues.append(("nettoyage", "Aucun arrêt nettoyage déclaré aujourd'hui."))

        if not issues:
            # Tout est OK → quitter directement
            self.root.destroy()
            return

        # Afficher un overlay de confirmation avec la liste des problèmes
        ov = tk.Frame(self.root, bg=WHITE)
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        ov.lift()

        hdr_ov = tk.Frame(ov, bg=C_RED, height=80)
        hdr_ov.pack(fill="x")
        hdr_ov.pack_propagate(False)
        tk.Frame(hdr_ov, bg=DARK, width=6).pack(side="left", fill="y")
        tk.Label(hdr_ov, text="⚠  Attention avant de quitter",
                 bg=C_RED, fg=WHITE,
                 font=("Arial", 20, "bold")).pack(side="left", padx=20, pady=20)

        body_ov = tk.Frame(ov, bg=WHITE)
        body_ov.pack(fill="both", expand=True, padx=60, pady=24)

        tk.Label(body_ov,
                 text="Les points suivants nécessitent votre attention :",
                 bg=WHITE, fg=DARK, font=("Arial", 13, "bold")).pack(anchor="w", pady=(0, 16))

        ICONS = {"prod": "🔴", "stops": "🟠", "pending": "🟡", "nettoyage": "🔵"}
        for kind, msg in issues:
            row = tk.Frame(body_ov, bg=WHITE)
            row.pack(anchor="w", pady=4, fill="x")
            tk.Label(row, text=ICONS.get(kind, "•"), bg=WHITE,
                     font=("Arial", 14)).pack(side="left", padx=(0, 10))
            tk.Label(row, text=msg, bg=WHITE, fg=DARK,
                     font=("Arial", 12), justify="left").pack(side="left")

        # Boutons
        btn_frame = tk.Frame(body_ov, bg=WHITE)
        btn_frame.pack(pady=(32, 0), anchor="w")

        def _force_quit():
            self._t_stop_all()
            self._tl_close_all()
            # Si production active, écrire la déclaration dans Excel avant de quitter
            if self._prod_active and self._of_start:
                try:
                    end_dt  = datetime.datetime.now()
                    of_s    = (end_dt - self._of_start).total_seconds()
                    stop_s  = self._t_wall_clock_stops()
                    # Récupérer les valeurs du formulaire (si disponibles)
                    v = {}
                    if hasattr(self, "fv"):
                        v = {k: var.get().strip() for k, var in self.fv.items()}
                    if hasattr(self, "_comment_txt"):
                        try:
                            v["comment"] = self._comment_txt.get("1.0", "end").strip()
                        except Exception:
                            pass
                    # Sinon, utiliser les données sauvegardées
                    for k, val in self._saved_form_data.items():
                        if k not in v:
                            v[k] = str(val)
                    def _n(k):
                        try: return int(v.get(k, 0) or 0)
                        except: return 0
                    qte_fab = _n("qte_fab")
                    nb_pers = max(1, _n("nb_pers") or 1)
                    of_min  = int(of_s) / 60
                    of_hrs  = int(of_s) / 3600
                    equiv = self._calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
                    c1 = round(equiv / of_hrs, 2)   if of_hrs > 0 else 0
                    c2 = round(equiv / (nb_pers * of_hrs), 2) if of_hrs > 0 else 0
                    prod_ref = self._get_prod_ref()
                    kit = 2 if getattr(self, "_v_kit", None) and self._v_kit.get() else 1
                    def _ts(key): return fmt(self._t_get(key))
                    row = [
                        v.get("of_num",""), datetime.date.today().strftime("%d/%m/%Y"),
                        v.get("poste",""), v.get("pilote", self._logged_in_pilot or ""),
                        v.get("copilote",""), v.get("nb_pers",""), v.get("taille",""),
                        v.get("code_prod",""), v.get("type_prod",""), v.get("poids",""),
                        v.get("fibre",""), v.get("of_taie",""), v.get("traca",""),
                        qte_fab, _n("qte_emb"), equiv, fmt(of_s),
                        self._of_start.strftime("%H:%M:%S"), end_dt.strftime("%H:%M:%S"),
                        c1, c2, kit, v.get("ref_taie",""),
                        _n("qte_init_taie"), _n("nb_taie2_choix"),
                        _n("nb_def_cout"), _n("mq_taie"),
                        _n("mq_housse_encart"),
                        _n("nb_pp_cousue"),
                        fmt(self._inter_of_s),
                        _min_str_to_hms(v.get("duree_mq_mp", "")),
                        _min_str_to_hms(v.get("manquant_pers", "")),
                        fmt(sum(
                            (ev["end"] - ev["start"]).total_seconds()
                            for ev in self._tl_events
                            if ev.get("key") == "nettoyage"
                               and ev.get("start") and ev.get("end")
                        )),
                        _ts("ratt_pochon"), _ts("ratt_couture"), _ts("ratt_emb"),
                        _ts("ratt_presse_soud"), _ts("ratt_presse_zip"),
                        _ts("pb_chargeuse"), _ts("pb_carde"), _ts("pb_etaleur"),
                        _ts("pb_coupe"), _ts("pb_tapis1"), _ts("pb_enrouleur"),
                        _ts("pb_pesee"), _ts("pb_deviation"), _ts("pb_enfileur"),
                        _ts("pb_kinna"), _ts("pb_tapeuse"), _ts("pb_table_rot"),
                        _ts("pb_h100"), _ts("pb_traversin"), _ts("pb_presse_orc"),
                        _ts("pb_presse_zip2"), _ts("pb_cercleuse"), _ts("pb_enrouleuse"),
                        v.get("comment",""),
                        fmt(self._interposte_s) if self._interposte_s > 0 else "",  # BF
                    ]
                except Exception:
                    pass
            self._save_session()
            self.root.destroy()

        tk.Button(btn_frame,
                  text="✕  Quitter quand même",
                  command=_force_quit,
                  bg=C_RED, fg=WHITE,
                  font=("Arial", 13, "bold"), relief="flat",
                  padx=20, pady=12, cursor="hand2").pack(side="left", padx=(0, 16))
        tk.Button(btn_frame,
                  text="↩  Retour",
                  command=ov.destroy,
                  bg=LGRAY, fg=DARK,
                  font=("Arial", 13), relief="flat",
                  padx=20, pady=12, cursor="hand2").pack(side="left")

    def _make_header(self, parent, title, subtitle=""):
        hdr = tk.Frame(parent, bg=NAVY, height=self._px(68))
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=title, bg=NAVY, fg=WHITE,
                 font=("Arial", 18, "bold")).pack(side="left", padx=22)
        if subtitle:
            tk.Label(hdr, text=subtitle, bg=NAVY, fg="#7a99c0",
                     font=("Arial", 11)).pack(side="left", padx=4)
        right_bar = tk.Frame(hdr, bg=NAVY)
        right_bar.pack(side="right", padx=12)
        BTN = dict(font=("Arial", 11, "bold"), relief="flat", padx=12, pady=4, cursor="hand2")
        # Quitter
        tk.Button(right_bar, text="⏻  Quitter", bg=C_RED, fg=WHITE,
                  command=self._confirm_quit, **BTN).pack(side="right", padx=4)
        # Actualiser
        tk.Button(right_bar, text="🔄  Actualiser", bg=NAVY_L, fg=WHITE,
                  command=self._refresh_all, **BTN).pack(side="right", padx=4)
        # Base données (remplace _db_widget)
        tk.Button(right_bar, text="⚙  Base", bg=NAVY_L, fg=WHITE,
                  command=self._select_db, **BTN).pack(side="right", padx=4)
        # Listes Excel
        tk.Button(right_bar, text="⚙  Listes", bg=NAVY_L, fg=WHITE,
                  command=self._show_excel_info, **BTN).pack(side="right", padx=4)
        # Paramètres
        tk.Button(right_bar, text="⚙  Paramètres", bg=NAVY_L, fg=WHITE,
                  font=("Arial", 11, "bold"), relief="flat", cursor="hand2",
                  command=self._show_settings).pack(side="right", padx=4)
        # Pilote (shows name)
        pilot_name = self._logged_in_pilot or "Non connecté"
        pilot_bg   = GREEN if self._logged_in_pilot else C_RED
        # Poste en cours (visible à côté du pilote)
        if self._logged_in_poste:
            tk.Label(right_bar, text=f"🕐 {self._logged_in_poste}",
                     bg=NAVY, fg="#4ade80",
                     font=("Arial", 12, "bold")).pack(side="right", padx=(0, 8))
        # Pilote (nom)
        tk.Button(right_bar, text=f"👤  {pilot_name}", bg=pilot_bg, fg=WHITE,
                  command=lambda: self._check_nettoyage_before_logout(
                      lambda: self._show_login_overlay(on_success=self._show_main)
                  ), **BTN).pack(side="right", padx=(4, 0))
        # Bouton déconnexion dédié (n'apparait que si connecté)
        if self._logged_in_pilot:
            tk.Button(right_bar, text="⏻", bg="#dc2626", fg=WHITE,
                      command=lambda: self._check_nettoyage_before_logout(
                          lambda: self._do_logout_to_main()
                      ), **BTN).pack(side="right", padx=(0, 4))
        return hdr

    # ── Paramètres ───────────────────────────────────────────────────────────
    RULES_TEXT = """\
RÈGLES DE FONCTIONNEMENT KPI-ORC

═══ DURÉES & CALCUL TRS ═══

TRS = (Temps productif net) / (Temps d'ouverture)
Temps d'ouverture = Durée totale du poste - Arrêts planifiés

Arrêts planifiés (exclus du TRS) :
  • Pauses opérateur (100% exclues)
  • Nettoyage court ≤ durée configurée (défaut 10 min)
  • Nettoyage long ≤ durée configurée (défaut 30 min)
  • Grand nettoyage ≤ durée configurée (défaut 60 min)
  • Réunion/manquant personnel ≤ tolérance (défaut 5 min/poste)

Arrêts imputés au TRS (temps perdu) :
  • Pannes techniques (100%)
  • Rattrapages (100%)
  • Nettoyage dépassant la durée planifiée (excédent seulement)
  • Réunion/manquant personnel dépassant la tolérance (excédent seulement)
  • Manquant matière première (100%)

═══ INTERVALLE ENTRE OF ═══
  • Si l'intervalle entre deux OF est ≤ durée configurée (défaut 5 min),
    il est automatiquement classé "Changement d'OF" dans les événements.
    Il est toujours imputé au TRS du prochain OF ouvert.
    Si l'intervalle dépasse la durée configurée, une confirmation est demandée au pilote.

═══ CHANGEMENT D'OF ═══
  • Un changement d'OF crée automatiquement un événement dans l'onglet
    Événements de l'Excel (colonne A = "Changement d'OF", col C = date,
    col Q = heure début, col R = heure fin, col S = durée, col T = commentaire).
  • Le nom du pilote qui a effectué le dernier OF est écrit en colonne E.

═══ PAUSES ═══
  • Les pauses sont déclarées via "Aller en pause" et enregistrées dans
    l'onglet Événements.
  • Les pauses sont exclues du calcul TRS (arrêt planifié).
  • Le temps de pause total est affiché dans le récapitulatif de fin de poste.

═══ POSTE EN COURS vs POSTE PRÉCÉDENT ═══
  • Le KPI "Poste en cours" est calculé sur toutes les déclarations du
    pilote connecté depuis 00:00 du jour courant.
  • Le KPI "Poste précédent" est calculé sur le dernier poste clôturé
    avant le poste en cours.

═══ NETTOYAGE ═══
  • Court (fin de poste) : tolérance configurable (défaut 10 min)
  • Long (ex: mercredi) : tolérance configurable (défaut 30 min)
  • Grand nettoyage : tolérance configurable (défaut 60 min)
  • Seul le temps dépassant la tolérance est imputé au TRS.

"""

    def _show_settings(self):
        """Fenêtre paramètres protégée par mot de passe."""
        # Étape 1 : mot de passe
        settings_pw = self.cfg.get("settings_password", "2026")
        top_pw = tk.Toplevel(self.root)
        top_pw.overrideredirect(True)
        top_pw.attributes("-topmost", True)
        top_pw.configure(bg=NAVY)
        self._center_on_root(top_pw, 380, 200)
        top_pw.grab_set()

        tk.Label(top_pw, text="⚙  PARAMÈTRES", bg=NAVY, fg=WHITE,
                 font=("Arial", 16, "bold")).pack(pady=(24, 4))
        tk.Label(top_pw, text="Mot de passe paramètres :", bg=NAVY, fg="#7a99c0",
                 font=("Arial", 11)).pack()
        err_lbl = tk.Label(top_pw, text="", bg=NAVY, fg=C_RED, font=("Arial", 10))
        err_lbl.pack()
        pv = tk.StringVar()
        pe = tk.Entry(top_pw, textvariable=pv, show="*",
                      font=("Arial", 18), width=10, justify="center",
                      relief="solid", bd=2, bg=WHITE)
        pe.pack(pady=4)
        pe.focus()
        allowed = [False]

        def _confirm_pw(ev=None):
            if pv.get() == str(self.cfg.get("settings_password", "2026")):
                allowed[0] = True
                top_pw.destroy()
            else:
                err_lbl.config(text="Mot de passe incorrect")
                pv.set("")

        pe.bind("<Return>", _confirm_pw)
        btn_row_pw = tk.Frame(top_pw, bg=NAVY)
        btn_row_pw.pack(pady=8)
        tk.Button(btn_row_pw, text="✔  Valider", command=_confirm_pw,
                  bg=GREEN, fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", padx=16, pady=6, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row_pw, text="Annuler", command=top_pw.destroy,
                  bg=LGRAY, fg=DARK, font=("Arial", 10),
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left")
        top_pw.wait_window()
        if not allowed[0]:
            return

        # Étape 2 : fenêtre paramètres
        self._modal_open = True
        win = tk.Toplevel(self.root)
        win.title("Paramètres KPI-ORC")
        win.grab_set()
        win.configure(bg=BG)
        win.resizable(True, True)
        self._center_on_root(win, 820, 620)
        win.bind("<Destroy>", lambda e: setattr(self, "_modal_open", False) if e.widget is win else None)

        tk.Frame(win, bg=NAVY, height=6).pack(fill="x")
        hdr_s = tk.Frame(win, bg=NAVY, height=50)
        hdr_s.pack(fill="x")
        hdr_s.pack_propagate(False)
        tk.Label(hdr_s, text="⚙  Paramètres KPI-ORC", bg=NAVY, fg=WHITE,
                 font=("Arial", 15, "bold")).pack(side="left", padx=20, pady=10)
        tk.Button(hdr_s, text="✕", bg=NAVY, fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", cursor="hand2",
                  command=win.destroy).pack(side="right", padx=12)

        nb_s = ttk.Notebook(win)
        nb_s.pack(fill="both", expand=True, padx=8, pady=8)

        # ── Tab Général ───────────────────────────────────────────────────────
        tab_gen = tk.Frame(nb_s, bg=WHITE)
        nb_s.add(tab_gen, text="  Général  ")
        cv_gen = tk.Canvas(tab_gen, bg=WHITE, highlightthickness=0)
        sb_gen = ttk.Scrollbar(tab_gen, orient="vertical", command=cv_gen.yview)
        cv_gen.configure(yscrollcommand=sb_gen.set)
        sb_gen.pack(side="right", fill="y")
        cv_gen.pack(fill="both", expand=True)
        gen_inner = tk.Frame(cv_gen, bg=WHITE)
        gen_win = cv_gen.create_window((0, 0), window=gen_inner, anchor="nw")
        gen_inner.bind("<Configure>", lambda e: cv_gen.configure(scrollregion=cv_gen.bbox("all")))
        cv_gen.bind("<Configure>", lambda e: cv_gen.itemconfig(gen_win, width=e.width))

        gen_fields = [
            ("Nettoyage court poste (minutes)",        "clean_short_min",   10),
            ("Nettoyage long ex: mercredi (minutes)",  "clean_long_min",    30),
            ("Grand nettoyage (minutes)",              "clean_grand_min",   60),
            ("Réunion tolérée par poste (minutes)",    "meeting_tol_min",    5),
            ("Temps de pause autorisé par poste (min)", "pause_max_min",    20),
            ("Mot de passe application",               "app_password",    "0000"),
            ("Mot de passe base de données",           "db_password",     "4594"),
            ("Mot de passe paramètres",                "settings_password","2026"),
        ]
        gen_vars = {}
        for i, (label, key, default) in enumerate(gen_fields):
            row_f = tk.Frame(gen_inner, bg=WHITE)
            row_f.pack(fill="x", padx=20, pady=6)
            tk.Label(row_f, text=label, bg=WHITE, fg=DARK,
                     font=("Arial", 11), width=38, anchor="w").pack(side="left")
            val = self.cfg.get(key, default)
            var = tk.StringVar(value=str(val))
            gen_vars[key] = var
            show = "*" if "password" in key else ""
            tk.Entry(row_f, textvariable=var, font=("Arial", 12, "bold"),
                     width=12, relief="solid", bd=1, show=show).pack(side="left", padx=8)

        # ── Tab Pilotes (avec mots de passe) ─────────────────────────────────
        tab_pil = tk.Frame(nb_s, bg=WHITE)
        nb_s.add(tab_pil, text="  Pilotes  ")

        pil_names_ref = self._get_list("Pilotes")[:]
        pil_pws_ref   = self._get_list("Mots de passe pilote")[:]
        # Aligner longueurs
        while len(pil_pws_ref) < len(pil_names_ref):
            pil_pws_ref.append("")

        pil_frm = tk.Frame(tab_pil, bg=WHITE)
        pil_frm.pack(fill="both", expand=True, padx=16, pady=12)
        pil_frm.columnconfigure(0, weight=2)
        pil_frm.columnconfigure(1, weight=1)
        pil_frm.rowconfigure(1, weight=1)

        tk.Label(pil_frm, text="Pilotes — Noms et mots de passe",
                 bg=WHITE, fg=NAVY,
                 font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=3,
                                                   sticky="w", pady=(0, 4))
        tk.Label(pil_frm, text="Nom", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w")
        tk.Label(pil_frm, text="Mot de passe", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold")).grid(row=1, column=1, sticky="w", padx=(8, 0))

        pil_scroll_frame = tk.Frame(pil_frm, bg=WHITE)
        pil_scroll_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(4, 8))
        pil_frm.rowconfigure(2, weight=1)

        pil_canvas = tk.Canvas(pil_scroll_frame, bg=WHITE, highlightthickness=0)
        pil_sb = ttk.Scrollbar(pil_scroll_frame, orient="vertical", command=pil_canvas.yview)
        pil_canvas.configure(yscrollcommand=pil_sb.set)
        pil_sb.pack(side="right", fill="y")
        pil_canvas.pack(side="left", fill="both", expand=True)

        pil_inner_frm = tk.Frame(pil_canvas, bg=WHITE)
        pil_win_id    = pil_canvas.create_window((0, 0), window=pil_inner_frm, anchor="nw")
        pil_inner_frm.bind("<Configure>",
                           lambda e: pil_canvas.configure(scrollregion=pil_canvas.bbox("all")))
        pil_canvas.bind("<Configure>",
                        lambda e: pil_canvas.itemconfig(pil_win_id, width=e.width))

        pil_row_widgets = []  # list of (name_var, pw_var, row_frame)

        def _rebuild_pil_rows():
            for w in pil_inner_frm.winfo_children():
                w.destroy()
            pil_row_widgets.clear()
            for i, (name, pw) in enumerate(zip(pil_names_ref, pil_pws_ref)):
                rf = tk.Frame(pil_inner_frm, bg=WHITE if i % 2 == 0 else "#f8f9fc")
                rf.pack(fill="x", pady=1)
                nv = tk.StringVar(value=name)
                pv = tk.StringVar(value=pw)
                tk.Entry(rf, textvariable=nv, font=("Arial", 11),
                         relief="solid", bd=1, width=22).pack(side="left", padx=(0, 6), pady=2)
                tk.Entry(rf, textvariable=pv, font=("Arial", 11),
                         relief="solid", bd=1, width=14).pack(side="left", padx=(0, 6), pady=2)
                def _del_row(idx=i):
                    if 0 <= idx < len(pil_names_ref):
                        pil_names_ref.pop(idx)
                        pil_pws_ref.pop(idx)
                        _rebuild_pil_rows()
                tk.Button(rf, text="✕", command=_del_row,
                          bg=C_RED, fg=WHITE, font=("Arial", 9, "bold"),
                          relief="flat", padx=4, pady=1, cursor="hand2").pack(side="left")
                pil_row_widgets.append((nv, pv))

        _rebuild_pil_rows()

        def _apply_pil_edits():
            pil_names_ref.clear()
            pil_pws_ref.clear()
            for nv, pv in pil_row_widgets:
                n = nv.get().strip()
                p = pv.get().strip()
                if n:
                    pil_names_ref.append(n)
                    pil_pws_ref.append(p)
            _rebuild_pil_rows()

        def _add_pil():
            _apply_pil_edits()
            pil_names_ref.append("Nouveau pilote")
            pil_pws_ref.append("")
            _rebuild_pil_rows()

        pil_add_f = tk.Frame(pil_frm, bg=WHITE)
        pil_add_f.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        tk.Button(pil_add_f, text="+ Ajouter un pilote", command=_add_pil,
                  bg=GREEN, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(pil_add_f, text="↺ Appliquer les modifications",
                  command=_apply_pil_edits,
                  bg=NAVY_L, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        # ── Tab Copilotes ─────────────────────────────────────────────────────
        tab_cop = tk.Frame(nb_s, bg=WHITE)
        nb_s.add(tab_cop, text="  Copilotes  ")

        def _make_list_tab(tab_frame, list_name):
            """Onglet générique gestion de liste simple."""
            frm = tk.Frame(tab_frame, bg=WHITE)
            frm.pack(fill="both", expand=True, padx=16, pady=12)
            frm.columnconfigure(0, weight=1)
            frm.rowconfigure(1, weight=1)
            tk.Label(frm, text=f"Liste : {list_name}", bg=WHITE, fg=NAVY,
                     font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2,
                                                       sticky="w", pady=(0, 6))
            lb = tk.Listbox(frm, font=("Arial", 11), bg=WHITE, fg=DARK,
                            relief="solid", bd=1, selectmode="single")
            lb.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
            sb_lb = ttk.Scrollbar(frm, orient="vertical", command=lb.yview)
            sb_lb.grid(row=1, column=1, sticky="ns", pady=(0, 8))
            lb.configure(yscrollcommand=sb_lb.set)
            items = self._get_list(list_name)[:]
            for it in items:
                lb.insert("end", it)
            def _reload():
                lb.delete(0, "end")
                for it2 in items:
                    lb.insert("end", it2)
            add_f = tk.Frame(frm, bg=WHITE)
            add_f.grid(row=2, column=0, columnspan=2, sticky="ew")
            add_var = tk.StringVar()
            tk.Entry(add_f, textvariable=add_var, font=("Arial", 11),
                     relief="solid", bd=1, width=24).pack(side="left", padx=(0, 6))
            def _add():
                val2 = add_var.get().strip()
                if val2 and val2 not in items:
                    items.append(val2)
                    _reload()
                    add_var.set("")
            def _del():
                sel = lb.curselection()
                if sel:
                    items.pop(sel[0])
                    _reload()
            tk.Button(add_f, text="+ Ajouter", command=_add,
                      bg=GREEN, fg=WHITE, font=("Arial", 10, "bold"),
                      relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left", padx=4)
            tk.Button(add_f, text="✕ Supprimer", command=_del,
                      bg=C_RED, fg=WHITE, font=("Arial", 10, "bold"),
                      relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")
            return items

        cop_items_ref = _make_list_tab(tab_cop, "Co-pilotes")

        # ── Tab Postes ────────────────────────────────────────────────────────
        tab_post = tk.Frame(nb_s, bg=WHITE)
        nb_s.add(tab_post, text="  Postes  ")

        postes_defaults = [("Matin", 8, 0), ("Midi", 8, 0), ("Nuit", 8, 0), ("Jour", 8, 0)]
        postes_vars = {}  # nom -> (h_var, m_var)
        poste_durees_cfg = self.cfg.get("postes_durees", {})

        post_frm = tk.Frame(tab_post, bg=WHITE)
        post_frm.pack(fill="both", expand=True, padx=16, pady=12)
        tk.Label(post_frm, text="Durées des postes", bg=WHITE, fg=NAVY,
                 font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 8))

        post_hdr = tk.Frame(post_frm, bg=WHITE)
        post_hdr.pack(fill="x", pady=(0, 4))
        tk.Label(post_hdr, text="Poste", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold"), width=14, anchor="w").pack(side="left")
        tk.Label(post_hdr, text="Heures", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold"), width=8, anchor="w").pack(side="left")
        tk.Label(post_hdr, text="Minutes", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold"), width=8, anchor="w").pack(side="left")

        for nom, dh, dm in postes_defaults:
            total_min = poste_durees_cfg.get(nom, dh * 60 + dm)
            h_val = total_min // 60
            m_val = total_min % 60
            h_var = tk.StringVar(value=str(h_val))
            m_var = tk.StringVar(value=str(m_val))
            postes_vars[nom] = (h_var, m_var)
            prow = tk.Frame(post_frm, bg=WHITE if postes_defaults.index((nom, dh, dm)) % 2 == 0 else "#f8f9fc")
            prow.pack(fill="x", pady=2)
            tk.Label(prow, text=nom, bg=prow.cget("bg"), fg=DARK,
                     font=("Arial", 11), width=14, anchor="w").pack(side="left")
            tk.Entry(prow, textvariable=h_var, font=("Arial", 11),
                     width=6, relief="solid", bd=1, justify="center").pack(side="left", padx=4)
            tk.Label(prow, text="h", bg=prow.cget("bg"), fg=GRAY,
                     font=("Arial", 11)).pack(side="left", padx=(0, 8))
            tk.Entry(prow, textvariable=m_var, font=("Arial", 11),
                     width=6, relief="solid", bd=1, justify="center").pack(side="left", padx=4)
            tk.Label(prow, text="min", bg=prow.cget("bg"), fg=GRAY,
                     font=("Arial", 11)).pack(side="left")

        # ── Tab Règles de calcul ──────────────────────────────────────────────
        tab_rules = tk.Frame(nb_s, bg=WHITE)
        nb_s.add(tab_rules, text="  Règles de calcul  ")
        rules_txt = tk.Text(tab_rules, font=("Courier", 10), bg="#f8f8f8", fg=DARK,
                            relief="flat", padx=12, pady=10, wrap="word",
                            state="normal")
        rules_txt.pack(fill="both", expand=True, padx=8, pady=8)
        rules_txt.insert("1.0", self.RULES_TEXT)
        rules_txt.config(state="disabled")

        # ── Barre de sauvegarde ───────────────────────────────────────────────
        btm_s = tk.Frame(win, bg=BG)
        btm_s.pack(fill="x", padx=8, pady=8)

        def _save_settings():
            # Appliquer les éditions en cours dans le tableau pilotes
            _apply_pil_edits()
            _pw_keys = {"app_password", "db_password", "settings_password"}
            for key, var in gen_vars.items():
                val2 = var.get().strip()
                if key in _pw_keys:
                    self.cfg[key] = val2  # toujours string pour les mots de passe
                else:
                    try:
                        self.cfg[key] = int(val2)
                    except (ValueError, TypeError):
                        self.cfg[key] = val2
            # Sauvegarder les durées de postes
            postes_dict = {}
            for nom, (hv, mv) in postes_vars.items():
                try:
                    total = int(hv.get() or 0) * 60 + int(mv.get() or 0)
                except Exception:
                    total = 480
                postes_dict[nom] = total
            self.cfg["postes_durees"] = postes_dict
            save_cfg(self.cfg)
            # Mettre à jour self.lists
            self.lists["Pilotes"] = list(pil_names_ref)
            self.lists["Mots de passe pilote"] = list(pil_pws_ref)
            self.lists["Co-pilotes"] = list(cop_items_ref)
            # Écrire dans Excel (nouvelle structure colonnes A-J)
            path = self.cfg.get("db_path", "")
            if path and os.path.exists(path):
                try:
                    wb_s = load_workbook(path)
                    if "Listes" not in wb_s.sheetnames:
                        wb_s.create_sheet("Listes")
                    ws_l = wb_s["Listes"]
                    # Lire les valeurs existantes pour les colonnes D-J
                    existing_rows = list(ws_l.iter_rows(min_row=2, values_only=True))
                    def _col(rows, ci):
                        return [r[ci] if r and len(r) > ci else None for r in rows]
                    taille_vals  = _col(existing_rows, 3)
                    type_vals    = _col(existing_rows, 4)
                    equiv_vals   = _col(existing_rows, 5)
                    postes_vals  = _col(existing_rows, 6)
                    nbpers_vals  = _col(existing_rows, 7)
                    prodref_vals = _col(existing_rows, 8)
                    fibres_vals  = _col(existing_rows, 9)
                    # Effacer et réécrire
                    ws_l.delete_rows(1, ws_l.max_row)
                    headers = ["Pilotes", "Mots de passe pilote", "Co-pilotes",
                               "Taille produit", "Type de produit", "Equivalence Coef",
                               "Postes", "Nb personnes", "Prod de reference", "Fibres"]
                    ws_l.append(headers)
                    self._format_row(ws_l, 1)
                    max_r = max(len(pil_names_ref), len(cop_items_ref),
                                len(taille_vals) if taille_vals else 0, 1)
                    for i in range(max_r):
                        row_data = [
                            pil_names_ref[i] if i < len(pil_names_ref) else None,
                            pil_pws_ref[i]   if i < len(pil_pws_ref)   else None,
                            cop_items_ref[i] if i < len(cop_items_ref)  else None,
                            taille_vals[i]   if i < len(taille_vals)    else None,
                            type_vals[i]     if i < len(type_vals)      else None,
                            equiv_vals[i]    if i < len(equiv_vals)     else None,
                            postes_vals[i]   if i < len(postes_vals)    else None,
                            nbpers_vals[i]   if i < len(nbpers_vals)    else None,
                            prodref_vals[i]  if i < len(prodref_vals)   else None,
                            fibres_vals[i]   if i < len(fibres_vals)    else None,
                        ]
                        ws_l.append(row_data)
                        self._format_row(ws_l, ws_l.max_row)
                    with self._excel_lock:
                        self._safe_excel_save(wb_s, path)
                    wb_s.close()
                    self._load_lists()  # Recharger
                except Exception:
                    pass
            _toast(self.root, "✔  Paramètres enregistrés", bg=GREEN, duration=2500)
            win.destroy()

        tk.Button(btm_s, text="Annuler", command=win.destroy,
                  bg=SHAD, fg=DARK, font=("Arial", 11),
                  relief="flat", padx=14, pady=5, cursor="hand2").pack(side="right", padx=4)
        tk.Button(btm_s, text="💾  Enregistrer", command=_save_settings,
                  bg=NAVY, fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", padx=20, pady=6, cursor="hand2").pack(side="right", padx=4)

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
        bar = tk.Frame(parent, bg="#d0d8e8", bd=0)
        bar.pack(fill="x")
        tk.Frame(bar, bg=NAVY, height=2).pack(fill="x", side="bottom")

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
            tab_bg  = WHITE if is_active else "#dce4f0"
            tab_fg  = NAVY  if is_active else (GRAY if enabled else LGRAY)
            relief  = "raised" if is_active else "groove"
            accent  = (GREEN if mode == "main" else C_RED) if is_active else tab_bg
            f = tk.Frame(bar, bg=tab_bg, relief=relief, bd=1)
            f.pack(side="left", padx=(0, 2), pady=(3, 0))
            tk.Frame(f, bg=accent, height=4).pack(fill="x")
            row = tk.Frame(f, bg=tab_bg)
            row.pack(padx=20, pady=8)
            lbl = tk.Label(row, text=label, bg=tab_bg, fg=tab_fg,
                           font=("Arial", 11, "bold" if is_active else "normal"))
            lbl.pack(side="left")
            if enabled and not is_active:
                _bind_recursive(f, "<Button-1>",
                                lambda e, a=action: self._transition(a))
            return row

        _tab("📊  Tableau de bord", "main", self._show_main, True)

        # Onglet Production
        is_prod_tab = (active == "production")
        fg_prod  = NAVY if is_prod_tab else (C_RED if self._prod_active else LGRAY)
        tab_bg_p = WHITE if is_prod_tab else ("#fff0f0" if self._prod_active else "#dce4f0")
        accent_p = C_RED if (is_prod_tab or self._prod_active) else tab_bg_p
        prod_f = tk.Frame(bar, bg=tab_bg_p,
                          relief="raised" if is_prod_tab else "groove", bd=1)
        prod_f.pack(side="left", padx=(0, 2), pady=(3, 0))
        tk.Frame(prod_f, bg=accent_p, height=4).pack(fill="x")
        prod_row = tk.Frame(prod_f, bg=tab_bg_p)
        prod_row.pack(padx=20, pady=8)
        if self._prod_active and not is_prod_tab:
            self._blink_dot = tk.Label(prod_row, text="● ", bg=tab_bg_p,
                                       fg=C_RED, font=("Arial", 11, "bold"))
            self._blink_dot.pack(side="left")
        tk.Label(prod_row,
                 text="⚡  Production en cours" if self._prod_active else "⚡  Production",
                 bg=tab_bg_p, fg=fg_prod,
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
        if self._mode == "production" and self._prod_active and self._of_start:
            of_s   = (now - self._of_start).total_seconds() + getattr(self, "_inter_of_s", 0)
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
        if self._mode == "main" and self._prod_active and self._of_start:
            of_s = (now - self._of_start).total_seconds() + getattr(self, "_inter_of_s", 0)
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
                and not self._modal_open
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

        # Régénère le HTML toutes les 15 secondes (supervision live)
        if self._tick_count % 15 == 1 and self.cfg.get("db_path"):
            self._generate_dashboard_html()

        if self._after_id: self.root.after_cancel(self._after_id)
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

        # Compteur pièces attendues (hors arrêts et pauses)
        _pause_elapsed = self._pause_total_s
        if getattr(self, "_is_paused", False) and self._pause_start:
            _pause_elapsed += (datetime.datetime.now() - self._pause_start).total_seconds()
        pure_s = max(0.0, of_s - stop_s - _pause_elapsed)
        prod_ref = self._get_prod_ref()
        # Adjust pcs_obj for current product equivalence coeff
        _equiv_coef = 1.0
        if hasattr(self, "fv"):
            _taille_v   = self.fv["taille"].get()   if "taille"    in self.fv else ""
            _type_pv    = self.fv["type_prod"].get() if "type_prod" in self.fv else ""
            _raw_equiv  = self._calc_equiv(1, _taille_v, _type_pv)
            if _raw_equiv > 0:
                _equiv_coef = _raw_equiv
        pcs_obj = int(prod_ref * pure_s / 28800.0 / _equiv_coef) if prod_ref > 0 else 0

        if any_running:
            cv.create_text(20, h // 2, text="⚠  EN ARRÊT",
                           font=("Arial", 20, "bold"), fill=WHITE, anchor="w")
            cv.create_text(w // 2, h // 2, text=fmt(stop_s),
                           font=("Arial", 32, "bold"), fill="#ffff44", anchor="center")
            cv.create_text(w - 20, h // 2,
                           text=f"OF: {fmt(of_s)}",
                           font=("Arial", 13), fill="#ffcccc", anchor="e")
        else:
            cv.create_text(20, h // 2, text="⏱",
                           font=("Arial", 16), fill=WHITE, anchor="w")
            cv.create_text(50, h // 2, text=fmt(of_s),
                           font=("Arial", 30, "bold"), fill=WHITE, anchor="w")
            # Pièces attendues
            if pcs_obj > 0:
                cv.create_text(260, h // 2 - 8, text="Objectif",
                               font=("Arial", 8), fill="#d4f5d4", anchor="w")
                cv.create_text(260, h // 2 + 8, text=f"≥ {pcs_obj} pcs",
                               font=("Arial", 13, "bold"), fill="#4ade80", anchor="w")

            # Barre PROD/ARRET 3D
            if of_s > 0:
                prod_r = max(0.0, (of_s - stop_s) / of_s)
                pct_p  = int(prod_r * 100)
                bx, bw2, by, bh2 = w - 320, 300, 8, h - 16
                # Ombre bas-droite
                cv.create_rectangle(bx+3, by+3, bx+bw2+3, by+bh2+3,
                                    fill="#1a0000", outline="")
                # Fond arrêt (rouge foncé)
                cv.create_rectangle(bx, by, bx+bw2, by+bh2,
                                    fill="#991b1b", outline="")
                # Part prod (vert foncé)
                prod_w = int(bw2 * prod_r)
                if prod_w > 0:
                    cv.create_rectangle(bx, by, bx+prod_w, by+bh2,
                                        fill="#166534", outline="")
                # Reflet lumineux haut (3D)
                ref_h = max(3, bh2 // 3)
                if prod_w > 0:
                    cv.create_rectangle(bx+1, by+1, bx+prod_w-1, by+ref_h,
                                        fill="#22c55e", outline="")
                cv.create_rectangle(bx+prod_w, by+1, bx+bw2-1, by+ref_h,
                                    fill="#ef4444", outline="")
                # Bord brillant gauche
                cv.create_rectangle(bx, by, bx+2, by+bh2, fill="#4ade80", outline="")
                # Textes (ombre simulée : décalage +1/+1 en noir puis texte blanc)
                cv.create_text(bx+11, by+bh2//2+1,
                               text=f"PROD {pct_p}%",
                               font=("Arial", 12, "bold"), fill="#003300", anchor="w")
                cv.create_text(bx+10, by+bh2//2,
                               text=f"PROD {pct_p}%",
                               font=("Arial", 12, "bold"), fill=WHITE, anchor="w")
                if pct_p < 95:
                    cv.create_text(bx+bw2-7, by+bh2//2+1,
                                   text=f"ARRÊT {100-pct_p}%",
                                   font=("Arial", 11, "bold"), fill="#330000", anchor="e")
                    cv.create_text(bx+bw2-8, by+bh2//2,
                                   text=f"ARRÊT {100-pct_p}%",
                                   font=("Arial", 11, "bold"), fill=WHITE, anchor="e")

    def _reset_activity(self, event=None):
        self._last_activity = datetime.datetime.now()

    def _transition(self, fn):
        fn()

    # ── Mot de passe ─────────────────────────────────────────────────────────
    def _check_password(self, action=""):
        top = tk.Toplevel(self.root)
        top.title("Mot de passe")
        top.resizable(False, False)
        top.grab_set()
        top.update_idletasks()
        self._center_on_root(top, 320, 160)

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
        btn_row_chk = tk.Frame(top, bg=top.cget("bg"))
        btn_row_chk.pack(pady=6)
        tk.Button(btn_row_chk, text="Valider", command=confirm,
                  bg=NAVY, fg=WHITE, font=("Arial", 10, "bold"),
                  relief="flat", padx=16, pady=4,
                  cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row_chk, text="Annuler", command=top.destroy,
                  bg=LGRAY, fg=DARK, font=("Arial", 10),
                  relief="flat", padx=12, pady=4,
                  cursor="hand2").pack(side="left")
        top.wait_window()
        return result.get()

    # ── Vérification fichier Excel au démarrage ───────────────────────────────
    def _check_db_on_startup(self):
        path = self.cfg.get("db_path", "")
        if path and os.path.exists(path):
            return
        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=C_RED)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        pw, ph = min(560, sw - 60), 280
        top.geometry(f"{pw}x{ph}+{(sw-pw)//2}+{(sh-ph)//2}")

        tk.Label(top, text="⚠  AUCUN FICHIER EXCEL CHARGÉ",
                 bg=C_RED, fg=WHITE, font=("Arial", 18, "bold")).pack(pady=(28, 4))
        tk.Label(top, text="Appelez le Bureau Méthodes\npour connecter la base de données.",
                 bg=C_RED, fg=WHITE, font=("Arial", 13),
                 justify="center").pack(pady=(0, 16))

        pw2 = tk.StringVar()
        err_lbl = tk.Label(top, text="", bg=C_RED, fg="#ffdddd", font=("Arial", 10))
        err_lbl.pack()
        pe = tk.Entry(top, textvariable=pw2, show="*",
                      font=("Arial", 20), width=8, justify="center",
                      relief="solid", bd=2)
        pe.pack(pady=4)
        pe.focus()

        def _confirm(ev=None):
            if pw2.get() == DB_PASSWORD:
                top.destroy()
                self._select_db()
            else:
                err_lbl.config(text="Mot de passe incorrect")
                pw2.set("")

        pe.bind("<Return>", _confirm)
        tk.Button(top, text="Valider  →  Charger le fichier",
                  command=_confirm, bg=DARK, fg=WHITE,
                  font=("Arial", 13, "bold"), relief="flat",
                  pady=10, cursor="hand2").pack(fill="x", padx=40, pady=(4, 20))
        top.wait_window()

    # =========================================================================
    #  ECRAN PRINCIPAL
    # =========================================================================
    def _show_main(self):
        if not self._reset_form_next:
            self._save_form_data()
        self._reset_form_next = False
        self._mode = "main"
        self._clear()
        self._show_loading("Chargement du tableau de bord…")
        path = self.cfg.get("db_path", "")

        def _bg():
            rows = []
            events = []
            trs_data = []
            if path and os.path.exists(path):
                try:
                    wb = load_workbook(path, read_only=True, data_only=True)
                    if "Data" in wb.sheetnames:
                        ws = wb["Data"]
                        min_r = 2 if str(ws.cell(1, 1).value or "").strip().upper() == "OF" else 1
                        for i, r in enumerate(ws.iter_rows(min_row=min_r, values_only=True), start=min_r):
                            if any(r):
                                rows.append((i, list(r) + [None] * 60))
                    if "Evenements" in wb.sheetnames:
                        ws_e = wb["Evenements"]
                        for r in ws_e.iter_rows(min_row=2, values_only=True):
                            if r and any(r):
                                events.append(list(r))
                    if "TRS" in wb.sheetnames:
                        ws_t = wb["TRS"]
                        for r in ws_t.iter_rows(min_row=2, values_only=True):
                            if r and any(r):
                                trs_data.append(list(r))
                    wb.close()
                except Exception:
                    pass
            final = rows[-50:] if len(rows) > 50 else rows
            self.root.after(0, lambda: self._show_main_done(final, events=events, trs_data=trs_data,
                                                             full_data=rows, full_evts=events))

        threading.Thread(target=_bg, daemon=True).start()

    def _show_main_done(self, rows, toast=None, events=None, trs_data=None,
                        full_data=None, full_evts=None):
        import time as _t
        self._data_rows_cache = rows
        self._events_cache    = events or []
        self._trs_cache       = trs_data or []
        if full_data is not None:
            self._review_full_data = full_data
            self._review_full_evts = full_evts or events or []
            self._review_full_trs  = trs_data or []
            self._review_cache_ts  = _t.time()
        self._hide_loading()
        self._build_main_ui()
        if toast:
            _toast(self.root, toast,
                   bg=GREEN if "✔" in toast else C_RATT, duration=3500)

    def _reload_and_refresh(self):
        """Relit Excel et rafraîchit tous les onglets du dashboard."""
        if self._mode != "main":
            return
        path = self._get_read_path()
        if not path or not os.path.exists(path):
            return

        def _bg():
            rows, events, trs_data = [], [], []
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                if "Data" in wb.sheetnames:
                    ws = wb["Data"]
                    min_r = 2 if str(ws.cell(1, 1).value or "").strip().upper() == "OF" else 1
                    for i, r in enumerate(ws.iter_rows(min_row=min_r, values_only=True), start=min_r):
                        if any(r):
                            rows.append((i, list(r) + [None] * 60))
                if "Evenements" in wb.sheetnames:
                    ws_e = wb["Evenements"]
                    for r in ws_e.iter_rows(min_row=2, values_only=True):
                        if r and any(r):
                            events.append(list(r))
                if "TRS" in wb.sheetnames:
                    ws_t = wb["TRS"]
                    for r in ws_t.iter_rows(min_row=2, values_only=True):
                        if r and any(r):
                            trs_data.append(list(r))
                wb.close()
            except Exception:
                pass
            final = rows[-50:] if len(rows) > 50 else rows
            self.root.after(0, lambda: self._apply_fresh_data(
                final, events, trs_data=trs_data,
                full_data=rows, full_evts=events))

        threading.Thread(target=_bg, daemon=True).start()

    def _apply_fresh_data(self, rows, events=None, trs_data=None,
                           full_data=None, full_evts=None):
        if self._mode != "main":
            return
        self._data_rows_cache = rows
        if full_data is not None:
            import time as _t
            self._review_full_data = full_data
            self._review_full_evts = full_evts or events or []
            self._review_full_trs  = trs_data or []
            self._review_cache_ts  = _t.time()
        if events is not None:
            self._events_cache = events
        if trs_data is not None:
            self._trs_cache = trs_data
        self._refresh_table()
        self._refresh_main_kpi()
        self._refresh_events_tab()
        self._refresh_postes_tab()
        try:
            self._generate_dashboard_html()
        except Exception:
            pass

    def _build_main_ui(self):
        self._cells = []
        self._db_labels.clear()
        self._last_activity = datetime.datetime.now()

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)
        self._outer_frame = outer
        outer.bind("<Motion>",  self._reset_activity)
        outer.bind("<Button-1>", self._reset_activity)

        self._make_header(outer, "KPI-ORC", "Ligne ORC1")
        self.root.after(200, self._check_db_on_startup)
        self.root.after(3000, self._check_pending)
        self._make_tabs(outer, "main")
        self._make_timeline(outer)

        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=12)

        top = tk.Frame(body, bg=BG, height=self._px(230))
        top.pack(fill="x", pady=(0, 8))
        top.pack_propagate(False)

        # ── Zone TRS : Poste en cours + Poste précédent ──────────────────────
        trs_wrap, trs_inner = shadow_frame(top, bg=WHITE)
        trs_wrap.pack(side="left", fill="y", padx=(0, 14))
        trs_inner.columnconfigure(0, weight=1)
        trs_inner.columnconfigure(1, weight=1)

        # Colonne gauche : poste en cours
        col_cur = tk.Frame(trs_inner, bg=WHITE)
        col_cur.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=6)
        tk.Frame(trs_inner, bg=LGRAY, width=1).grid(row=0, column=1, sticky="ns", pady=8)
        tk.Label(col_cur, text="Poste en cours", bg=WHITE, fg=NAVY,
                 font=("Arial", 9, "bold")).pack(pady=(6, 0))
        self._main_gauge = Gauge(col_cur, bg=WHITE,
                                 width=self._px(130), height=self._px(80), highlightthickness=0)
        self._main_gauge.pack(pady=(0, 2))
        self._pilot_name_lbl = tk.Label(col_cur, text="",
                                        bg=WHITE, fg=DARK, font=("Arial", 9, "bold"))
        self._pilot_name_lbl.pack()
        self._trs_calc_lbl = tk.Label(col_cur, text="TRS non calculé",
                                      bg=WHITE, fg=GRAY, font=("Arial", 8))
        self._trs_calc_lbl.pack(pady=(0, 4))

        # Colonne droite : poste précédent
        col_prev = tk.Frame(trs_inner, bg=WHITE)
        col_prev.grid(row=0, column=2, sticky="nsew", padx=(4, 8), pady=6)
        tk.Label(col_prev, text="Poste précédent", bg=WHITE, fg=GRAY,
                 font=("Arial", 9, "bold")).pack(pady=(6, 0))
        self._prev_gauge = Gauge(col_prev, bg=WHITE,
                                 width=self._px(130), height=self._px(80), highlightthickness=0)
        self._prev_gauge.pack(pady=(0, 2))
        self._prev_pilot_lbl = tk.Label(col_prev, text="—",
                                        bg=WHITE, fg=GRAY, font=("Arial", 9))
        self._prev_pilot_lbl.pack(pady=(0, 4))

        self._pilot_kpi_data = {}
        self._refresh_main_kpi()

        # Zone droite — remplit tout l'espace restant
        right_zone = tk.Frame(top, bg=BG)
        right_zone.pack(side="left", fill="both", expand=True)

        def _make_canvas_btn(parent, text, color, cmd, expand=True, pady=4, height=None):
            """Bouton canvas arrondi style Dodo."""
            cv = tk.Canvas(parent, highlightthickness=0, bg=BG,
                           **({"height": height} if height else {}))
            cv.pack(fill="x" if height else "both", expand=expand, padx=2, pady=pady)
            pressed = [False]
            def _draw(e=None):
                cv.delete("all")
                bw, bh = cv.winfo_width(), cv.winfo_height()
                if bw < 10 or bh < 10:
                    return
                c = _off(color, -60) if pressed[0] else color
                _rrect(cv, 4, 5, bw-1, bh, 14, fill=_off(c, -40))
                _rrect(cv, 0, 0, bw-5, bh-5, 14, fill=c)
                _rrect(cv, 2, 2, bw-7, bh//3, 14, fill=_off(c, +45))
                cv.create_text(bw//2-2, bh//2-2, text=text, fill=WHITE,
                               font=("Arial", 12, "bold"), justify="center")
            def _press(e): pressed[0] = True; _draw()
            def _release(e): pressed[0] = False; _draw(); cmd()
            cv.bind("<Configure>", _draw)
            cv.bind("<ButtonPress-1>",  _press)
            cv.bind("<ButtonRelease-1>", _release)
            cv.config(cursor="hand2")
            return cv

        if self._prod_active:
            # Panneau "Production en cours" — affiché à gauche de right_zone
            btn_zone = tk.Frame(right_zone, bg=BG, width=self._px(230))
            btn_zone.pack(side="left", fill="y", padx=(0, 8))
            btn_zone.pack_propagate(False)

            prod_wrap, prod_inner = shadow_frame(btn_zone, bg=NAVY)
            prod_wrap.pack(fill="both", expand=True)
            self._main_prod_panel = prod_wrap

            dot_row = tk.Frame(prod_inner, bg=NAVY)
            dot_row.pack(fill="x", padx=14, pady=(12, 2))
            self._blink_dot = tk.Label(dot_row, text="●", bg=NAVY,
                                       fg=C_RED, font=("Arial", 16))
            self._blink_dot.pack(side="left", padx=(0, 6))
            tk.Label(dot_row, text="EN COURS",
                     bg=NAVY, fg=WHITE,
                     font=("Arial", 13, "bold")).pack(side="left")

            debut_str2 = self._of_start.strftime('%H:%M:%S') if self._of_start else "--:--:--"
            tk.Label(prod_inner, text=f"Début : {debut_str2}",
                     bg=NAVY, fg="#7a99c0",
                     font=("Arial", 10)).pack(anchor="w", padx=14, pady=(0, 2))

            of_s_now = (datetime.datetime.now() - self._of_start).total_seconds() if self._of_start else 0.0
            self._elapsed_lbl = tk.Label(prod_inner,
                                          text=f"⏱  {fmt(of_s_now)}",
                                          bg=NAVY, fg="#4ade80",
                                          font=("Arial", 24, "bold"))
            self._elapsed_lbl.pack(anchor="center", pady=(6, 4))

            n_actifs = sum(1 for k in self._timers if self._t_running(k))
            col_arr = C_RED if n_actifs > 0 else "#7a99c0"
            self._stops_lbl = tk.Label(prod_inner,
                text=f"⚠  {n_actifs} arrêt(s)  |  {fmt(self._t_total_stops())}",
                bg=NAVY, fg=col_arr, font=("Arial", 10, "bold"))
            self._stops_lbl.pack(anchor="center", padx=10, pady=(0, 10))

        # Panneau KPI arrêts — toujours présent
        kpi_stops = tk.Frame(right_zone, bg=BG)
        kpi_stops.pack(side="left", fill="both", expand=True)
        self._build_stops_kpi_panel(kpi_stops)

        # ── Barre de boutons d'action — 4 boutons côte à côte, toujours visibles ──
        action_bar = tk.Frame(body, bg=BG, height=66)
        action_bar.pack(fill="x", pady=(6, 2))
        action_bar.pack_propagate(False)

        def _do_logout():
            self._check_nettoyage_before_logout(
                lambda: self._show_login_overlay(on_success=self._show_main))

        def _do_connect():
            self._show_login_overlay(on_success=self._show_main)

        # Btn 1 : DÉMARRER (vert) ou grisé si prod active
        b1 = tk.Frame(action_bar, bg=BG)
        b1.pack(side="left", fill="both", expand=True)
        if self._prod_active:
            _make_canvas_btn(b1, "▶  DÉMARRER\nUNE PROD", "#b0b8c8", lambda: None)
        else:
            _make_canvas_btn(b1, "▶  DÉMARRER\nUNE PROD", GREEN, self._start_production)

        # Btn 2 : SE CONNECTER / SE DÉCONNECTER
        b2 = tk.Frame(action_bar, bg=BG)
        b2.pack(side="left", fill="both", expand=True)
        if self._logged_in_pilot:
            _make_canvas_btn(b2, "👤  SE\nDÉCONNECTER", ORANGE, _do_logout)
        else:
            _make_canvas_btn(b2, "🔑  SE\nCONNECTER", GREEN, _do_connect)

        # Btn 4 : FIN DE MON POSTE
        b4 = tk.Frame(action_bar, bg=BG)
        b4.pack(side="left", fill="both", expand=True)
        _make_canvas_btn(b4, "🏁  FIN DE\nMON POSTE", NAVY_L, self._show_fin_de_poste)

        # ── Onglets Déclarations / Événements ──────────────────────────────────
        tab_bar = tk.Frame(body, bg=BG)
        tab_bar.pack(fill="x", pady=(0, 0))

        _active_tab = [0]   # 0 = Déclarations, 1 = Événements, 2 = Postes
        tab_frames  = [None, None, None]

        def _switch_tab(idx):
            _active_tab[0] = idx
            for i, btn in enumerate(tab_btns):
                if i == idx:
                    btn.config(bg=NAVY, fg=WHITE, relief="flat")
                else:
                    btn.config(bg=LGRAY, fg=DARK, relief="flat")
            for i, frm in enumerate(tab_frames):
                if frm:
                    if i == idx:
                        frm.pack(fill="both", expand=True)
                    else:
                        frm.pack_forget()

        tab_btns = []
        for i, lbl in enumerate(["📋  Déclarations", "📊  Événements", "🏁  Postes"]):
            btn = tk.Button(tab_bar, text=lbl, bg=NAVY if i == 0 else LGRAY,
                            fg=WHITE if i == 0 else DARK,
                            font=("Arial", 10, "bold"), relief="flat",
                            padx=18, pady=5, cursor="hand2",
                            command=lambda i2=i: _switch_tab(i2))
            btn.pack(side="left", padx=(0, 2))
            tab_btns.append(btn)

        # Container for both tab panels
        tab_container = tk.Frame(body, bg=BG)
        tab_container.pack(fill="both", expand=True)

        # ── Tab 0: Déclarations ──────────────────────────────────────────────
        decl_frame = tk.Frame(tab_container, bg=BG)
        tab_frames[0] = decl_frame

        tbl_wrap, tbl_inner = shadow_frame(decl_frame, bg=WHITE)
        tbl_wrap.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("KPI.Treeview",
                        background=WHITE, foreground=DARK,
                        fieldbackground=WHITE, rowheight=32,
                        font=("Arial", 10))
        style.configure("KPI.Treeview.Heading",
                        background=LGRAY, foreground=DARK,
                        font=("Arial", 10, "bold"), relief="flat")
        style.map("KPI.Treeview", background=[("selected", "#dbeafe")])

        cols = ("Date", "H.Début", "H.Fin", "OF", "Pilote", "Poste", "Qte Fab", "Qte Emb",
                "Equiv", "TRS %", "Duree OF", "Arrêts", "✏", "🗑")
        tree = ttk.Treeview(tbl_inner, columns=cols, show="headings",
                            height=15, style="KPI.Treeview")
        widths = {"Date": 80, "H.Début": 65, "H.Fin": 65, "OF": 90, "Pilote": 130, "Poste": 80,
                  "Qte Fab": 60, "Qte Emb": 60, "Equiv": 60,
                  "TRS %": 65, "Duree OF": 75, "Arrêts": 75,
                  "✏": 52, "🗑": 52}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 70), anchor="center",
                        stretch=(c not in ("✏", "🗑")))
        style.configure("TRS.Treeview", font=("Arial", 10, "bold"))
        tree.tag_configure("trs_hi",   background="#e6f7ee", foreground=GREEN)
        tree.tag_configure("trs_warn", background="#fff7e6", foreground=ORANGE)
        tree.tag_configure("trs_low",  background="#fde8e8", foreground=C_RED)

        sb_v = ttk.Scrollbar(tbl_inner, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb_v.set)
        tree.pack(side="left", fill="both", expand=True)
        sb_v.pack(side="right", fill="y")

        self._main_tree = tree
        self._col_ids   = cols
        self._load_table(tree)

        tree.bind("<Button-1>", self._on_tree_click)

        # Show decl_frame initially
        decl_frame.pack(fill="both", expand=True)

        # ── Tab 1: Événements ────────────────────────────────────────────────
        evt_frame = tk.Frame(tab_container, bg=BG)
        tab_frames[1] = evt_frame

        evt_wrap, evt_inner = shadow_frame(evt_frame, bg=WHITE)
        evt_wrap.pack(fill="both", expand=True)

        evt_cols2 = ("Type d'événement", "Pilote", "OF", "Date",
                     "Heure début", "Heure fin", "Durée", "Commentaire")
        evt_tree2 = ttk.Treeview(evt_inner, columns=evt_cols2, show="headings",
                                  height=15, style="KPI.Treeview")
        self._evt_tree2 = evt_tree2
        evt_widths2 = {"Type d'événement": 200, "Pilote": 120, "OF": 90, "Date": 80,
                       "Heure début": 80, "Heure fin": 80, "Durée": 70, "Commentaire": 200}
        for c2 in evt_cols2:
            evt_tree2.heading(c2, text=c2)
            evt_tree2.column(c2, width=evt_widths2.get(c2, 80), anchor="center",
                             stretch=(c2 == "Commentaire"))
        sb_evt = ttk.Scrollbar(evt_inner, orient="vertical", command=evt_tree2.yview)
        evt_tree2.configure(yscrollcommand=sb_evt.set)
        evt_tree2.pack(side="left", fill="both", expand=True)
        sb_evt.pack(side="right", fill="y")

        # Configure event color tags
        evt_tree2.tag_configure("evt_pause", background="#fff8e1", foreground="#92400e")
        evt_tree2.tag_configure("evt_panne", background="#fee2e2", foreground="#991b1b")
        evt_tree2.tag_configure("evt_ratt",  background="#fff3e0", foreground="#d97706")
        evt_tree2.tag_configure("evt_chg",   background="#ede9fe", foreground="#5b21b6")
        evt_tree2.tag_configure("evt_nett",  background="#e0f2fe", foreground="#0369a1")

        # Populate events tree
        for ev_row in reversed(self._events_cache[-100:]):
            try:
                ev_type = str(ev_row[0] or "")
                ev_of   = str(ev_row[1] or "")
                ev_date = str(ev_row[2] or "")[:10]
                ev_pil  = str(ev_row[4] or "")
                ev_hd   = str(ev_row[16] or "")[:8]
                ev_hf   = str(ev_row[17] or "")[:8]
                ev_dur  = str(ev_row[18] or "")
                ev_cmt  = str(ev_row[19] or "")
                ev_type_low = ev_type.lower()
                if "pause" in ev_type_low:
                    tag = "evt_pause"
                elif "panne" in ev_type_low or "pb" in ev_type_low or "problème" in ev_type_low:
                    tag = "evt_panne"
                elif "ratt" in ev_type_low:
                    tag = "evt_ratt"
                elif "changement" in ev_type_low:
                    tag = "evt_chg"
                elif "nettoyage" in ev_type_low:
                    tag = "evt_nett"
                else:
                    tag = ""
                evt_tree2.insert("", "end", values=(
                    ev_type, ev_pil, ev_of, ev_date, ev_hd, ev_hf, ev_dur, ev_cmt),
                    tags=(tag,) if tag else ())
            except Exception:
                pass

        # ── Tab 2: Postes (TRS par poste) ────────────────────────────────────
        postes_frame = tk.Frame(tab_container, bg=BG)
        tab_frames[2] = postes_frame

        pos_wrap, pos_inner = shadow_frame(postes_frame, bg=WHITE)
        pos_wrap.pack(fill="both", expand=True)

        pos_cols = ("Date", "Poste", "Pilote", "TRS", "Qté produite",
                    "Équivalence", "Total prod", "Total panne", "Total ratt")
        pos_tree = ttk.Treeview(pos_inner, columns=pos_cols, show="headings",
                                height=15, style="KPI.Treeview")
        self._pos_tree = pos_tree
        pos_widths = {"Date": 80, "Poste": 80, "Pilote": 130, "TRS": 70,
                      "Qté produite": 90, "Équivalence": 90,
                      "Total prod": 90, "Total panne": 90, "Total ratt": 90}
        for c3 in pos_cols:
            pos_tree.heading(c3, text=c3)
            pos_tree.column(c3, width=pos_widths.get(c3, 80), anchor="center",
                            stretch=(c3 == "Pilote"))
        pos_tree.tag_configure("trs_hi",   background="#e6f7ee", foreground=GREEN)
        pos_tree.tag_configure("trs_warn", background="#fff7e6", foreground=ORANGE)
        pos_tree.tag_configure("trs_low",  background="#fde8e8", foreground=C_RED)
        sb_pos = ttk.Scrollbar(pos_inner, orient="vertical", command=pos_tree.yview)
        pos_tree.configure(yscrollcommand=sb_pos.set)
        pos_tree.pack(side="left", fill="both", expand=True)
        sb_pos.pack(side="right", fill="y")

        # Populate from _trs_cache
        for trs_row in reversed(self._trs_cache[-100:]):
            try:
                t_date  = str(trs_row[0] or "")[:10]
                t_poste = str(trs_row[1] or "")
                t_pilot = str(trs_row[2] or "")
                t_trs   = str(trs_row[19] or "")
                t_qte   = str(trs_row[16] or "")
                t_equiv = str(trs_row[17] or "")
                t_prod  = str(trs_row[8] or "")
                t_panne = str(trs_row[11] or "")
                t_ratt  = str(trs_row[12] or "")
                # Tag TRS
                try:
                    trs_num = float(str(t_trs).replace("%", "").replace(",", "."))
                    pos_tag = ("trs_hi",) if trs_num >= 75 else (("trs_warn",) if trs_num >= 55 else ("trs_low",))
                except Exception:
                    pos_tag = ()
                pos_tree.insert("", "end", values=(
                    t_date, t_poste, t_pilot, t_trs, t_qte,
                    t_equiv, t_prod, t_panne, t_ratt), tags=pos_tag)
            except Exception:
                pass

        if self._after_id: self.root.after_cancel(self._after_id)
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
        # Pré-remplir le poste depuis la connexion si pas encore dans saved_form_data
        if self._logged_in_poste and "poste" not in self._saved_form_data:
            if "poste" in self.fv:
                try:
                    self.fv["poste"].set(self._logged_in_poste)
                except Exception:
                    pass
        if not self._saved_form_data:
            return
        for k, v in self._saved_form_data.items():
            if k.startswith("_"):
                continue
            if k == "pilote" and self._logged_in_pilot:
                continue  # Ne pas écraser le pilote connecté
            if k == "poste" and self._logged_in_poste:
                continue  # Le poste vient de la connexion
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
        # Re-forcer le disabled sur poste/pilote après restauration
        cb_widgets = getattr(self, "_form_cb_widgets", {})
        if self._logged_in_poste and "poste" in cb_widgets:
            try:
                self.fv["poste"].set(self._logged_in_poste)
                cb_widgets["poste"].config(state="disabled")
            except Exception:
                pass
        if self._logged_in_pilot and "pilote" in cb_widgets:
            try:
                self.fv["pilote"].set(self._logged_in_pilot)
                cb_widgets["pilote"].config(state="disabled")
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
            return equiv / expected * 100.0
        except Exception:
            return -1

    def _get_prod_ref(self):
        """Référence de production 8h — lue depuis config.json (priorité absolue)."""
        cfg_val = self.cfg.get("prod_ref", 0)
        try:
            v = float(cfg_val)
            if v > 0:
                return v
        except (TypeError, ValueError):
            pass
        return self._prod_ref_cached   # fallback cache I2

    def _build_stops_kpi_panel(self, parent):
        """Panneau KPI arrêts — Poste en cours / Poste précédent."""
        wrap, inner = shadow_frame(parent, bg=WHITE)
        wrap.pack(fill="both", expand=True)

        kpi = getattr(self, "_pilot_kpi_data", {})

        inner.columnconfigure(0, weight=1)
        inner.columnconfigure(1, weight=1)
        inner.rowconfigure(0, weight=0)
        inner.rowconfigure(1, weight=1)

        def _make_col(col_idx, title, pilot_name, trs_val, stops):
            accent = GREEN if trs_val >= 70 else (ORANGE if trs_val >= 50 else C_RED)
            # En-tête colonne
            hdr = tk.Frame(inner, bg=NAVY, height=36)
            hdr.grid(row=0, column=col_idx, sticky="ew",
                     padx=(0 if col_idx == 0 else 2, 0))
            hdr.grid_propagate(False)
            lbl_t = f"{title}"
            if pilot_name:
                lbl_t += f"  —  {pilot_name}"
            tk.Label(hdr, text=lbl_t, bg=NAVY, fg=WHITE,
                     font=("Arial", 9, "bold"), anchor="w").pack(
                     side="left", padx=8, pady=6)

            # Badge TRS
            trs_col = GREEN if trs_val >= 70 else (ORANGE if trs_val >= 50 else C_RED)
            tk.Label(hdr, text=f"TRS {trs_val:.0f}%", bg=trs_col, fg=WHITE,
                     font=("Arial", 9, "bold"), padx=6, pady=2).pack(
                     side="right", padx=6, pady=5)

            # Corps arrêts
            body_col = tk.Frame(inner, bg=WHITE)
            body_col.grid(row=1, column=col_idx, sticky="nsew",
                          padx=(0 if col_idx == 0 else 2, 0), pady=2)

            if not stops:
                tk.Label(body_col, text="Aucun arrêt déclaré",
                         bg=WHITE, fg=LGRAY,
                         font=("Arial", 9)).pack(pady=12, padx=6, anchor="w")
                return

            sorted_stops = sorted(stops.items(), key=lambda x: -x[1]["dur"])
            max_dur = max(d["dur"] for d in stops.values()) or 1

            for lbl, data in sorted_stops[:6]:  # Max 6 lignes
                c = C_RATT if data["cat"] == "ratt" else (NAVY_L if data["cat"] == "pause" else C_RED)
                rf = tk.Frame(body_col, bg=WHITE)
                rf.pack(fill="x", padx=4, pady=1)
                tk.Frame(rf, bg=c, width=4).pack(side="left", fill="y")
                # Nom court (enlève "Rattrapage: " / "PB Technique: ")
                short = lbl.split(": ", 1)[-1] if ": " in lbl else lbl
                tk.Label(rf, text=short, bg=WHITE, fg=DARK,
                         font=("Arial", 8), anchor="w").pack(
                         side="left", padx=(4, 0), fill="x", expand=True)
                tk.Label(rf, text=f"×{data['count']}  {fmt(data['dur'])}",
                         bg=WHITE, fg=GRAY, font=("Arial", 8),
                         anchor="e").pack(side="right", padx=4)

        last_p  = kpi.get("last_pilot") or ""
        prev_p  = kpi.get("prev_pilot") or ""
        last_s  = kpi.get("last_stops", {})
        prev_s  = kpi.get("prev_stops", {})
        last_t  = kpi.get("last_trs",  0.0)
        prev_t  = kpi.get("prev_trs",  0.0)

        _make_col(0, "Poste en cours",   last_p, last_t, last_s)
        _make_col(1, "Poste précédent",  prev_p, prev_t, prev_s)

    def _refresh_main_kpi(self):
        rows = [r for _, r in self._data_rows_cache]
        today = datetime.date.today().strftime("%d/%m/%Y")
        last_pilot  = ""
        last_dt_str = ""
        pilot_trs   = 0.0

        # Borne depuis connexion (gère postes de nuit chevauchant minuit)
        login_cutoff = self._login_time or (
            datetime.datetime.now() - datetime.timedelta(hours=12))

        def _row_dt_main(row):
            try:
                return datetime.datetime.strptime(
                    f"{_row_date(row[1])} {_row_time(row[17])}", "%d/%m/%Y %H:%M:%S")
            except Exception:
                return None

        last_pilot = (self._logged_in_pilot if self._logged_in_pilot
                      else "")
        if not last_pilot:
            for row in reversed(rows):
                dt = _row_dt_main(row)
                if dt and dt >= login_cutoff:
                    p = str(row[3] or "").strip()
                    if p:
                        last_pilot = p
                        break

        prod_ref = self._get_prod_ref()
        tot_eq, tot_s = 0.0, 0.0
        for row in rows:
            dt = _row_dt_main(row)
            p = str(row[3] or "").strip()
            if not dt or dt < login_cutoff or p != last_pilot:
                continue
            try:
                tot_eq += float(str(row[15] or 0).replace(",", "."))
                tot_s  += _hms_to_sec(_row_time(row[16]) or "00:00:00")
            except Exception:
                pass
            try:
                d_s = _row_date(row[1])
                t_s = _row_time(row[18])[:5]
                if d_s and t_s:
                    last_dt_str = f"{d_s} à {t_s}"
            except Exception:
                pass
        if prod_ref > 0 and tot_s > 0:
            pilot_trs = tot_eq / (prod_ref * tot_s / 28800.0) * 100.0

        self._load_pilot_kpi(rows, today)

        self._main_gauge.update_gauge(pilot_trs)
        calc_text = (f"TRS calculé le {last_dt_str}" if last_dt_str else "TRS non calculé")
        if hasattr(self, "_trs_calc_lbl") and self._trs_calc_lbl.winfo_exists():
            self._trs_calc_lbl.config(text=calc_text)
        kpi = getattr(self, "_pilot_kpi_data", {})
        cur_poste = kpi.get("last_poste", "")
        if hasattr(self, "_pilot_name_lbl") and self._pilot_name_lbl.winfo_exists():
            parts = []
            if last_pilot:
                parts.append(f"Pilote : {last_pilot}")
            if cur_poste:
                parts.append(cur_poste)
            self._pilot_name_lbl.config(text="  |  ".join(parts) if parts else "—")
        prev_trs = kpi.get("prev_trs", 0.0)
        if hasattr(self, "_prev_gauge") and self._prev_gauge:
            try:
                self._prev_gauge.update_gauge(prev_trs)
            except Exception:
                pass
        prev_pilot = kpi.get("prev_pilot", "")
        prev_poste = kpi.get("prev_poste", "")
        if hasattr(self, "_prev_pilot_lbl") and self._prev_pilot_lbl.winfo_exists():
            parts2 = []
            if prev_pilot:
                parts2.append(f"Pilote : {prev_pilot}")
            if prev_poste:
                parts2.append(prev_poste)
            self._prev_pilot_lbl.config(text="  |  ".join(parts2) if parts2 else "—")
        if hasattr(self, "_stops_lbl") and self._stops_lbl:
            try:
                self._stops_lbl.config(text="")
            except Exception:
                pass

    def _refresh_postes_tab(self):
        """Vide et repeuple l'onglet Postes depuis _trs_cache."""
        tree = getattr(self, "_pos_tree", None)
        if not tree or not tree.winfo_exists():
            return
        for item in tree.get_children():
            tree.delete(item)
        for trs_row in reversed(self._trs_cache[-100:]):
            try:
                t_date  = str(trs_row[0] or "")[:10]
                t_poste = str(trs_row[1] or "")
                t_pilot = str(trs_row[2] or "")
                t_trs   = str(trs_row[19] or "")
                t_qte   = str(trs_row[16] or "")
                t_equiv = str(trs_row[17] or "")
                t_prod  = str(trs_row[8] or "")
                t_panne = str(trs_row[11] or "")
                t_ratt  = str(trs_row[12] or "")
                try:
                    trs_num = float(str(t_trs).replace("%", "").replace(",", "."))
                    pos_tag = ("trs_hi",) if trs_num >= 75 else (("trs_warn",) if trs_num >= 55 else ("trs_low",))
                except Exception:
                    pos_tag = ()
                tree.insert("", "end", values=(
                    t_date, t_poste, t_pilot, t_trs, t_qte,
                    t_equiv, t_prod, t_panne, t_ratt), tags=pos_tag)
            except Exception:
                pass

    def _refresh_events_tab(self):
        """Vide et repeuple l'onglet Événements depuis _events_cache."""
        tree = getattr(self, "_evt_tree2", None)
        if not tree or not tree.winfo_exists():
            return
        for item in tree.get_children():
            tree.delete(item)
        for ev_row in reversed(self._events_cache[-100:]):
            try:
                ev_type     = str(ev_row[0] or "")
                ev_of       = str(ev_row[1] or "")
                ev_date     = str(ev_row[2] or "")[:10]
                ev_pil      = str(ev_row[4] or "")
                ev_hd       = str(ev_row[16] or "")[:8]
                ev_hf       = str(ev_row[17] or "")[:8]
                ev_dur      = str(ev_row[18] or "")
                ev_cmt      = str(ev_row[19] or "")
                ev_type_low = ev_type.lower()
                if "pause" in ev_type_low:
                    tag = "evt_pause"
                elif "panne" in ev_type_low or "pb" in ev_type_low or "problème" in ev_type_low:
                    tag = "evt_panne"
                elif "ratt" in ev_type_low:
                    tag = "evt_ratt"
                elif "changement" in ev_type_low:
                    tag = "evt_chg"
                elif "nettoyage" in ev_type_low:
                    tag = "evt_nett"
                else:
                    tag = ""
                tree.insert("", "end", values=(
                    ev_type, ev_pil, ev_of, ev_date, ev_hd, ev_hf, ev_dur, ev_cmt),
                    tags=(tag,) if tag else ())
            except Exception:
                pass

    def _load_pilot_kpi(self, rows, today):
        """Calcule les KPIs par pilote (connecté et précédent, 12 dernières heures)."""
        path   = self.cfg.get("db_path", "")
        cutoff = datetime.datetime.now() - datetime.timedelta(hours=12)

        # Si le pilote vient de se connecter, ne montrer que ses données depuis sa connexion
        login_cutoff = self._login_time if self._login_time else cutoff

        def _row_dt(row):
            try:
                return datetime.datetime.strptime(
                    f"{_row_date(row[1])} {_row_time(row[17])}",
                    "%d/%m/%Y %H:%M:%S")
            except Exception:
                return None

        # Pilote en cours = connecté, sinon dernier déclarant dans les 12h
        last_pilot = self._logged_in_pilot or None
        if not last_pilot:
            for row in reversed(rows):
                dt = _row_dt(row)
                if dt and dt >= cutoff:
                    p = str(row[3] or "").strip()
                    if p:
                        last_pilot = p
                        break

        # Pilote précédent = dernier différent de last_pilot AVANT la connexion actuelle
        prev_pilot = None
        for row in reversed(rows):
            dt = _row_dt(row)
            p  = str(row[3] or "").strip()
            if not dt or not p or p == last_pilot:
                continue
            # Le poste précédent doit être antérieur à la connexion du pilote actuel
            if dt < login_cutoff:
                prev_pilot = p
                break

        prod_ref = self._get_prod_ref()
        totals   = {p: {"eq": 0.0, "s": 0.0, "poste": ""}
                    for p in [last_pilot, prev_pilot] if p}

        for row in rows:
            dt = _row_dt(row)
            p  = str(row[3] or "").strip()
            if not dt or p not in totals:
                continue
            if p == last_pilot:
                # Poste en cours : depuis la connexion (gère minuit)
                if dt < login_cutoff:
                    continue
            else:
                # Poste précédent : entre 24h avant et la connexion actuelle
                if dt >= login_cutoff or dt < cutoff:
                    continue
            try:
                totals[p]["eq"] += float(str(row[15] or 0).replace(",", "."))
                totals[p]["s"]  += _hms_to_sec(_row_time(row[16]) or "00:00:00")
                if not totals[p]["poste"]:
                    totals[p]["poste"] = str(row[2] or "").strip()
            except Exception:
                pass

        def _trs_from_cache(pilot_name):
            """Lit le TRS poste depuis la feuille TRS (même valeur que l'onglet Postes).
            Colonnes: 0=Date, 1=Poste, 2=Pilote, 19=TRS"""
            for row in reversed(getattr(self, "_trs_cache", [])):
                if not row or len(row) < 20:
                    continue
                if str(row[2] or "").strip() == pilot_name:
                    try:
                        return float(str(row[19] or "").replace("%", "").replace(",", ".").strip())
                    except Exception:
                        pass
            return 0.0

        last_stops, prev_stops = {}, {}
        # Read stop totals directly from Data rows (cols 33-55)
        _stop_keys = [
            ("ratt_pochon",    "Rattrapage: Pochon / Fibre",      "ratt"),
            ("ratt_couture",   "Rattrapage: Couture",              "ratt"),
            ("ratt_emb",       "Rattrapage: Emballage",            "ratt"),
            ("ratt_presse_soud","Rattrapage: Presse Souder",       "ratt"),
            ("ratt_presse_zip","Rattrapage: Presse ZIP",           "ratt"),
            ("pb_chargeuse",   "PB Technique: Chargeuse",          "pb"),
            ("pb_carde",       "PB Technique: Carde",              "pb"),
            ("pb_etaleur",     "PB Technique: Etaleur / Tour",     "pb"),
            ("pb_coupe",       "PB Technique: Coupe / Circ.",      "pb"),
            ("pb_tapis1",      "PB Technique: Tapis Bascule",      "pb"),
            ("pb_enrouleur",   "PB Technique: Enrouleur Pochon",   "pb"),
            ("pb_pesee",       "PB Technique: Pesee / Tapis 2",    "pb"),
            ("pb_deviation",   "PB Technique: Deviation / Table",  "pb"),
            ("pb_enfileur",    "PB Technique: Enfileur Pochon",    "pb"),
            ("pb_kinna",       "PB Technique: Kinna / Stroebel",   "pb"),
            ("pb_tapeuse",     "PB Technique: Tapeuse",            "pb"),
            ("pb_table_rot",   "PB Technique: Table Rot. / Twin",  "pb"),
            ("pb_h100",        "PB Technique: Enfileuse H100",     "pb"),
            ("pb_traversin",   "PB Technique: Enfileuse Traversin","pb"),
            ("pb_presse_orc",  "PB Technique: Presse ORC",         "pb"),
            ("pb_presse_zip2", "PB Technique: Presse Housse ZIP",  "pb"),
            ("pb_cercleuse",   "PB Technique: Cercleuse",          "pb"),
            ("pb_enrouleuse",  "PB Technique: Enrouleuse Traversin","pb"),
        ]
        for row in rows:
            dt = _row_dt(row)
            p  = str(row[3] or "").strip()
            if not dt or dt < cutoff:
                continue
            target = (last_stops if p == last_pilot else (prev_stops if p == prev_pilot else None))
            if target is None:
                continue
            for ki, (key, label, cat) in enumerate(_stop_keys):
                col_idx = 33 + ki
                if col_idx < len(row) and row[col_idx]:
                    s = _hms_to_sec(str(row[col_idx]))
                    if s > 0:
                        if label not in target:
                            target[label] = {"count": 0, "dur": 0.0, "cat": cat}
                        target[label]["count"] += 1
                        target[label]["dur"]   += s
        # Nettoyage depuis col AG (index 32) du Data
        for row in rows:
            dt = _row_dt(row)
            p  = str(row[3] or "").strip()
            if not dt or dt < cutoff:
                continue
            target = (last_stops if p == last_pilot else (prev_stops if p == prev_pilot else None))
            if target is None:
                continue
            if len(row) > 32 and row[32]:
                s = _hms_to_sec(str(row[32]))
                if s > 0:
                    if "Nettoyage" not in target:
                        target["Nettoyage"] = {"count": 0, "dur": 0.0, "cat": "ratt"}
                    target["Nettoyage"]["count"] += 1
                    target["Nettoyage"]["dur"]   += s

        # Pauses depuis self._events_cache
        for ev_row in getattr(self, "_events_cache", []):
            if not ev_row or len(ev_row) < 19:
                continue
            evt_type = str(ev_row[0] or "").strip()
            if "pause" not in evt_type.lower():
                continue
            ev_date  = str(ev_row[2] or "").strip()[:10]
            ev_time  = str(ev_row[16] or "").strip()
            p        = str(ev_row[4] or "").strip()
            dur_s    = _hms_to_sec(str(ev_row[18] or ""))
            if dur_s <= 0:
                continue
            try:
                ev_dt = datetime.datetime.strptime(f"{ev_date} {ev_time}", "%d/%m/%Y %H:%M:%S")
            except Exception:
                continue
            if ev_dt < cutoff:
                continue
            target = (last_stops if p == last_pilot else (prev_stops if p == prev_pilot else None))
            if target is None:
                continue
            if "Pause pilote" not in target:
                target["Pause pilote"] = {"count": 0, "dur": 0.0, "cat": "pause"}
            target["Pause pilote"]["count"] += 1
            target["Pause pilote"]["dur"]   += dur_s

        self._pilot_kpi_data = {
            "last_pilot": last_pilot, "prev_pilot": prev_pilot,
            "last_trs": _trs_from_cache(last_pilot), "prev_trs": _trs_from_cache(prev_pilot),
            "last_stops": last_stops,     "prev_stops": prev_stops,
            "last_eq":    totals.get(last_pilot, {}).get("eq",    0.0),
            "last_s":     totals.get(last_pilot, {}).get("s",     0.0),
            "last_poste": totals.get(last_pilot, {}).get("poste", ""),
            "prev_poste": totals.get(prev_pilot, {}).get("poste", "") if prev_pilot else "",
        }

    def _load_table(self, tree):
        indexed = self._data_rows_cache
        if not indexed:
            return

        def _sd(row, idx):
            t = 0
            for i in idx:
                if i < len(row) and row[i]:
                    try:
                        t += _hms_to_sec(str(row[i]))
                    except Exception:
                        pass
            return t

        stop_cols = list(range(33, 56))
        for excel_row, row in list(reversed(indexed)):
            qte_fab = int(row[13]) if row[13] and str(row[13]).isdigit() else 0
            equiv   = float(str(row[15]).replace(",", ".")) if row[15] else 0.0
            of_s    = _hms_to_sec(str(row[16])) if row[16] else 0
            arr_s   = _sd(row, stop_cols)
            trs_str = ""
            trs_tag = ()
            if of_s > 0:
                trs_val = self._calc_trs_from_row(row)
                if trs_val >= 0:
                    trs_str = f"{trs_val:.0f}%"
                    trs_tag = ("trs_hi",) if trs_val >= 70 else (("trs_warn",) if trs_val >= 50 else ("trs_low",))
            date_str = _row_date(row[1])
            tree.insert("", "end", iid=str(excel_row), tags=trs_tag, values=(
                date_str,
                _row_time(row[17]),
                _row_time(row[18]),
                str(row[0])           if row[0]  else "",
                str(row[3])           if row[3]  else "",
                str(row[2])           if row[2]  else "",
                str(row[13])          if row[13] else "0",
                str(row[14])          if row[14] else "0",
                f"{equiv:.1f}"        if equiv   else "",
                trs_str,
                _row_time(row[16]),
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
        # Retrouver la vraie ligne Excel par contenu
        actual_idx, wb_found = self._find_excel_row_by_content(excel_row)
        if wb_found is not None:
            try:
                wb_found.close()
            except Exception:
                pass
        if not actual_idx:
            messagebox.showwarning("Introuvable",
                                   "Ligne introuvable dans le fichier Excel.\n"
                                   "Elle a peut-être déjà été supprimée.")
            return
        # Retirer du cache
        self._data_rows_cache = [(i, r) for i, r in self._data_rows_cache
                                  if i != excel_row]
        self._refresh_table()
        _toast(self.root, "⏳  Suppression en cours…", bg=NAVY_L, duration=5000)
        def _bg():
            try:
                with self._excel_lock:
                    wb = self._get_wb(path)
                    if wb is None:
                        return
                    wb["Data"].delete_rows(actual_idx)
                    self._safe_excel_save(wb, path)
                    self._wb_mtime_cache = os.path.getmtime(path)
                    self._invalidate_wb_cache()
                self.root.after(0, lambda: _toast(
                    self.root, "✔  Déclaration supprimée", bg=GREEN, duration=2500))
                self.root.after(0, self._reload_and_refresh)
            except Exception as e:
                self._invalidate_wb_cache()
                err_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror(
                    "Erreur", f"Suppression impossible:\n{err_msg}"))
        threading.Thread(target=_bg, daemon=True).start()

    # ── Edition declaration ───────────────────────────────────────────────────
    def _find_excel_row_by_content(self, cache_idx):
        """Retrouve l'index réel dans le fichier Excel en cherchant par OF+Date+H.Début."""
        cached = next((r for idx, r in self._data_rows_cache if idx == cache_idx), None)
        if cached is None:
            return None, None
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return None, cached
        of_val = str(cached[0] or "").strip()
        dt_val = str(cached[1] or "").strip()[:10]
        hs_val = str(cached[17] or "").strip()[:8]
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb["Data"]
            for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if (str(r[0] or "").strip() == of_val
                        and _row_date(r[1]) == dt_val
                        and _row_time(r[17])[:8] == hs_val):
                    return i, wb
            wb.close()
        except Exception:
            pass
        return None, None

    def _edit_declaration(self, excel_row):
        if not self._check_password("Modifier la declaration"):
            return
        path = self.cfg.get("db_path", "")

        # Récupérer la ligne du cache
        cached = next((r for idx, r in self._data_rows_cache if idx == excel_row), None)

        # Chercher la vraie ligne dans Excel par contenu
        actual_row = None
        row_data   = cached or []
        evt_rows   = []
        if path and os.path.exists(path):
            try:
                actual_idx, wb = self._find_excel_row_by_content(excel_row)
                if wb is not None:
                    if actual_idx:
                        actual_row = actual_idx
                        ws = wb["Data"]
                        row_data = [ws.cell(row=actual_idx, column=i).value
                                    for i in range(1, len(DATA_HEADERS) + 1)]
                        row_data += [None] * max(0, len(DATA_HEADERS) - len(row_data))
                    of_num  = str(row_data[0] or "")
                    of_date = str(row_data[1] or "")[:10]
                    # Normaliser la date en dd/mm/yyyy si openpyxl retourne un objet date
                    _d_raw = row_data[1]
                    if hasattr(_d_raw, 'strftime'):
                        of_date = _d_raw.strftime("%d/%m/%Y")
                    elif of_date.count('-') == 2:  # format ISO yyyy-mm-dd
                        _p = of_date.split('-')
                        of_date = f"{_p[2]}/{_p[1]}/{_p[0]}"
                    if "Evenements" in wb.sheetnames:
                        _row_pilot = str(row_data[3] or "").strip()
                        def _norm_evt_date(d):
                            s = str(d or "").strip()[:10]
                            # YYYY-MM-DD → DD/MM/YYYY
                            if len(s) == 10 and s[4] == '-':
                                p = s.split('-')
                                return f"{p[2]}/{p[1]}/{p[0]}"
                            return s
                        for r in wb["Evenements"].iter_rows(min_row=2, values_only=True):
                            if not r:
                                continue
                            _etype = str(r[0] or "").strip()
                            if _etype.lower().startswith("changement d"):
                                continue
                            _evt_of   = str(r[1] or "").strip()
                            _evt_date = _norm_evt_date(r[2])
                            _evt_pil  = str(r[4] or "").strip()
                            # Match par OF + date (inclut pauses avec OF)
                            if _evt_of == of_num and _evt_date == of_date:
                                evt_rows.append(list(r))
                            # Pauses sans OF : par date + pilote
                            elif ("pause" in _etype.lower() and not _evt_of
                                  and _evt_date == of_date and _evt_pil == _row_pilot):
                                evt_rows.append(list(r))
                    wb.close()
            except Exception as e:
                messagebox.showerror("Erreur", f"Lecture Excel:\n{e}")
                return
        elif not row_data:
            messagebox.showwarning("Introuvable", "Déclaration introuvable.")
            return

        self._open_edit_dialog(actual_row or excel_row, row_data, evt_rows)

    def _open_edit_dialog(self, excel_row, row_data, evt_rows):
        of_lbl = str(row_data[0] or "?")
        top = tk.Toplevel(self.root)
        top.title(f"Modifier OF {of_lbl}")
        top.grab_set()
        top.resizable(True, True)
        top.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        pw, ph = min(960, sw - 40), min(800, sh - 80)
        self._center_on_root(top, pw, ph)
        top.minsize(760, 500)

        # Boutons toujours visibles en bas (créés AVANT le notebook)
        btm = tk.Frame(top, bg=BG)
        btm.pack(side="bottom", fill="x", padx=10, pady=8)

        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

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
            e = tk.Entry(row_f, textvariable=var, bg="#f8f9fa", fg=DARK,
                         font=("Arial", 10), relief="flat", bd=1,
                         state="readonly", readonlybackground="#f8f9fa",
                         disabledforeground=DARK)
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
            # Reset stop columns in field_vars then recompute from evt_data
            _EVT_KEY_TO_COL = {
                "ratt_pochon": 33, "ratt_couture": 34, "ratt_emb": 35,
                "ratt_presse_soud": 36, "ratt_presse_zip": 37,
                "pb_chargeuse": 38, "pb_carde": 39, "pb_etaleur": 40,
                "pb_coupe": 41, "pb_tapis1": 42, "pb_enrouleur": 43,
                "pb_pesee": 44, "pb_deviation": 45, "pb_enfileur": 46,
                "pb_kinna": 47, "pb_tapeuse": 48, "pb_table_rot": 49,
                "pb_h100": 50, "pb_traversin": 51, "pb_presse_orc": 52,
                "pb_presse_zip2": 53, "pb_cercleuse": 54, "pb_enrouleuse": 55,
            }
            for col_idx in _EVT_KEY_TO_COL.values():
                if col_idx < len(field_vars):
                    field_vars[col_idx].set("00:00:00")
            cumul = {}
            for i, r in enumerate(evt_data):
                while len(r) < 20:
                    r.append("")
                label   = str(r[0] or "")
                hd_str  = str(r[16] or "")
                hf_str  = str(r[17] or "")
                dur_str = str(r[18] or "")
                evt_tree.insert("", "end", iid=str(i),
                                values=(label, hd_str, hf_str, dur_str))
                # Accumuler durée dans la colonne Data correspondante
                short = label.replace("Rattrapage: ", "").replace("PB Technique: ", "")
                for ev_label, ev_key, _ in EVENTS:
                    if ev_label.strip().lower() == short.strip().lower():
                        col_idx = _EVT_KEY_TO_COL.get(ev_key)
                        if col_idx is not None:
                            try:
                                cumul[col_idx] = cumul.get(col_idx, 0) + _hms_to_sec(dur_str)
                            except Exception:
                                pass
                        break
            for col_idx, secs in cumul.items():
                if col_idx < len(field_vars):
                    field_vars[col_idx].set(fmt(secs))

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

        def _save():
            path2 = self.cfg.get("db_path", "")
            if not path2:
                messagebox.showwarning("Attention", "Base de donnees non connectee.", parent=top)
                return
            # Recalculer equivalence avant de capturer les valeurs
            vals = [v.get() for v in field_vars]
            try:
                qte_f     = int(str(vals[13] or 0))
                new_equiv = self._calc_equiv(qte_f, str(vals[6] or ""), str(vals[8] or ""))
                vals[15]  = str(new_equiv)
                field_vars[15].set(str(new_equiv))
            except Exception:
                pass
            of_num2  = str(row_data[0] or "")
            # Normaliser la date OF en dd/mm/yyyy pour comparaison cohérente avec Evenements
            _d2_raw = row_data[1]
            if hasattr(_d2_raw, 'strftime'):
                of_date2 = _d2_raw.strftime("%d/%m/%Y")
            else:
                _d2_str = str(_d2_raw or "")[:10]
                if _d2_str.count('-') == 2:  # ISO yyyy-mm-dd → dd/mm/yyyy
                    _p2 = _d2_str.split('-')
                    of_date2 = f"{_p2[2]}/{_p2[1]}/{_p2[0]}"
                else:
                    of_date2 = _d2_str
            evt_snapshot = list(evt_data)  # capture avant thread
            # Fermer la fenêtre immédiatement
            top.destroy()
            _toast(self.root, "⏳  Enregistrement en cours…", bg=NAVY_L, duration=5000)

            def _bg():
                try:
                    with self._excel_lock:
                        wb2 = self._get_wb(path2)
                        if wb2 is None:
                            return
                        ws2 = wb2["Data"]
                        for col_i, val in enumerate(vals, start=1):
                            ws2.cell(row=excel_row, column=col_i).value = val
                        self._format_row(ws2, excel_row)
                        ws_e = self._ensure_events_sheet(wb2)
                        keep = []
                        for r in ws_e.iter_rows(min_row=2, values_only=True):
                            if not r:
                                continue
                            _etype = str(r[0] or "").strip().lower()
                            # Toujours garder "Changement d'OF" — pas lié à un OF spécifique
                            if _etype.startswith("changement d"):
                                keep.append(list(r))
                                continue
                            # Supprimer les événements de cet OF+date (remplacés par evt_snapshot)
                            _r_of   = str(r[1] or "")
                            _r_date = str(r[2] or "")[:10]
                            if _r_of == of_num2 and _r_date == of_date2:
                                continue  # supprimé — sera remplacé par evt_snapshot
                            keep.append(list(r))
                        if ws_e.max_row > 1:
                            ws_e.delete_rows(2, ws_e.max_row)
                        for r in keep:
                            ws_e.append(r)
                            self._format_row(ws_e, ws_e.max_row)
                        for r in evt_snapshot:
                            ws_e.append(r)
                            self._format_row(ws_e, ws_e.max_row)
                        self._safe_excel_save(wb2, path2)
                        self._wb_mtime_cache = os.path.getmtime(path2)
                    self.root.after(0, lambda: _toast(
                        self.root, "✔  Modifications enregistrées", bg=GREEN, duration=3000))
                    self.root.after(0, self._refresh_table)
                    self.root.after(0, self._refresh_main_kpi)
                except Exception as e:
                    self._invalidate_wb_cache()
                    err_msg = str(e)
                    self.root.after(0, lambda: messagebox.showerror(
                        "Erreur", f"Sauvegarde impossible:\n{err_msg}"))

            threading.Thread(target=_bg, daemon=True).start()

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
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.update_idletasks()
        self._center_on_root(dlg, 480, 230)

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
        hd_var   = tk.StringVar(value=existing[16] if existing and len(existing) > 16 else "")
        hf_var   = tk.StringVar(value=existing[17] if existing and len(existing) > 17 else "")

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
            # Construire une ligne 20 colonnes alignée sur EVT_HEADERS
            new_row = [
                type_var.get(),          # 0  Evenement
                str(row_data[0] or ""), # 1  OF
                str(row_data[1] or ""), # 2  Date
                str(row_data[2] or ""), # 3  Poste
                str(row_data[3] or ""), # 4  Pilote
                str(row_data[4] or ""), # 5  Co-Pilote
                str(row_data[5] or ""), # 6  Nb Personnes
                str(row_data[6] or ""), # 7  Taille
                str(row_data[8] or ""), # 8  Type Produit
                str(row_data[7] or ""), # 9  Code Produit
                str(row_data[10] or ""),# 10 Fibre
                str(row_data[9] or ""), # 11 Poids Garnissage
                str(row_data[11] or ""),# 12 OF Taie
                str(row_data[12] or ""),# 13 Traca Fibre
                str(row_data[22] or ""),# 14 Ref Taie
                str(row_data[21] or ""),# 15 Kit
                hd_var.get(),           # 16 Heure Debut
                hf_var.get(),           # 17 Heure Fin
                dur,                    # 18 Duree
                "",                     # 19 Commentaire
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
        if not self._logged_in_pilot:
            self._show_login_overlay(on_success=self._start_production)
            return
        now = datetime.datetime.now()
        self._inter_of_s = 0
        self._interposte_s = 0
        if (self._last_of_end is not None
                and self._last_of_pilot
                and self._last_of_pilot != (self._logged_in_pilot or "")):
            gap_p = (now - self._last_of_end).total_seconds()
            if 0 < gap_p <= 3600:
                self._interposte_s = gap_p
        if self._last_of_end is not None:
            gap = (now - self._last_of_end).total_seconds()
            if 30 < gap <= 28800:   # > 30s et <= 8h
                self._inter_of_s = gap
                do_changeof = False
                if self._last_of_pilot and self._last_of_pilot != (self._logged_in_pilot or ""):
                    h = int(gap // 3600)
                    m = int((gap % 3600) // 60)
                    ts = f"{h}h {m:02d}min" if h > 0 else f"{m}min"
                    # ── Overlay inter-poste 3 questions (changement de pilote) ─
                    result = {"same_of": None, "interposte": None}
                    OV_BG = WHITE
                    ov = tk.Frame(self.root, bg=OV_BG)
                    ov.place(relx=0, rely=0, relwidth=1, relheight=1)
                    ov.lift()

                    hdr_ov = tk.Frame(ov, bg=ORANGE, height=80)
                    hdr_ov.pack(fill="x")
                    hdr_ov.pack_propagate(False)
                    tk.Frame(hdr_ov, bg=DARK, width=6).pack(side="left", fill="y")
                    tk.Label(hdr_ov,
                             text=f"NOUVEAU POSTE  —  Pilote : {self._logged_in_pilot}",
                             bg=ORANGE, fg=WHITE,
                             font=("Arial", 20, "bold")).pack(
                             side="left", padx=20, pady=20)

                    body_ov = tk.Frame(ov, bg=OV_BG)
                    body_ov.pack(fill="both", expand=True, padx=60, pady=30)

                    tk.Label(body_ov,
                             text=f"Durée depuis fin du poste précédent :  {ts}",
                             bg=OV_BG, fg=DARK, font=("Arial", 14)).pack(pady=(0, 24))

                    # ── Question 1 ────────────────────────────────────────────
                    tk.Label(body_ov,
                             text="Question 1 : Même OF que le poste précédent ?",
                             bg=OV_BG, fg=NAVY, font=("Arial", 15, "bold")).pack(anchor="w")

                    q2_frame = tk.Frame(body_ov, bg=OV_BG)  # hidden until Q1 answered

                    def _same_of():
                        result["same_of"] = True
                        btn_same.config(relief="sunken", bg=_off(GREEN, -30))
                        btn_nouvel.config(state="disabled")
                        q2_frame.pack(fill="x", pady=(16, 0))

                    def _nouvel_of():
                        result["same_of"] = False
                        result["interposte"] = False
                        ov.destroy()

                    btn_row_q1 = tk.Frame(body_ov, bg=OV_BG)
                    btn_row_q1.pack(fill="x", pady=(10, 0))
                    btn_same = tk.Button(btn_row_q1,
                              text="✔  OUI — Même OF",
                              command=_same_of, bg=GREEN, fg=WHITE,
                              font=("Arial", 13, "bold"), relief="flat",
                              padx=20, pady=12, cursor="hand2")
                    btn_same.pack(side="left", padx=(0, 12))
                    btn_nouvel = tk.Button(btn_row_q1,
                              text="✕  NON — Nouvel OF",
                              command=_nouvel_of, bg=LGRAY, fg=DARK,
                              font=("Arial", 13), relief="flat",
                              padx=20, pady=12, cursor="hand2")
                    btn_nouvel.pack(side="left")

                    # ── Question 2 (affichée après Q1 = OUI) ─────────────────
                    tk.Label(q2_frame,
                             text=f"Question 2 : Ce temps de {ts} est-il un arrêt inter-poste ?",
                             bg=OV_BG, fg=NAVY, font=("Arial", 15, "bold")).pack(anchor="w")
                    btn_row_q2 = tk.Frame(q2_frame, bg=OV_BG)
                    btn_row_q2.pack(fill="x", pady=(10, 0))

                    def _oui_interposte():
                        result["interposte"] = True
                        ov.destroy()

                    def _non_interposte():
                        result["interposte"] = False
                        ov.destroy()

                    tk.Button(btn_row_q2,
                              text="✔  OUI — Déclarer comme arrêt inter-poste",
                              command=_oui_interposte, bg=GREEN, fg=WHITE,
                              font=("Arial", 13, "bold"), relief="flat",
                              padx=20, pady=12, cursor="hand2").pack(side="left", padx=(0, 12))
                    tk.Button(btn_row_q2,
                              text="✕  NON — Ignorer ce temps",
                              command=_non_interposte, bg=LGRAY, fg=DARK,
                              font=("Arial", 13), relief="flat",
                              padx=20, pady=12, cursor="hand2").pack(side="left")

                    self.root.wait_window(ov)

                    # Interpréter le résultat
                    if result["same_of"] and result["interposte"]:
                        do_changeof = True
                    else:
                        do_changeof = False
                else:
                    # ── Popup changement d'OF : toujours affiché, temps modifiable ─
                    h2 = int(gap // 3600); m2 = int((gap % 3600) // 60); s2 = int(gap % 60)
                    _alert_var = tk.BooleanVar(value=False)
                    self._modal_open = True
                    ov_a = tk.Frame(self.root, bg=WHITE)
                    ov_a.place(relx=0, rely=0, relwidth=1, relheight=1)
                    ov_a.lift()
                    tk.Frame(ov_a, bg=NAVY_L, height=6).pack(fill="x")
                    tk.Label(ov_a, text="⏱  Changement d'OF — Temps inter-OF",
                             bg=WHITE, fg=NAVY, font=("Arial", 16, "bold")).pack(pady=(36, 4))
                    tk.Label(ov_a,
                             text="Ce temps sera déclaré comme « Changement de série ».\n"
                                  "Vous pouvez le modifier si nécessaire.",
                             bg=WHITE, fg=DARK, font=("Arial", 12),
                             justify="center").pack(pady=(0, 16))
                    time_f = tk.Frame(ov_a, bg=WHITE)
                    time_f.pack(pady=4)
                    hv = tk.StringVar(value=str(h2))
                    mv = tk.StringVar(value=str(m2))
                    sv = tk.StringVar(value=str(s2))
                    _estyle = dict(font=("Arial", 18, "bold"), width=3,
                                   relief="solid", bd=1, justify="center")
                    tk.Entry(time_f, textvariable=hv, **_estyle).pack(side="left", padx=2)
                    tk.Label(time_f, text="h", bg=WHITE,
                             font=("Arial", 14)).pack(side="left")
                    tk.Entry(time_f, textvariable=mv, **_estyle).pack(side="left", padx=2)
                    tk.Label(time_f, text="min", bg=WHITE,
                             font=("Arial", 14)).pack(side="left")
                    tk.Entry(time_f, textvariable=sv, **_estyle).pack(side="left", padx=2)
                    tk.Label(time_f, text="sec", bg=WHITE,
                             font=("Arial", 14)).pack(side="left")
                    def _a_valider():
                        try:
                            new_gap = (int(hv.get() or 0) * 3600
                                       + int(mv.get() or 0) * 60
                                       + int(sv.get() or 0))
                            self._inter_of_s = max(0, new_gap)
                        except Exception:
                            self._inter_of_s = gap
                        ov_a.destroy()
                        _alert_var.set(True)
                    tk.Button(ov_a, text="✔  Valider — Changement de série",
                              command=_a_valider, bg=GREEN, fg=WHITE,
                              font=("Arial", 13, "bold"), relief="flat",
                              padx=24, pady=12, cursor="hand2").pack(pady=(20, 0))
                    self.root.wait_variable(_alert_var)
                    self._modal_open = False
                    do_changeof = True
                if do_changeof:
                    self._write_changement_of_excel(self._last_of_end, now)
                    # Choisir la clé timeline : interposte si existante, sinon changeof
                    interposte_key = next(
                        (e[1] for e in EVENTS if e[1] == "ratt_interposte"), "changeof")
                    self._tl_events.append({
                        "key": interposte_key, "cat": "ratt",
                        "start": self._last_of_end, "end": now
                    })
            # Si gap > 8h : nouveau départ, on ignore l'intervalle

        self._t_reset()
        self._of_start    = now
        self._prod_active = True
        self._cells       = []
        self._pause_start    = None
        self._pause_total_s  = 0.0
        self._pause_periods  = []
        self._is_paused      = False
        self._pause_overlay  = None
        if self._of_periods:
            self._of_changes.append(now)
        self._of_periods.append({"start": now, "end": None, "of_num": ""})

        # Vider le formulaire pour chaque nouveau lancement
        self._saved_form_data = {}

        # Pré-remplir réunion (5 min) si 1ère déclaration du pilote depuis >12h
        pilot = self._logged_in_pilot or ""
        if pilot:
            cutoff = now - datetime.timedelta(hours=12)
            has_recent = False
            for _, r in self._data_rows_cache:
                if str(r[3] or "").strip() != pilot:
                    continue
                try:
                    dt_r = datetime.datetime.strptime(
                        f"{str(r[1] or '').strip()[:10]} {str(r[17] or '').strip()}",
                        "%d/%m/%Y %H:%M:%S")
                    if dt_r >= cutoff:
                        has_recent = True
                        break
                except Exception:
                    pass
            if not has_recent:
                self._saved_form_data["manquant_pers"] = "5"

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
        hdr = tk.Frame(outer, bg=WHITE, height=self._px(62))
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        # Accent gauche coloré
        tk.Frame(hdr, bg=GREEN, width=6).pack(side="left", fill="y")
        tk.Label(hdr, text="KPI-ORC  |  ORC1",
                 bg=WHITE, fg=NAVY, font=("Arial", 15, "bold")).pack(
                 side="left", padx=16, pady=12)
        debut_str = self._of_start.strftime('%H:%M:%S') if self._of_start else "--:--:--"
        tk.Label(hdr, text=f"Debut : {debut_str}",
                 bg=WHITE, fg=GRAY, font=("Arial", 11)).pack(side="left")
        right_bar = tk.Frame(hdr, bg=WHITE)
        right_bar.pack(side="right", padx=12)
        BTN_P = dict(font=("Arial", 11, "bold"), relief="flat", padx=12, pady=4, cursor="hand2")
        # Quitter
        tk.Button(right_bar, text="⏻  Quitter", bg=C_RED, fg=WHITE,
                  command=self._confirm_quit, **BTN_P).pack(side="right", padx=4)
        # Actualiser
        tk.Button(right_bar, text="🔄  Actualiser", bg=NAVY_L, fg=WHITE,
                  command=self._refresh_all, **BTN_P).pack(side="right", padx=4)
        # Base données
        tk.Button(right_bar, text="⚙  Base", bg=NAVY_L, fg=WHITE,
                  command=self._select_db, **BTN_P).pack(side="right", padx=4)
        # Listes
        tk.Button(right_bar, text="⚙  Listes", bg=NAVY_L, fg=WHITE,
                  command=self._show_excel_info, **BTN_P).pack(side="right", padx=4)
        # Pilote
        pilot_name_p = self._logged_in_pilot or "Non connecté"
        pilot_bg_p   = GREEN if self._logged_in_pilot else C_RED
        tk.Button(right_bar, text=f"👤  {pilot_name_p}", bg=pilot_bg_p, fg=WHITE,
                  command=lambda: self._check_nettoyage_before_logout(
                      lambda: self._show_login_overlay(on_success=self._show_main)
                  ), **BTN_P).pack(side="right", padx=(4, 0))
        # Bouton déconnexion dédié
        if self._logged_in_pilot:
            tk.Button(right_bar, text="⏻", bg="#dc2626", fg=WHITE,
                      command=lambda: self._check_nettoyage_before_logout(
                          lambda: self._do_logout_to_main()
                      ), **BTN_P).pack(side="right", padx=(0, 4))

        # ── Barre de statut Canvas (chrono + KPI, change couleur) ────────────
        self._status_cv = tk.Canvas(outer, height=self._px(72), bg=BG, highlightthickness=0)
        self._status_cv.pack(fill="x", padx=8, pady=(4, 0))
        self._status_cv.bind("<Configure>",
                             lambda e: self._redraw_status(0, 0, False))

        self._make_tabs(outer, "production")
        self._make_timeline(outer)

        # ── Corps 33/33/33 ────────────────────────────────────────────────────
        body = tk.Frame(outer, bg=BG)
        body.pack(fill="both", expand=True, padx=6, pady=(4, 6))
        body.columnconfigure(0, weight=3)   # formulaire (50%)
        body.columnconfigure(1, weight=2)   # arrêts actifs (33%)
        body.columnconfigure(2, weight=1)   # récap (17%)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=0)

        # Zone 1 : formulaire
        left = tk.Frame(body, bg=WHITE)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 2))
        self._build_form(left)

        # Zone 2 : arrets actifs + boutons
        mid = tk.Frame(body, bg=BG)
        mid.grid(row=0, column=1, sticky="nsew", padx=2)
        self._build_right_panel(mid)

        # Zone 3 : récap arrêts de l'OF
        recap_panel = tk.Frame(body, bg=WHITE)
        recap_panel.grid(row=0, column=2, sticky="nsew", padx=(2, 0))
        self._recap_panel = recap_panel
        self._build_stops_recap(recap_panel)

        # ── Footer : bouton Annuler bas-droite ───────────────────────────────
        footer = tk.Frame(body, bg=BG)
        footer.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        tk.Frame(footer, bg=BG).pack(side="left", fill="both", expand=True)

        def _annuler_production():
            from tkinter import messagebox
            if not messagebox.askyesno(
                    "Annuler la production",
                    "Annuler TOUTES les déclarations de cette session ?\n\n"
                    "Les données seront perdues et vous revenez au tableau de bord.",
                    icon="warning"):
                return
            self._prod_active      = False
            self._of_start         = None
            self._timers.clear()
            self._tl_events        = []
            self._pause_total_s    = 0.0
            self._is_paused        = False
            self._pause_start      = None
            self._pause_periods    = []
            self._inter_of_s       = 0
            self._interposte_s     = 0
            self._of_periods       = []
            self._saved_form_data  = {}
            self._of_count_this_shift = 0
            self._last_of_end      = None
            self._delete_session()
            if self._after_id:
                self.root.after_cancel(self._after_id)
                self._after_id = None
            self._show_main()

        tk.Button(footer, text="✕  Annuler la production",
                  command=_annuler_production,
                  bg="#dc2626", fg=WHITE,
                  font=("Arial", 11, "bold"), relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="right")

        if self._after_id: self.root.after_cancel(self._after_id)
        self._after_id = self.root.after(1000, self._tick)

    # ── Formulaire compact 4 colonnes (pas de scroll) ────────────────────────
    def _build_form(self, parent):
        self.fv = {}
        self._form_cb_widgets = {}
        # Couleurs de fond par section
        BG_IDENT  = "#eef2fb"   # Bleu très clair — Identification
        BG_PROD   = "#f0f7f0"   # Vert très clair — Produit
        BG_ARRETS = "#fff8ee"   # Ambre très clair — Autres arrêts
        BG_QTE    = "#f5f0fb"   # Violet très clair — Quantités
        BG_CMT    = "#f7f7f7"   # Gris clair — Commentaire

        outer = tk.Frame(parent, bg=WHITE)
        outer.pack(fill="both", expand=True, padx=2, pady=2)

        def _make_section(bg_color, title, fg_color):
            """Crée un bloc section avec titre et fond coloré."""
            wrap = tk.Frame(outer, bg=bg_color,
                            highlightthickness=1, highlightbackground=_off(bg_color, -20))
            wrap.pack(fill="x", padx=4, pady=(4, 0))
            hdr_row = tk.Frame(wrap, bg=bg_color)
            hdr_row.pack(fill="x")
            tk.Frame(hdr_row, bg=fg_color, width=4).pack(side="left", fill="y")
            tk.Label(hdr_row, text=f"  {title}", bg=bg_color, fg=fg_color,
                     font=("Arial", 9, "bold"), pady=3).pack(side="left")
            body = tk.Frame(wrap, bg=bg_color)
            body.pack(fill="x", padx=4, pady=(0, 4))
            for col in range(4):
                body.columnconfigure(col, weight=1)
            return body

        LFONT  = ("Arial", 8)
        EFONT  = ("Arial", 10, "bold")
        CELL_H = 48
        ri     = [0]
        _cur_c = [None]  # frame courant pour placer les champs

        def fld(lbl_txt, key, ftype, lh=None, col=0, adv=True, suffix=None, bg=None):
            c   = _cur_c[0]
            cbg = bg or c.cget("bg")
            cell = tk.Frame(c, bg=cbg, height=CELL_H)
            cell.grid(row=ri[0], column=col, sticky="ew", padx=2, pady=1)
            cell.grid_propagate(False)
            cell.columnconfigure(0, weight=1)
            cell.rowconfigure(1, weight=1)
            tk.Label(cell, text=lbl_txt, bg=cbg, fg="#374151",
                     font=LFONT, anchor="w").grid(row=0, column=0, columnspan=2,
                                                  sticky="w", pady=(2, 0))
            var = tk.StringVar()
            self.fv[key] = var
            if ftype == "entry":
                e = tk.Entry(cell, textvariable=var, bg=WHITE, fg=DARK,
                             font=EFONT, relief="solid", bd=1,
                             insertbackground=DARK, width=1)
                e.grid(row=1, column=0, sticky="nsew", padx=(0, 2), pady=(0, 3))
                if suffix:
                    tk.Label(cell, text=suffix, bg=cbg, fg="#374151",
                             font=LFONT).grid(row=1, column=1, sticky="sw",
                                              padx=(2, 0), pady=(0, 3))
            else:
                cb = ttk.Combobox(cell, textvariable=var,
                                  values=self._get_list(lh) if lh else [],
                                  font=EFONT, state="readonly", height=6, width=1)
                cb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 3))
                self._form_cb_widgets[key] = cb
                if key == "pilote" and self._logged_in_pilot:
                    var.set(self._logged_in_pilot)
                    cb.config(state="disabled")
                if key == "poste" and self._logged_in_poste:
                    var.set(self._logged_in_poste)
                    cb.config(state="disabled")
            if adv:
                ri[0] += 1

        def row4(l1,k1,t1,h1, l2,k2,t2,h2, l3,k3,t3,h3, l4,k4,t4,h4,
                 s1=None, s2=None, s3=None, s4=None):
            fld(l1,k1,t1,h1, col=0, adv=False, suffix=s1)
            fld(l2,k2,t2,h2, col=1, adv=False, suffix=s2)
            fld(l3,k3,t3,h3, col=2, adv=False, suffix=s3)
            fld(l4,k4,t4,h4, col=3, adv=True,  suffix=s4)

        def row3(l1,k1,t1,h1, l2,k2,t2,h2, l3,k3,t3,h3,
                 s1=None, s2=None, s3=None):
            fld(l1,k1,t1,h1, col=0, adv=False, suffix=s1)
            fld(l2,k2,t2,h2, col=1, adv=False, suffix=s2)
            fld(l3,k3,t3,h3, col=2, adv=True,  suffix=s3)

        def row2(l1,k1,t1,h1, l2,k2,t2,h2, s1=None, s2=None):
            fld(l1,k1,t1,h1, col=0, adv=False, suffix=s1)
            fld(l2,k2,t2,h2, col=1, adv=True,  suffix=s2)

        # ── IDENTIFICATION ────────────────────────────────────────────────────
        _cur_c[0] = _make_section(BG_IDENT, "Identification", NAVY)
        ri[0] = 0
        row4("N° OF *",    "of_num",  "entry", None,
             "Poste *",    "poste",   "combo", "Postes",
             "Pilote *",   "pilote",  "combo", "Pilotes",
             "Co-Pilote",  "copilote","combo", "Co-pilotes")
        row2("Nb personnes", "nb_pers", "combo", "Nb personnes",
             "Fibre",        "fibre",   "combo", "Fibre")

        # ── PRODUIT ───────────────────────────────────────────────────────────
        _cur_c[0] = _make_section(BG_PROD, "Produit", NAVY_L)
        ri[0] = 0
        row4("Taille",         "taille",    "combo", "Taille produit",
             "Type produit",   "type_prod", "combo", "Type produit",
             "Code produit *", "code_prod", "entry", None,
             "Poids garnissage","poids",    "entry", None, s4="gr")
        c_prod = _cur_c[0]
        self._v_kit = tk.BooleanVar()
        fld("OF taie", "of_taie", "entry", None, col=0, adv=False)
        kit_cell = tk.Frame(c_prod, bg=BG_PROD, height=CELL_H)
        kit_cell.grid(row=ri[0], column=1, sticky="ew", padx=2, pady=1)
        kit_cell.grid_propagate(False)
        kit_cell.rowconfigure(1, weight=1)
        tk.Label(kit_cell, text="Options", bg=BG_PROD, fg=GRAY,
                 font=LFONT, anchor="w").grid(row=0, column=0, sticky="w", pady=(2, 0))
        tk.Checkbutton(kit_cell, text="Kit 2 pièces", variable=self._v_kit,
                       bg=BG_PROD, fg=DARK, font=("Arial", 9),
                       activebackground=BG_PROD, selectcolor=BG_PROD).grid(
                       row=1, column=0, sticky="w", padx=4)
        ri[0] += 1

        # ── AUTRES ARRÊTS ─────────────────────────────────────────────────────
        _cur_c[0] = _make_section(BG_ARRETS, "Autres arrêts", "#d97706")
        ri[0] = 0
        row2("Durée arrêt manquant MP (min)",         "duree_mq_mp",   "entry", None,
             "Arrêt manquant personne / Réunion (min)", "manquant_pers", "entry", None)

        # ── QUANTITÉS & QUALITÉ ───────────────────────────────────────────────
        _cur_c[0] = _make_section(BG_QTE, "Quantités & Qualité", GREEN)
        ri[0] = 0
        c_qte = _cur_c[0]
        # Qte fab + emb mis en valeur (vert)
        for col_i, (lbl_txt, key_s) in enumerate([
                ("Qte fabriquée *", "qte_fab"), ("Qte emballée", "qte_emb")]):
            cell = tk.Frame(c_qte, bg="#efffef", height=CELL_H)
            cell.grid(row=ri[0], column=col_i, sticky="ew", padx=2, pady=1)
            cell.grid_propagate(False)
            cell.columnconfigure(0, weight=1)
            cell.rowconfigure(1, weight=1)
            tk.Label(cell, text=lbl_txt, bg="#efffef", fg=GREEN,
                     font=("Arial", 8, "bold"), anchor="w").grid(
                     row=0, column=0, sticky="w", pady=(2, 0))
            var = tk.StringVar()
            self.fv[key_s] = var
            e = tk.Entry(cell, textvariable=var, bg="#efffef", fg=GREEN,
                         font=("Arial", 12, "bold"), relief="solid", bd=2,
                         insertbackground=GREEN, width=1,
                         highlightthickness=1, highlightbackground=GREEN,
                         highlightcolor=GREEN)
            e.grid(row=1, column=0, sticky="nsew", padx=(0, 2), pady=(0, 3))
        fld("Traça fibre",  "traca",    "entry", None, col=2, adv=False)
        fld("Réf. taie",    "ref_taie", "entry", None, col=3, adv=True)
        row4("Qte initiale taie", "qte_init_taie",   "entry", None,
             "Nb taie 2nd choix", "nb_taie2_choix",  "entry", None,
             "Nb déf. couture",   "nb_def_cout",     "entry", None,
             "Mq. taie",          "mq_taie",         "entry", None)
        row2("Mq. housse/encart (nb)", "mq_housse_encart", "entry", None,
             "Nb PP cousue",            "nb_pp_cousue",     "entry", None)

        # ── COMMENTAIRE ───────────────────────────────────────────────────────
        cmt_wrap = tk.Frame(outer, bg=BG_CMT,
                            highlightthickness=1, highlightbackground=_off(BG_CMT, -20))
        cmt_wrap.pack(fill="x", padx=4, pady=(4, 0))
        hdr_cmt = tk.Frame(cmt_wrap, bg=BG_CMT)
        hdr_cmt.pack(fill="x")
        tk.Frame(hdr_cmt, bg=GRAY, width=4).pack(side="left", fill="y")
        tk.Label(hdr_cmt, text="  Commentaire", bg=BG_CMT, fg=GRAY,
                 font=("Arial", 9, "bold"), pady=3).pack(side="left")
        self._comment_txt = tk.Text(cmt_wrap, height=2, bg=WHITE, fg=DARK,
                                     font=("Arial", 10), relief="solid", bd=1,
                                     wrap="word", insertbackground=DARK)
        self._comment_txt.pack(fill="x", padx=4, pady=(0, 4))
        # Restaurer les valeurs sauvegardées si disponibles
        self.root.after(50, self._restore_form_data)
        # Actualiser le compteur pièces quand taille/type_prod change
        def _on_prod_type_change(*_):
            if self._prod_active and self._status_cv:
                self.root.after(50, lambda: self._tick_force_redraw())
        for _key in ("taille", "type_prod"):
            if _key in self.fv:
                self.fv[_key].trace_add("write", _on_prod_type_change)
        # Sauvegarde auto session à chaque modification du formulaire (debounce 1s)
        _form_save_id = [None]
        def _on_form_change(*_):
            if _form_save_id[0]:
                try: self.root.after_cancel(_form_save_id[0])
                except Exception: pass
            _form_save_id[0] = self.root.after(1000, self._save_form_data)
        for _v in self.fv.values():
            _v.trace_add("write", _on_form_change)

    def _tick_force_redraw(self):
        """Force un recalcul du statut bar (pièces attendues) sans attendre le tick."""
        if not self._prod_active or not self._of_start:
            return
        now   = datetime.datetime.now()
        of_s  = (now - self._of_start).total_seconds()
        stop_s = self._t_wall_clock_stops()
        any_r  = any(self._t_running(k) for k in self._timers)
        self._redraw_status(of_s, stop_s, any_r)

    # ── Popup info structure fichier Excel ───────────────────────────────────
    def _show_excel_info(self):
        self._modal_open = True
        win = tk.Toplevel(self.root)
        win.title("Structure du fichier Excel")
        win.geometry("900x580")
        win.configure(bg=WHITE)
        win.attributes("-topmost", True)
        win.bind("<Destroy>", lambda e: setattr(self, "_modal_open", False) if e.widget is win else None)

        # En-tête
        hdr = tk.Frame(win, bg=NAVY, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📊  Structure du fichier Excel KPI-ORC",
                 bg=NAVY, fg=WHITE, font=("Arial", 14, "bold")).pack(
                 side="left", padx=16, pady=12)
        tk.Button(hdr, text="✕", bg=NAVY, fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", cursor="hand2",
                  command=win.destroy).pack(side="right", padx=12)

        def _open_listes():
            path = self.cfg.get("db_path", "")
            if not path or not os.path.exists(path):
                messagebox.showwarning("Attention", "Aucun fichier Excel chargé.")
                return
            pw_win = tk.Toplevel(win)
            pw_win.title("Mot de passe")
            pw_win.geometry("340x180")
            pw_win.configure(bg=WHITE)
            pw_win.attributes("-topmost", True)
            pw_win.grab_set()
            tk.Label(pw_win, text="Mot de passe requis", bg=WHITE, fg=NAVY,
                     font=("Arial", 13, "bold")).pack(pady=(20, 8))
            pw_var = tk.StringVar()
            pw_entry = tk.Entry(pw_win, textvariable=pw_var, show="*",
                                font=("Arial", 13), width=14, justify="center",
                                relief="solid", bd=1)
            pw_entry.pack(pady=4)
            pw_entry.focus()
            err = tk.Label(pw_win, text="", bg=WHITE, fg=C_RED, font=("Arial", 10))
            err.pack()
            def _confirm_pw(event=None):
                if pw_var.get() == PASSWORD:
                    pw_win.destroy()
                    try:
                        with self._excel_lock:
                            wb2 = load_workbook(path)
                            if "Listes" in wb2.sheetnames:
                                wb2.active = wb2["Listes"]
                            self._safe_excel_save(wb2, path)
                            wb2.close()
                    except Exception:
                        pass
                    try:
                        import subprocess as _sp
                        if sys.platform == "win32":
                            os.startfile(path)
                        elif sys.platform == "darwin":
                            _sp.Popen(["open", path])
                        else:
                            launched = False
                            for cmd in ["xdg-open", "libreoffice", "soffice",
                                        "gnome-open", "kde-open"]:
                                try:
                                    _sp.Popen([cmd, path])
                                    launched = True
                                    break
                                except FileNotFoundError:
                                    continue
                            if not launched:
                                messagebox.showinfo(
                                    "Ouvrez manuellement",
                                    f"Aucun programme trouvé pour ouvrir Excel.\n\n"
                                    f"Chemin du fichier :\n{path}")
                    except Exception:
                        messagebox.showinfo("Chemin", f"Ouvrez manuellement :\n{path}")
                else:
                    err.config(text="Mot de passe incorrect")
                    pw_var.set("")
            pw_entry.bind("<Return>", _confirm_pw)
            tk.Button(pw_win, text="Ouvrir Excel", command=_confirm_pw,
                      bg=GREEN, fg=WHITE, font=("Arial", 11, "bold"),
                      relief="flat", cursor="hand2", pady=6).pack(fill="x", padx=40, pady=8)

        tk.Button(hdr, text="📝  Modifier les listes Excel", bg=GREEN, fg=WHITE,
                  font=("Arial", 11, "bold"), relief="flat", cursor="hand2",
                  padx=16, pady=6,
                  command=_open_listes).pack(side="right", padx=16, pady=8)

        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        def col_letter(i):
            letters = ""
            i += 1
            while i > 0:
                i, r = divmod(i - 1, 26)
                letters = chr(65 + r) + letters
            return letters

        def make_tab(label, col_specs):
            # col_specs: list of (name, description) — col letter auto-assigned from index
            frame = tk.Frame(nb, bg=WHITE)
            nb.add(frame, text=f"  {label}  ")
            cv = tk.Canvas(frame, bg=WHITE, highlightthickness=0)
            sb = ttk.Scrollbar(frame, orient="vertical", command=cv.yview)
            cv.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            cv.pack(side="left", fill="both", expand=True)
            inner = tk.Frame(cv, bg=WHITE)
            win_id = cv.create_window((0, 0), window=inner, anchor="nw")
            inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
            cv.bind("<Configure>", lambda e: cv.itemconfig(win_id, width=e.width))
            inner.columnconfigure(1, weight=1)
            inner.columnconfigure(2, weight=2)
            # En-têtes
            for ci, txt in enumerate(["Col", "Nom de la colonne", "Description / Calcul"]):
                tk.Label(inner, text=txt, bg=LGRAY, fg=NAVY,
                         font=("Arial", 9, "bold"),
                         relief="flat", padx=4, pady=4,
                         anchor="w" if ci > 0 else "center",
                         width=5 if ci == 0 else 0).grid(
                         row=0, column=ci, sticky="nsew", padx=1, pady=1)
            for ri, (name, desc) in enumerate(col_specs, start=1):
                bg = WHITE if ri % 2 == 0 else BG
                tk.Label(inner, text=col_letter(ri - 1), bg=bg, fg=GRAY,
                         font=("Arial", 9, "bold"), width=5,
                         anchor="center", padx=4, pady=3).grid(
                         row=ri, column=0, sticky="nsew", padx=1, pady=0)
                tk.Label(inner, text=name, bg=bg, fg=DARK,
                         font=("Arial", 9, "bold"), anchor="w",
                         padx=6, pady=3).grid(row=ri, column=1, sticky="nsew", padx=1, pady=0)
                tk.Label(inner, text=desc, bg=bg, fg="#555555",
                         font=("Arial", 8), anchor="w",
                         padx=6, pady=3, wraplength=340).grid(
                         row=ri, column=2, sticky="nsew", padx=1, pady=0)
            return frame

        DATA_SPECS = [
            ("OF",                              "Numéro de l'ordre de fabrication"),
            ("Date",                            "Date de déclaration (JJ/MM/AAAA)"),
            ("Poste",                           "Nom du poste de travail"),
            ("Pilote",                          "Pilote principal du poste"),
            ("Co-Pilote",                       "Co-pilote associé"),
            ("Nb Personnes",                    "Nombre de personnes sur le poste"),
            ("Taille",                          "Taille du produit fabriqué"),
            ("Code Produit",                    "Code article du produit"),
            ("Type Produit",                    "Gamme / type de produit"),
            ("Poids Garnissage",                "Poids du garnissage (kg)"),
            ("Fibre",                           "Type de fibre utilisée"),
            ("OF Taie",                         "Numéro OF de la taie associée"),
            ("Traca Fibre",                     "Traçabilité fibre (numéro de lot)"),
            ("Qte Fabriquee",                   "Quantité de pièces produites déclarées"),
            ("Qte Emballee",                    "Quantité de pièces emballées"),
            ("Equivalence",                     "Qte Fabriquée × coef (Listes col F)"),
            ("Duree OF",                        "Durée totale OF + changement de série (HH:MM:SS)"),
            ("Heure Debut",                     "Heure de début production (HH:MM:SS)"),
            ("Heure Fin",                       "Heure de fin production (HH:MM:SS)"),
            ("Cadence/heure",                   "Qte Fabriquée ÷ Durée OF en heures"),
            ("Cadence/h/pers",                  "Cadence/heure ÷ Nb Personnes"),
            ("Kit",                             "Kit inclus (Oui / Non)"),
            ("Ref Taie",                        "Référence de la taie"),
            ("Qte Initiale Taie",               "Stock initial de taies au départ de l'OF"),
            ("Nb Taie 2nd Choix",               "Nombre de taies classées 2nd choix"),
            ("Nb Defaut Couture",               "Nombre de défauts couture détectés"),
            ("Mq Taie",                         "Pièces en attente : manque taie"),
            ("Mq Housse/Encart",                "Pièces en attente : manque housse ou encart"),
            ("Nb PP Cousue",                    "Nombre de points plastiques cousus"),
            ("Changement de Serie",             "Durée changement de série inter-OF (HH:MM:SS)"),
            ("Temps Arret Manquant MP",         "Durée arrêt manque matière première (HH:MM:SS)"),
            ("Temps Arret Manquant Pers/Réu.",  "Durée manque personnel ou réunion (HH:MM:SS)"),
            ("Nettoyage Fin de Poste",          "Durée nettoyage déclaré sur ce poste (HH:MM:SS)"),
            ("Ratt Pochon/Fibre",               "Rattrapage pochon / fibre (HH:MM:SS)"),
            ("Ratt Couture",                    "Rattrapage couture (HH:MM:SS)"),
            ("Ratt Emballage",                  "Rattrapage emballage (HH:MM:SS)"),
            ("Ratt Presse Souder",              "Rattrapage presse souder (HH:MM:SS)"),
            ("Ratt Presse ZIP",                 "Rattrapage presse ZIP (HH:MM:SS)"),
            ("PB Chargeuse",                    "Arrêt panne Chargeuse (HH:MM:SS)"),
            ("PB Carde",                        "Arrêt panne Carde (HH:MM:SS)"),
            ("PB Etaleur/Tour",                 "Arrêt panne Etaleur/Tour (HH:MM:SS)"),
            ("PB Coupe/Circ",                   "Arrêt panne Coupe/Circ (HH:MM:SS)"),
            ("PB Tapis Bascule",                "Arrêt panne Tapis Bascule (HH:MM:SS)"),
            ("PB Enrouleur Pochon",             "Arrêt panne Enrouleur Pochon (HH:MM:SS)"),
            ("PB Pesee/Tapis 2",                "Arrêt panne Pesée/Tapis 2 (HH:MM:SS)"),
            ("PB Deviation/Table",              "Arrêt panne Déviation/Table (HH:MM:SS)"),
            ("PB Enfileur Pochon",              "Arrêt panne Enfileur Pochon (HH:MM:SS)"),
            ("PB Kinna/Stroebel",               "Arrêt panne Kinna/Stroebel (HH:MM:SS)"),
            ("PB Tapeuse",                      "Arrêt panne Tapeuse (HH:MM:SS)"),
            ("PB Table Rot/Twin",               "Arrêt panne Table Rot/Twin (HH:MM:SS)"),
            ("PB Enfileuse H1",                 "Arrêt panne Enfileuse H1 (HH:MM:SS)"),
            ("PB Enfileuse Traversin",          "Arrêt panne Enfileuse Traversin (HH:MM:SS)"),
            ("PB Presse ORC",                   "Arrêt panne Presse ORC (HH:MM:SS)"),
            ("PB Presse Housse ZIP",            "Arrêt panne Presse Housse ZIP (HH:MM:SS)"),
            ("PB Cercleuse",                    "Arrêt panne Cercleuse (HH:MM:SS)"),
            ("PB Enrouleuse Traversin",         "Arrêt panne Enrouleuse Traversin (HH:MM:SS)"),
            ("Commentaire",                     "Commentaire libre du pilote"),
            ("Temps Interposte",                "Durée inter-poste : même OF, pilote différent (HH:MM:SS)"),
        ]

        EVT_SPECS = [
            ("Evenement",       "Type d'événement (Rattrapage, PB Technique, Pause pilote, Chgt de série…)"),
            ("OF",              "Numéro de l'OF en cours lors de l'événement"),
            ("Date",            "Date de l'événement (JJ/MM/AAAA)"),
            ("Poste",           "Nom du poste de travail"),
            ("Pilote",          "Pilote déclarant"),
            ("Co-Pilote",       "Co-pilote associé"),
            ("Nb Personnes",    "Nombre de personnes lors de l'événement"),
            ("Taille",          "Taille du produit en cours"),
            ("Type Produit",    "Gamme / type de produit"),
            ("Code Produit",    "Code article"),
            ("Fibre",           "Type de fibre"),
            ("Poids Garnissage","Poids du garnissage (kg)"),
            ("OF Taie",         "Numéro OF taie associé"),
            ("Traca Fibre",     "Traçabilité fibre (lot)"),
            ("Ref Taie",        "Référence taie"),
            ("Kit",             "Kit inclus (Oui/Non)"),
            ("Heure Debut",     "Heure de début de l'événement (HH:MM:SS)"),
            ("Heure Fin",       "Heure de fin de l'événement (HH:MM:SS)"),
            ("Duree",           "Durée calculée = Heure Fin − Heure Debut (HH:MM:SS)"),
            ("Commentaire",     "Commentaire libre"),
        ]

        TRS_SPECS = [
            ("Date",                        "Date du poste (JJ/MM/AAAA)"),
            ("Poste",                       "Nom du poste de travail"),
            ("Pilote",                      "Pilote principal du poste"),
            ("Co-Pilote",                   "Co-pilote associé"),
            ("Nb Personnes",                "Nombre de personnes sur le poste"),
            ("Total temps d'ouverture",     "Durée théorique du poste, paramétré dans Paramètres (HH:MM:SS)"),
            ("Total temps déclarés",        "Somme : production + arrêts + pauses + réunions (HH:MM:SS)"),
            ("Ecart ouverture/déclarés",    "max(0, F − G) — temps non déclaré (HH:MM:SS)"),
            ("Total temps de marche",       "Somme des durées OF de production pure (HH:MM:SS)"),
            ("Total arrêts prévu",          "Pauses ≤ tolérance + réunions ≤ tolérance (HH:MM:SS)"),
            ("Débordement arrêts prévu",    "Dépassement des tolérances pause/réunion (HH:MM:SS)"),
            ("Total arrêts (panne)",        "Somme tous les PB techniques du poste (HH:MM:SS)"),
            ("Total arrêt rattrapage",      "Somme tous les rattrapages du poste (HH:MM:SS)"),
            ("Total pauses",                "Durée totale pauses pilote (HH:MM:SS)"),
            ("Total réunions",              "Durée totale réunions / manque personnel (HH:MM:SS)"),
            ("Nombre d'OF",                 "Nombre d'OF déclarés sur le poste"),
            ("Nombre de pièce produites",   "Somme des quantités fabriquées (tous OF)"),
            ("Equivalence",                 "Somme des équivalences de tous les OF"),
            ("Nombre moyen de pièce par OF","R / P — moyenne pièces par OF"),
            ("TRS",                         "Equivalence / (Prod réf × Temps ouverture / 28800) × 100 (%)"),
        ]

        make_tab("Onglet Data", DATA_SPECS)
        make_tab("Onglet Evenements", EVT_SPECS)
        make_tab("Onglet TRS", TRS_SPECS)

        # ── Onglet Listes (structure réelle du fichier Excel) ─────────────────
        LISTES_SPECS = [
            ("Pilotes",             "Noms des pilotes (liste déroulante pilote)"),
            ("Mots de passe pilote","Mot de passe de chaque pilote (même ordre que col A)"),
            ("Co-pilotes",          "Noms des co-pilotes (liste déroulante)"),
            ("Taille produit",      "Valeurs de taille produit (liste déroulante)"),
            ("Type de produit",     "Gammes / types de produit (liste déroulante)"),
            ("Equivalence Coef",    "Coefficients d'équivalence par type produit (ligne 1 = en-tête)"),
            ("Postes",              "Noms des postes de travail (liste déroulante)"),
            ("Nb personnes",        "Valeurs nb personnes (liste déroulante)"),
            ("Prod de reference",   "★  Prod réf 8h — I2 = valeur lue pour le calcul TRS et cadences"),
            ("Fibres",              "Types de fibre (liste déroulante)"),
        ]
        frame_l = tk.Frame(nb, bg=WHITE)
        nb.add(frame_l, text="  Onglet Listes  ")
        cv_l = tk.Canvas(frame_l, bg=WHITE, highlightthickness=0)
        sb_l = ttk.Scrollbar(frame_l, orient="vertical", command=cv_l.yview)
        cv_l.configure(yscrollcommand=sb_l.set)
        sb_l.pack(side="right", fill="y")
        cv_l.pack(side="left", fill="both", expand=True)
        inner_l = tk.Frame(cv_l, bg=WHITE)
        win_id_l = cv_l.create_window((0, 0), window=inner_l, anchor="nw")
        inner_l.bind("<Configure>", lambda e: cv_l.configure(scrollregion=cv_l.bbox("all")))
        cv_l.bind("<Configure>", lambda e: cv_l.itemconfig(win_id_l, width=e.width))
        inner_l.columnconfigure(1, weight=1)
        inner_l.columnconfigure(2, weight=2)
        for ci, txt in enumerate(["Col", "Nom de la colonne", "Description"]):
            tk.Label(inner_l, text=txt, bg=LGRAY, fg=NAVY,
                     font=("Arial", 9, "bold"),
                     relief="flat", padx=4, pady=4,
                     anchor="w" if ci > 0 else "center",
                     width=5 if ci == 0 else 0).grid(
                     row=0, column=ci, sticky="nsew", padx=1, pady=1)
        for ri, (name, desc) in enumerate(LISTES_SPECS, start=1):
            bg = WHITE if ri % 2 == 0 else BG
            is_i2 = (ri == 9)  # Prod de reference = col I (index 8, ri=9)
            clet = col_letter(ri - 1)
            fg_col = GREEN if is_i2 else GRAY
            tk.Label(inner_l, text=clet, bg=bg, fg=fg_col,
                     font=("Arial", 9, "bold"), width=5,
                     anchor="center", padx=4, pady=3).grid(
                     row=ri, column=0, sticky="nsew", padx=1, pady=0)
            tk.Label(inner_l, text=name, bg=bg,
                     fg=GREEN if is_i2 else DARK,
                     font=("Arial", 9, "bold" if is_i2 else "normal"),
                     anchor="w", padx=6, pady=3).grid(
                     row=ri, column=1, sticky="nsew", padx=1, pady=0)
            tk.Label(inner_l, text=desc, bg=bg, fg="#555555",
                     font=("Arial", 8), anchor="w",
                     padx=6, pady=3, wraplength=340).grid(
                     row=ri, column=2, sticky="nsew", padx=1, pady=0)

    # ── Panneau droit : arrets actifs + boutons ───────────────────────────────
    def _build_right_panel(self, parent):
        # Zone arrets actifs (prend tout l'espace disponible)
        stops_frame = tk.Frame(parent, bg=BG)
        stops_frame.pack(fill="both", expand=True, padx=6, pady=(6, 4))
        self._active_stops_container = stops_frame
        self._refresh_active_stops()

        BTN_H    = 52
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

        # Couleurs modernisées : rouge / ambre / gris ardoise / vert
        _make_cv_btn("⚠   DÉCLARER UN ARRÊT / RATTRAPAGE", "#dc2626",
                     self._show_stop_selector)
        _make_cv_btn("🧹  DÉCLARER UN ARRÊT NETTOYAGE", "#f59e0b",
                     self._show_nettoyage_selector)
        _make_cv_btn("☕  JE VAIS EN PAUSE", "#64748b",
                     self._toggle_pause)
        _make_cv_btn("⏹   DÉCLARER LA FIN DE PRODUCTION", "#16a34a",
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
        hdr_f = tk.Frame(parent, bg="#c0392b", relief="raised", bd=2)
        hdr_f.pack(fill="x")
        tk.Frame(hdr_f, bg="#e74c3c", height=3).pack(fill="x", side="top")
        inner_hdr = tk.Frame(hdr_f, bg="#c0392b")
        inner_hdr.pack(fill="x", padx=6, pady=4)
        tk.Label(inner_hdr, text="RÉCAP ARRÊTS OF",
                 bg="#c0392b", fg=WHITE, font=("Arial", 9, "bold")).pack(side="left")
        tk.Button(inner_hdr, text="✏", bg="#e74c3c", fg=WHITE,
                  font=("Arial", 9, "bold"), relief="flat", cursor="hand2",
                  command=self._open_stops_editor).pack(side="right", padx=2)
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
        # Ajouter les pauses
        pause_total = self._pause_total_s
        if self._is_paused and self._pause_start:
            pause_total += (now - self._pause_start).total_seconds()

        has_any = bool(cumuls) or pause_total > 0
        if not has_any:
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
                     font=("Arial", 11), anchor="w",
                     wraplength=150).pack(anchor="w")
            tk.Label(name_f, text=f"× {info['n']}  —  {dur_s}",
                     bg=WHITE, fg=color, font=("Arial", 11, "bold"),
                     anchor="w").pack(anchor="w")
        if pause_total > 0:
            pm = int(pause_total // 60)
            ps = int(pause_total % 60)
            pn = len(self._pause_periods) + (1 if self._is_paused else 0)
            row_f = tk.Frame(inner, bg=WHITE)
            row_f.pack(fill="x", pady=1, padx=2)
            tk.Frame(row_f, bg=NAVY_L, width=4).pack(side="left", fill="y")
            name_f = tk.Frame(row_f, bg=WHITE)
            name_f.pack(side="left", fill="both", expand=True, padx=(4, 0))
            tk.Label(name_f, text="☕ Pause pilote", bg=WHITE, fg=DARK,
                     font=("Arial", 11), anchor="w").pack(anchor="w")
            tk.Label(name_f,
                     text=f"× {pn}  —  {pm}min {ps:02d}s",
                     bg=WHITE, fg=NAVY_L, font=("Arial", 11, "bold"),
                     anchor="w").pack(anchor="w")
        inter_of = getattr(self, "_inter_of_s", 0)
        if inter_of > 0:
            im = int(inter_of // 60)
            is_ = int(inter_of % 60)
            row_fi = tk.Frame(inner, bg=WHITE)
            row_fi.pack(fill="x", pady=1, padx=2)
            tk.Frame(row_fi, bg="#8b5cf6", width=4).pack(side="left", fill="y")
            name_fi = tk.Frame(row_fi, bg=WHITE)
            name_fi.pack(side="left", fill="both", expand=True, padx=(4, 0))
            tk.Label(name_fi, text="Chgmt. de série", bg=WHITE, fg=DARK,
                     font=("Arial", 11), anchor="w").pack(anchor="w")
            tk.Label(name_fi,
                     text=f"× 1  —  {im}min {is_:02d}s",
                     bg=WHITE, fg="#8b5cf6", font=("Arial", 11, "bold"),
                     anchor="w").pack(anchor="w")
        tk.Frame(inner, bg=LGRAY, height=1).pack(fill="x", pady=4)
        total = sum(v["dur"] for v in cumuls.values()) + pause_total + inter_of
        tm = int(total // 60)
        ts = int(total % 60)
        tk.Label(inner, text=f"Total : {tm}min {ts:02d}s",
                 bg=WHITE, fg=DARK, font=("Arial", 12, "bold")).pack(anchor="w", padx=6, pady=(2, 0))

    def _open_stops_editor(self):
        """Dialogue de modification des arrêts de l'OF en cours."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Modifier les arrêts")
        dlg.resizable(True, True)
        dlg.grab_set()
        self._center_on_root(dlg, 720, 500)
        dlg.configure(bg=WHITE)

        # Header
        hdr = tk.Frame(dlg, bg="#c0392b", height=40)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="✏  Modification des arrêts de l'OF en cours",
                 bg="#c0392b", fg=WHITE, font=("Arial", 11, "bold")).pack(side="left", padx=12, pady=8)

        # Colonnes
        cols_hdr = tk.Frame(dlg, bg=LGRAY)
        cols_hdr.pack(fill="x", padx=8, pady=(8, 0))
        for txt, w in [("Type d'arrêt", 200), ("Début", 140), ("Fin", 140), ("Durée", 80), ("", 80)]:
            tk.Label(cols_hdr, text=txt, bg=LGRAY, fg=DARK,
                     font=("Arial", 9, "bold"), width=w//8, anchor="w").pack(side="left", padx=4, pady=4)

        # Scrollable list
        scroll_frame = tk.Frame(dlg, bg=WHITE)
        scroll_frame.pack(fill="both", expand=True, padx=8, pady=4)
        canvas_s = tk.Canvas(scroll_frame, bg=WHITE, highlightthickness=0)
        sb = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas_s.yview)
        canvas_s.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas_s.pack(side="left", fill="both", expand=True)
        list_frame = tk.Frame(canvas_s, bg=WHITE)
        canvas_s.create_window((0, 0), window=list_frame, anchor="nw")
        list_frame.bind("<Configure>", lambda e: canvas_s.configure(scrollregion=canvas_s.bbox("all")))

        now = datetime.datetime.now()

        def _fmt_dt(dt):
            return dt.strftime("%d/%m %H:%M:%S") if dt else "—"

        def _rebuild():
            for w in list_frame.winfo_children():
                w.destroy()
            stops = [ev for ev in self._tl_events
                     if ev.get("cat") in ("ratt", "pb")
                     and (not self._of_start or ev["start"] >= self._of_start)]
            pauses_of = [(ps, pe) for ps, pe in self._pause_periods
                         if not self._of_start or ps >= self._of_start]
            if not stops and not pauses_of:
                tk.Label(list_frame, text="Aucun arrêt enregistré.", bg=WHITE, fg=GRAY,
                         font=("Arial", 10, "italic")).pack(pady=20)
            for i, ev in enumerate(stops):
                label = next((e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
                color = C_RATT if ev["cat"] == "ratt" else C_RED
                end_dt = ev.get("end") or now
                dur_s = int((end_dt - ev["start"]).total_seconds())
                dur_txt = fmt(dur_s)
                row = tk.Frame(list_frame, bg=WHITE if i % 2 == 0 else "#f9f9f9")
                row.pack(fill="x", pady=1)
                tk.Frame(row, bg=color, width=4).pack(side="left", fill="y")
                tk.Label(row, text=label, bg=row["bg"], fg=DARK,
                         font=("Arial", 10), width=25, anchor="w").pack(side="left", padx=4)
                tk.Label(row, text=_fmt_dt(ev["start"]), bg=row["bg"], fg=DARK,
                         font=("Arial", 10), width=16, anchor="w").pack(side="left", padx=2)
                tk.Label(row, text=_fmt_dt(ev.get("end")), bg=row["bg"], fg=DARK if ev.get("end") else GRAY,
                         font=("Arial", 10), width=16, anchor="w").pack(side="left", padx=2)
                tk.Label(row, text=dur_txt, bg=row["bg"], fg=color,
                         font=("Arial", 10, "bold"), width=8, anchor="w").pack(side="left", padx=2)

                def _edit(ev=ev):
                    _edit_stop_dlg(ev)
                def _delete(ev=ev):
                    if ev in self._tl_events:
                        self._tl_events.remove(ev)
                        # Recalculer elapsed du timer correspondant
                        key = ev["key"]
                        if key in self._timers and not self._timers[key]["running"]:
                            elapsed = sum(
                                (e.get("end") or now - e["start"]).total_seconds()
                                for e in self._tl_events
                                if e["key"] == key and e.get("end")
                            )
                            self._timers[key]["elapsed"] = elapsed
                    _rebuild()
                    self._refresh_stops_recap()

                tk.Button(row, text="✏", bg=LGRAY, fg=DARK, font=("Arial", 9),
                          relief="flat", cursor="hand2", padx=4,
                          command=_edit).pack(side="left", padx=2)
                tk.Button(row, text="🗑", bg="#fee2e2", fg="#dc2626", font=("Arial", 9),
                          relief="flat", cursor="hand2", padx=4,
                          command=_delete).pack(side="left", padx=2)

            for i2, (ps, pe) in enumerate(pauses_of):
                dur_s2 = int((pe - ps).total_seconds())
                row2 = tk.Frame(list_frame, bg=WHITE if (len(stops)+i2) % 2 == 0 else "#f9f9f9")
                row2.pack(fill="x", pady=1)
                tk.Frame(row2, bg="#3b82f6", width=4).pack(side="left", fill="y")
                tk.Label(row2, text="Pause pilote", bg=row2["bg"], fg=DARK,
                         font=("Arial", 10), width=25, anchor="w").pack(side="left", padx=4)
                tk.Label(row2, text=_fmt_dt(ps), bg=row2["bg"], fg=DARK,
                         font=("Arial", 10), width=16, anchor="w").pack(side="left", padx=2)
                tk.Label(row2, text=_fmt_dt(pe), bg=row2["bg"], fg=DARK,
                         font=("Arial", 10), width=16, anchor="w").pack(side="left", padx=2)
                tk.Label(row2, text=fmt(dur_s2), bg=row2["bg"], fg="#3b82f6",
                         font=("Arial", 10, "bold"), width=8, anchor="w").pack(side="left", padx=2)

        def _edit_stop_dlg(ev):
            """Mini-formulaire pour modifier heure début/fin d'un arrêt."""
            ed = tk.Toplevel(dlg)
            ed.title("Modifier l'arrêt")
            ed.resizable(False, False)
            ed.grab_set()
            self._center_on_root(ed, 380, 240)
            ed.configure(bg=WHITE)
            label = next((e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
            tk.Label(ed, text=label, bg=WHITE, fg=NAVY,
                     font=("Arial", 12, "bold")).pack(pady=(16, 8))

            def _dt_entry(parent, lbl, dt):
                f = tk.Frame(parent, bg=WHITE)
                f.pack(fill="x", padx=20, pady=4)
                tk.Label(f, text=lbl, bg=WHITE, fg=GRAY,
                         font=("Arial", 10), width=10, anchor="w").pack(side="left")
                sv = tk.StringVar(value=dt.strftime("%d/%m/%Y %H:%M:%S") if dt else "")
                tk.Entry(f, textvariable=sv, font=("Arial", 11), width=20,
                         relief="solid", bd=1).pack(side="left", padx=4)
                return sv

            sv_start = _dt_entry(ed, "Début :", ev["start"])
            sv_end   = _dt_entry(ed, "Fin :", ev.get("end"))

            def _apply():
                try:
                    new_start = datetime.datetime.strptime(sv_start.get().strip(), "%d/%m/%Y %H:%M:%S")
                    ev["start"] = new_start
                    end_s = sv_end.get().strip()
                    ev["end"] = datetime.datetime.strptime(end_s, "%d/%m/%Y %H:%M:%S") if end_s else None
                    # Recalculer elapsed
                    key = ev["key"]
                    if key in self._timers and not self._timers[key]["running"]:
                        self._timers[key]["elapsed"] = sum(
                            (e["end"] - e["start"]).total_seconds()
                            for e in self._tl_events
                            if e["key"] == key and e.get("end") and e.get("start")
                        )
                    ed.destroy()
                    _rebuild()
                    self._refresh_stops_recap()
                except ValueError:
                    _toast(ed, "Format : dd/mm/yyyy HH:MM:SS", bg=C_RED, duration=2000)

            bf = tk.Frame(ed, bg=WHITE)
            bf.pack(pady=12)
            tk.Button(bf, text="✔  Appliquer", command=_apply,
                      bg=GREEN, fg=WHITE, font=("Arial", 11, "bold"),
                      relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left", padx=4)
            tk.Button(bf, text="Annuler", command=ed.destroy,
                      bg=LGRAY, fg=DARK, font=("Arial", 10),
                      relief="flat", padx=8, pady=6).pack(side="left")

        _rebuild()

        # Footer
        foot = tk.Frame(dlg, bg=WHITE)
        foot.pack(fill="x", padx=8, pady=8)
        tk.Button(foot, text="✔  Fermer",
                  command=lambda: [dlg.destroy(), self._refresh_stops_recap()],
                  bg=NAVY, fg=WHITE, font=("Arial", 11, "bold"),
                  relief="flat", padx=14, pady=8, cursor="hand2").pack(side="right")

    def _show_nettoyage_selector(self):
        """Popup de sélection du type de nettoyage."""
        key = "nettoyage"
        if self._t_running(key):
            self._ask_stop_description(key)
            return
        clean_s = int(self.cfg.get("clean_short_min", 10))
        clean_l = int(self.cfg.get("clean_long_min", 30))
        clean_g = int(self.cfg.get("clean_grand_min", 60))

        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=WHITE)
        self._center_on_root(top, 480, 300)
        top.lift()
        top.grab_set()

        hdr = tk.Frame(top, bg="#f59e0b", height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🧹  TYPE DE NETTOYAGE",
                 bg="#f59e0b", fg=WHITE, font=("Arial", 15, "bold")).pack(
                 side="left", padx=20, pady=14)
        tk.Button(hdr, text="✕", bg="#f59e0b", fg=WHITE, font=("Arial", 12, "bold"),
                  relief="flat", cursor="hand2",
                  command=top.destroy).pack(side="right", padx=12)

        body = tk.Frame(top, bg=WHITE)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        def _choose(ntype):
            top.destroy()
            self._declare_nettoyage(ntype)

        options = [
            ("🧹 Nettoyage court (poste)", "court", f"{clean_s} min toléré", "#f59e0b"),
            ("🧽 Nettoyage long (ex: mercredi)", "long", f"{clean_l} min toléré", "#d97706"),
            ("✨ Grand nettoyage", "grand", f"{clean_g} min toléré", "#92400e"),
        ]
        for txt, ntype, sub, col in options:
            btn_f = tk.Frame(body, bg=col, cursor="hand2")
            btn_f.pack(fill="x", pady=4)
            btn_f.bind("<Button-1>", lambda e, t=ntype: _choose(t))
            inner_b = tk.Frame(btn_f, bg=col)
            inner_b.pack(fill="x", padx=12, pady=10)
            inner_b.bind("<Button-1>", lambda e, t=ntype: _choose(t))
            tk.Label(inner_b, text=txt, bg=col, fg=WHITE,
                     font=("Arial", 13, "bold"), anchor="w").pack(anchor="w")
            tk.Label(inner_b, text=sub, bg=col, fg="#fffde7",
                     font=("Arial", 10), anchor="w").pack(anchor="w")
            for w in inner_b.winfo_children():
                w.bind("<Button-1>", lambda e, t=ntype: _choose(t))

    def _declare_nettoyage(self, ntype="court"):
        """Démarre un timer nettoyage du type spécifié."""
        key = "nettoyage"
        self._t_start(key)
        ev = {
            "key": key,
            "cat": "ratt",
            "start": datetime.datetime.now(),
            "end": None,
            "nettoyage_type": ntype,
        }
        self._tl_events.append(ev)
        self._save_session()
        self._refresh_stops_recap()
        self._refresh_active_stops()

    def _start_nettoyage(self):
        """Compatibilité — délègue au sélecteur."""
        self._show_nettoyage_selector()

    # ── Pause pilote ─────────────────────────────────────────────────────────
    def _toggle_pause(self):
        if self._is_paused:
            # Fin de pause
            if self._pause_start:
                end_p = datetime.datetime.now()
                dur_p = (end_p - self._pause_start).total_seconds()
                self._pause_total_s += dur_p
                self._pause_periods.append((self._pause_start, end_p))
                # Hors production : écrire directement dans Evenements
                if not self._prod_active:
                    self._write_pause_event(self._pause_start, end_p)
                self._pause_start = None
            self._is_paused = False
            if self._pause_overlay:
                try:
                    self._pause_overlay.destroy()
                except Exception:
                    pass
                self._pause_overlay = None
            # Refresh stops recap after pause ends (refresh only, don't rebuild)
            try:
                self._refresh_stops_recap()
            except Exception:
                pass
        else:
            # Début de pause
            self._pause_start = datetime.datetime.now()
            self._is_paused = True
            self._show_pause_overlay()

    def _write_pause_event(self, start_dt, end_dt):
        """Écrit une pause pilote dans l'onglet Evenements (mode hors production)."""
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        dur = (end_dt - start_dt).total_seconds()
        pilot = self._logged_in_pilot or ""
        row_evt = [
            "Pause pilote",
            "",                                   # OF
            start_dt.strftime("%d/%m/%Y"),        # Date
            "", pilot, "", "", "", "", "", "", "", "", "", "", "",
            start_dt.strftime("%H:%M:%S"),
            end_dt.strftime("%H:%M:%S"),
            fmt(dur),
            "",
        ]
        def _bg():
            try:
                with self._excel_lock:
                    wb = self._get_wb(path)
                    if wb is None:
                        return
                    ws = self._ensure_events_sheet(wb)
                    ws.append(row_evt)
                    self._format_row(ws, ws.max_row)
                    self._safe_excel_save(wb, path)
                    self._wb_mtime_cache = os.path.getmtime(path)
                self.root.after(0, lambda: _toast(
                    self.root, f"☕  Pause enregistrée ({int(dur//60)}min)", bg=NAVY_L))
            except Exception:
                self._invalidate_wb_cache()
        threading.Thread(target=_bg, daemon=True).start()

    def _show_pause_overlay(self):
        ov = tk.Frame(self.root, bg=NAVY)
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        ov.lift()
        self._pause_overlay = ov

        tk.Label(ov, text="☕  EN PAUSE",
                 bg=NAVY, fg=WHITE, font=("Arial", 54, "bold")).pack(expand=False, pady=(80, 6))
        self._pause_timer_lbl = tk.Label(ov, text="00:00:00",
                                          bg=NAVY, fg=ORANGE, font=("Arial", 36, "bold"))
        self._pause_timer_lbl.pack(pady=(0, 40))
        tk.Button(ov, text="✔  JE SUIS REVENU", command=self._toggle_pause,
                  bg=GREEN, fg=WHITE, font=("Arial", 20, "bold"),
                  relief="flat", padx=60, pady=20, cursor="hand2").pack()
        self._update_pause_timer()

    def _update_pause_timer(self):
        if not self._is_paused or not self._pause_start:
            return
        elapsed = (datetime.datetime.now() - self._pause_start).total_seconds()
        lbl = getattr(self, "_pause_timer_lbl", None)
        if lbl:
            try:
                lbl.config(text=fmt(int(elapsed)))
            except Exception:
                return
        self.root.after(1000, self._update_pause_timer)

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
        self._center_on_root(top, 640, 360)
        top.lift()
        top.grab_set()

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
        # Terminer la pause si active
        if self._is_paused:
            self._toggle_pause()

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

        if not self._of_start:
            messagebox.showerror("Erreur", "Impossible de clôturer : heure de début inconnue.")
            return
        of_s_brut = (end_dt - self._of_start).total_seconds()
        # L'intervalle inter-OF impute toujours le TRS de l'OF suivant
        # Les pauses en excès du quota autorisé restent dans of_s (impactent TRS)
        _pause_max_s = int(self.cfg.get("pause_max_min", 20)) * 60
        _pause_excess_s = max(0.0, self._pause_total_s - _pause_max_s)
        of_s      = max(1, of_s_brut + self._inter_of_s - min(self._pause_total_s, _pause_max_s))
        stop_s    = self._t_wall_clock_stops()
        qte_fab   = _n("qte_fab")
        nb_pers   = max(1, _n("nb_pers") or 1)
        of_min    = int(of_s) / 60
        of_hrs    = int(of_s) / 3600
        equiv     = self._calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
        c1        = round(equiv / of_hrs, 2)   if of_hrs > 0 else 0
        c2        = round(equiv / (nb_pers * of_hrs), 2) if of_hrs > 0 else 0
        kit       = 2 if self._v_kit.get() else 1

        prod_ref = self._get_prod_ref()
        # ── Ajustements TRS nettoyage et réunion ─────────────────────────────
        _nett_s_tmp = sum(
            (ev["end"] - ev["start"]).total_seconds()
            for ev in self._tl_events
            if ev.get("key") == "nettoyage" and ev.get("start") and ev.get("end")
        )
        _nett_type_tmp = None
        for _ev in self._tl_events:
            if _ev.get("key") == "nettoyage" and _ev.get("start") and _ev.get("end"):
                _nett_type_tmp = _ev.get("nettoyage_type", "court")
                break
        _tol_key_map2 = {"court": "clean_short_min", "long": "clean_long_min", "grand": "clean_grand_min"}
        _tol_key2 = _tol_key_map2.get(_nett_type_tmp or "court", "clean_short_min")
        _tol_s2   = int(self.cfg.get(_tol_key2, 10)) * 60
        _nett_planned_s = min(_nett_s_tmp, _tol_s2)   # Portion planifiée exclue TRS
        _mtol_s2 = int(self.cfg.get("meeting_tol_min", 5)) * 60
        _reunion_s2 = _hms_to_sec(_min_str_to_hms(v.get("manquant_pers", "")))
        _reunion_planned_s = min(_reunion_s2, _mtol_s2)   # Portion planifiée exclue TRS
        # TRS OF : equiv / (prod_ref * of_s / 28800)
        # of_s inclut déjà la déduction des pauses planifiées — pas de déduction supplémentaire
        trs_pct  = -1.0
        if prod_ref > 0 and of_s > 0:
            expected = prod_ref * of_s / 28800.0
            trs_pct  = (equiv / expected * 100.0) if expected > 0 else -1.0

        # TRS du poste (12h) = historique Excel + déclaration actuelle
        kpi = getattr(self, "_pilot_kpi_data", {})
        hist_eq = kpi.get("last_eq", 0.0)
        hist_s  = kpi.get("last_s",  0.0)
        total_eq = hist_eq + equiv
        total_s  = hist_s  + of_s
        pilot_trs_12h = -1.0
        if prod_ref > 0 and total_s > 0:
            pilot_trs_12h = total_eq / (prod_ref * total_s / 28800.0) * 100.0

        def _ts(key):
            return fmt(self._t_get(key))

        _nett_s = sum(
            (ev["end"] - ev["start"]).total_seconds()
            for ev in self._tl_events
            if ev.get("key") == "nettoyage" and ev.get("start") and ev.get("end")
        )
        row = [
            v.get("of_num",""),                             # A
            datetime.date.today().strftime("%d/%m/%Y"),     # B
            v.get("poste",""),                              # C
            v.get("pilote",""),                             # D
            v.get("copilote",""),                           # E
            v.get("nb_pers",""),                            # F
            v.get("taille",""),                             # G
            v.get("code_prod",""),                          # H
            v.get("type_prod",""),                          # I
            v.get("poids",""),                              # J
            v.get("fibre",""),                              # K
            v.get("of_taie",""),                            # L
            v.get("traca",""),                              # M
            qte_fab,                                        # N
            _n("qte_emb"),                                  # O
            equiv,                                          # P
            fmt(of_s),                                      # Q
            self._of_start.strftime("%H:%M:%S"),            # R
            end_dt.strftime("%H:%M:%S"),                    # S
            c1, c2, kit,                                    # T, U, V
            v.get("ref_taie",""),                           # W
            _n("qte_init_taie"),                            # X
            _n("nb_taie2_choix"),                           # Y
            _n("nb_def_cout"),                              # Z
            _n("mq_taie"),                                  # AA
            _n("mq_housse_encart"),                          # AB
            _n("nb_pp_cousue"),                             # AC
            fmt(self._inter_of_s),                          # AD
            _min_str_to_hms(v.get("duree_mq_mp", "")),         # AE
            _min_str_to_hms(v.get("manquant_pers", "")),   # AF
            fmt(_nett_s),                                   # AG
            _ts("ratt_pochon"),   _ts("ratt_couture"),      # AH, AI
            _ts("ratt_emb"),      _ts("ratt_presse_soud"),  # AJ, AK
            _ts("ratt_presse_zip"),                         # AL
            _ts("pb_chargeuse"),  _ts("pb_carde"),          # AM, AN
            _ts("pb_etaleur"),    _ts("pb_coupe"),           # AO, AP
            _ts("pb_tapis1"),     _ts("pb_enrouleur"),      # AQ, AR
            _ts("pb_pesee"),      _ts("pb_deviation"),      # AS, AT
            _ts("pb_enfileur"),   _ts("pb_kinna"),           # AU, AV
            _ts("pb_tapeuse"),    _ts("pb_table_rot"),      # AW, AX
            _ts("pb_h100"),       _ts("pb_traversin"),      # AY, AZ
            _ts("pb_presse_orc"), _ts("pb_presse_zip2"),    # BA, BB
            _ts("pb_cercleuse"),  _ts("pb_enrouleuse"),     # BC, BD
            v.get("comment",""),                            # BE
            fmt(self._interposte_s) if self._interposte_s > 0 else "",  # BF
        ]

        # ── Alerte réunion non déclarée (1ère déclaration du poste) ──────────
        if self._of_count_this_shift == 0:
            mp_val = v.get("manquant_pers", "").strip()
            if mp_val in ("", "0", "00:00:00"):
                reunion_ok  = [False]
                reunion_var = tk.BooleanVar(value=False)
                ov_r = tk.Frame(self.root, bg=WHITE)
                ov_r.place(relx=0, rely=0, relwidth=1, relheight=1)
                ov_r.lift()
                tk.Frame(ov_r, bg=ORANGE, height=6).pack(fill="x")
                tk.Label(ov_r, text="⚠  Réunion non déclarée",
                         bg=WHITE, fg=ORANGE,
                         font=("Arial", 13, "bold")).pack(pady=(60, 4))
                tk.Label(ov_r,
                         text="Tu n'as pas déclaré de réunion.\nValider quand même ?",
                         bg=WHITE, fg=DARK,
                         font=("Arial", 11), justify="center").pack(pady=4)
                bf = tk.Frame(ov_r, bg=WHITE)
                bf.pack(pady=12)
                def _r_ok():
                    reunion_ok[0] = True
                    ov_r.destroy()
                    reunion_var.set(True)
                def _r_cancel():
                    ov_r.destroy()
                    reunion_var.set(True)
                tk.Button(bf, text="Valider quand même", command=_r_ok,
                          bg=GREEN, fg=WHITE, font=("Arial", 11, "bold"),
                          relief="flat", padx=14, pady=8, cursor="hand2").pack(
                          side="left", padx=6)
                tk.Button(bf, text="Annuler", command=_r_cancel,
                          bg=LGRAY, fg=DARK, font=("Arial", 11),
                          relief="flat", padx=14, pady=8, cursor="hand2").pack(
                          side="left", padx=6)
                self.root.wait_variable(reunion_var)
                if not reunion_ok[0]:
                    self._prod_active = True
                    if self._after_id: self.root.after_cancel(self._after_id)
                    self._after_id = self.root.after(1000, self._tick)
                    return

        # ── Calculs étendus pour l'affichage ──────────────────────────────────
        total_brut_s = of_s_brut
        pause_s      = self._pause_total_s
        _nett_type = None
        for _ev in self._tl_events:
            if _ev.get("key") == "nettoyage" and _ev.get("start") and _ev.get("end"):
                _nett_type = _ev.get("nettoyage_type", "court")
                break
        _tol_key_map = {"court": "clean_short_min", "long": "clean_long_min", "grand": "clean_grand_min"}
        _tol_key     = _tol_key_map.get(_nett_type or "court", "clean_short_min")
        _tol_s       = int(self.cfg.get(_tol_key, 10)) * 60
        _nett_planned = min(_nett_s, _tol_s)
        _nett_counted = max(0.0, _nett_s - _tol_s)
        _mtol_s      = int(self.cfg.get("meeting_tol_min", 5)) * 60
        _reunion_s   = _hms_to_sec(_min_str_to_hms(v.get("manquant_pers", "")))
        _reunion_planned = min(_reunion_s, _mtol_s)
        _reunion_counted = max(0.0, _reunion_s - _mtol_s)

        # ── Overlay plein écran recap avant confirmation ──────────────────────
        self._modal_open = True
        confirmed = [False]
        modified  = [False]
        recap_var = tk.BooleanVar(value=False)

        recap = tk.Frame(self.root, bg="#f0f4fb")
        recap.place(relx=0, rely=0, relwidth=1, relheight=1)
        recap.lift()

        # ── Header compact ────────────────────────────────────────────────────
        hdr_r = tk.Frame(recap, bg=NAVY, height=46)
        hdr_r.pack(fill="x")
        hdr_r.pack_propagate(False)
        tk.Label(hdr_r, text="📋  RÉCAPITULATIF DE L'OF",
                 bg=NAVY, fg=WHITE, font=("Arial", 13, "bold")).pack(
                 side="left", padx=16, pady=12)
        of_tag = v.get("of_num", "")
        if of_tag:
            tk.Label(hdr_r, text=f"  {of_tag}", bg=NAVY, fg="#7ab8f5",
                     font=("Arial", 13, "bold")).pack(side="left")

        # ── Corps : 2 colonnes ────────────────────────────────────────────────
        body_r = tk.Frame(recap, bg="#f0f4fb")
        body_r.pack(fill="both", expand=True, padx=16, pady=8)
        body_r.columnconfigure(0, weight=2)   # infos texte (40%)
        body_r.columnconfigure(1, weight=3)   # visuels (60%)
        body_r.rowconfigure(0, weight=1)

        # ── Colonne gauche : infos ─────────────────────────────────────────────
        info_card = tk.Frame(body_r, bg=WHITE,
                             highlightthickness=1, highlightbackground=LGRAY)
        info_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        tk.Frame(info_card, bg=NAVY, height=3).pack(fill="x")
        info_inner = tk.Frame(info_card, bg=WHITE)
        info_inner.pack(fill="both", expand=True, padx=16, pady=10)

        def _row(lbl, val, val_color=DARK, lbl_color=GRAY, bold=False, sep=False):
            if sep:
                tk.Frame(info_inner, bg=LGRAY, height=1).pack(fill="x", pady=(6, 4))
                return
            f = tk.Frame(info_inner, bg=WHITE)
            f.pack(fill="x", pady=2)
            tk.Label(f, text=lbl, bg=WHITE, fg=lbl_color,
                     font=("Arial", 10), width=22, anchor="w").pack(side="left")
            tk.Label(f, text=str(val), bg=WHITE, fg=val_color,
                     font=("Arial", 11, "bold" if bold else "normal")).pack(side="left")

        _row("N° OF",                 v.get("of_num", "—"),        NAVY,  bold=True)
        _row("Pilote",                self._logged_in_pilot or v.get("pilote","—"), NAVY_L)
        _row("Poste",                 v.get("poste", "—"),          DARK)
        _row(None, None, sep=True)
        _row("Qté fabriquée",         f"{qte_fab}",                 GREEN, bold=True)
        _row("Qté emballée",          f"{int(_n('qte_emb'))}",      GREEN)
        _row("Équivalence",           f"{int(round(equiv))}",       GREEN)
        _row(None, None, sep=True)
        _row("Durée brute",           fmt(total_brut_s),            DARK)
        _row("Durée comptée (TRS)",   fmt(of_s),                    NAVY_L, bold=True)
        _diff_s = of_s - total_brut_s
        if abs(_diff_s) >= 1:
            _diff_sign = "+" if _diff_s > 0 else "−"
            _row(f"Différence ({_diff_sign})", fmt(abs(_diff_s)),   C_RATT if _diff_s < 0 else GREEN, bold=True)
        if self._inter_of_s > 0:
            _row("Chgmt. de série",   fmt(self._inter_of_s),        C_RATT, bold=True)
        _row("Arrêts cumulés",        fmt(stop_s),
             C_RED if stop_s > 0 else DARK,                         bold=(stop_s > 0))
        _row("Pauses",                fmt(pause_s),                 GRAY)
        _pause_max_s_rc = int(self.cfg.get("pause_max_min", 20)) * 60
        _pause_excess_rc = max(0.0, self._pause_total_s - _pause_max_s_rc)
        _row("Pauses tolérées",       fmt(min(pause_s, _pause_max_s_rc)), GRAY)
        if _pause_excess_rc > 0:
            _row("Pauses excès TRS",  fmt(_pause_excess_rc),       C_RED, bold=True)
        _row(None, None, sep=True)
        if _nett_s > 0:
            _row("Nettoyage planifié",fmt(_nett_planned),           "#60a5fa")
            if _nett_counted > 0:
                _row("Nettoyage excès",fmt(_nett_counted),          C_RED, bold=True)
        _row("Nettoyage déclaré",     fmt(_nett_s) if _nett_s > 0 else "Aucun", GRAY)
        if _reunion_s > 0:
            _row("Réunion tolérée",   fmt(_reunion_planned),        "#fbbf24")
            if _reunion_counted > 0:
                _row("Réunion excès", fmt(_reunion_counted),        C_RED, bold=True)
        _row("Réunion déclarée",      fmt(_reunion_s) if _reunion_s > 0 else "Aucune", GRAY)
        _row("Commentaire",           v.get("comment","")[:60] or "—", GRAY)

        # ── Colonne droite : visuels ───────────────────────────────────────────
        vis_zone = tk.Frame(body_r, bg="#f0f4fb")
        vis_zone.grid(row=0, column=1, sticky="nsew")
        vis_zone.columnconfigure(0, weight=1)
        vis_zone.columnconfigure(1, weight=1)
        vis_zone.rowconfigure(0, weight=1)
        vis_zone.rowconfigure(1, weight=1)

        # Jauge 1 : TRS cet OF
        g1_card = tk.Frame(vis_zone, bg=WHITE,
                           highlightthickness=1, highlightbackground=LGRAY)
        g1_card.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        tk.Frame(g1_card, bg=GREEN, height=3).pack(fill="x")
        tk.Label(g1_card, text="TRS — cet OF", bg=WHITE, fg=GRAY,
                 font=("Arial", 11, "bold")).pack(pady=(6, 0))
        trs_disp  = max(0.0, trs_pct) if trs_pct >= 0 else 0.0
        trs_label = f"{trs_disp:.1f}%" if trs_pct >= 0 else "—"
        gauge_r   = Gauge(g1_card, bg=WHITE, width=210, height=160, highlightthickness=0)
        gauge_r.pack(fill="both", expand=True, padx=4, pady=4)
        gauge_r.update_gauge(trs_disp, trs_label)

        # Jauge 2 : TRS du poste
        g2_card = tk.Frame(vis_zone, bg=WHITE,
                           highlightthickness=1, highlightbackground=LGRAY)
        g2_card.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        tk.Frame(g2_card, bg=NAVY_L, height=3).pack(fill="x")
        _pilot_lbl = self._logged_in_pilot or v.get("pilote", "—")
        tk.Label(g2_card, text=f"TRS poste — {_pilot_lbl}",
                 bg=WHITE, fg=GRAY,
                 font=("Arial", 11, "bold"), wraplength=200).pack(pady=(6, 0))
        trs12_disp = max(0.0, pilot_trs_12h) if pilot_trs_12h >= 0 else 0.0
        trs12_lbl  = f"{trs12_disp:.1f}%" if pilot_trs_12h >= 0 else "—"
        gauge_r2   = Gauge(g2_card, bg=WHITE, width=210, height=160, highlightthickness=0)
        gauge_r2.pack(fill="both", expand=True, padx=4, pady=4)
        gauge_r2.update_gauge(trs12_disp, trs12_lbl)

        # Camembert — ligne du bas (pleine largeur)
        pie_card = tk.Frame(vis_zone, bg=WHITE,
                            highlightthickness=1, highlightbackground=LGRAY)
        pie_card.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
        tk.Frame(pie_card, bg=C_RATT, height=3).pack(fill="x")
        tk.Label(pie_card, text="Répartition du temps", bg=WHITE, fg=GRAY,
                 font=("Arial", 11, "bold")).pack(anchor="w", padx=12, pady=(4, 0))

        pie_cv = tk.Canvas(pie_card, bg=WHITE, highlightthickness=0)
        pie_cv.pack(fill="both", expand=True, padx=8, pady=4)

        def _draw_pie(e=None):
            pie_cv.delete("all")
            pw = pie_cv.winfo_width()
            ph = pie_cv.winfo_height()
            if pw < 40 or ph < 40:
                return
            r  = min(pw, ph) // 2 - 12
            cx = pw // 3          # camembert à gauche
            cy = ph // 2
            prod_eff_s = max(0.0, of_s - stop_s - _nett_counted - _reunion_counted)
            segments = [
                ("Prod. effective",     prod_eff_s,       "#1a8c4e"),
                ("Arrêts (TRS)",        stop_s,           "#e31e24"),
                ("Pauses",              pause_s,          "#94a3b8"),
                ("Nettoyage planifié",  _nett_planned,    "#60a5fa"),
                ("Nettoyage excès",     _nett_counted,    "#dc2626"),
                ("Réunion tolérée",     _reunion_planned, "#fbbf24"),
                ("Réunion excès",       _reunion_counted, "#f97316"),
                ("Changement de série", self._inter_of_s, "#a855f7"),
            ]
            total_pie = sum(s for _, s, _ in segments if s > 0)
            if total_pie <= 0:
                pie_cv.create_text(cx, cy, text="Aucune donnée",
                                   fill=GRAY, font=("Arial", 10))
                return
            start_ang = 90.0
            legend_x  = pw // 3 * 2 + 10
            legend_y  = 14
            for seg_lbl, seg_s, seg_col in segments:
                if seg_s <= 0:
                    continue
                extent = seg_s / total_pie * 360.0
                pie_cv.create_arc(cx - r, cy - r, cx + r, cy + r,
                                  start=start_ang, extent=-extent,
                                  fill=seg_col, outline=WHITE, width=1)
                # Légende à droite
                pie_cv.create_rectangle(legend_x, legend_y,
                                        legend_x + 14, legend_y + 14,
                                        fill=seg_col, outline="")
                pct_v = seg_s / total_pie * 100
                pie_cv.create_text(legend_x + 18, legend_y + 7,
                                   text=f"{seg_lbl}  {pct_v:.0f}%  ({fmt(int(seg_s))})",
                                   anchor="w", fill=DARK, font=("Arial", 9))
                legend_y += 20
                start_ang -= extent

        pie_cv.bind("<Configure>", _draw_pie)

        # ── Boutons bas ───────────────────────────────────────────────────────
        tk.Frame(recap, bg=LGRAY, height=1).pack(fill="x", padx=0)
        btn_row_r = tk.Frame(recap, bg=WHITE)
        btn_row_r.pack(fill="x", padx=20, pady=10)

        def _modifier():
            confirmed[0] = False
            modified[0]  = True
            recap.destroy()
            recap_var.set(True)

        def _confirmer():
            confirmed[0] = True
            recap.destroy()
            recap_var.set(True)

        tk.Button(btn_row_r, text="✏  MODIFIER",
                  command=_modifier, bg=LGRAY, fg=DARK,
                  font=("Arial", 12, "bold"), relief="flat",
                  padx=24, pady=10, cursor="hand2").pack(side="left", padx=(0, 10))
        tk.Button(btn_row_r, text="✔  CONFIRMER LA DÉCLARATION",
                  command=_confirmer, bg=GREEN, fg=WHITE,
                  font=("Arial", 12, "bold"), relief="flat",
                  padx=24, pady=10, cursor="hand2").pack(side="left")

        self.root.wait_variable(recap_var)
        self._modal_open = False

        if modified[0]:
            # Retour a la vue production sans rien perdre
            self._prod_active = True
            if self._after_id: self.root.after_cancel(self._after_id)
            self._after_id = self.root.after(1000, self._tick)
            return

        if not confirmed[0]:
            # Ferme sans confirmer → retour prod
            self._prod_active = True
            if self._after_id: self.root.after_cancel(self._after_id)
            self._after_id = self.root.after(1000, self._tick)
            return

        # ── Enregistrement + retour tableau de bord ───────────────────────────
        self._prod_active = False
        self._last_of_end = end_dt
        self._of_count_this_shift += 1
        self._last_of_pilot = v.get("pilote", self._logged_in_pilot or "")
        self._last_of_num   = v.get("of_num", "")
        self._interposte_s = 0
        self._inter_of_s = 0
        if self._of_periods:
            self._of_periods[-1]["end"]    = end_dt
            self._of_periods[-1]["of_num"] = v.get("of_num", "")
        self._delete_session()
        self._saved_form_data = {}
        self._reset_form_next = True

        # Afficher loading immédiatement (efface la vue production)
        self._clear()
        self._show_loading("Enregistrement de la déclaration…")

        path        = self.cfg.get("db_path", "")
        events_rows = self._build_events_rows(v)

        # Inter-poste : même OF, pilote différent → événement à écrire
        _cur_of_num    = v.get("of_num", "")
        _cur_pilot     = v.get("pilote", self._logged_in_pilot or "")
        _prev_of_num   = self._last_of_num
        _prev_pilot    = self._last_of_pilot
        _of_start_snap = self._of_start
        _last_end_snap = self._last_of_end
        _interposte_ev_row = None
        if (_cur_of_num and _cur_of_num == _prev_of_num
                and _prev_pilot and _prev_pilot != _cur_pilot
                and _of_start_snap and _last_end_snap
                and _of_start_snap > _last_end_snap):
            _gap_s = (_of_start_snap - _last_end_snap).total_seconds()
            if _gap_s > 0:
                _ip_row = ["Inter poste"] + [""] * 19
                _ip_row[2]  = datetime.date.today().strftime("%d/%m/%Y")
                _ip_row[4]  = _prev_pilot
                _ip_row[18] = fmt(int(_gap_s))
                _ip_row[19] = f"OF {_cur_of_num} — de {_prev_pilot} à {_cur_pilot}"
                _interposte_ev_row = _ip_row

        # Sauvegarde pending immédiate (backup si Excel planté)
        if path:
            payload = {"db_path": path, "row": row, "events_rows": events_rows}
            try:
                with open(PENDING_FILE, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, default=str)
            except Exception:
                pass

        def _bg_write_then_read():
            write_ok  = False
            write_err = ""
            if path:
                try:
                    with self._excel_lock:
                        wb = self._get_wb(path)
                        if wb is not None:
                            ws_d = self._ensure_data_sheet(wb)
                            ws_d.append(row)
                            self._format_row(ws_d, ws_d.max_row)
                            self._write_rows_to_events_sheet(wb, events_rows)
                            # Événement inter-poste si même OF, pilote différent
                            if _interposte_ev_row is not None:
                                self._write_rows_to_events_sheet(wb, [_interposte_ev_row])
                            self._safe_excel_save(wb, path)
                            self._wb_mtime_cache = os.path.getmtime(path)
                            try:
                                os.remove(PENDING_FILE)
                            except Exception:
                                pass
                            write_ok = True
                except PermissionError:
                    self._invalidate_wb_cache()
                    self.root.after(0, self._schedule_pending_retry)
                except Exception as e:
                    self._invalidate_wb_cache()
                    write_err = str(e)
                    self.root.after(0, self._schedule_pending_retry)
            else:
                self.root.after(0, self._schedule_pending_retry)

            # Relire Excel pour le dashboard
            self.root.after(0, lambda: self._loading_set_msg("Mise à jour du tableau…"))
            fresh_rows = []
            if path and os.path.exists(path):
                try:
                    wb2 = load_workbook(path, read_only=True, data_only=True)
                    if "Data" in wb2.sheetnames:
                        ws2 = wb2["Data"]
                        min_r = 2 if str(ws2.cell(1, 1).value or "").strip().upper() == "OF" else 1
                        for i, r in enumerate(ws2.iter_rows(min_row=min_r, values_only=True), start=min_r):
                            if any(r):
                                fresh_rows.append((i, list(r) + [None] * 60))
                    wb2.close()
                except Exception:
                    pass

            fresh_events = []
            fresh_trs    = []
            if path and os.path.exists(path):
                try:
                    wb3 = load_workbook(path, read_only=True, data_only=True)
                    if "Evenements" in wb3.sheetnames:
                        ws3 = wb3["Evenements"]
                        for r in ws3.iter_rows(min_row=2, values_only=True):
                            if r and any(r):
                                fresh_events.append(list(r))
                    if "TRS" in wb3.sheetnames:
                        ws_t3 = wb3["TRS"]
                        for r in ws_t3.iter_rows(min_row=2, values_only=True):
                            if r and any(r):
                                fresh_trs.append(list(r))
                    wb3.close()
                except Exception:
                    pass
            final = fresh_rows[-50:] if len(fresh_rows) > 50 else fresh_rows
            if write_ok:
                toast = "✔  Déclaration enregistrée dans Excel !"
            elif write_err:
                toast = f"⚠  Erreur Excel : {write_err[:60]}"
            else:
                toast = "⚠  Déclaration sauvegardée — Excel inaccessible"
            self.root.after(0, lambda: self._show_main_done(final, toast=toast, events=fresh_events, trs_data=fresh_trs))

        threading.Thread(target=_bg_write_then_read, daemon=True).start()

    def _show_fin_de_poste(self):
        """Affiche le récapitulatif complet du poste en cours."""
        if not self._logged_in_pilot:
            messagebox.showwarning("Accès refusé", "Aucun pilote connecté.\nConnectez-vous avant de déclarer une fin de poste.")
            return
        now = datetime.datetime.now()

        # Calcul durée théorique du poste depuis cfg
        poste_nom = self._logged_in_poste or ""
        if not poste_nom and hasattr(self, "fv") and self.fv:
            poste_nom = self.fv.get("poste", tk.StringVar()).get()
        if not poste_nom:
            poste_nom = self._saved_form_data.get("poste", "")

        durees_cfg = self.cfg.get("postes_durees", {})
        _override = getattr(self, "_fp_duree_override", None)
        duree_theorique_min = _override if _override else durees_cfg.get(poste_nom, 480)
        self._fp_duree_override = None

        pilot = self._logged_in_pilot or ""

        # Plage du poste : depuis la connexion jusqu'à maintenant (gère les nuits)
        login_dt  = self._login_time or (now - datetime.timedelta(hours=duree_theorique_min / 60))
        login_date_str = login_dt.strftime("%d/%m/%Y")
        today_str      = now.strftime("%d/%m/%Y")
        # Ensemble des dates à inclure (ex: veille + aujourd'hui pour les nuits)
        _shift_dates = {login_date_str, today_str}
        today = today_str  # conservé pour les écritures Excel

        total_prod_s   = 0.0
        total_panne_s  = 0.0
        total_ratt_s   = 0.0
        total_pause_s  = 0.0
        total_reunion_s = 0.0
        nb_of          = 0
        total_equiv    = 0.0
        total_qte      = 0
        total_qte_emb  = 0

        for _, row in self._data_rows_cache:
            row_date  = _row_date(row[1])
            row_pilot = str(row[3] or "")
            if row_date not in _shift_dates or row_pilot != pilot:
                continue
            try:
                total_prod_s  += _hms_to_sec(str(row[16] or ""))
                total_equiv   += float(str(row[15] or 0).replace(",", ".") or 0)
                total_qte     += int(float(str(row[13] or 0)))
                total_qte_emb += int(float(str(row[14] or 0)))
                nb_of         += 1
                panne_cols = range(38, 56)
                for ci in panne_cols:
                    if ci < len(row) and row[ci]:
                        total_panne_s += _hms_to_sec(str(row[ci]))
                ratt_cols = range(33, 38)
                for ci in ratt_cols:
                    if ci < len(row) and row[ci]:
                        total_ratt_s += _hms_to_sec(str(row[ci]))
                if len(row) > 31 and row[31]:
                    total_reunion_s += _hms_to_sec(str(row[31]))
            except Exception:
                pass

        # Ajouter prod en cours si active
        if self._prod_active and self._of_start:
            of_elapsed = (now - self._of_start).total_seconds()
            total_prod_s += of_elapsed
            nb_of += 1

        # Pauses depuis _events_cache aujourd'hui
        for ev_row in self._events_cache:
            try:
                if "pause" in str(ev_row[0] or "").lower():
                    ev_date = str(ev_row[2] or "")[:10]
                    ev_pil  = str(ev_row[4] or "")
                    if ev_date in _shift_dates and (not pilot or ev_pil == pilot):
                        total_pause_s += _hms_to_sec(str(ev_row[18] or ""))
            except Exception:
                pass
        total_pause_s += self._pause_total_s

        total_declare_s = total_prod_s + total_panne_s + total_ratt_s + total_pause_s + total_reunion_s
        duree_theorique_s = duree_theorique_min * 60
        non_declare_s = max(0.0, duree_theorique_s - total_declare_s)

        # TRS poste = equiv produite / (prod_ref × durée_théorique_poste / 28800)
        # Mesure le TRS sur l'intégralité du poste théorique configuré
        prod_ref = self._get_prod_ref()
        trs_poste = -1.0
        if prod_ref > 0 and total_equiv > 0 and duree_theorique_s > 0:
            expected2 = prod_ref * duree_theorique_s / 28800.0
            trs_poste = total_equiv / expected2 * 100.0 if expected2 > 0 else -1.0

        avg_of_poste = round(total_qte / nb_of, 1) if nb_of > 0 else 0

        # ── Sauvegarde TRS + refresh onglet Postes ──────────────────────────────
        def _fin_de_poste_save_trs():
            path = self.cfg.get("db_path", "")
            if not path or not os.path.exists(path):
                return
            _copilot = self._saved_form_data.get("copilote", "")
            _nb_pers = self._saved_form_data.get("nb_pers", "")
            _today_s = datetime.date.today().strftime("%d/%m/%Y")
            _pause_max_f = int(self.cfg.get("pause_max_min", 20)) * 60
            _meet_tol_f  = int(self.cfg.get("meeting_tol_min", 5)) * 60
            _total_decl_f = (total_prod_s + total_panne_s + total_ratt_s
                             + total_pause_s + total_reunion_s)
            _prevu_f = (min(total_pause_s, _pause_max_f)
                        + min(total_reunion_s, _meet_tol_f))
            _depasse_f = (max(0.0, total_pause_s - _pause_max_f)
                          + max(0.0, total_reunion_s - _meet_tol_f))
            def _bg_trs():
                try:
                    with self._excel_lock:
                        wb = self._get_wb(path)
                        if wb is None:
                            return
                        self._write_trs_sheet(
                            wb, _today_s, poste_nom, pilot, _copilot,
                            _nb_pers, total_prod_s, total_panne_s, total_ratt_s,
                            total_pause_s, total_reunion_s, nb_of,
                            total_qte, total_equiv, trs_poste,
                            _total_decl_f, duree_theorique_s,
                            _prevu_f, _depasse_f)
                        self._safe_excel_save(wb, path)
                        wb.close()
                    trs_fresh = []
                    try:
                        wb2 = load_workbook(path, read_only=True, data_only=True)
                        if "TRS" in wb2.sheetnames:
                            ws_t = wb2["TRS"]
                            hdrs = [c.value for c in next(ws_t.iter_rows(max_row=1))]
                            for r in ws_t.iter_rows(min_row=2, values_only=True):
                                if any(r):
                                    trs_fresh.append(list(r))
                        wb2.close()
                    except Exception:
                        pass
                    if trs_fresh:
                        self._trs_cache = trs_fresh
                        self.root.after(0, self._refresh_postes_tab)
                        try:
                            self._generate_dashboard_html()
                        except Exception:
                            pass
                except Exception:
                    pass
            threading.Thread(target=_bg_trs, daemon=True).start()

        # ── Overlay récap poste ──────────────────────────────────────────────────
        ov = tk.Frame(self.root, bg="#0d2040")
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        ov.lift()

        # Header
        hdr = tk.Frame(ov, bg=NAVY, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        _date_display = today_str if login_date_str == today_str else f"{login_date_str} → {today_str}"
        tk.Label(hdr, text="🏁  FIN DE POSTE — Récapitulatif",
                 bg=NAVY, fg=WHITE, font=("Arial", 13, "bold")).pack(side="left", padx=16, pady=12)
        tk.Label(hdr, text=f"{pilot}  ·  {poste_nom}  ·  {_date_display}",
                 bg=NAVY, fg="#93c5fd", font=("Arial", 10)).pack(side="left", padx=4)
        tk.Button(hdr, text="✕", command=ov.destroy,
                  bg=NAVY, fg=WHITE, font=("Arial", 12), relief="flat",
                  cursor="hand2", padx=8).pack(side="right", padx=10, pady=8)

        # Footer buttons
        footer = tk.Frame(ov, bg="#1e3a5f", height=52)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        def _modifier_duree():
            dlg = tk.Toplevel(self.root)
            dlg.title("Durée du poste")
            dlg.resizable(False, False)
            dlg.grab_set()
            self._center_on_root(dlg, 360, 200)
            dlg.configure(bg=WHITE)
            tk.Label(dlg, text="Modifier la durée de votre poste :",
                     bg=WHITE, fg=NAVY, font=("Arial", 12, "bold")).pack(pady=(20, 8))
            hv = tk.StringVar(value=str(duree_theorique_min // 60))
            mv = tk.StringVar(value=str(duree_theorique_min % 60))
            rf = tk.Frame(dlg, bg=WHITE)
            rf.pack(pady=4)
            tk.Entry(rf, textvariable=hv, width=4, font=("Arial", 14),
                     relief="solid", bd=1, justify="center").pack(side="left", padx=4)
            tk.Label(rf, text="h", bg=WHITE, font=("Arial", 12)).pack(side="left")
            tk.Entry(rf, textvariable=mv, width=4, font=("Arial", 14),
                     relief="solid", bd=1, justify="center").pack(side="left", padx=4)
            tk.Label(rf, text="min", bg=WHITE, font=("Arial", 12)).pack(side="left")
            def _apply_duree():
                try:
                    new_min = int(hv.get() or 0) * 60 + int(mv.get() or 0)
                    if new_min <= 0:
                        return
                    self._fp_duree_override = new_min
                    dlg.destroy()
                    ov.destroy()
                    self._show_fin_de_poste()
                except Exception:
                    pass
            btnf = tk.Frame(dlg, bg=WHITE)
            btnf.pack(pady=10)
            tk.Button(btnf, text="✔  Appliquer", command=_apply_duree,
                      bg=GREEN, fg=WHITE, font=("Arial", 11, "bold"),
                      relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=4)
            tk.Button(btnf, text="Annuler", command=dlg.destroy,
                      bg=LGRAY, fg=DARK, font=("Arial", 10),
                      relief="flat", padx=10, pady=6, cursor="hand2").pack(side="left")

        def _deconnecter_et_quitter():
            _fin_de_poste_save_trs()
            self._logged_in_pilot = None
            self._logged_in_poste = None
            self._login_time      = None
            ov.destroy()
            self._show_main()

        _dh = duree_theorique_min // 60
        _dm = duree_theorique_min % 60
        _dlbl = f"{_dh}h{_dm:02d}" if _dm else f"{_dh}h"
        tk.Button(footer,
                  text=f"⏱  Changer la durée du poste  (calculé sur {_dlbl})",
                  command=_modifier_duree, bg="#2563eb", fg=WHITE,
                  font=("Arial", 11, "bold"), relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left", padx=(12, 6), pady=8)
        tk.Button(footer, text="↩  Retour sans clôturer",
                  command=lambda: ov.destroy(),
                  bg="#475569", fg=WHITE, font=("Arial", 11), relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left", padx=6, pady=8)

        if non_declare_s > 60:
            _nd_h = int(non_declare_s) // 3600
            _nd_m = (int(non_declare_s) % 3600) // 60
            _nd_label = f"{_nd_h}h{_nd_m:02d}min" if _nd_h > 0 else f"{_nd_m}min"
            def _valider_etat():
                path = self.cfg.get("db_path", "")
                if path and os.path.exists(path):
                    try:
                        with self._excel_lock:
                            wb = self._get_wb(path)
                            if wb is not None:
                                if "Evenements" not in wb.sheetnames:
                                    wb.create_sheet("Evenements")
                                ws_e = wb["Evenements"]
                                nd_row = ["Non défini"] + [""] * 19
                                nd_row[2] = today
                                nd_row[3] = poste_nom
                                nd_row[4] = pilot
                                nd_row[18] = fmt(int(non_declare_s))
                                nd_row[19] = f"Durée théorique {fmt(duree_theorique_s)}"
                                ws_e.append(nd_row)
                                self._format_row(ws_e, ws_e.max_row)
                                self._safe_excel_save(wb, path)
                    except Exception as ex:
                        _toast(self.root, f"Erreur Excel : {ex}", bg=C_RED, duration=3000)
                _deconnecter_et_quitter()
            tk.Button(footer,
                      text=f"✔  Valider en l'état  ({_nd_label} → « Non défini »)",
                      command=_valider_etat, bg=GREEN, fg=WHITE,
                      font=("Arial", 11, "bold"), relief="flat",
                      padx=18, pady=8, cursor="hand2").pack(side="right", padx=12, pady=8)
        else:
            tk.Button(footer, text="✔  Valider et clôturer le poste",
                      command=_deconnecter_et_quitter, bg=GREEN, fg=WHITE,
                      font=("Arial", 11, "bold"), relief="flat",
                      padx=18, pady=8, cursor="hand2").pack(side="right", padx=12, pady=8)

        # ── Body — 3 colonnes : info (étroit) | TRS+KPI | Pareto — 3 lignes ────────
        body = tk.Frame(ov, bg="#0d2040")
        body.pack(fill="both", expand=True, padx=6, pady=4)

        # Colonnes : info étroit (0) | TRS+KPI (1) | Pareto (2)
        # Lignes   : résumé+pareto (0) | timeline (1) | OF table (2)
        body.columnconfigure(0, weight=1)   # info liste (étroite)
        body.columnconfigure(1, weight=2)   # TRS jauge + KPI
        body.columnconfigure(2, weight=2)   # Pareto
        body.rowconfigure(0, weight=3)      # zone haute
        body.rowconfigure(1, weight=0)      # timeline (hauteur fixe)
        body.rowconfigure(2, weight=2)      # tableau OF

        # ── ZONE 0 gauche : tableau infos ───────────────────────────────────────
        info_f = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground="#2d4a7a")
        info_f.grid(row=0, column=0, sticky="nsew", padx=(0, 3), pady=(0, 3))
        tk.Frame(info_f, bg=NAVY, height=3).pack(fill="x")
        li = tk.Frame(info_f, bg=WHITE)
        li.pack(fill="both", expand=True, padx=10, pady=6)

        def _ri(lbl, val, vc=DARK, bold=False, sep=False):
            if sep:
                tk.Frame(li, bg=LGRAY, height=1).pack(fill="x", pady=(4, 2))
                return
            f = tk.Frame(li, bg=WHITE)
            f.pack(fill="x", pady=1)
            tk.Label(f, text=lbl, bg=WHITE, fg=GRAY,
                     font=("Arial", 9), width=22, anchor="w").pack(side="left")
            tk.Label(f, text=str(val), bg=WHITE, fg=vc,
                     font=("Arial", 10, "bold" if bold else "normal")).pack(side="left")

        _ri("Date",              _date_display,              DARK)
        _ri("Pilote",            pilot or "—",               NAVY, bold=True)
        _ri("Poste",             poste_nom or "—",           NAVY_L)
        _ri("Durée théorique",   fmt(duree_theorique_s),     DARK)
        _ri(None, None, sep=True)
        _ri("OF déclarés",       nb_of,                      NAVY, bold=True)
        _ri("Qté fabriquée",     total_qte,                  GREEN, bold=True)
        _ri("Équivalence totale",f"{total_equiv:.0f}",       GREEN)
        _ri("Moy. pièces / OF",  f"{avg_of_poste:.1f}" if nb_of > 0 else "—", DARK)
        _ri(None, None, sep=True)
        _ri("Total production",  fmt(total_prod_s),          GREEN)
        _ri("Total arrêts panne",fmt(total_panne_s),
            C_RED if total_panne_s > 0 else DARK, bold=(total_panne_s > 0))
        _ri("Total rattrapages", fmt(total_ratt_s),
            C_RATT if total_ratt_s > 0 else DARK)
        _ri("Total pauses",      fmt(total_pause_s),         GRAY)
        _ri("Total réunions",    fmt(total_reunion_s),       GRAY)
        _ri("Total déclaré",     fmt(total_declare_s),       DARK, bold=True)
        nc_col = C_RED if non_declare_s > 60 else DARK
        _ri("Durée non déclarée",fmt(non_declare_s),         nc_col, bold=(non_declare_s > 60))
        _ri(None, None, sep=True)
        trs_col = GREEN if trs_poste >= 75 else (C_RATT if trs_poste >= 55 else C_RED)
        _ri("TRS du poste",
            f"{trs_poste:.1f}%" if trs_poste >= 0 else "—",
            trs_col, bold=True)

        # Alerte durée
        delta_s = total_declare_s - duree_theorique_s
        if abs(delta_s) > 300:
            if delta_s < 0:
                msg = f"⚠  DURÉE INSUFFISANTE — {fmt(abs(int(delta_s)))} de moins"
                alrt_bg, alrt_fg = "#fee2e2", "#991b1b"
            else:
                msg = f"⚠  DURÉE DÉPASSÉE — {fmt(int(delta_s))} de plus"
                alrt_bg, alrt_fg = "#fff3cd", "#92400e"
            alrt = tk.Frame(li, bg=alrt_bg, highlightthickness=2, highlightbackground=alrt_fg)
            alrt.pack(fill="x", pady=(6, 0))
            tk.Label(alrt, text=msg, bg=alrt_bg, fg=alrt_fg,
                     font=("Arial", 13, "bold"), justify="center",
                     padx=8, pady=6).pack(fill="x")

        # ── ZONE 0 milieu : jauge TRS + KPI cards ───────────────────────────────
        gauge_f = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground="#2d4a7a")
        gauge_f.grid(row=0, column=1, sticky="nsew", padx=3, pady=(0, 3))
        tk.Frame(gauge_f, bg=NAVY_L, height=3).pack(fill="x")
        tk.Label(gauge_f, text="TRS du poste", bg=WHITE, fg=GRAY,
                 font=("Arial", 10, "bold")).pack(pady=(6, 0))
        g_poste = Gauge(gauge_f, bg=WHITE, width=200, height=160, highlightthickness=0)
        g_poste.pack(pady=2)
        g_poste.update_gauge(max(0.0, trs_poste) if trs_poste >= 0 else 0.0,
                             f"{trs_poste:.1f}%" if trs_poste >= 0 else "—")
        tk.Frame(gauge_f, bg=LGRAY, height=1).pack(fill="x", padx=8, pady=4)
        kpi_data = [
            ("Quantité fabriquée",  str(total_qte),                          "#16a34a"),
            ("Quantité emballée",   str(total_qte_emb),                      "#0891b2"),
            ("Équivalence",         f"{total_equiv:.0f}",                    "#7c3aed"),
            ("Moyenne pièces / OF", f"{avg_of_poste:.1f}" if nb_of else "—", "#2563eb"),
        ]
        kg = tk.Frame(gauge_f, bg=WHITE)
        kg.pack(fill="x", padx=6, pady=4)
        for idx, (kl, kv, kc) in enumerate(kpi_data):
            kg.columnconfigure(idx % 2, weight=1)
            kf = tk.Frame(kg, bg=kc)
            kf.grid(row=idx // 2, column=idx % 2, sticky="ew", padx=3, pady=3, ipadx=4, ipady=4)
            tk.Label(kf, text=kv, bg=kc, fg=WHITE, font=("Arial", 13, "bold")).pack()
            tk.Label(kf, text=kl, bg=kc, fg="#e0f2fe", font=("Arial", 8)).pack()

        # ── Graphique barres répartition du temps (Canvas pur tkinter) ─────────
        tk.Frame(gauge_f, bg=LGRAY, height=1).pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(gauge_f, text="Répartition du temps", bg=WHITE, fg=GRAY,
                 font=("Arial", 9, "bold")).pack(anchor="w", padx=10)
        bars_data = [
            ("Production",   total_prod_s,            "#16a34a"),
            ("Arrêts panne", total_panne_s,           "#dc2626"),
            ("Rattrapages",  total_ratt_s,            "#f59e0b"),
            ("Pauses",       total_pause_s,           "#3b82f6"),
            ("Réunions",     total_reunion_s,         "#8b5cf6"),
            ("Non déclaré",  max(0.0, non_declare_s), "#94a3b8"),
        ]
        total_ref = max(duree_theorique_s, total_declare_s, 1)
        _bh = 14; _bgap = 6
        _cv_h = len(bars_data) * (_bh + _bgap) + 4
        bar_cv = tk.Canvas(gauge_f, bg=WHITE, height=_cv_h, highlightthickness=0)
        bar_cv.pack(fill="x", padx=10, pady=(2, 8))

        def _draw_bars(event=None):
            bar_cv.delete("all")
            W = bar_cv.winfo_width() or 300
            label_w = 95
            bar_w = W - label_w - 46
            for i, (lbl, sec, col) in enumerate(bars_data):
                if sec <= 0:
                    continue
                y = i * (_bh + _bgap) + 2
                pct = sec / total_ref * 100
                filled = int(bar_w * sec / total_ref)
                bar_cv.create_rectangle(label_w, y, label_w + bar_w, y + _bh,
                                        fill="#f1f5f9", outline="", width=0)
                bar_cv.create_rectangle(label_w, y, label_w + max(filled, 2), y + _bh,
                                        fill=col, outline="", width=0)
                bar_cv.create_text(label_w - 4, y + _bh // 2,
                                   text=lbl, anchor="e", font=("Arial", 8), fill="#475569")
                bar_cv.create_text(label_w + bar_w + 4, y + _bh // 2,
                                   text=f"{pct:.0f}%", anchor="w",
                                   font=("Arial", 8, "bold"), fill=col)

        bar_cv.bind("<Configure>", _draw_bars)
        bar_cv.after(50, _draw_bars)

        # ── ZONE 1 : Chronologie (pleine largeur) ────────────────────────────────
        chron_f = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground="#2d4a7a")
        chron_f.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(0, 3))
        tk.Frame(chron_f, bg=NAVY_L, height=3).pack(fill="x")
        tk.Label(chron_f, text="Chronologie du poste", bg=WHITE, fg=NAVY,
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(4, 0))

        # Canvas plus haut pour accueillir la légende intégrée et les labels OF
        chron_cv = tk.Canvas(chron_f, bg=WHITE, height=90, highlightthickness=0)
        chron_cv.pack(fill="x", padx=10, pady=(2, 0))

        _ct0 = login_dt.timestamp()
        _ct1 = now.timestamp()
        _cspan = max(_ct1 - _ct0, 1)

        _chron_legend = [
            ("Production", "#16a34a"),
            ("Panne",      "#dc2626"),
            ("Rattrapage", "#f59e0b"),
            ("Pause",      "#3b82f6"),
            ("Réunion",    "#8b5cf6"),
        ]

        def _draw_chron(event=None):
            chron_cv.delete("all")
            W = chron_cv.winfo_width() or 700
            H = chron_cv.winfo_height() or 86
            pad_l = 44; pad_r = 12
            of_lbl_h = 16   # hauteur zone labels OF (en haut)
            bar_y = of_lbl_h + 6; bar_h = 28
            tick_h = 16     # hauteur zone ticks + heures
            leg_h = 14      # hauteur légende (en bas)
            bw = W - pad_l - pad_r

            def _px(ts):
                return pad_l + (ts - _ct0) / _cspan * bw

            # Fond de la barre
            chron_cv.create_rectangle(pad_l, bar_y, pad_l + bw, bar_y + bar_h,
                                      fill="#e2e8f0", outline="", width=0)

            # Périodes OF (production)
            for p in self._of_periods:
                ps2 = p.get("start"); pe2 = p.get("end") or now
                of_label = str(p.get("of_num", "") or "")
                if ps2 and ps2 >= login_dt:
                    x0 = _px(ps2.timestamp()); x1 = _px(min(pe2, now).timestamp())
                    chron_cv.create_rectangle(x0, bar_y, max(x1, x0 + 2), bar_y + bar_h,
                                              fill="#16a34a", outline="", width=0)
                    # Numéro OF centré sur la période, tronqué si trop étroit
                    seg_w = max(x1, x0 + 2) - x0
                    if of_label and seg_w > 20:
                        short = of_label[-6:] if len(of_label) > 6 else of_label
                        chron_cv.create_text((x0 + max(x1, x0 + 2)) / 2, bar_y + bar_h / 2,
                                             text=short, anchor="center",
                                             font=("Arial", 7, "bold"), fill=WHITE)
                    # Trait de délimitation OF
                    chron_cv.create_line(x0, bar_y, x0, bar_y + bar_h,
                                        fill="white", width=1)

            # Arrêts / rattrapages
            for ev in self._tl_events:
                cat = ev.get("cat", "")
                es = ev.get("start"); ee = ev.get("end") or now
                if es and es >= login_dt and cat in ("ratt", "pb", "pause", "reunion"):
                    _cat_col = {"pb": "#dc2626", "ratt": "#f59e0b",
                                "pause": "#3b82f6", "reunion": "#8b5cf6"}.get(cat, "#94a3b8")
                    x0 = _px(es.timestamp()); x1 = _px(min(ee, now).timestamp())
                    chron_cv.create_rectangle(x0, bar_y, max(x1, x0 + 2), bar_y + bar_h,
                                              fill=_cat_col, outline="", width=0)

            # Pauses
            for ps2, pe2 in self._pause_periods:
                if ps2 >= login_dt:
                    x0 = _px(ps2.timestamp()); x1 = _px(min(pe2, now).timestamp())
                    chron_cv.create_rectangle(x0, bar_y, max(x1, x0 + 2), bar_y + bar_h,
                                              fill="#3b82f6", outline="", width=0)

            # Ticks horaires
            tick_y = bar_y + bar_h
            n_ticks = 8
            for i in range(n_ticks + 1):
                tx = pad_l + i * bw / n_ticks
                t_str = datetime.datetime.fromtimestamp(_ct0 + i * _cspan / n_ticks).strftime("%H:%M")
                chron_cv.create_line(tx, tick_y, tx, tick_y + 4, fill="#94a3b8")
                chron_cv.create_text(tx, tick_y + 10, text=t_str, anchor="center",
                                     font=("Arial", 7), fill="#64748b")

        chron_cv.bind("<Configure>", _draw_chron)
        chron_cv.after(50, _draw_chron)

        # Légende chronologie (sous le canvas)
        chron_leg_f = tk.Frame(chron_f, bg=WHITE)
        chron_leg_f.pack(anchor="w", padx=10, pady=(3, 6))
        for _cll, _clc in _chron_legend:
            _clf = tk.Frame(chron_leg_f, bg=WHITE)
            _clf.pack(side="left", padx=(0, 10))
            tk.Frame(_clf, bg=_clc, width=12, height=10).pack(side="left")
            tk.Label(_clf, text=_cll, bg=WHITE, fg="#475569",
                     font=("Arial", 8)).pack(side="left", padx=(2, 0))

        # ── ZONE 0 droite : Pareto arrêts (barres verticales) ───────────────────
        # Construire _stop_totals d'abord
        _stop_totals: dict = {}
        for _, row in self._data_rows_cache:
            rd = _row_date(row[1])
            if rd not in _shift_dates or str(row[3] or "").strip() != pilot:
                continue
            for ci in range(33, 56):
                if ci < len(row) and row[ci]:
                    _lbl_s = EVENTS[ci-33][0] if (ci-33) < len(EVENTS) else f"Col{ci}"
                    sv = _hms_to_sec(str(row[ci]))
                    if sv > 0:
                        _stop_totals[_lbl_s] = _stop_totals.get(_lbl_s, 0) + sv
            if len(row) > 32 and row[32]:
                sv = _hms_to_sec(str(row[32]))
                if sv > 0:
                    _stop_totals["Nettoyage"] = _stop_totals.get("Nettoyage", 0) + sv
        for _ev_r in self._events_cache:
            try:
                if "pause" in str(_ev_r[0] or "").lower():
                    _rd2 = str(_ev_r[2] or "")[:10]
                    _rp2 = str(_ev_r[4] or "").strip()
                    if _rd2 in _shift_dates and (not pilot or _rp2 == pilot):
                        sv = _hms_to_sec(str(_ev_r[18] or ""))
                        if sv > 0:
                            _stop_totals["Pause pilote"] = _stop_totals.get("Pause pilote", 0) + sv
            except Exception:
                pass
        if total_pause_s > 0 and "Pause pilote" not in _stop_totals:
            _stop_totals["Pause pilote"] = total_pause_s

        pareto_f = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground="#2d4a7a")
        pareto_f.grid(row=0, column=2, sticky="nsew", padx=(0, 0), pady=(0, 3))
        tk.Frame(pareto_f, bg=NAVY_L, height=3).pack(fill="x")
        tk.Label(pareto_f, text="Pareto arrêts", bg=WHITE, fg=NAVY,
                 font=("Arial", 9, "bold")).pack(anchor="w", padx=8, pady=(4, 0))
        pareto_cv = tk.Canvas(pareto_f, bg=WHITE, highlightthickness=0)
        pareto_cv.pack(fill="both", expand=True, padx=4, pady=4)

        _pareto_items = sorted(_stop_totals.items(), key=lambda x: -x[1])[:8]
        _pareto_colors = ["#dc2626","#ef4444","#f87171","#fb923c",
                          "#fbbf24","#facc15","#a3e635","#4ade80"]
        _pareto_max = max((v for _, v in _pareto_items), default=1)

        def _draw_pareto(event=None):
            pareto_cv.delete("all")
            W = pareto_cv.winfo_width() or 300
            H = pareto_cv.winfo_height() or 200
            if not _pareto_items:
                pareto_cv.create_text(W // 2, H // 2, text="Aucun arrêt",
                                      anchor="center", font=("Arial", 10), fill="#94a3b8")
                return
            n = len(_pareto_items)
            pad_t = 18; pad_b = 40; pad_l = 8; pad_r = 8
            bar_zone_h = H - pad_t - pad_b
            bar_zone_w = W - pad_l - pad_r
            bar_w = max(10, bar_zone_w / n - 4)
            for i, ((lbl, val), col) in enumerate(zip(_pareto_items, _pareto_colors)):
                x_ctr = pad_l + (i + 0.5) * bar_zone_w / n
                bar_h = int(bar_zone_h * val / _pareto_max)
                x0 = x_ctr - bar_w / 2; x1 = x_ctr + bar_w / 2
                y0 = pad_t + bar_zone_h - bar_h; y1 = pad_t + bar_zone_h
                pareto_cv.create_rectangle(x0, pad_t, x1, pad_t + bar_zone_h,
                                           fill="#f1f5f9", outline="", width=0)
                pareto_cv.create_rectangle(x0, y0, x1, y1, fill=col, outline="", width=0)
                mins = val / 60
                pareto_cv.create_text(x_ctr, y0 - 2,
                                      text=f"{mins:.0f}m", anchor="s",
                                      font=("Arial", 7, "bold"), fill=col)
                short = lbl[:10]
                pareto_cv.create_text(x_ctr, y1 + 2, text=short, anchor="n",
                                      font=("Arial", 7), fill="#374151",
                                      angle=0, width=int(bar_w + 10))

        pareto_cv.bind("<Configure>", _draw_pareto)
        pareto_cv.after(50, _draw_pareto)

        # ── ZONE 2 : Tableau des OF déclarés (pleine largeur) ────────────────────
        of_f = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground="#2d4a7a")
        of_f.grid(row=2, column=0, columnspan=3, sticky="nsew")
        tk.Frame(of_f, bg=NAVY_L, height=3).pack(fill="x")
        tk.Label(of_f, text="Détail des OF déclarés", bg=WHITE, fg=NAVY,
                 font=("Arial", 9, "bold")).pack(anchor="w", padx=10, pady=(3, 2))

        cols_of = ("N° OF","H.Début","H.Fin","Durée OF","Qté fab","Qté emb","Équiv","TRS OF","Arrêts","Ratt.")
        tv = ttk.Treeview(of_f, columns=cols_of, show="headings", height=4)
        cws = [90, 70, 70, 75, 65, 65, 65, 65, 80, 80]
        for col, w in zip(cols_of, cws):
            tv.heading(col, text=col)
            tv.column(col, width=w, anchor="center", minwidth=50, stretch=True)
        sb_of = ttk.Scrollbar(of_f, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb_of.set)
        sb_of.pack(side="right", fill="y")
        tv.pack(fill="both", expand=True)
        tv.tag_configure("trs_hi",   background="#e6f7ee", foreground=GREEN)
        tv.tag_configure("trs_warn", background="#fff7e6", foreground=ORANGE)
        tv.tag_configure("trs_low",  background="#fde8e8", foreground=C_RED)

        pr_ref = self._get_prod_ref()
        for _, row in self._data_rows_cache:
            try:
                if _row_date(row[1]) not in _shift_dates or str(row[3] or "").strip() != pilot:
                    continue
                arr_s = sum(_hms_to_sec(str(row[ci] or ""))
                            for ci in range(38, 56) if ci < len(row) and row[ci])
                rat_s = sum(_hms_to_sec(str(row[ci] or ""))
                            for ci in range(33, 38) if ci < len(row) and row[ci])
                eq = float(str(row[15] or 0).replace(",", ".") or 0)
                ps = _hms_to_sec(str(row[16] or ""))
                trs_v = (eq / (pr_ref * ps / 28800) * 100) if (pr_ref > 0 and ps > 0 and eq > 0) else -1
                trs_str = f"{trs_v:.0f}%" if trs_v >= 0 else "—"
                tag = ("trs_hi",) if trs_v >= 75 else (("trs_warn",) if trs_v >= 55 else ("trs_low",))
                tv.insert("", "end", tags=tag, values=(
                    str(row[0] or "—"),
                    str(row[17] or "—")[:8], str(row[18] or "—")[:8],
                    str(row[16] or "—")[:8],
                    str(row[13] or "—"), str(row[14] or "—"),
                    f"{eq:.0f}" if eq else "—", trs_str,
                    fmt(int(arr_s)) if arr_s else "—",
                    fmt(int(rat_s)) if rat_s else "—",
                ))
            except Exception:
                pass

    def _write_trs_sheet(self, wb, today_str, poste_nom, pilot_name, copilot_name,
                          nb_pers, total_prod_s, total_panne_s, total_ratt_s,
                          total_pause_s, total_reunion_s, nb_of, total_qte,
                          total_equiv, trs_val,
                          total_declare_s=0.0, duree_theo_s=0.0,
                          temps_prevu_s=0.0, depassement_prevu_s=0.0):
        """Écrit ou met à jour une ligne dans l'onglet TRS du workbook.
        Structure 20 colonnes A-T selon spécification v5.86."""
        trs_headers = [
            "Date",                          # A
            "Poste",                         # B
            "Pilote",                        # C
            "Co-Pilote",                     # D
            "Nb Personnes",                  # E
            "Total temps d'ouverture",       # F — durée théorique du poste
            "Total temps déclarés",          # G — prod + arrêts + pauses + réunions
            "Ecart ouverture/déclarés",      # H — F - G
            "Total temps de marche",         # I — temps de production pure
            "Total arrêts prévu",            # J — pauses + réunions dans les tolérances
            "Débordement arrêts prévu",      # K — dépassement des tolérances
            "Total arrêts (panne)",          # L
            "Total arrêt rattrapage",        # M
            "Total pauses",                  # N
            "Total réunions",                # O
            "Nombre d'OF",                   # P
            "Nombre de pièce produites",     # Q
            "Equivalence",                   # R
            "Nombre moyen de pièce par OF",  # S
            "TRS",                           # T
        ]
        if "TRS" not in wb.sheetnames:
            ws_trs = wb.create_sheet("TRS")
            ws_trs.append(trs_headers)
            self._format_row(ws_trs, 1)
        else:
            ws_trs = wb["TRS"]
            # Réécrire l'en-tête si la structure a changé
            existing_hdrs = [c.value for c in next(ws_trs.iter_rows(max_row=1))]
            if existing_hdrs != trs_headers:
                for ci, h in enumerate(trs_headers, start=1):
                    ws_trs.cell(1, ci).value = h
                self._format_row(ws_trs, 1)

        avg_of = round(total_qte / nb_of, 1) if nb_of > 0 else 0
        ecart_s = max(0.0, duree_theo_s - total_declare_s)

        existing_row_idx = None
        for i, r in enumerate(ws_trs.iter_rows(min_row=2, values_only=True), start=2):
            if (r and str(r[0] or "")[:10] == today_str
                    and str(r[2] or "") == pilot_name
                    and str(r[1] or "") == poste_nom):
                existing_row_idx = i
                break

        trs_row = [
            today_str, poste_nom, pilot_name, copilot_name, nb_pers,
            fmt(int(duree_theo_s)),          # F: Total temps d'ouverture
            fmt(int(total_declare_s)),       # G: Total temps déclarés
            fmt(int(ecart_s)),               # H: Ecart ouverture/déclarés
            fmt(int(total_prod_s)),          # I: Total temps de marche
            fmt(int(temps_prevu_s)),         # J: Total arrêts prévu
            fmt(int(depassement_prevu_s)),   # K: Débordement arrêts prévu
            fmt(int(total_panne_s)),         # L: Total arrêts (panne)
            fmt(int(total_ratt_s)),          # M: Total arrêt rattrapage
            fmt(int(total_pause_s)),         # N: Total pauses
            fmt(int(total_reunion_s)),       # O: Total réunions
            nb_of,                           # P: Nombre d'OF
            total_qte,                       # Q: Nombre de pièce produites
            round(total_equiv, 1),           # R: Equivalence
            avg_of,                          # S: Nombre moyen de pièce par OF
            f"{trs_val:.1f}%" if trs_val >= 0 else "—",  # T: TRS
        ]

        if existing_row_idx:
            for ci, val in enumerate(trs_row, start=1):
                ws_trs.cell(existing_row_idx, ci).value = val
            self._format_row(ws_trs, existing_row_idx)
        else:
            ws_trs.append(trs_row)
            self._format_row(ws_trs, ws_trs.max_row)

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
            self._invalidate_wb_cache()
            wb      = load_workbook(path, keep_links=False)
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
                with self._excel_lock:
                    self._safe_excel_save(wb, path)
            wb.close()
        except Exception:
            pass

    def _build_events_rows(self, v):
        """Retourne la liste de lignes à écrire dans Evenements (pour sérialisation)."""
        events_rows = []
        of_start = self._of_start
        try:
            kit_val = "Oui" if self._v_kit.get() else "Non"
        except Exception:
            kit_val = "Non"

        def _base_row(label, start, end):
            dur = (end - start).total_seconds()
            return [
                label,
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
                "",
            ]

        # Arrêts / rattrapages
        for ev in self._tl_events:
            if ev.get("cat") not in ("ratt", "pb"):
                continue
            if not ev.get("key") or ev["key"].startswith("_"):
                continue
            if of_start and ev["start"] < of_start:
                continue
            start    = ev["start"]
            end      = ev.get("end") or datetime.datetime.now()
            if ev["key"] == "nettoyage":
                _ntype = ev.get("nettoyage_type", "court")
                _nett_labels = {"court": "Nettoyage court", "long": "Nettoyage long", "grand": "Grand nettoyage"}
                label = _nett_labels.get(_ntype, "Nettoyage court")
                row = _base_row(label, start, end)
            else:
                cat_name = "Rattrapage" if ev["cat"] == "ratt" else "PB Technique"
                label    = next((e[0] for e in EVENTS if e[1] == ev["key"]), ev["key"])
                row = _base_row(f"{cat_name}: {label}", start, end)
            row[19] = ev.get("comment", "")
            events_rows.append(row)

        # Pauses pilote
        for ps, pe in self._pause_periods:
            if of_start and ps < of_start:
                continue
            events_rows.append(_base_row("Pause pilote", ps, pe))

        return events_rows

    def _save_pending_declaration(self, row, v):
        """Sauvegarde une déclaration temporairement quand Excel est verrouillé."""
        events_rows = self._build_events_rows(v)
        payload = {
            "db_path": self.cfg.get("db_path", ""),
            "row": row,
            "events_rows": events_rows,
        }
        try:
            with open(PENDING_FILE, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, default=str)
        except Exception:
            pass
        self._schedule_pending_retry()

    def _schedule_pending_retry(self):
        if hasattr(self, "_retry_pending_id"):
            try:
                self.root.after_cancel(self._retry_pending_id)
            except Exception:
                pass
        self._retry_pending_id = self.root.after(8000, self._check_pending)

    def _check_pending(self):
        if not os.path.exists(PENDING_FILE):
            return
        try:
            with open(PENDING_FILE, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception:
            return
        path = payload.get("db_path", "")
        if not path or not os.path.exists(path):
            self._schedule_pending_retry()
            return
        def _bg_pending():
            try:
                with self._excel_lock:
                    wb   = load_workbook(path)
                    ws_d = self._ensure_data_sheet(wb)
                    ws_d.append(payload["row"])
                    self._format_row(ws_d, ws_d.max_row)
                    ws_e = self._ensure_events_sheet(wb)
                    for ev_row in payload.get("events_rows", []):
                        ws_e.append(ev_row)
                        self._format_row(ws_e, ws_e.max_row)
                    self._safe_excel_save(wb, path)
                    wb.close()
                os.remove(PENDING_FILE)
                self.root.after(0, lambda: _toast(self.root, "✔  Déclaration enregistrée dans Excel !", bg=GREEN))
                self.root.after(0, self._reload_and_refresh)
            except PermissionError:
                self.root.after(0, self._schedule_pending_retry)
            except Exception:
                self.root.after(0, self._schedule_pending_retry)
        threading.Thread(target=_bg_pending, daemon=True).start()

    @staticmethod
    def _format_row(ws, row_idx):
        """Centre et encadre toutes les cellules d'une ligne Excel."""
        thin   = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        align  = Alignment(horizontal="center", vertical="center")
        for cell in ws[row_idx]:
            cell.alignment = align
            cell.border    = border


    # ── Dashboard HTML ────────────────────────────────────────────────────────
    def _generate_dashboard_html(self):
        """Génère KPI_Dashboard.html dans le même dossier que l'Excel."""
        import json as _json
        import datetime as _dt
        import traceback as _tb

        path = self.cfg.get("db_path", "")
        if not path:
            return
        out_dir = os.path.dirname(path) or "."
        log_path  = os.path.join(out_dir, "KPI_Dashboard_error.log")
        # Utilise la copie dashboard si disponible (lecture seule, non-bloquante)
        read_path = self._get_read_path()
        html_path = os.path.join(out_dir, "KPI_Dashboard.html")
        try:
            self._generate_dashboard_html_inner(_json, _dt, out_dir, read_path, html_path)
        except Exception:
            try:
                with open(log_path, "a", encoding="utf-8") as _lf:
                    _lf.write(f"\n[{_dt.datetime.now()}] ERREUR _generate_dashboard_html:\n")
                    _lf.write(_tb.format_exc())
            except Exception:
                pass

    def _generate_dashboard_html_inner(self, _json, _dt, out_dir, read_path, html_path):
        now_str = _dt.datetime.now().strftime("%d/%m/%Y à %H:%M:%S")

        # ── helpers ──────────────────────────────────────────────────────────
        def _s(v):
            s = str(v) if v is not None else ""
            return s if s.strip() else "—"

        def _hms_to_min(v):
            try:
                parts = str(v or "").strip().split(":")
                if len(parts) == 3:
                    return round(int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60, 1)
            except Exception:
                pass
            return 0.0

        def _trs_f(v):
            try:
                return float(str(v).replace("%", "").replace("—", "0").replace(",", ".").strip() or 0)
            except Exception:
                return 0.0

        def _esc(v):
            return str(v or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

        def _trs_color(val):
            return "#16a34a" if val >= 80 else ("#d97706" if val >= 60 else "#dc2626")

        def _badge_evt(label):
            l = str(label).lower()
            if "pb" in l or "panne" in l:
                return f'<span class="badge badge-pb">{_esc(label)}</span>'
            elif "ratt" in l:
                return f'<span class="badge badge-ratt">{_esc(label)}</span>'
            elif "nettoyage" in l:
                return f'<span class="badge badge-nett">{_esc(label)}</span>'
            elif "pause" in l:
                return f'<span class="badge badge-pause">{_esc(label)}</span>'
            elif "série" in l or "serie" in l or "chgt" in l or "changement" in l:
                return f'<span class="badge badge-chgt">{_esc(label)}</span>'
            return f'<span class="badge badge-other">{_esc(label)}</span>'

        def _js_safe(s):
            return s.replace("</", "<\\/").replace("<!--", "<\\!--")

        # ── collect last 3 postes ─────────────────────────────────────────────
        last3 = list(self._trs_cache[-3:])[::-1] if self._trs_cache else []

        postes = []
        for trs_row in last3:
            if not trs_row or len(trs_row) < 20:
                continue
            p = {
                "date":      _s(trs_row[0]),
                "poste":     _s(trs_row[1]),
                "pilote":    _s(trs_row[2]),
                "copilote":  _s(trs_row[3]),
                "nb_pers":   _s(trs_row[4]),
                "t_ouv":     _s(trs_row[5]),  "t_ouv_m":  _hms_to_min(trs_row[5]),
                "t_decl":    _s(trs_row[6]),  "t_decl_m": _hms_to_min(trs_row[6]),
                "ecart":     _s(trs_row[7]),  "ecart_m":  _hms_to_min(trs_row[7]),
                "t_marche":  _s(trs_row[8]),  "marche_m": _hms_to_min(trs_row[8]),
                "arr_prev":  _s(trs_row[9]),
                "debord":    _s(trs_row[10]),
                "pannes":    _s(trs_row[11]), "pannes_m": _hms_to_min(trs_row[11]),
                "ratt":      _s(trs_row[12]), "ratt_m":   _hms_to_min(trs_row[12]),
                "pauses":    _s(trs_row[13]), "pauses_m": _hms_to_min(trs_row[13]),
                "reunions":  _s(trs_row[14]), "reunions_m": _hms_to_min(trs_row[14]),
                "nb_of":     _s(trs_row[15]),
                "qte":       _s(trs_row[16]),
                "equiv":     _s(trs_row[17]),
                "moy_of":    _s(trs_row[18]),
                "trs_raw":   _s(trs_row[19]),
                "trs_val":   _trs_f(trs_row[19]),
            }
            d10 = _row_date(trs_row[0])
            p["ofs"] = [rd for _, rd in self._data_rows_cache
                        if _row_date(rd[1]) == d10
                        and str(rd[2] or "") == p["poste"]
                        and str(rd[3] or "") == p["pilote"]]
            p["evts"] = [r for r in self._events_cache
                         if _row_date(r[2]) == d10
                         and str(r[3] or "") == p["poste"]
                         and str(r[4] or "") == p["pilote"]]
            postes.append(p)

        # ── chart data ────────────────────────────────────────────────────────
        chart_labels = [f"{p['date'][:5]}  {p['poste']}  {p['pilote']}" for p in postes]
        chart_trs    = [round(p["trs_val"], 1) for p in postes]
        chart_colors = [_trs_color(p["trs_val"]) for p in postes]
        d_marche   = [p["marche_m"]  for p in postes]
        d_pannes   = [p["pannes_m"]  for p in postes]
        d_ratt     = [p["ratt_m"]    for p in postes]
        d_pauses   = [p["pauses_m"]  for p in postes]
        d_reunions = [p["reunions_m"] for p in postes]
        d_ecart    = [p["ecart_m"]   for p in postes]

        # pareto arrêts : sommer toutes les durées des événements
        stop_totals = {}
        for p in postes:
            for ev in p["evts"]:
                label = str(ev[0] or "").strip()
                if not label:
                    continue
                dur_m = _hms_to_min(ev[18] if len(ev) > 18 else 0)
                stop_totals[label] = stop_totals.get(label, 0.0) + dur_m
        pareto_sorted = sorted(stop_totals.items(), key=lambda x: -x[1])[:12]
        pareto_labels = [x[0] for x in pareto_sorted]
        pareto_values = [round(x[1], 1) for x in pareto_sorted]

        def _pareto_color(label):
            l = label.lower()
            if "pb" in l or "panne" in l:   return "#dc2626"
            if "ratt" in l:                  return "#d97706"
            if "nettoyage" in l:             return "#0284c7"
            if "pause" in l:                 return "#2563eb"
            if "série" in l or "serie" in l: return "#7c3aed"
            return "#64748b"

        pareto_colors = [_pareto_color(l) for l in pareto_labels]

        # Pareto par poste (1 par poste + 1 global)
        poste_paretos = []
        for p in postes:
            pt = {}
            for ev in p["evts"]:
                lbl = str(ev[0] or "").strip()
                if not lbl: continue
                dur_m = _hms_to_min(ev[18] if len(ev) > 18 else 0)
                pt[lbl] = pt.get(lbl, 0.0) + dur_m
            ps = sorted(pt.items(), key=lambda x: -x[1])[:8]
            poste_paretos.append({
                "poste": p["poste"],
                "labels": _json.dumps([x[0] for x in ps], ensure_ascii=False),
                "values": _json.dumps([round(x[1],1) for x in ps]),
                "colors": _json.dumps([_pareto_color(x[0]) for x in ps]),
            })

        # ── SECTION 1 : cartes postes ─────────────────────────────────────────
        cards_html = ""
        for i, p in enumerate(postes):
            tc   = _trs_color(p["trs_val"])
            logo = "🥇" if i == 0 else ("🥈" if i == 1 else "🥉")
            cop  = f' / {_esc(p["copilote"])}' if p["copilote"] not in ("—", "") else ""
            cards_html += f"""
    <div class="card">
      <div class="card-hdr">
        <span class="badge-poste">{logo} {_esc(p['poste'])}</span>
        <span class="badge-date">{_esc(p['date'])}</span>
        <span class="pilot">{_esc(p['pilote'])}{cop}</span>
        <span class="pers">{_esc(p['nb_pers'])} pers.</span>
      </div>
      <div class="card-body">
        <div class="gauge-col">
          <canvas id="gauge{i}" width="140" height="80"></canvas>
          <div class="trs-val" style="color:{tc}">{_esc(p['trs_raw'])}</div>
          <div class="trs-lbl">TRS</div>
        </div>
        <div class="kpi-col">
          <div class="kpi-cell"><div class="kpi-lbl">Tps ouverture</div><div class="kpi-v">{_esc(p['t_ouv'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Tps marche</div><div class="kpi-v">{_esc(p['t_marche'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Tps déclarés</div><div class="kpi-v">{_esc(p['t_decl'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Écart Tps déclaré/ouverture</div><div class="kpi-v">{_esc(p['ecart'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Pannes</div><div class="kpi-v red">{_esc(p['pannes'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Rattrapages</div><div class="kpi-v amber">{_esc(p['ratt'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Pauses</div><div class="kpi-v">{_esc(p['pauses'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Réunions</div><div class="kpi-v">{_esc(p['reunions'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Nb OF</div><div class="kpi-v blue">{_esc(p['nb_of'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Moy pièces/OF</div><div class="kpi-v">{_esc(p['moy_of'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Qte produite</div><div class="kpi-v blue">{_esc(p['qte'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Équivalence</div><div class="kpi-v green">{_esc(p['equiv'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Arrêts prévus</div><div class="kpi-v">{_esc(p['arr_prev'])}</div></div>
          <div class="kpi-cell"><div class="kpi-lbl">Débordement</div><div class="kpi-v red">{_esc(p['debord'])}</div></div>
        </div>
      </div>
    </div>"""

        # ── SECTION 3 : OF par poste ──────────────────────────────────────────
        of_sections_html = ""
        gauge_charts_js  = ""
        cadence_charts_js = ""
        for i, p in enumerate(postes):
            of_rows = ""
            cad_labels = []
            cad_vals   = []
            for rd in sorted(p["ofs"], key=lambda r: (str(r[1] or ""), str(r[17] or ""))):
                try:
                    cad_lbl = str(rd[0] or "—")[:10]
                    cad_v   = 0.0
                    try:
                        cad_v = float(str(rd[19] or "0").replace(",", ".") or 0)
                    except Exception:
                        pass
                    cad_labels.append(cad_lbl)
                    cad_vals.append(round(cad_v, 1))
                    row_json_dash = _json.dumps([str(v or "") for v in (list(rd) + [None]*60)[:60]], ensure_ascii=False).replace("'", "&#39;")
                    of_rows += f"""<tr style="cursor:pointer" onclick="openOfModal(JSON.parse(this.dataset.row))" data-row='{row_json_dash}' title="Cliquer pour voir tous les d&eacute;tails">
                      <td><b>{_esc(rd[0])}</b></td>
                      <td>{_esc(rd[17])}</td><td>{_esc(rd[18])}</td>
                      <td>{_esc(rd[6])}</td><td>{_esc(rd[7])}</td><td>{_esc(rd[8])}</td>
                      <td>{_esc(rd[13])}</td><td>{_esc(rd[14])}</td>
                      <td>{_esc(rd[15])}</td><td>{_esc(rd[16])}</td>
                      <td><b>{_esc(rd[19])}</b></td><td>{_esc(rd[20])}</td>
                      <td>{_esc(rd[29])}</td>
                      <td>{_esc(rd[30])}</td><td>{_esc(rd[31])}</td>
                      <td>{_esc(rd[32])}</td>
                      <td style="max-width:180px;white-space:normal">{_esc(rd[56])}</td>
                    </tr>"""
                except Exception:
                    continue
            if not of_rows:
                of_rows = '<tr><td colspan="17" class="empty">Aucun OF pour ce poste</td></tr>'

            tc = _trs_color(p["trs_val"])
            of_sections_html += f"""
    <div class="of-section">
      <div class="of-subtitle" style="border-color:{tc}">
        {_esc(p['poste'])} — {_esc(p['date'])} — {_esc(p['pilote'])}
        &nbsp;·&nbsp; {_esc(p['nb_of'])} OF &nbsp;·&nbsp; {_esc(p['qte'])} pièces &nbsp;·&nbsp; TRS : <b style="color:{tc}">{_esc(p['trs_raw'])}</b>
      </div>
      <div style="display:grid;grid-template-columns:1fr 280px;gap:14px;align-items:start">
        <div class="tbl-wrap">
          <table>
            <thead><tr>
              <th>OF</th><th>Début</th><th>Fin</th>
              <th>Taille</th><th>Code</th><th>Type</th>
              <th>Qte Fab</th><th>Qte Emb</th>
              <th>Equiv</th><th>Durée</th>
              <th>Cad/h</th><th>Cad/h/pers</th>
              <th>Chgt série</th>
              <th>Mq MP</th><th>Mq Pers</th>
              <th>Nettoyage</th>
              <th>Commentaire</th>
            </tr></thead>
            <tbody>{of_rows}</tbody>
          </table>
        </div>
        <div class="chart-card" style="padding:12px">
          <div class="chart-title">Cadence/h par OF</div>
          <canvas id="cad{i}" height="200"></canvas>
        </div>
      </div>
    </div>"""

            cadence_charts_js += f"""
new Chart(document.getElementById('cad{i}'), {{
  type: 'bar',
  data: {{
    labels: {_json.dumps(cad_labels)},
    datasets: [{{ label: 'Cad/h', data: {_json.dumps(cad_vals)},
      backgroundColor: '{tc}cc', borderColor: '{tc}', borderWidth: 1,
      borderRadius: 4, borderSkipped: false }}]
  }},
  options: {{
    plugins: {{ legend: {{ display: false }} }},
    scales: {{
      x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 9 }} }} }},
      y: {{ grid: {{ color: '#f1f5f9' }} }}
    }}
  }}
}});"""

            # gauge TRS mini (doughnut half)
            tv = p["trs_val"]
            gauge_charts_js += f"""
new Chart(document.getElementById('gauge{i}'), {{
  type: 'doughnut',
  data: {{
    datasets: [{{
      data: [{round(tv,1)}, {round(100-tv,1)}],
      backgroundColor: ['{tc}', '#e2e8f0'],
      borderWidth: 0,
      circumference: 180,
      rotation: 270
    }}]
  }},
  options: {{
    cutout: '70%',
    plugins: {{ legend: {{ display: false }}, tooltip: {{ enabled: false }} }},
    animation: {{ duration: 800 }}
  }}
}});"""

        # ── SECTION 4 : tous les événements ──────────────────────────────────
        all_evts = []
        for p in postes:
            for ev in p["evts"]:
                all_evts.append((p["date"], p["poste"], p["pilote"], ev))
        all_evts.sort(key=lambda x: (x[0], str(x[3][2] if len(x[3]) > 2 else ""), str(x[3][16] if len(x[3]) > 16 else "")))

        evts_html = ""
        for _pd, _pp, _pil, ev in all_evts:
            try:
                evts_html += f"""<tr>
                  <td>{_badge_evt(ev[0])}</td>
                  <td>{_esc(ev[1])}</td>
                  <td>{_esc(ev[2])}</td>
                  <td>{_esc(ev[3])}</td>
                  <td>{_esc(ev[4])}</td>
                  <td>{_esc(ev[16] if len(ev) > 16 else '')}</td>
                  <td>{_esc(ev[17] if len(ev) > 17 else '')}</td>
                  <td><b>{_esc(ev[18] if len(ev) > 18 else '')}</b></td>
                  <td style="max-width:200px;white-space:normal">{_esc(ev[19] if len(ev) > 19 else '')}</td>
                </tr>"""
            except Exception:
                continue
        if not evts_html:
            evts_html = '<tr><td colspan="9" class="empty">Aucun événement enregistré</td></tr>'

        # ── SUPERVISION : lecture session ─────────────────────────────────────
        import json as _json2
        session_data = {}
        try:
            if os.path.exists(SESSION_FILE):
                with open(SESSION_FILE, "r", encoding="utf-8") as _sf:
                    session_data = _json2.load(_sf)
        except Exception:
            pass

        sup_prod_active = bool(session_data.get("of_start"))
        sup_pilot       = session_data.get("logged_in_pilot") or "—"
        sup_form        = session_data.get("form_data", {})
        sup_poste       = str(sup_form.get("poste") or self._logged_in_poste or "—")
        sup_copilot     = str(sup_form.get("copilote") or "—")
        sup_nb_pers     = str(sup_form.get("nb_pers") or "—")
        sup_is_paused   = session_data.get("is_paused", False)
        sup_of_num      = str(sup_form.get("of_num") or "—")
        sup_taille      = str(sup_form.get("taille") or "—")
        sup_type_prod   = str(sup_form.get("type_prod") or "—")
        sup_code        = str(sup_form.get("code_prod") or "—")
        sup_fibre       = str(sup_form.get("fibre") or "—")
        sup_poids       = str(sup_form.get("poids") or "—")
        sup_of_taie     = str(sup_form.get("of_taie") or "—")
        sup_traca       = str(sup_form.get("traca") or "—")
        sup_ref_taie    = str(sup_form.get("ref_taie") or "—")
        sup_qte_fab     = str(sup_form.get("qte_fab") or "—")
        sup_qte_emb     = str(sup_form.get("qte_emb") or "—")
        sup_duree_mq_mp = str(sup_form.get("duree_mq_mp") or "—")
        sup_mq_pers     = str(sup_form.get("manquant_pers") or "—")
        sup_comment     = str(sup_form.get("_comment") or "")
        sup_kit         = bool(sup_form.get("_kit", False))
        sup_of_start_dt = None
        if session_data.get("of_start"):
            try:
                sup_of_start_dt = datetime.datetime.fromisoformat(session_data["of_start"])
            except Exception:
                pass
        sup_timers = session_data.get("timers", {})
        # Filtrer les live events par pilote connecté
        _all_sup_events = session_data.get("tl_events", [])
        if sup_pilot and sup_pilot != "—":
            sup_events = [e for e in _all_sup_events
                          if not e.get("pilot") or e.get("pilot") == sup_pilot]
        else:
            sup_events = _all_sup_events
        sup_of_count = int(session_data.get("of_count_shift", 0))

        # Index of_periods pour retrouver le numéro d'OF de chaque événement
        _of_periods_idx = session_data.get("of_periods", [])

        # Arrêts actifs
        EVENT_LABELS = {e[1]: e[0] for e in EVENTS}
        active_stops_info = []
        for k, tv in sup_timers.items():
            if tv.get("running"):
                elapsed_s = tv.get("elapsed", 0) or 0
                if tv.get("start"):
                    try:
                        s_dt = datetime.datetime.fromisoformat(tv["start"])
                        elapsed_s += (datetime.datetime.now() - s_dt).total_seconds()
                    except Exception:
                        pass
                lbl = EVENT_LABELS.get(k, k)
                cat = "pb" if k.startswith("pb_") else "ratt"
                active_stops_info.append({"label": lbl, "elapsed": elapsed_s, "cat": cat})

        has_active_stop = bool(active_stops_info)

        # Statut
        if has_active_stop:
            sup_status       = "ARRÊT EN COURS"
            sup_status_color = "#dc2626"
            sup_status_bg    = "#fee2e2"
            sup_status_icon  = "🔴"
        elif sup_is_paused:
            sup_status       = "EN PAUSE"
            sup_status_color = "#d97706"
            sup_status_bg    = "#fef3c7"
            sup_status_icon  = "⏸"
        elif sup_prod_active:
            sup_status       = "PRODUCTION EN COURS"
            sup_status_color = "#16a34a"
            sup_status_bg    = "#f0fdf4"
            sup_status_icon  = "▶"
        else:
            sup_status       = "EN ATTENTE"
            sup_status_color = "#64748b"
            sup_status_bg    = "#f1f5f9"
            sup_status_icon  = "⏺"

        # Durée de session
        sup_session_dur = "—"
        if sup_of_start_dt:
            sup_session_dur = fmt((datetime.datetime.now() - sup_of_start_dt).total_seconds())

        # Durée OF actuel
        sup_of_dur = "—"
        if sup_of_start_dt and not session_data.get("last_of_end"):
            # still in current OF
            periods = session_data.get("of_periods", [])
            if periods:
                last_p = periods[-1]
                try:
                    last_end_str = last_p.get("end")
                    last_end = datetime.datetime.fromisoformat(last_end_str) if last_end_str else None
                    if last_end is None:
                        sup_of_dur = fmt((datetime.datetime.now() - sup_of_start_dt).total_seconds())
                    else:
                        sup_of_dur = fmt((datetime.datetime.now() - last_end).total_seconds())
                except Exception:
                    pass

        # Déclarations d'aujourd'hui pour ce pilote/poste
        today_d = _dt.date.today().strftime("%d/%m/%Y")
        sup_decls = [rd for _, rd in self._data_rows_cache
                     if _row_date(rd[1]) == today_d
                     and (not sup_pilot or sup_pilot == "—"
                          or str(rd[3] or "").strip() == sup_pilot)]
        sup_decls_sorted = sorted(sup_decls,
            key=lambda r: str(r[17] or ""), reverse=True)

        # Calcul TRS live (depuis déclarations du jour)
        sup_trs_val  = 0.0
        sup_equiv_tot = 0.0
        sup_qte_tot   = 0
        prod_ref = self._get_prod_ref() if hasattr(self, "_get_prod_ref") else self._prod_ref_cached
        if sup_decls:
            for rd in sup_decls:
                try:
                    sup_equiv_tot += float(str(rd[15] or "0").replace(",", ".") or 0)
                    sup_qte_tot   += int(str(rd[13] or "0") or 0)
                except Exception:
                    pass
            if prod_ref and prod_ref > 0:
                total_decl_s = sum(_hms_to_sec(str(rd[16] or "0")) for rd in sup_decls)
                if total_decl_s > 0:
                    expected = prod_ref * total_decl_s / 28800.0
                    if expected > 0:
                        sup_trs_val = min(200, round(sup_equiv_tot / expected * 100, 1))

        # Rows HTML déclarations
        sup_decl_rows = ""
        for rd in sup_decls_sorted:
            try:
                trs_of = 0.0
                try:
                    eq  = float(str(rd[15] or "0").replace(",", ".") or 0)
                    ofs = _hms_to_sec(str(rd[16] or "0"))
                    if ofs > 0 and prod_ref and prod_ref > 0:
                        trs_of = min(200, round(eq / (prod_ref * ofs / 28800) * 100, 1))
                except Exception:
                    pass
                trs_col = "#16a34a" if trs_of >= 80 else ("#d97706" if trs_of >= 60 else "#dc2626")
                trs_cell = f'<span style="color:{trs_col};font-weight:700">{trs_of:.0f}%</span>' if trs_of > 0 else "—"
                row_json = _json.dumps([str(v or "") for v in (list(rd) + [None]*60)[:60]], ensure_ascii=False).replace("'", "&#39;")
                sup_decl_rows += f"""<tr style="cursor:pointer" onclick="openOfModal(JSON.parse(this.dataset.row))" data-row='{row_json}' title="Cliquer pour voir tous les d&eacute;tails">
                  <td><b>{_esc(rd[0])}</b></td>
                  <td>{_esc(rd[17])}</td><td>{_esc(rd[18])}</td>
                  <td>{_esc(rd[6])}</td><td>{_esc(rd[7])}</td>
                  <td>{_esc(rd[16])}</td>
                  <td>{trs_cell}</td>
                </tr>"""
            except Exception:
                continue
        if not sup_decl_rows:
            sup_decl_rows = '<tr><td colspan="7" class="empty">Aucune déclaration pour ce poste aujourd\'hui</td></tr>'

        # Rows HTML événements d'aujourd'hui
        sup_evts_today = [r for r in self._events_cache
                          if _row_date(r[2]) == today_d
                          and (not sup_pilot or sup_pilot == "—"
                               or str(r[4] or "").strip() == sup_pilot)]
        sup_evts_rows = ""
        for ev in sorted(sup_evts_today, key=lambda r: str(r[16] or ""), reverse=True):
            try:
                dur_str = str(ev[18] if len(ev) > 18 else "")
                is_open = not dur_str or dur_str.strip() in ("", "—", "00:00:00")
                row_style = ' style="background:#fee2e2"' if is_open else ""
                open_badge = ' <span style="background:#dc2626;color:white;border-radius:4px;padding:1px 6px;font-size:0.72em">EN COURS</span>' if is_open else ""
                sup_evts_rows += f"""<tr{row_style}>
                  <td>{_badge_evt(ev[0])}{open_badge}</td>
                  <td>{_esc(ev[1])}</td>
                  <td>{_esc(ev[16] if len(ev) > 16 else '')}</td>
                  <td>{_esc(ev[17] if len(ev) > 17 else '')}</td>
                  <td><b>{_esc(ev[18] if len(ev) > 18 else '')}</b></td>
                  <td style="max-width:180px;white-space:normal">{_esc(ev[19] if len(ev) > 19 else '')}</td>
                </tr>"""
            except Exception:
                continue
        if not sup_evts_rows:
            sup_evts_rows = '<tr><td colspan="6" class="empty">Aucun événement aujourd\'hui</td></tr>'

        # Événements de session en cours (tl_events, pas encore sauvés dans Excel)
        live_evts_rows = ""
        for ev in sorted(sup_events, key=lambda e: str(e.get("start") or ""), reverse=True):
            try:
                key = ev.get("key", "")
                lbl = EVENT_LABELS.get(key, key)
                start_s_str = "—"
                end_s_str   = "—"
                dur_s_str   = "—"
                try:
                    s_dt2 = datetime.datetime.fromisoformat(ev["start"])
                    start_s_str = s_dt2.strftime("%H:%M")
                    e_dt2 = datetime.datetime.fromisoformat(ev["end"]) if ev.get("end") else None
                    if e_dt2:
                        end_s_str = e_dt2.strftime("%H:%M")
                        dur_s_str = fmt((e_dt2 - s_dt2).total_seconds())
                    else:
                        end_s_str = "En cours"
                        dur_s_str = fmt((datetime.datetime.now() - s_dt2).total_seconds())
                except Exception:
                    pass
                is_open2 = not ev.get("end")
                row_style2 = ' style="background:#fee2e2"' if is_open2 else ""
                open_badge2 = ' <span style="background:#dc2626;color:white;border-radius:4px;padding:1px 6px;font-size:0.72em">EN COURS</span>' if is_open2 else ""
                # Numéro d'OF de l'événement : stocké directement, sinon retrouvé depuis of_periods
                ev_of_num = str(ev.get("of_num") or "")
                if not ev_of_num:
                    try:
                        ev_s_dt = datetime.datetime.fromisoformat(ev["start"])
                        for _op in _of_periods_idx:
                            _op_s = datetime.datetime.fromisoformat(_op["start"]) if _op.get("start") else None
                            _op_e = datetime.datetime.fromisoformat(_op["end"]) if _op.get("end") else datetime.datetime.now()
                            if _op_s and _op_s <= ev_s_dt <= _op_e:
                                ev_of_num = str(_op.get("of_num") or "")
                                break
                    except Exception:
                        pass
                ev_of_num = ev_of_num or "—"
                live_evts_rows += f"""<tr{row_style2}>
                  <td>{_badge_evt(lbl)}{open_badge2}</td>
                  <td>{_esc(ev_of_num)}</td>
                  <td>{start_s_str}</td>
                  <td>{end_s_str}</td>
                  <td><b>{dur_s_str}</b></td>
                  <td style="max-width:180px;white-space:normal">{_esc(ev.get('comment', ''))}</td>
                </tr>"""
            except Exception:
                continue
        if not live_evts_rows and not sup_prod_active:
            live_evts_rows = '<tr><td colspan="6" class="empty">Aucune session active</td></tr>'
        elif not live_evts_rows:
            live_evts_rows = '<tr><td colspan="6" class="empty">Aucun événement dans cette session</td></tr>'

        # Arrêts actifs HTML
        active_stops_html = ""
        for s in active_stops_info:
            cat_col = "#dc2626" if s["cat"] == "pb" else "#d97706"
            cat_lbl = "Panne" if s["cat"] == "pb" else "Rattrapage"
            active_stops_html += f"""
            <div class="active-stop-card" style="border-left-color:{cat_col}">
              <div class="stop-cat" style="color:{cat_col}">{cat_lbl}</div>
              <div class="stop-lbl">{_esc(s['label'])}</div>
              <div class="stop-dur" style="color:{cat_col}">{fmt(s['elapsed'])}</div>
            </div>"""

        # Alerte HTML
        alert_html = ""
        if has_active_stop:
            stops_list_txt = ", ".join(s["label"] for s in active_stops_info)
            alert_html = f"""
<div class="alarm-overlay" id="alarmBanner">
  <div class="alarm-icon">⚠</div>
  <div class="alarm-body">
    <div class="alarm-title">ARRÊT EN COURS — INTERVENTION REQUISE</div>
    <div class="alarm-detail">{_esc(stops_list_txt)}</div>
    <div class="alarm-timer" id="alarmTimer">En attente de clôture…</div>
  </div>
</div>"""

        # TRS color pour supervision
        sup_trs_color = _trs_color(sup_trs_val)
        sup_trs_display = f"{sup_trs_val:.1f}%" if sup_trs_val > 0 else "—"

        # ── Pareto arrêts supervision (live + sauvegardés) ───────────────────
        sup_pareto_dict = {}
        for ev in sup_events:
            key_p = ev.get("key", "")
            lbl_p = EVENT_LABELS.get(key_p, key_p)
            dur_p = 0.0
            try:
                s_p = datetime.datetime.fromisoformat(ev["start"])
                e_p = datetime.datetime.fromisoformat(ev["end"]) if ev.get("end") else datetime.datetime.now()
                dur_p = (e_p - s_p).total_seconds() / 60.0
            except Exception:
                pass
            sup_pareto_dict[lbl_p] = sup_pareto_dict.get(lbl_p, 0) + dur_p
        for ev in sup_evts_today:
            lbl_p = str(ev[0] or "").strip()
            if not lbl_p:
                continue
            sup_pareto_dict[lbl_p] = sup_pareto_dict.get(lbl_p, 0) + _hms_to_min(ev[18] if len(ev) > 18 else 0)
        sup_par_sorted = sorted(sup_pareto_dict.items(), key=lambda x: -x[1])[:8]
        def _par_col(l):
            ll = l.lower()
            if "pb" in ll or "panne" in ll: return "#dc2626"
            if "ratt" in ll: return "#d97706"
            if "nettoyage" in ll: return "#0284c7"
            if "pause" in ll: return "#2563eb"
            return "#7c3aed"
        sup_par_labels = _js_safe(_json.dumps([x[0] for x in sup_par_sorted], ensure_ascii=False))
        sup_par_values = _json.dumps([round(x[1], 1) for x in sup_par_sorted])
        sup_par_colors = _json.dumps([_par_col(x[0]) for x in sup_par_sorted])

        # ── Camembert production/arrêts/reste ────────────────────────────────
        shift_total_s = 8 * 3600
        elapsed_sess_s = (datetime.datetime.now() - sup_of_start_dt).total_seconds() if sup_of_start_dt else 0
        total_decl_s_pie = sum(_hms_to_sec(str(rd[16] or "0")) for rd in sup_decls)
        total_stop_s_pie = sum(v * 60 for v in sup_pareto_dict.values())
        prod_s_pie  = max(0.0, total_decl_s_pie - total_stop_s_pie)
        remaining_s_pie = max(0.0, shift_total_s - elapsed_sess_s)
        pie_values  = _json.dumps([round(prod_s_pie/60,1), round(total_stop_s_pie/60,1), round(remaining_s_pie/60,1)])

        # ── Timeline SVG 8h ─────────────────────────────────────────────────
        of_periods = session_data.get("of_periods", [])
        shift_total_s_tl = 8 * 3600
        if sup_of_start_dt:
            shift_start_tl = sup_of_start_dt
            now_s_tl = (datetime.datetime.now() - shift_start_tl).total_seconds()
            tl_segs = []
            for h in range(9):
                xp = h / 8 * 100
                hh = (shift_start_tl.hour + h) % 24
                tl_segs.append(f'<line x1="{xp:.1f}%" y1="0" x2="{xp:.1f}%" y2="54" stroke="#e2e8f0" stroke-width="1"/>')
                tl_segs.append(f'<text x="{xp:.1f}%" y="70" text-anchor="middle" fill="#94a3b8" font-size="9">{hh:02d}h</text>')
            for op in of_periods:
                try:
                    s_op = datetime.datetime.fromisoformat(op["start"])
                    e_op = datetime.datetime.fromisoformat(op["end"]) if op.get("end") else datetime.datetime.now()
                    xp2 = max(0.0, (s_op - shift_start_tl).total_seconds() / shift_total_s_tl * 100)
                    wp2 = min(100.0 - xp2, (e_op - s_op).total_seconds() / shift_total_s_tl * 100)
                    if wp2 <= 0: continue
                    col_op = "#16a34a" if op.get("end") else "#22c55e"
                    of_lbl_tl = _esc(str(op.get("of_num") or "OF")[:10])
                    tl_segs.append(f'<rect x="{xp2:.2f}%" y="4" width="{wp2:.2f}%" height="24" fill="{col_op}" rx="3" opacity="0.92"><title>{of_lbl_tl}</title></rect>')
                    if wp2 > 3:
                        tl_segs.append(f'<text x="{(xp2+wp2/2):.2f}%" y="20" text-anchor="middle" fill="white" font-size="8" font-weight="600">{of_lbl_tl[:9]}</text>')
                except Exception:
                    pass
            for ev in sup_events:
                try:
                    s_ev = datetime.datetime.fromisoformat(ev["start"])
                    e_ev = datetime.datetime.fromisoformat(ev["end"]) if ev.get("end") else datetime.datetime.now()
                    xp2 = max(0.0, (s_ev - shift_start_tl).total_seconds() / shift_total_s_tl * 100)
                    wp2 = min(100.0 - xp2, (e_ev - s_ev).total_seconds() / shift_total_s_tl * 100)
                    if wp2 <= 0: continue
                    key_tl = ev.get("key", "")
                    col_tl = "#dc2626" if key_tl.startswith("pb_") else "#d97706"
                    lbl_tl = _esc(EVENT_LABELS.get(key_tl, key_tl)[:14])
                    tl_segs.append(f'<rect x="{xp2:.2f}%" y="32" width="{wp2:.2f}%" height="14" fill="{col_tl}" rx="2" opacity="0.88"><title>{lbl_tl}</title></rect>')
                except Exception:
                    pass
            now_pct_tl = min(100.0, now_s_tl / shift_total_s_tl * 100)
            tl_segs.append(f'<line x1="{now_pct_tl:.2f}%" y1="0" x2="{now_pct_tl:.2f}%" y2="54" stroke="#2563eb" stroke-width="2" stroke-dasharray="4,2"/>')
            tl_segs.append(f'<text x="{now_pct_tl:.2f}%" y="56" text-anchor="middle" fill="#2563eb" font-size="8" font-weight="700">▲</text>')
            timeline_svg = '<svg width="100%" height="74" style="overflow:visible;display:block">' + "".join(tl_segs) + '</svg>'
        else:
            timeline_svg = '<div style="text-align:center;color:#94a3b8;padding:18px;font-style:italic">Aucune session active</div>'

        # OF en cours complets / périodes
        of_periods_html = ""
        for i, op in enumerate(reversed(of_periods[-10:]), 1):
            of_lbl = str(op.get("of_num") or f"OF #{i}")
            start_s = ""
            end_s = ""
            dur_s = "—"
            try:
                s_dt = datetime.datetime.fromisoformat(op["start"])
                start_s = s_dt.strftime("%H:%M")
                e_dt = datetime.datetime.fromisoformat(op["end"]) if op.get("end") else None
                if e_dt:
                    end_s = e_dt.strftime("%H:%M")
                    dur_s = fmt((e_dt - s_dt).total_seconds())
                else:
                    end_s = "En cours"
                    dur_s = fmt((datetime.datetime.now() - s_dt).total_seconds())
            except Exception:
                pass
            row_style = ' style="background:#f0fdf4"' if i == 1 and not op.get("end") else ""
            of_periods_html += f'<tr{row_style}><td><b>{_esc(of_lbl)}</b></td><td>{start_s}</td><td>{end_s}</td><td>{dur_s}</td></tr>'
        if not of_periods_html:
            of_periods_html = '<tr><td colspan="4" class="empty">Aucun OF cette session</td></tr>'

        # ── REVUE COMPLÈTE : préparer les données JSON pour le JS ─────────────
        import time as _time
        rev_all_data = self._review_full_data or []
        rev_all_evts = self._review_full_evts or []
        rev_all_trs  = self._review_full_trs or self._trs_cache or []

        # Si pas encore de données complètes, utiliser le cache courant
        if not rev_all_data:
            rev_all_data = [(i, rd) for i, rd in self._data_rows_cache]
        if not rev_all_evts:
            rev_all_evts = self._events_cache

        def _to_str_row(r):
            row = r[1] if isinstance(r, (list, tuple)) and len(r) == 2 and isinstance(r[0], int) else r
            result = list((list(row) + [None]*60)[:60])
            if len(result) > 1:
                result[1] = _row_date(result[1])  # Data col B = Date
            return [str(v or "") for v in result]

        def _to_str_evt(r):
            row = list(r)
            result = list((row + [""]*22)[:22])
            if len(result) > 2:
                result[2] = _row_date(result[2])  # Events col C = Date
            return [str(v or "") for v in result]

        def _to_str_trs(r):
            row = list(r)
            result = list((row + [""]*22)[:22])
            if len(result) > 0:
                result[0] = _row_date(result[0])  # TRS col A = Date
            return [str(v or "") for v in result]

        rev_data_js  = _js_safe(_json.dumps([_to_str_row(r) for r in rev_all_data], ensure_ascii=False))
        rev_evts_js  = _js_safe(_json.dumps([_to_str_evt(r) for r in rev_all_evts], ensure_ascii=False))
        rev_trs_js   = _js_safe(_json.dumps([_to_str_trs(r) for r in rev_all_trs], ensure_ascii=False))

        # Listes déroulantes filtre
        all_postes_rev  = sorted(set(str(r[1][2] if isinstance(r,(list,tuple)) and len(r)==2 else r[2] or "") for r in rev_all_data if r) - {""})
        all_pilotes_rev = sorted(set(str(r[1][3] if isinstance(r,(list,tuple)) and len(r)==2 else r[3] or "") for r in rev_all_data if r) - {""})
        rev_poste_opts  = "\n".join(f'<option value="{_esc(p)}">{_esc(p)}</option>' for p in all_postes_rev)
        rev_pilote_opts = "\n".join(f'<option value="{_esc(p)}">{_esc(p)}</option>' for p in all_pilotes_rev)

        rev_last_update = _dt.datetime.fromtimestamp(self._review_cache_ts).strftime("%d/%m/%Y %H:%M") if self._review_cache_ts else "—"

        # ── Pareto canvases HTML pour postes ────────────────────────────────────
        poste_pareto_canvases_html = ""
        for i, pp in enumerate(poste_paretos):
            poste_pareto_canvases_html += f"""    <div class="chart-card">
      <div class="chart-title">Par&eacute;to {_esc(pp['poste'])} (min)</div>
      <canvas id="chartParetoPoste{i}" height="200"></canvas>
    </div>\n"""

        # ── Pareto all postes + JS pareto postes ─────────────────────────────────
        pareto_all_totals = {}
        for p in postes:
            for ev in p["evts"]:
                lbl = str(ev[0] or "").strip()
                if not lbl: continue
                dur_m = _hms_to_min(ev[18] if len(ev) > 18 else 0)
                pareto_all_totals[lbl] = pareto_all_totals.get(lbl, 0.0) + dur_m
        pareto_all_sorted = sorted(pareto_all_totals.items(), key=lambda x: -x[1])[:12]
        pareto_all_labels = _js_safe(_json.dumps([x[0] for x in pareto_all_sorted], ensure_ascii=False))
        pareto_all_values = _json.dumps([round(x[1],1) for x in pareto_all_sorted])
        pareto_all_colors = _json.dumps([_pareto_color(x[0]) for x in pareto_all_sorted])

        pareto_poste_js = ""
        for i, pp in enumerate(poste_paretos):
            pareto_poste_js += f"""
new Chart(document.getElementById('chartParetoPoste{i}'), {{
  type: 'bar',
  data: {{
    labels: {pp['labels']},
    datasets: [{{ label: 'Minutes', data: {pp['values']}, backgroundColor: {pp['colors']}, borderRadius: 4, borderSkipped: false }}]
  }},
  options: {{
    plugins: {{ legend:{{ display:false }}, title:{{ display:false }} }},
    scales: {{
      x: {{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }}, maxRotation:40 }} }},
      y: {{ grid:{{ color:'#f1f5f9' }}, ticks:{{ callback:function(v){{ return v+'m'; }} }} }}
    }}
  }}
}});
"""

        # ─── Build new supervision HTML ───────────────────────────────────────
        # CSS animation class for blinking header on arrêt
        hdr_anim = 'animation:supBlink 1s step-start infinite;' if has_active_stop else ''
        hdr_bg   = sup_status_color if has_active_stop else sup_status_bg
        hdr_fg   = 'white' if has_active_stop else sup_status_color

        # Compact product info cells (hauteur réduite)
        def _inf(label, val, col='#1e3a5f'):
            empty = not val or val == '—'
            bg = '#f8fafc' if not empty else '#f1f5f9'
            vc = col if not empty else '#cbd5e1'
            return (f'<div style="background:{bg};border:1px solid #e2e8f0;border-radius:4px;padding:2px 5px">'
                    f'<div style="font-size:0.50em;color:#94a3b8;text-transform:uppercase;line-height:1.1">{label}</div>'
                    f'<div style="font-weight:700;color:{vc};font-size:0.78em;line-height:1.2">{_esc(val or "—")}</div></div>')

        prod_info_html = (
            _inf("Fibre", sup_fibre) +
            _inf("Poids", f"{sup_poids} gr" if sup_poids != '—' else '—') +
            _inf("OF Taie", sup_of_taie) +
            _inf("Réf. Taie", sup_ref_taie) +
            _inf("Traça fibre", sup_traca) +
            _inf("Kit 2 pièces", "✓ Oui" if sup_kit else "Non") +
            _inf("MQ MP (min)", sup_duree_mq_mp, '#dc2626') +
            _inf("MQ Pers (min)", sup_mq_pers, '#dc2626')
        )
        if sup_comment:
            prod_info_html += (f'<div style="grid-column:1/-1;background:#f0f9ff;border:1px solid #bae6fd;border-radius:6px;padding:5px 8px">'
                               f'<div style="font-size:0.58em;color:#94a3b8;text-transform:uppercase">Commentaire</div>'
                               f'<div style="font-size:0.82em;color:#1e3a5f;white-space:pre-wrap">{_esc(sup_comment)}</div></div>')

        stops_hdr_style = 'background:#dc2626;animation:supBlink 1s step-start infinite;' if has_active_stop else ''
        stops_body = (f'<div class="active-stops-grid">{active_stops_html}</div>' if active_stops_info
                      else '<div style="padding:10px;text-align:center;color:#16a34a;font-weight:700;font-size:0.9em">✓ Aucun arrêt en cours</div>')

        # ── HTML complet ──────────────────────────────────────────────────────
        html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<!-- auto-refresh géré en JS uniquement sur onglet supervision -->
<title>KPI-ORC — Dashboard & Supervision</title>
<script>
// Canvas 2D charts - aucun CDN nécessaire
function _drawDonut(canvas,values,colors){{
  var ctx=canvas.getContext('2d'),w=canvas.width,h=canvas.height,cx=w/2,cy=h/2;
  var r=Math.min(w,h)/2-3,inner=r*0.62;
  ctx.clearRect(0,0,w,h);
  var total=values.reduce(function(a,b){{return a+b;}},0);
  if(total<=0){{ctx.beginPath();ctx.arc(cx,cy,r,0,Math.PI*2);ctx.arc(cx,cy,inner,0,Math.PI*2,true);ctx.fillStyle='#e2e8f0';ctx.fill();return;}}
  var ang=-Math.PI/2;
  for(var i=0;i<values.length;i++){{
    var sl=values[i]/total*Math.PI*2;
    ctx.beginPath();ctx.moveTo(cx,cy);ctx.arc(cx,cy,r,ang,ang+sl);ctx.closePath();ctx.fillStyle=colors[i]||'#ccc';ctx.fill();
    ang+=sl;
  }}
  ctx.beginPath();ctx.arc(cx,cy,inner,0,Math.PI*2);ctx.fillStyle='white';ctx.fill();
}}
function _drawHBar(canvas,labels,values,colors){{
  var ctx=canvas.getContext('2d');
  var w=canvas.offsetWidth||canvas.width||300,h=canvas.offsetHeight||canvas.height||180;
  canvas.width=w;canvas.height=h;
  ctx.clearRect(0,0,w,h);
  if(!values||!values.length){{ctx.fillStyle='#94a3b8';ctx.font='12px sans-serif';ctx.textAlign='center';ctx.fillText('Aucun \u00e9v\u00e9nement',w/2,h/2);return;}}
  var maxV=Math.max.apply(null,values)||1,n=values.length;
  var barH=Math.min(26,(h-16)/n-4),lw=Math.min(100,w*0.38),bw=w-lw-36;
  ctx.font='10px Segoe UI,sans-serif';
  for(var i=0;i<n;i++){{
    var y=8+i*(barH+5);
    ctx.fillStyle='#475569';ctx.textAlign='right';
    ctx.fillText((labels[i]||'').substring(0,16),lw-4,y+barH-3);
    var fw=Math.max(2,values[i]/maxV*bw);
    ctx.fillStyle=colors[i]||'#7c3aed';ctx.beginPath();
    if(ctx.roundRect)ctx.roundRect(lw,y,fw,barH,3);else ctx.rect(lw,y,fw,barH);
    ctx.fill();
    ctx.fillStyle='#1e3a5f';ctx.textAlign='left';
    ctx.fillText(values[i].toFixed(1)+'m',lw+fw+4,y+barH-3);
  }}
}}
</script>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a2332; font-size: 14px; }}

#page-header {{ position: sticky; top: 0; z-index: 200; }}
.hdr {{ background: linear-gradient(135deg, #1e3a5f 0%, #2c5282 100%); color: white;
        padding: 14px 28px; display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 3px 16px rgba(0,0,0,0.3); }}
.hdr h1 {{ font-size: 1.35em; font-weight: 800; }}
.hdr .meta {{ font-size: 0.78em; color: rgba(255,255,255,0.75); margin-top: 2px; }}
.countdown {{ font-size: 0.75em; color: #fbbf24; margin-top: 4px; }}

/* ── Onglets ── */
.tab-bar {{ background: #1e3a5f; display: flex; gap: 0; border-bottom: 3px solid #fbbf24; }}
.tab-btn {{ padding: 12px 32px; border: none; background: transparent; color: rgba(255,255,255,0.6);
            font-size: 0.9em; font-weight: 700; cursor: pointer; letter-spacing: 0.5px;
            text-transform: uppercase; transition: all 0.2s; border-bottom: 3px solid transparent;
            margin-bottom: -3px; }}
.tab-btn:hover {{ color: white; background: rgba(255,255,255,0.08); }}
.tab-btn.active {{ color: #fbbf24; border-bottom-color: #fbbf24; background: rgba(251,191,36,0.1); }}
.tab-content {{ display: none; }}
.tab-content.visible {{ display: block; }}

/* ── Supervision ── */
.sup-status-bar {{ padding: 18px 28px; display: flex; align-items: center; gap: 16px;
                   border-bottom: 3px solid; transition: all 0.5s; }}
.sup-status-icon {{ font-size: 2.2em; }}
.sup-status-text {{ font-size: 1.6em; font-weight: 900; letter-spacing: 1px; }}
.sup-status-sub  {{ font-size: 0.82em; opacity: 0.75; margin-top: 3px; }}

.alarm-overlay {{ background: linear-gradient(135deg, #7f1d1d, #dc2626);
                  color: white; padding: 20px 28px; display: flex; align-items: center; gap: 20px;
                  animation: alarmPulse 1s infinite; border-bottom: 4px solid #fbbf24; }}
.alarm-icon {{ font-size: 3em; animation: alarmShake 0.5s infinite; }}
.alarm-title {{ font-size: 1.4em; font-weight: 900; letter-spacing: 1px; }}
.alarm-detail {{ font-size: 1em; opacity: 0.9; margin-top: 4px; }}
.alarm-timer {{ font-size: 0.82em; color: #fbbf24; margin-top: 6px; font-weight: 600; }}
@keyframes alarmPulse {{ 0%,100% {{ opacity:1; }} 50% {{ opacity:0.82; }} }}
@keyframes alarmShake {{
  0%,100% {{ transform: rotate(0deg); }}
  20%     {{ transform: rotate(-8deg); }}
  60%     {{ transform: rotate(8deg); }}
}}

.sup-grid {{ display: grid; grid-template-columns: 340px 1fr; gap: 20px;
             padding: 20px 28px; }}
.sup-left {{ display: flex; flex-direction: column; gap: 16px; }}
.sup-right {{ display: flex; flex-direction: column; gap: 16px; }}

.sup-card {{ background: white; border-radius: 14px; box-shadow: 0 2px 12px rgba(0,0,0,0.08);
             border: 1px solid #e2e8f0; overflow: hidden; }}
.sup-card-hdr {{ background: #1e3a5f; color: white; padding: 10px 16px;
                 font-size: 0.8em; font-weight: 700; text-transform: uppercase;
                 letter-spacing: 0.6px; display: flex; align-items: center; gap: 8px; }}
.sup-card-body {{ padding: 16px; }}

.sup-trs-big {{ font-size: 3.2em; font-weight: 900; text-align: center; padding: 12px 0 4px; }}
.sup-trs-label {{ text-align: center; font-size: 0.72em; color: #94a3b8;
                  text-transform: uppercase; letter-spacing: 1px; padding-bottom: 12px; }}
.sup-kpi-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1px;
                 background: #f1f5f9; border-top: 1px solid #f1f5f9; }}
.sup-kpi-cell {{ background: white; padding: 10px 12px; }}
.sup-kpi-lbl  {{ font-size: 0.63em; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.4px; }}
.sup-kpi-val  {{ font-size: 1.0em; font-weight: 700; color: #1e3a5f; margin-top: 3px; }}

.active-stop-card {{ background: white; border-left: 5px solid; border-radius: 8px;
                     padding: 12px 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                     display: flex; flex-direction: column; gap: 4px; }}
.stop-cat {{ font-size: 0.65em; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; }}
.stop-lbl {{ font-size: 0.95em; font-weight: 700; color: #1e3a5f; }}
.stop-dur {{ font-size: 1.3em; font-weight: 900; font-variant-numeric: tabular-nums; }}
.active-stops-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(180px,1fr));
                      gap: 10px; padding: 14px 16px; }}

/* ── Revue complète ── */
.rev-filters {{ background: #1e3a5f; padding: 14px 28px; display: flex; align-items: center;
                gap: 16px; flex-wrap: wrap; }}
.rev-filter-group {{ display: flex; align-items: center; gap: 6px; }}
.rev-filter-group label {{ color: rgba(255,255,255,0.75); font-size: 0.8em; font-weight: 600;
                           text-transform: uppercase; white-space: nowrap; }}
.rev-filter-group input, .rev-filter-group select {{
  padding: 5px 10px; border: none; border-radius: 6px; font-size: 0.85em;
  background: rgba(255,255,255,0.12); color: white; outline: none;
  border: 1px solid rgba(255,255,255,0.2); }}
.rev-filter-group input[type=date] {{ color-scheme: dark; }}
.rev-filter-group select option {{ background: #1e3a5f; }}
.rev-btn {{ padding: 6px 18px; background: #fbbf24; color: #1e3a5f; border: none;
            border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 0.85em;
            text-transform: uppercase; transition: background 0.15s; }}
.rev-btn:hover {{ background: #f59e0b; }}
.rev-last-update {{ font-size: 0.72em; color: rgba(255,255,255,0.5); margin-left: auto; }}

.rev-kpi-row {{ display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px;
                padding: 16px 28px; background: white; border-bottom: 1px solid #e2e8f0; }}
.rev-kpi-card {{ background: #f8fafc; border-radius: 10px; padding: 12px 14px;
                 border: 1px solid #e2e8f0; text-align: center; }}
.rev-kpi-num  {{ font-size: 1.6em; font-weight: 800; color: #1e3a5f; }}
.rev-kpi-lbl  {{ font-size: 0.67em; color: #94a3b8; text-transform: uppercase;
                 letter-spacing: 0.4px; margin-top: 2px; }}

.rev-body {{ padding: 20px 28px; }}
.rev-charts-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 18px;
                    margin-bottom: 24px; }}
.rev-charts-grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 18px;
                      margin-bottom: 24px; }}
.rev-chart-card {{ background: white; border-radius: 12px; padding: 18px;
                   box-shadow: 0 2px 10px rgba(0,0,0,0.07); border: 1px solid #e2e8f0; }}
.rev-chart-title {{ font-size: 0.8em; font-weight: 700; color: #1e3a5f; margin-bottom: 14px;
                    text-transform: uppercase; letter-spacing: 0.5px; display: flex;
                    justify-content: space-between; align-items: center; }}
.rev-tbl-section {{ background: white; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.07);
                    border: 1px solid #e2e8f0; margin-bottom: 20px; overflow: hidden; }}
.rev-tbl-hdr {{ background: #1e3a5f; color: white; padding: 12px 18px;
                font-size: 0.85em; font-weight: 700; text-transform: uppercase;
                letter-spacing: 0.5px; display: flex; align-items: center; justify-content: space-between; }}
.rev-tbl-wrap {{ overflow-x: auto; max-height: 420px; overflow-y: auto; }}
.rev-count {{ background: rgba(255,255,255,0.2); border-radius: 10px; padding: 2px 8px;
              font-size: 0.75em; }}

.section {{ padding: 20px 28px; }}
.section + .section {{ border-top: 2px solid #e2e8f0; }}
.sec-title {{ font-size: 1.05em; font-weight: 700; color: #1e3a5f; margin-bottom: 16px;
              padding-left: 12px; border-left: 4px solid #2563eb; display: flex; align-items: center; gap: 8px; }}

/* Cards */
.cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 20px; }}
.card {{ background: white; border-radius: 14px; box-shadow: 0 3px 16px rgba(0,0,0,0.09);
         overflow: hidden; border: 1px solid #e2e8f0; }}
.card-hdr {{ background: linear-gradient(90deg, #1e3a5f, #274e7f); color: white;
             padding: 10px 14px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
.badge-poste {{ background: #16a34a; border-radius: 6px; padding: 3px 10px;
                font-weight: 700; font-size: 0.88em; white-space: nowrap; }}
.badge-date {{ background: rgba(255,255,255,0.15); border-radius: 6px; padding: 3px 9px;
               font-size: 0.82em; color: #fbbf24; font-weight: 600; }}
.pilot {{ font-size: 0.86em; flex: 1; }}
.pers  {{ font-size: 0.78em; color: rgba(255,255,255,0.65); }}

.card-body {{ display: grid; grid-template-columns: 160px 1fr; }}
.gauge-col {{ padding: 16px 8px 12px; display: flex; flex-direction: column;
              align-items: center; justify-content: center; border-right: 1px solid #f0f4f8;
              background: #fafbfc; }}
.trs-val {{ font-size: 1.5em; font-weight: 800; margin-top: 6px; }}
.trs-lbl {{ font-size: 0.68em; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px; }}

.kpi-col {{ display: grid; grid-template-columns: 1fr 1fr; }}
.kpi-cell {{ padding: 7px 11px; border-right: 1px solid #f1f5f9; border-bottom: 1px solid #f1f5f9; }}
.kpi-cell:nth-child(even) {{ border-right: none; }}
.kpi-lbl {{ font-size: 0.65em; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.4px; }}
.kpi-v   {{ font-size: 0.88em; font-weight: 700; color: #1e3a5f; margin-top: 2px; }}
.kpi-v.red {{ color: #dc2626; }}
.kpi-v.amber {{ color: #d97706; }}
.kpi-v.green {{ color: #16a34a; }}
.kpi-v.blue {{ color: #2563eb; }}

/* Charts */
.charts-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 18px; }}
.chart-card {{ background: white; border-radius: 12px; padding: 18px;
               box-shadow: 0 2px 10px rgba(0,0,0,0.08); border: 1px solid #e2e8f0; }}
.chart-title {{ font-size: 0.8em; font-weight: 700; color: #1e3a5f; margin-bottom: 14px;
                text-transform: uppercase; letter-spacing: 0.5px; }}

/* OF sections */
.of-section {{ margin-bottom: 28px; }}
.of-subtitle {{ font-size: 0.9em; font-weight: 700; color: #1e3a5f; margin-bottom: 12px;
                padding: 8px 14px; background: #eff6ff; border-radius: 8px;
                border-left: 4px solid #2563eb; }}
.tbl-wrap {{ overflow-x: auto; border-radius: 10px; border: 1px solid #e2e8f0; }}
table {{ width: 100%; border-collapse: collapse; font-size: 0.78em; }}
thead tr {{ background: #1e3a5f; }}
thead th {{ padding: 9px 10px; color: white; text-align: left; font-weight: 600;
            white-space: nowrap; border-right: 1px solid rgba(255,255,255,0.12); }}
tbody tr:nth-child(even) {{ background: #f8fafc; }}
tbody tr:hover {{ background: #eff6ff; transition: background 0.15s; }}
tbody td {{ padding: 7px 10px; border-bottom: 1px solid #f1f5f9; white-space: nowrap;
            border-right: 1px solid #f1f5f9; }}

/* Badges événements */
.badge {{ display: inline-block; border-radius: 5px; padding: 2px 8px;
          font-size: 0.76em; font-weight: 600; white-space: nowrap; }}
.badge-pb    {{ background: #fee2e2; color: #dc2626; }}
.badge-ratt  {{ background: #fef3c7; color: #d97706; }}
.badge-nett  {{ background: #e0f2fe; color: #0284c7; }}
.badge-pause {{ background: #f0fdf4; color: #16a34a; }}
.badge-chgt  {{ background: #f5f3ff; color: #7c3aed; }}
.badge-other {{ background: #f1f5f9; color: #475569; }}

.empty {{ text-align: center; padding: 20px; color: #94a3b8; font-style: italic; }}
.footer {{ text-align: center; padding: 14px; font-size: 0.72em; color: #94a3b8;
           border-top: 1px solid #e2e8f0; background: white; margin-top: 8px; }}
</style>
</head>
<body>

<div id="page-header"><div class="hdr">
  <div>
    <h1>&#128202; KPI-ORC</h1>
    <div class="meta">{now_str}</div>
  </div>
  <div style="text-align:right">
    <div class="meta">Auto-refresh 15s &nbsp; <button onclick="location.reload()" style="background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.4);color:white;border-radius:4px;padding:2px 8px;cursor:pointer;font-size:0.85em">&#8635; Actualiser</button></div>
    <div class="countdown" id="cdown">&#8635; Mise &agrave; jour dans 15s</div>
  </div>
</div>

<div class="tab-bar">
  <button class="tab-btn" id="btn-supervision" onclick="showTab('supervision')">&#9881; Supervision</button>
  <button class="tab-btn" id="btn-dashboard" onclick="showTab('dashboard')">&#128202; R&eacute;cap 3 derniers postes</button>
  <button class="tab-btn" id="btn-review" onclick="showTab('review')">&#128218; Historique complet BDD</button>
</div>
</div><!-- /page-header -->

<!-- ══════════════════ ONGLET DASHBOARD ══════════════════ -->
<div class="tab-content" id="tab-dashboard">

<!-- ═══════════════════════ SECTION 1 : POSTES ═══════════════════════ -->
<div class="section">
  <div class="sec-title">&#127942; R&eacute;sum&eacute; des 3 derniers postes</div>
  <div class="cards">
{cards_html}
  </div>
</div>

<!-- ═══════════════════════ SECTION 2 : GRAPHIQUES ═══════════════════════ -->
<div class="section">
  <div class="sec-title">&#128200; Analyse comparative</div>
  <div class="charts-grid">
    <div class="chart-card">
      <div class="chart-title">TRS par poste (%)</div>
      <canvas id="chartTRS" height="200"></canvas>
    </div>
    <div class="chart-card">
      <div class="chart-title">R&eacute;partition du temps (min)</div>
      <canvas id="chartTemps" height="200"></canvas>
    </div>
    <div class="chart-card">
      <div class="chart-title">Par&eacute;to global arr&ecirc;ts (min)</div>
      <canvas id="chartPareto" height="200"></canvas>
    </div>
  </div>
  <!-- Paretos par poste -->
  <div class="charts-grid" style="margin-top:10px">
{poste_pareto_canvases_html}
    <div class="chart-card">
      <div class="chart-title">Par&eacute;to tous postes confondus (min)</div>
      <canvas id="chartParetoAll" height="200"></canvas>
    </div>
  </div>
</div>

<!-- ═══════════════════════ SECTION 3 : OF PAR POSTE ═══════════════════════ -->
<div class="section">
  <div class="sec-title">&#128203; D&eacute;tail des OF par poste</div>
{of_sections_html}
</div>

<!-- ═══════════════════════ SECTION 4 : ÉVÉNEMENTS ═══════════════════════ -->
<div class="section">
  <div class="sec-title">&#9888; Tous les &eacute;v&eacute;nements et arr&ecirc;ts</div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>Type</th><th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th>
        <th>D&eacute;but</th><th>Fin</th><th>Dur&eacute;e</th><th>Commentaire</th>
      </tr></thead>
      <tbody>{evts_html}</tbody>
    </table>
  </div>
</div>

<div class="footer">
  KPI-ORC &bull; G&eacute;n&eacute;r&eacute; le {now_str} &bull; Actualisation automatique toutes les 15 secondes
</div>

</div><!-- /tab-dashboard -->

<!-- ══════════════════ ONGLET SUPERVISION ══════════════════ -->
<div class="tab-content visible" id="tab-supervision">

<style>
@keyframes supBlink {{
  0%,49%{{ background:#dc2626; }}
  50%,100%{{ background:#7f1d1d; }}
}}
@keyframes supPageBlink {{
  0%,49%{{ background:#fca5a5; border:3px solid #dc2626; }}
  50%,100%{{ background:#fee2e2; border:3px solid #ef4444; }}
}}
</style>

<div id="supPageWrap" style="{'animation:supPageBlink 0.7s step-start infinite;border-radius:8px;padding:2px;margin:2px;' if has_active_stop else ''}">

<!-- HEADER 1 LIGNE -->
<div id="supHdr" style="background:{hdr_bg};border:2px solid {sup_status_color};border-radius:10px;
  margin:8px 12px 6px;padding:7px 16px;display:flex;align-items:center;gap:14px;
  color:{hdr_fg};{hdr_anim}flex-wrap:nowrap;overflow:hidden;min-height:50px">
  <span style="font-size:1.8em;flex-shrink:0">{sup_status_icon}</span>
  <span style="font-size:1.0em;text-transform:uppercase;font-weight:800;letter-spacing:1px;white-space:nowrap;flex-shrink:0">{_esc(sup_status)}</span>
  <span style="width:2px;height:28px;background:currentColor;opacity:0.3;flex-shrink:0"></span>
  <span style="font-size:2.0em;font-weight:900;white-space:nowrap;flex-shrink:0">{_esc(sup_of_num)}</span>
  <span style="width:2px;height:28px;background:currentColor;opacity:0.2;flex-shrink:0"></span>
  <div style="font-size:1.05em;display:flex;gap:16px;flex-wrap:nowrap;overflow:hidden;white-space:nowrap">
    <span>&#128203; <b>{_esc(sup_poste)}</b></span>
    <span>&#128100; <b>{_esc(sup_pilot)}</b>{'&nbsp;/&nbsp;' + _esc(sup_copilot) if sup_copilot not in ('—','') else ''}</span>
    <span>Taille: <b>{_esc(sup_taille)}</b></span>
    <span><b>{_esc(sup_type_prod)}</b></span>
    <span style="color:{'white' if has_active_stop else '#16a34a'}">Code: <b>{_esc(sup_code)}</b></span>
    {f'<span>{_esc(sup_nb_pers)} pers.</span>' if sup_nb_pers not in ('—','') else ''}
  </div>
  <span style="width:2px;height:28px;background:currentColor;opacity:0.2;flex-shrink:0"></span>
  <span style="font-size:1.0em;white-space:nowrap;flex-shrink:0">Durée OF: <b>{sup_of_dur}</b></span>
  <span style="font-size:1.0em;white-space:nowrap;flex-shrink:0">TRS: <b style="font-size:1.25em;color:{'white' if has_active_stop else sup_trs_color}">{_esc(sup_trs_display)}</b></span>
  <span style="margin-left:auto;font-size:0.75em;opacity:0.7;white-space:nowrap;flex-shrink:0">{now_str}</span>
</div>

<!-- TIMELINE PLEINE LARGEUR -->
<div class="sup-card" style="margin:0 12px 6px;padding:5px 12px">
  <div style="font-size:0.65em;font-weight:700;color:#1e3a5f;text-transform:uppercase;margin-bottom:3px">
    &#9654; Timeline 8h &nbsp;<span style="font-weight:400;color:#94a3b8">&#9632; OF &nbsp;<span style="color:#dc2626">&#9632;</span> Pannes &nbsp;<span style="color:#d97706">&#9632;</span> Ratt &nbsp;<span style="color:#2563eb">│</span> Maintenant</span>
  </div>
  {timeline_svg}
</div>

<!-- BODY 3 COLONNES -->
<div style="display:grid;grid-template-columns:1fr 1.2fr 1fr;gap:8px;padding:0 12px 8px;height:calc(100vh - 195px)">

  <!-- COL GAUCHE : Détails produit + Arrêts -->
  <div style="display:flex;flex-direction:column;gap:8px;min-height:0">

    <div class="sup-card" style="flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column">
      <div class="sup-card-hdr">&#128203; D&eacute;tails produit</div>
      <div style="padding:6px;display:grid;grid-template-columns:1fr 1fr;gap:4px;overflow-y:auto;flex:1">
        {prod_info_html}
      </div>
    </div>

    <div class="sup-card" style="flex-shrink:0;{'border:2px solid #dc2626;' if has_active_stop else ''}">
      <div class="sup-card-hdr" style="{stops_hdr_style}">
        &#9888; Arr&ecirc;ts actifs — {'<b>' + str(len(active_stops_info)) + ' en cours</b>' if active_stops_info else 'Aucun'}
      </div>
      {stops_body}
    </div>

  </div>

  <!-- COL CENTRE : TRS + KPIs + Pie + Évts session -->
  <div style="display:flex;flex-direction:column;gap:8px;min-height:0">

    <div class="sup-card" style="flex-shrink:0">
      <div class="sup-card-hdr">&#128200; TRS Actuel &amp; KPIs</div>
      <div style="padding:6px 8px;display:grid;grid-template-columns:auto 1fr;gap:8px;align-items:center">
        <div style="text-align:center">
          <div style="font-size:2.8em;font-weight:900;color:{sup_trs_color};line-height:1">{_esc(sup_trs_display)}</div>
          <div style="font-size:0.65em;color:#94a3b8">bas&eacute; OF d&eacute;clar&eacute;s</div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:0.8em">
          <div class="sup-kpi-cell"><div class="sup-kpi-lbl">Equiv. totale</div><div class="sup-kpi-val">{sup_equiv_tot:.1f}</div></div>
          <div class="sup-kpi-cell"><div class="sup-kpi-lbl">Qte produite</div><div class="sup-kpi-val">{sup_qte_tot}</div></div>
          <div class="sup-kpi-cell"><div class="sup-kpi-lbl">OF compl&eacute;t&eacute;s</div><div class="sup-kpi-val">{sup_of_count}</div></div>
          <div class="sup-kpi-cell"><div class="sup-kpi-lbl">Dur&eacute;e session</div><div class="sup-kpi-val">{sup_session_dur}</div></div>
        </div>
      </div>
    </div>

    <div class="sup-card" style="flex-shrink:0">
      <div class="sup-card-hdr">&#9685; R&eacute;partition du poste (min)</div>
      <div style="padding:8px;display:flex;align-items:center;gap:10px">
        <canvas id="supPieChart" width="120" height="120" style="flex-shrink:0;width:120px;height:120px"></canvas>
        <div style="font-size:0.74em;display:flex;flex-direction:column;gap:4px">
          <div><span style="display:inline-block;width:10px;height:10px;background:#16a34a;border-radius:2px;margin-right:5px"></span>Production d&eacute;clar&eacute;e</div>
          <div><span style="display:inline-block;width:10px;height:10px;background:#dc2626;border-radius:2px;margin-right:5px"></span>Arr&ecirc;ts</div>
          <div><span style="display:inline-block;width:10px;height:10px;background:#e2e8f0;border-radius:2px;margin-right:5px"></span>Reste du poste</div>
        </div>
      </div>
    </div>

    <div class="sup-card" style="flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column">
      <div class="sup-card-hdr" style="background:#7c3aed">&#128308; &Eacute;v&eacute;nements session ({len(sup_events)})</div>
      <div style="overflow-y:auto;flex:1">
        <table style="font-size:0.82em;width:100%">
          <thead><tr><th>Type</th><th>OF</th><th>D&eacute;but</th><th>Fin</th><th>Dur&eacute;e</th><th>Commentaire</th></tr></thead>
          <tbody>{live_evts_rows}</tbody>
        </table>
      </div>
    </div>

  </div>

  <!-- COL DROITE : Pareto + Déclarations -->
  <div style="display:flex;flex-direction:column;gap:8px;min-height:0">

    <div class="sup-card" style="flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column">
      <div class="sup-card-hdr">&#128200; Par&eacute;to arr&ecirc;ts &amp; &eacute;v&eacute;nements (min)</div>
      <div style="padding:8px;flex:1;min-height:0">
        <canvas id="supParetoChart" style="max-height:100%"></canvas>
      </div>
    </div>

    <div class="sup-card" style="flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column">
      <div class="sup-card-hdr">&#128221; D&eacute;clarations ({len(sup_decls)} OF) <span style="font-size:0.72em;opacity:0.7">— cliquer d&eacute;tails</span></div>
      <div style="overflow-y:auto;flex:1">
        <table id="supDeclTable" style="font-size:0.8em;width:100%">
          <thead><tr>
            <th>OF</th><th>D&eacute;but</th><th>Fin</th>
            <th>Taille</th><th>Code</th><th>Dur&eacute;e</th><th>TRS</th>
          </tr></thead>
          <tbody id="supDeclBody">{sup_decl_rows}</tbody>
        </table>
      </div>
    </div>

  </div>

</div><!-- /body-3col -->

<!-- Modal détail OF (shared with dashboard) -->
<div id="ofDetailModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.55);z-index:9999;align-items:center;justify-content:center">
  <div style="background:white;border-radius:16px;max-width:900px;width:95%;max-height:90vh;overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,0.3);display:flex;flex-direction:column">
    <div style="background:linear-gradient(135deg,#1e3a5f,#2c5282);color:white;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;gap:12px">
      <div style="font-size:1.1em;font-weight:800;flex:1">&#128203; D&eacute;tail de l'OF <span id="modalOfNum"></span></div>
      <button id="modalEditBtn" onclick="startEditModal()" style="background:#d97706;border:none;color:white;border-radius:8px;padding:6px 14px;cursor:pointer;font-weight:700;font-size:0.95em">&#9998; Modifier</button>
      <button id="modalSaveBtn" onclick="saveEditModal()" style="display:none;background:#16a34a;border:none;color:white;border-radius:8px;padding:6px 14px;cursor:pointer;font-weight:700;font-size:0.95em">&#10003; Sauvegarder</button>
      <button id="modalCancelBtn" onclick="cancelEditModal()" style="display:none;background:#6b7280;border:none;color:white;border-radius:8px;padding:6px 14px;cursor:pointer;font-weight:700;font-size:0.95em">&#10005; Annuler</button>
      <span id="modalEditMsg" style="font-size:0.8em;opacity:0.8"></span>
      <button onclick="closeOfModal()" style="background:rgba(255,255,255,0.15);border:none;color:white;border-radius:8px;padding:6px 12px;cursor:pointer;font-weight:700;font-size:1.1em">&#10005;</button>
    </div>
    <div id="modalBody" style="overflow-y:auto;padding:20px;display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px"></div>
    <div id="modalEditBody" style="display:none;overflow-y:auto;padding:20px"></div>
  </div>
</div>

<div class="footer">
  KPI-ORC Supervision &bull; {now_str} &bull; Rafra&icirc;chissement automatique 15s
</div>

</div><!-- /supPageWrap -->
</div><!-- /tab-supervision -->

<!-- ══════════════════ ONGLET REVUE COMPLÈTE ══════════════════ -->
<div class="tab-content" id="tab-review">

<div class="rev-filters">
  <div class="rev-filter-group">
    <label>Du</label>
    <input type="date" id="dtStart">
  </div>
  <div class="rev-filter-group">
    <label>Au</label>
    <input type="date" id="dtEnd">
  </div>
  <div class="rev-filter-group">
    <label>Poste</label>
    <select id="fPoste"><option value="">Tous les postes</option>{rev_poste_opts}</select>
  </div>
  <div class="rev-filter-group">
    <label>Pilote</label>
    <select id="fPilote"><option value="">Tous les pilotes</option>{rev_pilote_opts}</select>
  </div>
  <button class="rev-btn" onclick="applyFilters()">&#128260; Actualiser</button>
  <div class="rev-last-update">Données chargées le {rev_last_update}</div>
</div>

<div class="rev-kpi-row">
  <div class="rev-kpi-card"><div class="rev-kpi-num" id="kpiNbPostes">—</div><div class="rev-kpi-lbl">Postes</div></div>
  <div class="rev-kpi-card"><div class="rev-kpi-num" id="kpiNbOF">—</div><div class="rev-kpi-lbl">OF déclarés</div></div>
  <div class="rev-kpi-card"><div class="rev-kpi-num" id="kpiTrsAvg" style="color:#1e3a5f">—</div><div class="rev-kpi-lbl">TRS moyen</div></div>
  <div class="rev-kpi-card"><div class="rev-kpi-num" id="kpiQteTot">—</div><div class="rev-kpi-lbl">Qté produite</div></div>
  <div class="rev-kpi-card"><div class="rev-kpi-num" id="kpiEquivTot">—</div><div class="rev-kpi-lbl">Équivalence</div></div>
  <div class="rev-kpi-card"><div class="rev-kpi-num" id="kpiNbArrets">—</div><div class="rev-kpi-lbl">Arrêts déclarés</div></div>
</div>

<div class="rev-body">

  <!-- Ligne 1 : TRS + Répartition temps -->
  <div class="rev-charts-grid">
    <div class="rev-chart-card">
      <div class="rev-chart-title">TRS par poste (%) <span id="lblTRS"></span></div>
      <canvas id="revChartTRS" height="220"></canvas>
    </div>
    <div class="rev-chart-card">
      <div class="rev-chart-title">Répartition du temps (min) <span id="lblTemps"></span></div>
      <canvas id="revChartTemps" height="220"></canvas>
    </div>
    <div class="rev-chart-card">
      <div class="rev-chart-title">Pareto arrêts par type (min)</div>
      <canvas id="revChartPareto" height="220"></canvas>
    </div>
  </div>

  <!-- Ligne 2 : Cadence + Poste + Pilote -->
  <div class="rev-charts-grid">
    <div class="rev-chart-card">
      <div class="rev-chart-title">Cadence/h par OF (moy)</div>
      <canvas id="revChartCadence" height="220"></canvas>
    </div>
    <div class="rev-chart-card">
      <div class="rev-chart-title">TRS moyen par poste</div>
      <canvas id="revChartByPoste" height="220"></canvas>
    </div>
    <div class="rev-chart-card">
      <div class="rev-chart-title">TRS moyen par pilote</div>
      <canvas id="revChartByPilote" height="220"></canvas>
    </div>
  </div>

  <!-- Ligne 3 : Pauses/Nettoyages + Débordements + OF par jour -->
  <div class="rev-charts-grid">
    <div class="rev-chart-card">
      <div class="rev-chart-title">Pauses &amp; Nettoyages (min/jour)</div>
      <canvas id="revChartPauses" height="220"></canvas>
    </div>
    <div class="rev-chart-card">
      <div class="rev-chart-title">Débordements &amp; Changements série</div>
      <canvas id="revChartDebord" height="220"></canvas>
    </div>
    <div class="rev-chart-card">
      <div class="rev-chart-title">Nb OF par jour</div>
      <canvas id="revChartOFJour" height="220"></canvas>
    </div>
  </div>

  <!-- Tableau résumé postes (TRS) -->
  <div class="rev-tbl-section">
    <div class="rev-tbl-hdr">
      &#127942; R&eacute;sum&eacute; des postes (depuis feuille TRS)
      <span class="rev-count" id="cntTRS">0</span>
    </div>
    <div class="rev-tbl-wrap">
      <table id="revTblTRS">
        <thead><tr>
          <th>Date</th><th>Poste</th><th>Pilote</th><th>Co-Pilote</th><th>Nb Pers</th>
          <th>T. Ouv</th><th>T. D&eacute;cl</th><th>Écart Tps D/O</th><th>T. Marche</th>
          <th>Pannes</th><th>Ratt.</th><th>Pauses</th><th>R&eacute;unions</th>
          <th>Nb OF</th><th>Qte</th><th>Equiv</th><th>Moy/OF</th>
          <th style="min-width:70px">TRS</th>
          <th>Arr. Prévus</th><th>D&eacute;bord.</th>
        </tr></thead>
        <tbody id="bodyTRS"></tbody>
      </table>
    </div>
  </div>

  <!-- Tableau détail OF -->
  <div class="rev-tbl-section">
    <div class="rev-tbl-hdr">
      &#128203; D&eacute;tail de tous les OF
      <span class="rev-count" id="cntOF">0</span>
    </div>
    <div class="rev-tbl-wrap">
      <table id="revTblOF">
        <thead><tr>
          <th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th>
          <th>D&eacute;but</th><th>Fin</th><th>Taille</th><th>Code</th><th>Type</th>
          <th>Qte Fab</th><th>Qte Emb</th><th>Equiv</th><th>Dur&eacute;e</th>
          <th>Cad/h</th><th>Cad/h/pers</th>
          <th>Chgt S&eacute;rie</th><th>Mq MP</th><th>Mq Pers</th><th>Nettoyage</th>
          <th>Commentaire</th>
        </tr></thead>
        <tbody id="bodyOF"></tbody>
      </table>
    </div>
  </div>

  <!-- Tableau tous événements -->
  <div class="rev-tbl-section">
    <div class="rev-tbl-hdr">
      &#9888; Tous les &eacute;v&eacute;nements &amp; arr&ecirc;ts
      <span class="rev-count" id="cntEvts">0</span>
    </div>
    <div class="rev-tbl-wrap">
      <table id="revTblEvts">
        <thead><tr>
          <th>Type</th><th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th>
          <th>Taille</th><th>Type Prod</th>
          <th>D&eacute;but</th><th>Fin</th><th>Dur&eacute;e</th><th>Commentaire</th>
        </tr></thead>
        <tbody id="bodyEvts"></tbody>
      </table>
    </div>
  </div>

</div><!-- /rev-body -->

<div class="footer">
  KPI-ORC Revue Compl&egrave;te &bull; Donn&eacute;es au {rev_last_update} &bull; Auto-refresh 8h
</div>

</div><!-- /tab-review -->

<script>
var VALID_TABS = ['supervision','dashboard','review'];
function showTab(name) {{
  if (VALID_TABS.indexOf(name) === -1) name = 'supervision';
  document.querySelectorAll('.tab-content').forEach(function(el) {{
    el.classList.remove('visible');
  }});
  document.querySelectorAll('.tab-btn').forEach(function(el) {{
    el.classList.remove('active');
  }});
  var tc = document.getElementById('tab-' + name);
  var bc = document.getElementById('btn-' + name);
  if (tc) tc.classList.add('visible');
  if (bc) bc.classList.add('active');
  try {{ localStorage.setItem('kpi_orc_tab', name); }} catch(e) {{}}
}}

(function() {{
  var saved = '';
  try {{ saved = localStorage.getItem('kpi_orc_tab') || ''; }} catch(e) {{}}
  var hash = (location.hash || '').replace('#','');
  var tab = hash || saved || 'supervision';
  if (VALID_TABS.indexOf(tab) === -1) tab = 'supervision';
  showTab(tab);
}})();

// Countdown 15s – reload uniquement sur onglet supervision
(function() {{
  var s = 15, el = document.getElementById('cdown');
  if (!el) return;
  setInterval(function() {{
    s--;
    if (s <= 0) {{
      s = 15;
      var supTab = document.getElementById('tab-supervision');
      if (supTab && supTab.classList.contains('visible')) {{
        location.reload();
        return;
      }}
    }}
    el.textContent = '\\u21bb Mise \\u00e0 jour dans ' + s + 's';
  }}, 1000);
}})();


function _initDashboardCharts() {{
  if (typeof Chart === 'undefined') return;
  try {{
new Chart(document.getElementById('chartTRS'), {{
  type: 'bar',
  data: {{
    labels: {_json.dumps(chart_labels)},
    datasets: [{{
      label: 'TRS %', data: {_json.dumps(chart_trs)},
      backgroundColor: {_json.dumps(chart_colors)},
      borderRadius: 6, borderSkipped: false
    }}]
  }},
  options: {{
    indexAxis: 'y',
    plugins: {{ legend: {{ display: false }},
      tooltip: {{ callbacks: {{ label: function(c) {{ return c.raw.toFixed(1)+'%'; }} }} }} }},
    scales: {{
      x: {{ min:0, max:100, grid: {{ color:'#f1f5f9' }}, ticks: {{ callback: function(v){{ return v+'%'; }} }} }},
      y: {{ grid: {{ display:false }} }}
    }}
  }}
}});

new Chart(document.getElementById('chartTemps'), {{
  type: 'bar',
  data: {{
    labels: {_json.dumps(chart_labels)},
    datasets: [
      {{ label:'Production',  data:{_json.dumps(d_marche)},  backgroundColor:'#16a34a', borderWidth:0 }},
      {{ label:'Pannes',      data:{_json.dumps(d_pannes)},  backgroundColor:'#dc2626', borderWidth:0 }},
      {{ label:'Rattrapages', data:{_json.dumps(d_ratt)},    backgroundColor:'#d97706', borderWidth:0 }},
      {{ label:'Pauses',      data:{_json.dumps(d_pauses)},  backgroundColor:'#2563eb', borderWidth:0 }},
      {{ label:'R\\u00e9unions',    data:{_json.dumps(d_reunions)},backgroundColor:'#7c3aed', borderWidth:0 }},
      {{ label:'\\u00c9cart',       data:{_json.dumps(d_ecart)},   backgroundColor:'#94a3b8', borderWidth:0 }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    scales: {{
      x: {{ stacked:true, grid:{{ color:'#f1f5f9' }}, ticks:{{ callback:function(v){{ return v+'min'; }} }} }},
      y: {{ stacked:true, grid:{{ display:false }} }}
    }},
    plugins: {{ legend:{{ position:'bottom', labels:{{ boxWidth:12, font:{{ size:10 }} }} }} }}
  }}
}});

new Chart(document.getElementById('chartPareto'), {{
  type: 'bar',
  data: {{
    labels: {_json.dumps(pareto_labels)},
    datasets: [{{
      label: 'Minutes', data: {_json.dumps(pareto_values)},
      backgroundColor: {_json.dumps(pareto_colors)},
      borderRadius: 4, borderSkipped: false
    }}]
  }},
  options: {{
    plugins: {{ legend:{{ display:false }} }},
    scales: {{
      x: {{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }}, maxRotation:40 }} }},
      y: {{ grid:{{ color:'#f1f5f9' }}, ticks:{{ callback:function(v){{ return v+'m'; }} }} }}
    }}
  }}
}});

{gauge_charts_js}
{cadence_charts_js}

new Chart(document.getElementById('chartParetoAll'), {{
  type: 'bar',
  data: {{
    labels: {pareto_all_labels},
    datasets: [{{ label: 'Minutes', data: {pareto_all_values}, backgroundColor: {pareto_all_colors}, borderRadius: 4, borderSkipped: false }}]
  }},
  options: {{
    plugins: {{ legend:{{ display:false }} }},
    scales: {{
      x: {{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }}, maxRotation:40 }} }},
      y: {{ grid:{{ color:'#f1f5f9' }}, ticks:{{ callback:function(v){{ return v+'m'; }} }} }}
    }}
  }}
}});
{pareto_poste_js}

// ═══════════════ REVUE COMPLÈTE — données & logique ═══════════════
const REV_DATA = {rev_data_js};
const REV_EVTS = {rev_evts_js};
const REV_TRS  = {rev_trs_js};

var revCharts = {{}};

function parseRevDate(s) {{
  if (!s) return null;
  var p = String(s).trim().split('/');
  if (p.length === 3) return new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
  p = String(s).trim().split('-');
  if (p.length === 3) return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
  return null;
}}

function isoDate(d) {{
  if (!d) return '';
  var m = ('0'+(d.getMonth()+1)).slice(-2);
  var dd = ('0'+d.getDate()).slice(-2);
  return d.getFullYear()+'-'+m+'-'+dd;
}}

function trsColor(v) {{
  return v >= 80 ? '#16a34a' : (v >= 60 ? '#d97706' : '#dc2626');
}}

function hmsToMin(s) {{
  if (!s) return 0;
  var p = String(s).split(':');
  if (p.length === 3) return parseInt(p[0])*60 + parseInt(p[1]) + parseInt(p[2])/60;
  return 0;
}}

function parseTRS(s) {{
  if (!s) return 0;
  return parseFloat(String(s).replace('%','').replace(',','.') || '0') || 0;
}}

function badgeEvt(label) {{
  var l = String(label).toLowerCase();
  var cls = 'badge-other';
  if (l.indexOf('pb')>=0||l.indexOf('panne')>=0) cls='badge-pb';
  else if (l.indexOf('ratt')>=0) cls='badge-ratt';
  else if (l.indexOf('nettoyage')>=0) cls='badge-nett';
  else if (l.indexOf('pause')>=0) cls='badge-pause';
  else if (l.indexOf('serie')>=0||l.indexOf('série')>=0||l.indexOf('chgt')>=0) cls='badge-chgt';
  return '<span class="badge '+cls+'">'+esc(label)+'</span>';
}}

function esc(s) {{
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}}

function getFilters() {{
  var startVal = document.getElementById('dtStart').value;
  var endVal   = document.getElementById('dtEnd').value;
  var poste    = document.getElementById('fPoste').value;
  var pilote   = document.getElementById('fPilote').value;
  var startDt  = startVal ? new Date(startVal) : null;
  var endDt    = endVal   ? new Date(endVal+'T23:59:59') : null;
  return {{startDt:startDt, endDt:endDt, poste:poste, pilote:pilote}};
}}

function filterData(f) {{
  var data = REV_DATA.filter(function(r) {{
    var dt = parseRevDate(r[1]);
    if (f.startDt && dt && dt < f.startDt) return false;
    if (f.endDt   && dt && dt > f.endDt)   return false;
    if (f.poste   && r[2] !== f.poste)      return false;
    if (f.pilote  && r[3] !== f.pilote)     return false;
    return true;
  }});
  var evts = REV_EVTS.filter(function(r) {{
    var dt = parseRevDate(r[2]);
    if (f.startDt && dt && dt < f.startDt) return false;
    if (f.endDt   && dt && dt > f.endDt)   return false;
    if (f.poste   && r[3] !== f.poste)      return false;
    if (f.pilote  && r[4] !== f.pilote)     return false;
    return true;
  }});
  var trs = REV_TRS.filter(function(r) {{
    var dt = parseRevDate(r[0]);
    if (f.startDt && dt && dt < f.startDt) return false;
    if (f.endDt   && dt && dt > f.endDt)   return false;
    if (f.poste   && r[1] !== f.poste)      return false;
    if (f.pilote  && r[2] !== f.pilote)     return false;
    return true;
  }});
  return {{data:data, evts:evts, trs:trs}};
}}

function destroyChart(id) {{
  if (revCharts[id]) {{ try {{ revCharts[id].destroy(); }} catch(e) {{}} revCharts[id]=null; }}
}}

function makeChart(id, cfg) {{
  if (typeof Chart === 'undefined') return;
  destroyChart(id);
  var el = document.getElementById(id);
  if (!el) return;
  try {{ revCharts[id] = new Chart(el, cfg); }} catch(e) {{ console.warn('Chart error:', e.message); }}
}}

function applyFilters() {{
  var f = getFilters();
  var d = filterData(f);
  updateKPIs(d);
  updateCharts(d);
  updateTables(d);
}}

function updateKPIs(d) {{
  // Nb postes (lignes TRS)
  document.getElementById('kpiNbPostes').textContent = d.trs.length;
  // Nb OF
  document.getElementById('kpiNbOF').textContent = d.data.length;
  // TRS moyen
  var trsVals = d.trs.map(function(r){{return parseTRS(r[19]);}}).filter(function(v){{return v>0;}});
  var trsAvg = trsVals.length ? (trsVals.reduce(function(a,b){{return a+b;}},0)/trsVals.length).toFixed(1)+'%' : '—';
  var trsEl = document.getElementById('kpiTrsAvg');
  trsEl.textContent = trsAvg;
  if (trsVals.length) trsEl.style.color = trsColor(parseFloat(trsAvg));
  // Qté produite
  var qte = d.data.reduce(function(a,r){{return a+(parseInt(r[13])||0);}},0);
  document.getElementById('kpiQteTot').textContent = qte.toLocaleString('fr-FR');
  // Équivalence
  var equiv = d.data.reduce(function(a,r){{return a+(parseFloat(String(r[15]).replace(',','.'))||0);}},0);
  document.getElementById('kpiEquivTot').textContent = equiv.toFixed(1);
  // Arrêts
  document.getElementById('kpiNbArrets').textContent = d.evts.length;
}}

function updateCharts(d) {{
  // — Chart TRS par poste (en ordre chronologique)
  var trsLabels = d.trs.map(function(r){{return r[0].slice(0,5)+'  '+r[1];}});
  var trsVals   = d.trs.map(function(r){{return parseTRS(r[19]);}});
  var trsColors = trsVals.map(function(v){{return trsColor(v);}});
  makeChart('revChartTRS', {{
    type:'bar',
    data:{{labels:trsLabels, datasets:[{{
      label:'TRS %', data:trsVals, backgroundColor:trsColors,
      borderRadius:4, borderSkipped:false}}]}},
    options:{{
      indexAxis:'y',
      plugins:{{legend:{{display:false}},
        tooltip:{{callbacks:{{label:function(c){{return c.raw.toFixed(1)+'%';}}}}}}
      }},
      scales:{{
        x:{{min:0,max:100,grid:{{color:'#f1f5f9'}},ticks:{{callback:function(v){{return v+'%';}}}}}},
        y:{{grid:{{display:false}},ticks:{{font:{{size:9}}}}}}
      }}
    }}
  }});

  // — Chart Répartition temps
  var tempsLabels = d.trs.map(function(r){{return r[0].slice(0,5)+'  '+r[1];}});
  makeChart('revChartTemps', {{
    type:'bar',
    data:{{
      labels:tempsLabels,
      datasets:[
        {{label:'Production', data:d.trs.map(function(r){{return hmsToMin(r[8]);}}), backgroundColor:'#16a34a',borderWidth:0}},
        {{label:'Pannes',     data:d.trs.map(function(r){{return hmsToMin(r[11]);}}),backgroundColor:'#dc2626',borderWidth:0}},
        {{label:'Rattrapages',data:d.trs.map(function(r){{return hmsToMin(r[12]);}}),backgroundColor:'#d97706',borderWidth:0}},
        {{label:'Pauses',     data:d.trs.map(function(r){{return hmsToMin(r[13]);}}),backgroundColor:'#2563eb',borderWidth:0}},
        {{label:'Réunions',   data:d.trs.map(function(r){{return hmsToMin(r[14]);}}),backgroundColor:'#7c3aed',borderWidth:0}},
        {{label:'Écart Tps déclaré/ouverture',      data:d.trs.map(function(r){{return hmsToMin(r[7]); }}),backgroundColor:'#94a3b8',borderWidth:0}}
      ]
    }},
    options:{{
      indexAxis:'y',
      scales:{{
        x:{{stacked:true,grid:{{color:'#f1f5f9'}},ticks:{{callback:function(v){{return v+'m';}}}}}},
        y:{{stacked:true,grid:{{display:false}},ticks:{{font:{{size:9}}}}}}
      }},
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:9}}}}}}}}
    }}
  }});

  // — Pareto arrêts
  var stopTotals = {{}};
  d.evts.forEach(function(r) {{
    var lbl = String(r[0]||'').trim();
    if (!lbl) return;
    stopTotals[lbl] = (stopTotals[lbl]||0) + hmsToMin(r[18]);
  }});
  var pSorted = Object.entries(stopTotals).sort(function(a,b){{return b[1]-a[1];}}).slice(0,15);
  var pLabels = pSorted.map(function(x){{return x[0];}});
  var pVals   = pSorted.map(function(x){{return Math.round(x[1]*10)/10;}});
  var pColors = pLabels.map(function(l){{
    var ll=l.toLowerCase();
    if(ll.indexOf('pb')>=0||ll.indexOf('panne')>=0) return '#dc2626';
    if(ll.indexOf('ratt')>=0) return '#d97706';
    if(ll.indexOf('nettoyage')>=0) return '#0284c7';
    if(ll.indexOf('pause')>=0) return '#2563eb';
    if(ll.indexOf('serie')>=0||ll.indexOf('série')>=0) return '#7c3aed';
    return '#64748b';
  }});
  makeChart('revChartPareto', {{
    type:'bar',
    data:{{labels:pLabels, datasets:[{{label:'min',data:pVals,backgroundColor:pColors,borderRadius:4,borderSkipped:false}}]}},
    options:{{
      indexAxis:'y',
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{color:'#f1f5f9'}},ticks:{{callback:function(v){{return v+'m';}}}}}},
        y:{{grid:{{display:false}},ticks:{{font:{{size:9}}}}}}
      }}
    }}
  }});

  // — Cadence/h par OF (scatter by date)
  var cadData = d.data.map(function(r){{
    return {{x:parseRevDate(r[1]),y:parseFloat(String(r[19]).replace(',','.'))||0,of:r[0]}};
  }}).filter(function(x){{return x.x && x.y>0;}});
  cadData.sort(function(a,b){{return a.x-b.x;}});
  makeChart('revChartCadence', {{
    type:'line',
    data:{{
      labels:cadData.map(function(d){{return (d.x?String(d.x.getDate()).padStart(2,'0')+'/'+(String(d.x.getMonth()+1).padStart(2,'0')):'')+'  '+d.of;}}),
      datasets:[{{label:'Cad/h',data:cadData.map(function(d){{return d.y;}}),
        borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,0.08)',
        borderWidth:2,pointRadius:3,tension:0.3,fill:true}}]
    }},
    options:{{
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{display:false}},ticks:{{font:{{size:8}},maxRotation:45}}}},
        y:{{grid:{{color:'#f1f5f9'}}}}
      }}
    }}
  }});

  // — TRS par poste (agrégé)
  var byPoste = {{}};
  d.trs.forEach(function(r) {{
    var p = r[1]; var v = parseTRS(r[19]);
    if (!byPoste[p]) byPoste[p] = [];
    if (v > 0) byPoste[p].push(v);
  }});
  var bpLabels = Object.keys(byPoste).sort();
  var bpVals   = bpLabels.map(function(p){{
    var vals=byPoste[p]; return vals.length?(vals.reduce(function(a,b){{return a+b;}},0)/vals.length).toFixed(1):0;
  }});
  makeChart('revChartByPoste', {{
    type:'bar',
    data:{{labels:bpLabels, datasets:[{{label:'TRS moy %',data:bpVals,
      backgroundColor:bpVals.map(function(v){{return trsColor(parseFloat(v))+'cc';}}),
      borderRadius:6,borderSkipped:false}}]}},
    options:{{
      plugins:{{legend:{{display:false}},
        tooltip:{{callbacks:{{label:function(c){{return c.raw+'%';}}}}}}
      }},
      scales:{{
        x:{{grid:{{display:false}}}},
        y:{{min:0,max:100,grid:{{color:'#f1f5f9'}},ticks:{{callback:function(v){{return v+'%';}}}}}}
      }}
    }}
  }});

  // — TRS par pilote
  var byPilote = {{}};
  d.trs.forEach(function(r) {{
    var p = r[2]; var v = parseTRS(r[19]);
    if (!byPilote[p]) byPilote[p] = [];
    if (v > 0) byPilote[p].push(v);
  }});
  var pilLabels = Object.keys(byPilote).sort();
  var pilVals   = pilLabels.map(function(p){{
    var vals=byPilote[p]; return vals.length?(vals.reduce(function(a,b){{return a+b;}},0)/vals.length).toFixed(1):0;
  }});
  makeChart('revChartByPilote', {{
    type:'bar',
    data:{{labels:pilLabels, datasets:[{{label:'TRS moy %',data:pilVals,
      backgroundColor:pilVals.map(function(v){{return trsColor(parseFloat(v))+'cc';}}),
      borderRadius:6,borderSkipped:false}}]}},
    options:{{
      plugins:{{legend:{{display:false}},
        tooltip:{{callbacks:{{label:function(c){{return c.raw+'%';}}}}}}
      }},
      scales:{{
        x:{{grid:{{display:false}},ticks:{{font:{{size:9}}}}}},
        y:{{min:0,max:100,grid:{{color:'#f1f5f9'}},ticks:{{callback:function(v){{return v+'%';}}}}}}
      }}
    }}
  }});

  // — Pauses & Nettoyages par jour
  var pauseByDate = {{}};
  d.evts.forEach(function(r) {{
    var dt = String(r[2]).slice(0,5);
    var lbl = String(r[0]).toLowerCase();
    if (!pauseByDate[dt]) pauseByDate[dt] = {{pause:0, nett:0}};
    var dur = hmsToMin(r[18]);
    if (lbl.indexOf('pause')>=0) pauseByDate[dt].pause += dur;
    else if (lbl.indexOf('nettoyage')>=0) pauseByDate[dt].nett += dur;
  }});
  var pdDates = Object.keys(pauseByDate).sort();
  makeChart('revChartPauses', {{
    type:'bar',
    data:{{
      labels:pdDates,
      datasets:[
        {{label:'Pauses (min)', data:pdDates.map(function(d){{return Math.round(pauseByDate[d].pause*10)/10;}}), backgroundColor:'#2563ebcc',borderRadius:4,borderWidth:0}},
        {{label:'Nettoyages (min)', data:pdDates.map(function(d){{return Math.round(pauseByDate[d].nett*10)/10;}}), backgroundColor:'#0284c7cc',borderRadius:4,borderWidth:0}}
      ]
    }},
    options:{{
      scales:{{
        x:{{grid:{{display:false}},ticks:{{font:{{size:9}}}}}},
        y:{{grid:{{color:'#f1f5f9'}},ticks:{{callback:function(v){{return v+'m';}}}}}}
      }},
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:9}}}}}}}}
    }}
  }});

  // Helper: trier des clés DD/MM/YYYY par date réelle
  function sortDateKeys(keys) {{
    return keys.slice().sort(function(a,b) {{
      var pa=String(a).split('/'), pb=String(b).split('/');
      var da=new Date(parseInt(pa[2]||2000),parseInt(pa[1]||1)-1,parseInt(pa[0]||1));
      var db=new Date(parseInt(pb[2]||2000),parseInt(pb[1]||1)-1,parseInt(pb[0]||1));
      return da-db;
    }});
  }}

  // — Débordements & Changements série (clé = date complète DD/MM/YYYY)
  var debByDate = {{}};
  d.trs.forEach(function(r) {{
    var dt = String(r[0]).slice(0,10);
    if (!debByDate[dt]) debByDate[dt] = {{deb:0, chgt:0}};
    debByDate[dt].deb  += hmsToMin(r[10]);
    debByDate[dt].chgt += hmsToMin(r[9]);
  }});
  d.data.forEach(function(r) {{
    var dt = String(r[1]).slice(0,10);
    if (!debByDate[dt]) debByDate[dt] = {{deb:0, chgt:0}};
    if (String(r[29]).trim() && String(r[29]).trim() !== '0') debByDate[dt].chgt += 1;
  }});
  var dbKeys = sortDateKeys(Object.keys(debByDate));
  var dbLabels = dbKeys.map(function(k){{return k.slice(0,5);}});
  makeChart('revChartDebord', {{
    type:'bar',
    data:{{
      labels:dbLabels,
      datasets:[
        {{label:'Débord. (min)', data:dbKeys.map(function(k){{return Math.round(debByDate[k].deb*10)/10;}}), backgroundColor:'#dc2626cc',borderRadius:4,borderWidth:0}},
        {{label:'Chgt série (nb)', data:dbKeys.map(function(k){{return debByDate[k].chgt;}}), backgroundColor:'#7c3aedcc',borderRadius:4,borderWidth:0}}
      ]
    }},
    options:{{
      scales:{{
        x:{{grid:{{display:false}},ticks:{{font:{{size:9}}}},title:{{display:true,text:'Date',font:{{size:9}},color:'#94a3b8'}}}},
        y:{{grid:{{color:'#f1f5f9'}},title:{{display:true,text:'min / nb',font:{{size:9}},color:'#94a3b8'}}}}
      }},
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:9}}}}}}}}
    }}
  }});

  // — Nb OF par jour (clé = date complète DD/MM/YYYY)
  var ofByDate = {{}};
  d.data.forEach(function(r) {{
    var dt = String(r[1]).slice(0,10);
    ofByDate[dt] = (ofByDate[dt]||0) + 1;
  }});
  var ofKeys = sortDateKeys(Object.keys(ofByDate));
  var ofLabels = ofKeys.map(function(k){{return k.slice(0,5);}});
  makeChart('revChartOFJour', {{
    type:'bar',
    data:{{labels:ofLabels, datasets:[{{label:'Nb OF',data:ofKeys.map(function(k){{return ofByDate[k];}}) ,
      backgroundColor:'#1e3a5fcc',borderRadius:4,borderSkipped:false,borderWidth:0}}]}},
    options:{{
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{display:false}},ticks:{{font:{{size:9}}}},title:{{display:true,text:'Date',font:{{size:9}},color:'#94a3b8'}}}},
        y:{{grid:{{color:'#f1f5f9'}},ticks:{{stepSize:1}},title:{{display:true,text:'Nombre OF',font:{{size:9}},color:'#94a3b8'}}}}
      }}
    }}
  }});
}}

function updateTables(d) {{
  // TRS table
  var bTRS = document.getElementById('bodyTRS');
  if (bTRS) {{
    document.getElementById('cntTRS').textContent = d.trs.length;
    var rows = d.trs.slice().reverse().map(function(r) {{
      var tv = parseTRS(r[19]);
      var tc = trsColor(tv);
      return '<tr>'
        +'<td>'+esc(r[0])+'</td><td><b>'+esc(r[1])+'</b></td><td>'+esc(r[2])+'</td><td>'+esc(r[3])+'</td><td>'+esc(r[4])+'</td>'
        +'<td>'+esc(r[5])+'</td><td>'+esc(r[6])+'</td><td>'+esc(r[7])+'</td><td>'+esc(r[8])+'</td>'
        +'<td style="color:#dc2626">'+esc(r[11])+'</td>'
        +'<td style="color:#d97706">'+esc(r[12])+'</td>'
        +'<td>'+esc(r[13])+'</td><td>'+esc(r[14])+'</td>'
        +'<td style="color:#2563eb;font-weight:700">'+esc(r[15])+'</td>'
        +'<td>'+esc(r[16])+'</td><td>'+esc(r[17])+'</td><td>'+esc(r[18])+'</td>'
        +'<td style="color:'+tc+';font-weight:800;font-size:1.05em">'+esc(r[19])+'</td>'
        +'<td>'+esc(r[9])+'</td><td style="color:#dc2626">'+esc(r[10])+'</td>'
        +'</tr>';
    }});
    bTRS.innerHTML = rows.join('') || '<tr><td colspan="20" class="empty">Aucun poste</td></tr>';
  }}

  // OF table
  var bOF = document.getElementById('bodyOF');
  if (bOF) {{
    document.getElementById('cntOF').textContent = d.data.length;
    var rows = d.data.slice().reverse().map(function(r) {{
      var rj = JSON.stringify(r).replace(/'/g,"&#39;");
      return '<tr style="cursor:pointer" onclick="openOfModal(JSON.parse(this.dataset.row))" data-row=\''+rj+'\' title="Cliquer pour voir tous les détails">'
        +'<td><b>'+esc(r[0])+'</b></td><td>'+esc(r[1])+'</td><td>'+esc(r[2])+'</td><td>'+esc(r[3])+'</td>'
        +'<td>'+esc(r[17])+'</td><td>'+esc(r[18])+'</td>'
        +'<td>'+esc(r[6])+'</td><td>'+esc(r[7])+'</td><td>'+esc(r[8])+'</td>'
        +'<td>'+esc(r[13])+'</td><td>'+esc(r[14])+'</td><td>'+esc(r[15])+'</td><td>'+esc(r[16])+'</td>'
        +'<td><b>'+esc(r[19])+'</b></td><td>'+esc(r[20])+'</td>'
        +'<td>'+esc(r[29])+'</td><td>'+esc(r[30])+'</td><td>'+esc(r[31])+'</td><td>'+esc(r[32])+'</td>'
        +'<td style="max-width:180px;white-space:normal">'+esc(r[56])+'</td>'
        +'</tr>';
    }});
    bOF.innerHTML = rows.join('') || '<tr><td colspan="20" class="empty">Aucun OF</td></tr>';
  }}

  // Events table
  var bEvts = document.getElementById('bodyEvts');
  if (bEvts) {{
    document.getElementById('cntEvts').textContent = d.evts.length;
    var rows = d.evts.slice().reverse().map(function(r) {{
      return '<tr>'
        +'<td>'+badgeEvt(r[0])+'</td><td>'+esc(r[1])+'</td><td>'+esc(r[2])+'</td>'
        +'<td>'+esc(r[3])+'</td><td>'+esc(r[4])+'</td>'
        +'<td>'+esc(r[7])+'</td><td>'+esc(r[8])+'</td>'
        +'<td>'+esc(r[16])+'</td><td>'+esc(r[17])+'</td>'
        +'<td><b>'+esc(r[18])+'</b></td>'
        +'<td style="max-width:180px;white-space:normal">'+esc(r[19])+'</td>'
        +'</tr>';
    }});
    bEvts.innerHTML = rows.join('') || '<tr><td colspan="11" class="empty">Aucun événement</td></tr>';
  }}
}}

// DATA_HEADERS pour modal OF
var DATA_HEADERS = {_json.dumps(DATA_HEADERS, ensure_ascii=False)};

// ── Modal détail OF ────────────────────────────────────────────────────
var _modalRowData = null;
function openOfModal(rowData) {{
  _modalRowData = rowData;
  var modal = document.getElementById('ofDetailModal');
  var body  = document.getElementById('modalBody');
  var numEl = document.getElementById('modalOfNum');
  numEl.textContent = rowData[0] || '';
  document.getElementById('modalEditBtn').style.display = '';
  document.getElementById('modalSaveBtn').style.display = 'none';
  document.getElementById('modalCancelBtn').style.display = 'none';
  document.getElementById('modalEditMsg').textContent = '';
  var html = '';
  var sections = [
    {{title:'Identification', color:'#1e3a5f', bg:'#eff6ff', indices:[0,1,2,3,4,5]}},
    {{title:'Produit', color:'#16a34a', bg:'#f0fdf4', indices:[6,7,8,9,10,11,12,21]}},
    {{title:'Quantités & Qualité', color:'#7c3aed', bg:'#faf5ff', indices:[13,14,15,16,17,18,19,20,22,23,24,25,26,27,28]}},
    {{title:'Durées & Arrêts', color:'#d97706', bg:'#fffbeb', indices:[29,30,31,32,33,34,35,36,37,57]}},
    {{title:'Pannes', color:'#dc2626', bg:'#fef2f2', indices:[38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55]}},
    {{title:'Commentaire', color:'#64748b', bg:'#f8fafc', indices:[56]}}
  ];
  sections.forEach(function(sec) {{
    var secHtml = '';
    sec.indices.forEach(function(i) {{
      if (i >= DATA_HEADERS.length) return;
      var val = (i < rowData.length && rowData[i] !== '' && rowData[i] !== null && rowData[i] !== undefined) ? rowData[i] : '—';
      var isEmpty = (val === '—');
      secHtml += '<div style="background:white;border-radius:6px;padding:8px 10px;border:1px solid ' + (isEmpty ? '#f1f5f9' : '#e2e8f0') + '">'
        + '<div style="font-size:0.62em;color:#94a3b8;text-transform:uppercase;letter-spacing:0.3px">' + esc(DATA_HEADERS[i]) + '</div>'
        + '<div style="font-weight:700;color:' + (isEmpty ? '#cbd5e1' : '#1e3a5f') + ';margin-top:2px;word-break:break-word;font-size:0.95em">' + esc(val) + '</div>'
        + '</div>';
    }});
    if (!secHtml) return;
    html += '<div style="grid-column:1/-1;background:' + sec.bg + ';border-radius:10px;padding:12px;border-left:4px solid ' + sec.color + ';margin-bottom:4px">'
      + '<div style="font-size:0.72em;font-weight:700;color:' + sec.color + ';text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">' + esc(sec.title) + '</div>'
      + '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:6px">' + secHtml + '</div>'
      + '</div>';
  }});
  body.innerHTML = html || '<div style="color:#94a3b8;text-align:center;padding:20px">Aucune donnée</div>';
  modal.style.display = 'flex';
}}

function closeOfModal() {{
  document.getElementById('ofDetailModal').style.display = 'none';
}}

function startEditModal() {{
  if (!_modalRowData) return;
  var sections = [
    {{title:'Identification', color:'#1e3a5f', bg:'#eff6ff', indices:[0,1,2,3,4,5]}},
    {{title:'Produit', color:'#16a34a', bg:'#f0fdf4', indices:[6,7,8,9,10,11,12,21]}},
    {{title:'Quantités & Qualité', color:'#7c3aed', bg:'#faf5ff', indices:[13,14,15,16,17,18,19,20,22,23,24,25,26,27,28]}},
    {{title:'Durées & Arrêts', color:'#d97706', bg:'#fffbeb', indices:[29,30,31,32,33,34,35,36,37,57]}},
    {{title:'Pannes', color:'#dc2626', bg:'#fef2f2', indices:[38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55]}},
    {{title:'Commentaire', color:'#64748b', bg:'#f8fafc', indices:[56]}}
  ];
  var html = '<form id="editForm">';
  sections.forEach(function(sec) {{
    var secHtml = '';
    sec.indices.forEach(function(i) {{
      if (i >= DATA_HEADERS.length) return;
      var val = (i < _modalRowData.length && _modalRowData[i] !== null && _modalRowData[i] !== undefined) ? _modalRowData[i] : '';
      secHtml += '<div style="background:white;border-radius:6px;padding:8px 10px;border:1px solid #e2e8f0">'
        + '<label style="display:block;font-size:0.62em;color:#94a3b8;text-transform:uppercase;letter-spacing:0.3px;margin-bottom:4px" for="ef_' + i + '">' + esc(DATA_HEADERS[i]) + '</label>'
        + '<input id="ef_' + i + '" data-idx="' + i + '" type="text" value="' + esc(String(val)) + '" style="width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:4px;padding:4px 6px;font-size:0.92em;font-family:inherit" />'
        + '</div>';
    }});
    if (!secHtml) return;
    html += '<div style="grid-column:1/-1;background:' + sec.bg + ';border-radius:10px;padding:12px;border-left:4px solid ' + sec.color + ';margin-bottom:4px">'
      + '<div style="font-size:0.72em;font-weight:700;color:' + sec.color + ';text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">' + sec.title + '</div>'
      + '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:6px">' + secHtml + '</div>'
      + '</div>';
  }});
  html += '</form>';
  document.getElementById('modalBody').style.display = 'none';
  var eb = document.getElementById('modalEditBody');
  eb.innerHTML = html;
  eb.style.display = '';
  document.getElementById('modalEditBtn').style.display = 'none';
  document.getElementById('modalSaveBtn').style.display = '';
  document.getElementById('modalCancelBtn').style.display = '';
  document.getElementById('modalEditMsg').textContent = '';
}}

function cancelEditModal() {{
  document.getElementById('modalBody').style.display = '';
  document.getElementById('modalEditBody').style.display = 'none';
  document.getElementById('modalEditBtn').style.display = '';
  document.getElementById('modalSaveBtn').style.display = 'none';
  document.getElementById('modalCancelBtn').style.display = 'none';
  document.getElementById('modalEditMsg').textContent = '';
}}

function saveEditModal() {{
  if (!_modalRowData) return;
  var inputs = document.querySelectorAll('#editForm input[data-idx]');
  var newData = _modalRowData.slice();
  inputs.forEach(function(inp) {{
    var idx = parseInt(inp.getAttribute('data-idx'), 10);
    newData[idx] = inp.value;
  }});
  var ofNum = newData[0] || '';
  var date  = newData[1] || '';
  document.getElementById('modalSaveBtn').disabled = true;
  document.getElementById('modalEditMsg').textContent = 'Sauvegarde...';
  fetch('http://localhost:7892/edit', {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{of_num: ofNum, date: date, data: newData}})
  }}).then(function(r) {{ return r.json(); }}).then(function(res) {{
    if (res.ok) {{
      document.getElementById('modalEditMsg').textContent = '✔ Sauvegardé';
      document.getElementById('modalEditMsg').style.color = '#16a34a';
      _modalRowData = newData;
      setTimeout(function() {{ cancelEditModal(); }}, 1200);
    }} else {{
      document.getElementById('modalEditMsg').textContent = '✘ Erreur: ' + (res.error || 'inconnue');
      document.getElementById('modalEditMsg').style.color = '#dc2626';
      document.getElementById('modalSaveBtn').disabled = false;
    }}
  }}).catch(function(err) {{
    document.getElementById('modalEditMsg').textContent = '✘ ' + err;
    document.getElementById('modalEditMsg').style.color = '#dc2626';
    document.getElementById('modalSaveBtn').disabled = false;
  }});
}}

document.getElementById('ofDetailModal').addEventListener('click', function(e) {{
  if (e.target === this) closeOfModal();
}});

// ── Graphiques supervision (Canvas 2D, pas de CDN) ──────────────────
function _drawSupCharts() {{
  var supPie = document.getElementById('supPieChart');
  if (supPie) {{
    _drawDonut(supPie, {pie_values}, ['#16a34a','#dc2626','#94a3b8']);
  }}
  var supPar = document.getElementById('supParetoChart');
  if (supPar) {{
    supPar.width  = supPar.offsetWidth  || 300;
    supPar.height = supPar.offsetHeight || 180;
    _drawHBar(supPar, {sup_par_labels}, {sup_par_values}, {sup_par_colors});
  }}
}}
_drawSupCharts();
window.addEventListener('load', _drawSupCharts);

  }} catch(e) {{ console.warn('Chart init error:', e.message); }}
}}
// ── Initialisation filtres avec persistance localStorage ───────────────
(function() {{
  var today = new Date();
  var m30   = new Date(today); m30.setDate(m30.getDate()-30);
  var savedStart  = ''; try {{ savedStart  = localStorage.getItem('kpi_orc_dtStart')  || ''; }} catch(e) {{}}
  var savedEnd    = ''; try {{ savedEnd    = localStorage.getItem('kpi_orc_dtEnd')    || ''; }} catch(e) {{}}
  var savedPoste  = ''; try {{ savedPoste  = localStorage.getItem('kpi_orc_fPoste')  || ''; }} catch(e) {{}}
  var savedPilote = ''; try {{ savedPilote = localStorage.getItem('kpi_orc_fPilote') || ''; }} catch(e) {{}}

  document.getElementById('dtStart').value = savedStart  || isoDate(m30);
  document.getElementById('dtEnd').value   = savedEnd    || isoDate(today);
  var elPoste  = document.getElementById('fPoste');
  var elPilote = document.getElementById('fPilote');
  if (savedPoste  && Array.from(elPoste.options).some(function(o){{return o.value===savedPoste;}}))
    elPoste.value  = savedPoste;
  if (savedPilote && Array.from(elPilote.options).some(function(o){{return o.value===savedPilote;}}))
    elPilote.value = savedPilote;

  // Sauvegarder à chaque changement
  ['dtStart','dtEnd'].forEach(function(id) {{
    document.getElementById(id).addEventListener('change', function() {{
      try {{ localStorage.setItem('kpi_orc_'+id, this.value); }} catch(e) {{}}
    }});
  }});
  elPoste.addEventListener('change', function()  {{ try {{ localStorage.setItem('kpi_orc_fPoste',  this.value); }} catch(e) {{}} }});
  elPilote.addEventListener('change', function() {{ try {{ localStorage.setItem('kpi_orc_fPilote', this.value); }} catch(e) {{}} }});

  if (document.getElementById('tab-review').classList.contains('visible')) {{
    applyFilters();
  }}
}})();

// Appliquer filtres quand on active l'onglet review
var _origShowTab = showTab;
showTab = function(name) {{
  _origShowTab(name);
  if (name === 'review') {{
    setTimeout(function() {{ applyFilters(); }}, 50);
  }}
}};
_initDashboardCharts();
if (typeof Chart === 'undefined') {{
  var _cjsPoll = setInterval(function() {{
    if (typeof Chart !== 'undefined') {{ _initDashboardCharts(); clearInterval(_cjsPoll); }}
  }}, 500);
  setTimeout(function() {{ clearInterval(_cjsPoll); }}, 15000);
}}
</script>


</body>
</html>"""

        with open(html_path, "w", encoding="utf-8") as fh:
            fh.write(html)

    def _show_loading(self, msg="Chargement…"):
        if self._loading_anim_id:
            try:
                self.root.after_cancel(self._loading_anim_id)
            except Exception:
                pass
            self._loading_anim_id = None
        if self._loading_overlay:
            try:
                self._loading_overlay.destroy()
            except Exception:
                pass
        ov = tk.Frame(self.root, bg=NAVY)
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        ov.lift()
        center = tk.Frame(ov, bg=NAVY)
        center.place(relx=0.5, rely=0.45, anchor="center")
        cv = tk.Canvas(center, width=90, height=90, bg=NAVY, highlightthickness=0)
        cv.pack()
        lbl = tk.Label(center, text=msg, bg=NAVY, fg=WHITE,
                       font=("Arial", 17, "bold"))
        lbl.pack(pady=(18, 0))
        tk.Label(center, text="Veuillez patienter…", bg=NAVY, fg=NAVY_L,
                 font=("Arial", 11)).pack(pady=(4, 0))
        self._loading_overlay  = ov
        self._loading_canvas   = cv
        self._loading_msg_lbl  = lbl
        self._loading_angle    = 0
        self._animate_spinner()

    def _loading_set_msg(self, msg):
        if self._loading_msg_lbl:
            try:
                self._loading_msg_lbl.config(text=msg)
            except Exception:
                pass

    def _animate_spinner(self):
        if not self._loading_canvas:
            return
        try:
            cv = self._loading_canvas
            cv.delete("arc")
            a = self._loading_angle
            cv.create_arc(8, 8, 82, 82, start=a, extent=280,
                          style="arc", outline=GREEN, width=7, tags="arc")
            cv.create_arc(8, 8, 82, 82, start=a + 280, extent=80,
                          style="arc", outline=NAVY_L, width=7, tags="arc")
            self._loading_angle = (a + 9) % 360
            self._loading_anim_id = self.root.after(28, self._animate_spinner)
        except Exception:
            self._loading_anim_id = None

    def _hide_loading(self):
        if self._loading_anim_id:
            try:
                self.root.after_cancel(self._loading_anim_id)
            except Exception:
                pass
            self._loading_anim_id = None
        self._loading_canvas  = None
        self._loading_msg_lbl = None
        if self._loading_overlay:
            try:
                self._loading_overlay.destroy()
            except Exception:
                pass
            self._loading_overlay = None

    def _center_on_root(self, win, w, h):
        self.root.update_idletasks()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        if rw < 100:
            rw = self.root.winfo_screenwidth()
            rh = self.root.winfo_screenheight()
            rx, ry = 0, 0
        else:
            rx = self.root.winfo_x()
            ry = self.root.winfo_y()
        x = max(0, rx + (rw - w) // 2)
        y = max(0, ry + (rh - h) // 2)
        win.geometry(f"{w}x{h}+{x}+{y}")

    def _get_wb(self, path):
        """Charge le workbook pour écriture (sans cache — évite la corruption)."""
        try:
            if not os.path.exists(path):
                return None
            wb = load_workbook(path)
            self._wb_cache       = wb
            self._wb_path_cache  = path
            self._wb_mtime_cache = os.path.getmtime(path)
            return wb
        except Exception:
            return None

    def _invalidate_wb_cache(self):
        self._wb_cache = None
        self._wb_path_cache  = ""
        self._wb_mtime_cache = 0.0

    def _write_rows_to_events_sheet(self, wb, events_rows):
        """Écrit des lignes pré-construites dans l'onglet Evenements."""
        ws = self._ensure_events_sheet(wb)
        for ev_row in events_rows:
            ws.append(ev_row)
            self._format_row(ws, ws.max_row)

    def _write_excel(self, row, v):
        """Écriture Excel arrière-plan (pour les appels hors fin-de-prod)."""
        path = self.cfg.get("db_path", "")
        if not path:
            messagebox.showwarning("Attention", "Aucune base de données !")
            return False

        events_rows = self._build_events_rows(v)
        payload = {"db_path": path, "row": row, "events_rows": events_rows}
        try:
            with open(PENDING_FILE, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, default=str)
        except Exception:
            pass

        def _bg():
            try:
                with self._excel_lock:
                    wb = self._get_wb(path)
                    if wb is None:
                        self.root.after(0, self._schedule_pending_retry)
                        return
                    ws_d = self._ensure_data_sheet(wb)
                    ws_d.append(row)
                    self._format_row(ws_d, ws_d.max_row)
                    self._write_rows_to_events_sheet(wb, events_rows)
                    self._safe_excel_save(wb, path)
                    self._wb_mtime_cache = os.path.getmtime(path)
                    self._copy_excel_for_dashboard(path)
                    try:
                        os.remove(PENDING_FILE)
                    except Exception:
                        pass
                    self.root.after(0, self._reload_and_refresh)
                    self.root.after(0, lambda: _toast(
                        self.root, "✔  Déclaration enregistrée dans Excel",
                        bg=GREEN, duration=3000))
            except PermissionError:
                self._invalidate_wb_cache()
                self.root.after(0, self._schedule_pending_retry)
            except Exception as e:
                self._invalidate_wb_cache()
                err_msg = str(e)
                self.root.after(0, self._schedule_pending_retry)
                self.root.after(0, lambda: messagebox.showerror("Erreur Excel", err_msg))

        threading.Thread(target=_bg, daemon=True).start()
        return True

    def _ensure_events_sheet(self, wb):
        if "Evenements" not in wb.sheetnames:
            ws = wb.create_sheet("Evenements")
            for i, h in enumerate(EVT_HEADERS, start=1):
                ws.cell(1, i).value = h
            self._format_row(ws, 1)
        else:
            ws = wb["Evenements"]
            existing = [ws.cell(1, i).value for i in range(1, len(EVT_HEADERS) + 1)]
            if existing != EVT_HEADERS:
                for i, h in enumerate(EVT_HEADERS, start=1):
                    ws.cell(1, i).value = h
                self._format_row(ws, 1)
        return wb["Evenements"]

    def _ensure_data_sheet(self, wb):
        if "Data" not in wb.sheetnames:
            ws = wb.create_sheet("Data", 0)
            for i, h in enumerate(DATA_HEADERS, start=1):
                ws.cell(1, i).value = h
            self._format_row(ws, 1)
        else:
            ws = wb["Data"]
            for i, h in enumerate(DATA_HEADERS, start=1):
                if ws.cell(1, i).value is None:
                    ws.cell(1, i).value = h
        return wb["Data"]

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
            self._format_row(ws, ws.max_row)

    def _write_changement_of_excel(self, start_dt, end_dt):
        path = self.cfg.get("db_path", "")
        if not path or not os.path.exists(path):
            return
        pilot = self._last_of_pilot or self._logged_in_pilot or ""
        poste_co = self._logged_in_poste or ""
        row_evt = [
            "Changement d'OF",                       # A col 1
            "",                                       # B col 2 OF
            start_dt.strftime("%d/%m/%Y"),            # C col 3 Date
            poste_co,                                 # D col 4 Poste
            pilot,                                    # E col 5 pilote du dernier OF
            "", "", "", "", "", "", "", "", "", "", "", # F-P (cols 6-16)
            start_dt.strftime("%H:%M:%S"),            # Q col 17 Heure Début
            end_dt.strftime("%H:%M:%S"),              # R col 18 Heure Fin
            fmt((end_dt - start_dt).total_seconds()), # S col 19 Durée
            "",                                       # T col 20 Commentaire
        ]
        def _bg():
            try:
                with self._excel_lock:
                    wb = self._get_wb(path)
                    if wb is None:
                        return
                    ws = self._ensure_events_sheet(wb)
                    ws.append(row_evt)
                    self._format_row(ws, ws.max_row)
                    self._safe_excel_save(wb, path)
                    self._wb_mtime_cache = os.path.getmtime(path)
            except Exception:
                self._invalidate_wb_cache()
        threading.Thread(target=_bg, daemon=True).start()


    # ── Copie dashboard ───────────────────────────────────────────────────────
    def _dashboard_copy_path(self):
        """Retourne le chemin de la copie dashboard (à côté de l'Excel principal)."""
        path = self.cfg.get("db_path", "")
        if not path:
            return ""
        base, ext = os.path.splitext(path)
        return base + "_dashboard" + (ext or ".xlsx")

    def _copy_excel_for_dashboard(self, source_path=None):
        """Copie le fichier Excel vers la copie dashboard après chaque écriture."""
        import shutil
        if source_path is None:
            source_path = self.cfg.get("db_path", "")
        if not source_path or not os.path.exists(source_path):
            return
        dst = self._dashboard_copy_path()
        if not dst:
            return
        try:
            shutil.copy2(source_path, dst)
        except Exception:
            pass

    def _get_read_path(self):
        """Retourne la copie dashboard si elle existe, sinon le fichier principal."""
        copy = self._dashboard_copy_path()
        if copy and os.path.exists(copy):
            return copy
        return self.cfg.get("db_path", "")


if __name__ == "__main__":
    root = tk.Tk()
    root.configure(bg=BG)
    try:
        root.state("zoomed")
    except Exception:
        root.attributes("-fullscreen", True)
    App(root)
    root.mainloop()
