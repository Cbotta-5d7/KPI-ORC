"""KPI-ORC v6.3 - Flask + pywebview"""
import json, os, sys, datetime, threading, math, shutil, time, atexit, signal
from flask import Flask, request, jsonify, render_template_string
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill

CONFIG_FILE  = os.path.join(os.path.expanduser("~"), "kpi_orc_config.json")
SESSION_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_session.json")
PENDING_FILE = os.path.join(os.path.expanduser("~"), "kpi_orc_pending.json")

EVENTS = [
    ("Pochon / Fibre",       "ratt_pochon",      "ratt"),
    ("Couture",              "ratt_couture",     "ratt"),
    ("Emballage",            "ratt_emb",         "ratt"),
    ("Presse Souder",        "ratt_presse_soud", "ratt"),
    ("Presse ZIP",           "ratt_presse_zip",  "ratt"),
    ("Nettoyage",            "nettoyage",        "nettoyage"),
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
    ("Matiere premiere",     "arret_mp",         "pb"),
    ("Reunion",              "arret_reunion",    "ratt"),
]

# Nouveau schéma unifié - 39 colonnes
DECL_HEADERS = [
    "Type","OF","Date","Poste","Pilote","Co-Pilote","Nb Personnes",
    "Taille","Code Produit","Type Produit","Poids Garnissage","Fibre",
    "OF Taie","Traca Fibre","Ref Taie","Kit",
    "Heure Debut","Heure Fin","Duree",
    "Qte Fabriquee","Qte Emballee","Equivalence","Cadence/h","Cadence/h/pers","TRS%",
    "Qte Init Taie","Nb Taie 2nd Choix","Nb Defaut Couture",
    "Mq Taie","Mq Housse/Encart","Nb PP Cousue",
    "","Manquant MP","Manquant Personnel/Reunion",
    "","Commentaire","Prevu/Hors TRS",
    "Duree Arrets","Duree Prod Pure",
]

POSTES = ["Matin","Midi","Nuit","Jour"]

# Catégories interposte prédéfinies
INTERPOSTE_CATS = [
    ("Changement de série", "changement_serie", "interposte"),
    ("Réglage / Setup machine", "reglage_setup", "interposte"),
    ("Attente matière première", "attente_mp", "interposte"),
    ("Réunion / Formation", "reunion", "interposte"),
    ("Nettoyage interposte", "nettoyage_inter", "interposte"),
    ("Pause pilote interposte", "pause_inter", "interposte"),
    ("Autre (interposte)", "autre_inter", "interposte"),
]

def get_events_list():
    """Retourne la liste des arrêts configurés (Excel col K > cfg > EVENTS défaut)."""
    excel_evts = _lists.get("arrêts_k", [])
    if excel_evts:
        return excel_evts
    custom = cfg.get("events_list", [])
    if custom:
        return custom
    return [{"label": e[0], "key": e[1], "cat": e[2]} for e in EVENTS]

def write_events_to_excel(ev_list):
    """Écrit la liste des arrêts dans l'onglet Listes col K (format: label|cat)."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Listes" not in wb.sheetnames:
                    wb.create_sheet("Listes")
                ws = wb["Listes"]
                ws.cell(1, 11).value = "Arrêts"
                for ri in range(2, ws.max_row + 2):
                    ws.cell(ri, 11).value = None
                for ri, ev in enumerate(ev_list, start=2):
                    ws.cell(ri, 11).value = f"{ev.get('label','')}|{ev.get('cat','pb')}"
                _safe_excel_save(wb, path)
            threading.Thread(target=load_lists, daemon=True).start()
        except: pass
    threading.Thread(target=_bg, daemon=True).start()

def save_events_list(ev_list):
    cfg["events_list"] = ev_list
    save_cfg_data()
    write_events_to_excel(ev_list)

# ── État global ────────────────────────────────────────────────────────────────
_S = {
    "pilot": None, "poste": None,
    "prod_active": False,
    "of_start": None,
    "last_of_end": None,
    "last_of_pilot": "",
    "inter_of_s": 0.0, "interposte_s": 0.0,
    "timers": {},
    "tl_events": [],
    "of_periods": [], "of_changes": [],
    "form": {},
    "is_paused": False, "pause_start": None,
    "pause_total_s": 0.0, "pause_periods": [],
    "of_count_shift": 0,
    "shift_start": None,
}
_excel_lock = threading.Lock()
_lists = {}
_decl_cache = []   # liste de (row_num, row_data) - toutes déclarations (prod + events)
_prod_ref_cached = 0.0
_excel_busy = False  # True quand le fichier Excel est verrouillé (ouvert par Excel)
cfg = {}

flask_app = Flask(__name__)

# ── Utilitaires ────────────────────────────────────────────────────────────────
def fmt(seconds):
    s = max(0, int(seconds or 0))
    return f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}"

def _hms_to_sec(s):
    try:
        if hasattr(s,'hour'): return s.hour*3600+s.minute*60+getattr(s,'second',0)
        parts = str(s).strip().split(":")
        if len(parts)==3: return int(parts[0])*3600+int(parts[1])*60+float(parts[2])
        if len(parts)==2: return int(parts[0])*3600+float(parts[1])*60  # HH:MM
    except: pass
    return 0.0

def _row_date(v):
    if not v: return ""
    s = str(v)
    if len(s)>=10 and s[2]=="/" and s[5]=="/": return s[:10]
    try:
        if hasattr(v,"strftime"): return v.strftime("%d/%m/%Y")
    except: pass
    return s[:10]

def _row_time(v):
    if not v: return ""
    s = str(v)
    if ":" in s: return s[:8]
    try:
        if hasattr(v,"strftime"): return v.strftime("%H:%M:%S")
    except: pass
    return s

def _min_str_to_hms(val):
    try:
        mins = float(str(val).replace(",","."))
        if mins<=0: return ""
        return fmt(mins*60)
    except: return ""

def _n(v):
    try: return float(str(v).replace(",",".")) or 0
    except: return 0

def _dt_str(dt):
    return dt.isoformat() if dt else None

def _str_dt(s):
    if not s: return None
    try: return datetime.datetime.fromisoformat(s)
    except: return None

def load_cfg():
    try:
        with open(CONFIG_FILE) as f:
            d = json.load(f)
            d.setdefault("modeles_horaires", [])
            d.setdefault("pilot_passwords", {})
            return d
    except:
        return {"db_path":"","supervisor_pw":"1234","prod_ref":0,
                "pause_max_min":20,"clean_short_min":10,"clean_long_min":30,
                "clean_grand_min":60,"meeting_tol_min":5,
                "modeles_horaires":[],"pilot_passwords":{}}

def save_cfg_data():
    with open(CONFIG_FILE,"w") as f: json.dump(cfg,f)

def get_prod_ref():
    v = cfg.get("prod_ref",0)
    try:
        fv = float(v)
        if fv>0: return fv
    except: pass
    return _prod_ref_cached

def _get_model_day_cfg(poste, date_obj=None):
    """Retourne (debut_str, fin_str) du modèle horaire pour le poste/jour donné."""
    models = cfg.get("modeles_horaires", [])
    model = next((m for m in models if str(m.get("nom","")).strip().lower()==str(poste or "").strip().lower()), None)
    if not model: return None, None
    jours = model.get("jours", {})
    if date_obj and jours:
        day_map = {0:'lun',1:'mar',2:'mer',3:'jeu',4:'ven',5:'sam',6:'dim'}
        day_cfg = jours.get(day_map.get(date_obj.weekday(),'lun'))
    else:
        day_cfg = next((v for k,v in jours.items() if v and v.get("debut") and v.get("fin")), None) if jours else None
    if not day_cfg:
        return model.get("debut","05:00"), model.get("fin","13:00")
    return day_cfg.get("debut","05:00"), day_cfg.get("fin","13:00")

def get_shift_duration_s(poste, date_obj=None):
    """Durée nominale du poste en secondes selon le modèle horaire."""
    debut_str, fin_str = _get_model_day_cfg(poste, date_obj)
    if debut_str is None: return 28800
    def to_min(t):
        try:
            p=str(t).split(":"); return int(p[0])*60+int(p[1])
        except: return 0
    d=to_min(debut_str); f=to_min(fin_str)
    if f<=d: f+=1440
    return (f-d)*60

# ── Timers ────────────────────────────────────────────────────────────────────
def t_start(key):
    t = _S["timers"].setdefault(key,{"elapsed":0.0,"running":False,"start":None})
    if not t["running"]:
        t["start"] = datetime.datetime.now()
        t["running"] = True
    save_session()

def t_stop(key, end_time=None):
    t = _S["timers"].get(key)
    if t and t["running"]:
        end = end_time or datetime.datetime.now()
        t["elapsed"] += (end - t["start"]).total_seconds()
        t["running"] = False
        t["start"] = None
    save_session()

def t_get(key):
    t = _S["timers"].get(key)
    if not t: return 0.0
    el = t["elapsed"]
    if t["running"] and t["start"]:
        el += (datetime.datetime.now() - t["start"]).total_seconds()
    return el

def t_running(key):
    t = _S["timers"].get(key)
    return t and t["running"]

def t_stop_all():
    for key in list(_S["timers"].keys()):
        t_stop(key)

def t_reset():
    _S["timers"] = {}

def t_wall_clock_stops():
    total = 0.0
    for key, t in _S["timers"].items():
        if key.startswith("_"): continue
        total += t["elapsed"]
        if t["running"] and t["start"]:
            total += (datetime.datetime.now()-t["start"]).total_seconds()
    return total

# ── Timeline events ───────────────────────────────────────────────────────────
def tl_open(key, cat):
    existing = next((e for e in _S["tl_events"] if e["key"]==key and not e.get("end")), None)
    if not existing:
        _S["tl_events"].append({"key":key,"cat":cat,"start":datetime.datetime.now(),"end":None,"comment":"","hors_trs":False})
    save_session()

def tl_close(key, comment="", end_time=None):
    for ev in _S["tl_events"]:
        if ev["key"]==key and not ev.get("end"):
            ev["end"] = end_time or datetime.datetime.now()
            ev["comment"] = comment
            break
    save_session()

def tl_close_all():
    now = datetime.datetime.now()
    for ev in _S["tl_events"]:
        if not ev.get("end"):
            ev["end"] = now
    save_session()

def serialize_event(ev):
    return {
        "key": ev["key"], "cat": ev["cat"],
        "start": _dt_str(ev["start"]),
        "end": _dt_str(ev.get("end")),
        "comment": ev.get("comment",""),
        "nettoyage_type": ev.get("nettoyage_type",""),
        "hors_trs": ev.get("hors_trs",False),
    }

# ── Session ───────────────────────────────────────────────────────────────────
def save_session():
    try:
        timers_s = {}
        for k,t in _S["timers"].items():
            timers_s[k] = {
                "elapsed": t["elapsed"],
                "running": t["running"],
                "start": _dt_str(t["start"]),
            }
        d = {
            "pilot": _S["pilot"],
            "poste": _S["poste"],
            "prod_active": _S["prod_active"],
            "of_start": _dt_str(_S["of_start"]),
            "last_of_end": _dt_str(_S["last_of_end"]),
            "last_of_pilot": _S["last_of_pilot"],
            "inter_of_s": _S["inter_of_s"],
            "interposte_s": _S["interposte_s"],
            "timers": timers_s,
            "tl_events": [serialize_event(e) for e in _S["tl_events"]],
            "is_paused": _S["is_paused"],
            "pause_start": _dt_str(_S["pause_start"]),
            "pause_total_s": _S["pause_total_s"],
            "pause_periods": [[_dt_str(a),_dt_str(b)] for a,b in _S["pause_periods"]],
            "of_count_shift": _S["of_count_shift"],
            "form": _S["form"],
            "shift_start": _dt_str(_S.get("shift_start")),
        }
        with open(SESSION_FILE,"w",encoding="utf-8") as f: json.dump(d,f,default=str)
    except: pass

def load_session():
    try:
        with open(SESSION_FILE,encoding="utf-8") as f: d = json.load(f)
        _S["pilot"]         = d.get("pilot")
        _S["poste"]         = d.get("poste")
        _S["prod_active"]   = d.get("prod_active",False)
        _S["of_start"]      = _str_dt(d.get("of_start"))
        _S["last_of_end"]   = _str_dt(d.get("last_of_end"))
        _S["last_of_pilot"] = d.get("last_of_pilot","")
        _S["inter_of_s"]    = float(d.get("inter_of_s",0))
        _S["interposte_s"]  = float(d.get("interposte_s",0))
        _S["is_paused"]     = d.get("is_paused",False)
        _S["pause_start"]   = _str_dt(d.get("pause_start"))
        _S["pause_total_s"] = float(d.get("pause_total_s",0))
        _S["pause_periods"] = [[_str_dt(a),_str_dt(b)] for a,b in d.get("pause_periods",[])]
        _S["of_count_shift"]= d.get("of_count_shift",0)
        _S["form"]          = d.get("form",{})
        _S["shift_start"]   = _str_dt(d.get("shift_start"))
        raw_timers = d.get("timers",{})
        _S["timers"] = {}
        for k,t in raw_timers.items():
            _S["timers"][k] = {
                "elapsed": float(t.get("elapsed",0)),
                "running": t.get("running",False),
                "start": _str_dt(t.get("start")),
            }
        raw_tl = d.get("tl_events",[])
        _S["tl_events"] = []
        for ev in raw_tl:
            _S["tl_events"].append({
                "key": ev["key"], "cat": ev["cat"],
                "start": _str_dt(ev["start"]),
                "end": _str_dt(ev.get("end")),
                "comment": ev.get("comment",""),
                "nettoyage_type": ev.get("nettoyage_type",""),
                "hors_trs": ev.get("hors_trs",False),
            })
        return True
    except: return False

# Save session on any exit (window close, kill, etc.)
def _on_exit(*args):
    try: save_session()
    except: pass

atexit.register(_on_exit)
try: signal.signal(signal.SIGTERM, _on_exit)
except: pass
try: signal.signal(signal.SIGINT, _on_exit)
except: pass

# ── Listes depuis Excel ────────────────────────────────────────────────────────
def load_lists():
    global _lists, _prod_ref_cached
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if "Listes" in wb.sheetnames:
            ws = wb["Listes"]
            headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
            for ci, h in enumerate(headers, start=1):
                if not h: continue
                vals = []
                for ri in range(2, ws.max_row+1):
                    v = ws.cell(ri,ci).value
                    if v is not None and str(v).strip(): vals.append(str(v).strip())
                _lists[str(h).strip()] = vals
            # Production ref from list
            pr_vals = _lists.get("Prod ref 8h",[]) or _lists.get("prod_ref",[])
            if pr_vals:
                try: _prod_ref_cached = float(str(pr_vals[0]).replace(",","."))
                except: pass
            # Load pilot passwords from col A (pilote names) + col B (MDP) — Excel is source of truth
            pil_map = {}
            for ri in range(2, ws.max_row+1):
                pil_v = ws.cell(ri, 1).value
                pw_v = ws.cell(ri, 2).value
                if pil_v and str(pil_v).strip():
                    pil_map[str(pil_v).strip()] = str(pw_v or "").strip()
            if pil_map:
                cfg["pilot_passwords"] = pil_map
            # Also read fixed-position columns: C=copilotes, D=tailles, E=types_prod, F=equiv_coef, J=fibres
            for col_idx, list_key in [(3,"copilotes"),(4,"tailles_col"),(5,"types_prod_col"),(6,"equivalences_col"),(10,"fibres_col")]:
                vals = []
                for ri in range(2, ws.max_row+1):
                    v = ws.cell(ri, col_idx).value
                    if v is not None and str(v).strip():
                        vals.append(str(v).strip())
                if vals:
                    _lists[list_key] = vals
            # Col K (11): liste des arrêts configurables (format: "label|cat")
            evts_k = []
            for ri in range(2, ws.max_row+1):
                v = ws.cell(ri, 11).value
                if v is not None and str(v).strip():
                    parts = str(v).strip().split("|")
                    lbl = parts[0].strip()
                    cat = parts[1].strip() if len(parts) > 1 else "pb"
                    if lbl:
                        key = lbl.lower().replace(" ","_").replace("/","_").replace("é","e").replace("è","e").replace("ê","e").replace("à","a").replace("ç","c")[:28]
                        evts_k.append({"label": lbl, "key": key, "cat": cat})
            if evts_k:
                _lists["arrêts_k"] = evts_k
                cfg["events_list"] = evts_k
                save_cfg_data()
        wb.close()
    except: pass

def get_list(h):
    return _lists.get(h,[])

def load_history():
    global _decl_cache, _excel_busy
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        _excel_busy = False
        _decl_cache = []
        # Nouveau schéma unifié
        if "Declarations" in wb.sheetnames:
            ws = wb["Declarations"]
            for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if r and any(r):
                    _decl_cache.append((i, list(r)+[None]*5))
        # Rétro-compat: lire Data + Evenements si Declarations absent
        elif "Data" in wb.sheetnames:
            ws = wb["Data"]
            for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if r and any(r):
                    row = list(r)+[None]*5
                    # Convertir en format unifié (mettre "Production" en pos 0)
                    unified = [None]*37
                    unified[0] = "Production"
                    unified[1] = row[0]   # OF
                    unified[2] = row[1]   # Date
                    unified[3] = row[2]   # Poste
                    unified[4] = row[3]   # Pilote
                    unified[5] = row[4]   # Co-Pilote
                    unified[6] = row[5]   # Nb Personnes
                    unified[7] = row[6]   # Taille
                    unified[8] = row[7]   # Code Produit
                    unified[9] = row[8]   # Type Produit
                    unified[10] = row[9]  # Poids
                    unified[11] = row[10] # Fibre
                    unified[12] = row[11] # OF Taie
                    unified[13] = row[12] # Traca
                    unified[14] = row[22] # Ref Taie (col 23 dans Data)
                    unified[15] = row[21] # Kit (col 22 dans Data)
                    unified[16] = row[17] # Heure Debut
                    unified[17] = row[18] # Heure Fin
                    unified[18] = row[16] # Duree
                    unified[19] = row[13] # Qte Fab
                    unified[20] = row[14] # Qte Emb
                    unified[21] = row[15] # Equiv
                    unified[22] = row[19] # Cadence/h
                    unified[23] = row[20] # Cadence/h/pers
                    unified[35] = row[57] if len(row)>57 else None  # Commentaire
                    _decl_cache.append((i, unified))
        wb.close()
    except PermissionError:
        _excel_busy = True
        def _retry():
            import time as _t; _t.sleep(5)
            load_history()
        threading.Thread(target=_retry, daemon=True).start()
    except: pass

# ── Calcul équivalence ────────────────────────────────────────────────────────
def calc_equiv(qte, taille, type_prod):
    types  = get_list("types_prod_col") or get_list("Type produit") or get_list("types_prod")
    equivs = get_list("equivalences_col") or get_list("Equivalence coef") or get_list("Equivalence")
    if type_prod and types and equivs:
        for i,t in enumerate(types):
            if str(t).strip().lower()==str(type_prod).strip().lower():
                if i<len(equivs):
                    try: return round(qte*float(str(equivs[i]).replace(",",".")),2)
                    except: pass
    return round(float(qte or 0),2)

# ── Excel helpers ─────────────────────────────────────────────────────────────
def _get_wb(path):
    if not path or not os.path.exists(path):
        return None
    try: return load_workbook(path)
    except: return None

def _safe_excel_save(wb, path):
    bak = path+".bak"
    try: shutil.copy2(path,bak)
    except: pass
    wb.save(path)
    try: os.remove(bak)
    except: pass

def _format_row(ws, row_num):
    thin = Side(style="thin")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)
    for cell in ws[row_num]:
        cell.border = border
        cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)

def _ensure_decl_sheet(wb):
    if "Declarations" not in wb.sheetnames:
        ws = wb.create_sheet("Declarations",0)
        for i,h in enumerate(DECL_HEADERS,start=1): ws.cell(1,i).value=h
        _format_row(ws,1)
        # Couleur header
        fill = PatternFill("solid", fgColor="1a1f5e")
        from openpyxl.styles import Font
        for cell in ws[1]:
            cell.fill = fill
            cell.font = Font(color="FFFFFF", bold=True, size=10)
    else:
        ws = wb["Declarations"]
        for i,h in enumerate(DECL_HEADERS,start=1):
            if ws.cell(1,i).value is None: ws.cell(1,i).value=h
    return wb["Declarations"]

def build_decl_rows(v, tl_events, of_start, pause_periods):
    """Construit les lignes arrêts/pauses au format unifié (37 cols)."""
    rows = []
    kit_val = "Oui" if v.get("kit") else "Non"
    def _base_row(type_decl, start, end, comment="", hors_trs=""):
        dur = max(0,(end-start).total_seconds())
        return [
            type_decl,                          # 1 Type
            v.get("of_num",""),                 # 2 OF
            start.strftime("%d/%m/%Y"),          # 3 Date
            v.get("poste",""),                  # 4 Poste
            v.get("pilote",""),                 # 5 Pilote
            v.get("copilote",""),               # 6 Co-Pilote
            v.get("nb_pers",""),                # 7 Nb Personnes
            v.get("taille",""),                 # 8 Taille
            v.get("code_prod",""),              # 9 Code Produit
            v.get("type_prod",""),              # 10 Type Produit
            v.get("poids",""),                  # 11 Poids Garnissage
            v.get("fibre",""),                  # 12 Fibre
            v.get("of_taie",""),                # 13 OF Taie
            v.get("traca",""),                  # 14 Traca Fibre
            v.get("ref_taie",""),               # 15 Ref Taie
            kit_val,                            # 16 Kit
            start.strftime("%H:%M:%S"),          # 17 Heure Debut
            end.strftime("%H:%M:%S"),            # 18 Heure Fin
            fmt(dur),                           # 19 Duree
            "","","","","",                     # 20-24 prod only
            "",                                 # 25 TRS%
            "","","","","","",                  # 26-31 prod only
            "","","",                           # 32-34
            "",                                 # 35 Nettoyage
            comment,                            # 36 Commentaire
            hors_trs,                           # 37 Prevu/Hors TRS
        ]
    for ev in tl_events:
        if ev.get("cat") not in ("ratt","pb","nettoyage","autre"): continue
        if not ev.get("key") or ev["key"].startswith("_"): continue
        if of_start and ev["start"] < of_start and ev.get("key")!="arret_interposte": continue
        start = ev["start"]
        end = ev.get("end") or datetime.datetime.now()
        if ev["key"]=="nettoyage":
            ntype = ev.get("nettoyage_type","court")
            label = {"court":"Nettoyage court","long":"Nettoyage long","grand":"Grand nettoyage"}.get(ntype,"Nettoyage court")
        elif ev["cat"]=="autre":
            label = ev["key"]  # Custom stop name typed by user
        else:
            cat_name = "Rattrapage" if ev["cat"]=="ratt" else "PB Technique"
            lbl = next((e[0] for e in EVENTS if e[1]==ev["key"]),ev["key"])
            label = f"{cat_name}: {lbl}"
        rows.append(_base_row(label, start, end, ev.get("comment",""), "OUI" if ev.get("hors_trs") else ""))
    for ps, pe in pause_periods:
        if of_start and ps < of_start: continue
        rows.append(_base_row("Pause", ps, pe))
    return rows

def write_excel_bg(prod_row, evt_rows):
    """Écrit la ligne production + lignes arrêts dans la feuille Declarations."""
    path = cfg.get("db_path","")
    if not path: return
    try:
        with open(PENDING_FILE,"w",encoding="utf-8") as f:
            json.dump({"db_path":path,"prod_row":prod_row,"evt_rows":evt_rows},f,ensure_ascii=False,default=str)
    except: pass
    def _bg():
        for attempt in range(15):
            try:
                with _excel_lock:
                    wb = _get_wb(path)
                    if wb is None:
                        time.sleep(4)
                        continue
                    ws = _ensure_decl_sheet(wb)
                    ws.append(prod_row)
                    _format_row(ws, ws.max_row)
                    for er in evt_rows:
                        ws.append(er)
                        _format_row(ws, ws.max_row)
                    _safe_excel_save(wb, path)
                    try: os.remove(PENDING_FILE)
                    except: pass
                    break  # success — exit retry loop
            except PermissionError:
                # Excel a le fichier ouvert — on réessaie dans 5s
                time.sleep(5)
            except Exception:
                break
        threading.Thread(target=load_history, daemon=True).start()
    threading.Thread(target=_bg, daemon=True).start()

def write_changement_of(start_dt, end_dt, label=None, comment=""):
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    pilot = _S.get("last_of_pilot") or _S.get("pilot") or ""
    dur_s = (end_dt-start_dt).total_seconds()
    row_type = label or "Changement d'OF"
    row = [
        row_type,"",start_dt.strftime("%d/%m/%Y"),
        _S.get("poste",""),pilot,"","","","","","","","","","","",
        start_dt.strftime("%H:%M:%S"),end_dt.strftime("%H:%M:%S"),fmt(dur_s),
        "","","","","","","","","","","","","","","","",comment,
    ]
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _ensure_decl_sheet(wb)
                ws.append(row)
                _format_row(ws,ws.max_row)
                _safe_excel_save(wb,path)
            threading.Thread(target=load_history,daemon=True).start()
        except: pass
    threading.Thread(target=_bg,daemon=True).start()

POSTES_HEADERS = ["Date","Pilote","Co-Pilote","Poste","Nb OF","Prod Total (pièces)","Prod Totale (equiv)","TRS Poste %","Total Arrets (min)","Total Pauses (min)","Nettoyage (min)","Durée Prod Totale (min)","Durée Prod Sans Arrêt (min)","Durée poste théorique (min)","Commentaire"]

def write_pilots_to_excel(pilot_passwords):
    """Écrit la liste pilote+MDP dans l'onglet Listes col A+B."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return False
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Listes" not in wb.sheetnames:
                    ws = wb.create_sheet("Listes")
                    ws.cell(1,1).value = "Pilotes"
                    ws.cell(1,2).value = "MDP"
                else:
                    ws = wb["Listes"]
                    if not ws.cell(1,1).value: ws.cell(1,1).value = "Pilotes"
                    if not ws.cell(1,2).value: ws.cell(1,2).value = "MDP"
                # Clear existing pilot rows
                for ri in range(2, ws.max_row+2):
                    ws.cell(ri,1).value = None
                    ws.cell(ri,2).value = None
                # Write new pilot data
                for ri,(p,pw) in enumerate(pilot_passwords.items(),start=2):
                    ws.cell(ri,1).value = p
                    ws.cell(ri,2).value = pw
                _safe_excel_save(wb,path)
        except: pass
        threading.Thread(target=load_lists,daemon=True).start()
    threading.Thread(target=_bg,daemon=True).start()
    return True

def write_poste_row(data):
    """Écrit une ligne dans l'onglet Postes à la fin de chaque poste."""
    path = cfg.get("db_path","")
    if not path: return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Postes" not in wb.sheetnames:
                    ws = wb.create_sheet("Postes")
                    for i,h in enumerate(POSTES_HEADERS,start=1): ws.cell(1,i).value=h
                    _format_row(ws,1)
                    from openpyxl.styles import PatternFill, Font
                    fill=PatternFill("solid",fgColor="1a1f5e")
                    for cell in ws[1]:
                        cell.fill=fill
                        cell.font=Font(color="FFFFFF",bold=True,size=10)
                else:
                    ws = wb["Postes"]
                row = [
                    data.get("date",""),
                    data.get("pilot",""),
                    data.get("copilote",""),
                    data.get("poste",""),
                    data.get("nb_of",0),
                    round(float(data.get("prod_total",0) or 0),0),
                    round(float(data.get("tot_equiv",0) or 0),1),
                    data.get("trs_shift",""),
                    round(float(data.get("arret_min",0) or 0),1),
                    round(float(data.get("pause_min",0) or 0),1),
                    round(float(data.get("nett_min",0) or 0),1),
                    round(float(data.get("dur_prod_total_min",0) or 0),1),
                    round(float(data.get("dur_prod_sans_arret_min",0) or 0),1),
                    round(float(data.get("dur_poste_theorique_min",0) or 0),1),
                    data.get("comment",""),
                ]
                ws.append(row)
                _format_row(ws,ws.max_row)
                _safe_excel_save(wb,path)
        except: pass
    threading.Thread(target=_bg,daemon=True).start()

def _start_periodic_excel_sync():
    """Recharge les listes Excel toutes les 5 minutes pour éviter la perte de données."""
    def _loop():
        while True:
            time.sleep(300)  # 5 minutes
            try: load_lists()
            except: pass
    t = threading.Thread(target=_loop, daemon=True)
    t.start()

def toggle_hors_trs_excel(row_num, new_val):
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Declarations" in wb.sheetnames:
                    ws = wb["Declarations"]
                    ws.cell(row_num, 37).value = new_val  # col 37 = Prevu/Hors TRS
                    _safe_excel_save(wb,path)
        except: pass
    threading.Thread(target=_bg,daemon=True).start()

# ── Flask helpers ─────────────────────────────────────────────────────────────
def _check_pw(pw):
    return str(pw or "") == str(cfg.get("supervisor_pw","1234"))

def _state_json():
    active_stops = [k for k,t in _S["timers"].items() if t.get("running") and not k.startswith("_")]
    stop_wall_s = t_wall_clock_stops()
    of_elapsed = 0.0
    if _S["of_start"]:
        of_elapsed = (datetime.datetime.now()-_S["of_start"]).total_seconds()
        if _S["is_paused"] and _S["pause_start"]:
            of_elapsed -= (datetime.datetime.now()-_S["pause_start"]).total_seconds()
    timers_out = {}
    for k,t in _S["timers"].items():
        el = t["elapsed"]
        if t["running"] and t["start"]:
            el += (datetime.datetime.now()-t["start"]).total_seconds()
        timers_out[k] = {"elapsed":round(el,1),"running":t["running"]}
    return {
        "pilot": _S["pilot"],
        "poste": _S["poste"],
        "prod_active": _S["prod_active"],
        "of_start_iso": _dt_str(_S["of_start"]),
        "of_elapsed_s": round(of_elapsed,1),
        "inter_of_s": _S["inter_of_s"],
        "interposte_s": _S["interposte_s"],
        "is_paused": _S["is_paused"],
        "pause_start_iso": _dt_str(_S["pause_start"]),
        "pause_total_s": round(_S["pause_total_s"],1),
        "active_stops": active_stops,
        "stop_wall_s": round(stop_wall_s,1),
        "timers": timers_out,
        "form": _S["form"],
        "prod_ref": get_prod_ref(),
        "of_count_shift": _S["of_count_shift"],
        "tl_events": [serialize_event(e) for e in _S["tl_events"]],
        "shift_duration_s": get_shift_duration_s(_S["poste"]),
        "shift_start_iso": _dt_str(_S.get("shift_start")),
        "excel_busy": _excel_busy,
    }

@flask_app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@flask_app.route('/reset')
def page_reset():
    """Page de reset d'urgence — accessible directement dans le navigateur."""
    _S["prod_active"] = False
    _S["of_start"] = None
    _S["tl_events"] = []
    _S["timers"] = {}
    _S["is_paused"] = False
    _S["pause_start"] = None
    _S["pause_total_s"] = 0.0
    _S["pause_periods"] = []
    _S["inter_of_s"] = 0.0
    _S["form"] = {}
    save_session()
    return """<!DOCTYPE html><html><head><meta charset="utf-8"><title>KPI-ORC Reset</title>
    <style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f8fafc}
    .box{text-align:center;padding:40px;background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.1)}
    h2{color:#16a34a;margin-bottom:10px}p{color:#64748b;margin-bottom:20px}
    a{background:#1e3a5f;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:700}</style></head>
    <body><div class="box"><h2>✓ Production annulée</h2>
    <p>La prod bloquée a été réinitialisée.<br>Vous pouvez maintenant démarrer une nouvelle production.</p>
    <a href="/">Retour à l'application</a></div></body></html>"""

@flask_app.route('/api/state')
def api_state():
    return jsonify(_state_json())

@flask_app.route('/api/lists')
def api_lists():
    return jsonify({
        "pilotes": get_list("Pilotes") or get_list("pilotes") or get_list("Pilote") or get_list("pilote") or list(cfg.get("pilot_passwords",{}).keys()),
        "copilotes": get_list("copilotes") or get_list("Co-Pilote") or get_list("Copilote") or get_list("Pilotes") or get_list("pilotes") or list(cfg.get("pilot_passwords",{}).keys()),
        "tailles": get_list("tailles_col") or get_list("Tailles") or get_list("taille"),
        "types_prod": get_list("types_prod_col") or get_list("Type produit") or get_list("types_prod"),
        "fibres": get_list("fibres_col") or get_list("Fibres") or get_list("fibre"),
        "tracas": get_list("Traca") or get_list("tracas"),
        "equivalences": get_list("equivalences_col") or get_list("Equivalence coef") or get_list("Equivalence") or [],
    })

@flask_app.route('/api/login', methods=['POST'])
def api_login():
    data = request.json or {}
    pilot = data.get("pilot","").strip()
    poste = data.get("poste","").strip()
    pw = str(data.get("pw",""))
    if not pilot or not poste:
        return jsonify({"ok":False,"error":"Pilote et poste requis"}),400
    # Check pilot password if configured
    pilot_pws = cfg.get("pilot_passwords",{})
    if pilot in pilot_pws and pilot_pws[pilot]:
        if pw != str(pilot_pws[pilot]):
            return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    _S["pilot"] = pilot
    _S["poste"] = poste
    if not _S.get("shift_start"):
        _S["shift_start"] = datetime.datetime.now()
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/logout', methods=['POST'])
def api_logout():
    if _S["prod_active"]:
        return jsonify({"ok":False,"error":"Production en cours"}),400
    _S["pilot"] = None
    _S["poste"] = None
    _S["shift_start"] = None
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/force_reset_prod', methods=['POST'])
def api_force_reset_prod():
    """Reset d'urgence : annule la prod en cours sans écrire dans Excel."""
    _S["prod_active"] = False
    _S["of_start"] = None
    _S["tl_events"] = []
    _S["timers"] = {}
    _S["is_paused"] = False
    _S["pause_start"] = None
    _S["pause_total_s"] = 0.0
    _S["pause_periods"] = []
    _S["inter_of_s"] = 0.0
    _S["form"] = {}
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/start_prod', methods=['POST'])
def api_start_prod():
    if not _S["pilot"]:
        return jsonify({"ok":False,"error":"Connectez-vous d'abord"}),400
    if _S["prod_active"]:
        return jsonify({"ok":False,"error":"Production déjà en cours"}),400
    now = datetime.datetime.now()
    gap_s = 0.0
    if _S["last_of_end"]:
        gap_s = (now - _S["last_of_end"]).total_seconds()
        _S["interposte_s"] = gap_s
    is_first_of = (_S["of_count_shift"] == 0)
    _S["prod_active"] = True
    _S["of_start"] = now
    _S["inter_of_s"] = 0.0
    _S["pause_total_s"] = 0.0
    _S["pause_periods"] = []
    _S["tl_events"] = []
    _S["of_count_shift"] += 1
    _S["form"] = {"of_num": ""}
    if not _S.get("shift_start"):
        _S["shift_start"] = now
    t_reset()
    save_session()
    # Calcul du gap pré-poste (1er OF vs heure début modèle horaire)
    pre_shift_gap_s = 0.0
    shift_model_start_str = ""
    shift_model_start_iso = ""
    if is_first_of:
        poste = _S.get("poste","")
        day_keys = ["lun","mar","mer","jeu","ven","sam","dim"]
        dk = day_keys[now.weekday()]
        for m in cfg.get("modeles_horaires",[]):
            if m.get("nom","") == poste:
                jour = m.get("jours",{}).get(dk,{})
                debut_str = jour.get("debut","")
                if debut_str:
                    try:
                        h, mi = map(int, debut_str.split(":"))
                        shift_deb = now.replace(hour=h, minute=mi, second=0, microsecond=0)
                        # Si shift commence la veille (poste de nuit), on ajuste
                        if shift_deb > now:
                            shift_deb -= datetime.timedelta(days=1)
                        diff = (now - shift_deb).total_seconds()
                        if 120 < diff < 7200:  # entre 2 min et 2h de retard
                            pre_shift_gap_s = diff
                            shift_model_start_str = debut_str
                            shift_model_start_iso = shift_deb.isoformat()
                    except: pass
                break
    return jsonify({"ok":True,"gap_s":round(gap_s,0),
                    "pre_shift_gap_s":round(pre_shift_gap_s,0),
                    "shift_model_start":shift_model_start_str,
                    "shift_model_start_iso":shift_model_start_iso})

@flask_app.route('/api/set_of_start', methods=['POST'])
def api_set_of_start():
    """Rétrodate le début de l'OF en cours (et shift_start) à l'heure du modèle horaire."""
    data = request.json or {}
    iso = data.get("iso","")
    if not iso or not _S["prod_active"]:
        return jsonify({"ok":False,"error":"Pas de production active"}),400
    try:
        dt = datetime.datetime.fromisoformat(iso)
        _S["of_start"] = dt
        _S["shift_start"] = dt
        save_session()
        return jsonify({"ok":True})
    except Exception as e:
        return jsonify({"ok":False,"error":str(e)}),400

@flask_app.route('/api/inter_of_confirm', methods=['POST'])
def api_inter_of_confirm():
    data = request.json or {}
    _S["inter_of_s"] = float(data.get("inter_of_s",0))
    label = data.get("label","")
    comment = data.get("comment","")
    if _S["inter_of_s"] > 30 and _S["last_of_end"] and _S["of_start"]:
        write_changement_of(_S["last_of_end"], _S["of_start"], label=label or None, comment=comment)
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/events_cfg', methods=['GET'])
def api_events_cfg_get():
    return jsonify({"ok":True,"events":get_events_list()})

@flask_app.route('/api/events_cfg', methods=['POST'])
def api_events_cfg_post():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    ev_list = data.get("events",[])
    save_events_list(ev_list)
    return jsonify({"ok":True})

@flask_app.route('/api/interposte_cfg', methods=['GET'])
def api_interposte_cfg_get():
    default = [c[0] for c in INTERPOSTE_CATS]
    labels = cfg.get("interposte_labels", default)
    return jsonify({"ok":True,"labels":labels})

@flask_app.route('/api/interposte_cfg', methods=['POST'])
def api_interposte_cfg_post():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    labels = data.get("labels",[])
    cfg["interposte_labels"] = [str(l) for l in labels if str(l).strip()]
    save_cfg_data()
    return jsonify({"ok":True})

@flask_app.route('/api/end_prod', methods=['POST'])
def api_end_prod():
    if not _S["prod_active"] or not _S["of_start"]:
        return jsonify({"ok":False,"error":"Pas de production active"}),400
    data = request.json or {}
    v = data.get("form",{})
    if _S["is_paused"]:
        _toggle_pause_internal()
    t_stop_all()
    tl_close_all()
    end_dt = datetime.datetime.now()
    of_s_brut = (end_dt-_S["of_start"]).total_seconds()
    pause_max_s = int(cfg.get("pause_max_min",20))*60
    of_s = max(1, of_s_brut + _S["inter_of_s"] - min(_S["pause_total_s"],pause_max_s))
    stop_s = t_wall_clock_stops()
    qte_fab = _n(v.get("qte_fab",0))
    nb_pers = max(1,_n(v.get("nb_pers",1)) or 1)
    of_hrs = of_s/3600
    equiv = calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
    c1 = round(equiv/of_hrs,2) if of_hrs>0 else 0
    c2 = round(equiv/(nb_pers*of_hrs),2) if of_hrs>0 else 0
    _nett_s = sum(
        (ev["end"]-ev["start"]).total_seconds()
        for ev in _S["tl_events"]
        if ev.get("key")=="nettoyage" and ev.get("start") and ev.get("end")
    )
    prod_ref = get_prod_ref()
    trs = -1.0
    trs_str = ""
    if prod_ref>0 and of_s_brut>0:
        trs = round(equiv/(prod_ref*of_s_brut/28800)*100,1)
        trs_str = str(trs)

    # Ligne Production (37 cols, format unifié)
    prod_row = [
        "Production",
        v.get("of_num",""),
        datetime.date.today().strftime("%d/%m/%Y"),
        v.get("poste",_S["poste"] or ""),
        v.get("pilote",_S["pilot"] or ""),
        v.get("copilote",""),
        v.get("nb_pers",""),
        v.get("taille",""),
        v.get("code_prod",""),
        v.get("type_prod",""),
        v.get("poids",""),
        v.get("fibre",""),
        v.get("of_taie",""),
        v.get("traca",""),
        v.get("ref_taie",""),
        "Oui" if v.get("kit") else "Non",
        _S["of_start"].strftime("%H:%M:%S"),
        end_dt.strftime("%H:%M:%S"),
        fmt(of_s),
        qte_fab,
        _n(v.get("qte_emb",0)),
        equiv,
        c1,
        c2,
        trs_str,
        _n(v.get("qte_init_taie",0)),
        _n(v.get("nb_taie2_choix",0)),
        _n(v.get("nb_def_cout",0)),
        _n(v.get("mq_taie",0)),
        _n(v.get("mq_housse_encart",0)),
        _n(v.get("nb_pp_cousue",0)),
        "",
        _min_str_to_hms(v.get("duree_mq_mp","")),
        _min_str_to_hms(v.get("manquant_pers","")),
        "",
        v.get("comment",""),
        "",
        fmt(stop_s),
        fmt(max(0, of_s_brut - stop_s)),
    ]
    evt_rows = build_decl_rows(
        dict(v, pilote=v.get("pilote",_S["pilot"] or ""), poste=v.get("poste",_S["poste"] or "")),
        _S["tl_events"], _S["of_start"], _S["pause_periods"]
    )
    recap = {
        "of_num": v.get("of_num",""),
        "pilote": v.get("pilote",_S["pilot"] or ""),
        "poste": v.get("poste",_S["poste"] or ""),
        "taille": v.get("taille",""),
        "type_prod": v.get("type_prod",""),
        "qte_fab": qte_fab,
        "qte_emb": _n(v.get("qte_emb",0)),
        "equiv": equiv,
        "of_s": round(of_s,0),
        "of_s_brut": round(of_s_brut,0),
        "pause_s": round(_S["pause_total_s"],0),
        "inter_of_s": round(_S["inter_of_s"],0),
        "stop_wall_s": round(stop_s,0),
        "trs": trs,
        "debut": _S["of_start"].strftime("%H:%M:%S"),
        "fin": end_dt.strftime("%H:%M:%S"),
        "date": datetime.date.today().strftime("%d/%m/%Y"),
        "c1": c1, "c2": c2,
        "nett_s": round(_nett_s,0),
        "interposte_s": round(_S["interposte_s"],0),
    }
    _S["last_of_end"] = end_dt
    _S["last_of_pilot"] = _S["pilot"] or ""
    _S["prod_active"] = False
    _S["of_start"] = None
    _S["of_count_shift"] += 1
    if _S["of_periods"]:
        _S["of_periods"][-1]["end"] = end_dt
        _S["of_periods"][-1]["of_num"] = v.get("of_num","")
    t_reset()
    _S["inter_of_s"] = 0.0
    _S["interposte_s"] = 0.0
    _S["pause_total_s"] = 0.0
    _S["pause_periods"] = []
    _S["form"] = {}
    save_session()
    write_excel_bg(prod_row, evt_rows)
    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok":True,"recap":recap})

@flask_app.route('/api/preview_end_prod', methods=['POST'])
def api_preview_end_prod():
    if not _S["prod_active"] or not _S["of_start"]:
        return jsonify({"ok":False}),400
    data = request.json or {}
    v = data.get("form",{})
    now = datetime.datetime.now()
    of_s_brut = (now-_S["of_start"]).total_seconds()
    pause_max_s = int(cfg.get("pause_max_min",20))*60
    of_s = max(1, of_s_brut+_S["inter_of_s"]-min(_S["pause_total_s"],pause_max_s))
    stop_s = t_wall_clock_stops()
    qte_fab = _n(v.get("qte_fab",0))
    nb_pers = max(1,_n(v.get("nb_pers",1)) or 1)
    equiv = calc_equiv(qte_fab, v.get("taille",""), v.get("type_prod",""))
    of_hrs = of_s/3600
    c1 = round(equiv/of_hrs,2) if of_hrs>0 else 0
    c2 = round(equiv/(nb_pers*of_hrs),2) if of_hrs>0 else 0
    prod_ref = get_prod_ref()
    trs=-1.0
    if prod_ref>0 and of_s_brut>0:
        trs=round(equiv/(prod_ref*of_s_brut/28800)*100,1)
    return jsonify({
        "ok":True,
        "of_s":round(of_s,0),"of_s_brut":round(of_s_brut,0),
        "stop_s":round(stop_s,0),"prod_s":round(max(0,of_s-stop_s),0),
        "pause_s":round(_S["pause_total_s"],0),
        "equiv":equiv,"c1":c1,"c2":c2,"trs":trs,
        "qte_fab":qte_fab,"qte_emb":_n(v.get("qte_emb",0)),
        "tl_events":[serialize_event(e) for e in _S["tl_events"]],
        "debut":_S["of_start"].strftime("%H:%M:%S"),
        "now_str":now.strftime("%H:%M:%S"),
        "of_num":v.get("of_num",""),
    })

@flask_app.route('/api/start_stop', methods=['POST'])
def api_start_stop():
    data = request.json or {}
    key = data.get("key","")
    cat = data.get("cat","pb")
    if not key: return jsonify({"ok":False,"error":"Clé manquante"}),400
    t_start(key)
    tl_open(key,cat)
    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/end_stop', methods=['POST'])
def api_end_stop():
    data = request.json or {}
    key = data.get("key","")
    comment = data.get("comment","")
    t_stop(key)
    tl_close(key,comment)
    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok":True})

def _toggle_pause_internal():
    now = datetime.datetime.now()
    if not _S["is_paused"]:
        _S["is_paused"] = True
        _S["pause_start"] = now
    else:
        if _S["pause_start"]:
            dur = (now-_S["pause_start"]).total_seconds()
            _S["pause_total_s"] += dur
            _S["pause_periods"].append((_S["pause_start"],now))
        _S["is_paused"] = False
        _S["pause_start"] = None
    save_session()

@flask_app.route('/api/toggle_pause', methods=['POST'])
def api_toggle_pause():
    _toggle_pause_internal()
    return jsonify({"ok":True,"paused":_S["is_paused"]})

@flask_app.route('/api/start_nettoyage', methods=['POST'])
def api_start_nettoyage():
    data = request.json or {}
    ntype = data.get("ntype","court")
    t_start("nettoyage")
    tl_open("nettoyage","nettoyage")
    for ev in reversed(_S["tl_events"]):
        if ev["key"]=="nettoyage" and not ev.get("end"):
            ev["nettoyage_type"] = ntype
            break
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/end_nettoyage', methods=['POST'])
def api_end_nettoyage():
    data = request.json or {}
    t_stop("nettoyage")
    tl_close("nettoyage",data.get("comment",""))
    return jsonify({"ok":True})

@flask_app.route('/api/save_form', methods=['POST'])
def api_save_form():
    data = request.json or {}
    _S["form"] = data
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/history')
def api_history():
    rows = []
    date_from = request.args.get("from","")
    date_to = request.args.get("to","")
    date_single = request.args.get("date","")
    # Legacy single-date param
    if date_single and not date_from:
        date_from = date_to = date_single
    def _parse_date(s):
        try:
            if "-" in s: return datetime.datetime.strptime(s,"%Y-%m-%d").date()
            if "/" in s: return datetime.datetime.strptime(s,"%d/%m/%Y").date()
        except: pass
        return None
    d_from = _parse_date(date_from) if date_from else None
    d_to = _parse_date(date_to) if date_to else None
    all_rows = [(rn,r) for rn,r in _decl_cache]
    for rn, r in all_rows[-500:]:
        try:
            row_d = _parse_date(_row_date(r[2])) if r[2] else None
            if d_from and row_d and row_d < d_from: continue
            if d_to and row_d and row_d > d_to: continue
            trs = -1
            try:
                equiv_v = float(str(r[21] or 0).replace(",","."))
                pr = get_prod_ref()
                row_type = str(r[0] or "").strip().lower()
                if row_type in ("production","prod",""):
                    debut_s = _hms_to_sec(str(r[16] or "00:00:00"))
                    fin_s = _hms_to_sec(str(r[17] or "00:00:00"))
                    brut_s = fin_s - debut_s if fin_s > debut_s else _hms_to_sec(str(r[18] or "00:00:00"))
                    if pr>0 and brut_s>0 and equiv_v>0:
                        trs = round(equiv_v/(pr*brut_s/28800)*100,1)
                else:
                    trs_col = str(r[24] or "")
                    if trs_col:
                        try: trs = round(float(trs_col.replace(",",".")),1)
                        except: pass
            except: pass
            rows.append({
                "row_num": rn,
                "type": str(r[0] or ""),
                "date": _row_date(r[2]),
                "debut": str(r[16] or "")[:5],
                "fin": str(r[17] or "")[:5],
                "of": str(r[1] or ""),
                "pilote": str(r[4] or ""),
                "poste": str(r[3] or ""),
                "qte_fab": str(r[19] or ""),
                "qte_emb": str(r[20] or ""),
                "equiv": str(r[21] or ""),
                "trs": trs,
                "duree": str(r[18] or ""),
                "taille": str(r[7] or ""),
                "type_prod": str(r[9] or ""),
                "copilote": str(r[5] or ""),
                "nb_pers": str(r[6] or ""),
                "code_prod": str(r[8] or ""),
                "poids": str(r[10] or ""),
                "fibre": str(r[11] or ""),
                "of_taie": str(r[12] or ""),
                "traca": str(r[13] or ""),
                "ref_taie": str(r[14] or ""),
                "kit": str(r[15] or ""),
                "c1": str(r[22] or ""),
                "c2": str(r[23] or ""),
                "comment": str(r[35] or ""),
            })
        except: pass
    return jsonify(list(reversed(rows)))

@flask_app.route('/api/events_list')
def api_events_list():
    rows = []
    evt_rows = [(rn,r) for rn,r in _decl_cache if str(r[0] or "").strip().lower() not in ("production","prod","")]
    for rn, r in evt_rows[-100:]:
        try:
            hors = str(r[36] if len(r)>36 else "").strip().upper()
            type_str = str(r[0] or "").strip()
            tl = type_str.lower()
            if "nettoyage" in tl: cat = "nettoyage"
            elif tl == "pause": cat = "_pause"
            elif "rattrapage" in tl: cat = "ratt"
            elif tl.startswith("pb") or "panne" in tl: cat = "pb"
            elif tl: cat = "organisation"
            else: cat = "autre"
            rows.append({
                "row_num": rn,
                "type": type_str,
                "cat": cat,
                "of": str(r[1] or ""),
                "date": _row_date(r[2]),
                "poste": str(r[3] or ""),
                "pilote": str(r[4] or ""),
                "debut": str(r[16] or "")[:8],
                "fin": str(r[17] or "")[:8],
                "duree": str(r[18] or ""),
                "comment": str(r[35] or ""),
                "hors_trs": hors=="OUI",
            })
        except: pass
    return jsonify(list(reversed(rows)))

@flask_app.route('/api/config')
def api_config():
    return jsonify({
        "prod_ref": cfg.get("prod_ref",0),
        "pause_max_min": cfg.get("pause_max_min",20),
        "clean_short_min": cfg.get("clean_short_min",10),
        "clean_long_min": cfg.get("clean_long_min",30),
        "clean_grand_min": cfg.get("clean_grand_min",60),
        "meeting_tol_min": cfg.get("meeting_tol_min",5),
        "db_path": cfg.get("db_path",""),
        "db_name": os.path.basename(cfg.get("db_path","")) if cfg.get("db_path") else "",
        "modeles_horaires": cfg.get("modeles_horaires",[]),
        "pilot_passwords": cfg.get("pilot_passwords",{}),
    })

@flask_app.route('/api/settings', methods=['POST'])
def api_settings():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    for k in ["prod_ref","pause_max_min","clean_short_min","clean_long_min","clean_grand_min","meeting_tol_min"]:
        if k in data:
            try: cfg[k]=float(data[k]) if k=="prod_ref" else int(data[k])
            except: pass
    if "new_pw" in data and data["new_pw"]:
        cfg["supervisor_pw"] = str(data["new_pw"])
    if "modeles_horaires" in data:
        cfg["modeles_horaires"] = data["modeles_horaires"]
    if "pilot_passwords" in data:
        cfg["pilot_passwords"] = data["pilot_passwords"]
    save_cfg_data()
    return jsonify({"ok":True})

@flask_app.route('/api/set_db', methods=['POST'])
def api_set_db():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    path = data.get("path","").strip()
    if not path or not os.path.exists(path):
        return jsonify({"ok":False,"error":"Fichier introuvable"}),400
    cfg["db_path"] = path
    save_cfg_data()
    threading.Thread(target=load_lists,daemon=True).start()
    threading.Thread(target=load_history,daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/toggle_hors_trs', methods=['POST'])
def api_toggle_hors_trs():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    toggle_hors_trs_excel(data.get("row_num",0), data.get("new_val",""))
    return jsonify({"ok":True})

@flask_app.route('/api/update_model_today', methods=['POST'])
def api_update_model_today():
    """Met à jour les horaires du jour pour un modèle, sans mot de passe."""
    data = request.json or {}
    nom = data.get("nom","")
    day_key = data.get("day_key","")
    debut = data.get("debut","")
    fin = data.get("fin","")
    if not nom or not day_key or not debut or not fin:
        return jsonify({"ok":False,"error":"Paramètre manquant"}),400
    for m in cfg.get("modeles_horaires",[]):
        if m.get("nom","") == nom:
            if "jours" not in m:
                m["jours"] = {}
            if day_key not in m["jours"]:
                m["jours"][day_key] = {}
            m["jours"][day_key]["debut"] = debut
            m["jours"][day_key]["fin"] = fin
            break
    save_cfg_data()
    return jsonify({"ok":True})

@flask_app.route('/api/delete_row', methods=['POST'])
def api_delete_row():
    """Supprime n'importe quelle ligne de la feuille Declarations par row_num."""
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    row_num = data.get("row_num")
    path = cfg.get("db_path","")
    if not row_num or not path or not os.path.exists(path):
        return jsonify({"ok":False,"error":"Paramètre manquant"}),400
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                sheet = wb["Declarations"] if "Declarations" in wb.sheetnames else None
                if sheet is None: return
                sheet.delete_rows(row_num)
                _safe_excel_save(wb,path)
            threading.Thread(target=load_history,daemon=True).start()
        except: pass
    threading.Thread(target=_bg,daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/edit_row', methods=['POST'])
def api_edit_row():
    """Modifie une ligne dans Declarations."""
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    row_num = data.get("row_num")
    updates = data.get("updates",{})  # {col_idx: value} (1-indexed)
    path = cfg.get("db_path","")
    if not row_num or not path or not os.path.exists(path):
        return jsonify({"ok":False,"error":"Paramètre manquant"}),400
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Declarations" not in wb.sheetnames: return
                ws = wb["Declarations"]
                for col_str, val in updates.items():
                    try: ws.cell(row_num, int(col_str)).value = val
                    except: pass
                try:
                    row_type = str(ws.cell(row_num, 1).value or "").strip().lower()
                    if row_type in ("production","prod",""):
                        debut_s = _hms_to_sec(str(ws.cell(row_num, 17).value or "00:00:00"))
                        fin_s = _hms_to_sec(str(ws.cell(row_num, 18).value or "00:00:00"))
                        brut_s = fin_s - debut_s if fin_s > debut_s else 0
                        if brut_s > 0:
                            ws.cell(row_num, 19).value = fmt(brut_s)
                        qte_fab_v = 0.0
                        try: qte_fab_v = float(str(ws.cell(row_num, 20).value or 0).replace(",","."))
                        except: pass
                        taille_v = str(ws.cell(row_num, 8).value or "").strip()
                        type_prod_v = str(ws.cell(row_num, 10).value or "").strip()
                        new_equiv = calc_equiv(qte_fab_v, taille_v, type_prod_v)
                        if new_equiv > 0:
                            ws.cell(row_num, 22).value = new_equiv
                        nb_pers_v = 1
                        try: nb_pers_v = max(1, float(str(ws.cell(row_num, 7).value or 1).replace(",",".") or 1))
                        except: pass
                        of_hrs = brut_s / 3600 if brut_s > 0 else 0
                        if of_hrs > 0 and new_equiv > 0:
                            ws.cell(row_num, 23).value = round(new_equiv / of_hrs, 2)
                            ws.cell(row_num, 24).value = round(new_equiv / (nb_pers_v * of_hrs), 2)
                        pr = get_prod_ref()
                        if pr > 0 and brut_s > 0 and new_equiv > 0:
                            ws.cell(row_num, 25).value = str(round(new_equiv/(pr*brut_s/28800)*100,1))
                except: pass
                _safe_excel_save(wb,path)
            threading.Thread(target=load_history,daemon=True).start()
        except: pass
    threading.Thread(target=_bg,daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/fin_poste_data')
def api_fin_poste_data():
    today = datetime.date.today().strftime("%d/%m/%Y")
    shift_start_dt = _S.get("shift_start")
    shift_date = shift_start_dt.date() if shift_start_dt else datetime.date.today()
    shift_date_str = shift_date.strftime("%d/%m/%Y")
    pilot = _S["pilot"] or ""
    pilot_poste = _S["poste"] or ""
    prod_ref = get_prod_ref()
    tot_eq=0.0; tot_s=0.0
    of_list=[]
    prod_rows = [(rn,r) for rn,r in _decl_cache if str(r[0] or "").strip().lower() in ("production","prod","")]
    for rn, r in prod_rows:
        try:
            rd = _row_date(r[2])
            if rd != shift_date_str and rd != today: continue
            if str(r[4] or "")!=pilot: continue
            eq=float(str(r[21] or 0).replace(",","."))
            debut_s=_hms_to_sec(str(r[16] or "00:00:00"))
            fin_s=_hms_to_sec(str(r[17] or "00:00:00"))
            s = fin_s - debut_s if fin_s > debut_s else _hms_to_sec(str(r[18] or "00:00:00"))
            tot_eq+=eq; tot_s+=s
            trs=-1
            if prod_ref>0 and s>0 and eq>0:
                trs=round(eq/(prod_ref*s/28800)*100,1)
            of_list.append({
                "of":str(r[1] or ""),"taille":str(r[7] or ""),
                "type_prod":str(r[9] or ""),"qte_fab":str(r[19] or ""),
                "qte_emb":str(r[20] or ""),"equiv":str(r[21] or ""),
                "debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],
                "duree":str(r[18] or ""),"trs":trs,
            })
        except: pass
    trs_poste=-1.0
    if prod_ref>0 and tot_s>0: trs_poste=round(tot_eq/(prod_ref*tot_s/28800)*100,1)
    # TRS shift : même formule que l'accueil — elapsed = lastProdFin − modelDebut
    debut_str, _ = _get_model_day_cfg(pilot_poste, datetime.date.today())
    model_debut_s = _hms_to_sec(debut_str) if debut_str else None
    max_fin_s = 0.0
    for rn, r in prod_rows:
        try:
            rd = _row_date(r[2])
            if rd != shift_date_str and rd != today: continue
            if str(r[4] or "") != pilot: continue
            fs = _hms_to_sec(str(r[17] or "00:00:00"))
            if fs > max_fin_s: max_fin_s = fs
        except: pass
    trs_poste_shift = -1.0
    if model_debut_s is not None and max_fin_s > model_debut_s and prod_ref > 0:
        elapsed_s = max_fin_s - model_debut_s
        trs_poste_shift = round(tot_eq/(prod_ref*elapsed_s/28800)*100,1)
    elif prod_ref > 0 and tot_s > 0:
        trs_poste_shift = round(tot_eq/(prod_ref*tot_s/28800)*100,1)
    return jsonify({
        "pilot":pilot,"date":today,
        "nb_of":len(of_list),"trs":trs_poste,"trs_shift":trs_poste_shift,
        "tot_equiv":round(tot_eq,1),"tot_s":round(tot_s,0),
        "of_list":of_list,"of_count_shift":_S["of_count_shift"],
        "shift_start_iso": _dt_str(_S.get("shift_start")),
    })

@flask_app.route('/api/history_today')
def api_history_today():
    today = datetime.date.today().strftime("%d/%m/%Y")
    shift_start_dt = _S.get("shift_start")
    shift_date = shift_start_dt.date() if shift_start_dt else datetime.date.today()
    shift_date_str = shift_date.strftime("%d/%m/%Y")
    pilot = _S["pilot"] or ""
    poste = _S["poste"] or ""
    prod_ref = get_prod_ref()
    shift_s = get_shift_duration_s(poste, shift_date)
    rows = []
    tot_eq=0.0; tot_s=0.0
    for rn,r in _decl_cache:
        if str(r[0] or "").strip().lower() not in ("production","prod",""): continue
        rd = _row_date(r[2])
        if rd != shift_date_str and rd != today: continue
        if str(r[4] or "") != pilot: continue
        eq=float(str(r[21] or 0).replace(",",".") or 0)
        s=_hms_to_sec(str(r[18] or "00:00:00"))
        tot_eq+=eq; tot_s+=s
        trs_of=-1
        try:
            trs_col=str(r[24] or "")
            if trs_col: trs_of=round(float(trs_col.replace(",",".")),1)
            elif prod_ref>0 and s>0: trs_of=round(eq/(prod_ref*s/28800)*100,1)
        except: pass
        rows.append({"of":str(r[1] or ""),"debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],"trs":trs_of,"equiv":eq,"qte_fab":str(r[19] or "")})
    trs_shift=-1.0
    if prod_ref>0 and shift_s>0: trs_shift=round(tot_eq/(prod_ref*shift_s/28800)*100,1)
    trs_of_time=-1.0
    if prod_ref>0 and tot_s>0: trs_of_time=round(tot_eq/(prod_ref*tot_s/28800)*100,1)
    return jsonify({"rows":rows,"trs_shift":trs_shift,"trs_of":trs_of_time,"tot_eq":round(tot_eq,1),"shift_s":shift_s})

@flask_app.route('/api/reload', methods=['POST'])
def api_reload():
    threading.Thread(target=load_lists,daemon=True).start()
    threading.Thread(target=load_history,daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/save_poste', methods=['POST'])
def api_save_poste():
    data = request.json or {}
    if not data.get("dur_poste_theorique_min"):
        data["dur_poste_theorique_min"] = round(get_shift_duration_s(_S.get("poste","")) / 60, 1)
    write_poste_row(data)
    return jsonify({"ok":True})

@flask_app.route('/api/pilot_passwords_excel', methods=['POST'])
def api_pilot_passwords_excel():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    pilot_passwords = data.get("pilot_passwords",{})
    cfg["pilot_passwords"] = pilot_passwords
    save_cfg_data()
    ok = write_pilots_to_excel(pilot_passwords)
    return jsonify({"ok":ok})

@flask_app.route('/api/generate_dashboard', methods=['POST'])
def api_generate_dashboard():
    html_path, err = generate_dashboard_html()
    if err: return jsonify({"ok":False,"error":err})
    return jsonify({"ok":True,"path":html_path})

# ── Dashboard HTML superviseur ─────────────────────────────────────────────────
def generate_dashboard_html():
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path):
        return None, "Aucun fichier Excel configuré ou introuvable"
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return None, f"Erreur lecture Excel: {e}"

    prod_ref = get_prod_ref()
    from collections import defaultdict
    import math

    decl_rows = []
    if "Declarations" in wb.sheetnames:
        ws = wb["Declarations"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and any(r): decl_rows.append(list(r)+[None]*5)
    elif "Data" in wb.sheetnames:
        ws = wb["Data"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and any(r):
                row = list(r)+[None]*10; u=[None]*37
                u[0]="Production";u[1]=row[0];u[2]=row[1];u[3]=row[2];u[4]=row[3]
                u[16]=row[17];u[17]=row[18];u[18]=row[16];u[19]=row[13]
                u[20]=row[14];u[21]=row[15];u[22]=row[19];u[23]=row[20];u[24]=None
                decl_rows.append(u)
    wb.close()

    prod_rows_all = [r for r in decl_rows if str(r[0] or "").strip().lower() in ("production","prod","")]
    evt_rows_all  = [r for r in decl_rows if str(r[0] or "").strip().lower() not in ("production","prod","")]

    def trs_color(t):
        if t < 0: return "#94a3b8"
        if t >= 70: return "#22c55e"
        if t >= 50: return "#f59e0b"
        return "#ef4444"

    def hms2s(s):
        try:
            if hasattr(s,'hour'): return s.hour*3600+s.minute*60+getattr(s,'second',0)
            p = str(s).strip().split(":")
            if len(p)==3: return int(p[0])*3600+int(p[1])*60+float(p[2])
            if len(p)==2: return int(p[0])*3600+float(p[1])*60  # HH:MM
        except: pass
        return 0.0

    def fmt_s(s):
        s = max(0, int(s or 0))
        return f"{s//3600:02d}h{(s%3600)//60:02d}"

    today_str = datetime.date.today().strftime("%d/%m/%Y")
    pilot_now = _S.get("pilot","") or ""
    poste_now = _S.get("poste","") or ""
    prod_active = bool(_S.get("prod_active"))
    active_stops = [k for k,t in _S.get("timers",{}).items() if t.get("running") and not k.startswith("_")]
    of_num_now = (_S.get("form") or {}).get("of_num","") or "—"
    taille_now = (_S.get("form") or {}).get("taille","") or ""
    type_prod_now = (_S.get("form") or {}).get("type_prod","") or ""
    is_paused = _S.get("is_paused", False)

    today_prod = [r for r in prod_rows_all if _row_date(r[2])==today_str and str(r[4] or "")==pilot_now]
    today_evts = [r for r in evt_rows_all if _row_date(r[2])==today_str and str(r[4] or "")==pilot_now]

    shift_start_dt = _S.get("shift_start")

    trs_poste = -1.0
    elapsed_for_trs = 0.0

    # Use model horaire debut as shift reference (not shift_start which may include pre-shift events)
    model_debut_dt = None
    _day_map = {0:'lun',1:'mar',2:'mer',3:'jeu',4:'ven',5:'sam',6:'dim'}
    _dk = _day_map.get(datetime.date.today().weekday(), 'lun')
    model_fin_dt = None
    for _m in cfg.get("modeles_horaires", []):
        if str(_m.get("nom","")).strip() == str(poste_now).strip():
            _j = _m.get("jours",{}).get(_dk,{})
            _deb = _j.get("debut","") or _m.get("debut","")
            _fin = _j.get("fin","") or _m.get("fin","")
            if _deb:
                try:
                    _h, _mi = map(int, _deb.split(":"))
                    model_debut_dt = datetime.datetime.combine(datetime.date.today(), datetime.time(_h, _mi))
                except: pass
            if _fin:
                try:
                    _h2, _mi2 = map(int, _fin.split(":"))
                    model_fin_dt = datetime.datetime.combine(datetime.date.today(), datetime.time(_h2, _mi2))
                except: pass
            break
    # Filter to only prods within model horaire window (same logic as JS inShiftDecls)
    if model_debut_dt:
        _mdeb_s = model_debut_dt.hour * 3600 + model_debut_dt.minute * 60
        def _ts_s(s):
            try: p = str(s)[:5].split(':'); return int(p[0]) * 3600 + int(p[1]) * 60
            except: return 0
        in_shift_prod = [r for r in today_prod if _ts_s(r[17]) >= _mdeb_s or _ts_s(r[16]) >= _mdeb_s]
    else:
        in_shift_prod = today_prod
    tot_equiv = sum(float(str(r[21] or "0").replace(",",".") or 0) for r in in_shift_prod)
    nb_of_today = len(in_shift_prod)
    nb_pieces = sum(int(str(r[19] or 0).split('.')[0] or 0) for r in in_shift_prod)
    # Recompute last_fin_dt from in_shift_prod for TRS elapsed
    last_fin_dt = None
    for r in in_shift_prod:
        fin_str = str(r[17] or "")
        if fin_str and ":" in fin_str:
            try:
                t = datetime.datetime.strptime(f"{today_str} {fin_str[:8]}", "%d/%m/%Y %H:%M:%S")
                if last_fin_dt is None or t > last_fin_dt:
                    last_fin_dt = t
            except: pass
    # Match accueil formula: use last declared fin, not now()
    # (accueil uses _lastProdDeclTime which is fin of last declared OF)
    if prod_active and last_fin_dt is None:
        last_fin_dt = datetime.datetime.now()
    ref_start_dt = model_debut_dt or shift_start_dt
    if ref_start_dt and last_fin_dt and prod_ref > 0:
        elapsed_for_trs = (last_fin_dt - ref_start_dt).total_seconds()
        if elapsed_for_trs > 0:
            trs_poste = round(tot_equiv / (prod_ref * elapsed_for_trs / 28800) * 100, 1)

    prod_s_total = sum(hms2s(str(r[18] or "0")) for r in in_shift_prod)
    # Filter stop events to model horaire window (same logic as in_shift_prod)
    if model_debut_dt:
        in_shift_evts = [r for r in today_evts if _ts_s(r[17]) >= _mdeb_s or _ts_s(r[16]) >= _mdeb_s]
    else:
        in_shift_evts = today_evts
    stop_s_total = sum(hms2s(str(r[18] or "0")) for r in in_shift_evts
                       if str(r[0] or "").lower() not in ("pause pilote","changement d'of","interposte","changement de serie"))

    evt_dur = defaultdict(float)
    for r in in_shift_evts:
        t = str(r[0] or "")
        if not t or t.lower() in ("pause pilote","changement d'of","interposte"): continue
        evt_dur[t] += hms2s(str(r[18] or "0"))
    pareto = sorted(evt_dur.items(), key=lambda x: -x[1])[:8]

    all_stops_info = []  # list of (name, elapsed_s)
    if active_stops:
        evts_cfg_list = get_events_list()
        for k in active_stops:
            ev = next((e for e in evts_cfg_list if e.get("key")==k), None)
            stop_name = ev["label"] if ev else k
            t_data = _S.get("timers",{}).get(k, {})
            stop_elapsed = t_data.get("elapsed", 0)
            if t_data.get("running") and t_data.get("start"):
                try: stop_elapsed += (datetime.datetime.now() - t_data["start"]).total_seconds()
                except: pass
            all_stops_info.append((stop_name, stop_elapsed))
    elif is_paused:
        pause_elapsed = _S.get("pause_total_s", 0)
        if _S.get("pause_start"):
            try: pause_elapsed += (datetime.datetime.now() - _S["pause_start"]).total_seconds()
            except: pass
        all_stops_info.append(("Pause", pause_elapsed))
    # Legacy single-stop for compat
    active_stop_name = all_stops_info[0][0] if all_stops_info else ""
    active_stop_elapsed = all_stops_info[0][1] if all_stops_info else 0.0

    has_alert = bool(active_stops or is_paused)
    gen_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    elapsed_str = fmt_s(elapsed_for_trs) if elapsed_for_trs > 0 else "—"
    prod_pct = int(prod_s_total / (prod_s_total + stop_s_total) * 100) if (prod_s_total + stop_s_total) > 0 else 0
    trs_col = trs_color(trs_poste)

    # ── Gauge arc SVG ──
    def gauge_svg(trs, size=220):
        pct = max(0, min(100, trs)) if trs >= 0 else 0
        col = trs_color(trs)
        r_out=90; r_in=62; cx=100; cy=104
        def arc_pt(r, deg):
            rad = math.radians(deg)
            return cx+r*math.cos(rad), cy+r*math.sin(rad)
        x1,y1=arc_pt(r_out,180); x2,y2=arc_pt(r_out,0)
        xi1,yi1=arc_pt(r_in,180); xi2,yi2=arc_pt(r_in,0)
        bg=f'<path d="M{x1:.1f},{y1:.1f} A{r_out},{r_out} 0 0,1 {x2:.1f},{y2:.1f} L{xi2:.1f},{yi2:.1f} A{r_in},{r_in} 0 0,0 {xi1:.1f},{yi1:.1f} Z" fill="#e2e8f0"/>'
        fg=""
        if pct > 0:
            end_deg = 180 - pct * 1.8
            fx1,fy1=arc_pt(r_out,180); fx2,fy2=arc_pt(r_out,end_deg)
            fxi1,fyi1=arc_pt(r_in,180); fxi2,fyi2=arc_pt(r_in,end_deg)
            lg=1 if pct>50 else 0
            fg=f'<path d="M{fx1:.1f},{fy1:.1f} A{r_out},{r_out} 0 {lg},1 {fx2:.1f},{fy2:.1f} L{fxi2:.1f},{fyi2:.1f} A{r_in},{r_in} 0 {lg},0 {fxi1:.1f},{fyi1:.1f} Z" fill="{col}"/>'
        lbl_txt = f"{trs:.1f}%" if trs >= 0 else "—"
        h = int(size * 110 // 200)
        return (f'<svg width="{size}" height="{h}" viewBox="0 0 200 110">'
                f'{bg}{fg}'
                f'<text x="{cx}" y="{cy+6}" text-anchor="middle" font-size="28" font-weight="900" fill="{col}">{lbl_txt}</text>'
                f'</svg>')

    # ── Pie chart SVG ──
    def pie_svg(prod_s, stop_s, size=180):
        total = prod_s + stop_s
        if total <= 0:
            return f'<svg width="{size}" height="{size}"><text x="{size//2}" y="{size//2+6}" text-anchor="middle" font-size="16" fill="#475569">Pas de données</text></svg>'
        cx = cy = size // 2
        r = size // 2 - 8
        def seg(start_a, end_a, color):
            s = math.radians(start_a); e = math.radians(end_a)
            lg = 1 if (end_a - start_a) > 180 else 0
            x1,y1 = cx+r*math.cos(s), cy+r*math.sin(s)
            x2,y2 = cx+r*math.cos(e), cy+r*math.sin(e)
            return f'<path d="M{cx},{cy} L{x1:.1f},{y1:.1f} A{r},{r} 0 {lg},1 {x2:.1f},{y2:.1f} Z" fill="{color}"/>'
        prod_end = (prod_s / total) * 360 - 90
        parts = seg(-90, prod_end, "#22c55e") + seg(prod_end, 270, "#ef4444")
        pp = int(prod_s / total * 100)
        lbl = f'<text x="{cx}" y="{cy-6}" text-anchor="middle" font-size="20" font-weight="900" fill="#fff">{pp}%</text>'
        lbl += f'<text x="{cx}" y="{cy+16}" text-anchor="middle" font-size="13" fill="#cbd5e1">Prod</text>'
        return f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}">{parts}{lbl}</svg>'

    # ── Timeline SVG ──
    def timeline_svg(W=1200, H=64):
        now_ts = datetime.datetime.now()
        win_start = model_debut_dt if model_debut_dt else (shift_start_dt if shift_start_dt else now_ts - datetime.timedelta(hours=8))
        win_end = now_ts
        span = (win_end - win_start).total_seconds()
        if span <= 0: span = 28800
        Y=18; BH=38
        def to_x(dt_str):
            try:
                if hasattr(dt_str,'strftime'):
                    t = dt_str.strftime("%H:%M:%S")
                else:
                    t = str(dt_str).strip()[:8]
                    if len(t)==5: t += ":00"
                dt = datetime.datetime.strptime(f"{today_str} {t}", "%d/%m/%Y %H:%M:%S")
                return max(0, min(W, int((dt-win_start).total_seconds()/span*W)))
            except: return 0
        catcol = {"pb":"#ef4444","ratt":"#f59e0b","nettoyage":"#38bdf8","pause":"#64748b","organisation":"#a855f7"}
        svg = f'<svg width="100%" viewBox="0 0 {W} {H}" style="display:block" preserveAspectRatio="none">'
        svg += f'<rect x="0" y="{Y}" width="{W}" height="{BH}" fill="#e2e8f0" rx="4"/>'
        for r in today_prod:
            x1 = to_x(str(r[16] or ""))
            x2 = to_x(str(r[17] or "")) if r[17] else int((now_ts-win_start).total_seconds()/span*W)
            x2 = max(x2, x1+3)
            svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="#22c55e" rx="2" opacity="0.85"/>'
        # Current running OF (not yet declared) — show in lighter green
        if prod_active:
            _of_bar_start = last_fin_dt or model_debut_dt or shift_start_dt
            if _of_bar_start:
                x1 = to_x(_of_bar_start)
                x2 = int((now_ts - win_start).total_seconds() / span * W)
                x2 = max(x2, x1 + 3)
                svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="#4ade80" rx="2" opacity="0.65" stroke="#16a34a" stroke-width="1" stroke-dasharray="4,2"/>'
        for r in today_evts:
            t = str(r[0] or "").lower()
            x1 = to_x(str(r[16] or ""))
            x2 = to_x(str(r[17] or "")) if r[17] else int((now_ts-win_start).total_seconds()/span*W)
            x2 = max(x2, x1+3)
            col = catcol["pb"] if ("pb" in t or "panne" in t or "technique" in t) else catcol["ratt"] if "ratt" in t else catcol["nettoyage"] if "nett" in t else catcol["pause"] if "pause" in t else "#94a3b8"
            svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="{col}" rx="2" opacity="0.95"/>'
        h_span = span / 3600
        step = 1 if h_span <= 10 else 2
        cur = win_start.replace(minute=0, second=0, microsecond=0)
        if cur < win_start: cur += datetime.timedelta(hours=1)
        while cur <= win_end:
            frac = (cur-win_start).total_seconds()/span
            x = int(frac*W)
            svg += f'<line x1="{x}" y1="{Y}" x2="{x}" y2="{Y+BH}" stroke="#94a3b8" stroke-width="1"/>'
            svg += f'<text x="{x}" y="{Y-3}" font-size="12" fill="#475569" text-anchor="middle">{cur.strftime("%H:%M")}</text>'
            cur += datetime.timedelta(hours=step)
        now_x = int((now_ts-win_start).total_seconds()/span*W)
        svg += f'<line x1="{now_x}" y1="{Y-4}" x2="{now_x}" y2="{Y+BH+4}" stroke="#1e293b" stroke-width="2.5"/>'
        svg += '</svg>'
        return svg

    # ── Pareto bars HTML ──
    pareto_html = ""
    if pareto:
        max_dur = pareto[0][1]
        for lbl, dur in pareto:
            pct = dur / max_dur * 100 if max_dur > 0 else 0
            col = "#ef4444" if any(x in lbl.lower() for x in ["pb","panne","technique"]) else "#f59e0b" if "ratt" in lbl.lower() else "#38bdf8" if "nett" in lbl.lower() else "#a855f7"
            pareto_html += f'''<div style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
              <div style="width:120px;font-size:11px;text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0;color:#475569">{lbl[:22]}</div>
              <div style="flex:1;background:#e2e8f0;border-radius:3px;height:16px">
                <div style="width:{pct:.0f}%;height:16px;background:{col};border-radius:3px"></div>
              </div>
              <div style="width:40px;font-size:12px;font-weight:800;text-align:right;flex-shrink:0;color:#1e293b">{dur/60:.0f}m</div>
            </div>'''
    else:
        pareto_html = '<div style="color:#475569;font-size:12px;padding:10px;text-align:center">Aucun arrêt enregistré</div>'

    # ── Productions table ──
    prod_rows_html = ""
    for r in list(reversed(in_shift_prod))[:8]:
        trs_val = ""
        try:
            tv = float(str(r[24] or "").replace(",","."))
            tc = trs_color(tv)
            trs_val = f'<span style="color:{tc};font-weight:900">{tv:.1f}%</span>'
        except: pass
        prod_rows_html += f'''<tr>
          <td style="font-weight:800;font-size:12px;color:#1e293b">{r[1] or ""}</td>
          <td style="color:#475569">{str(r[16] or "")[:5]}</td><td style="color:#475569">{str(r[17] or "")[:5]}</td>
          <td style="color:#475569">{r[18] or ""}</td>
          <td style="color:#1e293b">{r[19] or "0"}</td><td style="font-weight:800;color:#0891b2">{r[21] or ""}</td>
          <td>{trs_val}</td>
        </tr>'''
    if not prod_rows_html:
        prod_rows_html = '<tr><td colspan="7" style="color:#94a3b8;padding:10px;text-align:center;font-size:12px">Aucune production déclarée</td></tr>'

    # ── ALERT BANNER HTML ──
    alert_html = ""
    if has_alert:
        stops_html = "".join(
            f'<div style="display:flex;align-items:center;justify-content:space-between;gap:16px;background:rgba(255,255,255,.12);border-radius:8px;padding:6px 14px;margin:3px 0;min-width:240px">'
            f'<span style="font-size:18px;font-weight:800;color:#fef2f2">{nm}</span>'
            f'<span style="font-size:24px;font-weight:900;color:#fecaca;font-variant-numeric:tabular-nums">{int(el/60)}<span style="font-size:13px">min</span></span>'
            f'</div>'
            for nm, el in all_stops_info
        )
        alert_html = f'''
<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;
  background:linear-gradient(135deg,#7f0000 0%,#b91c1c 50%,#ef4444 100%);
  animation:pulse 1.2s ease-in-out infinite;border-bottom:6px solid #fca5a5">
  <div style="font-size:42px;line-height:1;animation:wag .8s ease-in-out infinite">🚨</div>
  <div style="font-size:40px;font-weight:900;letter-spacing:4px;margin:8px 0;text-shadow:0 4px 16px rgba(0,0,0,.4);color:#fff">ARRÊT{"S" if len(all_stops_info)>1 else ""} EN COURS</div>
  {stops_html}
  <div style="font-size:42px;line-height:1;animation:wag .8s ease-in-out infinite reverse;margin-top:8px">🚨</div>
</div>'''
    else:
        # ── PROD EN COURS BIG CARD ──
        prod_card_bg = "#f0fdf4" if prod_active else "#f8fafc"
        prod_card_border = "2px solid #22c55e" if prod_active else "2px solid #e2e8f0"
        prod_status_label = "▶ PRODUCTION EN COURS" if prod_active else "○ EN ATTENTE"
        prod_status_col = "#16a34a" if prod_active else "#64748b"
        alert_html = f'''
<div style="background:{prod_card_bg};border:{prod_card_border};border-radius:10px;padding:10px 18px;display:flex;align-items:center;gap:18px;flex-shrink:0">
  <div style="flex:1">
    <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:2px;color:{prod_status_col};margin-bottom:4px">{prod_status_label}</div>
    <div style="font-size:28px;font-weight:900;color:#1e293b;line-height:1">OF {of_num_now}</div>
    <div style="font-size:13px;color:#64748b;margin-top:4px">{'👤 ' + pilot_now + '  |  ' + poste_now if pilot_now else poste_now}</div>
  </div>
  {f'<div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;margin-bottom:3px">Taille / Type</div><div style="font-size:16px;font-weight:800;color:#1e293b">{taille_now} — {type_prod_now}</div></div>' if taille_now or type_prod_now else ''}
  <div style="text-align:center">
    <div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;margin-bottom:3px">Éq. aujourd'hui</div>
    <div style="font-size:28px;font-weight:900;color:#0891b2;line-height:1">{tot_equiv:.1f}</div>
  </div>
  <div style="text-align:center">
    <div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;margin-bottom:3px">OF déclarés</div>
    <div style="font-size:28px;font-weight:900;color:#7c3aed;line-height:1">{nb_of_today}</div>
  </div>
</div>'''

    tl_svg = timeline_svg()

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="15">
<title>Dashboard Encadrant — ORC</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:100%;overflow:hidden;font-family:-apple-system,'Segoe UI',Arial,sans-serif;background:#f1f5f9;color:#1e293b;font-size:13px}}
.hdr{{height:42px;background:#1e3a8a;color:#fff;display:flex;align-items:center;justify-content:space-between;padding:0 16px;flex-shrink:0;border-bottom:2px solid #3b82f6}}
.hdr-title{{font-size:15px;font-weight:900;display:flex;align-items:center;gap:10px}}
.hdr-badge{{background:rgba(255,255,255,.15);border-radius:6px;padding:3px 10px;font-size:12px;font-weight:700}}
.hdr-badge.green{{background:#15803d}}
.hdr-badge.gray{{background:#475569}}
.hdr-time{{font-size:11px;opacity:.85}}
.outer{{height:calc(100vh - 42px);display:flex;flex-direction:column;gap:6px;padding:6px;overflow:hidden}}
/* MODE NORMAL */
.main-grid{{flex:1;display:grid;grid-template-columns:200px 0.7fr 280px;gap:6px;overflow:hidden;min-height:0}}
.panel{{background:#fff;border-radius:10px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
.panel-hdr{{padding:6px 12px;font-size:10px;font-weight:900;text-transform:uppercase;letter-spacing:1px;flex-shrink:0}}
.panel-body{{flex:1;overflow-y:auto;padding:8px 12px;min-height:0}}
/* TRS PANEL */
.trs-num{{font-size:34px;font-weight:900;line-height:1;color:{trs_col};text-align:center}}
.trs-lbl{{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#64748b;text-align:center;margin-bottom:4px}}
.trs-sub{{font-size:11px;color:#64748b;text-align:center;margin-top:3px}}
/* STAT CARDS */
.stat-grid{{display:grid;grid-template-columns:1fr 1fr;gap:6px}}
.stat-card{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;text-align:center}}
.stat-val{{font-size:26px;font-weight:900;line-height:1}}
.stat-lbl{{font-size:10px;font-weight:700;text-transform:uppercase;color:#64748b;margin-top:3px}}
/* TIMELINE CELL */
.tl-cell{{background:#fff;border-radius:10px;padding:6px 10px;flex-shrink:0;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
.tl-lbl{{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#64748b;margin-bottom:4px;display:flex;justify-content:space-between}}
.tl-legend{{display:flex;gap:10px;font-size:11px;color:#64748b;margin-top:4px;flex-wrap:wrap}}
/* TABLE */
.ktbl{{width:100%;border-collapse:collapse;font-size:12px}}
.ktbl th{{background:#f1f5f9;padding:5px 8px;font-weight:800;text-align:center;position:sticky;top:0;font-size:10px;text-transform:uppercase;color:#64748b;border-bottom:1px solid #e2e8f0}}
.ktbl td{{padding:5px 8px;border-bottom:1px solid #e2e8f0;text-align:center;color:#1e293b}}
.ktbl tr:hover td{{background:#f8fafc}}
/* ANIMATIONS */
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.85}}}}
@keyframes wag{{0%{{transform:rotate(-8deg)}}50%{{transform:rotate(8deg)}}100%{{transform:rotate(-8deg)}}}}
/* SCROLLBAR */
::-webkit-scrollbar{{width:6px}};::-webkit-scrollbar-track{{background:#f1f5f9}};::-webkit-scrollbar-thumb{{background:#cbd5e1;border-radius:3px}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-title">
    &#127981; Dashboard Encadrant — ORC
    {'<span class="hdr-badge">' + ('👤 ' + pilot_now + ' | ' + poste_now if pilot_now else poste_now) + '</span>' if (pilot_now or poste_now) else ''}
  </div>
  <div style="display:flex;align-items:center;gap:14px">
    {'<span class="hdr-badge green">▶ PROD EN COURS — OF ' + of_num_now + '</span>' if prod_active else '<span class="hdr-badge gray">○ En attente</span>'}
    <span class="hdr-time">🔄 15s | {gen_time}</span>
  </div>
</div>

<div class="outer">

  {alert_html}

  {'<!-- MODE ALERTE : mini stats bar -->' if has_alert else ''}
  {f'''<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr 1fr;gap:6px;flex-shrink:0">
    <div class="stat-card"><div class="stat-val" style="color:{trs_col}">{f"{trs_poste:.1f}%" if trs_poste>=0 else "—"}</div><div class="stat-lbl">TRS Poste</div></div>
    <div class="stat-card"><div class="stat-val" style="color:#7c3aed">{nb_of_today}</div><div class="stat-lbl">OF déclarés</div></div>
    <div class="stat-card"><div class="stat-val" style="color:#0891b2">{tot_equiv:.1f}</div><div class="stat-lbl">Équivalence</div></div>
    <div class="stat-card"><div class="stat-val" style="color:#ef4444">{stop_s_total/60:.0f}<span style="font-size:14px">min</span></div><div class="stat-lbl">Arrêts</div></div>
    <div class="stat-card"><div class="stat-val" style="color:#16a34a">{prod_s_total/60:.0f}<span style="font-size:14px">min</span></div><div class="stat-lbl">Production</div></div>
  </div>''' if has_alert else ''}

  {'<!-- MODE NORMAL -->' if not has_alert else ''}
  {'''<div class="main-grid">

    <!-- COLONNE GAUCHE : TRS + Pie -->
    <div style="display:flex;flex-direction:column;gap:10px;overflow:hidden">''' if not has_alert else ''}

    {f'''<!-- TRS Block -->
      <div class="panel" style="flex:1">
        <div class="panel-hdr" style="background:#1a1f5e;color:#fff">TRS Poste en cours</div>
        <div class="panel-body" style="display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;padding:10px 14px">
          <div class="trs-lbl">Taux de Rendement Synthétique</div>
          <div class="trs-num">{f"{trs_poste:.1f}%" if trs_poste >= 0 else "—"}</div>
          {'<div class="trs-sub" style="font-weight:700;color:#0369a1">⏱ Modèle : ' + model_debut_dt.strftime("%H:%M") + ' → ' + (model_fin_dt.strftime("%H:%M") if model_fin_dt else "—") + '</div>' if model_debut_dt else ''}
          <div class="trs-sub">Réf : {prod_ref:.0f} éq / 8h &nbsp;|&nbsp; {elapsed_str}</div>
          <div class="trs-sub" style="font-size:13px;font-weight:800;color:#0891b2;margin-top:4px">Éq. total : {tot_equiv:.1f}</div>
        </div>
      </div>

      <!-- Pie + répartition -->
      <div class="panel" style="flex-shrink:0">
        <div class="panel-hdr" style="background:#0c4a6e;color:#fff">Répartition temps</div>
        <div class="panel-body" style="display:flex;align-items:center;justify-content:space-around;padding:6px">
          {pie_svg(prod_s_total, stop_s_total, 110)}
          <div style="display:flex;flex-direction:column;gap:6px">
            <div>
              <div style="font-size:10px;color:#64748b;text-transform:uppercase;font-weight:700">Production</div>
              <div style="font-size:24px;font-weight:900;color:#22c55e">{prod_pct}%</div>
              <div style="font-size:11px;color:#64748b">{prod_s_total/60:.0f} min</div>
            </div>
            <div>
              <div style="font-size:10px;color:#64748b;text-transform:uppercase;font-weight:700">Arrêts</div>
              <div style="font-size:24px;font-weight:900;color:#ef4444">{100-prod_pct}%</div>
              <div style="font-size:11px;color:#64748b">{stop_s_total/60:.0f} min</div>
            </div>
          </div>
        </div>
      </div>''' if not has_alert else ''}

    {('</div>' if not has_alert else '')}

    {'<!-- COLONNE CENTRE : Timeline + Productions -->' if not has_alert else ''}
    {'''<div style="display:flex;flex-direction:column;gap:10px;overflow:hidden;min-height:0">''' if not has_alert else ''}

      {f'''<!-- Productions déclarées -->
      <div class="panel" style="flex:1;min-height:0">
        <div class="panel-hdr" style="background:#14532d;color:#fff">Productions déclarées — aujourd'hui</div>
        <div class="panel-body" style="padding:0">
          <table class="ktbl">
            <thead><tr><th>OF</th><th>Début</th><th>Fin</th><th>Durée</th><th>Qté</th><th>Éq.</th><th>TRS</th></tr></thead>
            <tbody>{prod_rows_html}</tbody>
          </table>
        </div>
      </div>''' if not has_alert else ''}

    {('</div>' if not has_alert else '')}

    {'<!-- COLONNE DROITE : Stats + Pareto -->' if not has_alert else ''}
    {'''<div style="display:flex;flex-direction:column;gap:10px;overflow:hidden">''' if not has_alert else ''}

      {f'''<!-- Stats KPI -->
      <div class="panel" style="flex-shrink:0">
        <div class="panel-hdr" style="background:#1a1f5e;color:#fff">Indicateurs clés</div>
        <div class="panel-body">
          <div class="stat-grid">
            <div class="stat-card"><div class="stat-val" style="color:#8b5cf6">{nb_pieces}</div><div class="stat-lbl">Pièces</div></div>
            <div class="stat-card"><div class="stat-val" style="color:#38bdf8">{tot_equiv:.1f}</div><div class="stat-lbl">Équivalence</div></div>
            <div class="stat-card"><div class="stat-val" style="color:#4ade80">{nb_of_today}</div><div class="stat-lbl">OF déclarés</div></div>
            <div class="stat-card"><div class="stat-val" style="color:#ef4444">{stop_s_total/60:.0f}<span style="font-size:13px">m</span></div><div class="stat-lbl">Arrêts</div></div>
          </div>
        </div>
      </div>

      <!-- Pareto arrêts -->
      <div class="panel" style="flex:1;min-height:0">
        <div class="panel-hdr" style="background:#78350f;color:#fff">Pareto arrêts</div>
        <div class="panel-body">
          {pareto_html}
        </div>
      </div>''' if not has_alert else ''}

    {('</div>' if not has_alert else '')}

  {('</div>' if not has_alert else '')}

  <!-- TIMELINE (toujours visible) -->
  <div class="tl-cell" style="flex-shrink:0">
    <div class="tl-lbl">
      <span>Timeline du poste — {poste_now or "en cours"}</span>
      <span style="font-size:12px">{(model_debut_dt or shift_start_dt).strftime("%H:%M") if (model_debut_dt or shift_start_dt) else "—"} → maintenant</span>
    </div>
    {tl_svg}
    <div class="tl-legend">
      <span>■ <span style="color:#22c55e">Production</span></span>
      <span>⬚ <span style="color:#4ade80">OF en cours</span></span>
      <span>■ <span style="color:#ef4444">PB Technique</span></span>
      <span>■ <span style="color:#f59e0b">Rattrapage</span></span>
      <span>■ <span style="color:#38bdf8">Nettoyage</span></span>
      <span>■ <span style="color:#64748b">Pause</span></span>
      <span style="color:#1e293b;font-weight:700">| Maintenant</span>
    </div>
  </div>

</div>

</body>
</html>"""

    try:
        html_path = os.path.join(os.path.dirname(path), "KPI_Dashboard.html")
        with open(html_path, "w", encoding="utf-8") as f: f.write(html)
        return html_path, None
    except Exception as e:
        return None, str(e)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>KPI-ORC</title>
<style>
:root{
  --navy:#1a1f5e;--navy2:#2d3480;--green:#16a34a;--red:#dc2626;
  --amber:#d97706;--purple:#7c3aed;--blue:#0891b2;
  --bg:#f0f2f8;--card:#fff;--border:#dde4ef;--text:#1e293b;--gray:#64748b;
  --lgray:#e2e8f0;--radius:10px;--shadow:0 2px 10px rgba(0,0,0,.07);
  --hdr-h:52px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:13px;height:100vh;overflow:hidden;display:flex;flex-direction:column}

/* ── STOP ACTIVE THEME ── */
body.stop-on #app-hdr{background:#7f0000!important;border-color:#b91c1c}

/* ── HEADER ── */
#app-hdr{height:var(--hdr-h);background:var(--navy);display:flex;align-items:center;padding:0 14px;gap:8px;flex-shrink:0;border-bottom:2px solid var(--navy2)}
.hdr-logo{color:#fff;font-weight:800;font-size:15px;letter-spacing:1px;margin-right:10px;white-space:nowrap}
.hdr-tabs{display:flex;gap:2px;flex:1}
.htab{background:none;border:none;color:rgba(255,255,255,.65);padding:6px 12px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:500;white-space:nowrap;transition:all .15s}
.htab:hover{background:rgba(255,255,255,.12);color:#fff}
.htab.on{background:rgba(255,255,255,.2);color:#fff;font-weight:700}
.htab.prod-on{background:var(--green)!important;color:#fff!important;font-weight:700;animation:pt 2s infinite}
#ht-prod{display:none}
#ht-prod.prod-visible{display:inline-block!important}
@keyframes pt{0%,100%{opacity:1}50%{opacity:.75}}
#hdr-right{display:flex;align-items:center;gap:8px;margin-left:auto;font-size:11px;color:rgba(255,255,255,.75)}
#hdr-pilot-lbl{font-weight:800;color:#fff;font-size:18px;letter-spacing:.3px}

/* ── ALERT STRIP ── */
#alert-strip{background:#b91c1c;color:#fff;text-align:center;padding:4px;font-weight:700;font-size:12px;flex-shrink:0;display:none;animation:blink .85s step-start infinite}
#alert-strip.on{display:block}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.35}}

/* ── VIEWS ── */
.view{display:none;flex:1;flex-direction:column;overflow:hidden}
.view.on{display:flex}

/* ── LOGIN ── */
#v-login{background:linear-gradient(135deg,#1a1f5e,#2d3480,#1e3a8a);align-items:center;justify-content:center}
.login-card{background:#fff;border-radius:14px;padding:32px;width:100%;max-width:380px;box-shadow:0 20px 60px rgba(0,0,0,.35)}
.lc-h1{color:var(--navy);font-size:24px;font-weight:800;margin-bottom:2px;text-align:center}
.lc-sub{color:#64748b;text-align:center;margin-bottom:20px;font-size:12px}
.lf{margin-bottom:12px}
.lf label{display:block;font-weight:600;margin-bottom:3px;color:#374151;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
.lf select,.lf input{width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:14px;outline:none;transition:border .2s}
.lf select:focus,.lf input:focus{border-color:var(--navy)}
.btn-login{width:100%;padding:11px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:14px;font-weight:700;cursor:pointer;margin-top:4px}
.btn-login:hover{opacity:.88}
.ln-err{color:#dc2626;text-align:center;margin-top:6px;font-size:12px;min-height:16px}

/* ── MAIN VIEW (Déclarations / Évts) ── */
#v-main{overflow:hidden}
.main-hdr{background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap}
.main-hdr .mtabs{display:flex;gap:3px}
.mtab{background:none;border:none;border-bottom:2px solid transparent;padding:5px 12px;cursor:pointer;font-size:12px;font-weight:600;color:var(--gray);transition:all .15s}
.mtab.on{border-color:var(--navy);color:var(--navy)}
.main-hdr .mbtns{display:flex;gap:6px;margin-left:auto}
.btn-sm{padding:6px 12px;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s}
.btn-green{background:var(--green);color:#fff}
.btn-green:hover{filter:brightness(.9)}
.btn-ghost{background:var(--lgray);color:var(--text)}
.btn-ghost:hover{filter:brightness(.93)}
.table-wrap{flex:1;overflow-y:auto}
.ktbl{width:100%;border-collapse:collapse;font-size:12px}
.ktbl th{text-align:left;padding:7px 10px;background:var(--navy);color:#fff;font-size:10px;text-transform:uppercase;letter-spacing:.5px;position:sticky;top:0}
.ktbl td{padding:6px 10px;border-bottom:1px solid var(--border)}
.ktbl tr:hover td{background:var(--lgray)}
.tg{color:var(--green);font-weight:700}
.tm{color:var(--amber);font-weight:700}
.tb{color:var(--red);font-weight:700}
.btn-tbl{padding:3px 8px;border:none;border-radius:4px;font-size:10px;cursor:pointer;font-weight:600}

/* ── PRODUCTION VIEW ── */
#v-prod{overflow:hidden}
/* pilot/OF banner */
.pob{background:var(--navy);color:#fff;padding:6px 14px;display:flex;align-items:center;gap:24px;flex-shrink:0}
.pob-item{display:flex;flex-direction:column;gap:1px}
.pob-lbl{font-size:9px;text-transform:uppercase;letter-spacing:.7px;opacity:.7;font-weight:600}
.pob-val{font-size:22px;font-weight:800;line-height:1}
.pob-item.of .pob-val{font-size:20px;color:#93c5fd}
.pob-item.trs .pob-val{color:#86efac}
/* status bar */
.sbar{display:flex;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0}
.sc{flex:1;padding:6px 14px;border-right:1px solid var(--border);text-align:center}
.sc:last-child{border:none}
.sc-lbl{font-size:9px;text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.6px}
.sc-val{font-size:22px;font-weight:800;font-variant-numeric:tabular-nums;color:var(--navy);margin-top:1px}
.sc-val.green{color:var(--green)}
.sc-val.red{color:var(--red)}
.sc-val.amber{color:var(--amber)}
/* main 2-col layout */
.prod-body{display:flex;flex:1;overflow:hidden}
/* CENTER form col (now left) */
.form-col{flex:1;overflow-y:auto;padding:8px;display:flex;flex-direction:column;gap:6px}
/* Action buttons row below timeline */
.prod-act-row{display:flex;gap:6px;flex-wrap:wrap;padding:6px 0 2px;border-top:1px solid var(--border);margin-top:2px;position:sticky;bottom:0;background:var(--card);z-index:10}
.act-btn{flex:1;min-width:100px;border:none;border-radius:8px;padding:24px 6px;cursor:pointer;font-size:13px;font-weight:700;text-align:center;transition:all .15s;white-space:nowrap;min-height:72px;display:flex;align-items:center;justify-content:center}
.act-btn:hover{filter:brightness(.9)}
.act-stop{background:linear-gradient(135deg,#b91c1c,#7f0000);color:#fff;font-size:13px;font-weight:800;box-shadow:0 3px 8px rgba(185,28,28,.3)}
.act-nett{background:#e0f2fe;color:var(--blue)}
.act-pause{background:#f3e8ff;color:var(--purple)}
.act-cancel{background:#f1f5f9;color:#64748b;border:1px solid #cbd5e1}
.act-endprod{background:linear-gradient(135deg,#d97706,#b45309);color:#fff;font-weight:800;box-shadow:0 3px 8px rgba(217,119,6,.3)}
/* 3-col form zones */
.form-3col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.fzone{border-radius:7px;padding:8px}
.fzone h4{font-size:9px;text-transform:uppercase;letter-spacing:.7px;font-weight:700;margin-bottom:6px;padding-bottom:3px;border-bottom:1px solid rgba(0,0,0,.08)}
.zi{background:#eef2ff;border:1px solid #c7d2fe}.zi h4{color:#3730a3}
.zp{background:#f0fdf4;border:1px solid #bbf7d0}.zp h4{color:#166534}
.zq{background:#fff7ed;border:1px solid #fed7aa}.zq h4{color:#9a3412}
.fr{display:flex;flex-direction:column;margin-bottom:4px}
.fr label{font-size:9px;font-weight:700;color:var(--gray);margin-bottom:2px;text-transform:uppercase;letter-spacing:.2px}
.fr input,.fr select,.fr textarea{padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:12px;background:var(--card);color:var(--text);width:100%;outline:none;transition:border .15s}
.fr input:focus,.fr select:focus{border-color:#6366f1}
.fr textarea{resize:none;height:42px}
.fr.comment-big textarea{height:80px;font-size:13px;border:2px solid #f59e0b;background:#fffbeb;font-weight:500}
.fr.comment-big label{color:#d97706;font-size:10px}
input[type=checkbox]{cursor:pointer}
input:not([type=checkbox]):not([type=radio]):not([type=button]):not([type=submit]),textarea{cursor:text!important}
input,select,textarea{cursor:auto}
input[type=text],input[type=number],input[type=password],input[type=time],input[type=date],textarea{cursor:text!important}
.tl-legend{display:flex;gap:12px;padding:2px 4px;font-size:10px;color:var(--gray);flex-wrap:wrap;align-items:center}
.tl-legend span{display:flex;align-items:center;gap:3px}
.tl-legend i{display:inline-block;width:12px;height:10px;border-radius:2px;flex-shrink:0}
select{cursor:default}
.fr.big input{font-size:16px;font-weight:700;padding:5px 6px;color:var(--green)}
.fr.ro input{background:#f8fafc;color:var(--gray)}
/* Timeline */
.tl-wrap{background:var(--card);border-radius:7px;padding:7px 8px;border:1px solid var(--border)}
.tl-wrap h5{font-size:9px;text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.7px;margin-bottom:4px}
/* RIGHT recap col (wider) */
.recap-col{width:280px;flex-shrink:0;border-left:1px solid var(--border);background:var(--card);display:flex;flex-direction:column;overflow:hidden}
.recap-hdr{font-size:10px;text-transform:uppercase;font-weight:700;color:var(--gray);letter-spacing:.6px;padding:8px 8px 4px}
.recap-body{flex:1;overflow-y:auto;padding:0 6px 6px}
.si{display:flex;align-items:center;gap:5px;padding:3px 0;border-bottom:1px solid var(--border);font-size:11px}
.si:last-child{border:none}
.sdot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.si-nm{flex:1;font-weight:600;font-size:10px;line-height:1.2}
.si-dur{font-size:9px;color:var(--gray);white-space:nowrap}
.btn-edit{background:none;border:none;cursor:pointer;font-size:11px;color:#6366f1;padding:1px 3px;border-radius:2px}
/* TRS gauge */
.gauge-box{padding:6px;border-top:1px solid var(--border);text-align:center}
.gauge-lbl{font-size:9px;text-transform:uppercase;color:var(--gray);font-weight:700;margin-top:2px}
/* Active stops bottom bar — chips */
#stop-bottom{display:none;background:#7f0000;color:#fff;padding:8px 14px;align-items:center;gap:8px;flex-shrink:0;border-top:2px solid #b91c1c;flex-wrap:wrap}
#stop-bottom.on{display:flex}
.stop-chip{display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.28);border-radius:8px;padding:6px 10px;border:1px solid rgba(255,255,255,.2)}
.chip-lbl{font-weight:800;font-size:13px;white-space:nowrap}
.chip-tim{font-size:18px;font-weight:800;font-variant-numeric:tabular-nums;min-width:52px;text-align:right}
.btn-endstop{background:#16a34a;color:#fff;border:none;border-radius:6px;padding:6px 12px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap}
.btn-endstop:hover{filter:brightness(.9)}
/* KPI shift cards */
.shift-kpis{display:flex;gap:8px;padding:10px 14px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0}
.skpi{flex:1;background:var(--bg);border-radius:var(--radius);padding:10px;border:1px solid var(--border);text-align:center}
.skpi.current{flex:2;background:var(--navy);color:#fff;border-color:var(--navy2);box-shadow:var(--shadow)}
.skpi.current .sk-lbl{color:rgba(255,255,255,.7)}.skpi.current .sk-val{color:#93c5fd;font-size:26px}.skpi.current .sk-sub{color:rgba(255,255,255,.75)}
.sk-lbl{font-size:9px;text-transform:uppercase;font-weight:700;color:var(--gray);letter-spacing:.7px;margin-bottom:3px}
.sk-val{font-size:20px;font-weight:800;color:var(--navy);line-height:1}.sk-sub{font-size:10px;color:var(--gray);margin-top:3px}
/* Merged table row types */
.row-prod td{background:#f0fdf4}.row-evt td{background:#fff7ed}
.row-prod:hover td,.row-evt:hover td{filter:brightness(.96)}
.row-tag{display:inline-block;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:700;text-transform:uppercase}
.tag-p{background:#bbf7d0;color:#166534}.tag-e{background:#fed7aa;color:#9a3412}.tag-n{background:#bfdbfe;color:#1e40af}

/* ── FIN DE POSTE ── */
#v-finposte{padding:0}
.fp-scroll{flex:1;overflow-y:auto;padding:14px}
.fp-top{text-align:center;padding-bottom:10px}
.fp-top h2{font-size:22px;font-weight:800;color:var(--navy)}
.fp-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px}
.fp-card{background:var(--card);border-radius:var(--radius);padding:12px;text-align:center;box-shadow:var(--shadow);border:1px solid var(--border)}
.fp-big{font-size:28px;font-weight:800;color:var(--navy)}
.fp-lbl{font-size:9px;text-transform:uppercase;color:var(--gray);font-weight:700;margin-top:3px}
.fp-acts{display:flex;justify-content:center;gap:10px;padding:14px 0}
.fp-tbl{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px}
.fp-tbl th{background:var(--navy);color:#fff;padding:5px 8px;font-size:9px;text-align:left;text-transform:uppercase}
.fp-tbl td{padding:5px 8px;border-bottom:1px solid var(--border)}

/* ── HISTORY ── */
#v-history{padding:0}

/* ── SETTINGS ── */
#v-settings{padding:0;overflow:hidden}
#settings-lock{flex:1;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px}
.lock-card{background:var(--card);border-radius:12px;padding:28px;width:100%;max-width:340px;text-align:center;box-shadow:var(--shadow)}
.lock-card h3{color:var(--navy);font-size:18px;font-weight:800;margin-bottom:8px}
.lock-card p{color:var(--gray);font-size:12px;margin-bottom:16px}
#v-settings-content{flex:1;overflow-y:auto;padding:14px;display:none}
.ss{background:var(--card);border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:var(--shadow);border:1px solid var(--border)}
.ss h3{font-size:12px;font-weight:700;margin-bottom:10px;color:var(--navy)}
.pr{display:flex;align-items:center;gap:6px;margin-bottom:6px;padding:5px;border-radius:5px;background:var(--bg)}
.pr .pn{font-weight:600;min-width:110px;font-size:11px}
.pr input{flex:1;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:12px}
.btn-eye{background:none;border:none;cursor:pointer;color:var(--gray);font-size:12px;padding:2px}
.day-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:6px}
.day-box{background:var(--bg);border-radius:4px;padding:4px;text-align:center}
.day-lbl{font-size:8px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px}
.day-box input{width:100%;padding:2px;border:1px solid var(--border);border-radius:3px;font-size:10px;text-align:center}
.model-card{border:1px solid var(--border);border-radius:7px;padding:10px;margin-bottom:8px}
.mch{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.mch input{flex:1;font-size:12px;font-weight:600;padding:4px 6px;border:1px solid var(--border);border-radius:4px}

/* ── MODALS ── */
.modal{display:none}
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:300;align-items:center;justify-content:center}
.overlay.on{display:flex}
.mbox{background:var(--card);border-radius:12px;width:90%;max-width:520px;box-shadow:0 20px 60px rgba(0,0,0,.3);max-height:92vh;display:flex;flex-direction:column}
.mbox.wide{max-width:860px}
.mhdr{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--border);flex-shrink:0}
.mhdr.red{background:#b91c1c;border-radius:12px 12px 0 0}
.mhdr h2{font-size:15px;font-weight:700;color:var(--navy)}
.mhdr.red h2,.mhdr.red button{color:#fff}
.mbody{flex:1;overflow-y:auto;padding:14px}
.mftr{display:flex;gap:8px;justify-content:flex-end;padding:12px 14px;border-top:1px solid var(--border);flex-shrink:0}
.btn{padding:7px 14px;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;gap:4px}
.btn:hover{filter:brightness(.9)}
.btn-prim{background:var(--navy);color:#fff}
.btn-danger{background:var(--red);color:#fff}
.btn-ok{background:var(--green);color:#fff}
.btn-sec{background:var(--lgray);color:var(--text)}
.btn-amber{background:var(--amber);color:#fff}
.btn-lg{font-size:14px;padding:10px 20px}

/* Stop modal */
.stop-section-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;margin:10px 0 5px;color:var(--navy)}
.stop-section-lbl:first-child{margin-top:0}
.stops-grid{display:grid;gap:5px;margin-bottom:6px}
.stops-grid.ratt{grid-template-columns:repeat(3,1fr)}
.stops-grid.pb{grid-template-columns:repeat(4,1fr)}
.stop-btn{border:none;border-radius:8px;padding:8px 5px;cursor:pointer;font-size:11px;font-weight:700;color:#fff;text-align:center;transition:all .12s;box-shadow:0 2px 0 rgba(0,0,0,.2)}
.stop-btn:active{transform:translateY(2px);box-shadow:none}
.stop-btn.ratt{background:#7c3aed}
.stop-btn.pb{background:#b91c1c}
.custom-row{display:flex;gap:6px;margin-top:4px}
.custom-row input{flex:1;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:12px;outline:none}
.custom-row input:focus{border-color:var(--navy)}

/* End prod modal */
.ep-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:12px}
.ep-stat{text-align:center;padding:10px;background:var(--bg);border-radius:6px}
.ep-stat .val{font-size:24px;font-weight:800;color:var(--navy)}
.ep-stat .lbl{font-size:9px;text-transform:uppercase;color:var(--gray);font-weight:700}
.ep-tbl{width:100%;border-collapse:collapse;font-size:11px}
.ep-tbl th{text-align:left;padding:4px 6px;background:var(--bg);font-size:9px;text-transform:uppercase;color:var(--gray)}
.ep-tbl td{padding:4px 6px;border-bottom:1px solid var(--border)}

/* Utils */
.hidden{display:none!important}
.flex{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.mt8{margin-top:8px}
.card{background:var(--card);border-radius:var(--radius);padding:12px;box-shadow:var(--shadow);border:1px solid var(--border)}

/* ── KPI VIEW ── */
#v-kpi{background:var(--bg)}
.kpi-prev-card{flex:1;background:var(--card);border:1px solid var(--border);border-radius:var(--radius);text-align:center;padding:6px 8px;min-width:0}

@media(max-width:900px){.form-3col{grid-template-columns:1fr 1fr}.recap-col{width:180px}}
@media(max-width:650px){.form-3col{grid-template-columns:1fr}.prod-body{flex-direction:column}.recap-col{width:100%}}
</style>
</head>
<body>

<!-- ════ LOGIN ════ -->
<div id="v-login" class="view on">
  <div class="login-card">
    <div class="lc-h1">⚙ KPI-ORC</div>
    <div class="lc-sub">Système de suivi de production</div>
    <div class="lf">
      <label>Pilote</label>
      <select id="ln-pilot"><option value="">-- Choisir --</option></select>
    </div>
    <div class="lf">
      <label>Modèle horaire (Poste)</label>
      <select id="ln-model" onchange="onLoginModelChange()"><option value="">-- Choisir --</option></select>
    </div>
    <!-- Horaires du jour -->
    <div id="ln-model-info" style="display:none;background:#f0f9ff;border:1px solid #bae6fd;border-radius:7px;padding:8px 10px;font-size:12px;margin-bottom:6px">
      <div style="color:#0369a1;font-weight:700;margin-bottom:3px">Horaires aujourd'hui :</div>
      <div id="ln-model-times" style="font-size:15px;font-weight:800;color:#0c4a6e;margin-bottom:5px"></div>
      <div id="ln-model-modify" style="display:none;margin-bottom:5px">
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <label style="font-size:11px;font-weight:600">Début <input type="time" id="ln-new-debut" style="padding:3px 6px;border:1.5px solid #bae6fd;border-radius:5px;font-size:12px"></label>
          <label style="font-size:11px;font-weight:600">Fin <input type="time" id="ln-new-fin" style="padding:3px 6px;border:1.5px solid #bae6fd;border-radius:5px;font-size:12px"></label>
        </div>
      </div>
      <button style="font-size:11px;padding:3px 8px;background:none;border:1px solid #0369a1;border-radius:5px;color:#0369a1;cursor:pointer" onclick="toggleLoginModelModify()">✏ Modifier les horaires d'aujourd'hui</button>
    </div>
    <div class="lf">
      <label>Mot de passe</label>
      <input type="password" id="ln-pw" placeholder="••••" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="btn-login" onclick="doLogin()">Valider</button>
    <div class="ln-err" id="ln-err"></div>
    <div style="margin-top:14px;text-align:center;font-size:11px;color:#94a3b8">
      Prod bloquée ? <a href="/reset" style="color:#dc2626;font-weight:700">Cliquer ici pour réinitialiser</a>
    </div>
  </div>
</div>

<!-- ════ APP ════ -->
<div id="app" class="hidden" style="display:none;flex:1;flex-direction:column;overflow:hidden">
  <div id="app-hdr">
    <div class="hdr-logo">⚙ KPI-ORC</div>
    <div class="hdr-tabs">
      <button class="htab on" id="ht-main" onclick="goTab('main')">Accueil</button>
      <button class="htab prod-on" id="ht-prod" onclick="goTab('prod')">▶ Prod en cours</button>
      <button class="htab" id="ht-hist" onclick="goTab('history')">Historique</button>
      <button class="htab" id="ht-kpi" onclick="goTab('kpi')">📊 KPI</button>
      <button class="htab" id="ht-cfg" onclick="goTab('settings')">Paramètres</button>
    </div>
    <div id="hdr-right">
      <span id="hdr-pilot-lbl"></span>
      <button class="btn-sm btn-ghost" onclick="doLogout()" style="font-size:11px">Déconnexion</button>
    </div>
  </div>
  <div id="alert-strip"></div>

  <!-- ════ MAIN VIEW ════ -->
  <div id="v-main" class="view" style="flex-direction:column">
    <!-- Bannière prod en cours (visible si prod_active mais sur vue accueil) -->
    <div id="main-prod-banner" style="display:none;background:#1e293b;color:#fff;padding:8px 14px;font-size:12px;align-items:center;gap:16px;cursor:pointer" onclick="goTab('prod')">
      <span style="font-weight:800;color:#86efac">▶ Prod en cours</span>
      <span>OF : <span id="mpb-of" style="font-weight:700">—</span></span>
      <span>Durée : <span id="mpb-dur" style="color:#67e8f9;font-weight:700">—</span></span>
      <span>Arrêts : <span id="mpb-stops" style="color:#fca5a5;font-weight:700">—</span></span>
      <span style="margin-left:auto;font-size:11px;opacity:.7">Cliquer → vue prod</span>
    </div>
    <!-- Barre Excel occupé -->
    <div id="excel-busy-bar" style="display:none;background:#92400e;color:#fef3c7;padding:5px 14px;font-size:11px;font-weight:700;text-align:center">
      ⚠ Fichier Excel ouvert par un autre programme — impossible de lire/écrire les données
    </div>
    <div class="main-hdr">
      <div class="mbtns" style="margin-left:0">
        <button class="btn btn-green" id="btn-start" onclick="doStartProd()" style="font-size:14px;padding:10px 18px;font-weight:800">▶ Démarrer production</button>
        <button class="btn btn-amber" onclick="doFinPoste()" style="font-size:14px;padding:10px 18px;font-weight:800">🏁 Fin de poste</button>
        <button class="btn btn-sec" onclick="loadMainDecl()" style="font-size:12px;padding:8px 14px">↺ Actualiser</button>
      </div>
    </div>
    <!-- KPI accueil — POSTE ACTUEL -->
    <div style="background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;padding:8px 10px;display:flex;gap:8px;align-items:stretch;flex-wrap:wrap">

      <!-- POSTE ACTUEL encart principal -->
      <div style="flex:3;min-width:280px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:8px 12px;display:flex;flex-direction:column;gap:6px">
        <!-- Titre + TRS jauge + valeur -->
        <div style="display:flex;align-items:center;gap:10px">
          <div style="flex-shrink:0;text-align:center">
            <svg viewBox="0 0 100 58" style="width:88px;display:block;margin:0 auto">
              <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="rgba(0,0,0,.12)" stroke-width="11" stroke-linecap="round"/>
              <path id="gauge-poste-acc-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="11" stroke-linecap="round" stroke-dasharray="0,132"/>
              <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#15803d" id="gauge-poste-acc-pct">—</text>
            </svg>
          </div>
          <div style="flex:1;min-width:0">
            <div style="font-size:10px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:.5px">TRS du Poste</div>
            <div style="font-size:10px;color:#64748b;margin-bottom:2px" id="gauge-poste-acc-lbl">—</div>
            <!-- Stats en ligne -->
            <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:2px">
              <div style="text-align:center">
                <div style="font-size:8px;font-weight:700;text-transform:uppercase;color:#64748b">Nombre OF</div>
                <div style="font-size:18px;font-weight:900;color:#1e40af;line-height:1" id="acc-nb-of">0</div>
              </div>
              <div style="text-align:center">
                <div style="font-size:8px;font-weight:700;text-transform:uppercase;color:#dc2626">Arrêts</div>
                <div style="font-size:18px;font-weight:900;color:#b91c1c;line-height:1" id="main-stat-arrets">0 min</div>
              </div>
              <div style="text-align:center">
                <div style="font-size:8px;font-weight:700;text-transform:uppercase;color:#16a34a">Prod</div>
                <div style="font-size:18px;font-weight:900;color:#15803d;line-height:1" id="main-stat-prod">0 min</div>
              </div>
            </div>
          </div>
        </div>
        <!-- Modèle horaire -->
        <div style="border-top:1px solid #bae6fd;padding-top:5px;display:flex;align-items:center;gap:5px;flex-wrap:wrap">
          <span style="font-size:10px;font-weight:700;color:#0369a1">Modèle :</span>
          <span style="font-size:12px;font-weight:800;color:#0c4a6e" id="main-model-times">—</span>
          <input type="time" id="main-model-debut" style="padding:2px 5px;border:1px solid #bae6fd;border-radius:4px;font-size:11px;color:#0c4a6e">
          <span style="font-size:11px;color:#0369a1">→</span>
          <input type="time" id="main-model-fin" style="padding:2px 5px;border:1px solid #bae6fd;border-radius:4px;font-size:11px;color:#0c4a6e">
          <button onclick="saveMainModelHours()" style="font-size:10px;padding:2px 7px;background:#0369a1;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:700">✓ Valider</button>
          <span style="font-size:10px;color:#64748b" id="main-ref-calc"></span>
        </div>
      </div>

      <!-- Poste précédent -->
      <div class="skpi" style="flex:1;min-width:100px">
        <div class="sk-lbl" id="kpi1-lbl">Poste précédent</div>
        <div class="sk-val" id="kpi1-trs">--%</div>
        <div class="sk-sub" id="kpi1-date" style="font-size:10px;opacity:.85"></div>
        <div class="sk-sub" id="kpi1-sub">0 OF</div>
      </div>
      <div class="skpi" style="flex:1;min-width:100px">
        <div class="sk-lbl" id="kpi2-lbl">Avant-dernier</div>
        <div class="sk-val" id="kpi2-trs">--%</div>
        <div class="sk-sub" id="kpi2-date" style="font-size:10px;opacity:.85"></div>
        <div class="sk-sub" id="kpi2-sub">0 OF</div>
      </div>

      <!-- IDs cachés compatibles JS existant -->
      <div style="display:none">
        <span id="kpi0-trs"></span><span id="kpi0-sub"></span>
        <span id="kpi0-date"></span><span id="kpi0-heure"></span>
        <span id="main-stat-pause"></span><span id="main-stat-nett"></span>
        <span id="main-stat-reunion"></span>
      </div>
    </div>
    <div class="table-wrap">
      <table class="ktbl">
        <thead><tr>
          <th>Type</th><th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th>
          <th>Début</th><th>Fin</th><th>Détails</th><th>Qté/Durée</th><th>TRS/Info</th><th>Actions</th>
        </tr></thead>
        <tbody id="main-body"></tbody>
      </table>
    </div>
  </div>

  <!-- ════ PRODUCTION VIEW ════ -->
  <div id="v-prod" class="view" style="flex-direction:column">
    <!-- Pilot/OF banner -->
    <div class="pob">
      <div class="pob-item">
        <div class="pob-lbl">Pilote</div>
        <div class="pob-val" id="pob-pilot">—</div>
      </div>
      <div class="pob-item of">
        <div class="pob-lbl">N° OF</div>
        <div class="pob-val" id="pob-of">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Poste</div>
        <div class="pob-val" style="font-size:16px" id="pob-poste">—</div>
      </div>
      <div class="pob-item" style="flex-direction:column;align-items:flex-start;gap:2px">
        <div class="pob-lbl">Modèle / Horaires</div>
        <div style="display:flex;align-items:center;gap:6px">
          <div class="pob-val" style="font-size:14px" id="pob-model">—</div>
          <button onclick="openPobModelEdit()" style="font-size:10px;padding:2px 6px;background:none;border:1px solid #94a3b8;border-radius:4px;cursor:pointer;color:#64748b">✏</button>
        </div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Départ OF</div>
        <div class="pob-val" style="font-size:16px;color:#fbbf24" id="pob-of-start">—</div>
      </div>
      <div class="pob-item trs">
        <div class="pob-lbl">TRS estimé</div>
        <div class="pob-val" id="pob-trs">—</div>
      </div>
    </div>
    <!-- Modal edit horaires depuis bandeau prod -->
    <div id="m-pobmodel" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:500;align-items:center;justify-content:center">
      <div class="mbox" style="max-width:340px;padding:20px">
        <div class="mhdr" style="margin:-20px -20px 14px;padding:14px 16px;border-radius:12px 12px 0 0"><h2>✏ Modifier horaires du poste</h2></div>
        <p id="pobm-info" style="font-size:12px;color:#64748b;margin-bottom:8px"></p>
        <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px">
          <label style="font-size:12px;font-weight:600">Début <input type="time" id="pobm-debut" style="padding:3px 6px;border:1.5px solid #cbd5e1;border-radius:5px;font-size:13px"></label>
          <label style="font-size:12px;font-weight:600">Fin <input type="time" id="pobm-fin" style="padding:3px 6px;border:1.5px solid #cbd5e1;border-radius:5px;font-size:13px"></label>
        </div>
        <div style="display:flex;gap:8px">
          <button class="btn btn-prim" onclick="savePobModelHours()">Enregistrer</button>
          <button class="btn btn-sec" onclick="closeM('m-pobmodel')">Annuler</button>
        </div>
      </div>
    </div>
    <!-- Status bar -->
    <div class="sbar">
      <div class="sc"><div class="sc-lbl">⏱ Durée OF</div><div class="sc-val green" id="sc-of">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">⛔ Arrêts</div><div class="sc-val red" id="sc-stops">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">⏸ Pauses</div><div class="sc-val amber" id="sc-pause">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">🎯 Pièces théo.</div><div class="sc-val" id="sc-theo" style="color:var(--blue)">—</div></div>
    </div>
    <!-- Body -->
    <div class="prod-body">
      <!-- LEFT: form (now full center, no act-col) -->
      <div class="form-col">
        <div class="form-3col">
          <!-- Zone Identification -->
          <div class="fzone zi">
            <h4>📋 Identification</h4>
            <div class="fr"><label>N° OF *</label><input id="f-of_num" oninput="scheduleAutoSave()"></div>
            <div class="fr ro"><label>Date</label><input id="f-date" readonly></div>
            <div class="fr ro"><label>Poste</label><input id="f-poste" readonly></div>
            <div class="fr ro"><label>Pilote</label><input id="f-pilote" readonly></div>
            <div class="fr"><label>Co-Pilote</label><select id="f-copilote" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Nb Personnes</label><input id="f-nb_pers" type="number" min="1" value="10" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Taille</label><select id="f-taille" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Code Produit</label><input id="f-code_prod" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Type Produit</label><select id="f-type_prod" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Kit</label><select id="f-kit" onchange="scheduleAutoSave()"><option value="">Non</option><option value="oui">Oui</option></select></div>
          </div>
          <!-- Zone Production -->
          <div class="fzone zp">
            <h4>🏭 Production</h4>
            <div class="fr big"><label>Qté Fabriquée *</label><input id="f-qte_fab" type="number" min="0" placeholder="0" oninput="scheduleAutoSave()"></div>
            <div class="fr big"><label>Qté Emballée</label><input id="f-qte_emb" type="number" min="0" placeholder="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Poids Garnissage (g)</label><input id="f-poids" type="number" min="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Fibre</label><select id="f-fibre" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>OF Taie</label><input id="f-of_taie" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Traca Fibre</label><input type="text" id="f-traca" oninput="scheduleAutoSave()" placeholder="n° de traca"></div>
            <div class="fr"><label>Réf Taie</label><input id="f-ref_taie" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq MP (min)</label><input id="f-duree_mq_mp" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Personnel (min)</label><input id="f-manquant_pers" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
          </div>
          <!-- Zone Qualité -->
          <div class="fzone zq">
            <h4>✅ Qualité</h4>
            <div class="fr"><label>Qté Init Taie</label><input id="f-qte_init_taie" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Taie 2nd Choix</label><input id="f-nb_taie2_choix" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Défaut Couture</label><input id="f-nb_def_cout" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Taie</label><input id="f-mq_taie" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Housse/Encart</label><input id="f-mq_housse_encart" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb PP Cousue</label><input id="f-nb_pp_cousue" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr comment-big"><label>💬 Commentaire</label><textarea id="f-comment" oninput="scheduleAutoSave()" placeholder="Commentaire libre…"></textarea></div>
          </div>
        </div>
        <!-- Timeline 4h -->
        <div class="tl-wrap">
          <h5>Timeline — 4 dernières heures</h5>
          <svg id="tl-svg" viewBox="0 0 800 40" preserveAspectRatio="none" style="width:100%;height:40px;display:block">
            <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
          </svg>
          <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
        </div>
        <!-- Action buttons row (below timeline) -->
        <div class="prod-act-row">
          <button class="act-btn act-stop" onclick="openStopModal()">⛔ Déclarer un arrêt</button>
          <button class="act-btn act-nett" onclick="doNettoyage()">🧹 Nettoyage</button>
          <button class="act-btn act-pause" id="btn-pause" onclick="doPause()">⏸ Pause</button>
          <button class="act-btn act-cancel" onclick="doCancelProd()">✖ Annuler prod</button>
          <button class="act-btn act-endprod" onclick="doEndProdPreview()">🏁 Fin d'OF/prod</button>
        </div>
      </div>
      <!-- RIGHT: recap + gauges + pie charts -->
      <div class="recap-col" style="width:310px">
        <div class="recap-hdr">Arrêts / pauses</div>
        <div class="recap-body" id="recap-list"></div>
        <!-- TRS OF gauge -->
        <div class="gauge-box" style="padding:8px 4px 4px">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;align-items:center">
            <div style="text-align:center">
              <svg viewBox="0 0 100 56" style="width:100%;max-width:110px">
                <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="10" stroke-linecap="round"/>
                <path id="gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="10" stroke-linecap="round" stroke-dasharray="0,1000"/>
                <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e" id="gauge-pct">—</text>
              </svg>
              <div class="gauge-lbl">TRS OF</div>
            </div>
            <div style="text-align:center">
              <svg viewBox="0 0 100 56" style="width:100%;max-width:110px">
                <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="10" stroke-linecap="round"/>
                <path id="gauge-poste-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#0891b2" stroke-width="10" stroke-linecap="round" stroke-dasharray="0,1000"/>
                <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#0c4a6e" id="gauge-poste-pct">—</text>
              </svg>
              <div class="gauge-lbl" id="gauge-poste-lbl">TRS Poste</div>
            </div>
          </div>
        </div>
        <!-- Pie charts -->
        <div style="padding:4px;border-top:1px solid var(--border);display:grid;grid-template-columns:1fr 1fr;gap:4px">
          <div style="text-align:center">
            <div style="font-size:8px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px">Poste — Prod/Arrêts</div>
            <svg id="pie-poste" viewBox="0 0 130 115" style="width:100%;height:auto;display:block"></svg>
          </div>
          <div style="text-align:center">
            <div style="font-size:8px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px">OF en cours</div>
            <svg id="pie-of" viewBox="0 0 130 115" style="width:100%;height:auto;display:block"></svg>
          </div>
        </div>
      </div>
    </div>
    <!-- Active stops bottom bar — chips -->
    <div id="stop-bottom">
      <div style="font-size:10px;opacity:.7;font-weight:700;white-space:nowrap">EN COURS :</div>
      <div id="stop-chips" style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;flex:1"></div>
    </div>
  </div>

  <!-- ════ FIN DE POSTE ════ -->
  <div id="v-finposte" class="view" style="flex-direction:column;overflow:hidden">
    <!-- Barre titre -->
    <div style="background:var(--navy);color:#fff;padding:7px 14px;flex-shrink:0;display:flex;align-items:center;justify-content:space-between">
      <div>
        <div style="font-size:15px;font-weight:800">🏁 Fin de poste</div>
        <div id="fp-who" style="font-size:11px;opacity:.8"></div>
      </div>
      <div style="text-align:right">
        <div style="font-size:14px;font-weight:800;color:#fbbf24" id="fp-horaire-display">—</div>
        <div id="fp-date" style="font-size:11px;font-weight:700;opacity:.8"></div>
      </div>
    </div>
    <!-- Graphiques + KPI (en haut, compact) -->
    <div style="display:flex;gap:12px;padding:10px 14px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;align-items:center;flex-wrap:wrap">
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">TRS Poste</div>
        <svg id="fp-gauge" viewBox="0 0 100 58" style="width:200px;display:block;margin:0 auto">
          <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
          <path id="fp-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
          <text x="50" y="46" text-anchor="middle" font-size="14" font-weight="800" fill="#1a1f5e" id="fp-gauge-pct">--%</text>
        </svg>
        <div style="font-size:18px;font-weight:800;color:var(--navy);margin-top:4px" id="fp-trs-lbl2">—</div>
        <div style="font-size:12px;color:var(--gray);margin-top:2px" id="fp-shift-hours">—</div>
      </div>
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">Répartition</div>
        <svg id="fp-pie" viewBox="0 0 130 115" style="width:200px;height:177px;display:block;margin:0 auto"></svg>
      </div>
      <div style="flex:1;display:grid;grid-template-columns:repeat(auto-fit,minmax(75px,1fr));gap:5px">
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:18px" id="fp-trs">--%</div><div class="fp-lbl">TRS Shift</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:18px" id="fp-trs-of">--%</div><div class="fp-lbl">TRS Prod</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:18px" id="fp-eq">0</div><div class="fp-lbl">Équivalence</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:18px" id="fp-nof">0</div><div class="fp-lbl">Nb OF</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:18px" id="fp-prod-t">0 min</div><div class="fp-lbl">Durée prod</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:18px" id="fp-stop-t">0 min</div><div class="fp-lbl">Arrêts</div></div>
      </div>
    </div>
    <!-- Timeline compact -->
    <div style="padding:5px 12px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:9px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:3px">Timeline du poste</div>
      <svg id="fp-tl" viewBox="0 0 800 42" preserveAspectRatio="none" style="width:100%;height:42px;display:block">
        <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
      </svg>
      <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
    </div>
    <!-- Corps défilant : productions + arrêts côte à côte -->
    <div style="flex:1;overflow-y:auto;padding:8px 12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div class="card" style="padding:8px">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:5px">Productions</div>
        <table class="fp-tbl" style="font-size:10px">
          <thead><tr><th>OF</th><th>Qté</th><th>Éq</th><th>Durée</th><th>TRS%</th></tr></thead>
          <tbody id="fp-prods"></tbody>
        </table>
      </div>
      <div class="card" style="padding:8px">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:5px">Arrêts du poste</div>
        <div id="fp-stops-list" style="font-size:11px"></div>
      </div>
    </div>
    <!-- Modèle horaire + recalcul TRS -->
    <div style="padding:6px 12px;background:var(--card);border-top:1px solid var(--border);flex-shrink:0;display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <span style="font-size:11px;font-weight:700;color:var(--navy)">Modèle horaire :</span>
      <select id="fp-model-sel" style="padding:4px 8px;border:1px solid var(--border);border-radius:5px;font-size:12px" onchange="recalcFPTRS()">
        <option value="">-- Choisir --</option>
      </select>
      <span id="fp-shift-info" style="font-size:11px;color:var(--gray)"></span>
      <button class="btn btn-ghost" style="font-size:11px;padding:4px 10px" onclick="openM('m-fp-horaires')">✏ Modifier horaires de mon poste</button>
    </div>
    <!-- Boutons -->
    <div style="padding:8px 12px;background:var(--card);border-top:1px solid var(--border);flex-shrink:0;display:flex;gap:10px;justify-content:flex-end">
      <button class="btn btn-sec" onclick="goTab('main')">← Retour</button>
      <button class="btn btn-danger btn-lg" onclick="confirmFinPoste()">⏹ Confirmer fin de poste &amp; Déconnexion</button>
    </div>
  </div>

  <!-- ════ MODAL : horaires de poste ════ -->
  <div id="m-fp-horaires" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:500;align-items:center;justify-content:center">
    <div class="card" style="width:340px;padding:20px;background:#fff;border-radius:12px">
      <div style="font-size:14px;font-weight:800;color:var(--navy);margin-bottom:14px">✏ Horaires de mon poste</div>
      <div style="font-size:11px;color:var(--gray);margin-bottom:10px">Ces horaires servent uniquement au calcul du TRS de poste (non sauvegardés).</div>
      <div class="lf" style="margin-bottom:10px">
        <label>Début de poste</label>
        <input type="datetime-local" id="fp-debut-dt" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:13px">
      </div>
      <div class="lf" style="margin-bottom:14px">
        <label>Fin de poste</label>
        <input type="datetime-local" id="fp-fin-dt" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:13px">
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="closeM('m-fp-horaires')">Annuler</button>
        <button class="btn btn-prim" onclick="applyFPHoraires()">✓ Appliquer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL PRÉ-POSTE (1er OF vs heure modèle) ════ -->
  <div id="m-preshift" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:601;align-items:center;justify-content:center">
    <div class="card" style="width:430px;padding:20px;background:#fff;border-radius:12px;border-top:4px solid var(--red)">
      <div style="font-size:15px;font-weight:800;color:var(--navy);margin-bottom:4px">⚠ Début de poste non déclaré</div>
      <div id="ps-text" style="font-size:13px;color:var(--red);font-weight:700;margin-bottom:10px"></div>
      <div style="font-size:12px;color:var(--gray);margin-bottom:14px">Que souhaitez-vous faire ?</div>
      <input type="hidden" id="ps-start-iso">
      <input type="hidden" id="ps-gap-s">
      <div style="display:flex;flex-direction:column;gap:8px">
        <button class="btn btn-prim" style="text-align:left;padding:10px 14px;font-size:13px" onclick="psChooseInterposte()">
          ⏱ Enregistrer comme temps d'arrêt Interposte<br>
          <span style="font-size:11px;font-weight:400;opacity:.85">Il n'y a pas eu de production pendant ce temps</span>
        </button>
        <button class="btn btn-green" style="text-align:left;padding:10px 14px;font-size:13px" onclick="psChooseBackdate()">
          ↩ Déclarer que cet OF a démarré à <span id="ps-backdate-time" style="font-weight:800">--h--</span> (début du poste)<br>
          <span id="ps-backdate-lbl" style="font-size:11px;font-weight:400;opacity:.85">L'OF sera rétro-daté à l'heure du modèle horaire</span>
        </button>
        <button class="btn btn-ghost" style="text-align:left;padding:10px 14px;font-size:13px" onclick="psShowModifyModel()">
          📅 Modifier les horaires de ce poste (aujourd'hui)<br>
          <span style="font-size:11px;font-weight:400;opacity:.75">Changer début et fin du modèle horaire, puis rétro-dater l'OF</span>
        </button>
        <!-- Formulaire inline modification modèle -->
        <div id="ps-model-form" style="display:none;background:#f8fafc;padding:10px;border-radius:7px;border:1px solid var(--border)">
          <div style="font-size:11px;font-weight:700;color:var(--navy);margin-bottom:7px">Nouvel horaire du poste :</div>
          <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
            <label style="font-size:11px">Début <input type="time" id="ps-new-debut" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:13px"></label>
            <label style="font-size:11px">Fin <input type="time" id="ps-new-fin" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:13px"></label>
            <button class="btn btn-prim" style="font-size:12px;padding:5px 14px" onclick="psConfirmModifyModel()">✓ Appliquer et démarrer</button>
          </div>
        </div>
        <button class="btn btn-ghost" style="font-size:12px" onclick="psChooseIgnore()">Ignorer (ne rien déclarer)</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL INTERPOSTE ════ -->
  <div id="m-interposte" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:420px;padding:20px;background:#fff;border-radius:12px;border-top:4px solid var(--amber)">
      <div style="font-size:15px;font-weight:800;color:var(--navy);margin-bottom:4px">⏱ Temps hors production</div>
      <div id="ip-duration" style="font-size:13px;color:var(--amber);font-weight:700;margin-bottom:12px"></div>
      <div style="font-size:12px;color:var(--gray);margin-bottom:10px">Que s'est-il passé pendant cette période ?</div>
      <div id="ip-btns" style="display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px"></div>
      <div style="margin-bottom:10px">
        <input id="ip-custom" placeholder="Ou saisir librement…" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:13px">
      </div>
      <div style="margin-bottom:10px">
        <input id="ip-comment" placeholder="Commentaire (optionnel)" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:12px">
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="skipInterposte()">Ignorer</button>
        <button class="btn btn-prim" onclick="confirmInterposte()">✓ Valider</button>
      </div>
    </div>
  </div>

  <!-- ════ HISTORY ════ -->
  <div id="v-history" class="view" style="flex-direction:column;overflow:hidden">
    <div style="background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap">
      <span style="font-size:12px;font-weight:700;color:var(--navy)">Historique</span>
      <label style="font-size:11px;font-weight:600;color:var(--gray)">Du <input type="date" id="hist-from" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:12px;margin-left:4px"></label>
      <label style="font-size:11px;font-weight:600;color:var(--gray)">Au <input type="date" id="hist-to" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:12px;margin-left:4px"></label>
      <button class="btn btn-primary" onclick="loadHist()" style="padding:5px 12px;font-size:12px">Charger</button>
    </div>
    <div style="flex:1;overflow-y:auto">
      <table class="ktbl"><thead><tr id="hist-hd"></tr></thead><tbody id="hist-bd"></tbody></table>
    </div>
  </div>

  <!-- ════ KPI VIEW ════ -->
  <div id="v-kpi" class="view" style="flex-direction:column;overflow:hidden;background:#f1f5f9">
    <!-- Barre titre -->
    <div style="background:#fff;color:#1e293b;padding:6px 16px;flex-shrink:0;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #e2e8f0">
      <div style="font-size:16px;font-weight:800;letter-spacing:.5px;color:#1e3a8a">📊 KPI — Vue d'ensemble</div>
      <button style="font-size:12px;padding:4px 12px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:6px;color:#475569;cursor:pointer" onclick="loadKPI()">↺ Actualiser</button>
    </div>
    <!-- Corps principal : 2 colonnes -->
    <div style="display:grid;grid-template-columns:1fr 380px;flex:1;overflow:hidden;min-height:0;gap:0">
      <!-- Colonne gauche : Jauges TRS + 4 Timelines -->
      <div style="display:flex;flex-direction:column;overflow:hidden;border-right:1px solid #e2e8f0">
        <!-- Jauges TRS : poste actuel + 3 précédents -->
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:0;flex-shrink:0;border-bottom:1px solid #e2e8f0">
          <!-- Jauge poste actuel (plus grande) -->
          <div style="background:linear-gradient(135deg,#dbeafe,#eff6ff);padding:14px 10px;text-align:center;border-right:1px solid #e2e8f0">
            <div style="font-size:11px;text-transform:uppercase;font-weight:700;color:#3b82f6;letter-spacing:1px;margin-bottom:6px">Poste actuel</div>
            <svg viewBox="0 0 120 70" style="width:110px;display:block;margin:0 auto">
              <path d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="rgba(0,0,0,.08)" stroke-width="14" stroke-linecap="round"/>
              <path id="kpi-g0-arc" d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#16a34a" stroke-width="14" stroke-linecap="round" stroke-dasharray="0,157"/>
              <text x="60" y="58" text-anchor="middle" font-size="20" font-weight="800" fill="#16a34a" id="kpi-g0-pct">--%</text>
            </svg>
            <div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:4px" id="kpi-cur-date">—</div>
            <div style="font-size:12px;color:#475569;margin-top:2px" id="kpi-cur-sub">0 OF</div>
          </div>
          <!-- 3 jauges précédentes -->
          <div id="kpi-p1-card" style="background:#fff;padding:12px 8px;text-align:center;border-right:1px solid #e2e8f0">
            <div style="font-size:10px;text-transform:uppercase;font-weight:700;color:#64748b;letter-spacing:.8px;margin-bottom:4px" id="kpi-p1-lbl">Poste précédent</div>
            <svg viewBox="0 0 120 70" style="width:90px;display:block;margin:0 auto">
              <path d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#e2e8f0" stroke-width="14" stroke-linecap="round"/>
              <path id="kpi-g1-arc" d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#16a34a" stroke-width="14" stroke-linecap="round" stroke-dasharray="0,157"/>
              <text x="60" y="58" text-anchor="middle" font-size="18" font-weight="800" fill="#1e293b" id="kpi-g1-pct">--%</text>
            </svg>
            <div style="font-size:10px;color:#64748b;margin-top:3px" id="kpi-p1-date"></div>
            <div style="font-size:11px;color:#94a3b8;margin-top:1px" id="kpi-p1-sub">—</div>
          </div>
          <div id="kpi-p2-card" style="background:#fff;padding:12px 8px;text-align:center;border-right:1px solid #e2e8f0">
            <div style="font-size:10px;text-transform:uppercase;font-weight:700;color:#64748b;letter-spacing:.8px;margin-bottom:4px" id="kpi-p2-lbl">Avant-dernier</div>
            <svg viewBox="0 0 120 70" style="width:90px;display:block;margin:0 auto">
              <path d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#e2e8f0" stroke-width="14" stroke-linecap="round"/>
              <path id="kpi-g2-arc" d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#16a34a" stroke-width="14" stroke-linecap="round" stroke-dasharray="0,157"/>
              <text x="60" y="58" text-anchor="middle" font-size="18" font-weight="800" fill="#1e293b" id="kpi-g2-pct">--%</text>
            </svg>
            <div style="font-size:10px;color:#64748b;margin-top:3px" id="kpi-p2-date"></div>
            <div style="font-size:11px;color:#94a3b8;margin-top:1px" id="kpi-p2-sub">—</div>
          </div>
          <div id="kpi-p3-card" style="background:#fff;padding:12px 8px;text-align:center">
            <div style="font-size:10px;text-transform:uppercase;font-weight:700;color:#64748b;letter-spacing:.8px;margin-bottom:4px" id="kpi-p3-lbl">Il y a 3 postes</div>
            <svg viewBox="0 0 120 70" style="width:90px;display:block;margin:0 auto">
              <path d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#e2e8f0" stroke-width="14" stroke-linecap="round"/>
              <path id="kpi-g3-arc" d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="#16a34a" stroke-width="14" stroke-linecap="round" stroke-dasharray="0,157"/>
              <text x="60" y="58" text-anchor="middle" font-size="18" font-weight="800" fill="#1e293b" id="kpi-g3-pct">--%</text>
            </svg>
            <div style="font-size:10px;color:#64748b;margin-top:3px" id="kpi-p3-date"></div>
            <div style="font-size:11px;color:#94a3b8;margin-top:1px" id="kpi-p3-sub">—</div>
          </div>
        </div>
        <!-- 4 Timelines -->
        <div style="flex:1;overflow-y:auto;display:flex;flex-direction:column;padding:10px 12px;gap:10px;min-height:0;background:#f8fafc">
          <!-- Timeline poste actuel -->
          <div style="flex-shrink:0">
            <div style="font-size:11px;font-weight:700;color:#1e40af;text-transform:uppercase;letter-spacing:.7px;margin-bottom:5px" id="kpi-tl0-lbl">Poste actuel</div>
            <svg id="kpi-tl0" viewBox="0 0 800 50" preserveAspectRatio="none" style="width:100%;height:72px;display:block;border-radius:6px">
              <rect x="0" y="0" width="800" height="50" fill="#e2e8f0" rx="4"/>
            </svg>
          </div>
          <!-- Timeline poste N-1 -->
          <div style="flex-shrink:0">
            <div style="font-size:10px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px" id="kpi-tl1-lbl">—</div>
            <svg id="kpi-tl1" viewBox="0 0 800 36" preserveAspectRatio="none" style="width:100%;height:48px;display:block;border-radius:4px">
              <rect x="0" y="0" width="800" height="36" fill="#e2e8f0" rx="4"/>
            </svg>
            <div id="kpi-p1-pareto" style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;min-height:14px"></div>
          </div>
          <!-- Timeline poste N-2 -->
          <div style="flex-shrink:0">
            <div style="font-size:10px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px" id="kpi-tl2-lbl">—</div>
            <svg id="kpi-tl2" viewBox="0 0 800 36" preserveAspectRatio="none" style="width:100%;height:48px;display:block;border-radius:4px">
              <rect x="0" y="0" width="800" height="36" fill="#e2e8f0" rx="4"/>
            </svg>
            <div id="kpi-p2-pareto" style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;min-height:14px"></div>
          </div>
          <!-- Timeline poste N-3 -->
          <div style="flex-shrink:0">
            <div style="font-size:10px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px" id="kpi-tl3-lbl">—</div>
            <svg id="kpi-tl3" viewBox="0 0 800 36" preserveAspectRatio="none" style="width:100%;height:48px;display:block;border-radius:4px">
              <rect x="0" y="0" width="800" height="36" fill="#e2e8f0" rx="4"/>
            </svg>
            <div id="kpi-p3-pareto" style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;min-height:14px"></div>
          </div>
          <!-- Légende -->
          <div style="display:flex;gap:14px;flex-shrink:0;font-size:11px;color:#475569">
            <span style="display:flex;align-items:center;gap:4px"><i style="display:inline-block;width:14px;height:10px;border-radius:2px;background:#dc2626"></i>PB/Panne</span>
            <span style="display:flex;align-items:center;gap:4px"><i style="display:inline-block;width:14px;height:10px;border-radius:2px;background:#3b82f6"></i>Organisation</span>
            <span style="display:flex;align-items:center;gap:4px"><i style="display:inline-block;width:14px;height:10px;border-radius:2px;background:#f59e0b"></i>Nettoyage</span>
            <span style="display:flex;align-items:center;gap:4px"><i style="display:inline-block;width:14px;height:10px;border-radius:2px;background:#94a3b8"></i>Pause</span>
            <span style="display:flex;align-items:center;gap:4px"><i style="display:inline-block;width:14px;height:10px;border-radius:2px;background:#16a34a"></i>Prod</span>
          </div>
        </div>
      </div>
      <!-- Colonne droite : Camembert + Stats + Pareto -->
      <div style="display:flex;flex-direction:column;overflow:hidden;gap:0;background:#fff">
        <!-- Camembert -->
        <div style="background:#f8fafc;padding:12px 16px;flex-shrink:0;border-bottom:1px solid #e2e8f0;text-align:center">
          <div style="font-size:11px;text-transform:uppercase;font-weight:700;color:#475569;letter-spacing:.8px;margin-bottom:6px">Prod / Arrêts — poste actuel</div>
          <svg id="kpi-pie" viewBox="0 0 130 115" style="width:150px;height:130px;display:block;margin:0 auto"></svg>
        </div>
        <!-- Stats chiffres clés -->
        <div id="kpi-stats" style="background:#fff;padding:10px 12px;flex-shrink:0;border-bottom:1px solid #e2e8f0;display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px"></div>
        <!-- Pareto -->
        <div style="flex:1;overflow:hidden;display:flex;flex-direction:column;padding:12px 16px;background:#fff">
          <div style="font-size:11px;text-transform:uppercase;font-weight:700;color:#475569;letter-spacing:.8px;margin-bottom:8px;flex-shrink:0">Pareto arrêts</div>
          <div id="kpi-pareto" style="flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:8px"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- ════ SETTINGS ════ -->
  <div id="v-settings" class="view" style="flex-direction:column">
    <div id="settings-lock">
      <div class="lock-card">
        <h3>🔒 Paramètres</h3>
        <p>Entrez le mot de passe administrateur pour accéder aux paramètres.</p>
        <input type="password" id="lock-pw" placeholder="MDP admin" style="width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:14px;outline:none;margin-bottom:10px" onkeydown="if(event.key==='Enter')unlockSettings()">
        <button onclick="unlockSettings()" style="width:100%;padding:10px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:14px;font-weight:700;cursor:pointer">Déverrouiller</button>
        <div id="lock-err" style="color:#dc2626;margin-top:6px;font-size:12px;text-align:center;min-height:14px"></div>
      </div>
    </div>
    <div id="v-settings-content">
      <div class="ss">
        <h3>📂 Fichier Excel de données</h3>
        <div style="font-size:11px;color:var(--gray);margin-bottom:8px">Chemin complet vers le fichier Excel (.xlsx) contenant les onglets Declarations et Listes.</div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:6px">
          <input id="cfg-db-path" placeholder="C:\chemin\vers\fichier.xlsx" style="flex:1;min-width:200px;padding:7px 10px;border:1.5px solid var(--border);border-radius:5px;font-size:12px">
          <button class="btn btn-prim" onclick="setDbPath()">💾 Enregistrer</button>
        </div>
        <div id="cfg-db-status" style="font-size:11px;color:var(--gray)"></div>
      </div>
      <div class="ss">
        <h3>🔐 Mots de passe pilotes</h3>
        <div id="pwd-list"></div>
        <div class="flex mt8">
          <input id="np-name" placeholder="Nom pilote" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:12px">
          <input id="np-pw" type="password" placeholder="MDP" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:12px">
          <button class="btn btn-prim" onclick="addPilot()">+ Ajouter</button>
        </div>
        <div class="flex mt8">
          <button class="btn btn-ok" onclick="savePwds()">💾 Enregistrer MDP</button>
        </div>
      </div>
      <div class="ss">
        <h3>🕐 Modèles horaires (par poste &amp; par jour)</h3>
        <div id="models-list"></div>
        <button class="btn btn-sec mt8" onclick="addModel()">+ Nouveau modèle</button>
        <div class="flex mt8">
          <button class="btn btn-ok" onclick="saveModels()">💾 Enregistrer modèles</button>
        </div>
      </div>
      <div class="ss">
        <h3>⚙ Référence production 8h</h3>
        <div class="flex">
          <label style="font-weight:600;font-size:12px">Prod ref (unités/8h):</label>
          <input id="cfg-pr" type="number" style="width:90px;padding:5px;border:1px solid var(--border);border-radius:5px">
          <button class="btn btn-ok" onclick="saveProdRef()">Enregistrer</button>
        </div>
      </div>
      <div class="ss">
        <h3>📊 Dashboard HTML superviseur</h3>
        <div class="flex">
          <button class="btn btn-prim" onclick="generateDashboard()">🔄 Générer le Dashboard HTML</button>
          <span id="dash-status" style="font-size:11px;color:var(--gray);margin-left:8px"></span>
        </div>
      </div>
      <div class="ss">
        <h3>⛔ Liste des arrêts configurables</h3>
        <div style="font-size:11px;color:var(--gray);margin-bottom:8px">Ajouter, modifier ou supprimer les boutons d'arrêt disponibles en production. Pris en compte immédiatement.</div>
        <div id="events-list-ui" style="margin-bottom:10px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;background:#f8fafc;padding:8px;border-radius:7px;border:1px solid var(--border)">
          <input id="ev-new-label" placeholder="Nom de l'arrêt" style="flex:1;min-width:120px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:12px">
          <select id="ev-new-cat" style="padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:12px">
            <option value="pb">🔴 Panne</option>
            <option value="ratt">🟠 Rattrapage</option>
            <option value="nettoyage">🟡 Nettoyage</option>
            <option value="organisation">🔵 Organisation</option>
            <option value="autre">⚫ Autre</option>
          </select>
          <button class="btn btn-green" style="font-size:11px;padding:5px 12px" onclick="addEvtItem()">+ Ajouter</button>
        </div>
        <button class="btn btn-prim" style="margin-top:8px;font-size:12px" onclick="saveEvtList()">💾 Enregistrer la liste</button>
      </div>
      <div class="ss">
        <h3>🔄 Labels "Entre 2 OFs" (interposte)</h3>
        <div style="font-size:11px;color:var(--gray);margin-bottom:8px">Boutons de choix affichés dans la fenêtre "Temps entre 2 OFs". Configurez ici vos motifs d'interposte.</div>
        <div id="interposte-list-ui" style="margin-bottom:10px"></div>
        <div style="display:flex;gap:6px;align-items:center;background:#f8fafc;padding:8px;border-radius:7px;border:1px solid var(--border)">
          <input id="ip-new-label" placeholder="Nouveau label interposte" style="flex:1;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:12px">
          <button class="btn btn-green" style="font-size:11px;padding:5px 12px" onclick="addInterposteLbl()">+ Ajouter</button>
        </div>
        <button class="btn btn-prim" style="margin-top:8px;font-size:12px" onclick="saveInterposteCfg()">💾 Enregistrer</button>
      </div>
    </div>
  </div>

</div><!-- /app -->

<!-- ════ MODAL: Déclarer un arrêt ════ -->
<div class="overlay" id="m-stop">
  <div class="mbox">
    <div class="mhdr red">
      <h2>⛔ Déclarer un arrêt</h2>
      <button style="background:none;border:none;cursor:pointer;color:#fff;font-size:16px" onclick="closeM('m-stop')">✕</button>
    </div>
    <div class="mbody">
      <div class="stop-section-lbl">🔄 Rattrapage</div>
      <div class="stops-grid ratt" id="sgrid-ratt"></div>
      <div class="stop-section-lbl">🔧 PB Technique</div>
      <div class="stops-grid pb" id="sgrid-pb"></div>
      <div class="stop-section-lbl" id="sgrid-nett-lbl" style="display:none">🟡 Nettoyage</div>
      <div class="stops-grid" id="sgrid-nettoyage" style="display:none"></div>
      <div class="stop-section-lbl" id="sgrid-org-lbl" style="display:none">🔵 Organisation</div>
      <div class="stops-grid" id="sgrid-organisation" style="display:none"></div>
      <div class="stop-section-lbl">✏ Arrêt libre / autre</div>
      <div class="custom-row">
        <input id="custom-stop-input" placeholder="Nom de l'arrêt…" maxlength="60">
        <button class="btn btn-amber" onclick="declareCustomStop()">Déclarer</button>
      </div>
    </div>
    <div class="mftr"><button class="btn btn-sec" onclick="closeM('m-stop')">Annuler</button></div>
  </div>
</div>

<!-- ════ MODAL: Fin d'OF/prod ════ -->
<div class="overlay" id="m-endprod">
  <div class="mbox wide">
    <div class="mhdr"><h2 id="ep-title">⏹ Fin d'OF/prod</h2></div>
    <div class="mbody">
      <div style="display:flex;gap:10px;margin-bottom:8px;align-items:flex-start;flex-wrap:wrap">
        <div style="flex:1;min-width:180px"><div class="ep-grid" id="ep-stats"></div></div>
        <div style="display:flex;gap:8px;flex-shrink:0">
          <div style="text-align:center">
            <div style="font-size:9px;text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:2px">Répartition</div>
            <svg id="ep-pie" viewBox="0 0 130 115" style="width:110px;height:97px;display:block"></svg>
          </div>
          <div style="text-align:center">
            <div style="font-size:9px;text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:2px">TRS</div>
            <svg id="ep-gauge" viewBox="0 0 100 58" style="width:90px;display:block;margin:0 auto;margin-top:8px">
              <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
              <path id="ep-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
              <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e" id="ep-gauge-pct">--%</text>
            </svg>
          </div>
        </div>
      </div>
      <div style="margin-bottom:10px">
        <div style="font-size:10px;text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:5px">Arrêts &amp; pauses</div>
        <table class="ep-tbl"><thead><tr><th>Type</th><th>Durée</th><th>%</th></tr></thead><tbody id="ep-stops"></tbody></table>
      </div>
      <div class="card" style="padding:8px;margin-bottom:0">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">Timeline</div>
        <svg id="ep-tl" viewBox="0 0 800 40" preserveAspectRatio="none" style="width:100%;height:40px;display:block">
          <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
        </svg>
        <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span></div>
      </div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-endprod')">Annuler</button>
      <button class="btn btn-danger btn-lg" onclick="confirmEndProd()">✓ Confirmer fin de production</button>
    </div>
  </div>
</div>

<!-- ════ MODAL: Éditer arrêt ════ -->
<div class="overlay" id="m-editstop">
  <div class="mbox">
    <div class="mhdr"><h2>✏ Modifier l'arrêt</h2></div>
    <div class="mbody">
      <input type="hidden" id="es-key">
      <div class="fr" style="margin-bottom:8px"><label>Type</label><select id="es-type"></select></div>
      <div class="fr" style="margin-bottom:8px"><label>Heure début</label><input type="time" id="es-deb" step="60"></div>
      <div class="fr" style="margin-bottom:8px"><label>Heure fin</label><input type="time" id="es-fin" step="60"></div>
      <div class="fr"><label>Commentaire</label><input type="text" id="es-cmt"></div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-editstop')">Annuler</button>
      <button class="btn btn-danger" onclick="deleteStop()">🗑 Supprimer</button>
      <button class="btn btn-ok" onclick="saveEditStop()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<!-- ════ MODAL: Commentaire fin d'arrêt ════ -->
<div class="overlay" id="m-stopcmt">
  <div class="mbox" style="max-width:400px">
    <div class="mhdr red"><h2>✓ Terminer l'arrêt</h2></div>
    <div class="mbody">
      <input type="hidden" id="cmt-stop-key">
      <div id="cmt-stop-lbl" style="font-size:13px;font-weight:700;color:var(--navy);margin-bottom:10px"></div>
      <div class="fr"><label>Commentaire (optionnel)</label><textarea id="cmt-stop-text" style="height:70px;resize:none;width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:13px" placeholder="Description de l'arrêt…"></textarea></div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-stopcmt')">Annuler</button>
      <button class="btn btn-ok btn-lg" onclick="confirmEndStop()">✓ Confirmer fin d'arrêt</button>
    </div>
  </div>
</div>

<!-- ════ MODAL: Éditer ligne ════ -->
<div class="overlay" id="m-editrow">
  <div class="mbox wide">
    <div class="mhdr">
      <h2 id="er-title">✏ Modifier la ligne</h2>
      <button style="background:none;border:none;cursor:pointer;font-size:16px" onclick="closeM('m-editrow')">✕</button>
    </div>
    <div class="mbody">
      <input type="hidden" id="er-rownum"><input type="hidden" id="er-rowtype">
      <div id="er-pw-row" style="display:none;margin-bottom:8px;padding:8px 10px;background:#fef3c7;border:1px solid #fbbf24;border-radius:6px">
        <div class="fr" style="max-width:220px"><label>🔒 Mot de passe Admin</label><input type="password" id="er-pw" placeholder="••••" onkeydown="if(event.key==='Enter')saveEditRow()"></div>
        <div style="font-size:10px;color:#92400e;margin-top:3px">Entrez le MDP puis cliquez à nouveau sur l'action.</div>
      </div>
      <div id="er-prod-fields">
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:6px">
          <div class="fr"><label>N° OF</label><input id="er-of"></div>
          <div class="fr"><label>Heure Début</label><input type="time" id="er-deb" step="60"></div>
          <div class="fr"><label>Heure Fin</label><input type="time" id="er-fin" step="60"></div>
          <div class="fr"><label>Poste</label><input id="er-poste"></div>
          <div class="fr"><label>Pilote</label><input id="er-pilote"></div>
          <div class="fr"><label>Co-Pilote</label><input id="er-copilote"></div>
          <div class="fr"><label>Nb Personnes</label><input type="number" id="er-nbpers"></div>
          <div class="fr"><label>Taille</label><select id="er-taille"><option value="">--</option></select></div>
          <div class="fr"><label>Code Produit</label><input id="er-codeprod"></div>
          <div class="fr"><label>Type Produit</label><select id="er-typeprod"><option value="">--</option></select></div>
          <div class="fr"><label>Qté Fabriquée</label><input type="number" id="er-qtefab"></div>
          <div class="fr"><label>Qté Emballée</label><input type="number" id="er-qteemb"></div>
          <div class="fr"><label>Poids Garnissage (g)</label><input type="number" id="er-poids"></div>
          <div class="fr"><label>Fibre</label><select id="er-fibre"><option value="">--</option></select></div>
          <div class="fr"><label>OF Taie</label><input id="er-oftaie"></div>
          <div class="fr"><label>Traca Fibre</label><select id="er-traca"><option value="">--</option></select></div>
          <div class="fr"><label>Réf Taie</label><input id="er-reftaie"></div>
          <div class="fr"><label>Kit</label><select id="er-kit"><option value="">Non</option><option value="oui">Oui</option></select></div>
          <div class="fr"><label>Qté Init Taie</label><input type="number" id="er-qteinit"></div>
          <div class="fr"><label>Nb Taie 2nd Choix</label><input type="number" id="er-nbtaie2"></div>
          <div class="fr"><label>Nb Défaut Couture</label><input type="number" id="er-nbdef"></div>
          <div class="fr"><label>Mq Taie</label><input type="number" id="er-mqtaie"></div>
          <div class="fr"><label>Mq Housse/Encart</label><input type="number" id="er-mqhousse"></div>
          <div class="fr"><label>Nb PP Cousue</label><input type="number" id="er-nbpp"></div>
          <div class="fr"><label>Duree MQ MP (min)</label><input type="number" id="er-dureemq"></div>
          <div class="fr"><label>Manquant Personnel (min)</label><input type="number" id="er-manqpers"></div>
        </div>
        <div class="fr comment-big"><label>💬 Commentaire</label><textarea id="er-comment-prod" style="height:60px;resize:none;width:100%;padding:4px 6px;border:2px solid #f59e0b;border-radius:4px;font-size:12px;background:#fffbeb"></textarea></div>
      </div>
      <div id="er-evt-fields" style="display:none">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:6px">
          <div class="fr"><label>Type d'arrêt</label><select id="er-evttype"><option value="">--</option></select></div>
          <div class="fr"><label>N° OF</label><input id="er-evtof"></div>
          <div class="fr"><label>Heure Début</label><input type="time" id="er-evtdeb" step="60"></div>
          <div class="fr"><label>Heure Fin</label><input type="time" id="er-evtfin" step="60"></div>
        </div>
        <div class="fr"><label>Commentaire</label><textarea id="er-evtcomment" style="height:44px;resize:none;width:100%;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:12px"></textarea></div>
        <div class="fr" style="margin-top:6px"><label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:600;font-size:12px"><input type="checkbox" id="er-horstrs"> Hors TRS</label></div>
      </div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-editrow')">Annuler</button>
      <button class="btn btn-danger" onclick="deleteRow(null,'er-rownum')">🗑 Supprimer</button>
      <button class="btn btn-ok" onclick="saveEditRow()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<script>
// ── EVENTS definition (matches Python EVENTS list) ──
const EVENTS = [
  ["Pochon / Fibre","ratt_pochon","ratt"],
  ["Couture","ratt_couture","ratt"],
  ["Emballage","ratt_emb","ratt"],
  ["Presse Souder","ratt_presse_soud","ratt"],
  ["Presse ZIP","ratt_presse_zip","ratt"],
  ["Réunion","arret_reunion","ratt"],
  ["Chargeuse","pb_chargeuse","pb"],
  ["Carde","pb_carde","pb"],
  ["Étaleur / Tour","pb_etaleur","pb"],
  ["Coupe / Circ.","pb_coupe","pb"],
  ["Tapis Bascule","pb_tapis1","pb"],
  ["Enrouleur Pochon","pb_enrouleur","pb"],
  ["Pesée / Tapis 2","pb_pesee","pb"],
  ["Déviation / Table","pb_deviation","pb"],
  ["Enfileur Pochon","pb_enfileur","pb"],
  ["Kinna / Stroebel","pb_kinna","pb"],
  ["Tapeuse","pb_tapeuse","pb"],
  ["Table Rot. / Twin","pb_table_rot","pb"],
  ["Enfileuse H100","pb_h100","pb"],
  ["Enfileuse Traversin","pb_traversin","pb"],
  ["Presse ORC","pb_presse_orc","pb"],
  ["Presse Housse ZIP","pb_presse_zip2","pb"],
  ["Cercleuse","pb_cercleuse","pb"],
  ["Enrouleuse Traversin","pb_enrouleuse","pb"],
  ["Matière première","arret_mp","pb"],
];

const STOP_COL = {
  ratt:"#dc2626",pb:"#dc2626",autre:"#64748b",
  nettoyage:"#f59e0b",
  organisation:"#3b82f6",
  "_pause":"#94a3b8","Pause":"#94a3b8"
};

function getStopColor(key, cat) {
  if (cat) return STOP_COL[cat]||'#64748b';
  if (key==='nettoyage') return STOP_COL.nettoyage;
  if (key==='_pause'||key==='Pause') return STOP_COL['Pause'];
  return '#94a3b8';
}

const FORM_FIELDS = ["of_num","copilote","nb_pers","taille","code_prod","type_prod","poids","fibre","of_taie","traca","ref_taie","kit","qte_fab","qte_emb","qte_init_taie","nb_taie2_choix","nb_def_cout","mq_taie","mq_housse_encart","nb_pp_cousue","duree_mq_mp","manquant_pers","comment"];

// ── State ──
let ST = {};
let gEvts = [];
let _curStopKey = null;
let _curStopElap = 0;
let _ofElapAtPoll = 0;
let _stopWallAtPoll = 0;
let _pauseTotalAtPoll = 0;
let _lastPoll = Date.now();
let _ticker = null;
let _autoSaveTimer = null;
let _curTab = 'main';
let _cfgPwds = {};
let _cfgModels = [];
let _settingsUnlocked = false;
let _adminPw = '';
window._evMap = {};

// ── INIT ──
document.addEventListener('DOMContentLoaded', async () => {
  await loadEvtsList(); // charge la liste dynamique des arrêts avant de construire les grilles
  await loadInterposteCfg();
  buildEditStopOpts();
  await loadLists();
  const s = await apiFetch('/api/state');
  if (s && (s.pilot || s.prod_active)) {
    document.getElementById('v-login').classList.remove('on');
    showApp(s);
    if (s.prod_active && !s.pilot) {
      toast('Prod en cours restaurée — vérifiez les infos','warn');
    }
  }
  setInterval(pollState, 5000);
  setInterval(pollEvts, 8000);
  setInterval(loadLists, 300000); // sync Excel lists every 5 min
  startTicker();
});

async function loadLists() {
  const d = await apiFetch('/api/lists');
  if (!d) return;
  popSel('f-taille', d.tailles||[]);
  popSel('f-type_prod', d.types_prod||[]);
  popSel('f-fibre', d.fibres||[]);
  popSel('f-copilote', d.copilotes||[]);
  popSel('er-taille', d.tailles||[]);
  popSel('er-typeprod', d.types_prod||[]);
  popSel('er-fibre', d.fibres||[]);
  popSel('er-traca', d.tracas||[]);
  // Build equivalence coef map: type_prod -> coef
  const eqs=d.equivalences||[]; const tps=d.types_prod||[];
  window._equivCoefs={};
  tps.forEach((t,i)=>{if(i<eqs.length&&eqs[i])window._equivCoefs[t]=parseFloat(String(eqs[i]).replace(',','.'))||1;});
  const pil = d.pilotes||[];
  const sel = document.getElementById('ln-pilot');
  // Clear existing options (except first placeholder)
  while(sel.options.length>1) sel.remove(1);
  pil.forEach(p => { const o=document.createElement('option'); o.value=p; o.textContent=p; sel.appendChild(o); });
  // If no pilots yet, retry after 1.5s (Excel may still be loading)
  if(!pil.length){
    setTimeout(async()=>{
      const d2=await apiFetch('/api/lists');
      if(!d2) return;
      const pil2=d2.pilotes||[];
      while(sel.options.length>1) sel.remove(1);
      pil2.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;sel.appendChild(o);});
    },1500);
  }
  // Load models for login select
  await loadModelsForLogin();
}

async function loadModelsForLogin() {
  const d = await apiFetch('/api/config');
  if (!d) return;
  _cfgModels = d.modeles_horaires||[];
  _cfgPwds = d.pilot_passwords||{};
  const sel = document.getElementById('ln-model');
  while (sel.options.length>1) sel.remove(1);
  _cfgModels.forEach(m => {
    const o=document.createElement('option'); o.value=m.nom||''; o.textContent=m.nom||''; sel.appendChild(o);
  });
  // Fallback: if no models, show POSTES
  if (!_cfgModels.length) {
    ["Matin","Après-midi","Nuit","Jour"].forEach(p => {
      const o=document.createElement('option'); o.value=p; o.textContent=p; sel.appendChild(o);
    });
  }
}

function popSel(id, vals) {
  const s = document.getElementById(id);
  if (!s) return;
  const cur = s.value;
  while (s.options.length>1) s.remove(1);
  vals.forEach(v => { const o=document.createElement('option'); o.value=v; o.textContent=v; s.appendChild(o); });
  if (cur) s.value = cur;
}

// ── Login model horaire info ──
function onLoginModelChange(){
  const nom=document.getElementById('ln-model').value;
  const info=document.getElementById('ln-model-info');
  const timesEl=document.getElementById('ln-model-times');
  const modifyDiv=document.getElementById('ln-model-modify');
  if(!nom){if(info)info.style.display='none';return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const todayKey=DAY_KEYS[new Date().getDay()];
  const model=_cfgModels.find(m=>m.nom===nom);
  let debut='',fin='';
  if(model&&model.jours&&model.jours[todayKey]){
    debut=model.jours[todayKey].debut||'';
    fin=model.jours[todayKey].fin||'';
  }
  if(timesEl) timesEl.textContent=debut&&fin?`${debut} → ${fin}`:'Aucun horaire défini pour aujourd\'hui';
  // Pre-fill modify inputs
  const nd=document.getElementById('ln-new-debut'),nf=document.getElementById('ln-new-fin');
  if(nd) nd.value=debut; if(nf) nf.value=fin;
  if(modifyDiv) modifyDiv.style.display='none';
  if(info) info.style.display='block';
}

function toggleLoginModelModify(){
  const d=document.getElementById('ln-model-modify');
  if(d) d.style.display=d.style.display==='none'?'block':'none';
}

// ── LOGIN ──
async function doLogin() {
  const pilot = document.getElementById('ln-pilot').value;
  const poste = document.getElementById('ln-model').value;
  const pw = document.getElementById('ln-pw').value;
  document.getElementById('ln-err').textContent = '';
  if (!pilot) { document.getElementById('ln-err').textContent='Choisir un pilote'; return; }
  if (!poste) { document.getElementById('ln-err').textContent='Choisir un modèle horaire'; return; }
  const r = await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pilot,poste,pw})});
  if (!r) return;
  const d = await r.json();
  if (d.ok) {
    // Si l'utilisateur a modifié les horaires du jour, les sauvegarder avant d'entrer
    const modifyDiv=document.getElementById('ln-model-modify');
    if(modifyDiv&&modifyDiv.style.display!=='none'){
      const newDebut=document.getElementById('ln-new-debut').value;
      const newFin=document.getElementById('ln-new-fin').value;
      if(newDebut&&newFin){
        const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
        const todayKey=DAY_KEYS[new Date().getDay()];
        const mi=_cfgModels.findIndex(m=>m.nom===poste);
        if(mi>=0){
          if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
          if(!_cfgModels[mi].jours[todayKey])_cfgModels[mi].jours[todayKey]={};
          _cfgModels[mi].jours[todayKey].debut=newDebut;
          _cfgModels[mi].jours[todayKey].fin=newFin;
          await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:todayKey,debut:newDebut,fin:newFin})});
        }
      }
    }
    const s = await apiFetch('/api/state');
    document.getElementById('v-login').classList.remove('on');
    showApp(s||{pilot,poste});
  } else {
    document.getElementById('ln-err').textContent = d.error||'Erreur connexion';
  }
}

function _getPobModelText(poste) {
  if(!poste||!_cfgModels) return poste||'—';
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const m=_cfgModels.find(x=>x.nom===poste);
  if(!m) return poste;
  const j=m.jours&&m.jours[dk];
  if(j&&j.debut&&j.fin) return `${poste} (${j.debut}→${j.fin})`;
  return poste;
}

function updatePobModel(poste){
  const el=document.getElementById('pob-model');
  if(el) el.textContent=_getPobModelText(poste||(_S_pilot_poste&&_S_pilot_poste.poste));
}

let _S_pilot_poste={poste:''};

function openPobModelEdit(){
  const poste=_S_pilot_poste&&_S_pilot_poste.poste;
  if(!poste){toast('Aucun modèle actif','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const m=_cfgModels.find(x=>x.nom===poste);
  const j=m&&m.jours&&m.jours[dk];
  document.getElementById('pobm-info').textContent=`Modèle : ${poste}`;
  document.getElementById('pobm-debut').value=(j&&j.debut)||'';
  document.getElementById('pobm-fin').value=(j&&j.fin)||'';
  openM('m-pobmodel');
}

async function savePobModelHours(){
  const poste=_S_pilot_poste&&_S_pilot_poste.poste;
  if(!poste) return;
  const debut=document.getElementById('pobm-debut').value;
  const fin=document.getElementById('pobm-fin').value;
  if(!debut||!fin){toast('Horaires incomplets','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mi=_cfgModels.findIndex(x=>x.nom===poste);
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=debut;
    _cfgModels[mi].jours[dk].fin=fin;
  }
  await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:dk,debut,fin})});
  updatePobModel(poste);
  closeM('m-pobmodel');
  toast('Horaires mis à jour','ok');
}

function showApp(s) {
  const app = document.getElementById('app');
  app.style.display = 'flex';
  app.classList.remove('hidden');
  if (s.pilot) { document.getElementById('f-pilote').value=s.pilot; document.getElementById('pob-pilot').textContent=s.pilot; }
  if (s.poste) {
    document.getElementById('f-poste').value=s.poste;
    document.getElementById('pob-poste').textContent=s.poste;
    _S_pilot_poste={poste:s.poste};
    updatePobModel(s.poste);
  }
  // Show prod tab button immediately if prod is active (don't wait for applyState)
  const tp=document.getElementById('ht-prod');
  if(tp) tp.classList.toggle('prod-visible',!!s.prod_active);
  // Also show the main banner immediately
  const mpb=document.getElementById('main-prod-banner');
  if(mpb) mpb.style.display=s.prod_active?'flex':'none';
  setToday();
  restoreFormFromStorage();
  pollState();
  pollEvts();
  loadCfg();
  goTab(s.prod_active ? 'prod' : 'main');
  const _today=new Date().toISOString().slice(0,10);
  const _hf=document.getElementById('hist-from'); if(_hf&&!_hf.value)_hf.value=_today;
  const _ht=document.getElementById('hist-to'); if(_ht&&!_ht.value)_ht.value=_today;
}

async function doLogout() {
  if (ST.prod_active) { toast('Terminer la production avant de déconnecter','err'); return; }
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  resetToLogin();
}

function resetToLogin() {
  ST={}; _curStopKey=null;
  document.getElementById('app').style.display='none';
  document.getElementById('ln-pw').value='';
  document.getElementById('ln-err').textContent='';
  document.getElementById('v-login').classList.add('on');
  _settingsUnlocked = false;
  // Reset horaires login aux valeurs config
  const _lnM=document.getElementById('ln-model');
  if(_lnM) _lnM.value='';
  const _lnInfo=document.getElementById('ln-model-info');
  if(_lnInfo) _lnInfo.style.display='none';
  const _lnMod=document.getElementById('ln-model-modify');
  if(_lnMod) _lnMod.style.display='none';
  loadModelsForLogin();
}

function setToday() {
  const d=document.getElementById('f-date');
  if(d) d.value=new Date().toLocaleDateString('fr-CA');
}

// ── NAVIGATION ──
function goTab(tab) {
  _curTab = tab;
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('on'));
  document.querySelectorAll('.htab').forEach(t=>t.classList.remove('on'));
  const vm={main:'v-main',prod:'v-prod',history:'v-history',settings:'v-settings',finposte:'v-finposte',kpi:'v-kpi'};
  const el=document.getElementById(vm[tab]);
  if(el) el.classList.add('on');
  const nt={main:'ht-main',prod:'ht-prod',history:'ht-hist',settings:'ht-cfg',kpi:'ht-kpi'};
  const ntEl=document.getElementById(nt[tab]);
  if(ntEl) ntEl.classList.add('on');
  if(tab==='history') loadHist();
  if(tab==='finposte') loadFPData();
  if(tab==='main') { loadMainDecl(); }
  if(tab==='kpi') loadKPI();
  if(tab==='settings') {
    // Show lock screen if not unlocked
    document.getElementById('settings-lock').style.display = _settingsUnlocked?'none':'flex';
    document.getElementById('v-settings-content').style.display = _settingsUnlocked?'block':'none';
    document.getElementById('lock-pw').value='';
    document.getElementById('lock-err').textContent='';
  }
}

function switchMTab(t) { loadMainDecl(); } // kept for compatibility

// ── SETTINGS LOCK ──
async function unlockSettings() {
  const pw = document.getElementById('lock-pw').value;
  // Try with entered password against /api/settings
  const r = await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,_check_only:true})});
  const d = r?await r.json():{};
  if(d.ok) {
    _settingsUnlocked = true;
    _adminPw = pw;
    document.getElementById('settings-lock').style.display='none';
    document.getElementById('v-settings-content').style.display='block';
    loadCfg();
  } else {
    document.getElementById('lock-err').textContent = d.error||'Mot de passe incorrect';
  }
}

// ── POLLING ──
async function pollState() {
  const s = await apiFetch('/api/state');
  if(!s) return;
  ST=s;
  _ofElapAtPoll = s.of_elapsed_s||0;
  _stopWallAtPoll = s.stop_wall_s||0;
  _pauseTotalAtPoll = s.pause_total_s||0;
  _lastPoll = Date.now();

  if(s.is_paused) {
    _curStopKey='_pause';
    _curStopElap = s.pause_total_s||0;
    if(s.pause_start_iso) _curStopElap += (Date.now()-new Date(s.pause_start_iso).getTime())/1000;
  } else if(s.active_stops&&s.active_stops.length>0) {
    const k=s.active_stops[0];
    _curStopKey=k;
    _curStopElap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
  } else {
    _curStopKey=null; _curStopElap=0;
  }

  applyState(s);
  if(_curTab==='main') loadMainDecl();
  if(_curTab==='kpi') loadKPI();
}

async function pollEvts() {
  const e = await apiFetch('/api/events_list');
  if(!Array.isArray(e)) return;
  // Filter to current pilot + today
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0');
  const todayPfx=dd+'/'+mm;
  const pilot=ST.pilot||'';
  gEvts=e.filter(ev=>{
    const dateOk=!ev.date||ev.date.startsWith(todayPfx);
    const pilotOk=!pilot||!ev.pilote||ev.pilote===pilot;
    return dateOk&&pilotOk;
  });
  // Don't override prod-view recap/timeline — applyState handles that from live tl_events
  if(!ST.prod_active){
    renderTL('tl-svg',gEvts);
    renderRecap(gEvts);
  }
}

function applyState(s) {
  // Header
  const hp=document.getElementById('hdr-pilot-lbl');
  if(hp&&s.pilot) hp.textContent=`${s.pilot} — ${s.poste||''}`;

  // Alert strip
  const al=document.getElementById('alert-strip');
  const stopOn=_curStopKey!==null;
  if(stopOn) {
    const lbl=_curStopKey==='_pause'?'PAUSE':getEvtLabel(_curStopKey);
    al.textContent=`⚠ ${lbl} EN COURS — cliquer pour terminer`;
    al.classList.add('on');
    al.onclick=()=>doEndStop();
    document.body.classList.add('stop-on');
  } else {
    al.classList.remove('on');
    al.onclick=null;
    document.body.classList.remove('stop-on');
  }

  // Prod tab visibility
  const tp=document.getElementById('ht-prod');
  if(tp) tp.classList.toggle('prod-visible',!!s.prod_active);

  // Main prod banner (shown when prod active but user is on main view)
  const mpb=document.getElementById('main-prod-banner');
  if(mpb) mpb.style.display=(s.prod_active&&_curTab==='main')?'flex':'none';
  if(s.prod_active&&s.form){
    const mpbOf=document.getElementById('mpb-of');
    if(mpbOf) mpbOf.textContent=s.form.of_num||'—';
  }

  // Excel busy bar
  const ebb=document.getElementById('excel-busy-bar');
  if(ebb) ebb.style.display=s.excel_busy?'block':'none';

  // Main btn-start
  const bs=document.getElementById('btn-start');
  if(bs) bs.disabled=s.prod_active;

  // POB (pilot/OF banner)
  if(s.prod_active&&s.form) {
    document.getElementById('pob-of').textContent=s.form.of_num||'—';
    // Don't override form while user is actively typing in it
    const af=document.activeElement;
    if(!af||!af.closest||!af.closest('.form-col')) fillFormFromState(s.form);
  }
  // OF start time in banner
  const pobOfStart=document.getElementById('pob-of-start');
  if(pobOfStart){
    if(s.of_start_iso){
      const os=new Date(s.of_start_iso);
      pobOfStart.textContent=String(os.getHours()).padStart(2,'0')+':'+String(os.getMinutes()).padStart(2,'0');
    } else {
      pobOfStart.textContent='—';
    }
  }

  // Render active stop chips
  renderStopChips(s);

  // Render live events for current prod (recap + timeline) from tl_events in state
  if(s.prod_active&&_curTab==='prod'){
    const le=tlEventsToDisplayFmt(s.tl_events||[]);
    renderTL('tl-svg',le);
    renderRecap(le);
  }

  // Pause button text
  const pbtn=document.getElementById('btn-pause');
  if(pbtn) pbtn.textContent=s.is_paused?'▶ Reprendre':'⏸ Pause';

  // TRS gauge
  updateGauge(s);
}

function getEvtLabel(key) {
  // Priorité : liste dynamique, puis EVENTS statique
  const dynEv=_evtsList.find(e=>e.key===key);
  if(dynEv) return (dynEv.cat==='ratt'?'Rattrapage: ':dynEv.cat==='pb'?'PB: ':'')+dynEv.label;
  const ev=EVENTS.find(e=>e[1]===key);
  if(ev) return (ev[2]==='ratt'?'Rattrapage: ':ev[2]==='pb'?'PB: ':'')+ev[0];
  if(key==='nettoyage') return 'Nettoyage';
  return key||'Arrêt';
}

function tlEventsToDisplayFmt(tlEvts){
  return (tlEvts||[]).map(ev=>{
    if(!ev.key||ev.key.startsWith('_')) return null;
    const s=ev.start?new Date(ev.start):null;
    const e=ev.end?new Date(ev.end):null;
    if(!s) return null;
    const pad=n=>String(n).padStart(2,'0');
    const toHMS=d=>pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
    const dur=e?(e.getTime()-s.getTime())/1000:0;
    let type;
    if(ev.cat==='nettoyage') type='Nettoyage'+(ev.nettoyage_type?' '+ev.nettoyage_type:'');
    else {
      const evDef=EVENTS.find(x=>x[1]===ev.key);
      if(evDef) type=(ev.cat==='ratt'?'Rattrapage: ':ev.cat==='pb'?'PB: ':'')+evDef[0];
      else type=ev.key;
    }
    return {type,cat:ev.cat||'autre',debut:toHMS(s),fin:e?toHMS(e):'',duree:dur>0?fmtDur(dur):'',comment:ev.comment||'',hors_trs:ev.hors_trs||false,_live:!ev.end};
  }).filter(Boolean);
}

function renderStopChips(s) {
  const cont=document.getElementById('stop-chips');
  const sb=document.getElementById('stop-bottom');
  if(!cont||!sb) return;
  const stops=s.active_stops||[];
  const hasAny=stops.length>0||s.is_paused;
  if(!hasAny){sb.classList.remove('on');cont.innerHTML='';return;}
  sb.classList.add('on');
  let html='';
  stops.forEach(k=>{
    const lbl=getEvtLabel(k);
    const elap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
    html+=`<div class="stop-chip"><span class="chip-lbl">⛔ ${esc(lbl)}</span><span class="chip-tim" id="chip-t-${esc(k)}">${fmtDur2(elap)}</span><button class="btn-endstop" onclick="doEndStop('${esc(k)}')">✓ Terminer</button></div>`;
  });
  if(s.is_paused){
    const pe=(s.pause_total_s||0)+(s.pause_start_iso?(Date.now()-new Date(s.pause_start_iso).getTime())/1000:0);
    html+=`<div class="stop-chip"><span class="chip-lbl">⏸ Pause</span><span class="chip-tim" id="chip-t-_pause">${fmtDur2(pe)}</span><button class="btn-endstop" onclick="doPause()">▶ Reprendre</button></div>`;
  }
  cont.innerHTML=html;
}

// ── TICKER ──
function startTicker() {
  if(_ticker) clearInterval(_ticker);
  _ticker=setInterval(()=>{
    if(!ST.prod_active) return;
    const dt=(Date.now()-_lastPoll)/1000;
    // OF timer
    const ofEl=_ofElapAtPoll+(!ST.is_paused?dt:0);
    const t=document.getElementById('sc-of');
    if(t) t.textContent=fmtDur(ofEl);
    // Stops total
    const sw=_stopWallAtPoll+((_curStopKey&&_curStopKey!=='_pause')?dt:0);
    const ts=document.getElementById('sc-stops');
    if(ts) ts.textContent=fmtDur(sw);
    // Pause total
    const pt=_pauseTotalAtPoll+(_curStopKey==='_pause'?dt:0);
    const tp=document.getElementById('sc-pause');
    if(tp) tp.textContent=fmtDur(pt);
    // Pièces théoriques : prod_ref / coef * (elapsed/28800)
    const thEl=document.getElementById('sc-theo');
    if(thEl&&ST.prod_ref){
      const typeProd=document.getElementById('f-type_prod')?.value||ST.form?.type_prod||'';
      const coef=(window._equivCoefs&&window._equivCoefs[typeProd])||1;
      const theo=Math.round(ST.prod_ref*(_ofElapAtPoll+dt)/28800/coef);
      thEl.textContent=theo>0?theo+' pièces':'—';
    }
    // Mise à jour bannière accueil (durée OF et arrêts)
    if(ST.prod_active&&_curTab==='main'){
      const mpbDur=document.getElementById('mpb-dur');
      const mpbStops=document.getElementById('mpb-stops');
      const ofElBanner=_ofElapAtPoll+(!ST.is_paused?dt:0);
      const swBanner=_stopWallAtPoll+((_curStopKey&&_curStopKey!=='_pause')?dt:0);
      if(mpbDur) mpbDur.textContent=fmtDur(ofElBanner);
      if(mpbStops) mpbStops.textContent=fmtDur(swBanner);
    }
    // Update stop chips timers
    if(ST.active_stops){
      ST.active_stops.forEach(k=>{
        const cel=document.getElementById('chip-t-'+k);
        if(cel&&ST.timers&&ST.timers[k]) cel.textContent=fmtDur2(ST.timers[k].elapsed+dt);
      });
    }
    if(ST.is_paused){
      const cel=document.getElementById('chip-t-_pause');
      if(cel) cel.textContent=fmtDur2(_pauseTotalAtPoll+dt);
    }
  },1000);
}

// ── MAIN VIEW ──
async function loadMainDecl() {
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0'),yyyy=now.getFullYear();
  const todayPfx=dd+'/'+mm;
  const [declData,evtData]=await Promise.all([apiFetch('/api/history'),apiFetch('/api/events_list')]);
  const _isProdRow=r=>['production','prod',''].includes((r.type||'').trim().toLowerCase());
  const decls=(Array.isArray(declData)?declData:(declData&&declData.rows?declData.rows:[])).filter(r=>
    (!r.date||r.date.startsWith(todayPfx))&&_isProdRow(r)
  );
  const evts=(Array.isArray(evtData)?evtData:[]).filter(r=>!r.date||r.date.startsWith(todayPfx));
  const allRows=[];
  decls.forEach(r=>allRows.push({...r,_rowType:'prod'}));
  evts.forEach(r=>allRows.push({...r,_rowType:'evt'}));
  allRows.sort((a,b)=>(b.debut||'').localeCompare(a.debut||''));
  const bd=document.getElementById('main-body');
  if(!bd) return;
  // Accumuler équivalences et arrêts du poste pour les jauges
  const curPilotD=ST.pilot||'';
  const pilotDecls=decls.filter(r=>!curPilotD||!r.pilote||r.pilote===curPilotD);
  const _hms2s=s=>{if(!s)return 0;const p=String(s).split(':');return p.length>=3?+p[0]*3600+ +p[1]*60+ +p[2]:p.length===2?+p[0]*60+ +p[1]:0;};
  // Filter to model horaire window only — exclude pre-shift prods (changement de série before model debut)
  const _dkLD=['dim','lun','mar','mer','jeu','ven','sam'][new Date().getDay()];
  const _mLD=_cfgModels&&_cfgModels.find(m=>m.nom===(ST.poste||''));
  const _jLD=_mLD&&_mLD.jours&&_mLD.jours[_dkLD];
  const _mDebS=_jLD&&_jLD.debut?_hms2s(_jLD.debut):0;
  // Set global _shiftRefDt so loadMainKPI and loadKPI use the same reference
  if(_jLD&&_jLD.debut){const[_hS,_mS]=_jLD.debut.split(':').map(Number);_shiftRefDt=new Date();_shiftRefDt.setHours(_hS,_mS,0,0);}
  else if(ST.shift_start_iso){_shiftRefDt=new Date(ST.shift_start_iso);}
  else{_shiftRefDt=null;}
  const inShiftDecls=_mDebS>0?pilotDecls.filter(r=>_hms2s(r.fin||'')>=_mDebS||_hms2s(r.debut||'')>=_mDebS):pilotDecls;
  _todayEquivAccum=inShiftDecls.reduce((a,r)=>a+parseFloat(r.equiv||0),0);
  const _pilotEvts=evts.filter(r=>!curPilotD||!r.pilote||r.pilote===curPilotD);
  const inShiftEvts=_mDebS>0?_pilotEvts.filter(r=>_hms2s(r.fin||'')>=_mDebS||_hms2s(r.debut||'')>=_mDebS):_pilotEvts;
  _todayStopAccum=inShiftEvts.reduce((a,r)=>a+_hms2s(r.duree||''),0);
  // Heure de la dernière déclaration prod enregistrée (dans la fenêtre du poste)
  if(inShiftDecls.length){
    const lastFin=inShiftDecls.map(r=>r.fin||'').filter(Boolean).sort().pop();
    if(lastFin){const[h,m,s]=(lastFin+'::').split(':').map(Number);const d=new Date();d.setHours(h,m,s||0,0);_lastProdDeclTime=d;}
  }
  if(!allRows.length){bd.innerHTML='<tr><td colspan="10" style="text-align:center;color:var(--gray);padding:16px">Aucune déclaration aujourd\'hui</td></tr>';loadMainKPI();updateGauge(ST);return;}
  window._rowMap={};
  bd.innerHTML=allRows.map(r=>{
    const key=r.row_num||r.debut;
    window._rowMap[String(key)]=r;
    const isProd=r._rowType==='prod';
    const t=parseFloat(r.trs||0);
    const tag=isProd?'<span class="row-tag tag-p">🏭 Prod</span>':(r.type&&r.type.toLowerCase().includes('nett')?'<span class="row-tag tag-n">🧹 Nett.</span>':'<span class="row-tag tag-e">⛔ Arrêt</span>');
    const details=isProd?esc(r.taille||''):esc(r.type||'');
    const qty=isProd?esc(String(r.qte_fab||'')):esc(r.duree||'');
    const info=isProd&&t>0?`<span class="${t>=90?'tg':t>=75?'tm':'tb'}">${fmtTRS(t)}</span>`:`<span style="color:var(--gray);font-size:10px">${esc(r.comment||'')}</span>`;
    return `<tr class="${isProd?'row-prod':'row-evt'}">
      <td>${tag}</td><td style="font-weight:600">${esc(r.of||'')}</td>
      <td style="font-size:10px">${esc(r.date||'')}</td><td style="font-size:10px">${esc(r.poste||'')}</td>
      <td>${esc(r.pilote||'')}</td><td>${esc(r.debut||'')}</td><td>${esc(r.fin||'')}</td>
      <td style="font-size:11px">${details}</td><td style="font-size:11px">${qty}</td><td>${info}</td>
      <td><button onclick="openEditRow('${esc(String(key))}')" style="background:#6366f1;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:15px;cursor:pointer;font-weight:700" title="Modifier">✏</button></td>
    </tr>`;
  }).join('');
  loadMainKPI();
  updateGauge(ST);
}

function saveMainModelHours(){
  const poste=ST.poste||'';if(!poste) return;
  const debut=document.getElementById('main-model-debut').value;
  const fin=document.getElementById('main-model-fin').value;
  if(!debut||!fin){toast('Horaires incomplets','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mi=_cfgModels.findIndex(m=>m.nom===poste);
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=debut;
    _cfgModels[mi].jours[dk].fin=fin;
  }
  fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:dk,debut,fin})});
  const el=document.getElementById('main-model-times');
  if(el) el.textContent=debut+' → '+fin;
  // Refresh Prod view model display and all TRS gauges/KPIs
  updatePobModel(poste);
  updateGauge(ST);
  loadMainKPI();
  loadKPI();
  toast('Horaires mis à jour','ok');
}

function updateMainModelDisplay(){
  const poste=ST.poste||'';
  const debutEl=document.getElementById('main-model-debut');
  const finEl=document.getElementById('main-model-fin');
  const timesEl=document.getElementById('main-model-times');
  const refCalcEl=document.getElementById('main-ref-calc');
  if(!poste) return;
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const m=_cfgModels.find(x=>x.nom===poste);
  const j=m&&m.jours&&m.jours[dk];
  const debut=(j&&j.debut)||'';const fin=(j&&j.fin)||'';
  if(timesEl) timesEl.textContent=debut&&fin?debut+' → '+fin:'Non défini';
  if(debutEl&&debut&&document.activeElement!==debutEl) debutEl.value=debut;
  if(finEl&&fin&&document.activeElement!==finEl) finEl.value=fin;
  // Show prod ref calculation
  if(refCalcEl&&ST.prod_ref>0){
    const ppm=(ST.prod_ref/480).toFixed(1);
    refCalcEl.textContent=`Réf: ${ST.prod_ref} pièces/8h = ${ppm} pcs/min`;
  }
}

async function loadMainKPI() {
  // Update model display
  updateMainModelDisplay();

  // Current shift
  const d=await apiFetch('/api/history_today');
  const evAll=await apiFetch('/api/events_list');
  const evArr=Array.isArray(evAll)?evAll:[];
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0');
  const todayPfx=dd+'/'+mm;
  const curPilot=ST.pilot||'';

  // Current shift stop time (today + current pilot)
  const curEvts=evArr.filter(e=>e.date&&e.date.startsWith(todayPfx)&&(!curPilot||!e.pilote||e.pilote===curPilot));
  const curStopS=curEvts.reduce((a,e)=>{try{const p=s=>s.split(':').reduce((acc,v,i)=>acc+(i===0?+v*3600:i===1?+v*60:+v),0);return a+Math.max(0,p(e.fin||'0:0:0')-p(e.debut||'0:0:0'));}catch(x){return a;}},0);
  // TRS à heure actuelle
  const heure=String(now.getHours()).padStart(2,'0')+'h'+String(now.getMinutes()).padStart(2,'0');
  const elHeure=document.getElementById('kpi0-heure');
  if(elHeure) elHeure.textContent=heure;
  if(d){
    // TRS poste actuel : même formule que la jauge accueil (equiv / (ref * elapsed/28800))
    let trs=-1;
    if(_todayEquivAccum>0&&_shiftRefDt&&ST.prod_ref>0){
      const refTime=_lastProdDeclTime||new Date();
      const shiftElap=(refTime.getTime()-_shiftRefDt.getTime())/1000;
      if(shiftElap>0) trs=Math.round(_todayEquivAccum/(ST.prod_ref*shiftElap/28800)*100*10)/10;
    }
    const el0t=document.getElementById('kpi0-trs'),el0s=document.getElementById('kpi0-sub'),el0d=document.getElementById('kpi0-date');
    if(el0t) el0t.textContent=fmtTRSv(trs);
    if(el0s) el0s.textContent=(d.rows?d.rows.length:0)+' OF | Arrêts '+Math.round(curStopS/60)+' min';
    if(el0d){
      const today=now.toLocaleDateString('fr-FR',{day:'2-digit',month:'2-digit',year:'numeric'});
      const shiftIso=ST.shift_start_iso;
      const shiftTime=shiftIso?new Date(shiftIso).toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'}):null;
      el0d.textContent=today+(shiftTime?' — depuis '+shiftTime:'');
    }
  }

  // Stats pause/nettoyage/réunion/prod du poste actuel
  const _hms2sFn=s=>{if(!s)return 0;const p=String(s).split(':');return p.length>=3?+p[0]*3600+ +p[1]*60+ +p[2]:p.length===2?+p[0]*60+ +p[1]:0;};
  let pauseS=0,nettS=0,reunionS=0,arretS=0,prodS=0;
  curEvts.forEach(e=>{
    const dur=_hms2sFn(e.duree||'');
    const t=(e.type||'').toLowerCase();
    if(t.includes('pause')) pauseS+=dur;
    else if(t.includes('nett')) nettS+=dur;
    else if(t.includes('réunion')||t.includes('reunion')) reunionS+=dur;
    arretS+=dur;
  });
  // Prod = sum of durations from decls (today+pilot)
  if(d&&d.rows) d.rows.forEach(r=>{ prodS+=_hms2sFn(r.duree||''); });
  const _fm=s=>Math.round(s/60)+' min';
  const spEl=document.getElementById('main-stat-pause'),snEl=document.getElementById('main-stat-nett');
  const srEl=document.getElementById('main-stat-reunion'),saEl=document.getElementById('main-stat-arrets'),sprodEl=document.getElementById('main-stat-prod');
  if(spEl) spEl.textContent=_fm(pauseS);
  if(snEl) snEl.textContent=_fm(nettS);
  if(srEl) srEl.textContent=_fm(reunionS);
  if(saEl) saEl.textContent=_fm(arretS);
  if(sprodEl) sprodEl.textContent=_fm(prodS);
  // Nb OF in new accueil block
  const nbOfEl=document.getElementById('acc-nb-of');
  if(nbOfEl) nbOfEl.textContent=d&&d.rows?d.rows.length:0;

  // Previous sessions from history (group by pilot+date, exclude current session)
  const allRows=await apiFetch('/api/history');
  const rows=Array.isArray(allRows)?allRows:[];
  // Build sessions: group by pilot+date, excluding today+curPilot
  const sessions={};
  rows.forEach(r=>{
    const key=(r.pilote||'?')+'|'+(r.date||'?');
    if(r.date&&r.date.startsWith(todayPfx)&&r.pilote===curPilot) return; // skip current session
    if(!sessions[key]) sessions[key]={pilot:r.pilote||'?',poste:r.poste||'?',date:r.date||'?',rows:[],eq:0,trs_sum:0,trs_cnt:0};
    sessions[key].rows.push(r);
    sessions[key].eq+=parseFloat(r.equiv||0);
    const t=parseFloat(r.trs||0);
    if(t>0){sessions[key].trs_sum+=t;sessions[key].trs_cnt++;}
  });
  const sessArr=Object.values(sessions).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,2);
  for(let i=0;i<2;i++){
    const lbl=document.getElementById('kpi'+(i+1)+'-lbl'),tv=document.getElementById('kpi'+(i+1)+'-trs'),sv=document.getElementById('kpi'+(i+1)+'-sub');
    if(!lbl||!tv||!sv) continue;
    if(sessArr[i]){
      const s=sessArr[i];
      const avgT=s.trs_cnt>0?s.trs_sum/s.trs_cnt:-1;
      // Count stops for this session
      const sEvts=evArr.filter(e=>e.date===s.date&&e.pilote===s.pilot);
      const sStopS=sEvts.reduce((a,e)=>{try{const p=t=>t.split(':').reduce((acc,v,ii)=>acc+(ii===0?+v*3600:ii===1?+v*60:+v),0);return a+Math.max(0,p(e.fin||'0:0:0')-p(e.debut||'0:0:0'));}catch(x){return a;}},0);
      const dt=document.getElementById('kpi'+(i+1)+'-date');
      lbl.textContent=s.pilot+' — '+s.poste;
      tv.textContent=fmtTRSv(avgT);
      sv.textContent=s.rows.length+' OF | Arrêts '+Math.round(sStopS/60)+' min';
      if(dt){
        // Get first debut and last fin from this session's rows
        const debs=s.rows.map(r=>r.debut_of||r.debut||'').filter(Boolean).sort();
        const fins=s.rows.map(r=>r.fin_of||r.fin||'').filter(Boolean).sort().reverse();
        const timeRange=(debs.length&&fins.length)?` ${debs[0].slice(0,5)}→${fins[0].slice(0,5)}`:'';
        dt.textContent=s.date+timeRange;
      }
    } else {
      lbl.textContent=i===0?'Poste précédent':'Avant-dernier';
      tv.textContent='--%'; sv.textContent='—';
      const dt=document.getElementById('kpi'+(i+1)+'-date');
      if(dt) dt.textContent='';
    }
  }
}

// ── START PROD ──
let _pendingGapS=0;
let _interposteLbls=["Changement de série","Réglage / Setup machine","Attente matière première","Réunion / Formation","Nettoyage interposte","Pause"];
let _interposteLblsEditing=[];

async function loadInterposteCfg(){
  const d=await apiFetch('/api/interposte_cfg');
  if(d&&d.labels&&d.labels.length) _interposteLbls=d.labels;
  _interposteLblsEditing=[..._interposteLbls];
  _renderInterposteLblsHTML();
}
function _renderInterposteLblsHTML(){
  const c=document.getElementById('interposte-list-ui');if(!c) return;
  c.innerHTML=_interposteLblsEditing.map((lbl,i)=>`
    <div style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:12px">
      <span style="flex:1;font-weight:600">${esc(lbl)}</span>
      <button class="btn btn-ghost" style="font-size:10px;padding:2px 6px" onclick="editInterposteLbl(${i})">✏</button>
      <button class="btn btn-danger" style="font-size:10px;padding:2px 6px" onclick="rmInterposteLbl(${i})">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:11px;padding:4px">Aucun label</div>';
}
function addInterposteLbl(){const lbl=document.getElementById('ip-new-label').value.trim();if(!lbl){toast('Nom requis','err');return;}_interposteLblsEditing.push(lbl);document.getElementById('ip-new-label').value='';_renderInterposteLblsHTML();}
function rmInterposteLbl(i){_interposteLblsEditing.splice(i,1);_renderInterposteLblsHTML();}
function editInterposteLbl(i){const lbl=prompt('Label :',_interposteLblsEditing[i]);if(lbl&&lbl.trim()){_interposteLblsEditing[i]=lbl.trim();_renderInterposteLblsHTML();}}
async function saveInterposteCfg(){
  const r=await fetch('/api/interposte_cfg',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,labels:_interposteLblsEditing})});
  const d=r?await r.json():{};
  if(d&&d.ok){_interposteLbls=[..._interposteLblsEditing];toast('Labels interposte enregistrés','ok');}
  else toast(d?.error||'Erreur','err');
}

function _fmtMin(s){const m=Math.round(s/60),h=Math.floor(m/60),mi=m%60;return h?`${h}h ${mi}min`:`${mi} min`;}

async function doStartProd() {
  saveFormToStorage();
  const r=await fetch('/api/start_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  if(!r) return;
  const d=await r.json();
  if(!d.ok){toast(d.error||'Erreur','err');return;}
  setToday();
  restoreFormFromStorage();
  await pollState();
  await pollEvts();
  _pendingGapS=d.gap_s||0;

  // 1er OF du poste : gap vs modèle horaire
  if((d.pre_shift_gap_s||0)>120 && d.shift_model_start){
    const m=_fmtMin(d.pre_shift_gap_s);
    document.getElementById('ps-text').textContent=
      `${m} non déclarées depuis le début de poste (${d.shift_model_start})`;
    document.getElementById('ps-start-iso').value=d.shift_model_start_iso||'';
    document.getElementById('ps-gap-s').value=d.pre_shift_gap_s||0;
    const bt=document.getElementById('ps-backdate-time');
    if(bt) bt.textContent=d.shift_model_start||'--h--';
    openM('m-preshift');
    return;
  }

  // OF suivant : gap interposte classique
  if(_pendingGapS>30){
    const bc=document.getElementById('ip-btns');bc.innerHTML='';
    document.getElementById('ip-duration').textContent=`Durée : ${_fmtMin(_pendingGapS)}`;
    document.getElementById('ip-custom').value='';
    document.getElementById('ip-comment').value='';
    _interposteLbls.forEach(lbl=>{
      const b=document.createElement('button');
      b.className='btn btn-ghost';b.style.fontSize='12px';b.textContent=lbl;
      b.onclick=()=>{document.getElementById('ip-custom').value=lbl;};
      bc.appendChild(b);
    });
    openM('m-interposte');
    return;
  }

  goTab('prod');
}

// ── Choix pré-poste ──
async function psChooseInterposte(){
  const gapS=parseFloat(document.getElementById('ps-gap-s').value)||0;
  const startIso=document.getElementById('ps-start-iso').value;
  closeM('m-preshift');
  // Ouvrir le modal interposte pour laisser choisir le label
  _pendingGapS=gapS;
  document.getElementById('ip-duration').textContent=`Durée : ${_fmtMin(gapS)} (début de poste → 1er OF)`;
  document.getElementById('ip-custom').value='Début de poste';
  document.getElementById('ip-comment').value='';
  const bc=document.getElementById('ip-btns');bc.innerHTML='';
  ['Mise en route machine','Réunion début de poste','Attente / Préparation','Nettoyage arrivée'].forEach(lbl=>{
    const b=document.createElement('button');
    b.className='btn btn-ghost';b.style.fontSize='12px';b.textContent=lbl;
    b.onclick=()=>{document.getElementById('ip-custom').value=lbl;};
    bc.appendChild(b);
  });
  // Override confirm pour aussi corriger shift_start
  window._psStartIso=startIso;
  openM('m-interposte');
}

async function psChooseBackdate(){
  const startIso=document.getElementById('ps-start-iso').value;
  closeM('m-preshift');
  if(startIso){
    await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:startIso})});
    await pollState();
  }
  toast('OF rétro-daté au début de poste','ok');
  goTab('prod');
}

function psChooseIgnore(){
  closeM('m-preshift');
  goTab('prod');
}

function psShowModifyModel(){
  // Récupérer le modèle horaire actuel pour pré-remplir les champs
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const todayKey=DAY_KEYS[new Date().getDay()];
  const model=_cfgModels.find(m=>m.nom===ST.poste);
  let debut='',fin='';
  if(model&&model.jours&&model.jours[todayKey]){
    debut=model.jours[todayKey].debut||'';
    fin=model.jours[todayKey].fin||'';
  }
  document.getElementById('ps-new-debut').value=debut;
  document.getElementById('ps-new-fin').value=fin;
  document.getElementById('ps-model-form').style.display='block';
}

async function psConfirmModifyModel(){
  const newDebut=document.getElementById('ps-new-debut').value;
  const newFin=document.getElementById('ps-new-fin').value;
  if(!newDebut||!newFin){toast('Remplissez le début et la fin','warn');return;}
  // Mettre à jour le modèle horaire dans la config pour aujourd'hui
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const todayKey=DAY_KEYS[new Date().getDay()];
  const modelIdx=_cfgModels.findIndex(m=>m.nom===ST.poste);
  if(modelIdx>=0){
    if(!_cfgModels[modelIdx].jours)_cfgModels[modelIdx].jours={};
    if(!_cfgModels[modelIdx].jours[todayKey])_cfgModels[modelIdx].jours[todayKey]={};
    _cfgModels[modelIdx].jours[todayKey].debut=newDebut;
    _cfgModels[modelIdx].jours[todayKey].fin=newFin;
    await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,modeles_horaires:_cfgModels})});
  }
  // Rétrodater l'OF au nouvel horaire de début
  const today=new Date();
  const [hh,mm]=newDebut.split(':');
  const newStart=new Date(today.getFullYear(),today.getMonth(),today.getDate(),parseInt(hh),parseInt(mm),0);
  const newStartIso=newStart.toISOString();
  closeM('m-preshift');
  await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:newStartIso})});
  await pollState();
  toast('Horaires mis à jour, OF rétro-daté','ok');
  goTab('prod');
}

async function confirmInterposte(){
  const lbl=document.getElementById('ip-custom').value.trim()||'Interposte';
  const cmt=document.getElementById('ip-comment').value.trim();
  closeM('m-interposte');
  await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:_pendingGapS,label:lbl,comment:cmt})});
  // Si on vient d'un popup pré-poste, rétrodater le shift_start aussi
  if(window._psStartIso){
    await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:window._psStartIso})});
    window._psStartIso=null;
  }
  await pollState();
  await pollEvts();
  loadMainDecl(); // interposte row must appear in accueil without waiting
  goTab('prod');
}

async function skipInterposte(){
  window._psStartIso=null;
  closeM('m-interposte');
  if(_pendingGapS>30){
    await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:_pendingGapS,label:'Interposte',comment:''})});
  }
  goTab('prod');
}

async function doCancelProd(){
  if(!confirm('Annuler cette production ? Aucune donnée ne sera écrite dans Excel.')) return;
  await fetch('/api/force_reset_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
  goTab('main');
  toast('Production annulée','ok');
}

// ── STOP/PAUSE ──
let _evtsList=[]; // dynamic events list from server

async function loadEvtsList(){
  const d=await apiFetch('/api/events_cfg');
  if(!d||!d.events) return;
  _evtsList=d.events;
  // Sync with JS EVENTS array for getEvtLabel compatibility
  // EVENTS stays as fallback, but we prefer _evtsList
  rebuildStopGrids();
  renderEvtListUI();
}

function rebuildStopGrids(){
  const evts=_evtsList.length?_evtsList:EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]}));
  const GRIDS={ratt:'sgrid-ratt',pb:'sgrid-pb',nettoyage:'sgrid-nettoyage',organisation:'sgrid-organisation'};
  const LABELS={nettoyage:'sgrid-nett-lbl',organisation:'sgrid-org-lbl'};
  Object.values(GRIDS).forEach(id=>{const g=document.getElementById(id);if(g)g.innerHTML='';});
  evts.forEach(e=>{
    const gid=GRIDS[e.cat]; if(!gid) return;
    const g=document.getElementById(gid); if(!g) return;
    const b=document.createElement('button');
    const col=STOP_COL[e.cat]||'#64748b';
    b.style.cssText=`background:${col};color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:12px;font-weight:600;cursor:pointer`;
    b.textContent=e.label;
    b.onclick=()=>{closeM('m-stop');doStartStop(e.key,e.cat);};
    g.appendChild(b);
    g.style.display='';
    if(LABELS[e.cat]){const l=document.getElementById(LABELS[e.cat]);if(l)l.style.display='';}
  });
}

function buildStopGrids(){rebuildStopGrids();}

function getEvtLabelDynamic(key){
  const ev=_evtsList.find(e=>e.key===key)||EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]})).find(e=>e.key===key);
  if(!ev) return key||'Arrêt';
  return (ev.cat==='ratt'?'Rattrapage: ':ev.cat==='pb'?'PB: ':'')+ev.label;
}

// ── EVENTS LIST UI (Settings) ──
let _evtsEditing=[];
function renderEvtListUI(){
  _evtsEditing=(_evtsList.length?_evtsList:EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]}))).map(e=>({...e}));
  _renderEvtListHTML();
}

function _renderEvtListHTML(){
  const c=document.getElementById('events-list-ui');if(!c) return;
  const catLbl={pb:'🔴 Panne',ratt:'🟠 Rattrapage',nettoyage:'🟡 Nettoyage',organisation:'🔵 Organisation',autre:'⚫ Autre'};
  c.innerHTML=_evtsEditing.map((e,i)=>`
    <div style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:12px">
      <span style="flex:1;font-weight:600">${esc(e.label)}</span>
      <span style="font-size:10px;color:var(--gray)">${catLbl[e.cat]||e.cat}</span>
      <button class="btn btn-ghost" style="font-size:10px;padding:2px 6px" onclick="editEvtItem(${i})">✏</button>
      <button class="btn btn-danger" style="font-size:10px;padding:2px 6px" onclick="rmEvtItem(${i})">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:11px;padding:4px">Aucun arrêt configuré</div>';
}

function addEvtItem(){
  const lbl=document.getElementById('ev-new-label').value.trim();
  const cat=document.getElementById('ev-new-cat').value;
  if(!lbl){toast('Nom requis','err');return;}
  const key=lbl.toLowerCase().replace(/[^a-z0-9]/g,'_').slice(0,32)+'_'+Date.now().toString(36);
  _evtsEditing.push({label:lbl,key,cat});
  document.getElementById('ev-new-label').value='';
  _renderEvtListHTML();
}

function rmEvtItem(i){_evtsEditing.splice(i,1);_renderEvtListHTML();}

function editEvtItem(i){
  const e=_evtsEditing[i];
  const lbl=prompt('Nom de l\'arrêt :',e.label);
  if(lbl&&lbl.trim()){_evtsEditing[i].label=lbl.trim();_renderEvtListHTML();}
}

async function saveEvtList(){
  const r=await fetch('/api/events_cfg',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,events:_evtsEditing})});
  const d=r?await r.json():{};
  if(d&&d.ok){
    _evtsList=_evtsEditing.map(e=>({...e}));
    rebuildStopGrids();
    toast('Liste des arrêts enregistrée','ok');
  } else toast(d?.error||'Erreur','err');
}

function openStopModal(){openM('m-stop');}

async function doPause(){
  try{await fetch('/api/toggle_pause',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});}
  catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
}

async function doNettoyage(){
  try{await fetch('/api/start_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ntype:'court'})});}
  catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
}

async function doStartStop(key,cat){
  try{
    const r=await fetch('/api/start_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,cat})});
    if(!r||!r.ok) toast('Erreur déclaration arrêt','err');
  }catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
  await pollEvts();
}

function doEndStop(key) {
  const k=key||_curStopKey;
  if(!k||k==='_pause'){doPause();return;}
  if(k==='nettoyage'){
    fetch('/api/end_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'}).then(()=>{pollState();pollEvts();});
    return;
  }
  // Show comment modal
  document.getElementById('cmt-stop-key').value=k;
  document.getElementById('cmt-stop-lbl').textContent='Arrêt : '+getEvtLabel(k);
  document.getElementById('cmt-stop-text').value='';
  openM('m-stopcmt');
  setTimeout(()=>document.getElementById('cmt-stop-text').focus(),100);
}

async function confirmEndStop() {
  const k=document.getElementById('cmt-stop-key').value;
  const cmt=document.getElementById('cmt-stop-text').value.trim();
  closeM('m-stopcmt');
  await fetch('/api/end_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:k,comment:cmt})});
  await pollState();
  await pollEvts();
}

async function declareCustomStop(){
  const v=document.getElementById('custom-stop-input').value.trim();
  if(!v) return;
  document.getElementById('custom-stop-input').value='';
  closeM('m-stop');
  await doStartStop(v,'autre');
}

// ── FORM ──
function collectForm(){
  const f={};
  FORM_FIELDS.forEach(k=>{
    const el=document.getElementById('f-'+k);
    if(!el) return;
    f[k]=el.type==='number'?(parseFloat(el.value)||0):el.value;
  });
  f.pilote=document.getElementById('f-pilote')?.value||ST.pilot||'';
  f.poste=document.getElementById('f-poste')?.value||ST.poste||'';
  return f;
}

function fillFormFromState(form){
  if(!form) return;
  FORM_FIELDS.forEach(k=>{
    const el=document.getElementById('f-'+k);
    if(!el) return;
    const v=form[k];
    if(v!==undefined&&v!==null&&v!=='') el.value=v;
  });
  const ofEl=document.getElementById('pob-of');
  if(ofEl) ofEl.textContent=form.of_num||'—';
}

// Form persistence in localStorage (persist across restarts until new prod)
function saveFormToStorage(){
  const f=collectForm();
  try{localStorage.setItem('kpiorc_form',JSON.stringify(f));}catch(e){}
}
function restoreFormFromStorage(){
  try{
    const raw=localStorage.getItem('kpiorc_form');
    if(!raw) return;
    const f=JSON.parse(raw);
    // Only restore if prod_active and form has of_num, or if not active (post-prod)
    FORM_FIELDS.forEach(k=>{
      const el=document.getElementById('f-'+k);
      if(!el||!(k in f)) return;
      const v=f[k];
      if(v!==undefined&&v!==null&&v!=='') el.value=v;
    });
  }catch(e){}
}

function scheduleAutoSave(){
  clearTimeout(_autoSaveTimer);
  _autoSaveTimer=setTimeout(()=>{
    const f=collectForm();
    fetch('/api/save_form',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(f)});
    saveFormToStorage();
  },1500);
}

// ── END PROD ──
async function doEndProdPreview(){
  const f=collectForm();
  const r=await fetch('/api/preview_end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:f})});
  if(!r||!r.ok){confirmEndProd();return;}
  const d=await r.json();
  renderEPModal(d,f);
  openM('m-endprod');
}

function renderEPModal(d,f){
  const ofNum=f.of_num||d.of_num||'';
  document.getElementById('ep-title').textContent=`⏹ Fin d'OF/prod${ofNum?' — '+ofNum:''}`;
  // Graphs
  drawPie('ep-pie',[
    {label:'Prod',value:d.prod_s||0,color:'#16a34a'},
    {label:'Arrêts',value:d.stop_s||0,color:'#dc2626'},
  ]);
  drawGauge('ep-gauge-arc','ep-gauge-pct',d.trs>=0?d.trs:0);
  const now=new Date();
  const dateStr=String(now.getDate()).padStart(2,'0')+'/'+String(now.getMonth()+1).padStart(2,'0')+'/'+now.getFullYear();
  document.getElementById('ep-stats').innerHTML=`
    <div class="ep-stat"><div class="val">${esc(ofNum||'—')}</div><div class="lbl">N° OF</div></div>
    <div class="ep-stat"><div class="val">${esc(f.type_prod||d.type_prod||'—')}</div><div class="lbl">Type produit</div></div>
    <div class="ep-stat"><div class="val">${esc(String(f.qte_fab||d.qte_fab||0))}</div><div class="lbl">Qté fabriquée</div></div>
    <div class="ep-stat"><div class="val">${esc(String(f.qte_emb||d.qte_emb||0))}</div><div class="lbl">Qté emballée</div></div>
    <div class="ep-stat"><div class="val">${dateStr}</div><div class="lbl">Date</div></div>
    <div class="ep-stat"><div class="val">${fmtTRS(d.trs)}</div><div class="lbl">TRS OF</div></div>
    <div class="ep-stat"><div class="val">${(d.equiv||0).toFixed(1)}</div><div class="lbl">Équivalence</div></div>
    <div class="ep-stat"><div class="val">${fmtD2(d.prod_s||0)}</div><div class="lbl">Durée prod</div></div>
    <div class="ep-stat"><div class="val">${fmtD2(d.stop_s||0)}</div><div class="lbl">Total arrêts</div></div>
  `;
  const evts=d.tl_events||[];
  const stopMap={};
  const totalS=d.stop_s||1;
  evts.forEach(e=>{if(!e.key||e.key==='prod')return;stopMap[e.key]=(stopMap[e.key]||0)+(e.dur_s||0);});
  const tbody=document.getElementById('ep-stops');
  tbody.innerHTML=Object.entries(stopMap).map(([k,s])=>`
    <tr><td>${getEvtLabel(k)}</td><td>${fmtD2(s)}</td><td>${Math.round(s/totalS*100)}%</td></tr>
  `).join('')||'<tr><td colspan="3" style="color:var(--gray)">Aucun arrêt</td></tr>';
  drawTL('ep-tl',evts,d.debut,d.now_str);
}

async function confirmEndProd(){
  const f=collectForm();
  closeM('m-endprod');
  const r=await fetch('/api/end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:f})});
  if(!r) return;
  const d=await r.json();
  if(d.ok){
    // Clear all form fields for next prod
    FORM_FIELDS.forEach(k=>{
      const el=document.getElementById('f-'+k);
      if(!el) return;
      el.value='';
    });
    document.getElementById('f-nb_pers').value='10';
    try{localStorage.removeItem('kpiorc_form');}catch(e){}
    await pollState();
    await pollEvts();
    goTab('main');
    toast('Production enregistrée','ok');
    // Auto-generate dashboard
    try{
      await fetch('/api/generate_dashboard',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
    }catch(e){}
  } else toast(d.error||'Erreur','err');
}

async function forceResetProd(){
  if(!confirm('Annuler la prod en cours ? Les données non enregistrées seront perdues.')) return;
  await fetch('/api/force_reset_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
  toast('Production annulée','warn');
}

// ── PIE & GAUGE CHARTS ──
function drawPie(svgId, segments) {
  const svg=document.getElementById(svgId);if(!svg) return;
  const total=segments.reduce((a,s)=>a+s.value,0);
  if(total<=0){svg.innerHTML='<text x="65" y="60" text-anchor="middle" font-size="9" fill="#94a3b8">Pas de données</text>';return;}
  const cx=65,cy=57,r=44,ir=24;let html='',startAngle=-Math.PI/2;
  segments.forEach(seg=>{
    if(seg.value<=0) return;
    const angle=(seg.value/total)*2*Math.PI;if(angle<0.001) return;
    const endAngle=startAngle+angle,large=angle>Math.PI?1:0;
    const x1=(cx+r*Math.cos(startAngle)).toFixed(2),y1=(cy+r*Math.sin(startAngle)).toFixed(2);
    const x2=(cx+r*Math.cos(endAngle)).toFixed(2),y2=(cy+r*Math.sin(endAngle)).toFixed(2);
    const ix1=(cx+ir*Math.cos(startAngle)).toFixed(2),iy1=(cy+ir*Math.sin(startAngle)).toFixed(2);
    const ix2=(cx+ir*Math.cos(endAngle)).toFixed(2),iy2=(cy+ir*Math.sin(endAngle)).toFixed(2);
    html+=`<path d="M${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} L${ix2},${iy2} A${ir},${ir} 0 ${large},0 ${ix1},${iy1} Z" fill="${seg.color}"/>`;
    startAngle=endAngle;
  });
  const m=segments[0],mp=total>0?Math.round(m.value/total*100):0;
  html+=`<text x="${cx}" y="${cy+5}" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e">${mp}%</text>`;
  html+=`<text x="${cx}" y="${cy+15}" text-anchor="middle" font-size="7" fill="#64748b">${esc(m.label)}</text>`;
  let lx=0;segments.filter(s=>s.value>0).forEach(s=>{
    const p=Math.round(s.value/total*100);
    html+=`<rect x="${lx}" y="108" width="7" height="7" fill="${s.color}" rx="1"/>`;
    html+=`<text x="${lx+9}" y="115" font-size="7" fill="#475569">${esc(s.label)} ${p}%</text>`;lx+=65;
  });
  svg.innerHTML=html;
}

function drawGauge(arcId,pctId,trs) {
  const arc=document.getElementById(arcId),pct=document.getElementById(pctId);if(!arc||!pct) return;
  const pArc=132,v=Math.max(0,Math.min(100,trs||0)),dash=(v/100)*pArc;
  const col=v>=90?'#16a34a':v>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash.toFixed(1)},${pArc}`);arc.setAttribute('stroke',col);
  pct.textContent=trs>=0?fmtTRS(trs):'--%';pct.setAttribute('fill',col);
}

// ── GAUGE (prod view) ──
function updateGauge(s){
  const arc=document.getElementById('gauge-arc');
  const pct=document.getElementById('gauge-pct');
  const pobTrs=document.getElementById('pob-trs');
  if(!arc||!pct) return;
  // Estimate live TRS using coefficient
  const prodRef=s.prod_ref||200;
  const ofS=s.of_elapsed_s||0;
  const qFab=s.form?parseFloat(s.form.qte_fab||0):0;
  const typeProd=s.form?s.form.type_prod||'':'';
  const coef=(window._equivCoefs&&typeProd&&window._equivCoefs[typeProd])||1;
  const equiv=qFab*coef;
  let trs=-1;
  if(ofS>0&&prodRef>0&&equiv>0){
    trs=Math.round(equiv/(prodRef*ofS/28800)*100*10)/10;
  }
  const trsStr=trs>=0?fmtTRS(trs):'—';
  pct.textContent=trsStr;
  if(pobTrs) pobTrs.textContent=trsStr;
  const pArc=132;
  const dash=trs>=0?Math.min(1,trs/100)*pArc:0;
  const col=trs>=90?'#16a34a':trs>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash},${pArc}`);
  arc.setAttribute('stroke',col);

  // TRS du poste (shift_start → dernière déclaration enregistrée)
  const arcPoste=document.getElementById('gauge-poste-arc');
  const pctPoste=document.getElementById('gauge-poste-pct');
  const lblPoste=document.getElementById('gauge-poste-lbl');
  // Même calcul pour l'accueil
  const arcPosteAcc=document.getElementById('gauge-poste-acc-arc');
  const pctPosteAcc=document.getElementById('gauge-poste-acc-pct');
  const lblPosteAcc=document.getElementById('gauge-poste-acc-lbl');
  // TRS Poste: use model horaire debut as reference (not shift_start_iso which may include pre-shift time)
  const _dayKeysG=['dim','lun','mar','mer','jeu','ven','sam'];
  const _dkG=_dayKeysG[new Date().getDay()];
  const _modelG=_cfgModels&&_cfgModels.find(m=>m.nom===(s.poste||ST.poste||''));
  const _jourG=_modelG&&_modelG.jours&&_modelG.jours[_dkG];
  // Update global _shiftRefDt (shared with loadMainKPI / loadKPI)
  if(_jourG&&_jourG.debut){const[_hG,_mG]=_jourG.debut.split(':').map(Number);_shiftRefDt=new Date();_shiftRefDt.setHours(_hG,_mG,0,0);}
  else if(s.shift_start_iso){_shiftRefDt=new Date(s.shift_start_iso);}
  else{_shiftRefDt=null;}
  if(_shiftRefDt&&s.prod_ref>0){
    const calcRef=_lastProdDeclTime||new Date();  // pour le calc TRS : fallback now si aucune décl
    const shiftElap=(calcRef.getTime()-_shiftRefDt.getTime())/1000;
    const todayEquiv=_todayEquivAccum||0;
    const trsPoste=shiftElap>0&&todayEquiv>0?Math.round(todayEquiv/(s.prod_ref*shiftElap/28800)*100*10)/10:-1;
    // Label : "Entre Xh et Yh" (Y = heure fin de la dernière déclaration, pas l'heure actuelle)
    let lbl;
    if(_lastProdDeclTime&&_shiftRefDt){
      const dh=String(_shiftRefDt.getHours()).padStart(2,'0'),dm=String(_shiftRefDt.getMinutes()).padStart(2,'0');
      const fh=String(_lastProdDeclTime.getHours()).padStart(2,'0'),fm=String(_lastProdDeclTime.getMinutes()).padStart(2,'0');
      lbl='Entre '+dh+'h'+dm+' et '+fh+'h'+fm;
    } else {
      lbl='TRS Poste';
    }
    const dashP=trsPoste>=0?Math.min(1,trsPoste/100)*pArc:0;
    const colP=trsPoste>=90?'#16a34a':trsPoste>=75?'#d97706':'#dc2626';
    [arcPoste,arcPosteAcc].forEach(el=>{if(el){el.setAttribute('stroke-dasharray',`${dashP},${pArc}`);el.setAttribute('stroke',colP);}});
    [pctPoste,pctPosteAcc].forEach(el=>{if(el){el.textContent=trsPoste>=0?fmtTRS(trsPoste):'—';el.setAttribute('fill',colP);}});
    [lblPoste,lblPosteAcc].forEach(el=>{if(el) el.textContent=lbl;});
  }

  // Pie charts
  const stopS=s.stop_wall_s||0;
  const ofDur=s.of_elapsed_s||0;
  const prodSof=Math.max(0,ofDur-stopS);
  drawPie('pie-of',[{label:'Prod',value:prodSof,color:'#16a34a'},{label:'Arrêts',value:stopS,color:'#dc2626'}]);
  // For poste pie — compute from shift start
  const shiftTotal=s.shift_start_iso?(Date.now()-new Date(s.shift_start_iso).getTime())/1000:0;
  const shiftStop=_todayStopAccum||0;
  const shiftProd=Math.max(0,shiftTotal-shiftStop);
  drawPie('pie-poste',[{label:'Prod',value:shiftProd,color:'#16a34a'},{label:'Arrêts',value:shiftStop,color:'#dc2626'}]);
}
// Accumulateurs poste (mis à jour à chaque loadMainDecl)
let _todayEquivAccum=0, _todayStopAccum=0, _lastProdDeclTime=null, _shiftRefDt=null;

// ── EDIT ROW (accueil) ──
function openEditRow(key) {
  const row=window._rowMap[String(key)];if(!row) return;
  const isProd=row._rowType==='prod';
  document.getElementById('er-rownum').value=row.row_num||'';
  document.getElementById('er-rowtype').value=row._rowType||'';
  document.getElementById('er-title').textContent=isProd?'✏ Modifier déclaration':'✏ Modifier événement';
  const pwRow=document.getElementById('er-pw-row'),pwEl=document.getElementById('er-pw');
  pwEl.value='';pwRow.style.display='none';
  document.getElementById('er-prod-fields').style.display=isProd?'':'none';
  document.getElementById('er-evt-fields').style.display=isProd?'none':'';
  if(isProd){
    document.getElementById('er-of').value=row.of||'';
    document.getElementById('er-deb').value=(row.debut||'').slice(0,5);
    document.getElementById('er-fin').value=(row.fin||'').slice(0,5);
    document.getElementById('er-poste').value=row.poste||'';
    document.getElementById('er-pilote').value=row.pilote||'';
    document.getElementById('er-copilote').value=row.copilote||'';
    document.getElementById('er-nbpers').value=row.nb_pers||'';
    document.getElementById('er-taille').value=row.taille||'';
    document.getElementById('er-codeprod').value=row.code_prod||'';
    document.getElementById('er-typeprod').value=row.type_prod||'';
    document.getElementById('er-qtefab').value=row.qte_fab||'';
    document.getElementById('er-qteemb').value=row.qte_emb||'';
    document.getElementById('er-poids').value=row.poids||'';
    document.getElementById('er-fibre').value=row.fibre||'';
    document.getElementById('er-oftaie').value=row.of_taie||'';
    document.getElementById('er-traca').value=row.traca||'';
    document.getElementById('er-reftaie').value=row.ref_taie||'';
    document.getElementById('er-kit').value=row.kit||'';
    document.getElementById('er-qteinit').value=row.qte_init_taie||'';
    document.getElementById('er-nbtaie2').value=row.nb_taie2_choix||'';
    document.getElementById('er-nbdef').value=row.nb_def_cout||'';
    document.getElementById('er-mqtaie').value=row.mq_taie||'';
    document.getElementById('er-mqhousse').value=row.mq_housse_encart||'';
    document.getElementById('er-nbpp').value=row.nb_pp_cousue||'';
    document.getElementById('er-dureemq').value=row.duree_mq_mp||'';
    document.getElementById('er-manqpers').value=row.manquant_pers||'';
    document.getElementById('er-comment-prod').value=row.comment||'';
  } else {
    document.getElementById('er-evttype').value=row.type||'';
    document.getElementById('er-evtof').value=row.of||'';
    document.getElementById('er-evtdeb').value=(row.debut||'').slice(0,5);
    document.getElementById('er-evtfin').value=(row.fin||'').slice(0,5);
    document.getElementById('er-evtcomment').value=row.comment||'';
    document.getElementById('er-horstrs').checked=row.hors_trs||false;
  }
  openM('m-editrow');
}

async function saveEditRow() {
  const rowNum=parseInt(document.getElementById('er-rownum').value);
  const rowType=document.getElementById('er-rowtype').value;
  const pwEl=document.getElementById('er-pw');
  const pwRow=document.getElementById('er-pw-row');
  const pw=pwEl.value||_adminPw;
  if(!rowNum){toast('Ligne invalide','err');return;}
  if(!pw){pwRow.style.display='';pwEl.focus();return;}
  const v=id=>document.getElementById(id)?.value||'';
  const n=id=>parseFloat(document.getElementById(id)?.value)||0;
  let updates={};
  if(rowType==='prod'){
    // Col indices per DECL_HEADERS (1-based):
    // 2=OF,4=Poste,5=Pilote,6=Co-Pilote,7=Nb Pers,8=Taille,9=Code Prod,10=Type Prod
    // 11=Poids,12=Fibre,13=OF Taie,14=Traca,15=Ref Taie,16=Kit
    // 17=Heure Debut,18=Heure Fin,20=Qte Fab,21=Qte Emb
    // 26=Qte Init Taie,27=Nb Taie 2nd,28=Nb Def Cout,29=Mq Taie,30=Mq Housse,31=Nb PP
    // 33=Manquant MP (duree),34=Manquant Pers,36=Commentaire
    updates={'2':v('er-of'),'4':v('er-poste'),'5':v('er-pilote'),'6':v('er-copilote'),
      '7':n('er-nbpers'),'8':v('er-taille'),'9':v('er-codeprod'),'10':v('er-typeprod'),
      '11':n('er-poids'),'12':v('er-fibre'),'13':v('er-oftaie'),'14':v('er-traca'),
      '15':v('er-reftaie'),'16':v('er-kit'),
      '17':v('er-deb'),'18':v('er-fin'),'20':n('er-qtefab'),'21':n('er-qteemb'),
      '26':n('er-qteinit'),'27':n('er-nbtaie2'),'28':n('er-nbdef'),
      '29':n('er-mqtaie'),'30':n('er-mqhousse'),'31':n('er-nbpp'),
      '33':n('er-dureemq'),'34':n('er-manqpers'),'36':v('er-comment-prod')};
  } else {
    updates={'1':v('er-evttype'),'2':v('er-evtof'),'17':v('er-evtdeb'),'18':v('er-evtfin'),
      '36':v('er-evtcomment'),'37':document.getElementById('er-horstrs').checked?'OUI':''};
  }
  const r=await fetch('/api/edit_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:rowNum,updates})});
  const d=r?await r.json():{};
  if(d&&d.ok){closeM('m-editrow');await loadMainDecl();if(_curTab==='history')await loadHist();loadKPI();toast('Ligne modifiée','ok');}
  else toast(d?.error||'Erreur modification','err');
}

async function deleteRow(key,rowNumId) {
  const rn=rowNumId?parseInt(document.getElementById(rowNumId)?.value):parseInt(window._rowMap[String(key)]?.row_num);
  if(!rn) return;
  const pwEl=document.getElementById('er-pw');
  const pwRow=document.getElementById('er-pw-row');
  const pw=(pwEl&&pwEl.value)||_adminPw;
  if(!pw){if(pwRow){pwRow.style.display='';pwEl&&pwEl.focus();}return;}
  if(!confirm('Supprimer cette ligne ?')) return;
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:rn})});
  const d=r?await r.json():{};
  if(d&&d.ok){closeM('m-editrow');await loadMainDecl();if(_curTab==='history')await loadHist();loadKPI();toast('Supprimé','ok');}
  else toast(d?.error||'Erreur suppression','err');
}

// ── EDIT STOP (recap prod view) ──
function buildEditStopOpts(){
  const s=document.getElementById('es-type');
  const s2=document.getElementById('er-evttype');
  const allOpts=[['Pause','_pause'],['Nettoyage','nettoyage'],...EVENTS.map(e=>[e[0],e[0]]),['Arrêt libre','Arrêt libre']];
  [s,s2].forEach(sel=>{
    if(!sel) return;
    allOpts.forEach(([lbl,val])=>{const o=document.createElement('option');o.value=val;o.textContent=lbl;sel.appendChild(o);});
  });
  // es-type uses key values (not display)
  if(s){s.innerHTML='';[['Pause','_pause'],['Nettoyage','nettoyage'],...EVENTS,['Arrêt libre','autre']].forEach(([lbl,key])=>{const o=document.createElement('option');o.value=key;o.textContent=lbl;s.appendChild(o);});}
}

function openEditStop(key){
  const ev=window._evMap[key];
  if(!ev) return;
  document.getElementById('es-key').value=key;
  document.getElementById('es-type').value=ev.type||key||'';
  const d=ev.debut||'',f2=ev.fin||'';
  document.getElementById('es-deb').value=d.length>=5?d.slice(0,5):d;
  document.getElementById('es-fin').value=f2.length>=5?f2.slice(0,5):f2;
  document.getElementById('es-cmt').value=ev.comment||'';
  openM('m-editstop');
}

async function saveEditStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev) return;
  const data={row_num:ev.row_num,type:document.getElementById('es-type').value,heure_debut:document.getElementById('es-deb').value,heure_fin:document.getElementById('es-fin').value,comment:document.getElementById('es-cmt').value};
  const r=await fetch('/api/edit_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  if(r&&r.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();loadKPI();toast('Modifié','ok');}
  else toast('Erreur','err');
}

async function deleteStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev||!confirm('Supprimer ?')) return;
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({row_num:ev.row_num})});
  if(r&&r.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();loadKPI();toast('Supprimé','ok');}
}

// ── TIMELINE ──
function renderTL(svgId,evts){
  const now=new Date(), s4h=new Date(now-4*3600*1000);
  drawTLFromISO(svgId,evts,s4h.toISOString(),now.toISOString());
}

function drawTL(svgId,tlEvts,debutHMS,finHMS){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=28,H=40;
  let html=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  if(!debutHMS||!finHMS){svg.innerHTML=html;return;}
  const base=new Date();base.setHours(0,0,0,0);
  const pHMS=s=>{const[h,m,sec]=(s||'').split(':');return base.getTime()+(+h||0)*3600000+(+m||0)*60000+(+sec||0)*1000;};
  const tS=pHMS(debutHMS),tE=pHMS(finHMS);
  const span=tE-tS;if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  html+=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  let cur=tS;
  (tlEvts||[]).forEach(e=>{
    if(!e.key||e.key==='prod') return;
    const cat=e.cat||'autre';
    const x1=toX(cur),x2=toX(cur+(e.dur_s||0)*1000);
    html+=`<rect x="${x1}" y="${Y}" width="${Math.max(1,x2-x1)}" height="${H2}" fill="${STOP_COL[cat]||'#94a3b8'}" rx="2" opacity=".9"/>`;
    cur+=(e.dur_s||0)*1000;
  });
  html+=`<text x="2" y="${H-1}" font-size="8" fill="#64748b">${debutHMS.slice(0,5)}</text>`;
  html+=`<text x="${W-30}" y="${H-1}" font-size="8" fill="#64748b">${finHMS.slice(0,5)}</text>`;
  svg.innerHTML=html;
}

function drawTLFromISO(svgId,evts,startIso,endIso,prodOfList){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=28,H=40;
  let html=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  const tS=new Date(startIso).getTime(),tE=new Date(endIso).getTime();
  const span=tE-tS;if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  // Draw historical prod blocks from of_list (fin de poste view)
  (prodOfList||[]).forEach(of=>{
    const t1=parseHMStoT(of.debut,of.date||null);
    const t2=of.fin?parseHMStoT(of.fin,of.date||null):null;
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  });
  // Prod background — show for current active OF
  if(ST.of_start_iso&&ST.prod_active){
    const ps=new Date(ST.of_start_iso).getTime();
    const x1=toX(ps),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  }
  // Events
  (evts||[]).forEach((ev,i)=>{
    const key=ev.debut||i;
    window._evMap[String(key)]=ev;
    const t1=parseHMStoT(ev.debut,ev.date),t2=parseHMStoT(ev.fin,ev.date);
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||tE);
    if(x2<=x1) return;
    const cat=ev.cat||'autre';
    html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${STOP_COL[cat]||'#94a3b8'}" rx="2" opacity=".85"/>`;
  });
  // Current live stop
  if(_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active){
    const se=_curStopElap+(Date.now()-_lastPoll)/1000;
    const sT=tE-se*1000;
    const x1=toX(sT),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${STOP_COL.pb||'#b91c1c'}" rx="2" opacity=".9"/>`;
  }
  const fT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  html+=`<text x="2" y="${H-1}" font-size="8" fill="#fff">${fT(tS)}</text>`;
  html+=`<text x="${W-30}" y="${H-1}" font-size="8" fill="#fff">${fT(tE)}</text>`;
  html+=`<line x1="${W/2}" y1="${Y}" x2="${W/2}" y2="${Y+H2}" stroke="#94a3b8" stroke-width=".5" stroke-dasharray="2,2"/>`;
  html+=`<text x="${W/2-10}" y="${H-1}" font-size="8" fill="#e2e8f0">${fT((tS+tE)/2)}</text>`;
  svg.innerHTML=html;
}

function parseHMStoT(hms,dateStr){
  if(!hms) return null;
  try{
    let base;
    if(dateStr){
      // Support both ISO (yyyy-mm-dd) and French (dd/mm/yyyy) formats
      const parts=dateStr.split('/');
      if(parts.length===3&&parts[2].length===4){
        base=new Date(parts[2]+'-'+parts[1]+'-'+parts[0]).getTime();
      } else {
        base=new Date(dateStr.slice(0,10)).getTime();
      }
    } else {
      const n=new Date();n.setHours(0,0,0,0);base=n.getTime();
    }
    const[h,m,s]=(hms||'').split(':').map(Number);return base+h*3600000+m*60000+(s||0)*1000;
  }catch(e){return null;}
}

// ── RECAP ──
function renderRecap(evts){
  const c=document.getElementById('recap-list');
  if(!c) return;
  window._evMap={};
  const stops=(evts||[]).filter(e=>e.type);
  if(!stops.length){c.innerHTML='<div style="color:var(--gray);font-size:10px;padding:4px">Aucun arrêt</div>';return;}
  let html='';
  stops.forEach((ev,i)=>{
    const key=ev.debut||i;
    window._evMap[String(key)]=ev;
    const cat=ev.cat||'autre';
    const col=STOP_COL[cat]||'#94a3b8';
    html+=`<div class="si">
      <div class="sdot" style="background:${col}"></div>
      <div class="si-nm">${ev.type||'?'}</div>
      <div class="si-dur">${ev.duree||calcDur(ev.debut,ev.fin)||'—'}</div>
      <button class="btn-edit" onclick="openEditStop('${esc(String(key))}')">✏</button>
    </div>`;
  });
  if(_curStopKey){
    const lbl=_curStopKey==='_pause'?'Pause':getEvtLabel(_curStopKey);
    const col=_curStopKey==='_pause'?STOP_COL['_pause']:STOP_COL.pb;
    html+=`<div class="si" style="animation:blink .85s step-start infinite">
      <div class="sdot" style="background:${col}"></div>
      <div class="si-nm" style="font-size:9px">${lbl}</div>
      <div class="si-dur" style="font-size:9px">…</div>
    </div>`;
  }
  c.innerHTML=html;
}

function calcDur(d,f){
  if(!d||!f) return '';
  try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
  const diff=p(f)-p(d);return diff>0?fmtDur(diff):'';}catch(e){return '';}
}

// ── FIN DE POSTE ──
async function doFinPoste(){
  if(ST.prod_active){toast('Terminer la production en cours avant de finir le poste','err');return;}
  goTab('finposte');
}

async function loadFPData(){
  const d=await apiFetch('/api/fin_poste_data');
  if(!d) return;
  document.getElementById('fp-trs').textContent=fmtTRS(d.trs_shift!==undefined?d.trs_shift:d.trs);
  document.getElementById('fp-trs-of').textContent=fmtTRS(d.trs);
  document.getElementById('fp-eq').textContent=(d.tot_equiv||0).toFixed(1);
  document.getElementById('fp-nof').textContent=d.nb_of||0;
  document.getElementById('fp-prod-t').textContent=Math.round((d.tot_s||0)/60)+' min';
  document.getElementById('fp-who').textContent=`${d.pilot||ST.pilot||''} — ${ST.poste||''}`;
  const now2=new Date();
  const fpDateEl=document.getElementById('fp-date');
  if(fpDateEl) fpDateEl.textContent=String(now2.getDate()).padStart(2,'0')+'/'+String(now2.getMonth()+1).padStart(2,'0')+'/'+now2.getFullYear();
  // Populate model selector
  const fpSel=document.getElementById('fp-model-sel');
  if(fpSel&&_cfgModels&&_cfgModels.length){
    fpSel.innerHTML='<option value="">— Modèle —</option>';
    _cfgModels.forEach(m=>{const o=document.createElement('option');o.value=m.nom||'';o.textContent=m.nom||'';if(m.nom===ST.poste)o.selected=true;fpSel.appendChild(o);});
  }
  window._fpData=d;

  // Count stops
  const stops=gEvts.filter(e=>e.type);
  const stopTotal=stops.reduce((a,e)=>{
    try{const p=s=>s.split(':').reduce((acc,v,i)=>acc+(i===0?+v*3600:i===1?+v*60:+v),0);
    return a+Math.max(0,p(e.fin||'00:00:00')-p(e.debut||'00:00:00'));}catch(ex){return a;}
  },0);
  document.getElementById('fp-stop-t').textContent=Math.round(stopTotal/60)+' min';

  // Timeline — refresh events first, then combine historical + live
  await pollEvts();
  const shiftStart=d.shift_start_iso||new Date(Date.now()-8*3600*1000).toISOString();
  // Use gEvts (historical) + convert live tl_events to display format
  const liveEvts=tlEventsToDisplayFmt(ST.tl_events||[]);
  const allEvtsForTL=[...gEvts,...liveEvts];
  // Pre-fill datetime-local inputs — use model horaire theoretical times when available
  const toLocalDT=dt=>{const y=dt.getFullYear(),mo=String(dt.getMonth()+1).padStart(2,'0'),dy=String(dt.getDate()).padStart(2,'0'),h=String(dt.getHours()).padStart(2,'0'),mi=String(dt.getMinutes()).padStart(2,'0');return `${y}-${mo}-${dy}T${h}:${mi}`;};
  const dayKeys2=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk2=dayKeys2[new Date().getDay()];
  const modelFP=_cfgModels&&_cfgModels.find(m=>m.nom===ST.poste);
  const jourFP=modelFP&&modelFP.jours&&modelFP.jours[dk2];
  // Debut: prefer model debut, fallback to shift_start_iso
  let debDt=new Date(shiftStart);
  if(jourFP&&jourFP.debut){const[hh,mm]=jourFP.debut.split(':').map(Number);debDt=new Date();debDt.setHours(hh,mm,0,0);}
  // Fin: prefer model fin, fallback to now
  let finDt=new Date();
  if(jourFP&&jourFP.fin){const[hh,mm]=jourFP.fin.split(':').map(Number);finDt=new Date();finDt.setHours(hh,mm,0,0);}
  const fpDeb=document.getElementById('fp-debut-dt');
  const fpFin=document.getElementById('fp-fin-dt');
  if(fpDeb) fpDeb.value=toLocalDT(debDt);
  if(fpFin) fpFin.value=toLocalDT(finDt);
  // Show model horaire info (in header and gauge footer)
  const fpShiftInfo=document.getElementById('fp-shift-info');
  const fpShiftH=document.getElementById('fp-shift-hours');
  const fpHoraireDisp=document.getElementById('fp-horaire-display');
  const hStr2=jourFP&&jourFP.debut&&jourFP.fin?`${jourFP.debut} → ${jourFP.fin}`:'—';
  if(fpShiftInfo) fpShiftInfo.textContent=hStr2;
  if(fpShiftH) fpShiftH.textContent=hStr2;
  if(fpHoraireDisp) fpHoraireDisp.textContent=hStr2 !== '—' ? `Poste : ${hStr2}` : '—';
  drawTLFromISO('fp-tl',allEvtsForTL,debDt.toISOString(),finDt.toISOString(),d.of_list||[]);

  // Productions
  const fpb=document.getElementById('fp-prods');
  if(fpb&&d.of_list){
    fpb.innerHTML=d.of_list.map(p=>`
      <tr>
        <td style="font-weight:600">${esc(p.of||'')}</td>
        <td>${esc(p.taille||'')}</td>
        <td>${esc(String(p.qte_fab||0))}</td>
        <td>${esc(String(p.equiv||''))}</td>
        <td>${esc(p.duree||'')}</td>
        <td class="${(p.trs||0)>=90?'tg':(p.trs||0)>=75?'tm':'tb'}">${fmtTRS(p.trs||0)}</td>
      </tr>`).join('')||'<tr><td colspan="6" style="color:var(--gray)">Aucune production</td></tr>';
  }

  // Stops list
  const fsl=document.getElementById('fp-stops-list');
  if(fsl){
    const stopsByType={};
    gEvts.forEach(e=>{if(!e.type)return;if(!stopsByType[e.type])stopsByType[e.type]=0;
    try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
    stopsByType[e.type]+=Math.max(0,p(e.fin||'00:00:00')-p(e.debut||'00:00:00'));}catch(ex){}});
    fsl.innerHTML=Object.entries(stopsByType).map(([t,s])=>`
      <div class="flex" style="padding:4px 0;border-bottom:1px solid var(--border);font-size:12px">
        <span style="flex:1;font-weight:600">${esc(t)}</span>
        <span style="color:var(--gray)">${Math.round(s/60)} min</span>
      </div>`).join('')||'<span style="color:var(--gray);font-size:11px">Aucun arrêt</span>';
  }

  // Graphs
  const prodS=d.tot_s||0;
  const stopS=stopTotal;
  drawPie('fp-pie',[
    {label:'Prod',value:prodS,color:'#16a34a'},
    {label:'Arrêts',value:stopS,color:'#dc2626'},
  ]);
  const trsS=d.trs_shift!==undefined?d.trs_shift:d.trs;
  drawGauge('fp-gauge-arc','fp-gauge-pct',trsS>=0?trsS:0);
  const fpL=document.getElementById('fp-trs-lbl2');if(fpL) fpL.textContent=fmtTRSv(trsS);
}

function applyFPHoraires(){
  const debStr=document.getElementById('fp-debut-dt').value;
  const finStr=document.getElementById('fp-fin-dt').value;
  if(!debStr||!finStr){toast('Renseigner début et fin','err');return;}
  const deb=new Date(debStr),fin=new Date(finStr);
  if(fin<=deb){toast('Fin doit être après début','err');return;}
  const shiftS=(fin.getTime()-deb.getTime())/1000;
  const shiftInfo=document.getElementById('fp-shift-info');
  const pad=n=>String(n).padStart(2,'0');
  const fmt=d=>pad(d.getHours())+':'+pad(d.getMinutes());
  if(shiftInfo) shiftInfo.textContent=fmt(deb)+'→'+fmt(fin);
  // Recalc TRS
  const d=window._fpData;
  if(d){
    const totEquiv=d.tot_equiv||0;
    const prodRef=d.prod_ref||ST.prod_ref||200;
    const trs=shiftS>0&&prodRef>0?Math.round(totEquiv/(prodRef*shiftS/28800)*1000)/10:0;
    document.getElementById('fp-trs').textContent=fmtTRS(trs);
  }
  // Redraw timeline with custom range
  drawTLFromISO('fp-tl',gEvts,deb.toISOString(),fin.toISOString(),(window._fpData&&window._fpData.of_list)||[]);
  closeM('m-fp-horaires');
  toast('Horaires appliqués','ok');
}

function recalcFPTRS(){
  const sel=document.getElementById('fp-model-sel');
  if(!sel||!window._fpData) return;
  const modelNom=sel.value;
  if(!modelNom) return;
  const model=_cfgModels.find(m=>m.nom===modelNom);
  if(!model) return;
  // Show model shift info
  const now=new Date();
  const dayKeys=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=dayKeys[now.getDay()];
  const jours=model.jours||{};
  const jour=jours[dk]||{};
  const shiftInfo=document.getElementById('fp-shift-info');
  const shiftHours=document.getElementById('fp-shift-hours');
  const fpHD=document.getElementById('fp-horaire-display');
  const hStr=jour.debut&&jour.fin?`${jour.debut} → ${jour.fin}`:'—';
  if(shiftInfo) shiftInfo.textContent=hStr;
  if(shiftHours) shiftHours.textContent=hStr;
  if(fpHD) fpHD.textContent=hStr !== '—' ? `Poste : ${hStr}` : '—';
  // Recalculate TRS using this model's shift duration
  if(jour.debut&&jour.fin){
    const p=s=>{const[h,m]=s.split(':').map(Number);return h*3600+m*60;};
    let shiftS=p(jour.fin)-p(jour.debut);
    if(shiftS<0) shiftS+=86400;
    const d=window._fpData;
    const totEquiv=d.tot_equiv||0;
    const prodRef=d.prod_ref||ST.prod_ref||200;
    const trs=shiftS>0&&prodRef>0?Math.round(totEquiv/(prodRef*shiftS/28800)*1000)/10:0;
    document.getElementById('fp-trs').textContent=fmtTRS(trs);
  }
}

async function confirmFinPoste(){
  const stops=gEvts.filter(e=>e.type);
  const pSec=s=>s?s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0):0;
  let arret_s=0,pause_s=0,nett_s=0;
  stops.forEach(e=>{
    const dur=Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0'));
    const t=(e.type||'').toLowerCase();
    if(t.includes('pause')) pause_s+=dur;
    else if(t.includes('nett')) nett_s+=dur;
    else arret_s+=dur;
  });
  const fpData=await apiFetch('/api/fin_poste_data');
  const ofList=fpData&&fpData.of_list||[];
  const prod_total=ofList.reduce((s,o)=>s+parseInt(o.qte_fab||0),0);
  const dur_prod_total_s=ofList.reduce((s,o)=>s+pSec(o.duree||''),0);
  const dur_prod_sans_arret_s=Math.max(0,dur_prod_total_s-arret_s);
  const posteRow={
    date:new Date().toLocaleDateString('fr-FR'),
    pilot:ST.pilot||'',
    copilote:ST.form&&ST.form.copilote||'',
    poste:ST.poste||'',
    nb_of:fpData?fpData.nb_of||0:0,
    prod_total,
    tot_equiv:fpData?fpData.tot_equiv||0:0,
    trs_shift:fpData?fpData.trs_shift||0:0,
    arret_min:Math.round(arret_s/60),
    pause_min:Math.round(pause_s/60),
    nett_min:Math.round(nett_s/60),
    dur_prod_total_min:Math.round(dur_prod_total_s/60),
    dur_prod_sans_arret_min:Math.round(dur_prod_sans_arret_s/60),
    dur_poste_theorique_min:(()=>{
      const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
      const dk=DAY_KEYS[new Date().getDay()];
      const model=_cfgModels.find(m=>m.nom===(ST.poste||''));
      if(model&&model.jours&&model.jours[dk]){
        const j=model.jours[dk];
        if(j.debut&&j.fin){
          const toS=s=>{const[h,m2]=s.split(':').map(Number);return h*3600+m2*60;};
          let s=toS(j.fin)-toS(j.debut);if(s<0)s+=86400;
          return Math.round(s/60);
        }
      }
      return 0;
    })(),
    comment:''
  };
  await fetch('/api/save_poste',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(posteRow)});
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  resetToLogin();
  toast('Bonne fin de poste !','ok');
}

// ── KPI TAB ──
function _kpiGauge(arcId,pctId,trs,pArc){
  // pArc=157 for bigger viewBox (radius 50), 132 for radius 42
  const P=pArc||157;
  const arc=document.getElementById(arcId),pct=document.getElementById(pctId);if(!arc||!pct)return;
  const v=Math.max(0,Math.min(100,trs||0)),dash=(v/100)*P;
  const col=v>=90?'#16a34a':v>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash.toFixed(1)},${P}`);arc.setAttribute('stroke',col);
  pct.textContent=trs>=0?fmtTRS(trs):'--%';pct.setAttribute('fill',col);
}
function _kpiNoData(arcId,pctId,pArc){
  const arc=document.getElementById(arcId),pct=document.getElementById(pctId);if(!arc||!pct)return;
  arc.setAttribute('stroke-dasharray',`0,${pArc||157}`);arc.setAttribute('stroke','#334155');
  pct.textContent='--%';pct.setAttribute('fill','#475569');
}
function _drawKpiTL(svgId,evts,startISO,endISO,isCurrent){
  const svg=document.getElementById(svgId);if(!svg)return;
  const H=isCurrent?50:36,W=800;
  const tS=new Date(startISO).getTime(),tE=new Date(endISO).getTime();
  const span=tE-tS;
  if(span<=0){svg.innerHTML=`<rect x="0" y="0" width="${W}" height="${H}" fill="#1e293b" rx="4"/>`;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  let html=`<rect x="0" y="0" width="${W}" height="${H}" fill="#1e2d48" rx="4"/>`;
  // Prod background (green)
  html+=`<rect x="0" y="0" width="${W}" height="${H}" fill="#166534" rx="4" opacity=".65"/>`;
  // Events (stops)
  (evts||[]).forEach(ev=>{
    const t1=parseHMStoT(ev.debut,ev.date),t2=parseHMStoT(ev.fin,ev.date);
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||(t1+300000));if(x2<=x1)return;
    const cat=ev.cat||'autre';const col=STOP_COL[cat]||'#94a3b8';
    html+=`<rect x="${x1}" y="0" width="${Math.max(2,x2-x1)}" height="${H}" fill="${col}" rx="2" opacity=".9"/>`;
  });
  // Current live stop for active session
  if(isCurrent&&_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active){
    const se=_curStopElap+(Date.now()-_lastPoll)/1000;
    const sT=tE-se*1000;
    const x1=toX(sT),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="0" width="${x2-x1}" height="${H}" fill="#dc2626" rx="2" opacity=".9"/>`;
  }
  const fT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  const fs=isCurrent?11:9;
  html+=`<text x="4" y="${H-4}" font-size="${fs}" fill="#fff">${fT(tS)}</text>`;
  html+=`<text x="${W-40}" y="${H-4}" font-size="${fs}" fill="#fff">${fT(tE)}</text>`;
  if(span>3600000){const mid=(tS+tE)/2;html+=`<line x1="${W/2}" y1="0" x2="${W/2}" y2="${H}" stroke="#475569" stroke-width="1" stroke-dasharray="3,3"/><text x="${W/2-18}" y="${H-4}" font-size="${fs}" fill="#e2e8f0">${fT(mid)}</text>`;}
  svg.innerHTML=html;
}

async function loadKPI(){
  const [histData,evtData,todayData]=await Promise.all([
    apiFetch('/api/history'),
    apiFetch('/api/events_list'),
    apiFetch('/api/history_today')
  ]);
  const rows=Array.isArray(histData)?histData:[];
  const evts=Array.isArray(evtData)?evtData:[];
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0'),yyyy=now.getFullYear();
  const todayPfx=dd+'/'+mm;
  const todayFR=dd+'/'+mm+'/'+yyyy;
  const curPilot=ST.pilot||'';
  const pSec=s=>s?s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0):0;

  // ── Poste actuel ──
  const curEvts=evts.filter(e=>e.date&&e.date.startsWith(todayPfx)&&(!curPilot||!e.pilote||e.pilote===curPilot));
  const curStopS=curEvts.reduce((a,e)=>a+Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0')),0);
  if(todayData){
    // Même formule que la jauge accueil : equiv / (prod_ref * elapsed / 28800)
    let trs=-1;
    if(_todayEquivAccum>0&&_shiftRefDt&&ST.prod_ref>0){
      const refTime=_lastProdDeclTime||new Date();
      const shiftElap=(refTime.getTime()-_shiftRefDt.getTime())/1000;
      if(shiftElap>0) trs=Math.round(_todayEquivAccum/(ST.prod_ref*shiftElap/28800)*100*10)/10;
    }
    _kpiGauge('kpi-g0-arc','kpi-g0-pct',trs>=0?trs:0,157);
    const cde=document.getElementById('kpi-cur-date');
    if(cde){const si=ST.shift_start_iso;const st=si?new Date(si).toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'}):null;cde.textContent=todayFR+(st?' depuis '+st:'');}
    const csu=document.getElementById('kpi-cur-sub');
    if(csu) csu.textContent=(todayData.nb_of||0)+' OF  ·  '+Math.round(curStopS/60)+' min arrêts';
  } else {
    _kpiNoData('kpi-g0-arc','kpi-g0-pct',157);
  }

  // ── 3 postes précédents ──
  const sessions={};
  rows.forEach(r=>{
    if(r.date&&r.date.startsWith(todayPfx)&&r.pilote===curPilot) return;
    const key=(r.pilote||'?')+'|'+(r.date||'?')+'|'+(r.poste||'?');
    if(!sessions[key]) sessions[key]={pilot:r.pilote||'?',poste:r.poste||'?',date:r.date||'?',rows:[],trs_sum:0,trs_cnt:0};
    sessions[key].rows.push(r);
    const t=parseFloat(r.trs||0);if(t>0){sessions[key].trs_sum+=t;sessions[key].trs_cnt++;}
  });
  const sessArr=Object.values(sessions).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,3);
  for(let i=0;i<3;i++){
    const arcId='kpi-g'+(i+1)+'-arc',pctId='kpi-g'+(i+1)+'-pct';
    const lbl=document.getElementById('kpi-p'+(i+1)+'-lbl');
    const dat=document.getElementById('kpi-p'+(i+1)+'-date');
    const sub=document.getElementById('kpi-p'+(i+1)+'-sub');
    if(sessArr[i]){
      const s=sessArr[i];
      const avgT=s.trs_cnt>0?s.trs_sum/s.trs_cnt:-1;
      const sEvts=evts.filter(e=>e.date===s.date&&(!s.pilot||!e.pilote||e.pilote===s.pilot));
      const sStopS=sEvts.reduce((a,e)=>a+Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0')),0);
      if(lbl) lbl.textContent=s.pilot+' — '+s.poste;
      if(dat){const debs=s.rows.map(r=>r.debut||'').filter(Boolean).sort();const fins=s.rows.map(r=>r.fin||'').filter(Boolean).sort().reverse();dat.textContent=s.date+(debs.length&&fins.length?' '+debs[0].slice(0,5)+'→'+fins[0].slice(0,5):'');}
      if(sub) sub.textContent=s.rows.length+' OF  ·  '+Math.round(sStopS/60)+' min arrêts';
      if(avgT>=0) _kpiGauge(arcId,pctId,avgT,157);
      else _kpiNoData(arcId,pctId,157);
    } else {
      if(lbl) lbl.textContent=['Poste précédent','Avant-dernier','Il y a 3 postes'][i];
      if(dat) dat.textContent=''; if(sub) sub.textContent='—';
      _kpiNoData(arcId,pctId,157);
    }
  }

  // ── 4 Timelines ──
  // TL0 : poste actuel
  const shiftStart=(_shiftRefDt?_shiftRefDt.toISOString():null)||ST.shift_start_iso||new Date(now.getTime()-8*3600*1000).toISOString();
  const tl0Lbl=document.getElementById('kpi-tl0-lbl');
  if(tl0Lbl) tl0Lbl.textContent='Poste actuel'+(ST.pilot?' — '+ST.pilot:'')+(ST.poste?' ('+ST.poste+')':'');
  _drawKpiTL('kpi-tl0',[...curEvts,...tlEventsToDisplayFmt(ST.tl_events||[])],shiftStart,now.toISOString(),true);
  // TL1-3 : sessions précédentes
  for(let i=0;i<3;i++){
    const svgId='kpi-tl'+(i+1);
    const lblId='kpi-tl'+(i+1)+'-lbl';
    const lbl=document.getElementById(lblId);
    if(sessArr[i]){
      const s=sessArr[i];
      if(lbl) lbl.textContent=s.pilot+' — '+s.poste+' ('+s.date+')';
      const sEvts=evts.filter(e=>e.date===s.date&&(!s.pilot||!e.pilote||e.pilote===s.pilot));
      const debs=s.rows.map(r=>r.debut||'').filter(Boolean).sort();
      const fins=s.rows.map(r=>r.fin||'').filter(Boolean).sort().reverse();
      const dateStr=s.date;
      // Compute start/end from model horaire for this session's poste+day
      const _DK=['dim','lun','mar','mer','jeu','ven','sam'];
      let sessionStartMs=null,sessionEndMs=null;
      if(dateStr&&s.poste){
        const _p=dateStr.split('/');
        if(_p.length===3){
          const _sd=new Date(parseInt(_p[2]),parseInt(_p[1])-1,parseInt(_p[0]));
          const _dk=_DK[_sd.getDay()];
          const _sm=_cfgModels&&_cfgModels.find(m=>m.nom===s.poste);
          const _sj=_sm&&_sm.jours&&_sm.jours[_dk];
          const _sdeb=_sj&&_sj.debut;const _sfin=_sj&&_sj.fin;
          if(_sdeb){const[_sh,_smm]=_sdeb.split(':').map(Number);sessionStartMs=new Date(parseInt(_p[2]),parseInt(_p[1])-1,parseInt(_p[0]),_sh,_smm,0).getTime();}
          if(_sfin){const[_fh,_fm]=_sfin.split(':').map(Number);sessionEndMs=new Date(parseInt(_p[2]),parseInt(_p[1])-1,parseInt(_p[0]),_fh,_fm,0).getTime();if(sessionEndMs<=sessionStartMs)sessionEndMs+=86400000;}
        }
      }
      if(sessionStartMs&&sessionEndMs){
        _drawKpiTL(svgId,sEvts,new Date(sessionStartMs).toISOString(),new Date(sessionEndMs).toISOString(),false);
      } else if(debs.length&&fins.length){
        const baseMs=parseHMStoT(debs[0],dateStr)||new Date().setHours(5,0,0,0);
        const endMs=sessionEndMs||parseHMStoT(fins[0],dateStr)||(baseMs+8*3600000);
        const startMs=sessionStartMs||(baseMs-600000);
        _drawKpiTL(svgId,sEvts,new Date(startMs).toISOString(),new Date(endMs).toISOString(),false);
      } else {
        _drawKpiTL(svgId,[],new Date().toISOString(),new Date(Date.now()+3600000).toISOString(),false);
      }
      // Mini-pareto for this session
      const _pEl=document.getElementById('kpi-p'+(i+1)+'-pareto');
      if(_pEl){
        const _sm={};sEvts.forEach(e=>{if(!e.type)return;const d=Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0'));_sm[e.type]=(_sm[e.type]||0)+d;});
        const _srt=Object.entries(_sm).sort((a,b)=>b[1]-a[1]).slice(0,4);
        if(_srt.length){
          const _mx=_srt[0][1];
          _pEl.innerHTML=_srt.map(([t,s])=>{
            const pct=Math.round(s/_mx*100),min=Math.round(s/60);
            const cat=sEvts.find(e=>e.type===t)?.cat||'autre';
            const col=STOP_COL[cat]||'#94a3b8';
            return `<div style="display:flex;align-items:center;gap:3px;font-size:9px;color:#64748b">
              <div style="width:${Math.round(pct*0.4)}px;min-width:3px;height:8px;background:${col};border-radius:2px"></div>
              <span style="max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(t)}</span>
              <span style="color:#94a3b8">${min}m</span></div>`;
          }).join('');
        } else { _pEl.innerHTML=''; }
      }
    } else {
      if(lbl) lbl.textContent='—';
      const svg=document.getElementById(svgId);
      if(svg) svg.innerHTML='<rect x="0" y="0" width="800" height="36" fill="#1e293b" rx="4"/><text x="400" y="22" text-anchor="middle" font-size="10" fill="#334155">Aucune donnée</text>';
      const _pEl2=document.getElementById('kpi-p'+(i+1)+'-pareto');
      if(_pEl2) _pEl2.innerHTML='';
    }
  }

  // ── Camembert Prod/Arrêts ──
  const prodS=(todayData&&todayData.tot_s)||0;
  // Use dark pie colors
  const kpiPieSvg=document.getElementById('kpi-pie');
  if(kpiPieSvg){
    const total=prodS+curStopS;
    if(total<=0){kpiPieSvg.innerHTML='<text x="65" y="60" text-anchor="middle" font-size="11" fill="#475569">Pas de données</text>';}
    else{
      const segs=[{label:'Prod',value:prodS,color:'#16a34a'},{label:'Arrêts',value:curStopS,color:'#dc2626'}];
      const cx=65,cy=57,r=44,ir=24;let html='',startAngle=-Math.PI/2;
      segs.forEach(seg=>{
        if(seg.value<=0)return;
        const angle=(seg.value/total)*2*Math.PI;if(angle<0.001)return;
        const endAngle=startAngle+angle,large=angle>Math.PI?1:0;
        const x1=(cx+r*Math.cos(startAngle)).toFixed(2),y1=(cy+r*Math.sin(startAngle)).toFixed(2);
        const x2=(cx+r*Math.cos(endAngle)).toFixed(2),y2=(cy+r*Math.sin(endAngle)).toFixed(2);
        const ix1=(cx+ir*Math.cos(startAngle)).toFixed(2),iy1=(cy+ir*Math.sin(startAngle)).toFixed(2);
        const ix2=(cx+ir*Math.cos(endAngle)).toFixed(2),iy2=(cy+ir*Math.sin(endAngle)).toFixed(2);
        html+=`<path d="M${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} L${ix2},${iy2} A${ir},${ir} 0 ${large},0 ${ix1},${iy1} Z" fill="${seg.color}"/>`;
        startAngle=endAngle;
      });
      const pp=total>0?Math.round(prodS/total*100):0;
      html+=`<text x="${cx}" y="${cy+6}" text-anchor="middle" font-size="16" font-weight="800" fill="#fff">${pp}%</text>`;
      html+=`<text x="${cx}" y="${cy+18}" text-anchor="middle" font-size="8" fill="#94a3b8">Prod</text>`;
      let lx=0;segs.forEach(s=>{const p=Math.round(s.value/total*100);html+=`<rect x="${lx}" y="108" width="9" height="9" fill="${s.color}" rx="1"/>`;html+=`<text x="${lx+12}" y="116" font-size="8" fill="#94a3b8">${esc(s.label)} ${p}%</text>`;lx+=65;});
      kpiPieSvg.innerHTML=html;
    }
  }

  // ── Stats ──
  const statsEl=document.getElementById('kpi-stats');
  if(statsEl){
    const mkStat=(lbl,val,col)=>`<div style="background:#1e293b;border-radius:8px;padding:8px 4px;text-align:center">
      <div style="font-size:17px;font-weight:800;color:${col||'#e2e8f0'}">${val}</div>
      <div style="font-size:9px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.5px;margin-top:2px;line-height:1.2">${lbl}</div>
    </div>`;
    if(todayData){
      const rows=todayData.rows||[];
      const qteProd=Math.round(rows.reduce((a,r)=>a+parseFloat(r.qte_fab||0),0));
      const prodMin=Math.round(rows.reduce((a,r)=>{const d=pSec(r.debut||'0:0:0'),f=pSec(r.fin||'0:0:0');return a+Math.max(0,f-d)/60;},0));
      const cad=prodMin>0?Math.round(qteProd/prodMin*60):0;
      const elapsedMin=_shiftRefDt?Math.round((now.getTime()-_shiftRefDt.getTime())/60000):0;
      statsEl.innerHTML=
        mkStat('Qté produite',qteProd,'#a78bfa')+
        mkStat('Qté équiv.',(todayData.tot_equiv||0).toFixed(1),'#93c5fd')+
        mkStat('Cadence pcs/h',cad,'#6ee7b7')+
        mkStat('Arrêts min',Math.round(curStopS/60),'#fca5a5')+
        mkStat('Prod min',prodMin,'#86efac')+
        mkStat('Écoulé min',elapsedMin+'','#fde68a');
    } else {
      statsEl.innerHTML='<div style="grid-column:1/-1;color:#475569;font-size:13px;text-align:center;padding:12px">Aucun poste actif</div>';
    }
  }

  // ── Pareto ──
  const stopMap={};const stopCat={};
  curEvts.forEach(e=>{if(!e.type)return;const dur=Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0'));stopMap[e.type]=(stopMap[e.type]||0)+dur;if(!stopCat[e.type])stopCat[e.type]=e.cat||'autre';});
  const sorted=Object.entries(stopMap).sort((a,b)=>b[1]-a[1]);
  const maxS=sorted.length?sorted[0][1]:1;
  const par=document.getElementById('kpi-pareto');
  if(par){
    if(!sorted.length){par.innerHTML='<div style="color:#475569;font-size:13px;padding:8px">Aucun arrêt ce poste</div>';}
    else{par.innerHTML=sorted.map(([type,s])=>{
      const pct=Math.round(s/maxS*100),min=Math.round(s/60);
      const col=STOP_COL[stopCat[type]]||'#94a3b8';
      const pctTotal=curStopS>0?Math.round(s/curStopS*100):0;
      return `<div style="display:flex;flex-direction:column;gap:3px">
        <div style="display:flex;justify-content:space-between;font-size:12px;font-weight:600;color:#cbd5e1">
          <span>${esc(type)}</span><span style="color:#94a3b8">${min} min (${pctTotal}%)</span>
        </div>
        <div style="background:#1e293b;border-radius:4px;height:16px;overflow:hidden">
          <div style="width:${pct}%;background:${col};height:100%;border-radius:4px;transition:width .4s"></div>
        </div>
      </div>`;
    }).join('');}
  }
}

// ── HISTORY ──
async function loadHist(){
  const today=new Date().toISOString().slice(0,10);
  const from=document.getElementById('hist-from').value||today;
  const to=document.getElementById('hist-to').value||today;
  const [declData,evtData]=await Promise.all([
    apiFetch(`/api/history?from=${from}&to=${to}`),
    apiFetch('/api/events_list')
  ]);
  const decls=Array.isArray(declData)?declData:(declData&&declData.rows?declData.rows:[]);
  const evts=Array.isArray(evtData)?evtData:[];
  // Filter events to the selected date range
  const fromD=new Date(from+'T00:00:00'),toD=new Date(to+'T23:59:59');
  const fmtFR=d=>{const[y,m,dy]=d.split('-');return `${dy}/${m}/${y}`;};
  const fmtFRfrom=fmtFR(from),fmtFRto=fmtFR(to);
  const evtsFiltered=evts.filter(e=>{
    if(!e.date) return false;
    // Support both dd/mm/yyyy and yyyy-mm-dd
    let dt;
    const parts=e.date.split('/');
    if(parts.length===3&&parts[2].length===4) dt=new Date(parts[2]+'-'+parts[1]+'-'+parts[0]);
    else dt=new Date(e.date);
    return dt>=fromD&&dt<=toD;
  });
  const hd=document.getElementById('hist-hd'),bd=document.getElementById('hist-bd');
  if(!hd||!bd) return;
  // Build unified row list same as loadMainDecl
  const allRows=[];
  decls.forEach(r=>allRows.push({...r,_rowType:'prod'}));
  evtsFiltered.forEach(r=>allRows.push({...r,_rowType:'evt'}));
  allRows.sort((a,b)=>{
    const da=a.date||'',db=b.date||'';
    if(da!==db) return db.localeCompare(da);
    return (b.debut||'').localeCompare(a.debut||'');
  });
  hd.innerHTML='<th>Type</th><th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th><th>Début</th><th>Fin</th><th>Détails</th><th>Qté/Durée</th><th>TRS/Info</th><th>Actions</th>';
  if(!allRows.length){bd.innerHTML='<tr><td colspan="11" style="text-align:center;color:var(--gray);padding:16px">Aucune donnée sur cette période</td></tr>';return;}
  window._rowMap=window._rowMap||{};
  bd.innerHTML=allRows.map(r=>{
    const key=r.row_num||r.debut;
    window._rowMap[String(key)]=r;
    const isProd=r._rowType==='prod';
    const t=parseFloat(r.trs||0);
    const tag=isProd?'<span class="row-tag tag-p">🏭 Prod</span>':(r.type&&r.type.toLowerCase().includes('nett')?'<span class="row-tag tag-n">🧹 Nett.</span>':'<span class="row-tag tag-e">⛔ Arrêt</span>');
    const details=isProd?esc(r.taille||''):esc(r.type||'');
    const qty=isProd?esc(String(r.qte_fab||'')):esc(r.duree||'');
    const info=isProd&&t>0?`<span class="${t>=90?'tg':t>=75?'tm':'tb'}">${fmtTRS(t)}</span>`:`<span style="color:var(--gray);font-size:10px">${esc(r.comment||'')}</span>`;
    return `<tr class="${isProd?'row-prod':'row-evt'}">
      <td>${tag}</td><td style="font-weight:600">${esc(r.of||'')}</td>
      <td style="font-size:10px">${esc(r.date||'')}</td><td style="font-size:10px">${esc(r.poste||'')}</td>
      <td>${esc(r.pilote||'')}</td><td>${esc(r.debut||'')}</td><td>${esc(r.fin||'')}</td>
      <td style="font-size:11px">${details}</td><td style="font-size:11px">${qty}</td><td>${info}</td>
      <td><button onclick="openEditRow('${esc(String(key))}')" style="background:#6366f1;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:15px;cursor:pointer;font-weight:700" title="Modifier">✏</button></td>
    </tr>`;
  }).join('');
}

// ── SETTINGS ──
async function setDbPath(){
  const path=document.getElementById('cfg-db-path').value.trim();
  if(!path){toast('Chemin requis','err');return;}
  const st=document.getElementById('cfg-db-status');
  if(st){st.textContent='Enregistrement…';st.style.color='var(--amber)';}
  const r=await fetch('/api/set_db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,path})});
  const d=r?await r.json():{};
  if(d&&d.ok){
    if(st){st.textContent='✓ Fichier Excel configuré';st.style.color='var(--green)';}
    toast('Fichier Excel enregistré','ok');
    loadCfg();
  } else {
    if(st){st.textContent='✗ '+(d&&d.error||'Erreur — vérifiez le chemin');st.style.color='var(--red)';}
    toast(d&&d.error||'Chemin invalide','err');
  }
}

async function loadCfg(){
  const d=await apiFetch('/api/config');
  if(!d) return;
  _cfgPwds=d.pilot_passwords||{};
  _cfgModels=d.modeles_horaires||[];
  const prEl=document.getElementById('cfg-pr');
  if(prEl) prEl.value=d.prod_ref||200;
  // Show current db path
  const dbEl=document.getElementById('cfg-db-path');
  const dbSt=document.getElementById('cfg-db-status');
  if(dbEl&&d.db_path) dbEl.value=d.db_path;
  if(dbSt&&d.db_name){dbSt.textContent='Fichier actuel : '+d.db_name;dbSt.style.color='var(--green)';}
  renderPwdList();
  renderModelList();
  await loadEvtsList();
  await loadInterposteCfg();
  // Refresh login model dropdown
  const sel=document.getElementById('ln-model');
  if(sel){
    while(sel.options.length>1) sel.remove(1);
    _cfgModels.forEach(m=>{const o=document.createElement('option');o.value=m.nom||'';o.textContent=m.nom||'';sel.appendChild(o);});
  }
}

function renderPwdList(){
  const c=document.getElementById('pwd-list');
  if(!c) return;
  c.innerHTML=Object.entries(_cfgPwds).map(([nm,pw])=>`
    <div class="pr">
      <div class="pn">${esc(nm)}</div>
      <input type="password" id="pwi-${esc(nm)}" value="${esc(String(pw))}" data-n="${esc(nm)}">
      <button class="btn-eye" onclick="toggleEye('pwi-${esc(nm)}')">👁</button>
      <button class="btn btn-sec" style="font-size:10px;padding:2px 6px" onclick="rmPilot('${esc(nm)}')">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:11px;padding:3px">Aucun pilote configuré</div>';
}

function toggleEye(id){const i=document.getElementById(id);if(i)i.type=i.type==='password'?'text':'password';}
function addPilot(){
  const n=document.getElementById('np-name').value.trim(),pw=document.getElementById('np-pw').value;
  if(!n){toast('Nom requis','err');return;}
  _cfgPwds[n]=pw;
  document.getElementById('np-name').value='';document.getElementById('np-pw').value='';
  renderPwdList();
}
function rmPilot(n){delete _cfgPwds[n];renderPwdList();}
async function savePwds(){
  document.querySelectorAll('#pwd-list input[data-n]').forEach(i=>_cfgPwds[i.dataset.n]=i.value);
  const r=await fetch('/api/pilot_passwords_excel',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,pilot_passwords:_cfgPwds})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('MDP enregistrés et sauvegardés dans Excel','ok');loadCfg();}else toast(d&&d.error||'Erreur','err');
}

const DAYS=[{k:'lun',l:'Lun'},{k:'mar',l:'Mar'},{k:'mer',l:'Mer'},{k:'jeu',l:'Jeu'},{k:'ven',l:'Ven'},{k:'sam',l:'Sam'},{k:'dim',l:'Dim'}];
function renderModelList(){
  const c=document.getElementById('models-list');
  if(!c) return;
  c.innerHTML=_cfgModels.map((m,mi)=>{
    const j=m.jours||{};
    return `<div class="model-card">
      <div class="mch">
        <input value="${esc(m.nom||'Poste '+(mi+1))}" onchange="_cfgModels[${mi}].nom=this.value" placeholder="Nom du poste">
        <button class="btn btn-danger" style="font-size:10px;padding:2px 6px" onclick="_cfgModels.splice(${mi},1);renderModelList()">✕</button>
      </div>
      <div class="day-grid">${DAYS.map(d=>{const dc=j[d.k]||{};
        return `<div class="day-box"><div class="day-lbl">${d.l}</div>
          <input type="time" onchange="setDay(${mi},'${d.k}','debut',this.value)" value="${dc.debut||'05:00'}" style="margin-bottom:2px">
          <input type="time" onchange="setDay(${mi},'${d.k}','fin',this.value)" value="${dc.fin||'13:00'}">
        </div>`;}).join('')}
      </div>
    </div>`;
  }).join('')||'<div style="color:var(--gray);font-size:11px">Aucun modèle horaire</div>';
}

function setDay(mi,day,field,val){if(!_cfgModels[mi])return;if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};if(!_cfgModels[mi].jours[day])_cfgModels[mi].jours[day]={};_cfgModels[mi].jours[day][field]=val;}
function addModel(){_cfgModels.push({nom:'Nouveau poste',jours:{lun:{debut:'05:00',fin:'13:00'},mar:{debut:'05:00',fin:'13:00'},mer:{debut:'05:00',fin:'13:00'},jeu:{debut:'05:00',fin:'13:00'},ven:{debut:'05:00',fin:'13:00'},sam:{debut:'',fin:''},dim:{debut:'',fin:''}}});renderModelList();}
async function saveModels(){
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,modeles_horaires:_cfgModels})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Modèles enregistrés','ok');loadCfg();}else toast(d&&d.error||'Erreur','err');
}
async function saveProdRef(){
  const v=parseFloat(document.getElementById('cfg-pr').value)||200;
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,prod_ref:v})});
  const d=r?await r.json():{};
  d&&d.ok?toast('Enregistré','ok'):toast(d&&d.error||'Erreur','err');
}

async function generateDashboard(){
  const st=document.getElementById('dash-status');
  if(st){st.textContent='Génération en cours…';st.style.color='var(--amber)';}
  let d={};
  try{
    const r=await fetch('/api/generate_dashboard',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
    if(r&&r.ok) d=await r.json();
  }catch(e){}
  if(d&&d.ok){
    if(st){st.textContent='✓ Dashboard généré : '+esc(d.path||'');st.style.color='var(--green)';}
    toast('Dashboard généré !','ok');
  } else {
    const errMsg=d&&d.error?d.error:'Erreur inconnue — vérifiez que le fichier Excel est configuré dans les paramètres';
    if(st){st.textContent='✗ '+errMsg;st.style.color='var(--red)';}
    toast('Erreur dashboard: '+errMsg,'err');
  }
}

// ── MODALS ──
function openM(id){const m=document.getElementById(id);if(m){m.classList.add('on');m.style.display='flex';}}
function closeM(id){const m=document.getElementById(id);if(m){m.classList.remove('on');m.style.display='';}}
document.addEventListener('click',e=>{if(e.target.classList.contains('overlay'))closeM(e.target.id);});

// ── UTILS ──
async function apiFetch(url,retries=2){
  try{
    const r=await fetch(url);
    if(r.status===429){
      if(retries>0){await new Promise(res=>setTimeout(res,1200));return apiFetch(url,retries-1);}
      return null;
    }
    return r.ok?await r.json():null;
  }catch(e){return null;}
}
function fmtDur2(s){if(!s||s<0)return'0:00';const m=Math.floor(s/60),sec=Math.floor(s%60);return m+':'+String(sec).padStart(2,'0');}
function fmtTRSv(v){return(v===null||v===undefined||isNaN(v)||v<0)?'--%':parseFloat(v).toFixed(1)+'%';}
function fmtDur(s){if(!s||s<0)return'00:00:00';const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sec=Math.floor(s%60);return[h,m,sec].map(x=>String(x).padStart(2,'0')).join(':');}
function fmtD2(s){if(!s||s<0)return'0 min';const h=Math.floor(s/3600),m=Math.floor((s%3600)/60);return h?h+'h'+String(m).padStart(2,'0'):m+' min';}
function fmtTRS(v){return(v===null||v===undefined||isNaN(v))?'--%':parseFloat(v).toFixed(1)+'%';}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function toast(msg,type){
  let t=document.getElementById('_toast');
  if(!t){t=document.createElement('div');t.id='_toast';t.style.cssText='position:fixed;bottom:16px;right:16px;padding:8px 14px;border-radius:7px;font-size:13px;font-weight:600;z-index:999;transition:opacity .3s;box-shadow:0 4px 12px rgba(0,0,0,.18)';document.body.appendChild(t);}
  t.textContent=msg;t.style.background=type==='ok'?'#16a34a':'#dc2626';t.style.color='#fff';t.style.opacity='1';
  clearTimeout(t._to);t._to=setTimeout(()=>t.style.opacity='0',3000);
}
</script>
</body>
</html>"""

def _session_autosave():
    while True:
        time.sleep(30)
        try: save_session()
        except: pass

def _dashboard_autogen():
    while True:
        time.sleep(30)
        try:
            if cfg.get("db_path") and _S.get("pilot"):
                generate_dashboard_html()
        except: pass

def main():
    global cfg
    cfg = load_cfg()
    load_session()
    threading.Thread(target=load_lists, daemon=True).start()
    threading.Thread(target=load_history, daemon=True).start()
    threading.Thread(target=_session_autosave, daemon=True).start()
    threading.Thread(target=_dashboard_autogen, daemon=True).start()
    _start_periodic_excel_sync()

    # Recover pending Excel write after crash
    try:
        if os.path.exists(PENDING_FILE):
            with open(PENDING_FILE, encoding="utf-8") as f:
                pending = json.load(f)
            if pending.get("db_path") and pending.get("prod_row"):
                write_excel_bg(pending["prod_row"], pending.get("evt_rows", []))
    except:
        pass

    try:
        import webview
        from threading import Thread

        def run_flask():
            flask_app.run(host="127.0.0.1", port=5001, debug=False, use_reloader=False)

        t = Thread(target=run_flask, daemon=True)
        t.start()
        import time; time.sleep(0.8)
        _w = webview.create_window(
            "KPI-ORC",
            "http://127.0.0.1:5001",
            min_size=(900, 600),
            resizable=True,
        )
        def _on_shown():
            try: _w.maximize()
            except: pass
        webview.start(_on_shown)
    except ImportError:
        # Fallback: run as plain Flask server (dev mode)
        print("pywebview non disponible — démarrage en mode serveur sur http://127.0.0.1:5001")
        flask_app.run(host="127.0.0.1", port=5001, debug=True, use_reloader=False)


if __name__ == "__main__":
    main()
