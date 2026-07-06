"""KPI-ORC v6.3 - Flask + pywebview"""
import json, os, sys, datetime, threading, math, shutil
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

# Nouveau schéma unifié - 37 colonnes (remplace Data + Evenements)
DECL_HEADERS = [
    "Type","OF","Date","Poste","Pilote","Co-Pilote","Nb Personnes",
    "Taille","Code Produit","Type Produit","Poids Garnissage","Fibre",
    "OF Taie","Traca Fibre","Ref Taie","Kit",
    "Heure Debut","Heure Fin","Duree",
    "Qte Fabriquee","Qte Emballee","Equivalence","Cadence/h","Cadence/h/pers","TRS%",
    "Qte Init Taie","Nb Taie 2nd Choix","Nb Defaut Couture",
    "Mq Taie","Mq Housse/Encart","Nb PP Cousue",
    "Changement de Serie","Manquant MP","Manquant Personnel/Reunion",
    "Nettoyage Fin de Poste","Commentaire","Prevu/Hors TRS",
]

POSTES = ["Matin","Midi","Nuit","Jour"]

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
cfg = {}

flask_app = Flask(__name__)

# ── Utilitaires ────────────────────────────────────────────────────────────────
def fmt(seconds):
    s = max(0, int(seconds or 0))
    return f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}"

def _hms_to_sec(s):
    try:
        parts = str(s).split(":")
        if len(parts)==3: return int(parts[0])*3600+int(parts[1])*60+float(parts[2])
        if len(parts)==2: return int(parts[0])*60+float(parts[1])
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

def get_shift_duration_s(poste, date_obj=None):
    """Durée nominale du poste en secondes selon le modèle horaire."""
    models = cfg.get("modeles_horaires", [])
    if not models: return 28800
    model = next((m for m in models if str(m.get("nom","")).strip().lower()==str(poste or "").strip().lower()), None)
    if not model: return 28800
    jours = model.get("jours", {})
    if date_obj and jours:
        day_map = {0:'lun',1:'mar',2:'mer',3:'jeu',4:'ven',5:'sam',6:'dim'}
        day_cfg = jours.get(day_map.get(date_obj.weekday(),'lun'))
    else:
        day_cfg = next((v for k,v in jours.items() if v and v.get("debut") and v.get("fin")), None) if jours else None
    if not day_cfg:
        debut_str = model.get("debut","05:00"); fin_str = model.get("fin","13:00")
    else:
        debut_str = day_cfg.get("debut","05:00"); fin_str = day_cfg.get("fin","13:00")
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
        wb.close()
    except: pass

def get_list(h):
    return _lists.get(h,[])

def load_history():
    global _decl_cache
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
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
    except: pass

# ── Calcul équivalence ────────────────────────────────────────────────────────
def calc_equiv(qte, taille, type_prod):
    types  = get_list("Type produit")
    equivs = get_list("Equivalence coef") or get_list("Equivalence")
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
        if ev.get("cat") not in ("ratt","pb","nettoyage"): continue
        if not ev.get("key") or ev["key"].startswith("_"): continue
        if of_start and ev["start"] < of_start and ev.get("key")!="arret_interposte": continue
        start = ev["start"]
        end = ev.get("end") or datetime.datetime.now()
        if ev["key"]=="nettoyage":
            ntype = ev.get("nettoyage_type","court")
            label = {"court":"Nettoyage court","long":"Nettoyage long","grand":"Grand nettoyage"}.get(ntype,"Nettoyage court")
        else:
            cat_name = "Rattrapage" if ev["cat"]=="ratt" else "PB Technique"
            lbl = next((e[0] for e in EVENTS if e[1]==ev["key"]),ev["key"])
            label = f"{cat_name}: {lbl}"
        rows.append(_base_row(label, start, end, ev.get("comment",""), "OUI" if ev.get("hors_trs") else ""))
    for ps, pe in pause_periods:
        if of_start and ps < of_start: continue
        rows.append(_base_row("Pause pilote", ps, pe))
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
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _ensure_decl_sheet(wb)
                ws.append(prod_row)
                _format_row(ws, ws.max_row)
                for er in evt_rows:
                    ws.append(er)
                    _format_row(ws, ws.max_row)
                _safe_excel_save(wb, path)
                try: os.remove(PENDING_FILE)
                except: pass
        except: pass
        threading.Thread(target=load_history, daemon=True).start()
    threading.Thread(target=_bg, daemon=True).start()

def write_changement_of(start_dt, end_dt):
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    pilot = _S.get("last_of_pilot") or _S.get("pilot") or ""
    dur_s = (end_dt-start_dt).total_seconds()
    row = [
        "Changement d'OF","",start_dt.strftime("%d/%m/%Y"),
        _S.get("poste",""),pilot,"","","","","","","","","","","",
        start_dt.strftime("%H:%M:%S"),end_dt.strftime("%H:%M:%S"),fmt(dur_s),
        "","","","","","","","","","","","","","","","","",
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
        except: pass
    threading.Thread(target=_bg,daemon=True).start()

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
        of_elapsed = (datetime.datetime.now()-_S["of_start"]).total_seconds() + _S["inter_of_s"]
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
    }

@flask_app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@flask_app.route('/api/state')
def api_state():
    return jsonify(_state_json())

@flask_app.route('/api/lists')
def api_lists():
    return jsonify({
        "pilotes": get_list("Pilotes") or get_list("pilotes"),
        "tailles": get_list("Tailles") or get_list("taille"),
        "types_prod": get_list("Type produit") or get_list("types_prod"),
        "fibres": get_list("Fibres") or get_list("fibre"),
        "tracas": get_list("Traca") or get_list("tracas"),
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
    _S["prod_active"] = True
    _S["of_start"] = now
    _S["inter_of_s"] = 0.0
    _S["pause_total_s"] = 0.0
    _S["pause_periods"] = []
    _S["tl_events"] = []
    _S["of_count_shift"] += 1
    _S["form"] = {"of_num": ""}  # OF vide au démarrage
    if not _S.get("shift_start"):
        _S["shift_start"] = now
    t_reset()
    save_session()
    return jsonify({"ok":True,"gap_s":round(gap_s,0)})

@flask_app.route('/api/inter_of_confirm', methods=['POST'])
def api_inter_of_confirm():
    data = request.json or {}
    _S["inter_of_s"] = float(data.get("inter_of_s",0))
    if _S["inter_of_s"] > 30 and _S["last_of_end"] and _S["of_start"]:
        write_changement_of(_S["last_of_end"], _S["of_start"])
    save_session()
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
    if prod_ref>0 and of_s>0:
        trs = round(equiv/(prod_ref*of_s/28800)*100,1)
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
        fmt(_S["inter_of_s"]) if _S["inter_of_s"]>0 else "",
        _min_str_to_hms(v.get("duree_mq_mp","")),
        _min_str_to_hms(v.get("manquant_pers","")),
        fmt(_nett_s),
        v.get("comment",""),
        "",
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
    if prod_ref>0 and of_s>0:
        trs=round(equiv/(prod_ref*of_s/28800)*100,1)
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
    return jsonify({"ok":True})

@flask_app.route('/api/end_stop', methods=['POST'])
def api_end_stop():
    data = request.json or {}
    key = data.get("key","")
    comment = data.get("comment","")
    t_stop(key)
    tl_close(key,comment)
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
    prod_rows = [(rn,r) for rn,r in _decl_cache if str(r[0] or "").strip().lower() in ("production","prod","")]
    for rn, r in prod_rows[-100:]:
        try:
            trs = -1
            try:
                equiv_v = float(str(r[21] or 0).replace(",","."))
                of_s_v = _hms_to_sec(str(r[18] or "00:00:00"))
                pr = get_prod_ref()
                trs_col = str(r[24] or "")
                if trs_col:
                    try: trs = round(float(trs_col.replace(",",".")),1)
                    except: pass
                elif pr>0 and of_s_v>0 and equiv_v>0:
                    trs = round(equiv_v/(pr*of_s_v/28800)*100,1)
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
            rows.append({
                "row_num": rn,
                "type": str(r[0] or ""),
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
                _safe_excel_save(wb,path)
            threading.Thread(target=load_history,daemon=True).start()
        except: pass
    threading.Thread(target=_bg,daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/fin_poste_data')
def api_fin_poste_data():
    today = datetime.date.today().strftime("%d/%m/%Y")
    pilot = _S["pilot"] or ""
    pilot_poste = _S["poste"] or ""
    prod_ref = get_prod_ref()
    tot_eq=0.0; tot_s=0.0
    of_list=[]
    prod_rows = [(rn,r) for rn,r in _decl_cache if str(r[0] or "").strip().lower() in ("production","prod","")]
    for rn, r in prod_rows:
        try:
            if _row_date(r[2])!=today: continue
            if str(r[4] or "")!=pilot: continue
            eq=float(str(r[21] or 0).replace(",","."))
            s=_hms_to_sec(str(r[18] or "00:00:00"))
            tot_eq+=eq; tot_s+=s
            trs=-1
            trs_col = str(r[24] or "")
            if trs_col:
                try: trs=round(float(trs_col.replace(",",".")),1)
                except: pass
            elif prod_ref>0 and s>0 and eq>0:
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
    shift_s = get_shift_duration_s(pilot_poste, datetime.date.today())
    trs_poste_shift = -1.0
    if prod_ref > 0 and shift_s > 0:
        trs_poste_shift = round(tot_eq/(prod_ref*shift_s/28800)*100,1)
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
    pilot = _S["pilot"] or ""
    poste = _S["poste"] or ""
    prod_ref = get_prod_ref()
    shift_s = get_shift_duration_s(poste, datetime.date.today())
    rows = []
    tot_eq=0.0; tot_s=0.0
    for rn,r in _decl_cache:
        if str(r[0] or "").strip().lower() not in ("production","prod",""): continue
        if _row_date(r[2]) != today: continue
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

    def trs_color_hex(t):
        if t<0: return "#94a3b8"
        if t>=70: return "#1a8c4e"
        if t>=50: return "#d97706"
        return "#e31e24"

    def trs_bg(t):
        if t<0: return "#f1f5f9"
        if t>=70: return "#d1fae5"
        if t>=50: return "#fef3c7"
        return "#fee2e2"

    def calc_trs_of(rows):
        eq=sum(float(str(r[21] or 0).replace(",",".") or 0) for r in rows)
        s=sum(_hms_to_sec(str(r[18] or "00:00:00")) for r in rows)
        if prod_ref>0 and s>0: return round(eq/(prod_ref*s/28800)*100,1)
        return -1

    def calc_trs_shift(rows, poste, date_str):
        try: d=datetime.datetime.strptime(date_str,"%d/%m/%Y").date()
        except: d=None
        shift_s=get_shift_duration_s(poste,d)
        eq=sum(float(str(r[21] or 0).replace(",",".") or 0) for r in rows)
        if prod_ref>0 and shift_s>0: return round(eq/(prod_ref*shift_s/28800)*100,1)
        return -1

    # Grouper par (date, poste, pilote)
    sessions_map = defaultdict(list)
    for r in prod_rows_all:
        key=(str(_row_date(r[2])),str(r[3] or ""),str(r[4] or ""))
        sessions_map[key].append(r)

    # Trier par date desc
    sess_sorted = sorted(sessions_map.items(), key=lambda x:x[0][0], reverse=True)

    # Session en cours (live)
    today_str  = datetime.date.today().strftime("%d/%m/%Y")
    pilot_now  = _S.get("pilot","") or ""
    poste_now  = _S.get("poste","") or ""
    prod_active= bool(_S.get("prod_active"))
    active_stops=[k for k,t in _S.get("timers",{}).items() if t.get("running") and not k.startswith("_")]
    of_num_now = (_S.get("form") or {}).get("of_num","") or "—"
    of_start_str=""
    if _S.get("of_start"):
        try: of_start_str=_S["of_start"].strftime("%H:%M:%S")
        except: pass

    # Sessions du jour (live)
    today_rows = sessions_map.get((today_str,poste_now,pilot_now),[])
    today_trs_shift = calc_trs_shift(today_rows,poste_now,today_str) if today_rows else -1
    today_trs_of    = calc_trs_of(today_rows) if today_rows else -1
    today_eq = sum(float(str(r[21] or 0).replace(",",".") or 0) for r in today_rows)

    # 3 derniers postes terminés
    last3 = [(k,v) for k,v in sess_sorted if not (k[0]==today_str and k[1]==poste_now and k[2]==pilot_now)][:3]

    # Pareto arrêts (tous)
    evt_dur = defaultdict(float)
    for r in evt_rows_all:
        t=str(r[0] or "")
        if not t or t.lower() in ("pause pilote","changement d'of"): continue
        evt_dur[t] += _hms_to_sec(str(r[18] or "00:00:00"))
    pareto = sorted(evt_dur.items(),key=lambda x:-x[1])[:12]

    # Événements du poste en cours
    today_evts = [r for r in evt_rows_all if _row_date(r[2])==today_str and str(r[4] or "")==pilot_now]

    # Arrêts actifs live
    stops_live_html=""
    for k in active_stops:
        ev=next((e for e in EVENTS if e[1]==k),None)
        lbl=ev[0] if ev else k
        t=_S.get("timers",{}).get(k,{})
        el=t.get("elapsed",0)
        stops_live_html+=f'<div class="slv-card"><div class="slv-name">{lbl}</div><div class="slv-timer">{fmt(el)}</div></div>'

    gen_time=datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    sup_col="#e31e24" if prod_active else "#1a8c4e"
    sup_txt="⬤ PRODUCTION EN COURS" if prod_active else "○ Aucune production active"

    # Gauge SVG helper
    def gauge_svg(trs, size=120):
        pct=max(0,min(100,trs)) if trs>=0 else 0
        col=trs_color_hex(trs)
        r_out=50; r_in=33; cx=60; cy=65
        import math
        def arc_pt(r,deg):
            rad=math.radians(deg)
            return cx+r*math.cos(rad), cy+r*math.sin(rad)
        x1,y1=arc_pt(r_out,180); x2,y2=arc_pt(r_out,0)
        xi1,yi1=arc_pt(r_in,180); xi2,yi2=arc_pt(r_in,0)
        bg=f'<path d="M{x1},{y1} A{r_out},{r_out} 0 0,1 {x2},{y2} L{xi2},{yi2} A{r_in},{r_in} 0 0,0 {xi1},{yi1} Z" fill="#e2e8f0"/>'
        if pct>0:
            end_deg=180-pct*1.8
            fx1,fy1=arc_pt(r_out,180); fx2,fy2=arc_pt(r_out,end_deg)
            fxi1,fyi1=arc_pt(r_in,180); fxi2,fyi2=arc_pt(r_in,end_deg)
            lg=1 if pct>50 else 0
            fg=f'<path d="M{fx1},{fy1} A{r_out},{r_out} 0 {lg},1 {fx2},{fy2} L{fxi2},{fyi2} A{r_in},{r_in} 0 {lg},0 {fxi1},{fyi1} Z" fill="{col}"/>'
        else: fg=""
        lbl=f"{trs:.1f}%" if trs>=0 else "—"
        txt=f'<text x="{cx}" y="{cy+10}" text-anchor="middle" font-size="18" font-weight="900" fill="{col}">{lbl}</text>'
        return f'<svg width="{size}" height="{size*65//120}" viewBox="0 0 120 65">{bg}{fg}{txt}</svg>'

    # Build 3 last sessions cards
    last3_html=""
    for (date,poste,pilot),rows in last3:
        t_shift=calc_trs_shift(rows,poste,date)
        t_of=calc_trs_of(rows)
        eq=sum(float(str(r[21] or 0).replace(",",".") or 0) for r in rows)
        nb_of=len(rows)
        last3_html+=f"""
        <div class="sess-card">
          <div class="sess-hdr" style="background:{trs_bg(t_shift)};border-left:4px solid {trs_color_hex(t_shift)}">
            <div class="sess-title">{poste} — {date}</div>
            <div class="sess-pilot">👤 {pilot}</div>
          </div>
          {gauge_svg(t_shift)}
          <div class="sess-stats">
            <div><span class="sl">OF</span><span class="sv">{nb_of}</span></div>
            <div><span class="sl">Equiv</span><span class="sv">{round(eq,1)}</span></div>
            <div><span class="sl">TRS(OF)</span><span class="sv" style="color:{trs_color_hex(t_of)}">{f"{t_of:.1f}%" if t_of>=0 else "—"}</span></div>
          </div>
        </div>"""

    # Productions table du poste en cours
    prod_table_html=""
    for r in list(reversed(today_rows))[:20]:
        trs_v=str(r[24] or "")
        try: tv=float(trs_v.replace(",","."))
        except: tv=-1
        tc=trs_color_hex(tv)
        prod_table_html+=f"""<tr>
          <td><strong>{r[1] or ''}</strong></td>
          <td>{str(r[16] or '')[:5]}</td><td>{str(r[17] or '')[:5]}</td>
          <td>{r[18] or ''}</td><td>{r[7] or ''}</td>
          <td>{r[19] or ''}</td><td>{r[20] or ''}</td><td>{r[21] or ''}</td>
          <td style="color:{tc};font-weight:800">{trs_v+'%' if trs_v else '—'}</td>
        </tr>"""

    # Arrêts table du poste en cours
    evts_table_html=""
    for r in list(reversed(today_evts))[:20]:
        t=str(r[0] or "")
        cls="badge-red" if "PB" in t or "Technique" in t else "badge-amber" if "Ratt" in t else "badge-cyan" if "Nett" in t else "badge-navy"
        evts_table_html+=f"""<tr>
          <td><span class="badge {cls}">{t}</span></td>
          <td>{str(r[16] or '')[:8]}</td><td>{str(r[17] or '')[:8]}</td>
          <td>{r[18] or ''}</td><td style="text-align:left;max-width:150px;overflow:hidden">{r[35] or ''}</td>
        </tr>"""

    # Pareto bars (inline SVG)
    pareto_html=""
    if pareto:
        max_dur=pareto[0][1]/60
        for lbl,dur in pareto[:8]:
            pct=dur/60/max_dur*100
            col="#e31e24" if "PB" in lbl or "Technique" in lbl else "#d97706" if "Ratt" in lbl else "#0891b2" if "Nett" in lbl else "#7c3aed"
            pareto_html+=f"""<div class="pareto-row">
              <div class="pareto-lbl">{lbl[:28]}</div>
              <div class="pareto-bar-wrap"><div class="pareto-bar" style="width:{pct:.1f}%;background:{col}"></div></div>
              <div class="pareto-val">{dur/60:.0f}min</div>
            </div>"""

    # Build timeline SVG for today
    def build_tl_svg(events, width=700):
        import math
        now_ts=datetime.datetime.now()
        win_end=now_ts
        win_start=now_ts-datetime.timedelta(hours=8)
        def to_x(dt_str):
            try:
                dt=datetime.datetime.strptime(f"{today_str} {str(dt_str)[:8]}","%d/%m/%Y %H:%M:%S")
                frac=(dt-win_start).total_seconds()/(8*3600)
                return max(0,min(width,int(frac*width)))
            except: return 0
        catcol={"pb":"#e31e24","ratt":"#d97706","nettoyage":"#0891b2","pause":"#7c3aed"}
        svg=f'<svg width="{width}" height="52" viewBox="0 0 {width} 52" style="display:block">'
        svg+=f'<rect x="0" y="16" width="{width}" height="20" fill="#e2e8f0" rx="4"/>'
        for r in events:
            x1=to_x(str(r[16] or ""))
            x2=to_x(str(r[17] or "")) if r[17] else int((datetime.datetime.now()-win_start).total_seconds()/(8*3600)*width)
            x2=max(x2,x1+2)
            t=str(r[0] or "")
            col="#e31e24" if "PB" in t or "Technique" in t else "#d97706" if "Ratt" in t else "#0891b2" if "Nett" in t else "#7c3aed"
            svg+=f'<rect x="{x1}" y="16" width="{x2-x1}" height="20" fill="{col}" rx="2" opacity="0.85"/>'
        for r in today_rows:
            x1=to_x(str(r[16] or ""))
            x2=to_x(str(r[17] or "")) if r[17] else int((datetime.datetime.now()-win_start).total_seconds()/(8*3600)*width)
            svg+=f'<rect x="{x1}" y="16" width="{max(2,x2-x1)}" height="20" fill="#1a8c4e" rx="2" opacity="0.4"/>'
        # Hour marks
        for h in range(9):
            dt=win_start+datetime.timedelta(hours=h)
            x=int(h/8*width)
            svg+=f'<line x1="{x}" y1="12" x2="{x}" y2="36" stroke="#94a3b8" stroke-width="0.5"/>'
            svg+=f'<text x="{x}" y="10" font-size="9" fill="#64748b" text-anchor="middle">{dt.strftime("%H:%M")}</text>'
        # Now marker
        now_x=int((datetime.datetime.now()-win_start).total_seconds()/(8*3600)*width)
        svg+=f'<line x1="{now_x}" y1="8" x2="{now_x}" y2="44" stroke="#1a1f5e" stroke-width="2"/>'
        svg+='<defs><svg><rect x="0" y="44" width="60" height="8" fill="#1a1f5e" rx="2"/></svg></defs>'
        # Legend
        for i,(lbl,col) in enumerate([("Prod","#1a8c4e"),("PB Technique","#e31e24"),("Rattrapage","#d97706"),("Nettoyage","#0891b2"),("Pause","#7c3aed")]):
            svg+=f'<rect x="{i*140}" y="44" width="10" height="8" fill="{col}" rx="2"/>'
            svg+=f'<text x="{i*140+14}" y="52" font-size="8" fill="#475569">{lbl}</text>'
        svg+='</svg>'
        return svg

    tl_svg=build_tl_svg(today_evts)

    shift_duration_s=get_shift_duration_s(poste_now,datetime.date.today())
    shift_h=shift_duration_s/3600

    html=f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="30">
<title>KPI Supervision — ORC1</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:#f0f4fb;color:#0f172a;font-size:12px}}
.page-hdr{{background:#1a1f5e;color:#fff;padding:10px 20px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}}
.page-hdr h1{{font-size:15px;font-weight:800;letter-spacing:.3px}}
.page-hdr .gen-info{{font-size:11px;opacity:.75}}
{'div.alert-strip{background:#e31e24;color:#fff;padding:8px 20px;font-weight:800;font-size:13px;text-align:center;animation:blink 1s infinite}' if active_stops else ''}
@keyframes blink{{0%,100%{{opacity:1}}50%{{opacity:.7}}}}
.grid-main{{display:grid;grid-template-columns:320px 1fr;gap:12px;padding:12px 16px;min-height:0}}
.col-left{{display:flex;flex-direction:column;gap:10px}}
.col-right{{display:flex;flex-direction:column;gap:10px}}
.block{{background:#fff;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.07);overflow:hidden}}
.block-hdr{{background:#1a1f5e;color:#fff;padding:10px 14px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;display:flex;align-items:center;justify-content:space-between}}
.block-hdr.green{{background:#1a8c4e}}
.block-hdr.red{{background:#e31e24}}
.block-hdr.amber{{background:#d97706}}
.block-body{{padding:12px 14px}}
/* Supervision live */
.sup-hdr{{background:{sup_col};color:#fff;padding:14px 16px;display:flex;align-items:center;justify-content:space-between}}
.sup-status{{font-size:16px;font-weight:900}}
.sup-meta{{font-size:11px;opacity:.85}}
.sup-kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;padding:12px 14px}}
.sup-kpi{{background:#f8fafc;border-radius:8px;padding:8px 12px;text-align:center}}
.sup-kpi-lbl{{font-size:9px;font-weight:700;text-transform:uppercase;color:#64748b}}
.sup-kpi-val{{font-size:20px;font-weight:900;margin-top:2px}}
.stops-live-row{{display:flex;gap:8px;flex-wrap:wrap;padding:0 14px 12px}}
.slv-card{{background:#fee2e2;border:2px solid #e31e24;border-radius:8px;padding:8px 14px;text-align:center}}
.slv-name{{font-size:11px;font-weight:700;color:#991b1b}}
.slv-timer{{font-size:18px;font-weight:900;color:#e31e24;font-variant-numeric:tabular-nums}}
/* Sessions cards */
.sessions-row{{display:flex;gap:8px;padding:12px 14px;flex-wrap:wrap}}
.sess-card{{background:#f8fafc;border-radius:10px;padding:10px 12px;flex:1;min-width:180px;text-align:center}}
.sess-hdr{{border-radius:8px 8px 0 0;padding:8px 10px;margin:-10px -12px 8px}}
.sess-title{{font-size:11px;font-weight:800;color:#1a1f5e}}
.sess-pilot{{font-size:10px;color:#64748b;margin-top:2px}}
.sess-stats{{display:flex;justify-content:space-around;margin-top:8px;font-size:10px}}
.sl{{color:#64748b;display:block}}
.sv{{font-weight:800;font-size:13px;display:block}}
/* Tables */
table{{width:100%;border-collapse:collapse;font-size:11px}}
th{{background:#dde4ef;padding:7px 8px;font-weight:700;text-align:center;border-bottom:2px solid #c0cde0;white-space:nowrap}}
td{{padding:5px 8px;text-align:center;border-bottom:1px solid #edf0f7;white-space:nowrap}}
tr:hover td{{background:#f8fafc}}
.badge{{display:inline-block;padding:1px 6px;border-radius:12px;font-size:10px;font-weight:700}}
.badge-red{{background:#fee2e2;color:#991b1b}}
.badge-amber{{background:#fef3c7;color:#92400e}}
.badge-cyan{{background:#cffafe;color:#0e7490}}
.badge-navy{{background:#dbeafe;color:#1e40af}}
/* Pareto */
.pareto-row{{display:flex;align-items:center;gap:8px;margin-bottom:6px}}
.pareto-lbl{{width:160px;font-size:10px;text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0}}
.pareto-bar-wrap{{flex:1;background:#f0f4fb;border-radius:4px;height:16px;overflow:hidden}}
.pareto-bar{{height:16px;border-radius:4px;transition:width .3s}}
.pareto-val{{width:42px;font-size:10px;font-weight:700;text-align:right;flex-shrink:0}}
/* Timeline */
.tl-wrap{{padding:12px 14px;overflow-x:auto}}
/* Footer */
.footer{{padding:10px 20px;color:#94a3b8;font-size:10px;text-align:center;border-top:1px solid #e2e8f0}}
</style>
</head>
<body>
<div class="page-hdr">
  <h1>&#127981; KPI Supervision ORC1 — Dashboard encadrant</h1>
  <div class="gen-info">Généré le {gen_time} &nbsp;|&nbsp; Actualisation auto. 30s</div>
</div>
{'<div class="alert-strip">&#9888; ' + str(len(active_stops)) + ' ARRÊT(S) EN COURS — Poste ' + poste_now + ' — ' + pilot_now + '</div>' if active_stops else ''}

<div class="grid-main">
  <!-- Colonne gauche -->
  <div class="col-left">

    <!-- Supervision live -->
    <div class="block">
      <div class="sup-hdr">
        <div>
          <div class="sup-status">{sup_txt}</div>
          {'<div class="sup-meta">Pilote : ' + pilot_now + ' | Poste : ' + poste_now + ' | OF : ' + of_num_now + ' | Début : ' + of_start_str + '</div>' if prod_active else '<div class="sup-meta">En attente de déclaration</div>'}
        </div>
      </div>
      <div class="sup-kpis">
        <div class="sup-kpi">
          <div class="sup-kpi-lbl">TRS Poste</div>
          <div class="sup-kpi-val" style="color:{trs_color_hex(today_trs_shift)}">{f"{today_trs_shift:.1f}%" if today_trs_shift>=0 else "—"}</div>
        </div>
        <div class="sup-kpi">
          <div class="sup-kpi-lbl">Equiv. totale</div>
          <div class="sup-kpi-val">{round(today_eq,1)}</div>
        </div>
        <div class="sup-kpi">
          <div class="sup-kpi-lbl">Nb OF</div>
          <div class="sup-kpi-val">{len(today_rows)}</div>
        </div>
        <div class="sup-kpi">
          <div class="sup-kpi-lbl">Durée poste</div>
          <div class="sup-kpi-val">{shift_h:.1f}h</div>
        </div>
      </div>
      {('<div class="stops-live-row">' + stops_live_html + '</div>') if active_stops else '<div style="padding:0 14px 12px;color:#1a8c4e;font-weight:700;font-size:12px">✔ Aucun arrêt actif</div>'}
    </div>

    <!-- TRS gauge poste actuel -->
    <div class="block">
      <div class="block-hdr">TRS Poste en cours (vs modèle horaire)</div>
      <div class="block-body" style="text-align:center">
        {gauge_svg(today_trs_shift, 160)}
        <div style="font-size:11px;color:#64748b;margin-top:6px">Poste {poste_now} — {today_str}</div>
        <div style="font-size:11px;color:#64748b">Durée poste : {shift_h:.1f}h | Équiv : {round(today_eq,1)}</div>
      </div>
    </div>

    <!-- 3 derniers postes -->
    <div class="block">
      <div class="block-hdr">3 derniers postes</div>
      <div class="sessions-row">
        {last3_html if last3_html else '<div style="color:#94a3b8;padding:10px">Aucun historique</div>'}
      </div>
    </div>

    <!-- Pareto arrêts -->
    <div class="block">
      <div class="block-hdr">Pareto des arrêts (toutes sessions)</div>
      <div class="block-body">
        {pareto_html if pareto_html else '<div style="color:#94a3b8">Aucun arrêt enregistré</div>'}
      </div>
    </div>

  </div>

  <!-- Colonne droite -->
  <div class="col-right">

    <!-- Timeline 8h du poste -->
    <div class="block">
      <div class="block-hdr">Timeline du poste (8 dernières heures)</div>
      <div class="tl-wrap">{tl_svg}</div>
    </div>

    <!-- Productions du poste en cours -->
    <div class="block">
      <div class="block-hdr green">Productions du poste en cours</div>
      <div style="overflow-x:auto">
        <table>
          <thead><tr><th>OF</th><th>Début</th><th>Fin</th><th>Durée</th><th>Taille</th><th>Qté Fab</th><th>Qté Emb</th><th>Equiv</th><th>TRS%</th></tr></thead>
          <tbody>{prod_table_html if prod_table_html else '<tr><td colspan="9" style="color:#94a3b8;padding:14px">Aucune production déclarée</td></tr>'}</tbody>
        </table>
      </div>
    </div>

    <!-- Arrêts du poste en cours -->
    <div class="block">
      <div class="block-hdr red">Arrêts / événements du poste</div>
      <div style="overflow-x:auto">
        <table>
          <thead><tr><th>Type</th><th>Début</th><th>Fin</th><th>Durée</th><th>Commentaire</th></tr></thead>
          <tbody>{evts_table_html if evts_table_html else '<tr><td colspan="5" style="color:#94a3b8;padding:14px">Aucun arrêt</td></tr>'}</tbody>
        </table>
      </div>
    </div>

  </div>
</div>

<div class="footer">KPI-ORC v6.4 — Généré le {gen_time} — Actualisation toutes les 30 secondes</div>
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
  --c-bg:#f4f6fb;--c-card:#fff;--c-border:#e2e8f0;--c-text:#1e293b;--c-muted:#64748b;
  --c-primary:#1a1f5e;--c-primary-light:#e8eaf6;--c-accent:#3b82f6;
  --c-green:#16a34a;--c-red:#dc2626;--c-yellow:#d97706;--c-orange:#ea580c;
  --radius:10px;--shadow:0 2px 12px rgba(0,0,0,.08);--header-h:54px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--c-bg);color:var(--c-text);font-size:13px;line-height:1.4}
body.stop-active{--c-bg:#1a0808;--c-card:#2d1010;--c-border:#5a2020;--c-text:#f8d7d7;--c-muted:#c98888;--c-primary-light:#3d1010}
body.stop-active #app-header{background:#7f0000!important}

#app-header{position:fixed;top:0;left:0;right:0;height:var(--header-h);background:var(--c-primary);display:flex;align-items:center;padding:0 12px;gap:6px;z-index:100;border-bottom:2px solid #2d3480}
#app-header .logo{color:#fff;font-weight:800;font-size:16px;letter-spacing:1px;margin-right:10px;white-space:nowrap}
#app-header nav{display:flex;gap:2px;flex:1}
.nav-tab{background:none;border:none;color:rgba(255,255,255,.7);padding:6px 12px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:500;transition:all .15s;white-space:nowrap}
.nav-tab:hover{background:rgba(255,255,255,.12);color:#fff}
.nav-tab.active{background:rgba(255,255,255,.2);color:#fff;font-weight:700}
.tab-prod{background:#16a34a!important;color:#fff!important;animation:pt 2s infinite}
@keyframes pt{0%,100%{opacity:1}50%{opacity:.75}}
#hdr-info{color:rgba(255,255,255,.8);font-size:11px;text-align:right;line-height:1.3;margin-left:auto}
.alert-strip{background:#b91c1c;color:#fff;text-align:center;padding:5px;font-weight:700;font-size:12px;position:fixed;top:var(--header-h);left:0;right:0;z-index:99;display:none;animation:blink .9s step-start infinite}
.alert-strip.on{display:block}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.4}}

.view{display:none;padding-top:calc(var(--header-h) + 10px);min-height:100vh}
.view.on{display:block}
.ctr{max-width:1400px;margin:0 auto;padding:10px 14px}

/* LOGIN */
#v-login{display:flex;align-items:center;justify-content:center;min-height:100vh;background:linear-gradient(135deg,#1a1f5e,#2d3480,#1e3a8a);padding:20px}
.login-card{background:#fff;border-radius:16px;padding:36px;width:100%;max-width:400px;box-shadow:0 20px 60px rgba(0,0,0,.35)}
.login-card h1{color:var(--c-primary);font-size:26px;font-weight:800;margin-bottom:2px;text-align:center}
.login-card .sub{color:#64748b;text-align:center;margin-bottom:24px;font-size:12px}
.lf{margin-bottom:14px}
.lf label{display:block;font-weight:600;margin-bottom:4px;color:#374151;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
.lf select,.lf input{width:100%;padding:10px 12px;border:2px solid #e5e7eb;border-radius:7px;font-size:14px;outline:none;transition:border .2s;background:#fff;color:#1e293b}
.lf select:focus,.lf input:focus{border-color:var(--c-primary)}
.btn-login{width:100%;padding:12px;background:var(--c-primary);color:#fff;border:none;border-radius:7px;font-size:15px;font-weight:700;cursor:pointer;margin-top:6px;transition:opacity .2s}
.btn-login:hover{opacity:.88}
.login-err{color:#dc2626;text-align:center;margin-top:8px;font-size:12px;min-height:18px}

/* MAIN */
.main-row{display:grid;grid-template-columns:3fr 1fr;gap:14px;margin-bottom:16px;height:170px}
.btn-start{background:linear-gradient(135deg,#16a34a,#15803d);color:#fff;border:none;border-radius:var(--radius);font-size:26px;font-weight:800;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:10px;transition:all .15s;box-shadow:0 8px 28px rgba(22,163,74,.3)}
.btn-start:hover{transform:translateY(-2px);box-shadow:0 12px 36px rgba(22,163,74,.4)}
.btn-start:disabled{opacity:.4;cursor:not-allowed;transform:none}
.btn-fp{background:linear-gradient(135deg,#1a1f5e,#2d3480);color:#fff;border:none;border-radius:var(--radius);font-size:14px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:4px;transition:all .15s;box-shadow:0 4px 14px rgba(26,31,94,.25)}
.btn-fp:hover{opacity:.88;transform:translateY(-1px)}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;margin-bottom:14px}
.card{background:var(--c-card);border-radius:var(--radius);padding:14px;box-shadow:var(--shadow);border:1px solid var(--c-border)}
.card h3{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--c-muted);margin-bottom:8px;font-weight:700}
.prod-active-card{background:var(--c-card);border-radius:var(--radius);padding:12px;box-shadow:var(--shadow);border:2px solid var(--c-green);margin-bottom:12px}

/* PROD VIEW */
.ph-row{display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap}
.of-badge{background:var(--c-primary);color:#fff;padding:5px 14px;border-radius:18px;font-size:17px;font-weight:800;letter-spacing:1px}
.of-timer{font-size:36px;font-weight:800;color:var(--c-green);font-variant-numeric:tabular-nums}
.stop-timer-big{font-size:24px;font-weight:800;color:#dc2626;font-variant-numeric:tabular-nums}
.stop-lbl{background:#fee2e2;color:#b91c1c;border:1px solid #fca5a5;padding:3px 10px;border-radius:14px;font-weight:700;font-size:11px;animation:blink .8s step-start infinite}

/* 3-col form */
.form-3col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:8px}
.fzone{border-radius:8px;padding:10px 10px}
.fzone h4{font-size:9px;text-transform:uppercase;letter-spacing:.8px;font-weight:700;margin-bottom:7px;padding-bottom:3px;border-bottom:1px solid rgba(0,0,0,.1)}
.zi{background:#eef2ff;border:1px solid #c7d2fe}
.zi h4{color:#3730a3}
.zp{background:#f0fdf4;border:1px solid #bbf7d0}
.zp h4{color:#166534}
.zq{background:#fff7ed;border:1px solid #fed7aa}
.zq h4{color:#9a3412}
body.stop-active .zi{background:#1e1b3a;border-color:#4c4a8a}
body.stop-active .zp{background:#0a1f0a;border-color:#1a4d1a}
body.stop-active .zq{background:#1f1208;border-color:#4d2a00}
.fr{display:flex;flex-direction:column;margin-bottom:5px}
.fr label{font-size:9px;font-weight:700;color:var(--c-muted);margin-bottom:2px;text-transform:uppercase;letter-spacing:.3px}
.fr input,.fr select,.fr textarea{padding:5px 7px;border:1px solid var(--c-border);border-radius:5px;font-size:12px;background:var(--c-card);color:var(--c-text);width:100%;outline:none;transition:border .15s}
.fr input:focus,.fr select:focus,.fr textarea:focus{border-color:var(--c-accent)}
.fr textarea{resize:none;height:46px}
.fr.big input{font-size:17px;font-weight:700;padding:6px 7px;color:var(--c-green)}
.fr.ro input{background:#f8fafc;color:var(--c-muted)}

/* Controls */
.ctrls{display:flex;gap:7px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
.btn{padding:8px 16px;border:none;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;gap:5px}
.btn:hover{filter:brightness(.92)}
.btn-pause{background:#7c3aed;color:#fff}
.btn-arret{background:#dc2626;color:#fff}
.btn-endstop{background:#16a34a;color:#fff}
.btn-endprod{background:linear-gradient(135deg,#d97706,#b45309);color:#fff;font-size:14px;font-weight:800;padding:9px 20px;box-shadow:0 4px 14px rgba(217,119,6,.3)}
.btn-sec{background:var(--c-border);color:var(--c-text)}
.btn-danger{background:#dc2626;color:#fff}
.btn-prim{background:var(--c-primary);color:#fff}
.btn-ok{background:#16a34a;color:#fff}

/* Stop bar */
.stop-bar{background:#7f0000;color:#fff;border-radius:8px;padding:9px 14px;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;gap:10px}
.stop-bar .stype{font-weight:700;font-size:13px}
.stop-bar .selap{font-size:20px;font-weight:800;font-variant-numeric:tabular-nums}

/* Timeline */
.tl-wrap{background:var(--c-card);border-radius:var(--radius);padding:9px 10px;margin-bottom:8px;box-shadow:var(--shadow);border:1px solid var(--c-border)}
.tl-wrap h4{font-size:9px;text-transform:uppercase;letter-spacing:.8px;color:var(--c-muted);margin-bottom:5px;font-weight:700}

/* Recap stops */
.recap{background:var(--c-card);border-radius:var(--radius);padding:9px 10px;box-shadow:var(--shadow);border:1px solid var(--c-border);margin-bottom:8px}
.recap h4{font-size:9px;text-transform:uppercase;letter-spacing:.8px;color:var(--c-muted);margin-bottom:5px;font-weight:700}
.si{display:flex;align-items:center;gap:7px;padding:4px 0;border-bottom:1px solid var(--c-border);font-size:12px}
.si:last-child{border:none}
.sdot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.si .sname{flex:1;font-weight:600}
.si .sdur{color:var(--c-muted);min-width:50px;text-align:right;font-size:11px}
.btn-edit{background:none;border:none;cursor:pointer;padding:2px 4px;border-radius:3px;color:var(--c-accent);font-size:13px}
.btn-edit:hover{background:var(--c-primary-light)}

/* Modals */
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:200;align-items:center;justify-content:center}
.overlay.on{display:flex}
.mbox{background:var(--c-card);border-radius:12px;padding:22px;width:90%;max-width:480px;box-shadow:0 20px 60px rgba(0,0,0,.3);max-height:90vh;overflow-y:auto}
.mbox.wide{max-width:880px}
.mbox h2{font-size:16px;font-weight:700;margin-bottom:14px;color:var(--c-primary)}
.m-acts{display:flex;gap:8px;justify-content:flex-end;margin-top:14px;flex-wrap:wrap}
.stop-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-bottom:14px}
.stop-btn{padding:11px;border:2px solid var(--c-border);border-radius:7px;background:var(--c-card);cursor:pointer;font-size:12px;font-weight:600;text-align:center;transition:all .15s;color:var(--c-text)}
.stop-btn:hover{border-color:var(--c-red);background:#fee2e2;color:#b91c1c}

/* End-prod modal */
.ep-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px}
.ep-stat{text-align:center;padding:10px;background:var(--c-bg);border-radius:7px}
.ep-stat .val{font-size:26px;font-weight:800;color:var(--c-primary)}
.ep-stat .lbl{font-size:9px;text-transform:uppercase;color:var(--c-muted);font-weight:700}
.ep-tbl{width:100%;border-collapse:collapse;font-size:11px}
.ep-tbl th{text-align:left;padding:4px 7px;background:var(--c-bg);font-size:9px;text-transform:uppercase;color:var(--c-muted)}
.ep-tbl td{padding:4px 7px;border-bottom:1px solid var(--c-border)}

/* Settings */
.ss{background:var(--c-card);border-radius:var(--radius);padding:14px;margin-bottom:14px;box-shadow:var(--shadow);border:1px solid var(--c-border)}
.ss h3{font-size:12px;font-weight:700;margin-bottom:10px;color:var(--c-primary)}
.pr{display:flex;align-items:center;gap:7px;margin-bottom:7px;padding:5px;border-radius:5px;background:var(--c-bg)}
.pr .pn{font-weight:600;min-width:110px;font-size:11px}
.pr input{flex:1;padding:5px 7px;border:1px solid var(--c-border);border-radius:4px;font-size:12px}
.btn-eye{background:none;border:none;cursor:pointer;padding:3px;color:var(--c-muted);font-size:13px}
.day-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:5px;margin-bottom:8px}
.day-box{background:var(--c-bg);border-radius:5px;padding:5px;text-align:center}
.day-box .dl{font-size:8px;font-weight:700;text-transform:uppercase;color:var(--c-muted);margin-bottom:3px}
.day-box input{width:100%;padding:3px;border:1px solid var(--c-border);border-radius:3px;font-size:10px;text-align:center}
.model-card{border:1px solid var(--c-border);border-radius:7px;padding:10px;margin-bottom:8px}
.mch{display:flex;align-items:center;gap:7px;margin-bottom:7px}
.mch input{flex:1;font-size:12px;font-weight:600;padding:4px 7px;border:1px solid var(--c-border);border-radius:4px}

/* Fin de poste */
.fp-top{text-align:center;padding:16px 0 8px}
.fp-top h2{font-size:22px;font-weight:800;color:var(--c-primary)}
.fp-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}
.fp-sc{background:var(--c-card);border-radius:var(--radius);padding:14px;text-align:center;box-shadow:var(--shadow);border:1px solid var(--c-border)}
.fp-sc .big{font-size:30px;font-weight:800;color:var(--c-primary)}
.fp-sc .lbl{font-size:9px;text-transform:uppercase;color:var(--c-muted);font-weight:700;margin-top:3px}
.fp-acts{display:flex;justify-content:center;gap:10px;padding:16px 0}

/* History */
.htbl{width:100%;border-collapse:collapse;font-size:12px}
.htbl th{text-align:left;padding:7px 9px;background:var(--c-primary);color:#fff;font-size:9px;text-transform:uppercase;letter-spacing:.5px}
.htbl td{padding:6px 9px;border-bottom:1px solid var(--c-border)}
.htbl tr:hover td{background:var(--c-primary-light)}
.tg{color:#16a34a;font-weight:700}
.tm{color:#d97706;font-weight:700}
.tb{color:#dc2626;font-weight:700}

/* Utils */
.hidden{display:none!important}
.flex{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.mt8{margin-top:8px}
.sl{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--c-muted);font-weight:700;margin-bottom:7px}

@media(max-width:880px){.form-3col{grid-template-columns:1fr 1fr}}
@media(max-width:580px){.form-3col{grid-template-columns:1fr};.main-row{grid-template-columns:1fr;height:auto}}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="v-login" style="display:flex;align-items:center;justify-content:center;min-height:100vh;background:linear-gradient(135deg,#1a1f5e,#2d3480)">
  <div class="login-card">
    <h1>⚙ KPI-ORC</h1>
    <p class="sub">Système de suivi de production</p>
    <div class="lf">
      <label>Pilote</label>
      <select id="ln-pilot"><option value="">-- Choisir --</option></select>
    </div>
    <div class="lf">
      <label>Poste</label>
      <select id="ln-poste">
        <option value="Matin">Matin (05h–13h)</option>
        <option value="Après-midi">Après-midi (13h–21h)</option>
        <option value="Nuit">Nuit (21h–05h)</option>
      </select>
    </div>
    <div class="lf">
      <label>Mot de passe</label>
      <input type="password" id="ln-pw" placeholder="••••" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="btn-login" onclick="doLogin()">Valider</button>
    <div class="login-err" id="ln-err"></div>
  </div>
</div>

<!-- APP -->
<div id="app" class="hidden">
  <div id="app-header">
    <div class="logo">⚙ KPI-ORC</div>
    <nav>
      <button class="nav-tab" id="nt-main" onclick="goTab('main')">Accueil</button>
      <button class="nav-tab" id="nt-prod" style="display:none" onclick="goTab('prod')">▶ Prod en cours</button>
      <button class="nav-tab" id="nt-hist" onclick="goTab('history')">Historique</button>
      <button class="nav-tab" id="nt-cfg" onclick="goTab('settings')">Paramètres</button>
    </nav>
    <div id="hdr-info"></div>
  </div>
  <div id="alert-strip" class="alert-strip"></div>

  <!-- MAIN -->
  <div id="v-main" class="view">
    <div class="ctr">
      <div id="mc-active" class="prod-active-card hidden">
        <div class="flex">
          <span class="of-badge" id="mc-of">OF ---</span>
          <span id="mc-inf" style="font-weight:600;color:var(--c-muted)"></span>
          <span style="margin-left:auto">
            <button class="btn btn-prim" onclick="goTab('prod')">▶ Reprendre la production</button>
          </span>
        </div>
      </div>
      <div class="main-row">
        <button class="btn-start" id="btn-start" onclick="doStartProd()">▶ Démarrer production</button>
        <button class="btn-fp" onclick="doFinPoste()">
          <span style="font-size:20px">⏹</span>
          <span>Fin de poste</span>
        </button>
      </div>
      <div class="kpi-grid">
        <div class="card"><h3>Poste en cours — TRS</h3><div id="mkpi-cur" style="min-height:70px"></div></div>
        <div class="card"><h3>Dernier poste</h3><div id="mkpi-prev" style="min-height:70px"></div></div>
        <div class="card"><h3>Derniers OF</h3><div id="mkpi-of" style="min-height:70px"></div></div>
      </div>
    </div>
  </div>

  <!-- PROD -->
  <div id="v-prod" class="view">
    <div class="ctr">
      <div class="ph-row">
        <span class="of-badge" id="ph-of">OF ---</span>
        <div>
          <div style="font-size:9px;text-transform:uppercase;color:var(--c-muted);font-weight:700">Durée OF</div>
          <div class="of-timer" id="ph-timer">00:00:00</div>
        </div>
        <div id="ph-stop-area" class="hidden">
          <div style="font-size:9px;text-transform:uppercase;color:var(--c-muted);font-weight:700">Arrêt / Pause</div>
          <div class="stop-timer-big" id="ph-stop-tmr">00:00:00</div>
        </div>
        <div id="ph-slbl" class="stop-lbl hidden">ARRÊT</div>
        <div style="margin-left:auto">
          <button class="btn btn-endprod" onclick="doEndProdPreview()">⏹ Fin de production</button>
        </div>
      </div>

      <div id="stop-bar" class="stop-bar hidden">
        <div><div class="stype" id="sb-type">—</div><div style="font-size:10px;opacity:.8">en cours</div></div>
        <div class="selap" id="sb-elap">0:00</div>
        <button class="btn btn-endstop" onclick="doEndStop()">✓ Terminer</button>
      </div>

      <div class="ctrls" id="prod-ctrls">
        <button class="btn btn-pause" id="btn-pause" onclick="doPause()">⏸ Pause</button>
        <button class="btn btn-arret" onclick="openStopModal()">⚠ Arrêt</button>
        <button class="btn btn-sec" onclick="saveFormNow()" style="margin-left:auto">💾 Sauvegarder</button>
      </div>

      <div class="form-3col">
        <!-- Zone Identification -->
        <div class="fzone zi">
          <h4>📋 Identification</h4>
          <div class="fr"><label>N° OF *</label><input id="f-of_num" placeholder="OF123456"></div>
          <div class="fr ro"><label>Date</label><input id="f-date" readonly></div>
          <div class="fr ro"><label>Poste</label><input id="f-poste" readonly></div>
          <div class="fr ro"><label>Pilote</label><input id="f-pilote" readonly></div>
          <div class="fr"><label>Co-Pilote</label><input id="f-copilote" placeholder="Nom"></div>
          <div class="fr"><label>Nb Personnes</label><input id="f-nb_pers" type="number" min="1" value="2"></div>
          <div class="fr"><label>Taille</label><select id="f-taille"><option value="">--</option></select></div>
          <div class="fr"><label>Code Produit</label><input id="f-code_prod" placeholder="CODE01"></div>
          <div class="fr"><label>Type Produit</label><select id="f-type_prod"><option value="">--</option></select></div>
          <div class="fr"><label>Kit</label><select id="f-kit"><option value="">Non</option><option value="oui">Oui</option></select></div>
        </div>

        <!-- Zone Production -->
        <div class="fzone zp">
          <h4>🏭 Production</h4>
          <div class="fr big"><label>Qté Fabriquée *</label><input id="f-qte_fab" type="number" min="0" placeholder="0"></div>
          <div class="fr big"><label>Qté Emballée</label><input id="f-qte_emb" type="number" min="0" placeholder="0"></div>
          <div class="fr"><label>Poids Garnissage (g)</label><input id="f-poids" type="number" min="0" placeholder="350"></div>
          <div class="fr"><label>Fibre</label><select id="f-fibre"><option value="">--</option></select></div>
          <div class="fr"><label>OF Taie</label><input id="f-of_taie" placeholder="OF-T001"></div>
          <div class="fr"><label>Traca Fibre</label><select id="f-traca"><option value="">--</option></select></div>
          <div class="fr"><label>Réf Taie</label><input id="f-ref_taie" placeholder="REF-T01"></div>
          <div class="fr"><label>Manquant MP (min)</label><input id="f-duree_mq_mp" type="number" min="0" placeholder="0"></div>
          <div class="fr"><label>Manquant Personnel (min)</label><input id="f-manquant_pers" type="number" min="0" placeholder="0"></div>
        </div>

        <!-- Zone Qualité -->
        <div class="fzone zq">
          <h4>✅ Qualité</h4>
          <div class="fr"><label>Qté Init Taie</label><input id="f-qte_init_taie" type="number" min="0" value="0"></div>
          <div class="fr"><label>Nb Taie 2nd Choix</label><input id="f-nb_taie2_choix" type="number" min="0" value="0"></div>
          <div class="fr"><label>Nb Défaut Couture</label><input id="f-nb_def_cout" type="number" min="0" value="0"></div>
          <div class="fr"><label>Mq Taie</label><input id="f-mq_taie" type="number" min="0" value="0"></div>
          <div class="fr"><label>Mq Housse/Encart</label><input id="f-mq_housse_encart" type="number" min="0" value="0"></div>
          <div class="fr"><label>Nb PP Cousue</label><input id="f-nb_pp_cousue" type="number" min="0" value="0"></div>
          <div class="fr"><label>Commentaire</label><textarea id="f-comment" placeholder="Observations..."></textarea></div>
        </div>
      </div>

      <!-- Timeline 4h -->
      <div class="tl-wrap">
        <h4>Timeline — 4 dernières heures</h4>
        <svg id="tl-svg" viewBox="0 0 800 48" preserveAspectRatio="none" style="width:100%;height:48px;display:block">
          <rect x="0" y="8" width="800" height="32" fill="#e2e8f0" rx="4"/>
          <text x="2" y="46" font-size="9" fill="#94a3b8">-4h</text>
          <text x="770" y="46" font-size="9" fill="#94a3b8">Maintenant</text>
        </svg>
      </div>

      <!-- Recap stops -->
      <div class="recap">
        <h4>Arrêts &amp; pauses du poste</h4>
        <div id="recap-list"><span style="color:var(--c-muted);font-size:11px">Aucun arrêt enregistré</span></div>
      </div>
    </div>
  </div>

  <!-- FIN DE POSTE -->
  <div id="v-finposte" class="view">
    <div class="ctr">
      <div class="fp-top"><h2>Fin de poste</h2><p id="fp-who" style="color:var(--c-muted);margin-top:3px"></p></div>
      <div class="fp-stats" id="fp-stats">
        <div class="fp-sc"><div class="big" id="fp-trs">--%</div><div class="lbl">TRS Poste (shift)</div></div>
        <div class="fp-sc"><div class="big" id="fp-eq">0</div><div class="lbl">Équivalence totale</div></div>
        <div class="fp-sc"><div class="big" id="fp-nof">0</div><div class="lbl">Nb OF</div></div>
        <div class="fp-sc"><div class="big" id="fp-st">0 min</div><div class="lbl">Durée prod totale</div></div>
      </div>
      <div class="tl-wrap"><h4>Timeline du poste</h4>
        <svg id="fp-tl" viewBox="0 0 800 48" preserveAspectRatio="none" style="width:100%;height:48px;display:block">
          <rect x="0" y="8" width="800" height="32" fill="#e2e8f0" rx="4"/>
        </svg>
      </div>
      <div class="card" style="margin-bottom:14px">
        <h3>Productions du poste</h3>
        <div id="fp-prods"></div>
      </div>
      <div class="fp-acts">
        <button class="btn btn-sec" onclick="goTab('main')">← Retour</button>
        <button class="btn btn-danger" onclick="confirmFinPoste()" style="font-size:14px;padding:11px 22px">⏹ Confirmer fin de poste &amp; Déconnexion</button>
      </div>
    </div>
  </div>

  <!-- HISTORY -->
  <div id="v-history" class="view">
    <div class="ctr">
      <div class="flex" style="margin-bottom:10px">
        <div class="sl" style="margin:0;font-size:13px">Historique</div>
        <input type="date" id="hist-dt" style="padding:5px 9px;border:1px solid var(--c-border);border-radius:5px;font-size:12px" onchange="loadHist()">
      </div>
      <div style="overflow-x:auto">
        <table class="htbl"><thead><tr id="hist-hd"></tr></thead><tbody id="hist-bd"></tbody></table>
      </div>
    </div>
  </div>

  <!-- SETTINGS -->
  <div id="v-settings" class="view">
    <div class="ctr">
      <div class="sl" style="font-size:13px;margin-bottom:12px">Paramètres</div>

      <div class="ss">
        <h3>🔐 Mots de passe pilotes</h3>
        <div id="pwd-list"></div>
        <div class="flex mt8">
          <input id="np-name" placeholder="Nom pilote" style="flex:1;padding:6px 9px;border:1px solid var(--c-border);border-radius:5px;font-size:12px">
          <input id="np-pw" type="password" placeholder="Mot de passe" style="flex:1;padding:6px 9px;border:1px solid var(--c-border);border-radius:5px;font-size:12px">
          <button class="btn btn-prim" onclick="addPilot()">+ Ajouter</button>
        </div>
        <div class="flex mt8">
          <label style="font-size:11px;font-weight:600">MDP admin requis :</label>
          <input id="adm-pw" type="password" value="1234" style="width:80px;padding:5px 7px;border:1px solid var(--c-border);border-radius:5px;font-size:12px">
          <button class="btn btn-ok" onclick="savePwds()">💾 Enregistrer MDP</button>
        </div>
      </div>

      <div class="ss">
        <h3>🕐 Modèles horaires (par poste &amp; par jour)</h3>
        <div id="models-list"></div>
        <button class="btn btn-sec mt8" onclick="addModel()">+ Nouveau modèle</button>
        <div class="flex mt8">
          <label style="font-size:11px;font-weight:600">MDP admin :</label>
          <input id="adm-pw2" type="password" value="1234" style="width:80px;padding:5px 7px;border:1px solid var(--c-border);border-radius:5px;font-size:12px">
          <button class="btn btn-ok" onclick="saveModels()">💾 Enregistrer modèles</button>
        </div>
      </div>

      <div class="ss">
        <h3>⚙ Référence production 8h</h3>
        <div class="flex">
          <label style="font-weight:600;font-size:12px">Prod ref (unités/8h) :</label>
          <input id="cfg-pr" type="number" style="width:90px;padding:5px 7px;border:1px solid var(--c-border);border-radius:5px">
          <input id="adm-pw3" type="password" value="1234" placeholder="MDP admin" style="width:80px;padding:5px 7px;border:1px solid var(--c-border);border-radius:5px;font-size:12px">
          <button class="btn btn-ok" onclick="saveProdRef()">Enregistrer</button>
        </div>
      </div>
    </div>
  </div>

</div><!-- /app -->

<!-- MODAL: type d'arrêt -->
<div class="overlay" id="m-stop">
  <div class="mbox">
    <h2>⚠ Type d'arrêt</h2>
    <div class="stop-grid" id="stop-grid"></div>
    <div class="m-acts"><button class="btn btn-sec" onclick="closeM('m-stop')">Annuler</button></div>
  </div>
</div>

<!-- MODAL: fin de production -->
<div class="overlay" id="m-endprod">
  <div class="mbox wide">
    <h2>⏹ Fin de production — Récapitulatif</h2>
    <div class="ep-grid" id="ep-stats"></div>
    <div style="margin-bottom:10px">
      <div class="sl">Arrêts &amp; pauses</div>
      <table class="ep-tbl"><thead><tr><th>Type</th><th>Durée</th><th>%</th></tr></thead><tbody id="ep-stops"></tbody></table>
    </div>
    <div class="tl-wrap"><h4>Timeline</h4>
      <svg id="ep-tl" viewBox="0 0 800 48" preserveAspectRatio="none" style="width:100%;height:48px;display:block">
        <rect x="0" y="8" width="800" height="32" fill="#e2e8f0" rx="4"/>
      </svg>
    </div>
    <div class="m-acts">
      <button class="btn btn-sec" onclick="closeM('m-endprod')">Annuler</button>
      <button class="btn btn-danger" onclick="confirmEndProd()" style="font-size:14px;padding:10px 20px">✓ Confirmer fin de production</button>
    </div>
  </div>
</div>

<!-- MODAL: éditer arrêt -->
<div class="overlay" id="m-editstop">
  <div class="mbox">
    <h2>✏ Modifier l'arrêt</h2>
    <input type="hidden" id="es-key">
    <div class="fr" style="margin-bottom:9px"><label>Type</label><select id="es-type"></select></div>
    <div class="fr" style="margin-bottom:9px"><label>Heure début</label><input type="time" id="es-deb" step="60"></div>
    <div class="fr" style="margin-bottom:9px"><label>Heure fin</label><input type="time" id="es-fin" step="60"></div>
    <div class="fr" style="margin-bottom:9px"><label>Commentaire</label><input type="text" id="es-cmt"></div>
    <div class="m-acts">
      <button class="btn btn-sec" onclick="closeM('m-editstop')">Annuler</button>
      <button class="btn btn-danger" onclick="deleteStop()">🗑 Supprimer</button>
      <button class="btn btn-ok" onclick="saveEditStop()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<script>
// ── Constants ──
const STOP_TYPES = [
  "PB Technique: Carde","PB Technique: Gainonnage","PB Technique: Finition",
  "PB Technique: Encartage","PB Qualité","Manque MP",
  "Formation / Réunion","Nettoyage","Divers"
];
const STOP_COL = {
  "PB Technique: Carde":"#b91c1c","PB Technique: Gainonnage":"#dc2626",
  "PB Technique: Finition":"#ef4444","PB Technique: Encartage":"#f87171",
  "PB Qualité":"#ea580c","Manque MP":"#d97706",
  "Formation / Réunion":"#7c3aed","Nettoyage":"#0891b2",
  "Divers":"#64748b","Pause Pilote":"#7c3aed","nettoyage":"#0891b2"
};
const FORM_FIELDS = ["of_num","copilote","nb_pers","taille","code_prod","type_prod","poids","fibre","of_taie","traca","ref_taie","kit","qte_fab","qte_emb","qte_init_taie","nb_taie2_choix","nb_def_cout","mq_taie","mq_housse_encart","nb_pp_cousue","duree_mq_mp","manquant_pers","comment"];

// ── State ──
let ST = {};
let gEvts = [];
let _curStopKey = null;
let _curStopElap = 0;
let _ofElapAtPoll = 0;
let _lastPoll = Date.now();
let _ticker = null;
let _curTab = 'main';
let _cfgPwds = {};
let _cfgModels = [];
window._evMap = {};

// ── Init ──
document.addEventListener('DOMContentLoaded', async () => {
  buildStopGrid();
  buildStopTypeOpts();
  await loadLists();
  const s = await apiFetch('/api/state');
  if (s && s.pilot) showApp(s);
  else { document.getElementById('v-login').style.display=''; }
  setInterval(pollState, 5000);
  setInterval(pollEvts, 8000);
});

async function loadLists() {
  const d = await apiFetch('/api/lists');
  if (!d) return;
  popSel('f-taille', d.tailles||[]);
  popSel('f-type_prod', d.types_prod||[]);
  popSel('f-fibre', d.fibres||[]);
  popSel('f-traca', d.tracas||[]);
  const pil = d.pilotes||[];
  const sel = document.getElementById('ln-pilot');
  pil.forEach(p => { const o=document.createElement('option'); o.value=p; o.textContent=p; sel.appendChild(o); });
}

function popSel(id, vals) {
  const s = document.getElementById(id);
  if (!s) return;
  const cur = s.value;
  while (s.options.length>1) s.remove(1);
  vals.forEach(v => { const o=document.createElement('option'); o.value=v; o.textContent=v; s.appendChild(o); });
  if (cur) s.value = cur;
}

// ── Login ──
async function doLogin() {
  const pilot = document.getElementById('ln-pilot').value;
  const poste = document.getElementById('ln-poste').value;
  const pw = document.getElementById('ln-pw').value;
  document.getElementById('ln-err').textContent = '';
  if (!pilot) { document.getElementById('ln-err').textContent='Choisir un pilote'; return; }
  const r = await fetch('/api/login', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pilot,poste,pw})});
  if (!r) return;
  const d = await r.json();
  if (d.ok) {
    const s = await apiFetch('/api/state');
    showApp(s||{pilot,poste});
  } else {
    document.getElementById('ln-err').textContent = d.error||'Erreur connexion';
  }
}

function showApp(s) {
  document.getElementById('v-login').style.display='none';
  document.getElementById('app').classList.remove('hidden');
  if (s.poste) { document.getElementById('f-poste').value=s.poste; }
  if (s.pilot) { document.getElementById('f-pilote').value=s.pilot; }
  setToday();
  pollState();
  pollEvts();
  startTicker();
  loadCfg();
  goTab(s.prod_active ? 'prod' : 'main');
  document.getElementById('hist-dt').value = new Date().toISOString().slice(0,10);
}

function setToday() {
  const d = document.getElementById('f-date');
  if (d) d.value = new Date().toLocaleDateString('fr-CA');
}

// ── Navigation ──
function goTab(tab) {
  _curTab = tab;
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('on'));
  document.querySelectorAll('.nav-tab').forEach(t=>t.classList.remove('active','tab-prod'));
  const vm = {main:'v-main',prod:'v-prod',history:'v-history',settings:'v-settings',finposte:'v-finposte'};
  const el = document.getElementById(vm[tab]);
  if (el) el.classList.add('on');
  const ntm = {main:'nt-main',prod:'nt-prod',history:'nt-hist',settings:'nt-cfg'};
  const nt = document.getElementById(ntm[tab]);
  if (nt) nt.classList.add('active');
  if (tab==='history') loadHist();
  if (tab==='finposte') loadFPData();
}

// ── State polling ──
async function pollState() {
  const s = await apiFetch('/api/state');
  if (!s) return;
  ST = s;
  _ofElapAtPoll = s.of_elapsed_s||0;
  _lastPoll = Date.now();

  // Determine current stop
  if (s.is_paused) {
    _curStopKey = '_pause';
    _curStopElap = s.pause_total_s||0;
    if (s.pause_start_iso) {
      _curStopElap += (Date.now() - new Date(s.pause_start_iso).getTime())/1000;
    }
  } else if (s.active_stops && s.active_stops.length>0) {
    const k = s.active_stops[0];
    _curStopKey = k;
    _curStopElap = s.timers&&s.timers[k] ? s.timers[k].elapsed : 0;
  } else {
    _curStopKey = null;
    _curStopElap = 0;
  }

  applyState(s);
  renderMainKpi();
}

async function pollEvts() {
  const e = await apiFetch('/api/events_list');
  if (!Array.isArray(e)) return;
  gEvts = e;
  renderTL('tl-svg', gEvts);
  renderRecap(gEvts);
}

function applyState(s) {
  // Header info
  const hi = document.getElementById('hdr-info');
  if (hi && s.pilot) hi.innerHTML = `${s.pilot}<br><span style="opacity:.7">${s.poste||''}</span>`;

  // Stop alert
  const al = document.getElementById('alert-strip');
  const stopOn = _curStopKey !== null;
  if (stopOn) {
    const lbl = _curStopKey==='_pause' ? 'PAUSE' : _curStopKey;
    al.textContent = `⚠ ${lbl} EN COURS`;
    al.classList.add('on');
    document.body.classList.add('stop-active');
  } else {
    al.classList.remove('on');
    document.body.classList.remove('stop-active');
  }

  // Prod tab
  const tp = document.getElementById('nt-prod');
  if (tp) tp.style.display = s.prod_active ? '' : 'none';

  // Main: active card
  const mac = document.getElementById('mc-active');
  const mof = document.getElementById('mc-of');
  const minf = document.getElementById('mc-inf');
  const bstart = document.getElementById('btn-start');
  if (s.prod_active) {
    mac&&mac.classList.remove('hidden');
    mof&&(mof.textContent=s.form&&s.form.of_num?s.form.of_num:'OF ---');
    minf&&(minf.textContent=`${(s.form&&s.form.taille)||''} ${(s.form&&s.form.type_prod)||''} • ${s.pilot||''}`);
    bstart&&(bstart.disabled=true);
  } else {
    mac&&mac.classList.add('hidden');
    bstart&&(bstart.disabled=false);
  }

  // Prod header
  if (s.prod_active && s.form) {
    const el = document.getElementById('ph-of');
    if (el) el.textContent = s.form.of_num||'OF ---';
    fillFormFromState(s.form);
  }

  // Stop/pause bar
  const sb = document.getElementById('stop-bar');
  const pc = document.getElementById('prod-ctrls');
  const psa = document.getElementById('ph-stop-area');
  const psl = document.getElementById('ph-slbl');
  const pbtn = document.getElementById('btn-pause');
  if (_curStopKey) {
    sb&&sb.classList.remove('hidden');
    pc&&pc.classList.add('hidden');
    psa&&psa.classList.remove('hidden');
    psl&&psl.classList.remove('hidden');
    const sbt = document.getElementById('sb-type');
    if (sbt) sbt.textContent = _curStopKey==='_pause'?'Pause Pilote':_curStopKey;
  } else {
    sb&&sb.classList.add('hidden');
    psa&&psa.classList.add('hidden');
    psl&&psl.classList.add('hidden');
    if (s.prod_active) {
      pc&&pc.classList.remove('hidden');
    }
  }
  // Pause button label
  if (pbtn) pbtn.textContent = s.is_paused ? '▶ Reprendre' : '⏸ Pause';
}

// ── Local ticker ──
function startTicker() {
  if (_ticker) clearInterval(_ticker);
  _ticker = setInterval(() => {
    if (!ST.prod_active) return;
    const dt = (Date.now()-_lastPoll)/1000;

    // OF timer ticks unless paused
    const ofEl = _ofElapAtPoll + (!ST.is_paused ? dt : 0);
    const phT = document.getElementById('ph-timer');
    if (phT) phT.textContent = fmtDur(ofEl);

    // Stop/pause elapsed
    if (_curStopKey) {
      const stopEl = _curStopElap + dt;
      const phS = document.getElementById('ph-stop-tmr');
      if (phS) phS.textContent = fmtDur(stopEl);
      const sbE = document.getElementById('sb-elap');
      if (sbE) sbE.textContent = fmtDur2(stopEl);
    }
  }, 1000);
}

// ── Start prod ──
async function doStartProd() {
  const d = await fetch('/api/start_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
  if (!d) return;
  const r = await d.json();
  if (r.ok) {
    setToday();
    await pollState();
    await pollEvts();
    goTab('prod');
  } else { toast(r.error||'Erreur','err'); }
}

// ── Stop/Pause ──
function buildStopGrid() {
  const g = document.getElementById('stop-grid');
  if (!g) return;
  STOP_TYPES.forEach(t => {
    const b = document.createElement('button');
    b.className='stop-btn';
    b.textContent=t;
    b.onclick=()=>{ closeM('m-stop'); doStartStop(t); };
    g.appendChild(b);
  });
}

function openStopModal() { openM('m-stop'); }

async function doPause() {
  await fetch('/api/toggle_pause',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
}

async function doStartStop(key) {
  const cat = key.toLowerCase().startsWith('pb') ? 'pb' : key==='Pause Pilote'?'pause':key.toLowerCase().replace(/\s+/g,'_').slice(0,12);
  await fetch('/api/start_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,cat})});
  await pollState();
  await pollEvts();
}

async function doEndStop() {
  const key = _curStopKey==='_pause' ? null : _curStopKey;
  const body = key ? JSON.stringify({key}) : '{}';
  if (_curStopKey==='_pause') {
    await fetch('/api/toggle_pause',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  } else {
    await fetch('/api/end_stop',{method:'POST',headers:{'Content-Type':'application/json'},body});
  }
  await pollState();
  await pollEvts();
}

// ── Form ──
function collectForm() {
  const f = {};
  FORM_FIELDS.forEach(k => {
    const el = document.getElementById('f-'+k);
    if (!el) return;
    f[k] = el.type==='number' ? (parseFloat(el.value)||0) : el.value;
  });
  f.pilote = document.getElementById('f-pilote')?.value||ST.pilot||'';
  f.poste = document.getElementById('f-poste')?.value||ST.poste||'';
  return f;
}

function fillFormFromState(form) {
  if (!form) return;
  FORM_FIELDS.forEach(k => {
    const el = document.getElementById('f-'+k);
    if (!el) return;
    const v = form[k];
    if (v!==undefined && v!==null) el.value = v;
  });
}

async function saveFormNow() {
  const f = collectForm();
  await fetch('/api/save_form',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(f)});
  toast('Formulaire sauvegardé','ok');
}

// Auto-save form every 30s when prod active
setInterval(()=>{ if(ST.prod_active&&_curTab==='prod') { const f=collectForm(); fetch('/api/save_form',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(f)}); }}, 30000);

// ── End production ──
async function doEndProdPreview() {
  const f = collectForm();
  const r = await fetch('/api/preview_end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:f})});
  if (!r||!r.ok) { if(confirm('Confirmer fin de production ?')) confirmEndProd(); return; }
  const d = await r.json();
  renderEPModal(d, f);
  openM('m-endprod');
}

function renderEPModal(d, f) {
  document.getElementById('ep-stats').innerHTML = `
    <div class="ep-stat"><div class="val">${fmtTRS(d.trs)}</div><div class="lbl">TRS OF</div></div>
    <div class="ep-stat"><div class="val">${(d.equiv||0).toFixed(1)}</div><div class="lbl">Équivalence</div></div>
    <div class="ep-stat"><div class="val">${fmtD2(d.prod_s||0)}</div><div class="lbl">Durée prod</div></div>
    <div class="ep-stat"><div class="val">${fmtD2(d.stop_s||0)}</div><div class="lbl">Total arrêts</div></div>
    <div class="ep-stat"><div class="val">${d.c1||0}</div><div class="lbl">Cadence/h</div></div>
    <div class="ep-stat"><div class="val">${d.c2||0}</div><div class="lbl">Cad/h/pers</div></div>
  `;
  // Build stop summary from tl_events
  const evts = d.tl_events||[];
  const stopMap = {};
  const totalS = d.stop_s||1;
  evts.forEach(e=>{
    if (!e.key||e.key==='prod') return;
    const dur = (e.dur_s||0);
    stopMap[e.key] = (stopMap[e.key]||0)+dur;
  });
  const tbody = document.getElementById('ep-stops');
  tbody.innerHTML = Object.entries(stopMap).map(([k,s])=>`
    <tr><td>${k}</td><td>${fmtD2(s)}</td><td>${totalS>0?Math.round(s/totalS*100):0}%</td></tr>
  `).join('')||'<tr><td colspan="3" style="color:var(--c-muted)">Aucun arrêt</td></tr>';
  // Timeline in modal
  drawTL('ep-tl', evts, d.debut, d.now_str);
}

async function confirmEndProd() {
  const f = collectForm();
  closeM('m-endprod');
  const r = await fetch('/api/end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:f})});
  if (!r) return;
  const d = await r.json();
  if (d.ok) { await pollState(); await pollEvts(); goTab('main'); toast('Production enregistrée','ok'); }
  else toast(d.error||'Erreur','err');
}

// ── Edit stop ──
function buildStopTypeOpts() {
  const s = document.getElementById('es-type');
  if (!s) return;
  ['Pause Pilote',...STOP_TYPES,'nettoyage'].forEach(t=>{
    const o=document.createElement('option'); o.value=t; o.textContent=t; s.appendChild(o);
  });
}

function openEditStop(key) {
  const ev = window._evMap[key];
  if (!ev) return;
  document.getElementById('es-key').value=key;
  document.getElementById('es-type').value=ev.type||'';
  const d=ev.debut||'',f2=ev.fin||'';
  document.getElementById('es-deb').value=d.length>=5?d.slice(0,5):d;
  document.getElementById('es-fin').value=f2.length>=5?f2.slice(0,5):f2;
  document.getElementById('es-cmt').value=ev.comment||'';
  openM('m-editstop');
}

async function saveEditStop() {
  const key = document.getElementById('es-key').value;
  const ev = window._evMap[key];
  if (!ev) return;
  const data = {
    row_num: ev.row_num,
    type: document.getElementById('es-type').value,
    heure_debut: document.getElementById('es-deb').value,
    heure_fin: document.getElementById('es-fin').value,
    comment: document.getElementById('es-cmt').value,
  };
  const r = await fetch('/api/edit_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  if (r&&r.ok) { closeM('m-editstop'); await pollEvts(); toast('Modifié','ok'); }
  else toast('Erreur','err');
}

async function deleteStop() {
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if (!ev||!confirm('Supprimer ?')) return;
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({row_num:ev.row_num})});
  if (r&&r.ok) { closeM('m-editstop'); await pollEvts(); toast('Supprimé','ok'); }
}

// ── Timeline ──
function renderTL(svgId, evts) {
  const now = new Date();
  const s4h = new Date(now-4*3600*1000);
  drawTLFromISO(svgId, evts, s4h.toISOString(), now.toISOString());
}

function drawTL(svgId, tlEvts, debutHMS, finHMS) {
  // tlEvts are {key, dur_s, ...} from preview
  const svg = document.getElementById(svgId);
  if (!svg) return;
  const W=800,Y=8,H2=32,H=48;
  let html = `<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  // If debut/fin provided as HH:MM:SS, reconstruct
  if (!debutHMS||!finHMS) { svg.innerHTML=html; return; }
  const base = new Date(); base.setHours(0,0,0,0);
  const parseHMS = s => { const [h,m,sec]=(s||'').split(':'); return base.getTime()+(parseInt(h)||0)*3600000+(parseInt(m)||0)*60000+(parseInt(sec)||0)*1000; };
  const tS=parseHMS(debutHMS), tE=parseHMS(finHMS);
  const span=tE-tS; if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  // Prod background
  html+=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  // Events
  let cur=tS;
  (tlEvts||[]).forEach(e=>{
    if (!e.key||e.key==='prod') return;
    const x1=toX(cur), x2=toX(cur+e.dur_s*1000);
    const col=STOP_COL[e.key]||'#94a3b8';
    html+=`<rect x="${x1}" y="${Y}" width="${Math.max(1,x2-x1)}" height="${H2}" fill="${col}" rx="2" opacity="0.9"/>`;
    cur+=e.dur_s*1000;
  });
  const fmtHM=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  html+=`<text x="2" y="${H-2}" font-size="9" fill="#64748b">${debutHMS.slice(0,5)}</text>`;
  html+=`<text x="${W-35}" y="${H-2}" font-size="9" fill="#64748b">${finHMS.slice(0,5)}</text>`;
  svg.innerHTML=html;
}

function drawTLFromISO(svgId, evts, startIso, endIso) {
  const svg = document.getElementById(svgId);
  if (!svg) return;
  const W=800,Y=8,H2=32,H=48;
  let html = `<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  const tS=new Date(startIso).getTime(), tE=new Date(endIso).getTime();
  const span=tE-tS; if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));

  // Prod background
  if (ST.prod_active) {
    const ps = ST.of_start_iso ? new Date(ST.of_start_iso).getTime() : tS;
    const x1=toX(ps),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  }

  // Events (stops/pauses)
  (evts||[]).forEach((ev,i)=>{
    const key=ev.debut||i;
    window._evMap[key]=ev;
    const t1=parseHMStoT(ev.debut, ev.date), t2=parseHMStoT(ev.fin, ev.date);
    if (!t1) return;
    const x1=toX(t1),x2=toX(t2||tE);
    if(x2<=x1) return;
    const col=STOP_COL[ev.type||'']||'#94a3b8';
    html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${col}" rx="2" opacity="0.85"/>`;
  });

  // Current stop
  if (_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active) {
    const stopStart=tE-((_curStopElap+(Date.now()-_lastPoll)/1000)*1000);
    const x1=toX(stopStart),x2=toX(tE);
    if(x2>x1){const col=STOP_COL[_curStopKey]||'#dc2626';html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${col}" rx="2" opacity="0.9"/>`;}
  }

  const fmtT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  html+=`<text x="2" y="${H-2}" font-size="9" fill="#64748b">${fmtT(tS)}</text>`;
  html+=`<text x="${W-35}" y="${H-2}" font-size="9" fill="#64748b">${fmtT(tE)}</text>`;
  html+=`<line x1="${W/2}" y1="${Y}" x2="${W/2}" y2="${Y+H2}" stroke="#94a3b8" stroke-width="0.5" stroke-dasharray="2,2"/>`;
  html+=`<text x="${W/2-12}" y="${H-2}" font-size="9" fill="#94a3b8">${fmtT((tS+tE)/2)}</text>`;
  svg.innerHTML=html;
}

function parseHMStoT(hms, dateStr) {
  if (!hms) return null;
  try {
    const base = dateStr ? new Date(dateStr.slice(0,10)).getTime() : new Date().setHours(0,0,0,0);
    const [h,m,s]=(hms||'00:00:00').split(':').map(Number);
    return base+h*3600000+m*60000+(s||0)*1000;
  } catch(e){return null;}
}

// ── Recap stops list ──
function renderRecap(evts) {
  const c=document.getElementById('recap-list');
  if (!c) return;
  window._evMap={};
  const stops=(evts||[]).filter(e=>e.type&&!['Production','prod',''].includes((e.type||'').toLowerCase()));
  if (!stops.length) { c.innerHTML='<span style="color:var(--c-muted);font-size:11px">Aucun arrêt enregistré</span>'; return; }
  let html='';
  stops.forEach((ev,i)=>{
    const key=ev.debut||i;
    window._evMap[String(key)]=ev;
    const col=STOP_COL[ev.type||'']||'#94a3b8';
    html+=`<div class="si">
      <div class="sdot" style="background:${col}"></div>
      <div class="sname">${ev.type||'?'}</div>
      <div class="sdur">${ev.duree||calcDur(ev.debut,ev.fin)||'?'}</div>
      <button class="btn-edit" onclick="openEditStop('${esc(String(key))}')">✏</button>
    </div>`;
  });
  // Current live stop
  if (_curStopKey) {
    const lbl=_curStopKey==='_pause'?'Pause Pilote':_curStopKey;
    html+=`<div class="si" style="animation:blink .9s step-start infinite">
      <div class="sdot" style="background:${STOP_COL[lbl]||'#dc2626'}"></div>
      <div class="sname">${lbl} <span style="font-size:9px;color:var(--c-muted)">(en cours)</span></div>
      <div class="sdur" id="recap-live">--:--</div>
    </div>`;
  }
  c.innerHTML=html;
}

function calcDur(d,f){
  if(!d||!f) return '';
  try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
  const diff=p(f)-p(d); return diff>0?fmtDur(diff):'';}catch(e){return '';}
}

// ── Main KPI ──
async function renderMainKpi() {
  const d = await apiFetch('/api/history_today');
  if (!d) return;
  const cur=document.getElementById('mkpi-cur');
  const prev=document.getElementById('mkpi-prev');
  const ofL=document.getElementById('mkpi-of');
  if (!cur||!prev||!ofL) return;
  const trs=d.trs_shift||0;
  const rows=d.rows||[];
  if (rows.length) {
    cur.innerHTML=`<div style="font-size:30px;font-weight:800;color:${trs>=90?'var(--c-green)':trs>=75?'var(--c-yellow)':'var(--c-red)'}">${fmtTRS(trs)}</div>
    <div style="font-size:10px;color:var(--c-muted);margin-top:3px">${rows.length} OF • Éq: ${(d.tot_eq||0).toFixed(1)}</div>`;
  } else cur.innerHTML='<span style="color:var(--c-muted);font-size:11px">Aucune production ce poste</span>';
  prev.innerHTML='<span style="color:var(--c-muted);font-size:11px">—</span>';
  ofL.innerHTML=rows.slice(-5).reverse().map(r=>`
    <div class="flex" style="padding:3px 0;border-bottom:1px solid var(--c-border);font-size:11px">
      <span style="font-weight:600;flex:1">${r.of||'?'}</span>
      <span class="${(r.trs||0)>=90?'tg':(r.trs||0)>=75?'tm':'tb'}">${fmtTRS(r.trs)}</span>
    </div>`).join('')||'<span style="color:var(--c-muted)">--</span>';
}

// ── Fin de poste ──
async function doFinPoste() {
  if (ST.prod_active) { toast('Terminer la production en cours avant de finir le poste','err'); return; }
  goTab('finposte');
}

async function loadFPData() {
  const d = await apiFetch('/api/fin_poste_data');
  if (!d) return;
  const trs = d.trs_shift!==undefined?d.trs_shift:d.trs;
  document.getElementById('fp-trs').textContent = fmtTRS(trs);
  document.getElementById('fp-eq').textContent = (d.tot_equiv||0).toFixed(1);
  document.getElementById('fp-nof').textContent = d.nb_of||0;
  document.getElementById('fp-st').textContent = Math.round((d.tot_s||0)/60)+' min';
  document.getElementById('fp-who').textContent = `${d.pilot||ST.pilot||''} — ${ST.poste||''}`;

  // Timeline
  const shiftStart = d.shift_start_iso || new Date(Date.now()-8*3600*1000).toISOString();
  drawTLFromISO('fp-tl', gEvts, shiftStart, new Date().toISOString());

  // Productions list
  const fp = document.getElementById('fp-prods');
  if (fp && d.of_list) {
    fp.innerHTML=d.of_list.map(p=>`
      <div class="flex" style="padding:4px 0;border-bottom:1px solid var(--c-border);font-size:11px">
        <span style="font-weight:600;min-width:80px">${p.of||'?'}</span>
        <span style="color:var(--c-muted)">${p.taille||''} ${p.type_prod||''}</span>
        <span style="margin-left:auto">${p.qte_fab||0} pcs</span>
        <span class="${(p.trs||0)>=90?'tg':(p.trs||0)>=75?'tm':'tb'}">${fmtTRS(p.trs||0)}</span>
      </div>`).join('')||'<span style="color:var(--c-muted)">Aucune production</span>';
  }
}

async function confirmFinPoste() {
  if (!confirm('Confirmer fin de poste et se déconnecter ?')) return;
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  ST={}; _curStopKey=null;
  document.getElementById('app').classList.add('hidden');
  document.getElementById('ln-pw').value='';
  document.getElementById('ln-err').textContent='';
  document.getElementById('v-login').style.display='';
  toast('Bonne fin de poste !','ok');
}

// ── History ──
async function loadHist() {
  const dt=document.getElementById('hist-dt').value||new Date().toISOString().slice(0,10);
  const d=await apiFetch('/api/history?date='+dt);
  const rows=Array.isArray(d)?(d):(d&&d.rows?d.rows:[]);
  const hd=document.getElementById('hist-hd'), bd=document.getElementById('hist-bd');
  if (!hd||!bd) return;
  if (!rows.length) { bd.innerHTML='<tr><td colspan="9" style="text-align:center;color:var(--c-muted);padding:18px">Aucune donnée</td></tr>'; return; }
  const ks=['type','of','poste','pilote','debut','fin','qte_fab','equiv','trs'];
  const lb={type:'Type',of:'OF',poste:'Poste',pilote:'Pilote',debut:'Début',fin:'Fin',qte_fab:'Qté',equiv:'Éq',trs:'TRS'};
  hd.innerHTML=ks.map(k=>`<th>${lb[k]||k}</th>`).join('');
  bd.innerHTML=rows.map(row=>{
    const t=parseFloat(row.trs||row['trs%']||0);
    return '<tr>'+ks.map(k=>{
      const v=row[k]||'';
      if(k==='trs') return `<td class="${t>=90?'tg':t>=75?'tm':t>0?'tb':''}">${t>0?fmtTRS(t):''}</td>`;
      return `<td>${esc(String(v))}</td>`;
    }).join('')+'</tr>';
  }).join('');
}

// ── Settings ──
async function loadCfg() {
  const d=await apiFetch('/api/config');
  if (!d) return;
  _cfgPwds=d.pilot_passwords||{};
  _cfgModels=d.modeles_horaires||[];
  document.getElementById('cfg-pr').value=d.prod_ref||200;
  renderPwdList();
  renderModelList();
}

function renderPwdList() {
  const c=document.getElementById('pwd-list');
  if (!c) return;
  c.innerHTML=Object.entries(_cfgPwds).map(([nm,pw])=>`
    <div class="pr">
      <div class="pn">${esc(nm)}</div>
      <input type="password" id="pwi-${esc(nm)}" value="${esc(pw)}" data-n="${esc(nm)}">
      <button class="btn-eye" onclick="toggleEye('pwi-${esc(nm)}')">👁</button>
      <button class="btn btn-sec" style="font-size:10px;padding:3px 7px" onclick="rmPilot('${esc(nm)}')">✕</button>
    </div>`).join('')||'<div style="color:var(--c-muted);font-size:11px;padding:3px">Aucun pilote</div>';
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
  const pw=document.getElementById('adm-pw').value||'1234';
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,pilot_passwords:_cfgPwds})});
  const d=r?await r.json():{};
  d&&d.ok?toast('MDP enregistrés','ok'):toast(d&&d.error||'Erreur (vérifier MDP admin)','err');
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
        <button class="btn btn-danger" style="font-size:10px;padding:3px 7px" onclick="_cfgModels.splice(${mi},1);renderModelList()">✕</button>
      </div>
      <div class="day-grid">${DAYS.map(d=>{
        const dc=j[d.k]||{};
        return `<div class="day-box"><div class="dl">${d.l}</div>
          <input type="time" onchange="setDay(${mi},'${d.k}','debut',this.value)" value="${dc.debut||'05:00'}" style="margin-bottom:2px">
          <input type="time" onchange="setDay(${mi},'${d.k}','fin',this.value)" value="${dc.fin||'13:00'}">
        </div>`;
      }).join('')}</div>
    </div>`;
  }).join('')||'<div style="color:var(--c-muted);font-size:11px">Aucun modèle</div>';
}

function setDay(mi,day,field,val){
  if(!_cfgModels[mi]) return;
  if(!_cfgModels[mi].jours) _cfgModels[mi].jours={};
  if(!_cfgModels[mi].jours[day]) _cfgModels[mi].jours[day]={};
  _cfgModels[mi].jours[day][field]=val;
}
function addModel(){_cfgModels.push({nom:'Nouveau poste',jours:{lun:{debut:'05:00',fin:'13:00'},mar:{debut:'05:00',fin:'13:00'},mer:{debut:'05:00',fin:'13:00'},jeu:{debut:'05:00',fin:'13:00'},ven:{debut:'05:00',fin:'13:00'},sam:{debut:'',fin:''},dim:{debut:'',fin:''}}});renderModelList();}
async function saveModels(){
  const pw=document.getElementById('adm-pw2').value||'1234';
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,modeles_horaires:_cfgModels})});
  const d=r?await r.json():{};
  d&&d.ok?toast('Modèles enregistrés','ok'):toast(d&&d.error||'Erreur (vérifier MDP admin)','err');
}
async function saveProdRef(){
  const v=parseFloat(document.getElementById('cfg-pr').value)||200;
  const pw=document.getElementById('adm-pw3').value||'1234';
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,prod_ref:v})});
  const d=r?await r.json():{};
  d&&d.ok?toast('Référence enregistrée','ok'):toast(d&&d.error||'Erreur','err');
}

// ── Modals ──
function openM(id){const m=document.getElementById(id);if(m){m.classList.add('on');m.style.display='flex';}}
function closeM(id){const m=document.getElementById(id);if(m){m.classList.remove('on');m.style.display='';}}
document.addEventListener('click',e=>{if(e.target.classList.contains('overlay'))closeM(e.target.id);});

// ── Utils ──
async function apiFetch(url){
  try{const r=await fetch(url);return r.ok?await r.json():null;}catch(e){return null;}
}
function fmtDur(s){
  if(!s||s<0) return '00:00:00';
  const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sec=Math.floor(s%60);
  return [h,m,sec].map(x=>String(x).padStart(2,'0')).join(':');
}
function fmtD2(s){
  if(!s||s<0) return '0 min';
  const h=Math.floor(s/3600),m=Math.floor((s%3600)/60);
  return h?h+'h'+String(m).padStart(2,'0'):m+' min';
}
function fmtTRS(v){return(v===null||v===undefined||isNaN(v))?'--%':parseFloat(v).toFixed(1)+'%';}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}

function toast(msg,type){
  let t=document.getElementById('_toast');
  if(!t){t=document.createElement('div');t.id='_toast';t.style.cssText='position:fixed;bottom:18px;right:18px;padding:9px 16px;border-radius:7px;font-size:13px;font-weight:600;z-index:999;transition:opacity .3s;box-shadow:0 4px 14px rgba(0,0,0,.18)';document.body.appendChild(t);}
  t.textContent=msg;t.style.background=type==='ok'?'#16a34a':'#dc2626';t.style.color='#fff';t.style.opacity='1';
  clearTimeout(t._to);t._to=setTimeout(()=>t.style.opacity='0',3000);
}
</script>
</body>
</html>"""

def main():
    global cfg
    cfg = load_cfg()
    load_session()
    threading.Thread(target=load_lists, daemon=True).start()
    threading.Thread(target=load_history, daemon=True).start()

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
        webview.create_window(
            "KPI-ORC",
            "http://127.0.0.1:5001",
            width=1200, height=820,
            min_size=(900, 600),
            resizable=True,
        )
        webview.start()
    except ImportError:
        # Fallback: run as plain Flask server (dev mode)
        print("pywebview non disponible — démarrage en mode serveur sur http://127.0.0.1:5001")
        flask_app.run(host="127.0.0.1", port=5001, debug=True, use_reloader=False)


if __name__ == "__main__":
    main()
