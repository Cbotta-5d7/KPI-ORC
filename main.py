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
    if not pilot or not poste:
        return jsonify({"ok":False,"error":"Pilote et poste requis"}),400
    _S["pilot"] = pilot
    _S["poste"] = poste
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/logout', methods=['POST'])
def api_logout():
    if _S["prod_active"]:
        return jsonify({"ok":False,"error":"Production en cours"}),400
    _S["pilot"] = None
    _S["poste"] = None
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
    return jsonify({
        "pilot":pilot,"date":today,
        "nb_of":len(of_list),"trs":trs_poste,
        "tot_equiv":round(tot_eq,1),"tot_s":round(tot_s,0),
        "of_list":of_list,"of_count_shift":_S["of_count_shift"],
    })

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

# ── Dashboard HTML ─────────────────────────────────────────────────────────────
def generate_dashboard_html():
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path):
        return None, "Aucun fichier Excel configuré ou introuvable"
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return None, f"Erreur lecture Excel: {e}"

    today_str = datetime.date.today().strftime("%d/%m/%Y")
    prod_ref = get_prod_ref()
    from collections import defaultdict

    # Lire depuis la feuille unifiée Declarations
    decl_rows = []
    if "Declarations" in wb.sheetnames:
        ws = wb["Declarations"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and any(r):
                decl_rows.append(list(r)+[None]*5)
    # Rétro-compat: lire Data si Declarations absent
    elif "Data" in wb.sheetnames:
        ws = wb["Data"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and any(r):
                row = list(r)+[None]*10
                unified = [None]*37
                unified[0]="Production"; unified[1]=row[0]; unified[2]=row[1]; unified[3]=row[2]
                unified[4]=row[3]; unified[5]=row[4]; unified[6]=row[5]; unified[7]=row[6]
                unified[8]=row[7]; unified[9]=row[8]; unified[16]=row[17]; unified[17]=row[18]
                unified[18]=row[16]; unified[19]=row[13]; unified[20]=row[14]; unified[21]=row[15]
                unified[22]=row[19]; unified[23]=row[20]; unified[35]=row[57] if len(row)>57 else None
                decl_rows.append(unified)
    wb.close()

    prod_rows_all = [r for r in decl_rows if str(r[0] or "").strip().lower() in ("production","prod","")]
    evt_rows_all  = [r for r in decl_rows if str(r[0] or "").strip().lower() not in ("production","prod","")]

    # Calcul TRS global d'une liste de lignes prod
    def calc_trs(rows):
        eq = sum(float(str(r[21] or 0).replace(",",".") or 0) for r in rows)
        s  = sum(_hms_to_sec(str(r[18] or "00:00:00")) for r in rows)
        if prod_ref>0 and s>0: return round(eq/(prod_ref*s/28800)*100,1)
        return -1

    # TRS couleur
    def trs_color(t):
        if t<0: return "#94a3b8"
        if t>=70: return "#1a8c4e"
        if t>=50: return "#d97706"
        return "#e31e24"

    # Grouper par (poste, date)
    sessions = defaultdict(list)
    for r in prod_rows_all:
        key = (str(r[3] or ""), _row_date(r[2]))
        sessions[key].append(r)

    sess_list = sorted(sessions.items(), key=lambda x: (x[0][1],x[0][0]), reverse=True)[:6]

    # KPI cards (3 dernières séances)
    poste_data = []
    for (poste,date),rows in sess_list[:3]:
        trs  = calc_trs(rows)
        tot_eq = sum(float(str(r[21] or 0).replace(",",".") or 0) for r in rows)
        of_count = len(rows)
        pilots = list({str(r[4] or "") for r in rows if r[4]})
        poste_data.append({"poste":poste,"date":date,"trs":trs,"of_count":of_count,"equiv":round(tot_eq,1),"pilots":", ".join(pilots)})

    gauge_js_list = []
    cards_html = ""
    for i,pd in enumerate(poste_data):
        t   = pd["trs"]
        col = trs_color(t)
        trs_lbl = f"{t:.1f}%" if t>=0 else "—"
        raw  = max(0, t if t>=0 else 0)
        rest = max(0, 100-raw)
        cards_html += f"""
        <div class="kpi-card">
          <div class="kpi-card-hdr">{pd['poste']} — {pd['date']}</div>
          <canvas id="gauge-{i}" width="160" height="100"></canvas>
          <div class="kpi-trs" style="color:{col}">{trs_lbl}</div>
          <div class="kpi-sub">{pd['of_count']} OF &nbsp;|&nbsp; Equiv: {pd['equiv']}</div>
          <div class="kpi-sub" style="margin-top:3px">{pd['pilots']}</div>
        </div>"""
        gauge_js_list.append(f"""
        new Chart(document.getElementById('gauge-{i}'),{{
          type:'doughnut',
          data:{{datasets:[{{data:[{raw},{rest}],backgroundColor:['{col}','#dde4ef'],borderWidth:0}}]}},
          options:{{circumference:180,rotation:-90,cutout:'70%',plugins:{{legend:{{display:false}},tooltip:{{enabled:false}}}}}}
        }});""")

    # Pareto arrêts
    evt_durations = defaultdict(float)
    for r in evt_rows_all:
        t = str(r[0] or "")
        if not t or t.lower() in ("pause pilote","changement d'of"): continue
        dur = _hms_to_sec(str(r[18] or "00:00:00"))
        evt_durations[t] += dur
    pareto = sorted(evt_durations.items(), key=lambda x:-x[1])[:14]
    pareto_labels = json.dumps([p[0][:25] for p in pareto])
    pareto_vals   = json.dumps([round(p[1]/60,1) for p in pareto])
    def _pcol(lbl):
        if "PB" in lbl or "Technique" in lbl: return "#e31e24"
        if "Ratt" in lbl: return "#d97706"
        if "Nettoyage" in lbl: return "#0891b2"
        return "#7c3aed"
    pareto_colors = json.dumps([_pcol(p[0]) for p in pareto])

    # TRS par pilote (semaine courante)
    pilot_sessions = defaultdict(lambda: {"eq":0.0,"s":0.0,"nb":0})
    for r in prod_rows_all:
        p = str(r[4] or "")
        if not p: continue
        pilot_sessions[p]["eq"] += float(str(r[21] or 0).replace(",",".") or 0)
        pilot_sessions[p]["s"]  += _hms_to_sec(str(r[18] or "00:00:00"))
        pilot_sessions[p]["nb"] += 1
    pilot_trs_labels = []; pilot_trs_vals = []; pilot_trs_colors = []
    for pname, d in sorted(pilot_sessions.items(), key=lambda x:-x[1]["nb"])[:10]:
        t = calc_trs([]) if d["s"]==0 else round(d["eq"]/(prod_ref*d["s"]/28800)*100,1) if prod_ref>0 else -1
        pilot_trs_labels.append(pname)
        pilot_trs_vals.append(max(0,t) if t>=0 else 0)
        pilot_trs_colors.append(trs_color(t))
    pilot_trs_labels_j = json.dumps(pilot_trs_labels)
    pilot_trs_vals_j   = json.dumps(pilot_trs_vals)
    pilot_trs_colors_j = json.dumps(pilot_trs_colors)

    # Table derniers événements (50)
    evt_table_rows = ""
    for r in list(reversed(evt_rows_all))[:50]:
        typ = str(r[0] or "")
        if "PB" in typ or "Technique" in typ: col_cls="badge-red"
        elif "Ratt" in typ: col_cls="badge-amber"
        elif "Nettoyage" in typ: col_cls="badge-cyan"
        else: col_cls="badge-navy"
        evt_table_rows += f"""<tr>
          <td><span class='badge {col_cls}'>{typ}</span></td>
          <td>{r[1] or ''}</td><td>{_row_date(r[2])}</td>
          <td>{r[4] or ''}</td><td>{str(r[16] or '')[:8]}</td>
          <td>{str(r[17] or '')[:8]}</td><td>{r[18] or ''}</td>
          <td style='text-align:left;max-width:200px;overflow:hidden;text-overflow:ellipsis'>{r[35] or ''}</td>
        </tr>"""

    # Table dernières productions (30)
    prod_table_rows = ""
    for r in list(reversed(prod_rows_all))[:30]:
        trs_v = str(r[24] or "")
        try: tv = float(trs_v.replace(",","."))
        except: tv = -1
        trs_cls = "color:#1a8c4e;font-weight:800" if tv>=70 else "color:#d97706;font-weight:800" if tv>=50 else "color:#e31e24;font-weight:800" if tv>=0 else "color:#94a3b8"
        prod_table_rows += f"""<tr>
          <td>{str(r[1] or '')}</td><td>{_row_date(r[2])}</td>
          <td>{r[3] or ''}</td><td>{r[4] or ''}</td>
          <td>{str(r[16] or '')[:5]}</td><td>{str(r[17] or '')[:5]}</td>
          <td>{r[7] or ''}</td><td>{r[19] or ''}</td><td>{r[20] or ''}</td><td>{r[21] or ''}</td>
          <td style="{trs_cls}">{trs_v or '—'}{'%' if trs_v else ''}</td>
        </tr>"""

    # Supervision live
    session_active = _S.get("prod_active", False)
    active_stops = [k for k,t in _S.get("timers",{}).items() if t.get("running") and not k.startswith("_")]
    pilot_now  = _S.get("pilot","—") or "—"
    poste_now  = _S.get("poste","—") or "—"
    of_num_now = (_S.get("form") or {}).get("of_num","") or "—"
    of_start_str = ""
    if _S.get("of_start"):
        try: of_start_str = _S["of_start"].strftime("%H:%M:%S")
        except: pass

    stops_html = ""
    for k in active_stops:
        ev = next((e for e in EVENTS if e[1]==k), None)
        lbl = ev[0] if ev else k
        t = _S.get("timers",{}).get(k,{})
        elapsed = t.get("elapsed",0)
        if t.get("running") and t.get("start"):
            try: elapsed += (datetime.datetime.now()-t["start"]).total_seconds()
            except: pass
        stops_html += f"""<div class='stop-live-card'>
          <div class='slc-lbl'>{lbl}</div>
          <div class='slc-timer'>{fmt(elapsed)}</div>
        </div>"""

    gen_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    gauge_js = "\n".join(gauge_js_list)
    sup_bg = "#e31e24" if session_active else "#1a8c4e"
    sup_status = "● PRODUCTION EN COURS" if session_active else "○ Aucune production active"

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="60">
<title>KPI Dashboard — ORC1</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:#f0f4fb;color:#0f172a}}
.page-hdr{{background:#1a1f5e;color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}}
.page-hdr h1{{font-size:18px;font-weight:800;letter-spacing:.3px}}
.page-hdr .gen-time{{font-size:12px;opacity:.7}}
.section{{padding:18px 24px}}
.section h2{{font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:#1a1f5e;margin-bottom:14px;display:flex;align-items:center;gap:8px}}
.section h2::after{{content:'';flex:1;height:2px;background:#dde4ef}}
.kpi-row{{display:flex;gap:14px;flex-wrap:wrap}}
.kpi-card{{background:#fff;border-radius:14px;box-shadow:0 2px 10px rgba(0,0,0,.08);padding:16px 20px;min-width:200px;text-align:center;flex:1}}
.kpi-card-hdr{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#64748b;margin-bottom:8px}}
.kpi-trs{{font-size:30px;font-weight:900;margin:4px 0}}
.kpi-sub{{font-size:11px;color:#64748b;margin-top:2px}}
.charts-row{{display:grid;grid-template-columns:2fr 1fr;gap:14px;margin-bottom:14px}}
.charts-row-2{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.chart-card{{background:#fff;border-radius:14px;box-shadow:0 2px 10px rgba(0,0,0,.08);padding:16px}}
.chart-card h3{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#64748b;margin-bottom:12px}}
table{{width:100%;border-collapse:collapse;background:#fff;font-size:11px}}
th{{background:#dde4ef;font-weight:700;padding:8px 10px;text-align:center;white-space:nowrap;border-bottom:2px solid #c0cde0}}
td{{padding:6px 10px;text-align:center;border-bottom:1px solid #edf0f7;white-space:nowrap}}
tr:hover td{{background:#f0f4fb}}
.table-wrap{{border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.07)}}
.badge{{display:inline-block;padding:2px 7px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap}}
.badge-red{{background:#fee2e2;color:#991b1b}}
.badge-amber{{background:#fef3c7;color:#92400e}}
.badge-navy{{background:#dbeafe;color:#1e40af}}
.badge-cyan{{background:#cffafe;color:#0e7490}}
.supervision-block{{background:#fff;border-radius:14px;box-shadow:0 2px 10px rgba(0,0,0,.08);overflow:hidden;margin:0 0 14px}}
.sup-hdr{{background:{sup_bg};color:#fff;padding:14px 20px;font-weight:800;font-size:15px;display:flex;align-items:center;justify-content:space-between}}
.sup-hdr .sup-meta{{font-size:13px;opacity:.85;font-weight:600}}
.sup-body{{padding:16px 20px;display:flex;gap:16px;flex-wrap:wrap;align-items:flex-start}}
.sup-info-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;flex:1}}
.sup-cell{{background:#f8fafc;border-radius:8px;padding:10px 14px;text-align:center}}
.sup-cell-lbl{{font-size:10px;font-weight:700;text-transform:uppercase;color:#64748b}}
.sup-cell-val{{font-size:18px;font-weight:800;margin-top:2px;color:#0f172a}}
.stops-live{{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px}}
.stop-live-card{{background:#fee2e2;border:2px solid #e31e24;border-radius:10px;padding:10px 16px;text-align:center}}
.slc-lbl{{font-size:12px;font-weight:700;color:#991b1b}}
.slc-timer{{font-size:20px;font-weight:900;color:#e31e24;font-variant-numeric:tabular-nums}}
.no-stops{{color:#1a8c4e;font-weight:700;padding:8px 0}}
{'div.alert-banner{background:#e31e24;color:#fff;padding:10px 24px;font-weight:700;font-size:14px;text-align:center;animation:blink 1.2s infinite}' if active_stops else ''}
@keyframes blink{{0%,100%{{opacity:1}}50%{{opacity:.75}}}}
.footer{{padding:14px 24px;color:#94a3b8;font-size:11px;text-align:center;border-top:1px solid #e2e8f0;margin-top:8px}}
</style>
</head>
<body>
<div class="page-hdr">
  <h1>&#127981; KPI Dashboard — ORC1</h1>
  <div class="gen-time">Généré le {gen_time} &nbsp;|&nbsp; Actualisation auto. 60s</div>
</div>
{'<div class="alert-banner">&#9888; ' + str(len(active_stops)) + " arrêt(s) en cours — Poste " + poste_now + "</div>" if active_stops else ""}

<!-- Supervision live -->
<div class="section">
  <h2>&#128308; Supervision live</h2>
  <div class="supervision-block">
    <div class="sup-hdr">
      <span>{sup_status}</span>
      {f'<span class="sup-meta">Pilote: {pilot_now} | Poste: {poste_now} | OF: {of_num_now} | Début: {of_start_str}</span>' if session_active else ''}
    </div>
    <div class="sup-body">
      <div style="flex:1">
        <div class="sup-info-grid">
          <div class="sup-cell"><div class="sup-cell-lbl">Pilote</div><div class="sup-cell-val">{pilot_now}</div></div>
          <div class="sup-cell"><div class="sup-cell-lbl">Poste</div><div class="sup-cell-val">{poste_now}</div></div>
          <div class="sup-cell"><div class="sup-cell-lbl">OF</div><div class="sup-cell-val">{of_num_now}</div></div>
          <div class="sup-cell"><div class="sup-cell-lbl">Début</div><div class="sup-cell-val">{of_start_str or "—"}</div></div>
        </div>
        <div style="margin-top:12px">
          {"<div style='font-weight:700;color:#991b1b;margin-bottom:8px;font-size:12px'>Arrêts actifs :</div><div class='stops-live'>" + stops_html + "</div>" if active_stops else "<div class='no-stops'>&#10004; Aucun arrêt actif</div>"}
        </div>
      </div>
    </div>
  </div>
</div>

<!-- KPI TRS par séance -->
<div class="section">
  <h2>&#128202; TRS par séance (dernières sessions)</h2>
  <div class="kpi-row">
    {cards_html if cards_html else "<div style='color:#94a3b8;padding:20px'>Aucune donnée</div>"}
  </div>
</div>

<!-- Pareto + TRS pilotes -->
<div class="section">
  <h2>&#128200; Analyse des arrêts &amp; TRS pilotes</h2>
  <div class="charts-row">
    <div class="chart-card">
      <h3>Pareto des arrêts — Temps total (min)</h3>
      <canvas id="pareto-chart" height="220"></canvas>
    </div>
    <div class="chart-card">
      <h3>Répartition par catégorie</h3>
      <canvas id="pie-chart" height="220"></canvas>
    </div>
  </div>
  <div class="chart-card">
    <h3>TRS par pilote (toutes sessions)</h3>
    <canvas id="pilot-chart" height="100"></canvas>
  </div>
</div>

<!-- Dernières productions -->
<div class="section">
  <h2>&#128203; Dernières productions</h2>
  <div class="table-wrap">
    <table>
      <thead><tr><th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th><th>Début</th><th>Fin</th><th>Taille</th><th>Qté Fab</th><th>Qté Emb</th><th>Equiv</th><th>TRS%</th></tr></thead>
      <tbody>{prod_table_rows if prod_table_rows else "<tr><td colspan='11' style='color:#94a3b8;padding:20px'>Aucune production</td></tr>"}</tbody>
    </table>
  </div>
</div>

<!-- Derniers événements -->
<div class="section">
  <h2>&#128221; Derniers événements / arrêts</h2>
  <div class="table-wrap">
    <table>
      <thead><tr><th>Type</th><th>OF</th><th>Date</th><th>Pilote</th><th>Début</th><th>Fin</th><th>Durée</th><th>Commentaire</th></tr></thead>
      <tbody>{evt_table_rows if evt_table_rows else "<tr><td colspan='8' style='color:#94a3b8;padding:20px'>Aucun événement</td></tr>"}</tbody>
    </table>
  </div>
</div>

<div class="footer">KPI-ORC v6.3 — Généré le {gen_time} — Actualisation toutes les 60 secondes</div>

<script>
{gauge_js}

new Chart(document.getElementById('pareto-chart'),{{
  type:'bar',
  data:{{
    labels:{pareto_labels},
    datasets:[{{label:'Temps (min)',data:{pareto_vals},backgroundColor:{pareto_colors},borderRadius:4}}]
  }},
  options:{{
    indexAxis:'y',
    plugins:{{legend:{{display:false}}}},
    scales:{{x:{{grid:{{color:'#f0f4fb'}},ticks:{{font:{{size:10}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:11}}}}}}}}
  }}
}});

new Chart(document.getElementById('pie-chart'),{{
  type:'doughnut',
  data:{{
    labels:{pareto_labels},
    datasets:[{{data:{pareto_vals},backgroundColor:{pareto_colors},borderWidth:2,borderColor:'#fff'}}]
  }},
  options:{{plugins:{{legend:{{position:'right',labels:{{font:{{size:10}},boxWidth:10,padding:8}}}}}}}}
}});

new Chart(document.getElementById('pilot-chart'),{{
  type:'bar',
  data:{{
    labels:{pilot_trs_labels_j},
    datasets:[{{label:'TRS%',data:{pilot_trs_vals_j},backgroundColor:{pilot_trs_colors_j},borderRadius:4}}]
  }},
  options:{{
    plugins:{{legend:{{display:false}}}},
    scales:{{
      y:{{min:0,max:110,grid:{{color:'#f0f4fb'}},ticks:{{callback:v=>v+'%',font:{{size:10}}}}}},
      x:{{grid:{{display:false}},ticks:{{font:{{size:11}}}}}}
    }}
  }}
}});
</script>
</body>
</html>"""
    try:
        html_path = os.path.join(os.path.dirname(path), "KPI_Dashboard.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        return html_path, None
    except Exception as e:
        return None, str(e)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>KPI-ORC | ORC1</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --navy:#1a1f5e;--navy-l:#2d3490;--red:#e31e24;--green:#1a8c4e;
  --amber:#d97706;--cyan:#0891b2;--purple:#7c3aed;
  --bg:#f0f4fb;--white:#ffffff;--gray:#64748b;--lgray:#dde4ef;--dark:#0f172a;
  --shadow:0 2px 8px rgba(0,0,0,.10);--shadow-lg:0 4px 20px rgba(0,0,0,.14);
}
/* Dark theme quand arrêt actif */
body.stop-active{
  --bg:#0f172a;--white:#1e293b;--lgray:#334155;--gray:#94a3b8;--dark:#f1f5f9;
}
body{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--dark);overflow:hidden;height:100vh;transition:background .3s}
.view{display:none;height:100vh;flex-direction:column;overflow:hidden}
.view.active{display:flex}

/* Header */
.hdr{background:var(--white);border-bottom:2px solid var(--lgray);display:flex;align-items:center;padding:0 12px;height:52px;flex-shrink:0;box-shadow:var(--shadow);transition:background .3s}
body.stop-active .hdr{background:#1e293b;border-bottom-color:#334155}
.hdr-logo{font-size:15px;font-weight:800;color:var(--navy);letter-spacing:.5px;margin-right:10px}
body.stop-active .hdr-logo{color:#93c5fd}
.hdr-accent{width:4px;height:32px;border-radius:2px;background:var(--green);margin-right:10px;flex-shrink:0}
.hdr-accent.red{background:var(--red)}
.hdr-right{margin-left:auto;display:flex;gap:6px;align-items:center}
.hdr-pilot{font-size:12px;color:var(--gray);white-space:nowrap}
.hdr-pilot strong{color:var(--dark)}

/* Tab nav */
.hdr-tabs{display:flex;gap:2px;align-items:center;margin:0 8px}
.hdr-tab{padding:6px 14px;font-size:12px;font-weight:700;cursor:pointer;border-radius:6px;color:var(--gray);transition:all .15s;border:none;background:none}
.hdr-tab.active{background:var(--navy);color:#fff}
.hdr-tab:not(.active):hover{background:var(--lgray);color:var(--dark)}
.hdr-tab .tab-dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:4px;background:var(--red);animation:pulse 1.2s infinite}

/* Buttons */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:7px 14px;border:none;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer;transition:all .15s;white-space:nowrap}
.btn:active{transform:scale(.97)}
.btn-navy{background:var(--navy);color:#fff}.btn-navy:hover{background:var(--navy-l)}
.btn-green{background:var(--green);color:#fff}.btn-green:hover{filter:brightness(1.1)}
.btn-red{background:var(--red);color:#fff}.btn-red:hover{filter:brightness(1.1)}
.btn-amber{background:var(--amber);color:#fff}.btn-amber:hover{filter:brightness(1.1)}
.btn-cyan{background:var(--cyan);color:#fff}.btn-cyan:hover{filter:brightness(1.1)}
.btn-ghost{background:var(--lgray);color:var(--dark)}.btn-ghost:hover{background:#c8d4e8}
.btn-lg{padding:12px 20px;font-size:14px;border-radius:10px}
.btn-xl{padding:18px 28px;font-size:16px;border-radius:12px;width:100%}
.btn-sm{padding:4px 10px;font-size:11px;border-radius:6px}
.btn-icon{background:none;border:none;cursor:pointer;padding:4px 6px;border-radius:6px;font-size:15px;line-height:1;transition:background .15s}
.btn-icon:hover{background:var(--lgray)}
.btn-icon.edit{color:#2563eb}
.btn-icon.del{color:#e31e24}

/* Cards */
.card{background:var(--white);border-radius:12px;box-shadow:var(--shadow);padding:14px;transition:background .3s}
.card-hdr{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--gray);margin-bottom:10px}

/* Status bar */
.status-bar{display:flex;gap:8px;padding:6px 12px;flex-shrink:0}
.status-cell{background:var(--white);border-radius:10px;padding:8px 16px;text-align:center;box-shadow:var(--shadow);flex:1;transition:background .3s}
.status-cell.running{background:var(--red);color:#fff}
.status-cell.ok{background:var(--green);color:#fff}
.status-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.75}
.status-value{font-size:22px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.1}
.status-sub{font-size:10px;opacity:.7}

/* Production layout */
.prod-body{display:grid;grid-template-columns:240px 1fr 220px;gap:8px;padding:8px 12px;flex:1;min-height:0;overflow:hidden}
.prod-col{min-height:0;overflow:hidden}

/* Action col */
.action-col{display:flex;flex-direction:column;gap:8px}
.big-stop-btn{
  background:var(--red);color:#fff;border:none;border-radius:14px;
  padding:0;cursor:pointer;font-weight:900;font-size:17px;
  box-shadow:0 5px 0 #9b0f14;transition:all .12s;
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  height:100px;width:100%;gap:4px
}
body.stop-active .big-stop-btn{
  height:120px;font-size:20px;animation:pulse 1.2s infinite;
  box-shadow:0 5px 0 #9b0f14,0 0 30px rgba(227,30,36,.5)
}
.big-stop-btn:active{transform:translateY(4px);box-shadow:0 1px 0 #9b0f14}
.big-stop-btn .stop-sub{font-size:11px;font-weight:600;opacity:.85}
.action-btn{
  display:flex;align-items:center;justify-content:center;gap:8px;
  border:none;border-radius:12px;padding:14px;cursor:pointer;
  font-size:14px;font-weight:700;color:#fff;width:100%;
  box-shadow:0 3px 0 rgba(0,0,0,.25);transition:all .12s
}
.action-btn:active{transform:translateY(2px);box-shadow:0 1px 0 rgba(0,0,0,.25)}
.action-btn.nett{background:var(--cyan)}
.action-btn.pause{background:var(--navy)}
.action-btn.end{background:var(--dark)}
.action-spacer{flex:1}

/* Form col */
.form-col{display:flex;flex-direction:column;overflow:hidden}
.form-toggle{
  background:var(--white);border:2px solid var(--lgray);border-radius:10px;
  padding:10px 14px;cursor:pointer;display:flex;align-items:center;justify-content:space-between;
  font-size:12px;font-weight:700;color:var(--dark);flex-shrink:0;margin-bottom:6px;
  transition:border-color .15s
}
.form-toggle:hover{border-color:var(--navy-l)}
.form-toggle .toggle-arrow{transition:transform .2s;font-size:14px}
.form-toggle.open .toggle-arrow{transform:rotate(180deg)}
.form-body{background:var(--white);border-radius:10px;padding:12px;overflow-y:auto;flex:1;display:none}
.form-body.visible{display:block}
.form-essential{background:var(--white);border-radius:10px;padding:12px;flex-shrink:0}
/* Pilote + OF très grands */
.pilot-of-banner{background:var(--white);border-radius:10px;padding:10px 14px;flex-shrink:0;margin-bottom:6px;display:flex;gap:16px;align-items:center}
body.stop-active .pilot-of-banner{background:#1e293b}
.pob-item{flex:1;text-align:center}
.pob-label{font-size:10px;font-weight:700;text-transform:uppercase;color:var(--gray);letter-spacing:.5px}
.pob-value{font-size:28px;font-weight:900;color:var(--navy);letter-spacing:1px;line-height:1.1}
body.stop-active .pob-value{color:#93c5fd}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.form-group{display:flex;flex-direction:column;gap:2px}
.form-group.full{grid-column:1/-1}
.form-group label{font-size:10px;font-weight:700;color:var(--gray);text-transform:uppercase;letter-spacing:.4px}
.form-group input,.form-group select,.form-group textarea{
  border:1.5px solid var(--lgray);border-radius:6px;padding:6px 8px;
  font-size:12px;color:var(--dark);background:var(--white);outline:none;
  transition:border-color .15s}
.form-group input:focus,.form-group select:focus{border-color:var(--navy-l)}
.form-group input[disabled],.form-group select[disabled]{background:#f0f4fb;color:var(--gray)}
.form-group textarea{resize:none;height:48px;font-size:12px}
.checkbox-row{display:flex;align-items:center;gap:6px;padding:6px 0}
.checkbox-row input[type=checkbox]{width:16px;height:16px;cursor:pointer}

/* Recap col */
.recap-col{display:flex;flex-direction:column;gap:6px;overflow:hidden}
.recap-stop-row{display:flex;align-items:center;gap:6px;padding:5px 7px;border-radius:6px;margin-bottom:3px;background:#f8fafc;cursor:pointer;transition:background .15s}
.recap-stop-row:hover{background:#e2e8f0}
body.stop-active .recap-stop-row{background:#1e293b}
body.stop-active .recap-stop-row:hover{background:#334155}
.recap-stop-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.recap-stop-name{font-size:11px;flex:1;font-weight:600}
.recap-stop-time{font-size:11px;color:var(--gray);font-variant-numeric:tabular-nums}
.recap-stop-row.running{background:#fff0f0;animation:pulse 1.5s infinite}
body.stop-active .recap-stop-row.running{background:#3f1212}

/* Active stops bottom — beaucoup plus visibles */
.active-stops-bottom{
  background:#0f172a;border-top:4px solid var(--red);
  padding:10px 12px;flex-shrink:0;
  display:none;align-items:stretch;gap:10px;flex-wrap:nowrap;overflow-x:auto
}
body.stop-active .active-stops-bottom{border-top-color:#ff0000}
.active-stops-bottom.has-stops{display:flex}
.active-stop-card{
  background:var(--red);color:#fff;border-radius:12px;
  padding:12px 20px;display:flex;flex-direction:column;align-items:center;
  justify-content:center;cursor:pointer;min-width:170px;flex-shrink:0;
  box-shadow:0 4px 0 #9b0f14,0 0 20px rgba(227,30,36,.4);transition:all .12s;
  animation:pulse 1.4s infinite
}
.active-stop-card:active{transform:translateY(3px);box-shadow:0 1px 0 #9b0f14}
.active-stop-card.nett{background:var(--cyan);box-shadow:0 4px 0 #065981,0 0 16px rgba(8,145,178,.3);animation:none}
.active-stop-card .asc-label{font-size:13px;font-weight:800;letter-spacing:.3px}
.active-stop-card .asc-timer{font-size:26px;font-weight:900;font-variant-numeric:tabular-nums;letter-spacing:2px}
.active-stop-card .asc-hint{font-size:10px;opacity:.7;margin-top:3px}

/* TRS gauge */
.gauge-wrap{position:relative;width:140px;height:80px;margin:0 auto}
.gauge-svg{width:140px;height:80px}
.gauge-text{position:absolute;bottom:0;left:50%;transform:translateX(-50%);text-align:center}
.gauge-pct{font-size:24px;font-weight:800}
.gauge-lbl{font-size:10px;color:var(--gray);font-weight:700;text-transform:uppercase}

/* Tables */
.table-wrap{overflow-x:auto;border-radius:10px;box-shadow:var(--shadow)}
.ktable{width:100%;border-collapse:collapse;background:var(--white);font-size:11px;transition:background .3s}
.ktable th{background:var(--lgray);color:var(--dark);font-weight:700;padding:8px 8px;text-align:center;white-space:nowrap;border-bottom:2px solid #c0cde0}
.ktable td{padding:6px 8px;text-align:center;border-bottom:1px solid #edf0f7;white-space:nowrap}
.ktable tr:hover td{background:#f0f4fb}
body.stop-active .ktable tr:hover td{background:#1e293b}
.trs-hi{background:#f0fdf4}.trs-hi .trs-cell{color:var(--green);font-weight:800}
.trs-warn{background:#fffbeb}.trs-warn .trs-cell{color:var(--amber);font-weight:800}
.trs-low{background:#fef2f2}.trs-low .trs-cell{color:var(--red);font-weight:800}

/* Modal */
.modal{display:none;position:fixed;inset:0;z-index:100;background:rgba(15,23,42,.6);align-items:center;justify-content:center}
.modal.active{display:flex}
.modal-box{background:var(--white);border-radius:16px;box-shadow:var(--shadow-lg);width:min(700px,95vw);max-height:92vh;overflow-y:auto;display:flex;flex-direction:column}
.modal-box.narrow{width:min(480px,95vw)}
.modal-box.wide{width:min(900px,98vw)}
.modal-hdr{background:var(--navy);color:#fff;padding:16px 22px;border-radius:16px 16px 0 0;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.modal-hdr h2{font-size:15px;font-weight:800}
.modal-hdr.green{background:var(--green)}
.modal-hdr.red{background:var(--red)}
.modal-hdr.amber{background:var(--amber)}
.modal-body{padding:18px 22px;flex:1}
.modal-footer{padding:14px 22px;border-top:1px solid var(--lgray);display:flex;gap:8px;justify-content:flex-end;flex-shrink:0}

/* Settings tabs */
.set-tabs{display:flex;gap:2px;border-bottom:2px solid var(--lgray);margin-bottom:16px}
.set-tab{padding:8px 18px;font-size:12px;font-weight:700;cursor:pointer;border-radius:8px 8px 0 0;color:var(--gray);transition:all .15s;border:none;background:none}
.set-tab.active{background:var(--navy);color:#fff}
.set-panel{display:none}.set-panel.active{display:block}

/* Stop selector */
.stops-grid{display:grid;gap:6px;margin-bottom:10px}
.stops-grid.ratt{grid-template-columns:repeat(3,1fr)}
.stops-grid.pb{grid-template-columns:repeat(4,1fr)}
.stops-grid.special{grid-template-columns:repeat(2,1fr)}
.stop-btn{
  border:none;border-radius:10px;padding:10px 6px;cursor:pointer;
  font-size:11px;font-weight:700;color:#fff;text-align:center;
  transition:all .15s;position:relative;box-shadow:0 3px 0 rgba(0,0,0,.25)}
.stop-btn:active{transform:translateY(2px);box-shadow:0 1px 0 rgba(0,0,0,.25)}
.stop-btn.ratt{background:var(--amber)}
.stop-btn.pb{background:var(--navy)}
.stop-btn.running{animation:pulse 1.2s infinite}
.stop-btn.running::after{content:'● EN COURS';display:block;font-size:8px;margin-top:3px;opacity:.85}
.custom-stop-row{display:flex;gap:6px;margin-top:6px}
.custom-stop-row input{flex:1;border:2px solid var(--lgray);border-radius:8px;padding:8px 12px;font-size:12px;outline:none}
.custom-stop-row input:focus{border-color:var(--navy)}

/* Modèles horaires */
.modele-row{display:grid;grid-template-columns:1fr 1fr 1fr 1fr 40px;gap:8px;align-items:center;margin-bottom:8px}
.modele-row input{border:1.5px solid var(--lgray);border-radius:6px;padding:6px 8px;font-size:12px;width:100%}

/* Fin poste table */
.fp-table{width:100%;border-collapse:collapse;font-size:12px}
.fp-table th{background:var(--lgray);padding:8px;font-weight:700;text-align:center}
.fp-table td{padding:7px 8px;border-bottom:1px solid #edf0f7;text-align:center}

/* Edit row modal fields */
.edit-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.edit-group{display:flex;flex-direction:column;gap:3px}
.edit-group.full{grid-column:1/-1}
.edit-group label{font-size:10px;font-weight:700;color:var(--gray);text-transform:uppercase;letter-spacing:.4px}
.edit-group input,.edit-group select,.edit-group textarea{
  border:1.5px solid var(--lgray);border-radius:6px;padding:6px 8px;
  font-size:12px;color:var(--dark);background:var(--white);outline:none}
.edit-group input:focus{border-color:var(--navy)}

@keyframes pulse{0%,100%{opacity:1}50%{opacity:.65}}
@keyframes pulseScale{0%,100%{transform:scale(1)}50%{transform:scale(1.03)}}
</style>
</head>
<body>

<!-- ══ VIEW: LOGIN ══════════════════════════════════════════════════════════ -->
<div class="view active" id="view-login">
  <div style="flex:1;display:flex;align-items:center;justify-content:center;background:var(--bg)">
    <div style="width:340px">
      <div style="text-align:center;margin-bottom:30px">
        <div style="font-size:36px">🏭</div>
        <div style="font-size:22px;font-weight:900;color:var(--navy);margin-top:8px">KPI-ORC</div>
        <div style="font-size:13px;color:var(--gray);margin-top:4px">ORC1 — Gestion de production</div>
      </div>
      <div class="card" style="padding:24px">
        <div class="form-group" style="margin-bottom:14px">
          <label>Pilote</label>
          <select id="login-pilot" style="height:38px;font-size:14px">
            <option value="">— Sélectionner —</option>
          </select>
        </div>
        <div class="form-group" style="margin-bottom:14px">
          <label>Poste</label>
          <select id="login-poste" style="height:38px;font-size:14px">
            <option value="">— Sélectionner —</option>
            <option>Matin</option><option>Midi</option><option>Nuit</option><option>Jour</option>
          </select>
        </div>
        <button class="btn btn-navy btn-xl" onclick="doLogin()" style="margin-top:6px">
          Connexion
        </button>
        <div id="login-err" style="color:var(--red);font-size:12px;margin-top:8px;text-align:center"></div>
      </div>
      <div style="text-align:center;margin-top:14px">
        <button class="btn btn-ghost btn-sm" onclick="openSettings()">⚙ Paramètres</button>
      </div>
    </div>
  </div>
</div>

<!-- ══ VIEW: MAIN ═══════════════════════════════════════════════════════════ -->
<div class="view" id="view-main">
  <div class="hdr">
    <div class="hdr-accent" id="hdr-accent-main"></div>
    <div class="hdr-logo">KPI-ORC</div>
    <div class="hdr-tabs">
      <button class="hdr-tab active" id="main-tab-decl" onclick="switchMainTab('decl')">Déclarations</button>
      <button class="hdr-tab" id="main-tab-evt" onclick="switchMainTab('evt')">Événements</button>
    </div>
    <div class="hdr-right">
      <span class="hdr-pilot">Pilote: <strong id="main-pilot-name">—</strong> | Poste: <strong id="main-poste-name">—</strong></span>
      <button class="btn btn-green btn-sm" onclick="openStartProd()">▶ Démarrer prod</button>
      <button class="btn btn-ghost btn-sm" onclick="openFinPoste()">🏁 Fin de poste</button>
      <button class="btn btn-ghost btn-sm" onclick="openSettings()">⚙</button>
      <button class="btn btn-ghost btn-sm" onclick="doLogout()">Déconnexion</button>
    </div>
  </div>

  <!-- Déclarations tab -->
  <div id="main-panel-decl" style="flex:1;overflow:auto;padding:12px">
    <div class="table-wrap">
      <table class="ktable" id="hist-table">
        <thead><tr>
          <th>OF</th><th>Date</th><th>Poste</th><th>Pilote</th>
          <th>Début</th><th>Fin</th><th>Taille</th><th>Qté Fab</th>
          <th>Qté Emb</th><th>Equiv</th><th>TRS%</th><th>Actions</th>
        </tr></thead>
        <tbody id="hist-body"></tbody>
      </table>
    </div>
  </div>

  <!-- Événements tab -->
  <div id="main-panel-evt" style="flex:1;overflow:auto;padding:12px;display:none">
    <div class="table-wrap">
      <table class="ktable" id="evt-table">
        <thead><tr>
          <th>Type</th><th>OF</th><th>Date</th><th>Pilote</th>
          <th>Début</th><th>Fin</th><th>Durée</th><th>Commentaire</th><th>Actions</th>
        </tr></thead>
        <tbody id="evt-body"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ══ VIEW: PRODUCTION ════════════════════════════════════════════════════ -->
<div class="view" id="view-prod">
  <div class="hdr">
    <div class="hdr-accent red" id="hdr-accent-prod"></div>
    <div class="hdr-logo">KPI-ORC</div>
    <div class="hdr-tabs">
      <button class="hdr-tab" onclick="showView('main')">← Déclarations</button>
      <button class="hdr-tab active"><span class="tab-dot"></span>Prod en cours</button>
    </div>
    <div class="hdr-right">
      <span class="hdr-pilot">Pilote: <strong id="prod-hdr-pilot">—</strong></span>
      <button class="btn btn-ghost btn-sm" onclick="openSettings()">⚙</button>
    </div>
  </div>

  <!-- Pilot / OF banner -->
  <div class="pilot-of-banner">
    <div class="pob-item">
      <div class="pob-label">Pilote</div>
      <div class="pob-value" id="pob-pilot">—</div>
    </div>
    <div class="pob-item">
      <div class="pob-label">OF</div>
      <div class="pob-value" id="pob-of">—</div>
    </div>
    <div class="pob-item">
      <div class="pob-label">Poste</div>
      <div class="pob-value" id="pob-poste" style="font-size:22px">—</div>
    </div>
  </div>

  <!-- Status bar -->
  <div class="status-bar">
    <div class="status-cell running" id="sc-of">
      <div class="status-label">Durée OF</div>
      <div class="status-value" id="sc-of-val">00:00:00</div>
    </div>
    <div class="status-cell" id="sc-stop">
      <div class="status-label">⛔ Arrêts</div>
      <div class="status-value" id="sc-stop-val">00:00:00</div>
    </div>
    <div class="status-cell" id="sc-trs">
      <div class="status-label">TRS%</div>
      <div class="status-value" id="sc-trs-val">—</div>
    </div>
  </div>

  <!-- Main 3-col layout -->
  <div class="prod-body">
    <!-- LEFT: actions -->
    <div class="prod-col action-col">
      <button class="big-stop-btn" onclick="openStopModal()">
        <span>⛔</span>
        <span>Déclarer un arrêt</span>
        <span class="stop-sub" id="stop-count-lbl"></span>
      </button>
      <button class="action-btn nett" onclick="openNettModal()">🧹 Nettoyage</button>
      <button class="action-btn pause" id="pause-btn" onclick="togglePause()">⏸ Pause pilote</button>
      <div class="action-spacer"></div>
      <button class="action-btn end" onclick="openEndProd()">🏁 Fin de production</button>
    </div>

    <!-- CENTER: form -->
    <div class="prod-col form-col">
      <!-- Infos essentielles always visible -->
      <div class="form-essential">
        <div class="form-grid">
          <div class="form-group">
            <label>N° OF</label>
            <input id="of-num" placeholder="OF123456" oninput="autoSaveForm()">
          </div>
          <div class="form-group">
            <label>Taille</label>
            <select id="of-taille" onchange="autoSaveForm()"><option value="">—</option></select>
          </div>
          <div class="form-group">
            <label>Code produit</label>
            <input id="of-code" placeholder="CODE" oninput="autoSaveForm()">
          </div>
          <div class="form-group">
            <label>Type produit</label>
            <select id="of-type" onchange="autoSaveForm()"><option value="">—</option></select>
          </div>
          <div class="form-group">
            <label>Nb personnes</label>
            <input id="of-nbpers" type="number" min="1" value="1" oninput="autoSaveForm()">
          </div>
          <div class="form-group">
            <label>Co-Pilote</label>
            <input id="of-copilote" oninput="autoSaveForm()">
          </div>
        </div>
      </div>

      <!-- Collapsible details -->
      <div class="form-toggle" id="form-toggle-btn" onclick="toggleFormBody()">
        <span>Détails OF (taie, fibre, qualité…)</span>
        <span class="toggle-arrow">▼</span>
      </div>
      <div class="form-body" id="form-body">
        <div class="form-grid">
          <div class="form-group"><label>Poids garnissage</label><input id="of-poids" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Fibre</label><select id="of-fibre" onchange="autoSaveForm()"><option value="">—</option></select></div>
          <div class="form-group"><label>OF Taie</label><input id="of-taie" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Traca Fibre</label><select id="of-traca" onchange="autoSaveForm()"><option value="">—</option></select></div>
          <div class="form-group"><label>Réf Taie</label><input id="of-ref-taie" oninput="autoSaveForm()"></div>
          <div class="form-group checkbox-row"><input type="checkbox" id="of-kit" onchange="autoSaveForm()"><label for="of-kit" style="text-transform:none;font-size:12px;letter-spacing:0">Kit</label></div>
          <div class="form-group"><label>Qté fab.</label><input id="of-qtefab" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Qté emb.</label><input id="of-qteemb" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Qté init. taie</label><input id="of-qte-init-taie" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Nb taie 2nd choix</label><input id="of-nb-taie2" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Nb défaut couture</label><input id="of-nb-def-cout" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Mq taie</label><input id="of-mq-taie" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Mq housse/encart</label><input id="of-mq-housse" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Nb PP cousue</label><input id="of-nb-pp" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Mq MP (min)</label><input id="of-duree-mq-mp" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group"><label>Mq Personnel (min)</label><input id="of-mq-pers" type="number" min="0" oninput="autoSaveForm()"></div>
          <div class="form-group full"><label>Commentaire</label><textarea id="of-comment" oninput="autoSaveForm()"></textarea></div>
        </div>
      </div>
    </div>

    <!-- RIGHT: stops list -->
    <div class="prod-col recap-col">
      <div class="card" style="flex:1;overflow:hidden;display:flex;flex-direction:column">
        <div class="card-hdr">Arrêts / pauses</div>
        <div id="recap-stops-list" style="overflow-y:auto;flex:1"></div>
      </div>
      <!-- TRS gauge -->
      <div class="card">
        <div class="card-hdr">TRS estimé</div>
        <div class="gauge-wrap">
          <svg class="gauge-svg" viewBox="0 0 140 80">
            <path d="M10,70 A60,60 0 0,1 130,70" fill="none" stroke="#dde4ef" stroke-width="14" stroke-linecap="round"/>
            <path id="gauge-arc" d="M10,70 A60,60 0 0,1 130,70" fill="none" stroke="#1a8c4e" stroke-width="14" stroke-linecap="round" stroke-dasharray="0,1000"/>
          </svg>
          <div class="gauge-text">
            <div class="gauge-pct" id="gauge-pct">—</div>
            <div class="gauge-lbl">TRS</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Active stops bottom bar -->
  <div class="active-stops-bottom" id="active-stops-bottom"></div>
</div>

<!-- ══ MODAL: START PROD ═══════════════════════════════════════════════════ -->
<div class="modal" id="modal-start-prod">
  <div class="modal-box narrow">
    <div class="modal-hdr green"><h2>▶ Démarrer une production</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-start-prod')">✕</button></div>
    <div class="modal-body">
      <p style="font-size:13px;color:var(--gray);margin-bottom:14px">La production démarrera immédiatement avec les infos du formulaire.</p>
      <div id="start-prod-gap" style="background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:10px 14px;font-size:12px;color:#92400e;margin-bottom:12px;display:none"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-start-prod')">Annuler</button>
      <button class="btn btn-green btn-lg" onclick="doStartProd()">▶ Démarrer</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: STOP SELECTOR ════════════════════════════════════════════════ -->
<div class="modal" id="modal-stop">
  <div class="modal-box">
    <div class="modal-hdr red"><h2>⛔ Déclarer un arrêt</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-stop')">✕</button></div>
    <div class="modal-body">
      <div style="font-size:11px;font-weight:700;color:var(--amber);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px">Rattrapages</div>
      <div class="stops-grid ratt" id="stop-grid-ratt"></div>
      <div style="font-size:11px;font-weight:700;color:var(--navy);text-transform:uppercase;letter-spacing:.6px;margin:10px 0 6px">PB Technique</div>
      <div class="stops-grid pb" id="stop-grid-pb"></div>
      <div style="font-size:11px;font-weight:700;color:var(--gray);text-transform:uppercase;letter-spacing:.6px;margin:10px 0 6px">Arrêt libre / autre</div>
      <div class="custom-stop-row">
        <input id="custom-stop-input" placeholder="Nom de l'arrêt personnalisé…" maxlength="60">
        <button class="btn btn-amber" onclick="declareCustomStop()">Déclarer</button>
      </div>
    </div>
  </div>
</div>

<!-- ══ MODAL: END STOP ═════════════════════════════════════════════════════ -->
<div class="modal" id="modal-end-stop">
  <div class="modal-box narrow">
    <div class="modal-hdr amber"><h2 id="end-stop-title">Terminer l'arrêt</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-end-stop')">✕</button></div>
    <div class="modal-body">
      <div style="font-size:13px;color:var(--gray);margin-bottom:12px" id="end-stop-timer-display"></div>
      <div class="form-group">
        <label>Commentaire (optionnel)</label>
        <textarea id="end-stop-comment" style="height:72px;width:100%;border:1.5px solid var(--lgray);border-radius:8px;padding:8px;font-size:13px;resize:none;outline:none"></textarea>
      </div>
      <div style="margin-top:10px;display:flex;align-items:center;gap:8px">
        <input type="checkbox" id="end-stop-hors-trs" style="width:16px;height:16px">
        <label for="end-stop-hors-trs" style="font-size:13px;cursor:pointer">Hors TRS (prévu)</label>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-end-stop')">Annuler</button>
      <button class="btn btn-amber btn-lg" id="end-stop-confirm-btn">✔ Terminer l'arrêt</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: NETTOYAGE ═══════════════════════════════════════════════════ -->
<div class="modal" id="modal-nett">
  <div class="modal-box narrow">
    <div class="modal-hdr" style="background:var(--cyan)"><h2>🧹 Nettoyage</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-nett')">✕</button></div>
    <div class="modal-body">
      <div id="nett-running-info" style="display:none;background:#cffafe;border-radius:8px;padding:12px;margin-bottom:12px">
        <div style="font-weight:700;color:#0e7490;font-size:13px">Nettoyage en cours…</div>
        <div id="nett-timer-disp" style="font-size:22px;font-weight:900;color:#0891b2;font-variant-numeric:tabular-nums"></div>
      </div>
      <div id="nett-start-section">
        <p style="font-size:13px;color:var(--gray);margin-bottom:14px">Choisissez le type de nettoyage :</p>
        <div style="display:flex;flex-direction:column;gap:8px">
          <button class="btn btn-cyan btn-lg" onclick="startNett('court')">Court (~10 min)</button>
          <button class="btn btn-cyan btn-lg" onclick="startNett('long')">Long (~30 min)</button>
          <button class="btn btn-cyan btn-lg" onclick="startNett('grand')">Grand nettoyage (~60 min)</button>
        </div>
      </div>
      <div id="nett-end-section" style="display:none;margin-top:14px">
        <div class="form-group">
          <label>Commentaire</label>
          <textarea id="nett-comment" style="height:60px;width:100%;border:1.5px solid var(--lgray);border-radius:8px;padding:8px;font-size:13px;resize:none;outline:none"></textarea>
        </div>
        <div class="modal-footer" style="padding:12px 0 0;border:none">
          <button class="btn btn-cyan btn-lg" onclick="endNett()">✔ Terminer le nettoyage</button>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ══ MODAL: END PROD ════════════════════════════════════════════════════ -->
<div class="modal" id="modal-end-prod">
  <div class="modal-box wide">
    <div class="modal-hdr"><h2>🏁 Fin de production</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-end-prod')">✕</button></div>
    <div class="modal-body">
      <div id="end-prod-recap" style="background:#f8fafc;border-radius:10px;padding:14px;margin-bottom:16px;display:grid;grid-template-columns:repeat(4,1fr);gap:10px"></div>
      <div class="form-grid" style="gap:10px">
        <div class="form-group"><label>Pilote</label><input id="ep-pilote"></div>
        <div class="form-group"><label>Poste</label><select id="ep-poste"><option>Matin</option><option>Midi</option><option>Nuit</option><option>Jour</option></select></div>
        <div class="form-group"><label>Qté fabriquée</label><input id="ep-qtefab" type="number" min="0" oninput="updateEpEquiv()"></div>
        <div class="form-group"><label>Qté emballée</label><input id="ep-qteemb" type="number" min="0"></div>
        <div class="form-group"><label>Mq MP (min)</label><input id="ep-duree-mq-mp" type="number" min="0"></div>
        <div class="form-group"><label>Mq Personnel (min)</label><input id="ep-manquant-pers" type="number" min="0"></div>
        <div class="form-group"><label>Commentaire</label><textarea id="ep-comment" style="height:48px"></textarea></div>
        <div class="form-group"><label>Équivalence calculée</label><input id="ep-equiv" readonly style="background:#f0f4fb;font-weight:700;color:var(--navy)"></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-end-prod')">Annuler</button>
      <button class="btn btn-navy btn-lg" onclick="doEndProd()">✔ Valider la production</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: END PROD RECAP ═════════════════════════════════════════════ -->
<div class="modal" id="modal-recap">
  <div class="modal-box">
    <div class="modal-hdr green"><h2>✔ Production enregistrée</h2><button class="btn btn-sm btn-ghost" onclick="closeRecap()">✕</button></div>
    <div class="modal-body" id="recap-body"></div>
    <div class="modal-footer">
      <button class="btn btn-green btn-lg" onclick="closeRecap()">Nouvelle production</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: FIN DE POSTE ════════════════════════════════════════════════ -->
<div class="modal" id="modal-fin-poste">
  <div class="modal-box">
    <div class="modal-hdr"><h2>🏁 Fin de mon poste</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-fin-poste')">✕</button></div>
    <div class="modal-body" id="fin-poste-body">
      <div style="text-align:center;padding:20px;color:var(--gray)">Chargement…</div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-fin-poste')">Fermer</button>
      <button class="btn btn-navy btn-lg" onclick="generateDashboard()">📊 Générer supervision</button>
      <button class="btn btn-green btn-lg" onclick="validateFinPoste()">✔ Valider fin de poste</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: EDIT ROW (déclaration ou événement) ═════════════════════════ -->
<div class="modal" id="modal-edit-row">
  <div class="modal-box wide">
    <div class="modal-hdr amber"><h2 id="edit-row-title">Modifier la déclaration</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-edit-row')">✕</button></div>
    <div class="modal-body">
      <div id="edit-row-pw-block" style="margin-bottom:14px">
        <div class="form-group">
          <label>Mot de passe administrateur</label>
          <input type="password" id="edit-row-pw" placeholder="Mot de passe…" style="max-width:260px">
        </div>
        <div id="edit-row-pw-err" style="color:var(--red);font-size:12px;margin-top:6px"></div>
      </div>
      <div id="edit-row-fields" style="display:none">
        <div class="edit-grid" id="edit-row-fields-grid"></div>
        <div style="margin-top:14px;display:flex;gap:10px">
          <button class="btn btn-red" onclick="confirmDeleteRow()">🗑 Supprimer cette ligne</button>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-edit-row')">Annuler</button>
      <button class="btn btn-ghost" id="edit-row-auth-btn" onclick="authenticateEditRow()">Vérifier MDP</button>
      <button class="btn btn-amber btn-lg" id="edit-row-save-btn" style="display:none" onclick="saveEditRow()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: EDIT STOP (depuis liste arrêts droite) ═════════════════════ -->
<div class="modal" id="modal-edit-stop">
  <div class="modal-box narrow">
    <div class="modal-hdr amber"><h2>Modifier l'arrêt</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-edit-stop')">✕</button></div>
    <div class="modal-body">
      <div id="edit-stop-pw-block">
        <div class="form-group" style="margin-bottom:10px">
          <label>Mot de passe administrateur</label>
          <input type="password" id="edit-stop-pw" placeholder="Mot de passe…">
        </div>
        <div id="edit-stop-pw-err" style="color:var(--red);font-size:12px"></div>
      </div>
      <div id="edit-stop-fields" style="display:none">
        <input type="hidden" id="edit-stop-row-num">
        <div class="form-group" style="margin-bottom:10px">
          <label>Type d'arrêt</label>
          <input id="edit-stop-type">
        </div>
        <div class="form-grid" style="gap:10px">
          <div class="form-group">
            <label>Heure début</label>
            <input id="edit-stop-debut" type="time" step="1">
          </div>
          <div class="form-group">
            <label>Heure fin</label>
            <input id="edit-stop-fin" type="time" step="1">
          </div>
        </div>
        <div class="form-group" style="margin-top:10px">
          <label>Commentaire</label>
          <textarea id="edit-stop-comment" style="height:60px;width:100%;border:1.5px solid var(--lgray);border-radius:8px;padding:8px;font-size:13px;resize:none;outline:none"></textarea>
        </div>
        <div style="margin-top:10px">
          <button class="btn btn-red btn-sm" onclick="deleteStopFromEdit()">🗑 Supprimer cet arrêt</button>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-edit-stop')">Annuler</button>
      <button class="btn btn-ghost" id="edit-stop-auth-btn" onclick="authenticateEditStop()">Vérifier MDP</button>
      <button class="btn btn-amber btn-lg" id="edit-stop-save-btn" style="display:none" onclick="saveEditStop()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<!-- ══ MODAL: SETTINGS ════════════════════════════════════════════════════ -->
<div class="modal" id="modal-settings">
  <div class="modal-box wide">
    <div class="modal-hdr"><h2>⚙ Paramètres</h2><button class="btn btn-sm btn-ghost" onclick="closeModal('modal-settings')">✕</button></div>
    <div class="modal-body">
      <div class="set-tabs">
        <button class="set-tab active" onclick="switchSetTab(0)">Général</button>
        <button class="set-tab" onclick="switchSetTab(1)">Modèles horaires</button>
        <button class="set-tab" onclick="switchSetTab(2)">Pilotes</button>
        <button class="set-tab" onclick="switchSetTab(3)">Administrateur</button>
      </div>
      <!-- Général -->
      <div class="set-panel active" id="set-panel-0">
        <div class="form-group" style="margin-bottom:12px">
          <label>Fichier Excel (base de données)</label>
          <div style="display:flex;gap:8px;align-items:center">
            <input id="set-db-path" placeholder="Chemin vers le fichier .xlsx" style="flex:1;border:1.5px solid var(--lgray);border-radius:6px;padding:7px 10px;font-size:12px">
            <button class="btn btn-navy btn-sm" onclick="setDb()">Appliquer</button>
          </div>
          <div id="set-db-name" style="font-size:11px;color:var(--gray);margin-top:4px"></div>
        </div>
        <div class="form-grid" style="gap:12px">
          <div class="form-group"><label>Prod. ref 8h (equiv/8h)</label><input id="set-prod-ref" type="number" min="0"></div>
          <div class="form-group"><label>Pause max imputée (min)</label><input id="set-pause-max" type="number" min="0"></div>
          <div class="form-group"><label>Nett. court (min)</label><input id="set-clean-short" type="number" min="0"></div>
          <div class="form-group"><label>Nett. long (min)</label><input id="set-clean-long" type="number" min="0"></div>
          <div class="form-group"><label>Grand nettoyage (min)</label><input id="set-clean-grand" type="number" min="0"></div>
          <div class="form-group"><label>Tolérance réunion (min)</label><input id="set-meeting-tol" type="number" min="0"></div>
        </div>
        <div style="margin-top:14px">
          <button class="btn btn-navy" onclick="saveSettings()">Enregistrer</button>
          <span id="set-save-ok" style="color:var(--green);font-size:12px;margin-left:10px"></span>
        </div>
      </div>
      <!-- Modèles horaires -->
      <div class="set-panel" id="set-panel-1">
        <p style="font-size:12px;color:var(--gray);margin-bottom:14px">Définissez les plages horaires de chaque poste (pour calcul TRS par poste).</p>
        <div id="modeles-list"></div>
        <button class="btn btn-navy btn-sm" style="margin-top:10px" onclick="addModele()">+ Ajouter un modèle</button>
        <div style="margin-top:14px">
          <button class="btn btn-navy" onclick="saveSettings()">Enregistrer</button>
        </div>
        <div style="margin-top:14px;background:#f8fafc;border-radius:8px;padding:12px;font-size:12px;color:var(--gray)">
          <strong>Formule TRS :</strong><br>
          TRS% = Equiv / (Prod_ref × Durée_OF_s / 28800) × 100<br>
          Durée_OF_s = (Fin − Début) + Changements_OF − min(Pauses, Pause_max)
        </div>
      </div>
      <!-- Pilotes -->
      <div class="set-panel" id="set-panel-2">
        <p style="font-size:12px;color:var(--gray);margin-bottom:14px">Les pilotes sont définis dans la feuille "Listes" du fichier Excel. Rechargez le fichier pour mettre à jour.</p>
        <button class="btn btn-ghost" onclick="reloadLists()">🔄 Recharger les listes</button>
        <div id="set-pilots-list" style="margin-top:14px;font-size:12px"></div>
      </div>
      <!-- Admin -->
      <div class="set-panel" id="set-panel-3">
        <div style="margin-bottom:12px">
          <div class="form-group"><label>Mot de passe actuel</label><input type="password" id="set-cur-pw" style="max-width:260px"></div>
          <div class="form-group" style="margin-top:8px"><label>Nouveau mot de passe</label><input type="password" id="set-new-pw" style="max-width:260px"></div>
          <button class="btn btn-navy" style="margin-top:10px" onclick="saveSettings()">Changer le mot de passe</button>
          <div id="set-pw-err" style="color:var(--red);font-size:12px;margin-top:6px"></div>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-settings')">Fermer</button>
    </div>
  </div>
</div>

<script>
// ─── State ──────────────────────────────────────────────────────────────────
let st = {}, lists = {}, histRows = [], evtRows = [];
let pollTimer = null;
let tickInterval = null;  // 1-second tick for active stop timers
let stopActiveKeys = [];  // clés des arrêts actifs
let currentEditRowNum = null;
let currentEditRowData = null;
let currentEndStopKey = null;
let nettType = null;
let formSaveTimer = null;
let setTabIdx = 0;

// ─── Init ────────────────────────────────────────────────────────────────────
async function init() {
  const r = await fetch('/api/state');
  st = await r.json();
  const lr = await fetch('/api/lists');
  lists = await lr.json();
  populateLists();
  const cr = await fetch('/api/config');
  const cfg = await cr.json();
  window._cfg = cfg;
  if (st.pilot) {
    if (st.prod_active) { showView('prod'); restoreForm(); }
    else showView('main');
  } else {
    showView('login');
  }
  startTick();
  startPoll();
  buildStopGrid();
}

function startPoll() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(doPoll, 5000);
}

function startTick() {
  if (tickInterval) clearInterval(tickInterval);
  tickInterval = setInterval(tickTimers, 1000);
}

async function doPoll() {
  try {
    const r = await fetch('/api/state');
    st = await r.json();
    updateUI();
  } catch(e){}
}

// Tick every second — updates only active stop timers locally
function tickTimers() {
  if (!st || !st.timers) return;
  const now = Date.now() / 1000;
  stopActiveKeys = st.active_stops || [];
  for (const k of stopActiveKeys) {
    const t = st.timers[k];
    if (t && t.running) t.elapsed += 1;
  }
  // Also increment of_elapsed if prod active
  if (st.prod_active && !st.is_paused) {
    st.of_elapsed_s = (st.of_elapsed_s || 0) + 1;
  }
  updateTimerDisplays();
  updateDarkTheme();
}

function updateDarkTheme() {
  const hasStop = stopActiveKeys.length > 0;
  document.body.classList.toggle('stop-active', hasStop);
}

// ─── UI Update ───────────────────────────────────────────────────────────────
function updateUI() {
  // Header pilot
  const pilot = st.pilot || '—';
  const poste = st.poste || '—';
  document.getElementById('main-pilot-name').textContent = pilot;
  document.getElementById('main-poste-name').textContent = poste;
  document.getElementById('prod-hdr-pilot').textContent = pilot;
  document.getElementById('pob-pilot').textContent = pilot;
  document.getElementById('pob-poste').textContent = poste;
  // OF from form
  const ofNum = (st.form || {}).of_num || '—';
  document.getElementById('pob-of').textContent = ofNum || '—';

  updateTimerDisplays();
  updateDarkTheme();
  renderActiveStopsBottom();
  renderRecapStops();
  updateStopBtnStates();
  stopActiveKeys = st.active_stops || [];
}

function updateTimerDisplays() {
  // OF timer
  const ofEl = document.getElementById('sc-of-val');
  if (ofEl && st.prod_active) ofEl.textContent = fmtS(st.of_elapsed_s || 0);

  // Stop total
  const stEl = document.getElementById('sc-stop-val');
  if (stEl) {
    let total = 0;
    for (const k in (st.timers || {})) {
      if (k.startsWith('_')) continue;
      const t = st.timers[k];
      total += t.elapsed || 0;
    }
    stEl.textContent = fmtS(total);
  }

  // TRS estimate
  const pr = st.prod_ref || 0;
  const ofS = st.of_elapsed_s || 0;
  const form = st.form || {};
  const qte = parseFloat(form.qte_fab || 0) || 0;
  const trsEl = document.getElementById('sc-trs-val');
  const gaugePct = document.getElementById('gauge-pct');
  const gaugeArc = document.getElementById('gauge-arc');
  if (pr > 0 && ofS > 0 && qte > 0) {
    const trs = Math.round(qte / (pr * ofS / 28800) * 100 * 10) / 10;
    if (trsEl) { trsEl.textContent = trs + '%'; trsEl.style.color = trsColor(trs); }
    if (gaugePct) { gaugePct.textContent = trs + '%'; gaugePct.style.color = trsColor(trs); }
    if (gaugeArc) {
      const pct = Math.min(100, trs);
      const total = Math.PI * 60;
      gaugeArc.style.strokeDasharray = (pct/100*total) + ',' + total;
      gaugeArc.style.stroke = trsColor(trs);
    }
  } else {
    if (trsEl) { trsEl.textContent = '—'; trsEl.style.color = ''; }
    if (gaugePct) { gaugePct.textContent = '—'; }
  }

  // Active stop cards bottom
  renderActiveStopsBottom();
}

function renderActiveStopsBottom() {
  const bar = document.getElementById('active-stops-bottom');
  if (!bar) return;
  const active = st.active_stops || [];
  bar.innerHTML = '';
  if (active.length === 0) {
    bar.classList.remove('has-stops');
    return;
  }
  bar.classList.add('has-stops');
  for (const k of active) {
    const t = st.timers[k] || {};
    const el = t.elapsed || 0;
    const isNett = k === 'nettoyage';
    const label = k === 'nettoyage' ? '🧹 Nettoyage' : (getEvtLabel(k) || k);
    const cls = isNett ? 'active-stop-card nett' : 'active-stop-card';
    const div = document.createElement('div');
    div.className = cls;
    div.innerHTML = `<div class="asc-label">${label}</div><div class="asc-timer">${fmtS(el)}</div><div class="asc-hint">Cliquer pour terminer</div>`;
    div.onclick = () => openEndStop(k);
    bar.appendChild(div);
  }
}

function renderRecapStops() {
  const el = document.getElementById('recap-stops-list');
  if (!el) return;
  const events = st.tl_events || [];
  el.innerHTML = '';
  const colors = {ratt:'#d97706', pb:'#1a1f5e', nettoyage:'#0891b2'};
  events.forEach(ev => {
    if (ev.key.startsWith('_')) return;
    const running = !ev.end;
    const dur = running ? ((Date.now()/1000) - new Date(ev.start).getTime()/1000) : ((new Date(ev.end)-new Date(ev.start))/1000);
    const color = colors[ev.cat] || '#64748b';
    const label = ev.key === 'nettoyage' ? 'Nettoyage' : (getEvtLabel(ev.key) || ev.key);
    const row = document.createElement('div');
    row.className = 'recap-stop-row' + (running ? ' running' : '');
    row.title = 'Cliquer pour modifier';
    row.innerHTML = `<div class="recap-stop-dot" style="background:${color}"></div>
      <div class="recap-stop-name">${label}</div>
      <div class="recap-stop-time">${fmtS(Math.abs(dur))}</div>
      <button class="btn-icon edit" onclick="openEditStopFromList(event, ${JSON.stringify(ev)})">✏️</button>`;
    el.appendChild(row);
  });
}

function getEvtLabel(key) {
  const evtMap = {
    ratt_pochon:'Pochon / Fibre', ratt_couture:'Couture', ratt_emb:'Emballage',
    ratt_presse_soud:'Presse Souder', ratt_presse_zip:'Presse ZIP',
    nettoyage:'Nettoyage', pb_chargeuse:'Chargeuse', pb_carde:'Carde',
    pb_etaleur:'Etaleur / Tour', pb_coupe:'Coupe / Circ.', pb_tapis1:'Tapis Bascule',
    pb_enrouleur:'Enrouleur Pochon', pb_pesee:'Pesée / Tapis 2',
    pb_deviation:'Déviation / Table', pb_enfileur:'Enfileur Pochon',
    pb_kinna:'Kinna / Stroebel', pb_tapeuse:'Tapeuse', pb_table_rot:'Table Rot. / Twin',
    pb_h100:'Enfileuse H100', pb_traversin:'Enfileuse Traversin',
    pb_presse_orc:'Presse ORC', pb_presse_zip2:'Presse Housse ZIP',
    pb_cercleuse:'Cercleuse', pb_enrouleuse:'Enrouleuse Traversin',
    arret_mp:'Matière première', arret_reunion:'Réunion',
  };
  return evtMap[key] || key;
}

function updateStopBtnStates() {
  const active = st.active_stops || [];
  document.querySelectorAll('.stop-btn').forEach(btn => {
    const k = btn.dataset.key;
    const isActive = active.includes(k);
    btn.classList.toggle('running', isActive);
    btn.onclick = isActive ? () => openEndStop(k) : () => doStartStop(k, btn.dataset.cat);
  });
  const lbl = document.getElementById('stop-count-lbl');
  if (lbl) lbl.textContent = active.length ? `${active.length} en cours` : '';
  const pauseBtn = document.getElementById('pause-btn');
  if (pauseBtn) pauseBtn.textContent = st.is_paused ? '▶ Reprendre' : '⏸ Pause pilote';
}

// ─── Views ───────────────────────────────────────────────────────────────────
function showView(name) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById('view-' + name).classList.add('active');
  if (name === 'main') {
    loadHistory();
    loadEvents();
  }
}

function switchMainTab(tab) {
  document.getElementById('main-panel-decl').style.display = tab === 'decl' ? 'block' : 'none';
  document.getElementById('main-panel-evt').style.display  = tab === 'evt'  ? 'block' : 'none';
  document.getElementById('main-tab-decl').classList.toggle('active', tab === 'decl');
  document.getElementById('main-tab-evt').classList.toggle('active', tab === 'evt');
}

// ─── Login ───────────────────────────────────────────────────────────────────
function populateLists() {
  const pilotSel = document.getElementById('login-pilot');
  (lists.pilotes || []).forEach(p => {
    const o = document.createElement('option');
    o.value = o.textContent = p;
    pilotSel.appendChild(o);
  });
  const tailleSel = document.getElementById('of-taille');
  (lists.tailles || []).forEach(t => { const o=document.createElement('option'); o.value=o.textContent=t; tailleSel.appendChild(o); });
  const typeSel = document.getElementById('of-type');
  (lists.types_prod || []).forEach(t => { const o=document.createElement('option'); o.value=o.textContent=t; typeSel.appendChild(o); });
  const fibreSel = document.getElementById('of-fibre');
  (lists.fibres || []).forEach(t => { const o=document.createElement('option'); o.value=o.textContent=t; fibreSel.appendChild(o); });
  const tracaSel = document.getElementById('of-traca');
  (lists.tracas || []).forEach(t => { const o=document.createElement('option'); o.value=o.textContent=t; tracaSel.appendChild(o); });
}

async function doLogin() {
  const pilot = document.getElementById('login-pilot').value;
  const poste = document.getElementById('login-poste').value;
  if (!pilot || !poste) { document.getElementById('login-err').textContent = 'Sélectionnez pilote et poste'; return; }
  const r = await fetch('/api/login', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({pilot, poste})});
  const d = await r.json();
  if (d.ok) {
    st.pilot = pilot; st.poste = poste;
    showView('main');
    // Prefill form
    document.getElementById('of-num').value = '';
  } else {
    document.getElementById('login-err').textContent = d.error || 'Erreur';
  }
}

async function doLogout() {
  const r = await fetch('/api/logout', {method:'POST'});
  const d = await r.json();
  if (d.ok) showView('login');
  else alert(d.error);
}

// ─── History / Events ────────────────────────────────────────────────────────
async function loadHistory() {
  const r = await fetch('/api/history');
  histRows = await r.json();
  const tbody = document.getElementById('hist-body');
  tbody.innerHTML = '';
  for (const row of histRows) {
    const trs = row.trs >= 0 ? row.trs + '%' : '—';
    const trsCls = row.trs >= 70 ? 'trs-hi' : row.trs >= 50 ? 'trs-warn' : row.trs >= 0 ? 'trs-low' : '';
    const tr = document.createElement('tr');
    tr.className = trsCls;
    tr.innerHTML = `
      <td><strong>${row.of||'—'}</strong></td>
      <td>${row.date||''}</td><td>${row.poste||''}</td><td>${row.pilote||''}</td>
      <td>${row.debut||''}</td><td>${row.fin||''}</td><td>${row.taille||''}</td>
      <td>${row.qte_fab||''}</td><td>${row.qte_emb||''}</td><td>${row.equiv||''}</td>
      <td class="trs-cell">${trs}</td>
      <td>
        <button class="btn-icon edit" title="Modifier" onclick="openEditRow(${row.row_num}, 'prod')">✏️</button>
        <button class="btn-icon del" title="Supprimer" onclick="openEditRow(${row.row_num}, 'prod')">🗑️</button>
      </td>`;
    tbody.appendChild(tr);
  }
}

async function loadEvents() {
  const r = await fetch('/api/events_list');
  evtRows = await r.json();
  const tbody = document.getElementById('evt-body');
  tbody.innerHTML = '';
  for (const row of evtRows) {
    const tr = document.createElement('tr');
    const cls = row.hors_trs ? '' : '';
    tr.innerHTML = `
      <td><strong>${row.type||''}</strong></td>
      <td>${row.of||''}</td><td>${row.date||''}</td><td>${row.pilote||''}</td>
      <td>${row.debut||''}</td><td>${row.fin||''}</td><td>${row.duree||''}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${row.comment||''}</td>
      <td>
        <button class="btn-icon edit" title="Modifier" onclick="openEditRow(${row.row_num}, 'evt')">✏️</button>
        <button class="btn-icon del" title="Supprimer" onclick="openEditRow(${row.row_num}, 'evt')">🗑️</button>
      </td>`;
    tbody.appendChild(tr);
  }
}

// ─── Edit Row Modal ───────────────────────────────────────────────────────────
const DECL_HEADERS = [
  "Type","OF","Date","Poste","Pilote","Co-Pilote","Nb Personnes",
  "Taille","Code Produit","Type Produit","Poids Garnissage","Fibre",
  "OF Taie","Traca Fibre","Ref Taie","Kit",
  "Heure Debut","Heure Fin","Duree",
  "Qte Fabriquee","Qte Emballee","Equivalence","Cadence/h","Cadence/h/pers","TRS%",
  "Qte Init Taie","Nb Taie 2nd Choix","Nb Defaut Couture",
  "Mq Taie","Mq Housse/Encart","Nb PP Cousue",
  "Changement de Serie","Manquant MP","Manquant Personnel/Reunion",
  "Nettoyage Fin de Poste","Commentaire","Prevu/Hors TRS"
];

function openEditRow(rowNum, mode) {
  currentEditRowNum = rowNum;
  document.getElementById('edit-row-title').textContent = mode === 'prod' ? 'Modifier la déclaration' : "Modifier l'événement";
  document.getElementById('edit-row-pw').value = '';
  document.getElementById('edit-row-pw-err').textContent = '';
  document.getElementById('edit-row-fields').style.display = 'none';
  document.getElementById('edit-row-save-btn').style.display = 'none';
  document.getElementById('edit-row-auth-btn').style.display = '';
  document.getElementById('edit-row-fields-grid').innerHTML = '';
  openModal('modal-edit-row');
}

async function authenticateEditRow() {
  const pw = document.getElementById('edit-row-pw').value;
  // Try a dummy edit to verify password
  const r = await fetch('/api/edit_row', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({pw, row_num: currentEditRowNum, updates: {}})});
  const d = await r.json();
  if (!d.ok) {
    document.getElementById('edit-row-pw-err').textContent = d.error || 'Mot de passe incorrect';
    return;
  }
  // Find row data
  let rowData = histRows.find(r => r.row_num === currentEditRowNum) || evtRows.find(r => r.row_num === currentEditRowNum);
  document.getElementById('edit-row-pw-block').style.display = 'none';
  document.getElementById('edit-row-auth-btn').style.display = 'none';
  document.getElementById('edit-row-save-btn').style.display = '';
  document.getElementById('edit-row-fields').style.display = 'block';
  buildEditRowFields(rowData);
}

function buildEditRowFields(row) {
  const grid = document.getElementById('edit-row-fields-grid');
  grid.innerHTML = '';
  if (!row) {
    // Fallback: show all 37 headers as blank inputs
    DECL_HEADERS.forEach((h, i) => {
      const div = document.createElement('div');
      div.className = 'edit-group';
      div.innerHTML = `<label>${h}</label><input data-col="${i+1}" value="">`;
      grid.appendChild(div);
    });
    return;
  }
  const isProd = (row.type || '').toLowerCase() === 'production' || !row.type;
  // Build fields from known row keys
  const fieldMap = [
    {col:1, label:'Type', val: row.type||''},
    {col:2, label:'OF', val: row.of||''},
    {col:3, label:'Date (jj/mm/aaaa)', val: row.date||''},
    {col:4, label:'Poste', val: row.poste||''},
    {col:5, label:'Pilote', val: row.pilote||''},
    {col:6, label:'Co-Pilote', val: row.copilote||''},
    {col:7, label:'Nb Personnes', val: row.nb_pers||''},
    {col:8, label:'Taille', val: row.taille||''},
    {col:9, label:'Code Produit', val: row.code_prod||''},
    {col:10,label:'Type Produit', val: row.type_prod||''},
    {col:11,label:'Poids Garnissage', val: row.poids||''},
    {col:12,label:'Fibre', val: row.fibre||''},
    {col:13,label:'OF Taie', val: row.of_taie||''},
    {col:14,label:'Traca Fibre', val: row.traca||''},
    {col:15,label:'Ref Taie', val: row.ref_taie||''},
    {col:16,label:'Kit', val: row.kit||''},
    {col:17,label:'Heure Debut', val: row.debut||''},
    {col:18,label:'Heure Fin', val: row.fin||''},
    {col:19,label:'Durée', val: row.duree||''},
    ...(isProd ? [
      {col:20,label:'Qté Fabriquée', val: row.qte_fab||''},
      {col:21,label:'Qté Emballée', val: row.qte_emb||''},
      {col:22,label:'Équivalence', val: row.equiv||''},
      {col:23,label:'Cadence/h', val: row.c1||''},
      {col:24,label:'Cadence/h/pers', val: row.c2||''},
      {col:25,label:'TRS%', val: row.trs >= 0 ? String(row.trs) : ''},
    ] : []),
    {col:36,label:'Commentaire', val: row.comment||''},
    {col:37,label:'Prevu/Hors TRS', val: row.hors_trs ? 'OUI' : ''},
  ];
  fieldMap.forEach(f => {
    const div = document.createElement('div');
    div.className = 'edit-group' + (f.col===36 ? ' full' : '');
    div.innerHTML = `<label>${f.label}</label><input data-col="${f.col}" value="${(f.val||'').toString().replace(/"/g,'&quot;')}">`;
    grid.appendChild(div);
  });
}

async function saveEditRow() {
  const pw = document.getElementById('edit-row-pw').value;
  const updates = {};
  document.querySelectorAll('#edit-row-fields-grid input[data-col]').forEach(inp => {
    updates[inp.dataset.col] = inp.value;
  });
  const r = await fetch('/api/edit_row', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({pw, row_num: currentEditRowNum, updates})});
  const d = await r.json();
  if (d.ok) {
    closeModal('modal-edit-row');
    setTimeout(() => { loadHistory(); loadEvents(); }, 800);
  } else alert(d.error || 'Erreur');
}

async function confirmDeleteRow() {
  if (!confirm('Supprimer définitivement cette ligne ?')) return;
  const pw = document.getElementById('edit-row-pw').value;
  const r = await fetch('/api/delete_row', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({pw, row_num: currentEditRowNum})});
  const d = await r.json();
  if (d.ok) {
    closeModal('modal-edit-row');
    setTimeout(() => { loadHistory(); loadEvents(); }, 800);
  } else alert(d.error || 'Erreur');
}

// ─── Edit Stop from list (right panel) ───────────────────────────────────────
function openEditStopFromList(ev, evData) {
  ev.stopPropagation();
  document.getElementById('edit-stop-pw').value = '';
  document.getElementById('edit-stop-pw-err').textContent = '';
  document.getElementById('edit-stop-fields').style.display = 'none';
  document.getElementById('edit-stop-save-btn').style.display = 'none';
  document.getElementById('edit-stop-auth-btn').style.display = '';
  window._editStopEvData = evData;
  openModal('modal-edit-stop');
}

async function authenticateEditStop() {
  const pw = document.getElementById('edit-stop-pw').value;
  const r = await fetch('/api/edit_row', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({pw, row_num: 1, updates: {}})});
  const d = await r.json();
  if (!d.ok && d.error && d.error.includes('Mot de passe')) {
    document.getElementById('edit-stop-pw-err').textContent = 'Mot de passe incorrect';
    return;
  }
  // For live session events there's no row_num yet (not saved to Excel)
  // But we authenticate and show the edit fields
  const ev = window._editStopEvData;
  document.getElementById('edit-stop-pw-block').style.display = 'none';
  document.getElementById('edit-stop-auth-btn').style.display = 'none';
  document.getElementById('edit-stop-save-btn').style.display = '';
  document.getElementById('edit-stop-fields').style.display = 'block';
  document.getElementById('edit-stop-type').value = getEvtLabel(ev.key) || ev.key;
  document.getElementById('edit-stop-debut').value = ev.start ? ev.start.substring(11,19) : '';
  document.getElementById('edit-stop-fin').value = ev.end ? ev.end.substring(11,19) : '';
  document.getElementById('edit-stop-comment').value = ev.comment || '';
  document.getElementById('edit-stop-row-num').value = ev.row_num || '';
}

async function saveEditStop() {
  const pw = document.getElementById('edit-stop-pw').value;
  const rowNum = parseInt(document.getElementById('edit-stop-row-num').value) || 0;
  if (rowNum) {
    const updates = {
      1: document.getElementById('edit-stop-type').value,
      17: document.getElementById('edit-stop-debut').value,
      18: document.getElementById('edit-stop-fin').value,
      36: document.getElementById('edit-stop-comment').value,
    };
    const r = await fetch('/api/edit_row', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({pw, row_num: rowNum, updates})});
    const d = await r.json();
    if (!d.ok) { alert(d.error); return; }
  }
  closeModal('modal-edit-stop');
}

async function deleteStopFromEdit() {
  if (!confirm('Supprimer cet arrêt ?')) return;
  const pw = document.getElementById('edit-stop-pw').value;
  const rowNum = parseInt(document.getElementById('edit-stop-row-num').value) || 0;
  if (rowNum) {
    const r = await fetch('/api/delete_row', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({pw, row_num: rowNum})});
    const d = await r.json();
    if (!d.ok) { alert(d.error); return; }
  }
  closeModal('modal-edit-stop');
}

// ─── Stop grid ────────────────────────────────────────────────────────────────
function buildStopGrid() {
  const evts = [
    {lbl:'Pochon / Fibre', key:'ratt_pochon', cat:'ratt'},
    {lbl:'Couture', key:'ratt_couture', cat:'ratt'},
    {lbl:'Emballage', key:'ratt_emb', cat:'ratt'},
    {lbl:'Presse Souder', key:'ratt_presse_soud', cat:'ratt'},
    {lbl:'Presse ZIP', key:'ratt_presse_zip', cat:'ratt'},
    {lbl:'Matière Première', key:'arret_mp', cat:'ratt'},
  ];
  const pbs = [
    {lbl:'Chargeuse', key:'pb_chargeuse', cat:'pb'},
    {lbl:'Carde', key:'pb_carde', cat:'pb'},
    {lbl:'Etaleur/Tour', key:'pb_etaleur', cat:'pb'},
    {lbl:'Coupe/Circ.', key:'pb_coupe', cat:'pb'},
    {lbl:'Tapis Bascule', key:'pb_tapis1', cat:'pb'},
    {lbl:'Enrouleur Pochon', key:'pb_enrouleur', cat:'pb'},
    {lbl:'Pesée/Tapis 2', key:'pb_pesee', cat:'pb'},
    {lbl:'Déviation/Table', key:'pb_deviation', cat:'pb'},
    {lbl:'Enfileur Pochon', key:'pb_enfileur', cat:'pb'},
    {lbl:'Kinna/Stroebel', key:'pb_kinna', cat:'pb'},
    {lbl:'Tapeuse', key:'pb_tapeuse', cat:'pb'},
    {lbl:'Table Rot./Twin', key:'pb_table_rot', cat:'pb'},
    {lbl:'Enfileuse H100', key:'pb_h100', cat:'pb'},
    {lbl:'Enfileuse Traversin', key:'pb_traversin', cat:'pb'},
    {lbl:'Presse ORC', key:'pb_presse_orc', cat:'pb'},
    {lbl:'Presse Housse ZIP', key:'pb_presse_zip2', cat:'pb'},
    {lbl:'Cercleuse', key:'pb_cercleuse', cat:'pb'},
    {lbl:'Enrouleuse Traversin', key:'pb_enrouleuse', cat:'pb'},
  ];
  const rattGrid = document.getElementById('stop-grid-ratt');
  const pbGrid = document.getElementById('stop-grid-pb');
  evts.forEach(e => rattGrid.appendChild(makeStopBtn(e)));
  pbs.forEach(e => pbGrid.appendChild(makeStopBtn(e)));
}

function makeStopBtn(e) {
  const btn = document.createElement('button');
  btn.className = 'stop-btn ' + e.cat;
  btn.dataset.key = e.key;
  btn.dataset.cat = e.cat;
  btn.textContent = e.lbl;
  btn.onclick = () => doStartStop(e.key, e.cat);
  return btn;
}

async function doStartStop(key, cat) {
  // Start timer immediately (before closing modal)
  const now = Date.now();
  if (!st.timers) st.timers = {};
  st.timers[key] = {elapsed: 0, running: true, start: now/1000};
  if (!st.active_stops) st.active_stops = [];
  if (!st.active_stops.includes(key)) st.active_stops.push(key);
  stopActiveKeys = st.active_stops;
  updateDarkTheme();
  renderActiveStopsBottom();
  closeModal('modal-stop');

  await fetch('/api/start_stop', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({key, cat})});
  doPoll();
}

async function declareCustomStop() {
  const inp = document.getElementById('custom-stop-input');
  const name = inp.value.trim();
  if (!name) { inp.focus(); return; }
  const key = 'custom_' + name.toLowerCase().replace(/\s+/g,'_').substring(0,30);
  if (!st.timers) st.timers = {};
  st.timers[key] = {elapsed: 0, running: true};
  if (!st.active_stops) st.active_stops = [];
  if (!st.active_stops.includes(key)) st.active_stops.push(key);
  stopActiveKeys = st.active_stops;
  // Register custom label
  window._customLabels = window._customLabels || {};
  window._customLabels[key] = name;
  updateDarkTheme();
  renderActiveStopsBottom();
  closeModal('modal-stop');
  inp.value = '';
  await fetch('/api/start_stop', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({key, cat:'pb', label: name})});
  doPoll();
}

// ─── End Stop ─────────────────────────────────────────────────────────────────
function openEndStop(key) {
  currentEndStopKey = key;
  const lbl = key === 'nettoyage' ? '🧹 Nettoyage' : (getEvtLabel(key) || key);
  document.getElementById('end-stop-title').textContent = 'Terminer : ' + lbl;
  document.getElementById('end-stop-comment').value = '';
  document.getElementById('end-stop-hors-trs').checked = false;
  const t = (st.timers || {})[key] || {};
  document.getElementById('end-stop-timer-display').textContent = 'Durée : ' + fmtS(t.elapsed || 0);
  document.getElementById('end-stop-confirm-btn').onclick = doEndStop;
  openModal('modal-end-stop');
}

async function doEndStop() {
  const comment = document.getElementById('end-stop-comment').value;
  const horsTrs = document.getElementById('end-stop-hors-trs').checked;
  if (currentEndStopKey === 'nettoyage') {
    await fetch('/api/end_nettoyage', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({comment})});
  } else {
    await fetch('/api/end_stop', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({key: currentEndStopKey, comment, hors_trs: horsTrs})});
  }
  closeModal('modal-end-stop');
  doPoll();
}

// ─── Nettoyage ────────────────────────────────────────────────────────────────
function openNettModal() {
  const isRunning = (st.active_stops || []).includes('nettoyage');
  document.getElementById('nett-running-info').style.display = isRunning ? 'block' : 'none';
  document.getElementById('nett-start-section').style.display = isRunning ? 'none' : 'block';
  document.getElementById('nett-end-section').style.display = isRunning ? 'block' : 'none';
  if (isRunning) {
    const t = (st.timers || {}).nettoyage || {};
    document.getElementById('nett-timer-disp').textContent = fmtS(t.elapsed || 0);
  }
  openModal('modal-nett');
}

async function startNett(type) {
  nettType = type;
  await fetch('/api/start_nettoyage', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({ntype: type})});
  closeModal('modal-nett');
  doPoll();
}

async function endNett() {
  const comment = document.getElementById('nett-comment').value;
  await fetch('/api/end_nettoyage', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({comment})});
  closeModal('modal-nett');
  doPoll();
}

// ─── Start Prod ───────────────────────────────────────────────────────────────
function openStartProd() {
  document.getElementById('start-prod-gap').style.display = 'none';
  openModal('modal-start-prod');
}

async function doStartProd() {
  const r = await fetch('/api/start_prod', {method:'POST'});
  const d = await r.json();
  if (!d.ok) { alert(d.error); return; }
  closeModal('modal-start-prod');
  // Clear OF field immediately
  document.getElementById('of-num').value = '';
  st.prod_active = true;
  st.form = {of_num: ''};
  // Switch directly to production view
  showView('prod');
  doPoll();
}

// ─── Pause ────────────────────────────────────────────────────────────────────
async function togglePause() {
  await fetch('/api/toggle_pause', {method:'POST'});
  doPoll();
}

// ─── Form save ────────────────────────────────────────────────────────────────
function autoSaveForm() {
  if (formSaveTimer) clearTimeout(formSaveTimer);
  formSaveTimer = setTimeout(doSaveForm, 600);
  // Update OF banner immediately
  const ofVal = document.getElementById('of-num').value;
  const el = document.getElementById('pob-of');
  if (el) el.textContent = ofVal || '—';
}

async function doSaveForm() {
  const form = collectForm();
  st.form = form;
  await fetch('/api/save_form', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify(form)});
}

function collectForm() {
  return {
    of_num: document.getElementById('of-num').value,
    taille: document.getElementById('of-taille').value,
    code_prod: document.getElementById('of-code').value,
    type_prod: document.getElementById('of-type').value,
    nb_pers: document.getElementById('of-nbpers').value,
    copilote: document.getElementById('of-copilote').value,
    poids: document.getElementById('of-poids').value,
    fibre: document.getElementById('of-fibre').value,
    of_taie: document.getElementById('of-taie').value,
    traca: document.getElementById('of-traca').value,
    ref_taie: document.getElementById('of-ref-taie').value,
    kit: document.getElementById('of-kit').checked,
    qte_fab: document.getElementById('of-qtefab').value,
    qte_emb: document.getElementById('of-qteemb').value,
    qte_init_taie: document.getElementById('of-qte-init-taie').value,
    nb_taie2_choix: document.getElementById('of-nb-taie2').value,
    nb_def_cout: document.getElementById('of-nb-def-cout').value,
    mq_taie: document.getElementById('of-mq-taie').value,
    mq_housse_encart: document.getElementById('of-mq-housse').value,
    nb_pp_cousue: document.getElementById('of-nb-pp').value,
    duree_mq_mp: document.getElementById('of-duree-mq-mp').value,
    manquant_pers: document.getElementById('of-mq-pers').value,
    comment: document.getElementById('of-comment').value,
  };
}

function restoreForm() {
  const f = st.form || {};
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
  set('of-num', f.of_num); set('of-taille', f.taille); set('of-code', f.code_prod);
  set('of-type', f.type_prod); set('of-nbpers', f.nb_pers || 1); set('of-copilote', f.copilote);
  set('of-poids', f.poids); set('of-fibre', f.fibre); set('of-taie', f.of_taie);
  set('of-traca', f.traca); set('of-ref-taie', f.ref_taie);
  const kit = document.getElementById('of-kit'); if (kit) kit.checked = !!f.kit;
  set('of-qtefab', f.qte_fab); set('of-qteemb', f.qte_emb);
  set('of-qte-init-taie', f.qte_init_taie); set('of-nb-taie2', f.nb_taie2_choix);
  set('of-nb-def-cout', f.nb_def_cout); set('of-mq-taie', f.mq_taie);
  set('of-mq-housse', f.mq_housse_encart); set('of-nb-pp', f.nb_pp_cousue);
  set('of-duree-mq-mp', f.duree_mq_mp); set('of-mq-pers', f.manquant_pers);
  set('of-comment', f.comment);
  const ofEl = document.getElementById('pob-of');
  if (ofEl) ofEl.textContent = f.of_num || '—';
}

function toggleFormBody() {
  const body = document.getElementById('form-body');
  const btn = document.getElementById('form-toggle-btn');
  body.classList.toggle('visible');
  btn.classList.toggle('open');
}

// ─── End Production ──────────────────────────────────────────────────────────
function openEndProd() {
  const f = collectForm();
  document.getElementById('ep-pilote').value = st.pilot || '';
  document.getElementById('ep-poste').value = st.poste || '';
  document.getElementById('ep-qtefab').value = f.qte_fab || '';
  document.getElementById('ep-qteemb').value = f.qte_emb || '';
  document.getElementById('ep-comment').value = f.comment || '';
  document.getElementById('ep-duree-mq-mp').value = f.duree_mq_mp || '';
  document.getElementById('ep-manquant-pers').value = f.manquant_pers || '';
  updateEpEquiv();

  // Recap
  const recapEl = document.getElementById('end-prod-recap');
  recapEl.innerHTML = `
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">OF</div><div style="font-size:18px;font-weight:800">${f.of_num||'—'}</div></div>
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Durée</div><div style="font-size:18px;font-weight:800">${fmtS(st.of_elapsed_s||0)}</div></div>
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Arrêts</div><div style="font-size:18px;font-weight:800">${fmtS(st.stop_wall_s||0)}</div></div>
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Pauses</div><div style="font-size:18px;font-weight:800">${fmtS(st.pause_total_s||0)}</div></div>`;
  openModal('modal-end-prod');
}

function updateEpEquiv() {
  const qte = parseFloat(document.getElementById('ep-qtefab').value) || 0;
  // Simplified equiv (1:1 if no type prod selected)
  document.getElementById('ep-equiv').value = qte.toFixed(2);
}

async function doEndProd() {
  const form = collectForm();
  form.pilote = document.getElementById('ep-pilote').value;
  form.poste  = document.getElementById('ep-poste').value;
  form.qte_fab = document.getElementById('ep-qtefab').value;
  form.qte_emb = document.getElementById('ep-qteemb').value;
  form.comment = document.getElementById('ep-comment').value;
  form.duree_mq_mp = document.getElementById('ep-duree-mq-mp').value;
  form.manquant_pers = document.getElementById('ep-manquant-pers').value;
  const r = await fetch('/api/end_prod', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({form})});
  const d = await r.json();
  if (!d.ok) { alert(d.error); return; }
  closeModal('modal-end-prod');
  st.prod_active = false;
  document.body.classList.remove('stop-active');
  showRecap(d.recap);
  showView('main');
  setTimeout(() => { loadHistory(); loadEvents(); }, 800);
}

function showRecap(recap) {
  const trs = recap.trs >= 0 ? recap.trs + '%' : '—';
  const col = recap.trs >= 70 ? '#1a8c4e' : recap.trs >= 50 ? '#d97706' : '#e31e24';
  document.getElementById('recap-body').innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:16px">
      <div class="card" style="text-align:center"><div class="card-hdr">TRS</div><div style="font-size:40px;font-weight:900;color:${col}">${trs}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Qté fabriquée</div><div style="font-size:28px;font-weight:800">${recap.qte_fab}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Équivalence</div><div style="font-size:28px;font-weight:800">${recap.equiv}</div></div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px">
      <div class="card" style="text-align:center"><div class="card-hdr">OF</div><div style="font-size:16px;font-weight:800">${recap.of_num||'—'}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Pilote</div><div style="font-size:16px;font-weight:800">${recap.pilote}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Début</div><div style="font-size:16px;font-weight:800">${recap.debut}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Fin</div><div style="font-size:16px;font-weight:800">${recap.fin}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Durée OF</div><div style="font-size:16px;font-weight:800">${fmtS(recap.of_s)}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Pauses</div><div style="font-size:16px;font-weight:800">${fmtS(recap.pause_s)}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Arrêts</div><div style="font-size:16px;font-weight:800">${fmtS(recap.stop_wall_s)}</div></div>
      <div class="card" style="text-align:center"><div class="card-hdr">Qté emb.</div><div style="font-size:16px;font-weight:800">${recap.qte_emb}</div></div>
    </div>`;
  openModal('modal-recap');
}

function closeRecap() {
  closeModal('modal-recap');
}

// ─── Fin de poste ────────────────────────────────────────────────────────────
async function openFinPoste() {
  openModal('modal-fin-poste');
  const r = await fetch('/api/fin_poste_data');
  const d = await r.json();
  let html = `<div style="margin-bottom:16px;padding:14px;background:#f8fafc;border-radius:10px;display:grid;grid-template-columns:repeat(4,1fr);gap:12px">
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Pilote</div><div style="font-size:18px;font-weight:800">${d.pilot||'—'}</div></div>
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Date</div><div style="font-size:18px;font-weight:800">${d.date||'—'}</div></div>
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Nb OF</div><div style="font-size:18px;font-weight:800">${d.nb_of}</div></div>
    <div style="text-align:center"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">TRS Poste</div><div style="font-size:18px;font-weight:800;color:${d.trs>=70?'#1a8c4e':d.trs>=50?'#d97706':'#e31e24'}">${d.trs>=0?d.trs+'%':'—'}</div></div>
    <div style="text-align:center;grid-column:1/-1"><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">Équivalence totale</div><div style="font-size:20px;font-weight:800">${d.tot_equiv}</div></div>
  </div>`;
  if (d.of_list && d.of_list.length) {
    html += `<table class="fp-table"><thead><tr><th>OF</th><th>Taille</th><th>Type</th><th>Qté Fab</th><th>Qté Emb</th><th>Equiv</th><th>Début</th><th>Fin</th><th>TRS%</th></tr></thead><tbody>`;
    d.of_list.forEach(of => {
      const tc = of.trs>=70?'color:#1a8c4e':of.trs>=50?'color:#d97706':'color:#e31e24';
      html += `<tr><td>${of.of}</td><td>${of.taille}</td><td>${of.type_prod}</td><td>${of.qte_fab}</td><td>${of.qte_emb}</td><td>${of.equiv}</td><td>${of.debut}</td><td>${of.fin}</td><td style="${tc};font-weight:700">${of.trs>=0?of.trs+'%':'—'}</td></tr>`;
    });
    html += '</tbody></table>';
  }
  document.getElementById('fin-poste-body').innerHTML = html;
}

function validateFinPoste() {
  alert('Fin de poste validée. Bonne fin de journée !');
  closeModal('modal-fin-poste');
}

async function generateDashboard() {
  const r = await fetch('/api/generate_dashboard', {method:'POST'});
  const d = await r.json();
  if (d.ok) alert('Supervision générée : ' + d.path);
  else alert('Erreur : ' + d.error);
}

// ─── Open Stop Modal ──────────────────────────────────────────────────────────
function openStopModal() {
  document.getElementById('custom-stop-input').value = '';
  updateStopBtnStates();
  openModal('modal-stop');
}

// ─── Settings ─────────────────────────────────────────────────────────────────
async function openSettings() {
  const r = await fetch('/api/config');
  const cfg = await r.json();
  window._cfg = cfg;
  document.getElementById('set-prod-ref').value = cfg.prod_ref || 0;
  document.getElementById('set-pause-max').value = cfg.pause_max_min || 20;
  document.getElementById('set-clean-short').value = cfg.clean_short_min || 10;
  document.getElementById('set-clean-long').value = cfg.clean_long_min || 30;
  document.getElementById('set-clean-grand').value = cfg.clean_grand_min || 60;
  document.getElementById('set-meeting-tol').value = cfg.meeting_tol_min || 5;
  document.getElementById('set-db-path').value = cfg.db_path || '';
  document.getElementById('set-db-name').textContent = cfg.db_name ? '📄 ' + cfg.db_name : '';
  buildModelesList(cfg.modeles_horaires || []);
  buildPilotsList(cfg);
  openModal('modal-settings');
}

function buildModelesList(modeles) {
  const el = document.getElementById('modeles-list');
  el.innerHTML = '';
  modeles.forEach((m, i) => {
    const row = document.createElement('div');
    row.className = 'modele-row';
    row.innerHTML = `
      <input placeholder="Nom (ex: Matin)" value="${m.nom||''}" data-mi="${i}" data-f="nom">
      <input placeholder="Début (ex: 05:00)" value="${m.debut||''}" data-mi="${i}" data-f="debut">
      <input placeholder="Fin (ex: 13:00)" value="${m.fin||''}" data-mi="${i}" data-f="fin">
      <input placeholder="Durée OF max (h)" type="number" value="${m.duree_max||8}" data-mi="${i}" data-f="duree_max">
      <button class="btn btn-red btn-sm" onclick="removeModele(${i})">✕</button>`;
    el.appendChild(row);
  });
}

function addModele() {
  const cfg = window._cfg || {};
  const modeles = cfg.modeles_horaires || [];
  modeles.push({nom:'', debut:'', fin:'', duree_max: 8});
  cfg.modeles_horaires = modeles;
  window._cfg = cfg;
  buildModelesList(modeles);
}

function removeModele(i) {
  const cfg = window._cfg || {};
  const modeles = cfg.modeles_horaires || [];
  modeles.splice(i, 1);
  cfg.modeles_horaires = modeles;
  window._cfg = cfg;
  buildModelesList(modeles);
}

function buildPilotsList(cfg) {
  const el = document.getElementById('set-pilots-list');
  const pilots = (lists.pilotes || []);
  el.innerHTML = pilots.length ? pilots.map(p => `<div style="padding:4px 0;border-bottom:1px solid var(--lgray)">👤 ${p}</div>`).join('') : '<div style="color:var(--gray)">Aucun pilote dans le fichier Excel.</div>';
}

async function setDb() {
  const path = document.getElementById('set-db-path').value.trim();
  const pw = document.getElementById('set-cur-pw').value || (window._cfg||{}).supervisor_pw || '1234';
  const r = await fetch('/api/set_db', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({pw, path})});
  const d = await r.json();
  if (d.ok) {
    document.getElementById('set-db-name').textContent = '✔ ' + path.split(/[\\/]/).pop();
    setTimeout(() => { init(); }, 1200);
  } else alert(d.error);
}

async function reloadLists() {
  await fetch('/api/reload', {method:'POST'});
  const lr = await fetch('/api/lists');
  lists = await lr.json();
  alert('Listes rechargées.');
}

function switchSetTab(i) {
  document.querySelectorAll('.set-tab').forEach((t,j) => t.classList.toggle('active', j===i));
  document.querySelectorAll('.set-panel').forEach((p,j) => p.classList.toggle('active', j===i));
  setTabIdx = i;
}

async function saveSettings() {
  const cfg = window._cfg || {};
  // Collect modeles
  const modeles = [];
  document.querySelectorAll('.modele-row').forEach(row => {
    const m = {};
    row.querySelectorAll('input[data-f]').forEach(inp => { m[inp.dataset.f] = inp.value; });
    if (m.nom) modeles.push(m);
  });
  const payload = {
    pw: document.getElementById('set-cur-pw').value || '1234',
    prod_ref: document.getElementById('set-prod-ref').value,
    pause_max_min: document.getElementById('set-pause-max').value,
    clean_short_min: document.getElementById('set-clean-short').value,
    clean_long_min: document.getElementById('set-clean-long').value,
    clean_grand_min: document.getElementById('set-clean-grand').value,
    meeting_tol_min: document.getElementById('set-meeting-tol').value,
    modeles_horaires: modeles,
  };
  const newPw = document.getElementById('set-new-pw').value;
  if (newPw) payload.new_pw = newPw;
  const r = await fetch('/api/settings', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify(payload)});
  const d = await r.json();
  if (d.ok) {
    document.getElementById('set-save-ok').textContent = '✔ Enregistré';
    setTimeout(() => document.getElementById('set-save-ok').textContent = '', 2000);
  } else {
    alert(d.error || 'Erreur');
  }
}

// ─── Modal helpers ────────────────────────────────────────────────────────────
function openModal(id) {
  document.getElementById(id).classList.add('active');
}
function closeModal(id) {
  document.getElementById(id).classList.remove('active');
}

// ─── Utilities ────────────────────────────────────────────────────────────────
function fmtS(s) {
  const t = Math.max(0, Math.round(s || 0));
  return String(Math.floor(t/3600)).padStart(2,'0') + ':' +
         String(Math.floor((t%3600)/60)).padStart(2,'0') + ':' +
         String(t%60).padStart(2,'0');
}

function trsColor(t) {
  if (t < 0) return '#94a3b8';
  if (t >= 70) return '#1a8c4e';
  if (t >= 50) return '#d97706';
  return '#e31e24';
}

// Close modal on backdrop click
document.querySelectorAll('.modal').forEach(m => {
  m.addEventListener('click', e => { if (e.target === m) closeModal(m.id); });
});

// ─── Start ────────────────────────────────────────────────────────────────────
init();
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
