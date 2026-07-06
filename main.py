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
            # Load pilot passwords from col A (pilote names) + col B (MDP) — Excel is source of truth
            pil_map = {}
            for ri in range(2, ws.max_row+1):
                pil_v = ws.cell(ri, 1).value
                pw_v = ws.cell(ri, 2).value
                if pil_v and str(pil_v).strip():
                    pil_map[str(pil_v).strip()] = str(pw_v or "").strip()
            if pil_map:
                cfg["pilot_passwords"] = pil_map
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

POSTES_HEADERS = ["Date","Pilote","Co-Pilote","Poste","Nb OF","Prod Totale (equiv)","TRS Poste %","TRS Prod %","Total Arrets (min)","Total Pauses (min)","Nettoyage (min)","Commentaire"]

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
                    round(float(data.get("tot_equiv",0) or 0),1),
                    data.get("trs_shift",""),
                    data.get("trs_of",""),
                    round(float(data.get("arret_min",0) or 0),1),
                    round(float(data.get("pause_min",0) or 0),1),
                    round(float(data.get("nett_min",0) or 0),1),
                    data.get("comment",""),
                ]
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
        "equivalences": get_list("Equivalence coef") or get_list("Equivalence") or [],
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
/* main 3-col layout */
.prod-body{display:flex;flex:1;overflow:hidden}
/* LEFT actions col */
.act-col{width:190px;flex-shrink:0;background:var(--card);border-right:1px solid var(--border);display:flex;flex-direction:column;padding:8px;gap:6px}
.big-stop-btn{width:100%;background:linear-gradient(135deg,#b91c1c,#7f0000);color:#fff;border:none;border-radius:10px;padding:0;cursor:pointer;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;flex:0 0 90px;font-size:13px;font-weight:800;box-shadow:0 4px 12px rgba(185,28,28,.35);transition:all .15s;letter-spacing:.3px}
.big-stop-btn:hover{filter:brightness(.9);transform:translateY(-1px)}
.big-stop-btn .ico{font-size:22px}
.act-btn{width:100%;border:none;border-radius:8px;padding:8px 4px;cursor:pointer;font-size:11px;font-weight:700;text-align:center;transition:all .15s}
.act-btn:hover{filter:brightness(.9)}
.act-nett{background:#e0f2fe;color:var(--blue)}
.act-pause{background:#f3e8ff;color:var(--purple)}
.act-spacer{flex:1}
.act-endprod{background:linear-gradient(135deg,#d97706,#b45309);color:#fff;font-size:12px;font-weight:800;padding:10px 4px;border-radius:8px;box-shadow:0 3px 8px rgba(217,119,6,.3)}
/* CENTER form col */
.form-col{flex:1;overflow-y:auto;padding:8px;display:flex;flex-direction:column;gap:6px}
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
input:not([type=checkbox]):not([type=radio]):not([type=button]):not([type=submit]),textarea{cursor:text}
select{cursor:default}
.fr.big input{font-size:16px;font-weight:700;padding:5px 6px;color:var(--green)}
.fr.ro input{background:#f8fafc;color:var(--gray)}
/* Timeline */
.tl-wrap{background:var(--card);border-radius:7px;padding:7px 8px;border:1px solid var(--border)}
.tl-wrap h5{font-size:9px;text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.7px;margin-bottom:4px}
/* RIGHT recap col */
.recap-col{width:215px;flex-shrink:0;border-left:1px solid var(--border);background:var(--card);display:flex;flex-direction:column;overflow:hidden}
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

@media(max-width:900px){.form-3col{grid-template-columns:1fr 1fr}.recap-col{width:130px}.act-col{width:110px}}
@media(max-width:650px){.form-3col{grid-template-columns:1fr}.prod-body{flex-direction:column}.act-col,.recap-col{width:100%;flex-direction:row;flex-wrap:wrap}.act-col{height:auto}}
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
      <select id="ln-model"><option value="">-- Choisir --</option></select>
    </div>
    <div class="lf">
      <label>Mot de passe</label>
      <input type="password" id="ln-pw" placeholder="••••" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="btn-login" onclick="doLogin()">Valider</button>
    <div class="ln-err" id="ln-err"></div>
  </div>
</div>

<!-- ════ APP ════ -->
<div id="app" class="hidden" style="display:none;flex:1;flex-direction:column;overflow:hidden">
  <div id="app-hdr">
    <div class="hdr-logo">⚙ KPI-ORC</div>
    <div class="hdr-tabs">
      <button class="htab on" id="ht-main" onclick="goTab('main')">Accueil</button>
      <button class="htab prod-on" id="ht-prod" style="display:none" onclick="goTab('prod')">▶ Prod en cours</button>
      <button class="htab" id="ht-hist" onclick="goTab('history')">Historique</button>
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
    <div class="main-hdr">
      <div class="mbtns" style="margin-left:0">
        <button class="btn btn-green" id="btn-start" onclick="doStartProd()" style="font-size:14px;padding:10px 18px;font-weight:800">▶ Démarrer production</button>
        <button class="btn btn-amber" onclick="doFinPoste()" style="font-size:14px;padding:10px 18px;font-weight:800">🏁 Fin de poste</button>
        <button class="btn btn-sec" onclick="loadMainDecl()" style="font-size:12px;padding:8px 14px">↺ Actualiser</button>
      </div>
    </div>
    <!-- KPI 3 derniers postes -->
    <div class="shift-kpis">
      <div class="skpi current">
        <div class="sk-lbl">Poste actuel — TRS</div>
        <div class="sk-val" id="kpi0-trs">--%</div>
        <div class="sk-sub" id="kpi0-sub">0 OF</div>
      </div>
      <div class="skpi">
        <div class="sk-lbl" id="kpi1-lbl">Hier</div>
        <div class="sk-val" id="kpi1-trs">--%</div>
        <div class="sk-sub" id="kpi1-sub">0 OF</div>
      </div>
      <div class="skpi">
        <div class="sk-lbl" id="kpi2-lbl">Avant-hier</div>
        <div class="sk-val" id="kpi2-trs">--%</div>
        <div class="sk-sub" id="kpi2-sub">0 OF</div>
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
      <div class="pob-item trs">
        <div class="pob-lbl">TRS estimé</div>
        <div class="pob-val" id="pob-trs">—</div>
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
      <!-- LEFT: action buttons -->
      <div class="act-col">
        <button class="big-stop-btn" onclick="openStopModal()">
          <span class="ico">⛔</span>
          <span>Déclarer<br>un arrêt</span>
        </button>
        <button class="act-btn act-nett" onclick="doNettoyage()">🧹 Nettoyage</button>
        <button class="act-btn act-pause" id="btn-pause" onclick="doPause()">⏸ Pause pilote</button>
        <div class="act-spacer"></div>
        <button class="act-btn act-endprod" onclick="doEndProdPreview()">🏁 Fin de<br>production</button>
      </div>
      <!-- CENTER: form -->
      <div class="form-col">
        <div class="form-3col">
          <!-- Zone Identification -->
          <div class="fzone zi">
            <h4>📋 Identification</h4>
            <div class="fr"><label>N° OF *</label><input id="f-of_num" oninput="scheduleAutoSave()"></div>
            <div class="fr ro"><label>Date</label><input id="f-date" readonly></div>
            <div class="fr ro"><label>Poste</label><input id="f-poste" readonly></div>
            <div class="fr ro"><label>Pilote</label><input id="f-pilote" readonly></div>
            <div class="fr"><label>Co-Pilote</label><input id="f-copilote" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Personnes</label><input id="f-nb_pers" type="number" min="1" value="2" oninput="scheduleAutoSave()"></div>
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
            <div class="fr"><label>Traca Fibre</label><select id="f-traca" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
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
        </div>
      </div>
      <!-- RIGHT: recap + gauge -->
      <div class="recap-col">
        <div class="recap-hdr">Arrêts / pauses</div>
        <div class="recap-body" id="recap-list"></div>
        <div class="gauge-box">
          <svg viewBox="0 0 100 56" style="width:100%;max-width:120px">
            <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="10" stroke-linecap="round"/>
            <path id="gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="10" stroke-linecap="round" stroke-dasharray="0,1000"/>
            <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e" id="gauge-pct">—</text>
          </svg>
          <div class="gauge-lbl">TRS estimé</div>
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
      <div id="fp-date" style="font-size:12px;font-weight:700;opacity:.9"></div>
    </div>
    <!-- Graphiques + KPI (en haut, compact) -->
    <div style="display:flex;gap:8px;padding:8px 12px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;align-items:center;flex-wrap:wrap">
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px">TRS Poste</div>
        <svg id="fp-gauge" viewBox="0 0 100 58" style="width:80px;display:block;margin:0 auto">
          <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
          <path id="fp-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
          <text x="50" y="46" text-anchor="middle" font-size="14" font-weight="800" fill="#1a1f5e" id="fp-gauge-pct">--%</text>
        </svg>
        <div style="font-size:10px;font-weight:700;margin-top:2px" id="fp-trs-lbl2">—</div>
      </div>
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px">Répartition</div>
        <svg id="fp-pie" viewBox="0 0 130 115" style="width:80px;height:71px;display:block;margin:0 auto"></svg>
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
      <svg id="fp-tl" viewBox="0 0 800 32" preserveAspectRatio="none" style="width:100%;height:32px;display:block">
        <rect x="0" y="2" width="800" height="24" fill="#e2e8f0" rx="4"/>
      </svg>
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
    </div>
    <!-- Boutons -->
    <div style="padding:8px 12px;background:var(--card);border-top:1px solid var(--border);flex-shrink:0;display:flex;gap:10px;justify-content:flex-end">
      <button class="btn btn-sec" onclick="goTab('main')">← Retour</button>
      <button class="btn btn-danger btn-lg" onclick="confirmFinPoste()">⏹ Confirmer fin de poste &amp; Déconnexion</button>
    </div>
  </div>

  <!-- ════ HISTORY ════ -->
  <div id="v-history" class="view" style="flex-direction:column;overflow:hidden">
    <div style="background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0">
      <span style="font-size:12px;font-weight:700;color:var(--navy)">Historique</span>
      <input type="date" id="hist-dt" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:12px" onchange="loadHist()">
    </div>
    <div style="flex:1;overflow-y:auto">
      <table class="ktbl"><thead><tr id="hist-hd"></tr></thead><tbody id="hist-bd"></tbody></table>
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
      <div class="stop-section-lbl">✏ Arrêt libre / autre</div>
      <div class="custom-row">
        <input id="custom-stop-input" placeholder="Nom de l'arrêt…" maxlength="60">
        <button class="btn btn-amber" onclick="declareCustomStop()">Déclarer</button>
      </div>
    </div>
    <div class="mftr"><button class="btn btn-sec" onclick="closeM('m-stop')">Annuler</button></div>
  </div>
</div>

<!-- ════ MODAL: Fin de production ════ -->
<div class="overlay" id="m-endprod">
  <div class="mbox wide">
    <div class="mhdr"><h2 id="ep-title">⏹ Fin de production</h2></div>
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
      <div id="er-pw-row" style="margin-bottom:8px">
        <div class="fr" style="max-width:200px"><label>MDP Admin</label><input type="password" id="er-pw" placeholder="••••"></div>
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
  ratt:"#7c3aed",pb:"#b91c1c",nettoyage:"#0891b2",
  "Pause pilote":"#7c3aed","_pause":"#7c3aed"
};

function getStopColor(key, cat) {
  if (cat) return STOP_COL[cat]||'#64748b';
  if (key==='nettoyage') return STOP_COL.nettoyage;
  if (key==='_pause'||key==='Pause pilote') return STOP_COL['Pause pilote'];
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
  buildStopGrids();
  buildEditStopOpts();
  await loadLists();
  const s = await apiFetch('/api/state');
  if (s && s.pilot) {
    document.getElementById('v-login').classList.remove('on');
    showApp(s);
  }
  setInterval(pollState, 5000);
  setInterval(pollEvts, 8000);
  startTicker();
});

async function loadLists() {
  const d = await apiFetch('/api/lists');
  if (!d) return;
  popSel('f-taille', d.tailles||[]);
  popSel('f-type_prod', d.types_prod||[]);
  popSel('f-fibre', d.fibres||[]);
  popSel('f-traca', d.tracas||[]);
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
  pil.forEach(p => { const o=document.createElement('option'); o.value=p; o.textContent=p; sel.appendChild(o); });
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
    const s = await apiFetch('/api/state');
    document.getElementById('v-login').classList.remove('on');
    showApp(s||{pilot,poste});
  } else {
    document.getElementById('ln-err').textContent = d.error||'Erreur connexion';
  }
}

function showApp(s) {
  const app = document.getElementById('app');
  app.style.display = 'flex';
  app.classList.remove('hidden');
  if (s.pilot) { document.getElementById('f-pilote').value=s.pilot; document.getElementById('pob-pilot').textContent=s.pilot; }
  if (s.poste) { document.getElementById('f-poste').value=s.poste; document.getElementById('pob-poste').textContent=s.poste; }
  setToday();
  restoreFormFromStorage();
  pollState();
  pollEvts();
  loadCfg();
  goTab(s.prod_active ? 'prod' : 'main');
  document.getElementById('hist-dt').value = new Date().toISOString().slice(0,10);
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
  const vm={main:'v-main',prod:'v-prod',history:'v-history',settings:'v-settings',finposte:'v-finposte'};
  const el=document.getElementById(vm[tab]);
  if(el) el.classList.add('on');
  const nt={main:'ht-main',prod:'ht-prod',history:'ht-hist',settings:'ht-cfg'};
  const ntEl=document.getElementById(nt[tab]);
  if(ntEl) ntEl.classList.add('on');
  if(tab==='history') loadHist();
  if(tab==='finposte') loadFPData();
  if(tab==='main') { loadMainDecl(); }
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
  if(tp) tp.style.display=s.prod_active?'':'none';

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
  if(pbtn) pbtn.textContent=s.is_paused?'▶ Reprendre':'⏸ Pause pilote';

  // TRS gauge
  updateGauge(s);
}

function getEvtLabel(key) {
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
    html+=`<div class="stop-chip"><span class="chip-lbl">⏸ Pause pilote</span><span class="chip-tim" id="chip-t-_pause">${fmtDur2(pe)}</span><button class="btn-endstop" onclick="doPause()">▶ Reprendre</button></div>`;
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
    // Équivalence théorique (basée sur type produit)
    const thEl=document.getElementById('sc-theo');
    if(thEl&&ST.prod_ref){
      const typeProd=ST.form&&ST.form.type_prod||'';
      const coef=(window._equivCoefs&&window._equivCoefs[typeProd])||1;
      const theo=Math.round(ST.prod_ref*(_ofElapAtPoll+dt)/28800*coef);
      thEl.textContent=theo>0?theo+' éq.':'—';
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
  const decls=(Array.isArray(declData)?declData:(declData&&declData.rows?declData.rows:[])).filter(r=>!r.date||r.date.startsWith(todayPfx));
  const evts=(Array.isArray(evtData)?evtData:[]).filter(r=>!r.date||r.date.startsWith(todayPfx));
  const allRows=[];
  decls.forEach(r=>allRows.push({...r,_rowType:'prod'}));
  evts.forEach(r=>allRows.push({...r,_rowType:'evt'}));
  allRows.sort((a,b)=>(b.debut||'').localeCompare(a.debut||''));
  const bd=document.getElementById('main-body');
  if(!bd) return;
  if(!allRows.length){bd.innerHTML='<tr><td colspan="10" style="text-align:center;color:var(--gray);padding:16px">Aucune déclaration aujourd\'hui</td></tr>';loadMainKPI();return;}
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
}

async function loadMainKPI() {
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
  if(d){
    const trs=d.trs_shift!==undefined?d.trs_shift:d.trs;
    const el0t=document.getElementById('kpi0-trs'),el0s=document.getElementById('kpi0-sub');
    if(el0t) el0t.textContent=fmtTRSv(trs);
    if(el0s) el0s.textContent=(d.rows?d.rows.length:0)+' OF | Arrêts '+Math.round(curStopS/60)+' min';
  }

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
      lbl.textContent=s.pilot+' — '+s.poste;
      tv.textContent=fmtTRSv(avgT);
      sv.textContent=s.rows.length+' OF | Arrêts '+Math.round(sStopS/60)+' min';
    } else {
      lbl.textContent=i===0?'Poste précédent':'Avant-dernier';
      tv.textContent='--%'; sv.textContent='—';
    }
  }
}

// ── START PROD ──
async function doStartProd() {
  saveFormToStorage(); // save current form before clearing
  const r=await fetch('/api/start_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  if(!r) return;
  const d=await r.json();
  if(d.ok){
    // DON'T clear form - restore from storage or keep current
    setToday();
    restoreFormFromStorage();
    await pollState();
    await pollEvts();
    goTab('prod');
  } else toast(d.error||'Erreur','err');
}

// ── STOP/PAUSE ──
function buildStopGrids() {
  const rattGrid=document.getElementById('sgrid-ratt');
  const pbGrid=document.getElementById('sgrid-pb');
  if(!rattGrid||!pbGrid) return;
  EVENTS.filter(e=>e[2]==='ratt').forEach(e=>{
    const b=document.createElement('button');
    b.className='stop-btn ratt'; b.textContent=e[0];
    b.onclick=()=>{closeM('m-stop');doStartStop(e[1],'ratt');};
    rattGrid.appendChild(b);
  });
  EVENTS.filter(e=>e[2]==='pb').forEach(e=>{
    const b=document.createElement('button');
    b.className='stop-btn pb'; b.textContent=e[0];
    b.onclick=()=>{closeM('m-stop');doStartStop(e[1],'pb');};
    pbGrid.appendChild(b);
  });
}

function openStopModal(){openM('m-stop');}

async function doPause(){
  await fetch('/api/toggle_pause',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
}

async function doNettoyage(){
  await fetch('/api/start_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ntype:'court'})});
  await pollState();
}

async function doStartStop(key,cat){
  await fetch('/api/start_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,cat})});
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
  if(!r||!r.ok){if(confirm('Confirmer fin de production?'))confirmEndProd();return;}
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
    try{localStorage.removeItem('kpiorc_form');}catch(e){}
    await pollState();
    await pollEvts();
    goTab('main');
    toast('Production enregistrée','ok');
  } else toast(d.error||'Erreur','err');
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
  // Estimate live TRS
  const prodRef=s.prod_ref||200;
  const ofS=s.of_elapsed_s||0;
  const qFab=s.form?parseFloat(s.form.qte_fab||0):0;
  const taille=s.form?s.form.taille:'';
  let trs=-1;
  // Can't compute equiv client-side, show from state if available
  if(ofS>0&&prodRef>0&&qFab>0){
    trs=Math.round(qFab/(prodRef*ofS/28800)*100*10)/10;
  }
  const trsStr=trs>=0?fmtTRS(trs):'—';
  pct.textContent=trsStr;
  if(pobTrs) pobTrs.textContent=trsStr;
  // Arc: 0-100% maps to 0-132 (half circle perimeter ≈ π*42 ≈ 132)
  const pArc=132;
  const dash=trs>=0?Math.min(1,trs/100)*pArc:0;
  const col=trs>=90?'#16a34a':trs>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash},${pArc}`);
  arc.setAttribute('stroke',col);
}

// ── EDIT ROW (accueil) ──
function openEditRow(key) {
  const row=window._rowMap[String(key)];if(!row) return;
  const isProd=row._rowType==='prod';
  document.getElementById('er-rownum').value=row.row_num||'';
  document.getElementById('er-rowtype').value=row._rowType||'';
  document.getElementById('er-title').textContent=isProd?'✏ Modifier déclaration':'✏ Modifier événement';
  const pwRow=document.getElementById('er-pw-row'),pwEl=document.getElementById('er-pw');
  if(_adminPw){pwEl.value=_adminPw;pwRow.style.display='none';}else{pwEl.value='';pwRow.style.display='';}
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
  const pw=document.getElementById('er-pw').value||_adminPw;
  if(!rowNum){toast('Ligne invalide','err');return;}
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
  if(d&&d.ok){closeM('m-editrow');await loadMainDecl();toast('Ligne modifiée','ok');}
  else toast(d?.error||'Erreur modification','err');
}

async function deleteRow(key,rowNumId) {
  const rn=rowNumId?parseInt(document.getElementById(rowNumId)?.value):parseInt(window._rowMap[String(key)]?.row_num);
  if(!rn||!confirm('Supprimer cette ligne ?')) return;
  const pw=document.getElementById('er-pw')?.value||_adminPw||'';
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:rn})});
  const d=r?await r.json():{};
  if(d&&d.ok){closeM('m-editrow');await loadMainDecl();toast('Supprimé','ok');}
  else toast(d?.error||'Erreur suppression','err');
}

// ── EDIT STOP (recap prod view) ──
function buildEditStopOpts(){
  const s=document.getElementById('es-type');
  const s2=document.getElementById('er-evttype');
  const allOpts=[['Pause pilote','_pause'],['Nettoyage','nettoyage'],...EVENTS.map(e=>[e[0],e[0]]),['Arrêt libre','Arrêt libre']];
  [s,s2].forEach(sel=>{
    if(!sel) return;
    allOpts.forEach(([lbl,val])=>{const o=document.createElement('option');o.value=val;o.textContent=lbl;sel.appendChild(o);});
  });
  // es-type uses key values (not display)
  if(s){s.innerHTML='';[['Pause pilote','_pause'],['Nettoyage','nettoyage'],...EVENTS,['Arrêt libre','autre']].forEach(([lbl,key])=>{const o=document.createElement('option');o.value=key;o.textContent=lbl;s.appendChild(o);});}
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
  if(r&&r.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();toast('Modifié','ok');}
  else toast('Erreur','err');
}

async function deleteStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev||!confirm('Supprimer ?')) return;
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({row_num:ev.row_num})});
  if(r&&r.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();toast('Supprimé','ok');}
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

function drawTLFromISO(svgId,evts,startIso,endIso){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=28,H=40;
  let html=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  const tS=new Date(startIso).getTime(),tE=new Date(endIso).getTime();
  const span=tE-tS;if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  // Prod background — show for current prod AND keep events from previous prods visible
  if(ST.of_start_iso){
    const ps=new Date(ST.of_start_iso).getTime();
    const pe=ST.prod_active?tE:(ST.last_of_end_iso?new Date(ST.last_of_end_iso).getTime():tE);
    const x1=toX(ps),x2=toX(pe);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${ST.prod_active?'#bbf7d0':'#e0f2fe'}" rx="4"/>`;
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
  html+=`<text x="2" y="${H-1}" font-size="8" fill="#64748b">${fT(tS)}</text>`;
  html+=`<text x="${W-30}" y="${H-1}" font-size="8" fill="#64748b">${fT(tE)}</text>`;
  html+=`<line x1="${W/2}" y1="${Y}" x2="${W/2}" y2="${Y+H2}" stroke="#94a3b8" stroke-width=".5" stroke-dasharray="2,2"/>`;
  html+=`<text x="${W/2-10}" y="${H-1}" font-size="8" fill="#94a3b8">${fT((tS+tE)/2)}</text>`;
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
    const lbl=_curStopKey==='_pause'?'Pause pilote':getEvtLabel(_curStopKey);
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

  // Timeline
  const shiftStart=d.shift_start_iso||new Date(Date.now()-8*3600*1000).toISOString();
  drawTLFromISO('fp-tl',gEvts,shiftStart,new Date().toISOString());

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
  if(shiftInfo) shiftInfo.textContent=jour.debut&&jour.fin?`${jour.debut}→${jour.fin}`:'—';
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
  // Gather stop/pause/nettoyage totals from gEvts
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
  const posteRow={
    date:new Date().toLocaleDateString('fr-FR'),
    pilot:ST.pilot||'',
    copilote:ST.form&&ST.form.copilote||'',
    poste:ST.poste||'',
    nb_of:fpData?fpData.nb_of||0:0,
    tot_equiv:fpData?fpData.tot_equiv||0:0,
    trs_shift:fpData?fpData.trs_shift||0:0,
    trs_of:fpData?fpData.trs||0:0,
    arret_min:Math.round(arret_s/60),
    pause_min:Math.round(pause_s/60),
    nett_min:Math.round(nett_s/60),
    comment:''
  };
  await fetch('/api/save_poste',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(posteRow)});
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  resetToLogin();
  toast('Bonne fin de poste !','ok');
}

// ── HISTORY ──
async function loadHist(){
  const dt=document.getElementById('hist-dt').value||new Date().toISOString().slice(0,10);
  const d=await apiFetch('/api/history?date='+dt);
  const rows=Array.isArray(d)?d:(d&&d.rows?d.rows:[]);
  const hd=document.getElementById('hist-hd'),bd=document.getElementById('hist-bd');
  if(!hd||!bd) return;
  if(!rows.length){bd.innerHTML='<tr><td colspan="9" style="text-align:center;color:var(--gray);padding:16px">Aucune donnée</td></tr>';return;}
  const ks=['type','of','poste','pilote','debut','fin','qte_fab','equiv','trs'];
  const lb={type:'Type',of:'OF',poste:'Poste',pilote:'Pilote',debut:'Début',fin:'Fin',qte_fab:'Qté',equiv:'Éq',trs:'TRS'};
  hd.innerHTML=ks.map(k=>`<th>${lb[k]||k}</th>`).join('');
  bd.innerHTML=rows.map(row=>{
    const t=parseFloat(row.trs||0);
    return '<tr>'+ks.map(k=>{const v=row[k]||'';
    if(k==='trs') return `<td class="${t>=90?'tg':t>=75?'tm':t>0?'tb':''}">${t>0?fmtTRS(t):''}</td>`;
    return `<td>${esc(String(v))}</td>`;}).join('')+'</tr>';
  }).join('');
}

// ── SETTINGS ──
async function loadCfg(){
  const d=await apiFetch('/api/config');
  if(!d) return;
  _cfgPwds=d.pilot_passwords||{};
  _cfgModels=d.modeles_horaires||[];
  const prEl=document.getElementById('cfg-pr');
  if(prEl) prEl.value=d.prod_ref||200;
  renderPwdList();
  renderModelList();
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
  if(st) st.textContent='Génération…';
  const r=await fetch('/api/generate_dashboard',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  const d=r?await r.json():{};
  if(d&&d.ok){if(st) st.textContent='Dashboard généré : '+esc(d.path||'');}
  else {if(st) st.textContent='Erreur: '+(d&&d.error||'inconnue');}
}

// ── MODALS ──
function openM(id){const m=document.getElementById(id);if(m){m.classList.add('on');m.style.display='flex';}}
function closeM(id){const m=document.getElementById(id);if(m){m.classList.remove('on');m.style.display='';}}
document.addEventListener('click',e=>{if(e.target.classList.contains('overlay'))closeM(e.target.id);});

// ── UTILS ──
async function apiFetch(url){try{const r=await fetch(url);return r.ok?await r.json():null;}catch(e){return null;}}
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
