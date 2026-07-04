"""KPI-ORC v6.1 - Flask + pywebview (design HTML, logique v5.81 conservée)"""
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

DATA_HEADERS = [
    "OF","Date","Poste","Pilote","Co-Pilote","Nb Personnes",
    "Taille","Code Produit","Type Produit","Poids Garnissage","Fibre",
    "OF Taie","Traca Fibre","Qte Fabriquee","Qte Emballee","Equivalence",
    "Duree OF","Heure Debut","Heure Fin","Cadence/heure","Cadence/h/pers",
    "Kit","Ref Taie","Qte Initiale Taie","Nb Taie 2nd Choix",
    "Nb Defaut Couture","Mq Taie","Mq Housse/Encart",
    "Nb PP Cousue","Changement de Serie",
    "Temps Arret Manquant MP","Temps Arret Manquant Personnel/Reunion",
    "Nettoyage Fin de Poste",
    "Ratt Pochon/Fibre","Ratt Couture","Ratt Emballage",
    "Ratt Presse Souder","Ratt Presse ZIP",
    "PB Chargeuse","PB Carde","PB Etaleur/Tour","PB Coupe/Circ",
    "PB Tapis Bascule","PB Enrouleur Pochon","PB Pesee/Tapis 2",
    "PB Deviation/Table","PB Enfileur Pochon","PB Kinna/Stroebel",
    "PB Tapeuse","PB Table Rot/Twin","PB Enfileuse H1",
    "PB Enfileuse Traversin","PB Presse ORC","PB Presse Housse ZIP",
    "PB Cercleuse","PB Enrouleuse Traversin",
    "Commentaire","Temps Interposte",
]

EVT_HEADERS = [
    "Evenement","OF","Date","Poste","Pilote","Co-Pilote",
    "Nb Personnes","Taille","Type Produit","Code Produit","Fibre",
    "Poids Garnissage","OF Taie","Traca Fibre","Ref Taie","Kit",
    "Heure Debut","Heure Fin","Duree","Commentaire","Prevu/Hors TRS",
]

POSTES = ["Matin","Midi","Nuit","Jour"]

# ── État global ───────────────────────────────────────────────────────────────
_S = {
    "pilot": None, "poste": None,
    "prod_active": False,
    "of_start": None,
    "last_of_end": None,
    "last_of_pilot": "",
    "last_of_modele": None,
    "last_of_poste": None,
    "inter_of_s": 0.0,
    "interposte_s": 0.0,
    "timers": {},
    "tl_events": [],
    "of_periods": [],
    "of_changes": [],
    "form": {},
    "is_paused": False,
    "pause_start": None,
    "pause_total_s": 0.0,
    "pause_periods": [],
    "of_count_shift": 0,
}
_excel_lock = threading.Lock()
_lists = {}
_data_rows_cache = []
_events_cache = []
_prod_ref_cached = 0.0
cfg = {}

# ── Utilitaires ───────────────────────────────────────────────────────────────
def fmt(seconds):
    h, r = divmod(int(max(0, seconds)), 3600)
    m, s = divmod(r, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def _hms_to_sec(s):
    try:
        p = str(s).split(":")
        return int(p[0])*3600 + int(p[1])*60 + int(p[2])
    except Exception:
        return 0

def _row_date(v):
    if v is None: return ""
    if hasattr(v, 'strftime'): return v.strftime("%d/%m/%Y")
    s = str(v).strip()[:10]
    if len(s)==10 and s[4]=='-':
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s

def _row_time(v):
    if v is None: return ""
    if hasattr(v, 'strftime'): return v.strftime("%H:%M:%S")
    return str(v).strip()

def _min_str_to_hms(val):
    if not val or str(val).strip()=="": return ""
    try:
        total_s = int(float(str(val).replace(",",".")) * 60)
        h,r = divmod(total_s,3600); m,s = divmod(r,60)
        return f"{h:02d}:{m:02d}:{s:02d}"
    except Exception:
        return str(val)

def _n(v):
    try: return int(str(v).strip() or 0)
    except: return 0

def _dt_str(dt):
    return dt.isoformat() if dt else None

def _str_dt(s):
    if not s: return None
    try: return datetime.datetime.fromisoformat(s)
    except: return None

def load_cfg():
    try:
        with open(CONFIG_FILE) as f: return json.load(f)
    except: return {"db_path":"","supervisor_pw":"1234","prod_ref":0,
                    "pause_max_min":20,"clean_short_min":10,"clean_long_min":30,
                    "clean_grand_min":60,"meeting_tol_min":5}

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
    t = _S["timers"].get(key,{"elapsed":0.0,"running":False,"start":None})
    total = t["elapsed"]
    if t["running"] and t["start"]:
        total += (datetime.datetime.now()-t["start"]).total_seconds()
    return total

def t_running(key):
    return _S["timers"].get(key,{}).get("running",False)

def t_stop_all():
    for k in list(_S["timers"]): t_stop(k)

def t_reset():
    _S["timers"].clear()

def t_wall_clock_stops():
    if not _S["of_start"]: return 0.0
    now = datetime.datetime.now()
    intervals = []
    for ev in _S["tl_events"]:
        if ev.get("cat") not in ("ratt","pb"): continue
        if ev.get("hors_trs"): continue
        if _S["of_start"] and ev["start"] < _S["of_start"] and ev.get("key")!="arret_interposte": continue
        intervals.append((ev["start"], ev.get("end") or now))
    if not intervals: return 0.0
    intervals.sort(key=lambda x:x[0])
    merged,cs,ce = [],intervals[0][0],intervals[0][1]
    for s,e in intervals[1:]:
        if s<=ce: ce=max(ce,e)
        else: merged.append((cs,ce)); cs,ce=s,e
    merged.append((cs,ce))
    return sum((e-s).total_seconds() for s,e in merged)

# ── Timeline events ───────────────────────────────────────────────────────────
def tl_open(key, cat):
    of_num = _S["form"].get("of_num","")
    _S["tl_events"].append({
        "key":key,"cat":cat,
        "start":datetime.datetime.now(),"end":None,
        "comment":"","of_num":of_num,
        "pilot":_S["pilot"] or "",
    })
    save_session()

def tl_close(key, comment="", end_time=None):
    for ev in reversed(_S["tl_events"]):
        if ev["key"]==key and ev["end"] is None:
            ev["end"] = end_time or datetime.datetime.now()
            ev["comment"] = comment
            break
    save_session()

def tl_close_all():
    now = datetime.datetime.now()
    for ev in _S["tl_events"]:
        if ev["end"] is None: ev["end"] = now
    save_session()

def serialize_event(ev):
    return {
        "key": ev["key"], "cat": ev["cat"],
        "start": ev["start"].isoformat() if ev.get("start") else None,
        "end": ev["end"].isoformat() if ev.get("end") else None,
        "comment": ev.get("comment",""),
        "of_num": ev.get("of_num",""),
        "pilot": ev.get("pilot",""),
        "hors_trs": ev.get("hors_trs",False),
        "nettoyage_type": ev.get("nettoyage_type",""),
    }

# ── Session ───────────────────────────────────────────────────────────────────
def save_session():
    if not _S["prod_active"]:
        try: os.remove(SESSION_FILE)
        except: pass
        return
    try:
        timers_s = {}
        for k,t in _S["timers"].items():
            timers_s[k]={"elapsed":t["elapsed"],"running":t["running"],"start":_dt_str(t.get("start"))}
        events_s = []
        for ev in _S["tl_events"]:
            events_s.append({
                "key":ev["key"],"cat":ev["cat"],
                "start":_dt_str(ev["start"]),"end":_dt_str(ev.get("end")),
                "comment":ev.get("comment",""),
                "hors_trs":ev.get("hors_trs",False),
                "nettoyage_type":ev.get("nettoyage_type",""),
                "of_num":ev.get("of_num",""),
                "pilot":ev.get("pilot",""),
            })
        periods_s = [{"start":_dt_str(p["start"]),"end":_dt_str(p.get("end")),"of_num":p.get("of_num","")} for p in _S["of_periods"]]
        changes_s = [_dt_str(d) for d in _S["of_changes"]]
        pause_periods_s = [[_dt_str(a),_dt_str(b)] for a,b in _S["pause_periods"]]
        data = {
            "of_start":_dt_str(_S["of_start"]),
            "last_of_end":_dt_str(_S["last_of_end"]),
            "timers":timers_s,"tl_events":events_s,
            "of_periods":periods_s,"of_changes":changes_s,
            "form":_S["form"],"pilot":_S["pilot"],"poste":_S["poste"],
            "inter_of_s":_S["inter_of_s"],"interposte_s":_S["interposte_s"],
            "of_count_shift":_S["of_count_shift"],
            "pause_total_s":_S["pause_total_s"],
            "pause_start":_dt_str(_S["pause_start"]),
            "is_paused":_S["is_paused"],
            "pause_periods":pause_periods_s,
        }
        tmp = SESSION_FILE+".tmp"
        with open(tmp,"w",encoding="utf-8") as f: json.dump(data,f,ensure_ascii=False,indent=2)
        os.replace(tmp,SESSION_FILE)
    except Exception: pass

def load_session():
    if not os.path.exists(SESSION_FILE): return False
    try:
        with open(SESSION_FILE,"r",encoding="utf-8") as f: data=json.load(f)
        if not data.get("of_start"): return False
        _S["of_start"] = _str_dt(data["of_start"])
        _S["last_of_end"] = _str_dt(data.get("last_of_end"))
        _S["prod_active"] = True
        _S["timers"] = {}
        for k,t in data.get("timers",{}).items():
            _S["timers"][k]={"elapsed":float(t["elapsed"]),"running":bool(t["running"]),"start":_str_dt(t.get("start"))}
        _S["tl_events"] = []
        for ev in data.get("tl_events",[]):
            _S["tl_events"].append({
                "key":ev["key"],"cat":ev["cat"],
                "start":_str_dt(ev["start"]),"end":_str_dt(ev.get("end")),
                "comment":ev.get("comment",""),
                "hors_trs":ev.get("hors_trs",False),
                "nettoyage_type":ev.get("nettoyage_type",""),
                "of_num":ev.get("of_num",""),
                "pilot":ev.get("pilot",""),
            })
        _S["of_periods"] = [{"start":_str_dt(p["start"]),"end":_str_dt(p.get("end")),"of_num":p.get("of_num","")} for p in data.get("of_periods",[])]
        _S["of_changes"] = [_str_dt(d) for d in data.get("of_changes",[]) if d]
        _S["form"] = data.get("form",{})
        _S["pilot"] = data.get("pilot")
        _S["poste"] = data.get("poste")
        _S["inter_of_s"] = float(data.get("inter_of_s",0))
        _S["interposte_s"] = float(data.get("interposte_s",0))
        _S["of_count_shift"] = int(data.get("of_count_shift",0))
        _S["pause_total_s"] = float(data.get("pause_total_s",0))
        _S["is_paused"] = bool(data.get("is_paused",False))
        _S["pause_start"] = _str_dt(data.get("pause_start"))
        _S["pause_periods"] = []
        for pair in data.get("pause_periods",[]):
            if len(pair)==2: _S["pause_periods"].append((_str_dt(pair[0]),_str_dt(pair[1])))
        return True
    except Exception: return False

# ── Listes depuis Excel ────────────────────────────────────────────────────────
def load_lists():
    global _lists, _prod_ref_cached
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    try:
        wb = load_workbook(path,read_only=True,data_only=True)
        if "Listes" not in wb.sheetnames: return
        ws = wb["Listes"]
        headers = {}
        for cell in next(ws.iter_rows(max_row=1)):
            if cell.value: headers[cell.column]=str(cell.value)
        _lists = {h:[] for h in headers.values()}
        rows_listes = list(ws.iter_rows(min_row=2,values_only=True))
        for row in rows_listes:
            for ci,val in enumerate(row,1):
                if ci in headers and val is not None:
                    _lists[headers[ci]].append(str(val))
        try:
            for row_ix in ws.iter_rows(min_row=2,min_col=9,max_col=9,values_only=True):
                if row_ix and row_ix[0] is not None:
                    v_ref = str(row_ix[0]).replace(",",".")
                    if v_ref.replace(".","",1).isdigit():
                        _prod_ref_cached = float(v_ref)
                        break
        except: pass
        wb.close()
    except Exception: pass

def get_list(h):
    return _lists.get(h,[])

def load_history():
    global _data_rows_cache, _events_cache
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    try:
        wb = load_workbook(path,read_only=True,data_only=True)
        _data_rows_cache = []
        if "Data" in wb.sheetnames:
            ws = wb["Data"]
            min_r = 2 if str(ws.cell(1,1).value or "").strip().upper()=="OF" else 1
            for i,r in enumerate(ws.iter_rows(min_row=min_r,values_only=True),start=min_r):
                if any(r): _data_rows_cache.append((i,list(r)+[None]*60))
        _events_cache = []
        if "Evenements" in wb.sheetnames:
            ws_e = wb["Evenements"]
            for r in ws_e.iter_rows(min_row=2,values_only=True):
                if r and any(r): _events_cache.append(list(r))
        wb.close()
    except Exception: pass

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
    return qte

# ── Excel helpers ─────────────────────────────────────────────────────────────
def _get_wb(path):
    try: return load_workbook(path,keep_links=False)
    except: return None

def _safe_excel_save(wb, path):
    tmp_path = path+".tmp_kpi"
    bak_path = path+".bak_kpi"
    wb.save(tmp_path)
    try:
        if os.path.exists(path): shutil.copy2(path,bak_path)
    except: pass
    os.replace(tmp_path,path)

def _format_row(ws, row_num):
    try:
        thin = Side(border_style="thin",color="BBBBBB")
        border = Border(left=thin,right=thin,top=thin,bottom=thin)
        for cell in ws[row_num]:
            cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=False)
            cell.border = border
    except: pass

def _ensure_data_sheet(wb):
    if "Data" not in wb.sheetnames:
        ws = wb.create_sheet("Data",0)
        for i,h in enumerate(DATA_HEADERS,start=1): ws.cell(1,i).value=h
        _format_row(ws,1)
    else:
        ws = wb["Data"]
        for i,h in enumerate(DATA_HEADERS,start=1):
            if ws.cell(1,i).value is None: ws.cell(1,i).value=h
    return wb["Data"]

def _ensure_events_sheet(wb):
    if "Evenements" not in wb.sheetnames:
        ws = wb.create_sheet("Evenements")
        for i,h in enumerate(EVT_HEADERS,start=1): ws.cell(1,i).value=h
        _format_row(ws,1)
    else:
        ws = wb["Evenements"]
        existing = [ws.cell(1,i).value for i in range(1,len(EVT_HEADERS)+1)]
        if existing!=EVT_HEADERS:
            for i,h in enumerate(EVT_HEADERS,start=1): ws.cell(1,i).value=h
            _format_row(ws,1)
    return wb["Evenements"]

def build_events_rows(v, tl_events, of_start, pause_periods):
    rows = []
    kit_val = "Oui" if v.get("kit") else "Non"
    def _base_row(label,start,end):
        dur = (end-start).total_seconds()
        return [
            label,v.get("of_num",""),start.strftime("%d/%m/%Y"),
            v.get("poste",""),v.get("pilote",""),v.get("copilote",""),v.get("nb_pers",""),
            v.get("taille",""),v.get("type_prod",""),v.get("code_prod",""),
            v.get("fibre",""),v.get("poids",""),v.get("of_taie",""),v.get("traca",""),
            v.get("ref_taie",""),kit_val,
            start.strftime("%H:%M:%S"),end.strftime("%H:%M:%S"),fmt(dur),"","",
        ]
    for ev in tl_events:
        if ev.get("cat") not in ("ratt","pb","nettoyage"): continue
        if not ev.get("key") or ev["key"].startswith("_"): continue
        if of_start and ev["start"]<of_start and ev.get("key")!="arret_interposte": continue
        start = ev["start"]
        end = ev.get("end") or datetime.datetime.now()
        if ev["key"]=="nettoyage":
            ntype = ev.get("nettoyage_type","court")
            labels_nett={"court":"Nettoyage court","long":"Nettoyage long","grand":"Grand nettoyage"}
            label = labels_nett.get(ntype,"Nettoyage court")
            row = _base_row(label,start,end)
        else:
            cat_name = "Rattrapage" if ev["cat"]=="ratt" else "PB Technique"
            lbl = next((e[0] for e in EVENTS if e[1]==ev["key"]),ev["key"])
            row = _base_row(f"{cat_name}: {lbl}",start,end)
        row[19] = ev.get("comment","")
        row[20] = "OUI" if ev.get("hors_trs") else ""
        rows.append(row)
    for ps,pe in pause_periods:
        if of_start and ps<of_start: continue
        rows.append(_base_row("Pause pilote",ps,pe))
    return rows

def write_excel_bg(row, events_rows):
    path = cfg.get("db_path","")
    if not path: return
    try:
        with open(PENDING_FILE,"w",encoding="utf-8") as f:
            json.dump({"db_path":path,"row":row,"events_rows":events_rows},f,ensure_ascii=False,default=str)
    except: pass
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws_d = _ensure_data_sheet(wb)
                ws_d.append(row)
                _format_row(ws_d,ws_d.max_row)
                ws_e = _ensure_events_sheet(wb)
                for er in events_rows:
                    ws_e.append(er)
                    _format_row(ws_e,ws_e.max_row)
                _safe_excel_save(wb,path)
                try: os.remove(PENDING_FILE)
                except: pass
        except Exception:
            pass
        threading.Thread(target=load_history,daemon=True).start()
    threading.Thread(target=_bg,daemon=True).start()

def write_changement_of(start_dt, end_dt):
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    pilot = _S.get("last_of_pilot") or _S.get("pilot") or ""
    row_evt = [
        "Changement d'OF","",start_dt.strftime("%d/%m/%Y"),
        _S.get("poste",""),pilot,
        "","","","","","","","","","","",
        start_dt.strftime("%H:%M:%S"),end_dt.strftime("%H:%M:%S"),
        fmt((end_dt-start_dt).total_seconds()),"","",
    ]
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _ensure_events_sheet(wb)
                ws.append(row_evt)
                _format_row(ws,ws.max_row)
                _safe_excel_save(wb,path)
        except: pass
    threading.Thread(target=_bg,daemon=True).start()

def toggle_hors_trs_excel(ev_date,ev_of,ev_pilote,ev_hd,new_val):
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None or "Evenements" not in wb.sheetnames: return
                ws = wb["Evenements"]
                for row in ws.iter_rows(min_row=2):
                    r_date=str(row[2].value or "")[:10]
                    r_of=str(row[1].value or "")
                    r_pil=str(row[4].value or "")
                    r_hd=str(row[16].value or "")[:8]
                    if r_date==ev_date and r_of==ev_of and r_pil==ev_pilote and r_hd==ev_hd:
                        row[20].value=new_val; break
                _safe_excel_save(wb,path)
            threading.Thread(target=load_history,daemon=True).start()
        except: pass
    threading.Thread(target=_bg,daemon=True).start()

# ── Flask app ─────────────────────────────────────────────────────────────────
flask_app = Flask(__name__)
flask_app.config['JSON_AS_ASCII'] = False

def _check_pw(pw):
    return pw == str(cfg.get("supervisor_pw","1234"))

def _state_json():
    now = datetime.datetime.now()
    of_elapsed = 0.0
    if _S["of_start"]:
        of_elapsed = (now-_S["of_start"]).total_seconds() + _S["inter_of_s"]
        if _S["is_paused"] and _S["pause_start"]:
            # ne pas compter le temps de pause en cours dans l'elapsed affiché
            pass
    pause_elapsed = _S["pause_total_s"]
    if _S["is_paused"] and _S["pause_start"]:
        pause_elapsed += (now-_S["pause_start"]).total_seconds()
    prod_ref = get_prod_ref()
    stop_wall = t_wall_clock_stops()
    active_stops = [k for k in _S["timers"] if t_running(k)]
    timers_out = {}
    for k in _S["timers"]:
        timers_out[k] = {"elapsed":round(t_get(k),1),"running":t_running(k)}
    recent_events = [serialize_event(ev) for ev in reversed(_S["tl_events"][-30:])]
    return {
        "pilot": _S["pilot"],
        "poste": _S["poste"],
        "prod_active": _S["prod_active"],
        "is_paused": _S["is_paused"],
        "of_start_iso": _dt_str(_S["of_start"]),
        "of_elapsed_s": round(of_elapsed,1),
        "inter_of_s": _S["inter_of_s"],
        "pause_total_s": round(pause_elapsed,1),
        "pause_start_iso": _dt_str(_S["pause_start"]),
        "stop_wall_s": round(stop_wall,1),
        "timers": timers_out,
        "tl_events": recent_events,
        "active_stops": active_stops,
        "prod_ref": prod_ref,
        "db_path": cfg.get("db_path",""),
        "db_name": os.path.basename(cfg.get("db_path","")) if cfg.get("db_path") else "",
        "interposte_s": _S["interposte_s"],
        "of_count_shift": _S["of_count_shift"],
        "form": _S["form"],
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
        "pilotes": get_list("Pilotes"),
        "tailles": get_list("Taille"),
        "types_prod": get_list("Type produit"),
        "fibres": get_list("Fibre"),
        "tracas": get_list("Traca"),
        "postes": POSTES,
    })

@flask_app.route('/api/login', methods=['POST'])
def api_login():
    data = request.json or {}
    pilot = str(data.get("pilot","")).strip()
    poste = str(data.get("poste","")).strip()
    if not pilot or not poste:
        return jsonify({"ok":False,"error":"Pilote et poste obligatoires"}),400
    _S["pilot"] = pilot
    _S["poste"] = poste
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/logout', methods=['POST'])
def api_logout():
    if _S["prod_active"]:
        return jsonify({"ok":False,"error":"Production en cours — clôturez d'abord"}),400
    _S["pilot"] = None
    _S["poste"] = None
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/start_prod', methods=['POST'])
def api_start_prod():
    if not _S["pilot"]:
        return jsonify({"ok":False,"error":"Connectez-vous d'abord"}),400
    data = request.json or {}
    now = datetime.datetime.now()
    _S["inter_of_s"] = 0.0
    _S["interposte_s"] = 0.0
    # Calcul inter-OF si dernier OF connu
    if _S["last_of_end"] is not None:
        gap = (now-_S["last_of_end"]).total_seconds()
        if 30 < gap <= 28800:
            _S["inter_of_s"] = data.get("inter_of_s_override", gap)
        if (_S["last_of_pilot"] and _S["last_of_pilot"]!=_S["pilot"]
                and 0 < gap <= 3600):
            _S["interposte_s"] = gap
    t_reset()
    _S["of_start"] = now
    _S["prod_active"] = True
    _S["is_paused"] = False
    _S["pause_start"] = None
    _S["pause_total_s"] = 0.0
    _S["pause_periods"] = []
    _S["form"] = {}
    _S["tl_events"] = [e for e in _S["tl_events"] if e.get("start") and e["start"] < now]
    _S["of_periods"].append({"start":now,"end":None,"of_num":""})
    save_session()
    return jsonify({"ok":True,"inter_of_s":_S["inter_of_s"],"gap_s":0 if _S["last_of_end"] is None else (now-_S["last_of_end"]).total_seconds()})

@flask_app.route('/api/inter_of_confirm', methods=['POST'])
def api_inter_of_confirm():
    data = request.json or {}
    _S["inter_of_s"] = float(data.get("inter_of_s",_S["inter_of_s"]))
    write_changement_of(_S["last_of_end"], _S["of_start"])
    return jsonify({"ok":True})

@flask_app.route('/api/end_prod', methods=['POST'])
def api_end_prod():
    if not _S["prod_active"] or not _S["of_start"]:
        return jsonify({"ok":False,"error":"Pas de production active"}),400
    data = request.json or {}
    v = data.get("form",{})
    # Terminer pause si active
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
    kit = 2 if v.get("kit") else 1
    _nett_s = sum(
        (ev["end"]-ev["start"]).total_seconds()
        for ev in _S["tl_events"]
        if ev.get("key")=="nettoyage" and ev.get("start") and ev.get("end")
    )
    def _ts(key): return fmt(t_get(key))
    row = [
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
        qte_fab,
        _n(v.get("qte_emb",0)),
        equiv,
        fmt(of_s),
        _S["of_start"].strftime("%H:%M:%S"),
        end_dt.strftime("%H:%M:%S"),
        c1, c2, kit,
        v.get("ref_taie",""),
        _n(v.get("qte_init_taie",0)),
        _n(v.get("nb_taie2_choix",0)),
        _n(v.get("nb_def_cout",0)),
        _n(v.get("mq_taie",0)),
        _n(v.get("mq_housse_encart",0)),
        _n(v.get("nb_pp_cousue",0)),
        fmt(_S["inter_of_s"]),
        _min_str_to_hms(v.get("duree_mq_mp","")),
        _min_str_to_hms(v.get("manquant_pers","")),
        fmt(_nett_s),
        _ts("ratt_pochon"), _ts("ratt_couture"),
        _ts("ratt_emb"),    _ts("ratt_presse_soud"),
        _ts("ratt_presse_zip"),
        _ts("pb_chargeuse"), _ts("pb_carde"),
        _ts("pb_etaleur"),   _ts("pb_coupe"),
        _ts("pb_tapis1"),    _ts("pb_enrouleur"),
        _ts("pb_pesee"),     _ts("pb_deviation"),
        _ts("pb_enfileur"),  _ts("pb_kinna"),
        _ts("pb_tapeuse"),   _ts("pb_table_rot"),
        _ts("pb_h100"),      _ts("pb_traversin"),
        _ts("pb_presse_orc"),_ts("pb_presse_zip2"),
        _ts("pb_cercleuse"), _ts("pb_enrouleuse"),
        v.get("comment",""),
        fmt(_S["interposte_s"]) if _S["interposte_s"]>0 else "",
    ]
    events_rows = build_events_rows(
        dict(v, pilote=v.get("pilote",_S["pilot"] or ""), poste=v.get("poste",_S["poste"] or "")),
        _S["tl_events"], _S["of_start"], _S["pause_periods"]
    )
    prod_ref = get_prod_ref()
    trs = -1.0
    if prod_ref>0 and of_s>0:
        trs = equiv/(prod_ref*of_s/28800)*100
    # Préparer données recap avant reset
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
        "trs": round(trs,1),
        "debut": _S["of_start"].strftime("%H:%M:%S"),
        "fin": end_dt.strftime("%H:%M:%S"),
        "date": datetime.date.today().strftime("%d/%m/%Y"),
        "c1": c1, "c2": c2,
        "nett_s": round(_nett_s,0),
        "interposte_s": round(_S["interposte_s"],0),
    }
    # Reset état
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
    write_excel_bg(row, events_rows)
    return jsonify({"ok":True,"recap":recap})

@flask_app.route('/api/start_stop', methods=['POST'])
def api_start_stop():
    data = request.json or {}
    key = data.get("key","")
    cat = data.get("cat","pb")
    if not key: return jsonify({"ok":False}),400
    t_start(key)
    tl_open(key, cat)
    return jsonify({"ok":True})

@flask_app.route('/api/end_stop', methods=['POST'])
def api_end_stop():
    data = request.json or {}
    key = data.get("key","")
    comment = data.get("comment","")
    if not key: return jsonify({"ok":False}),400
    t_stop(key)
    tl_close(key, comment)
    return jsonify({"ok":True,"elapsed":round(t_get(key),0)})

def _toggle_pause_internal():
    now = datetime.datetime.now()
    if not _S["is_paused"]:
        _S["is_paused"] = True
        _S["pause_start"] = now
    else:
        if _S["pause_start"]:
            dur = (now-_S["pause_start"]).total_seconds()
            _S["pause_total_s"] += dur
            _S["pause_periods"].append((_S["pause_start"], now))
        _S["is_paused"] = False
        _S["pause_start"] = None
    save_session()

@flask_app.route('/api/toggle_pause', methods=['POST'])
def api_toggle_pause():
    _toggle_pause_internal()
    return jsonify({"ok":True,"is_paused":_S["is_paused"]})

@flask_app.route('/api/start_nettoyage', methods=['POST'])
def api_start_nettoyage():
    data = request.json or {}
    ntype = data.get("ntype","court")
    key = "nettoyage"
    if t_running(key):
        return jsonify({"ok":False,"error":"Nettoyage déjà en cours"}),400
    t_start(key)
    of_num = _S["form"].get("of_num","")
    _S["tl_events"].append({
        "key":key,"cat":"nettoyage",
        "start":datetime.datetime.now(),"end":None,
        "comment":"","of_num":of_num,"pilot":_S["pilot"] or "",
        "nettoyage_type":ntype,
    })
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/end_nettoyage', methods=['POST'])
def api_end_nettoyage():
    data = request.json or {}
    comment = data.get("comment","")
    t_stop("nettoyage")
    tl_close("nettoyage", comment)
    return jsonify({"ok":True})

@flask_app.route('/api/save_form', methods=['POST'])
def api_save_form():
    data = request.json or {}
    _S["form"].update(data)
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/history')
def api_history():
    rows = []
    for _,r in _data_rows_cache[-100:]:
        try:
            trs=-1
            try:
                equiv_v=float(str(r[15] or 0).replace(",","."))
                of_s_v=_hms_to_sec(str(r[16] or "00:00:00"))
                pr=get_prod_ref()
                if pr>0 and of_s_v>0 and equiv_v>0:
                    trs=round(equiv_v/(pr*of_s_v/28800)*100,1)
            except: pass
            rows.append({
                "date":_row_date(r[1]),"debut":_row_time(r[17])[:5] if r[17] else "",
                "fin":_row_time(r[18])[:5] if r[18] else "",
                "of":str(r[0] or ""),"pilote":str(r[3] or ""),
                "poste":str(r[2] or ""),"qte_fab":str(r[13] or ""),
                "qte_emb":str(r[14] or ""),"equiv":str(r[15] or ""),
                "trs":trs,"duree":str(r[16] or ""),
                "taille":str(r[6] or ""),"type_prod":str(r[8] or ""),
            })
        except: pass
    return jsonify(list(reversed(rows)))

@flask_app.route('/api/events_list')
def api_events_list():
    rows = []
    for r in _events_cache[-100:]:
        try:
            hors = str(r[20] if len(r)>20 else "").strip().upper()
            rows.append({
                "type":str(r[0] or ""),"of":str(r[1] or ""),
                "date":str(r[2] or "")[:10],"poste":str(r[3] or ""),
                "pilote":str(r[4] or ""),"debut":str(r[16] or "")[:8],
                "fin":str(r[17] or "")[:8],"duree":str(r[18] or ""),
                "comment":str(r[19] or ""),"hors_trs": hors=="OUI",
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
    toggle_hors_trs_excel(
        data.get("ev_date",""),data.get("ev_of",""),
        data.get("ev_pilote",""),data.get("ev_hd",""),
        data.get("new_val","")
    )
    return jsonify({"ok":True})

@flask_app.route('/api/fin_poste_data')
def api_fin_poste_data():
    today = datetime.date.today().strftime("%d/%m/%Y")
    pilot = _S["pilot"] or ""
    rows_today = []
    for _,r in _data_rows_cache:
        if _row_date(r[1])==today and str(r[3] or "")==pilot:
            rows_today.append(r)
    prod_ref = get_prod_ref()
    tot_eq=0.0; tot_s=0.0; tot_stop=0.0
    of_list=[]
    for r in rows_today:
        try:
            eq=float(str(r[15] or 0).replace(",","."))
            s=_hms_to_sec(str(r[16] or "00:00:00"))
            tot_eq+=eq; tot_s+=s
            trs=-1
            if prod_ref>0 and s>0 and eq>0: trs=round(eq/(prod_ref*s/28800)*100,1)
            of_list.append({
                "of":str(r[0] or ""),"taille":str(r[6] or ""),
                "type_prod":str(r[8] or ""),"qte_fab":str(r[13] or ""),
                "qte_emb":str(r[14] or ""),"equiv":str(r[15] or ""),
                "debut":_row_time(r[17])[:5],"fin":_row_time(r[18])[:5],
                "duree":str(r[16] or ""),"trs":trs,
            })
        except: pass
    trs_poste=-1.0
    if prod_ref>0 and tot_s>0: trs_poste=round(tot_eq/(prod_ref*tot_s/28800)*100,1)
    return jsonify({
        "pilot":pilot,"date":today,
        "nb_of":len(rows_today),"trs":trs_poste,
        "tot_equiv":round(tot_eq,1),"tot_s":round(tot_s,0),
        "of_list":of_list,
        "of_count_shift":_S["of_count_shift"],
    })

@flask_app.route('/api/reload', methods=['POST'])
def api_reload():
    threading.Thread(target=load_lists,daemon=True).start()
    threading.Thread(target=load_history,daemon=True).start()
    return jsonify({"ok":True})

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
body{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--dark);overflow:hidden;height:100vh}
.view{display:none;height:100vh;flex-direction:column;overflow:hidden}
.view.active{display:flex}
/* ── Header ── */
.hdr{background:var(--white);border-bottom:2px solid var(--lgray);display:flex;align-items:center;padding:0 16px;height:56px;flex-shrink:0;box-shadow:var(--shadow)}
.hdr-logo{font-size:17px;font-weight:800;color:var(--navy);letter-spacing:.5px;margin-right:12px}
.hdr-accent{width:5px;height:36px;border-radius:3px;background:var(--green);margin-right:12px}
.hdr-accent.red{background:var(--red)}
.hdr-right{margin-left:auto;display:flex;gap:8px;align-items:center}
.hdr-pilot{font-size:13px;color:var(--gray)}
.hdr-pilot strong{color:var(--dark)}
/* ── Buttons ── */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;padding:8px 16px;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;transition:all .15s;white-space:nowrap}
.btn:active{transform:scale(.97)}
.btn-navy{background:var(--navy);color:#fff}.btn-navy:hover{background:var(--navy-l)}
.btn-green{background:var(--green);color:#fff}.btn-green:hover{filter:brightness(1.1)}
.btn-red{background:var(--red);color:#fff}.btn-red:hover{filter:brightness(1.1)}
.btn-amber{background:var(--amber);color:#fff}.btn-amber:hover{filter:brightness(1.1)}
.btn-cyan{background:var(--cyan);color:#fff}.btn-cyan:hover{filter:brightness(1.1)}
.btn-purple{background:var(--purple);color:#fff}.btn-purple:hover{filter:brightness(1.1)}
.btn-ghost{background:var(--lgray);color:var(--dark)}.btn-ghost:hover{background:#c8d4e8}
.btn-lg{padding:14px 24px;font-size:15px;border-radius:10px}
.btn-xl{padding:20px 32px;font-size:18px;border-radius:12px;width:100%}
.btn-sm{padding:5px 12px;font-size:12px;border-radius:6px}
/* ── Cards ── */
.card{background:var(--white);border-radius:12px;box-shadow:var(--shadow);padding:16px}
.card-hdr{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--gray);margin-bottom:12px}
/* ── Status bar ── */
.status-bar{display:flex;gap:12px;padding:8px 16px;flex-shrink:0}
.status-cell{background:var(--white);border-radius:10px;padding:10px 20px;text-align:center;box-shadow:var(--shadow);flex:1}
.status-cell.running{background:var(--red);color:#fff}
.status-cell.ok{background:var(--green);color:#fff}
.status-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;opacity:.75}
.status-value{font-size:26px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.1}
.status-sub{font-size:11px;opacity:.7;margin-top:2px}
/* ── Production layout ── */
.prod-body{display:grid;grid-template-columns:2fr 1.2fr 1.3fr;gap:10px;padding:10px 12px;flex:1;min-height:0;overflow:hidden}
.prod-col{overflow-y:auto;min-height:0}
/* ── Form ── */
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.form-group{display:flex;flex-direction:column;gap:3px}
.form-group.full{grid-column:1/-1}
.form-group label{font-size:11px;font-weight:700;color:var(--gray);text-transform:uppercase;letter-spacing:.5px}
.form-group input,.form-group select,.form-group textarea{
  border:1.5px solid var(--lgray);border-radius:7px;padding:7px 10px;
  font-size:13px;color:var(--dark);background:var(--white);outline:none;
  transition:border-color .15s}
.form-group input:focus,.form-group select:focus,.form-group textarea:focus{border-color:var(--navy-l)}
.form-group input[disabled],.form-group select[disabled]{background:#f0f4fb;color:var(--gray)}
.form-group textarea{resize:vertical;min-height:56px}
.checkbox-row{display:flex;align-items:center;gap:8px}
.checkbox-row input[type=checkbox]{width:18px;height:18px;cursor:pointer}
/* ── Stop selector ── */
.stops-grid{display:grid;gap:8px}
.stops-grid.ratt{grid-template-columns:repeat(3,1fr)}
.stops-grid.pb{grid-template-columns:repeat(4,1fr)}
.stops-grid.special{grid-template-columns:repeat(2,1fr)}
.stop-btn{
  border:none;border-radius:10px;padding:12px 8px;cursor:pointer;
  font-size:12px;font-weight:700;color:#fff;text-align:center;
  transition:all .15s;position:relative;box-shadow:0 3px 0 rgba(0,0,0,.25)}
.stop-btn:active{transform:translateY(2px);box-shadow:0 1px 0 rgba(0,0,0,.25)}
.stop-btn.ratt{background:var(--amber)}
.stop-btn.pb{background:var(--navy)}
.stop-btn.special{background:var(--purple)}
.stop-btn.running{animation:pulse 1.2s infinite}
.stop-btn.running::after{content:'● EN COURS';display:block;font-size:9px;margin-top:4px;opacity:.85}
.stop-btn .timer{font-size:11px;font-weight:600;margin-top:3px}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.7}}
/* ── Recap stops ── */
.recap-stop-row{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:6px;margin-bottom:4px;background:#f8fafc}
.recap-stop-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.recap-stop-name{font-size:12px;flex:1;font-weight:600}
.recap-stop-time{font-size:12px;color:var(--gray);font-variant-numeric:tabular-nums}
.recap-stop-row.running{background:#fff0f0;animation:pulse 1.5s infinite}
/* ── TRS gauge ── */
.gauge-wrap{position:relative;width:160px;height:90px;margin:0 auto}
.gauge-svg{width:160px;height:90px}
.gauge-text{position:absolute;bottom:0;left:50%;transform:translateX(-50%);text-align:center}
.gauge-pct{font-size:28px;font-weight:800}
.gauge-lbl{font-size:10px;color:var(--gray);font-weight:700;text-transform:uppercase}
/* ── Tables ── */
.table-wrap{overflow-x:auto;border-radius:10px;box-shadow:var(--shadow)}
.ktable{width:100%;border-collapse:collapse;background:var(--white);font-size:12px}
.ktable th{background:var(--lgray);color:var(--dark);font-weight:700;padding:9px 10px;text-align:center;white-space:nowrap;border-bottom:2px solid #c0cde0}
.ktable td{padding:8px 10px;text-align:center;border-bottom:1px solid #edf0f7;white-space:nowrap}
.ktable tr:hover td{background:#f0f4fb}
.trs-hi td:first-child,.trs-hi td:nth-child(9){color:var(--green);font-weight:800}
.trs-warn td:first-child,.trs-warn td:nth-child(9){color:var(--amber);font-weight:800}
.trs-low td:first-child,.trs-low td:nth-child(9){color:var(--red);font-weight:800}
/* ── Modal / Overlay ── */
.modal{display:none;position:fixed;inset:0;z-index:100;background:rgba(15,23,42,.5);align-items:center;justify-content:center}
.modal.active{display:flex}
.modal-box{background:var(--white);border-radius:16px;box-shadow:var(--shadow-lg);width:min(700px,95vw);max-height:90vh;overflow-y:auto;display:flex;flex-direction:column}
.modal-box.wide{width:min(960px,98vw)}
.modal-box.narrow{width:min(460px,95vw)}
.modal-hdr{background:var(--navy);color:#fff;padding:18px 24px;border-radius:16px 16px 0 0;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.modal-hdr h2{font-size:16px;font-weight:800}
.modal-hdr.green{background:var(--green)}
.modal-hdr.red{background:var(--red)}
.modal-hdr.amber{background:var(--amber)}
.modal-body{padding:20px 24px;flex:1}
.modal-footer{padding:16px 24px;border-top:1px solid var(--lgray);display:flex;gap:10px;justify-content:flex-end;flex-shrink:0}
/* ── Full-screen overlays ── */
.fullscreen{display:none;position:fixed;inset:0;z-index:200;flex-direction:column;align-items:center;justify-content:center;color:#fff}
.fullscreen.active{display:flex}
.fullscreen.pause-ov{background:#1a1f5e}
.fullscreen.reunion-ov{background:#5b21b6}
.fullscreen .fs-icon{font-size:72px;margin-bottom:16px}
.fullscreen .fs-title{font-size:36px;font-weight:800;margin-bottom:8px}
.fullscreen .fs-timer{font-size:56px;font-weight:900;font-variant-numeric:tabular-nums;letter-spacing:2px;margin:16px 0}
/* ── Tabs ── */
.tabs{display:flex;gap:2px;padding:0 16px;background:var(--white);border-bottom:2px solid var(--lgray);flex-shrink:0}
.tab{padding:10px 20px;font-size:13px;font-weight:700;cursor:pointer;border-bottom:3px solid transparent;color:var(--gray);transition:all .15s}
.tab.active{color:var(--navy);border-bottom-color:var(--navy)}
/* ── Category headers ── */
.cat-hdr{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;margin:12px 0 6px}
.cat-bar{width:4px;height:18px;border-radius:2px}
/* ── Alerts / badges ── */
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}
.badge-green{background:#dcfce7;color:#15803d}
.badge-amber{background:#fef3c7;color:#92400e}
.badge-red{background:#fee2e2;color:#991b1b}
.badge-navy{background:#dbeafe;color:#1e40af}
/* ── Interposte alert ── */
.interposte-alert{background:#ef4444;color:#fff;border-radius:8px;padding:8px 12px;font-weight:700;font-size:12px;animation:pulse 1.2s infinite}
/* ── Toast ── */
#toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#1a8c4e;color:#fff;padding:12px 28px;border-radius:12px;font-weight:700;font-size:14px;opacity:0;transition:opacity .3s;z-index:9999;pointer-events:none}
#toast.show{opacity:1}
/* ── Fin de poste stats ── */
.fin-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.fin-stat{background:var(--white);border-radius:10px;padding:16px;text-align:center;box-shadow:var(--shadow)}
.fin-stat-value{font-size:28px;font-weight:800;color:var(--navy)}
.fin-stat-label{font-size:11px;color:var(--gray);font-weight:700;text-transform:uppercase;margin-top:4px}
/* ── Scrollbar ── */
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:#f0f4fb}
::-webkit-scrollbar-thumb{background:#c0cde0;border-radius:3px}
/* ── Active stop banner ── */
.active-stops-bar{background:#1e293b;padding:6px 16px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;flex-shrink:0}
.active-stop-chip{background:var(--red);color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;display:flex;align-items:center;gap:6px;cursor:pointer;animation:pulse 1.2s infinite}
.active-stop-chip.nett{background:var(--cyan);animation:none}
/* ── Responsive ── */
@media(max-width:900px){.prod-body{grid-template-columns:1fr}}
</style>
</head>
<body>

<!-- ═══════════════════════ VUE PRINCIPALE ═══════════════════════ -->
<div id="view-main" class="view active">
  <div class="hdr">
    <div class="hdr-accent" id="hdr-accent-main"></div>
    <div class="hdr-logo">KPI-ORC &nbsp;|&nbsp; ORC1</div>
    <div id="main-prod-badge" style="display:none;margin-left:8px"></div>
    <div class="hdr-right">
      <div class="hdr-pilot">Pilote : <strong id="main-pilot-lbl">—</strong> &nbsp;|&nbsp; <span id="main-poste-lbl">—</span></div>
      <button class="btn btn-ghost btn-sm" onclick="openModal('modal-settings')">⚙ Paramètres</button>
      <button class="btn btn-navy btn-sm" onclick="doReload()">↻ Actualiser</button>
      <button class="btn btn-ghost btn-sm" onclick="openModal('modal-login')">⇄ Changer pilote</button>
    </div>
  </div>

  <!-- KPI row -->
  <div style="display:grid;grid-template-columns:200px 1fr;gap:12px;padding:12px 16px 0;flex-shrink:0">
    <!-- Gauge TRS -->
    <div class="card" style="text-align:center;padding:12px">
      <div class="card-hdr">TRS Poste</div>
      <div class="gauge-wrap">
        <svg class="gauge-svg" viewBox="0 0 160 90">
          <path d="M20,80 A60,60,0,0,1,140,80" fill="none" stroke="#dde4ef" stroke-width="14" stroke-linecap="round"/>
          <path id="gauge-arc" d="M20,80 A60,60,0,0,1,140,80" fill="none" stroke="#1a8c4e" stroke-width="14" stroke-linecap="round" stroke-dasharray="188.5" stroke-dashoffset="188.5"/>
        </svg>
        <div class="gauge-text">
          <div class="gauge-pct" id="main-trs-pct">—</div>
          <div class="gauge-lbl">TRS %</div>
        </div>
      </div>
    </div>
    <!-- Boutons action -->
    <div class="card" style="display:flex;flex-direction:column;justify-content:center;gap:10px;padding:14px">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
        <button id="btn-start-prod" class="btn btn-green btn-lg" onclick="startProduction()" style="font-size:16px">
          ▶ DÉMARRER UNE PROD
        </button>
        <button class="btn btn-navy btn-lg" onclick="showFinDePoste()" style="font-size:16px">
          🏁 FIN DE MON POSTE
        </button>
      </div>
    </div>
  </div>

  <!-- Onglets déclarations / événements -->
  <div class="tabs" style="margin-top:10px">
    <div class="tab active" onclick="switchTab('tab-decl',this)">📋 Déclarations</div>
    <div class="tab" onclick="switchTab('tab-evts',this)">📊 Événements</div>
  </div>

  <div style="flex:1;overflow:hidden;padding:10px 16px 16px">
    <div id="tab-decl" style="height:100%;overflow-y:auto">
      <div class="table-wrap">
        <table class="ktable" id="decl-table">
          <thead><tr>
            <th>Date</th><th>H.Début</th><th>H.Fin</th><th>OF</th>
            <th>Pilote</th><th>Poste</th><th>Taille</th><th>Type</th>
            <th>Qte Fab</th><th>Equiv</th><th>TRS %</th><th>Durée</th>
          </tr></thead>
          <tbody id="decl-tbody"><tr><td colspan="12" style="color:#94a3b8;padding:20px">Chargement…</td></tr></tbody>
        </table>
      </div>
    </div>
    <div id="tab-evts" style="display:none;height:100%;overflow-y:auto">
      <div class="table-wrap">
        <table class="ktable" id="evts-table">
          <thead><tr>
            <th>Type</th><th>Pilote</th><th>OF</th><th>Date</th>
            <th>Début</th><th>Fin</th><th>Durée</th><th>Commentaire</th><th>Hors TRS</th>
          </tr></thead>
          <tbody id="evts-tbody"><tr><td colspan="9" style="color:#94a3b8;padding:20px">Chargement…</td></tr></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ═══════════════════════ VUE PRODUCTION ═══════════════════════ -->
<div id="view-production" class="view">
  <div class="hdr" id="prod-hdr">
    <div class="hdr-accent red" id="hdr-accent-prod"></div>
    <div class="hdr-logo">KPI-ORC &nbsp;|&nbsp; Production en cours</div>
    <span id="prod-debut-lbl" style="font-size:12px;color:var(--gray);margin-left:8px"></span>
    <div class="hdr-right">
      <div class="hdr-pilot">Pilote : <strong id="prod-pilot-lbl">—</strong></div>
      <button class="btn btn-ghost btn-sm" onclick="doReload()">↻</button>
      <button class="btn btn-ghost btn-sm" onclick="openModal('modal-settings')">⚙</button>
    </div>
  </div>

  <!-- Status bar -->
  <div class="status-bar" id="prod-status-bar">
    <div class="status-cell" id="sc-of">
      <div class="status-label">Durée OF</div>
      <div class="status-value" id="sc-of-val">00:00:00</div>
    </div>
    <div class="status-cell" id="sc-stop">
      <div class="status-label">Arrêts</div>
      <div class="status-value" id="sc-stop-val">00:00:00</div>
      <div class="status-sub" id="sc-stop-count"></div>
    </div>
    <div class="status-cell" id="sc-pcs">
      <div class="status-label">Pièces attendues</div>
      <div class="status-value" id="sc-pcs-val">—</div>
    </div>
    <div class="status-cell" id="sc-trs">
      <div class="status-label">TRS live</div>
      <div class="status-value" id="sc-trs-val">—%</div>
    </div>
  </div>

  <!-- Arrêts actifs banner -->
  <div class="active-stops-bar" id="active-stops-bar" style="display:none"></div>

  <!-- Corps 3 colonnes -->
  <div class="prod-body">
    <!-- Col 1: Formulaire -->
    <div class="prod-col card" style="overflow-y:auto">
      <div class="card-hdr">Informations OF</div>
      <div class="form-grid" id="prod-form">
        <div class="form-group full">
          <label>N° OF *</label>
          <input type="text" id="f-of_num" placeholder="Ex: OF-12345" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Taille</label>
          <select id="f-taille" onchange="saveForm()"><option value="">—</option></select>
        </div>
        <div class="form-group">
          <label>Type produit</label>
          <select id="f-type_prod" onchange="saveForm()"><option value="">—</option></select>
        </div>
        <div class="form-group">
          <label>Code produit</label>
          <input type="text" id="f-code_prod" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Nb personnes</label>
          <input type="number" id="f-nb_pers" min="1" max="20" value="1" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Pilote</label>
          <input type="text" id="f-pilote" disabled>
        </div>
        <div class="form-group">
          <label>Poste</label>
          <input type="text" id="f-poste" disabled>
        </div>
        <div class="form-group">
          <label>Co-pilote</label>
          <input type="text" id="f-copilote" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Fibre</label>
          <select id="f-fibre" onchange="saveForm()"><option value="">—</option></select>
        </div>
        <div class="form-group">
          <label>Poids garnissage (g)</label>
          <input type="number" id="f-poids" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>OF Taie</label>
          <input type="text" id="f-of_taie" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Traca fibre</label>
          <select id="f-traca" onchange="saveForm()"><option value="">—</option></select>
        </div>
        <div class="form-group">
          <label>Ref taie</label>
          <input type="text" id="f-ref_taie" onchange="saveForm()">
        </div>
        <div class="form-group" style="align-items:flex-start;justify-content:flex-end">
          <label>Kit</label>
          <div class="checkbox-row" style="margin-top:8px">
            <input type="checkbox" id="f-kit" onchange="saveForm()">
            <span style="font-size:13px">Avec kit</span>
          </div>
        </div>
        <div class="form-group">
          <label>Qte fabriquée *</label>
          <input type="number" id="f-qte_fab" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Qte emballée</label>
          <input type="number" id="f-qte_emb" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Qte initiale taie</label>
          <input type="number" id="f-qte_init_taie" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Nb taie 2nd choix</label>
          <input type="number" id="f-nb_taie2_choix" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Nb défaut couture</label>
          <input type="number" id="f-nb_def_cout" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Mq taie</label>
          <input type="number" id="f-mq_taie" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Mq housse/encart</label>
          <input type="number" id="f-mq_housse_encart" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Nb PP cousue</label>
          <input type="number" id="f-nb_pp_cousue" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Manq. personnel/réunion (min)</label>
          <input type="number" id="f-manquant_pers" min="0" onchange="saveForm()">
        </div>
        <div class="form-group">
          <label>Manq. MP (min)</label>
          <input type="number" id="f-duree_mq_mp" min="0" onchange="saveForm()">
        </div>
        <div class="form-group full">
          <label>Commentaire</label>
          <textarea id="f-comment" rows="3" onchange="saveForm()"></textarea>
        </div>
      </div>
    </div>

    <!-- Col 2: Boutons action -->
    <div class="prod-col" style="display:flex;flex-direction:column;gap:10px">
      <!-- GROS bouton arrêt -->
      <button class="btn btn-red" onclick="openModal('modal-stop')"
        style="font-size:20px;font-weight:900;padding:24px;border-radius:14px;width:100%;box-shadow:0 6px 0 #9b0f14">
        ⛔ DÉCLARER UN ARRÊT
      </button>
      <button class="btn btn-cyan btn-lg" onclick="openNettoyage()" style="width:100%;font-size:15px">
        🧹 NETTOYAGE
      </button>
      <button class="btn" id="btn-pause" onclick="togglePause()"
        style="width:100%;font-size:15px;padding:14px;border-radius:10px;background:var(--navy);color:#fff;font-weight:700">
        ⏸ PAUSE
      </button>
      <button class="btn btn-purple btn-lg" onclick="openModal('modal-reunion-confirm')" style="width:100%;font-size:15px">
        🤝 RÉUNION
      </button>
      <div style="flex:1"></div>
      <button class="btn btn-amber btn-lg" onclick="openEndProd()" style="width:100%;font-size:15px;margin-top:auto">
        ✅ FIN DE PRODUCTION
      </button>
    </div>

    <!-- Col 3: Récap arrêts OF -->
    <div class="prod-col card">
      <div class="card-hdr">Récap arrêts OF en cours</div>
      <div id="recap-stops-list"></div>
      <div style="margin-top:12px;padding-top:12px;border-top:1px solid var(--lgray)">
        <div style="font-size:11px;color:var(--gray);font-weight:700;text-transform:uppercase">Pauses</div>
        <div id="recap-pauses" style="margin-top:6px;font-size:12px;color:var(--gray)">—</div>
      </div>
    </div>
  </div>
</div>

<!-- ══════════════ FULLSCREEN: PAUSE ══════════════ -->
<div class="fullscreen pause-ov" id="fs-pause">
  <div class="fs-icon">⏸</div>
  <div class="fs-title">PAUSE EN COURS</div>
  <div class="fs-timer" id="fs-pause-timer">00:00:00</div>
  <button class="btn btn-green btn-xl" style="max-width:300px" onclick="togglePause()">
    ▶ REPRENDRE LA PRODUCTION
  </button>
</div>

<!-- ══════════════ FULLSCREEN: RÉUNION ══════════════ -->
<div class="fullscreen reunion-ov" id="fs-reunion">
  <div class="fs-icon">🤝</div>
  <div class="fs-title">RÉUNION EN COURS</div>
  <div class="fs-timer" id="fs-reunion-timer">00:00:00</div>
  <p style="opacity:.7;margin-bottom:20px">Arrêt inclus dans Manquant Personnel</p>
  <button class="btn btn-xl" style="max-width:300px;background:#fff;color:#5b21b6;font-weight:900" onclick="stopReunion()">
    ✅ FIN DE RÉUNION
  </button>
</div>

<!-- ══════════════ MODAL: LOGIN ══════════════ -->
<div class="modal active" id="modal-login">
  <div class="modal-box narrow">
    <div class="modal-hdr green">
      <h2>👤 Connexion Pilote</h2>
    </div>
    <div class="modal-body">
      <div class="form-group" style="margin-bottom:14px">
        <label>Pilote *</label>
        <select id="login-pilot" style="font-size:15px;padding:10px"><option value="">Sélectionner…</option></select>
      </div>
      <div class="form-group" style="margin-bottom:14px">
        <label>Poste *</label>
        <select id="login-poste" style="font-size:15px;padding:10px">
          <option value="">Sélectionner…</option>
          <option>Matin</option><option>Midi</option><option>Nuit</option><option>Jour</option>
        </select>
      </div>
      <div id="login-err" style="color:var(--red);font-size:13px;margin-bottom:8px;display:none"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-green btn-lg" onclick="doLogin()" style="width:100%;font-size:16px">✔ SE CONNECTER</button>
    </div>
    <div style="text-align:center;padding:10px 24px 16px">
      <small id="login-db-name" style="color:var(--gray);cursor:pointer" onclick="openSetDb()">📂 Aucun fichier chargé — cliquer pour changer</small>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: STOP SELECTOR ══════════════ -->
<div class="modal" id="modal-stop">
  <div class="modal-box wide">
    <div class="modal-hdr red">
      <h2>⛔ Choisir un arrêt</h2>
      <button class="btn btn-sm" style="background:rgba(255,255,255,.2);color:#fff" onclick="closeModal('modal-stop')">✕ Retour</button>
    </div>
    <div class="modal-body">
      <!-- Rattrapages -->
      <div class="cat-hdr"><div class="cat-bar" style="background:var(--amber)"></div>ARRÊTS RATTRAPAGE</div>
      <div class="stops-grid ratt" id="stop-grid-ratt"></div>
      <!-- PB Techniques -->
      <div class="cat-hdr" style="margin-top:14px"><div class="cat-bar" style="background:var(--red)"></div>PROBLÈMES TECHNIQUES</div>
      <div class="stops-grid pb" id="stop-grid-pb"></div>
      <!-- Spéciaux -->
      <div class="cat-hdr" style="margin-top:14px"><div class="cat-bar" style="background:var(--purple)"></div>ARRÊTS SPÉCIAUX</div>
      <div class="stops-grid special" id="stop-grid-special"></div>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: STOP DESCRIPTION ══════════════ -->
<div class="modal" id="modal-stop-desc">
  <div class="modal-box narrow">
    <div class="modal-hdr green">
      <h2>✅ Clôturer l'arrêt</h2>
    </div>
    <div class="modal-body">
      <div id="stop-desc-name" style="font-weight:700;font-size:16px;color:var(--navy);margin-bottom:8px"></div>
      <div id="stop-desc-timer" style="font-size:28px;font-weight:900;color:var(--red);font-variant-numeric:tabular-nums;margin-bottom:16px"></div>
      <div class="form-group">
        <label>Commentaire (optionnel)</label>
        <textarea id="stop-desc-comment" rows="3" placeholder="Description de l'arrêt…"></textarea>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-stop-desc')">Annuler</button>
      <button class="btn btn-green btn-lg" onclick="confirmEndStop()">✔ Valider l'arrêt</button>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: NETTOYAGE ══════════════ -->
<div class="modal" id="modal-nettoyage">
  <div class="modal-box narrow">
    <div class="modal-hdr amber">
      <h2>🧹 Type de nettoyage</h2>
      <button class="btn btn-sm" style="background:rgba(255,255,255,.2);color:#fff" onclick="closeModal('modal-nettoyage')">✕</button>
    </div>
    <div class="modal-body" id="nett-body"></div>
  </div>
</div>

<!-- ══════════════ MODAL: RÉUNION CONFIRM ══════════════ -->
<div class="modal" id="modal-reunion-confirm">
  <div class="modal-box narrow">
    <div class="modal-hdr" style="background:var(--purple)">
      <h2>🤝 Déclarer une réunion ?</h2>
    </div>
    <div class="modal-body">
      <p style="font-size:14px;color:var(--gray)">La durée de la réunion sera comptée dans "Manquant Personnel/Réunion". Après la réunion, vous saisirez la durée dans le formulaire.</p>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-reunion-confirm')">Annuler</button>
      <button class="btn btn-purple btn-lg" onclick="startReunion()">✔ Commencer la réunion</button>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: INTER-OF ══════════════ -->
<div class="modal" id="modal-interof">
  <div class="modal-box narrow">
    <div class="modal-hdr amber">
      <h2>⏱ Changement d'OF — Temps inter-OF</h2>
    </div>
    <div class="modal-body">
      <p style="margin-bottom:14px;color:var(--gray);font-size:14px">Ce temps sera déclaré comme <strong>Changement de série</strong>. Vous pouvez le modifier si nécessaire.</p>
      <div style="display:flex;gap:8px;align-items:center;justify-content:center">
        <input type="number" id="iof-h" min="0" style="width:70px;font-size:20px;text-align:center;padding:8px;border:1.5px solid var(--lgray);border-radius:8px"> <span>h</span>
        <input type="number" id="iof-m" min="0" max="59" style="width:70px;font-size:20px;text-align:center;padding:8px;border:1.5px solid var(--lgray);border-radius:8px"> <span>min</span>
        <input type="number" id="iof-s" min="0" max="59" style="width:70px;font-size:20px;text-align:center;padding:8px;border:1.5px solid var(--lgray);border-radius:8px"> <span>sec</span>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-green btn-lg" onclick="confirmInterOf()">✔ Valider — Changement de série</button>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: RECAP FIN PROD ══════════════ -->
<div class="modal" id="modal-recap">
  <div class="modal-box wide">
    <div class="modal-hdr navy" style="background:var(--navy)">
      <h2>📋 Récapitulatif de l'OF — Confirmer ?</h2>
    </div>
    <div class="modal-body">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
        <div>
          <table style="width:100%;font-size:13px;border-collapse:collapse" id="recap-table"></table>
        </div>
        <div>
          <div class="card-hdr">Arrêts de l'OF</div>
          <div id="recap-stops-detail" style="font-size:12px"></div>
          <div style="margin-top:12px;padding:12px;border-radius:10px;text-align:center" id="recap-trs-box">
            <div style="font-size:12px;color:var(--gray);font-weight:700">TRS calculé</div>
            <div style="font-size:40px;font-weight:900" id="recap-trs-val">—</div>
          </div>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost btn-lg" onclick="closeModal('modal-recap')">↩ Modifier</button>
      <button class="btn btn-green btn-lg" onclick="confirmEndProd()">✔ Valider et enregistrer</button>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: FIN DE POSTE ══════════════ -->
<div class="modal" id="modal-fin-poste">
  <div class="modal-box wide">
    <div class="modal-hdr" style="background:var(--navy)">
      <h2>🏁 Fin de poste</h2>
      <button class="btn btn-sm" style="background:rgba(255,255,255,.2);color:#fff" onclick="closeModal('modal-fin-poste')">✕ Fermer</button>
    </div>
    <div class="modal-body">
      <div class="fin-stats" id="fin-stats-row"></div>
      <div class="card-hdr">Détail des OF du poste</div>
      <div class="table-wrap" style="margin-top:8px">
        <table class="ktable"><thead><tr>
          <th>OF</th><th>Taille</th><th>Type</th><th>Qte Fab</th><th>Equiv</th><th>Début</th><th>Fin</th><th>Durée</th><th>TRS%</th>
        </tr></thead>
        <tbody id="fin-poste-tbody"></tbody></table>
      </div>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: PARAMÈTRES ══════════════ -->
<div class="modal" id="modal-settings">
  <div class="modal-box narrow">
    <div class="modal-hdr">
      <h2>⚙ Paramètres</h2>
      <button class="btn btn-sm" style="background:rgba(255,255,255,.2);color:#fff" onclick="closeModal('modal-settings')">✕</button>
    </div>
    <div class="modal-body">
      <div class="form-group" style="margin-bottom:12px">
        <label>Mot de passe admin *</label>
        <input type="password" id="set-pw" placeholder="Requis pour sauvegarder">
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
        <div class="form-group">
          <label>Réf. prod 8h (pièces)</label>
          <input type="number" id="set-prod_ref" min="0">
        </div>
        <div class="form-group">
          <label>Pause max (min)</label>
          <input type="number" id="set-pause_max_min" min="0">
        </div>
        <div class="form-group">
          <label>Nett. court (min)</label>
          <input type="number" id="set-clean_short_min" min="0">
        </div>
        <div class="form-group">
          <label>Nett. long (min)</label>
          <input type="number" id="set-clean_long_min" min="0">
        </div>
        <div class="form-group">
          <label>Grand nett. (min)</label>
          <input type="number" id="set-clean_grand_min" min="0">
        </div>
        <div class="form-group">
          <label>Tolérance réunion (min)</label>
          <input type="number" id="set-meeting_tol_min" min="0">
        </div>
        <div class="form-group full">
          <label>Nouveau mot de passe admin</label>
          <input type="password" id="set-new_pw" placeholder="Laisser vide pour ne pas changer">
        </div>
      </div>
      <div id="settings-err" style="color:var(--red);font-size:13px;margin-top:8px;display:none"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-settings')">Annuler</button>
      <button class="btn btn-navy btn-lg" onclick="saveSettings()">💾 Sauvegarder</button>
    </div>
  </div>
</div>

<!-- ══════════════ MODAL: SET DB ══════════════ -->
<div class="modal" id="modal-set-db">
  <div class="modal-box narrow">
    <div class="modal-hdr">
      <h2>📂 Fichier Excel base de données</h2>
    </div>
    <div class="modal-body">
      <div class="form-group" style="margin-bottom:12px">
        <label>Mot de passe admin</label>
        <input type="password" id="setdb-pw">
      </div>
      <div class="form-group">
        <label>Chemin du fichier (.xlsx)</label>
        <input type="text" id="setdb-path" placeholder="C:\chemin\vers\fichier.xlsx">
      </div>
      <div id="setdb-err" style="color:var(--red);font-size:13px;margin-top:8px;display:none"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-set-db')">Annuler</button>
      <button class="btn btn-navy btn-lg" onclick="doSetDb()">Confirmer</button>
    </div>
  </div>
</div>

<!-- ══════════════ TOAST ══════════════ -->
<div id="toast"></div>

<!-- ══════════════════════ JAVASCRIPT ══════════════════════ -->
<script>
// ── État global ──
let S = {};
let localOfElapsed = 0;    // compteur OF local (s), incrémenté chaque seconde
let pauseLocalElapsed = 0;
let reunionRunning = false;
let reunionStartTs = null;
let stopDescKey = null;     // arrêt en cours de clôture
let pendingEndProdData = null; // data recap en attente de confirmation
let endProdFormData = null;

// ── Init ──
fetchState();
loadLists();
setInterval(fetchState, 5000);
setInterval(tickTimers, 1000);

async function fetchState() {
  try {
    const r = await fetch('/api/state');
    S = await r.json();
    // Sync local elapsed avec le serveur
    if (S.of_start_iso) {
      const serverSideElapsed = S.of_elapsed_s;
      localOfElapsed = serverSideElapsed;
    } else {
      localOfElapsed = 0;
    }
    if (S.pause_start_iso) {
      const pStart = new Date(S.pause_start_iso);
      pauseLocalElapsed = Math.round(S.pause_total_s - ((Date.now() - pStart)/1000));
      if(pauseLocalElapsed<0) pauseLocalElapsed=0;
    } else {
      pauseLocalElapsed = S.pause_total_s || 0;
    }
    renderAll();
  } catch(e) { /* offline */ }
}

async function loadLists() {
  try {
    const r = await fetch('/api/lists');
    const lists = await r.json();
    fillSelect('login-pilot', lists.pilotes);
    fillSelect('f-taille', lists.tailles);
    fillSelect('f-type_prod', lists.types_prod);
    fillSelect('f-fibre', lists.fibres);
    fillSelect('f-traca', lists.tracas);
  } catch(e) {}
  try {
    const r = await fetch('/api/config');
    const conf = await r.json();
    document.getElementById('set-prod_ref').value = conf.prod_ref||'';
    document.getElementById('set-pause_max_min').value = conf.pause_max_min||20;
    document.getElementById('set-clean_short_min').value = conf.clean_short_min||10;
    document.getElementById('set-clean_long_min').value = conf.clean_long_min||30;
    document.getElementById('set-clean_grand_min').value = conf.clean_grand_min||60;
    document.getElementById('set-meeting_tol_min').value = conf.meeting_tol_min||5;
    const dbLbl = document.getElementById('login-db-name');
    if(dbLbl) dbLbl.textContent = conf.db_name ? '📂 '+conf.db_name : '📂 Aucun fichier chargé — cliquer pour changer';
  } catch(e) {}
}

function fillSelect(id, items) {
  const sel = document.getElementById(id);
  if(!sel) return;
  const prev = sel.value;
  while(sel.options.length>1) sel.remove(1);
  (items||[]).forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent=v; sel.appendChild(o); });
  if(prev) sel.value=prev;
}

function tickTimers() {
  if (S.prod_active && !S.is_paused && S.of_start_iso) {
    localOfElapsed++;
    updateStatusBar();
  }
  if (S.is_paused && S.pause_start_iso) {
    const pStart = new Date(S.pause_start_iso);
    const pauseCur = S.pause_total_s + (Date.now()/1000 - pStart.getTime()/1000);
    document.getElementById('fs-pause-timer').textContent = fmtTime(pauseCur);
  }
  if (reunionRunning && reunionStartTs) {
    const dur = (Date.now() - reunionStartTs) / 1000;
    document.getElementById('fs-reunion-timer').textContent = fmtTime(dur);
  }
  // Update active stop timers in stop-desc modal
  if (stopDescKey && S.timers && S.timers[stopDescKey]) {
    const el = document.getElementById('stop-desc-timer');
    if(el) {
      const base = S.timers[stopDescKey].elapsed || 0;
      const extra = S.timers[stopDescKey].running ? (Date.now()/1000 - (S.timers[stopDescKey]._startLocal||Date.now()/1000)) : 0;
      el.textContent = fmtTime(base);
    }
  }
}

function updateStatusBar() {
  const ofEl = document.getElementById('sc-of-val');
  const stopEl = document.getElementById('sc-stop-val');
  const pcsEl = document.getElementById('sc-pcs-val');
  const trsEl = document.getElementById('sc-trs-val');
  const scOf = document.getElementById('sc-of');
  const scStop = document.getElementById('sc-stop');
  if(!ofEl) return;

  const pauseS = S.pause_total_s||0;
  const interOfS = S.inter_of_s||0;
  // of_s = elapsed + inter_of_s (pause déjà déduite côté serveur dans elapsed)
  const ofS = Math.max(0, localOfElapsed);
  ofEl.textContent = fmtTime(ofS);

  const stopS = S.stop_wall_s||0;
  stopEl.textContent = fmtTime(stopS);

  const activeCount = (S.active_stops||[]).length;
  document.getElementById('sc-stop-count').textContent = activeCount>0 ? `${activeCount} en cours` : '';

  const prod_ref = S.prod_ref||0;
  if (prod_ref>0 && ofS>0) {
    const pureS = Math.max(0, ofS - stopS - pauseS);
    const pcs = Math.floor(prod_ref * pureS / 28800);
    pcsEl.textContent = pcs.toLocaleString('fr');
    // TRS live
    const equiv = parseFloat(document.getElementById('f-qte_fab')?.value||0) || 0;
    if(equiv>0) {
      const trs = equiv/(prod_ref*ofS/28800)*100;
      trsEl.textContent = trs.toFixed(1)+'%';
    } else { trsEl.textContent='—%'; }
  } else { pcsEl.textContent='—'; trsEl.textContent='—%'; }

  // Couleur status bar selon arrêts actifs
  if(activeCount>0) {
    scOf.className='status-cell running';
    scStop.className='status-cell running';
  } else {
    scOf.className='status-cell ok';
    scStop.className=(stopS>0)?'status-cell':'status-cell';
  }
}

function renderAll() {
  // Header state
  document.getElementById('main-pilot-lbl').textContent = S.pilot||'—';
  document.getElementById('main-poste-lbl').textContent = S.poste||'—';
  const prodPilotEl = document.getElementById('prod-pilot-lbl');
  if(prodPilotEl) prodPilotEl.textContent = S.pilot||'—';

  // Views
  if (S.prod_active) {
    showView('production');
    const debutEl = document.getElementById('prod-debut-lbl');
    if(debutEl && S.of_start_iso) {
      const dt = new Date(S.of_start_iso);
      debutEl.textContent = 'Début: ' + dt.toLocaleTimeString('fr');
    }
    renderRecapStops();
    renderActiveStopsBanner();
    restoreFormFields();
  } else {
    showView('main');
    loadHistory();
  }

  // Pause fullscreen
  if (S.is_paused) {
    document.getElementById('fs-pause').classList.add('active');
    document.getElementById('btn-pause').textContent='▶ REPRENDRE';
    document.getElementById('btn-pause').style.background='var(--green)';
  } else {
    document.getElementById('fs-pause').classList.remove('active');
    document.getElementById('btn-pause').textContent='⏸ PAUSE';
    document.getElementById('btn-pause').style.background='var(--navy)';
  }

  // Prod badge
  const badge = document.getElementById('main-prod-badge');
  if(badge) {
    if(S.prod_active){
      badge.style.display='inline';
      badge.innerHTML='<span class="badge badge-red" style="animation:pulse 1.2s infinite">● EN COURS</span>';
    } else { badge.style.display='none'; }
  }

  // Header accent
  const acc = document.getElementById('hdr-accent-main');
  if(acc) acc.style.background = S.prod_active?'var(--red)':'var(--green)';

  updateStatusBar();
  renderStopSelector();
}

function showView(name) {
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  const el = document.getElementById('view-'+name);
  if(el) el.classList.add('active');
}

function openModal(id) {
  document.getElementById(id).classList.add('active');
  if(id==='modal-stop') renderStopSelector();
  if(id==='modal-nettoyage') renderNettoyageModal();
  if(id==='modal-settings') loadLists();
}
function closeModal(id) {
  document.getElementById(id).classList.remove('active');
}

// ── Login ──
async function doLogin() {
  const pilot = document.getElementById('login-pilot').value;
  const poste = document.getElementById('login-poste').value;
  const err = document.getElementById('login-err');
  if(!pilot||!poste){
    err.textContent='Veuillez sélectionner pilote et poste'; err.style.display='block'; return;
  }
  const r = await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pilot,poste})});
  const d = await r.json();
  if(d.ok){ closeModal('modal-login'); fetchState(); }
  else { err.textContent=d.error||'Erreur'; err.style.display='block'; }
}

// ── Start production ──
async function startProduction() {
  if(!S.pilot){ openModal('modal-login'); return; }
  const r = await fetch('/api/start_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
  const d = await r.json();
  if(!d.ok){ toast(d.error||'Erreur',true); return; }
  // Inter-OF dialog?
  if(d.gap_s > 30 && d.gap_s <= 28800) {
    const h=Math.floor(d.gap_s/3600), m=Math.floor((d.gap_s%3600)/60), s=Math.floor(d.gap_s%60);
    document.getElementById('iof-h').value=h;
    document.getElementById('iof-m').value=m;
    document.getElementById('iof-s').value=s;
    openModal('modal-interof');
  }
  fetchState();
}

async function confirmInterOf() {
  const h=parseInt(document.getElementById('iof-h').value)||0;
  const m=parseInt(document.getElementById('iof-m').value)||0;
  const s=parseInt(document.getElementById('iof-s').value)||0;
  const total=h*3600+m*60+s;
  await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:total})});
  closeModal('modal-interof');
  fetchState();
}

// ── Stop selector ──
function renderStopSelector() {
  const rattEvents = EVENTS_JS.filter(e=>e.cat==='ratt');
  const pbEvents   = EVENTS_JS.filter(e=>e.cat==='pb');
  const specEvents = EVENTS_JS.filter(e=>e.key==='arret_mp'||e.key==='arret_reunion');

  renderStopGrid('stop-grid-ratt',  rattEvents.filter(e=>e.key!=='arret_reunion'), 'ratt');
  renderStopGrid('stop-grid-pb',    pbEvents.filter(e=>e.key!=='arret_mp'),        'pb');
  renderStopGrid('stop-grid-special',specEvents,                                   'special');
}

function renderStopGrid(gridId, events, cat) {
  const grid = document.getElementById(gridId);
  if(!grid) return;
  grid.innerHTML='';
  events.forEach(ev=>{
    const running = (S.timers||{})[ev.key]?.running;
    const elapsed = (S.timers||{})[ev.key]?.elapsed||0;
    const btn = document.createElement('button');
    btn.className='stop-btn '+cat+(running?' running':'');
    btn.innerHTML=`<span>${ev.label}</span>${running?`<div class="timer">${fmtTime(elapsed)}</div>`:''}`;
    btn.onclick=()=>handleStopClick(ev.key, ev.cat, ev.label);
    grid.appendChild(btn);
  });
}

async function handleStopClick(key, cat, label) {
  const isRunning = (S.timers||{})[key]?.running;
  if(isRunning) {
    // Ouvrir dialog commentaire pour clôturer
    stopDescKey = key;
    document.getElementById('stop-desc-name').textContent = label;
    const elapsed = (S.timers||{})[key]?.elapsed||0;
    document.getElementById('stop-desc-timer').textContent = fmtTime(elapsed);
    document.getElementById('stop-desc-comment').value='';
    closeModal('modal-stop');
    openModal('modal-stop-desc');
  } else {
    await fetch('/api/start_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,cat})});
    fetchState();
    closeModal('modal-stop');
    toast('Arrêt démarré : '+label);
  }
}

async function confirmEndStop() {
  const comment = document.getElementById('stop-desc-comment').value.trim();
  await fetch('/api/end_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:stopDescKey,comment})});
  closeModal('modal-stop-desc');
  stopDescKey=null;
  fetchState();
  toast('Arrêt clôturé');
}

// ── Pause ──
async function togglePause() {
  const r = await fetch('/api/toggle_pause',{method:'POST'});
  const d = await r.json();
  fetchState();
}

// ── Réunion ──
function startReunion() {
  closeModal('modal-reunion-confirm');
  reunionRunning=true;
  reunionStartTs=Date.now();
  document.getElementById('fs-reunion').classList.add('active');
}
function stopReunion() {
  if(reunionStartTs) {
    const durMin = Math.round((Date.now()-reunionStartTs)/60000);
    const el = document.getElementById('f-manquant_pers');
    if(el) { const cur=parseInt(el.value||0)||0; el.value=cur+durMin; saveForm(); }
    toast(`Réunion terminée : ${durMin} min ajoutées`);
  }
  reunionRunning=false; reunionStartTs=null;
  document.getElementById('fs-reunion').classList.remove('active');
}

// ── Nettoyage ──
function openNettoyage() {
  if((S.timers||{})['nettoyage']?.running) {
    // Clôturer nettoyage
    stopDescKey='nettoyage';
    document.getElementById('stop-desc-name').textContent='Nettoyage';
    document.getElementById('stop-desc-timer').textContent=fmtTime((S.timers['nettoyage']?.elapsed||0));
    document.getElementById('stop-desc-comment').value='';
    openModal('modal-stop-desc');
    return;
  }
  openModal('modal-nettoyage');
}

function renderNettoyageModal() {
  const cfg = {short:10,long:30,grand:60};
  const body = document.getElementById('nett-body');
  const opts=[
    {type:'court',label:'🧹 Nettoyage court (poste)',sub:`${cfg.short} min tolérés`,color:'#f59e0b'},
    {type:'long', label:'🧽 Nettoyage long (ex: mercredi)',sub:`${cfg.long} min tolérés`,color:'#d97706'},
    {type:'grand',label:'✨ Grand nettoyage',sub:`${cfg.grand} min tolérés`,color:'#92400e'},
  ];
  body.innerHTML=opts.map(o=>`
    <button onclick="startNettoyage('${o.type}')" style="width:100%;margin-bottom:10px;padding:18px;background:${o.color};color:#fff;border:none;border-radius:12px;cursor:pointer;font-size:15px;font-weight:700;box-shadow:0 4px 0 rgba(0,0,0,.25)">
      ${o.label}<br><span style="font-size:12px;opacity:.8">${o.sub}</span>
    </button>
  `).join('');
}

async function startNettoyage(ntype) {
  await fetch('/api/start_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ntype})});
  closeModal('modal-nettoyage');
  fetchState();
  toast('Nettoyage démarré');
}

// ── Fin de production ──
function openEndProd() {
  // Collecter les données du formulaire
  const form = collectForm();
  if(!form.of_num) { toast('N° OF requis', true); return; }

  // Afficher recap
  pendingEndProdData = form;
  const recap_table = document.getElementById('recap-table');
  recap_table.innerHTML=`
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">OF</th><td style="padding:4px 8px;font-weight:700">${form.of_num}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Pilote</th><td style="padding:4px 8px">${form.pilote||S.pilot}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Poste</th><td style="padding:4px 8px">${form.poste||S.poste}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Taille</th><td style="padding:4px 8px">${form.taille}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Type prod</th><td style="padding:4px 8px">${form.type_prod}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Qte fabriquée</th><td style="padding:4px 8px;font-weight:700">${form.qte_fab||'—'}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Durée OF (estimée)</th><td style="padding:4px 8px">${fmtTime(localOfElapsed)}</td></tr>
    <tr><th style="text-align:left;padding:4px 8px;color:var(--gray);font-size:11px;text-transform:uppercase">Pause totale</th><td style="padding:4px 8px">${fmtTime(S.pause_total_s||0)}</td></tr>
  `;

  // Recap stops
  const stopsDetail = document.getElementById('recap-stops-detail');
  const timers = S.timers||{};
  const stopRows = EVENTS_JS.filter(e=>timers[e.key]&&timers[e.key].elapsed>0).map(e=>{
    const dur=timers[e.key].elapsed;
    return `<div style="display:flex;justify-content:space-between;margin-bottom:4px;font-size:12px"><span>${e.label}</span><strong>${fmtTime(dur)}</strong></div>`;
  }).join('');
  stopsDetail.innerHTML=stopRows||'<span style="color:var(--gray)">Aucun arrêt</span>';

  // TRS estimé
  const prod_ref=S.prod_ref||0;
  const qte_fab=parseInt(form.qte_fab)||0;
  const trsBox=document.getElementById('recap-trs-box');
  const trsVal=document.getElementById('recap-trs-val');
  if(prod_ref>0&&localOfElapsed>0&&qte_fab>0){
    const trs=qte_fab/(prod_ref*localOfElapsed/28800)*100;
    trsVal.textContent=trs.toFixed(1)+'%';
    const col=trs>=70?'var(--green)':trs>=50?'var(--amber)':'var(--red)';
    trsBox.style.background=col; trsBox.style.color='#fff';
  } else { trsVal.textContent='—'; trsBox.style.background='var(--lgray)'; trsBox.style.color='var(--dark)'; }

  openModal('modal-recap');
}

async function confirmEndProd() {
  if(!pendingEndProdData) return;
  const r = await fetch('/api/end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:pendingEndProdData})});
  const d = await r.json();
  if(d.ok){
    closeModal('modal-recap');
    toast('✔ Production enregistrée dans Excel');
    pendingEndProdData=null;
    fetchState();
  } else { toast(d.error||'Erreur',true); }
}

// ── Fin de poste ──
async function showFinDePoste() {
  const r = await fetch('/api/fin_poste_data');
  const d = await r.json();
  const stats = document.getElementById('fin-stats-row');
  const trsCol = (d.trs>=70)?'var(--green)':(d.trs>=50)?'var(--amber)':'var(--red)';
  stats.innerHTML=`
    <div class="fin-stat"><div class="fin-stat-value" style="color:${trsCol}">${d.trs>=0?d.trs.toFixed(1)+'%':'—'}</div><div class="fin-stat-label">TRS Poste</div></div>
    <div class="fin-stat"><div class="fin-stat-value">${d.nb_of}</div><div class="fin-stat-label">OF déclarés</div></div>
    <div class="fin-stat"><div class="fin-stat-value">${d.tot_equiv||'—'}</div><div class="fin-stat-label">Équiv. totale</div></div>
    <div class="fin-stat"><div class="fin-stat-value">${d.pilot||'—'}</div><div class="fin-stat-label">Pilote</div></div>
  `;
  const tbody = document.getElementById('fin-poste-tbody');
  tbody.innerHTML=(d.of_list||[]).map(row=>{
    const trsC=(row.trs>=70)?'color:var(--green)':(row.trs>=50)?'color:var(--amber)':'color:var(--red)';
    return `<tr>
      <td>${row.of}</td><td>${row.taille}</td><td>${row.type_prod}</td>
      <td>${row.qte_fab}</td><td>${row.equiv}</td>
      <td>${row.debut}</td><td>${row.fin}</td><td>${row.duree}</td>
      <td style="font-weight:700;${trsC}">${row.trs>=0?row.trs+'%':'—'}</td>
    </tr>`;
  }).join('');
  openModal('modal-fin-poste');
}

// ── Formulaire ──
function collectForm() {
  const get = id => { const el=document.getElementById(id); return el?el.value:''; };
  return {
    of_num: get('f-of_num'), taille: get('f-taille'),
    type_prod: get('f-type_prod'), code_prod: get('f-code_prod'),
    nb_pers: get('f-nb_pers'),
    pilote: S.pilot||'', poste: S.poste||'',
    copilote: get('f-copilote'), fibre: get('f-fibre'),
    poids: get('f-poids'), of_taie: get('f-of_taie'),
    traca: get('f-traca'), ref_taie: get('f-ref_taie'),
    kit: document.getElementById('f-kit')?.checked||false,
    qte_fab: get('f-qte_fab'), qte_emb: get('f-qte_emb'),
    qte_init_taie: get('f-qte_init_taie'),
    nb_taie2_choix: get('f-nb_taie2_choix'),
    nb_def_cout: get('f-nb_def_cout'),
    mq_taie: get('f-mq_taie'),
    mq_housse_encart: get('f-mq_housse_encart'),
    nb_pp_cousue: get('f-nb_pp_cousue'),
    manquant_pers: get('f-manquant_pers'),
    duree_mq_mp: get('f-duree_mq_mp'),
    comment: get('f-comment'),
  };
}

async function saveForm() {
  const data = collectForm();
  _S_form = data;
  await fetch('/api/save_form',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
}

function restoreFormFields() {
  const f = S.form||{};
  const set=(id,v)=>{ const el=document.getElementById(id); if(el&&v!==undefined&&v!=='')el.value=v; };
  set('f-of_num',f.of_num); set('f-taille',f.taille);
  set('f-type_prod',f.type_prod); set('f-code_prod',f.code_prod);
  set('f-nb_pers',f.nb_pers); set('f-copilote',f.copilote);
  set('f-fibre',f.fibre); set('f-poids',f.poids);
  set('f-of_taie',f.of_taie); set('f-traca',f.traca);
  set('f-ref_taie',f.ref_taie); set('f-qte_fab',f.qte_fab);
  set('f-qte_emb',f.qte_emb); set('f-qte_init_taie',f.qte_init_taie);
  set('f-nb_taie2_choix',f.nb_taie2_choix); set('f-nb_def_cout',f.nb_def_cout);
  set('f-mq_taie',f.mq_taie); set('f-mq_housse_encart',f.mq_housse_encart);
  set('f-nb_pp_cousue',f.nb_pp_cousue); set('f-manquant_pers',f.manquant_pers);
  set('f-duree_mq_mp',f.duree_mq_mp); set('f-comment',f.comment);
  // Pilote/poste auto
  const fPilote=document.getElementById('f-pilote');
  const fPoste=document.getElementById('f-poste');
  if(fPilote) fPilote.value=S.pilot||'';
  if(fPoste) fPoste.value=S.poste||'';
  if(f.kit) { const el=document.getElementById('f-kit'); if(el) el.checked=true; }
}

// ── Recap stops (vue production) ──
function renderRecapStops() {
  const el=document.getElementById('recap-stops-list');
  if(!el) return;
  const timers=S.timers||{};
  const rows=EVENTS_JS.filter(e=>timers[e.key]&&(timers[e.key].elapsed>0||timers[e.key].running)).map(e=>{
    const running=timers[e.key].running;
    const elapsed=timers[e.key].elapsed||0;
    const col=e.cat==='ratt'?'var(--amber)':e.cat==='nettoyage'?'var(--cyan)':'var(--red)';
    return `<div class="recap-stop-row${running?' running':''}">
      <div class="recap-stop-dot" style="background:${col}"></div>
      <div class="recap-stop-name">${e.label}</div>
      <div class="recap-stop-time">${fmtTime(elapsed)}</div>
    </div>`;
  });
  el.innerHTML=rows.length?rows.join(''):'<div style="color:var(--gray);font-size:12px;padding:8px">Aucun arrêt</div>';

  const pauseEl=document.getElementById('recap-pauses');
  if(pauseEl) pauseEl.textContent=fmtTime(S.pause_total_s||0)+' total';
}

// ── Active stops banner ──
function renderActiveStopsBanner() {
  const bar=document.getElementById('active-stops-bar');
  const active=S.active_stops||[];
  if(active.length===0){ bar.style.display='none'; return; }
  bar.style.display='flex';
  bar.innerHTML='<span style="color:#94a3b8;font-size:11px;font-weight:700;margin-right:4px">ARRÊTS ACTIFS:</span>'+
    active.map(k=>{
      const ev=EVENTS_JS.find(e=>e.key===k);
      const lbl=ev?ev.label:k;
      const elapsed=(S.timers||{})[k]?.elapsed||0;
      const isnett=k==='nettoyage';
      return `<div class="active-stop-chip${isnett?' nett':''}" onclick="handleStopClick('${k}','${ev?.cat||'pb'}','${lbl}')">${lbl} — ${fmtTime(elapsed)}</div>`;
    }).join('');
}

// ── History tables ──
async function loadHistory() {
  try {
    const r=await fetch('/api/history'); const rows=await r.json();
    const tbody=document.getElementById('decl-tbody');
    if(!tbody) return;
    if(!rows.length){ tbody.innerHTML='<tr><td colspan="12" style="color:#94a3b8;padding:20px">Aucune déclaration</td></tr>'; return; }
    tbody.innerHTML=rows.slice(0,80).map(row=>{
      const trs=row.trs;
      const cls=trs>=70?'trs-hi':trs>=50?'trs-warn':trs>=0?'trs-low':'';
      return `<tr class="${cls}">
        <td>${row.date}</td><td>${row.debut}</td><td>${row.fin}</td>
        <td><strong>${row.of}</strong></td><td>${row.pilote}</td><td>${row.poste}</td>
        <td>${row.taille}</td><td>${row.type_prod}</td>
        <td>${row.qte_fab}</td><td>${row.equiv}</td>
        <td><strong>${trs>=0?trs.toFixed(1)+'%':'—'}</strong></td><td>${row.duree}</td>
      </tr>`;
    }).join('');
  } catch(e){}
}

async function loadEventsTable() {
  try {
    const r=await fetch('/api/events_list'); const rows=await r.json();
    const tbody=document.getElementById('evts-tbody');
    if(!tbody) return;
    if(!rows.length){ tbody.innerHTML='<tr><td colspan="9" style="color:#94a3b8;padding:20px">Aucun événement</td></tr>'; return; }
    tbody.innerHTML=rows.slice(0,80).map(row=>{
      const isHors=row.hors_trs;
      return `<tr style="${isHors?'background:#dcfce7':''}">
        <td>${row.type}</td><td>${row.pilote}</td><td>${row.of}</td>
        <td>${row.date}</td><td>${row.debut}</td><td>${row.fin}</td>
        <td>${row.duree}</td><td style="text-align:left">${row.comment}</td>
        <td>${isHors?'<span class="badge badge-green">✔ Hors TRS</span>':'—'}</td>
      </tr>`;
    }).join('');
  } catch(e){}
}

// ── Tabs ──
function switchTab(tabId, btn) {
  ['tab-decl','tab-evts'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.style.display=id===tabId?'block':'none';
  });
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  btn.classList.add('active');
  if(tabId==='tab-evts') loadEventsTable();
}

// ── Settings ──
async function saveSettings() {
  const pw=document.getElementById('set-pw').value;
  const err=document.getElementById('settings-err');
  const data={
    pw,
    prod_ref: document.getElementById('set-prod_ref').value,
    pause_max_min: document.getElementById('set-pause_max_min').value,
    clean_short_min: document.getElementById('set-clean_short_min').value,
    clean_long_min: document.getElementById('set-clean_long_min').value,
    clean_grand_min: document.getElementById('set-clean_grand_min').value,
    meeting_tol_min: document.getElementById('set-meeting_tol_min').value,
    new_pw: document.getElementById('set-new_pw').value,
  };
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  const d=await r.json();
  if(d.ok){ closeModal('modal-settings'); toast('Paramètres sauvegardés'); }
  else { err.textContent=d.error||'Erreur'; err.style.display='block'; }
}

function openSetDb() { closeModal('modal-login'); openModal('modal-set-db'); }
async function doSetDb() {
  const pw=document.getElementById('setdb-pw').value;
  const path=document.getElementById('setdb-path').value;
  const err=document.getElementById('setdb-err');
  const r=await fetch('/api/set_db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,path})});
  const d=await r.json();
  if(d.ok){ closeModal('modal-set-db'); toast('Base de données chargée'); fetchState(); loadLists(); openModal('modal-login'); }
  else { err.textContent=d.error||'Erreur'; err.style.display='block'; }
}

async function doReload() {
  await fetch('/api/reload',{method:'POST'});
  fetchState(); loadLists(); toast('Données actualisées');
}

// ── TRS gauge ──
function updateGauge(trs) {
  const arc=document.getElementById('gauge-arc');
  const pct=document.getElementById('main-trs-pct');
  if(!arc||!pct) return;
  if(trs<0){ pct.textContent='—'; arc.style.stroke='#dde4ef'; return; }
  const totalLen=188.5;
  const offset=totalLen-(Math.min(100,trs)/100*totalLen);
  arc.style.strokeDashoffset=offset;
  const col=trs>=70?'#1a8c4e':trs>=50?'#d97706':'#e31e24';
  arc.style.stroke=col;
  pct.textContent=trs.toFixed(1)+'%';
  pct.style.color=col;
}

// ── Utils ──
function fmtTime(s){
  s=Math.max(0,Math.floor(s||0));
  const h=Math.floor(s/3600), m=Math.floor((s%3600)/60), sec=s%60;
  return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(sec).padStart(2,'0')}`;
}

function toast(msg, isErr=false) {
  const el=document.getElementById('toast');
  el.textContent=msg;
  el.style.background=isErr?'#e31e24':'#1a8c4e';
  el.classList.add('show');
  setTimeout(()=>el.classList.remove('show'),3000);
}

// Constantes events côté JS
const EVENTS_JS = [
  {label:'Pochon / Fibre',key:'ratt_pochon',cat:'ratt'},
  {label:'Couture',key:'ratt_couture',cat:'ratt'},
  {label:'Emballage',key:'ratt_emb',cat:'ratt'},
  {label:'Presse Souder',key:'ratt_presse_soud',cat:'ratt'},
  {label:'Presse ZIP',key:'ratt_presse_zip',cat:'ratt'},
  {label:'Nettoyage',key:'nettoyage',cat:'nettoyage'},
  {label:'Chargeuse',key:'pb_chargeuse',cat:'pb'},
  {label:'Carde',key:'pb_carde',cat:'pb'},
  {label:'Etaleur / Tour',key:'pb_etaleur',cat:'pb'},
  {label:'Coupe / Circ.',key:'pb_coupe',cat:'pb'},
  {label:'Tapis Bascule',key:'pb_tapis1',cat:'pb'},
  {label:'Enrouleur Pochon',key:'pb_enrouleur',cat:'pb'},
  {label:'Pesee / Tapis 2',key:'pb_pesee',cat:'pb'},
  {label:'Deviation / Table',key:'pb_deviation',cat:'pb'},
  {label:'Enfileur Pochon',key:'pb_enfileur',cat:'pb'},
  {label:'Kinna / Stroebel',key:'pb_kinna',cat:'pb'},
  {label:'Tapeuse',key:'pb_tapeuse',cat:'pb'},
  {label:'Table Rot. / Twin',key:'pb_table_rot',cat:'pb'},
  {label:'Enfileuse H100',key:'pb_h100',cat:'pb'},
  {label:'Enfileuse Traversin',key:'pb_traversin',cat:'pb'},
  {label:'Presse ORC',key:'pb_presse_orc',cat:'pb'},
  {label:'Presse Housse ZIP',key:'pb_presse_zip2',cat:'pb'},
  {label:'Cercleuse',key:'pb_cercleuse',cat:'pb'},
  {label:'Enrouleuse Traversin',key:'pb_enrouleuse',cat:'pb'},
  {label:'Matiere premiere',key:'arret_mp',cat:'pb'},
  {label:'Reunion',key:'arret_reunion',cat:'ratt'},
];
</script>
</body>
</html>"""


# ── Point d'entrée ────────────────────────────────────────────────────────────
def main():
    global cfg
    cfg = load_cfg()
    load_lists()
    threading.Thread(target=load_history, daemon=True).start()
    if not load_session():
        pass  # pas de session en cours

    # Vérifier si session interrompue
    session_restored = os.path.exists(SESSION_FILE)

    flask_thread = threading.Thread(
        target=lambda: flask_app.run(
            host='127.0.0.1', port=5001,
            debug=False, use_reloader=False, threaded=True
        ),
        daemon=True
    )
    flask_thread.start()

    import time
    time.sleep(0.8)

    try:
        import webview
        window = webview.create_window(
            'KPI-ORC | Ligne ORC1',
            'http://127.0.0.1:5001',
            width=1440, height=860,
            min_size=(1024, 640),
            resizable=True,
        )
        webview.start(debug=False)
    except ImportError:
        # Fallback si pywebview pas dispo : ouvrir dans le navigateur système
        import webbrowser, time as _t
        _t.sleep(0.3)
        webbrowser.open('http://127.0.0.1:5001')
        print("KPI-ORC démarré sur http://127.0.0.1:5001")
        print("Appuyez sur Ctrl+C pour quitter.")
        try:
            while True: _t.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == '__main__':
    main()
