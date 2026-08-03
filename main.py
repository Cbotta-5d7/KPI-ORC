"""KPI-ORC v6.4 - Flask + pywebview"""
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

# Nouveau schéma unifié - 40 colonnes
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
    "Duree Arrets","Duree Prod Pure","Date_poste",
    "","Degrade_min","Nbr pièces théorique",
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

def _get_or_create_listes_ws(wb):
    if "Listes" not in wb.sheetnames:
        wb.create_sheet("Listes")
    return wb["Listes"]

def write_events_to_excel(ev_list):
    """Écrit la liste des arrêts dans l'onglet Listes col K=label, L=cat."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _get_or_create_listes_ws(wb)
                ws.cell(1, 11).value = "Arrêts"
                ws.cell(1, 12).value = "Type arrêt"
                max_r = max(ws.max_row, len(ev_list) + 2)
                for ri in range(2, max_r + 2):
                    ws.cell(ri, 11).value = None
                    ws.cell(ri, 12).value = None
                for ri, ev in enumerate(ev_list, start=2):
                    ws.cell(ri, 11).value = ev.get("label","")
                    ws.cell(ri, 12).value = ev.get("cat","pb")
                _safe_excel_save(wb, path)
            threading.Thread(target=load_lists, daemon=True).start()
        except: pass
    threading.Thread(target=_bg, daemon=True).start()

def write_interposte_to_excel(labels):
    """Écrit les labels interposte dans l'onglet Listes col M."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _get_or_create_listes_ws(wb)
                ws.cell(1, 13).value = "Interposte"
                max_r = max(ws.max_row, len(labels) + 2)
                for ri in range(2, max_r + 2):
                    ws.cell(ri, 13).value = None
                for ri, lbl in enumerate(labels, start=2):
                    ws.cell(ri, 13).value = lbl
                _safe_excel_save(wb, path)
        except: pass
    threading.Thread(target=_bg, daemon=True).start()

_ARRETS_PREVUS_KEYS = ["clean_short_min","clean_long_min","clean_grand_min","meeting_tol_min","pause_min"]

def write_arrets_prevus_to_excel():
    """Écrit les budgets arrêts prévus dans l'onglet Listes col N (clé=valeur)."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _get_or_create_listes_ws(wb)
                ws.cell(1, 14).value = "Arrêts prévus"
                for ri in range(2, len(_ARRETS_PREVUS_KEYS) + 3):
                    ws.cell(ri, 14).value = None
                for ri, k in enumerate(_ARRETS_PREVUS_KEYS, start=2):
                    ws.cell(ri, 14).value = f"{k}={cfg.get(k, 0)}"
                _safe_excel_save(wb, path)
        except: pass
    threading.Thread(target=_bg, daemon=True).start()

def write_degrade_list_to_excel():
    """Écrit les motifs mode dégradé dans l'onglet Listes col O."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg2():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _get_or_create_listes_ws(wb)
                ws.cell(1, 15).value = "Mode dégradé"
                motifs = cfg.get("degrade_motifs", [])
                for ri in range(2, max(len(motifs)+3, 20)):
                    ws.cell(ri, 15).value = None
                for ri, m in enumerate(motifs, start=2):
                    ws.cell(ri, 15).value = m
                _safe_excel_save(wb, path)
        except: pass
    threading.Thread(target=_bg2, daemon=True).start()

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
    "postes_row_num": None,
    "shift_debut_dt": None,
    "shift_fin_dt": None,
    "tot_prod_s": 0.0,
    "budget_overrides": {},
    "degrade_active": False,
    "degrade_type": "",
    "degrade_start_dt": None,
    "degrade_periods": [],
}
_excel_lock = threading.Lock()
_lists = {}
_decl_cache = []   # liste de (row_num, row_data) - toutes déclarations (prod + events)
_prod_ref_cached = 0.0
_excel_busy = False  # True quand le fichier Excel est verrouillé (ouvert par Excel)
cfg = {}

flask_app = Flask(__name__)

@flask_app.after_request
def _add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

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

def _sec_to_hm(s):
    """Convert seconds to HH:MM string."""
    s = int(max(0, s)) % 86400
    return f"{s//3600:02d}:{(s%3600)//60:02d}"

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
    """Retourne (debut_str, fin_str) du modèle horaire pour le poste/jour donné.
    Vérifie d'abord les surcharges de session (_S['model_overrides']) avant cfg."""
    day_map = {0:'lun',1:'mar',2:'mer',3:'jeu',4:'ven',5:'sam',6:'dim'}
    day_key = day_map.get((date_obj or datetime.date.today()).weekday(), 'lun')
    # Session override (temporary, never saved to disk)
    overrides = _S.get("model_overrides", {})
    if poste and poste in overrides and day_key in overrides[poste]:
        ov = overrides[poste][day_key]
        return ov.get("debut","05:00"), ov.get("fin","13:00")
    models = cfg.get("modeles_horaires", [])
    model = next((m for m in models if str(m.get("nom","")).strip().lower()==str(poste or "").strip().lower()), None)
    if not model: return None, None
    jours = model.get("jours", {})
    if date_obj and jours:
        day_cfg = jours.get(day_key)
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

def _get_arret_budget_key(label):
    """Associe un libellé d'arrêt à une clé de budget plannifié."""
    l = str(label or "").lower()
    if "nettoyage" in l or "nett" in l:
        if "très long" in l or "tres long" in l or "grand" in l: return "clean_grand_min"
        if "long" in l: return "clean_long_min"
        if "court" in l: return "clean_short_min"
    if "réunion" in l or "reunion" in l or "meeting" in l: return "meeting_tol_min"
    if "pause" in l: return "pause_min"
    return None

def _is_degrade_type(t):
    """Retourne True si le type est un motif de mode dégradé configuré."""
    motifs = cfg.get("degrade_motifs", [])
    return bool(motifs) and str(t or "").strip() in motifs

def _norm_fin(deb_s, fin_s):
    """Normalise fin_s pour les événements chevauchant minuit (fin < deb → +86400)."""
    return fin_s + 86400 if fin_s < deb_s else fin_s

def _merged_degrade_s(rows):
    """Retourne les secondes de dégradé dédupliquées en fusionnant les intervalles qui se chevauchent.
    Évite le double-comptage quand api_stop_degrade ET build_decl_rows écrivent des lignes Formation pour la même période.
    rows : liste de lignes brutes (pas de paires (rn, row))."""
    raw = [(_hms_to_sec(str(r[16] or "00:00:00")), _hms_to_sec(str(r[17] or "00:00:00")))
           for r in rows if _is_degrade_type(str(r[0] or ""))]
    ivs = sorted((s, _norm_fin(s, f)) for s, f in raw if _norm_fin(s, f) > s)
    mg = []
    for s, f in ivs:
        if mg and s <= mg[-1][1]: mg[-1] = (mg[-1][0], max(mg[-1][1], f))
        else: mg.append((s, f))
    return sum(f - s for s, f in mg)

_pers_pct_map = {}  # {nb_pers_int: pct_float}  e.g. {1: 0.10, 10: 1.00}

def _merged_degrade_ivs(rows):
    """Returns merged dégradé intervals [(start_s, end_s)] from raw rows."""
    raw = [(_hms_to_sec(str(r[16] or "00:00:00")), _hms_to_sec(str(r[17] or "00:00:00")))
           for r in rows if _is_degrade_type(str(r[0] or ""))]
    ivs = sorted((s, _norm_fin(s, f)) for s, f in raw if _norm_fin(s, f) > s)
    mg = []
    for s, f in ivs:
        if mg and s <= mg[-1][1]: mg[-1] = (mg[-1][0], max(mg[-1][1], f))
        else: mg.append((s, f))
    return mg

def get_pct_cadence(nb_pers):
    """Returns pct_cadence factor for nb_pers (1.0 = 100% cadence). Falls back to 1.0."""
    if not _pers_pct_map: return 1.0
    try:
        n = int(float(str(nb_pers or 1)))
        if n in _pers_pct_map: return float(_pers_pct_map[n])
    except: pass
    return 1.0

def _deg_overlap_s(of_start_s, of_end_s, deg_ivs):
    """Returns seconds of dégradé overlap with OF interval [of_start_s, of_end_s]."""
    total = 0.0
    for ds, df in deg_ivs:
        o0 = max(ds, of_start_s); o1 = min(df, of_end_s)
        if o1 > o0: total += o1 - o0
    return total

def _option_b_trs(prod_raw_rows, deg_ivs, prod_ref, plan_ivs=None):
    """Option B TRS: per-OF adjusted time × pct_cadence(nb_pers).
    Returns (trs_float, sum_expected_equiv)."""
    if prod_ref <= 0: return -1.0, 0.0
    tot_equiv = 0.0; sum_expected = 0.0
    for r in prod_raw_rows:
        try:
            eq = float(str(r[21] or 0).replace(",", "."))
            deb_s = _hms_to_sec(str(r[16] or "00:00:00"))
            fin_s = _norm_fin(deb_s, _hms_to_sec(str(r[17] or "00:00:00")))
            dur_s = fin_s - deb_s if fin_s > deb_s else _hms_to_sec(str(r[18] or "00:00:00"))
            if dur_s <= 0: continue
            nb_p = r[6] if len(r) > 6 else 1
            pct = get_pct_cadence(nb_p)
            ovl = _deg_overlap_s(deb_s, fin_s, deg_ivs)
            plan_ovl = _deg_overlap_s(deb_s, fin_s, plan_ivs) if plan_ivs else 0.0
            adj_s = max(1.0, dur_s - ovl - plan_ovl)
            sum_expected += prod_ref * pct * adj_s / 28800
            tot_equiv += eq
        except: pass
    if sum_expected <= 0: return -1.0, 0.0
    return round(tot_equiv / sum_expected * 100, 1), sum_expected

def _compute_planned_deduction_s(evt_rows, overrides=None):
    """Calcule les secondes à déduire de l'elapsed TRS pour les arrêts planifiés.
    evt_rows : liste de tuples (rn, r) issus de _decl_cache OU liste de dicts {"type","duree"}.
    overrides : dict optionnel {key: minutes} — prioritaire sur cfg (session en cours uniquement).
    """
    _ov = overrides if overrides is not None else {}
    budgets = {}
    for k in ("clean_short_min", "clean_long_min", "clean_grand_min", "meeting_tol_min", "pause_min"):
        raw = _ov.get(k)
        budgets[k] = float((raw if raw is not None else cfg.get(k, 0)) or 0) * 60
    if all(v == 0 for v in budgets.values()):
        return 0.0
    actual = {}
    for row in evt_rows:
        if isinstance(row, dict):
            lbl  = str(row.get("type","") or "")
            dur_s = _hms_to_sec(str(row.get("duree","00:00:00") or "00:00:00"))
        elif isinstance(row, (list, tuple)):
            # format: (rn, r) or just r
            r = row[1] if len(row) == 2 and isinstance(row[0], int) else row
            lbl  = str(r[0] or "")
            dur_s = _hms_to_sec(str(r[18] or "00:00:00"))
        else:
            continue
        key = _get_arret_budget_key(lbl)
        if key:
            actual[key] = actual.get(key, 0.0) + dur_s
    return sum(min(actual.get(k, 0.0), b) for k, b in budgets.items())

def _compute_budget_state_now():
    """Calcule l'état des budgets arrêts prévus pour le poste en cours.
    Retourne per_type (consumed_s, budget_s, of_consumed_s, of/shift_deductible_s)
    + total_shift_deductible_s + total_of_deductible_s.
    """
    BUDGET_KEYS = {
        "pause_min":       "Pause",
        "meeting_tol_min": "Réunion",
        "clean_short_min": "Nettoyage court",
        "clean_long_min":  "Nettoyage long",
        "clean_grand_min": "Nettoyage très long",
    }
    _ov = _S.get("budget_overrides", {})
    budgets_s = {lbl: float((_ov.get(k) if _ov.get(k) is not None else cfg.get(k, 0)) or 0) * 60 for k, lbl in BUDGET_KEYS.items()}
    shift_consumed = {lbl: 0.0 for lbl in BUDGET_KEYS.values()}
    of_consumed    = {lbl: 0.0 for lbl in BUDGET_KEYS.values()}

    pilot = _S.get("pilot", "")
    shift_start = _S.get("shift_start")
    today_str = datetime.date.today().strftime("%d/%m/%Y")
    shift_date_str = (shift_start.date() if shift_start else datetime.date.today()).strftime("%d/%m/%Y")

    # Key → label mapping pour les arrêts configurés
    try:
        _evts_cfg = get_events_list()
        _key_lbl = {e.get("key",""): e.get("label","") for e in _evts_cfg}
    except:
        _key_lbl = {}

    def _planned_lbl_from_raw(raw_lbl):
        bk = _get_arret_budget_key(raw_lbl)
        return BUDGET_KEYS.get(bk)

    def _planned_lbl_from_tl(ev):
        cat = ev.get("cat","")
        if cat == "nettoyage":
            ntype = ev.get("nettoyage_type","court")
            bk = {"court":"clean_short_min","long":"clean_long_min","grand":"clean_grand_min"}.get(ntype,"clean_short_min")
            return BUDGET_KEYS.get(bk)
        if cat == "reunion":
            return BUDGET_KEYS.get("meeting_tol_min")
        lbl = _key_lbl.get(ev.get("key",""), ev.get("key",""))
        bk = _get_arret_budget_key(lbl)
        return BUDGET_KEYS.get(bk)

    # 1. Événements des OFs passés (déjà écrits dans Excel)
    for rn, r in _decl_cache:
        rd = _row_date(r[2])
        if rd != shift_date_str and rd != today_str: continue
        if str(r[4] or "") != pilot: continue
        if str(r[0] or "").strip().lower() in ("production","prod",""): continue
        pl = _planned_lbl_from_raw(str(r[0] or "").strip())
        if pl:
            shift_consumed[pl] = shift_consumed.get(pl, 0.0) + _hms_to_sec(str(r[18] or "00:00:00"))

    # 2. OF en cours : tl_events (terminés et actifs)
    for ev in _S.get("tl_events", []):
        if not ev.get("start"): continue
        pl = _planned_lbl_from_tl(ev)
        if pl:
            end = ev.get("end") or datetime.datetime.now()
            dur = (end - ev["start"]).total_seconds()
            of_consumed[pl] = of_consumed.get(pl, 0.0) + dur
            shift_consumed[pl] = shift_consumed.get(pl, 0.0) + dur

    # 3. Pause OF en cours (hors tl_events)
    pause_s = _S.get("pause_total_s", 0.0)
    if _S.get("is_paused") and _S.get("pause_start"):
        pause_s += (datetime.datetime.now() - _S["pause_start"]).total_seconds()
    of_consumed["Pause"] = of_consumed.get("Pause", 0.0) + pause_s
    shift_consumed["Pause"] = shift_consumed.get("Pause", 0.0) + pause_s

    # 4. Calcul des déductibles
    per_type = {}
    total_shift_ded = 0.0
    total_of_ded = 0.0
    for lbl in BUDGET_KEYS.values():
        budget = budgets_s.get(lbl, 0.0)
        s_cons = shift_consumed.get(lbl, 0.0)
        of_cons = of_consumed.get(lbl, 0.0)
        past = s_cons - of_cons
        remaining = max(0.0, budget - past)
        of_ded = min(of_cons, remaining)
        shift_ded = min(s_cons, budget)
        total_of_ded += of_ded
        total_shift_ded += shift_ded
        per_type[lbl] = {
            "consumed_s": round(s_cons, 1),
            "budget_s": round(budget, 0),
            "of_consumed_s": round(of_cons, 1),
            "of_deductible_s": round(of_ded, 1),
            "shift_deductible_s": round(shift_ded, 1),
        }
    # Déductible hors OF en cours (pour éviter le bug TRS 150% quand arrêt prévu pendant OF actif)
    total_shift_ded_before_of = 0.0
    for lbl in BUDGET_KEYS.values():
        budget = budgets_s.get(lbl, 0.0)
        s_cons_before = max(0.0, shift_consumed.get(lbl, 0.0) - of_consumed.get(lbl, 0.0))
        total_shift_ded_before_of += min(s_cons_before, budget)
    return {
        "per_type": per_type,
        "total_shift_deductible_s": round(total_shift_ded, 1),
        "total_of_deductible_s": round(total_of_ded, 1),
        "shift_deductible_before_of_s": round(total_shift_ded_before_of, 1),
    }

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
            "postes_row_num": _S.get("postes_row_num"),
            "shift_debut_dt": _dt_str(_S.get("shift_debut_dt")),
            "shift_fin_dt": _dt_str(_S.get("shift_fin_dt")),
            "tot_prod_s": _S.get("tot_prod_s", 0.0),
            "budget_overrides": _S.get("budget_overrides", {}),
            "degrade_active": _S.get("degrade_active", False),
            "degrade_type": _S.get("degrade_type", ""),
            "degrade_start_dt": _dt_str(_S.get("degrade_start_dt")),
            "degrade_periods": [{"start": _dt_str(p["start"]), "end": _dt_str(p["end"]), "type": p["type"]} for p in _S.get("degrade_periods", [])],
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
        _S["postes_row_num"] = d.get("postes_row_num")
        _S["shift_debut_dt"] = _str_dt(d.get("shift_debut_dt"))
        _S["shift_fin_dt"]   = _str_dt(d.get("shift_fin_dt"))
        _S["tot_prod_s"]     = float(d.get("tot_prod_s", 0))
        _S["budget_overrides"] = d.get("budget_overrides", {})
        _S["degrade_active"] = d.get("degrade_active", False)
        _S["degrade_type"]   = d.get("degrade_type", "")
        _S["degrade_start_dt"] = _str_dt(d.get("degrade_start_dt"))
        _S["degrade_periods"] = [{"start": _str_dt(p.get("start")), "end": _str_dt(p.get("end")), "type": p.get("type","")} for p in d.get("degrade_periods", [])]
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
            # Col K (11)=label, L (12)=cat : liste des arrêts configurables
            evts_k = []
            for ri in range(2, ws.max_row+1):
                lbl_v = ws.cell(ri, 11).value
                cat_v = ws.cell(ri, 12).value
                if lbl_v is not None and str(lbl_v).strip():
                    lbl = str(lbl_v).strip()
                    # Rétro-compat : ancien format "label|cat" en col K seule
                    if "|" in lbl and not cat_v:
                        parts = lbl.split("|"); lbl = parts[0].strip(); cat_v = parts[1].strip()
                    cat = str(cat_v or "pb").strip() or "pb"
                    key = lbl.lower().replace(" ","_").replace("/","_").replace("é","e").replace("è","e").replace("ê","e").replace("à","a").replace("ç","c")[:28]
                    evts_k.append({"label": lbl, "key": key, "cat": cat})
            if evts_k:
                _lists["arrêts_k"] = evts_k
                cfg["events_list"] = evts_k
                save_cfg_data()
            # Col M (13) : labels interposte
            ipl = []
            for ri in range(2, ws.max_row+1):
                v = ws.cell(ri, 13).value
                if v is not None and str(v).strip():
                    ipl.append(str(v).strip())
            if ipl:
                cfg["interposte_labels"] = ipl
                save_cfg_data()
            # Col N (14) : arrêts prévus (format "clé=valeur")
            for ri in range(2, ws.max_row+1):
                v = ws.cell(ri, 14).value
                if v is not None and str(v).strip():
                    try:
                        k, val = str(v).strip().split("=", 1)
                        k = k.strip(); val = val.strip()
                        if k in _ARRETS_PREVUS_KEYS:
                            cfg[k] = float(val)
                    except: pass
            save_cfg_data()
            # Col O (15) : motifs mode dégradé (depuis row 1 pour accepter saisie manuelle)
            _deg_motifs = []
            for ri in range(1, ws.max_row+1):
                _ov2 = ws.cell(ri, 15).value
                if _ov2 is None: continue
                _ov2_s = str(_ov2).strip()
                if not _ov2_s or _ov2_s.lower() in ("mode dégradé", "mode degrade"): continue
                _deg_motifs.append(_ov2_s)
            if _deg_motifs:
                cfg["degrade_motifs"] = _deg_motifs
                save_cfg_data()
            # Col P (16) = nb_pers, Col Q (17) = % cadence attendu
            _pers_map_new = {}
            for ri in range(2, ws.max_row+1):
                _pv = ws.cell(ri, 16).value
                _qv = ws.cell(ri, 17).value
                if _pv is None and _qv is None: continue
                try:
                    _np = int(float(str(_pv or "").strip()))
                    _pct_raw = float(str(_qv or "").strip().replace(",",".").replace("%","").strip())
                    if 1 <= _np <= 10 and 0 < _pct_raw <= 200:
                        _pers_map_new[_np] = _pct_raw / 100.0 if _pct_raw > 2 else _pct_raw
                except: pass
            if _pers_map_new:
                _pers_pct_map.clear(); _pers_pct_map.update(_pers_map_new)
            # Col G (7) row 2 : MDP admin
            _adm_pw_v = ws.cell(2, 7).value
            if _adm_pw_v is not None and str(_adm_pw_v).strip():
                cfg["supervisor_pw"] = str(_adm_pw_v).strip()
                save_cfg_data()
        wb.close()
    except: pass

def write_pers_pct_to_excel():
    """Persiste _pers_pct_map dans l'onglet Listes, colonnes P (16) et Q (17)."""
    path = cfg.get("db_path", "")
    if not path or not os.path.exists(path): return
    try:
        wb = load_workbook(path, read_only=False, data_only=False)
        if "Listes" not in wb.sheetnames: wb.close(); return
        ws = wb["Listes"]
        if not ws.cell(1, 16).value: ws.cell(1, 16).value = "Nombre de personne"
        if not ws.cell(1, 17).value: ws.cell(1, 17).value = "% cadence attendu"
        for ri in range(2, 15):
            ws.cell(ri, 16).value = None; ws.cell(ri, 17).value = None
        ri = 2
        for np_k in sorted(_pers_pct_map.keys()):
            ws.cell(ri, 16).value = np_k
            ws.cell(ri, 17).value = round(_pers_pct_map[np_k] * 100, 1)
            ri += 1
        wb.save(path); wb.close()
    except: pass

def write_admin_pw_to_excel(pw):
    """Écrit le MDP admin en G2 de l'onglet Listes."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return False
    try:
        with _excel_lock:
            wb = _get_wb(path)
            if wb is None: return False
            if "Listes" not in wb.sheetnames: wb.close(); return False
            ws = wb["Listes"]
            if not ws.cell(1,7).value: ws.cell(1,7).value = "MDP Admin"
            ws.cell(2,7).value = pw
            _safe_excel_save(wb, path)
        return True
    except: return False

def write_simple_list_to_excel(col_idx, header, items):
    """Écrit une liste dans une colonne de l'onglet Listes (row 2+), efface l'ancien contenu."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return False
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Listes" not in wb.sheetnames: wb.close(); return
                ws = wb["Listes"]
                if not ws.cell(1, col_idx).value: ws.cell(1, col_idx).value = header
                clear_until = max(ws.max_row, len(items) + 5)
                for ri in range(2, clear_until + 1):
                    ws.cell(ri, col_idx).value = None
                for ri, val in enumerate(items, start=2):
                    ws.cell(ri, col_idx).value = val
                _safe_excel_save(wb, path)
        except: pass
        threading.Thread(target=load_lists, daemon=True).start()
    threading.Thread(target=_bg, daemon=True).start()
    return True

def write_equiv_list_to_excel(items):
    """Écrit la liste type_produit (col E=5) + coeff (col F=6) de façon synchronisée."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return False
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                if "Listes" not in wb.sheetnames: wb.close(); return
                ws = wb["Listes"]
                if not ws.cell(1,5).value: ws.cell(1,5).value = "Type Produit"
                if not ws.cell(1,6).value: ws.cell(1,6).value = "Equivalence"
                clear_until = max(ws.max_row, len(items) + 5)
                for ri in range(2, clear_until + 1):
                    ws.cell(ri, 5).value = None
                    ws.cell(ri, 6).value = None
                for ri, it in enumerate(items, start=2):
                    ws.cell(ri, 5).value = it.get("type","")
                    ws.cell(ri, 6).value = it.get("coeff","")
                _safe_excel_save(wb, path)
        except: pass
        threading.Thread(target=load_lists, daemon=True).start()
    threading.Thread(target=_bg, daemon=True).start()
    return True

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
    """Construit les lignes arrêts/pauses au format unifié (40 cols)."""
    rows = []
    kit_val = "Oui" if v.get("kit") else "Non"
    def _base_row(type_decl, start, end, comment="", hors_trs=""):
        dur = max(0,(end-start).total_seconds())
        shift_dt = _S.get("shift_start") or start
        shift_date_str = shift_dt.strftime("%d/%m/%Y")
        return [
            type_decl,                          # 0 Type
            v.get("of_num",""),                 # 1 OF
            start.strftime("%d/%m/%Y"),          # 2 Date
            v.get("poste",""),                  # 3 Poste
            v.get("pilote",""),                 # 4 Pilote
            v.get("copilote",""),               # 5 Co-Pilote
            v.get("nb_pers",""),                # 6 Nb Personnes
            v.get("taille",""),                 # 7 Taille
            v.get("code_prod",""),              # 8 Code Produit
            v.get("type_prod",""),              # 9 Type Produit
            v.get("poids",""),                  # 10 Poids Garnissage
            v.get("fibre",""),                  # 11 Fibre
            v.get("of_taie",""),                # 12 OF Taie
            v.get("traca",""),                  # 13 Traca Fibre
            v.get("ref_taie",""),               # 14 Ref Taie
            kit_val,                            # 15 Kit
            start.strftime("%H:%M:%S"),          # 16 Heure Debut
            end.strftime("%H:%M:%S"),            # 17 Heure Fin
            fmt(dur),                           # 18 Duree
            "","","","","",                     # 19-23 prod only
            "",                                 # 24 TRS%
            "","","","","","",                  # 25-30 prod only
            "","","",                           # 31-33
            "",                                 # 34
            comment,                            # 35 Commentaire
            hors_trs,                           # 36 Prevu/Hors TRS
            "","",                              # 37-38 Duree Arrets, Duree Prod Pure
            shift_date_str,                     # 39 Date_poste
            "",                                 # 40 AO
            "",                                 # 41 Degrade_min AP
        ]
    for ev in tl_events:
        if ev.get("cat") not in ("ratt","pb","nettoyage","reunion","autre","interposte"): continue
        if not ev.get("key") or ev["key"].startswith("_"): continue
        start = ev.get("start")
        if not start: continue  # event sans timestamp = invalide
        if of_start and start < of_start and ev.get("key")!="arret_interposte": continue
        end = ev.get("end") or datetime.datetime.now()
        if ev["cat"]=="reunion":
            label = "Réunion"
        elif ev["key"]=="nettoyage":
            ntype = ev.get("nettoyage_type","court")
            label = {"court":"Nettoyage court","long":"Nettoyage long","grand":"Grand nettoyage"}.get(ntype,"Nettoyage court")
        elif ev["cat"]=="autre":
            label = ev["key"]  # Custom stop name typed by user
        elif ev["cat"]=="interposte":
            _dyn = get_events_list()
            lbl = (next((e["label"] for e in _dyn if isinstance(e,dict) and e.get("key")==ev["key"]), None)
                   or next((e[0] for e in INTERPOSTE_CATS if e[1]==ev["key"]), None)
                   or ev["key"])
            label = lbl
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
                    if prod_row:
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
    shift_dt = _S.get("shift_start") or start_dt
    shift_date_str = shift_dt.strftime("%d/%m/%Y")
    row = [
        row_type,"",start_dt.strftime("%d/%m/%Y"),
        _S.get("poste",""),pilot,"","","","","","","","","","","",
        start_dt.strftime("%H:%M:%S"),end_dt.strftime("%H:%M:%S"),fmt(dur_s),
        "","","","","","","","","","","","","","","","",comment,
        "","","",shift_date_str,"","",
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

POSTES_HEADERS = ["Date","Pilote","Co-Pilote","Poste","Nb OF","Prod Total (pièces)","Prod Totale (equiv)","TRS Poste %","Cadence (equiv/h)","Total Pauses (min)","Nettoyage (min)","Réunion (min)","Dépassement arrêts (min)","Nb chgt fibre","Commentaire","Début Poste","Fin Poste","Temps ouverture (min)","Temps utile (min)","Temps fonctionnement (min)","Temps en arrêt (min)","Réf cadence (pcs/min)","Perte cadence (min)","Temps dégradé (min)","Pièces théoriques"]

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

def _ensure_postes_sheet(wb):
    if "Postes" not in wb.sheetnames:
        ws = wb.create_sheet("Postes")
        for i, h in enumerate(POSTES_HEADERS, start=1): ws.cell(1, i).value = h
        _format_row(ws, 1)
        from openpyxl.styles import PatternFill, Font
        fill = PatternFill("solid", fgColor="1a1f5e")
        for cell in ws[1]:
            cell.fill = fill
            cell.font = Font(color="FFFFFF", bold=True, size=10)
    else:
        ws = wb["Postes"]
        for i, h in enumerate(POSTES_HEADERS, start=1):
            if ws.cell(1, i).value is None: ws.cell(1, i).value = h
    return wb["Postes"]

def write_poste_login_row(pilot, poste, debut_dt, fin_dt):
    """Écrit en arrière-plan. Stocke postes_row_num dans _S pour que update_poste_horaires le retrouve directement."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _ensure_postes_sheet(wb)
                new_row = ws.max_row + 1
                ws.cell(new_row, 2).value = pilot
                ws.cell(new_row, 4).value = poste
                ws.cell(new_row, 16).value = debut_dt.isoformat() if debut_dt else None
                ws.cell(new_row, 17).value = fin_dt.isoformat() if fin_dt else None
                _format_row(ws, new_row)
                _safe_excel_save(wb, path)
                _S["postes_row_num"] = new_row
                save_session()
        except: pass
    threading.Thread(target=_bg, daemon=True).start()

def find_postes_row_num(pilot, debut_dt):
    """Cherche dans POSTES la ligne correspondant à ce pilote + date debut. Retourne row_num ou None."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path) or not debut_dt: return None
    try:
        with _excel_lock:
            wb = _get_wb(path)
            if wb is None: return None
            if "Postes" not in wb.sheetnames: return None
            ws = wb["Postes"]
            target_date = debut_dt.date()
            for row in ws.iter_rows(min_row=2, values_only=False):
                try:
                    b = row[1].value if len(row) > 1 else None  # col B pilot
                    p = row[15].value if len(row) > 15 else None  # col P debut
                    if str(b or "").strip().lower() != pilot.lower(): continue
                    if p is None: continue
                    p_dt = datetime.datetime.fromisoformat(str(p)) if isinstance(p, str) else p
                    if hasattr(p_dt, 'date') and p_dt.date() == target_date:
                        return row[0].row
                except: continue
    except: pass
    return None

def update_poste_horaires(row_num, debut_dt, fin_dt):
    """Écrit P et Q directement (pas de thread interne — à appeler depuis un thread background)."""
    if not row_num or row_num <= 1: return
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return
    try:
        with _excel_lock:
            wb = _get_wb(path)
            if wb is None: return
            if "Postes" not in wb.sheetnames: return
            ws = wb["Postes"]
            ws.cell(row_num, 16).value = debut_dt.isoformat() if debut_dt else None
            ws.cell(row_num, 17).value = fin_dt.isoformat() if fin_dt else None
            _safe_excel_save(wb, path)
    except: pass

def write_poste_row(data, row_num=None):
    """Écrit ou met à jour une ligne dans l'onglet Postes à la fin de chaque poste."""
    path = cfg.get("db_path","")
    if not path: return
    def _bg():
        try:
            with _excel_lock:
                wb = _get_wb(path)
                if wb is None: return
                ws = _ensure_postes_sheet(wb)
                vals = [
                    data.get("date",""),                                          # col 1  (A) Date
                    data.get("pilot",""),                                         # col 2  (B) Pilote
                    data.get("copilote",""),                                      # col 3  (C) Co-Pilote
                    data.get("poste",""),                                         # col 4  (D) Poste
                    data.get("nb_of",0),                                          # col 5  (E) Nb OF
                    round(float(data.get("prod_total",0) or 0),0),               # col 6  (F) Prod Total
                    round(float(data.get("tot_equiv",0) or 0),1),                # col 7  (G) Prod Equiv
                    data.get("trs_shift",""),                                     # col 8  (H) TRS Poste %
                    round(float(data.get("cadence_h",0) or 0),0),               # col 9  (I) Cadence/h
                    round(float(data.get("pause_min",0) or 0),1),                # col 10 (J) Total Pauses
                    round(float(data.get("nett_min",0) or 0),1),                 # col 11 (K) Nettoyage
                    round(float(data.get("reunion_min",0) or 0),1),              # col 12 (L) Temps en réunion (min)
                    round(float(data.get("depassement_min",0) or 0),1),          # col 13 (M) Temps hors budget (min)
                    int(data.get("nb_fibre_chg",0) or 0),                        # col 14 (N) Nb changements fibre
                    data.get("comment",""),                                       # col 15 (O) Commentaire
                    None,                                                          # col 16 (P) Début Poste — géré par write_poste_login_row/update_poste_horaires
                    None,                                                          # col 17 (Q) Fin Poste   — géré par write_poste_login_row/update_poste_horaires
                    round(float(data.get("temps_ouverture_min",0) or 0),1),      # col 18 (R) Temps ouverture
                    round(float(data.get("temps_utile_min",0) or 0),1),          # col 19 (S) Temps utile
                    round(float(data.get("temps_fonctionnement_min",0) or 0),1), # col 20 (T) Temps fonctionnement
                    round(float(data.get("temps_arret_min",0) or 0),1),          # col 21 (U) Temps en arrêt
                    round(float(data.get("cadence_ref_pcs_min",0) or 0),4),      # col 22 (V) Réf cadence
                    round(float(data.get("perte_cadence_min",0) or 0),1),        # col 23 (W) Perte cadence
                    round(float(data.get("degrade_min",0) or 0),1),              # col 24 (X) Temps en mode dégradé
                    round(float(data.get("pcs_theorique",0) or 0),1),             # col 25 (Y) Pièces théoriques
                ]
                if row_num and row_num > 1:
                    for ci, v in enumerate(vals, start=1):
                        if ci in (16, 17): continue  # Début/Fin Poste: ne pas écraser les timestamps
                        ws.cell(row_num, ci).value = v
                    _format_row(ws, row_num)
                else:
                    ws.append(vals)
                    _format_row(ws, ws.max_row)
                _safe_excel_save(wb, path)
        except: pass
    threading.Thread(target=_bg, daemon=True).start()

def get_current_shift_duration_s():
    sd = _S.get("shift_debut_dt")
    sf = _S.get("shift_fin_dt")
    if sd and sf:
        return (sf - sd).total_seconds()
    return get_shift_duration_s(_S.get("poste",""))

def load_postes_shift_map():
    """Lit l'onglet Postes et retourne un dict (pilot_lower, date_dmy) -> dict de toutes les valeurs Excel."""
    path = cfg.get("db_path","")
    if not path or not os.path.exists(path): return {}
    result = {}
    def _flt(v):
        try: return float(str(v).replace('%','').replace(',','.').strip()) if v not in (None,'') else None
        except: return None
    def _parse_dt(v):
        if v is None: return None
        if isinstance(v, datetime.datetime): return v
        if isinstance(v, datetime.date): return datetime.datetime.combine(v, datetime.time())
        s = str(v).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M','%d/%m/%Y %H:%M'):
            try: return datetime.datetime.strptime(s, fmt)
            except: pass
        try: return datetime.datetime.fromisoformat(s)
        except: return None
    try:
        with _excel_lock:
            wb = _get_wb(path)
            if wb is None: return {}
            if "Postes" not in wb.sheetnames:
                wb.close(); return {}
            ws = wb["Postes"]
            for ri in range(2, ws.max_row + 1):
                pilot_v = ws.cell(ri, 2).value   # col B: Pilote
                deb_v   = ws.cell(ri, 16).value  # col P: Debut Poste (datetime)
                if not pilot_v or not deb_v: continue
                deb_dt = _parse_dt(deb_v)
                if deb_dt is None: continue
                fin_v  = ws.cell(ri, 17).value   # col Q: Fin Poste (datetime)
                fin_dt = _parse_dt(fin_v)
                date_str = deb_dt.strftime("%d/%m/%Y")
                pk = (str(pilot_v).strip().lower(), date_str)
                result[pk] = {
                    'deb_dt':        deb_dt,
                    'fin_dt':        fin_dt,
                    'date_str':      date_str,
                    'pilot':         str(pilot_v).strip(),
                    'poste':         str(ws.cell(ri, 4).value or '').strip(),  # col D
                    'trs':           _flt(ws.cell(ri, 8).value),   # col H: TRS Poste %
                    'cadence_h':     _flt(ws.cell(ri, 9).value),   # col I: Cadence/h
                    'ouverture_min': _flt(ws.cell(ri, 18).value),  # col R: Temps ouverture
                    'utile_min':     _flt(ws.cell(ri, 19).value),  # col S: Temps utile
                    'fonct_min':     _flt(ws.cell(ri, 20).value),  # col T: Temps fonctionnement
                    'arret_min':     _flt(ws.cell(ri, 21).value),  # col U: Temps arret
                    'perte_min':     _flt(ws.cell(ri, 23).value),  # col W: Perte cadence (min)
                    'degrade_min':   _flt(ws.cell(ri, 24).value),  # col X: Temps degrade
                    'pcs_theorique': _flt(ws.cell(ri, 25).value),  # col Y: Pièces théoriques
                    'depassement_min': _flt(ws.cell(ri, 13).value), # col M: Dépassement arrêts
                    'row_idx':       ri,
                }
            wb.close()
    except: pass
    return result

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
        "pause_periods": [[_dt_str(a), _dt_str(b)] for a, b in _S.get("pause_periods", [])],
        "shift_debut_iso": _dt_str(_S.get("shift_debut_dt")),
        "shift_fin_iso": _dt_str(_S.get("shift_fin_dt")),
        "budget_state": _compute_budget_state_now(),
        "degrade_active": _S.get("degrade_active", False),
        "degrade_type": _S.get("degrade_type", ""),
        "degrade_start_iso": _dt_str(_S.get("degrade_start_dt")),
        "degrade_periods_iso": [{"start": _dt_str(p["start"]), "end": _dt_str(p["end"]), "type": p["type"]} for p in _S.get("degrade_periods", [])],
        "degrade_motifs": cfg.get("degrade_motifs", []),
        "budget_overrides": _S.get("budget_overrides", {}),
        "pers_pct_map": {str(k): round(v*100,1) for k,v in _pers_pct_map.items()},
        "reunion_active": t_running("reunion"),
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
    try:
        _pp = cfg.get("pilot_passwords", {})
        _pp_keys = list(_pp.keys()) if isinstance(_pp, dict) else []
    except Exception:
        _pp_keys = []
    return jsonify({
        "pilotes": get_list("Pilotes") or get_list("pilotes") or get_list("Pilote") or get_list("pilote") or _pp_keys,
        "copilotes": get_list("copilotes") or get_list("Co-Pilote") or get_list("Copilote") or get_list("Pilotes") or get_list("pilotes") or _pp_keys,
        "tailles": get_list("tailles_col") or get_list("Taille produit") or get_list("Taille") or get_list("Tailles") or get_list("taille"),
        "types_prod": get_list("types_prod_col") or get_list("Type de produit") or get_list("Type produit") or get_list("Type Produit") or get_list("types_prod"),
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
    _S["of_count_shift"] = 0
    _S["last_of_end"] = None
    _S["last_of_pilot"] = ""
    _S["interposte_s"] = 0.0
    _S["postes_row_num"] = None
    _S["tot_prod_s"] = 0.0
    if not _S.get("shift_start"):
        _S["shift_start"] = datetime.datetime.now()
    now = datetime.datetime.now()
    debut_str, fin_str = _get_model_day_cfg(poste)
    shift_debut_dt = None
    shift_fin_dt = None
    if debut_str and fin_str:
        try:
            h, m = map(int, debut_str.split(':'))
            sd = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if sd > now: sd -= datetime.timedelta(days=1)
            fh, fm = map(int, fin_str.split(':'))
            sf = sd.replace(hour=fh, minute=fm, second=0, microsecond=0)
            if sf <= sd: sf += datetime.timedelta(days=1)
            shift_debut_dt = sd
            shift_fin_dt = sf
        except: pass
    _S["shift_debut_dt"] = shift_debut_dt
    _S["shift_fin_dt"] = shift_fin_dt
    write_poste_login_row(pilot, poste, shift_debut_dt, shift_fin_dt)  # async, non bloquant
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/logout', methods=['POST'])
def api_logout():
    if _S["prod_active"]:
        return jsonify({"ok":False,"error":"Production en cours"}),400
    _S["pilot"] = None
    _S["poste"] = None
    _S["shift_start"] = None
    _S["postes_row_num"] = None
    _S["budget_overrides"] = {}
    save_session()
    # Reload cfg from disk so temporary session horaire overrides are cleared
    # (login screen will show original Paramètres values again)
    global cfg
    cfg.clear()
    cfg.update(load_cfg())
    return jsonify({"ok":True})

@flask_app.route('/api/update_shift_horaires', methods=['POST'])
def api_update_shift_horaires():
    data = request.json or {}
    debut_iso = data.get("debut_iso","")
    fin_iso = data.get("fin_iso","")
    try:
        debut_dt = datetime.datetime.fromisoformat(debut_iso)
        fin_dt = datetime.datetime.fromisoformat(fin_iso)
        # Convertir en local tz-naive si le JS envoie de l'UTC (".toISOString()")
        # pour éviter TypeError lors des soustractions avec datetime.now() (tz-naive)
        if debut_dt.tzinfo is not None:
            debut_dt = datetime.datetime.fromtimestamp(debut_dt.timestamp())
        if fin_dt.tzinfo is not None:
            fin_dt = datetime.datetime.fromtimestamp(fin_dt.timestamp())
    except:
        return jsonify({"ok":False,"error":"Format invalide"}),400
    if fin_dt <= debut_dt:
        return jsonify({"ok":False,"error":"Fin doit être après début"}),400
    old_debut = _S.get("shift_debut_dt") or debut_dt  # capturer AVANT écrasement
    _S["shift_debut_dt"] = debut_dt
    _S["shift_fin_dt"] = fin_dt
    pilot_snap = _S.get("pilot","")
    row_num_snap = _S.get("postes_row_num")
    save_session()
    # Mise à jour Excel en arrière-plan pour ne pas bloquer
    def _bg():
        rn = row_num_snap or find_postes_row_num(pilot_snap, old_debut)
        if rn and not row_num_snap:
            _S["postes_row_num"] = rn
        update_poste_horaires(rn, debut_dt, fin_dt)
    threading.Thread(target=_bg, daemon=True).start()
    dur_s = (fin_dt - debut_dt).total_seconds()
    return jsonify({"ok":True,"shift_dur_s":round(dur_s,0)})

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

def _get_uncovered_gaps(from_dt, to_dt, pilot):
    """Compute actual uncovered time gaps between from_dt and to_dt,
    accounting for ALL declarations in _decl_cache (OFs + stops) for this pilot."""
    if not from_dt or not to_dt: return []
    from_s = _hms_to_sec(from_dt) if not hasattr(from_dt,'hour') else from_dt.hour*3600+from_dt.minute*60
    to_s = _hms_to_sec(to_dt) if not hasattr(to_dt,'hour') else to_dt.hour*3600+to_dt.minute*60
    if to_s <= from_s + 59: return []
    date_strs = set()
    if hasattr(from_dt,'strftime'): date_strs.add(from_dt.strftime("%d/%m/%Y"))
    date_strs.add(datetime.date.today().strftime("%d/%m/%Y"))
    all_slots = []
    for rn, r in _decl_cache:
        rd = _row_date(r[2])
        if rd not in date_strs: continue
        if pilot and str(r[4] or "") != pilot: continue
        ds = _hms_to_sec(str(r[16] or "00:00:00"))
        fs = _hms_to_sec(str(r[17] or "00:00:00"))
        if fs > ds and ds >= 0:
            all_slots.append([ds, fs])
    all_slots.sort()
    merged = []
    for s, e in all_slots:
        if merged and s <= merged[-1][1] + 30:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    gaps = []
    covered = from_s
    for s, e in merged:
        s2 = max(s, from_s); e2 = min(e, to_s)
        if e2 <= s2: continue
        if s2 >= covered + 60:
            gaps.append({"debut": _sec_to_hm(covered), "fin": _sec_to_hm(s2),
                         "duree_s": round(s2 - covered)})
        covered = max(covered, e2)
    if to_s >= covered + 60:
        gaps.append({"debut": _sec_to_hm(covered), "fin": _sec_to_hm(to_s),
                     "duree_s": round(to_s - covered)})
    return gaps

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
    model_debut_dt = _S.get("shift_debut_dt")
    model_debut_str = model_debut_dt.strftime("%H:%M") if model_debut_dt else ""
    ip_debut_hms = ""
    ip_fin_hms = now.strftime("%H:%M")
    ip_debut_iso = ""
    pre_shift_gap_s = 0.0
    shift_model_start_str = ""
    shift_model_start_iso = ""
    pilot = _S.get("pilot","")
    gaps = []  # actual uncovered intervals
    if is_first_of and model_debut_dt:
        gap_s = max(0.0, (now - model_debut_dt).total_seconds())
        _S["interposte_s"] = gap_s
        ip_debut_hms = model_debut_str
        ip_debut_iso = model_debut_dt.isoformat()
        if gap_s >= 120:
            shift_model_start_str = model_debut_str
            shift_model_start_iso = model_debut_dt.isoformat()
            gaps = _get_uncovered_gaps(model_debut_dt, now, pilot)
            pre_shift_gap_s = sum(g["duree_s"] for g in gaps) if gaps else 0.0
    elif _S["last_of_end"]:
        gap_s = (now - _S["last_of_end"]).total_seconds()
        ip_debut_hms = _S["last_of_end"].strftime("%H:%M")
        ip_debut_iso = _S["last_of_end"].isoformat()
        if gap_s >= 120:
            gaps = _get_uncovered_gaps(_S["last_of_end"], now, pilot)
    return jsonify({"ok":True,"gap_s":round(gap_s,0),
                    "pre_shift_gap_s":round(pre_shift_gap_s,0),
                    "gaps": gaps,
                    "shift_model_start":shift_model_start_str,
                    "shift_model_start_iso":shift_model_start_iso,
                    "ip_debut_hms":ip_debut_hms,
                    "ip_fin_hms":ip_fin_hms,
                    "ip_debut_iso":ip_debut_iso})

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
    debut_hms = data.get("debut_hms","")
    fin_hms = data.get("fin_hms","")
    start_dt = _S["last_of_end"]
    end_dt = _S["of_start"]
    today = datetime.date.today()
    if debut_hms:
        try:
            h, mi = map(int, debut_hms.split(":"))
            start_dt = datetime.datetime.combine(today, datetime.time(h, mi))
        except: pass
    if fin_hms:
        try:
            h, mi = map(int, fin_hms.split(":"))
            end_dt = datetime.datetime.combine(today, datetime.time(h, mi))
            if end_dt < start_dt: end_dt += datetime.timedelta(days=1)
        except: pass
    # 1er OF du poste : last_of_end est None → utiliser shift_debut_dt comme début du gap
    if start_dt is None:
        start_dt = _S.get("shift_debut_dt")
    if _S["inter_of_s"] >= 60 and start_dt and end_dt:
        write_changement_of(start_dt, end_dt, label=label or None, comment=comment)
        # Mise à jour synchrone du cache pour éviter le race-condition fin-de-poste
        # (write_changement_of écrit en background → _decl_cache pas encore rafraîchi)
        global _decl_cache
        shift_dt2 = _S.get("shift_start") or start_dt
        _sd2 = shift_dt2.strftime("%d/%m/%Y")
        _dur2_s = (end_dt - start_dt).total_seconds()
        _cache_row = [
            label or "Changement d'OF", "", start_dt.strftime("%d/%m/%Y"),
            _S.get("poste",""), _S.get("pilot",""),
            "","","","","","","","","","","",
            start_dt.strftime("%H:%M:%S"), end_dt.strftime("%H:%M:%S"), fmt(_dur2_s),
            "","","","","","","","","","","","","","","","",comment,
            "","",_sd2,
        ]
        _decl_cache.append((-1, _cache_row))
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
    write_interposte_to_excel(cfg["interposte_labels"])
    return jsonify({"ok":True})

    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/set_budget_override', methods=['POST'])
def api_set_budget_override():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    ov = data.get("overrides",{})
    valid_keys = {"clean_short_min","clean_long_min","clean_grand_min","meeting_tol_min","pause_min"}
    _S["budget_overrides"] = {k: float(v) for k,v in ov.items() if k in valid_keys}
    save_session()
    return jsonify({"ok":True})

@flask_app.route('/api/start_degrade', methods=['POST'])
def api_start_degrade():
    data = request.json or {}
    motif = str(data.get("motif","")).strip()
    if not motif:
        return jsonify({"ok":False,"error":"Motif requis"}),400
    if _S.get("degrade_active"):
        return jsonify({"ok":False,"error":"Mode dégradé déjà actif"}),400
    _S["degrade_active"] = True
    _S["degrade_type"] = motif
    _S["degrade_start_dt"] = datetime.datetime.now()
    save_session()
    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/stop_degrade', methods=['POST'])
def api_stop_degrade():
    if not _S.get("degrade_active"):
        return jsonify({"ok":False,"error":"Mode dégradé non actif"}),400
    end_dt_deg = datetime.datetime.now()
    start_dt_deg = _S["degrade_start_dt"]
    motif = _S["degrade_type"]
    dur_s = max(0, (end_dt_deg - start_dt_deg).total_seconds())
    _S["degrade_periods"].append({"start": start_dt_deg, "end": end_dt_deg, "type": motif})
    _S["degrade_active"] = False; _S["degrade_type"] = ""; _S["degrade_start_dt"] = None
    # Si durée < 30s (ex: fin de poste juste après fin d'OF), on ne génère pas de ligne parasite
    if dur_s < 30:
        save_session()
        threading.Thread(target=generate_dashboard_html, daemon=True).start()
        return jsonify({"ok":True})
    pilot = _S.get("pilot",""); poste = _S.get("poste","")
    shift_dt = _S.get("shift_start") or start_dt_deg
    _row = [
        motif, "Mode dégradé",
        start_dt_deg.strftime("%d/%m/%Y"), poste, pilot,
        "","","","","","","","","","","",
        start_dt_deg.strftime("%H:%M:%S"), end_dt_deg.strftime("%H:%M:%S"), fmt(dur_s),
        "","","","","","","","","","","","","","","","","","","","",
        shift_dt.strftime("%d/%m/%Y"),
    ]
    write_excel_bg([], [_row])
    try:
        _nrn = max((rn for rn,_ in _decl_cache), default=1)+1
        _decl_cache.append((_nrn, tuple(_row)+("",)*max(0,40-len(_row))))
    except: pass
    save_session()
    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/save_degrade_list', methods=['POST'])
def api_save_degrade_list():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    motifs = [str(m).strip() for m in data.get("motifs",[]) if str(m).strip()]
    cfg["degrade_motifs"] = motifs
    save_cfg_data()
    write_degrade_list_to_excel()
    return jsonify({"ok":True})

@flask_app.route('/api/save_pers_pct', methods=['POST'])
def api_save_pers_pct():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw): return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    entries = data.get("entries",[])
    _pers_pct_map.clear()
    for e in entries:
        try:
            np_k = int(e.get("nb_pers",0))
            pct_raw = float(e.get("pct",0))
            if 1 <= np_k <= 10 and 0 < pct_raw <= 200:
                _pers_pct_map[np_k] = pct_raw / 100.0 if pct_raw > 2 else pct_raw
        except: pass
    write_pers_pct_to_excel()
    return jsonify({"ok":True})

@flask_app.route('/api/end_prod', methods=['POST'])
def api_end_prod():
    if not _S["prod_active"] or not _S["of_start"]:
        return jsonify({"ok":False,"error":"Pas de production active"}),400
    data = request.json or {}
    v = data.get("form",{})
    if _S["is_paused"]:
        _toggle_pause_internal()
    # Mode dégradé : enregistrer la portion de CET OF uniquement — NE PAS fermer le mode
    # Le dégradé persiste jusqu'à ce que l'utilisateur l'arrête explicitement
    _degrade_end_row = None
    _dg_dur_for_trs = 0.0  # durée dégradé dans cet OF, pour calcul TRS ci-dessous
    if _S.get("degrade_active") and _S.get("degrade_start_dt"):
        _dg_of_start = max(_S["degrade_start_dt"], _S["of_start"])
        _dg_of_end = datetime.datetime.now()
        _dg_motif = _S["degrade_type"]
        _dg_dur_of = max(0.0, (_dg_of_end - _dg_of_start).total_seconds())
        _dg_dur_for_trs = _dg_dur_of  # sauvegarde AVANT d'avancer le pointeur
        _sh_dt = _S.get("shift_start") or _dg_of_start
        if _dg_dur_of >= 1:
            _degrade_end_row = [
                _dg_motif, "Mode dégradé",
                _dg_of_start.strftime("%d/%m/%Y"), _S.get("poste",""), _S.get("pilot",""),
                "","","","","","","","","","","",
                _dg_of_start.strftime("%H:%M:%S"), _dg_of_end.strftime("%H:%M:%S"), fmt(_dg_dur_of),
                "","","","","","","","","","","","","","","","","","","","",
                _sh_dt.strftime("%d/%m/%Y"),
            ]
            # Avancer le start pour que api_stop_degrade ne couvre que la période restante
            _S["degrade_start_dt"] = _dg_of_end
        # degrade_active / degrade_type restent inchangés — l'utilisateur arrête explicitement
    t_stop_all()
    tl_close_all()
    # Arrêter la pause si elle est encore active (non gérée par tl_events)
    if _S.get("is_paused") and _S.get("pause_start"):
        _pnow = datetime.datetime.now()
        _S["pause_total_s"] += (_pnow - _S["pause_start"]).total_seconds()
        _S["pause_periods"].append((_S["pause_start"], _pnow))
        _S["is_paused"] = False
        _S["pause_start"] = None
    end_dt = datetime.datetime.now()
    of_s_brut = (end_dt-_S["of_start"]).total_seconds()
    # Budget arrêts prévus — calculé après tl_close_all (tous les événements sont terminés)
    _of_budget = _compute_budget_state_now()
    _of_planned_ded_s = _of_budget["total_of_deductible_s"]
    pause_max_s = int(cfg.get("pause_max_min",20))*60
    of_s = max(1, of_s_brut - min(_S["pause_total_s"],pause_max_s))
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
    _objectif_pcs = ""
    if prod_ref>0 and of_s_brut>0:
        # Calculer le temps en mode dégradé pendant cet OF
        _deg_s = 0.0
        for _dp in _S.get("degrade_periods", []):
            if _dp.get("start") and _dp.get("end"):
                _d0 = max(_dp["start"], _S["of_start"])
                _d1 = min(_dp["end"], end_dt)
                if _d1 > _d0: _deg_s += (_d1 - _d0).total_seconds()
        # Dégradé actif pendant cet OF — valeur déjà calculée avant l'avance du pointeur
        _deg_s += _dg_dur_for_trs
        _eff_s = max(1.0, of_s_brut - _of_planned_ded_s)
        _adj_s = max(1.0, _eff_s)
        _pct_ep = get_pct_cadence(v.get("nb_pers", 1))
        trs = round(equiv/(prod_ref*_pct_ep*_adj_s/28800)*100,1)
        trs_str = str(trs)
        # Pièces théoriques = objectif OF (même base que TRS)
        _pcoef_ep = (equiv / qte_fab) if (qte_fab and qte_fab > 0 and equiv and equiv > 0) else 1.0
        _objectif_pcs = round(prod_ref * _pct_ep * _adj_s / 28800 / _pcoef_ep, 1)

    # Ligne Production (40 cols, format unifié)
    _shift_dt = _S.get("shift_start") or datetime.datetime.now()
    _shift_date_str = _shift_dt.strftime("%d/%m/%Y")
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
        _shift_date_str,
        "",
        round(_deg_s / 60.0, 2),
        _objectif_pcs,
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
    _S["is_paused"] = False
    _S["pause_start"] = None
    _S["form"] = {}
    _S["degrade_periods"] = []
    save_session()
    _extra_evt_rows = [_degrade_end_row] if _degrade_end_row else []
    # Mise à jour immédiate de _decl_cache avant write_excel_bg (évite race condition avec generate_dashboard_html)
    try:
        _next_rn_ep = max((rn for rn, _ in _decl_cache), default=0) + 1
        for _ep_row in ([prod_row] if prod_row else []) + evt_rows + _extra_evt_rows:
            _padded_ep = tuple(_ep_row) + ("",) * max(0, 40 - len(_ep_row))
            _decl_cache.append((_next_rn_ep, _padded_ep))
            _next_rn_ep += 1
    except: pass
    write_excel_bg(prod_row, evt_rows + _extra_evt_rows)
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
        _prev_ded = _compute_budget_state_now()["total_of_deductible_s"]
        _eff_s = max(1.0, of_s_brut - _prev_ded)
        _prev_deg_s = 0.0
        for _dp in _S.get("degrade_periods", []):
            if _dp.get("start") and _dp.get("end"):
                _d0 = max(_dp["start"], _S["of_start"])
                _d1 = min(_dp["end"], now)
                if _d1 > _d0: _prev_deg_s += (_d1 - _d0).total_seconds()
        if _S.get("degrade_active") and _S.get("degrade_start_dt"):
            _d0 = max(_S["degrade_start_dt"], _S["of_start"])
            _prev_deg_s += max(0.0, (now - _d0).total_seconds())
        _adj_s_prev = max(1.0, _eff_s)
        _pct_prv = get_pct_cadence(_S.get("form",{}).get("nb_pers",1))
        trs=round(equiv/(prod_ref*_pct_prv*_adj_s_prev/28800)*100,1)
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
    # Si pas de prod active : écrire la déclaration directement en Excel
    if not _S.get("prod_active"):
        ev = next((e for e in reversed(_S["tl_events"]) if e.get("key")==key and e.get("end")), None)
        if ev and ev.get("start") and ev.get("end"):
            _start = ev["start"]; _end = ev["end"]
            _dur = max(0,(_end-_start).total_seconds())
            _cat = ev.get("cat","pb"); _ntype = ev.get("nettoyage_type","court")
            if key=="nettoyage":
                _lbl = {"court":"Nettoyage court","long":"Nettoyage long","grand":"Grand nettoyage"}.get(_ntype,"Nettoyage court")
            elif _cat=="autre":
                _lbl = key
            else:
                _cat_n = "Rattrapage" if _cat=="ratt" else "PB Technique"
                _evlbl = next((e[0] for e in EVENTS if e[1]==key), key)
                _lbl = f"{_cat_n}: {_evlbl}"
            _sh = _S.get("shift_start") or _start
            _row = [
                _lbl, _S.get("form",{}).get("of_num",""),
                _start.strftime("%d/%m/%Y"), _S.get("poste",""), _S.get("pilot",""),
                "","","","","","","","","","","Oui" if _S.get("form",{}).get("kit") else "Non",
                _start.strftime("%H:%M:%S"), _end.strftime("%H:%M:%S"), fmt(_dur),
                "","","","","","","","","","","","","","","","",comment,"",
                _sh.strftime("%d/%m/%Y"),
            ]
            write_excel_bg([], [_row])
            try:
                _nrn = max((rn for rn,_ in _decl_cache), default=1)+1
                _decl_cache.append((_nrn, tuple(_row)+('',)*max(0,40-len(_row))))
            except: pass
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

@flask_app.route('/api/toggle_reunion', methods=['POST'])
def api_toggle_reunion():
    if t_running("reunion"):
        t_stop("reunion")
        # Fermer tous les événements réunion ouverts (gère les clés legacy)
        now = datetime.datetime.now()
        for ev in _S["tl_events"]:
            k = ev.get("key","")
            if ("reunion" in k.lower() or "meeting" in k.lower()) and not ev.get("end"):
                ev["end"] = now
                ev["comment"] = ""
        save_session()
        reunion_active = False
        # Écrire en Excel si hors production (en prod : écrit à la fin de l'OF via build_decl_rows)
        if not _S.get("prod_active"):
            ev = next((e for e in reversed(_S["tl_events"]) if e.get("key")=="reunion" and e.get("end")), None)
            if ev and ev.get("start") and ev.get("end"):
                _start = ev["start"]; _end = ev["end"]
                _dur = max(0, (_end - _start).total_seconds())
                _sh = _S.get("shift_start") or _start
                _row = [
                    "Réunion", _S.get("form",{}).get("of_num",""),
                    _start.strftime("%d/%m/%Y"), _S.get("poste",""), _S.get("pilot",""),
                    "","","","","","","","","","","Oui" if _S.get("form",{}).get("kit") else "Non",
                    _start.strftime("%H:%M:%S"), _end.strftime("%H:%M:%S"), fmt(_dur),
                    "","","","","","","","","","","","","","","","","","",
                    _sh.strftime("%d/%m/%Y"),
                ]
                write_excel_bg([], [_row])
                # Pas de _decl_cache.append ici : l'événement est encore dans tl_events
                # → évite le double comptage dans _compute_budget_state_now
    else:
        t_start("reunion")
        tl_open("reunion", "reunion")
        reunion_active = True
    threading.Thread(target=generate_dashboard_html, daemon=True).start()
    return jsonify({"ok": True, "reunion_active": reunion_active})

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
    all_rows = [(rn,r) for rn,r in _decl_cache if str(r[0] or "").strip().lower() in ("production","prod","")]
    for rn, r in all_rows[-500:]:
        try:
            row_d = _parse_date(_row_date(r[2])) if r[2] else None
            if d_from and row_d and row_d < d_from: continue
            if d_to and row_d and row_d > d_to: continue
            trs = -1
            try:
                # r[24] = TRS stocké par api_end_prod (avec correction dégradé)
                trs_col = str(r[24] or "")
                if trs_col:
                    try: trs = round(float(trs_col.replace(",",".")),1)
                    except: pass
                if trs < 0:
                    # Fallback : calcul depuis timing (anciens enregistrements sans r[24])
                    equiv_v = float(str(r[21] or 0).replace(",","."))
                    pr = get_prod_ref()
                    debut_s = _hms_to_sec(str(r[16] or "00:00:00"))
                    fin_s = _hms_to_sec(str(r[17] or "00:00:00"))
                    brut_s = fin_s - debut_s if fin_s > debut_s else _hms_to_sec(str(r[18] or "00:00:00"))
                    if pr>0 and brut_s>0 and equiv_v>0:
                        trs = round(equiv_v/(pr*brut_s/28800)*100,1)
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
                "qte_init_taie": str(r[25] if len(r)>25 else ""),
                "nb_taie2": str(r[26] if len(r)>26 else ""),
                "nb_def_cout": str(r[27] if len(r)>27 else ""),
                "mq_taie": str(r[28] if len(r)>28 else ""),
                "mq_housse": str(r[29] if len(r)>29 else ""),
                "nb_pp": str(r[30] if len(r)>30 else ""),
                "duree_mq_mp": str(r[32] if len(r)>32 else ""),
                "manquant_pers": str(r[33] if len(r)>33 else ""),
                "comment": str(r[35] or ""),
            })
        except: pass
    return jsonify(list(reversed(rows)))

@flask_app.route('/api/events_list')
def api_events_list():
    rows = []
    date_from_el = request.args.get("from","")
    date_to_el   = request.args.get("to","")
    def _parse_date_el(s):
        try:
            if "-" in s: return datetime.datetime.strptime(s,"%Y-%m-%d").date()
            if "/" in s: return datetime.datetime.strptime(s,"%d/%m/%Y").date()
        except: pass
        return None
    d_from_el = _parse_date_el(date_from_el) if date_from_el else None
    d_to_el   = _parse_date_el(date_to_el)   if date_to_el   else None
    evt_rows = [(rn,r) for rn,r in _decl_cache if str(r[0] or "").strip().lower() not in ("production","prod","")]
    # With date range: scan all; without: limit to last 500
    scan_rows = evt_rows if (d_from_el or d_to_el) else evt_rows[-500:]
    for rn, r in scan_rows:
        try:
            if d_from_el or d_to_el:
                row_d_el = _parse_date_el(_row_date(r[2])) if r[2] else None
                if d_from_el and row_d_el and row_d_el < d_from_el: continue
                if d_to_el   and row_d_el and row_d_el > d_to_el:   continue
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
                "is_degrade": _is_degrade_type(type_str),
            })
        except: pass
    return jsonify(list(reversed(rows)))

def _apply_model_overrides(models):
    """Retourne une copie des modèles horaires avec les surcharges de session appliquées."""
    overrides = _S.get("model_overrides", {})
    if not overrides:
        return models
    import copy
    result = copy.deepcopy(models)
    day_map = {0:'lun',1:'mar',2:'mer',3:'jeu',4:'ven',5:'sam',6:'dim'}
    dk = day_map.get(datetime.date.today().weekday(), 'lun')
    for m in result:
        nom = m.get("nom","")
        if nom in overrides and dk in overrides[nom]:
            ov = overrides[nom][dk]
            if "jours" not in m:
                m["jours"] = {}
            if dk not in m["jours"]:
                m["jours"][dk] = {}
            m["jours"][dk]["debut"] = ov.get("debut","")
            m["jours"][dk]["fin"] = ov.get("fin","")
    return result

@flask_app.route('/api/config')
def api_config():
    arrets_prevus = {
        "clean_short_min": cfg.get("clean_short_min", 0),
        "clean_long_min":  cfg.get("clean_long_min",  0),
        "clean_grand_min": cfg.get("clean_grand_min", 0),
        "meeting_tol_min": cfg.get("meeting_tol_min", 0),
        "pause_min":       cfg.get("pause_min",       0),
    }
    return jsonify({
        "prod_ref": cfg.get("prod_ref",0),
        "pause_max_min": cfg.get("pause_max_min",20),
        "clean_short_min": cfg.get("clean_short_min",0),
        "clean_long_min": cfg.get("clean_long_min",0),
        "clean_grand_min": cfg.get("clean_grand_min",0),
        "meeting_tol_min": cfg.get("meeting_tol_min",0),
        "pause_min": cfg.get("pause_min",0),
        "arrets_prevus": arrets_prevus,
        "db_path": cfg.get("db_path",""),
        "db_name": os.path.basename(cfg.get("db_path","")) if cfg.get("db_path") else "",
        "modeles_horaires": _apply_model_overrides(cfg.get("modeles_horaires",[])),
        "modeles_horaires_base": cfg.get("modeles_horaires",[]),
        "pilot_passwords": cfg.get("pilot_passwords",{}),
        "degrade_motifs": cfg.get("degrade_motifs",[]),
        "pers_pct_map": {str(k): round(v*100,1) for k,v in _pers_pct_map.items()},
    })

@flask_app.route('/api/settings', methods=['POST'])
def api_settings():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    for k in ["prod_ref","pause_max_min","clean_short_min","clean_long_min","clean_grand_min","meeting_tol_min","pause_min"]:
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
    # Persister arrêts prévus dans Excel si concerné
    if any(k in data for k in _ARRETS_PREVUS_KEYS):
        write_arrets_prevus_to_excel()
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
    # Stocker dans _S uniquement — cfg n'est jamais modifié, Paramètres reste intact
    if "model_overrides" not in _S:
        _S["model_overrides"] = {}
    if nom not in _S["model_overrides"]:
        _S["model_overrides"][nom] = {}
    _S["model_overrides"][nom][day_key] = {"debut": debut, "fin": fin}
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
    # Suppression synchrone du cache pour éviter stale data
    global _decl_cache
    _decl_cache = [(rn, row) for rn, row in _decl_cache if rn != row_num]
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
    of_list=[]; _filtered_prod_raw_fp=[]
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
            _filtered_prod_raw_fp.append(r)
            try: _r24fp=float(str(r[24] if len(r)>24 else '').strip() or '-1')
            except: _r24fp=-1.0
            trs = _r24fp if _r24fp>=0 else (round(eq/(prod_ref*s/28800)*100,1) if prod_ref>0 and s>0 and eq>0 else -1)
            of_list.append({
                "of":str(r[1] or ""),"taille":str(r[7] or ""),
                "type_prod":str(r[9] or ""),"qte_fab":str(r[19] or ""),
                "qte_emb":str(r[20] or ""),"equiv":str(r[21] or ""),
                "debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],
                "duree":str(r[18] or ""),"trs":trs,"fibre":str(r[11] or ""),
                "nb_pers":str(r[6] or ""),
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
    # Gather stop events for this shift (for planned deduction + écart)
    shift_evt_rows = []
    declared_stop_s = 0.0
    for rn, r in _decl_cache:
        rd = _row_date(r[2])
        if rd != shift_date_str and rd != today: continue
        if str(r[4] or "") != pilot: continue
        row_type = str(r[0] or "").strip().lower()
        if row_type not in ("production","prod",""):
            try:
                dur_s = _hms_to_sec(str(r[18] or "00:00:00"))
                declared_stop_s += dur_s
                shift_evt_rows.append((rn, r))
            except: pass
    _ses_ov = _S.get("budget_overrides") or {}
    planned_ded = _compute_planned_deduction_s(shift_evt_rows, _ses_ov)
    model_dur_s = get_current_shift_duration_s()
    ecart_s = max(0.0, model_dur_s - (tot_s + declared_stop_s))
    trs_poste_shift = -1.0
    if model_dur_s > 0 and prod_ref > 0 and tot_eq > 0:
        elapsed_s = max(1.0, model_dur_s - planned_ded)
        trs_poste_shift = round(tot_eq/(prod_ref*elapsed_s/28800)*100,1)
    # Compute gap intervals (plages non justifiées)
    # Use session shift times (set at login from Excel) — more reliable than day-cfg lookup
    _sd = _S.get("shift_debut_dt")
    _sf = _S.get("shift_fin_dt")
    if _sd and _sf:
        debut_str2 = _sd.strftime("%H:%M")
        fin_str2 = _sf.strftime("%H:%M")
    else:
        debut_str2, fin_str2 = _get_model_day_cfg(pilot_poste, datetime.date.today())
    gap_intervals = []
    model_debut_hm = debut_str2 or ""
    model_fin_hm = fin_str2 or ""
    if debut_str2 and fin_str2:
        md_s = _hms_to_sec(debut_str2)
        mf_s = _hms_to_sec(fin_str2)
        if mf_s <= md_s: mf_s += 86400
        all_slots = []
        for rn2, r2 in _decl_cache:
            rd2 = _row_date(r2[2])
            if rd2 != shift_date_str and rd2 != today: continue
            if str(r2[4] or "") != pilot: continue
            ds2 = _hms_to_sec(str(r2[16] or "00:00:00"))
            fs2 = _norm_fin(ds2, _hms_to_sec(str(r2[17] or "00:00:00")))
            if fs2 > ds2 and ds2 >= 0: all_slots.append([ds2, fs2])
        all_slots.sort()
        merged = []
        for s2, e2 in all_slots:
            if merged and s2 <= merged[-1][1] + 60:
                merged[-1][1] = max(merged[-1][1], e2)
            else:
                merged.append([s2, e2])
        overflow_s = 0.0
        for s2_raw, e2_raw in merged:
            if e2_raw > mf_s:
                overflow_s = max(overflow_s, e2_raw - mf_s)
        covered = md_s
        for s2, e2 in merged:
            s2 = max(s2, md_s); e2 = min(e2, mf_s)
            if s2 >= covered + 120:
                gap_intervals.append({"debut": _sec_to_hm(covered), "fin": _sec_to_hm(s2), "duree_min": round((s2-covered)/60)})
            covered = max(covered, e2)
        if mf_s >= covered + 120:
            gap_intervals.append({"debut": _sec_to_hm(covered), "fin": _sec_to_hm(mf_s), "duree_min": round((mf_s-covered)/60)})
    else:
        overflow_s = 0.0
    # ── Nouvelles métriques pour l'onglet Postes Excel ──
    # Intervalles d'arrêts fusionnés (sans chevauchement, hors dégradé)
    _degrade_s_fp = _merged_degrade_s([r_s for _, r_s in shift_evt_rows])
    _stop_raw = [(_hms_to_sec(str(r_s[16] or "00:00:00")), _hms_to_sec(str(r_s[17] or "00:00:00")))
                 for _, r_s in shift_evt_rows if not _is_degrade_type(str(r_s[0] or ""))]
    _stop_ivs = sorted((s2, _norm_fin(s2, f2)) for s2, f2 in _stop_raw if _norm_fin(s2, f2) > s2)
    _merged_s = []
    for _ds, _fs in _stop_ivs:
        if _merged_s and _ds <= _merged_s[-1][1]:
            _merged_s[-1] = (_merged_s[-1][0], max(_merged_s[-1][1], _fs))
        else:
            _merged_s.append((_ds, _fs))
    net_stop_min_fp = round(sum(f - s for s, f in _merged_s) / 60, 1)
    ouverture_min_fp = round(model_dur_s / 60, 1)
    temps_fonctionnement_fp = round(max(0.0, ouverture_min_fp - net_stop_min_fp), 1)
    # Budget arrêts prévus — utilise l'implémentation centralisée pour arrets_prevu
    arrets_prevu_fp = _compute_planned_deduction_s(shift_evt_rows, _ses_ov) / 60  # minutes
    temps_utile_fp = round(max(0.0, ouverture_min_fp - arrets_prevu_fp), 1)
    # Détail par type (nécessaire pour réunion et dépassement uniquement)
    _blab = {"pause_min":"Pause","meeting_tol_min":"Réunion","clean_short_min":"Nettoyage court","clean_long_min":"Nettoyage long","clean_grand_min":"Nettoyage très long"}
    # budget_min utilise l'override de session si disponible, sinon cfg
    _bdata = {bk:{"budget_min":float((_ses_ov.get(bk) if _ses_ov.get(bk) is not None else cfg.get(bk,0)) or 0),"used_min":0.0} for bk in _blab}
    for _, r_e in shift_evt_rows:
        _bk2 = _get_arret_budget_key(str(r_e[0] or ''))
        if _bk2 and _bk2 in _bdata:
            _dp2 = str(r_e[18] or ''); _pp2 = (_dp2+':00:00').split(':')
            try: _bs2 = int(_pp2[0] or 0)*3600+int(_pp2[1] or 0)*60+int(_pp2[2] or 0)
            except: _bs2 = 0
            _bdata[_bk2]['used_min'] += _bs2/60
    reunion_min_fp = round(_bdata.get("meeting_tol_min", {}).get("used_min", 0.0), 1)
    depassement_min_fp = round(sum(max(0.0, v["used_min"] - v["budget_min"]) for v in _bdata.values()), 1)
    cadence_ref_fp = round(prod_ref / 480, 4) if prod_ref > 0 else 0.0
    _elapsed_fp = max(1.0, model_dur_s - arrets_prevu_fp * 60)
    _adj_fp = max(1.0, _elapsed_fp)
    _plan_bdata_fp = {bk: float((_ses_ov.get(bk) if _ses_ov.get(bk) is not None else cfg.get(bk, 0)) or 0) * 60 for bk in _blab}
    _plan_used_fp = {bk: 0.0 for bk in _blab}
    _plan_ivs_fp = []
    for _, r_e2 in shift_evt_rows:
        _bk_p = _get_arret_budget_key(str(r_e2[0] or ''))
        if _bk_p and _bk_p in _plan_bdata_fp:
            _ds_p = _hms_to_sec(str(r_e2[16] or '00:00:00'))
            _fs_p = _norm_fin(_ds_p, _hms_to_sec(str(r_e2[17] or '00:00:00')))
            _dur_p = _fs_p - _ds_p
            if _dur_p > 0 and _plan_used_fp[_bk_p] < _plan_bdata_fp[_bk_p]:
                _cap_p = min(_dur_p, _plan_bdata_fp[_bk_p] - _plan_used_fp[_bk_p])
                _plan_ivs_fp.append((_ds_p, _ds_p + _cap_p))
            _plan_used_fp[_bk_p] += _dur_p
    _sum_exp_fp = None
    if _pers_pct_map and _filtered_prod_raw_fp:
        _deg_ivs_fp = _merged_degrade_ivs([r_s for _, r_s in shift_evt_rows])
        trs_poste_shift, _sum_exp_fp = _option_b_trs(_filtered_prod_raw_fp, _deg_ivs_fp, prod_ref, _plan_ivs_fp)
        perte_cadence_fp = round((_sum_exp_fp - tot_eq) / cadence_ref_fp, 1) if cadence_ref_fp > 0 and _sum_exp_fp > 0 else 0.0
    else:
        if prod_ref > 0 and _adj_fp > 0 and tot_eq > 0:
            trs_poste_shift = round(tot_eq / (prod_ref * _adj_fp / 28800) * 100, 1)
        perte_cadence_fp = round((prod_ref * _adj_fp / 28800 - tot_eq) / cadence_ref_fp, 1) if cadence_ref_fp > 0 else 0.0
    tot_pcs_fp = sum(float(str(r[19] or 0).replace(",",".") or 0) for r in _filtered_prod_raw_fp)
    cadence_h_fp = round(tot_eq * 60 / temps_utile_fp) if temps_utile_fp > 0 else 0
    _sorted_of_fib = sorted([o for o in of_list if o.get("fibre")], key=lambda x: x.get("debut",""))
    nb_fibre_chg_fp = sum(1 for i in range(1, len(_sorted_of_fib)) if _sorted_of_fib[i]["fibre"] != _sorted_of_fib[i-1]["fibre"])
    return jsonify({
        "pilot":pilot,"date":today,
        "nb_of":len(of_list),"trs":trs_poste,"trs_shift":trs_poste_shift,
        "tot_equiv":round(tot_eq,1),"tot_s":round(tot_s,0),
        "declared_stop_s": round(declared_stop_s, 0),
        "of_list":of_list,"of_count_shift":_S["of_count_shift"],
        "shift_start_iso": _dt_str(_S.get("shift_start")),
        "ecart_s": round(ecart_s, 0),
        "model_dur_s": round(model_dur_s, 0),
        "planned_ded_s": round(planned_ded, 0),
        "gap_intervals": gap_intervals,
        "model_debut": model_debut_hm,
        "model_fin": model_fin_hm,
        "overflow_min": round(overflow_s / 60),
        "overflow_s": round(overflow_s, 0),
        "ouverture_min": ouverture_min_fp,
        "temps_utile_min": temps_utile_fp,
        "temps_fonctionnement_min": temps_fonctionnement_fp,
        "net_stop_min": net_stop_min_fp,
        "cadence_ref_pcs_min": cadence_ref_fp,
        "perte_cadence_min": perte_cadence_fp,
        "degrade_min": round(_degrade_s_fp / 60, 1),
        "cadence_h": cadence_h_fp,
        "pcs_theorique": round(_sum_exp_fp, 1) if _sum_exp_fp is not None else round(prod_ref * max(0.0, model_dur_s - planned_ded) / 28800, 1),
        "reunion_min": reunion_min_fp,
        "depassement_min": depassement_min_fp,
        "nb_fibre_chg": nb_fibre_chg_fp,
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
    rows = []; _htd_prod_raw = []
    tot_eq=0.0; tot_s=0.0; _htd_ded_s=0.0; _htd_deg_ivs=[]
    for rn,r in _decl_cache:
        rd = _row_date(r[2])
        if rd != shift_date_str and rd != today: continue
        if str(r[4] or "") != pilot: continue
        _rtype_htd = str(r[0] or "").strip().lower()
        if _rtype_htd in ("production","prod",""):
            eq=float(str(r[21] or 0).replace(",",".") or 0)
            s=_hms_to_sec(str(r[18] or "00:00:00"))
            tot_eq+=eq; tot_s+=s
            _htd_prod_raw.append(r)
            trs_of=-1
            try:
                trs_col=str(r[24] or "")
                if trs_col: trs_of=round(float(trs_col.replace(",",".")),1)
                elif prod_ref>0 and s>0: trs_of=round(eq/(prod_ref*s/28800)*100,1)
            except: pass
            rows.append({"of":str(r[1] or ""),"debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],"trs":trs_of,"equiv":eq,"qte_fab":str(r[19] or "")})
        elif _is_degrade_type(str(r[0] or "").strip()):
            _htd_deg_ivs.append((_hms_to_sec(str(r[16] or "00:00:00")), _hms_to_sec(str(r[17] or "00:00:00"))))
        else:
            _rtype_full = str(r[0] or "").strip()
            if any(k in _rtype_full.lower() for k in ["pause","nettoyage","réunion","reunion","meeting"]):
                _htd_ded_s += _hms_to_sec(str(r[18] or "00:00:00"))
    _htd_deg_ivs_s = sorted((s, f) for s, f in _htd_deg_ivs if f > s)
    _htd_deg_mg = []
    for _s, _f in _htd_deg_ivs_s:
        if _htd_deg_mg and _s <= _htd_deg_mg[-1][1]: _htd_deg_mg[-1] = (_htd_deg_mg[-1][0], max(_htd_deg_mg[-1][1], _f))
        else: _htd_deg_mg.append((_s, _f))
    _htd_deg_s = sum(_f - _s for _s, _f in _htd_deg_mg)
    trs_shift=-1.0
    if _pers_pct_map and _htd_prod_raw and tot_eq > 0:
        trs_shift, _ = _option_b_trs(_htd_prod_raw, _htd_deg_mg, prod_ref)
    elif prod_ref>0 and shift_s>0 and tot_eq>0:
        _htd_adj = max(1.0, shift_s - _htd_ded_s)
        trs_shift=round(tot_eq/(prod_ref*_htd_adj/28800)*100,1)
    trs_of_time=-1.0
    if prod_ref>0 and tot_s>0: trs_of_time=round(tot_eq/(prod_ref*tot_s/28800)*100,1)
    return jsonify({"rows":rows,"trs_shift":trs_shift,"trs_of":trs_of_time,"tot_eq":round(tot_eq,1),"shift_s":shift_s})

@flask_app.route('/api/past_sessions')
def api_past_sessions():
    prod_ref = get_prod_ref()
    sessions = {}
    for rn, r in _decl_cache:
        # Use Date_poste (col 39) if available, else fall back to row date (col 2)
        date_str = str(r[39] if len(r) > 39 else "") .strip() or _row_date(r[2])
        if not date_str: continue
        pilot = str(r[4] or "")
        poste = str(r[3] or "")
        row_type = str(r[0] or "").strip().lower()
        key = f"{date_str}||{pilot}||{poste}"
        if key not in sessions:
            sessions[key] = {"date":date_str,"pilot":pilot,"poste":poste,"nb_of":0,"tot_equiv":0.0,"max_fin_s":0.0,"max_rn":0,"prod_raws":[]}
        if rn > sessions[key]["max_rn"]: sessions[key]["max_rn"] = rn
        if row_type in ("production","prod",""):
            try:
                eq = float(str(r[21] or 0).replace(",","."))
                fin_s = _hms_to_sec(str(r[17] or "00:00:00"))
                sessions[key]["nb_of"] += 1
                sessions[key]["tot_equiv"] += eq
                sessions[key]["prod_raws"].append(r)
                if fin_s > sessions[key]["max_fin_s"]: sessions[key]["max_fin_s"] = fin_s
            except: pass
    # Also gather stop events per session for planned deduction
    session_evts = {}
    for rn, r in _decl_cache:
        date_str2 = str(r[39] if len(r) > 39 else "").strip() or _row_date(r[2])
        if not date_str2: continue
        pilot2 = str(r[4] or ""); poste2 = str(r[3] or "")
        row_type2 = str(r[0] or "").strip().lower()
        key2 = f"{date_str2}||{pilot2}||{poste2}"
        if key2 not in session_evts: session_evts[key2] = []
        if row_type2 not in ("production","prod",""):
            session_evts[key2].append((rn, r))
    postes_map = load_postes_shift_map()
    # Clé de la session active en cours (pour calcul TRS live identique à api_fin_poste_data)
    _live_pilot = (_S.get("pilot") or "").strip()
    _live_poste = (_S.get("poste") or "").strip()
    _live_ss = _S.get("shift_start")
    _live_date_str = _live_ss.date().strftime("%d/%m/%Y") if _live_ss else datetime.date.today().strftime("%d/%m/%Y")
    _live_key = f"{_live_date_str}||{_live_pilot}||{_live_poste}" if _live_pilot else None
    result = []
    for key, s in sessions.items():
        trs = -1.0
        if prod_ref > 0 and s["tot_equiv"] > 0:
            try:
                parts = s["date"].split('/'); date_obj = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
            except: date_obj = None
            _evts_ps = session_evts.get(key, [])
            _prod_raws_ps = s.get("prod_raws", [])
            _pk = (s["pilot"].lower(), s["date"])
            _xl_trs_ps = None
            if _pk in postes_map:
                _pm_ps = postes_map[_pk]
                _xl_trs_ps = _pm_ps.get('trs')
                _pdeb_ps = _pm_ps['deb_dt']; _pfin_ps = _pm_ps['fin_dt']
                _mdur2 = max(0.0, (_pfin_ps - _pdeb_ps).total_seconds()) if _pfin_ps else get_shift_duration_s(s["poste"], date_obj)
            else:
                _mdur2 = get_shift_duration_s(s["poste"], date_obj)
            if _live_key and key == _live_key:
                # Poste en cours : TRS live identique à api_fin_poste_data (Option B + plan_ivs)
                _ses_ov = _S.get("budget_overrides") or {}
                _live_dur = get_current_shift_duration_s()
                if _pers_pct_map and _prod_raws_ps:
                    _blab_lv = ("pause_min","meeting_tol_min","clean_short_min","clean_long_min","clean_grand_min")
                    _plan_bdata_lv = {bk: float((_ses_ov.get(bk) if _ses_ov.get(bk) is not None else cfg.get(bk, 0)) or 0) * 60 for bk in _blab_lv}
                    _plan_used_lv = {bk: 0.0 for bk in _blab_lv}
                    _plan_ivs_lv = []
                    for _, r_lv in _evts_ps:
                        _bk_lv = _get_arret_budget_key(str(r_lv[0] or ''))
                        if _bk_lv and _bk_lv in _plan_bdata_lv:
                            _ds_lv = _hms_to_sec(str(r_lv[16] or '00:00:00'))
                            _fs_lv = _norm_fin(_ds_lv, _hms_to_sec(str(r_lv[17] or '00:00:00')))
                            _dur_lv = _fs_lv - _ds_lv
                            if _dur_lv > 0 and _plan_used_lv[_bk_lv] < _plan_bdata_lv[_bk_lv]:
                                _plan_ivs_lv.append((_ds_lv, _ds_lv + min(_dur_lv, _plan_bdata_lv[_bk_lv] - _plan_used_lv[_bk_lv])))
                            _plan_used_lv[_bk_lv] += _dur_lv
                    _deg_ivs_lv = _merged_degrade_ivs([re for _, re in _evts_ps])
                    trs, _ = _option_b_trs(_prod_raws_ps, _deg_ivs_lv, prod_ref, _plan_ivs_lv)
                elif _live_dur > 0 and prod_ref > 0 and s["tot_equiv"] > 0:
                    _live_ded = _compute_planned_deduction_s(_evts_ps, _ses_ov)
                    _live_el = max(1.0, _live_dur - _live_ded)
                    trs = round(s["tot_equiv"] / (prod_ref * _live_el / 28800) * 100, 1)
            elif _xl_trs_ps is not None and _xl_trs_ps > 0:
                trs = _xl_trs_ps
            elif _pers_pct_map and _prod_raws_ps:
                _deg_ivs_ps = _merged_degrade_ivs([re for _, re in _evts_ps])
                trs, _ = _option_b_trs(_prod_raws_ps, _deg_ivs_ps, prod_ref)
            else:
                planned_ded = _compute_planned_deduction_s(_evts_ps)
                if _mdur2 > 0 and prod_ref > 0 and s["tot_equiv"] > 0:
                    _el2 = max(1.0, _mdur2 - planned_ded)
                    trs = round(s["tot_equiv"] / (prod_ref * _el2 / 28800) * 100, 1)
        _pk_check = (s["pilot"].lower(), s["date"])
        if _pk_check not in postes_map:
            continue
        result.append({"date":s["date"],"pilot":s["pilot"],"poste":s["poste"],"nb_of":s["nb_of"],"tot_equiv":round(s["tot_equiv"],1),"trs":trs,"_rn":s["max_rn"]})
    def _date_sort_key(x):
        d = x["date"]
        try:
            p=d.split('/'); return (int(p[2]),int(p[1]),int(p[0]),x.get("_rn",0))
        except: return (0,0,0,0)
    result.sort(key=_date_sort_key, reverse=True)
    for x in result: x.pop("_rn", None)
    return jsonify(result[:60])

@flask_app.route('/api/period_report')
def api_period_report():
    """Rapport agrégé sur une période : tous les postes Postes-sheet dans la plage."""
    date_from_str = request.args.get('date_from', '').strip()  # yyyy-mm-dd
    date_to_str   = request.args.get('date_to',   '').strip()
    filter_pilot  = request.args.get('pilot', '').strip().lower()
    filter_poste  = request.args.get('poste', '').strip().lower()
    max_sessions  = int(request.args.get('max_sessions', 0) or 0)  # 0 = pas de limite
    def _parse_ymd(s):
        try: p=s.split('-'); return datetime.date(int(p[0]),int(p[1]),int(p[2]))
        except: return None
    def _parse_dmy(s):
        try: p=s.split('/'); return datetime.date(int(p[2]),int(p[1]),int(p[0]))
        except: return None
    dt_from = _parse_ymd(date_from_str)
    dt_to   = _parse_ymd(date_to_str)
    prod_ref    = get_prod_ref()
    postes_map  = load_postes_shift_map()
    # ── Build sessions ──
    sessions = {}
    for rn, r in _decl_cache:
        date_str = str(r[39] if len(r) > 39 else '').strip() or _row_date(r[2])
        if not date_str: continue
        pilot = str(r[4] or ''); poste = str(r[3] or '')
        if filter_pilot and pilot.lower() != filter_pilot: continue
        if filter_poste and poste.lower() != filter_poste: continue
        _pk = (pilot.lower(), date_str)
        if _pk not in postes_map: continue
        d_obj = _parse_dmy(date_str)
        if d_obj is None: continue
        if dt_from and d_obj < dt_from: continue
        if dt_to   and d_obj > dt_to:   continue
        row_type = str(r[0] or '').strip().lower()
        key = f"{date_str}||{pilot}||{poste}"
        if row_type in ('production','prod',''):
            if key not in sessions:
                sessions[key] = {'date':date_str,'pilot':pilot,'poste':poste,
                                 'nb_of':0,'tot_equiv':0.0,'tot_pcs':0,'evt_rows':[],'prod_rows':[],'prod_raws':[]}
            try:
                eq  = float(str(r[21] or 0).replace(',','.'))
                pcs = float(str(r[19] or 0).replace(',','.'))
                sessions[key]['nb_of']     += 1
                sessions[key]['tot_equiv'] += eq
                sessions[key]['tot_pcs']   += pcs
                _deb_f = _hms_to_sec(str(r[16] or '00:00:00'))
                _fib_f = str(r[11] or '').strip()
                sessions[key]['prod_rows'].append((_deb_f, _fib_f))
                sessions[key]['prod_raws'].append(r)
            except: pass
        else:
            # Event rows may have empty/different poste → attach to matching prod session by date+pilot
            _ev_key = key if key in sessions else next(
                (k for k in sessions if k.startswith(f"{date_str}||{pilot}||")), None
            )
            if _ev_key:
                sessions[_ev_key]['evt_rows'].append((rn, r))
    # Exclure la session en cours si demandé
    if request.args.get('skip_current') and _S.get('pilot') and _S.get('shift_debut_dt'):
        _cur_pilot_l = _S.get('pilot','').lower()
        _cur_date_str = _S['shift_debut_dt'].strftime('%d/%m/%Y')
        sessions = {k: v for k, v in sessions.items()
                    if not (v.get('pilot','').lower() == _cur_pilot_l and v.get('date') == _cur_date_str)}
    # Limiter aux N sessions les plus récentes si max_sessions > 0
    if max_sessions > 0 and len(sessions) > max_sessions:
        def _key_row(kv):
            _pk3 = (kv[1]['pilot'].lower(), kv[1]['date'])
            return postes_map.get(_pk3, {}).get('row_idx', 0)
        sessions = dict(sorted(sessions.items(), key=_key_row, reverse=True)[:max_sessions])
    # ── Aggregate ──
    _blab = {'pause_min','meeting_tol_min','clean_short_min','clean_long_min','clean_grand_min'}
    agg_ouv=0.0; agg_utile=0.0; agg_fonct=0.0; agg_stop=0.0; agg_perte=0.0
    agg_equiv=0.0; agg_pcs=0; agg_of=0; agg_elapsed_s=0.0; agg_sum_expected=0.0
    agg_sum_theorique=0.0
    agg_fibre_chg=0; agg_depassement=0.0; agg_degrade_min=0.0; stop_by_type={}; sessions_detail=[]
    jours=set(); pilotes=set(); postes_set=set()
    trs_by_day = {}
    cadence_ref = round(prod_ref/480, 4) if prod_ref > 0 else 0.0
    for key, s in sessions.items():
        _pk = (s['pilot'].lower(), s['date'])
        if _pk not in postes_map: continue
        _xl = postes_map[_pk]
        _pdeb = _xl['deb_dt']; _pfin = _xl['fin_dt']
        model_dur_s = max(0.0, (_pfin - _pdeb).total_seconds()) if _pfin else 0.0
        # Lire directement depuis Excel (colonnes R, S, T, U)
        ouv_min   = _xl['ouverture_min'] if _xl.get('ouverture_min') is not None else round(model_dur_s / 60, 1)
        fonct_min = _xl['fonct_min']     if _xl.get('fonct_min')     is not None else round(max(0.0, ouv_min), 1)
        utile_min = _xl['utile_min']     if _xl.get('utile_min')     is not None else ouv_min
        net_stop_min = _xl['arret_min'] if _xl.get('arret_min') is not None else round(max(0.0, ouv_min - fonct_min), 1)
        _xl_trs   = _xl.get('trs')
        _xl_perte = _xl.get('perte_min')
        # Merged degrade for evt_rows (still needed for perte/TRS fallback)
        _deg_s = _merged_degrade_s([re2 for _, re2 in s['evt_rows']])
        # Planned stops budget tracking (for budget_data display)
        _bdata = {bk:{'budget_min':float(cfg.get(bk,0) or 0),'used_min':0.0} for bk in _blab}
        for _, re3 in s['evt_rows']:
            _bk2 = _get_arret_budget_key(str(re3[0] or ''))
            if _bk2 and _bk2 in _bdata:
                _dp2=str(re3[18] or ''); _pp2=(_dp2+':00:00').split(':')
                try: _bs2=int(_pp2[0] or 0)*3600+int(_pp2[1] or 0)*60+int(_pp2[2] or 0)
                except: _bs2=0
                _bdata[_bk2]['used_min'] += _bs2/60
        arrets_prevu = sum(min(v['budget_min'],v['used_min']) for v in _bdata.values())
        planned_ded = _compute_planned_deduction_s(s['evt_rows'])
        elapsed_s = max(1.0, model_dur_s - planned_ded)
        adj_s = max(1.0, elapsed_s)
        _deg_ivs_pr = _merged_degrade_ivs([re2 for _, re2 in s['evt_rows']])
        _plan_bdata_pr = {bk: float(cfg.get(bk, 0) or 0) * 60 for bk in _blab}
        _plan_used_pr = {bk: 0.0 for bk in _blab}
        _plan_ivs_pr = []
        for _, re_ev in s['evt_rows']:
            _bk_pr = _get_arret_budget_key(str(re_ev[0] or ''))
            if _bk_pr and _bk_pr in _plan_bdata_pr:
                _ds_pr = _hms_to_sec(str(re_ev[16] or '00:00:00'))
                _fs_pr = _norm_fin(_ds_pr, _hms_to_sec(str(re_ev[17] or '00:00:00')))
                _dur_pr = _fs_pr - _ds_pr
                if _dur_pr > 0 and _plan_used_pr[_bk_pr] < _plan_bdata_pr[_bk_pr]:
                    _cap_pr = min(_dur_pr, _plan_bdata_pr[_bk_pr] - _plan_used_pr[_bk_pr])
                    _plan_ivs_pr.append((_ds_pr, _ds_pr + _cap_pr))
                _plan_used_pr[_bk_pr] += _dur_pr
        if _xl_trs is not None and _xl_trs > 0:
            _trs_s = _xl_trs
            _sum_exp_pr = s['tot_equiv'] * 100.0 / _xl_trs
            perte = _xl_perte if _xl_perte is not None else 0.0
        elif _pers_pct_map and s.get('prod_raws'):
            _trs_s, _sum_exp_pr = _option_b_trs(s['prod_raws'], _deg_ivs_pr, prod_ref, _plan_ivs_pr)
            perte = round((_sum_exp_pr - s['tot_equiv']) / cadence_ref, 1) if cadence_ref > 0 and _sum_exp_pr > 0 else 0.0
        else:
            _sum_exp_pr = prod_ref * adj_s / 28800
            _trs_s = round(s['tot_equiv']/(prod_ref*adj_s/28800)*100,1) if prod_ref>0 and adj_s>0 and s['tot_equiv']>0 else -1.0
            perte = round((_sum_exp_pr - s['tot_equiv']) / cadence_ref, 1) if cadence_ref>0 else 0.0
        agg_sum_expected += _sum_exp_pr
        depassement = sum(max(0.0, v['used_min'] - v['budget_min']) for v in _bdata.values())
        _xl_theorique = _xl.get('pcs_theorique') or _sum_exp_pr
        agg_sum_theorique += _xl_theorique
        agg_depassement += (_xl.get('depassement_min') if _xl.get('depassement_min') is not None else depassement)
        _pf = sorted(s.get('prod_rows', []), key=lambda x: x[0])
        nb_chg = sum(1 for i in range(1, len(_pf)) if _pf[i][1] and _pf[i-1][1] and _pf[i][1] != _pf[i-1][1])
        agg_fibre_chg += nb_chg
        for _, re_p in s['evt_rows']:
            _stype = str(re_p[0] or '').strip()
            if _stype and not _is_degrade_type(_stype):
                _ds_p = _hms_to_sec(str(re_p[16] or '00:00:00'))
                _fs_p = _hms_to_sec(str(re_p[17] or '00:00:00'))
                stop_by_type[_stype] = stop_by_type.get(_stype, 0.0) + max(0.0, _fs_p - _ds_p)
        agg_ouv     += ouv_min
        agg_utile   += utile_min
        agg_fonct   += fonct_min
        agg_stop    += net_stop_min
        agg_perte   += perte
        agg_equiv   += s['tot_equiv']
        agg_pcs     += s['tot_pcs']
        agg_of      += s['nb_of']
        agg_elapsed_s += adj_s
        jours.add(s['date']); pilotes.add(s['pilot']); postes_set.add(s['poste'])
        day = s['date']
        if day not in trs_by_day: trs_by_day[day]={'equiv':0.0,'elapsed_s':0.0,'sum_expected':0.0}
        trs_by_day[day]['equiv']    += s['tot_equiv']
        trs_by_day[day]['elapsed_s'] += elapsed_s
        trs_by_day[day]['sum_expected'] += _xl_theorique
        _cad_s = _xl.get('cadence_h') or (round(s['tot_equiv']/utile_min*60) if utile_min>0 else 0)
        _of_rows_sd = []
        for _rp in s.get('prod_raws', []):
            try:
                _deb_rp = _hms_to_sec(str(_rp[16] or "00:00:00"))
                _fin_rp = _hms_to_sec(str(_rp[17] or "00:00:00"))
                _dur_rp = _fin_rp - _deb_rp if _fin_rp > _deb_rp else _hms_to_sec(str(_rp[18] or "00:00:00"))
                _plan_rp = _deg_overlap_s(_deb_rp, _fin_rp, _plan_ivs_pr)
                _deg_rp  = _deg_overlap_s(_deb_rp, _fin_rp, _deg_ivs_pr)
                _nb_p_rp = max(1, int(float(str(_rp[6] or 1) or 1)))
                _pct_rp  = get_pct_cadence(_nb_p_rp)
                _adj_rp  = max(1.0, _dur_rp - _plan_rp - _deg_rp)
                _eq_rp   = float(str(_rp[21] or 0).replace(",", "."))
                _qte_rp  = float(str(_rp[19] or 0).replace(",", "."))
                _pcoef_rp = _eq_rp / _qte_rp if _qte_rp > 0 and _eq_rp > 0 else 1.0
                _exp_rp  = prod_ref * _pct_rp * _adj_rp / 28800 if prod_ref > 0 else 0.0
                _obj_rp  = round(_exp_rp / _pcoef_rp, 1) if _exp_rp > 0 else -1
            except:
                _plan_rp = 0; _deg_rp = 0; _obj_rp = -1
            _of_rows_sd.append({
                "of":str(_rp[1] or ""),"debut":str(_rp[16] or "")[:5],"fin":str(_rp[17] or "")[:5],
                "duree":str(_rp[18] or ""),"qte_fab":str(_rp[19] or ""),"qte_emb":str(_rp[20] or ""),
                "equiv":str(_rp[21] or ""),"cadence_h":str(_rp[22] if len(_rp)>22 else ""),
                "cadence_h_pers":str(_rp[23] if len(_rp)>23 else ""),"fibre":str(_rp[11] or ""),
                "taille":str(_rp[7] or ""),"code_prod":str(_rp[8] or ""),"type_prod":str(_rp[9] or ""),
                "poids":str(_rp[10] or ""),"of_taie":str(_rp[12] or ""),"traca":str(_rp[13] or ""),
                "ref_taie":str(_rp[14] or ""),"nb_pers":str(_rp[6] or ""),"copilote":str(_rp[5] or ""),
                "trs":str(_rp[24] if len(_rp)>24 else ""),
                "comment":str(_rp[35] if len(_rp)>35 else ""),
                "prevu_hors_trs":str(_rp[36] if len(_rp)>36 else ""),
                "qte_init_taie":str(_rp[25] if len(_rp)>25 else ""),
                "nb_taie2":str(_rp[26] if len(_rp)>26 else ""),
                "nb_def_cout":str(_rp[27] if len(_rp)>27 else ""),
                "mq_taie":str(_rp[28] if len(_rp)>28 else ""),
                "mq_housse":str(_rp[29] if len(_rp)>29 else ""),
                "nb_pp":str(_rp[30] if len(_rp)>30 else ""),
                "duree_mq_mp":str(_rp[32] if len(_rp)>32 else ""),
                "manquant_pers":str(_rp[33] if len(_rp)>33 else ""),
                "degrade_min":round(_deg_rp/60,1),"plan_stop_s":round(_plan_rp),"objectif":_obj_rp,
            })
        _evt_rows_sd = [{"type":str(re2[0] or ""),"of":str(re2[1] or ""),"debut":str(re2[16] or "")[:5],"fin":str(re2[17] or "")[:5],"duree":str(re2[18] or ""),"comment":str(re2[35] or ""),"is_degrade":_is_degrade_type(str(re2[0] or ""))} for _rn2, re2 in s.get('evt_rows',[])]
        agg_degrade_min += (_xl.get('degrade_min') if _xl.get('degrade_min') is not None else _deg_s / 60.0)
        sessions_detail.append({'date':s['date'],'pilot':s['pilot'],'poste':s['poste'],'trs':_trs_s,'cadence_h':_cad_s,'equiv':round(s['tot_equiv'],1),'degrade_min':round(_xl.get('degrade_min') if _xl.get('degrade_min') is not None else _deg_s/60.0, 1),'of_rows':_of_rows_sd,'evt_rows':_evt_rows_sd,'_deb_dt':_pdeb})
    _trs_denom = agg_sum_theorique if agg_sum_theorique > 0 else agg_sum_expected
    trs_periode = round(agg_equiv/_trs_denom*100,1) if _trs_denom>0 and agg_equiv>0 else -1.0
    def _sort_dmy(d):
        try: p=d.split('/'); return (int(p[2]),int(p[1]),int(p[0]))
        except: return (0,0,0)
    trs_by_day_list = [
        {'date':day,'trs':round(v['equiv']/v['sum_expected']*100,1) if v['sum_expected']>0 and v['equiv']>0 else -1.0}
        for day,v in sorted(trs_by_day.items(), key=lambda x:_sort_dmy(x[0]))
    ]
    cadence_h = round(agg_equiv*60/agg_utile) if agg_utile>0 else 0
    sessions_detail_sorted = sorted(sessions_detail, key=lambda x: x['_deb_dt'] or datetime.datetime.min)
    for _sd_item in sessions_detail_sorted: _sd_item.pop('_deb_dt', None)
    _agg_perte_xl = round((agg_sum_theorique - agg_equiv) / cadence_ref, 1) if cadence_ref > 0 and agg_sum_theorique > 0 else round(agg_perte, 1)
    return jsonify({
        'ok':True,
        'trs_periode':trs_periode,
        'nb_sessions':len(sessions),
        'nb_of':agg_of,
        'nb_jours':len(jours),
        'nb_pilotes':len(pilotes),
        'nb_postes':len(postes_set),
        'tot_equiv':round(agg_equiv,1),
        'tot_pcs':round(agg_pcs),
        'ouverture_min':round(agg_ouv,1),
        'temps_utile_min':round(agg_utile,1),
        'temps_fonctionnement_min':round(agg_fonct,1),
        'net_stop_min':round(agg_stop,1),
        'tot_degrade_min':round(agg_degrade_min,1),
        'perte_cadence_min':_agg_perte_xl,
        'cadence_ref_pcs_min':cadence_ref,
        'cadence_h':cadence_h,
        'trs_by_day':trs_by_day_list,
        'nb_fibre_chg':agg_fibre_chg,
        'depassement_min':round(agg_depassement,1),  # sum col M
        'sessions_detail':sessions_detail_sorted,
        'degrade_min_total': round(sum(_merged_degrade_s([re2 for _, re2 in s['evt_rows']]) for s in sessions.values()) / 60, 1),
        'stop_pareto':[{'type':k,'cat':(_t:=k.lower()) and ('nettoyage' if 'nettoyage' in _t else ('_pause' if _t=='pause' else ('ratt' if 'rattrapage' in _t else ('pb' if _t.startswith('pb') or 'panne' in _t else 'organisation')))),'min':round(v/60,1)} for k,v in sorted(stop_by_type.items(),key=lambda x:-x[1])[:15]],
    })

@flask_app.route('/api/session_report')
def api_session_report():
    date_str = request.args.get('date','')
    pilot = request.args.get('pilot','')
    poste = request.args.get('poste','')
    prod_ref = get_prod_ref()
    prod_rows = []; evt_rows = []; tot_eq = 0.0; tot_s = 0.0; max_fin_s = 0.0; stop_s = 0.0; _deg_ivs_sr = []; _prod_raws_sr = []
    all_debut_s = []; all_fin_s = []
    for rn, r in _decl_cache:
        row_date_key = str(r[39] if len(r) > 39 else "").strip() or _row_date(r[2])
        if row_date_key != date_str: continue
        if str(r[4] or "") != pilot: continue
        if str(r[3] or "") != poste: continue
        row_type = str(r[0] or "").strip().lower()
        deb_raw = str(r[16] or ""); fin_raw = str(r[17] or "")
        deb_s_r = _hms_to_sec(deb_raw) if deb_raw else -1
        fin_s_r = _hms_to_sec(fin_raw) if fin_raw else -1
        if deb_s_r >= 0: all_debut_s.append(deb_s_r)
        if fin_s_r >= 0: all_fin_s.append(fin_s_r)
        if row_type in ("production","prod",""):
            try:
                eq = float(str(r[21] or 0).replace(",","."))
                deb_s = _hms_to_sec(str(r[16] or "00:00:00"))
                fin_s = _hms_to_sec(str(r[17] or "00:00:00"))
                dur_s = fin_s - deb_s if fin_s > deb_s else _hms_to_sec(str(r[18] or "00:00:00"))
                try: _r24=float(str(r[24] if len(r)>24 else '').strip() or '-1')
                except: _r24=-1.0
                trs = _r24 if _r24>=0 else (round(eq/(prod_ref*dur_s/28800)*100,1) if prod_ref>0 and dur_s>0 and eq>0 else -1)
                tot_eq += eq; tot_s += dur_s
                if fin_s > max_fin_s: max_fin_s = fin_s
                _prod_raws_sr.append(r)
                prod_rows.append({"of":str(r[1] or ""),"taille":str(r[7] or ""),"code_prod":str(r[8] or ""),"type_prod":str(r[9] or ""),"poids":str(r[10] or ""),"fibre":str(r[11] or ""),"of_taie":str(r[12] or ""),"traca":str(r[13] or ""),"ref_taie":str(r[14] or ""),"kit":str(r[15] or ""),"qte_fab":str(r[19] or ""),"qte_emb":str(r[20] or ""),"equiv":str(r[21] or ""),"cadence_h":str(r[22] if len(r)>22 else ""),"cadence_h_pers":str(r[23] if len(r)>23 else ""),"debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],"duree":str(r[18] or ""),"trs":trs,"comment":str(r[35] or ""),"prevu_hors_trs":str(r[36] if len(r)>36 else ""),"nb_pers":str(r[6] or ""),"copilote":str(r[5] or ""),"qte_init_taie":str(r[25] if len(r)>25 else ""),"nb_taie2":str(r[26] if len(r)>26 else ""),"nb_def_cout":str(r[27] if len(r)>27 else ""),"mq_taie":str(r[28] if len(r)>28 else ""),"mq_housse":str(r[29] if len(r)>29 else ""),"nb_pp":str(r[30] if len(r)>30 else ""),"duree_mq_mp":str(r[32] if len(r)>32 else ""),"manquant_pers":str(r[33] if len(r)>33 else "")})
            except: pass
        elif _is_degrade_type(str(r[0] or "").strip()):
            try:
                _deg_ivs_sr.append((_hms_to_sec(str(r[16] or "00:00:00")), _hms_to_sec(str(r[17] or "00:00:00"))))
                evt_rows.append({"type":str(r[0] or ""),"of":str(r[1] or ""),"taille":str(r[7] or ""),"type_prod":str(r[9] or ""),"debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],"duree":str(r[18] or ""),"comment":str(r[35] or ""),"is_degrade":True})
            except: pass
        else:
            try:
                dur_s = _hms_to_sec(str(r[18] or "00:00:00"))
                stop_s += dur_s
                evt_rows.append({"type":str(r[0] or ""),"of":str(r[1] or ""),"taille":str(r[7] or ""),"type_prod":str(r[9] or ""),"debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],"duree":str(r[18] or ""),"comment":str(r[35] or ""),"is_degrade":False})
            except: pass
    _deg_ivs_sr_s = sorted((s, f) for s, f in _deg_ivs_sr if f > s)
    _deg_mg_sr = []
    for _s, _f in _deg_ivs_sr_s:
        if _deg_mg_sr and _s <= _deg_mg_sr[-1][1]: _deg_mg_sr[-1] = (_deg_mg_sr[-1][0], max(_deg_mg_sr[-1][1], _f))
        else: _deg_mg_sr.append((_s, _f))
    degrade_s = sum(_f - _s for _s, _f in _deg_mg_sr)
    actual_debut = _sec_to_hm(min(all_debut_s)) if all_debut_s else ""
    actual_fin = _sec_to_hm(max(all_fin_s)) if all_fin_s else ""
    try:
        _date_obj_rpt = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
    except Exception:
        _date_obj_rpt = None
    debut_str, fin_str = _get_model_day_cfg(poste, _date_obj_rpt)
    _sr_ss = _S.get("shift_start")
    _sr_live_date = _sr_ss.date().strftime("%d/%m/%Y") if _sr_ss else datetime.date.today().strftime("%d/%m/%Y")
    _is_live_sr = bool(_S.get("pilot") and _S.get("pilot") == pilot and _S.get("poste") == poste and date_str == _sr_live_date)
    _sr_live_ov = (_S.get("budget_overrides") or {}) if _is_live_sr else {}
    planned_ded = _compute_planned_deduction_s(evt_rows, _sr_live_ov if _is_live_sr else None)
    # Pour session en cours : intervalles planifiés capés au budget (comme api_fin_poste_data)
    _plan_ivs_sr = []
    _blab_sr = ("pause_min","meeting_tol_min","clean_short_min","clean_long_min","clean_grand_min")
    _plan_bdata_sr = {bk: float((_sr_live_ov.get(bk) if _sr_live_ov.get(bk) is not None else cfg.get(bk, 0)) or 0) * 60 for bk in _blab_sr}
    _plan_used_sr = {bk: 0.0 for bk in _blab_sr}
    for rn2, r2 in _decl_cache:
        _rd2 = str(r2[39] if len(r2) > 39 else "").strip() or _row_date(r2[2])
        if _rd2 != date_str: continue
        if str(r2[4] or "") != pilot or str(r2[3] or "") != poste: continue
        if str(r2[0] or "").strip().lower() in ("production","prod",""): continue
        _bk_sr = _get_arret_budget_key(str(r2[0] or ''))
        if _bk_sr and _bk_sr in _plan_bdata_sr:
            _ds_sr = _hms_to_sec(str(r2[16] or '00:00:00'))
            _fs_sr = _norm_fin(_ds_sr, _hms_to_sec(str(r2[17] or '00:00:00')))
            _dur_sr = _fs_sr - _ds_sr
            if _dur_sr > 0 and _plan_used_sr[_bk_sr] < _plan_bdata_sr[_bk_sr]:
                _plan_ivs_sr.append((_ds_sr, _ds_sr + min(_dur_sr, _plan_bdata_sr[_bk_sr] - _plan_used_sr[_bk_sr])))
            _plan_used_sr[_bk_sr] += _dur_sr
    # Per-OF: calcul arrêts prévus et objectif pièces
    for pi, raw_r in enumerate(_prod_raws_sr):
        try:
            _deb_of = _hms_to_sec(str(raw_r[16] or "00:00:00"))
            _fin_of = _hms_to_sec(str(raw_r[17] or "00:00:00"))
            _dur_of = _fin_of - _deb_of if _fin_of > _deb_of else _hms_to_sec(str(raw_r[18] or "00:00:00"))
            _plan_of_s = _deg_overlap_s(_deb_of, _fin_of, _plan_ivs_sr)
            _deg_of_s = _deg_overlap_s(_deb_of, _fin_of, _deg_mg_sr)
            _nb_p = max(1, int(float(str(raw_r[6] or 1) or 1)))
            _pct_of = get_pct_cadence(_nb_p)
            _adj_of_s = max(1.0, _dur_of - _plan_of_s - _deg_of_s)
            _eq_of = float(str(raw_r[21] or 0).replace(",", "."))
            _qte_of = float(str(raw_r[19] or 0).replace(",", "."))
            _pcoef_of = _eq_of / _qte_of if _qte_of > 0 and _eq_of > 0 else 1.0
            _exp_of = prod_ref * _pct_of * _adj_of_s / 28800 if prod_ref > 0 else 0.0
            _obj_of = round(_exp_of / _pcoef_of, 1) if _exp_of > 0 else -1
            prod_rows[pi]["plan_stop_s"] = round(_plan_of_s)
            prod_rows[pi]["objectif"] = _obj_of
        except: pass
    _postes_map2 = load_postes_shift_map()
    _pk2 = (pilot.lower(), date_str)
    _xl_trs_sr = None
    _xl_perte_sr = None
    if _pk2 in _postes_map2:
        _pm2 = _postes_map2[_pk2]
        _pdeb2 = _pm2['deb_dt']; _pfin2 = _pm2['fin_dt']
        _xl_trs_sr = _pm2.get('trs')
        _xl_perte_sr = _pm2.get('perte_min')
        model_dur_s = max(0.0, (_pfin2 - _pdeb2).total_seconds()) if _pfin2 else (get_current_shift_duration_s() if _is_live_sr else 0.0)
        if not debut_str:
            debut_str = _pdeb2.strftime("%H:%M")
            fin_str = _pfin2.strftime("%H:%M") if _pfin2 else ''
    else:
        model_dur_s = get_shift_duration_s(poste)
    ecart_s = max(0.0, model_dur_s - (tot_s + stop_s))
    trs_shift = -1.0
    perte_cadence_s = 0.0
    if _xl_trs_sr is not None and _xl_trs_sr > 0 and tot_eq > 0:
        trs_shift = _xl_trs_sr
        _cadence_ref_s = prod_ref / 28800
        if _cadence_ref_s > 0:
            _sum_exp_xl = tot_eq * 100.0 / _xl_trs_sr
            perte_cadence_s = (_sum_exp_xl - tot_eq) / _cadence_ref_s
            if _xl_perte_sr is not None:
                perte_cadence_s = _xl_perte_sr * 60.0
    elif _pers_pct_map and _prod_raws_sr and tot_eq > 0:
        trs_shift, _sum_exp_sr = _option_b_trs(_prod_raws_sr, _deg_mg_sr, prod_ref, _plan_ivs_sr)
        _cadence_ref_s = prod_ref / 28800
        if _cadence_ref_s > 0 and _sum_exp_sr > 0:
            perte_cadence_s = (_sum_exp_sr - tot_eq) / _cadence_ref_s
    elif model_dur_s > 0 and prod_ref > 0 and tot_eq > 0:
        elapsed_s = max(1.0, model_dur_s - planned_ded)
        adj_s = max(1.0, elapsed_s)
        trs_shift = round(tot_eq/(prod_ref*adj_s/28800)*100,1)
        _cadence_ref_s = prod_ref / 28800
        if _cadence_ref_s > 0:
            perte_cadence_s = (prod_ref * adj_s / 28800 - tot_eq) / _cadence_ref_s
    trs_of = round(tot_eq/(prod_ref*tot_s/28800)*100,1) if prod_ref>0 and tot_s>0 and tot_eq>0 else -1
    _budget_labels = {"pause_min":"Pause","meeting_tol_min":"Réunion","clean_short_min":"Nettoyage court","clean_long_min":"Nettoyage long","clean_grand_min":"Nettoyage très long"}
    budget_data = {bk:{"label":bl,"budget_min":float(cfg.get(bk,0) or 0),"used_min":0.0} for bk,bl in _budget_labels.items()}
    for _er in evt_rows:
        _bk = _get_arret_budget_key(_er.get('type','') or _er.get('comment',''))
        if _bk and _bk in budget_data:
            _dp = (_er.get('duree') or ''); _pp = (_dp+':00:00').split(':')
            try: _bs = int(_pp[0] or 0)*3600+int(_pp[1] or 0)*60+int(_pp[2] or 0)
            except: _bs = 0
            budget_data[_bk]['used_min'] += _bs/60
    _pm2_xl = _postes_map2.get(_pk2, {})
    return jsonify({"date":date_str,"pilot":pilot,"poste":poste,"prod_rows":prod_rows,"evt_rows":evt_rows,"budget_data":budget_data,
                    "trs_shift":trs_shift,"trs":trs_of,"tot_equiv":round(tot_eq,1),"tot_s":round(tot_s,0),
                    "stop_s":round(stop_s,0),"nb_of":len(prod_rows),
                    "degrade_s":round(degrade_s,0),
                    "perte_cadence_min":round(perte_cadence_s/60,1),
                    "model_debut":debut_str or "","model_fin":fin_str or "",
                    "actual_debut":actual_debut,"actual_fin":actual_fin,
                    "ecart_s":round(ecart_s,0),"model_dur_s":round(model_dur_s,0),
                    "planned_ded_s":round(planned_ded,0),"prod_ref":round(prod_ref,1),
                    "is_live":_is_live_sr,
                    "ouverture_min":_pm2_xl.get('ouverture_min'),
                    "utile_min":_pm2_xl.get('utile_min'),
                    "fonct_min":_pm2_xl.get('fonct_min'),
                    "arret_min":_pm2_xl.get('arret_min'),
                    "degrade_min":_pm2_xl.get('degrade_min'),
                    "cadence_h":_pm2_xl.get('cadence_h')})

@flask_app.route('/api/add_stop_decl', methods=['POST'])
def api_add_stop_decl():
    """Ajoute rétroactivement une déclaration d'arrêt (pause/nettoyage/réunion) pour le pilote connecté."""
    data = request.json or {}
    stop_type = str(data.get("type","")).strip()
    debut_hms = str(data.get("debut_hms","")).strip()
    fin_hms   = str(data.get("fin_hms","")).strip()
    comment   = str(data.get("comment","")).strip()
    if not stop_type or not debut_hms or not fin_hms:
        return jsonify({"ok":False,"error":"type/debut/fin requis"}),400
    pilot = _S.get("pilot","")
    poste = _S.get("poste","")
    if not pilot:
        return jsonify({"ok":False,"error":"Pas de pilote connecté"}),400
    now = datetime.datetime.now()
    try:
        dh,dm = [int(x) for x in debut_hms.split(":")[:2]]
        fh,fm = [int(x) for x in fin_hms.split(":")[:2]]
        start_dt = now.replace(hour=dh, minute=dm, second=0, microsecond=0)
        end_dt   = now.replace(hour=fh, minute=fm, second=0, microsecond=0)
        if end_dt <= start_dt: end_dt += datetime.timedelta(days=1)  # poste de nuit
        dur_s = max(0, (end_dt - start_dt).total_seconds())
        shift_dt2 = _S.get("shift_start") or now
        shift_date2 = shift_dt2.strftime("%d/%m/%Y")
        row = [
            stop_type, _S.get("form",{}).get("of_num",""),
            start_dt.strftime("%d/%m/%Y"), poste, pilot,
            "","","","","","","","","","",
            "",                              # col P : vide pour les arrêts
            start_dt.strftime("%H:%M:%S"), end_dt.strftime("%H:%M:%S"), fmt(dur_s),
            "","","","","","","","","","","","","","","","",comment,"","","",
            shift_date2,
        ]
        write_excel_bg([], [row])
        # Synchronously update cache so fin_poste_data sees it immediately
        try:
            next_rn = max((rn for rn,_ in _decl_cache), default=1) + 1
            padded = tuple(row) + ('',) * max(0, 40 - len(row))
            _decl_cache.append((next_rn, padded))
        except: pass
    except Exception as e:
        return jsonify({"ok":False,"error":str(e)}),500
    return jsonify({"ok":True})

@flask_app.route('/api/reload_excel', methods=['POST'])
def api_reload_excel():
    """Force un rechargement du cache Excel depuis le disque."""
    load_history()
    return jsonify({"ok": True, "rows": len(_decl_cache)})

@flask_app.route('/api/update_of_time', methods=['POST'])
def api_update_of_time():
    """Modifie l'heure début/fin d'un OF déclaré (pour correction écart fin de poste)."""
    data = request.json or {}
    of_num = str(data.get("of_num","")).strip()
    old_debut = str(data.get("old_debut","")).strip()  # HH:MM
    new_debut = str(data.get("new_debut","")).strip()  # HH:MM
    new_fin   = str(data.get("new_fin","")).strip()    # HH:MM
    if not of_num or not new_debut or not new_fin:
        return jsonify({"ok":False,"error":"of_num/new_debut/new_fin requis"}),400
    pilot = _S.get("pilot","")
    now = datetime.datetime.now()
    today = now.strftime("%d/%m/%Y")
    shift_start_dt = _S.get("shift_start")
    shift_date_str = shift_start_dt.strftime("%d/%m/%Y") if shift_start_dt else today
    # Find matching row in cache
    target_rn = None
    for rn, r in _decl_cache:
        rd = _row_date(r[2])
        if rd not in (today, shift_date_str): continue
        if str(r[4] or "") != pilot: continue
        if str(r[0] or "").strip().lower() not in ("production","prod",""): continue
        row_of = str(r[1] or "")
        row_debut = str(r[16] or "")[:5]
        if row_of == of_num and (not old_debut or row_debut == old_debut):
            target_rn = rn
            break
    if target_rn is None:
        # Fallback: search directly in Excel file (cache may not be updated yet after background write)
        path = cfg.get("db_path","")
        if path:
            try:
                with _excel_lock:
                    wb = _get_wb(path)
                    if wb is not None:
                        ws = wb["Declarations"] if "Declarations" in wb.sheetnames else wb.active
                        for row in ws.iter_rows(min_row=2, values_only=False):
                            rn_fb = row[0].row
                            r_fb = tuple(c.value for c in row)
                            if len(r_fb) < 18: continue
                            rd = _row_date(r_fb[2])
                            if rd not in (today, shift_date_str): continue
                            if str(r_fb[4] or "") != pilot: continue
                            if str(r_fb[0] or "").strip().lower() not in ("production","prod",""): continue
                            row_of = str(r_fb[1] or "")
                            row_debut = str(r_fb[16] or "")[:5]
                            if row_of == of_num and (not old_debut or row_debut == old_debut):
                                target_rn = rn_fb
                                break
            except Exception as _e:
                pass
    if target_rn is None:
        return jsonify({"ok":False,"error":"OF non trouvé"}),404
    try:
        dh,dm = [int(x) for x in new_debut.split(":")[:2]]
        fh,fm = [int(x) for x in new_fin.split(":")[:2]]
        start_dt2 = now.replace(hour=dh, minute=dm, second=0, microsecond=0)
        end_dt2   = now.replace(hour=fh, minute=fm, second=0, microsecond=0)
        if end_dt2 <= start_dt2: end_dt2 += datetime.timedelta(days=1)
        dur_s2 = max(0, (end_dt2 - start_dt2).total_seconds())
        new_debut_hms = f"{dh:02d}:{dm:02d}:00"
        new_fin_hms   = f"{fh:02d}:{fm:02d}:00"
        dur_hms = fmt(dur_s2)
    except Exception as e:
        return jsonify({"ok":False,"error":str(e)}),400
    # Update in-memory cache
    for i, (rn, r) in enumerate(_decl_cache):
        if rn == target_rn:
            r_list = list(r)
            r_list[16] = new_debut_hms
            r_list[17] = new_fin_hms
            r_list[18] = dur_hms
            _decl_cache[i] = (rn, tuple(r_list))
            break
    # Update Excel in background
    path = cfg.get("db_path","")
    if path:
        def _bg():
            try:
                with _excel_lock:
                    wb = _get_wb(path)
                    if wb is None: return
                    ws = wb["Declarations"] if "Declarations" in wb.sheetnames else wb.active
                    ws.cell(target_rn, 17).value = new_debut_hms
                    ws.cell(target_rn, 18).value = new_fin_hms
                    ws.cell(target_rn, 19).value = dur_hms
                    _safe_excel_save(wb, path)
            except: pass
        threading.Thread(target=_bg, daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/reload', methods=['POST'])
def api_reload():
    threading.Thread(target=load_lists,daemon=True).start()
    threading.Thread(target=load_history,daemon=True).start()
    return jsonify({"ok":True})

@flask_app.route('/api/save_poste', methods=['POST'])
def api_save_poste():
    data = request.json or {}
    if not data.get("dur_poste_theorique_min"):
        data["dur_poste_theorique_min"] = round(get_current_shift_duration_s() / 60, 1)
    # Utiliser postes_row_num stocké au login (évite find_postes_row_num qui peut rater)
    row_num = _S.get("postes_row_num") or find_postes_row_num(_S.get("pilot",""), _S.get("shift_debut_dt"))
    write_poste_row(data, row_num=row_num)
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

@flask_app.route('/api/change_admin_pw', methods=['POST'])
def api_change_admin_pw():
    data = request.json or {}
    old_pw = data.get("old_pw","")
    new_pw = str(data.get("new_pw","")).strip()
    if not _check_pw(old_pw):
        return jsonify({"ok":False,"error":"Mot de passe actuel incorrect"}),403
    if not new_pw:
        return jsonify({"ok":False,"error":"Le nouveau MDP ne peut pas être vide"}),400
    cfg["supervisor_pw"] = new_pw
    save_cfg_data()
    ok = write_admin_pw_to_excel(new_pw)
    return jsonify({"ok":ok})

@flask_app.route('/api/save_list', methods=['POST'])
def api_save_list():
    data = request.json or {}
    pw = data.get("pw","")
    if not _check_pw(pw):
        return jsonify({"ok":False,"error":"Mot de passe incorrect"}),403
    list_type = data.get("list_type","")
    items = data.get("items",[])
    if list_type == "copilotes":
        _lists["copilotes"] = items
        ok = write_simple_list_to_excel(3, "Co-pilotes", items)
    elif list_type == "tailles":
        _lists["tailles_col"] = items
        ok = write_simple_list_to_excel(4, "Taille produit", items)
    elif list_type == "fibres":
        _lists["fibres_col"] = items
        ok = write_simple_list_to_excel(10, "Fibres", items)
    elif list_type == "equiv":
        _lists["types_prod_col"] = [it.get("type","") for it in items if it.get("type","")]
        _lists["equivalences_col"] = [it.get("coeff","") for it in items if it.get("type","")]
        ok = write_equiv_list_to_excel(items)
    else:
        return jsonify({"ok":False,"error":"Type de liste inconnu"}),400
    return jsonify({"ok":bool(ok)})

@flask_app.route('/api/generate_dashboard', methods=['POST'])
def api_generate_dashboard():
    html_path, err = generate_dashboard_html()
    if err: return jsonify({"ok":False,"error":err})
    return jsonify({"ok":True,"path":html_path})

@flask_app.route('/dashboard')
def dashboard_view():
    """Serve the dashboard HTML via Flask so JS fetch() is same-origin (no CORS block)."""
    html_path, err = generate_dashboard_html()
    if err:
        return f"<html><body style='font-family:sans-serif;padding:40px;color:#dc2626'><h2>Erreur Dashboard</h2><p>{err}</p></body></html>", 500
    try:
        with open(html_path, encoding='utf-8') as f:
            content = f.read()
        return content, 200, {'Content-Type': 'text/html; charset=utf-8'}
    except Exception as e:
        return f"<html><body>Erreur lecture fichier: {e}</body></html>", 500

def _dash_budget_bars_html():
    """Génère les barres de budget arrêts prévus pour le dashboard (rendu serveur)."""
    try:
        bs = _compute_budget_state_now()
    except:
        return ''
    LABELS = ["Pause","Réunion","Nettoyage court","Nettoyage long","Nettoyage très long"]
    def _fmt_dur_s(s):
        s = int(round(s))
        m, sec = divmod(s, 60)
        return f"{m}m {sec:02d}s" if m else f"{sec}s"
    rows = ""
    any_budget = False
    for lbl in LABELS:
        d = bs["per_type"].get(lbl, {})
        budget = d.get("budget_s", 0)
        if budget <= 0: continue
        any_budget = True
        consumed = d.get("consumed_s", 0)
        pct = min(100, consumed / budget * 100) if budget > 0 else 0
        over = max(0, consumed - budget)
        color = "#dc2626" if pct >= 100 else ("#d97706" if pct >= 70 else "#16a34a")
        val_str = f"<b>+{_fmt_dur_s(over)}</b>" if over > 0 else f"{_fmt_dur_s(consumed)} / {_fmt_dur_s(budget)}"
        rows += (f'<div style="margin-bottom:6px">'
            f'<div style="display:flex;justify-content:space-between;font-size:calc(11px*var(--zf,1));font-weight:700;margin-bottom:2px">'
            f'<span style="color:#374151">{lbl}</span>'
            f'<span style="color:{color}">{val_str}</span></div>'
            f'<div style="background:#e5e7eb;border-radius:4px;height:7px">'
            f'<div style="background:{color};width:{min(100,pct):.0f}%;height:7px;border-radius:4px"></div>'
            f'</div></div>')
    if not any_budget:
        return ''
    return (f'<div style="background:#fefce8;border:1px solid #fde68a;border-radius:8px;padding:8px 12px;min-width:170px">'
        f'<div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#92400e;letter-spacing:.5px;margin-bottom:8px">⏱ Arrêts prévus</div>'
        f'{rows}</div>')

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
    kit_now = "Oui" if (_S.get("form") or {}).get("kit") else "Non"
    _form_now = _S.get("form") or {}
    copilote_now = str(_form_now.get("copilote","") or "")
    nb_pers_now  = str(_form_now.get("nb_pers","") or "")
    code_prod_now= str(_form_now.get("code_prod","") or "")
    poids_now    = str(_form_now.get("poids","") or "")
    fibre_now    = str(_form_now.get("fibre","") or "")
    of_taie_now  = str(_form_now.get("of_taie","") or "")
    traca_now    = str(_form_now.get("traca","") or "")
    ref_taie_now = str(_form_now.get("ref_taie","") or "")
    qte_fab_now  = str(_form_now.get("qte_fab","") or "")
    qte_emb_now  = str(_form_now.get("qte_emb","") or "")
    comment_now  = str(_form_now.get("comment","") or "")
    of_start_dt = _S.get("of_start")
    is_paused = _S.get("is_paused", False)
    # Temps écoulé depuis début OF (en cours)
    of_elapsed_s = 0.0
    if prod_active and of_start_dt:
        try: of_elapsed_s = (datetime.datetime.now() - of_start_dt).total_seconds()
        except: pass
    # Arrêts en cours cumulés sur cet OF (depuis of_start via tl_events live)
    of_stop_s = 0.0
    for ev in (_S.get("tl_events") or []):
        if of_start_dt and ev.get("key") != "_prod":
            try:
                ev_start = ev.get("start") or ev.get("t_start")
                if ev_start and ev_start >= of_start_dt:
                    ev_end = ev.get("end") or ev.get("t_end") or datetime.datetime.now()
                    of_stop_s += max(0, (ev_end - ev_start).total_seconds())
            except: pass

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
    _last_fin_s = 0.0
    for r in in_shift_prod:
        fs = hms2s(r[17])
        if fs > _last_fin_s:
            _last_fin_s = fs
    if _last_fin_s > 0:
        try:
            _h = int(_last_fin_s // 3600)
            _mi = int((_last_fin_s % 3600) // 60)
            _sc = int(_last_fin_s % 60)
            last_fin_dt = datetime.datetime.combine(datetime.date.today(), datetime.time(_h, _mi, _sc))
        except: pass
    # Match accueil formula: use last declared fin, not now()
    # (accueil uses _lastProdDeclTime which is fin of last declared OF)
    if prod_active and last_fin_dt is None:
        last_fin_dt = datetime.datetime.now()
    prod_s_total = sum(hms2s(str(r[18] or "0")) for r in in_shift_prod)
    # Filter stop events to model horaire window (same logic as in_shift_prod)
    if model_debut_dt:
        in_shift_evts = [r for r in today_evts if _ts_s(r[17]) >= _mdeb_s or _ts_s(r[16]) >= _mdeb_s]
    else:
        in_shift_evts = today_evts
    stop_s_total = sum(hms2s(str(r[18] or "0")) for r in in_shift_evts
                       if str(r[0] or "").lower() not in ("pause pilote","changement d'of","interposte","changement de serie"))
    ref_start_dt = _S.get("shift_debut_dt") or model_debut_dt or shift_start_dt
    if ref_start_dt and last_fin_dt and prod_ref > 0:
        elapsed_for_trs = (last_fin_dt - ref_start_dt).total_seconds()
        if elapsed_for_trs > 0:
            _bgt = _compute_budget_state_now()
            _shift_ded = _bgt["total_shift_deductible_s"]
            _adj_elapsed = max(1.0, elapsed_for_trs - _shift_ded)
            trs_poste = round(tot_equiv / (prod_ref * _adj_elapsed / 28800) * 100, 1)

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
                f'<text x="{cx}" y="{cy+6}" text-anchor="middle" font-size="35" font-weight="900" fill="{col}">{lbl_txt}</text>'
                f'</svg>')

    # ── Pie chart SVG ──
    def pie_svg(prod_s, stop_s, size=180):
        total = prod_s + stop_s
        if total <= 0:
            return f'<svg width="{size}" height="{size}"><text x="{size//2}" y="{size//2+6}" text-anchor="middle" font-size="20" fill="#475569">Pas de données</text></svg>'
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
        lbl = f'<text x="{cx}" y="{cy-6}" text-anchor="middle" font-size="25" font-weight="900" fill="#fff">{pp}%</text>'
        lbl += f'<text x="{cx}" y="{cy+16}" text-anchor="middle" font-size="16" fill="#cbd5e1">Prod</text>'
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
        catcol = {"pb":"#ef4444","ratt":"#f59e0b","nettoyage":"#f97316","pause":"#64748b","organisation":"#3b82f6","reunion":"#8b5cf6","degrade":"url(#deg-pat)"}
        svg = f'<svg width="100%" viewBox="0 0 {W} {H}" style="display:block" preserveAspectRatio="none">'
        svg += '<defs><pattern id="deg-pat" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="4" height="8" fill="#16a34a"/><rect x="4" y="0" width="4" height="8" fill="#fef08a"/></pattern></defs>'
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
            col = catcol["pb"] if ("pb" in t or "panne" in t or "technique" in t) else catcol["ratt"] if "ratt" in t else catcol["nettoyage"] if "nett" in t else catcol["reunion"] if ("réunion" in t or "reunion" in t or "meeting" in t) else catcol["pause"] if "pause" in t else catcol["degrade"] if ("dégr" in t or "degrad" in t or "mode" in t) else "#94a3b8"
            svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="{col}" rx="2" opacity="0.95"/>'
        # Live events from _S (in-memory, not yet in Excel)
        _catcol2 = {"pb":"#ef4444","ratt":"#f59e0b","nettoyage":"#f97316","pause":"#64748b","organisation":"#3b82f6","reunion":"#8b5cf6"}
        for _dp in (_S.get("degrade_periods") or []):
            _d0 = _dp.get("start"); _d1 = _dp.get("end") or now_ts
            if _d0:
                x1 = to_x(_d0); x2 = to_x(_d1); x2 = max(x2, x1+3)
                svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="url(#deg-pat)" rx="2" opacity="0.85"/>'
        if _S.get("degrade_active") and _S.get("degrade_start_dt"):
            x1 = to_x(_S["degrade_start_dt"]); x2 = int((now_ts-win_start).total_seconds()/span*W); x2 = max(x2, x1+3)
            svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="url(#deg-pat)" rx="2" opacity="0.85"/>'
        for _ev in (_S.get("tl_events") or []):
            _key = _ev.get("key","")
            if not _key or _key.startswith("_"): continue
            _ev_start = _ev.get("start")
            if not _ev_start: continue
            _ev_end = _ev.get("end") or now_ts
            _cat = _ev.get("cat","autre")
            _col = _catcol2.get(_cat, "#94a3b8")
            try:
                x1 = to_x(_ev_start)
                x2 = to_x(_ev_end)
                x2 = max(x2, x1+3)
                if x2 > x1:
                    svg += f'<rect x="{x1}" y="{Y}" width="{x2-x1}" height="{BH}" fill="{_col}" rx="2" opacity="0.95"/>'
            except: pass
        h_span = span / 3600
        step = 1 if h_span <= 10 else 2
        cur = win_start.replace(minute=0, second=0, microsecond=0)
        if cur < win_start: cur += datetime.timedelta(hours=1)
        while cur <= win_end:
            frac = (cur-win_start).total_seconds()/span
            x = int(frac*W)
            svg += f'<line x1="{x}" y1="{Y}" x2="{x}" y2="{Y+BH}" stroke="#94a3b8" stroke-width="1"/>'
            svg += f'<text x="{x}" y="{Y-3}" font-size="15" fill="#475569" text-anchor="middle">{cur.strftime("%H:%M")}</text>'
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
            col = "#ef4444" if any(x in lbl.lower() for x in ["pb","panne","technique"]) else "#f59e0b" if "ratt" in lbl.lower() else "#f97316" if "nett" in lbl.lower() else "#8b5cf6" if any(x in lbl.lower() for x in ["réunion","reunion","meeting"]) else "#3b82f6"
            pareto_html += f'''<div style="margin-bottom:4px">
              <div style="display:flex;justify-content:space-between;font-size:calc(12px*var(--zf,1));color:#475569;margin-bottom:2px">
                <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:75%">{lbl[:30]}</span>
                <span style="font-weight:800;color:#1e293b;flex-shrink:0">{dur/60:.0f}m</span>
              </div>
              <div style="background:#e2e8f0;border-radius:3px;height:12px;width:100%">
                <div style="width:{pct:.0f}%;height:12px;background:{col};border-radius:3px"></div>
              </div>
            </div>'''
    else:
        pareto_html = '<div style="color:#475569;font-size:calc(13px*var(--zf,1));padding:6px;text-align:center">Aucun arrêt enregistré</div>'
    # ── Liste arrêts individuels ──
    _stop_list_html = ""
    for _sr in today_evts:
        _st = str(_sr[0] or "").strip()
        _sd = str(_sr[16] or "")[:5]
        _sf = str(_sr[17] or "")[:5]
        _sdur = str(_sr[18] or "")
        _scmt = str(_sr[35] or "").strip()
        _scol = "#ef4444" if any(x in _st.lower() for x in ["pb","panne","technique"]) else "#f59e0b" if "ratt" in _st.lower() else "#f97316" if "nett" in _st.lower() else "#8b5cf6" if any(x in _st.lower() for x in ["réunion","reunion","meeting"]) else "#3b82f6"
        _stop_list_html += (f'<tr>'
            f'<td style="padding:3px 6px;font-size:calc(11px*var(--zf,1));font-weight:700;color:{_scol};white-space:nowrap;max-width:100px;overflow:hidden;text-overflow:ellipsis">{_st}</td>'
            f'<td style="padding:3px 6px;font-size:calc(11px*var(--zf,1));color:#64748b;white-space:nowrap">{_sd}→{_sf}</td>'
            f'<td style="padding:3px 6px;font-size:calc(11px*var(--zf,1));font-weight:800;color:#1e293b;white-space:nowrap">{_sdur}</td>'
            f'<td style="padding:3px 6px;font-size:calc(10px*var(--zf,1));color:#94a3b8;max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{_scmt}">{_scmt}</td>'
            f'</tr>')
    if not _stop_list_html:
        _stop_list_html = '<tr><td colspan="4" style="padding:6px;text-align:center;color:#94a3b8;font-size:calc(12px*var(--zf,1))">Aucun arrêt</td></tr>'

    # ── Productions table (avec données détaillées pour popup) ──
    import json as _json2
    _dash_prod_list = []
    prod_rows_html = ""
    for _pi, r in enumerate(list(reversed(in_shift_prod))[:8]):
        _trs_p = -1.0
        trs_val = ""
        try:
            tv = float(str(r[24] or "").replace(",","."))
            _trs_p = tv
            tc = trs_color(tv)
            trs_val = f'<span style="color:{tc};font-weight:900">{tv:.1f}%</span>'
        except: pass
        kit_val = str(r[15] or "").strip()
        if kit_val.lower() in ("oui","yes","1","true","x"): kit_disp = '<span style="color:#16a34a;font-weight:800">Oui</span>'
        elif kit_val.lower() in ("non","no","0","false",""): kit_disp = '<span style="color:#94a3b8">Non</span>'
        else: kit_disp = kit_val
        cad_val = str(r[22] or "").strip()
        cmt_val = str(r[35] or "").strip()
        fibre_val = str(r[11] or "").strip()
        fibre_short = fibre_val[:9]+('…' if len(fibre_val)>9 else '')
        # Arrêts pendant cet OF
        _of_date = str(r[2] or "")[:10]
        _of_pilot = str(r[4] or "")
        _of_deb_s2 = hms2s(r[16]); _of_fin_s2 = hms2s(r[17]) or 86400
        _of_stops2 = [{"type":str(_er[0] or ""),"debut":str(_er[16] or "")[:5],"fin":str(_er[17] or "")[:5],"duree":str(_er[18] or ""),"comment":str(_er[35] or "")}
                      for _er in evt_rows_all if str(_er[2] or "")[:10]==_of_date and str(_er[4] or "")==_of_pilot and _of_deb_s2<=hms2s(_er[16])<=_of_fin_s2]
        _dash_prod_list.append({"of":str(r[1] or ""),"date":str(r[2] or "")[:10],"pilot":str(r[4] or ""),"poste":str(r[3] or ""),
            "copilote":str(r[5] or ""),"nb_pers":str(r[6] or ""),"taille":str(r[7] or ""),"code_prod":str(r[8] or ""),
            "type_prod":str(r[9] or ""),"poids":str(r[10] or ""),"fibre":str(r[11] or ""),"of_taie":str(r[12] or ""),
            "traca":str(r[13] or ""),"ref_taie":str(r[14] or ""),"kit":str(r[15] or ""),
            "debut":str(r[16] or "")[:5],"fin":str(r[17] or "")[:5],"duree":str(r[18] or ""),
            "qte_fab":str(r[19] or ""),"qte_emb":str(r[20] or ""),"equiv":str(r[21] or ""),
            "qte_init_taie":str(r[25] if len(r)>25 else ""),"nb_taie2_choix":str(r[26] if len(r)>26 else ""),
            "nb_def_cout":str(r[27] if len(r)>27 else ""),"mq_taie":str(r[28] if len(r)>28 else ""),
            "mq_housse_encart":str(r[29] if len(r)>29 else ""),"nb_pp_cousue":str(r[30] if len(r)>30 else ""),
            "duree_mq_mp":str(r[32] if len(r)>32 else ""),"manquant_pers":str(r[33] if len(r)>33 else ""),
            "comment":cmt_val,"trs":_trs_p,"stops":_of_stops2})
        prod_rows_html += f'''<tr style="cursor:pointer" onclick="showDashProdOf({_pi})" title="Voir détail OF">
          <td style="font-weight:800;font-size:calc(15px*var(--zf,1));color:#1e3a8a;text-decoration:underline">{r[1] or ""}</td>
          <td style="color:#6366f1;font-weight:700;font-size:calc(13px*var(--zf,1));cursor:pointer" title="{fibre_val}" onclick="event.stopPropagation();if(this.title)alert(\'Fibre : \'+this.title)">{fibre_short}</td>
          <td style="color:#475569;font-size:calc(14px*var(--zf,1))">{r[9] or ""}</td>
          <td style="color:#475569;font-size:calc(14px*var(--zf,1))">{r[7] or ""}</td>
          <td style="text-align:center">{kit_disp}</td>
          <td style="color:#475569">{str(r[16] or "")[:5]}</td><td style="color:#475569">{str(r[17] or "")[:5]}</td>
          <td style="color:#475569">{r[18] or ""}</td>
          <td style="color:#1e293b">{r[19] or "0"}</td>
          <td style="color:#0369a1;font-size:calc(14px*var(--zf,1))">{cad_val}</td>
          <td style="font-weight:800;color:#0891b2">{r[21] or ""}</td>
          <td>{trs_val}</td>
          <td style="color:#64748b;font-size:calc(14px*var(--zf,1));max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{cmt_val}">{cmt_val}</td>
        </tr>'''
    if not prod_rows_html:
        prod_rows_html = '<tr><td colspan="12" style="color:#94a3b8;padding:10px;text-align:center;font-size:calc(15px*var(--zf,1))">Aucune production déclarée</td></tr>'
    _dash_prod_json = _json2.dumps(_dash_prod_list, ensure_ascii=True, default=str)

    # ── ALERT BANNER HTML ──
    alert_html = ""
    if has_alert:
        stops_html = "".join(
            f'<div style="display:flex;align-items:center;justify-content:space-between;gap:16px;background:rgba(255,255,255,.12);border-radius:8px;padding:6px 14px;margin:3px 0;min-width:240px">'
            f'<span style="font-size:calc(22px*var(--zf,1));font-weight:800;color:#fef2f2">{nm}</span>'
            f'<span style="font-size:calc(30px*var(--zf,1));font-weight:900;color:#fecaca;font-variant-numeric:tabular-nums">{int(el/60)}<span style="font-size:calc(16px*var(--zf,1))">min</span></span>'
            f'</div>'
            for nm, el in all_stops_info
        )
        alert_html = f'''
<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;
  background:linear-gradient(135deg,#7f0000 0%,#b91c1c 50%,#ef4444 100%);
  animation:pulse 1.2s ease-in-out infinite;border-bottom:6px solid #fca5a5">
  <div style="font-size:calc(52px*var(--zf,1));line-height:1;animation:wag .8s ease-in-out infinite">🚨</div>
  <div style="font-size:calc(50px*var(--zf,1));font-weight:900;letter-spacing:4px;margin:8px 0;text-shadow:0 4px 16px rgba(0,0,0,.4);color:#fff">ARRÊT{"S" if len(all_stops_info)>1 else ""} EN COURS</div>
  {stops_html}
  <div style="font-size:calc(52px*var(--zf,1));line-height:1;animation:wag .8s ease-in-out infinite reverse;margin-top:8px">🚨</div>
</div>'''
    else:
        # ── PROD EN COURS + PRODUCTIONS DU POSTE ──
        prod_status_col = "#16a34a" if prod_active else "#64748b"
        prod_status_label = "▶ PRODUCTION EN COURS" if prod_active else "○ EN ATTENTE"
        of_debut_str = of_start_dt.strftime("%H:%M") if of_start_dt else "—"
        of_elapsed_str = f"{int(of_elapsed_s//3600):02d}h{int((of_elapsed_s%3600)//60):02d}" if of_elapsed_s > 0 else "—"
        of_stop_min = f"{of_stop_s/60:.0f}" if of_stop_s > 0 else "0"
        kit_col = "#16a34a" if kit_now == "Oui" else "#94a3b8"
        of_border_col = "#22c55e" if prod_active else "#e2e8f0"
        # Chips ligne 1 : timing courant
        _of_chips = ""
        for _lbl, _val, _col in [("Début", of_debut_str, "#1e293b"),
                                   ("Écoulé", of_elapsed_str, "#0891b2"),
                                   ("Arrêts", of_stop_min+" min", "#ef4444")]:
            _of_chips += (f'<div style="text-align:center;flex-shrink:0">'
                          f'<div style="font-size:calc(10px*var(--zf,1));color:#64748b;font-weight:700;text-transform:uppercase">{_lbl}</div>'
                          f'<div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:{_col}">{_val}</div></div>')
        # Chips ligne 2 : champs formulaire (uniquement si renseigné)
        _form_detail_chips = ""
        _form_detail_fields = [
            ("Taille", taille_now, "#1e293b"), ("Type", type_prod_now, "#1e293b"),
            ("Code prod.", code_prod_now, "#374151"), ("Lots de 2", kit_now if prod_active else "", kit_col),
            ("Fibre", fibre_now, "#6366f1"), ("Poids (g)", poids_now, "#1e293b"),
            ("Co-pilote", copilote_now, "#1e293b"), ("Nb pers.", nb_pers_now, "#1e293b"),
            ("OF Taie", of_taie_now, "#374151"), ("Traca", traca_now, "#374151"),
            ("Réf Taie", ref_taie_now, "#374151"),
            ("Qté Fab.", qte_fab_now, "#1e293b"), ("Qté Emb.", qte_emb_now, "#1e293b"),
        ]
        for _lbl2, _val2, _col2 in _form_detail_fields:
            if not _val2 or _val2 == "Non": continue
            _form_detail_chips += (f'<div style="text-align:center;flex-shrink:0;padding:2px 5px;background:#f0fdf4;border:1px solid #d1fae5;border-radius:5px">'
                                   f'<div style="font-size:calc(9px*var(--zf,1));color:#64748b;font-weight:700;text-transform:uppercase">{_lbl2}</div>'
                                   f'<div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:{_col2};white-space:nowrap">{_val2}</div></div>')
        _form_section = (f'<div style="border-top:1px solid #d1fae5;padding-top:4px;display:flex;flex-wrap:wrap;gap:4px">{_form_detail_chips}</div>'
                         if _form_detail_chips else '')
        _poste_chips = ""
        for _i, (_lbl, _val, _col) in enumerate([("Éq.", f"{tot_equiv:.1f}", "#0891b2"),
                                                   ("Qté", str(nb_pieces), "#7c3aed"),
                                                   ("OF", str(nb_of_today), "#1e3a8a"),
                                                   ("Prod", f"{prod_s_total/3600:.1f}h", "#16a34a"),
                                                   ("Arrêts", f"{stop_s_total/60:.0f}min", "#ef4444")]):
            _bl = "border-left:1px solid #e2e8f0;" if _i > 0 else ""
            _poste_chips += (f'<div style="text-align:center;flex:1;padding:0 3px;{_bl}">'
                             f'<div style="font-size:calc(20px*var(--zf,1));font-weight:900;color:{_col};line-height:1">{_val}</div>'
                             f'<div style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#64748b;text-transform:uppercase;margin-top:1px">{_lbl}</div></div>')
        _budget_bars = _dash_budget_bars_html()
        _budget_col = f'  {_budget_bars}' if _budget_bars else ''
        _grid_cols = '2fr 1fr auto' if _budget_bars else '2fr 1fr'
        alert_html = f'''
<div style="display:grid;grid-template-columns:{_grid_cols};gap:6px;flex-shrink:0;align-items:stretch">
  <div style="background:#f0fdf4;border:2px solid {of_border_col};border-radius:8px;padding:5px 12px;display:flex;flex-direction:column;gap:4px">
    <div style="display:flex;align-items:center;gap:12px">
      <div style="flex-shrink:0">
        <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:{prod_status_col};letter-spacing:1px">{prod_status_label}</div>
        <div style="font-size:calc(24px*var(--zf,1));font-weight:900;color:#1e293b;line-height:1.1;font-family:monospace">OF {of_num_now}</div>
      </div>
      <div style="height:32px;width:1px;background:#d1fae5;flex-shrink:0"></div>
      {_of_chips}
    </div>
    {_form_section}
  </div>
  <div style="background:#f8fafc;border:2px solid #e2e8f0;border-radius:8px;padding:3px 8px;display:flex;align-items:center;gap:0">
    <div style="font-size:calc(10px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#1e3a8a;writing-mode:vertical-rl;transform:rotate(180deg);letter-spacing:1px;flex-shrink:0;margin-right:6px">Poste entier</div>
    {_poste_chips}
  </div>
{_budget_col}
</div>'''

    tl_svg = timeline_svg()

    # ── Historique: all prods sorted by date+time desc (last 30) ──
    hist_rows = sorted(prod_rows_all, key=lambda r: (str(r[2] or ''), str(r[16] or '')), reverse=True)[:30]
    import json as _json
    _dash_of_list = []
    hist_html = ""
    _prev_session_key = None
    for _hidx, r in enumerate(hist_rows):
        tc_hist = ""
        try:
            tv_h = float(str(r[24] or "").replace(",","."))
            tc_hist = f'<span style="color:{trs_color(tv_h)};font-weight:900">{tv_h:.1f}%</span>'
        except: pass
        # Compute stops for this OF
        _of_date = str(r[2] or "")[:10]
        _of_pilot = str(r[4] or "")
        _of_deb_s = hms2s(r[16])
        _of_fin_s = hms2s(r[17]) or 86400
        _of_stops = []
        for _er in evt_rows_all:
            if str(_er[2] or "")[:10] != _of_date: continue
            if str(_er[4] or "") != _of_pilot: continue
            _er_s = hms2s(_er[16])
            if _of_deb_s <= _er_s <= _of_fin_s:
                _of_stops.append({"type":str(_er[0] or ""),"debut":str(_er[16] or "")[:5],"fin":str(_er[17] or "")[:5],"duree":str(_er[18] or ""),"comment":str(_er[35] or "")})
        _trs_of = -1.0
        try:
            _tv_h2 = float(str(r[24] or "").replace(",","."))
            _trs_of = _tv_h2
        except: pass
        _dash_of_list.append({
            "of":str(r[1] or ""),
            "date":_of_date,
            "pilot":_of_pilot,
            "poste":str(r[3] or ""),
            "copilote":str(r[5] or ""),
            "nb_pers":str(r[6] or ""),
            "taille":str(r[7] or ""),
            "code_prod":str(r[8] or ""),
            "type_prod":str(r[9] or ""),
            "kit":str(r[15] or ""),
            "debut":str(r[16] or "")[:5],
            "fin":str(r[17] or "")[:5],
            "duree":str(r[18] or ""),
            "qte_fab":str(r[19] or ""),
            "qte_emb":str(r[20] or ""),
            "equiv":str(r[21] or ""),
            "poids":str(r[10] or ""),
            "fibre":str(r[11] or ""),
            "of_taie":str(r[12] or ""),
            "traca":str(r[13] or ""),
            "ref_taie":str(r[14] or ""),
            "qte_init_taie":str(r[22] or ""),
            "nb_taie2_choix":str(r[23] or ""),
            "trs":_trs_of,
            "duree_mq_mp":str(r[26] or "") if len(r)>26 else "",
            "manquant_pers":str(r[27] or "") if len(r)>27 else "",
            "nb_def_cout":str(r[28] or "") if len(r)>28 else "",
            "mq_taie":str(r[29] or "") if len(r)>29 else "",
            "mq_housse_encart":str(r[30] or "") if len(r)>30 else "",
            "nb_pp_cousue":str(r[31] or "") if len(r)>31 else "",
            "comment":str(r[35] or "") if len(r)>35 else "",
            "stops":_of_stops
        })
        _sess_key = f"{str(r[2] or '')[:10]}|{str(r[3] or '')}|{str(r[4] or '')}"
        if _sess_key != _prev_session_key:
            _prev_session_key = _sess_key
            hist_html += (f'<tr class="hist-sep" style="background:#f0f4fa;border-top:2px solid #c7d2e8">'
                f'<td colspan="12" style="padding:3px 10px;font-size:calc(10px*var(--zf,1));font-weight:600;color:#334155;letter-spacing:.2px">'
                f'📅 {str(r[2] or "")[:10]} &nbsp;·&nbsp; 🏭 {str(r[3] or "")} &nbsp;·&nbsp; 👤 {str(r[4] or "")}'
                f'</td></tr>')
        _fibre_h = str(r[11] or "").strip()
        _fibre_short_h = _fibre_h[:9] + ('…' if len(_fibre_h) > 9 else '')
        hist_html += (f'<tr class="hist-row" data-date="{str(r[2] or "")[:10]}" data-of="{str(r[1] or "")}" data-pilot="{str(r[4] or "")}" style="cursor:pointer" onclick="showDashOf({_hidx})" title="Voir détail OF">'
            f'<td style="font-weight:800">{str(r[2] or "")[:10]}</td>'
            f'<td style="font-weight:800;color:#1e3a8a;text-decoration:underline">{r[1] or ""}</td>'
            f'<td style="color:#6366f1;font-weight:700;font-size:calc(12px*var(--zf,1));cursor:pointer" title="{_fibre_h}" onclick="event.stopPropagation();if(this.title)alert(\'Fibre : \'+this.title)">{_fibre_short_h}</td>'
            f'<td>{r[9] or ""}</td>'
            f'<td>{str(r[3] or "")}</td>'
            f'<td>{str(r[4] or "")}</td>'
            f'<td>{str(r[16] or "")[:5]}</td><td>{str(r[17] or "")[:5]}</td>'
            f'<td>{r[18] or ""}</td>'
            f'<td>{r[19] or "0"}</td>'
            f'<td style="color:#0891b2;font-weight:800">{r[21] or ""}</td>'
            f'<td>{tc_hist}</td>'
            f'</tr>')
    if not hist_html:
        hist_html = '<tr><td colspan="12" style="text-align:center;color:#94a3b8;padding:12px">Aucune production</td></tr>'
    _dash_of_json = _json.dumps(_dash_of_list, ensure_ascii=True, default=str)

    # ── Rapports: groupé par type_prod ──
    from collections import defaultdict as _dd2
    _rpt = _dd2(lambda: {'count':0,'equiv':0.0,'qty':0,'trs_sum':0.0,'trs_cnt':0})
    for r in prod_rows_all:
        _tp = str(r[9] or '').strip() or '(sans type)'
        _rpt[_tp]['count'] += 1
        try: _rpt[_tp]['equiv'] += float(str(r[21] or '0').replace(',','.') or 0)
        except: pass
        try: _rpt[_tp]['qty'] += int(str(r[19] or '0').split('.')[0] or 0)
        except: pass
        try:
            _tv = float(str(r[24] or '').replace(',','.'))
            if _tv >= 0: _rpt[_tp]['trs_sum'] += _tv; _rpt[_tp]['trs_cnt'] += 1
        except: pass
    rpt_html = ""
    for _tpk, _dr in sorted(_rpt.items(), key=lambda x: -x[1]['equiv']):
        _atrs = _dr['trs_sum']/_dr['trs_cnt'] if _dr['trs_cnt'] else -1
        _tcrpt = (f'<span style="color:{trs_color(_atrs)};font-weight:900">{_atrs:.1f}%</span>'
                  if _atrs >= 0 else '<span style="color:#94a3b8">—</span>')
        rpt_html += (f'<tr>'
            f'<td style="font-weight:800;text-align:left;padding:6px 10px">{_tpk}</td>'
            f'<td style="text-align:center">{_dr["count"]}</td>'
            f'<td style="text-align:center;color:#0891b2;font-weight:800">{_dr["equiv"]:.1f}</td>'
            f'<td style="text-align:center">{_dr["qty"]}</td>'
            f'<td style="text-align:center">{_tcrpt}</td>'
            f'</tr>')
    if not rpt_html:
        rpt_html = '<tr><td colspan="5" style="text-align:center;color:#94a3b8;padding:12px">Aucune donnée</td></tr>'

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard Encadrant — ORC</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:100%;overflow:hidden;font-family:-apple-system,'Segoe UI',Arial,sans-serif;background:#eef2f7;color:#1e293b;font-size:calc(15px*var(--zf,1))}}
.hdr{{height:48px;background:linear-gradient(135deg,#1e3a8a 0%,#1e40af 100%);color:#fff;display:flex;align-items:center;justify-content:space-between;padding:0 18px;flex-shrink:0;box-shadow:0 2px 8px rgba(30,58,138,.3)}}
.hdr-title{{font-size:calc(18px*var(--zf,1));font-weight:900;display:flex;align-items:center;gap:10px;letter-spacing:.3px}}
.hdr-badge{{background:rgba(255,255,255,.15);border-radius:20px;padding:4px 12px;font-size:calc(13px*var(--zf,1));font-weight:700}}
.hdr-badge.green{{background:#15803d;box-shadow:0 0 0 2px #22c55e44}}
.hdr-badge.gray{{background:rgba(255,255,255,.15)}}
.hdr-time{{font-size:calc(12px*var(--zf,1));opacity:.75}}
.outer{{height:calc(100vh - 48px);display:flex;flex-direction:column;gap:8px;padding:8px;overflow:hidden}}
.panel{{background:#fff;border-radius:10px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
.panel-hdr{{padding:6px 12px;font-size:calc(12px*var(--zf,1));font-weight:900;text-transform:uppercase;letter-spacing:1px;flex-shrink:0}}
.panel-body{{flex:1;overflow-y:auto;padding:8px 12px;min-height:0}}
.trs-num{{font-size:calc(42px*var(--zf,1));font-weight:900;line-height:1;text-align:center}}
.trs-lbl{{font-size:calc(12px*var(--zf,1));font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#64748b;text-align:center;margin-bottom:4px}}
.trs-sub{{font-size:calc(14px*var(--zf,1));color:#64748b;text-align:center;margin-top:3px}}
/* HERO ROW */
.dash-hero{{display:flex;gap:10px;align-items:stretch;flex-shrink:0}}
.dash-trs-card{{background:#fff;border-radius:12px;padding:12px 16px;display:flex;flex-direction:column;align-items:center;justify-content:center;box-shadow:0 1px 3px rgba(0,0,0,.1);min-width:180px}}
.dash-kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;flex:1}}
.dash-kpi{{background:#fff;border-radius:10px;padding:10px 8px;display:flex;flex-direction:column;align-items:center;justify-content:center;box-shadow:0 1px 3px rgba(0,0,0,.1);text-align:center}}
.dash-kpi-val{{font-size:calc(24px*var(--zf,1));font-weight:900;line-height:1.1}}
.dash-kpi-lbl{{font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;margin-top:4px;letter-spacing:.4px}}
.dash-pie-card{{background:#fff;border-radius:12px;padding:10px 14px;display:flex;align-items:center;gap:12px;box-shadow:0 1px 3px rgba(0,0,0,.1);flex-shrink:0}}
/* CONTENT GRID */
.dash-content{{display:grid;grid-template-columns:1fr 320px;gap:8px;flex:1;overflow:hidden;min-height:0}}
.dash-right{{display:flex;flex-direction:column;gap:8px;overflow:hidden}}
.dash-card{{background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.08);display:flex;flex-direction:column}}
.dash-card-hdr{{padding:7px 14px;font-size:calc(11px*var(--zf,1));font-weight:900;text-transform:uppercase;letter-spacing:.7px;color:#475569;border-bottom:1px solid #f1f5f9;flex-shrink:0;display:flex;align-items:center;gap:6px}}
.dash-card-hdr .dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.dash-card-body{{flex:1;overflow-y:auto;min-height:0}}
/* STAT CARDS (alert mode) */
.stat-grid{{display:grid;grid-template-columns:1fr 1fr;gap:6px}}
.stat-card{{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;text-align:center;box-shadow:0 1px 2px rgba(0,0,0,.05)}}
.stat-val{{font-size:calc(28px*var(--zf,1));font-weight:900;line-height:1}}
.stat-lbl{{font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;margin-top:3px;letter-spacing:.4px}}
/* TIMELINE */
.tl-cell{{background:#fff;border-radius:12px;padding:8px 12px;flex-shrink:0;box-shadow:0 1px 3px rgba(0,0,0,.08)}}
.tl-lbl{{font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:#64748b;margin-bottom:5px;display:flex;justify-content:space-between}}
.tl-legend{{display:flex;gap:10px;font-size:calc(12px*var(--zf,1));color:#64748b;margin-top:4px;flex-wrap:wrap;align-items:center}}
.tl-legend span i{{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:2px;vertical-align:middle}}
/* FP CARDS (used in _render_rpt_panel) */
.fp-card{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;text-align:center}}
.fp-big{{font-size:calc(20px*var(--zf,1));font-weight:900;color:#1e3a8a;line-height:1.1}}
.fp-lbl{{font-size:calc(9px*var(--zf,1));text-transform:uppercase;font-weight:700;color:#64748b;margin-top:2px}}
.rpt-card{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:10px}}
/* TABLE */
.ktbl{{width:100%;border-collapse:collapse;font-size:calc(14px*var(--zf,1))}}
.ktbl th{{background:#f8fafc;padding:6px 10px;font-weight:800;text-align:center;position:sticky;top:0;font-size:calc(11px*var(--zf,1));text-transform:uppercase;color:#475569;border-bottom:2px solid #e2e8f0;white-space:nowrap}}
.ktbl td{{padding:6px 8px;border-bottom:1px solid #f1f5f9;text-align:center;color:#1e293b}}
.ktbl tr:hover td{{background:#f8fafc}}
/* HISTORIQUE FILTER */
.hist-filter{{display:flex;gap:8px;align-items:center;padding:8px 10px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);flex-shrink:0;flex-wrap:wrap}}
.hist-filter label{{font-size:calc(11px*var(--zf,1));font-weight:700;color:#64748b;white-space:nowrap}}
.hist-filter input{{border:1px solid #e2e8f0;border-radius:6px;padding:5px 10px;font-size:calc(12px*var(--zf,1));color:#1e293b;outline:none;background:#f8fafc}}
.hist-filter input:focus{{border-color:#3b82f6;background:#fff}}
/* TABS */
.tab-bar{{display:flex;gap:0;flex-shrink:0;border-bottom:2px solid #e2e8f0}}
.tab-btn{{background:none;border:none;border-bottom:3px solid transparent;padding:9px 22px;font-size:calc(13px*var(--zf,1));font-weight:700;color:#64748b;cursor:pointer;transition:all .15s;margin-bottom:-2px}}
.tab-btn:hover{{color:#1e3a8a}}
.tab-btn.active{{color:#1e3a8a;border-bottom-color:#1e3a8a;background:rgba(30,58,138,.04)}}
.tab-pane{{flex:1;display:flex;flex-direction:column;gap:8px;overflow:hidden;min-height:0}}
/* RAPPORTS SIDEBAR */
.rpt-wrap{{display:flex;flex:1;overflow:hidden;min-height:0;position:relative;gap:0}}
#rpt-sidebar{{width:260px;min-width:0;transition:width .25s ease,opacity .2s ease;overflow:hidden;flex-shrink:0;display:flex;flex-direction:column;border-right:1px solid #e2e8f0;background:#fafbfc}}
#rpt-sidebar.col{{width:0;opacity:0;border-right:none}}
#rpt-tog{{position:absolute;left:260px;top:50%;transform:translateY(-50%);z-index:20;width:18px;height:40px;background:#fff;border:1px solid #e2e8f0;border-left:none;border-radius:0 6px 6px 0;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#64748b;transition:left .25s ease;box-shadow:2px 0 4px rgba(0,0,0,.06);font-size:calc(13px*var(--zf,1))}}
#rpt-sidebar.col+#rpt-tog{{left:0}}
/* MODAL */
.dash-modal-overlay{{position:fixed;inset:0;background:rgba(0,0,0,.5);display:none;align-items:center;justify-content:center;z-index:9999}}
.dash-modal-box{{background:#fff;border-radius:16px;padding:24px;max-width:680px;width:95%;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.25)}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.85}}}}
@keyframes wag{{0%{{transform:rotate(-8deg)}}50%{{transform:rotate(8deg)}}100%{{transform:rotate(-8deg)}}}}
::-webkit-scrollbar{{width:5px}}::-webkit-scrollbar-track{{background:#f1f5f9}}::-webkit-scrollbar-thumb{{background:#cbd5e1;border-radius:3px}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-title">
    &#127981; Dashboard Encadrant — ORC
  </div>
  <div style="display:flex;align-items:center;gap:18px">
    {f'<div style="display:flex;flex-direction:column;align-items:center;line-height:1.1"><span style="font-size:calc(22px*var(--zf,1));font-weight:900;color:#fff;letter-spacing:.5px">👤 {pilot_now}</span><span style="font-size:calc(14px*var(--zf,1));font-weight:700;color:#93c5fd;text-transform:uppercase">{poste_now}</span></div>' if pilot_now else ''}
    {f'<div style="background:rgba(255,255,255,.12);border-radius:8px;padding:4px 12px;text-align:center"><div style="font-size:calc(12px*var(--zf,1));color:#93c5fd;font-weight:700;text-transform:uppercase">Modèle horaire</div><div style="font-size:calc(18px*var(--zf,1));font-weight:900;color:#fff">{model_debut_dt.strftime("%H:%M")} → {model_fin_dt.strftime("%H:%M")}</div></div>' if (model_debut_dt and model_fin_dt) else (f'<div style="background:rgba(255,255,255,.12);border-radius:8px;padding:4px 12px"><div style="font-size:calc(12px*var(--zf,1));color:#93c5fd;font-weight:700">Modèle</div><div style="font-size:calc(18px*var(--zf,1));font-weight:900;color:#fff">{model_debut_dt.strftime("%H:%M")} →</div></div>' if model_debut_dt else '')}
    {'<span class="hdr-badge green">▶ PROD — OF ' + of_num_now + '</span>' if prod_active else '<span class="hdr-badge gray">○ En attente</span>'}
    <span class="hdr-time">🔄 15s | {gen_time}</span>
  </div>
</div>

<div class="outer">

  <!-- ONGLETS -->
  <div class="tab-bar">
    <button class="tab-btn active" id="tb-accueil" onclick="showTab('accueil')">🏠 Accueil</button>
    <button class="tab-btn" id="tb-historique" onclick="showTab('historique')">📋 Historique</button>
    <button class="tab-btn" id="tb-rapports" onclick="showTab('rapports')">📊 Rapports postes</button>
    <button class="tab-btn" id="tb-rpt-jour" onclick="showTab('rpt-jour')">📅 Rapports jour</button>
  </div>

  <!-- ONGLET ACCUEIL -->
  <div id="tab-accueil" class="tab-pane">

  {alert_html}

  <!-- HEADER BAR like EXE Reports -->
  <div style="background:#1e3a8a;color:#fff;padding:10px 16px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;border-radius:10px;margin-bottom:8px">
    <div>
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800">📋 Tableau de bord — {poste_now or "—"}</div>
      <div style="font-size:calc(11px*var(--zf,1));opacity:.8">{pilot_now or "—"} · Aujourd'hui · {elapsed_str}</div>
    </div>
    <div style="text-align:right">
      <div style="font-size:calc(28px*var(--zf,1));font-weight:900;color:{trs_col}">{f"{trs_poste:.1f}%" if trs_poste>=0 else "—"}</div>
      <div style="font-size:calc(11px*var(--zf,1));opacity:.7">TRS Poste</div>
    </div>
  </div>

  <!-- KPI BAR: gauge + pie + fp-cards -->
  <div style="display:flex;gap:10px;padding:10px 12px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:8px;align-items:center;flex-wrap:wrap;flex-shrink:0">
    <div style="text-align:center;flex-shrink:0">
      <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#94a3b8;margin-bottom:4px">TRS Poste</div>
      {gauge_svg(trs_poste, 130)}
    </div>
    <div style="text-align:center;flex-shrink:0">
      <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#94a3b8;margin-bottom:4px">Répartition</div>
      {pie_svg(prod_s_total, stop_s_total, 65)}
    </div>
    <div style="flex:1;display:flex;flex-direction:column;gap:5px;min-width:300px">
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:5px">
        <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(24px*var(--zf,1));color:#7c3aed;font-weight:900">{nb_of_today}</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">OF déclarés</div></div>
        <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(24px*var(--zf,1));color:#0891b2;font-weight:900">{tot_equiv:.1f}</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Équivalence</div></div>
        <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(24px*var(--zf,1));color:{trs_col};font-weight:900">{f"{trs_poste:.1f}%" if trs_poste>=0 else "—"}</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">TRS Poste</div></div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px">
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1))">{(model_debut_dt.strftime("%H:%M")+"→"+last_fin_dt.strftime("%H:%M")) if (model_debut_dt and last_fin_dt) else "—"}</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Plage</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#16a34a">{prod_s_total/60:.0f} min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Prod.</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#dc2626">{stop_s_total/60:.0f} min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts</div></div>
        <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#16a34a">{prod_pct}%</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">% prod</div></div>
      </div>
    </div>
  </div>

  <!-- TIMELINE -->
  <div style="padding:6px 12px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:8px;flex-shrink:0">
    <div class="tl-lbl">
      <span>Timeline — {poste_now or "en cours"}</span>
      <span style="font-size:calc(12px*var(--zf,1))">{(model_debut_dt or shift_start_dt).strftime("%H:%M") if (model_debut_dt or shift_start_dt) else "—"} → maintenant</span>
    </div>
    {tl_svg}
    <div class="tl-legend">
      <span><i style="background:#22c55e"></i>Production</span>
      <span><i style="background:#4ade80"></i>OF en cours</span>
      <span><i style="background:#ef4444"></i>PB Technique</span>
      <span><i style="background:#f59e0b"></i>Rattrapage</span>
      <span><i style="background:#f97316"></i>Nettoyage</span>
      <span><i style="background:#64748b"></i>Pause</span>
      <span><i style="background:#8b5cf6"></i>Réunion</span>
      <span><i style="background:repeating-linear-gradient(45deg,#16a34a,#16a34a 4px,#fef08a 4px,#fef08a 8px)"></i>Mode dégradé</span>
    </div>
  </div>

  <!-- BODY: Productions + Pareto/Arrêts side by side -->
  <div style="display:grid;grid-template-columns:1fr 340px;gap:8px;flex:1;min-height:0;overflow:hidden">
    <!-- Productions -->
    <div style="background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);display:flex;flex-direction:column;overflow:hidden">
      <div style="padding:8px 12px;font-size:calc(12px*var(--zf,1));font-weight:800;color:#1e3a8a;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;gap:6px;flex-shrink:0">
        <span style="width:8px;height:8px;background:#16a34a;border-radius:50%;display:inline-block"></span>Productions déclarées
      </div>
      <div style="flex:1;overflow-y:auto">
        <table class="ktbl" style="font-size:calc(12px*var(--zf,1))">
          <thead><tr><th>OF</th><th>Fibre</th><th>Type</th><th>Format</th><th>Lots 2</th><th>Début</th><th>Fin</th><th>Durée</th><th>Qté</th><th>Cad./h</th><th>Éq.</th><th>TRS</th><th>Comm.</th></tr></thead>
          <tbody>{prod_rows_html}</tbody>
        </table>
      </div>
    </div>
    <!-- Pareto + Arrêts -->
    <div style="display:flex;flex-direction:column;gap:8px;overflow:hidden">
      <div style="background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:10px 12px;flex-shrink:0">
        <div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:#1e3a8a;margin-bottom:8px;display:flex;align-items:center;gap:6px">
          <span style="width:8px;height:8px;background:#d97706;border-radius:50%;display:inline-block"></span>Pareto arrêts
        </div>
        {pareto_html}
      </div>
      <div style="background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);flex:1;overflow:hidden;display:flex;flex-direction:column">
        <div style="padding:8px 12px;font-size:calc(12px*var(--zf,1));font-weight:800;color:#1e3a8a;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;gap:6px;flex-shrink:0">
          <span style="width:8px;height:8px;background:#ef4444;border-radius:50%;display:inline-block"></span>Détail arrêts
        </div>
        <div style="flex:1;overflow-y:auto">
          <table style="width:100%;border-collapse:collapse"><tbody>{_stop_list_html}</tbody></table>
        </div>
      </div>
    </div>
  </div>

  </div><!-- /tab-accueil -->

  <!-- ONGLET HISTORIQUE -->
  <div id="tab-historique" class="tab-pane" style="display:none;flex-direction:column">

  <div class="hist-filter">
    <label>Du :</label>
    <input type="date" id="h-from" oninput="filterHist()">
    <label>Au :</label>
    <input type="date" id="h-to" oninput="filterHist()">
    <label>Recherche OF :</label>
    <input type="text" id="h-q" placeholder="Numéro d'OF..." oninput="filterHist()" style="width:160px">
    <button onclick="document.getElementById('h-from').value='';document.getElementById('h-to').value='';document.getElementById('h-q').value='';filterHist()" style="background:none;border:1px solid #e2e8f0;border-radius:6px;padding:4px 10px;font-size:calc(12px*var(--zf,1));color:#64748b;cursor:pointer">✕ Réinitialiser</button>
    <span id="h-count" style="font-size:calc(11px*var(--zf,1));color:#94a3b8;margin-left:auto">{len(hist_rows)} productions</span>
  </div>

  <div class="dash-card" style="flex:1;min-height:0">
    <div class="dash-card-body" style="padding:0">
      <table class="ktbl">
        <thead><tr>
          <th>Date</th><th>OF</th><th>Fibre</th><th>Type</th><th>Poste</th><th>Pilote</th>
          <th>Début</th><th>Fin</th><th>Durée</th><th>Qté</th><th>Éq.</th><th>TRS</th>
        </tr></thead>
        <tbody id="hist-tbody">{hist_html}</tbody>
      </table>
    </div>
  </div>

  </div><!-- /tab-historique -->

  <!-- ONGLET RAPPORTS -->
  <div id="tab-rapports" class="tab-pane" style="display:none;flex-direction:column;overflow:hidden">
    <div class="rpt-wrap">
      <div id="rpt-sidebar">
        <div style="padding:10px 14px;font-size:calc(13px*var(--zf,1));font-weight:800;color:#1e3a8a;border-bottom:1px solid #e2e8f0;flex-shrink:0;display:flex;align-items:center;justify-content:space-between">
          <span>📋 Rapports postes</span>
          <button onclick="loadRapports()" style="font-size:calc(11px*var(--zf,1));padding:3px 8px;background:none;border:1px solid #cbd5e1;border-radius:4px;cursor:pointer;color:#64748b">↺</button>
        </div>
        <div id="rpt-list" style="flex:1;overflow-y:auto">
          __RPT_LIST__
        </div>
      </div>
      <button id="rpt-tog" onclick="toggleRptSidebar()" title="Réduire/Agrandir">❮</button>
      <div id="rpt-detail" style="overflow-y:auto;flex:1;padding:0">
        __RPT_DET__
      </div>
    </div>
  </div><!-- /tab-rapports -->

  <!-- ONGLET RAPPORTS JOUR -->
  <div id="tab-rpt-jour" class="tab-pane" style="display:none;flex-direction:column;overflow:hidden;height:100%">
    <div class="hist-filter" style="flex-shrink:0">
      <label>Du :</label>
      <input type="date" id="rj-from" style="font-size:calc(12px*var(--zf,1))">
      <label>Au :</label>
      <input type="date" id="rj-to" style="font-size:calc(12px*var(--zf,1))">
      <label>Pilote :</label>
      <select id="rj-pilot" style="font-size:calc(12px*var(--zf,1));padding:4px 8px;border:1px solid #e2e8f0;border-radius:6px;background:#fff"><option value="">Tous</option></select>
      <label>Poste :</label>
      <select id="rj-poste" style="font-size:calc(12px*var(--zf,1));padding:4px 8px;border:1px solid #e2e8f0;border-radius:6px;background:#fff"><option value="">Tous</option></select>
      <button onclick="calcPeriodReport()" style="background:#1e3a8a;color:#fff;border:none;border-radius:6px;padding:5px 14px;font-size:calc(12px*var(--zf,1));font-weight:700;cursor:pointer">🔄 Actualiser</button>
      <button onclick="resetPeriodReport()" style="background:none;border:1px solid #e2e8f0;border-radius:6px;padding:4px 10px;font-size:calc(12px*var(--zf,1));color:#64748b;cursor:pointer">✕ Réinitialiser</button>
    </div>
    <div id="rj-auto-banner" style="display:none;padding:5px 16px;background:#eff6ff;border-bottom:1px solid #bfdbfe;flex-shrink:0"></div>
    <div id="rj-result" style="flex:1;overflow-y:auto;padding:12px 16px">
      <div style="padding:60px;text-align:center;color:#94a3b8">
        <div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">📅</div>
        <div style="font-size:calc(14px*var(--zf,1));font-weight:600">Sélectionnez une période puis cliquez sur Calculer</div>
      </div>
    </div>
  </div><!-- /tab-rpt-jour -->

</div><!-- /outer -->

<script>
var _dashOf={_dash_of_json};
var _dashProdOf={_dash_prod_json};
function _renderDashOf(r){{
  if(!r)return;
  var tc=r.trs>=90?'#16a34a':r.trs>=70?'#f59e0b':r.trs>=0?'#dc2626':'#94a3b8';
  var kit=(r.kit||'').toLowerCase()==='oui'?'<span style="color:#16a34a;font-weight:800">✓ Oui</span>':'Non';
  var chips=[
    ['Date',r.date],['Poste',r.poste],['Pilote',r.pilot||r.pilote||''],['Co-Pilote',r.copilote||''],
    ['Taille',r.taille||''],['Type',r.type_prod||''],['Lots de 2',kit,'raw'],
    ['Nb Pers.',r.nb_pers||''],['Code Produit',r.code_prod||''],
    ['Début',r.debut||''],['Fin',r.fin||''],['Durée',r.duree||''],
    ['Qté Fab.',r.qte_fab||''],['Qté Emb.',r.qte_emb||''],['Équivalence',r.equiv||''],
    ['Poids (g)',r.poids||''],['Fibre',r.fibre||''],['OF Taie',r.of_taie||''],
    ['Traca',(r.traca||'').split(';').filter(function(t){{return t.trim();}}).join(' · ')],['Réf Taie',r.ref_taie||''],
    ['Qté Init Taie',r.qte_init_taie||''],['Nb Taie 2nd',r.nb_taie2_choix||''],
    ['Nb déf. coût',r.nb_def_cout||''],['Mq taie',r.mq_taie||''],
    ['Mq housse',r.mq_housse_encart||''],['PP cousu emb.',r.nb_pp_cousue||''],
    ['Durée MQ MP',r.duree_mq_mp||''],
    ['TRS OF',r.trs>=0?r.trs.toFixed(1)+'%':''],
  ];
  chips=chips.filter(function(c){{return c[1]&&c[1]!=='—'&&c[1]!==''||c[0]==='TRS OF';}});
  var esc=function(s){{return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;');}};
  var chipsHtml=chips.map(function(c){{return '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;text-align:center"><div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b">'+c[0]+'</div><div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:#1e293b">'+(c[2]==='raw'?c[1]:esc(c[1]))+'</div></div>';}}).join('');
  var stopsHtml=(r.stops&&r.stops.length)?r.stops.map(function(e){{return '<tr><td style="padding:4px 8px;font-size:calc(12px*var(--zf,1));font-weight:600">'+esc(e.type||'')+'</td><td style="padding:4px 8px;font-size:calc(11px*var(--zf,1));white-space:nowrap">'+e.debut+'→'+e.fin+'</td><td style="padding:4px 8px;font-weight:700">'+e.duree+'</td><td style="padding:4px 8px;font-size:calc(11px*var(--zf,1));color:#64748b">'+esc(e.comment||'')+'</td></tr>';}}).join(''):'<tr><td colspan="4" style="padding:8px;text-align:center;color:#94a3b8">Aucun arrêt</td></tr>';
  document.getElementById('dash-of-detail-content').innerHTML='<div style="font-size:calc(22px*var(--zf,1));font-weight:900;color:#1e3a8a;margin-bottom:14px;font-family:monospace">OF '+esc(r.of||'—')+'</div><div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:14px">'+chipsHtml+'</div>'+(r.comment?'<div style="background:#fffbeb;border:1px solid #fef08a;border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:calc(13px*var(--zf,1))">💬 '+esc(r.comment)+'</div>':'')+'<div style="font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;margin-bottom:6px">Arrêts pendant cet OF</div><table style="width:100%;border-collapse:collapse"><thead><tr style="background:#f1f5f9"><th style="padding:4px 8px;text-align:left;font-size:calc(11px*var(--zf,1))">Type</th><th style="padding:4px 8px;font-size:calc(11px*var(--zf,1))">Plage</th><th style="padding:4px 8px;font-size:calc(11px*var(--zf,1))">Durée</th><th style="padding:4px 8px;font-size:calc(11px*var(--zf,1))">Comm.</th></tr></thead><tbody>'+stopsHtml+'</tbody></table>';
  document.getElementById('dash-of-modal').style.display='flex';
}}
function showDashProdOf(i){{ _renderDashOf(_dashProdOf[i]); }}
function showDashOf(i){{ _renderDashOf(_dashOf[i]); }}
var _rptSideCol=false;
function toggleRptSidebar(){{
  var sb=document.getElementById('rpt-sidebar');
  var tog=document.getElementById('rpt-tog');
  if(!sb||!tog) return;
  _rptSideCol=!_rptSideCol;
  sb.classList.toggle('col',_rptSideCol);
  tog.textContent=_rptSideCol?'❯':'❮';
}}
function filterHist(){{
  var from=document.getElementById('h-from').value;
  var to=document.getElementById('h-to').value;
  var q=(document.getElementById('h-q').value||'').trim().toLowerCase();
  var rows=document.querySelectorAll('#hist-tbody .hist-row');
  var vis=0;
  rows.forEach(function(tr){{
    var d=(tr.dataset.date||'');
    var iso=d.split('/').reverse().join('-');
    var of=(tr.dataset.of||'').toLowerCase();
    var ok=true;
    if(from&&iso<from) ok=false;
    if(to&&iso>to) ok=false;
    if(q&&!of.includes(q)) ok=false;
    tr.style.display=ok?'':'none';
    if(ok) vis++;
  }});
  var seps=document.querySelectorAll('#hist-tbody .hist-sep');
  seps.forEach(function(sep){{
    var next=sep.nextElementSibling;
    var show=false;
    while(next&&next.classList.contains('hist-row')){{if(next.style.display!=='none'){{show=true;break;}}next=next.nextElementSibling;}}
    sep.style.display=show?'':'none';
  }});
  var cnt=document.getElementById('h-count');
  if(cnt) cnt.textContent=vis+' production'+(vis>1?'s':'');
}}
var _currentDashTab='accueil';
function showTab(name){{
  _currentDashTab=name;
  ['accueil','historique','rapports','rpt-jour'].forEach(function(n){{
    var p=document.getElementById('tab-'+n);
    var b=document.getElementById('tb-'+n);
    if(p) p.style.display=(n===name)?'flex':'none';
    if(b) b.classList.toggle('active',n===name);
  }});
  if(name==='rapports') loadRapports();
  if(name==='rpt-jour') loadRptJour();
}}
setInterval(function(){{if(_currentDashTab==='accueil') location.reload();}},15000);
</script>

<div id="dash-of-modal" class="dash-modal-overlay" onclick="if(event.target.id==='dash-of-modal')this.style.display='none'">
  <div class="dash-modal-box">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <span style="font-size:calc(14px*var(--zf,1));font-weight:800;color:#1e3a8a">📋 Détail OF</span>
      <button onclick="document.getElementById('dash-of-modal').style.display='none'" style="background:none;border:none;font-size:calc(18px*var(--zf,1));cursor:pointer;color:#64748b">✕</button>
    </div>
    <div id="dash-of-detail-content"></div>
  </div>
</div>

</body>
</html>"""

    # ── Compute embedded sessions & reports (no fetch calls needed) ──
    import json as _json_rpt
    _sess_map_r = {}
    _sess_evts_map_r = {}
    _sess_prods_map_r = {}
    for _, _r in _decl_cache:  # utilise _decl_cache (toujours à jour) au lieu de decl_rows (Excel potentiellement en retard)
        _dkey = str(_r[39] if len(_r) > 39 else "").strip() or _row_date(_r[2])
        if not _dkey: continue
        _pilot_r = str(_r[4] or ""); _poste_r = str(_r[3] or "")
        _rtype_r = str(_r[0] or "").strip().lower()
        _sk_r = f"{_dkey}||{_pilot_r}||{_poste_r}"
        if _sk_r not in _sess_map_r:
            _sess_map_r[_sk_r] = {"date":_dkey,"pilot":_pilot_r,"poste":_poste_r,"nb_of":0,"tot_equiv":0.0,"max_fin_s":0.0}
            _sess_evts_map_r[_sk_r] = []
            _sess_prods_map_r[_sk_r] = []
        if _rtype_r in ("production","prod",""):
            try:
                _eq_r = float(str(_r[21] or 0).replace(",","."))
                _fs_r = hms2s(_r[17])
                _sess_map_r[_sk_r]["nb_of"] += 1
                _sess_map_r[_sk_r]["tot_equiv"] += _eq_r
                _sess_prods_map_r[_sk_r].append(_r)
                if _fs_r > _sess_map_r[_sk_r]["max_fin_s"]: _sess_map_r[_sk_r]["max_fin_s"] = _fs_r
            except: pass
        else:
            _sess_evts_map_r[_sk_r].append(_r)
    _embedded_sessions_list = []
    for _sk_r, _s_r in _sess_map_r.items():
        _trs_r = -1.0
        try:
            _p2 = _s_r["date"].split('/'); _do2 = datetime.date(int(_p2[2]),int(_p2[1]),int(_p2[0]))
        except: _do2 = None
        _deb2, _ = _get_model_day_cfg(_s_r["poste"], _do2)
        _mds2 = hms2s(_deb2) if _deb2 else None
        _sess_evts2 = _sess_evts_map_r.get(_sk_r, [])
        _sess_prods2 = _sess_prods_map_r.get(_sk_r, [])
        _mdur2 = get_shift_duration_s(_s_r["poste"], _do2)
        if _pers_pct_map and _sess_prods2 and _s_r["tot_equiv"] > 0:
            _deg_ivs2 = _merged_degrade_ivs(_sess_evts2)
            _trs_r, _ = _option_b_trs(_sess_prods2, _deg_ivs2, prod_ref)
        elif _mdur2 > 0 and prod_ref > 0 and _s_r["tot_equiv"] > 0:
            _ded2 = sum(hms2s(_er[18]) for _er in _sess_evts2 if any(k in str(_er[0] or "").lower() for k in ["pause","nettoyage","réunion","reunion","meeting"]))
            _deg2 = _merged_degrade_s(_sess_evts2)
            _el2 = max(1.0, _mdur2 - _ded2)
            _trs_r = round(_s_r["tot_equiv"] / (prod_ref * _el2 / 28800) * 100, 1)
        _embedded_sessions_list.append({"date":_s_r["date"],"pilot":_s_r["pilot"],"poste":_s_r["poste"],"nb_of":_s_r["nb_of"],"tot_equiv":round(_s_r["tot_equiv"],1),"trs":_trs_r})
    _embedded_sessions_list.sort(key=lambda x: (lambda p: (int(p[2]),int(p[1]),int(p[0])) if len(p)==3 else (0,0,0))(x["date"].split('/')), reverse=True)
    _embedded_sessions_list = _embedded_sessions_list[:60]
    _embedded_reports_dict = {}
    for _s_r in _embedded_sessions_list:
        _sk3 = f"{_s_r['date']}||{_s_r['pilot']}||{_s_r['poste']}"
        _pr3=[]; _er3=[]; _pr3_raw=[]; _deg_ivs3=[]; _teq3=0.0; _ts3=0.0; _mfs3=0.0; _sts3=0.0; _ads3=[]; _afs3=[]
        for _, _r3 in _decl_cache:  # utilise _decl_cache (toujours à jour)
            _dk3 = str(_r3[39] if len(_r3)>39 else "").strip() or _row_date(_r3[2])
            if _dk3 != _s_r["date"] or str(_r3[4] or "") != _s_r["pilot"]: continue
            if _s_r["poste"] and str(_r3[3] or "") != _s_r["poste"]: continue
            _rt3 = str(_r3[0] or "").strip().lower()
            _dbs3 = hms2s(_r3[16]) if _r3[16] else -1; _fbs3 = hms2s(_r3[17]) if _r3[17] else -1
            if _dbs3 >= 0: _ads3.append(_dbs3)
            if _fbs3 >= 0: _afs3.append(_fbs3)
            if _rt3 in ("production","prod",""):
                try:
                    _eq3 = float(str(_r3[21] or 0).replace(",","."))
                    _ds3 = hms2s(_r3[16]); _fs3b = hms2s(_r3[17])
                    _dur3 = _fs3b - _ds3 if _fs3b > _ds3 else hms2s(_r3[18])
                    try: _r3_24=float(str(_r3[24] if len(_r3)>24 else '').strip() or '-1')
                    except: _r3_24=-1.0
                    _trs3 = _r3_24 if _r3_24>=0 else (round(_eq3/(prod_ref*_dur3/28800)*100,1) if prod_ref>0 and _dur3>0 and _eq3>0 else -1)
                    _teq3 += _eq3; _ts3 += _dur3
                    if _fs3b > _mfs3: _mfs3 = _fs3b
                    _pr3_raw.append(_r3)
                    _pr3.append({"of":str(_r3[1] or ""),"taille":str(_r3[7] or ""),"type_prod":str(_r3[9] or ""),"kit":str(_r3[15] or ""),"qte_fab":str(_r3[19] or ""),"equiv":str(_r3[21] or ""),"debut":str(_r3[16] or "")[:5],"fin":str(_r3[17] or "")[:5],"duree":str(_r3[18] or ""),"trs":_trs3,"comment":str(_r3[35] or ""),"nb_pers":str(_r3[6] or "")})
                except: pass
            else:
                try:
                    _durs3 = hms2s(_r3[18]); _sts3 += _durs3
                    _is_deg3 = _is_degrade_type(str(_r3[0] or "").strip())
                    if _is_deg3 and _dbs3 >= 0 and _fbs3 > _dbs3: _deg_ivs3.append((_dbs3, _fbs3))
                    _er3.append({"type":str(_r3[0] or ""),"of":str(_r3[1] or ""),"taille":str(_r3[7] or ""),"type_prod":str(_r3[9] or ""),"debut":str(_r3[16] or "")[:5],"fin":str(_r3[17] or "")[:5],"duree":str(_r3[18] or ""),"comment":str(_r3[35] or ""),"is_degrade":_is_deg3})
                except: pass
        _acd3 = _sec_to_hm(min(_ads3)) if _ads3 else ""
        _acf3 = _sec_to_hm(max(_afs3)) if _afs3 else ""
        try:
            _dp3 = _s_r["date"].split('/'); _dpo3 = datetime.date(int(_dp3[2]),int(_dp3[1]),int(_dp3[0]))
        except: _dpo3 = None
        _mdeb3, _mfin3 = _get_model_day_cfg(_s_r["poste"], _dpo3)
        _ded3 = sum(hms2s(_e3r.get("duree","")) for _e3r in _er3 if any(k in str(_e3r.get("type","")).lower() for k in ["pause","nettoyage","réunion","reunion","meeting"]))
        _mdur3 = get_shift_duration_s(_s_r["poste"], _dpo3)
        _ecart3 = max(0.0, _mdur3 - (_ts3 + _sts3))
        _trs_sh3 = -1.0
        if _pers_pct_map and _pr3_raw and _teq3 > 0:
            _deg_mg3 = []
            for _si3, _fi3 in sorted(_deg_ivs3):
                if _deg_mg3 and _si3 <= _deg_mg3[-1][1]: _deg_mg3[-1] = (_deg_mg3[-1][0], max(_deg_mg3[-1][1], _fi3))
                else: _deg_mg3.append((_si3, _fi3))
            _trs_sh3, _ = _option_b_trs(_pr3_raw, _deg_mg3, prod_ref)
        elif _mdur3 > 0 and prod_ref > 0 and _teq3 > 0:
            _deg3 = sum(hms2s(_e3r.get("duree","")) for _e3r in _er3 if _e3r.get("is_degrade"))
            _el3 = max(1.0, _mdur3 - _ded3)
            _trs_sh3 = round(_teq3/(prod_ref*_el3/28800)*100,1)
        _trs_of3 = round(_teq3/(prod_ref*_ts3/28800)*100,1) if prod_ref>0 and _ts3>0 and _teq3>0 else -1
        _rpt_key3 = f"{_s_r['date']}|{_s_r['pilot']}|{_s_r['poste']}"
        _embedded_reports_dict[_rpt_key3] = {"date":_s_r["date"],"pilot":_s_r["pilot"],"poste":_s_r["poste"],"prod_rows":_pr3,"evt_rows":_er3,"trs_shift":_trs_sh3,"trs":_trs_of3,"tot_equiv":round(_teq3,1),"tot_s":round(_ts3,0),"stop_s":round(_sts3,0),"nb_of":len(_pr3),"model_debut":_mdeb3 or "","model_fin":_mfin3 or "","actual_debut":_acd3,"actual_fin":_acf3,"ecart_s":round(_ecart3,0),"model_dur_s":round(_mdur3,0),"planned_ded_s":round(_ded3,0)}
    # --- STATIC RAPPORTS HTML GENERATION ---
    def _resc(s):
        return str(s or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

    def _svg_gauge_static(trs):
        v = max(0.0, min(100.0, float(trs) if (trs is not None and trs >= 0) else 0.0))
        dash = (v / 100.0) * 132
        col = '#16a34a' if v >= 90 else '#d97706' if v >= 75 else '#dc2626'
        trs_str = f'{trs:.1f}%' if (trs is not None and trs >= 0) else '--%'
        return (
            f'<svg width="110" height="72" viewBox="0 0 110 72">'
            f'<circle cx="55" cy="54" r="38" fill="none" stroke="#e2e8f0" stroke-width="11"'
            f' stroke-dasharray="119,200" stroke-linecap="round" transform="rotate(134 55 54)"/>'
            f'<circle cx="55" cy="54" r="38" fill="none" stroke="{col}" stroke-width="11"'
            f' stroke-dasharray="{dash:.1f},200" stroke-linecap="round" transform="rotate(134 55 54)"/>'
            f'<text x="55" y="58" text-anchor="middle" font-size="14" font-weight="900" fill="{col}">{_resc(trs_str)}</text>'
            f'<text x="55" y="69" text-anchor="middle" font-size="8" fill="#64748b">TRS</text>'
            f'</svg>'
        )

    def _svg_pie_static(prod_min, stop_min, total_min):
        autre = max(0.0, total_min - prod_min - stop_min)
        segs = [(prod_min, '#16a34a', 'Prod'), (stop_min, '#dc2626', 'Arrêts'), (autre, '#94a3b8', 'Autre')]
        total = sum(s[0] for s in segs)
        if total <= 0:
            return '<svg width="130" height="120" viewBox="0 0 130 120"><text x="65" y="60" text-anchor="middle" font-size="9" fill="#94a3b8">Pas de données</text></svg>'
        cx, cy, r, ir = 65.0, 57.0, 44.0, 24.0
        parts = []
        start = -math.pi / 2
        for val, color, label in segs:
            if val <= 0: continue
            angle = (val / total) * 2 * math.pi
            if angle < 0.001: continue
            end = start + angle
            large = 1 if angle > math.pi else 0
            x1 = cx + r * math.cos(start); y1 = cy + r * math.sin(start)
            x2 = cx + r * math.cos(end);   y2 = cy + r * math.sin(end)
            ix1 = cx + ir * math.cos(start); iy1 = cy + ir * math.sin(start)
            ix2 = cx + ir * math.cos(end);   iy2 = cy + ir * math.sin(end)
            parts.append(f'<path d="M{x1:.2f},{y1:.2f} A{r:.0f},{r:.0f} 0 {large},1 {x2:.2f},{y2:.2f} L{ix2:.2f},{iy2:.2f} A{ir:.0f},{ir:.0f} 0 {large},0 {ix1:.2f},{iy1:.2f} Z" fill="{color}"/>')
            start = end
        prod_pct = round(prod_min / total * 100) if total > 0 else 0
        parts.append(f'<text x="65" y="62" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e">{prod_pct}%</text>')
        parts.append(f'<text x="65" y="72" text-anchor="middle" font-size="7" fill="#64748b">Prod</text>')
        lx = 0
        for val, color, label in segs:
            if val <= 0: continue
            p = round(val / total * 100)
            parts.append(f'<rect x="{lx}" y="108" width="7" height="7" fill="{color}" rx="1"/>')
            parts.append(f'<text x="{lx+9}" y="115" font-size="7" fill="#475569">{_resc(label)} {p}%</text>')
            lx += 65
        return f'<svg width="130" height="120" viewBox="0 0 130 120">{"".join(parts)}</svg>'

    def _render_rpt_panel(idx, d, visible=False):
        trs_s = d.get('trs_shift', -1) if (d.get('trs_shift') is not None and d.get('trs_shift', -1) >= 0) else d.get('trs', -1)
        trs_col = '#16a34a' if trs_s >= 90 else '#d97706' if trs_s >= 70 else '#dc2626' if trs_s >= 0 else '#94a3b8'
        trs_str = f'{trs_s:.1f}%' if trs_s >= 0 else '—'
        trs_of = d.get('trs', -1)
        trs_of_str = f'{trs_of:.1f}%' if trs_of >= 0 else '--'
        stop_min = round((d.get('stop_s', 0) or 0) / 60)
        prod_min = round(max(0, (d.get('tot_s', 0) or 0) - (d.get('stop_s', 0) or 0)) / 60)
        total_min = round((d.get('tot_s', 0) or 0) / 60) + stop_min
        ecart_mn = round((d.get('ecart_s', 0) or 0) / 60)
        stop_map = {}
        for _er in (d.get('evt_rows') or []):
            k = _er.get('type') or 'Inconnu'
            _dur = (_er.get('duree') or '')
            _p2 = (_dur + ':00:00').split(':')
            try: _s2 = int(_p2[0] or 0)*3600 + int(_p2[1] or 0)*60 + int(_p2[2] or 0)
            except: _s2 = 0
            stop_map[k] = stop_map.get(k, 0) + _s2 / 60
        stop_arr = sorted(stop_map.items(), key=lambda x: -x[1])
        max_stop = stop_arr[0][1] if stop_arr else 1.0
        pareto_h = ''.join(
            f'<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">'
            f'<div style="font-size:calc(10px*var(--zf,1));width:100px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">{_resc(e[0])}</div>'
            f'<div style="flex:1;background:#f1f5f9;border-radius:4px;height:14px;overflow:hidden">'
            f'<div style="height:100%;background:#dc2626;border-radius:4px;width:{round(e[1]/max_stop*100) if max_stop>0 else 0}%;opacity:.8"></div></div>'
            f'<div style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#dc2626;width:36px;text-align:right;flex-shrink:0">{round(e[1])}mn</div></div>'
            for e in stop_arr
        ) if stop_arr else '<div style="color:#64748b;font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>'
        prod_h = ''
        for _pr in (d.get('prod_rows') or []):
            _to = _pr.get('trs', -1)
            _tc = '#16a34a' if _to >= 90 else '#d97706' if _to >= 70 else '#dc2626' if _to >= 0 else '#94a3b8'
            _ts = f'{_to:.1f}%' if _to >= 0 else '—'
            prod_h += (
                f'<tr style="border-bottom:1px solid #e2e8f0">'
                f'<td style="padding:4px 6px;font-weight:700;color:#1e3a8a">{_resc(_pr.get("of",""))}</td>'
                f'<td style="padding:4px 6px;font-size:calc(11px*var(--zf,1))">{_resc(_pr.get("taille",""))} {_resc(_pr.get("type_prod",""))}</td>'
                f'<td style="padding:4px 6px;text-align:center">{"✓" if str(_pr.get("kit","")).lower()=="oui" else ""}</td>'
                f'<td style="padding:4px 6px">{_resc(_pr.get("qte_fab",""))}</td>'
                f'<td style="padding:4px 6px;color:#0891b2;font-weight:700">{_resc(_pr.get("equiv",""))}</td>'
                f'<td style="padding:4px 6px;white-space:nowrap">{_resc(_pr.get("debut",""))} → {_resc(_pr.get("fin",""))}</td>'
                f'<td style="padding:4px 6px;font-weight:800;color:{_tc}">{_ts}</td>'
                f'<td style="padding:4px 6px;font-size:calc(10px*var(--zf,1));color:#64748b">{_resc(_pr.get("comment",""))}</td>'
                f'</tr>'
            )
        if not prod_h:
            prod_h = '<tr><td colspan="8" style="padding:8px;text-align:center;color:#94a3b8">Aucune production</td></tr>'
        evts_h = ''
        for _ev in (d.get('evt_rows') or []):
            evts_h += (
                f'<tr style="border-bottom:1px solid #e2e8f0">'
                f'<td style="padding:3px 5px;font-weight:600">{_resc(_ev.get("type",""))}</td>'
                f'<td style="padding:3px 5px;color:#0369a1">{_resc(_ev.get("of","—"))}</td>'
                f'<td style="padding:3px 5px;white-space:nowrap;color:#64748b">{_resc(_ev.get("debut",""))} → {_resc(_ev.get("fin",""))}</td>'
                f'<td style="padding:3px 5px;font-weight:700">{_resc(_ev.get("duree",""))}</td>'
                f'<td style="padding:3px 5px;color:#64748b">{_resc(_ev.get("comment","—"))}</td>'
                f'</tr>'
            )
        evts_section = (
            f'<table style="width:100%;border-collapse:collapse;font-size:calc(10px*var(--zf,1))"><thead><tr style="background:#f8fafc;border-bottom:1px solid #e2e8f0">'
            f'<th style="padding:3px 5px;text-align:left;font-weight:700;color:#64748b">Arrêt</th>'
            f'<th style="padding:3px 5px;font-weight:700;color:#64748b">OF</th>'
            f'<th style="padding:3px 5px;font-weight:700;color:#64748b">Plage</th>'
            f'<th style="padding:3px 5px;font-weight:700;color:#64748b">Durée</th>'
            f'<th style="padding:3px 5px;font-weight:700;color:#64748b">Commentaire</th>'
            f'</tr></thead><tbody>{evts_h}</tbody></table>'
        ) if evts_h else '<div style="color:#64748b;font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>'
        _bgt_labels = {"pause_min":"Pause","meeting_tol_min":"Réunion","clean_short_min":"Nettoyage court","clean_long_min":"Nettoyage long","clean_grand_min":"Nettoyage très long"}
        _bgt_used = {bk:0.0 for bk in _bgt_labels}
        for _er in (d.get('evt_rows') or []):
            _bk2 = None
            _rl = str(_er.get('type','') or '').lower()
            if 'nettoyage' in _rl or 'nett' in _rl:
                if 'très long' in _rl or 'tres long' in _rl or 'grand' in _rl: _bk2='clean_grand_min'
                elif 'long' in _rl: _bk2='clean_long_min'
                elif 'court' in _rl: _bk2='clean_short_min'
            elif 'réunion' in _rl or 'reunion' in _rl or 'meeting' in _rl: _bk2='meeting_tol_min'
            elif 'pause' in _rl: _bk2='pause_min'
            if _bk2:
                _dp2=str(_er.get('duree') or ''); _pp2=(_dp2+':00:00').split(':')
                try: _bs2=int(_pp2[0] or 0)*3600+int(_pp2[1] or 0)*60+int(_pp2[2] or 0)
                except: _bs2=0
                _bgt_used[_bk2]+=_bs2/60
        _budget_bars_h=''
        for _bk3,_bl3 in [('clean_short_min','Nettoyage court'),('clean_long_min','Nettoyage long'),('clean_grand_min','Nettoyage très long'),('meeting_tol_min','Réunion'),('pause_min','Pause')]:
            _bm=float(cfg.get(_bk3,0) or 0); _bu=_bgt_used.get(_bk3,0.0)
            if _bm<=0: continue
            _pct=min(100,round(_bu/_bm*100)) if _bm>0 else 0
            _bc='#dc2626' if _bu>_bm else '#d97706' if _bu/_bm>=0.8 else '#16a34a'
            _budget_bars_h+=(
                f'<div style="margin-bottom:7px">'
                f'<div style="display:flex;justify-content:space-between;font-size:calc(11px*var(--zf,1));margin-bottom:3px">'
                f'<span>{_resc(_bl3)}</span>'
                f'<span style="font-weight:700;color:{_bc}">{round(_bu)}/{round(_bm)} min</span>'
                f'</div>'
                f'<div style="background:#f1f5f9;border-radius:4px;height:12px;overflow:hidden">'
                f'<div style="height:100%;background:{_bc};border-radius:4px;width:{_pct}%;opacity:.85"></div>'
                f'</div></div>'
            )
        if not _budget_bars_h: _budget_bars_h='<div style="color:#94a3b8;font-size:calc(12px*var(--zf,1))">Aucun budget configuré</div>'
        _tot_qte_fab = sum(float(str(pr.get("qte_fab","") or 0).replace(",",".")) for pr in (d.get("prod_rows") or []))
        _eff_s = max(1.0, float(d.get("model_dur_s",0) or 0) - float(d.get("planned_ded_s",0) or 0))
        _cad_h = round(_tot_qte_fab / _eff_s * 3600) if _eff_s > 0 and _tot_qte_fab > 0 else 0
        _sorted_f = sorted([pr for pr in (d.get("prod_rows") or []) if pr.get("fibre")], key=lambda x: x.get("debut",""))
        _nb_chg_f = sum(1 for _i in range(1, len(_sorted_f)) if _sorted_f[_i]["fibre"] != _sorted_f[_i-1]["fibre"])
        plage_str = ''
        if d.get('actual_debut') and d.get('actual_fin'):
            plage_str = f' · {_resc(d.get("actual_debut",""))} → {_resc(d.get("actual_fin",""))}'
        elif d.get('model_debut') and d.get('model_fin'):
            plage_str = f' · Modèle : {_resc(d.get("model_debut",""))} → {_resc(d.get("model_fin",""))}'
        ecart_div = (
            f'<div class="fp-card" style="padding:7px;border:1.5px solid #f59e0b">'
            f'<div class="fp-big" style="font-size:calc(16px*var(--zf,1));color:#d97706">{ecart_mn} min</div>'
            f'<div class="fp-lbl">Non déclaré</div></div>'
        ) if ecart_mn > 0 else ''
        disp = 'flex' if visible else 'none'
        g_svg = _svg_gauge_static(trs_s)
        p_svg = _svg_pie_static(prod_min, stop_min, total_min)
        return (
            f'<div class="rpt-det-panel" id="rpt-det-{idx}" style="display:{disp};flex-direction:column;overflow-y:auto">'
            f'<div style="background:#1e3a8a;color:#fff;padding:10px 16px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0">'
            f'<div><div style="font-size:calc(15px*var(--zf,1));font-weight:800">📋 Rapport — {_resc(d.get("poste",""))}</div>'
            f'<div style="font-size:calc(11px*var(--zf,1));opacity:.8">{_resc(d.get("pilot",""))} · {_resc(d.get("date",""))}{plage_str}</div></div>'
            f'<div style="text-align:right"><div style="font-size:calc(26px*var(--zf,1));font-weight:900;color:{trs_col}">{trs_str}</div>'
            f'<div style="font-size:calc(11px*var(--zf,1));opacity:.7">TRS Shift</div></div></div>'
            f'<div style="display:flex;gap:12px;padding:10px 14px;background:#fff;border-bottom:1px solid #e2e8f0;align-items:center;flex-wrap:wrap">'
            f'<div style="text-align:center;flex-shrink:0"><div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;margin-bottom:4px">TRS Poste</div>{g_svg}</div>'
            f'<div style="text-align:center;flex-shrink:0"><div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;margin-bottom:4px">Répartition</div>{p_svg}</div>'
            f'<div style="flex:1;display:flex;flex-direction:column;gap:5px">'
            f'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:5px">'
            f'<div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#059669;font-weight:900">{round(_tot_qte_fab)}</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Nb pièces prod.</div></div>'
            f'<div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#0891b2;font-weight:900">{round(d.get("tot_equiv",0) or 0)}</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Équivalence</div></div>'
            f'<div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#0369a1;font-weight:900">{_cad_h}</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Cadence/h</div></div>'
            f'</div>'
            f'<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:5px">'
            f'<div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1))">{(_resc(d.get("model_debut",""))+"→"+_resc(d.get("model_fin",""))) if d.get("model_debut") and d.get("model_fin") else (str(round((d.get("model_dur_s",0) or 0)/60))+" min")}</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Durée ouverture</div></div>'
            f'<div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#16a34a">{prod_min} min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Durée prod</div></div>'
            f'<div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#dc2626">{stop_min} min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts total</div></div>'
            f'<div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#16a34a">{round((d.get("planned_ded_s",0) or 0)/60)} min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts prévus</div></div>'
            f'<div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#7c3aed">{d.get("nb_of",0)}</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Nb OF</div></div>'
            f'<div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(13px*var(--zf,1));color:#8b5cf6">{_nb_chg_f}</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Chg. fibre</div></div>'
            f'</div></div></div>'
            f'<div style="flex:1;overflow-y:auto;padding:8px 12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">'
            f'<div class="rpt-card" style="padding:10px"><div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#64748b;margin-bottom:6px">Productions</div>'
            f'<table style="width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1))"><thead><tr style="background:#f8fafc">'
            f'<th style="padding:4px 6px;text-align:left">OF</th><th style="padding:4px 6px;text-align:left">Taille</th>'
            f'<th style="padding:4px 6px">Kit</th><th style="padding:4px 6px">Qté</th><th style="padding:4px 6px">Éq.</th>'
            f'<th style="padding:4px 6px">Heures</th><th style="padding:4px 6px">TRS</th><th style="padding:4px 6px">Comm.</th>'
            f'</tr></thead><tbody>{prod_h}</tbody></table></div>'
            f'<div class="rpt-card" style="padding:10px"><div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#92400e;margin-bottom:8px">⏱ Arrêts prévus</div>{_budget_bars_h}</div>'
            f'<div style="display:flex;flex-direction:column;gap:8px">'
            f'<div class="rpt-card" style="padding:10px"><div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#64748b;margin-bottom:8px">Pareto arrêts</div>{pareto_h}</div>'
            f'<div class="rpt-card" style="padding:10px"><div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#64748b;margin-bottom:6px">Détail arrêts</div>{evts_section}</div>'
            f'</div></div></div>'
        )

    _rpt_list_html = ''
    _rpt_det_html = ''
    if not _embedded_sessions_list:
        _rpt_list_html = '<div style="padding:20px;text-align:center;color:#94a3b8;font-size:calc(12px*var(--zf,1))">Aucun poste disponible</div>'
        _rpt_det_html = '<div style="padding:60px;text-align:center;color:#94a3b8"><div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">📋</div><div style="font-size:calc(14px*var(--zf,1));font-weight:600">Aucun rapport disponible</div></div>'
    else:
        for _ri, _sr in enumerate(_embedded_sessions_list):
            _tv = _sr.get('trs', -1)
            _ts2 = f'{_tv:.1f}%' if _tv >= 0 else '—'
            _tc2 = '#16a34a' if _tv >= 90 else '#f59e0b' if _tv >= 70 else '#dc2626' if _tv >= 0 else '#94a3b8'
            _sel_st = ' style="padding:10px 14px;border-bottom:1px solid #e2e8f0;cursor:pointer;transition:background .15s;background:#eff6ff"' if _ri == 0 else ' style="padding:10px 14px;border-bottom:1px solid #e2e8f0;cursor:pointer;transition:background .15s"'
            _rpt_list_html += (
                f'<div class="rpt-item" id="rpt-item-{_ri}" onclick="showRptPanel({_ri})"{_sel_st}>'
                f'<div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:#1e3a8a">{_resc(_sr.get("date",""))} — {_resc(_sr.get("poste",""))}</div>'
                f'<div style="font-size:calc(11px*var(--zf,1));color:#64748b;margin-top:2px">{_resc(_sr.get("pilot","?"))} | {_sr.get("nb_of",0)} OF | Éq. {_sr.get("tot_equiv",0)}</div>'
                f'<div style="font-size:calc(16px*var(--zf,1));font-weight:900;color:{_tc2};margin-top:2px">{_ts2}</div>'
                f'</div>'
            )
            _rk = f"{_sr['date']}|{_sr['pilot']}|{_sr['poste']}"
            _rd = _embedded_reports_dict.get(_rk, _sr)
            _rpt_det_html += _render_rpt_panel(_ri, _rd, visible=(_ri == 0))

    # Inject rapports JS (minimal – all content is pre-rendered)
    _rapports_js = """<script>
function showRptPanel(idx){
  document.querySelectorAll('.rpt-det-panel').forEach(function(el){el.style.display='none';});
  document.querySelectorAll('.rpt-item').forEach(function(el){el.style.background='';});
  var panel=document.getElementById('rpt-det-'+idx);
  if(panel){panel.style.display='flex';}
  var item=document.getElementById('rpt-item-'+idx);
  if(item){item.style.background='#eff6ff';}
}
function loadRapports(){}
</script>"""
    html = html.replace('__RPT_LIST__', _rpt_list_html, 1)
    html = html.replace('__RPT_DET__', _rpt_det_html, 1)
    html = html.replace('</body>', _rapports_js + '\n</body>', 1)

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
<title>PRODUCTION ORC — DODO</title>
<style>
:root{
  --navy:#1a1f5e;--navy2:#2d3480;--green:#16a34a;--red:#dc2626;
  --amber:#d97706;--purple:#7c3aed;--blue:#0891b2;
  --bg:#f0f2f8;--card:#fff;--border:#dde4ef;--text:#1e293b;--gray:#64748b;
  --lgray:#e2e8f0;--radius:10px;--shadow:0 2px 10px rgba(0,0,0,.07);
  --hdr-h:65px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:calc(13px*var(--zf,1));height:100vh;overflow:hidden;display:flex;flex-direction:column}

/* ── STOP ACTIVE THEME ── */
body.stop-on #app-hdr{background:linear-gradient(90deg,#fff 0px,#fee2e2 160px,#7f0000 280px,#7f0000 100%)!important;border-color:#b91c1c}

/* ── HEADER ── */
#app-hdr{height:var(--hdr-h);background:linear-gradient(90deg,#ffffff 0px,#ffffff 150px,#dde4f7 230px,var(--navy) 320px,var(--navy) 100%);display:flex;align-items:flex-end;padding:0 14px;gap:10px;flex-shrink:0;box-shadow:0 4px 18px rgba(26,31,94,.32),inset 0 -1px 0 rgba(255,255,255,.12)}
.hdr-logo{display:none}
.hdr-tabs{display:flex;gap:4px;flex:1;align-self:flex-end}
.htab{background:rgba(26,31,94,0.72);border:1.5px solid rgba(255,255,255,.25);border-bottom:3px solid transparent;color:rgba(255,255,255,.82);padding:7px 16px 9px;border-radius:8px 8px 0 0;cursor:pointer;font-size:calc(12px*var(--zf,1));font-weight:700;white-space:nowrap;transition:all .15s;position:relative;top:3px;box-shadow:inset 0 1px 0 rgba(255,255,255,.18),0 -2px 6px rgba(0,0,0,.15)}
.htab:hover{background:linear-gradient(180deg,rgba(255,255,255,.28) 0%,rgba(255,255,255,.12) 100%);border-color:rgba(255,255,255,.42);border-bottom-color:transparent;color:#fff;box-shadow:inset 0 1px 0 rgba(255,255,255,.28),0 -3px 8px rgba(0,0,0,.2)}
.htab.on{background:linear-gradient(180deg,#fff 0%,#f0f4ff 100%);color:#1e3a8a;font-weight:800;border-color:rgba(255,255,255,.5);border-bottom:3px solid #fff;box-shadow:0 -4px 10px rgba(0,0,0,.15),inset 0 1px 0 #fff,inset 0 -1px 0 rgba(30,58,138,.1)}
.htab.prod-on{background:#22c55e!important;color:#fff!important;font-weight:800;animation:pt 1.4s ease-in-out infinite;border-color:transparent!important;border-bottom-color:transparent!important;letter-spacing:.3px}
#ht-prod{display:none}
#ht-prod.prod-visible{display:inline-block!important}
@keyframes pt{0%,100%{opacity:1;transform:scale(1);box-shadow:0 0 5px 2px rgba(34,197,94,.4)}50%{opacity:.45;transform:scale(1.08);box-shadow:0 0 16px 6px rgba(34,197,94,.9)}}
#hdr-right{display:flex;align-items:center;gap:8px;margin-left:auto;font-size:calc(11px*var(--zf,1));color:rgba(255,255,255,.75);align-self:center}
#hdr-pilot-lbl{font-weight:800;color:#fff;font-size:calc(18px*var(--zf,1));letter-spacing:.3px}
#ht-guest-badge{display:none!important}
#main-prod-banner{display:none!important}

/* ── ALERT STRIP ── */
#alert-strip{background:#b91c1c;color:#fff;text-align:center;padding:4px;font-weight:700;font-size:calc(12px*var(--zf,1));flex-shrink:0;display:none;animation:blink .85s step-start infinite}
#alert-strip.on{display:block}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.35}}
@keyframes rpt-arr{0%,100%{background:#e5e7eb}50%{background:#f9fafb}}

/* ── VIEWS ── */
.view{display:none;flex:1;flex-direction:column;overflow:hidden}
.view.on{display:flex}

/* ── LOGIN ── */
#v-login{background:linear-gradient(135deg,#1a1f5e,#2d3480,#1e3a8a);align-items:center;justify-content:center}
.login-card{background:#fff;border-radius:14px;padding:32px;width:100%;max-width:380px;box-shadow:0 20px 60px rgba(0,0,0,.35)}
.lc-h1{color:var(--navy);font-size:calc(24px*var(--zf,1));font-weight:800;margin-bottom:2px;text-align:center}
.lc-sub{color:#64748b;text-align:center;margin-bottom:20px;font-size:calc(12px*var(--zf,1))}
.lf{margin-bottom:12px}
.lf label{display:block;font-weight:600;margin-bottom:3px;color:#374151;font-size:calc(11px*var(--zf,1));text-transform:uppercase;letter-spacing:.5px}
.lf select,.lf input{width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:calc(14px*var(--zf,1));outline:none;transition:border .2s}
.lf select:focus,.lf input:focus{border-color:var(--navy)}
.btn-login{width:100%;padding:11px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:calc(14px*var(--zf,1));font-weight:700;cursor:pointer;margin-top:4px}
.btn-login:hover{opacity:.88}
.ln-err{color:#dc2626;text-align:center;margin-top:6px;font-size:calc(12px*var(--zf,1));min-height:16px}

/* ── MAIN VIEW (Déclarations / Évts) ── */
#v-main{overflow:hidden}
.main-hdr{background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap}
.main-hdr .mtabs{display:flex;gap:3px}
.mtab{background:none;border:none;border-bottom:2px solid transparent;padding:5px 12px;cursor:pointer;font-size:calc(12px*var(--zf,1));font-weight:600;color:var(--gray);transition:all .15s}
.mtab.on{border-color:var(--navy);color:var(--navy)}
.main-hdr .mbtns{display:flex;gap:6px;margin-left:auto}
.btn-sm{padding:6px 12px;border:none;border-radius:6px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer;transition:all .15s}
.btn-green{background:var(--green);color:#fff}
.btn-green:hover{filter:brightness(.9)}
.btn-ghost{background:var(--lgray);color:var(--text)}
.btn-ghost:hover{filter:brightness(.93)}
.table-wrap{flex:1;overflow-y:auto}
.ktbl{width:100%;border-collapse:collapse;font-size:calc(12px*var(--zf,1));table-layout:auto}
.ktbl th{text-align:left;padding:8px 10px;background:linear-gradient(180deg,#2d3480 0%,var(--navy) 100%);color:#fff;font-size:calc(10px*var(--zf,1));text-transform:uppercase;letter-spacing:.5px;position:sticky;top:0;font-weight:800;border-right:1px solid rgba(255,255,255,.1);box-shadow:0 2px 4px rgba(0,0,0,.2),inset 0 1px 0 rgba(255,255,255,.15)}
.ktbl th:last-child{border-right:none}
.ktbl td{padding:6px 10px;border-bottom:1px solid var(--border);border-right:1px solid #f1f5f9;transition:background .1s}
.ktbl td:last-child{border-right:none}
.ktbl tbody tr:nth-child(even) td{background:#f7f9ff}
.ktbl tbody tr:nth-child(odd) td{background:#fff}
.ktbl tbody tr:hover td{background:#e8eeff!important;box-shadow:inset 0 0 0 9999px rgba(99,102,241,.07)}
.tg{color:var(--green);font-weight:700}
.tm{color:var(--amber);font-weight:700}
.tb{color:var(--red);font-weight:700}
.btn-tbl{padding:3px 8px;border:none;border-radius:4px;font-size:calc(10px*var(--zf,1));cursor:pointer;font-weight:600}

/* ── PRODUCTION VIEW ── */
#v-prod{overflow:hidden}
/* pilot/OF banner */
.pob{background:var(--navy);color:#fff;padding:6px 14px;display:flex;align-items:center;gap:24px;flex-shrink:0}
.pob-item{display:flex;flex-direction:column;gap:1px}
.pob-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.7px;opacity:.7;font-weight:600}
.pob-val{font-size:calc(22px*var(--zf,1));font-weight:800;line-height:1}
.pob-item.of .pob-val{font-size:calc(20px*var(--zf,1));color:#93c5fd}
.pob-item.trs .pob-val{color:#86efac}
/* status bar */
.sbar{display:flex;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0}
.sc{flex:1;padding:6px 14px;border-right:1px solid var(--border);text-align:center}
.sc:last-child{border:none}
.sc-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.6px}
.sc-val{font-size:calc(22px*var(--zf,1));font-weight:800;font-variant-numeric:tabular-nums;color:var(--navy);margin-top:1px}
.sc-val.green{color:var(--green)}
.sc-val.red{color:var(--red)}
.sc-val.amber{color:var(--amber)}
/* main 2-col layout */
.prod-body{display:flex;flex:1;overflow:hidden}
/* CENTER form col (now left) */
.form-col{flex:1;overflow-y:auto;padding:8px;display:flex;flex-direction:column;gap:6px}
/* Action buttons row below timeline */
.prod-act-row{display:flex;gap:4px;flex-wrap:nowrap;padding:6px 0 2px;border-top:1px solid var(--border);margin-top:2px;position:sticky;bottom:0;background:var(--card);z-index:10;height:120px;align-items:stretch}
.act-btn{flex:1;min-width:100px;border:none;border-radius:14px;padding:8px 6px 10px;cursor:pointer;font-size:calc(14px*var(--zf,1));font-weight:700;text-align:center;transition:all .12s;white-space:nowrap;min-height:56px;display:flex;align-items:center;justify-content:center;gap:3px;flex-direction:column;line-height:1.25;box-shadow:0 8px 0 rgba(0,0,0,.3),0 10px 16px rgba(0,0,0,.25),inset 0 2px 3px rgba(255,255,255,.35),inset 0 -3px 6px rgba(0,0,0,.2);transform:translateY(0);position:relative;overflow:hidden}
.act-btn-sm{aspect-ratio:1!important;height:100%!important;min-height:0!important;min-width:0!important;padding:6px 4px!important;font-size:calc(16px*var(--zf,1))!important;flex:1!important}
.act-btn-sm .act-icon{font-size:calc(30px*var(--zf,1))!important;margin-bottom:2px}
.stop-icon{position:relative;display:inline-flex;align-items:center;justify-content:center}
.stop-icon-x{position:absolute;font-size:.55em;font-weight:900;color:#fff;text-shadow:none;line-height:1}
.act-btn::before{content:'';position:absolute;top:0;left:0;right:0;height:50%;background:linear-gradient(180deg,rgba(255,255,255,.22) 0%,rgba(255,255,255,0) 100%);border-radius:14px 14px 0 0;pointer-events:none}
.act-btn:hover{filter:brightness(1.08)}
.act-btn:active{transform:translateY(6px);box-shadow:0 2px 0 rgba(0,0,0,.3),0 3px 6px rgba(0,0,0,.2),inset 0 1px 2px rgba(255,255,255,.2),inset 0 -1px 3px rgba(0,0,0,.15)}
.act-stop{background:radial-gradient(ellipse at 50% 25%,#f87171 0%,#dc2626 55%,#991b1b 100%);color:#fff;font-size:calc(16px*var(--zf,1));font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.4)}
.act-nett{background:radial-gradient(ellipse at 50% 25%,#fdba74 0%,#f97316 55%,#c2410c 100%);color:#fff;font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.4)}
.act-pause{background:radial-gradient(ellipse at 50% 25%,#cbd5e1 0%,#64748b 55%,#334155 100%);color:#fff;font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.4)}
.act-cancel{background:radial-gradient(ellipse at 50% 25%,#94a3b8 0%,#64748b 55%,#334155 100%);color:#fff;font-weight:700;text-shadow:0 1px 3px rgba(0,0,0,.3)}
.act-endprod{background:radial-gradient(ellipse at 50% 25%,#4ade80 0%,#16a34a 55%,#14532d 100%);color:#fff;font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.4)}
.act-icon{font-size:calc(34px*var(--zf,1));line-height:1;display:block;margin-bottom:3px}
/* Accueil 3D buttons */
.acc-btn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;border:none;border-radius:16px;padding:14px 22px;min-height:80px;min-width:130px;cursor:pointer;font-size:calc(14px*var(--zf,1));font-weight:800;color:#fff;text-shadow:0 1px 3px rgba(0,0,0,.4);box-shadow:0 8px 0 rgba(0,0,0,.3),0 10px 16px rgba(0,0,0,.25),inset 0 2px 3px rgba(255,255,255,.35),inset 0 -3px 6px rgba(0,0,0,.2);transform:translateY(0);transition:transform .1s,box-shadow .1s;position:relative;overflow:hidden}
.acc-btn::before{content:'';position:absolute;top:0;left:0;right:0;height:50%;background:linear-gradient(180deg,rgba(255,255,255,.22) 0%,rgba(255,255,255,0) 100%);border-radius:16px 16px 0 0;pointer-events:none}
.acc-btn:hover{filter:brightness(1.08)}
.acc-btn:active{transform:translateY(6px);box-shadow:0 2px 0 rgba(0,0,0,.3),0 3px 6px rgba(0,0,0,.2),inset 0 1px 2px rgba(255,255,255,.2),inset 0 -1px 3px rgba(0,0,0,.15)}
.acc-green{background:radial-gradient(ellipse at 50% 25%,#4ade80 0%,#16a34a 55%,#14532d 100%)}
.acc-red{background:radial-gradient(ellipse at 50% 25%,#f87171 0%,#dc2626 55%,#991b1b 100%)}
.acc-amber{background:radial-gradient(ellipse at 50% 25%,#fde68a 0%,#f59e0b 55%,#92400e 100%)}
/* Field highlight when form incomplete */
@keyframes flash-field{0%{box-shadow:0 0 0 3px rgba(220,38,38,0.1);}50%{box-shadow:0 0 0 4px rgba(220,38,38,0.7);}100%{box-shadow:0 0 0 3px rgba(220,38,38,0.35);}}
.field-missing{border-color:#dc2626!important;background:#fff0f0!important;box-shadow:0 0 0 3px rgba(220,38,38,0.35)!important;animation:flash-field 0.6s ease 3;}
/* 3-col form zones */
.form-3col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.fzone{border-radius:7px;padding:8px}
.fzone h4{font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.7px;font-weight:700;margin-bottom:6px;padding-bottom:3px;border-bottom:1px solid rgba(0,0,0,.08)}
.zi{background:#d4ddff;border:1px solid #a5b4fc}.zi h4{color:#3730a3}
.zp{background:#bbf7d0;border:1px solid #86efac}.zp h4{color:#166534}
.zq{background:#fed7aa;border:1px solid #fb923c}.zq h4{color:#9a3412}
.fr{display:flex;flex-direction:column;margin-bottom:4px}
.fr label{font-size:calc(12.65px*var(--zf,1));font-weight:700;color:var(--gray);margin-bottom:2px;text-transform:uppercase;letter-spacing:.2px}
.fr input,.fr select,.fr textarea{padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1));background:var(--card);color:var(--text);width:100%;outline:none;transition:border .15s}
.fr input:focus,.fr select:focus{border-color:#6366f1}
.fr textarea{resize:none;height:42px}
.fr.comment-big textarea{height:80px;font-size:calc(13px*var(--zf,1));border:2px solid #f59e0b;background:#fffbeb;font-weight:500}
.fr.comment-big label{color:#d97706;font-size:calc(10px*var(--zf,1))}
input[type=checkbox]{cursor:pointer}
input:not([type=checkbox]):not([type=radio]):not([type=button]):not([type=submit]),textarea{cursor:text!important}
input,select,textarea{cursor:auto}
input[type=text],input[type=number],input[type=password],input[type=time],input[type=date],textarea{cursor:text!important}
.tl-legend{display:flex;gap:12px;padding:2px 4px;font-size:calc(12.5px*var(--zf,1));color:var(--gray);flex-wrap:wrap;align-items:center}
.tl-legend span{display:flex;align-items:center;gap:4px}
.tl-legend i{display:inline-block;width:15px;height:13px;border-radius:2px;flex-shrink:0}
select{cursor:default}
.fr.big input{font-size:calc(16px*var(--zf,1));font-weight:700;padding:5px 6px;color:var(--green)}
.fr.ro input{background:#f8fafc;color:var(--gray)}
/* Timeline */
.tl-wrap{background:var(--card);border-radius:7px;padding:7px 8px;border:1px solid var(--border)}
.tl-wrap h5{font-size:calc(11.25px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.7px;margin-bottom:4px}
/* RIGHT recap col (wider) */
.recap-col{width:280px;flex-shrink:0;border-left:1px solid var(--border);background:var(--card);display:flex;flex-direction:column;overflow:hidden}
.recap-hdr{font-size:calc(10px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);letter-spacing:.6px;padding:8px 8px 4px}
.recap-body{flex:1;overflow-y:auto;padding:0 6px 6px}
.si{display:flex;align-items:center;gap:5px;padding:3px 0;border-bottom:1px solid var(--border);font-size:calc(11px*var(--zf,1))}
.si:last-child{border:none}
.sdot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.si-nm{flex:1;font-weight:600;font-size:calc(10px*var(--zf,1));line-height:1.2}
.si-dur{font-size:calc(9px*var(--zf,1));color:var(--gray);white-space:nowrap}
.btn-edit{background:none;border:none;cursor:pointer;font-size:calc(11px*var(--zf,1));color:#6366f1;padding:1px 3px;border-radius:2px}
/* TRS gauge */
.gauge-box{padding:6px;border-top:1px solid var(--border);text-align:center}
.gauge-lbl{font-size:calc(11px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;margin-top:3px}
.cs{width:calc(36px*var(--zf,1));height:calc(46px*var(--zf,1));border:2px solid #cbd5e1;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:calc(22px*var(--zf,1));font-weight:900;font-family:monospace;color:#c8d5e0;background:#f8fafc;transition:all .12s;user-select:none;cursor:default}
.cs.cs-filled{color:#2563eb;border-color:#3b82f6;background:#eff6ff}
.cs.cs-active{border-color:#4f46e5;box-shadow:0 0 0 3px rgba(99,102,241,.2)}
.cs.cs-done{border-color:#16a34a;background:#f0fdf4}
/* Active stops bottom bar — chips */
#stop-bottom,#stop-bottom-main{display:none;background:#7f0000;color:#fff;padding:8px 14px;align-items:center;gap:8px;flex-shrink:0;border-top:2px solid #b91c1c;flex-wrap:wrap}
#stop-bottom.on,#stop-bottom-main.on{display:flex}
.stop-chip{display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.28);border-radius:8px;padding:6px 10px;border:1px solid rgba(255,255,255,.2)}
.chip-lbl{font-weight:800;font-size:calc(13px*var(--zf,1));white-space:nowrap}
.chip-tim{font-size:calc(18px*var(--zf,1));font-weight:800;font-variant-numeric:tabular-nums;min-width:52px;text-align:right}
.btn-endstop{background:#16a34a;color:#fff;border:none;border-radius:6px;padding:6px 12px;font-size:calc(12px*var(--zf,1));font-weight:700;cursor:pointer;white-space:nowrap}
.btn-endstop:hover{filter:brightness(.9)}
/* KPI shift cards */
.shift-kpis{display:flex;gap:8px;padding:10px 14px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0}
.skpi{flex:1;background:var(--bg);border-radius:var(--radius);padding:10px;border:1px solid var(--border);text-align:center}
.skpi.current{flex:2;background:var(--navy);color:#fff;border-color:var(--navy2);box-shadow:var(--shadow)}
.skpi.current .sk-lbl{color:rgba(255,255,255,.7)}.skpi.current .sk-val{color:#93c5fd;font-size:calc(26px*var(--zf,1))}.skpi.current .sk-sub{color:rgba(255,255,255,.75)}
.sk-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);letter-spacing:.7px;margin-bottom:3px}
.sk-val{font-size:calc(20px*var(--zf,1));font-weight:800;color:var(--navy);line-height:1}.sk-sub{font-size:calc(10px*var(--zf,1));color:var(--gray);margin-top:3px}
/* Merged table row types */
.row-prod td{background:#f0fdf4}.row-evt td{background:#fff7ed}
.row-prod:hover td,.row-evt:hover td{filter:brightness(.96)}
.row-tag{display:inline-block;padding:1px 6px;border-radius:4px;font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase}
.tag-p{background:#bbf7d0;color:#166534}.tag-e{background:#fed7aa;color:#9a3412}.tag-n{background:#bfdbfe;color:#1e40af}

/* ── FIN DE POSTE ── */
#v-finposte{padding:0}
.fp-scroll{flex:1;overflow-y:auto;padding:14px}
.fp-top{text-align:center;padding-bottom:10px}
.fp-top h2{font-size:calc(22px*var(--zf,1));font-weight:800;color:var(--navy)}
.fp-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px}
.fp-card{background:var(--card);border-radius:var(--radius);padding:12px;text-align:center;box-shadow:var(--shadow);border:1px solid var(--border)}
.fp-big{font-size:calc(28px*var(--zf,1));font-weight:800;color:var(--navy)}
.fp-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;margin-top:3px}
.fp-acts{display:flex;justify-content:center;gap:10px;padding:14px 0}
.fp-tbl{width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1));margin-bottom:12px}
.fp-tbl th{background:var(--navy);color:#fff;padding:5px 8px;font-size:calc(9px*var(--zf,1));text-align:left;text-transform:uppercase}
.fp-tbl td{padding:5px 8px;border-bottom:1px solid var(--border)}

/* ── HISTORY ── */
#v-history{padding:0}

/* ── SETTINGS ── */
#v-settings{padding:0;overflow:hidden}
#settings-lock{flex:1;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px}
.lock-card{background:var(--card);border-radius:12px;padding:28px;width:100%;max-width:340px;text-align:center;box-shadow:var(--shadow)}
.lock-card h3{color:var(--navy);font-size:calc(18px*var(--zf,1));font-weight:800;margin-bottom:8px}
.lock-card p{color:var(--gray);font-size:calc(12px*var(--zf,1));margin-bottom:16px}
#v-settings-content{flex:1;overflow-y:auto;padding:14px;display:none}
.ss{background:var(--card);border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:var(--shadow);border:1px solid var(--border)}
.ss h3{font-size:calc(12px*var(--zf,1));font-weight:700;margin-bottom:10px;color:var(--navy)}
.pr{display:flex;align-items:center;gap:6px;margin-bottom:6px;padding:5px;border-radius:5px;background:var(--bg)}
.pr .pn{font-weight:600;min-width:110px;font-size:calc(11px*var(--zf,1))}
.pr input{flex:1;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1))}
.btn-eye{background:none;border:none;cursor:pointer;color:var(--gray);font-size:calc(12px*var(--zf,1));padding:2px}
.day-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:6px}
.day-box{background:var(--bg);border-radius:4px;padding:4px;text-align:center}
.day-lbl{font-size:calc(8px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px}
.day-box input{width:100%;padding:2px;border:1px solid var(--border);border-radius:3px;font-size:calc(10px*var(--zf,1));text-align:center}
.model-card{border:1px solid var(--border);border-radius:7px;padding:10px;margin-bottom:8px}
.mch{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.mch input{flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;padding:4px 6px;border:1px solid var(--border);border-radius:4px}

/* ── MODALS ── */
.modal{display:none}
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:300;align-items:center;justify-content:center}
.overlay.on{display:flex}
.mbox{background:var(--card);border-radius:12px;width:90%;max-width:520px;box-shadow:0 20px 60px rgba(0,0,0,.3);max-height:92vh;display:flex;flex-direction:column}
.mbox.wide{max-width:1100px}
.mhdr{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--border);flex-shrink:0}
.mhdr.red{background:#b91c1c;border-radius:12px 12px 0 0}
.mhdr h2{font-size:calc(15px*var(--zf,1));font-weight:700;color:var(--navy)}
.mhdr.red h2,.mhdr.red button{color:#fff}
.mbody{flex:1;overflow-y:auto;padding:14px}
.mftr{display:flex;gap:8px;justify-content:flex-end;padding:12px 14px;border-top:1px solid var(--border);flex-shrink:0}
.btn{padding:7px 14px;border:none;border-radius:6px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;gap:4px}
.btn:hover{filter:brightness(.9)}
.btn-prim{background:var(--navy);color:#fff}
.btn-danger{background:var(--red);color:#fff}
.btn-ok{background:var(--green);color:#fff}
.btn-sec{background:var(--lgray);color:var(--text)}
.btn-amber{background:var(--amber);color:#fff}
.btn-lg{font-size:calc(14px*var(--zf,1));padding:10px 20px}

/* Stop modal */
.stop-section-lbl{font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;letter-spacing:.8px;margin:10px 0 5px;color:var(--navy)}
.stop-section-lbl:first-child{margin-top:0}
.stops-grid{display:grid;gap:5px;margin-bottom:6px}
.stops-grid.ratt{grid-template-columns:repeat(3,1fr)}
.stops-grid.pb{grid-template-columns:repeat(4,1fr)}
.stop-btn{border:none;border-radius:8px;padding:8px 5px;cursor:pointer;font-size:calc(11px*var(--zf,1));font-weight:700;color:#fff;text-align:center;transition:all .12s;box-shadow:0 2px 0 rgba(0,0,0,.2)}
.stop-btn:active{transform:translateY(2px);box-shadow:none}
.stop-btn.ratt{background:#7c3aed}
.stop-btn.pb{background:#b91c1c}
.custom-row{display:flex;gap:6px;margin-top:4px}
.custom-row input{flex:1;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1));outline:none}
.custom-row input:focus{border-color:var(--navy)}

/* End prod modal */
.ep-grid{display:flex;flex-direction:column;gap:4px;margin-bottom:12px}
.ep-stat{display:flex;align-items:center;justify-content:space-between;padding:7px 12px;background:var(--bg);border-radius:6px;border-left:3px solid #c7d2fe}
.ep-stat .val{font-size:calc(16px*var(--zf,1));font-weight:800;color:var(--navy);white-space:nowrap}
.ep-stat .lbl{font-size:calc(10px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;margin-right:10px}
.ep-tbl{width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1))}
.ep-tbl th{text-align:left;padding:4px 6px;background:var(--bg);font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray)}
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
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;margin-bottom:16px">
      <div style="background:#fff;border-radius:11px;padding:3px 8px;box-shadow:0 2px 10px rgba(0,0,0,.35);display:inline-flex;align-items:center"><img src="data:image/png;base64,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" height="46" style="display:block;border-radius:8px" alt="DODO"></div>
      <div class="lc-h1" style="margin-bottom:0;white-space:nowrap;font-size:calc(22px*var(--zf,1));letter-spacing:.5px">PRODUCTION ORC1</div>
    </div>
    <div class="lf">
      <label>Pilote</label>
      <select id="ln-pilot"><option value="">-- Choisir --</option></select>
    </div>
    <div class="lf">
      <label>Modèle horaire (Poste)</label>
      <select id="ln-model" onchange="onLoginModelChange()"><option value="">-- Choisir --</option></select>
    </div>
    <!-- Horaires du jour -->
    <div id="ln-model-info" style="display:none;background:#f0f9ff;border:1px solid #bae6fd;border-radius:7px;padding:8px 10px;font-size:calc(12px*var(--zf,1));margin-bottom:6px">
      <div style="color:#0369a1;font-weight:700;margin-bottom:3px">Horaires aujourd'hui :</div>
      <div id="ln-model-times" style="font-size:calc(15px*var(--zf,1));font-weight:800;color:#0c4a6e;margin-bottom:5px"></div>
      <div id="ln-model-modify" style="display:none;margin-bottom:5px">
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <label style="font-size:calc(11px*var(--zf,1));font-weight:600">Début <input type="time" id="ln-new-debut" style="padding:3px 6px;border:1.5px solid #bae6fd;border-radius:5px;font-size:calc(12px*var(--zf,1))"></label>
          <label style="font-size:calc(11px*var(--zf,1));font-weight:600">Fin <input type="time" id="ln-new-fin" style="padding:3px 6px;border:1.5px solid #bae6fd;border-radius:5px;font-size:calc(12px*var(--zf,1))"></label>
        </div>
      </div>
      <button style="font-size:calc(11px*var(--zf,1));padding:3px 8px;background:none;border:1px solid #0369a1;border-radius:5px;color:#0369a1;cursor:pointer" onclick="toggleLoginModelModify()">✏ Modifier les horaires d'aujourd'hui</button>
    </div>
    <div class="lf">
      <label>Mot de passe</label>
      <input type="password" id="ln-pw" placeholder="••••" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="btn-login" onclick="doLogin()">Valider</button>
    <div class="ln-err" id="ln-err"></div>
    <div style="margin-top:12px;border-top:1px solid rgba(255,255,255,.15);padding-top:12px;text-align:center">
      <button onclick="doGuestLogin()" style="background:rgba(255,255,255,.12);color:#cbd5e1;border:1px solid rgba(255,255,255,.25);border-radius:7px;padding:7px 18px;font-size:calc(12px*var(--zf,1));cursor:pointer;width:100%;font-weight:600">👁 Consulter en tant qu'invité</button>
    </div>
    <div style="margin-top:10px;text-align:center;font-size:calc(11px*var(--zf,1));color:#94a3b8">
      Prod bloquée ? <a href="/reset" style="color:#dc2626;font-weight:700">Cliquer ici pour réinitialiser</a>
    </div>
    <div style="margin-top:10px;border-top:1px solid #e5e7eb;padding-top:8px">
      <div style="text-align:right;margin-bottom:4px">
        <button id="btn-excel-toggle" onclick="toggleExcelPanel()" style="background:none;border:none;cursor:pointer;color:#64748b;font-size:calc(12px*var(--zf,1));padding:2px 6px">&#x1F4C1; Excel</button>
      </div>
      <div id="excel-panel" style="display:none;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px">
        <div style="margin-bottom:8px">
          <label style="color:#475569;font-size:calc(11px*var(--zf,1));display:block;margin-bottom:3px">Chemin du fichier Excel</label>
          <input id="excel-path-inp" type="text" placeholder="C:\...\fichier.xlsx" style="width:100%;box-sizing:border-box;padding:7px 9px;border-radius:5px;border:1px solid #cbd5e1;font-size:calc(11px*var(--zf,1));outline:none;color:#1e293b">
        </div>
        <div style="margin-bottom:8px">
          <label style="color:#475569;font-size:calc(11px*var(--zf,1));display:block;margin-bottom:3px">Mot de passe superviseur</label>
          <input id="excel-pw-inp" type="password" placeholder="1234" style="width:100%;box-sizing:border-box;padding:7px 9px;border-radius:5px;border:1px solid #cbd5e1;font-size:calc(11px*var(--zf,1));outline:none;color:#1e293b">
        </div>
        <div id="excel-panel-err" style="color:#dc2626;font-size:calc(11px*var(--zf,1));min-height:14px;margin-bottom:6px"></div>
        <div style="display:flex;gap:6px;justify-content:flex-end">
          <button onclick="toggleExcelPanel()" style="padding:5px 12px;border-radius:5px;border:1px solid #cbd5e1;background:#fff;color:#64748b;cursor:pointer;font-size:calc(11px*var(--zf,1))">Fermer</button>
          <button onclick="saveExcelPath()" style="padding:5px 12px;border-radius:5px;border:none;background:#1d4ed8;color:#fff;cursor:pointer;font-size:calc(11px*var(--zf,1));font-weight:600">OK</button>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ════ APP ════ -->
<div id="app" class="hidden" style="display:none;flex:1;flex-direction:column;overflow:hidden">
  <div id="app-hdr">
    <div style="display:inline-flex;align-items:center;align-self:center;padding:4px 0;flex-shrink:0"><img src="data:image/png;base64,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" height="52" style="display:block;border-radius:8px;filter:drop-shadow(0 2px 8px rgba(0,0,0,.18))" alt="DODO"></div>
    <div class="hdr-logo">PRODUCTION ORC</div>
    <div class="hdr-tabs">
      <button class="htab on" id="ht-main" onclick="goTab('main')">Accueil</button>
      <button class="htab prod-on" id="ht-prod" onclick="goTab('prod')">▶ Prod en cours</button>
      <span id="ht-guest-badge" style="display:none;font-size:calc(11px*var(--zf,1));font-weight:700;color:#94a3b8;padding:4px 10px;border:1px solid #94a3b8;border-radius:12px;margin-left:4px">👁 Invité</span>
      <button class="htab" id="ht-hist" onclick="goTab('history')">Historique</button>
      <button class="htab" id="ht-rapports" onclick="goTab('rapports')">📋 Rapports poste</button>
      <button class="htab" id="ht-rpt-jour" onclick="goTab('rpt-jour')">📅 Rapports jour</button>
      <button class="htab" id="ht-kpi" onclick="goTab('kpi')">📈 Evolution perf.</button>
    </div>
    <div id="hdr-right">
      <span id="hdr-pilot-lbl"></span>
      <button class="btn-sm btn-ghost" onclick="toggleZoomPop()" id="zoom-btn" style="font-size:calc(11px*var(--zf,1));display:flex;align-items:center;gap:4px" title="Zoom texte">🔍 Zoom</button>
      <button class="btn-sm btn-ghost" onclick="doLogout()" style="font-size:calc(11px*var(--zf,1))">Déconnexion</button>
      <button class="btn-sm btn-ghost" id="ht-cfg" onclick="goTab('settings')" style="font-size:calc(18px*var(--zf,1));padding:4px 8px;line-height:1" title="Paramètres">⚙</button>
    </div>
  </div>
  <!-- Zoom popover -->
  <div id="zoom-pop" style="display:none;position:fixed;top:44px;right:80px;z-index:9000;background:#1e293b;border:1px solid #334155;border-radius:10px;padding:12px 16px;box-shadow:0 4px 20px rgba(0,0,0,.4);min-width:220px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
      <span style="color:#fff;font-size:calc(12px*var(--zf,1));font-weight:700">🔍 Zoom texte</span>
      <button onclick="resetZoom()" style="background:none;border:1px solid #475569;border-radius:5px;color:#94a3b8;font-size:calc(10px*var(--zf,1));padding:2px 7px;cursor:pointer">Reset</button>
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      <button onclick="stepZoom(-5)" style="background:#334155;border:none;border-radius:5px;color:#fff;font-size:calc(14px*var(--zf,1));width:26px;height:26px;cursor:pointer;line-height:1">−</button>
      <input type="range" id="zoom-slider" min="70" max="150" value="135" step="5"
        oninput="applyZoom(+this.value)"
        style="flex:1;accent-color:#3b82f6;cursor:pointer">
      <button onclick="stepZoom(+5)" style="background:#334155;border:none;border-radius:5px;color:#fff;font-size:calc(14px*var(--zf,1));width:26px;height:26px;cursor:pointer;line-height:1">+</button>
    </div>
    <div style="text-align:center;margin-top:6px;color:#93c5fd;font-size:calc(13px*var(--zf,1));font-weight:700" id="zoom-val">135%</div>
  </div>
  <div id="alert-strip"></div>
  <div id="degrade-strip" style="display:none;background:#854d0e;color:#fef9c3;text-align:center;padding:4px;font-weight:700;font-size:calc(12px*var(--zf,1));flex-shrink:0;animation:blink .85s step-start infinite"></div>

  <!-- ════ MAIN VIEW ════ -->
  <div id="v-main" class="view" style="flex-direction:column">
    <!-- Bannière prod en cours (visible si prod_active mais sur vue accueil) -->
    <div id="main-prod-banner" style="display:none;background:#1e293b;color:#fff;padding:8px 14px;font-size:calc(12px*var(--zf,1));align-items:center;gap:16px;cursor:pointer" onclick="goTab('prod')">
      <span style="font-weight:800;color:#86efac">▶ Prod en cours</span>
      <span>OF : <span id="mpb-of" style="font-weight:700">—</span></span>
      <span>Durée : <span id="mpb-dur" style="color:#67e8f9;font-weight:700">—</span></span>
      <span>Arrêts : <span id="mpb-stops" style="color:#fca5a5;font-weight:700">—</span></span>
    </div>
    <!-- Barre Excel occupé -->
    <div id="excel-busy-bar" style="display:none;background:#92400e;color:#fef3c7;padding:5px 14px;font-size:calc(11px*var(--zf,1));font-weight:700;text-align:center">
      ⚠ Fichier Excel ouvert par un autre programme — impossible de lire/écrire les données
    </div>
    <div class="main-hdr">
      <div class="mbtns" style="margin-left:0" id="main-action-btns">
        <button class="acc-btn acc-green" id="btn-start" onclick="doStartProd()"><span class="act-icon">▶</span><span>Démarrer production</span></button>
        <button class="acc-btn acc-red" onclick="openStopModal()"><span class="act-icon"><span class="stop-icon">🛑<span class="stop-icon-x">✕</span></span></span><span>Déclarer un arrêt</span></button>
        <button id="btn-degrade-acc" class="acc-btn acc-amber" onclick="toggleDegrade()"><span class="act-icon">🐌</span><span>Mode dégradé</span></button>
        <button class="acc-btn acc-green" onclick="doFinPoste()"><span class="act-icon">🏁</span><span>Fin de poste</span></button>
      </div>
    </div>
    <!-- KPI accueil — POSTE ACTUEL -->
    <div style="background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;padding:4px 8px;display:flex;gap:6px;align-items:stretch;flex-wrap:wrap">

      <!-- POSTE ACTUEL encart principal -->
      <div style="flex:1.5;min-width:220px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:4px 10px;display:flex;flex-direction:column;gap:3px">
        <!-- Titre + TRS jauge + valeur -->
        <div style="display:flex;align-items:center;gap:8px">
          <div style="flex-shrink:0;text-align:center">
            <svg viewBox="0 0 100 58" style="width:100px;display:block;margin:0 auto">
              <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="rgba(0,0,0,.12)" stroke-width="11" stroke-linecap="round"/>
              <path id="gauge-poste-acc-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="11" stroke-linecap="round" stroke-dasharray="0,132"/>
              <text x="50" y="46" text-anchor="middle" font-size="17" font-weight="800" fill="#15803d" id="gauge-poste-acc-pct">—</text>
            </svg>
          </div>
          <div style="flex:1;min-width:0">
            <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:#0369a1;text-transform:uppercase;letter-spacing:.5px">TRS du Poste</div>
            <div style="font-size:calc(13px*var(--zf,1));font-weight:700;color:#0369a1;margin-bottom:2px" id="gauge-poste-acc-lbl">—</div>
            <!-- Stats en grille uniforme -->
            <div style="display:grid;grid-template-columns:auto 1fr;gap:1px 6px;margin-top:3px;align-items:baseline">
              <span style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#15803d;white-space:nowrap">Prod</span>
              <span style="font-size:calc(12px*var(--zf,1));font-weight:900;color:#15803d;line-height:1.1" id="acc-prod-total">— pcs / — éq</span>
              <span style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#1e40af;white-space:nowrap">Nb OF</span>
              <span style="font-size:calc(12px*var(--zf,1));font-weight:900;color:#1e40af;line-height:1.1" id="acc-nb-of">0</span>
              <span style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#dc2626;white-space:nowrap">Arrêts</span>
              <span style="font-size:calc(12px*var(--zf,1));font-weight:900;color:#b91c1c;line-height:1.1" id="main-stat-arrets">0 min</span>
              <span style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;white-space:nowrap">Fonct.</span>
              <span style="font-size:calc(12px*var(--zf,1));font-weight:900;color:#15803d;line-height:1.1" id="main-stat-prod">0 min</span>
            </div>
          </div>
          <!-- Répartition temps (pie) -->
          <div style="flex-shrink:0;text-align:center;border-left:1px solid #bae6fd;padding-left:8px;min-width:110px">
            <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">Répartition temps</div>
            <svg id="pie-poste-acc" viewBox="0 0 130 130" style="width:160px;height:auto;display:block;margin:0 auto"></svg>
          </div>
        </div>
        <!-- Horaire temporaire -->
        <div id="acc-model-info" style="display:none;border-top:1px solid #bae6fd;padding-top:5px;margin-top:2px;font-size:calc(11px*var(--zf,1));color:#0369a1"></div>
      </div>

      <!-- Arrêts prévus — barres budget -->
      <div style="flex:2;min-width:190px;background:#fefce8;border:1px solid #fde68a;border-radius:10px;padding:4px 10px;display:flex;flex-direction:column">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
          <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:.5px">⏱ Arrêts prévus du poste</span>
          <button onclick="openBudgetOverrideModal()" id="btn-bov-acc" style="display:none;background:none;border:1px solid #92400e;border-radius:4px;color:#92400e;font-size:calc(10px*var(--zf,1));padding:1px 7px;cursor:pointer" title="Modifier le budget pour ce poste">✏️</button>
        </div>
        <div id="budget-bars-acc" style="flex:1"></div>
      </div>

      <!-- Poste précédent -->
      <div class="skpi" style="flex:1;min-width:100px;max-width:180px">
        <div class="sk-lbl" id="kpi1-lbl">Poste précédent</div>
        <div class="sk-val" id="kpi1-trs">--%</div>
        <div class="sk-sub" id="kpi1-date" style="font-size:calc(10px*var(--zf,1));opacity:.85"></div>
        <div class="sk-sub" id="kpi1-sub">0 OF</div>
      </div>
      <div class="skpi" style="flex:1;min-width:100px;max-width:180px">
        <div class="sk-lbl" id="kpi2-lbl">Avant-dernier</div>
        <div class="sk-val" id="kpi2-trs">--%</div>
        <div class="sk-sub" id="kpi2-date" style="font-size:calc(10px*var(--zf,1));opacity:.85"></div>
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
          <th>Type</th><th>OF</th><th>Fibre</th><th>Date</th><th>Poste</th><th>Pilote</th>
          <th>Début</th><th>Fin</th><th>Détails</th><th>Qté/Durée</th><th>TRS/Info</th><th>Commentaire</th><th>Actions</th>
        </tr></thead>
        <tbody id="main-body"></tbody>
      </table>
    </div>
    <!-- Active stops bottom bar — chips (accueil) -->
    <div id="stop-bottom-main">
      <div style="font-size:calc(10px*var(--zf,1));opacity:.7;font-weight:700;white-space:nowrap">EN COURS :</div>
      <div id="stop-chips-main" style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;flex:1"></div>
    </div>
  </div>

  <!-- ════ PRODUCTION VIEW ════ -->
  <div id="v-prod" class="view" style="flex-direction:column">
    <!-- Pilot/OF banner -->
    <div class="pob" style="display:none">
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
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1))" id="pob-poste">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Départ OF</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1));color:#fbbf24" id="pob-of-start">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Fin OF</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1));color:#94a3b8" id="rc-fin">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Durée OF</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1));color:#0891b2" id="rc-duree">—</div>
      </div>
      <div class="pob-item trs">
        <div class="pob-lbl">TRS estimé</div>
        <div class="pob-val" id="pob-trs">—</div>
      </div>
    </div>
    <!-- Status bar -->
    <div class="sbar">
      <div class="sc"><div class="sc-lbl">⏱ Durée de l'OF en cours</div><div class="sc-val green" id="sc-of">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">⛔ Arrêt de l'OF en cours</div><div class="sc-val red" id="sc-stops">00:00:00</div></div>
      <div class="sc" id="sc-deg-wrap" style="display:none"><div class="sc-lbl">🐌 Dégradé</div><div class="sc-val amber" id="sc-deg">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">⛔ Arrêt du poste entier</div><div class="sc-val red" id="sc-all-stops">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">🎯 Objectif instantané OF en cours</div><div class="sc-val" id="sc-theo" style="color:var(--blue)">—</div></div>
    </div>
    <!-- Body -->
    <div class="prod-body">
      <!-- LEFT: form (now full center, no act-col) -->
      <div class="form-col">
        <div class="form-3col">
          <!-- Zone Identification -->
          <div class="fzone zi">
            <h4>📋 Identification</h4>
            <div class="fr"><label>N° OF *</label><input id="f-of_num" oninput="scheduleAutoSave()" onfocus="openCodeInput('of_num','N° OF','9')"></div>
            <div class="fr"><label>Code Produit</label><input id="f-code_prod" oninput="scheduleAutoSave()" onfocus="openCodeInput('code_prod','Code Produit')"></div>
            <div class="fr"><label>Type Produit</label><select id="f-type_prod" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr ro"><label>Date</label><input id="f-date" readonly></div>
            <div class="fr ro"><label>Poste</label><input id="f-poste" readonly></div>
            <div class="fr ro"><label>Pilote</label><input id="f-pilote" readonly></div>
            <div class="fr"><label>Co-Pilote</label><select id="f-copilote" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Nb Personnes</label><input id="f-nb_pers" type="number" min="1" value="10" oninput="scheduleAutoSave()"></div>
          </div>
          <!-- Zone Production -->
          <div class="fzone zp">
            <h4>🏭 Production</h4>
            <div class="fr big"><label>Qté Fabriquée *</label><input id="f-qte_fab" type="number" min="0" placeholder="0" oninput="scheduleAutoSave()"></div>
            <div class="fr big"><label>Qté Emballée</label><input id="f-qte_emb" type="number" min="0" placeholder="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Lots de 2</label><select id="f-kit" onchange="scheduleAutoSave()"><option value="">Non</option><option value="oui">Oui</option></select></div>
            <div class="fr"><label>Poids Garnissage (g)</label><input id="f-poids" type="number" min="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Taille</label><select id="f-taille" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Fibre</label><select id="f-fibre" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Traca Fibre</label><input type="text" id="f-traca" style="display:none"><div id="f-traca-ui" style="display:flex;flex-direction:column;gap:2px;margin-bottom:3px"></div><button type="button" onclick="addTracaRow()" style="align-self:flex-start;background:#eff6ff;border:1.5px solid #bfdbfe;border-radius:5px;color:#1d4ed8;font-size:calc(10px*var(--zf,1));font-weight:700;padding:3px 10px;cursor:pointer">Ajouter une autre traça</button></div>
            <div class="fr" style="display:none"><input id="f-of_taie" oninput="scheduleAutoSave()"></div>
            <div class="fr" style="display:none"><input id="f-duree_mq_mp" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr" style="display:none"><input id="f-manquant_pers" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
          </div>
          <!-- Zone Qualité -->
          <div class="fzone zq">
            <h4>✅ Qualité</h4>
            <div class="fr" style="display:none"><input id="f-qte_init_taie" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Code Taie</label><input id="f-ref_taie" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Taie 2nd Choix</label><input id="f-nb_taie2_choix" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Défaut Couture</label><input id="f-nb_def_cout" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Taie</label><input id="f-mq_taie" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Housse/Encart</label><input id="f-mq_housse_encart" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>PP cousu et emballé</label><input id="f-nb_pp_cousue" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr comment-big"><label>💬 Commentaire</label><textarea id="f-comment" oninput="scheduleAutoSave()" placeholder="Commentaire libre…"></textarea></div>
          </div>
        </div>
        <!-- Timeline 4h -->
        <div class="tl-wrap">
          <h5>Timeline — 4 dernières heures</h5>
          <svg id="tl-svg" viewBox="0 0 800 64" preserveAspectRatio="none" style="width:100%;height:64px;display:block">
            <rect x="0" y="4" width="800" height="35" fill="#e2e8f0" rx="4"/>
          </svg>
          <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f97316"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#8b5cf6"></i>Réunion</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span><span><i style="background:repeating-linear-gradient(45deg,#16a34a,#16a34a 4px,#fef08a 4px,#fef08a 8px)"></i>Prod dégradé</span></div>
        </div>
        <!-- Action buttons row (below timeline) -->
        <div class="prod-act-row" style="justify-content:center">
          <button class="act-btn act-btn-sm act-stop" style="flex:1" onclick="openStopModal()"><span class="act-icon"><span class="stop-icon">🛑<span class="stop-icon-x">✕</span></span></span><span>Déclarer un arrêt</span></button>
          <button id="btn-degrade-prod" class="act-btn act-btn-sm" onclick="toggleDegrade()" style="flex:1;background:radial-gradient(ellipse at 50% 25%,#fde68a 0%,#f59e0b 55%,#92400e 100%);color:#fff;font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.4);border:none"><span class="act-icon">🐌</span><span>Mode dégradé</span></button>
          <button class="act-btn act-btn-sm act-nett" style="flex:1" onclick="doNettoyage()"><span class="act-icon">🧹</span><span>Nettoyage</span></button>
          <button class="act-btn act-btn-sm act-pause" id="btn-pause" style="flex:1" onclick="doPause()"><span class="act-icon">☕</span><span>Pause</span></button>
          <button class="act-btn act-btn-sm" id="btn-reunion" onclick="doReunion()" style="flex:1;background:radial-gradient(ellipse at 50% 25%,#c4b5fd 0%,#8b5cf6 55%,#5b21b6 100%);color:#fff;font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.4);border:none"><span class="act-icon">🗣️</span><span>Réunion</span></button>
        </div>
      </div>
      <!-- RIGHT: recap arrêts + gauges + pie charts -->
      <div class="recap-col" style="width:310px">
        <div class="recap-hdr">Arrêts / pauses</div>
        <div class="recap-body" id="recap-list" style="max-height:120px;flex:none;overflow-y:auto"></div>
        <!-- Budget arrêts prévus -->
        <div style="padding:5px 8px;border-top:1px solid var(--border);flex-shrink:0;background:#fffbeb">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
            <span style="font-size:calc(9px*var(--zf,1));font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:.5px">⏱ Arrêts prévus du poste</span>
            <button onclick="openBudgetOverrideModal()" id="btn-bov-prod" style="display:none;background:none;border:1px solid #92400e;border-radius:4px;color:#92400e;font-size:calc(9px*var(--zf,1));padding:1px 6px;cursor:pointer" title="Modifier le budget pour ce poste">✏️</button>
          </div>
          <div id="budget-bars-prod"></div>
        </div>
        <!-- Gauge + Pie: OF uniquement -->
        <div style="padding:6px;border-top:1px solid var(--border);display:flex;flex-direction:column;align-items:center;flex-shrink:0">
          <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;margin-bottom:2px;letter-spacing:.4px">TRS — OF en cours</div>
          <svg viewBox="0 0 100 56" style="width:100%;max-width:140px">
            <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="10" stroke-linecap="round"/>
            <path id="gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="10" stroke-linecap="round" stroke-dasharray="0,1000"/>
            <text x="50" y="46" text-anchor="middle" font-size="17" font-weight="800" fill="#1a1f5e" id="gauge-pct">—</text>
          </svg>
          <div class="gauge-lbl" style="font-size:calc(9px*var(--zf,1))">TRS OF</div>
          <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-top:4px;margin-bottom:2px">Répartition temps OF</div>
          <svg id="pie-of" viewBox="0 0 130 150" style="width:100%;max-width:220px;height:auto;display:block;margin:0 auto"></svg>
        </div>
        <!-- Bottom action buttons -->
        <div style="padding:8px;border-top:1px solid var(--border);display:flex;flex-direction:column;gap:6px;flex-shrink:0;background:var(--card)">
          <button class="act-btn act-cancel" onclick="doCancelProd()" style="width:100%;min-height:44px;padding:8px 12px;font-size:calc(14px*var(--zf,1))"><span class="act-icon" style="font-size:calc(18px*var(--zf,1))">✖</span>Annuler prod</button>
          <button class="act-btn act-endprod" id="btn-endprod" onclick="doEndProdPreview()" title="Remplir le formulaire" style="width:100%;min-height:121px;padding:10px 12px;font-size:calc(20px*var(--zf,1));margin-top:8px"><span class="act-icon" style="font-size:calc(33px*var(--zf,1))">✅</span>Fin d'OF/prod</button>
        </div>
      </div>
    </div>
    <!-- Active stops bottom bar — chips -->
    <div id="stop-bottom">
      <div style="font-size:calc(10px*var(--zf,1));opacity:.7;font-weight:700;white-space:nowrap">EN COURS :</div>
      <div id="stop-chips" style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;flex:1"></div>
    </div>
  </div>

  <!-- ════ FIN DE POSTE ════ -->
  <div id="v-finposte" class="view" style="flex-direction:column;overflow:hidden">
    <!-- Barre titre -->
    <div style="background:var(--navy);color:#fff;padding:7px 14px;flex-shrink:0;display:flex;align-items:center;justify-content:space-between">
      <div>
        <div style="font-size:calc(15px*var(--zf,1));font-weight:800">🏁 Fin de poste</div>
        <div id="fp-who" style="font-size:calc(11px*var(--zf,1));opacity:.8"></div>
      </div>
      <div style="text-align:right">
        <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:#fbbf24" id="fp-horaire-display">—</div>
        <div id="fp-date" style="font-size:calc(11px*var(--zf,1));font-weight:700;opacity:.8"></div>
      </div>
    </div>
    <!-- Graphiques + KPI (en haut, compact) -->
    <div style="display:flex;gap:12px;padding:10px 14px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;align-items:center;flex-wrap:wrap">
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">TRS Poste</div>
        <svg id="fp-gauge" viewBox="0 0 100 58" style="width:200px;display:block;margin:0 auto">
          <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
          <path id="fp-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
          <text x="50" y="46" text-anchor="middle" font-size="14" font-weight="800" fill="#1a1f5e" id="fp-gauge-pct">--%</text>
        </svg>
        <div style="font-size:calc(12px*var(--zf,1));color:var(--gray);margin-top:2px" id="fp-shift-hours">—</div>
      </div>
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">Répartition</div>
        <svg id="fp-pie" viewBox="0 0 130 130" style="width:200px;height:200px;display:block;margin:0 auto"></svg>
      </div>
      <div style="flex:1;display:flex;flex-direction:column;gap:5px">
        <div style="display:none"><span id="fp-trs"></span><span id="fp-trs-of"></span></div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:5px">
          <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#059669;font-weight:900" id="fp-pieces">--</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Nb pièces prod.</div></div>
          <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#0891b2;font-weight:900" id="fp-eq">0</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Équivalence</div></div>
          <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#0369a1;font-weight:900" id="fp-cadence">--</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Cadence/h</div></div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:5px">
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1))" id="fp-ouverture">--</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Durée ouverture</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#16a34a" id="fp-prod-t">0 min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Durée prod</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#dc2626" id="fp-stop-t">0 min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts total</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#16a34a" id="fp-ded">0 min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts prévus</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#7c3aed" id="fp-nof">0</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Nb OF</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#8b5cf6" id="fp-chg-fibre">--</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Chg. fibre</div></div>
        </div>
      </div>
    </div>
    <!-- Timeline compact -->
    <div style="padding:5px 12px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:3px">Timeline du poste</div>
      <svg id="fp-tl" viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">
        <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
      </svg>
      <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f97316"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#8b5cf6"></i>Réunion</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
    </div>
    <!-- Corps défilant : productions + arrêts côte à côte -->
    <div style="flex:1;overflow-y:auto;padding:8px 12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div class="card" style="padding:8px">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:5px">Productions</div>
        <table class="fp-tbl" style="font-size:calc(10px*var(--zf,1))">
          <thead><tr><th>OF</th><th>Début</th><th>Fin</th><th>Taille</th><th>Qté</th><th>Éq</th><th>Durée</th><th>TRS%</th></tr></thead>
          <tbody id="fp-prods"></tbody>
        </table>
      </div>
      <div class="card" style="padding:8px">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:5px">Arrêts du poste</div>
        <div id="fp-stops-list" style="font-size:calc(11px*var(--zf,1))"></div>
      </div>
    </div>
    <!-- Boutons -->
    <div style="padding:8px 12px;background:var(--card);border-top:1px solid var(--border);flex-shrink:0;display:flex;gap:10px;justify-content:flex-end">
      <button class="btn btn-sec" onclick="goTab('main')">← Retour</button>
      <button onclick="confirmFinPoste()" style="background:#16a34a;color:#fff;border:none;border-radius:10px;font-size:calc(18px*var(--zf,1));font-weight:900;padding:18px 40px;cursor:pointer;box-shadow:0 4px 18px rgba(22,163,74,.4);letter-spacing:.3px;transition:all .15s" onmouseover="this.style.filter='brightness(.9)'" onmouseout="this.style.filter=''">✅ Confirmer fin de poste &amp; Déconnexion</button>
    </div>
  </div>

  <!-- ════ MODAL : horaires de poste ════ -->
  <div id="m-fp-horaires" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:500;align-items:center;justify-content:center">
    <div class="card" style="width:340px;padding:20px;background:#fff;border-radius:12px">
      <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:14px">✏ Horaires de mon poste</div>
      <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Ces horaires servent uniquement au calcul du TRS de poste (non sauvegardés).</div>
      <div class="lf" style="margin-bottom:10px">
        <label>Début de poste</label>
        <input type="datetime-local" id="fp-debut-dt" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))">
      </div>
      <div class="lf" style="margin-bottom:14px">
        <label>Fin de poste</label>
        <input type="datetime-local" id="fp-fin-dt" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))">
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="closeM('m-fp-horaires')">Annuler</button>
        <button class="btn btn-prim" onclick="applyFPHoraires()">✓ Appliquer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL PRÉ-POSTE (gaps non déclarés avant 1er OF ou entre OFs) ════ -->
  <div id="m-preshift" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:601;align-items:center;justify-content:center">
    <div class="card" style="width:min(880px,95vw);padding:20px;background:#fff;border-radius:12px;border-top:4px solid var(--red)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
        <div id="ps-title" style="font-size:calc(15px*var(--zf,1));font-weight:800;color:var(--navy)">⚠ Période non déclarée</div>
        <div id="ps-counter" style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#94a3b8;background:#f1f5f9;border-radius:8px;padding:3px 10px"></div>
      </div>
      <div id="ps-text" style="font-size:calc(13px*var(--zf,1));color:var(--red);font-weight:700;margin-bottom:14px"></div>
      <input type="hidden" id="ps-start-iso">
      <input type="hidden" id="ps-gap-s">
      <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:var(--gray);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Déclarer comme :</div>
      <div id="ps-stop-btns" style="margin-bottom:10px"></div>
      <div style="margin-bottom:10px;display:flex;gap:6px">
        <input id="ps-custom" placeholder="Ou saisir librement…" style="flex:1;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))" onkeydown="if(event.key==='Enter')confirmPsAsStop()">
        <button class="btn btn-prim" onclick="confirmPsAsStop()" style="flex-shrink:0">✓ Valider</button>
      </div>
      <hr style="border:none;border-top:1px solid var(--border);margin-bottom:14px">
      <div style="display:flex;gap:8px;align-items:center;justify-content:space-between">
        <div id="ps-backdate-row" style="flex:1">
          <button class="btn btn-green" style="width:100%;text-align:left;padding:10px 14px;font-size:calc(13px*var(--zf,1))" onclick="psChooseBackdate()">
            ↩ Rétrodater le début de cet OF à <span id="ps-backdate-time" style="font-weight:800">--h--</span>
          </button>
        </div>
        <button class="btn btn-sec" style="flex-shrink:0;padding:10px 18px;font-size:calc(13px*var(--zf,1))" onclick="psIgnorer()">Ignorer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL INTERPOSTE ════ -->
  <div id="m-interposte" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:min(880px,95vw);padding:20px;background:#fff;border-radius:12px;border-top:4px solid var(--amber)">
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:4px">⏱ Temps hors production</div>
      <div id="ip-duration" style="font-size:calc(13px*var(--zf,1));color:var(--amber);font-weight:700;margin-bottom:8px"></div>
      <div style="display:flex;gap:10px;margin-bottom:12px;align-items:flex-end">
        <div style="flex:1">
          <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Début</label>
          <input type="time" id="ip-debut" style="width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
        </div>
        <div style="flex:1">
          <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Fin</label>
          <input type="time" id="ip-fin" style="width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
        </div>
      </div>
      <div style="font-size:calc(12px*var(--zf,1));color:var(--gray);margin-bottom:10px">Que s'est-il passé pendant cette période ?</div>
      <div id="ip-arrprev-group" style="display:none;margin-bottom:10px"></div>
      <div id="ip-btns" style="display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px"></div>
      <div style="margin-bottom:10px">
        <input id="ip-custom" placeholder="Ou saisir librement…" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))">
      </div>
      <div style="margin-bottom:10px">
        <input id="ip-comment" placeholder="Commentaire (optionnel)" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1))">
      </div>
      <!-- Formulaire modification plage horaire (caché par défaut) -->
      <div id="ip-model-form" style="display:none;background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:8px;padding:10px;margin-bottom:10px">
        <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#0369a1;margin-bottom:8px;text-transform:uppercase">Modifier la plage horaire du poste</div>
        <div style="display:flex;gap:10px;margin-bottom:8px">
          <div style="flex:1">
            <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Début modèle</label>
            <input type="time" id="ip-model-debut" style="width:100%;padding:6px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
          </div>
          <div style="flex:1">
            <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Fin modèle</label>
            <input type="time" id="ip-model-fin" style="width:100%;padding:6px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
          </div>
        </div>
        <div style="display:flex;gap:6px;justify-content:flex-end">
          <button class="btn btn-sec" style="font-size:calc(11px*var(--zf,1));padding:4px 10px" onclick="document.getElementById('ip-model-form').style.display='none'">Annuler</button>
          <button class="btn btn-prim" style="font-size:calc(11px*var(--zf,1));padding:4px 10px" onclick="ipConfirmModifyModel()">✓ Appliquer</button>
        </div>
      </div>
      <div style="display:flex;gap:8px;justify-content:space-between;align-items:center">
        <button class="btn btn-ghost" style="font-size:calc(11px*var(--zf,1));padding:4px 10px;color:#0369a1;border-color:#bae6fd" onclick="ipShowModifyModel()">✏ Modifier plage horaire</button>
        <div style="display:flex;gap:8px">
          <button class="btn btn-sec" onclick="skipInterposte()">Ignorer</button>
          <button class="btn btn-prim" onclick="confirmInterposte()">✓ Valider</button>
        </div>
      </div>
    </div>
  </div>

  <!-- ════ MODAL ARRÊTS MANQUANTS ════ -->
  <div id="m-missing-decl" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:480px;max-height:85vh;overflow-y:auto;padding:20px;background:#fff;border-radius:12px;border-top:4px solid #f59e0b">
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:#92400e;margin-bottom:4px">⚠ Arrêts non déclarés</div>
      <div style="font-size:calc(12px*var(--zf,1));color:#78350f;margin-bottom:14px">Vous n'avez pas déclaré les arrêts prévus suivants. Souhaitez-vous les ajouter avant de terminer le poste ?</div>
      <div id="md-rows" style="display:flex;flex-direction:column;gap:10px;margin-bottom:14px"></div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="skipMissingDecl()">Ignorer et continuer</button>
        <button class="btn btn-prim" style="background:#f59e0b;border-color:#f59e0b" onclick="skipMissingDecl()">✓ Continuer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL ÉCART FIN DE POSTE ════ -->
  <div id="m-ecart-poste" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:min(1160px,95vw);max-height:92vh;overflow-y:auto;padding:20px;background:#fff;border-radius:12px;border-top:4px solid #dc2626">
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:6px">📊 Réconciliation fin de poste</div>
      <div id="ecart-guide" style="font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;line-height:1.5"></div>
      <div id="ecart-info" style="background:#f8fafc;border:1px solid var(--border);border-radius:8px;padding:10px 12px;margin-bottom:12px;font-size:calc(12px*var(--zf,1))"></div>
      <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);letter-spacing:.06em;margin-bottom:6px">Plages non justifiées</div>
      <div id="ecart-gaps" style="margin-bottom:14px"></div>
      <div id="ecart-of-panel" style="margin-bottom:10px">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);letter-spacing:.06em;margin-bottom:6px">Modifier les horaires des OFs</div>
        <div id="ecart-of-list" style="display:flex;flex-direction:column;gap:6px"></div>
      </div>
      <!-- Modifier la plage horaire du poste -->
      <div id="ecart-plage-panel" style="margin-bottom:14px">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#0369a1;letter-spacing:.06em;margin-bottom:6px">Modifier la plage horaire de mon poste</div>
        <div style="background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:8px;padding:10px;display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
          <div>
            <label style="font-size:calc(10px*var(--zf,1));color:var(--gray);font-weight:600;display:block;margin-bottom:3px">Heure de début</label>
            <input type="time" id="ecart-plage-debut" style="padding:5px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700;width:110px">
          </div>
          <div>
            <label style="font-size:calc(10px*var(--zf,1));color:var(--gray);font-weight:600;display:block;margin-bottom:3px">Heure de fin</label>
            <input type="time" id="ecart-plage-fin" style="padding:5px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700;width:110px">
          </div>
          <button class="btn btn-prim" style="font-size:calc(12px*var(--zf,1));padding:6px 16px;background:#0369a1;border-color:#0369a1" onclick="saveEcartPlageHoraire()">✓ Appliquer</button>
          <span style="font-size:calc(10px*var(--zf,1));color:#64748b;align-self:center">Modifie uniquement pour ce poste aujourd'hui</span>
        </div>
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end;align-items:center">
        <button class="btn btn-ghost" style="font-size:calc(12px*var(--zf,1))" onclick="closeM('m-ecart-poste')">Annuler</button>
        <button id="ecart-valider-btn" class="btn btn-sec" onclick="skipEcartPoste()">Valider et terminer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL : "Autre" type d'arrêt ════ -->
  <div id="m-ecart-autre" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:700;align-items:center;justify-content:center">
    <div class="card" style="width:340px;padding:20px;background:#fff;border-radius:12px;border-top:4px solid #64748b">
      <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:10px">✏ Saisir le type d'arrêt</div>
      <input id="ecart-autre-label" placeholder="Ex : Changement outillage, Défaut matière…" style="width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:calc(13px*var(--zf,1));outline:none;margin-bottom:14px" onkeydown="if(event.key==='Enter')confirmEcartAutre()">
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="closeM('m-ecart-autre')">Annuler</button>
        <button class="btn btn-prim" onclick="confirmEcartAutre()">✓ Valider</button>
      </div>
    </div>
  </div>

<!-- Modal saisie code formaté (DDDDDD_DDD) -->
<div id="m-code-input" class="modal" style="position:fixed;inset:0;z-index:9999;backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);background:rgba(0,0,0,0.75);align-items:center;justify-content:center" onclick="if(event.target===this)closeM('m-code-input')">
  <div style="background:#fff;border-radius:14px;padding:24px 28px;width:90%;max-width:480px;text-align:center;box-shadow:0 30px 80px rgba(0,0,0,.5)">
    <div style="font-size:calc(13px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:8px;letter-spacing:.05em" id="code-input-lbl">CODE</div>
    <div id="code-slots" tabindex="0" onkeydown="_csKeydown(event)"
      style="display:flex;align-items:center;justify-content:center;gap:5px;margin-bottom:16px;outline:none;cursor:text">
      <div class="cs" id="cs-0">X</div><div class="cs" id="cs-1">X</div><div class="cs" id="cs-2">X</div>
      <div class="cs" id="cs-3">X</div><div class="cs" id="cs-4">X</div><div class="cs" id="cs-5">X</div>
      <div id="cs-sep" style="font-size:calc(26px*var(--zf,1));font-weight:900;color:#94a3b8;line-height:1;align-self:center;margin:0 3px">_</div>
      <div class="cs" id="cs-6">X</div><div class="cs" id="cs-7">X</div><div class="cs" id="cs-8">X</div>
    </div>
    <div id="cs-hint" style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:16px">Tapez les 6 premiers puis les 3 derniers chiffres</div>
    <div style="display:flex;gap:8px;justify-content:center">
      <button class="btn btn-ghost" onclick="closeM('m-code-input')">Annuler</button>
      <button class="btn btn-primary" onclick="_codeInputConfirm()">Confirmer ✓</button>
    </div>
  </div>
</div>

<!-- Modal détail OF -->
<div id="m-of-detail" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:650;align-items:center;justify-content:center;padding:12px" onclick="if(event.target===this)closeM('m-of-detail')">
  <div id="of-detail-box" style="width:min(1280px,98vw);height:92vh;background:#f8fafc;border-radius:18px;box-shadow:0 32px 100px rgba(0,0,0,.55),0 8px 24px rgba(0,0,0,.25);display:flex;flex-direction:column;overflow:hidden">
    <div id="of-detail-content" style="display:flex;flex-direction:column;min-height:0;flex:1;overflow:hidden"></div>
  </div>
</div>

  <!-- ════ RAPPORTS DES POSTES ════ -->
  <div id="v-rapports" class="view" style="flex-direction:column;overflow:hidden">
    <!-- Bandeau filtre date -->
    <div style="background:linear-gradient(180deg,#f8faff 0%,#fff 100%);border-bottom:2px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:8px;flex-shrink:0;flex-wrap:wrap;box-shadow:0 2px 6px rgba(0,0,0,.04)">
      <span style="font-size:calc(13px*var(--zf,1));font-weight:800;color:var(--navy);letter-spacing:.3px">📋 Rapports postes</span>
      <div style="display:flex;align-items:center;gap:6px;background:#fff;border:1.5px solid #c7d2fe;border-radius:10px;padding:4px 10px;box-shadow:0 1px 4px rgba(99,102,241,.12)">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#4338ca;text-transform:uppercase">📆 Du</span>
        <input type="date" id="rpt-from" onchange="loadRapports()" style="padding:3px 6px;border:none;border-radius:5px;font-size:calc(12px*var(--zf,1));color:#1e3a8a;font-weight:600;outline:none;background:transparent">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#4338ca;text-transform:uppercase">au</span>
        <input type="date" id="rpt-to" onchange="loadRapports()" style="padding:3px 6px;border:none;border-radius:5px;font-size:calc(12px*var(--zf,1));color:#1e3a8a;font-weight:600;outline:none;background:transparent">
      </div>
      <div style="margin-left:auto;display:flex;gap:6px">
        <button onclick="_rptSetLast7();loadRapports();" style="background:linear-gradient(180deg,#34d399,#059669);color:#fff;border:none;border-radius:8px;padding:5px 12px;font-size:calc(11px*var(--zf,1));font-weight:700;cursor:pointer;box-shadow:0 3px 8px rgba(5,150,105,.35),inset 0 1px 0 rgba(255,255,255,.18)">7 derniers jours</button>
      </div>
    </div>
    <div style="display:flex;flex:1;overflow:hidden;min-height:0">
      <!-- Panneau gauche : liste OU résumé KPI selon mode -->
      <div id="rpt-left-exe" style="width:280px;min-width:0;flex-shrink:0;border-right:1px solid var(--border);display:flex;flex-direction:column;background:#f8fafc;overflow:hidden;transition:width .25s ease">
        <!-- Mode liste (par défaut) -->
        <div id="rpt-left-list" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
          <div style="padding:10px 14px;font-size:calc(13px*var(--zf,1));font-weight:800;color:var(--navy);border-bottom:1px solid var(--border);flex-shrink:0;display:flex;align-items:center;justify-content:space-between">
            <span>📋 Tous les postes</span>
            <button class="btn btn-ghost" style="font-size:calc(11px*var(--zf,1));padding:3px 8px" onclick="reloadAndLoadRapports()">↺</button>
          </div>
          <div id="rpt-list" style="flex:1;overflow-y:auto">
            <div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Chargement…</div>
          </div>
        </div>
        <!-- Mode KPI (caché jusqu'à sélection d'un poste) -->
        <div id="rpt-left-kpi" style="display:none;flex-direction:column;flex:1;overflow-y:auto"></div>
      </div>
      <!-- Détail du rapport : tables + timeline -->
      <div id="rpt-detail" style="overflow-y:auto;flex:1;padding:0">
        <div style="padding:60px;text-align:center;color:var(--gray)">
          <div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">📋</div>
          <div style="font-size:calc(14px*var(--zf,1));font-weight:600">Sélectionner un poste dans la liste</div>
        </div>
      </div>
    </div>
  </div>

  <!-- ════ RAPPORTS JOUR ════ -->
  <div id="v-rpt-jour" class="view" style="flex-direction:column;overflow:hidden">
    <div id="rj-period-banner" style="background:linear-gradient(90deg,#1e3a8a,#2d3480);color:#fff;text-align:center;padding:9px 14px;font-size:calc(15px*var(--zf,1));font-weight:800;letter-spacing:.4px;flex-shrink:0;box-shadow:0 2px 8px rgba(30,58,138,.3)">📅 Rapport des 3 derniers postes</div>
    <div style="background:linear-gradient(180deg,#f8faff 0%,#fff 100%);border-bottom:2px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:8px;flex-shrink:0;flex-wrap:wrap;box-shadow:0 2px 6px rgba(0,0,0,.04)">
      <span style="font-size:calc(13px*var(--zf,1));font-weight:800;color:var(--navy);letter-spacing:.3px">📅 Rapports jour</span>
      <div style="display:flex;align-items:center;gap:6px;background:#fff;border:1.5px solid #c7d2fe;border-radius:10px;padding:4px 10px;box-shadow:0 1px 4px rgba(99,102,241,.12)">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#4338ca;text-transform:uppercase">📆 Du</span>
        <input type="date" id="rj-from" style="padding:3px 6px;border:none;border-radius:5px;font-size:calc(12px*var(--zf,1));color:#1e3a8a;font-weight:600;outline:none;background:transparent">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#4338ca;text-transform:uppercase">au</span>
        <input type="date" id="rj-to" style="padding:3px 6px;border:none;border-radius:5px;font-size:calc(12px*var(--zf,1));color:#1e3a8a;font-weight:600;outline:none;background:transparent">
      </div>
      <select id="rj-pilot" style="font-size:calc(12px*var(--zf,1));padding:5px 10px;border:1.5px solid #c7d2fe;border-radius:8px;background:#fff;color:#1e3a8a;font-weight:600;outline:none"><option value="">Tous les pilotes</option></select>
      <select id="rj-poste" style="font-size:calc(12px*var(--zf,1));padding:5px 10px;border:1.5px solid #c7d2fe;border-radius:8px;background:#fff;color:#1e3a8a;font-weight:600;outline:none"><option value="">Tous les postes</option></select>
      <button onclick="calcPeriodReport()" style="background:linear-gradient(180deg,#2d3480,#1a1f5e);color:#fff;border:none;border-radius:8px;padding:6px 16px;font-size:calc(12px*var(--zf,1));font-weight:700;cursor:pointer;box-shadow:0 3px 8px rgba(26,31,94,.35),inset 0 1px 0 rgba(255,255,255,.18)">🔄 Actualiser</button>
      <button onclick="resetPeriodReport()" style="background:#fff;border:1.5px solid var(--border);border-radius:8px;padding:5px 12px;font-size:calc(12px*var(--zf,1));color:var(--gray);cursor:pointer;box-shadow:0 1px 3px rgba(0,0,0,.08)">✕ Réinitialiser</button>
      <div style="margin-left:auto;display:flex;gap:6px">
        <button onclick="rjLast3()" style="background:linear-gradient(180deg,#0ea5e9,#0369a1);color:#fff;border:none;border-radius:8px;padding:5px 12px;font-size:calc(11px*var(--zf,1));font-weight:700;cursor:pointer;box-shadow:0 3px 8px rgba(3,105,161,.35),inset 0 1px 0 rgba(255,255,255,.18)">3 derniers postes</button>
        <button onclick="rjLast7Days()" style="background:linear-gradient(180deg,#34d399,#059669);color:#fff;border:none;border-radius:8px;padding:5px 12px;font-size:calc(11px*var(--zf,1));font-weight:700;cursor:pointer;box-shadow:0 3px 8px rgba(5,150,105,.35),inset 0 1px 0 rgba(255,255,255,.18)">7 derniers jours</button>
      </div>
    </div>
    <div id="rj-result" style="flex:1;overflow-y:auto;padding:14px 18px">
      <div style="padding:60px;text-align:center;color:var(--gray)">
        <div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">📅</div>
        <div style="font-size:calc(14px*var(--zf,1));font-weight:600">Sélectionnez une période puis cliquez sur Calculer</div>
      </div>
    </div>
  </div>

  <!-- ════ HISTORY ════ -->
  <div id="v-history" class="view" style="flex-direction:column;overflow:hidden">
    <div style="background:linear-gradient(180deg,#f8faff 0%,#fff 100%);border-bottom:2px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:8px;flex-shrink:0;flex-wrap:wrap;box-shadow:0 2px 6px rgba(0,0,0,.04)">
      <span style="font-size:calc(13px*var(--zf,1));font-weight:800;color:var(--navy);letter-spacing:.3px">🕘 Historique</span>
      <div style="display:flex;align-items:center;gap:6px;background:#fff;border:1.5px solid #c7d2fe;border-radius:10px;padding:4px 10px;box-shadow:0 1px 4px rgba(99,102,241,.12)">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#4338ca;text-transform:uppercase">📆 Du</span>
        <input type="date" id="hist-from" style="padding:3px 6px;border:none;border-radius:5px;font-size:calc(12px*var(--zf,1));color:#1e3a8a;font-weight:600;outline:none;background:transparent">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#4338ca;text-transform:uppercase">au</span>
        <input type="date" id="hist-to" style="padding:3px 6px;border:none;border-radius:5px;font-size:calc(12px*var(--zf,1));color:#1e3a8a;font-weight:600;outline:none;background:transparent">
      </div>
      <button class="btn btn-primary" onclick="loadHist()" style="padding:5px 14px;font-size:calc(12px*var(--zf,1));background:linear-gradient(180deg,#2d3480,#1a1f5e);border:none;border-radius:8px;color:#fff;font-weight:700;cursor:pointer;box-shadow:0 3px 8px rgba(26,31,94,.35),inset 0 1px 0 rgba(255,255,255,.18)">Charger</button>
      <div style="width:1px;height:22px;background:var(--border);flex-shrink:0"></div>
      <input id="hist-search" type="text" placeholder="🔍 Rechercher N° OF…" style="padding:4px 9px;border:1.5px solid var(--border);border-radius:7px;font-size:calc(12px*var(--zf,1));min-width:160px" oninput="filterHistBySearch()">
      <div style="width:1px;height:22px;background:var(--border);flex-shrink:0"></div>
      <div style="display:flex;gap:5px;flex-wrap:wrap;align-items:center">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:var(--gray);text-transform:uppercase">Filtrer :</span>
        <button class="hf-btn" data-hf="production" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #16a34a;color:#16a34a;background:none;cursor:pointer;font-weight:700;transition:all .15s">🏭 Production</button>
        <button class="hf-btn" data-hf="arret" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #dc2626;color:#dc2626;background:none;cursor:pointer;font-weight:700;transition:all .15s">⛔ Arrêts</button>
        <button class="hf-btn" data-hf="degrade" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #ca8a04;color:#ca8a04;background:none;cursor:pointer;font-weight:700;transition:all .15s">🟡 Mode dégradé</button>
        <button class="hf-btn" data-hf="nettoyage" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #0891b2;color:#0891b2;background:none;cursor:pointer;font-weight:700;transition:all .15s">🧹 Nettoyage</button>
        <button class="hf-btn" data-hf="pause" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #f59e0b;color:#f59e0b;background:none;cursor:pointer;font-weight:700;transition:all .15s">⏸ Pause</button>
        <button class="hf-btn" data-hf="reunion" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #8b5cf6;color:#8b5cf6;background:none;cursor:pointer;font-weight:700;transition:all .15s">👥 Réunion</button>
      </div>
    </div>
    <div style="flex:1;overflow-y:auto">
      <table class="ktbl"><thead><tr id="hist-hd"></tr></thead><tbody id="hist-bd"></tbody></table>
    </div>
  </div>

  <!-- ════ KPI VIEW ════ -->
  <div id="v-kpi" class="view" style="flex-direction:column;overflow:hidden;background:#f1f5f9">
    <!-- Header -->
    <div style="background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;flex-shrink:0;display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <span style="font-size:calc(12px*var(--zf,1));font-weight:700;color:var(--navy)">📈 Evolution perf.</span>
      <label style="font-size:calc(11px*var(--zf,1));font-weight:600;color:var(--gray)">Du <input type="date" id="kpi-from" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));margin-left:4px"></label>
      <label style="font-size:calc(11px*var(--zf,1));font-weight:600;color:var(--gray)">Au <input type="date" id="kpi-to" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));margin-left:4px"></label>
      <button onclick="loadKPI()" style="background:#1e3a8a;color:#fff;border:none;border-radius:6px;padding:6px 16px;font-size:calc(12px*var(--zf,1));font-weight:700;cursor:pointer">↺ Actualiser</button>
    </div>
    <!-- 3 courbes côte à côte (compact) -->
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;height:200px;flex-shrink:0;background:#fff;border-bottom:1px solid #e2e8f0">
      <div style="display:flex;flex-direction:column;overflow:hidden;padding:6px 10px 4px;border-right:1px solid #f1f5f9">
        <div style="font-size:calc(12px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;letter-spacing:.5px;margin-bottom:2px;flex-shrink:0">📈 TRS par poste (%)</div>
        <div id="kpi-trs-chart" style="flex:1;min-height:0;overflow:hidden"></div>
      </div>
      <div style="display:flex;flex-direction:column;overflow:hidden;padding:6px 10px 4px;border-right:1px solid #f1f5f9;background:#fafafa">
        <div style="font-size:calc(12px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#dc2626;letter-spacing:.5px;margin-bottom:2px;flex-shrink:0">🛑 Arrêts cumulés (min)</div>
        <div id="kpi-arr-chart" style="flex:1;min-height:0;overflow:hidden"></div>
      </div>
      <div style="display:flex;flex-direction:column;overflow:hidden;padding:6px 10px 4px">
        <div style="font-size:calc(12px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#6366f1;letter-spacing:.5px;margin-bottom:2px;flex-shrink:0">📦 Nombre d'OF par poste</div>
        <div id="kpi-of-chart" style="flex:1;min-height:0;overflow:hidden"></div>
      </div>
    </div>
    <!-- Corps principal -->
    <div style="flex:1;overflow:hidden;display:grid;grid-template-columns:1fr 210px;min-height:0">
      <!-- Gauche : cadence + 2 évolutions -->
      <div style="display:flex;flex-direction:column;overflow:hidden;border-right:1px solid #e2e8f0;min-height:0">
        <div style="flex:2;padding:6px 10px;overflow:hidden;display:flex;flex-direction:column;background:#fff;border-bottom:1px solid #f1f5f9;min-height:0">
          <div style="font-size:calc(12px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#f59e0b;letter-spacing:.5px;margin-bottom:3px;flex-shrink:0">⚡ Évolution cadence (éq./h)</div>
          <div id="kpi-cad-chart" style="flex:1;min-height:0;overflow:hidden"></div>
        </div>
        <div style="flex:1;padding:6px 10px;overflow:hidden;display:flex;flex-direction:column;background:#f8fafc;border-bottom:1px solid #f1f5f9;min-height:0">
          <div style="font-size:calc(12px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;letter-spacing:.5px;margin-bottom:3px;flex-shrink:0">📦 Équivalence</div>
          <div id="kpi-qte-chart" style="flex:1;min-height:0;overflow:hidden"></div>
        </div>
        <div style="flex:1;padding:6px 10px;overflow:hidden;display:flex;flex-direction:column;background:#fff;min-height:0">
          <div style="font-size:calc(12px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#8b5cf6;letter-spacing:.5px;margin-bottom:3px;flex-shrink:0">🧵 Changements de fibre par poste</div>
          <div id="kpi-fibre-chart" style="flex:1;min-height:0;overflow:hidden"></div>
        </div>
      </div>
      <!-- Droite : donut + pareto -->
      <div style="display:flex;flex-direction:column;overflow:hidden;background:#fff;min-height:0">
        <div style="padding:7px 10px;border-bottom:1px solid #f1f5f9;flex-shrink:0;display:flex;flex-direction:column;align-items:center">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;letter-spacing:.5px;margin-bottom:3px;align-self:flex-start">Prod vs Arrêts</div>
          <div style="display:flex;align-items:center;gap:8px">
            <svg id="kpi-donut" viewBox="0 0 100 100" style="width:78px;height:78px;flex-shrink:0"></svg>
            <div id="kpi-donut-legend" style="display:flex;flex-direction:column;gap:3px;font-size:calc(10px*var(--zf,1))"></div>
          </div>
        </div>
        <div style="flex:1;overflow-y:auto;padding:8px 10px;display:flex;flex-direction:column;min-height:0">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;letter-spacing:.5px;margin-bottom:6px;flex-shrink:0">Pareto arrêts</div>
          <div id="kpi-pareto-new" style="display:flex;flex-direction:column;gap:5px"></div>
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
        <input type="password" id="lock-pw" placeholder="MDP admin" style="width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:calc(14px*var(--zf,1));outline:none;margin-bottom:10px" onkeydown="if(event.key==='Enter')unlockSettings()">
        <button onclick="unlockSettings()" style="width:100%;padding:10px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:calc(14px*var(--zf,1));font-weight:700;cursor:pointer">Déverrouiller</button>
        <div id="lock-err" style="color:#dc2626;margin-top:6px;font-size:calc(12px*var(--zf,1));text-align:center;min-height:14px"></div>
      </div>
    </div>
    <div id="v-settings-content">
      <div class="ss">
        <h3>🔑 Changer le mot de passe administrateur</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Le MDP est enregistré dans la cellule G2 de l'onglet Listes du fichier Excel.</div>
        <div style="display:flex;flex-direction:column;gap:8px;max-width:360px">
          <div style="display:flex;align-items:center;gap:8px">
            <label style="width:150px;font-size:calc(12px*var(--zf,1));font-weight:600">MDP actuel</label>
            <input type="password" id="adm-old-pw" placeholder="Mot de passe actuel" style="flex:1;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1))">
          </div>
          <div style="display:flex;align-items:center;gap:8px">
            <label style="width:150px;font-size:calc(12px*var(--zf,1));font-weight:600">Nouveau MDP</label>
            <input type="password" id="adm-new-pw" placeholder="Nouveau mot de passe" style="flex:1;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1))">
          </div>
          <div style="display:flex;align-items:center;gap:8px">
            <label style="width:150px;font-size:calc(12px*var(--zf,1));font-weight:600">Confirmer MDP</label>
            <input type="password" id="adm-confirm-pw" placeholder="Confirmer le nouveau MDP" style="flex:1;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1))">
          </div>
          <div style="display:flex;align-items:center;gap:8px">
            <button class="btn btn-prim" onclick="changeAdminPw()" style="font-size:calc(12px*var(--zf,1))">🔒 Changer le MDP</button>
            <span id="adm-pw-msg" style="font-size:calc(11px*var(--zf,1));font-weight:600"></span>
          </div>
        </div>
      </div>
      <div class="ss">
        <h3>📂 Fichier Excel de données</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Chemin complet vers le fichier Excel (.xlsx) contenant les onglets Declarations et Listes.</div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:6px">
          <input id="cfg-db-path" placeholder="C:\chemin\vers\fichier.xlsx" style="flex:1;min-width:200px;padding:7px 10px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-prim" onclick="setDbPath()">💾 Enregistrer</button>
        </div>
        <div id="cfg-db-status" style="font-size:calc(11px*var(--zf,1));color:var(--gray)"></div>
      </div>
      <div class="ss">
        <h3>🔐 Mots de passe pilotes</h3>
        <div id="pwd-list"></div>
        <div class="flex mt8">
          <input id="np-name" placeholder="Nom pilote" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <input id="np-pw" type="password" placeholder="MDP" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-prim" onclick="addPilot()">+ Ajouter</button>
        </div>
        <div class="flex mt8">
          <button class="btn btn-ok" onclick="savePwds()">💾 Enregistrer MDP</button>
        </div>
      </div>
      <div class="ss">
        <h3>🧑‍🤝‍🧑 Co-pilotes</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Liste lue/écrite dans la colonne C de l'onglet Listes. Glisser pour réordonner.</div>
        <div id="copilotes-list-ui" style="margin-bottom:8px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
          <input id="cp-new-name" placeholder="Nom co-pilote" style="flex:1;min-width:140px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-green" onclick="addCopilote()">+ Ajouter</button>
          <button class="btn btn-ok" onclick="saveCopilotes()">💾 Enregistrer</button>
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
          <label style="font-weight:600;font-size:calc(12px*var(--zf,1))">Prod ref (unités/8h):</label>
          <input id="cfg-pr" type="number" style="width:90px;padding:5px;border:1px solid var(--border);border-radius:5px">
          <button class="btn btn-ok" onclick="saveProdRef()">Enregistrer</button>
        </div>
      </div>
      <div class="ss">
        <h3>⛔ Liste des arrêts configurables</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Ajouter, modifier ou supprimer les boutons d'arrêt disponibles en production. Pris en compte immédiatement.</div>
        <div id="events-list-ui" style="margin-bottom:10px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;background:#f8fafc;padding:8px;border-radius:7px;border:1px solid var(--border)">
          <input id="ev-new-label" placeholder="Nom de l'arrêt" style="flex:1;min-width:120px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <select id="ev-new-cat" style="padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
            <option value="manquants">🟣 Manquants</option>
            <option value="ratt">🟠 Rattrapage</option>
            <option value="organisation">🔵 Organisationnel</option>
            <option value="pb">🔴 Technique</option>
            <option value="autre">⚫ Autre</option>
          </select>
          <button class="btn btn-green" style="font-size:calc(11px*var(--zf,1));padding:5px 12px" onclick="addEvtItem()">+ Ajouter</button>
        </div>
        <button class="btn btn-prim" style="margin-top:8px;font-size:calc(12px*var(--zf,1))" onclick="saveEvtList()">💾 Enregistrer la liste</button>
      </div>
      <div class="ss" style="display:none"><div id="interposte-list-ui"></div></div>
      <div class="ss">
        <h3>⏱ Arrêts prévus (budget planifié)</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Les durées planifiées sont <b>déduites du temps de référence TRS</b> si le pilote les a réellement déclarées. Tout dépassement reste impactant. Mettre 0 pour désactiver.</div>
        <div style="display:flex;flex-direction:column;gap:6px">
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">🧹 Nettoyage court</span>
            <input type="number" id="ap-clean-short" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">🧹 Nettoyage long</span>
            <input type="number" id="ap-clean-long" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">🧹 Nettoyage très long</span>
            <input type="number" id="ap-clean-grand" min="0" max="240" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">📋 Réunion</span>
            <input type="number" id="ap-meeting" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">⏸ Pause</span>
            <input type="number" id="ap-pause" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
        </div>
        <div style="font-size:calc(10px*var(--zf,1));color:var(--gray);margin-top:6px">Pour l'affectation automatique : le libellé doit contenir "nettoyage court/long/très long", "réunion" ou "pause".</div>
        <button class="btn btn-prim" style="margin-top:10px;font-size:calc(12px*var(--zf,1))" onclick="saveArretsPrevus()">💾 Enregistrer arrêts prévus</button>
      </div>
      <div class="ss">
        <h3>🟡 Mode dégradé</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Configurez les motifs disponibles pour le mode dégradé. La durée est enregistrée par OF pour suivi.</div>
        <div id="degrade-list-ui" style="margin-bottom:10px"></div>
        <div style="display:flex;gap:6px;margin-bottom:10px">
          <input id="deg-new-label" placeholder="Nouveau motif dégradé" style="flex:1;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-green" onclick="addDegradeItem()">+ Ajouter</button>
        </div>
        <button class="btn btn-prim" style="font-size:calc(12px*var(--zf,1))" onclick="saveDegradeList()">💾 Enregistrer liste dégradé</button>
      </div>
      <div class="ss">
        <h3>👥 Influence nb opérateur</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Configurez le % de cadence attendu selon le nombre de personnes.<br>Données lues depuis les colonnes P/Q de l'onglet Listes du fichier Excel.</div>
        <table id="pers-pct-table" style="width:100%;border-collapse:collapse;font-size:calc(12px*var(--zf,1));margin-bottom:10px">
          <thead><tr style="background:#f1f5f9">
            <th style="padding:6px 10px;text-align:left;border:1px solid var(--border)">Nb personnes</th>
            <th style="padding:6px 10px;text-align:left;border:1px solid var(--border)">% cadence</th>
          </tr></thead>
          <tbody id="pers-pct-tbody"></tbody>
        </table>
        <button class="btn btn-prim" style="font-size:calc(12px*var(--zf,1))" onclick="savePersPct()">💾 Enregistrer</button>
      </div>
      <div class="ss">
        <h3>📐 Coefficients d'équivalence (Type produit + Coeff)</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Colonnes E (type) + F (coeff) de l'onglet Listes. Supprimer décale vers le haut. Glisser pour réordonner.</div>
        <div id="equiv-list-ui" style="margin-bottom:8px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;background:#f8fafc;padding:8px;border-radius:7px;border:1px solid var(--border)">
          <input id="eq-new-type" placeholder="Type produit" style="flex:2;min-width:120px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <input id="eq-new-coeff" placeholder="Coeff (ex: 1.66)" style="flex:1;min-width:80px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-green" onclick="addEquiv()">+ Ajouter</button>
        </div>
        <button class="btn btn-ok" style="margin-top:8px;font-size:calc(12px*var(--zf,1))" onclick="saveEquiv()">💾 Enregistrer coefficients</button>
      </div>
      <div class="ss">
        <h3>📏 Tailles de produit</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Colonne D de l'onglet Listes. Glisser pour réordonner.</div>
        <div id="tailles-list-ui" style="margin-bottom:8px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
          <input id="tl-new-val" placeholder="Nouvelle taille" style="flex:1;min-width:120px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-green" onclick="addTaille()">+ Ajouter</button>
          <button class="btn btn-ok" onclick="saveTailles()">💾 Enregistrer</button>
        </div>
      </div>
      <div class="ss">
        <h3>🧵 Fibres</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Colonne J de l'onglet Listes. Glisser pour réordonner.</div>
        <div id="fibres-list-ui" style="margin-bottom:8px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
          <input id="fb-new-val" placeholder="Nouvelle fibre" style="flex:1;min-width:120px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-green" onclick="addFibre()">+ Ajouter</button>
          <button class="btn btn-ok" onclick="saveFibres()">💾 Enregistrer</button>
        </div>
      </div>
    </div>
  </div>

</div><!-- /app -->

<!-- ════ MODAL: Choix type nettoyage ════ -->
<div class="overlay" id="m-nett-type" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:700;align-items:center;justify-content:center;backdrop-filter:blur(4px)">
  <div class="card" style="width:380px;padding:24px;background:linear-gradient(145deg,#1c1408 0%,#2d1f07 60%,#1c1408 100%);border-radius:20px;border:1px solid rgba(251,191,36,.25);box-shadow:0 0 60px rgba(245,158,11,.35),0 20px 60px rgba(0,0,0,.7),inset 0 1px 0 rgba(255,255,255,.08)">
    <div style="font-size:calc(17px*var(--zf,1));font-weight:900;color:#fcd34d;margin-bottom:18px;text-align:center;text-shadow:0 0 20px rgba(251,191,36,.8),0 2px 4px rgba(0,0,0,.6);letter-spacing:.05em">🧹 TYPE DE NETTOYAGE</div>
    <div style="display:flex;flex-direction:column;gap:10px">
      <button class="btn" id="nett-btn-court" style="text-align:left;padding:14px 18px;border-radius:14px;border:1.5px solid rgba(253,230,138,.4);background:linear-gradient(145deg,rgba(253,230,138,.18) 0%,rgba(217,119,6,.22) 100%);font-size:calc(14px*var(--zf,1));font-weight:800;color:#fde68a;transition:all .18s;box-shadow:0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.12)"
        onmouseover="this.style.boxShadow='0 0 24px rgba(253,230,138,.45),0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.18)'"
        onmouseout="this.style.boxShadow='0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.12)'"
        onclick="doStartNettoyage('court')">
        🧹 Nettoyage court
        <span id="nett-lbl-court" style="font-size:calc(11px*var(--zf,1));font-weight:500;color:#fbbf24;display:block;margin-top:3px;opacity:.85"></span>
      </button>
      <button class="btn" id="nett-btn-long" style="text-align:left;padding:14px 18px;border-radius:14px;border:1.5px solid rgba(252,211,77,.4);background:linear-gradient(145deg,rgba(252,211,77,.18) 0%,rgba(180,83,9,.22) 100%);font-size:calc(14px*var(--zf,1));font-weight:800;color:#fcd34d;transition:all .18s;box-shadow:0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.12)"
        onmouseover="this.style.boxShadow='0 0 24px rgba(252,211,77,.45),0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.18)'"
        onmouseout="this.style.boxShadow='0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.12)'"
        onclick="doStartNettoyage('long')">
        🧹 Nettoyage long
        <span id="nett-lbl-long" style="font-size:calc(11px*var(--zf,1));font-weight:500;color:#fbbf24;display:block;margin-top:3px;opacity:.85"></span>
      </button>
      <button class="btn" id="nett-btn-grand" style="text-align:left;padding:14px 18px;border-radius:14px;border:1.5px solid rgba(245,158,11,.5);background:linear-gradient(145deg,rgba(245,158,11,.22) 0%,rgba(124,45,18,.28) 100%);font-size:calc(14px*var(--zf,1));font-weight:800;color:#f59e0b;transition:all .18s;box-shadow:0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.12)"
        onmouseover="this.style.boxShadow='0 0 24px rgba(245,158,11,.5),0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.18)'"
        onmouseout="this.style.boxShadow='0 4px 16px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.12)'"
        onclick="doStartNettoyage('grand')">
        🧹 Nettoyage très long
        <span id="nett-lbl-grand" style="font-size:calc(11px*var(--zf,1));font-weight:500;color:#fbbf24;display:block;margin-top:3px;opacity:.85"></span>
      </button>
    </div>
    <button class="btn btn-sec" style="margin-top:16px;width:100%;font-size:calc(13px*var(--zf,1));background:rgba(255,255,255,.07);border-color:rgba(255,255,255,.15);color:#cbd5e1;border-radius:10px;padding:10px" onclick="closeM('m-nett-type')">✕ Annuler</button>
  </div>
</div>

<!-- ════ MODAL: Déclarer un arrêt ════ -->
<div class="overlay" id="m-stop">
  <div class="mbox" style="width:70%;max-width:70vw">
    <div class="mhdr red">
      <h2>⛔ Déclarer un arrêt</h2>
      <button style="background:none;border:none;cursor:pointer;color:#fff;font-size:calc(16px*var(--zf,1))" onclick="closeM('m-stop')">✕</button>
    </div>
    <div class="mbody">
      <div id="sgrid-all" style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:6px"></div>
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
        <div style="flex:1;min-width:160px"><div class="ep-grid" id="ep-stats"></div></div>
        <div style="display:flex;gap:8px;flex-shrink:0">
          <div style="text-align:center">
            <div style="font-size:calc(9px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:2px">Répartition</div>
            <svg id="ep-pie" viewBox="0 0 130 115" style="width:160px;height:142px;display:block"></svg>
          </div>
          <div style="text-align:center">
            <div style="font-size:calc(11px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:2px">TRS</div>
            <svg id="ep-gauge" viewBox="0 0 100 58" style="width:130px;display:block;margin:0 auto;margin-top:8px">
              <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
              <path id="ep-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
              <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e" id="ep-gauge-pct">--%</text>
            </svg>
          </div>
        </div>
      </div>
      <div style="margin-bottom:10px">
        <div style="font-size:calc(10px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:5px">Arrêts &amp; pauses</div>
        <table class="ep-tbl"><thead><tr><th>Type</th><th>Durée</th><th>%</th></tr></thead><tbody id="ep-stops"></tbody></table>
      </div>
      <div class="card" style="padding:8px;margin-bottom:0">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">Timeline</div>
        <svg id="ep-tl" viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">
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
      <div class="fr" style="margin-bottom:10px"><label style="font-weight:800;color:#dc2626">🔑 Mot de passe admin</label><input type="password" id="es-pw" placeholder="Mot de passe requis"></div>
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
      <div id="cmt-stop-lbl" style="font-size:calc(13px*var(--zf,1));font-weight:700;color:var(--navy);margin-bottom:10px"></div>
      <div class="fr"><label>Commentaire (optionnel)</label><textarea id="cmt-stop-text" style="height:70px;resize:none;width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))" placeholder="Description de l'arrêt…"></textarea></div>
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
      <button style="background:none;border:none;cursor:pointer;font-size:calc(16px*var(--zf,1))" onclick="closeM('m-editrow')">✕</button>
    </div>
    <div class="mbody">
      <input type="hidden" id="er-rownum"><input type="hidden" id="er-rowtype">
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
          <div class="fr"><label>Traca Fibre</label><input type="text" id="er-traca" placeholder="n° de lot(s), séparés par ;"></div>
          <div class="fr"><label>Réf Taie</label><input id="er-reftaie"></div>
          <div class="fr"><label>Lots de 2</label><select id="er-kit"><option value="">Non</option><option value="oui">Oui</option></select></div>
          <div class="fr"><label>Qté Init Taie</label><input type="number" id="er-qteinit"></div>
          <div class="fr"><label>Nb Taie 2nd Choix</label><input type="number" id="er-nbtaie2"></div>
          <div class="fr"><label>Nb Défaut Couture</label><input type="number" id="er-nbdef"></div>
          <div class="fr"><label>Mq Taie</label><input type="number" id="er-mqtaie"></div>
          <div class="fr"><label>Mq Housse/Encart</label><input type="number" id="er-mqhousse"></div>
          <div class="fr"><label>PP cousu et emballé</label><input type="number" id="er-nbpp"></div>
          <div class="fr"><label>Duree MQ MP (min)</label><input type="number" id="er-dureemq"></div>
          <div class="fr" style="display:none"><input type="number" id="er-manqpers"></div>
        </div>
        <div class="fr comment-big"><label>💬 Commentaire</label><textarea id="er-comment-prod" style="height:60px;resize:none;width:100%;padding:4px 6px;border:2px solid #f59e0b;border-radius:4px;font-size:calc(12px*var(--zf,1));background:#fffbeb"></textarea></div>
      </div>
      <div id="er-evt-fields" style="display:none">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:6px">
          <div class="fr"><label>Type d'arrêt</label><select id="er-evttype"><option value="">--</option></select></div>
          <div class="fr"><label>N° OF</label><input id="er-evtof"></div>
          <div class="fr"><label>Heure Début</label><input type="time" id="er-evtdeb" step="60"></div>
          <div class="fr"><label>Heure Fin</label><input type="time" id="er-evtfin" step="60"></div>
        </div>
        <div class="fr"><label>Commentaire</label><textarea id="er-evtcomment" style="height:44px;resize:none;width:100%;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1))"></textarea></div>
        <div class="fr" style="margin-top:6px"><label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:600;font-size:calc(12px*var(--zf,1))"><input type="checkbox" id="er-horstrs"> Hors TRS</label></div>
      </div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-editrow')">Annuler</button>
      <button class="btn btn-danger" onclick="deleteRow(null,'er-rownum')">🗑 Supprimer</button>
      <button class="btn btn-ok" onclick="saveEditRow()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<div id="m-degrade" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:5100;align-items:center;justify-content:center;backdrop-filter:blur(4px)" onclick="if(event.target===this)closeM('m-degrade')">
  <div style="max-width:400px;width:90%;padding:24px;background:linear-gradient(145deg,#1a1200 0%,#271900 60%,#1a1200 100%);border-radius:20px;border:1px solid rgba(202,138,4,.3);box-shadow:0 0 60px rgba(202,138,4,.4),0 20px 60px rgba(0,0,0,.8),inset 0 1px 0 rgba(255,255,255,.07)" onclick="event.stopPropagation()">
    <div id="m-degrade-title" style="font-size:calc(17px*var(--zf,1));font-weight:900;color:#fbbf24;margin-bottom:18px;text-align:center;text-shadow:0 0 24px rgba(202,138,4,.9),0 2px 4px rgba(0,0,0,.6);letter-spacing:.05em">🟡 MODE DÉGRADÉ</div>
    <div id="m-degrade-body" style="margin-bottom:16px"></div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-prim" id="m-degrade-confirm" onclick="_confirmDegrade()" style="flex:1;background:linear-gradient(145deg,#ca8a04,#92400e);border:none;color:#fff;font-weight:800;font-size:calc(14px*var(--zf,1));border-radius:12px;padding:11px;box-shadow:0 4px 18px rgba(202,138,4,.5),inset 0 1px 0 rgba(255,255,255,.2)">✓ Confirmer</button>
      <button class="btn btn-sec" onclick="closeM('m-degrade')" style="flex:1;background:rgba(255,255,255,.07);border:1.5px solid rgba(255,255,255,.15);color:#cbd5e1;border-radius:12px;padding:11px;font-size:calc(14px*var(--zf,1))">✕ Annuler</button>
    </div>
  </div>
</div>

<div id="m-budget-override" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:5100;align-items:center;justify-content:center" onclick="if(event.target===this)closeM('m-budget-override')">
  <div class="mbox" style="max-width:360px;padding:20px" onclick="event.stopPropagation()">
    <div class="mhdr" style="margin:-20px -20px 14px;padding:14px 16px;border-radius:12px 12px 0 0"><h2>✏️ Budget arrêts — surcharge ponctuelle</h2></div>
    <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Modification temporaire pour ce poste uniquement. Revient aux valeurs globales en fin de poste.</div>
    <div class="fr" style="margin-bottom:12px"><label>Code admin</label><input type="password" id="bov-pw" placeholder="••••" style="padding:6px 10px;border:1.5px solid #cbd5e1;border-radius:6px;font-size:calc(14px*var(--zf,1));width:100%"></div>
    <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px" id="bov-fields"></div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-prim" onclick="saveBudgetOverride()">✓ Appliquer</button>
      <button class="btn btn-sec" onclick="closeM('m-budget-override')">Annuler</button>
    </div>
  </div>
</div>

<div id="m-pw-action" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:5000;align-items:center;justify-content:center">
  <div class="mbox" style="max-width:320px;padding:20px" onclick="event.stopPropagation()">
    <div class="mhdr" style="margin:-20px -20px 14px;padding:14px 16px;border-radius:12px 12px 0 0"><h2 id="pwa-title">🔒 Mot de passe Admin</h2></div>
    <div class="fr" style="margin-bottom:14px"><label>Mot de passe</label><input type="password" id="pwa-pw" placeholder="••••" style="padding:6px 10px;border:1.5px solid #cbd5e1;border-radius:6px;font-size:calc(14px*var(--zf,1));width:100%" onkeydown="if(event.key==='Enter')_pwaConfirm()"></div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-prim" onclick="_pwaConfirm()">✓ Confirmer</button>
      <button class="btn btn-sec" onclick="closeM('m-pw-action')">Annuler</button>
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
  nettoyage:"#f97316",
  organisation:"#3b82f6",
  manquants:"#9333ea",
  reunion:"#8b5cf6",
  "_pause":"#94a3b8","Pause":"#94a3b8"
};
const _STOP_GRAD = {
  ratt:        ['#f87171','#dc2626','#991b1b'],
  pb:          ['#f87171','#dc2626','#991b1b'],
  manquants:   ['#d8b4fe','#9333ea','#6b21a8'],
  organisation:['#93c5fd','#3b82f6','#1d4ed8'],
  nettoyage:   ['#fdba74','#f97316','#c2410c'],
  autre:       ['#94a3b8','#64748b','#334155'],
  _pause:      ['#cbd5e1','#64748b','#334155'],
  Pause:       ['#cbd5e1','#64748b','#334155'],
  reunion:     ['#c4b5fd','#8b5cf6','#5b21b6'],
};

function getStopColor(key, cat) {
  if (cat) return STOP_COL[cat]||'#64748b';
  if (key==='nettoyage') return STOP_COL.nettoyage;
  if (key==='_pause'||key==='Pause') return STOP_COL['Pause'];
  return '#94a3b8';
}

const FORM_FIELDS = ["of_num","copilote","nb_pers","taille","code_prod","type_prod","poids","fibre","of_taie","traca","ref_taie","kit","qte_fab","qte_emb","qte_init_taie","nb_taie2_choix","nb_def_cout","mq_taie","mq_housse_encart","nb_pp_cousue","duree_mq_mp","comment"];

// ── State ──
let ST = {};
let gEvts = [];
let _curStopKey = null;
let _curStopElap = 0;
let _ofElapAtPoll = 0;
let _stopWallAtPoll = 0;
let _pauseTotalAtPoll = 0;
let _pauseElapAtPoll = 0;
let _pauseBaseS = 0;
let _pauseStartMs = 0;
let _lastPoll = Date.now();
let _ticker = null;
let _autoSaveTimer = null;
let _curTab = 'main';
let _cfgPwds = {};
let _cfgModels = [];
let _cfgModelsBase = [];
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
  // Populate Paramètres list editors
  _copilotesList=(d.copilotes||[]).slice();renderCopilotesList();
  _taillesList=(d.tailles||[]).slice();renderTaillesList();
  _fibresList=(d.fibres||[]).slice();renderFibresList();
  const _tps=d.types_prod||[],_eqs=d.equivalences||[];
  _equivList=_tps.map((t,i)=>({type:t,coeff:String(_eqs[i]||'')}));renderEquivList();
  // er-traca est maintenant un input text (multi-lots), pas de popSel
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
  _cfgModelsBase = d.modeles_horaires_base||JSON.parse(JSON.stringify(_cfgModels));
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
  const model=_cfgModelsBase.find(m=>m.nom===nom);
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
        // Mettre à jour shift_debut_dt / shift_fin_dt et POSTES P/Q avec les horaires modifiés
        const today2=new Date();
        const [dH,dM]=newDebut.split(':').map(Number);
        const debDt=new Date(today2.getFullYear(),today2.getMonth(),today2.getDate(),dH,dM,0);
        const [fH,fM]=newFin.split(':').map(Number);
        const finDt2=new Date(today2.getFullYear(),today2.getMonth(),today2.getDate(),fH,fM,0);
        if(finDt2<=debDt) finDt2.setDate(finDt2.getDate()+1);
        await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:debDt.toISOString(),fin_iso:finDt2.toISOString()})});
      }
    }
    const s = await apiFetch('/api/state');
    document.getElementById('v-login').classList.remove('on');
    showApp(s||{pilot,poste});
  } else {
    document.getElementById('ln-err').textContent = d.error||'Erreur connexion';
  }
}

function toggleExcelPanel(){
  var p=document.getElementById('excel-panel');
  if(!p) return;
  if(p.style.display==='none'){
    p.style.display='block';
    fetch('/api/config').then(function(r){return r.json();}).then(function(d){
      var inp=document.getElementById('excel-path-inp');
      if(inp) inp.value=d.db_path||'';
    }).catch(function(){});
  } else {
    p.style.display='none';
  }
}
function saveExcelPath(){
  var path=document.getElementById('excel-path-inp').value.trim();
  var pw=document.getElementById('excel-pw-inp').value;
  var errEl=document.getElementById('excel-panel-err');
  if(errEl) errEl.textContent='';
  if(!path){if(errEl) errEl.textContent='Chemin requis';return;}
  fetch('/api/set_db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path:path,pw:pw})})
    .then(function(r){return r.json();})
    .then(function(d){
      if(d.ok){
        var p=document.getElementById('excel-panel');
        if(p) p.style.display='none';
        if(typeof loadLists==='function') loadLists();
        toast('Excel OK','ok');
      } else {
        if(errEl) errEl.textContent=d.error||'Erreur';
      }
    })
    .catch(function(){if(errEl) errEl.textContent='Erreur';});
}
function doGuestLogin(){
  window._guestMode=true;
  document.getElementById('v-login').classList.remove('on');
  const app=document.getElementById('app');
  app.style.display='flex';app.classList.remove('hidden');
  setToday();
  loadCfg();
  _applyGuestMode();
  goTab('main');
  loadMainDecl();
}

function _applyGuestMode(){
  const actionBtns=document.getElementById('main-action-btns');
  if(actionBtns) actionBtns.style.display='none';
  const prodTab=document.getElementById('ht-prod');
  if(prodTab) prodTab.style.display='none';
  const guestBadge=document.getElementById('ht-guest-badge');
  if(guestBadge) guestBadge.style.display='';
  const pilotLbl=document.getElementById('hdr-pilot-lbl');
  if(pilotLbl) pilotLbl.textContent='👁 Invité';
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
  const _30ago=new Date(Date.now()-30*86400000).toISOString().slice(0,10);
  const _hf=document.getElementById('hist-from'); if(_hf&&!_hf.value)_hf.value=_30ago;
  const _ht=document.getElementById('hist-to'); if(_ht&&!_ht.value)_ht.value=_today;
}

async function doLogout() {
  if (ST.prod_active) { toast('Terminer la production avant de déconnecter','err'); return; }
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  resetToLogin();
}

// ── Zoom texte ──
var _zoomOpen=false;
function applyZoom(pct){
  pct=Math.min(150,Math.max(70,pct));
  document.documentElement.style.removeProperty('zoom');
  document.documentElement.style.setProperty('--zf', pct/100);
  var sl=document.getElementById('zoom-slider');
  var vl=document.getElementById('zoom-val');
  if(sl) sl.value=pct;
  if(vl) vl.textContent=pct+'%';
  try{localStorage.setItem('kpi_zoom_pct',pct);}catch(e){}
}
function stepZoom(delta){
  var cur=+((document.getElementById('zoom-slider')||{value:100}).value);
  applyZoom(cur+delta);
}
function resetZoom(){applyZoom(135);}
function toggleZoomPop(){
  _zoomOpen=!_zoomOpen;
  var pop=document.getElementById('zoom-pop');
  if(pop) pop.style.display=_zoomOpen?'block':'none';
}
document.addEventListener('click',function(e){
  if(!_zoomOpen) return;
  var pop=document.getElementById('zoom-pop');
  var btn=document.getElementById('zoom-btn');
  if(pop&&btn&&!pop.contains(e.target)&&!btn.contains(e.target)){
    _zoomOpen=false;pop.style.display='none';
  }
});
(function(){
  try{var z=+localStorage.getItem('kpi_zoom_pct');applyZoom((z>=70&&z<=150)?z:135);}catch(e){applyZoom(135);}
})();

function resetToLogin() {
  ST={}; _curStopKey=null;
  window._guestMode=false;
  _missingDeclChecked=false;
  _ecartChecked=false;
  window._finPosteMode=false;
  document.getElementById('app').style.display='none';
  document.getElementById('ln-pw').value='';
  document.getElementById('ln-err').textContent='';
  document.getElementById('v-login').classList.add('on');
  _settingsUnlocked = false;
  // Restaurer les éléments cachés en mode invité
  const actionBtns=document.getElementById('main-action-btns');if(actionBtns) actionBtns.style.display='';
  const _btnStart=document.getElementById('btn-start');
  if(_btnStart){if(ST.prod_active){_btnStart.innerHTML='<span class="act-icon">▶</span><span>Production en cours</span>';_btnStart.onclick=()=>goTab('prod');}else{_btnStart.innerHTML='<span class="act-icon">▶</span><span>Démarrer production</span>';_btnStart.onclick=doStartProd;}}
  const prodTab=document.getElementById('ht-prod');if(prodTab) prodTab.style.display='';
  const guestBadge=document.getElementById('ht-guest-badge');if(guestBadge) guestBadge.style.display='none';
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
  if(window._guestMode && (tab==='prod'||tab==='finposte')){toast('Mode consultation — accès restreint','warn');return;}
  const _prevTab=_curTab;
  _curTab = tab;
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('on'));
  document.querySelectorAll('.htab').forEach(t=>t.classList.remove('on'));
  const vm={main:'v-main',prod:'v-prod',history:'v-history',settings:'v-settings',finposte:'v-finposte',kpi:'v-kpi',rapports:'v-rapports','rpt-jour':'v-rpt-jour'};
  const el=document.getElementById(vm[tab]);
  if(el) el.classList.add('on');
  const nt={main:'ht-main',prod:'ht-prod',history:'ht-hist',settings:'ht-cfg',kpi:'ht-kpi',rapports:'ht-rapports','rpt-jour':'ht-rpt-jour'};
  const ntEl=document.getElementById(nt[tab]);
  if(ntEl) ntEl.classList.add('on');
  if(tab==='history') loadHist();
  else if(_prevTab==='history') _resetHistFilters();
  if(tab==='finposte') loadFPData();
  if(tab==='rapports'){_rptSetLast7();loadRapports();rptBackToList();}
  if(tab==='rpt-jour') loadRptJour();
  if(tab==='main') { loadMainDecl(); }
  if(tab!=='prod') _clearFieldHighlights();
  if(tab==='kpi') loadKPI();
  if(tab==='settings') {
    _settingsUnlocked = false;
    document.getElementById('settings-lock').style.display = 'flex';
    document.getElementById('v-settings-content').style.display = 'none';
    document.getElementById('lock-pw').value='';
    document.getElementById('lock-err').textContent='';
    setTimeout(()=>document.getElementById('lock-pw')?.focus(), 80);
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
    renderDegradeList(_degradeListLocal);
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
  _pauseBaseS = _pauseTotalAtPoll;
  _pauseStartMs = (s.is_paused && s.pause_start_iso) ? new Date(s.pause_start_iso).getTime() : 0;
  _pauseElapAtPoll = _pauseBaseS + (_pauseStartMs > 0 ? (Date.now()-_pauseStartMs)/1000 : 0);
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
  _checkEndProdBtn();
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
  // Sync pers_pct_map depuis state (toujours à jour)
  if(s.pers_pct_map) _persPctMapLocal=s.pers_pct_map;
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

  // Main btn-start: texte + action selon état prod
  const bs=document.getElementById('btn-start');
  if(bs){
    if(s.prod_active){bs.innerHTML='<span class="act-icon">▶</span><span>Production en cours</span>';bs.onclick=()=>goTab('prod');}
    else{bs.innerHTML='<span class="act-icon">▶</span><span>Démarrer production</span>';bs.onclick=doStartProd;}
  }

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

  // OF timing in pob banner
  const rcFin=document.getElementById('rc-fin');
  if(rcFin) rcFin.textContent=s.prod_active?'en cours':'—';

  // Render active stop chips (prod view + main view)
  renderStopChips(s);
  renderStopChipsMain(s);

  // Render live events for current prod (recap + timeline) from tl_events in state
  if(s.prod_active&&_curTab==='prod'){
    const le=tlEventsToDisplayFmt(s.tl_events||[]);
    // Add completed pauses from pause_periods
    (s.pause_periods||[]).forEach(([pStart,pEnd])=>{
      const s0=pStart?new Date(pStart):null;
      const e0=pEnd?new Date(pEnd):null;
      if(!s0) return;
      const pad=n=>String(n).padStart(2,'0');
      const toHMS=d=>pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
      const dur=e0?(e0.getTime()-s0.getTime())/1000:0;
      le.push({type:'Pause',cat:'_pause',debut:toHMS(s0),fin:e0?toHMS(e0):'',duree:dur>0?fmtDur(dur):'',comment:'',hors_trs:false,_live:!pEnd});
    });
    le.sort((a,b)=>a.debut.localeCompare(b.debut));
    renderTL('tl-svg',le);
    renderRecap(le);
  }

  // Mode dégradé strip
  const ds=document.getElementById('degrade-strip');
  const bdAcc=document.getElementById('btn-degrade-acc');
  const bdProd=document.getElementById('btn-degrade-prod');
  window._degradeActive=s.degrade_active||false;
  window._degradeMotifs=s.degrade_motifs||[];
  window._degradePeriodsIso=s.degrade_periods_iso||[];
  const degWrap=document.getElementById('sc-deg-wrap');
  if(s.degrade_active){
    if(ds){ds.textContent='🟡 MODE DÉGRADÉ EN COURS : '+esc(s.degrade_type||'');ds.style.display='block';ds.onclick=stopDegrade;}
    if(bdAcc){bdAcc.style.background='#ca8a04';bdAcc.style.color='#fff';bdAcc.innerHTML='<span class="act-icon">🐌</span><span>Désactiver dégradé</span>';bdAcc.style.animation='blink .85s step-start infinite';}
    if(bdProd){bdProd.style.background='#ca8a04';bdProd.style.color='#fff';bdProd.innerHTML='<span class="act-icon">🐌</span><span>Désactiver dégradé</span>';bdProd.style.animation='blink .85s step-start infinite';}
    if(degWrap) degWrap.style.display='';
  } else {
    if(ds){ds.style.display='none';ds.onclick=null;}
    if(bdAcc){bdAcc.style.background='';bdAcc.style.color='';bdAcc.innerHTML='<span class="act-icon">🐌</span><span>Mode dégradé</span>';bdAcc.style.animation='none';}
    if(bdProd){bdProd.style.background='radial-gradient(ellipse at 50% 25%,#fde68a 0%,#f59e0b 55%,#92400e 100%)';bdProd.style.color='#fff';bdProd.innerHTML='<span class="act-icon">🐌</span><span>Mode dégradé</span>';bdProd.style.animation='none';}
    if(degWrap) degWrap.style.display='none';
  }
  // Afficher bouton ✏️ budget si pilote connecté
  const _showBov=!!(s.pilot);
  const bovA=document.getElementById('btn-bov-acc');if(bovA)bovA.style.display=_showBov?'':'none';
  const bovP=document.getElementById('btn-bov-prod');if(bovP)bovP.style.display=_showBov?'':'none';
  window._budgetOverrides=s.budget_overrides||{};
  window._lastBudgetState=s.budget_state||null;
  // Pause button text
  const pbtn=document.getElementById('btn-pause');
  if(pbtn){pbtn.innerHTML=s.is_paused?'<span class="act-icon">▶</span>Reprendre':'<span class="act-icon">☕</span>Pause';}
  const rbtn=document.getElementById('btn-reunion');
  if(rbtn){rbtn.innerHTML=s.reunion_active?'<span class="act-icon">✓</span>Fin réunion':'<span class="act-icon">🗣️</span>Réunion';}

  // TRS gauge
  updateGauge(s);
  updateAccModelInfo();
}

function getEvtLabel(key) {
  // Priorité : liste dynamique, puis EVENTS statique
  const dynEv=_evtsList.find(e=>e.key===key);
  if(dynEv) return (dynEv.cat==='ratt'?'Rattrapage: ':dynEv.cat==='pb'?'PB: ':'')+dynEv.label;
  const ev=EVENTS.find(e=>e[1]===key);
  if(ev) return (ev[2]==='ratt'?'Rattrapage: ':ev[2]==='pb'?'PB: ':'')+ev[0];
  if(key==='nettoyage') return 'Nettoyage';
  // Fallback par mot-clé (clés non normalisées ou issues d'une ancienne version)
  const kl=(key||'').toLowerCase();
  if(kl.includes('reunion')||kl.includes('union')||kl.includes('meeting')) return 'Réunion';
  if(kl.includes('pause')) return 'Pause';
  if(kl.includes('nettoyage')||kl.includes('nett')) return 'Nettoyage';
  // Recherche fuzzy dans _evtsList
  const fuzzy=_evtsList.find(e=>(e.label||'').toLowerCase().split(' ').some(w=>w.length>3&&kl.includes(w)));
  if(fuzzy) return fuzzy.label;
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

let _lastMainChipKeys='';
function renderStopChipsMain(s) {
  const cont=document.getElementById('stop-chips-main');
  const sb=document.getElementById('stop-bottom-main');
  if(!cont||!sb) return;
  const stops=s.active_stops||[];
  const hasAny=stops.length>0||s.is_paused;
  if(!hasAny){sb.classList.remove('on');cont.innerHTML='';_lastMainChipKeys='';return;}
  sb.classList.add('on');
  const newKeys=stops.join(',')+(s.is_paused?'|pause':'');
  if(newKeys===_lastMainChipKeys){
    // Même set de stops : mettre à jour seulement les timers
    stops.forEach(k=>{
      const elap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
      const el=document.getElementById('main-chip-t-'+k);
      if(el) el.textContent=fmtDur2(elap);
    });
    if(s.is_paused){
      const pe=(s.pause_total_s||0)+(s.pause_start_iso?(Date.now()-new Date(s.pause_start_iso).getTime())/1000:0);
      const el=document.getElementById('main-chip-t-_pause');
      if(el) el.textContent=fmtDur2(pe);
    }
    return;
  }
  _lastMainChipKeys=newKeys;
  let html='';
  stops.forEach(k=>{
    const lbl=getEvtLabel(k);
    const elap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
    html+=`<div class="stop-chip"><span class="chip-lbl">⛔ ${esc(lbl)}</span><span class="chip-tim" id="main-chip-t-${esc(k)}">${fmtDur2(elap)}</span><button class="btn-endstop" onclick="doEndStop('${esc(k)}')">✓ Terminer</button></div>`;
  });
  if(s.is_paused){
    const pe=(s.pause_total_s||0)+(s.pause_start_iso?(Date.now()-new Date(s.pause_start_iso).getTime())/1000:0);
    html+=`<div class="stop-chip"><span class="chip-lbl">⏸ Pause</span><span class="chip-tim" id="main-chip-t-_pause">${fmtDur2(pe)}</span><button class="btn-endstop" onclick="doPause()">▶ Reprendre</button></div>`;
  }
  cont.innerHTML=html;
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
    // Tick main-chip timers even without active prod
    {const dt0=(Date.now()-_lastPoll)/1000;
    (ST.active_stops||[]).forEach(k=>{
      const cel=document.getElementById('main-chip-t-'+k);
      if(cel&&ST.timers&&ST.timers[k]) cel.textContent=fmtDur2(ST.timers[k].elapsed+dt0);
    });
    if(ST.is_paused){const cel=document.getElementById('main-chip-t-_pause');
      if(cel) cel.textContent=fmtDur2(_pauseStartMs>0?_pauseBaseS+(Date.now()-_pauseStartMs)/1000:_pauseBaseS);}
    }
    if(!ST.prod_active) return;
    const dt=(Date.now()-_lastPoll)/1000;
    // OF timer
    const ofEl=_ofElapAtPoll+(!ST.is_paused?dt:0);
    const t=document.getElementById('sc-of');
    if(t) t.textContent=fmtDur(ofEl);
    const rd=document.getElementById('rc-duree');
    if(rd) rd.textContent=fmtDur(ofEl);
    // Stops total
    const sw=_stopWallAtPoll+((_curStopKey&&_curStopKey!=='_pause')?dt:0);
    const ts=document.getElementById('sc-stops');
    if(ts) ts.textContent=fmtDur(sw);
    // Pause total (kept for arrêt totaux computation)
    const pauseNow=_pauseStartMs>0?_pauseBaseS+(Date.now()-_pauseStartMs)/1000:_pauseBaseS;
    // Dégradé timer
    if(ST&&ST.degrade_active&&ST.degrade_start_iso){
      const _degElapsed=(Date.now()-new Date(ST.degrade_start_iso).getTime())/1000;
      const tdeg=document.getElementById('sc-deg');
      if(tdeg) tdeg.textContent=fmtDur(Math.max(0,_degElapsed));
    }
    // Arrêt totaux du poste (stops wall + pauses, merged — no double-counting)
    const tall=document.getElementById('sc-all-stops');
    if(tall) tall.textContent=fmtDur(sw+pauseNow);
    // Pièces théoriques : avec déduction budget + ajustement mode dégradé
    const thEl=document.getElementById('sc-theo');
    if(thEl&&ST.prod_ref){
      const typeProd=document.getElementById('f-type_prod')?.value||ST.form?.type_prod||'';
      const coef=(window._equivCoefs&&window._equivCoefs[typeProd])||1;
      const _ofDedT=(ST.budget_state&&ST.budget_state.total_of_deductible_s)||0;
      const effOfElT=Math.max(1,_ofElapAtPoll+dt-_ofDedT);
      // Calculer les secondes en mode dégradé (cadence ÷ 2)
      let _degST=0;
      const _ofStartMsT=ST.of_start_iso?new Date(ST.of_start_iso).getTime():0;
      const _nowMsT=Date.now();
      if(ST.degrade_active&&ST.degrade_start_iso&&_ofStartMsT>0){
        const _dsTmp=Math.max(new Date(ST.degrade_start_iso).getTime(),_ofStartMsT);
        _degST=Math.max(0,(_nowMsT-_dsTmp)/1000);
      }
      (ST.degrade_periods_iso||[]).forEach(function(p){
        if(!p.start||!p.end||!_ofStartMsT) return;
        const _d0=Math.max(new Date(p.start).getTime(),_ofStartMsT);
        const _d1=Math.min(new Date(p.end).getTime(),_nowMsT);
        if(_d1>_d0) _degST+=(_d1-_d0)/1000;
      });
      const _adjEffOfElT=Math.max(1,effOfElT);
      const _nbPersTheo=parseInt(document.getElementById('f-nb_pers')?.value||'1')||1;
      const _pctTheo=(_persPctMapLocal&&_persPctMapLocal[String(_nbPersTheo)]!=null)?(_persPctMapLocal[String(_nbPersTheo)]/100):1.0;
      const theo=Math.round(ST.prod_ref*_pctTheo*_adjEffOfElT/28800/coef);
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
      if(cel) cel.textContent=fmtDur2(_pauseStartMs>0?_pauseBaseS+(Date.now()-_pauseStartMs)/1000:_pauseBaseS);
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
  window._mainEvtsAll=evts;
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
  window._lastMainRows=inShiftEvts; // pour checkMissingDecls()
  _todayStopAccum=inShiftEvts.reduce((a,r)=>a+_hms2s(r.duree||''),0);
  // Heure de la dernière déclaration prod enregistrée (dans la fenêtre du poste)
  if(inShiftDecls.length){
    const lastFin=inShiftDecls.map(r=>r.fin||'').filter(Boolean).sort().pop();
    if(lastFin){const[h,m,s]=(lastFin+'::').split(':').map(Number);const d=new Date();d.setHours(h,m,s||0,0);_lastProdDeclTime=d;}
  }
  if(!allRows.length){bd.innerHTML='<tr><td colspan="11" style="text-align:center;color:var(--gray);padding:16px">Aucune déclaration aujourd\'hui</td></tr>';loadMainKPI();updateGauge(ST);return;}
  window._rowMap={};
  bd.innerHTML=allRows.map(r=>{
    const key=r.row_num||r.debut;
    window._rowMap[String(key)]=r;
    const isProd=r._rowType==='prod';
    const t=parseFloat(r.trs||0);
    const tag=isProd?'<span class="row-tag tag-p">🏭 Prod</span>':(r.is_degrade?'<span class="row-tag tag-e" style="border-color:#ca8a04;color:#ca8a04">🟡 Dégradé</span>':(r.type&&r.type.toLowerCase().includes('nett')?'<span class="row-tag tag-n">🧹 Nett.</span>':'<span class="row-tag tag-e">⛔ Arrêt</span>'));
    const details=isProd?esc(r.taille||''):esc(r.type||'');
    const qty=isProd?esc(String(r.qte_fab||'')):esc(r.duree||'');
    const info=isProd&&t>0?`<span class="${t>=90?'tg':t>=75?'tm':'tb'}">${fmtTRS(t)}</span>`:'—';
    const cmt=esc(r.comment||'');
    const fibre=r.fibre||'';const fibreShort=esc(fibre.slice(0,9));
    return `<tr class="${isProd?'row-prod':'row-evt'}">
      <td>${tag}</td><td style="font-weight:700;color:${isProd?'#1e3a8a':'#dc2626'};text-decoration:underline;cursor:pointer" onclick="showMainRowDetail('${esc(String(key))}')" title="Voir détail">${esc(r.of||r.type||'—')}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:#6366f1;font-weight:600;cursor:${fibre?'pointer':''}" title="${esc(fibre)}" onclick="${fibre?'showFibre(\''+esc(fibre)+'\')':''}">${fibreShort}${fibre.length>9?'…':''}</td>
      <td style="font-size:calc(10px*var(--zf,1))">${esc(r.date||'')}</td><td style="font-size:calc(10px*var(--zf,1))">${esc(r.poste||'')}</td>
      <td>${esc(r.pilote||'')}</td><td>${esc(r.debut||'')}</td><td>${esc(r.fin||'')}</td>
      <td style="font-size:calc(11px*var(--zf,1))">${details}</td><td style="font-size:calc(11px*var(--zf,1))">${qty}</td><td>${info}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:var(--gray);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${cmt}">${cmt}</td>
      <td><button onclick="openEditRow('${esc(String(key))}')" style="background:#6366f1;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:calc(15px*var(--zf,1));cursor:pointer;font-weight:700" title="Modifier">✏</button></td>
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
    // TRS Actuel : début de plage → fin de la dernière déclaration de prod, avec déduction budget
    let trs=-1;
    if(_todayEquivAccum>0&&_shiftRefDt&&ST.prod_ref>0){
      const refTime=_lastProdDeclTime||new Date();
      const shiftElap=(refTime.getTime()-_shiftRefDt.getTime())/1000;
      const _shDed=(ST.budget_state&&ST.budget_state.total_shift_deductible_s)||0;
      const effElap=Math.max(1,shiftElap-_shDed);
      if(effElap>0) trs=Math.round(_todayEquivAccum/(ST.prod_ref*effElap/28800)*100*10)/10;
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
  // Nb OF + total pcs/equiv in new accueil block
  const nbOfEl=document.getElementById('acc-nb-of');
  if(nbOfEl) nbOfEl.textContent=d&&d.rows?d.rows.length:0;
  const prodTotEl=document.getElementById('acc-prod-total');
  if(prodTotEl&&d&&d.rows){
    let tPcs=0,tEq=0;
    d.rows.forEach(r=>{
      tPcs+=parseFloat((r.qte_fab||'0').toString().replace(',','.'))||0;
      tEq+=parseFloat(r.equiv||0)||0;
    });
    prodTotEl.textContent=Math.round(tPcs)+' pcs / '+Math.round(tEq*10)/10+' éq';
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

// ── Remplit le groupe "Arrêts prévus" dans la popup interposte ──
function _fillIpArretsPrevus(){
  const g=document.getElementById('ip-arrprev-group');
  if(!g) return;
  g.innerHTML='';
  const ap=_cfgArretsPrevus||{};
  const ARRETS=[
    {key:'pause_min',lbl:'Pause',color:'#94a3b8'},
    {key:'clean_short_min',lbl:'Nettoyage court',color:'#f59e0b'},
    {key:'clean_long_min',lbl:'Nettoyage long',color:'#d97706'},
    {key:'clean_grand_min',lbl:'Nettoyage très long',color:'#b45309'},
    {key:'meeting_tol_min',lbl:'Réunion',color:'#3b82f6'},
  ];
  const visible=ARRETS.filter(a=>(ap[a.key]||0)>0);
  if(!visible.length){g.style.display='none';return;}
  g.style.display='block';
  const title=document.createElement('div');
  title.style.cssText='font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:6px;letter-spacing:.04em';
  title.textContent='Arrêts prévus';
  g.appendChild(title);
  const row=document.createElement('div');
  row.style.cssText='display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;padding-bottom:10px;border-bottom:1px solid var(--border)';
  visible.forEach(a=>{
    const b=document.createElement('button');
    b.className='btn btn-ghost';
    b._origColor=a.color;
    b.style.cssText=`font-size:calc(12px*var(--zf,1));padding:5px 10px;border:2px solid ${a.color};color:${a.color};background:transparent;transition:all .15s`;
    b.textContent=a.lbl+' ('+ap[a.key]+'min max autorisé pendant ce poste)';
    b.onclick=()=>{
      document.getElementById('ip-custom').value=a.lbl;
      document.querySelectorAll('#ip-btns .btn, #ip-arrprev-group .btn').forEach(x=>{
        x.style.background='';x.style.color=x._origColor||'';x.style.borderColor=x._origColor||'var(--border)';x.style.transform='';
      });
      b.style.background=a.color;b.style.color='#fff';b.style.borderColor=a.color;
      b.style.transform='scale(0.93)';setTimeout(()=>{b.style.transform='';},150);
    };
    row.appendChild(b);
  });
  g.appendChild(row);
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
    <div style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1))">
      <span style="flex:1;font-weight:600">${esc(lbl)}</span>
      <button class="btn btn-ghost" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="editInterposteLbl(${i})">✏</button>
      <button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="rmInterposteLbl(${i})">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucun label</div>';
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

// ── Gestion des gaps multiples (trous non déclarés avant chaque OF) ──
let _pendingGaps=[];
let _pendingGapIdx=0;
let _isFirstOfGaps=false;

function _showNextGap(){
  if(_pendingGapIdx>=_pendingGaps.length){goTab('prod');return;}
  const g=_pendingGaps[_pendingGapIdx];
  const total=_pendingGaps.length;
  const idx=_pendingGapIdx+1;
  const dur=_fmtMin(g.duree_s||0);
  // Titre + compteur
  const titleEl=document.getElementById('ps-title');
  if(titleEl) titleEl.textContent='⚠ Période non déclarée';
  const cntEl=document.getElementById('ps-counter');
  if(cntEl) cntEl.textContent=total>1?`Trou ${idx} / ${total}`:'';
  // Plage et durée
  const debut=g.debut||'';const fin=g.fin||'';
  document.getElementById('ps-text').textContent=`${dur} non déclarées : ${debut.replace(':','h')} → ${fin.replace(':','h')}`;
  document.getElementById('ps-gap-s').value=g.duree_s||0;
  document.getElementById('ps-start-iso').value='';
  document.getElementById('ps-custom').value='';
  // Bouton rétrodatage : seulement sur 1er OF + 1er gap
  const bdRow=document.getElementById('ps-backdate-row');
  const bt=document.getElementById('ps-backdate-time');
  const showBd=_isFirstOfGaps&&_pendingGapIdx===0;
  if(bdRow) bdRow.style.display=showBd?'':'none';
  if(bt) bt.textContent=debut.replace(':','h');
  psFillStopBtns();
  openM('m-preshift');
}

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
  _pendingGaps=d.gaps||[];
  _pendingGapIdx=0;
  _isFirstOfGaps=!!(d.shift_model_start);
  // Fallback : si _get_uncovered_gaps n'a rien renvoyé mais gap > 1 min, créer un gap synthétique
  if(_pendingGaps.length===0 && _pendingGapS>=120){
    _pendingGaps=[{debut:d.ip_debut_hms||'',fin:d.ip_fin_hms||'',duree_s:_pendingGapS}];
  }
  _showNextGap();
}

// ── Ignorer ce trou sans déclarer, passer au suivant ──
async function psIgnorer(){
  closeM('m-preshift');
  _pendingGapIdx++;
  _showNextGap();
}

// ── Choix pré-poste ──
function psFillStopBtns(){
  const bc=document.getElementById('ps-stop-btns');
  if(!bc) return;
  bc.innerHTML='';

  const _makeBtn=(lbl,accent)=>{
    const b=document.createElement('button');
    b.className='btn btn-ghost';
    b.style.cssText=`font-size:calc(12px*var(--zf,1));transition:all .15s;border:2px solid ${accent||'var(--border)'};color:${accent||'var(--text)'};margin-bottom:3px`;
    b.textContent=lbl;
    b.onclick=()=>{
      document.getElementById('ps-custom').value=lbl;
      bc.querySelectorAll('.btn').forEach(x=>{x.style.background='';x.style.borderColor='var(--border)';x.style.color='var(--text)';x.style.transform='';});
      b.style.background='var(--navy)';b.style.color='#fff';b.style.borderColor='var(--navy)';
      b.style.transform='scale(0.93)';
      setTimeout(()=>confirmPsAsStop(),180);
    };
    return b;
  };

  const _makeSection=(title,labels,accent)=>{
    if(!labels.length) return;
    const hdr=document.createElement('div');
    hdr.style.cssText='font-size:calc(10px*var(--zf,1));font-weight:800;text-transform:uppercase;color:'+accent+';letter-spacing:.4px;margin:6px 0 4px;border-bottom:1px solid #e5e7eb;padding-bottom:2px';
    hdr.textContent=title;
    bc.appendChild(hdr);
    const row=document.createElement('div');
    row.style.cssText='display:flex;flex-wrap:wrap;gap:5px';
    labels.forEach(lbl=>row.appendChild(_makeBtn(lbl,accent)));
    bc.appendChild(row);
  };

  // Section "Arrêts prévus" depuis _cfgArretsPrevus
  const ap=_cfgArretsPrevus||{};
  const PREVUS_CFG=[
    {key:'pause_min',lbl:'Pause'},
    {key:'meeting_tol_min',lbl:'Réunion'},
    {key:'clean_short_min',lbl:'Nettoyage court'},
    {key:'clean_long_min',lbl:'Nettoyage long'},
    {key:'clean_grand_min',lbl:'Nettoyage très long'},
  ];
  const prevusList=PREVUS_CFG.filter(a=>(ap[a.key]||0)>0).map(a=>a.lbl);
  if(!prevusList.length) prevusList.push(...['Pause','Réunion','Nettoyage court','Nettoyage long']);
  _makeSection('⏱ Arrêts prévus',prevusList,'#d97706');

  // Sections par catégorie depuis _evtsList (paramètres)
  const evts=_evtsList.length?_evtsList:EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]}));
  const prevusSet=new Set(prevusList.map(l=>l.toLowerCase()));
  const cats={pb:[],ratt:[],nettoyage:[],organisation:[],autre:[]};
  const seen=new Set(prevusList.map(l=>l.toLowerCase()));
  evts.forEach(e=>{
    const lbl=e.label||'';if(!lbl||seen.has(lbl.toLowerCase())) return;
    seen.add(lbl.toLowerCase());
    const c=e.cat||'autre';
    if(cats[c]!==undefined) cats[c].push(lbl);
    else cats.autre.push(lbl);
  });
  if(cats.pb.length||cats.ratt.length) _makeSection('🔴 Pannes / Rattrapages',[...cats.pb,...cats.ratt],'#dc2626');
  if(cats.nettoyage.length) _makeSection('🧹 Nettoyage',cats.nettoyage,'#f59e0b');
  if(cats.organisation.length) _makeSection('📋 Organisation',cats.organisation,'#3b82f6');
  if(cats.autre.length) _makeSection('⚫ Autre',cats.autre,'#64748b');

  document.getElementById('ps-custom').value='';
}

async function confirmPsAsStop(){
  const lbl=(document.getElementById('ps-custom').value||'').trim()||'Interposte';
  const gapS=parseFloat(document.getElementById('ps-gap-s').value)||0;
  const g=_pendingGaps[_pendingGapIdx]||{};
  closeM('m-preshift');
  await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({inter_of_s:gapS,label:lbl,comment:'',debut_hms:g.debut||'',fin_hms:g.fin||''})});
  await pollState();
  await pollEvts();
  loadMainDecl();
  _pendingGapIdx++;
  _showNextGap();
}

async function psChooseBackdate(){
  const startIso=document.getElementById('ps-start-iso').value;
  closeM('m-preshift');
  if(startIso){
    await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:startIso})});
    await pollState();
    const bt=document.getElementById('ps-backdate-time');
    const tStr=bt?bt.textContent:'';
    toast('OF rétro-daté à '+(tStr||'l\'heure indiquée'),'ok');
  }
  // Rétrodatage = couvre toute la période → skip tous les gaps restants
  _pendingGapIdx=_pendingGaps.length;
  _showNextGap();
}

function ipShowModifyModel(){
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const model=_cfgModels&&_cfgModels.find(m=>m.nom===(ST.poste||''));
  const jour=model&&model.jours&&model.jours[dk];
  const toHM=iso=>{if(!iso)return'';const d=new Date(iso);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  const deb=(jour&&jour.debut)||(model&&model.debut)||toHM(ST.shift_debut_iso)||'';
  const fin=(jour&&jour.fin)||(model&&model.fin)||toHM(ST.shift_fin_iso)||'';
  document.getElementById('ip-model-debut').value=deb;
  document.getElementById('ip-model-fin').value=fin;
  document.getElementById('ip-model-form').style.display='block';
}

async function ipConfirmModifyModel(){
  const newDebut=document.getElementById('ip-model-debut').value;
  const newFin=document.getElementById('ip-model-fin').value;
  if(!newDebut||!newFin){toast('Remplissez le début et la fin','warn');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mi=_cfgModels.findIndex(m=>m.nom===(ST.poste||''));
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=newDebut;
    _cfgModels[mi].jours[dk].fin=newFin;
    await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:ST.poste||'',day_key:dk,debut:newDebut,fin:newFin})});
  }
  document.getElementById('ip-model-form').style.display='none';
  // Recalculer le gap selon le contexte (fin de poste ou inter-OF)
  if(window._finPosteMode){
    // Gap = fin nouveau modèle - dernière déclaration
    if(_lastProdDeclTime){
      const[fh,fm]=newFin.split(':').map(Number);
      const modelFin=new Date();modelFin.setHours(fh,fm,0,0);
      if(modelFin<_lastProdDeclTime) modelFin.setDate(modelFin.getDate()+1); // poste de nuit
      const newGapS=(modelFin.getTime()-_lastProdDeclTime.getTime())/1000;
      if(newGapS>60){
        _pendingGapS=newGapS;
        document.getElementById('ip-duration').textContent=`Durée non déclarée : ${_fmtMin(newGapS)} (fin OF → fin modèle)`;
        const dh=String(_lastProdDeclTime.getHours()).padStart(2,'0'),dm=String(_lastProdDeclTime.getMinutes()).padStart(2,'0');
        document.getElementById('ip-debut').value=dh+':'+dm;
        document.getElementById('ip-fin').value=newFin;
        toast('Plage mise à jour','ok');
        return; // garder popup ouverte
      }
    }
    // Pas d'écart → aller directement à fin de poste
    window._finPosteMode=false;
    closeM('m-interposte');
    goTab('finposte');
  } else {
    // Contexte interposte normal : recalculer depuis model debut vers maintenant
    const[dh2,dm2]=newDebut.split(':').map(Number);
    const modelDebut=new Date();modelDebut.setHours(dh2,dm2,0,0);
    const now2=new Date();
    const newGapS=(now2.getTime()-modelDebut.getTime())/1000;
    if(newGapS>60){
      _pendingGapS=newGapS;
      document.getElementById('ip-duration').textContent=`Durée : ${_fmtMin(newGapS)} (début de poste → maintenant)`;
      document.getElementById('ip-debut').value=newDebut;
      const nh=String(now2.getHours()).padStart(2,'0'),nm=String(now2.getMinutes()).padStart(2,'0');
      document.getElementById('ip-fin').value=nh+':'+nm;
      toast('Plage mise à jour','ok');
      return; // garder popup ouverte
    }
    // Pas d'écart
    closeM('m-interposte');
    goTab('prod');
  }
}

async function confirmInterposte(){
  const lbl=document.getElementById('ip-custom').value.trim()||'Interposte';
  const cmt=document.getElementById('ip-comment').value.trim();
  const debutVal=document.getElementById('ip-debut').value;
  const finVal=document.getElementById('ip-fin').value;
  let interS=_pendingGapS;
  if(debutVal&&finVal){
    const[dh,dm]=debutVal.split(':').map(Number);const[fh,fm]=finVal.split(':').map(Number);
    let ds=dh*3600+dm*60,fs=fh*3600+fm*60;
    if(fs<ds) fs+=86400;
    interS=Math.max(0,fs-ds);
  }
  closeM('m-interposte');
  await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:interS,label:lbl,comment:cmt,debut_hms:debutVal,fin_hms:finVal})});
  // Si on vient d'un popup pré-poste, rétrodater le shift_start aussi
  if(window._psStartIso){
    await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:window._psStartIso})});
    window._psStartIso=null;
  }
  await pollState();
  await pollEvts();
  loadMainDecl(); // interposte row must appear in accueil without waiting
  if(window._finPosteMode){window._finPosteMode=false;await _doGoFinPoste();}else{goTab('prod');}
}

async function skipInterposte(){
  window._psStartIso=null;
  closeM('m-interposte');
  if(_pendingGapS>30){
    await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:_pendingGapS,label:'Interposte',comment:''})});
  }
  if(window._finPosteMode){window._finPosteMode=false;await _doGoFinPoste();}else{goTab('prod');}
}

async function doCancelProd(){
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
  const g=document.getElementById('sgrid-all'); if(!g) return;
  g.style.cssText='display:block;margin-bottom:6px';
  g.innerHTML='';
  const CAT_ORDER=[
    {key:'manquants',label:'🟣 Manquants',col:'#9333ea'},
    {key:'ratt',label:'🟠 Rattrapage',col:'#dc2626'},
    {key:'organisation',label:'🔵 Organisationnel',col:'#3b82f6'},
    {key:'pb',label:'🔴 Technique',col:'#dc2626'},
    {key:'nettoyage',label:'🧹 Nettoyage',col:'#f59e0b'},
    {key:'autre',label:'⚫ Autre',col:'#64748b'},
  ];
  const bycat={};
  evts.forEach(e=>{const c=e.cat||'autre';if(!bycat[c])bycat[c]=[];bycat[c].push(e);});
  const _bs0='0 8px 20px rgba(0,0,0,.35),inset 0 2px 3px rgba(255,255,255,.35),inset 0 -3px 6px rgba(0,0,0,.25)';
  const _bs1='0 3px 10px rgba(0,0,0,.3),inset 0 1px 2px rgba(255,255,255,.2),inset 0 -1px 3px rgba(0,0,0,.2)';
  let first=true;
  CAT_ORDER.forEach(({key,label,col})=>{
    const items=bycat[key]; if(!items||!items.length) return;
    const hdr=document.createElement('div');
    hdr.style.cssText=`font-size:calc(10px*var(--zf,1));font-weight:800;color:${col};text-transform:uppercase;letter-spacing:.07em;padding:4px 2px;border-bottom:2px solid ${col}40;margin-bottom:6px;${first?'':'margin-top:10px;'}`;
    hdr.textContent=label; g.appendChild(hdr); first=false;
    const grid=document.createElement('div');
    grid.style.cssText='display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:4px';
    items.forEach(e=>{
      const [lt,md,dk]=_STOP_GRAD[e.cat]||_STOP_GRAD.autre;
      const b=document.createElement('button');
      b.style.cssText=`color:#fff;border:none;border-radius:12px;padding:10px 6px;min-height:44px;width:100%;font-size:calc(10px*var(--zf,1));font-weight:700;cursor:pointer;background:radial-gradient(ellipse at 50% 25%,${lt} 0%,${md} 55%,${dk} 100%);box-shadow:${_bs0};transition:transform .12s,box-shadow .12s;position:relative;overflow:hidden;text-shadow:0 1px 3px rgba(0,0,0,.4)`;
      const sheen=document.createElement('span');
      sheen.style.cssText='position:absolute;top:0;left:0;right:0;height:50%;background:linear-gradient(180deg,rgba(255,255,255,.28) 0%,rgba(255,255,255,0) 100%);border-radius:14px 14px 0 0;pointer-events:none';
      const txt=document.createElement('span');
      txt.style.cssText='position:relative;display:block;text-align:center';
      txt.textContent=e.label;
      b.appendChild(sheen); b.appendChild(txt);
      b.onmousedown=()=>{b.style.transform='scale(.95)';b.style.boxShadow=_bs1;};
      b.onmouseup=()=>{b.style.transform='';b.style.boxShadow=_bs0;};
      b.onmouseleave=()=>{b.style.transform='';b.style.boxShadow=_bs0;};
      b.onclick=()=>{closeM('m-stop');doStartStop(e.key,e.cat);};
      grid.appendChild(b);
    });
    g.appendChild(grid);
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

// ── Drag-to-reorder shared state ──
let _dIdx=null;
// Arrêts list drag
function _edDS(i){_dIdx=i;}
function _edDO(e){e.preventDefault();}
function _edDrop(i){if(_dIdx===null||_dIdx===i)return;const m=_evtsEditing.splice(_dIdx,1)[0];_evtsEditing.splice(i,0,m);_dIdx=null;_renderEvtListHTML();}
// Dégradé list drag
function _ddDS(i){_dIdx=i;}
function _ddDO(e){e.preventDefault();}
function _ddDrop(i){if(_dIdx===null||_dIdx===i)return;const m=_degradeListLocal.splice(_dIdx,1)[0];_degradeListLocal.splice(i,0,m);_dIdx=null;renderDegradeList(_degradeListLocal);}
// Pwd list drag
function _pdDS(i){_dIdx=i;}
function _pdDO(e){e.preventDefault();}
function _pdDrop(i){if(_dIdx===null||_dIdx===i)return;const keys=Object.keys(_cfgPwds);const[rm]=keys.splice(_dIdx,1);keys.splice(i,0,rm);const nw={};keys.forEach(k=>{nw[k]=_cfgPwds[k];});_cfgPwds=nw;_dIdx=null;renderPwdList();}

function _renderEvtListHTML(){
  const c=document.getElementById('events-list-ui');if(!c) return;
  const catLbl={pb:'🔴 Technique',ratt:'🟠 Rattrapage',nettoyage:'🟡 Nettoyage',organisation:'🔵 Organisationnel',manquants:'🟣 Manquants',autre:'⚫ Autre'};
  c.innerHTML=_evtsEditing.map((e,i)=>`
    <div draggable="true" ondragstart="_edDS(${i})" ondragover="_edDO(event)" ondrop="_edDrop(${i})" style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1));cursor:default">
      <span style="cursor:grab;color:#94a3b8;font-size:16px;padding:0 2px;user-select:none" title="Déplacer">⠿</span>
      <span style="flex:1;font-weight:600">${esc(e.label)}</span>
      <span style="font-size:calc(10px*var(--zf,1));color:var(--gray)">${catLbl[e.cat]||e.cat}</span>
      <button class="btn btn-ghost" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="editEvtItem(${i})">✏</button>
      <button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="rmEvtItem(${i})">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucun arrêt configuré</div>';
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
  if(_curTab==='main') loadMainDecl();
}

async function doReunion(){
  try{await fetch('/api/toggle_reunion',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});}
  catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
  if(_curTab==='main') loadMainDecl();
}

function doNettoyage(){
  // Populate budget labels from config
  const ap=_cfgArretsPrevus||{};
  const fmtMin=m=>m>0?`Budget autorisé : ${m} min`:'Non limité';
  const sc=document.getElementById('nett-lbl-court'),sl=document.getElementById('nett-lbl-long'),sg=document.getElementById('nett-lbl-grand');
  if(sc) sc.textContent=fmtMin(ap.clean_short_min||0);
  if(sl) sl.textContent=fmtMin(ap.clean_long_min||0);
  if(sg) sg.textContent=fmtMin(ap.clean_grand_min||0);
  openM('m-nett-type');
}

async function doStartNettoyage(ntype){
  closeM('m-nett-type');
  try{await fetch('/api/start_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ntype})});}
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
    fetch('/api/end_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'}).then(()=>{pollState();pollEvts();if(_curTab==='main')loadMainDecl();});
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
  if(_curTab==='main') loadMainDecl();
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
    f[k]=el.type==='number'?(el.value===''?'':parseFloat(el.value)||0):el.value;
  });
  f.pilote=document.getElementById('f-pilote')?.value||ST.pilot||'';
  f.poste=document.getElementById('f-poste')?.value||ST.poste||'';
  return f;
}

// ── Traca Fibre multi-lots ────────────────────────────────────────────────────
function addTracaRow(val){
  if(val===undefined)val='';
  const ui=document.getElementById('f-traca-ui');if(!ui)return;
  const row=document.createElement('div');
  row.className='traca-row';row.style.cssText='display:flex;gap:4px;align-items:center';
  const inp=document.createElement('input');
  inp.type='text';inp.className='traca-input';inp.value=val;
  inp.placeholder='n° de lot / traca';
  inp.style.cssText='flex:1;padding:5px 7px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));min-width:0';
  inp.oninput=syncTracaField;
  const btn=document.createElement('button');
  btn.type='button';btn.textContent='×';
  btn.style.cssText='background:#fee2e2;border:none;border-radius:4px;color:#dc2626;font-weight:900;padding:3px 8px;cursor:pointer;flex-shrink:0;font-size:calc(13px*var(--zf,1));line-height:1';
  btn.onclick=function(){removeTracaRow(this);};
  row.appendChild(inp);row.appendChild(btn);ui.appendChild(row);
}
function removeTracaRow(btn){
  const ui=document.getElementById('f-traca-ui');if(!ui)return;
  if(ui.querySelectorAll('.traca-row').length<=1)return;
  btn.closest('.traca-row').remove();
  syncTracaField();scheduleAutoSave();
}
function syncTracaField(){
  const val=[...document.querySelectorAll('#f-traca-ui .traca-input')].map(i=>i.value.trim()).filter(Boolean).join(';');
  const h=document.getElementById('f-traca');if(h)h.value=val;
  scheduleAutoSave();
}
function fillTracaUI(val){
  const ui=document.getElementById('f-traca-ui');if(!ui)return;
  ui.innerHTML='';
  const parts=(val||'').split(';').map(s=>s.trim()).filter(Boolean);
  if(!parts.length)parts.push('');
  parts.forEach(p=>addTracaRow(p));
}
// ── fin Traca multi-lots ───────────────────────────────────────────────────────

function fillFormFromState(form){
  if(!form) return;
  FORM_FIELDS.forEach(k=>{
    const el=document.getElementById('f-'+k);
    if(!el) return;
    const v=form[k];
    if(v!==undefined&&v!==null&&v!=='') el.value=v;
  });
  // Ensure nb_pers defaults to 10 if not set or 0
  const npEl=document.getElementById('f-nb_pers');
  if(npEl&&(!npEl.value||npEl.value==='0')) npEl.value='10';
  const ofEl=document.getElementById('pob-of');
  if(ofEl) ofEl.textContent=form.of_num||'—';
  fillTracaUI(form.traca||'');
}

// Form persistence in localStorage (persist across restarts until new prod)
function saveFormToStorage(){
  const f=collectForm();
  try{localStorage.setItem('kpiorc_form',JSON.stringify(f));}catch(e){}
}
function restoreFormFromStorage(){
  try{
    const raw=localStorage.getItem('kpiorc_form');
    if(raw){
      const f=JSON.parse(raw);
      // Only restore if prod_active and form has of_num, or if not active (post-prod)
      FORM_FIELDS.forEach(k=>{
        const el=document.getElementById('f-'+k);
        if(!el||!(k in f)) return;
        const v=f[k];
        if(v!==undefined&&v!==null&&v!=='') el.value=v;
      });
      const npEl=document.getElementById('f-nb_pers');
      if(npEl&&(!npEl.value||npEl.value==='0')) npEl.value='10';
    }
  }catch(e){}
  fillTracaUI(document.getElementById('f-traca')?.value||'');
}

function _checkEndProdBtn(){
  const req=['f-of_num','f-code_prod','f-type_prod','f-nb_pers','f-qte_fab','f-qte_emb','f-poids','f-taille','f-fibre'];
  const ok=req.every(id=>{const el=document.getElementById(id);return el&&(el.value||'').trim()!==''&&el.value!=='0';});
  const btn=document.getElementById('btn-endprod'); if(!btn) return;
  btn.style.opacity=ok?'1':'0.38';
  btn.style.cursor=ok?'pointer':'default';
  // Auto-clear highlight when field is now filled
  req.forEach(id=>{
    const el=document.getElementById(id);
    if(el&&(el.value||'').trim()!==''&&el.value!=='0') el.classList.remove('field-missing');
  });
}
function _clearFieldHighlights(){
  ['f-of_num','f-code_prod','f-type_prod','f-nb_pers','f-qte_fab','f-qte_emb','f-poids','f-taille','f-fibre'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.classList.remove('field-missing');
  });
}
function scheduleAutoSave(){
  clearTimeout(_autoSaveTimer);
  _checkEndProdBtn();
  _autoSaveTimer=setTimeout(()=>{
    const f=collectForm();
    fetch('/api/save_form',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(f)});
    saveFormToStorage();
  },1500);
}

// ── END PROD ──
async function doEndProdPreview(){
  const req=['f-of_num','f-code_prod','f-type_prod','f-nb_pers','f-qte_fab','f-qte_emb','f-poids','f-taille','f-fibre'];
  const missing=req.filter(id=>{const el=document.getElementById(id);return !el||(el.value||'').trim()===''||el.value==='0';});
  if(missing.length){
    missing.forEach(id=>{const el=document.getElementById(id);if(el){el.classList.remove('field-missing');void el.offsetWidth;el.classList.add('field-missing');}});
    toast('⚠ Formulaire incomplet','err',1000);
    return;
  }
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
    <div class="ep-stat"><span class="lbl">N° OF</span><span class="val">${esc(ofNum||'—')}</span></div>
    <div class="ep-stat"><span class="lbl">Type produit</span><span class="val">${esc(f.type_prod||d.type_prod||'—')}</span></div>
    <div class="ep-stat"><span class="lbl">Qté fabriquée</span><span class="val">${esc(String(f.qte_fab||d.qte_fab||0))}</span></div>
    <div class="ep-stat"><span class="lbl">Qté emballée</span><span class="val">${esc(String(f.qte_emb||d.qte_emb||0))}</span></div>
    <div class="ep-stat"><span class="lbl">Date</span><span class="val">${dateStr}</span></div>
    <div class="ep-stat"><span class="lbl">TRS OF</span><span class="val">${fmtTRS(d.trs)}</span></div>
    <div class="ep-stat"><span class="lbl">Équivalence</span><span class="val">${(d.equiv||0).toFixed(1)}</span></div>
    <div class="ep-stat"><span class="lbl">Durée prod</span><span class="val">${fmtD2(d.prod_s||0)}</span></div>
    <div class="ep-stat"><span class="lbl">Total arrêts</span><span class="val">${fmtD2(d.stop_s||0)}</span></div>
  `;
  const evts=d.tl_events||[];
  const stopMap={};
  const totalS=d.stop_s||1;
  evts.forEach(e=>{
    if(!e.key||e.key==='prod'||e.key.startsWith('_')) return;
    const dur=(e.start&&e.end)?(new Date(e.end).getTime()-new Date(e.start).getTime())/1000:(e.dur_s||0);
    stopMap[e.key]=(stopMap[e.key]||0)+dur;
  });
  const tbody=document.getElementById('ep-stops');
  tbody.innerHTML=Object.entries(stopMap).map(([k,s])=>`
    <tr><td>${getEvtLabel(k)}</td><td>${fmtD2(s)}</td><td>${Math.round(s/totalS*100)}%</td></tr>
  `).join('')||'<tr><td colspan="3" style="color:var(--gray)">Aucun arrêt</td></tr>';
  const epDisplayEvts=tlEventsToDisplayFmt(evts);
  const _epS=parseHMStoT(d.debut,null)||(Date.now()-3600000);
  const _epE=parseHMStoT(d.now_str,null)||Date.now();
  drawTLFromISO('ep-tl',epDisplayEvts,new Date(_epS).toISOString(),new Date(_epE).toISOString());
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
  await fetch('/api/force_reset_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
  toast('Production annulée','warn');
}

// ── PIE & GAUGE CHARTS ──
function drawPie(svgId, segments, opts) {
  const svg=document.getElementById(svgId);if(!svg) return;
  const fCenter=(opts&&opts.fCenter)||13;
  const fSub=(opts&&opts.fSub)||7;
  const fLeg=(opts&&opts.fLeg)||7;
  const total=segments.reduce((a,s)=>a+s.value,0);
  if(total<=0){svg.innerHTML=`<text x="65" y="60" text-anchor="middle" font-size="${fCenter}" fill="#94a3b8">Pas de données</text>`;return;}
  const cx=65,cy=57,r=44,ir=24;let html='',startAngle=-Math.PI/2;
  const visSegs=segments.filter(s=>s.value>0&&(s.value/total)*2*Math.PI>=0.001);
  if(visSegs.length===1){
    // 100% — arc dégénéré : dessiner un anneau plein
    html+=`<circle cx="${cx}" cy="${cy}" r="${r}" fill="${visSegs[0].color}"/>`;
    html+=`<circle cx="${cx}" cy="${cy}" r="${ir}" fill="#fff"/>`;
  } else {
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
  }
  const m=segments[0],mp=total>0?Math.round(m.value/total*100):0;
  html+=`<text x="${cx}" y="${cy+5}" text-anchor="middle" font-size="${fCenter}" font-weight="800" fill="#1a1f5e">${mp}%</text>`;
  html+=`<text x="${cx}" y="${cy+15}" text-anchor="middle" font-size="${fSub}" fill="#64748b">${esc(m.label)}</text>`;
  const legBase=(opts&&opts.legY)||108;
  let lx=0;segments.filter(s=>s.value>0).forEach(s=>{
    const p=Math.round(s.value/total*100);
    html+=`<rect x="${lx}" y="${legBase}" width="8" height="8" fill="${s.color}" rx="1"/>`;
    html+=`<text x="${lx+11}" y="${legBase+8}" font-size="${fLeg}" fill="#475569" font-weight="600">${esc(s.label)} ${p}%</text>`;lx+=65;
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
function renderBudgetBars(containerId,bs){
  const el=document.getElementById(containerId);
  if(!el||!bs||!bs.per_type) return;
  const TYPES=['Pause','Réunion','Nettoyage court','Nettoyage long','Nettoyage très long'];
  const ap=_cfgArretsPrevus||{};
  let html='';
  let anyBar=false;
  TYPES.forEach(lbl=>{
    const d=bs.per_type[lbl]||{};
    const budget=d.budget_s||0;
    if(budget<=0) return;
    anyBar=true;
    const consumed=d.consumed_s||0;
    const pct=budget>0?Math.min(100,consumed/budget*100):0;
    const over=Math.max(0,consumed-budget);
    const color=pct>=100?'#dc2626':pct>=70?'#d97706':'#16a34a';
    const valStr=over>0?`<b>+${fmtDurShort(over)}</b>`:`${fmtDurShort(consumed)} / ${fmtDurShort(budget)}`;
    html+=`<div style="margin-bottom:5px">
      <div style="display:flex;justify-content:space-between;font-size:calc(10px*var(--zf,1));font-weight:700;margin-bottom:2px">
        <span style="color:#374151">${esc(lbl)}</span>
        <span style="color:${color}">${valStr}</span></div>
      <div style="background:#e5e7eb;border-radius:4px;height:6px">
        <div style="background:${color};width:${Math.min(100,pct).toFixed(0)}%;height:6px;border-radius:4px"></div>
      </div></div>`;
  });
  el.innerHTML=anyBar?html:'<div style="color:#92400e;font-size:calc(10px*var(--zf,1));opacity:.7">Aucun budget configuré</div>';
}

// ── Mode dégradé ──────────────────────────────────────────────────────────────
function toggleDegrade(){
  if(window._degradeActive) stopDegrade();
  else openDegradeModal();
}
function openDegradeModal(){
  const motifs=window._degradeMotifs||[];
  const bd=document.getElementById('m-degrade-body');
  const tl=document.getElementById('m-degrade-title');
  if(tl) tl.textContent='🟡 MODE DÉGRADÉ — Choisir le motif';
  if(!bd) return;
  if(!motifs.length){
    bd.innerHTML='<div style="color:#fbbf24;font-size:calc(12px*var(--zf,1));text-align:center;padding:12px;background:rgba(220,38,38,.15);border-radius:10px;border:1px solid rgba(220,38,38,.3)">Aucun motif configuré.<br>Veuillez d\'abord les ajouter dans les Paramètres.</div>';
    document.getElementById('m-degrade-confirm').style.display='none';
  } else {
    let html='<div style="display:flex;flex-direction:column;gap:8px">';
    motifs.forEach(function(m,i){
      html+=`<label style="display:flex;align-items:center;gap:10px;padding:12px 16px;border:1.5px solid rgba(202,138,4,.35);border-radius:12px;cursor:pointer;background:linear-gradient(135deg,rgba(202,138,4,.15) 0%,rgba(146,64,14,.18) 100%);box-shadow:0 4px 14px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.08);transition:all .15s" onmouseover="this.style.boxShadow='0 0 20px rgba(202,138,4,.4),0 4px 14px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.12)'" onmouseout="this.style.boxShadow='0 4px 14px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.08)'"><input type="radio" name="deg-motif" value="${esc(m)}" ${i===0?'checked':''}> <span style="font-weight:800;color:#fcd34d;font-size:calc(13px*var(--zf,1));text-shadow:0 0 10px rgba(202,138,4,.5)">${esc(m)}</span></label>`;
    });
    html+='</div>';
    bd.innerHTML=html;
    document.getElementById('m-degrade-confirm').style.display='';
  }
  openM('m-degrade');
}
async function _confirmDegrade(){
  const r=document.querySelector('input[name="deg-motif"]:checked');
  if(!r) return;
  closeM('m-degrade');
  try{
    await fetch('/api/start_degrade',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({motif:r.value})});
  }catch(e){toast('Erreur connexion','err');}
  await pollState();
}
async function stopDegrade(){
  try{
    await fetch('/api/stop_degrade',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
  }catch(e){toast('Erreur connexion','err');}
  await pollState();
}
// ── Budget override ─────────────────────────────────────────────────────────
function openBudgetOverrideModal(){
  const bs=window._lastBudgetState;
  const ov=window._budgetOverrides||{};
  const LABELS={'clean_short_min':'🧹 Nettoyage court','clean_long_min':'🧹 Nettoyage long','clean_grand_min':'🧹 Nettoyage très long','meeting_tol_min':'📋 Réunion','pause_min':'⏸ Pause'};
  let html='';
  Object.entries(LABELS).forEach(function([k,lbl]){
    const bgt=bs&&bs.per_type&&bs.per_type[lbl.replace(/^[^\w]*\s/,'').trim()];
    const defVal=ov[k]!==undefined?ov[k]:(bgt?Math.round(bgt.budget_s/60):0);
    html+=`<div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:6px 10px"><span style="flex:1;font-size:calc(11px*var(--zf,1));font-weight:600">${lbl}</span><input type="number" id="bov-${k}" value="${defVal}" min="0" max="240" style="width:65px;padding:4px 7px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));text-align:right"><span style="font-size:calc(10px*var(--zf,1));color:var(--gray)">min</span></div>`;
  });
  const f=document.getElementById('bov-fields');
  if(f) f.innerHTML=html;
  const p=document.getElementById('bov-pw');if(p)p.value='';
  openM('m-budget-override');
}
async function saveBudgetOverride(){
  const pw=document.getElementById('bov-pw')?.value||'';
  const overrides={};
  ['clean_short_min','clean_long_min','clean_grand_min','meeting_tol_min','pause_min'].forEach(function(k){
    const el=document.getElementById('bov-'+k);
    if(el) overrides[k]=parseFloat(el.value)||0;
  });
  try{
    const r=await fetch('/api/set_budget_override',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,overrides})});
    const d=await r.json();
    if(d&&d.ok){closeM('m-budget-override');toast('Budget modifié pour ce poste');await pollState();}
    else toast(d?.error||'Mot de passe incorrect','err');
  }catch(e){toast('Erreur connexion','err');}
}
// ── Liste dégradé dans Paramètres ───────────────────────────────────────────
function renderDegradeList(motifs){
  const ul=document.getElementById('degrade-list-ui');if(!ul)return;
  if(!motifs||!motifs.length){ul.innerHTML='<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:6px 0">Aucun motif configuré.</div>';return;}
  ul.innerHTML=motifs.map(function(m,i){
    return `<div draggable="true" ondragstart="_ddDS(${i})" ondragover="_ddDO(event)" ondrop="_ddDrop(${i})" style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1));cursor:default"><span style="cursor:grab;color:#94a3b8;font-size:16px;padding:0 2px;user-select:none" title="Déplacer">⠿</span><span style="flex:1;font-weight:600">${esc(m)}</span><button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="removeDegradeItem(${i})">✕</button></div>`;
  }).join('');
}
function addDegradeItem(){
  const inp=document.getElementById('deg-new-label');
  if(!inp||!inp.value.trim()) return;
  _degradeListLocal=_degradeListLocal||[];
  _degradeListLocal.push(inp.value.trim());
  inp.value='';
  renderDegradeList(_degradeListLocal);
}
function removeDegradeItem(i){
  if(!_degradeListLocal) return;
  _degradeListLocal.splice(i,1);
  renderDegradeList(_degradeListLocal);
}
let _degradeListLocal=[];
async function saveDegradeList(){
  const pw=_adminPw||'';
  try{
    const r=await fetch('/api/save_degrade_list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,motifs:_degradeListLocal})});
    const d=await r.json();
    if(d&&d.ok) toast('Liste dégradé enregistrée');
    else toast(d?.error||'Erreur','err');
  }catch(e){toast('Erreur connexion','err');}
}
// ── Influence nb opérateur JS ────────────────────────────────────────────────
let _persPctMapLocal={};
function renderPersPctTable(){
  const tb=document.getElementById('pers-pct-tbody');
  if(!tb) return;
  tb.innerHTML='';
  for(let n=1;n<=10;n++){
    const pct=_persPctMapLocal[String(n)]??'';
    const tr=document.createElement('tr');
    tr.innerHTML=`<td style="padding:5px 10px;border:1px solid var(--border);font-weight:600">${n} pers.</td>`+
      `<td style="padding:5px 10px;border:1px solid var(--border)"><input type="number" id="pct-n-${n}" min="0" max="200" step="0.1" value="${pct}" style="width:80px;padding:3px 6px;border:1.5px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1))"> %</td>`;
    tb.appendChild(tr);
  }
}
async function savePersPct(){
  const entries=[];
  for(let n=1;n<=10;n++){
    const el=document.getElementById('pct-n-'+n);
    if(!el||!el.value) continue;
    const pct=parseFloat(el.value);
    if(isNaN(pct)||pct<=0) continue;
    entries.push({nb_pers:n,pct:pct});
  }
  try{
    const r=await fetch('/api/save_pers_pct',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,entries})});
    const d=await r.json();
    if(d&&d.ok){toast('Tableau nb opérateur enregistré','ok');await loadCfg();}
    else toast(d?.error||'Erreur','err');
  }catch(e){toast('Erreur connexion','err');}
}
// ── fin Influence nb opérateur JS ────────────────────────────────────────────

function updateGauge(s){
  const arc=document.getElementById('gauge-arc');
  const pct=document.getElementById('gauge-pct');
  const pobTrs=document.getElementById('pob-trs');
  if(!arc||!pct) return;
  // Budget arrêts prévus — déduction pour TRS OF et TRS Poste
  const bs=s.budget_state||{};
  const ofDed=(bs.total_of_deductible_s)||0;
  const shiftDed=(bs.total_shift_deductible_s)||0;
  // Estimate live TRS using coefficient
  const prodRef=s.prod_ref||200;
  const ofS=s.of_elapsed_s||0;
  const effOfS=Math.max(1,ofS-ofDed);
  const qFab=s.form?parseFloat(s.form.qte_fab||0):0;
  const typeProd=s.form?s.form.type_prod||'':'';
  const coef=(window._equivCoefs&&typeProd&&window._equivCoefs[typeProd])||1;
  const equiv=qFab*coef;
  // Calculer le temps dégradé pour cet OF (live)
  let _degS=0;
  const _ofStartMs=s.of_start_iso?new Date(s.of_start_iso).getTime():0;
  const _nowMs=Date.now();
  if(s.degrade_active&&s.degrade_start_iso&&_ofStartMs>0){
    const _ds=Math.max(new Date(s.degrade_start_iso).getTime(),_ofStartMs);
    _degS=Math.max(0,(_nowMs-_ds)/1000);
  }
  (s.degrade_periods_iso||[]).forEach(function(p){
    if(!p.start||!p.end||!_ofStartMs) return;
    const _d0=Math.max(new Date(p.start).getTime(),_ofStartMs);
    const _d1=Math.min(new Date(p.end).getTime(),_nowMs);
    if(_d1>_d0) _degS+=(_d1-_d0)/1000;
  });
  const _adjS=Math.max(1,effOfS-_degS/2);
  const _nbPersLive=s.form?parseInt(s.form.nb_pers||1)||1:1;
  const _pctLive=(_persPctMapLocal&&_persPctMapLocal[String(_nbPersLive)])?(_persPctMapLocal[String(_nbPersLive)]/100):1.0;
  let trs=-1;
  if(ofS>0&&prodRef>0&&equiv>0){
    trs=Math.round(equiv/(prodRef*_pctLive*_adjS/28800)*100*10)/10;
  }
  // Mise à jour barres budget accueil + prod en cours
  renderBudgetBars('budget-bars-acc',bs);
  renderBudgetBars('budget-bars-prod',bs);
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
  // TRS Poste: prefer shift_debut_iso (set at login from model), fall back to model config
  if(s.shift_debut_iso){_shiftRefDt=new Date(s.shift_debut_iso);}
  else{
    const _dayKeysG=['dim','lun','mar','mer','jeu','ven','sam'];
    const _dkG=_dayKeysG[new Date().getDay()];
    const _modelG=_cfgModels&&_cfgModels.find(m=>m.nom===(s.poste||ST.poste||''));
    const _jourG=_modelG&&_modelG.jours&&_modelG.jours[_dkG];
    if(_jourG&&_jourG.debut){const[_hG,_mG]=_jourG.debut.split(':').map(Number);_shiftRefDt=new Date();_shiftRefDt.setHours(_hG,_mG,0,0);}
    else if(s.shift_start_iso){_shiftRefDt=new Date(s.shift_start_iso);}
    else{_shiftRefDt=null;}
  }
  // TRS Accueil : début de plage → fin de la dernière déclaration de prod, avec déduction budget
  // Quand OF en cours, n'utiliser que les arrêts AVANT l'OF actif pour éviter le bug TRS 150%
  if(_shiftRefDt&&s.prod_ref>0){
    const calcRef=_lastProdDeclTime||new Date();
    const shiftElap=(calcRef.getTime()-_shiftRefDt.getTime())/1000;
    const shiftDedForPoste=s.prod_active?((bs.shift_deductible_before_of_s)||0):shiftDed;
    const effShiftElap=Math.max(1,shiftElap-shiftDedForPoste);
    const todayEquiv=_todayEquivAccum||0;
    const trsPoste=effShiftElap>0&&todayEquiv>0?Math.round(todayEquiv/(s.prod_ref*effShiftElap/28800)*100*10)/10:-1;
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
  drawPie('pie-of',[{label:'Prod',value:prodSof,color:'#16a34a'},{label:'Arrêts',value:stopS,color:'#dc2626'}],{fCenter:20,fSub:12,fLeg:12,legY:118});
  // For poste pie — compute from shift start
  const shiftTotal=s.shift_start_iso?(Date.now()-new Date(s.shift_start_iso).getTime())/1000:0;
  const shiftStop=_todayStopAccum||0;
  const shiftProd=Math.max(0,shiftTotal-shiftStop);
  const postePieData=[{label:'Prod',value:shiftProd,color:'#16a34a'},{label:'Arrêts',value:shiftStop,color:'#dc2626'}];
  drawPie('pie-poste',postePieData,{fCenter:16,fSub:9,fLeg:9});
  drawPie('pie-poste-acc',postePieData,{fCenter:16,fSub:9,fLeg:9,legY:118});
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

let _pwaCallback=null;
function _showPwAction(title,cb){
  document.getElementById('pwa-title').textContent=title;
  document.getElementById('pwa-pw').value='';
  _pwaCallback=cb;
  openM('m-pw-action');
  setTimeout(()=>document.getElementById('pwa-pw').focus(),80);
}
function _pwaConfirm(){
  const pw=document.getElementById('pwa-pw').value;
  closeM('m-pw-action');
  if(_pwaCallback){_pwaCallback(pw);_pwaCallback=null;}
}

function saveEditRow() {
  _showPwAction('🔒 Confirmer modification',async function(pw){
  const rowNum=parseInt(document.getElementById('er-rownum').value);
  const rowType=document.getElementById('er-rowtype').value;
  if(!rowNum){toast('Ligne invalide','err');return;}
  if(!pw){toast('Mot de passe requis','err');return;}
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
  });
}

function deleteRow(key,rowNumId) {
  const rn=rowNumId?parseInt(document.getElementById(rowNumId)?.value):parseInt(window._rowMap[String(key)]?.row_num);
  if(!rn) return;
  _showPwAction('🗑 Confirmer suppression',async function(pw){
    if(!pw){toast('Mot de passe requis','err');return;}
    const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:rn})});
    const d=r?await r.json():{};
    if(d&&d.ok){closeM('m-editrow');await loadMainDecl();if(_curTab==='history')await loadHist();loadKPI();toast('Supprimé','ok');}
    else toast(d?.error||'Erreur suppression','err');
  });
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
  document.getElementById('es-pw').value='';
  document.getElementById('es-type').value=ev.type||key||'';
  const d=ev.debut||'',f2=ev.fin||'';
  document.getElementById('es-deb').value=d.length>=5?d.slice(0,5):d;
  document.getElementById('es-fin').value=f2.length>=5?f2.slice(0,5):f2;
  document.getElementById('es-cmt').value=ev.comment||'';
  openM('m-editstop');
  setTimeout(()=>{const pw=document.getElementById('es-pw');if(pw)pw.focus();},80);
}

async function saveEditStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev) return;
  const pw=document.getElementById('es-pw').value||'';
  const data={pw,row_num:ev.row_num,type:document.getElementById('es-type').value,heure_debut:document.getElementById('es-deb').value,heure_fin:document.getElementById('es-fin').value,comment:document.getElementById('es-cmt').value};
  const r=await fetch('/api/edit_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  const d=r?await r.json():{};
  if(d&&d.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();loadKPI();toast('Modifié','ok');}
  else toast(d?.error||'Mot de passe incorrect','err');
}

async function deleteStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev) return;
  const pw=document.getElementById('es-pw').value||'';
  if(!pw){toast('Mot de passe requis','err');return;}
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:ev.row_num})});
  const d=r?await r.json():{};
  if(d&&d.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();loadKPI();toast('Supprimé','ok');}
  else toast(d?.error||'Mot de passe incorrect','err');
}

// ── TIMELINE ──
function renderTL(svgId,evts){
  const now=new Date(), s4h=new Date(now-4*3600*1000);
  drawTLFromISO(svgId,evts,s4h.toISOString(),now.toISOString());
}

function drawTL(svgId,tlEvts,debutHMS,finHMS){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=28,H=52;
  const _dpId='dpat_'+svgId;
  let html=`<defs><pattern id="${_dpId}" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="4" height="8" fill="#16a34a"/><rect x="4" y="0" width="4" height="8" fill="#fef08a"/></pattern></defs>`;
  html+=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
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
  // Dégradé overlay (periodes closes + active)
  const _ofStartMs=ST.of_start_iso?new Date(ST.of_start_iso).getTime():tS;
  (window._degradePeriodsIso||[]).forEach(function(p){
    if(!p.start||!p.end) return;
    const _d0=Math.max(new Date(p.start).getTime(),_ofStartMs);
    const _d1=new Date(p.end).getTime();
    if(_d1<=_d0) return;
    const x1=toX(_d0),x2=toX(_d1);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="url(#${_dpId})" rx="2" opacity=".85"/>`;
  });
  if(window._degradeActive&&ST.degrade_start_iso){
    const _d0=Math.max(new Date(ST.degrade_start_iso).getTime(),_ofStartMs);
    const _d1=Date.now();
    if(_d1>_d0){const x1=toX(_d0),x2=toX(_d1);if(x2>x1)html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="url(#${_dpId})" rx="2" opacity=".85"/>`;}
  }
  html+=`<text x="2" y="${H-2}" font-size="10" fill="#374151" font-weight="600">${debutHMS.slice(0,5)}</text>`;
  html+=`<text x="${W-36}" y="${H-2}" font-size="10" fill="#374151" font-weight="600">${finHMS.slice(0,5)}</text>`;
  svg.innerHTML=html;
}

function drawTLFromISO(svgId,evts,startIso,endIso,prodOfList){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=35,H=64;
  const _dpId='dpat_'+svgId;
  let html=`<defs><pattern id="${_dpId}" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="4" height="8" fill="#16a34a"/><rect x="4" y="0" width="4" height="8" fill="#fef08a"/></pattern></defs>`;
  html+=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
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
    const fillCol=ev.is_degrade?`url(#${_dpId})`:(STOP_COL[cat]||'#94a3b8');
    html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${fillCol}" rx="2" opacity=".85"/>`;
  });
  // Current live stop
  if(_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active){
    const se=_curStopElap+(Date.now()-_lastPoll)/1000;
    const sT=tE-se*1000;
    const x1=toX(sT),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${STOP_COL.pb||'#b91c1c'}" rx="2" opacity=".9"/>`;
  }
  // Dégradé overlay (periodes closes + active)
  const _ofStartMsTL=ST.of_start_iso?new Date(ST.of_start_iso).getTime():tS;
  (window._degradePeriodsIso||[]).forEach(function(p){
    if(!p.start||!p.end) return;
    const _d0=Math.max(new Date(p.start).getTime(),_ofStartMsTL);
    const _d1=new Date(p.end).getTime();
    if(_d1<=_d0) return;
    const x1=toX(_d0),x2=toX(_d1);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="url(#${_dpId})" rx="2" opacity=".85"/>`;
  });
  if(window._degradeActive&&ST.degrade_start_iso&&ST.prod_active){
    const _d0=Math.max(new Date(ST.degrade_start_iso).getTime(),_ofStartMsTL);
    const _d1=Date.now();
    if(_d1>_d0){const x1=toX(_d0),x2=toX(_d1);if(x2>x1)html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="url(#${_dpId})" rx="2" opacity=".85"/>`;}
  }
  const fT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  // Hourly tick marks — labels centered on tick, guarded 58px from each edge to avoid overlap with start/end labels
  const TLGUARD=58;
  let tickT=Math.ceil(tS/3600000)*3600000;
  while(tickT<tE){
    const tx=toX(tickT);
    const hr=new Date(tickT).getHours();
    html+=`<line x1="${tx}" y1="${Y}" x2="${tx}" y2="${Y+H2}" stroke="rgba(0,0,0,.2)" stroke-width="1"/>`;
    if(tx>TLGUARD && tx<W-TLGUARD)
      html+=`<text x="${tx}" y="${Y+H2+16}" text-anchor="middle" font-size="11" fill="#374151" font-weight="600">${String(hr).padStart(2,'0')}h</text>`;
    tickT+=3600000;
  }
  // Start and end labels always drawn at edges (never overlap with guarded ticks)
  html+=`<text x="2" y="${Y+H2+16}" font-size="13" fill="#374151" font-weight="600">${fT(tS)}</text>`;
  html+=`<text x="${W-2}" y="${Y+H2+16}" text-anchor="end" font-size="13" fill="#374151" font-weight="600">${fT(tE)}</text>`;
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
  if(!stops.length){c.innerHTML='<div style="color:var(--gray);font-size:calc(10px*var(--zf,1));padding:4px">Aucun arrêt</div>';return;}
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
      <div class="si-nm" style="font-size:calc(9px*var(--zf,1))">${lbl}</div>
      <div class="si-dur" style="font-size:calc(9px*var(--zf,1))">…</div>
    </div>`;
  }
  c.innerHTML=html;
}

function calcDur(d,f){
  if(!d||!f) return '';
  try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
  const diff=p(f)-p(d);return diff>0?fmtDur(diff):'';}catch(e){return '';}
}

// ── CODE INPUT OVERLAY ──
let _codeInputTarget = null;
let _csDigits = ['','','','','','','','',''];
let _csFormat = '6_3'; // '9' for 9-digit straight, '6_3' for 6+separator+3
function _csRender(){
  let allFilled=true;
  for(let i=0;i<9;i++){
    const el=document.getElementById('cs-'+i);if(!el) continue;
    const d=_csDigits[i]||'';
    el.textContent=d||'X';
    el.className='cs'+(d?' cs-filled':'');
    if(!d) allFilled=false;
  }
  if(allFilled){for(let i=0;i<9;i++){const e=document.getElementById('cs-'+i);if(e)e.classList.add('cs-done');}}
  else{
    const ai=_csDigits.findIndex(d=>!d);
    const activeEl=document.getElementById('cs-'+ai);
    if(activeEl) activeEl.classList.add('cs-active');
  }
}
function _csKeydown(e){
  if(/^[0-9]$/.test(e.key)){
    e.preventDefault();
    const pos=_csDigits.findIndex(d=>!d);
    if(pos!==-1){_csDigits[pos]=e.key;_csRender();if(_csDigits.every(d=>d))setTimeout(_codeInputConfirm,80);}
  } else if(e.key==='Backspace'){
    e.preventDefault();
    let pos=_csDigits.findIndex(d=>!d);
    if(pos===-1) pos=9;
    if(pos>0){_csDigits[pos-1]='';_csRender();}
  } else if(e.key==='Enter'){
    e.preventDefault();_codeInputConfirm();
  } else if(e.key==='Escape'){
    e.preventDefault();closeM('m-code-input');
  }
}
function openCodeInput(fieldId, label, fmt) {
  _codeInputTarget = fieldId;
  _csFormat = fmt || '6_3';
  document.getElementById('code-input-lbl').textContent = label || 'Code';
  _csDigits = Array(9).fill('');
  const sep=document.getElementById('cs-sep');
  const hint=document.getElementById('cs-hint');
  if(sep) sep.style.display=_csFormat==='9'?'none':'';
  if(hint) hint.textContent=_csFormat==='9'?'Tapez les 9 chiffres du N° OF':'Tapez les 6 premiers puis les 3 derniers chiffres';
  _csRender();
  openM('m-code-input');
  setTimeout(()=>{const s=document.getElementById('code-slots');if(s)s.focus();},80);
}
function _codeInputConfirm() {
  const all=_csDigits.join('');
  if(_csFormat==='9'){
    if(!/^[0-9]{9}$/.test(all)){toast('Format requis : 9 chiffres (ex : 123456789)','err');return;}
    const field=document.getElementById('f-'+_codeInputTarget);
    if(field){field.value=all;scheduleAutoSave();field.blur();}
  } else {
    const d6=_csDigits.slice(0,6).join('');
    const d3=_csDigits.slice(6,9).join('');
    const v=d6+'_'+d3;
    if(!/^[0-9]{6}_[0-9]{3}$/.test(v)){toast('Format requis : 6 chiffres_3 chiffres (ex : 123456_789)','err');return;}
    const field=document.getElementById('f-'+_codeInputTarget);
    if(field){field.value=v;scheduleAutoSave();field.blur();}
  }
  closeM('m-code-input');
  _checkFormAutoConfirm();
}
function _checkFormAutoConfirm(){
  if(window.ST&&window.ST.prod_active) return;
  const ofVal=(document.getElementById('f-of_num')||{}).value||'';
  const cpVal=(document.getElementById('f-code_prod')||{}).value||'';
  if(/^[0-9]{9}$/.test(ofVal)&&/^[0-9]{6}_[0-9]{3}$/.test(cpVal)){
    setTimeout(doStartProd,500);
  }
}

// ── OF DETAIL REPORT ──
window._rptProdRows = [];
window._rptEvtRows = [];

// Shared OF/arrêt detail modal renderer
function _renderAndOpenOfDetail(r, ofEvts) {
  const isProd = r._rowType !== 'evt';
  const tc = r.trs>=90?'#16a34a':r.trs>=70?'#f59e0b':r.trs>=0?'#dc2626':'#94a3b8';
  const kitStr = (r.kit||'').toLowerCase();

  function _row(lbl,val,col){return val&&String(val).trim()&&String(val).trim()!=='0'?`<div style="display:flex;justify-content:space-between;align-items:baseline;padding:3px 0;border-bottom:1px solid #f1f5f9"><span style="font-size:calc(9px*var(--zf,1));font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.04em;flex-shrink:0;padding-right:8px">${lbl}</span><span style="font-size:calc(11px*var(--zf,1));font-weight:700;color:${col};text-align:right">${esc(String(val))}</span></div>`:'';}
  function _sec(lbl){return `<div style="font-size:calc(9px*var(--zf,1));font-weight:900;color:#475569;text-transform:uppercase;letter-spacing:.08em;padding:8px 0 4px;border-bottom:2px solid #e2e8f0;margin-bottom:4px">${lbl}</div>`;}

  if(!isProd){
    const commentHtmlE=r.comment?`<div style="background:#fffbeb;border-left:3px solid #fbbf24;padding:7px 10px;margin:8px 0;font-size:calc(11px*var(--zf,1));color:#92400e;border-radius:0 8px 8px 0">💬 ${esc(r.comment)}</div>`:'';
    document.getElementById('of-detail-content').innerHTML=`
      <div style="background:linear-gradient(135deg,#991b1b 0%,#dc2626 100%);color:#fff;padding:14px 20px;border-radius:18px 18px 0 0;display:flex;align-items:center;justify-content:space-between;flex-shrink:0">
        <div><div style="font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.12em;opacity:.6;margin-bottom:3px">Détail Arrêt</div><div style="font-size:calc(18px*var(--zf,1));font-weight:900">${esc(r.type||'—')}</div></div>
        <button onclick="closeM('m-of-detail')" style="background:rgba(255,255,255,.15);border:none;color:#fff;border-radius:10px;width:36px;height:36px;font-size:calc(17px*var(--zf,1));cursor:pointer;display:flex;align-items:center;justify-content:center">✕</button>
      </div>
      <div style="padding:18px 22px;overflow-y:auto;flex:1;background:#fff">
        ${[_row('Heure début',r.debut,'#374151'),_row('Heure fin',r.fin,'#374151'),_row('Durée',r.duree,'#059669'),_row('Date',r.date,'#374151'),_row('Poste',r.poste,'#374151'),_row('Pilote',r.pilote,'#374151'),_row('OF associé',r.of,'#1e3a8a')].join('')}
        ${commentHtmlE}
      </div>`;
    const box=document.getElementById('of-detail-box');
    if(box){box.style.width='min(480px,96vw)';box.style.height='auto';box.style.maxHeight='80vh';}
    openM('m-of-detail');
    return;
  }

  // ── Timeline bandeau plein-largeur ──
  const hm2ms=hm=>{if(!hm)return 0;const p=(hm+':0:0').split(':').map(Number);return((p[0]||0)*3600+(p[1]||0)*60)*1000;};
  const hm2s2=hm=>{if(!hm)return 0;const p=(hm+':0').split(':').map(Number);return p[0]*3600+p[1]*60;};
  let tlBandHtml='';
  if(r.debut&&r.fin){
    const dMs=hm2ms(r.debut),fMs=hm2ms(r.fin);
    const span=Math.max(1,fMs-dMs);
    const pct=ms=>Math.max(0,Math.min(100,Math.round((ms-dMs)/span*1000)/10));
    const segs=ofEvts.map(ev=>({s:hm2ms(ev.debut),e:hm2ms(ev.fin||ev.debut),c:ev.is_degrade?'#f59e0b':'#dc2626',l:ev.is_degrade?'Dégradé':(ev.type||'Arrêt')})).filter(sg=>sg.e>sg.s);
    const tlBar=segs.map(sg=>`<div title="${esc(sg.l)} ${Math.round((sg.e-sg.s)/60000)}min" style="position:absolute;top:0;bottom:0;left:${pct(sg.s)}%;width:${Math.max(.6,pct(sg.e)-pct(sg.s))}%;background:${sg.c};border-radius:3px;opacity:.9;box-shadow:0 2px 5px rgba(0,0,0,.22)"></div>`).join('');
    tlBandHtml=`<div style="padding:10px 22px 8px;background:#f8fafc;border-bottom:2px solid #e2e8f0;flex-shrink:0">
      <div style="font-size:calc(10px*var(--zf,1));font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:.07em;margin-bottom:7px">Timeline OF · ${r.debut} → ${r.fin}</div>
      <div style="position:relative;height:38px;background:linear-gradient(135deg,#dcfce7,#bbf7d0);border-radius:10px;overflow:hidden;box-shadow:inset 0 2px 5px rgba(0,0,0,.07),0 2px 8px rgba(0,0,0,.08)">${tlBar}</div>
      <div style="display:flex;justify-content:space-between;font-size:calc(9px*var(--zf,1));color:#94a3b8;margin-top:4px;margin-bottom:6px"><span>${r.debut}</span><span>${r.fin}</span></div>
      <div style="display:flex;gap:16px;flex-wrap:wrap">
        <div style="display:flex;align-items:center;gap:5px;font-size:calc(10px*var(--zf,1));color:#374151"><div style="width:12px;height:12px;border-radius:3px;background:#16a34a"></div>Production</div>
        <div style="display:flex;align-items:center;gap:5px;font-size:calc(10px*var(--zf,1));color:#374151"><div style="width:12px;height:12px;border-radius:3px;background:#dc2626"></div>Arrêt non prévu</div>
        <div style="display:flex;align-items:center;gap:5px;font-size:calc(10px*var(--zf,1));color:#374151"><div style="width:12px;height:12px;border-radius:3px;background:#60a5fa"></div>Arrêt prévu</div>
        <div style="display:flex;align-items:center;gap:5px;font-size:calc(10px*var(--zf,1));color:#374151"><div style="width:12px;height:12px;border-radius:3px;background:#f59e0b"></div>Mode dégradé</div>
      </div>
    </div>`;
  }

  // ── Donut chart (GROS) ──
  const dS=hm2s2(r.debut),fS=hm2s2(r.fin);
  const totalMin=Math.max(1,Math.round((fS-dS)/60));
  const {netMin,stopMin}=(window._rptNetProd?window._rptNetProd(r.debut,r.fin):{netMin:0,stopMin:0});
  const degMin=window._rptDegMin?window._rptDegMin(r.debut,r.fin):0;
  const planMin=Math.round((r.plan_stop_s||0)/60);
  const unplanMin=Math.max(0,stopMin-planMin);
  const prodMin=Math.max(0,netMin-degMin);
  const slices=[{v:prodMin,c:'#16a34a',l:'Production'},{v:planMin,c:'#60a5fa',l:'Arrêts prévus'},{v:unplanMin,c:'#dc2626',l:'Arrêts non prévus'},{v:degMin,c:'#f59e0b',l:'Mode dégradé'}].filter(s=>s.v>0);
  const tot=slices.reduce((a,s)=>a+s.v,0)||1;
  let sA=-Math.PI/2,dpaths='';
  const visSlices=slices.filter(sl=>sl.v>0&&(sl.v/tot)*2*Math.PI>=0.001);
  if(visSlices.length===1){
    // 100% — arc dégénéré : anneau plein de la couleur du seul segment
    dpaths=`<circle cx="105" cy="105" r="90" fill="${visSlices[0].c}" opacity=".93" filter="url(#ds3)"/><circle cx="105" cy="105" r="44" fill="white" filter="url(#ds3)"/>`;
  } else {
    slices.forEach(sl=>{
      const a=sl.v/tot*2*Math.PI,cx=105,cy=105,r2=90,ri=44;
      const x1=cx+r2*Math.cos(sA),y1=cy+r2*Math.sin(sA);
      const x2=cx+r2*Math.cos(sA+a),y2=cy+r2*Math.sin(sA+a);
      const xi1=cx+ri*Math.cos(sA),yi1=cy+ri*Math.sin(sA);
      const xi2=cx+ri*Math.cos(sA+a),yi2=cy+ri*Math.sin(sA+a);
      const lg=a>Math.PI?1:0;
      dpaths+=`<path d="M${xi1.toFixed(1)},${yi1.toFixed(1)} L${x1.toFixed(1)},${y1.toFixed(1)} A${r2},${r2} 0 ${lg},1 ${x2.toFixed(1)},${y2.toFixed(1)} L${xi2.toFixed(1)},${yi2.toFixed(1)} A${ri},${ri} 0 ${lg},0 ${xi1.toFixed(1)},${yi1.toFixed(1)}" fill="${sl.c}" opacity=".93" filter="url(#ds3)"/>`;
      sA+=a;
    });
  }
  const legend2=slices.map(sl=>`<div style="display:flex;align-items:center;gap:8px;font-size:calc(12px*var(--zf,1));padding:5px 0;border-bottom:1px solid #f1f5f9"><div style="width:14px;height:14px;border-radius:4px;background:${sl.c};flex-shrink:0;box-shadow:0 1px 4px rgba(0,0,0,.2)"></div><span style="color:#374151;flex:1">${sl.l}</span><span style="color:#94a3b8;font-size:calc(10px*var(--zf,1));margin-right:6px">${Math.round(sl.v/tot*100)}%</span><b style="color:${sl.c};white-space:nowrap">${sl.v} min</b></div>`).join('');
  const chartHtml2=r.debut&&r.fin?`
    <svg viewBox="0 0 210 210" width="210" height="210" style="display:block;margin:0 auto">
      <defs><filter id="ds3" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="3" stdDeviation="3" flood-opacity=".22"/></filter></defs>
      ${dpaths||'<circle cx="105" cy="105" r="90" fill="#e2e8f0"/><circle cx="105" cy="105" r="44" fill="white" filter="url(#ds3)"/>'}
    </svg>
    <div style="width:100%;padding:0 6px;margin-top:8px">${legend2}</div>
    <div style="font-size:calc(11px*var(--zf,1));color:#94a3b8;text-align:center;margin-top:10px;font-weight:600">${totalMin} min total OF</div>`
    :'<div style="color:#94a3b8;font-size:calc(11px*var(--zf,1));padding:30px 0;text-align:center">Pas de données</div>';

  // ── Événements enrichis ──
  function _isPlanned(type){const t=(type||'').toLowerCase();return t.includes('pause')||t.includes('nettoyage')||t.includes('nett')||t.includes('réunion')||t.includes('reunion')||t.includes('meeting');}
  let totalPlannedActMin=0;
  ofEvts.forEach(ev=>{if(!ev.is_degrade&&_isPlanned(ev.type)){const p=(ev.duree||'0:0:0').split(':').map(Number);totalPlannedActMin+=(p[0]||0)*60+(p[1]||0)+(p[2]||0)/60;}});
  const budgetExcess=Math.round(Math.max(0,totalPlannedActMin-planMin));
  const budgetWarnHtml2=budgetExcess>0&&planMin>0?`<div style="background:#fef3c7;border:1.5px solid #fde68a;border-radius:8px;padding:7px 10px;margin-bottom:8px;font-size:calc(10px*var(--zf,1));display:flex;align-items:center;gap:6px">⚠️ <span>Dépassement budget arrêts prévus : <b style="color:#b45309">+${budgetExcess} min</b> <span style="color:#94a3b8">(budget ${planMin} min)</span></span></div>`:'';
  const evtsHtml2=ofEvts.length?ofEvts.map(ev=>{
    const planned=ev.is_degrade?null:_isPlanned(ev.type);
    const evDurParts=(ev.duree||'0:0:0').split(':').map(Number);
    const evMin=Math.round((evDurParts[0]||0)*60+(evDurParts[1]||0)+(evDurParts[2]||0)/60);
    const bgColor=ev.is_degrade?'#fffbeb':planned?'#eff6ff':'#fff7f7';
    const borderColor=ev.is_degrade?'#fde68a':planned?'#bfdbfe':'#fecaca';
    const typeColor=ev.is_degrade?'#b45309':planned?'#1d4ed8':'#dc2626';
    const badge=ev.is_degrade?`<span style="background:#fef3c7;color:#b45309;font-size:calc(9px*var(--zf,1));font-weight:800;padding:2px 7px;border-radius:10px;border:1px solid #fde68a">DÉGRADÉ</span>`:planned?`<span style="background:#dbeafe;color:#1d4ed8;font-size:calc(9px*var(--zf,1));font-weight:800;padding:2px 7px;border-radius:10px;border:1px solid #bfdbfe">✓ PRÉVU</span>`:`<span style="background:#fee2e2;color:#dc2626;font-size:calc(9px*var(--zf,1));font-weight:800;padding:2px 7px;border-radius:10px;border:1px solid #fecaca">✗ NON PRÉVU</span>`;
    const commentLine=ev.comment?`<div style="display:flex;align-items:flex-start;gap:5px;font-size:calc(10px*var(--zf,1));color:#92400e;background:#fffbeb;border-radius:5px;padding:4px 7px;margin-top:5px"><span style="flex-shrink:0">💬</span><span>${esc(ev.comment)}</span></div>`:'';
    return `<div style="background:${bgColor};border:1.5px solid ${borderColor};border-radius:9px;padding:9px 11px;margin-bottom:7px"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px"><div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap"><span style="font-size:calc(12px*var(--zf,1));font-weight:800;color:${typeColor}">${esc(ev.type||'Arrêt')}</span>${badge}</div><span style="font-size:calc(14px*var(--zf,1));font-weight:900;color:${typeColor};white-space:nowrap;margin-left:8px">${evMin} min</span></div><div style="font-size:calc(10px*var(--zf,1));color:#64748b;margin-bottom:2px">${esc(ev.debut||'')} → ${esc(ev.fin||'')} · durée ${esc(ev.duree||'—')}</div>${commentLine}</div>`;
  }).join(''):`<div style="color:#94a3b8;font-size:calc(11px*var(--zf,1));text-align:center;padding:16px 0;border:1.5px dashed #e2e8f0;border-radius:8px">Aucun arrêt ni événement</div>`;

  // ── Colonnes infos ──
  const commentHtml3=r.comment?`<div style="background:#fffbeb;border-left:3px solid #fbbf24;padding:7px 10px;margin-top:10px;font-size:calc(11px*var(--zf,1));color:#92400e;border-radius:0 8px 8px 0;box-shadow:0 2px 5px rgba(251,191,36,.15)">💬 ${esc(r.comment)}</div>`:'';
  const col1Html=[_sec('Identité'),_row('OF',r.of,'#1e3a8a'),_row('Date',r.date,'#374151'),_row('Poste',r.poste,'#374151'),_row('Pilote',r.pilote,'#374151'),_row('Co-Pilote',r.copilote,'#374151'),_row('Nb Personnes',r.nb_pers,'#374151'),_sec('Produit'),_row('Taille',r.taille,'#374151'),_row('Type produit',r.type_prod,'#374151'),_row('Code produit',r.code_prod,'#374151'),_row('Fibre',r.fibre,'#6366f1'),_row('Poids garnissage (g)',r.poids,'#374151'),_row('OF Taie',r.of_taie,'#374151'),_row('Réf Taie',r.ref_taie,'#374151'),_row('Traca',r.traca?(r.traca.split(';').filter(t=>t.trim()).join(' · ')):'' ,'#374151'),_row('Lots de 2',kitStr==='oui'?'✓ Oui':'Non',kitStr==='oui'?'#16a34a':'#94a3b8')].join('');
  const col2Html=[_sec('Production'),_row('Heure début',r.debut,'#374151'),_row('Heure fin',r.fin,'#374151'),_row('Durée',r.duree,'#059669'),_row('Qté fabriquée',r.qte_fab,'#1e3a8a'),_row('Qté emballée',r.qte_emb,'#374151'),_row('Équivalence',r.equiv,'#0891b2'),_row('Cadence/h',r.cadence_h,'#374151'),_row('Cadence/h/pers',r.cadence_h_pers,'#374151'),_row('TRS %',r.trs>=0?r.trs.toFixed(1)+'%':'—',tc),_row('Objectif pièces',r.objectif!=null&&r.objectif>=0?String(r.objectif):'','#0369a1'),_row('Prévu/Hors TRS',r.prevu_hors_trs,'#374151'),commentHtml3,`<div style="margin-top:10px">${_sec('Qualité')}${[_row('Qté init Taie',r.qte_init_taie,'#374151'),_row('Nb Taie 2nd choix',r.nb_taie2,'#f59e0b'),_row('Nb défauts couture',r.nb_def_cout,'#dc2626'),_row('Mq Taie',r.mq_taie,'#dc2626'),_row('Mq Housse/Encart',r.mq_housse,'#dc2626'),_row('Nb PP cousue',r.nb_pp,'#374151')].join('')}</div>`,`<div style="margin-top:4px">${_sec('Manquants')}${[_row('Manquant MP',r.duree_mq_mp,'#dc2626'),_row('Manquant Personnel/Réunion',r.manquant_pers,'#374151')].join('')}</div>`].join('');
  const col3Html=`${_sec('Événements ('+ofEvts.length+')')}${budgetWarnHtml2}${evtsHtml2}`;

  const trsBlock2=r.trs>=0?`<div style="text-align:right;background:#fff;border-radius:10px;padding:8px 14px;box-shadow:0 4px 12px rgba(0,0,0,.25)"><div style="font-size:calc(28px*var(--zf,1));font-weight:900;color:${tc};line-height:1">${r.trs.toFixed(1)}%</div><div style="font-size:calc(9px*var(--zf,1));color:#94a3b8;text-transform:uppercase;letter-spacing:.1em">TRS</div></div>`:'';

  document.getElementById('of-detail-content').innerHTML=`
    <!-- Header bleu -->
    <div style="background:linear-gradient(135deg,#1e3a8a 0%,#1e40af 50%,#2563eb 100%);color:#fff;padding:14px 22px;border-radius:18px 18px 0 0;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;box-shadow:0 4px 16px rgba(30,58,138,.4)">
      <div>
        <div style="font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.12em;opacity:.6;margin-bottom:3px">Détail OF</div>
        <div style="font-size:calc(22px*var(--zf,1));font-weight:900;font-family:monospace;letter-spacing:.05em;text-shadow:0 2px 8px rgba(0,0,0,.25)">${esc(r.of||'—')}</div>
        ${r.taille||r.type_prod?`<div style="font-size:calc(12px*var(--zf,1));opacity:.75;margin-top:3px">${esc(r.taille||'')} ${esc(r.type_prod||'')}</div>`:''}
      </div>
      <div style="display:flex;align-items:center;gap:12px">${trsBlock2}<button onclick="closeM('m-of-detail')" style="background:rgba(255,255,255,.15);border:none;color:#fff;border-radius:10px;width:38px;height:38px;font-size:calc(18px*var(--zf,1));cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,.2);transition:background .15s" onmouseover="this.style.background='rgba(255,255,255,.25)'" onmouseout="this.style.background='rgba(255,255,255,.15)'">✕</button></div>
    </div>
    <!-- Timeline bandeau plein-largeur -->
    ${tlBandHtml}
    <!-- Corps 4 colonnes -->
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr minmax(250px,310px);flex:1;min-height:0;overflow:hidden">
      <div style="padding:12px 12px;border-right:1px solid #e2e8f0;overflow-y:auto;background:#fff">${col1Html}</div>
      <div style="padding:12px 12px;border-right:1px solid #e2e8f0;overflow-y:auto;background:#fff">${col2Html}</div>
      <div style="padding:12px 12px;border-right:1px solid #e2e8f0;overflow-y:auto;background:#fff">${col3Html}</div>
      <div style="padding:12px 10px;overflow-y:auto;background:linear-gradient(180deg,#f8fafc 0%,#fff 100%);display:flex;flex-direction:column;align-items:center">
        <div style="font-size:calc(11px*var(--zf,1));font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:.07em;margin-bottom:12px;text-align:center">Répartition temps</div>
        ${chartHtml2}
      </div>
    </div>`;
  const box=document.getElementById('of-detail-box');
  if(box){box.style.width='min(1400px,99vw)';box.style.height='92vh';box.style.maxHeight='92vh';}
  openM('m-of-detail');
}

function showRptEvtDetail(ri) {
  const r = window._rptEvtRows[ri];
  if(!r) return;
  _renderAndOpenOfDetail({...r, _rowType:'evt'}, []);
}

function showOfDetail(ri) {
  const r = window._rptProdRows[ri];
  const evtRows = window._rptEvtRows || [];
  if(!r) return;
  function _hm2s(hm) { if(!hm) return 0; const p=(hm+':0:0').split(':').map(Number); return(p[0]||0)*3600+(p[1]||0)*60+(p[2]||0); }
  const debS = _hm2s(r.debut), finS = _hm2s(r.fin)||86400;
  const ofEvts = evtRows.filter(e => { const t = _hm2s(e.debut); return t >= debS && t <= finS; });
  _renderAndOpenOfDetail(r, ofEvts);
}

function showMainRowDetail(key) {
  const r = window._rowMap && window._rowMap[String(key)];
  if(!r) return;
  const isProd = r._rowType === 'prod';
  let ofEvts = [];
  if(isProd) {
    const _hm2s = hm=>{if(!hm)return 0;const[h,m]=(hm+':00').split(':').map(Number);return(h||0)*3600+(m||0)*60;};
    const debS=_hm2s(r.debut), finS=_hm2s(r.fin)||86400;
    ofEvts=(window._mainEvtsAll||[]).filter(e=>{const t=_hm2s(e.debut);return t>=debS&&t<=finS;});
  }
  _renderAndOpenOfDetail(r, ofEvts);
}


// ── ARRÊTS MANQUANTS ──
let _missingDeclChecked = false;
let _ecartChecked = false;

function _hasDeclaredType(keywords){
  // Cherche dans les déclarations du poste actuel (accueil)
  const rows=window._lastMainRows||[];
  return rows.some(r=>{
    const t=String(r.type||r._type||r[0]||'').toLowerCase();
    return keywords.some(k=>t.includes(k));
  });
}

async function checkMissingDecls(){
  // Refresh des déclarations avant check
  await loadMainDecl();
  const ap=_cfgArretsPrevus||{};
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const model=_cfgModels&&_cfgModels.find(m=>m.nom===(ST.poste||''));
  const jour=model&&model.jours&&model.jours[dk];
  const missing=[];
  if((ap.pause_min||0)>0 && !_hasDeclaredType(['pause'])){
    missing.push({label:'Pause',budget:ap.pause_min||0,keywords:['pause'],id:'md-pause'});
  }
  if((ap.clean_short_min||0)>0 && !_hasDeclaredType(['nettoyage','nett'])){
    missing.push({label:'Nettoyage court',budget:ap.clean_short_min||0,keywords:['nettoyage','nett'],id:'md-nett'});
  }
  if((ap.meeting_tol_min||0)>0 && !_hasDeclaredType(['réunion','reunion','meeting'])){
    missing.push({label:'Réunion',budget:ap.meeting_tol_min||0,keywords:['réunion','reunion','meeting'],id:'md-meet'});
  }
  if(!missing.length) return false;
  // Construire les lignes du modal
  const lastFin=_lastProdDeclTime;
  const container=document.getElementById('md-rows');
  container.innerHTML='';
  missing.forEach(item=>{
    // Heure de début suggérée = fin dernière prod ou fin modèle - budget
    let sugStart='',sugEnd='';
    if(jour&&jour.fin){
      const[fh,fm]=jour.fin.split(':').map(Number);
      const budMin=item.budget;
      const endM=fh*60+fm;
      const startM=endM-budMin;
      const sh=Math.floor(((startM%1440)+1440)%1440/60),sm=((startM%1440)+1440)%1440%60;
      sugStart=String(sh).padStart(2,'0')+':'+String(sm).padStart(2,'0');
      sugEnd=jour.fin;
    } else if(lastFin){
      sugStart=String(lastFin.getHours()).padStart(2,'0')+':'+String(lastFin.getMinutes()).padStart(2,'0');
      const e=new Date(lastFin.getTime()+item.budget*60000);
      sugEnd=String(e.getHours()).padStart(2,'0')+':'+String(e.getMinutes()).padStart(2,'0');
    }
    const div=document.createElement('div');
    div.style.cssText='background:#fffbeb;border:1.5px solid #fde68a;border-radius:8px;padding:10px 12px';
    div.innerHTML=`<div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:#92400e;margin-bottom:6px">${esc(item.label)} <span style="font-weight:400;color:#a16207">(prévu ${item.budget} min)</span></div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <label style="font-size:calc(11px*var(--zf,1));font-weight:600">De <input type="time" id="${item.id}-debut" value="${sugStart}" style="padding:4px 6px;border:1.5px solid #fde68a;border-radius:5px;font-size:calc(12px*var(--zf,1));font-weight:700;margin-left:4px"></label>
        <label style="font-size:calc(11px*var(--zf,1));font-weight:600">à <input type="time" id="${item.id}-fin" value="${sugEnd}" style="padding:4px 6px;border:1.5px solid #fde68a;border-radius:5px;font-size:calc(12px*var(--zf,1));font-weight:700;margin-left:4px"></label>
        <button onclick="addMissingDecl('${item.id}','${esc(item.label)}')" style="background:#f59e0b;color:#fff;border:none;border-radius:6px;padding:5px 12px;font-size:calc(11px*var(--zf,1));font-weight:700;cursor:pointer">+ Ajouter</button>
        <span id="${item.id}-ok" style="display:none;color:#16a34a;font-weight:700;font-size:calc(12px*var(--zf,1))">✓ Ajouté</span>
      </div>`;
    container.appendChild(div);
  });
  openM('m-missing-decl');
  return true;
}

async function addMissingDecl(rowId, label){
  const debut=document.getElementById(rowId+'-debut')?.value;
  const fin=document.getElementById(rowId+'-fin')?.value;
  if(!debut||!fin){toast('Remplissez les heures','warn');return;}
  const r=await fetch('/api/add_stop_decl',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({type:label,debut_hms:debut,fin_hms:fin,comment:'Déclaré rétroactivement'})});
  const d=r?await r.json():{};
  if(d.ok){
    const ok=document.getElementById(rowId+'-ok');if(ok) ok.style.display='';
    // Griser la ligne
    const btn=document.querySelector(`button[onclick="addMissingDecl('${rowId}','${label.replace(/'/g,"\\'")}')"]`);
    if(btn){btn.disabled=true;btn.style.opacity='.4';}
    toast(label+' ajouté','ok');
  } else {
    toast('Erreur : '+(d.error||'?'),'err');
  }
}

async function skipMissingDecl(){
  closeM('m-missing-decl');
  _missingDeclChecked=true;
  await doFinPoste();
}

// ── FIN DE POSTE ──
async function doFinPoste(){
  if(ST.prod_active){toast('Terminer la production en cours avant de finir le poste','err');return;}
  // Arrêt automatique du mode dégradé s'il est actif
  if(window._degradeActive){
    const rd=await fetch('/api/stop_degrade',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
    if(rd&&rd.ok){window._degradeActive=false;toast('Mode dégradé arrêté automatiquement','ok');}
  }
  const fpd=await apiFetch('/api/fin_poste_data');
  if(fpd){
    window._ecartFpData=fpd;
    const hasIssue=(fpd.gap_intervals&&fpd.gap_intervals.length>0)||((fpd.overflow_s||0)>60)||((fpd.ecart_s||0)>60);
    if(hasIssue){_showEcartModal(fpd);}else{goTab('finposte');}
  } else goTab('finposte');
}

async function _doGoFinPoste(){
  if(!_ecartChecked){
    const fpd=await apiFetch('/api/fin_poste_data');
    if(fpd&&(fpd.ecart_s||0)>60){
      window._ecartFpData=fpd;
      _showEcartModal(fpd);
      return;
    }
  }
  _ecartChecked=false;
  goTab('finposte');
}

function _showEcartModal(fpd){
  const ecart_min=Math.round((fpd.ecart_s||0)/60);
  const model_min=Math.round((fpd.model_dur_s||0)/60);
  const prod_min=Math.round((fpd.tot_s||0)/60);
  const stop_min=Math.max(0,model_min-prod_min-ecart_min);
  const modelDebut=fpd.model_debut||'';
  const modelFin=fpd.model_fin||'';
  const overflow_min=Math.round(fpd.overflow_min||0);
  const gaps=fpd.gap_intervals||[];
  const hasGaps=gaps.length>0;
  // Réinitialiser le panneau plage horaire au style normal avant d'appliquer la surcharge
  const _pP=document.getElementById('ecart-plage-panel');
  if(_pP){
    _pP.style.cssText='margin-bottom:14px';
    const _pD=_pP.querySelectorAll('div');
    if(_pD[0]) _pD[0].style.color='#0369a1';
    if(_pD[1]){_pD[1].style.background='#f0f9ff';_pD[1].style.borderColor='#bae6fd';}
  }
  // Guide header — basé sur les gaps réels, pas sur ecart_min
  const guidEl=document.getElementById('ecart-guide');
  if(guidEl){
    if(overflow_min>0){
      guidEl.style.cssText='font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;background:#fef2f2;border:1px solid #fca5a5;line-height:1.5';
      guidEl.innerHTML='<span style="color:#dc2626;font-weight:800;font-size:calc(13px*var(--zf,1))">⚠ Dépassement de plage : +'+overflow_min+' min au-delà de '+esc(modelFin)+'</span><br>'+
        '<span style="color:#7f1d1d">Un ou plusieurs OFs se terminent après la fin de la plage modèle. Modifiez la plage ci-dessous.</span>';
      // Mettre le panneau de modification de plage en rouge pour signaler l'urgence
      const plagePanel=document.getElementById('ecart-plage-panel');
      if(plagePanel){
        plagePanel.style.cssText='margin-bottom:14px;border-radius:10px;border:2px solid #dc2626;padding:10px;background:#fef2f2';
        const divs=plagePanel.querySelectorAll('div');
        if(divs[0]) divs[0].style.color='#dc2626'; // label titre
        if(divs[1]){ divs[1].style.background='#fef2f2'; divs[1].style.borderColor='#fca5a5'; }
      }
    } else if(!hasGaps){
      guidEl.style.cssText='font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;background:#f0fdf4;border:1px solid #bbf7d0;line-height:1.5';
      guidEl.innerHTML='<span style="color:#16a34a;font-weight:800;font-size:calc(13px*var(--zf,1))">✓ Toute la plage '+esc(modelDebut)+' → '+esc(modelFin)+' est couverte !</span>';
    } else {
      const gapTotal=gaps.reduce((a,g)=>a+(g.duree_min||0),0);
      guidEl.style.cssText='font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;background:#fff7ed;border:1px solid #fed7aa;line-height:1.5';
      guidEl.innerHTML='Objectif : couvrir <b>'+esc(modelDebut)+' → '+esc(modelFin)+'</b> ('+model_min+' min).<br>'+
        '<span style="color:#dc2626;font-weight:700">'+gapTotal+' min de plages non couvertes ('+gaps.length+' écart'+(gaps.length>1?'s':'')+').</span> '+
        '<span style="color:var(--gray)">Déclarez les arrêts manquants ou corrigez les horaires des OFs.</span>';
    }
  }
  // Stats
  document.getElementById('ecart-info').innerHTML=
    '<div style="display:grid;grid-template-columns:auto 1fr auto 1fr;gap:3px 16px;font-size:calc(12px*var(--zf,1))">'+
    '<span style="color:var(--gray)">Durée modèle :</span><b>'+model_min+' min</b>'+
    '<span style="color:var(--gray)">Prod déclarée :</span><b>'+prod_min+' min</b>'+
    '<span style="color:var(--gray)">Arrêts déclarés :</span><b>'+stop_min+' min</b>'+
    '<span style="color:#dc2626;font-weight:700">Écart :</span><b style="color:#dc2626;font-weight:900">'+ecart_min+' min</b>'+
    '</div>';
  // Prefill plage horaire du poste
  const pdebut=document.getElementById('ecart-plage-debut');
  const pfin=document.getElementById('ecart-plage-fin');
  if(pdebut)pdebut.value=fpd.model_debut||'';
  if(pfin)pfin.value=fpd.model_fin||'';
  // Gap intervals
  const gapsEl=document.getElementById('ecart-gaps');
  if(gapsEl){
    if(!hasGaps){
      // Ne pas afficher "Aucune plage non couverte" quand il y a un dépassement
      gapsEl.innerHTML=overflow_min>0?''
        :'<div style="color:#16a34a;font-size:calc(12px*var(--zf,1));font-weight:700;padding:4px 0">✓ Aucune plage non couverte</div>';
    } else {
      gapsEl.innerHTML='';
      gaps.forEach((g,gi)=>{
        const div=document.createElement('div');
        div.style.cssText='background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:8px 10px;margin-bottom:6px';
        div.innerHTML=
          '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px">'+
            '<span style="font-size:calc(12px*var(--zf,1));font-weight:700;color:#dc2626">⚠ '+esc(g.debut)+' → '+esc(g.fin)+
              ' <span style="font-weight:400;color:#9f1239">('+g.duree_min+' min)</span></span>'+
            '<button class="btn btn-ghost" style="font-size:calc(11px*var(--zf,1));padding:3px 10px;color:#dc2626;border-color:#fca5a5" onclick="ecartToggleGapForm('+gi+')">+ Déclarer un arrêt</button>'+
          '</div>'+
          '<div id="ecart-gap-form-'+gi+'" style="display:none;margin-top:8px;border-top:1px solid #fca5a5;padding-top:8px">'+
            '<div style="display:flex;gap:6px;align-items:flex-end;flex-wrap:wrap">'+
              '<div><label style="font-size:calc(10px*var(--zf,1));color:var(--gray);display:block;margin-bottom:2px">Début</label>'+
                '<input type="time" id="ecart-gap-debut-'+gi+'" value="'+esc(g.debut)+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px"></div>'+
              '<div><label style="font-size:calc(10px*var(--zf,1));color:var(--gray);display:block;margin-bottom:2px">Fin</label>'+
                '<input type="time" id="ecart-gap-fin-'+gi+'" value="'+esc(g.fin)+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px"></div>'+
              '<div style="flex:1;min-width:160px"><label style="font-size:calc(10px*var(--zf,1));color:var(--gray);display:block;margin-bottom:2px">Type d\'arrêt</label>'+
                _buildEcartStopSelect(gi)+'</div>'+
              '<button class="btn btn-prim" style="font-size:calc(11px*var(--zf,1));padding:5px 12px" onclick="saveEcartGapStop('+gi+')">✓ Ajouter</button>'+
            '</div>'+
          '</div>';
        gapsEl.appendChild(div);
      });
    }
  }
  // OFs list
  const list=document.getElementById('ecart-of-list');
  list.innerHTML='';
  (fpd.of_list||[]).forEach(of=>{
    const div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:8px;padding:8px;background:#f8fafc;border-radius:6px;border:1px solid var(--border);flex-wrap:wrap';
    div.innerHTML='<span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:700;min-width:80px">'+(of.of||'OF')+
      ' <span style="font-weight:400;color:var(--gray);font-size:calc(11px*var(--zf,1))">'+esc(of.taille||'')+'</span> '+
      '<span style="font-size:calc(10px*var(--zf,1));color:#64748b">'+esc(of.debut||'')+'→'+esc(of.fin||'')+'</span></span>'+
      '<div style="display:flex;align-items:center;gap:4px">'+
      '<input type="time" class="ecart-debut" value="'+(of.debut||'')+'" data-oldebut="'+(of.debut||'')+'" data-ofnum="'+(of.of||'')+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px">'+
      '<span style="color:var(--gray)">→</span>'+
      '<input type="time" class="ecart-fin" value="'+(of.fin||'')+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px">'+
      '<button class="btn btn-prim" style="font-size:calc(11px*var(--zf,1));padding:3px 10px" onclick="saveEcartOf(this)">✓</button>'+
      '</div>';
    list.appendChild(div);
  });
  // Griser "Valider et terminer" tant que tout n'est pas résolu
  const valBtn=document.getElementById('ecart-valider-btn');
  if(valBtn){
    const locked=hasGaps||overflow_min>0;
    valBtn.disabled=locked;
    valBtn.style.opacity=locked?'0.4':'1';
    valBtn.style.cursor=locked?'not-allowed':'pointer';
    valBtn.title=locked?(hasGaps?'Justifiez toutes les plages non couvertes avant de terminer':'Résolvez le dépassement de plage avant de terminer'):'';
  }
  openM('m-ecart-poste');
}

function ecartToggleGapForm(gi){
  const f=document.getElementById('ecart-gap-form-'+gi);
  if(f) f.style.display=f.style.display==='none'?'block':'none';
}

async function saveEcartGapStop(gi){
  const debut=(document.getElementById('ecart-gap-debut-'+gi)||{}).value||'';
  const fin=(document.getElementById('ecart-gap-fin-'+gi)||{}).value||'';
  const type=((document.getElementById('ecart-gap-type-'+gi)||{}).value||'').trim();
  if(!debut||!fin||!type){toast('Renseigner début, fin et type d\'arrêt','err');return;}
  const r=await fetch('/api/add_stop_decl',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({type,debut_hms:debut,fin_hms:fin,comment:'Déclaré depuis réconciliation fin de poste'})});
  const d=r?await r.json():{};
  if(d.ok){
    toast(type+' ajouté','ok');
    const fpd=await apiFetch('/api/fin_poste_data');
    if(fpd){window._ecartFpData=fpd;_showEcartModal(fpd);}
  } else {
    toast('Erreur : '+(d.error||'?'),'err');
  }
}

function _buildEcartStopSelect(gi){
  // Rassemble tous les types d'arrêt depuis paramètres
  const pauses=new Set(['Pause']);
  const nettoyage=new Set(['Nettoyage court','Nettoyage long','Nettoyage très long']);
  const pannes=new Set();
  const organisation=new Set();
  (_evtsList||[]).forEach(e=>{
    const lbl=e.label||'';if(!lbl)return;
    if(e.cat==='pb'||e.cat==='ratt')pannes.add(lbl);
    else if(e.cat==='nettoyage')nettoyage.add(lbl);
    else if(e.cat==='organisation')organisation.add(lbl);
    else if(e.cat==='autre')organisation.add(lbl);
  });
  const mkGrp=(name,items)=>{
    const arr=[...items];if(!arr.length)return'';
    return`<optgroup label="${esc(name)}">`+arr.map(l=>`<option value="${esc(l)}">${esc(l)}</option>`).join('')+'</optgroup>';
  };
  const html='<option value="">-- Choisir un type d\'arrêt --</option>'+
    mkGrp('⏸ Pauses',pauses)+
    mkGrp('🧹 Nettoyage',nettoyage)+
    mkGrp('🔴 Pannes / Rattrapages',pannes)+
    mkGrp('🔵 Organisation',organisation)+
    '<optgroup label="⚫ Autre"><option value="Autre (à justifier)">Autre (à justifier)</option></optgroup>';
  return`<select id="ecart-gap-type-${gi}" onchange="ecartCheckAutreType(this,${gi})" style="width:100%;padding:4px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));background:#fff">${html}</select>`;
}

let _ecartGapAutreIdx=-1;
function ecartCheckAutreType(sel,gi){
  if(sel.value==='Autre (à justifier)'){
    _ecartGapAutreIdx=gi;
    const inp=document.getElementById('ecart-autre-label');if(inp)inp.value='';
    sel.value='';
    openM('m-ecart-autre');
  }
}

function confirmEcartAutre(){
  const inp=document.getElementById('ecart-autre-label');
  const lbl=(inp&&inp.value||'').trim();
  if(!lbl){toast('Saisir un libellé','err');return;}
  const gi=_ecartGapAutreIdx;
  const sel=document.getElementById('ecart-gap-type-'+gi);
  if(sel){
    let opt=[...sel.options].find(o=>o.value===lbl);
    if(!opt){opt=document.createElement('option');opt.value=lbl;opt.textContent=lbl;sel.appendChild(opt);}
    sel.value=lbl;
  }
  closeM('m-ecart-autre');
}

async function saveEcartPlageHoraire(){
  const debut=(document.getElementById('ecart-plage-debut')||{}).value||'';
  const fin=(document.getElementById('ecart-plage-fin')||{}).value||'';
  if(!debut||!fin){toast('Renseigner début et fin','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const poste=ST.poste||'';
  if(!poste){toast('Poste non défini','err');return;}
  const mi=_cfgModels.findIndex(m=>m.nom===poste);
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=debut;
    _cfgModels[mi].jours[dk].fin=fin;
  }
  await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:dk,debut,fin})});
  // Mettre à jour shift_debut_dt / shift_fin_dt → écrit aussi dans Excel POSTES P/Q
  const _today=new Date();
  const [_dh,_dm]=(debut||'00:00').split(':').map(Number);
  const [_fh,_fm]=(fin||'00:00').split(':').map(Number);
  const _debDt=new Date(_today.getFullYear(),_today.getMonth(),_today.getDate(),_dh,_dm,0);
  let _finDt=new Date(_today.getFullYear(),_today.getMonth(),_today.getDate(),_fh,_fm,0);
  if(_finDt<=_debDt) _finDt=new Date(_finDt.getTime()+86400000);
  await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:_debDt.toISOString(),fin_iso:_finDt.toISOString()})});
  toast('Plage mise à jour','ok');
  const fpd=await apiFetch('/api/fin_poste_data');
  if(fpd){window._ecartFpData=fpd;_showEcartModal(fpd);}
}

async function skipEcartPoste(){
  closeM('m-ecart-poste');
  _ecartChecked=true;
  await _doGoFinPoste();
}

function ecartOpenOfPanel(){
  // OF panel is always visible now - kept for backward compat
}

async function saveEcartOf(btn){
  const row=btn.closest('div');
  const debutIn=row.querySelector('.ecart-debut');
  const finIn=row.querySelector('.ecart-fin');
  const ofNum=debutIn?debutIn.dataset.ofnum:'';
  const oldDebut=debutIn?debutIn.dataset.oldebut:'';
  const newDebut=debutIn?debutIn.value:'';
  const newFin=finIn?finIn.value:'';
  if(!newDebut||!newFin){toast('Renseigner début et fin','err');return;}
  const r=await apiFetch('/api/update_of_time',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({of_num:ofNum,old_debut:oldDebut,new_debut:newDebut,new_fin:newFin})});
  if(r&&r.ok){
    toast('OF mis à jour','ok');
    if(debutIn) debutIn.dataset.oldebut=newDebut;
    // Recalculer l'écart
    const fpd=await apiFetch('/api/fin_poste_data');
    if(fpd){window._ecartFpData=fpd;_showEcartModal(fpd);}
  } else {
    toast('Erreur mise à jour','err');
  }
}

async function loadFPData(){
  const d=await apiFetch('/api/fin_poste_data');
  if(!d) return;
  document.getElementById('fp-trs').textContent=fmtTRS(d.trs_shift!==undefined?d.trs_shift:d.trs);
  document.getElementById('fp-trs-of').textContent=fmtTRS(d.trs);
  document.getElementById('fp-eq').textContent=Math.round(d.tot_equiv||0);
  document.getElementById('fp-nof').textContent=d.nb_of||0;
  // fp-prod-t sera mis à jour après le calcul de stopTotal ci-dessous
  const fpOuvEl=document.getElementById('fp-ouverture');
  if(fpOuvEl) fpOuvEl.textContent=(d.model_debut&&d.model_fin)?(d.model_debut+'→'+d.model_fin):(Math.round((d.model_dur_s||0)/60)+' min');
  const fpDedEl=document.getElementById('fp-ded');
  if(fpDedEl) fpDedEl.textContent=Math.round((d.planned_ded_s||0)/60)+' min';
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
  const _hmsS=hms=>{if(!hms)return 0;const p=(hms+':0:0').split(':').map(Number);return(p[0]||0)*3600+(p[1]||0)*60+(p[2]||0);};
  const stopTotal=stops.reduce((a,e)=>a+Math.max(0,_hmsS(e.fin)-_hmsS(e.debut)),0);
  document.getElementById('fp-stop-t').textContent=fmtDurMS(stopTotal);
  // Net prod = sum of each OF's time minus overlapping stops within that period
  const _ofPs=(d.of_list||[]).map(p=>({s:_hmsS(p.debut),e:_hmsS(p.fin)})).filter(p=>p.e>p.s);
  const _stPs=stops.map(e=>({s:_hmsS(e.debut),e:_hmsS(e.fin)})).filter(e=>e.e>e.s);
  let _netProdS=0;
  _ofPs.forEach(pp=>{
    const ov=_stPs.map(sv=>({s:Math.max(sv.s,pp.s),e:Math.min(sv.e,pp.e)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    _netProdS+=Math.max(0,pp.e-pp.s-mg.reduce((a,o)=>a+(o.e-o.s),0));
  });
  document.getElementById('fp-prod-t').textContent=Math.round(_netProdS/60)+' min';

  // New metrics: Cadence/h, Nb pièces, Chg. fibre
  const ofList=d.of_list||[];
  const totQteFabFP=ofList.reduce((s,r)=>s+parseFloat(r.qte_fab||0),0);
  const cadenceHFP=(d.cadence_h!=null&&d.cadence_h>0)?Math.round(d.cadence_h):0;
  const sortedProdFFP=ofList.filter(r=>r.fibre).sort((a,b)=>(a.debut||'').localeCompare(b.debut||''));
  let nbChangFibreFP=0;for(let i=1;i<sortedProdFFP.length;i++){if(sortedProdFFP[i].fibre!==sortedProdFFP[i-1].fibre)nbChangFibreFP++;}
  document.getElementById('fp-cadence').textContent=cadenceHFP;
  document.getElementById('fp-pieces').textContent=Math.round(totQteFabFP);
  document.getElementById('fp-chg-fibre').textContent=nbChangFibreFP;

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
        <td>${esc(p.debut||'')}</td>
        <td>${esc(p.fin||'')}</td>
        <td>${esc(p.taille||'')}</td>
        <td>${esc(String(p.qte_fab||0))}</td>
        <td>${esc(String(p.equiv||''))}</td>
        <td>${esc(p.duree||'')}</td>
        <td class="${(p.trs||0)>=90?'tg':(p.trs||0)>=75?'tm':'tb'}">${fmtTRS(p.trs||0)}</td>
      </tr>`).join('')||'<tr><td colspan="8" style="color:var(--gray)">Aucune production</td></tr>';
  }

  // Stops list
  const fsl=document.getElementById('fp-stops-list');
  if(fsl){
    const stopsByType={};
    gEvts.forEach(e=>{if(!e.type)return;if(!stopsByType[e.type])stopsByType[e.type]=0;
    try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
    stopsByType[e.type]+=Math.max(0,p(e.fin||'00:00:00')-p(e.debut||'00:00:00'));}catch(ex){}});
    fsl.innerHTML=Object.entries(stopsByType).map(([t,s])=>`
      <div class="flex" style="padding:4px 0;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1))">
        <span style="flex:1;font-weight:600">${esc(t)}</span>
        <span style="color:var(--gray)">${fmtDurMS(s)}</span>
      </div>`).join('')||'<span style="color:var(--gray);font-size:calc(11px*var(--zf,1))">Aucun arrêt</span>';
  }

  // Graphs
  const prodS=d.tot_s||0;
  const stopS=stopTotal;
  drawPie('fp-pie',[
    {label:'Prod',value:prodS,color:'#16a34a'},
    {label:'Arrêts',value:stopS,color:'#dc2626'},
  ],{fLeg:9,legY:118});
  const trsS=d.trs_shift!==undefined?d.trs_shift:d.trs;
  drawGauge('fp-gauge-arc','fp-gauge-pct',trsS>=0?trsS:0);
}

async function applyFPHoraires(){
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
  await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:deb.toISOString(),fin_iso:fin.toISOString()})});
  // Recalc TRS
  const d=window._fpData;
  if(d){
    const totEquiv=d.tot_equiv||0;
    const prodRef=d.prod_ref||ST.prod_ref||200;
    const dedS=d.planned_ded_s||0;
    const netS=Math.max(1,shiftS-dedS);
    const trs=netS>0&&prodRef>0?Math.round(totEquiv/(prodRef*netS/28800)*1000)/10:0;
    document.getElementById('fp-trs').textContent=fmtTRS(trs);
  }
  // Redraw timeline with custom range
  drawTLFromISO('fp-tl',gEvts,deb.toISOString(),fin.toISOString(),(window._fpData&&window._fpData.of_list)||[]);
  closeM('m-fp-horaires');
  toast('Horaires appliqués','ok');
}

async function recalcFPTRS(){
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
    // Build debut datetime and call server
    const[dH,dM]=jour.debut.split(':').map(Number);
    const debDt=new Date(now);debDt.setHours(dH,dM,0,0);
    if(debDt>now) debDt.setDate(debDt.getDate()-1);
    const finDt=new Date(debDt.getTime()+shiftS*1000);
    await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:debDt.toISOString(),fin_iso:finDt.toISOString()})});
    const d=window._fpData;
    const totEquiv=d.tot_equiv||0;
    const prodRef=d.prod_ref||ST.prod_ref||200;
    const dedS=d.planned_ded_s||0;
    const netS=Math.max(1,shiftS-dedS);
    const trs=netS>0&&prodRef>0?Math.round(totEquiv/(prodRef*netS/28800)*1000)/10:0;
    document.getElementById('fp-trs').textContent=fmtTRS(trs);
  }
}

async function confirmFinPoste(){
  const stops=gEvts.filter(e=>e.type&&!e.is_degrade);
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
  // Net prod = each OF's (fin-debut) minus overlapping stops within that period
  const _fpOfPs=ofList.map(o=>({s:pSec(o.debut||'0:0:0'),e:pSec(o.fin||'0:0:0')})).filter(p=>p.e>p.s);
  const _fpStPs=stops.map(e=>({s:pSec(e.debut||'0:0:0'),e:pSec(e.fin||'0:0:0')})).filter(e=>e.e>e.s);
  let _fpNet=0;
  _fpOfPs.forEach(pp=>{
    const ov=_fpStPs.map(sv=>({s:Math.max(sv.s,pp.s),e:Math.min(sv.e,pp.e)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    _fpNet+=Math.max(0,pp.e-pp.s-mg.reduce((a,o)=>a+(o.e-o.s),0));
  });
  const dur_prod_sans_arret_s=_fpNet;
  const posteRow={
    date:new Date().toLocaleDateString('fr-FR'),
    pilot:ST.pilot||'',
    copilote:ST.form&&ST.form.copilote||'',
    poste:ST.poste||'',
    nb_of:fpData?fpData.nb_of||0:0,
    prod_total,
    tot_equiv:fpData?fpData.tot_equiv||0:0,
    trs_shift:fpData?fpData.trs_shift||0:0,
    cadence_h:fpData&&fpData.cadence_h||0,
    pause_min:Math.round(pause_s/60),
    nett_min:Math.round(nett_s/60),
    reunion_min:fpData&&fpData.reunion_min||0,
    depassement_min:fpData&&fpData.depassement_min||0,
    nb_fibre_chg:fpData&&fpData.nb_fibre_chg||0,
    comment:'',
    temps_ouverture_min:fpData&&fpData.ouverture_min||0,
    temps_utile_min:fpData&&fpData.temps_utile_min||0,
    temps_fonctionnement_min:fpData&&fpData.temps_fonctionnement_min||0,
    temps_arret_min:fpData&&fpData.net_stop_min||0,
    cadence_ref_pcs_min:fpData&&fpData.cadence_ref_pcs_min||0,
    perte_cadence_min:fpData&&fpData.perte_cadence_min||0,
    degrade_min:fpData&&fpData.degrade_min||0,
    pcs_theorique:fpData&&fpData.pcs_theorique||0,
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
  const _dpId='dpat_'+svgId;
  const tS=new Date(startISO).getTime(),tE=new Date(endISO).getTime();
  const span=tE-tS;
  if(span<=0){svg.innerHTML=`<rect x="0" y="0" width="${W}" height="${H}" fill="#1e293b" rx="4"/>`;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  let html=`<defs><pattern id="${_dpId}" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="4" height="8" fill="#16a34a"/><rect x="4" y="0" width="4" height="8" fill="#fef08a"/></pattern></defs>`;
  html+=`<rect x="0" y="0" width="${W}" height="${H}" fill="#1e2d48" rx="4"/>`;
  // Prod background (green)
  html+=`<rect x="0" y="0" width="${W}" height="${H}" fill="#166534" rx="4" opacity=".65"/>`;
  // Events (stops)
  (evts||[]).forEach(ev=>{
    const t1=parseHMStoT(ev.debut,ev.date),t2=parseHMStoT(ev.fin,ev.date);
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||(t1+300000));if(x2<=x1)return;
    const cat=ev.cat||'autre';
    const col=ev.is_degrade?`url(#${_dpId})`:(STOP_COL[cat]||'#94a3b8');
    html+=`<rect x="${x1}" y="0" width="${Math.max(2,x2-x1)}" height="${H}" fill="${col}" rx="2" opacity=".9"/>`;
  });
  // Current live stop for active session
  if(isCurrent&&_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active){
    const se=_curStopElap+(Date.now()-_lastPoll)/1000;
    const sT=tE-se*1000;
    const x1=toX(sT),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="0" width="${x2-x1}" height="${H}" fill="#dc2626" rx="2" opacity=".9"/>`;
  }
  // Dégradé overlay
  const _ofStartMsKpi=ST.of_start_iso?new Date(ST.of_start_iso).getTime():tS;
  (window._degradePeriodsIso||[]).forEach(function(p){
    if(!p.start||!p.end) return;
    const _d0=Math.max(new Date(p.start).getTime(),_ofStartMsKpi);
    const _d1=new Date(p.end).getTime();
    if(_d1<=_d0) return;
    const x1=toX(_d0),x2=toX(_d1);
    if(x2>x1) html+=`<rect x="${x1}" y="0" width="${x2-x1}" height="${H}" fill="url(#${_dpId})" rx="2" opacity=".85"/>`;
  });
  if(isCurrent&&window._degradeActive&&ST.degrade_start_iso&&ST.prod_active){
    const _d0=Math.max(new Date(ST.degrade_start_iso).getTime(),_ofStartMsKpi);
    const _d1=Date.now();
    if(_d1>_d0){const x1=toX(_d0),x2=toX(_d1);if(x2>x1)html+=`<rect x="${x1}" y="0" width="${x2-x1}" height="${H}" fill="url(#${_dpId})" rx="2" opacity=".85"/>`;}
  }
  const fT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  const fs=isCurrent?11:9;
  html+=`<text x="4" y="${H-4}" font-size="${fs}" fill="#fff">${fT(tS)}</text>`;
  html+=`<text x="${W-40}" y="${H-4}" font-size="${fs}" fill="#fff">${fT(tE)}</text>`;
  if(span>3600000){const mid=(tS+tE)/2;html+=`<line x1="${W/2}" y1="0" x2="${W/2}" y2="${H}" stroke="#475569" stroke-width="1" stroke-dasharray="3,3"/><text x="${W/2-18}" y="${H-4}" font-size="${fs}" fill="#e2e8f0">${fT(mid)}</text>`;}
  svg.innerHTML=html;
}

// ── KPI helpers ──
function _kpiParseFR(fr){if(!fr)return null;const p=fr.split('/');if(p.length!==3)return null;return new Date(parseInt(p[2]),parseInt(p[1])-1,parseInt(p[0])).getTime();}
function _kpiTrsColor(t){return t>=90?'#16a34a':t>=70?'#f59e0b':t>=0?'#dc2626':'#94a3b8';}

function _kpiDrawDonut(svgId,legendId,segments){
  const svg=document.getElementById(svgId);const leg=document.getElementById(legendId);
  if(!svg)return;
  const total=segments.reduce((a,s)=>a+s.v,0);
  if(total<=0){svg.innerHTML='<text x="50" y="55" text-anchor="middle" font-size="9" fill="#94a3b8">Aucune donnée</text>';if(leg)leg.innerHTML='';return;}
  const cx=50,cy=50,R=42,r=22;let html='',a=-Math.PI/2;
  segments.forEach(s=>{
    if(s.v<=0)return;
    const ang=(s.v/total)*2*Math.PI;if(ang<0.002)return;
    const ea=a+ang,lg=ang>Math.PI?1:0;
    const x1=(cx+R*Math.cos(a)).toFixed(1),y1=(cy+R*Math.sin(a)).toFixed(1);
    const x2=(cx+R*Math.cos(ea)).toFixed(1),y2=(cy+R*Math.sin(ea)).toFixed(1);
    const ix1=(cx+r*Math.cos(a)).toFixed(1),iy1=(cy+r*Math.sin(a)).toFixed(1);
    const ix2=(cx+r*Math.cos(ea)).toFixed(1),iy2=(cy+r*Math.sin(ea)).toFixed(1);
    html+=`<path d="M${x1},${y1} A${R},${R} 0 ${lg},1 ${x2},${y2} L${ix2},${iy2} A${r},${r} 0 ${lg},0 ${ix1},${iy1} Z" fill="${s.col}"/>`;
    a=ea;
  });
  const top=segments.reduce((a,s)=>s.v>a.v?s:a,segments[0]);
  const topPct=Math.round(top.v/total*100);
  html+=`<text x="${cx}" y="${cy+4}" text-anchor="middle" font-size="14" font-weight="800" fill="#1e293b">${topPct}%</text>`;
  html+=`<text x="${cx}" y="${cy+13}" text-anchor="middle" font-size="8" fill="#64748b">${esc(top.label)}</text>`;
  svg.innerHTML=html;
  if(leg) leg.innerHTML=segments.map(s=>{const p=Math.round(s.v/total*100);return `<div style="display:flex;align-items:center;gap:4px"><div style="width:8px;height:8px;border-radius:2px;background:${s.col};flex-shrink:0"></div><span style="color:#475569">${esc(s.label)}</span><b style="margin-left:3px;color:#1e293b">${p}%</b></div>`;}).join('');
}

function _kpiLineChart(containerId,items,valueKey,colorFn,unit,yMin,yMax){
  const el=document.getElementById(containerId);if(!el)return;
  const rect=el.getBoundingClientRect();
  const W=Math.max(rect.width||400,200);
  const H=Math.max(rect.height||100,60);
  if(!items.length){el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W/2}" y="${H/2}" text-anchor="middle" font-size="10" fill="#94a3b8">Aucune donnée</text></svg>`;return;}
  const vals=items.map(it=>it[valueKey]||0);
  const minV=yMin!==undefined?yMin:Math.max(0,Math.min(...vals)-5);
  const maxV=yMax!==undefined?yMax:Math.max(...vals,1)+2;
  const padL=36,padR=65,padT=22,padB=80;
  const gW=W-padL-padR,gH=H-padT-padB;
  const toX=i=>padL+i/(Math.max(items.length-1,1))*gW;
  const toY=v=>padT+gH*(1-(v-minV)/(maxV-minV||1));
  // grid
  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%">`;
  const nTicks=4;
  for(let t=0;t<=nTicks;t++){
    const v=minV+(maxV-minV)*t/nTicks;
    const y=toY(v);
    svg+=`<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W-padR}" y2="${y.toFixed(1)}" stroke="#f1f5f9" stroke-width="1"/>`;
    svg+=`<text x="${padL-3}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="13" fill="#64748b">${v%1?v.toFixed(0):v}${unit||''}</text>`;
  }
  // area fill
  let areaD=`M${toX(0).toFixed(1)},${(H-padB).toFixed(1)}`;
  items.forEach((it,i)=>{areaD+=` L${toX(i).toFixed(1)},${toY(it[valueKey]||0).toFixed(1)}`;});
  areaD+=` L${toX(items.length-1).toFixed(1)},${(H-padB).toFixed(1)} Z`;
  const areaCol=colorFn?colorFn(items[0]):'#6366f1';
  svg+=`<path d="${areaD}" fill="${areaCol}" opacity=".08"/>`;
  // line
  let lineD='';
  items.forEach((it,i)=>{lineD+=(i===0?'M':'L')+toX(i).toFixed(1)+','+toY(it[valueKey]||0).toFixed(1)+' ';});
  svg+=`<path d="${lineD}" fill="none" stroke="${areaCol}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
  // dots + value labels
  items.forEach((it,i)=>{
    const x=toX(i),y=toY(it[valueKey]||0);
    const col=colorFn?colorFn(it):areaCol;
    svg+=`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
    const v=it[valueKey]||0;
    svg+=`<text x="${x.toFixed(1)}" y="${(y-6).toFixed(1)}" text-anchor="middle" font-size="13" font-weight="700" fill="${col}">${v%1?v.toFixed(1):v}</text>`;
    // X labels: date (line1) + poste (line2) — bigger, black
    const lblParts=(it._xLabel||'').split('\n');
    svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+3).toFixed(1)}) rotate(35)" font-size="13" font-weight="700" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[0]||'')}</text>`;
    if(lblParts[1]) svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+17).toFixed(1)}) rotate(35)" font-size="12" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[1]||'')}</text>`;
  });
  svg+='</svg>';
  el.innerHTML=svg;
}

function _kpiBarChart(containerId,items,valueKey,colorFn,unit){
  const el=document.getElementById(containerId);if(!el)return;
  const rect=el.getBoundingClientRect();
  const W=Math.max(rect.width||400,200);
  const H=Math.max(rect.height||100,60);
  if(!items.length){el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W/2}" y="${H/2}" text-anchor="middle" font-size="10" fill="#94a3b8">Aucune donnée</text></svg>`;return;}
  const vals=items.map(it=>it[valueKey]||0);
  const maxV=Math.max(...vals,1);
  const padL=30,padR=8,padT=8,padB=46;
  const gW=W-padL-padR,gH=H-padT-padB;
  const n=items.length;
  const barW=Math.max(3,Math.min(30,gW/n-4));
  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%">`;
  [0,0.5,1].forEach(t=>{
    const v=maxV*t;const y=padT+gH*(1-t);
    svg+=`<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W-padR}" y2="${y.toFixed(1)}" stroke="#f1f5f9" stroke-width="1"/>`;
    svg+=`<text x="${padL-3}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="11" fill="#94a3b8">${Math.round(v)}${t>0&&unit?unit:''}</text>`;
  });
  items.forEach((it,i)=>{
    const v=it[valueKey]||0;
    const x=padL+(i/n)*gW+(gW/n-barW)/2;
    const bH=v>0?(v/maxV)*gH:0;
    const y=padT+gH-bH;
    const col=colorFn?colorFn(it):'#f59e0b';
    svg+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW}" height="${Math.max(0,bH).toFixed(1)}" fill="${col}" rx="2" opacity=".85"/>`;
    if(v>0)svg+=`<text x="${(x+barW/2).toFixed(1)}" y="${(y-2).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="700" fill="${col}">${v%1?v.toFixed(1):v}</text>`;
    const lbl=(it._xLabel||'').split('\n')[0];
    svg+=`<text transform="translate(${(x+barW/2).toFixed(1)},${(H-padB+4).toFixed(1)}) rotate(35)" font-size="11" fill="#64748b" dominant-baseline="hanging">${esc(lbl)}</text>`;
  });
  svg+='</svg>';
  el.innerHTML=svg;
}

function _kpiDualLineChart(containerId,items,series,opts){
  const el=document.getElementById(containerId);if(!el)return;
  const rect=el.getBoundingClientRect();
  const W=Math.max(rect.width||400,200);
  const H=Math.max(rect.height||100,60);
  if(!items.length){el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W/2}" y="${H/2}" text-anchor="middle" font-size="10" fill="#94a3b8">Aucune donnée</text></svg>`;return;}
  const refLine=opts&&opts.refLine!=null?opts.refLine:null;
  const allVals=series.flatMap(s=>items.map(it=>parseFloat(it[s.key]||0)));
  if(refLine!=null) allVals.push(refLine);
  const minV=Math.max(0,Math.min(...allVals)-2);
  const maxV=Math.max(...allVals,1)+2;
  const padL=36,padR=70,padT=12,padB=68;
  const gW=W-padL-padR,gH=H-padT-padB;
  const toX=i=>padL+i/(Math.max(items.length-1,1))*gW;
  const toY=v=>padT+gH*(1-(v-minV)/(maxV-minV||1));
  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%">`;
  [0,0.25,0.5,0.75,1].forEach(t=>{
    const v=minV+(maxV-minV)*t;const y=toY(v);
    svg+=`<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W-padR}" y2="${y.toFixed(1)}" stroke="#f1f5f9" stroke-width="1"/>`;
    svg+=`<text x="${padL-3}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="13" fill="#94a3b8">${v.toFixed(0)}</text>`;
  });
  // Ligne cible (refLine)
  if(refLine!=null){
    const ry=toY(refLine);
    svg+=`<line x1="${padL}" y1="${ry.toFixed(1)}" x2="${W-padR}" y2="${ry.toFixed(1)}" stroke="#dc2626" stroke-width="1.5" stroke-dasharray="6,3"/>`;
    svg+=`<text x="${W-padR+3}" y="${(ry+4).toFixed(1)}" font-size="13" fill="#dc2626" font-weight="700">Cible ${refLine}</text>`;
  }
  series.forEach(ser=>{
    let lineD='';
    items.forEach((it,i)=>{lineD+=(i===0?'M':'L')+toX(i).toFixed(1)+','+toY(parseFloat(it[ser.key]||0)).toFixed(1)+' ';});
    svg+=`<path d="${lineD}" fill="none" stroke="${ser.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" ${ser.dash?`stroke-dasharray="${ser.dash}"`:''}/>`;
    items.forEach((it,i)=>{
      const x=toX(i),y=toY(parseFloat(it[ser.key]||0));
      const v=parseFloat(it[ser.key]||0);
      svg+=`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3" fill="${ser.color}" stroke="#fff" stroke-width="1.5"/>`;
      svg+=`<text x="${x.toFixed(1)}" y="${(y-7).toFixed(1)}" text-anchor="middle" font-size="13" fill="${ser.color}" font-weight="700">${v%1?v.toFixed(1):v}</text>`;
    });
  });
  items.forEach((it,i)=>{
    const x=toX(i);
    const lblParts=(it._xLabel||'').split('\n');
    svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+3).toFixed(1)}) rotate(35)" font-size="13" font-weight="700" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[0]||'')}</text>`;
    if(lblParts[1]) svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+17).toFixed(1)}) rotate(35)" font-size="12" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[1]||'')}</text>`;
  });
  let lx=padL;
  series.forEach(ser=>{
    svg+=`<line x1="${lx}" y1="${H-7}" x2="${lx+14}" y2="${H-7}" stroke="${ser.color}" stroke-width="2" ${ser.dash?`stroke-dasharray="${ser.dash}"`:''}/>`;
    svg+=`<text x="${lx+16}" y="${H-3}" font-size="13" fill="#475569">${esc(ser.label||ser.key)}</text>`;
    lx+=90;
  });
  if(refLine!=null){svg+=`<line x1="${lx}" y1="${H-7}" x2="${lx+14}" y2="${H-7}" stroke="#dc2626" stroke-width="1.5" stroke-dasharray="6,3"/>`;svg+=`<text x="${lx+16}" y="${H-3}" font-size="13" fill="#dc2626">Cible</text>`;}
  svg+='</svg>';
  el.innerHTML=svg;
}

function _kpiInitDates(){
  const fi=document.getElementById('kpi-from'),ti=document.getElementById('kpi-to');
  if(!fi||!ti)return;
  if(!fi.value){const d=new Date();d.setMonth(d.getMonth()-6);fi.value=d.toISOString().slice(0,10);}
  if(!ti.value){ti.value=new Date().toISOString().slice(0,10);}
}

async function loadKPI(){
  _kpiInitDates();
  const fi=document.getElementById('kpi-from'),ti=document.getElementById('kpi-to');
  const _parseLocalDate=s=>{if(!s)return null;const p=s.split('-');return new Date(+p[0],+p[1]-1,+p[2]).getTime();};
  const fromMs=fi&&fi.value?_parseLocalDate(fi.value):0;
  const toMs=ti&&ti.value?_parseLocalDate(ti.value)+86399999:Date.now();
  const pSec=s=>s?s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0):0;

  const prParams=new URLSearchParams();
  if(fi&&fi.value)prParams.set('date_from',fi.value);
  if(ti&&ti.value)prParams.set('date_to',ti.value);
  const [histData,evtData,prData]=await Promise.all([apiFetch('/api/history'),apiFetch('/api/events_list'),apiFetch('/api/period_report?'+prParams.toString())]);
  const allRows=Array.isArray(histData)?histData:[];
  const allEvts=Array.isArray(evtData)?evtData:[];

  const inRange=dateStr=>{const ms=_kpiParseFR(dateStr);return ms&&ms>=fromMs&&ms<=toMs;};
  const rows=allRows.filter(r=>inRange(r.date)&&(r.type||'').trim().toLowerCase()!=='');
  const prodRows=allRows.filter(r=>inRange(r.date)&&['production','prod',''].includes((r.type||'').trim().toLowerCase()));
  const evts=allEvts.filter(e=>inRange(e.date));

  // Build sessions (one per pilot+date+poste)
  const sessMap={};
  const sessRowsMap={};
  prodRows.forEach(r=>{
    const key=(r.pilote||'')+'||'+(r.date||'')+'||'+(r.poste||'');
    if(!sessMap[key]) sessMap[key]={pilot:r.pilote||'',date:r.date||'',poste:r.poste||'',nb_of:0,tot_equiv:0,tot_s:0,tot_qte:0,trs_vals:[]};
    if(!sessRowsMap[key]) sessRowsMap[key]=[];
    const s=sessMap[key];
    s.nb_of++;
    const eq=parseFloat(r.equiv||0)||0; s.tot_equiv+=eq;
    const qte=parseFloat(r.qte_fab||0)||0; s.tot_qte+=qte;
    const ds=pSec(r.debut||'0:0:0'),fs=pSec(r.fin||'0:0:0'),dur=Math.max(0,fs-ds); s.tot_s+=dur;
    const t=parseFloat(r.trs||0);if(t>0)s.trs_vals.push(t);
    sessRowsMap[key].push(r);
  });
  // Order chronologically
  const sessArr=Object.values(sessMap).sort((a,b)=>{const da=_kpiParseFR(a.date)||0,db=_kpiParseFR(b.date)||0;return da-db;});
  sessArr.forEach(s=>{
    s.trs=s.trs_vals.length?Math.round(s.trs_vals.reduce((a,v)=>a+v,0)/s.trs_vals.length*10)/10:-1;
    s.cad=s.tot_s>60?Math.round(s.tot_equiv/(s.tot_s/3600)*10)/10:0;
    s.cad_qte=s.tot_s>60&&s.tot_qte>0?Math.round(s.tot_qte/(s.tot_s/3600)*10)/10:0;
    const evtKey=s.pilot+'||'+s.date+'||'+s.poste;
    const sessEvts=evts.filter(e=>(e.pilote||'')+'||'+(e.date||'')+'||'+(e.poste||'')==evtKey);
    const _ivs=sessEvts.map(e=>[pSec(e.debut||'0:0:0'),pSec(e.fin||'0:0:0')]).filter(([d,f])=>f>d).sort((a,b)=>a[0]-b[0]);
    const _mg=[];for(const [d,f] of _ivs){if(_mg.length&&d<=_mg[_mg.length-1][1])_mg[_mg.length-1][1]=Math.max(_mg[_mg.length-1][1],f);else _mg.push([d,f]);}
    s.stop_min=Math.round(_mg.reduce((a,[d,f])=>a+(f-d),0)/60);
    // Fibre changes: count transitions between different fibre values
    const sRows=(sessRowsMap[evtKey]||[]).sort((a,b)=>pSec(a.debut||'0:0')-pSec(b.debut||'0:0'));
    let fibChg=0;
    for(let i=1;i<sRows.length;i++){const fa=(sRows[i-1].fibre||'').trim(),fb=(sRows[i].fibre||'').trim();if(fa&&fb&&fa!==fb)fibChg++;}
    s.nb_fibre_chg=fibChg;
    // Prod/arrêt ratio
    const tot_dur=s.tot_s+s.stop_min*60;
    s.prod_pct=tot_dur>0?Math.round(s.tot_s/tot_dur*100):0;
    s.stop_pct=100-s.prod_pct;
    // Label X axis
    const dParts=(s.date||'').split('/');
    const dateShort=dParts.length===3?dParts[0]+'/'+dParts[1]:s.date;
    s._xLabel=dateShort+'\n'+s.poste;
  });

  const nbSess=sessArr.length;
  const totalEquiv=sessArr.reduce((a,s)=>a+s.tot_equiv,0);
  const totalOF=sessArr.reduce((a,s)=>a+s.nb_of,0);
  const totalQte=sessArr.reduce((a,s)=>a+s.tot_qte,0);
  const totalProdS=sessArr.reduce((a,s)=>a+s.tot_s,0);
  const totalStopMin=sessArr.reduce((a,s)=>a+s.stop_min,0);
  // Use period_report for accurate aggregate values (Postes-filtered, merged intervals, weighted TRS)
  const pr=prData&&prData.ok?prData:{};
  const avgTRS=pr.trs_periode!==undefined?pr.trs_periode:-1;
  const avgCadH=pr.cadence_h!==undefined?pr.cadence_h:(totalProdS>0?Math.round(totalEquiv/(totalProdS/3600)*10)/10:0);
  const avgOFperSess=nbSess?Math.round(totalOF/nbSess*10)/10:0;
  const avgEquivPerSess=nbSess?Math.round(totalEquiv/nbSess*10)/10:0;



  // ── 3 courbes (compact, côte à côte) ──
  // TRS : utiliser sessions_detail du period_report (calcul pondéré correct)
  const prSd=(prData&&prData.sessions_detail)||[];
  prSd.forEach(s=>{const dp=(s.date||'').split('/');s._xLabel=(dp[0]||'')+'/'+(dp[1]||'')+'\n'+(s.poste||'');});
  const trsItems=prSd.length>0?prSd.filter(s=>s.trs>=0):sessArr.filter(s=>s.trs>=0);
  _kpiLineChart('kpi-trs-chart',trsItems,'trs',it=>_kpiTrsColor(it.trs),'%',0,100);
  _kpiLineChart('kpi-arr-chart',sessArr,'stop_min',()=>'#dc2626','min');
  _kpiLineChart('kpi-of-chart',sessArr,'nb_of',()=>'#6366f1','');

  // ── Cadence dual-line chart (éq./h + pièces/h) ──
  const _cadRefH=pr.cadence_ref_pcs_min>0?Math.round(pr.cadence_ref_pcs_min*60):null;
  _kpiDualLineChart('kpi-cad-chart',sessArr,[
    {key:'cad',color:'#f59e0b',label:'Éq./h'}
  ],{refLine:_cadRefH});

  // ── Évolution changements de fibre par poste ──
  _kpiLineChart('kpi-fibre-chart',sessArr,'nb_fibre_chg',()=>'#8b5cf6','');

  // ── Évolution équivalence ──
  _kpiLineChart('kpi-qte-chart',sessArr,'tot_equiv',()=>'#16a34a','');

  // ── Donut Prod/Arrêts ──
  _kpiDrawDonut('kpi-donut','kpi-donut-legend',[
    {label:'Prod',v:totalProdS,col:'#16a34a'},
    {label:'Arrêts',v:totalStopMin*60,col:'#dc2626'},
  ]);

  // ── Pareto ──
  const stopMap={};const stopCat={};
  evts.forEach(e=>{
    if(!e.type)return;
    const dur=Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0'));
    stopMap[e.type]=(stopMap[e.type]||0)+dur;
    if(!stopCat[e.type])stopCat[e.type]=e.cat||'autre';
  });
  const paretoArr=Object.entries(stopMap).sort((a,b)=>b[1]-a[1]);
  const paretoTotal=paretoArr.reduce((a,[,v])=>a+v,0);
  const parEl=document.getElementById('kpi-pareto-new');
  if(parEl){
    if(!paretoArr.length){parEl.innerHTML='<div style="color:#94a3b8;font-size:calc(11px*var(--zf,1))">Aucun arrêt</div>';}
    else{
      const maxP=paretoArr[0][1];
      let cumul=0;
      parEl.innerHTML=paretoArr.slice(0,12).map(([type,s])=>{
        const pct=Math.round(s/maxP*100);
        const min=Math.round(s/60);
        const pctTot=paretoTotal>0?Math.round(s/paretoTotal*100):0;
        cumul+=pctTot;
        const col=STOP_COL[stopCat[type]]||'#94a3b8';
        return `<div>
          <div style="display:flex;justify-content:space-between;font-size:calc(10px*var(--zf,1));color:#374151;margin-bottom:2px">
            <div style="display:flex;align-items:center;gap:3px;min-width:0"><div style="width:7px;height:7px;border-radius:1px;background:${col};flex-shrink:0"></div><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:110px">${esc(type)}</span></div>
            <span style="white-space:nowrap;color:#6b7280;flex-shrink:0">${min}m <b style="color:#1e293b">${pctTot}%</b></span>
          </div>
          <div style="background:#f1f5f9;border-radius:3px;height:9px;overflow:hidden">
            <div style="width:${pct}%;background:${col};height:100%;border-radius:3px"></div>
          </div>
        </div>`;
      }).join('');
    }
  }
}

// ── HISTORY ROW DETAIL ──
function showHistRowDetail(key){
  let r=window._rowMap&&window._rowMap[String(key)];
  if(!r) return;
  const isProd=r._rowType==='prod';
  let ofEvts=[];
  if(isProd){
    const _hm2s=hm=>{if(!hm)return 0;const[h,m,s2]=((hm||'')+':0:0').split(':').map(Number);return(h||0)*3600+(m||0)*60+(s2||0);};
    const debS=_hm2s(r.debut), finS=_hm2s(r.fin)||86400;
    ofEvts=(window._histEvtsAll||[]).filter(e=>{
      if(!e.date||e.date!==r.date) return false;
      if(e.pilote&&r.pilote&&e.pilote!==r.pilote) return false;
      const t=_hm2s(e.debut); return t>=debS&&t<=finS;
    });
    // Set up net prod helpers from this OF's events (same as showRjOfDetail)
    const _h2ms=hm=>{if(!hm)return 0;const[h,m,s3]=((hm||'')+':0:0').split(':').map(Number);return((h||0)*3600+(m||0)*60+(s3||0))*1000;};
    const _stE=ofEvts.filter(e=>!e.is_degrade).map(e=>({s:_h2ms(e.debut),e:_h2ms(e.fin)})).filter(e=>e.e>e.s);
    const _dgE=ofEvts.filter(e=>e.is_degrade).map(e=>({s:_h2ms(e.debut),e:_h2ms(e.fin)})).filter(e=>e.e>e.s);
    const _mgH=evs=>{const ss=[...evs].sort((a,b)=>a.s-b.s);const m=[];ss.forEach(o=>{if(m.length&&o.s<=m[m.length-1].e)m[m.length-1].e=Math.max(m[m.length-1].e,o.e);else m.push({...o});});return m;};
    window._rptDegMin=(dH,fH)=>{const dM=_h2ms(dH),fM=_h2ms(fH);if(fM<=dM)return 0;const mg=_mgH(_dgE.map(e=>({s:Math.max(e.s,dM),e:Math.min(e.e,fM)})).filter(e=>e.e>e.s));return Math.round(mg.reduce((a,o)=>a+(o.e-o.s),0)/60000);};
    window._rptNetProd=(dH,fH)=>{const dM=_h2ms(dH),fM=_h2ms(fH);if(fM<=dM)return{netMin:0,stopMin:0};const mg=_mgH(_stE.map(e=>({s:Math.max(e.s,dM),e:Math.min(e.e,fM)})).filter(e=>e.e>e.s));const bl=mg.reduce((a,o)=>a+(o.e-o.s),0);return{netMin:Math.round(Math.max(0,fM-dM-bl)/60000),stopMin:Math.round(bl/60000)};};
    // Compute plan_stop_s from ofEvts + budget config if not set in row (historique rows lack this field)
    if(!r.plan_stop_s){
      const _getBK=t=>{const tl=(t||'').toLowerCase();if(tl.includes('grand')||tl.includes('très long'))return 'clean_grand_min';if(tl.includes('long'))return 'clean_long_min';if(tl.includes('nettoyage')||tl.includes('nett'))return 'clean_short_min';if(tl.includes('pause'))return 'pause_min';if(tl.includes('réunion')||tl.includes('reunion')||tl.includes('meeting'))return 'meeting_tol_min';return null;};
      const ap=_cfgArretsPrevus||{};
      const byK={};
      ofEvts.forEach(ev=>{const k=_getBK(ev.type);if(!k)return;const p=((ev.duree||'0:0:0')+':0').split(':').map(Number);const dm=(p[0]||0)*60+(p[1]||0)+(p[2]||0)/60;byK[k]=(byK[k]||0)+dm;});
      const planMinCalc=Object.entries(byK).reduce((a,[k,v])=>a+Math.min(v,(ap[k]||0)),0);
      r={...r,plan_stop_s:planMinCalc*60};
    }
  }
  _renderAndOpenOfDetail({...r,_rowType:isProd?'prod':'evt',trs:(r.trs!==''&&r.trs!=null)?parseFloat(r.trs):-1}, ofEvts);
}

// ── HISTORY ──
let _histFilters=new Set(); // filtres actifs

function toggleHistFilter(btn){
  const key=btn.dataset.hf;
  if(_histFilters.has(key)){
    _histFilters.delete(key);
    btn.style.background='none';
    btn.style.color=btn.style.borderColor; // remettre couleur texte
    btn.style.opacity='1';
  } else {
    _histFilters.add(key);
    btn.style.background=btn.style.borderColor;
    btn.style.color='#fff';
  }
  _applyHistFilter();
}

function _histMatchFilter(r){
  if(!_histFilters.size) return true;
  const t=String(r.type||r._rowType||'').toLowerCase();
  const isProd=r._rowType==='prod'||t==='production'||t==='prod'||t==='';
  if(_histFilters.has('production')&&isProd) return true;
  if(_histFilters.has('nettoyage')&&(t.includes('nettoyage')||t.includes('nett'))) return true;
  if(_histFilters.has('pause')&&t.includes('pause')) return true;
  if(_histFilters.has('reunion')&&(t.includes('réunion')||t.includes('reunion')||t.includes('meeting'))) return true;
  if(_histFilters.has('arret')&&!isProd&&!t.includes('nettoyage')&&!t.includes('pause')&&!(t.includes('réunion')||t.includes('reunion')||t.includes('meeting'))) return true;
  return false;
}

function filterHistBySearch(){
  const q=(document.getElementById('hist-search')?.value||'').toLowerCase().trim();
  document.querySelectorAll('#hist-bd tr[data-hftype]').forEach(tr=>{
    if(!q){tr.dataset.searchHidden='';return;}
    const ofTxt=(tr.cells[1]?.textContent||'').toLowerCase();
    tr.dataset.searchHidden=ofTxt.includes(q)?'':'1';
  });
  _applyHistFilter();
}

function _applyHistFilter(){
  const rows=document.querySelectorAll('#hist-bd tr[data-hftype]');
  rows.forEach(tr=>{
    if(tr.dataset.searchHidden==='1'){tr.style.display='none';return;}
    const t=tr.dataset.hftype||'';
    const isProd=t==='prod';
    const isDegrade=t==='degrade';
    const isNett=t.includes('nettoyage')||t.includes('nett');
    const isPause=t.includes('pause');
    const isReunion=t.includes('réunion')||t.includes('reunion')||t.includes('meeting');
    const isArret=!isProd&&!isDegrade&&!isNett&&!isPause&&!isReunion;
    let show=!_histFilters.size;
    if(!show){
      if(_histFilters.has('production')&&isProd) show=true;
      if(_histFilters.has('degrade')&&isDegrade) show=true;
      if(_histFilters.has('nettoyage')&&isNett) show=true;
      if(_histFilters.has('pause')&&isPause) show=true;
      if(_histFilters.has('reunion')&&isReunion) show=true;
      if(_histFilters.has('arret')&&isArret) show=true;
    }
    tr.style.display=show?'':'none';
  });
}

function _resetHistFilters(){
  _histFilters.clear();
  document.querySelectorAll('.hf-btn').forEach(b=>{
    b.style.background='none';
    b.style.color=b.style.borderColor;
    b.style.opacity='1';
  });
}

async function loadHist(){
  const today=new Date().toISOString().slice(0,10);
  const from=document.getElementById('hist-from').value||today;
  const to=document.getElementById('hist-to').value||today;
  const [declData,evtData]=await Promise.all([
    apiFetch(`/api/history?from=${from}&to=${to}`),
    apiFetch(`/api/events_list?from=${from}&to=${to}`)
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
  hd.innerHTML='<th>Type</th><th>OF</th><th>Fibre</th><th>Date</th><th>Poste</th><th>Pilote</th><th>Nb pers</th><th>Début</th><th>Fin</th><th>Détails</th><th>Qté/Durée</th><th>TRS/Info</th><th>Commentaire</th><th>Actions</th>';
  if(!allRows.length){bd.innerHTML='<tr><td colspan="12" style="text-align:center;color:var(--gray);padding:16px">Aucune donnée sur cette période</td></tr>';return;}
  window._rowMap=window._rowMap||{};
  window._histEvtsAll=evtsFiltered; // pour showHistRowDetail
  bd.innerHTML=allRows.map(r=>{
    const key=r.row_num||r.debut;
    window._rowMap[String(key)]=r;
    const isProd=r._rowType==='prod';
    const rt=String(r.type||'').toLowerCase();
    const hftype=isProd?'prod':(r.is_degrade?'degrade':rt||'arret');
    const t=parseFloat(r.trs||0);
    const tag=isProd?'<span class="row-tag tag-p">🏭 Prod</span>':(r.is_degrade?'<span class="row-tag tag-e" style="border-color:#ca8a04;color:#ca8a04">🟡 Dégradé</span>':(rt.includes('nett')?'<span class="row-tag tag-n">🧹 Nett.</span>':rt.includes('pause')?'<span class="row-tag tag-n" style="border-color:#f59e0b;color:#f59e0b">⏸ Pause</span>':(rt.includes('réunion')||rt.includes('reunion'))?'<span class="row-tag tag-n" style="border-color:#8b5cf6;color:#8b5cf6">👥 Réunion</span>':'<span class="row-tag tag-e">⛔ Arrêt</span>'));
    const details=isProd?esc(r.taille||''):esc(r.type||'');
    const qty=isProd?esc(String(r.qte_fab||'')):esc(r.duree||'');
    const info=isProd&&t>0?`<span class="${t>=90?'tg':t>=75?'tm':'tb'}">${fmtTRS(t)}</span>`:'—';
    const cmt=esc(r.comment||'');
    const fbrH=r.fibre||'';const fbrShH=esc(fbrH.slice(0,9));
    return `<tr class="${isProd?'row-prod':'row-evt'}" data-hftype="${esc(hftype)}">
      <td>${tag}</td><td style="font-weight:700;color:#1e3a8a;text-decoration:underline;cursor:pointer" onclick="showHistRowDetail('${esc(String(key))}')" title="Voir détail">${esc(r.of||r.type||'—')}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:#6366f1;font-weight:600;cursor:${fbrH?'pointer':''}" title="${esc(fbrH)}" onclick="${fbrH?'showFibre(\''+esc(fbrH)+'\')':''}">${fbrShH}${fbrH.length>9?'…':''}</td>
      <td style="font-size:calc(10px*var(--zf,1))">${esc(r.date||'')}</td><td style="font-size:calc(10px*var(--zf,1))">${esc(r.poste||'')}</td>
      <td>${esc(r.pilote||'')}</td><td style="text-align:center;font-size:calc(10px*var(--zf,1));color:#374151">${isProd?esc(r.nb_pers||''):'—'}</td><td>${esc(r.debut||'')}</td><td>${esc(r.fin||'')}</td>
      <td style="font-size:calc(11px*var(--zf,1))">${details}</td><td style="font-size:calc(11px*var(--zf,1))">${qty}</td><td>${info}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:var(--gray);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${cmt}">${cmt}</td>
      <td><button onclick="openEditRow('${esc(String(key))}')" style="background:#6366f1;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:calc(15px*var(--zf,1));cursor:pointer;font-weight:700" title="Modifier">✏</button></td>
    </tr>`;
  }).join('');
  _applyHistFilter();
}

// ── RAPPORTS DES POSTES ──
async function reloadAndLoadRapports(){
  toast('Rechargement depuis Excel…','ok');
  await apiFetch('/api/reload_excel',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await loadRapports();
}
async function loadRptJour(){
  const _rjf=document.getElementById('rj-from'),_rjt=document.getElementById('rj-to');
  // Toujours réinitialiser les filtres à l'ouverture de l'onglet
  if(_rjf) _rjf.value='';
  if(_rjt) _rjt.value='';
  const _pSel=document.getElementById('rj-pilot'); if(_pSel) _pSel.value='';
  const _qSel=document.getElementById('rj-poste'); if(_qSel) _qSel.value='';
  _rjSetBanner('Rapport des 3 derniers postes');
  // Populate pilot/poste selects from past_sessions
  const sessions=await apiFetch('/api/past_sessions');
  if(!sessions) return;
  const pilots=[...new Set(sessions.map(s=>s.pilot).filter(Boolean))].sort();
  const postes=[...new Set(sessions.map(s=>s.poste).filter(Boolean))].sort();
  const pSel=document.getElementById('rj-pilot');
  const qSel=document.getElementById('rj-poste');
  if(pSel){
    const cur=pSel.value;
    pSel.innerHTML='<option value="">Tous</option>'+pilots.map(p=>`<option value="${esc(p)}" ${p===cur?'selected':''}>${esc(p)}</option>`).join('');
  }
  if(qSel){
    const cur=qSel.value;
    qSel.innerHTML='<option value="">Tous</option>'+postes.map(p=>`<option value="${esc(p)}" ${p===cur?'selected':''}>${esc(p)}</option>`).join('');
  }
  // Auto-load: trouver les 3 derniers postes et stocker pour bandeau + limit API
  window._rjAutoLast3=[];
  if(sessions&&sessions.length){
    const _sorted=[...sessions].sort((a,b)=>{
      const _dmy=s=>{try{const p=(s.date||'').split('/');return new Date(p[2]+'-'+p[1]+'-'+p[0]).getTime();}catch(e){return 0;}};
      return _dmy(b)-_dmy(a);
    });
    window._rjAutoLast3=_sorted.slice(0,3);
    if(window._rjAutoLast3.length){
      const _toDate=window._rjAutoLast3[0].date;
      const _fromDate=window._rjAutoLast3[window._rjAutoLast3.length-1].date;
      const _toISO=_toDate?_toDate.split('/').reverse().join('-'):'';
      const _fromISO=_fromDate?_fromDate.split('/').reverse().join('-'):'';
      if(_rjf&&_fromISO)_rjf.value=_fromISO;
      if(_rjt&&_toISO)_rjt.value=_toISO;
    }
  }
  calcPeriodReport(true);
}
async function calcPeriodReport(autoLoad){
  // En mode autoLoad, on n'utilise PAS les filtres date/pilote/poste
  // On laisse l'API renvoyer les max_sessions=3 derniers postes sans restriction de date
  const from=autoLoad?'':document.getElementById('rj-from').value;
  const to=autoLoad?'':document.getElementById('rj-to').value;
  const pilot=autoLoad?'':document.getElementById('rj-pilot').value;
  const poste=autoLoad?'':document.getElementById('rj-poste').value;
  const resultEl=document.getElementById('rj-result');
  if(!resultEl) return;
  if(autoLoad){_rjSetBanner('Rapport des 3 derniers postes');}
  else if(from||to){const _fp=from?from.split('-').reverse().join('/'):'…';const _tp=to?to.split('-').reverse().join('/'):'…';_rjSetBanner('Rapport du '+_fp+' au '+_tp);}
  resultEl.innerHTML='<div style="padding:40px;text-align:center;color:var(--gray)">Calcul en cours…</div>';
  let url='/api/period_report?';
  if(from) url+='date_from='+encodeURIComponent(from)+'&';
  if(to)   url+='date_to='+encodeURIComponent(to)+'&';
  if(pilot) url+='pilot='+encodeURIComponent(pilot)+'&';
  if(poste) url+='poste='+encodeURIComponent(poste)+'&';
  if(autoLoad) url+='max_sessions=3&skip_current=1&';
  const d=await apiFetch(url);
  if(!d||!d.ok){resultEl.innerHTML='<div style="padding:40px;text-align:center;color:#dc2626">Erreur ou aucune donnée</div>';return;}
  if(d.nb_sessions===0){resultEl.innerHTML='<div style="padding:60px;text-align:center;color:#94a3b8"><div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">🔍</div><div style="font-size:calc(14px*var(--zf,1));font-weight:600">Aucun poste trouvé pour cette période</div></div>';return;}
  const trsCol=d.trs_periode>=90?'#16a34a':d.trs_periode>=70?'#f59e0b':d.trs_periode>=0?'#dc2626':'#94a3b8';
  const pertRaw=d.perte_cadence_min||0;
  const pertHtml=pertRaw<0?`<span style="color:#16a34a;font-weight:900">${Math.abs(Math.round(pertRaw))} min de gain</span>`:pertRaw>0?`<span style="color:#dc2626;font-weight:900">${Math.round(pertRaw)} min de perte</span>`:`<span style="color:#64748b">0 min</span>`;
  // ── Chart A : TRS par équipe — barres verticales SVG ──────────────────────────
  let chartTrsHtml='';
  if(d.sessions_detail&&d.sessions_detail.length>0){
    const sd=d.sessions_detail;
    const maxTrs=Math.max(...sd.filter(s=>s.trs>=0).map(s=>s.trs),100);
    const CH=280,padT=20,padB=56,padL=4,padR=4;
    const gH=CH-padT-padB;
    const n=sd.length;
    const WB=Math.max(22,Math.min(60,Math.floor((420-padL-padR-n*4)/n)));
    const GP=5;
    const svgW=Math.max(200,n*(WB+GP)+padL+padR);
    let svgB='',svgL='';
    sd.forEach((s,i)=>{
      const x=padL+i*(WB+GP);
      const trs=s.trs>=0?s.trs:0;
      const bh=Math.max(2,Math.round(trs/maxTrs*gH));
      const by=padT+gH-bh;
      const col=s.trs>=90?'#16a34a':s.trs>=70?'#f59e0b':s.trs>=0?'#dc2626':'#94a3b8';
      svgB+=`<rect x="${x}" y="${by}" width="${WB}" height="${bh}" fill="${col}" opacity=".85" rx="2"/>`;
      if(s.trs>=0)svgB+=`<text x="${x+WB/2}" y="${Math.max(by-3,12)}" text-anchor="middle" font-size="13" font-weight="700" fill="${col}">${s.trs.toFixed(0)}%</text>`;
      const dp=s.date.split('/');
      svgL+=`<text x="${x+WB/2}" y="${padT+gH+14}" text-anchor="middle" font-size="12" font-weight="600" fill="#374151">${esc((dp[0]||'')+'/'+(dp[1]||''))}</text>`;
      svgL+=`<text x="${x+WB/2}" y="${padT+gH+27}" text-anchor="middle" font-size="11" fill="#6366f1">${esc((s.pilot||'').slice(0,9))}</text>`;
      svgL+=`<text x="${x+WB/2}" y="${padT+gH+40}" text-anchor="middle" font-size="11" fill="#94a3b8">${esc((s.poste||'').slice(0,9))}</text>`;
    });
    const yBase=padT+gH;
    chartTrsHtml=`<div style="background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:8px;padding:8px 10px;flex:1;min-width:0"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#16a34a;text-transform:uppercase;margin-bottom:4px;letter-spacing:.3px">📈 TRS par équipe</div><div style="overflow-x:auto"><svg width="${svgW}" height="${CH}" style="display:block"><line x1="0" y1="${yBase}" x2="${svgW}" y2="${yBase}" stroke="#e2e8f0" stroke-width="1"/>${svgB}${svgL}</svg></div></div>`;
  }
  // ── Chart B : Cadence vs référence — barres verticales SVG ─────────────────
  let chartCadHtml='';
  if(d.sessions_detail&&d.sessions_detail.length>0){
    const sd2=d.sessions_detail;
    const cadRef=Math.round((d.cadence_ref_pcs_min||0)*60);
    const maxCad=Math.max(...sd2.map(s=>s.cadence_h||0),cadRef,1);
    const CH2=280,padT2=20,padB2=56,padL2=4,padR2=4;
    const gH2=CH2-padT2-padB2;
    const n2=sd2.length;
    const WB2=Math.max(22,Math.min(60,Math.floor((420-padL2-padR2-n2*4)/n2)));
    const GP2=5;
    const svgW2=Math.max(200,n2*(WB2+GP2)+padL2+padR2);
    const tY=cadRef>0?padT2+gH2-Math.round(cadRef/maxCad*gH2):null;
    let svgB2='',svgL2='';
    sd2.forEach((s,i)=>{
      const x=padL2+i*(WB2+GP2);
      const v=s.cadence_h||0;
      const bh=Math.max(2,Math.round(v/maxCad*gH2));
      const by=padT2+gH2-bh;
      const col=v>=cadRef?'#16a34a':'#f59e0b';
      svgB2+=`<rect x="${x}" y="${by}" width="${WB2}" height="${bh}" fill="${col}" opacity=".85" rx="2"/>`;
      if(v>0)svgB2+=`<text x="${x+WB2/2}" y="${Math.max(by-3,12)}" text-anchor="middle" font-size="13" font-weight="700" fill="${col}">${v}</text>`;
      const dp=s.date.split('/');
      svgL2+=`<text x="${x+WB2/2}" y="${padT2+gH2+14}" text-anchor="middle" font-size="12" font-weight="600" fill="#374151">${esc((dp[0]||'')+'/'+(dp[1]||''))}</text>`;
      svgL2+=`<text x="${x+WB2/2}" y="${padT2+gH2+27}" text-anchor="middle" font-size="11" fill="#6366f1">${esc((s.pilot||'').slice(0,9))}</text>`;
      svgL2+=`<text x="${x+WB2/2}" y="${padT2+gH2+40}" text-anchor="middle" font-size="11" fill="#94a3b8">${esc((s.poste||'').slice(0,9))}</text>`;
    });
    const yBase2=padT2+gH2;
    const tLine=tY!==null?`<line x1="0" y1="${tY}" x2="${svgW2}" y2="${tY}" stroke="#dc2626" stroke-width="2" stroke-dasharray="6,3"/>`:'';
    const legCad=cadRef>0?`<span style="display:inline-flex;align-items:center;gap:4px;font-size:calc(10px*var(--zf,1));color:#dc2626;font-weight:600;margin-left:8px"><svg width="18" height="6" style="flex-shrink:0"><line x1="0" y1="3" x2="18" y2="3" stroke="#dc2626" stroke-width="2" stroke-dasharray="5,3"/></svg>Cible ${cadRef} éq/h</span>`:'';
    chartCadHtml=`<div style="background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:8px;padding:8px 10px;flex:1;min-width:0"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#f59e0b;text-transform:uppercase;margin-bottom:4px;letter-spacing:.3px;display:flex;align-items:center;flex-wrap:wrap;gap:2px">⚡ Cadence vs référence${legCad}</div><div style="overflow-x:auto"><svg width="${svgW2}" height="${CH2}" style="display:block"><line x1="0" y1="${yBase2}" x2="${svgW2}" y2="${yBase2}" stroke="#e2e8f0" stroke-width="1"/>${svgB2}${tLine}${svgL2}</svg></div></div>`;
  }
  // ── Pareto arrêts — barres horizontales ────────────────────────────────────
  let paretoRjHtml='';
  if(d.stop_pareto&&d.stop_pareto.length){
    const maxMp=d.stop_pareto[0].min,totMp=d.stop_pareto.reduce((a,e)=>a+e.min,0);
    const rows3=d.stop_pareto.map(e=>{
      const pct=Math.round(e.min/maxMp*100),col=STOP_COL[e.cat]||'#94a3b8',pctTot=totMp>0?Math.round(e.min/totMp*100):0;
      return `<div style="display:flex;align-items:center;gap:5px;margin-bottom:4px">
        <div style="flex:0 0 130px;display:flex;align-items:center;gap:4px;min-width:0">
          <div style="width:8px;height:8px;border-radius:2px;background:${col};flex-shrink:0"></div>
          <span style="font-size:calc(10px*var(--zf,1));color:#374151;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(e.type)}</span>
        </div>
        <div style="flex:1;background:#f1f5f9;border-radius:3px;height:14px;min-width:40px">
          <div style="width:${pct}%;background:${col};height:100%;border-radius:3px;opacity:.8"></div>
        </div>
        <div style="flex:0 0 56px;font-size:calc(10px*var(--zf,1));color:#6b7280;text-align:right">${Math.round(e.min)}m <b style="color:#1e293b">${pctTot}%</b></div>
      </div>`;
    }).join('');
    paretoRjHtml=`<div style="background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:8px;padding:8px 10px;flex-shrink:0"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#dc2626;text-transform:uppercase;margin-bottom:6px;letter-spacing:.3px">🛑 Pareto des arrêts</div>${rows3}</div>`;
  }
  // OF list table
  let ofListHtml='';
  const allOfs=[];
  (d.sessions_detail||[]).forEach(s=>{
    (s.of_rows||[]).forEach(r=>{
      const _ofEvts=(s.evt_rows||[]).filter(e=>e.of&&String(e.of)===String(r.of));
      allOfs.push({...r,date:s.date,poste:s.poste,pilot:s.pilot,pilote:s.pilot,_rowType:'prod',_ofEvts:_ofEvts});
    });
  });
  allOfs.reverse();
  window._rjOfs=allOfs;
  if(allOfs.length){
    const ofRows=allOfs.map((r,i)=>`<tr style="border-bottom:1px solid var(--border);font-size:calc(10px*var(--zf,1));cursor:pointer;transition:background .12s" onclick="showRjOfDetail(${i})" title="Voir détail OF">
      <td style="padding:4px 6px;font-weight:700;color:#1d4ed8;text-decoration:underline">${esc(r.of)}</td>
      <td style="padding:4px 6px">${esc(r.date)} · ${esc(r.poste)}</td>
      <td style="padding:4px 6px">${esc(r.pilot)}</td>
      <td style="padding:4px 6px">${esc(r.fibre)}</td>
      <td style="padding:4px 6px;text-align:center">${esc(r.debut)}→${esc(r.fin)}</td>
      <td style="padding:4px 6px;text-align:right;font-weight:700">${esc(r.qte_fab)}</td>
      <td style="padding:4px 6px;text-align:right">${esc(r.equiv)}</td>
      <td style="padding:4px 6px;text-align:right;font-weight:700;color:${r.trs&&parseFloat(r.trs)>=90?'#16a34a':r.trs&&parseFloat(r.trs)>=70?'#f59e0b':'#dc2626'}">${r.trs?parseFloat(r.trs).toFixed(1)+'%':'—'}</td>
      <td style="padding:4px 6px;text-align:right;color:${r.degrade_min>0?'#f59e0b':'#94a3b8'}">${r.degrade_min>0?Math.round(r.degrade_min)+' min':'—'}</td>
    </tr>`).join('');
    ofListHtml=`<div style="background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:8px;padding:8px 10px;margin-top:8px;overflow-x:auto"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#0369a1;text-transform:uppercase;margin-bottom:6px;letter-spacing:.3px">📋 Liste des OF fabriqués</div><table style="width:100%;border-collapse:collapse"><thead><tr style="background:#f1f5f9;font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:#64748b"><th style="padding:4px 6px;text-align:left">OF</th><th style="padding:4px 6px;text-align:left">Poste · Date</th><th style="padding:4px 6px;text-align:left">Pilote</th><th style="padding:4px 6px;text-align:left">Fibre</th><th style="padding:4px 6px;text-align:center">Plage</th><th style="padding:4px 6px;text-align:right">Qté</th><th style="padding:4px 6px;text-align:right">Équiv</th><th style="padding:4px 6px;text-align:right">TRS</th><th style="padding:4px 6px;text-align:right">Dégradé</th></tr></thead><tbody>${ofRows}</tbody></table></div>`;
  }
  // Events list (all non-prod events)
  let eventsListHtml='';
  const allEvts=[];
  (d.sessions_detail||[]).forEach(s=>{
    (s.evt_rows||[]).forEach(r=>{allEvts.push({...r,date:s.date,poste:s.poste,pilote:s.pilot,_rowType:'evt'});});
  });
  allEvts.reverse();
  if(allEvts.length){
    const catCol=t=>{const tl=(t||'').toLowerCase();return tl.includes('nett')?'#f97316':tl.includes('pause')?'#94a3b8':(tl.includes('réunion')||tl.includes('reunion'))?'#8b5cf6':tl.includes('dégrad')?'#ca8a04':'#dc2626';};
    const evtRows=allEvts.map(r=>`<tr style="border-bottom:1px solid var(--border);font-size:calc(10px*var(--zf,1))">
      <td style="padding:4px 6px;font-weight:700;color:${catCol(r.type)}">${esc(r.type||'—')}</td>
      <td style="padding:4px 6px">${esc(r.date||'')} · ${esc(r.poste||'')}</td>
      <td style="padding:4px 6px;color:#0369a1;font-weight:700">${esc(r.of||'—')}</td>
      <td style="padding:4px 6px;text-align:center;white-space:nowrap">${esc(r.debut||'')}→${esc(r.fin||'')}</td>
      <td style="padding:4px 6px;text-align:right;font-weight:700">${esc(r.duree||'—')}</td>
      <td style="padding:4px 6px;color:var(--gray);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(r.comment||'')}">${esc(r.comment||'—')}</td>
    </tr>`).join('');
    eventsListHtml=`<div style="background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:8px;padding:8px 10px;margin-top:8px;overflow-x:auto"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#dc2626;text-transform:uppercase;margin-bottom:6px;letter-spacing:.3px">⏱ Liste des événements (${allEvts.length})</div><table style="width:100%;border-collapse:collapse"><thead><tr style="background:#f1f5f9;font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:#64748b"><th style="padding:4px 6px;text-align:left">Type</th><th style="padding:4px 6px;text-align:left">Poste · Date</th><th style="padding:4px 6px;text-align:left">OF</th><th style="padding:4px 6px;text-align:center">Plage</th><th style="padding:4px 6px;text-align:right">Durée</th><th style="padding:4px 6px;text-align:left">Commentaire</th></tr></thead><tbody>${evtRows}</tbody></table></div>`;
  }
  // Pie chart: fonctionnement vs arrêts
  const fonctMin=d.temps_fonctionnement_min||0;
  const stopMin=d.net_stop_min||0;
  const pieTotal=fonctMin+stopMin;
  let pieHtml='';
  if(pieTotal>0){
    const r=60,cx=65,cy=65;
    const slices=[{v:fonctMin,c:'#16a34a',l:'Prod'},{v:stopMin,c:'#dc2626',l:'Arrêts'}];
    let startA=-Math.PI/2,svgPaths='';
    slices.forEach(sl=>{
      const a=sl.v/pieTotal*2*Math.PI;
      const x1=cx+r*Math.cos(startA),y1=cy+r*Math.sin(startA);
      const x2=cx+r*Math.cos(startA+a),y2=cy+r*Math.sin(startA+a);
      const lg=a>Math.PI?1:0;
      svgPaths+=`<path d="M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${lg},1 ${x2.toFixed(1)},${y2.toFixed(1)} Z" fill="${sl.c}" opacity=".85"/>`;
      startA+=a;
    });
    const legHtml=slices.map(sl=>`<div style="display:flex;align-items:center;gap:5px;font-size:calc(10px*var(--zf,1))"><div style="width:10px;height:10px;border-radius:2px;background:${sl.c};flex-shrink:0"></div>${sl.l}: <b>${Math.round(sl.v)} min</b></div>`).join('');
    pieHtml=`<div style="text-align:center">
      <svg viewBox="0 0 130 130" style="width:130px;height:130px;display:block;margin:0 auto"><circle cx="65" cy="65" r="60" fill="#e2e8f0"/>${svgPaths}</svg>
      <div style="margin-top:6px;display:flex;flex-direction:column;gap:3px;align-items:center">${legHtml}</div>
    </div>`;
  }
  // TRS par jour supprimé
  // Pie compact avec titre + légende
  let pieSmall='';
  if(pieTotal>0){
    const r=34,cx=38,cy=38;let sA=-Math.PI/2,paths='';
    [{v:fonctMin,c:'#16a34a'},{v:stopMin,c:'#dc2626'}].forEach(sl=>{
      const a=sl.v/pieTotal*2*Math.PI;
      const x1=cx+r*Math.cos(sA),y1=cy+r*Math.sin(sA);
      const x2=cx+r*Math.cos(sA+a),y2=cy+r*Math.sin(sA+a);
      paths+=`<path d="M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${a>Math.PI?1:0},1 ${x2.toFixed(1)},${y2.toFixed(1)} Z" fill="${sl.c}" opacity=".85"/>`;
      sA+=a;
    });
    pieSmall=`<div style="display:flex;flex-direction:column;align-items:center;gap:3px;flex-shrink:0">
      <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#374151;white-space:nowrap">Prod / Arrêts</div>
      <svg viewBox="0 0 76 76" style="width:74px;height:74px"><circle cx="38" cy="38" r="34" fill="#e2e8f0"/>${paths}</svg>
      <div style="font-size:calc(10px*var(--zf,1));display:flex;flex-direction:column;gap:2px;align-self:flex-start">
        <div style="display:flex;align-items:center;gap:3px"><div style="width:9px;height:9px;border-radius:2px;background:#16a34a;flex-shrink:0"></div><span style="color:#374151;white-space:nowrap">Prod : <b>${Math.round(fonctMin)} min</b></span></div>
        <div style="display:flex;align-items:center;gap:3px"><div style="width:9px;height:9px;border-radius:2px;background:#dc2626;flex-shrink:0"></div><span style="color:#374151;white-space:nowrap">Arrêts : <b>${Math.round(stopMin)} min</b></span></div>
      </div>
    </div>`;
  }
  // Bandeau postes chargés (auto-load)
  const _rjBanner=document.getElementById('rj-auto-banner');
  if(_rjBanner){
    if(autoLoad&&d.sessions_detail&&d.sessions_detail.length){
      const _lbls=d.sessions_detail.map(s=>esc(s.poste)+' '+esc(s.date)+(s.pilot?' ('+esc(s.pilot)+')':'')).join(' · ');
      _rjBanner.innerHTML=`<span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#0369a1">📋 ${d.sessions_detail.length} dernier${d.sessions_detail.length>1?'s':''} poste${d.sessions_detail.length>1?'s':''} chargé${d.sessions_detail.length>1?'s':''} :</span> <span style="font-size:calc(10px*var(--zf,1));color:#374151">${_lbls}</span>`;
      _rjBanner.style.display='block';
    } else {
      _rjBanner.style.display='none';
    }
  }
  resultEl.innerHTML=`
    <div style="display:flex;gap:10px;height:100%;min-height:0;align-items:stretch">
      <!-- Colonne gauche : KPI synthèse — 320px -->
      <div style="flex:0 0 320px;display:flex;flex-direction:column;gap:5px;overflow:hidden">
        <!-- TRS + pie côte à côte -->
        <div style="display:flex;align-items:center;gap:10px;background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:8px;padding:10px;flex-shrink:0">
          ${pieSmall}
          <div style="flex:1;text-align:center">
            <div style="font-size:calc(36px*var(--zf,1));font-weight:900;color:${trsCol};line-height:1">${d.trs_periode>=0?d.trs_periode.toFixed(1)+'%':'—'}</div>
            <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);font-weight:600;margin-top:3px">TRS période</div>
            <div style="font-size:calc(10px*var(--zf,1));color:#94a3b8;margin-top:2px">${d.nb_jours}j · ${d.nb_sessions} postes · ${d.nb_of} OF</div>
          </div>
        </div>
        <!-- Stats 2×2 -->
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;flex-shrink:0">
          <div class="fp-card" style="padding:6px 8px;text-align:center"><div style="font-size:calc(17px*var(--zf,1));color:#059669;font-weight:900">${Math.round(d.tot_pcs||0)}</div><div style="font-size:calc(10px*var(--zf,1));color:#94a3b8">Pièces</div></div>
          <div class="fp-card" style="padding:6px 8px;text-align:center"><div style="font-size:calc(17px*var(--zf,1));color:#0891b2;font-weight:900">${Math.round(d.tot_equiv||0)}</div><div style="font-size:calc(10px*var(--zf,1));color:#94a3b8">Équiv.</div></div>
          <div class="fp-card" style="padding:6px 8px;text-align:center"><div style="font-size:calc(17px*var(--zf,1));color:#0369a1;font-weight:900">${d.cadence_h||0}</div><div style="font-size:calc(10px*var(--zf,1));color:#94a3b8">Cad./h</div></div>
          <div class="fp-card" style="padding:6px 8px;text-align:center"><div style="font-size:calc(16px*var(--zf,1));color:#0369a1;font-weight:900">${Math.round((d.cadence_ref_pcs_min||0)*100)/100}</div><div style="font-size:calc(10px*var(--zf,1));color:#94a3b8">Réf/min</div></div>
        </div>
        <!-- Lignes info -->
        <div style="display:flex;flex-direction:column;gap:3px;flex-shrink:0">
          ${[['Temps d\'ouverture',Math.round(d.ouverture_min||0)+' min','#374151'],['Temps utile',Math.round(d.temps_utile_min||0)+' min','#059669'],['Fonctionnement',Math.round(d.temps_fonctionnement_min||0)+' min','#16a34a'],['Temps d\'arrêt',Math.round(d.net_stop_min||0)+' min','#dc2626'],['Temps dégradé',Math.round(d.tot_degrade_min||0)+' min','#f59e0b'],['Perte cadence',pertRaw>0?Math.round(pertRaw)+' min de perte':pertRaw<0?Math.abs(Math.round(pertRaw))+' min de gain':'0 min',pertRaw>0?'#dc2626':pertRaw<0?'#16a34a':'#64748b'],['Postes',d.nb_sessions,'#0891b2'],['OF',d.nb_of,'#0891b2'],['Chgt fibre',d.nb_fibre_chg||0,'#8b5cf6'],['Dépass. arrêts prévu',(d.depassement_min||0)>0?Math.round(d.depassement_min)+' min':'✓ OK',(d.depassement_min||0)>0?'#dc2626':'#16a34a']].map(([l,v,c])=>`<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 9px;background:var(--card-bg,#fff);border:1px solid var(--border);border-radius:4px"><span style="font-size:calc(11px*var(--zf,1));color:#64748b">${l}</span><span style="font-size:calc(12px*var(--zf,1));font-weight:700;color:${c}">${v}</span></div>`).join('')}
        </div>
      </div>
      <!-- Colonne droite : graphiques -->
      <div style="flex:1;min-width:0;display:flex;flex-direction:column;gap:6px;overflow-y:auto">
        <div style="display:flex;gap:6px;flex-shrink:0;align-items:flex-start">
          <div style="flex:7;min-width:0;display:flex;flex-direction:column;gap:6px">
            <div style="display:flex;gap:6px">${chartTrsHtml}${chartCadHtml}</div>
          </div>
          <div style="flex:3;min-width:0">${paretoRjHtml}</div>
        </div>
        ${ofListHtml}
        ${eventsListHtml}
      </div>
    </div>
  `;
}
function showRjOfDetail(i){
  const r=window._rjOfs&&window._rjOfs[i];
  if(!r) return;
  const ofEvts=r._ofEvts||[];
  const _h2ms=hm=>{if(!hm)return 0;const[h,m,s]=(hm+':0:0').split(':').map(Number);return((h||0)*3600+(m||0)*60+(s||0))*1000;};
  const _stE=ofEvts.filter(e=>!e.is_degrade).map(e=>({s:_h2ms(e.debut),e:_h2ms(e.fin)})).filter(e=>e.e>e.s);
  const _dgE=ofEvts.filter(e=>e.is_degrade).map(e=>({s:_h2ms(e.debut),e:_h2ms(e.fin)})).filter(e=>e.e>e.s);
  const _mg=evs=>{const ss=[...evs].sort((a,b)=>a.s-b.s);const m=[];ss.forEach(o=>{if(m.length&&o.s<=m[m.length-1].e)m[m.length-1].e=Math.max(m[m.length-1].e,o.e);else m.push({...o});});return m;};
  window._rptDegMin=(dH,fH)=>{const dM=_h2ms(dH),fM=_h2ms(fH);if(fM<=dM)return 0;const mg=_mg(_dgE.map(e=>({s:Math.max(e.s,dM),e:Math.min(e.e,fM)})).filter(e=>e.e>e.s));return Math.round(mg.reduce((a,o)=>a+(o.e-o.s),0)/60000);};
  window._rptNetProd=(dH,fH)=>{const dM=_h2ms(dH),fM=_h2ms(fH);if(fM<=dM)return{netMin:0,stopMin:0};const mg=_mg(_stE.map(e=>({s:Math.max(e.s,dM),e:Math.min(e.e,fM)})).filter(e=>e.e>e.s));const bl=mg.reduce((a,o)=>a+(o.e-o.s),0);return{netMin:Math.round(Math.max(0,fM-dM-bl)/60000),stopMin:Math.round(bl/60000)};};
  _renderAndOpenOfDetail({...r, trs:(r.trs!==''&&r.trs!=null)?parseFloat(r.trs):-1}, ofEvts);
}
function _rjSetBanner(txt){const b=document.getElementById('rj-period-banner');if(b)b.textContent='📅 '+txt;}
function rjLast3(){
  document.getElementById('rj-from').value='';
  document.getElementById('rj-to').value='';
  document.getElementById('rj-pilot').value='';
  document.getElementById('rj-poste').value='';
  _rjSetBanner('Rapport des 3 derniers postes');
  loadRptJour();
}
function rjLast7Days(){
  const _t=new Date(),_f=new Date(_t);
  _f.setDate(_t.getDate()-6);
  document.getElementById('rj-from').value=_f.toISOString().slice(0,10);
  document.getElementById('rj-to').value=_t.toISOString().slice(0,10);
  _rjSetBanner('Rapport des 7 derniers jours');
  calcPeriodReport(false);
}
function resetPeriodReport(){
  document.getElementById('rj-from').value='';
  document.getElementById('rj-to').value='';
  document.getElementById('rj-pilot').value='';
  document.getElementById('rj-poste').value='';
  const r=document.getElementById('rj-result');
  if(r) r.innerHTML='<div style="padding:60px;text-align:center;color:#94a3b8"><div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">📅</div><div style="font-size:calc(14px*var(--zf,1));font-weight:600">Sélectionnez une période puis cliquez sur Calculer</div></div>';
}
function _rptSetLast7(){
  const today=new Date();
  const d7=new Date(today);d7.setDate(today.getDate()-6);
  const fmt=d=>d.toISOString().slice(0,10);
  const f=document.getElementById('rpt-from');const t=document.getElementById('rpt-to');
  if(f)f.value=fmt(d7);if(t)t.value=fmt(today);
}
async function loadRapports(){
  const listEl=document.getElementById('rpt-list');
  if(!listEl) return;
  listEl.innerHTML='<div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Chargement…</div>';
  const sessions=await apiFetch('/api/past_sessions');
  if(!sessions||!sessions.length){
    listEl.innerHTML='<div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun poste disponible</div>';
    return;
  }
  // Filtre par dates (rpt-from / rpt-to en yyyy-mm-dd, sessions en dd/mm/yyyy)
  const fromVal=(document.getElementById('rpt-from')||{}).value||'';
  const toVal=(document.getElementById('rpt-to')||{}).value||'';
  function _dmy2ymd(d){const p=d.split('/');return p.length===3?p[2]+'-'+p[1].padStart(2,'0')+'-'+p[0].padStart(2,'0'):'';}
  const filtered=(fromVal||toVal)?sessions.filter(s=>{const y=_dmy2ymd(s.date);return(!fromVal||y>=fromVal)&&(!toVal||y<=toVal);}):sessions;
  if(!filtered.length){
    listEl.innerHTML='<div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun poste sur cette période</div>';
    return;
  }
  listEl.innerHTML=filtered.map((s,i)=>{
    const trsStr=s.trs>=0?s.trs.toFixed(1)+'%':'—';
    const trsCol=s.trs>=90?'#16a34a':s.trs>=70?'#f59e0b':s.trs>=0?'#dc2626':'#94a3b8';
    return `<div class="rpt-item" id="rpt-item-${i}" onclick="loadSessionReport('${esc(s.date)}','${esc(s.pilot)}','${esc(s.poste)}','rpt-item-${i}')"
      style="padding:10px 14px;border-bottom:1px solid var(--border);cursor:pointer;transition:background .15s">
      <div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:var(--navy)">${esc(s.date)} — ${esc(s.poste)}</div>
      <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-top:2px">${esc(s.pilot||'?')} | ${s.nb_of} OF | Éq. ${s.tot_equiv}</div>
      <div style="font-size:calc(16px*var(--zf,1));font-weight:900;color:${trsCol};margin-top:2px">${trsStr}</div>
    </div>`;
  }).join('');
}

async function loadSessionReport(date,pilot,poste,itemId){
  document.querySelectorAll('.rpt-item').forEach(el=>el.style.background='');
  const sel=document.getElementById(itemId);if(sel) sel.style.background='#eff6ff';
  const detailEl=document.getElementById('rpt-detail');
  if(!detailEl) return;
  detailEl.innerHTML='<div style="padding:40px;text-align:center;color:var(--gray)">Chargement…</div>';
  const d=await apiFetch(`/api/session_report?date=${encodeURIComponent(date)}&pilot=${encodeURIComponent(pilot)}&poste=${encodeURIComponent(poste)}`);
  if(!d){detailEl.innerHTML='<div style="padding:40px;text-align:center;color:#dc2626">Erreur chargement</div>';return;}
  const trsS=d.trs_shift>=0?d.trs_shift:d.trs;
  const trsCol=trsS>=90?'#16a34a':trsS>=70?'#f59e0b':trsS>=0?'#dc2626':'#94a3b8';
  const stopMin=Math.round((d.stop_s||0)/60);
  const prodMin=Math.round(Math.max(0,(d.tot_s||0)-(d.stop_s||0))/60);
  const totalMin=Math.round((d.tot_s||0)/60)+stopMin;
  // Pareto des arrêts
  const stopMap={};
  (d.evt_rows||[]).filter(r=>!r.is_degrade).forEach(r=>{
    const k=r.type||'Inconnu';
    if(!stopMap[k]) stopMap[k]=0;
    const p=r.duree?r.duree.split(':'):[0,0,0];
    stopMap[k]+=(parseInt(p[0]||0)*3600+parseInt(p[1]||0)*60+parseInt(p[2]||0))/60;
  });
  const stopArr=Object.entries(stopMap).sort((a,b)=>b[1]-a[1]);
  const maxStopMin=stopArr.length?stopArr[0][1]:1;
  const paretoHtml=stopArr.length?stopArr.map(([k,v])=>`
    <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
      <div style="font-size:calc(10px*var(--zf,1));width:100px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text)">${esc(k)}</div>
      <div style="flex:1;background:#f1f5f9;border-radius:4px;height:14px;overflow:hidden">
        <div style="height:100%;background:#dc2626;border-radius:4px;width:${Math.round(v/maxStopMin*100)}%;opacity:.8"></div>
      </div>
      <div style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#dc2626;width:36px;text-align:right;flex-shrink:0">${Math.round(v)}mn</div>
    </div>`).join(''):'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>';
  window._rptProdRows = d.prod_rows || [];
  window._rptEvtRows = d.evt_rows || [];
  // Helper: compute net prod and stop overlap for each OF against evt_rows
  const _rptHmsMs=hm=>{if(!hm)return 0;const[h,m,s]=(hm+':0:0').split(':').map(Number);return(h||0)*3600000+(m||0)*60000+(s||0)*1000;};
  const _rptStEvts=(d.evt_rows||[]).filter(e=>!e.is_degrade).map(e=>({s:_rptHmsMs(e.debut),e:_rptHmsMs(e.fin)})).filter(e=>e.e>e.s);
  const _rptDgEvts=(d.evt_rows||[]).filter(e=>e.is_degrade).map(e=>({s:_rptHmsMs(e.debut),e:_rptHmsMs(e.fin)})).filter(e=>e.e>e.s);
  window._rptDegMin=function(debHm,finHm){
    const dMs=_rptHmsMs(debHm),fMs=_rptHmsMs(finHm);
    if(fMs<=dMs) return 0;
    const ov=_rptDgEvts.map(sv=>({s:Math.max(sv.s,dMs),e:Math.min(sv.e,fMs)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    return Math.round(mg.reduce((a,o)=>a+(o.e-o.s),0)/60000);
  };
  window._rptNetProd=function(debHm,finHm){
    const dMs=_rptHmsMs(debHm),fMs=_rptHmsMs(finHm);
    if(fMs<=dMs) return {netMin:0,stopMin:0};
    const ov=_rptStEvts.map(sv=>({s:Math.max(sv.s,dMs),e:Math.min(sv.e,fMs)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    const blocked=mg.reduce((a,o)=>a+(o.e-o.s),0);
    return {netMin:Math.round(Math.max(0,fMs-dMs-blocked)/60000),stopMin:Math.round(blocked/60000)};
  };
  const prodsHtml=(d.prod_rows||[]).map((r,ri)=>{
    const tc=r.trs>=90?'#16a34a':r.trs>=70?'#f59e0b':r.trs>=0?'#dc2626':'#94a3b8';
    const kitStr=(r.kit||'').toLowerCase();
    const kitDisp=kitStr==='oui'?'<span style="color:#16a34a;font-weight:800">✓</span>':'';
    const {netMin,stopMin}=_rptNetProd(r.debut,r.fin);
    const degMinOf=_rptDegMin(r.debut,r.fin);
    const ofDurMin=r.debut&&r.fin?Math.round((_rptHmsMs(r.fin)-_rptHmsMs(r.debut))/60000):0;
    const nbPers=r.nb_pers||'';
    const planMin=Math.round((r.plan_stop_s||0)/60);
    const unplanMin=Math.max(0,stopMin-planMin);
    const td='padding:4px 6px;text-align:center;font-size:calc(11px*var(--zf,1))';
    return `<tr style="border-bottom:1px solid var(--border);cursor:pointer" onclick="showOfDetail(${ri})" title="Voir détail OF">
      <td style="${td};font-weight:700;color:#1e3a8a;text-decoration:underline">${esc(r.of||'')}</td>
      <td style="${td}">${esc(r.taille||'')} ${esc(r.type_prod||'')}</td>
      <td style="${td}">${kitDisp}</td>
      <td style="${td}">${esc(r.qte_fab||'')}</td>
      <td style="${td};color:#0891b2;font-weight:700">${esc(r.equiv||'')}</td>
      <td style="${td};white-space:nowrap">${esc(r.debut||'')} → ${esc(r.fin||'')}</td>
      <td style="${td};color:#94a3b8">${ofDurMin>0?ofDurMin+' min':'—'}</td>
      <td style="${td};color:#16a34a;font-weight:700">${netMin} min</td>
      <td style="${td};color:#16a34a;font-weight:700">${planMin>0?planMin+' min':'—'}</td>
      <td style="${td};color:#dc2626;font-weight:700">${unplanMin>0?unplanMin+' min':'—'}</td>
      <td style="${td};color:${degMinOf>0?'#b45309':'#94a3b8'};font-weight:${degMinOf>0?'700':'400'}">${degMinOf>0?degMinOf+' min':'—'}</td>
      <td style="${td};color:#374151;font-weight:600">${esc(String(nbPers))}</td>
      <td style="${td};color:#0891b2;font-weight:700">${(r.objectif!=null&&r.objectif>=0)?r.objectif:'—'}</td>
      <td style="${td};font-weight:800;color:${tc}">${r.trs>=0?r.trs.toFixed(1)+'%':'—'}</td>
      <td style="padding:4px 6px;font-size:calc(10px*var(--zf,1));color:var(--gray)">${esc(r.comment||'')}</td>
    </tr>`;
  }).join('');
  // Timeline inline builder (ne dépend pas de ST)
  function buildTL(prodRows,evtRows,dateStr,modelDebut,modelFin){
    const W=800,Y=4,H2=28,H=40;
    // Convertir dd/mm/yyyy → base ms
    const parts=dateStr.split('/');
    const baseMs=parts.length===3?new Date(parseInt(parts[2]),parseInt(parts[1])-1,parseInt(parts[0])).getTime():Date.now();
    const hm2ms=hm=>{if(!hm)return null;const[h,m]=(hm+':00').split(':').map(Number);return baseMs+h*3600000+m*60000;};
    const mdMs=hm2ms(modelDebut),mfMs=hm2ms(modelFin);
    // Compute range
    let allMs=[];
    prodRows.forEach(r=>{if(r.debut)allMs.push(hm2ms(r.debut));if(r.fin)allMs.push(hm2ms(r.fin));});
    evtRows.forEach(r=>{if(r.debut)allMs.push(hm2ms(r.debut));if(r.fin)allMs.push(hm2ms(r.fin));});
    const tS=mdMs||Math.min(...allMs.filter(Boolean));
    const tE=mfMs||Math.max(...allMs.filter(Boolean));
    if(!tS||!tE||tE<=tS) return `<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
    const span=tE-tS;
    const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
    let html=`<defs><pattern id="dpat_rpt" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="4" height="8" fill="#16a34a"/><rect x="4" y="0" width="4" height="8" fill="#fef08a"/></pattern></defs>`;
    html+=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
    prodRows.forEach(r=>{
      const t1=hm2ms(r.debut),t2=hm2ms(r.fin);
      if(!t1) return;
      const x1=toX(t1),x2=toX(t2||tE);
      if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="3"/>`;
    });
    evtRows.forEach(r=>{
      const t1=hm2ms(r.debut),t2=hm2ms(r.fin);
      if(!t1) return;
      const x1=toX(t1),x2=toX(t2||t1+1800000);
      if(x2<=x1) return;
      const tl=(r.type||'').toLowerCase();
      const col=r.is_degrade?'url(#dpat_rpt)':(tl.includes('nett')?'#f97316':tl.includes('pause')?'#94a3b8':(tl.includes('réunion')||tl.includes('reunion')||tl.includes('meeting'))?'#8b5cf6':tl.includes('ratt')?'#f59e0b':'#dc2626');
      html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${col}" rx="2" opacity=".75"/>`;
    });
    const fmt=ms=>{const d=new Date(ms);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
    // Hourly tick marks
    let tickT=Math.ceil(tS/3600000)*3600000;
    while(tickT<tE){
      const tx=toX(tickT);
      const hr=new Date(tickT).getHours();
      html+=`<line x1="${tx}" y1="${Y}" x2="${tx}" y2="${Y+H2}" stroke="rgba(0,0,0,.2)" stroke-width="1"/>`;
      if(tx>20&&tx<W-40) html+=`<text x="${tx+2}" y="${Y+H2+9}" font-size="7" fill="#374151">${String(hr).padStart(2,'0')}h</text>`;
      tickT+=3600000;
    }
    html+=`<text x="2" y="${Y+H2+9}" font-size="8" fill="#374151">${fmt(tS)}</text>`;
    html+=`<text x="${W-30}" y="${Y+H2+9}" font-size="8" fill="#374151">${fmt(tE)}</text>`;
    return html;
  }
  // Métriques supplémentaires — priorité aux valeurs Excel (onglet Postes)
  const totQteFab=(d.prod_rows||[]).reduce((s,r)=>s+parseFloat(r.qte_fab||0),0);
  const elapsedEffS=Math.max(1,(d.model_dur_s||0)-(d.planned_ded_s||0));
  const cadenceH=(d.cadence_h!=null&&d.cadence_h>0)?Math.round(d.cadence_h):(elapsedEffS>0?Math.round(totQteFab/elapsedEffS*3600):0);
  const sortedProdF=(d.prod_rows||[]).filter(r=>r.fibre).sort((a,b)=>(a.debut||'').localeCompare(b.debut||''));
  let nbChangFibre=0;for(let i=1;i<sortedProdF.length;i++){if(sortedProdF[i].fibre!==sortedProdF[i-1].fibre)nbChangFibre++;}
  const prodRef=d.prod_ref||200;
  const cadenceRefPcsMin=prodRef/480;
  const ouvertureMin=(d.ouverture_min!=null)?Math.round(d.ouverture_min):Math.round((d.model_dur_s||0)/60);
  // Intervalles fusionnés (arrêts sans chevauchement) — fallback si Excel absent
  const _allEvtIv=(d.evt_rows||[]).filter(e=>!e.is_degrade).map(e=>({s:_rptHmsMs(e.debut),e:_rptHmsMs(e.fin)})).filter(o=>o.e>o.s).sort((a,b)=>a.s-b.s);
  const _merged=[];_allEvtIv.forEach(iv=>{if(_merged.length&&iv.s<=_merged[_merged.length-1].e)_merged[_merged.length-1].e=Math.max(_merged[_merged.length-1].e,iv.e);else _merged.push({s:iv.s,e:iv.e});});
  const netStopMin=(d.arret_min!=null)?Math.round(d.arret_min):Math.round(_merged.reduce((a,o)=>a+(o.e-o.s),0)/60000);
  const tempsFonctionnement=(d.fonct_min!=null)?Math.round(d.fonct_min):Math.max(0,ouvertureMin-netStopMin);
  // Temps utile depuis Excel ou calcul budget
  const budgetData=d.budget_data||{};
  const arretsPrevu=Object.values(budgetData).reduce((a,b)=>a+Math.min(b.budget_min||0,b.used_min||0),0);
  const tempsUtile=(d.utile_min!=null)?Math.round(d.utile_min):Math.max(0,ouvertureMin-Math.round(arretsPrevu));
  // Perte cadence: utilise la valeur serveur (Option B, tient compte dégradé + nb_pers)
  const perteCadenceRaw=Math.round(d.perte_cadence_min||0);
  const perteCadenceHtml=perteCadenceRaw<0?`<span style="color:#16a34a;font-weight:800">${Math.abs(perteCadenceRaw)} min de gain</span>`:perteCadenceRaw>0?`<span style="color:#dc2626;font-weight:800">${perteCadenceRaw} min de perte</span>`:`<span style="color:#64748b">0 min</span>`;
  const degMin=(d.degrade_min!=null)?Math.round(d.degrade_min):Math.round((d.degrade_s||0)/60);
  const tlDebut=d.actual_debut||d.model_debut;
  const tlFin=d.actual_fin||d.model_fin;
  const tlContent=buildTL(d.prod_rows||[],d.evt_rows||[],date,tlDebut,tlFin);
  const plageStr=(tlDebut&&tlFin)?(' · '+esc(tlDebut)+' → '+esc(tlFin)):'';
  // ── Panneau gauche : passe en mode KPI ──
  const leftList=document.getElementById('rpt-left-list');
  const leftKpi=document.getElementById('rpt-left-kpi');
  const leftExe=document.getElementById('rpt-left-exe');
  if(leftList&&leftKpi&&leftExe){
    leftList.style.display='none';
    leftExe.style.width='246px';
    leftKpi.style.cssText='display:flex;flex-direction:row;flex:1;overflow:hidden';
    leftKpi.innerHTML=`
      <!-- Fine bande retour (gauche) -->
      <div onclick="rptBackToList()" title="Retour à la liste"
        style="width:28px;background:#e5e7eb;border-right:2px solid #d1d5db;cursor:pointer;flex-shrink:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;transition:background .15s;animation:rpt-arr .9s step-start infinite"
        onmouseenter="this.style.background='#f9fafb';this.style.animation='none'" onmouseleave="this.style.background='#e5e7eb';this.style.animation='rpt-arr .9s step-start infinite'">
        <span style="font-size:20px;color:#374151;user-select:none;line-height:1;font-weight:900">❮</span>
      </div>
      <!-- Contenu KPI -->
      <div style="flex:1;overflow-y:auto;display:flex;flex-direction:column">
      <div style="background:var(--navy);color:#fff;padding:10px 12px;flex-shrink:0">
        <div style="font-size:calc(12px*var(--zf,1));font-weight:800;opacity:.9">${esc(poste)}${((d.actual_debut||d.model_debut)&&(d.actual_fin||d.model_fin))?' — '+(d.actual_debut||d.model_debut)+' → '+(d.actual_fin||d.model_fin):''}</div>
        <div style="font-size:calc(10px*var(--zf,1));opacity:.75;margin-top:2px">${esc(pilot)} · ${esc(date)}</div>
        <div style="font-size:calc(9px*var(--zf,1));opacity:.65;margin-top:6px;font-weight:600;text-transform:uppercase;letter-spacing:.05em">TRS :</div>
        <div style="font-size:calc(36px*var(--zf,1));font-weight:900;color:${trsCol};line-height:1.1;text-shadow:0 1px 4px rgba(0,0,0,.3)">${trsS>=0?trsS.toFixed(1)+'%':'—'}</div>
      </div>
      ${d.is_live?`<style>@keyframes live-banner{0%,49%{background:#22c55e;color:#000}50%,100%{background:#fff;color:#15803d}}</style><div style="font-size:calc(20px*var(--zf,1));font-weight:900;text-align:center;padding:10px 8px;letter-spacing:.08em;animation:live-banner 1.2s step-start infinite;flex-shrink:0;border-bottom:2px solid #22c55e">▶ POSTE EN COURS</div>`:''}
      <div style="padding:6px 8px;display:flex;flex-direction:column;gap:5px">
        <div style="text-align:center">
          <svg id="rpt-pie" viewBox="0 0 130 130" style="width:150px;height:150px;display:block;margin:0 auto"></svg>
          ${((d.actual_debut||d.model_debut)&&(d.actual_fin||d.model_fin))?`<div style="font-size:calc(9px*var(--zf,1));color:var(--gray);margin-top:3px;font-weight:600">${esc(d.actual_debut||d.model_debut)} → ${esc(d.actual_fin||d.model_fin)}</div>`:''}
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
          <div class="fp-card" style="padding:6px;text-align:center"><div class="fp-big" style="font-size:calc(16px*var(--zf,1));color:#059669;font-weight:900">${Math.round(totQteFab)}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Pièces</div></div>
          <div class="fp-card" style="padding:6px;text-align:center"><div class="fp-big" style="font-size:calc(16px*var(--zf,1));color:#0891b2;font-weight:900">${Math.round(d.tot_equiv||0)}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Équiv.</div></div>
        </div>
        <div style="display:flex;align-items:center;gap:5px">
          <div class="fp-card" style="padding:6px;text-align:center;flex:1"><div class="fp-big" style="font-size:calc(16px*var(--zf,1));color:#0369a1;font-weight:900">${d.is_live?'—':cadenceH}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Cad./h</div></div>
          <div style="text-align:center;flex-shrink:0"><div style="font-size:calc(13px*var(--zf,1));font-weight:800;color:#0369a1">${Math.round(cadenceRefPcsMin*10)/10}</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">réf pcs/min</div></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
          <div class="fp-card" style="padding:6px;text-align:center"><div class="fp-big" style="font-size:calc(16px*var(--zf,1));color:#7c3aed;font-weight:900">${d.nb_of||0}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Nb OF</div></div>
          <div class="fp-card" style="padding:6px;text-align:center"><div class="fp-big" style="font-size:calc(16px*var(--zf,1));color:#8b5cf6;font-weight:900">${nbChangFibre}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Chg. fibre</div></div>
        </div>
        <div style="display:flex;flex-direction:column;gap:4px">
          <div class="fp-card" style="padding:5px 6px"><div class="fp-big" style="font-size:calc(11px*var(--zf,1));color:#374151">${d.is_live?'—':ouvertureMin+' min'}</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">Temps d\'ouverture</div></div>
          <div class="fp-card" style="padding:5px 6px"><div class="fp-big" style="font-size:calc(11px*var(--zf,1));color:#059669">${d.is_live?'—':tempsUtile+' min'}</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">Temps utile</div></div>
          <div class="fp-card" style="padding:5px 6px"><div class="fp-big" style="font-size:calc(11px*var(--zf,1));color:#16a34a">${d.is_live?'—':tempsFonctionnement+' min'}</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">Temps de fonctionnement</div></div>
          <div class="fp-card" style="padding:5px 6px"><div class="fp-big" style="font-size:calc(11px*var(--zf,1));color:#dc2626">${netStopMin} min</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">Temps en arrêt</div></div>
          ${degMin>0?`<div class="fp-card" style="padding:5px 6px;border-left:3px solid #ca8a04"><div class="fp-big" style="font-size:calc(11px*var(--zf,1));color:#b45309">${degMin} min</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">Temps en mode dégradé</div></div>`:''}
          <div class="fp-card" style="padding:5px 6px"><div class="fp-big" style="font-size:calc(11px*var(--zf,1))">${d.is_live?'<span style="color:#94a3b8">—</span>':perteCadenceHtml}</div><div class="fp-lbl" style="font-size:calc(8px*var(--zf,1))">Perte cadence</div></div>
        </div>
      </div>
      </div>`;
  }
  // ── Panneau droit : timeline + tables (pleine largeur) ──
  detailEl.innerHTML=`
    <!-- Timeline -->
    <div style="padding:5px 12px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:3px">Timeline${plageStr}</div>
      <svg viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">${tlContent}</svg>
      <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Rattrapage</span><span><i style="background:#38bdf8"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
    </div>
    <!-- Productions (pleine largeur) -->
    <div style="padding:8px 12px;border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:var(--gray);margin-bottom:6px">Productions</div>
      <table style="width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1))">
        <thead><tr style="background:#f8fafc;border-bottom:1px solid var(--border)">
          <th style="padding:4px 6px;text-align:center">OF</th><th style="padding:4px 6px;text-align:center">Taille</th>
          <th style="padding:4px 6px;text-align:center">Lots×2</th><th style="padding:4px 6px;text-align:center">Qté</th><th style="padding:4px 6px;text-align:center">Éq.</th>
          <th style="padding:4px 6px;text-align:center">Heures</th><th style="padding:4px 6px;text-align:center;color:#94a3b8">Durée OF</th><th style="padding:4px 6px;text-align:center;color:#16a34a">Durée prod</th><th style="padding:4px 6px;text-align:center;color:#16a34a">Arrêts prévus</th><th style="padding:4px 6px;text-align:center;color:#dc2626">Arrêts non prévus</th><th style="padding:4px 6px;text-align:center;color:#b45309">Dégradé</th><th style="padding:4px 6px;text-align:center">Nb pers</th><th style="padding:4px 6px;text-align:center;color:#0891b2">Objectif</th><th style="padding:4px 6px;text-align:center">TRS</th><th style="padding:4px 6px;text-align:left">Comm.</th>
        </tr></thead>
        <tbody>${prodsHtml||'<tr><td colspan="15" style="padding:8px;text-align:center;color:var(--gray)">Aucune production</td></tr>'}</tbody>
      </table>
    </div>
    <!-- Pareto + Arrêts côte à côte -->
    <div style="flex:1;overflow-y:auto;padding:8px 12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div style="display:flex;flex-direction:column;gap:8px">
        <div class="card" style="padding:10px">
          <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:var(--gray);margin-bottom:8px">Pareto des arrêts</div>
          ${paretoHtml||'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>'}
        </div>
        <div class="card" style="padding:10px">
          <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#92400e;margin-bottom:8px">⏱ Arrêts prévus</div>
          ${(()=>{
            const bd=d.budget_data||{};
            const bKeys=['clean_short_min','clean_long_min','clean_grand_min','meeting_tol_min','pause_min'];
            const rows=bKeys.map(bk=>{
              const b=bd[bk];if(!b||b.budget_min<=0)return '';
              const pct=Math.min(100,Math.round(b.used_min/b.budget_min*100));
              const col=b.used_min>b.budget_min?'#dc2626':b.used_min/b.budget_min>=0.8?'#d97706':'#16a34a';
              return '<div style="margin-bottom:7px">'
                +'<div style="display:flex;justify-content:space-between;font-size:calc(11px*var(--zf,1));margin-bottom:3px">'
                +'<span style="color:var(--text)">'+esc(b.label)+'</span>'
                +'<span style="font-weight:700;color:'+col+'">'+Math.round(b.used_min)+'/'+Math.round(b.budget_min)+' min</span>'
                +'</div>'
                +'<div style="background:#f1f5f9;border-radius:4px;height:12px;overflow:hidden">'
                +'<div style="height:100%;background:'+col+';border-radius:4px;width:'+pct+'%;opacity:.85"></div>'
                +'</div></div>';
            }).filter(Boolean).join('');
            return rows||'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun budget configuré</div>';
          })()}
        </div>
      </div>
      <div class="card" style="padding:10px;display:flex;flex-direction:column;gap:6px">
        <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:var(--gray);margin-bottom:6px">Détail arrêts</div>
        ${(d.evt_rows||[]).length?`<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:calc(10px*var(--zf,1))">
          <thead><tr style="background:#f8fafc;border-bottom:1px solid var(--border)">
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray);white-space:nowrap">Arrêt</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">OF</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">Type</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray);white-space:nowrap">Plage</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">Durée</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">Commentaire</th>
          </tr></thead>
          <tbody>${(d.evt_rows||[]).filter(r=>!r.is_degrade).map((r,ri)=>`<tr style="border-bottom:1px solid var(--border);cursor:pointer;transition:background .12s" onclick="showRptEvtDetail(${ri})" title="Voir détail">
            <td style="padding:4px 5px;font-weight:700;white-space:nowrap;max-width:90px;overflow:hidden;text-overflow:ellipsis;color:#dc2626;text-decoration:underline">${esc(r.type||'')}</td>
            <td style="padding:4px 5px;color:#0369a1;font-weight:700">${esc(r.of||'—')}</td>
            <td style="padding:4px 5px;color:var(--text)">${esc(r.type_prod||'—')}</td>
            <td style="padding:4px 5px;white-space:nowrap;color:var(--gray)">${esc(r.debut||'')} → ${esc(r.fin||'')}</td>
            <td style="padding:4px 5px;font-weight:700;white-space:nowrap">${esc(r.duree||'')}</td>
            <td style="padding:4px 5px;color:var(--gray);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(r.comment||'')}">${esc(r.comment||'—')}</td>
          </tr>`).join('')}</tbody>
        </table></div>`:'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>'}
      </div>
    </div>`;
  // Dessiner gauge et pie (éléments maintenant dans le DOM)
  drawPie('rpt-pie',[
    {label:'Prod',value:tempsFonctionnement,color:'#16a34a'},
    {label:'Arrêts',value:netStopMin,color:'#dc2626'}
  ],{fCenter:16,fSub:10,fLeg:10,legY:118});
}

function rptBackToList(){
  const leftExe=document.getElementById('rpt-left-exe');
  const leftList=document.getElementById('rpt-left-list');
  const leftKpi=document.getElementById('rpt-left-kpi');
  const detailEl=document.getElementById('rpt-detail');
  if(leftExe) leftExe.style.width='280px';
  if(leftList) leftList.style.display='flex';
  if(leftKpi){leftKpi.style.display='none';leftKpi.innerHTML='';}
  if(detailEl) detailEl.innerHTML='<div style="padding:24px;color:var(--gray);text-align:center">Sélectionnez un poste</div>';
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

let _cfgArretsPrevus = {};

async function loadCfg(){
  const d=await apiFetch('/api/config');
  if(!d) return;
  _cfgPwds=d.pilot_passwords||{};
  _cfgModels=d.modeles_horaires||[];
  _cfgModelsBase=d.modeles_horaires_base||JSON.parse(JSON.stringify(_cfgModels));
  _cfgArretsPrevus=d.arrets_prevus||{};
  const prEl=document.getElementById('cfg-pr');
  if(prEl) prEl.value=d.prod_ref||200;
  // Show current db path
  const dbEl=document.getElementById('cfg-db-path');
  const dbSt=document.getElementById('cfg-db-status');
  if(dbEl&&d.db_path) dbEl.value=d.db_path;
  if(dbSt&&d.db_name){dbSt.textContent='Fichier actuel : '+d.db_name;dbSt.style.color='var(--green)';}
  // Arrêts prévus
  const apMap={'ap-clean-short':'clean_short_min','ap-clean-long':'clean_long_min','ap-clean-grand':'clean_grand_min','ap-meeting':'meeting_tol_min','ap-pause':'pause_min'};
  Object.entries(apMap).forEach(([elId,key])=>{const el=document.getElementById(elId);if(el)el.value=(_cfgArretsPrevus[key]||0);});
  if(Array.isArray(d.degrade_motifs)){_degradeListLocal=d.degrade_motifs.slice();renderDegradeList(_degradeListLocal);}
  if(d.pers_pct_map){_persPctMapLocal=d.pers_pct_map;renderPersPctTable();}
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
  updateAccModelInfo();
}

function updateAccModelInfo(){
  const el=document.getElementById('acc-model-info');
  if(!el) return;
  const poste=ST&&ST.poste;
  if(!poste){el.style.display='none';return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mEff=_cfgModels.find(m=>m.nom===poste);
  const mBase=_cfgModelsBase.find(m=>m.nom===poste);
  const jEff=mEff&&mEff.jours&&mEff.jours[dk]||{};
  const jBase=mBase&&mBase.jours&&mBase.jours[dk]||{};
  if(!jEff.debut&&!jEff.fin){el.style.display='none';return;}
  const isOverridden=(jEff.debut!==jBase.debut)||(jEff.fin!==jBase.fin);
  el.style.display='block';
  if(isOverridden){
    el.innerHTML=`⏰ <b>${esc(poste)}</b> : <b style="color:#d97706">${esc(jEff.debut)} → ${esc(jEff.fin)}</b> <span style="background:#fef3c7;color:#92400e;font-size:calc(10px*var(--zf,1));padding:1px 5px;border-radius:4px;font-weight:700">⚠ temporaire</span>`;
  } else {
    el.innerHTML=`⏰ <b>${esc(poste)}</b> : ${esc(jEff.debut)} → ${esc(jEff.fin)}`;
  }
}

async function saveArretsPrevus(){
  const clean_short=parseFloat(document.getElementById('ap-clean-short')?.value||0)||0;
  const clean_long=parseFloat(document.getElementById('ap-clean-long')?.value||0)||0;
  const clean_grand=parseFloat(document.getElementById('ap-clean-grand')?.value||0)||0;
  const meeting=parseFloat(document.getElementById('ap-meeting')?.value||0)||0;
  const pause_m=parseFloat(document.getElementById('ap-pause')?.value||0)||0;
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,clean_short_min:clean_short,clean_long_min:clean_long,clean_grand_min:clean_grand,meeting_tol_min:meeting,pause_min:pause_m})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Arrêts prévus enregistrés','ok');await loadCfg();}
  else toast(d?.error||'Erreur','err');
}

function renderPwdList(){
  const c=document.getElementById('pwd-list');
  if(!c) return;
  c.innerHTML=Object.entries(_cfgPwds).map(([nm,pw],i)=>`
    <div class="pr" draggable="true" ondragstart="_pdDS(${i})" ondragover="_pdDO(event)" ondrop="_pdDrop(${i})" style="cursor:default">
      <span style="cursor:grab;color:#94a3b8;font-size:16px;padding:0 2px;user-select:none;flex-shrink:0" title="Déplacer">⠿</span>
      <div class="pn">${esc(nm)}</div>
      <input type="password" id="pwi-${esc(nm)}" value="${esc(String(pw))}" data-n="${esc(nm)}">
      <button class="btn-eye" onclick="toggleEye('pwi-${esc(nm)}')">👁</button>
      <button class="btn btn-sec" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="rmPilot('${esc(nm)}')">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:3px">Aucun pilote configuré</div>';
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

async function changeAdminPw(){
  const oldPw=document.getElementById('adm-old-pw').value;
  const newPw=document.getElementById('adm-new-pw').value;
  const confirmPw=document.getElementById('adm-confirm-pw').value;
  const msgEl=document.getElementById('adm-pw-msg');
  if(!newPw){if(msgEl){msgEl.textContent='Nouveau MDP vide';msgEl.style.color='#dc2626';}return;}
  if(newPw!==confirmPw){if(msgEl){msgEl.textContent='Les MDP ne correspondent pas';msgEl.style.color='#dc2626';}return;}
  const r=await fetch('/api/change_admin_pw',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({old_pw:oldPw,new_pw:newPw})});
  const d=r?await r.json():{};
  if(d&&d.ok){
    _adminPw=newPw;
    document.getElementById('adm-old-pw').value='';document.getElementById('adm-new-pw').value='';document.getElementById('adm-confirm-pw').value='';
    if(msgEl){msgEl.textContent='✓ MDP changé et enregistré dans Excel';msgEl.style.color='#16a34a';}
  } else {
    if(msgEl){msgEl.textContent=d&&d.error||'Erreur';msgEl.style.color='#dc2626';}
  }
}

// ── Listes Paramètres : co-pilotes, tailles, fibres, équivalences ──
let _copilotesList=[], _taillesList=[], _fibresList=[], _equivList=[];

// Drag handlers co-pilotes
function _cpDS(i){_dIdx=i;}function _cpDO(e){e.preventDefault();}
function _cpDrop(i){if(_dIdx===null||_dIdx===i)return;const m=_copilotesList.splice(_dIdx,1)[0];_copilotesList.splice(i,0,m);_dIdx=null;renderCopilotesList();}
// Drag handlers tailles
function _tlDS(i){_dIdx=i;}function _tlDO(e){e.preventDefault();}
function _tlDrop(i){if(_dIdx===null||_dIdx===i)return;const m=_taillesList.splice(_dIdx,1)[0];_taillesList.splice(i,0,m);_dIdx=null;renderTaillesList();}
// Drag handlers fibres
function _fbDS(i){_dIdx=i;}function _fbDO(e){e.preventDefault();}
function _fbDrop(i){if(_dIdx===null||_dIdx===i)return;const m=_fibresList.splice(_dIdx,1)[0];_fibresList.splice(i,0,m);_dIdx=null;renderFibresList();}
// Drag handlers equiv
function _eqDS(i){_dIdx=i;}function _eqDO(e){e.preventDefault();}
function _eqDrop(i){if(_dIdx===null||_dIdx===i)return;const m=_equivList.splice(_dIdx,1)[0];_equivList.splice(i,0,m);_dIdx=null;renderEquivList();}

const _listRowStyle='display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1));cursor:default';
const _dragHandle='<span style="cursor:grab;color:#94a3b8;font-size:16px;padding:0 2px;user-select:none" title="Déplacer">⠿</span>';

function renderCopilotesList(){
  const c=document.getElementById('copilotes-list-ui');if(!c)return;
  c.innerHTML=_copilotesList.map((v,i)=>`<div draggable="true" ondragstart="_cpDS(${i})" ondragover="_cpDO(event)" ondrop="_cpDrop(${i})" style="${_listRowStyle}">${_dragHandle}<span style="flex:1;font-weight:600">${esc(v)}</span><button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="_copilotesList.splice(${i},1);renderCopilotesList()">✕</button></div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucun co-pilote</div>';
}
function addCopilote(){const v=document.getElementById('cp-new-name').value.trim();if(!v){toast('Nom requis','err');return;}_copilotesList.push(v);document.getElementById('cp-new-name').value='';renderCopilotesList();}
async function saveCopilotes(){
  const r=await fetch('/api/save_list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,list_type:'copilotes',items:_copilotesList})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Co-pilotes enregistrés dans Excel','ok');}else toast(d&&d.error||'Erreur','err');
}

function renderTaillesList(){
  const c=document.getElementById('tailles-list-ui');if(!c)return;
  c.innerHTML=_taillesList.map((v,i)=>`<div draggable="true" ondragstart="_tlDS(${i})" ondragover="_tlDO(event)" ondrop="_tlDrop(${i})" style="${_listRowStyle}">${_dragHandle}<span style="flex:1;font-weight:600">${esc(v)}</span><button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="_taillesList.splice(${i},1);renderTaillesList()">✕</button></div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucune taille</div>';
}
function addTaille(){const v=document.getElementById('tl-new-val').value.trim();if(!v){toast('Valeur requise','err');return;}_taillesList.push(v);document.getElementById('tl-new-val').value='';renderTaillesList();}
async function saveTailles(){
  const r=await fetch('/api/save_list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,list_type:'tailles',items:_taillesList})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Tailles enregistrées dans Excel','ok');}else toast(d&&d.error||'Erreur','err');
}

function renderFibresList(){
  const c=document.getElementById('fibres-list-ui');if(!c)return;
  c.innerHTML=_fibresList.map((v,i)=>`<div draggable="true" ondragstart="_fbDS(${i})" ondragover="_fbDO(event)" ondrop="_fbDrop(${i})" style="${_listRowStyle}">${_dragHandle}<span style="flex:1;font-weight:600">${esc(v)}</span><button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="_fibresList.splice(${i},1);renderFibresList()">✕</button></div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucune fibre</div>';
}
function addFibre(){const v=document.getElementById('fb-new-val').value.trim();if(!v){toast('Valeur requise','err');return;}_fibresList.push(v);document.getElementById('fb-new-val').value='';renderFibresList();}
async function saveFibres(){
  const r=await fetch('/api/save_list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,list_type:'fibres',items:_fibresList})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Fibres enregistrées dans Excel','ok');}else toast(d&&d.error||'Erreur','err');
}

function renderEquivList(){
  const c=document.getElementById('equiv-list-ui');if(!c)return;
  c.innerHTML=_equivList.map((it,i)=>`<div draggable="true" ondragstart="_eqDS(${i})" ondragover="_eqDO(event)" ondrop="_eqDrop(${i})" style="${_listRowStyle}">${_dragHandle}<input value="${esc(it.type||'')}" onchange="_equivList[${i}].type=this.value" style="flex:2;padding:3px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1));font-weight:600"><input value="${esc(it.coeff||'')}" onchange="_equivList[${i}].coeff=this.value" style="flex:1;padding:3px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1));text-align:center" placeholder="Coeff"><button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="_equivList.splice(${i},1);renderEquivList()">✕</button></div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucun coefficient</div>';
}
function addEquiv(){
  const t=document.getElementById('eq-new-type').value.trim(),c=document.getElementById('eq-new-coeff').value.trim();
  if(!t){toast('Type requis','err');return;}
  _equivList.push({type:t,coeff:c});
  document.getElementById('eq-new-type').value='';document.getElementById('eq-new-coeff').value='';
  renderEquivList();
}
async function saveEquiv(){
  const r=await fetch('/api/save_list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,list_type:'equiv',items:_equivList})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Coefficients enregistrés dans Excel','ok');}else toast(d&&d.error||'Erreur','err');
}

const DAYS=[{k:'lun',l:'Lun'},{k:'mar',l:'Mar'},{k:'mer',l:'Mer'},{k:'jeu',l:'Jeu'},{k:'ven',l:'Ven'},{k:'sam',l:'Sam'},{k:'dim',l:'Dim'}];
function renderModelList(){
  const c=document.getElementById('models-list');
  if(!c) return;
  c.innerHTML=_cfgModelsBase.map((m,mi)=>{
    const j=m.jours||{};
    return `<div class="model-card">
      <div class="mch">
        <input value="${esc(m.nom||'Poste '+(mi+1))}" onchange="_cfgModelsBase[${mi}].nom=this.value" placeholder="Nom du poste">
        <button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="_cfgModelsBase.splice(${mi},1);renderModelList()">✕</button>
      </div>
      <div class="day-grid">${DAYS.map(d=>{const dc=j[d.k]||{};
        return `<div class="day-box"><div class="day-lbl">${d.l}</div>
          <input type="time" onchange="setDay(${mi},'${d.k}','debut',this.value)" value="${dc.debut||'05:00'}" style="margin-bottom:2px">
          <input type="time" onchange="setDay(${mi},'${d.k}','fin',this.value)" value="${dc.fin||'13:00'}">
        </div>`;}).join('')}
      </div>
    </div>`;
  }).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1))">Aucun modèle horaire</div>';
}

function setDay(mi,day,field,val){if(!_cfgModelsBase[mi])return;if(!_cfgModelsBase[mi].jours)_cfgModelsBase[mi].jours={};if(!_cfgModelsBase[mi].jours[day])_cfgModelsBase[mi].jours[day]={};_cfgModelsBase[mi].jours[day][field]=val;}
function addModel(){_cfgModelsBase.push({nom:'Nouveau poste',jours:{lun:{debut:'05:00',fin:'13:00'},mar:{debut:'05:00',fin:'13:00'},mer:{debut:'05:00',fin:'13:00'},jeu:{debut:'05:00',fin:'13:00'},ven:{debut:'05:00',fin:'13:00'},sam:{debut:'',fin:''},dim:{debut:'',fin:''}}});renderModelList();}
async function saveModels(){
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,modeles_horaires:_cfgModelsBase})});
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
    if(st){st.textContent='✓ Dashboard prêt — cliquez pour ouvrir';st.style.color='var(--green)';}
    toast('Dashboard généré !','ok');
    // Ouvrir via Flask (même origine, pas de CORS)
    window.open('http://127.0.0.1:5001/dashboard','_blank');
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
function showFibre(name){
  if(!name)return;
  let m=document.getElementById('m-fibre-info');
  if(!m){
    m=document.createElement('div');
    m.id='m-fibre-info';
    m.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:900;display:flex;align-items:center;justify-content:center';
    m.innerHTML='<div style="background:#fff;border-radius:14px;padding:28px 32px;min-width:260px;max-width:420px;box-shadow:0 20px 60px rgba(0,0,0,.3);text-align:center">'+
      '<div style="font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#6366f1;letter-spacing:.6px;margin-bottom:8px">Fibre</div>'+
      '<div id="m-fibre-name" style="font-size:calc(18px*var(--zf,1));font-weight:800;color:#1e293b;word-break:break-all;margin-bottom:18px"></div>'+
      '<button onclick="document.getElementById(\'m-fibre-info\').style.display=\'none\'" style="background:#6366f1;color:#fff;border:none;border-radius:8px;padding:8px 24px;font-size:calc(13px*var(--zf,1));font-weight:700;cursor:pointer">Fermer</button>'+
      '</div>';
    m.onclick=e=>{if(e.target===m)m.style.display='none';};
    document.body.appendChild(m);
  }
  document.getElementById('m-fibre-name').textContent=name;
  m.style.display='flex';
}
function fmtDur2(s){if(!s||s<0)return'0:00';const m=Math.floor(s/60),sec=Math.floor(s%60);return m+':'+String(sec).padStart(2,'0');}
function fmtTRSv(v){return(v===null||v===undefined||isNaN(v)||v<0)?'--%':parseFloat(v).toFixed(1)+'%';}
function fmtDur(s){if(!s||s<0)return'00:00:00';const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sec=Math.floor(s%60);return[h,m,sec].map(x=>String(x).padStart(2,'0')).join(':');}
function fmtD2(s){if(!s||s<0)return'0 min';const h=Math.floor(s/3600),m=Math.floor((s%3600)/60);return h?h+'h'+String(m).padStart(2,'0'):m+' min';}
function fmtDurMS(s){s=Math.round(s||0);const m=Math.floor(s/60),sec=s%60;return m>0?m+' min'+(sec?' '+sec+'s':''):sec+'s';}
function fmtDurShort(s){s=Math.round(s||0);const m=Math.floor(s/60),sec=s%60;return m>0?m+'m'+(sec?' '+String(sec).padStart(2,'0')+'s':''):sec+'s';}
function fmtTRS(v){return(v===null||v===undefined||isNaN(v))?'--%':parseFloat(v).toFixed(1)+'%';}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function toast(msg,type,dur){
  let t=document.getElementById('_toast');
  if(!t){t=document.createElement('div');t.id='_toast';t.style.cssText='position:fixed;bottom:16px;right:16px;padding:8px 14px;border-radius:7px;font-size:calc(13px*var(--zf,1));font-weight:600;z-index:999;transition:opacity .3s;box-shadow:0 4px 12px rgba(0,0,0,.18)';document.body.appendChild(t);}
  t.textContent=msg;t.style.background=type==='ok'?'#16a34a':type==='warn'?'#d97706':'#dc2626';t.style.color='#fff';t.style.opacity='1';
  clearTimeout(t._to);t._to=setTimeout(()=>t.style.opacity='0',dur||3000);
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
