#!/usr/bin/env python3
"""KPI ORC v6.0 — Flask + pywebview"""

import sys
import os
import json
import threading
import time
from datetime import datetime, date

import webview
from flask import Flask, jsonify, request
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────────────────────────────────────
VERSION = "6.0"
PORT = 5000

# Colonnes unifiées — 1 ligne par déclaration
EVENEMENTS_HEADERS = [
    "Type", "OF", "Date", "Poste", "Pilote", "Co-Pilote", "Nb Personnes",
    "Taille", "Code Produit", "Type Produit", "Poids Garnissage", "Fibre",
    "OF Taie", "Traca Fibre", "Ref Taie", "Kit",
    "Heure Debut", "Heure Fin", "Duree",
    "Qte Fabriquee", "Qte Emballee", "Equivalence",
    "Cadence/heure", "Cadence/h/pers",
    "Qte Initiale Taie", "Nb Taie 2nd Choix",
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
    "Temps Interposte", "Commentaire", "Prévu/Hors TRS",
]

IDX = {h: i for i, h in enumerate(EVENEMENTS_HEADERS)}

# ──────────────────────────────────────────────────────────────────────────────
# ÉTAT GLOBAL
# ──────────────────────────────────────────────────────────────────────────────
_cfg = {"db_path": "", "supervisor_pw": "admin", "html_output": ""}

_session = {
    "logged_in": False,
    "poste": "", "pilote": "", "copilote": "",
    "nb_personnes": 1,
    "horaire_nom": "", "horaire_debut": "08:00", "horaire_fin": "16:00",
    "last_heure_fin": None,
}

_lists = {
    "pilotes": [], "pilotes_pw": [], "copilotes": [],
    "tailles": [], "types_produit": [], "equiv_coef": {},
    "postes": [], "nb_personnes": [],
    "prod_ref": 0.0, "fibres": [], "arrets": [], "horaires": {},
}

_app_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_app_dir, "kpi_orc_config.json")

# ──────────────────────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                _cfg.update(json.load(f))
        except Exception:
            pass

def save_config():
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(_cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ──────────────────────────────────────────────────────────────────────────────
# LISTES (depuis onglet Listes du fichier Excel)
# ──────────────────────────────────────────────────────────────────────────────
def load_lists():
    path = _cfg.get("db_path", "")
    if not path or not os.path.exists(path):
        return
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "Listes" not in wb.sheetnames:
            wb.close()
            return
        ws = wb["Listes"]
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not all_rows:
            return
        header_row = all_rows[0]
        data_rows = all_rows[1:]

        def col(idx):
            return [str(r[idx]) for r in data_rows if len(r) > idx and r[idx] is not None]

        _lists["pilotes"]      = col(0)   # A
        _lists["pilotes_pw"]   = col(1)   # B
        _lists["copilotes"]    = col(2)   # C
        _lists["tailles"]      = col(3)   # D
        _lists["types_produit"]= col(4)   # E
        _lists["postes"]       = col(6)   # G
        _lists["nb_personnes"] = col(7)   # H
        _lists["fibres"]       = col(9)   # J
        _lists["arrets"]       = col(10)  # K — liste des types d'arrêt

        # Prod de référence : première valeur non nulle de la col I (idx 8)
        for r in data_rows:
            if len(r) > 8 and r[8] is not None:
                try:
                    _lists["prod_ref"] = float(str(r[8]).replace(",", "."))
                    break
                except Exception:
                    pass

        # Coef d'équivalence : type produit (E) → coef (F)
        types = col(4)
        coefs = col(5)
        _lists["equiv_coef"] = {t: c for t, c in zip(types, coefs) if t and c}

        # Horaires : en-têtes cols L/M/N/O (idx 11–14), lignes 2 et 3 = heure début/fin
        horaires = {}
        for ci in (11, 12, 13, 14):
            if ci < len(header_row) and header_row[ci]:
                nom = str(header_row[ci])
                vals = [r[ci] for r in data_rows if len(r) > ci and r[ci] is not None]
                if len(vals) >= 2:
                    horaires[nom] = {"debut": str(vals[0]), "fin": str(vals[1])}
        _lists["horaires"] = horaires

    except Exception as e:
        print(f"Erreur chargement listes: {e}")

# ──────────────────────────────────────────────────────────────────────────────
# EXCEL — ÉCRITURE
# ──────────────────────────────────────────────────────────────────────────────
def _ensure_sheet(wb):
    if "Evenements" not in wb.sheetnames:
        ws = wb.create_sheet("Evenements")
        for i, h in enumerate(EVENEMENTS_HEADERS, 1):
            c = ws.cell(1, i, h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1e3a5f")
            c.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
    return wb["Evenements"]

def write_declaration(data: dict):
    path = _cfg.get("db_path", "")
    if not path:
        raise ValueError("Fichier Excel non configuré")
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    ws = _ensure_sheet(wb)
    ws.append([data.get(h, "") for h in EVENEMENTS_HEADERS])
    wb.save(path)
    hf = data.get("Heure Fin", "")
    if hf:
        _session["last_heure_fin"] = hf

def toggle_hors_trs(row_num: int, value: str):
    path = _cfg.get("db_path", "")
    if not path:
        return
    wb = openpyxl.load_workbook(path)
    ws = _ensure_sheet(wb)
    ws.cell(row_num, IDX["Prévu/Hors TRS"] + 1).value = value
    wb.save(path)

# ──────────────────────────────────────────────────────────────────────────────
# EXCEL — LECTURE
# ──────────────────────────────────────────────────────────────────────────────
def get_declarations_today():
    path = _cfg.get("db_path", "")
    if not path or not os.path.exists(path):
        return []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "Evenements" not in wb.sheetnames:
            wb.close()
            return []
        ws = wb["Evenements"]
        today = date.today().strftime("%d/%m/%Y")
        poste = _session.get("poste", "")
        result = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[IDX["Type"]]:
                continue
            rd = str(row[IDX["Date"]] or "")
            rp = str(row[IDX["Poste"]] or "")
            if rd == today and (not poste or rp == poste):
                result.append([str(v) if v is not None else "" for v in row])
        wb.close()
        return result
    except Exception as e:
        print(f"Erreur lecture: {e}")
        return []

# ──────────────────────────────────────────────────────────────────────────────
# TRS
# ──────────────────────────────────────────────────────────────────────────────
def calc_trs():
    rows = get_declarations_today()
    prod_ref = _lists.get("prod_ref", 0)
    if not prod_ref:
        return {"trs": None, "equiv_total": 0, "objectif": 0, "nb_prod": 0}
    equiv_total = 0
    nb_prod = 0
    for row in rows:
        if row[IDX["Type"]] != "Production":
            continue
        if row[IDX["Prévu/Hors TRS"]] == "OUI":
            continue
        try:
            equiv_total += float(row[IDX["Equivalence"]].replace(",", ".") or 0)
            nb_prod += 1
        except Exception:
            pass
    try:
        h1, m1 = map(int, _session["horaire_debut"].split(":"))
        h2, m2 = map(int, _session["horaire_fin"].split(":"))
        duree_h = (h2 * 60 + m2 - h1 * 60 - m1) / 60
    except Exception:
        duree_h = 8.0
    objectif = prod_ref * duree_h / 8
    trs = (equiv_total / objectif * 100) if objectif > 0 else None
    return {
        "trs": round(trs, 1) if trs is not None else None,
        "equiv_total": round(equiv_total, 2),
        "objectif": round(objectif, 1),
        "nb_prod": nb_prod,
    }

# ──────────────────────────────────────────────────────────────────────────────
# HTML SUPERVISION
# ──────────────────────────────────────────────────────────────────────────────
def generate_supervision_html():
    rows = get_declarations_today()
    trs_data = calc_trs()
    trs = trs_data.get("trs")
    trs_color = "#22c55e" if trs and trs >= 85 else "#f59e0b" if trs and trs >= 70 else "#ef4444"
    trs_str = f"{trs}%" if trs is not None else "N/A"
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    rows_html = ""
    for r in rows:
        typ = r[IDX["Type"]]
        hd = r[IDX["Heure Debut"]]
        hf = r[IDX["Heure Fin"]]
        dur = r[IDX["Duree"]]
        equiv = r[IDX["Equivalence"]]
        of_ = r[IDX["OF"]]
        is_interposte = "interposte" in typ.lower()
        is_long = False
        if is_interposte and dur:
            try:
                p = dur.split(":")
                mins = int(p[0]) * 60 + int(p[1]) if len(p) >= 2 else int(p[0])
                is_long = mins > 20
            except Exception:
                pass
        if is_long:
            style = "animation:pulse 1s ease-in-out infinite;font-weight:700;"
        elif typ == "Production":
            style = "background:#f0fdf4;"
        else:
            style = "background:#fff7ed;"
        rows_html += f'<tr style="{style}"><td>{typ}</td><td>{of_}</td><td>{hd}</td><td>{hf}</td><td>{dur}</td><td>{equiv}</td></tr>'
    html = f"""<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8"><meta http-equiv="refresh" content="30">
<title>KPI ORC — Supervision</title>
<style>
@keyframes pulse{{0%,100%{{background:#dc2626;color:#fff}}50%{{background:#7f1d1d;color:#fca5a5}}}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#0f172a;color:#e2e8f0;margin:0;padding:20px}}
.hdr{{background:#1e3a5f;border-radius:12px;padding:20px 28px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center}}
.trs{{font-size:3em;font-weight:800;color:{trs_color}}}
table{{width:100%;border-collapse:collapse}}
th{{background:#1e3a5f;padding:10px 12px;text-align:left;font-weight:600}}
td{{padding:8px 12px;border-bottom:1px solid #334155}}
</style></head><body>
<div class="hdr">
<div><div style="font-size:1.6em;font-weight:800">KPI ORC — Supervision</div>
<div style="margin-top:6px;color:#94a3b8">Poste: {_session['poste']} | Pilote: {_session['pilote']} | {now_str}</div></div>
<div class="trs">TRS {trs_str}</div></div>
<table><tr><th>Type</th><th>OF</th><th>Début</th><th>Fin</th><th>Durée</th><th>Équivalence</th></tr>
{rows_html}</table></body></html>"""
    out = _cfg.get("html_output", "")
    if not out and _cfg.get("db_path"):
        out = os.path.join(os.path.dirname(_cfg["db_path"]), "supervision.html")
    if out:
        try:
            with open(out, "w", encoding="utf-8") as f:
                f.write(html)
        except Exception as e:
            print(f"Erreur écriture HTML: {e}")
    return html

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def calc_duration(h_debut: str, h_fin: str) -> str:
    try:
        h1, m1 = map(int, h_debut.split(":"))
        h2, m2 = map(int, h_fin.split(":"))
        total = h2 * 60 + m2 - h1 * 60 - m1
        if total < 0:
            total += 24 * 60
        return f"{total // 60:02d}:{total % 60:02d}"
    except Exception:
        return ""

# ──────────────────────────────────────────────────────────────────────────────
# FLASK
# ──────────────────────────────────────────────────────────────────────────────
flask_app = Flask(__name__)

@flask_app.route("/")
def index():
    return HTML

@flask_app.route("/api/config", methods=["GET"])
def api_get_config():
    return jsonify({"db_path": _cfg.get("db_path", ""), "html_output": _cfg.get("html_output", "")})

@flask_app.route("/api/config", methods=["POST"])
def api_set_config():
    d = request.json or {}
    if d.get("supervisor_pw", "") != _cfg.get("supervisor_pw", "admin"):
        return jsonify({"error": "Mot de passe incorrect"}), 403
    if "db_path" in d:
        _cfg["db_path"] = d["db_path"]
    if "html_output" in d:
        _cfg["html_output"] = d["html_output"]
    if d.get("new_supervisor_pw"):
        _cfg["supervisor_pw"] = d["new_supervisor_pw"]
    save_config()
    load_lists()
    return jsonify({"ok": True})

@flask_app.route("/api/lists")
def api_lists():
    load_lists()
    return jsonify({
        "pilotes": _lists["pilotes"],
        "copilotes": _lists["copilotes"],
        "tailles": _lists["tailles"],
        "types_produit": _lists["types_produit"],
        "equiv_coef": _lists["equiv_coef"],
        "postes": _lists["postes"],
        "nb_personnes": _lists["nb_personnes"],
        "fibres": _lists["fibres"],
        "arrets": _lists["arrets"],
        "horaires": _lists["horaires"],
        "prod_ref": _lists["prod_ref"],
    })

@flask_app.route("/api/session")
def api_session():
    return jsonify(_session)

@flask_app.route("/api/login", methods=["POST"])
def api_login():
    d = request.json or {}
    pilote = d.get("pilote", "")
    pw = d.get("password", "")
    # Vérification mot de passe pilote
    if pilote in _lists["pilotes"]:
        idx = _lists["pilotes"].index(pilote)
        expected = _lists["pilotes_pw"][idx] if idx < len(_lists["pilotes_pw"]) else ""
        if expected and pw != expected:
            return jsonify({"error": "Mot de passe incorrect"}), 403
    horaire_debut = d.get("horaire_debut", "08:00")
    _session.update({
        "logged_in": True,
        "poste": d.get("poste", ""),
        "pilote": pilote,
        "copilote": d.get("copilote", ""),
        "nb_personnes": d.get("nb_personnes", 1),
        "horaire_nom": d.get("horaire_nom", ""),
        "horaire_debut": horaire_debut,
        "horaire_fin": d.get("horaire_fin", "16:00"),
        "last_heure_fin": horaire_debut,
    })
    # Interposte automatique si connexion en retard
    now_hm = datetime.now().strftime("%H:%M")
    if now_hm > horaire_debut:
        try:
            write_declaration({
                "Type": "Interposte",
                "Date": date.today().strftime("%d/%m/%Y"),
                "Poste": _session["poste"],
                "Pilote": pilote,
                "Co-Pilote": _session["copilote"],
                "Nb Personnes": _session["nb_personnes"],
                "Heure Debut": horaire_debut,
                "Heure Fin": now_hm,
                "Duree": calc_duration(horaire_debut, now_hm),
                "Commentaire": "Interposte connexion",
            })
        except Exception:
            pass
    return jsonify({"ok": True, "session": _session})

@flask_app.route("/api/logout", methods=["POST"])
def api_logout():
    _session.update({
        "logged_in": False, "poste": "", "pilote": "", "copilote": "",
        "nb_personnes": 1, "horaire_nom": "", "horaire_debut": "08:00",
        "horaire_fin": "16:00", "last_heure_fin": None,
    })
    return jsonify({"ok": True})

@flask_app.route("/api/declaration", methods=["POST"])
def api_declaration():
    if not _session.get("logged_in"):
        return jsonify({"error": "Non connecté"}), 401
    d = request.json or {}
    d.setdefault("Date", date.today().strftime("%d/%m/%Y"))
    d.setdefault("Poste", _session["poste"])
    d.setdefault("Pilote", _session["pilote"])
    d.setdefault("Co-Pilote", _session["copilote"])
    d.setdefault("Nb Personnes", _session["nb_personnes"])
    hd = d.get("Heure Debut", "")
    hf = d.get("Heure Fin", "")
    if hd and hf and not d.get("Duree"):
        d["Duree"] = calc_duration(hd, hf)
    # Cadences auto pour production
    if d.get("Type") == "Production" and hd and hf:
        try:
            h1, m1 = map(int, hd.split(":"))
            h2, m2 = map(int, hf.split(":"))
            duree_h = (h2 * 60 + m2 - h1 * 60 - m1) / 60
            qte = float(str(d.get("Qte Fabriquee", 0) or 0).replace(",", "."))
            nb = max(1, int(d.get("Nb Personnes", 1) or 1))
            if duree_h > 0 and qte > 0:
                d["Cadence/heure"] = round(qte / duree_h, 1)
                d["Cadence/h/pers"] = round(qte / duree_h / nb, 2)
        except Exception:
            pass
    try:
        write_declaration(d)
        generate_supervision_html()
        return jsonify({"ok": True, "trs": calc_trs()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@flask_app.route("/api/declarations")
def api_declarations():
    rows = get_declarations_today()
    result = []
    for i, row in enumerate(rows):
        rd = {EVENEMENTS_HEADERS[j]: (row[j] if j < len(row) else "")
              for j in range(len(EVENEMENTS_HEADERS))}
        rd["_row_num"] = i + 2  # +1 header, +1 pour 1-based
        result.append(rd)
    return jsonify(result)

@flask_app.route("/api/hors_trs", methods=["POST"])
def api_hors_trs():
    d = request.json or {}
    if d.get("supervisor_pw", "") != _cfg.get("supervisor_pw", "admin"):
        return jsonify({"error": "Mot de passe incorrect"}), 403
    row_num = d.get("row_num")
    value = d.get("value", "")
    if not row_num:
        return jsonify({"error": "row_num manquant"}), 400
    try:
        toggle_hors_trs(int(row_num), value)
        generate_supervision_html()
        return jsonify({"ok": True, "trs": calc_trs()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@flask_app.route("/api/trs")
def api_trs():
    return jsonify(calc_trs())

@flask_app.route("/api/last_heure_fin")
def api_last_heure_fin():
    return jsonify({
        "last_heure_fin": _session.get("last_heure_fin") or _session.get("horaire_debut", "08:00")
    })

@flask_app.route("/api/select_db", methods=["POST"])
def api_select_db():
    try:
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Fichiers Excel (*.xlsx)",),
        )
        if result:
            _cfg["db_path"] = result[0]
            save_config()
            load_lists()
            return jsonify({"ok": True, "path": result[0]})
    except Exception as e:
        print(f"Erreur select_db: {e}")
    return jsonify({"ok": False})

@flask_app.route("/api/select_html", methods=["POST"])
def api_select_html():
    try:
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="supervision.html",
            file_types=("Fichiers HTML (*.html)",),
        )
        if result:
            path = result if isinstance(result, str) else result[0]
            _cfg["html_output"] = path
            save_config()
            return jsonify({"ok": True, "path": path})
    except Exception as e:
        print(f"Erreur select_html: {e}")
    return jsonify({"ok": False})

@flask_app.route("/supervision")
def supervision():
    return generate_supervision_html()

# ──────────────────────────────────────────────────────────────────────────────
# HTML / CSS / JS (interface complète)
# ──────────────────────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>KPI ORC</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --navy:#1e3a5f; --navy-d:#0f172a; --blue:#2563eb;
  --green:#22c55e; --red:#ef4444; --orange:#f59e0b;
  --bg:#f1f5f9; --card:#fff; --text:#0f172a;
  --muted:#64748b; --border:#e2e8f0; --r:12px;
}
body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh}

/* VUES */
.view{display:none}.view.active{display:block}

/* LOGIN */
.login-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;background:var(--navy-d)}
.login-card{background:var(--card);border-radius:16px;padding:40px;width:440px;box-shadow:0 20px 60px rgba(0,0,0,.4)}
.login-logo{text-align:center;font-size:2.2em;font-weight:900;color:var(--navy);margin-bottom:28px;letter-spacing:2px}
.login-logo span{color:var(--blue)}

/* HEADER */
.hdr{background:var(--navy);color:#fff;padding:0 24px;height:60px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.3)}
.hdr-title{font-size:1.3em;font-weight:800;letter-spacing:1px}
.hdr-right{display:flex;gap:10px;align-items:center}

/* LAYOUT */
.main-grid{display:grid;grid-template-columns:300px 1fr;gap:20px;padding:20px}
.sidebar{display:flex;flex-direction:column;gap:16px}

/* CARD */
.card{background:var(--card);border-radius:var(--r);padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.card-title{font-size:.73em;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:14px}

/* FORM */
.fg{margin-bottom:14px}
.fg label{display:block;font-size:.78em;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px}
.fg input,.fg select{width:100%;padding:9px 11px;border:1.5px solid var(--border);border-radius:8px;font-size:.93em;background:#fff;outline:none;transition:border-color .15s}
.fg input:focus,.fg select:focus{border-color:var(--blue)}
.fr{display:grid;grid-template-columns:1fr 1fr;gap:12px}

/* BOUTONS */
.btn{padding:9px 18px;border:none;border-radius:8px;font-size:.93em;font-weight:700;cursor:pointer;transition:opacity .15s,transform .1s}
.btn:hover{opacity:.85}.btn:active{transform:scale(.97)}
.btn-navy{background:var(--navy);color:#fff}
.btn-green{background:var(--green);color:#fff}
.btn-red{background:var(--red);color:#fff}
.btn-orange{background:var(--orange);color:#fff}
.btn-ghost{background:transparent;color:var(--muted);border:1.5px solid var(--border)}
.btn-full{width:100%}
.btn-lg{padding:15px 20px;font-size:1em}

/* TRS */
.trs-num{font-size:3em;font-weight:900;text-align:center}
.trs-bar-bg{background:var(--border);border-radius:999px;height:10px;margin-top:12px;overflow:hidden}
.trs-bar{height:100%;border-radius:999px;transition:width .5s,background .5s}

/* ACTION BUTTONS */
.ab{display:flex;align-items:center;gap:14px;padding:16px;border:none;border-radius:10px;width:100%;text-align:left;cursor:pointer;font-size:1em;font-weight:700;transition:transform .1s,opacity .15s}
.ab:hover{opacity:.9}.ab:active{transform:scale(.97)}
.ab .ic{font-size:1.6em}
.ab .lbl{line-height:1.2}
.ab .sub{font-size:.75em;font-weight:400;opacity:.75}

/* TABLEAU */
.tbl{width:100%;border-collapse:collapse;font-size:.87em}
.tbl th{padding:7px 10px;text-align:left;font-size:.73em;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);border-bottom:2px solid var(--border)}
.tbl td{padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:middle}
.tbl tr.prod{background:#f0fdf4}
.tbl tr.stop{background:#fff7ed}
.tbl tr.interposte-alert{animation:rowpulse 1s ease-in-out infinite}
@keyframes rowpulse{0%,100%{background:#fecaca}50%{background:#fee2e2}}

/* BADGE */
.bdg{display:inline-block;padding:2px 8px;border-radius:999px;font-size:.72em;font-weight:700}
.bdg-prod{background:#dcfce7;color:#15803d}
.bdg-stop{background:#ffedd5;color:#c2410c}
.bdg-interposte{background:#fee2e2;color:#b91c1c}
.bdg-hors{background:#f0fdf4;color:#166534;border:1px solid #bbf7d0}

/* MODAL */
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;z-index:1000;backdrop-filter:blur(2px)}
.overlay.hidden{display:none}
.modal{background:#fff;border-radius:16px;padding:28px;width:580px;max-width:96vw;max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal-title{font-size:1.15em;font-weight:800;margin-bottom:20px;color:var(--navy)}
.modal-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:20px}
hr.sep{border:none;border-top:1px solid var(--border);margin:14px 0}

/* GRILLE TYPES D'ARRÊT */
.stop-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:10px}
.stop-btn{padding:12px;border:2px solid var(--border);border-radius:10px;background:#fff;cursor:pointer;font-size:.87em;font-weight:600;text-align:center;transition:all .15s}
.stop-btn:hover{border-color:var(--red);background:#fff1f2;color:var(--red)}
.stop-btn.sel{border-color:var(--red);background:#fee2e2;color:#991b1b}

/* TOAST */
.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#1e293b;color:#fff;padding:11px 24px;border-radius:999px;font-weight:600;z-index:9999;animation:tin .25s ease;box-shadow:0 4px 20px rgba(0,0,0,.3)}
.toast.err{background:var(--red)}
@keyframes tin{from{opacity:0;transform:translateX(-50%) translateY(10px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}

/* MISC */
.hors-btn{font-size:.73em;padding:3px 8px;border-radius:6px;cursor:pointer;background:none;border:1px solid var(--border);color:var(--muted)}
.hors-btn:hover{border-color:var(--green);color:var(--green)}
.warn-box{background:#fff7ed;border-radius:8px;padding:12px;margin-bottom:16px;font-size:.88em;color:#92400e}
</style>
</head>
<body>

<!-- ══════════════════════════════════════ VUE LOGIN -->
<div id="v-login" class="view active">
<div class="login-wrap"><div class="login-card">
  <div class="login-logo">KPI <span>ORC</span></div>
  <div id="no-db" class="warn-box" style="display:none">
    ⚠️ Aucun fichier Excel configuré.
    <a href="#" onclick="openAdmin()" style="color:#92400e;font-weight:700">Configurer →</a>
  </div>
  <div class="fg"><label>Poste</label><select id="l-poste"></select></div>
  <div class="fg"><label>Pilote</label><select id="l-pilote"></select></div>
  <div class="fg"><label>Mot de passe pilote</label>
    <input type="password" id="l-pw" placeholder="Laisser vide si aucun"></div>
  <div class="fr">
    <div class="fg"><label>Co-Pilote</label><select id="l-copilote"></select></div>
    <div class="fg"><label>Nb Personnes</label><select id="l-nb"></select></div>
  </div>
  <div class="fg"><label>Modèle Horaire</label><select id="l-horaire" onchange="updateHoraireInfo()"></select></div>
  <div id="horaire-info" style="font-size:.82em;color:var(--muted);margin-bottom:14px"></div>
  <button class="btn btn-navy btn-full btn-lg" onclick="doLogin()">Connexion →</button>
  <div style="margin-top:12px;text-align:center">
    <a href="#" onclick="openAdmin()" style="font-size:.8em;color:var(--muted)">⚙️ Administration</a>
  </div>
</div></div>
</div>

<!-- ══════════════════════════════════════ VUE PRINCIPALE -->
<div id="v-main" class="view">
  <div class="hdr">
    <div class="hdr-title">KPI ORC <span style="font-size:.55em;opacity:.5">v""" + VERSION + """</span></div>
    <div id="hdr-info" style="font-size:.88em;color:#93c5fd;text-align:center"></div>
    <div class="hdr-right">
      <span id="hdr-clock" style="color:#93c5fd;font-size:.9em"></span>
      <button class="btn btn-ghost" style="color:#fff;border-color:rgba(255,255,255,.3)" onclick="openAdmin()">⚙️</button>
      <button class="btn btn-red" onclick="doLogout()">Déconnexion</button>
    </div>
  </div>
  <div class="main-grid">
    <div class="sidebar">
      <!-- TRS -->
      <div class="card">
        <div class="card-title">TRS du poste</div>
        <div class="trs-num" id="trs-val">—</div>
        <div style="text-align:center;font-size:.78em;color:var(--muted);margin-top:4px" id="trs-detail"></div>
        <div class="trs-bar-bg"><div class="trs-bar" id="trs-bar" style="width:0;background:var(--green)"></div></div>
      </div>
      <!-- ACTIONS -->
      <div class="card">
        <div class="card-title">Déclarer</div>
        <div style="display:flex;flex-direction:column;gap:10px">
          <button class="ab" style="background:#dbeafe;color:#1e3a5f" onclick="openProdModal()">
            <span class="ic">🏭</span>
            <div class="lbl">Déclarer Production<div class="sub">Saisir OF + quantités</div></div>
          </button>
          <button class="ab" style="background:#fee2e2;color:#991b1b" onclick="openStopModal()">
            <span class="ic">⛔</span>
            <div class="lbl">Déclarer un Arrêt<div class="sub">Arrêt / Rattrapage / PB</div></div>
          </button>
          <button class="ab" style="background:#fefce8;color:#713f12" onclick="openFinPoste()">
            <span class="ic">🏁</span>
            <div class="lbl">Fin de Poste<div class="sub">Clôturer la session</div></div>
          </button>
        </div>
      </div>
      <!-- SESSION -->
      <div class="card">
        <div class="card-title">Session en cours</div>
        <div id="sess-info" style="font-size:.87em;line-height:2"></div>
      </div>
    </div>
    <!-- TABLEAU -->
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">
        <div class="card-title" style="margin:0">Déclarations du jour</div>
        <button class="btn btn-ghost" onclick="loadDecls()" style="font-size:.8em">↻ Actualiser</button>
      </div>
      <div style="overflow-x:auto">
        <table class="tbl">
          <thead><tr>
            <th>Type</th><th>OF</th><th>Début</th><th>Fin</th>
            <th>Durée</th><th>Qté Fab.</th><th>Equiv.</th><th>Action</th>
          </tr></thead>
          <tbody id="decl-tbody"></tbody>
        </table>
      </div>
      <div id="decl-empty" style="text-align:center;padding:30px;color:var(--muted);display:none">
        Aucune déclaration aujourd'hui
      </div>
    </div>
  </div>
</div>

<!-- ══════════════════════════════════════ MODAL PRODUCTION -->
<div class="overlay hidden" id="m-prod">
<div class="modal">
  <div class="modal-title">🏭 Déclarer une Production</div>
  <div class="fr">
    <div class="fg"><label>Heure Début</label><input type="time" id="p-hdebut"></div>
    <div class="fg"><label>Heure Fin</label><input type="time" id="p-hfin"></div>
  </div>
  <div class="fr">
    <div class="fg"><label>OF (numéro)</label><input type="text" id="p-of" placeholder="Ex: 123456"></div>
    <div class="fg"><label>Taille</label><select id="p-taille"></select></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Code Produit</label><input type="text" id="p-code"></div>
    <div class="fg"><label>Type Produit</label><select id="p-type-prod" onchange="calcEquiv()"></select></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Fibre</label><select id="p-fibre"></select></div>
    <div class="fg"><label>Poids Garnissage (kg)</label><input type="number" id="p-poids" step="0.1"></div>
  </div>
  <hr class="sep">
  <div class="fr">
    <div class="fg"><label>Qté Fabriquée</label><input type="number" id="p-qte-fab" step="1" oninput="calcEquiv()"></div>
    <div class="fg"><label>Qté Emballée</label><input type="number" id="p-qte-emb" step="1"></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Équivalence (auto)</label><input type="number" id="p-equiv" step="0.01"></div>
    <div class="fg"><label>Changement de Série</label>
      <select id="p-chgt"><option value="">Non</option><option value="Oui">Oui</option></select></div>
  </div>
  <hr class="sep">
  <div class="fr">
    <div class="fg"><label>OF Taie</label><input type="text" id="p-of-taie"></div>
    <div class="fg"><label>Ref Taie</label><input type="text" id="p-ref-taie"></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Traca Fibre</label><input type="text" id="p-traca"></div>
    <div class="fg"><label>Kit</label>
      <select id="p-kit"><option value="">Non</option><option value="1">Kit 1 pièce</option><option value="2">Kit 2 pièces</option></select></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Qté Initiale Taie</label><input type="number" id="p-qte-taie"></div>
    <div class="fg"><label>Nb Taie 2nd Choix</label><input type="number" id="p-taie2"></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Nb Défaut Couture</label><input type="number" id="p-def-cout"></div>
    <div class="fg"><label>Mq Taie</label><input type="number" id="p-mq-taie"></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Mq Housse/Encart</label><input type="number" id="p-mq-housse"></div>
    <div class="fg"><label>Nb PP Cousue</label><input type="number" id="p-pp"></div>
  </div>
  <hr class="sep">
  <div class="fr">
    <div class="fg"><label>Ratt Pochon/Fibre</label><input type="number" id="p-r-pochon"></div>
    <div class="fg"><label>Ratt Couture</label><input type="number" id="p-r-couture"></div>
  </div>
  <div class="fr">
    <div class="fg"><label>Ratt Emballage</label><input type="number" id="p-r-emb"></div>
    <div class="fg"><label>Ratt Presse Souder</label><input type="number" id="p-r-presse-s"></div>
  </div>
  <div class="fg"><label>Ratt Presse ZIP</label><input type="number" id="p-r-presse-z"></div>
  <div class="fg"><label>Commentaire</label><input type="text" id="p-comment" placeholder="Optionnel"></div>
  <div class="modal-actions">
    <button class="btn btn-ghost" onclick="hide('m-prod')">Annuler</button>
    <button class="btn btn-green" onclick="submitProd()">✓ Enregistrer</button>
  </div>
</div></div>

<!-- ══════════════════════════════════════ MODAL ARRÊT — ÉTAPE 1: TYPE -->
<div class="overlay hidden" id="m-stop-type">
<div class="modal">
  <div class="modal-title">⛔ Choisir le type d'arrêt</div>
  <div class="stop-grid" id="stop-grid"></div>
  <div class="modal-actions">
    <button class="btn btn-ghost" onclick="hide('m-stop-type')">Annuler</button>
    <button class="btn btn-red" id="btn-stop-next" onclick="openStopDetails()" disabled>Suivant →</button>
  </div>
</div></div>

<!-- ══════════════════════════════════════ MODAL ARRÊT — ÉTAPE 2: DÉTAILS -->
<div class="overlay hidden" id="m-stop-det">
<div class="modal">
  <div class="modal-title">⛔ Arrêt : <span id="stop-type-lbl"></span></div>
  <div class="fr">
    <div class="fg"><label>Heure Début</label><input type="time" id="s-hdebut"></div>
    <div class="fg"><label>Heure Fin</label><input type="time" id="s-hfin"></div>
  </div>
  <div class="fg"><label>OF (optionnel)</label><input type="text" id="s-of"></div>
  <div class="fg"><label>Commentaire</label><input type="text" id="s-comment"></div>
  <div class="fg" style="display:flex;align-items:center;gap:10px;flex-direction:row">
    <input type="checkbox" id="s-hors-trs" style="width:auto;width:16px;height:16px">
    <label for="s-hors-trs" style="margin:0;font-size:.93em;text-transform:none;letter-spacing:0;color:var(--text);font-weight:600">
      Arrêt Prévu / Hors TRS
    </label>
  </div>
  <div class="modal-actions">
    <button class="btn btn-ghost" onclick="openStopModal()">← Retour</button>
    <button class="btn btn-red" onclick="submitStop()">✓ Enregistrer</button>
  </div>
</div></div>

<!-- ══════════════════════════════════════ MODAL FIN DE POSTE -->
<div class="overlay hidden" id="m-fin">
<div class="modal">
  <div class="modal-title">🏁 Fin de Poste</div>
  <p style="color:var(--muted);margin-bottom:18px">Confirmer la clôture du poste. La session sera fermée.</p>
  <div class="fg"><label>Nettoyage fin de poste</label>
    <select id="fin-nett"><option value="">Non effectué</option><option value="Oui">Oui — effectué</option></select></div>
  <div class="fg"><label>Commentaire</label><input type="text" id="fin-comment" placeholder="Optionnel"></div>
  <div class="modal-actions">
    <button class="btn btn-ghost" onclick="hide('m-fin')">Annuler</button>
    <button class="btn btn-red" onclick="submitFin()">✓ Confirmer Fin de Poste</button>
  </div>
</div></div>

<!-- ══════════════════════════════════════ MODAL ADMIN -->
<div class="overlay hidden" id="m-admin">
<div class="modal">
  <div class="modal-title">⚙️ Administration</div>
  <div class="fg"><label>Mot de passe admin</label>
    <input type="password" id="adm-pw" placeholder="Mot de passe administrateur">
  </div>
  <button class="btn btn-navy" onclick="adminAuth()">Valider</button>
  <div id="adm-content" style="display:none">
    <hr class="sep">
    <div style="margin-bottom:16px">
      <div class="card-title">Fichier Excel</div>
      <div style="font-size:.83em;color:var(--muted);margin-bottom:8px;word-break:break-all" id="adm-db">Non configuré</div>
      <button class="btn btn-navy" onclick="selectDb()">📂 Choisir le fichier Excel</button>
    </div>
    <hr class="sep">
    <div style="margin-bottom:16px">
      <div class="card-title">Fichier HTML Supervision</div>
      <div style="font-size:.83em;color:var(--muted);margin-bottom:8px;word-break:break-all" id="adm-html">Non configuré</div>
      <button class="btn btn-ghost" onclick="selectHtml()">📂 Emplacement supervision.html</button>
    </div>
    <hr class="sep">
    <div>
      <div class="card-title">Changer le mot de passe admin</div>
      <div class="fr">
        <div class="fg"><label>Nouveau MDP</label><input type="password" id="adm-np1"></div>
        <div class="fg"><label>Confirmer</label><input type="password" id="adm-np2"></div>
      </div>
      <button class="btn btn-ghost" onclick="changeAdmPw()">Mettre à jour</button>
    </div>
  </div>
  <div class="modal-actions">
    <button class="btn btn-ghost" onclick="hide('m-admin')">Fermer</button>
  </div>
</div></div>

<script>
// ── ÉTAT ──
let lists = {};
let selStopType = null;

// ── INIT ──
async function init() {
  try { lists = await api('/api/lists'); } catch(e) {}
  fillLogin();
  const sess = await api('/api/session').catch(() => ({}));
  if (sess.logged_in) showMain(sess);
  else showLogin();
  setInterval(tickClock, 1000);
  tickClock();
}

function fillLogin() {
  sel('l-poste', lists.postes || []);
  sel('l-pilote', lists.pilotes || []);
  sel('l-copilote', ['', ...(lists.copilotes || [])]);
  sel('l-nb', lists.nb_personnes || ['1','2','3','4','5']);
  const h = document.getElementById('l-horaire');
  h.innerHTML = '';
  Object.entries(lists.horaires || {}).forEach(([nom, v]) => {
    const o = document.createElement('option');
    o.value = nom; o.textContent = `${nom}  (${v.debut} – ${v.fin})`;
    o.dataset.debut = v.debut; o.dataset.fin = v.fin;
    h.appendChild(o);
  });
  updateHoraireInfo();
  if (!lists.postes || !lists.postes.length)
    document.getElementById('no-db').style.display = '';
}

function updateHoraireInfo() {
  const h = document.getElementById('l-horaire');
  const o = h.options[h.selectedIndex];
  document.getElementById('horaire-info').textContent = o && o.dataset.debut
    ? `Horaire: ${o.dataset.debut} → ${o.dataset.fin}` : '';
}

// ── VUES ──
function showLogin() {
  setView('v-login');
}
function showMain(sess) {
  setView('v-main');
  document.getElementById('hdr-info').innerHTML =
    `${sess.poste} | <strong>${sess.pilote}</strong> | ${sess.horaire_debut} – ${sess.horaire_fin}`;
  document.getElementById('sess-info').innerHTML =
    `<div><b>Poste:</b> ${sess.poste}</div>
     <div><b>Pilote:</b> ${sess.pilote}</div>
     ${sess.copilote ? `<div><b>Co-Pilote:</b> ${sess.copilote}</div>` : ''}
     <div><b>Personnes:</b> ${sess.nb_personnes}</div>
     <div><b>Horaire:</b> ${sess.horaire_debut} – ${sess.horaire_fin}</div>`;
  // Pré-remplir listes des formulaires
  sel('p-taille', ['', ...(lists.tailles || [])]);
  sel('p-type-prod', ['', ...(lists.types_produit || [])]);
  sel('p-fibre', ['', ...(lists.fibres || [])]);
  loadDecls();
  loadTrs();
  setInterval(loadTrs, 60000);
}
function setView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

// ── CLOCK ──
function tickClock() {
  const d = new Date();
  const t = pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  const el = document.getElementById('hdr-clock');
  if (el) el.textContent = t;
}
function pad(n) { return String(n).padStart(2,'0'); }

// ── LOGIN ──
async function doLogin() {
  const hSel = document.getElementById('l-horaire');
  const opt = hSel.options[hSel.selectedIndex];
  const data = {
    poste: v('l-poste'), pilote: v('l-pilote'), password: v('l-pw'),
    copilote: v('l-copilote'), nb_personnes: v('l-nb') || 1,
    horaire_nom: v('l-horaire'),
    horaire_debut: opt ? opt.dataset.debut : '08:00',
    horaire_fin: opt ? opt.dataset.fin : '16:00',
  };
  const res = await api('/api/login', 'POST', data).catch(() => ({error: 'Erreur réseau'}));
  if (res.error) { toast(res.error, 1); return; }
  lists = await api('/api/lists').catch(() => lists);
  showMain(res.session);
  toast('Connexion réussie !');
}

async function doLogout() {
  if (!confirm('Déconnecter le pilote ?')) return;
  await api('/api/logout', 'POST');
  showLogin();
  toast('Déconnecté');
}

// ── TRS ──
async function loadTrs() {
  const t = await api('/api/trs').catch(() => ({}));
  const val = t.trs;
  const el = document.getElementById('trs-val');
  const bar = document.getElementById('trs-bar');
  const det = document.getElementById('trs-detail');
  if (val === null || val === undefined) {
    el.textContent = '—'; el.style.color = '#94a3b8';
    bar.style.width = '0'; det.textContent = 'Aucune production';
    return;
  }
  const col = val >= 85 ? '#22c55e' : val >= 70 ? '#f59e0b' : '#ef4444';
  el.textContent = val + '%'; el.style.color = col;
  bar.style.width = Math.min(val, 100) + '%'; bar.style.background = col;
  det.textContent = `Equiv: ${t.equiv_total} / Obj: ${t.objectif} (${t.nb_prod} prod)`;
}

function applyTrs(t) {
  if (t && t.trs !== undefined) {
    const val = t.trs;
    if (val === null) return;
    const col = val >= 85 ? '#22c55e' : val >= 70 ? '#f59e0b' : '#ef4444';
    document.getElementById('trs-val').textContent = val + '%';
    document.getElementById('trs-val').style.color = col;
    document.getElementById('trs-bar').style.width = Math.min(val,100) + '%';
    document.getElementById('trs-bar').style.background = col;
    document.getElementById('trs-detail').textContent =
      `Equiv: ${t.equiv_total} / Obj: ${t.objectif} (${t.nb_prod} prod)`;
  }
}

// ── TABLEAU DÉCLARATIONS ──
async function loadDecls() {
  const rows = await api('/api/declarations').catch(() => []);
  const tbody = document.getElementById('decl-tbody');
  const empty = document.getElementById('decl-empty');
  tbody.innerHTML = '';
  if (!rows.length) { empty.style.display = ''; return; }
  empty.style.display = 'none';
  [...rows].reverse().forEach(r => {
    const typ = r['Type'] || '';
    const isProd = typ === 'Production';
    const isInter = typ.toLowerCase().includes('interposte');
    const isHors = r['Prévu/Hors TRS'] === 'OUI';
    let interAlert = false;
    if (isInter && r['Duree']) {
      const p = r['Duree'].split(':');
      const mins = p.length >= 2 ? parseInt(p[0])*60 + parseInt(p[1]) : parseInt(p[0]);
      interAlert = mins > 20;
    }
    const cls = isProd ? 'prod' : 'stop';
    const badge = isProd ? `<span class="bdg bdg-prod">${esc(typ)}</span>`
      : isInter ? `<span class="bdg bdg-interposte">${esc(typ)}</span>`
      : `<span class="bdg bdg-stop">${esc(typ)}</span>`;
    const action = isHors
      ? `<span class="bdg bdg-hors">Prévu ✓</span>`
      : `<button class="hors-btn" onclick="markHorsTrs(${r['_row_num']})">Marquer Prévu</button>`;
    const tr = document.createElement('tr');
    tr.className = cls + (interAlert ? ' interposte-alert' : '');
    tr.innerHTML = `<td>${badge}</td><td>${esc(r['OF'])}</td>`
      + `<td>${esc(r['Heure Debut'])}</td><td>${esc(r['Heure Fin'])}</td>`
      + `<td>${esc(r['Duree'])}</td><td>${esc(r['Qte Fabriquee'])}</td>`
      + `<td>${esc(r['Equivalence'])}</td><td>${action}</td>`;
    tbody.appendChild(tr);
  });
}

async function markHorsTrs(rowNum) {
  const pw = prompt('Mot de passe superviseur :');
  if (pw === null) return;
  const res = await api('/api/hors_trs', 'POST', {supervisor_pw: pw, row_num: rowNum, value: 'OUI'});
  if (res.error) { toast(res.error, 1); return; }
  toast('Mis à jour !');
  loadDecls(); applyTrs(res.trs);
}

// ── PRODUCTION ──
async function openProdModal() {
  const last = await api('/api/last_heure_fin').catch(() => ({}));
  document.getElementById('p-hdebut').value = last.last_heure_fin || nowHM();
  document.getElementById('p-hfin').value = nowHM();
  // Reset fields
  ['p-of','p-code','p-poids','p-qte-fab','p-qte-emb','p-equiv',
   'p-of-taie','p-ref-taie','p-traca','p-qte-taie','p-taie2',
   'p-def-cout','p-mq-taie','p-mq-housse','p-pp',
   'p-r-pochon','p-r-couture','p-r-emb','p-r-presse-s','p-r-presse-z','p-comment']
    .forEach(id => { const el = document.getElementById(id); if(el) el.value = ''; });
  show('m-prod');
}

function calcEquiv() {
  const tp = v('p-type-prod');
  const qte = parseFloat(v('p-qte-fab')) || 0;
  const coef = parseFloat((lists.equiv_coef || {})[tp]) || 0;
  if (qte && coef)
    document.getElementById('p-equiv').value = Math.round(qte * coef * 100) / 100;
}

async function submitProd() {
  if (!v('p-of')) { toast('Numéro OF obligatoire', 1); return; }
  const data = {
    'Type': 'Production',
    'OF': v('p-of'), 'Taille': v('p-taille'), 'Code Produit': v('p-code'),
    'Type Produit': v('p-type-prod'), 'Fibre': v('p-fibre'), 'Poids Garnissage': v('p-poids'),
    'Heure Debut': v('p-hdebut'), 'Heure Fin': v('p-hfin'),
    'Qte Fabriquee': v('p-qte-fab'), 'Qte Emballee': v('p-qte-emb'), 'Equivalence': v('p-equiv'),
    'OF Taie': v('p-of-taie'), 'Ref Taie': v('p-ref-taie'), 'Traca Fibre': v('p-traca'),
    'Kit': v('p-kit'), 'Qte Initiale Taie': v('p-qte-taie'), 'Nb Taie 2nd Choix': v('p-taie2'),
    'Nb Defaut Couture': v('p-def-cout'), 'Mq Taie': v('p-mq-taie'),
    'Mq Housse/Encart': v('p-mq-housse'), 'Nb PP Cousue': v('p-pp'),
    'Changement de Serie': v('p-chgt'),
    'Ratt Pochon/Fibre': v('p-r-pochon'), 'Ratt Couture': v('p-r-couture'),
    'Ratt Emballage': v('p-r-emb'), 'Ratt Presse Souder': v('p-r-presse-s'),
    'Ratt Presse ZIP': v('p-r-presse-z'), 'Commentaire': v('p-comment'),
  };
  const res = await api('/api/declaration', 'POST', data);
  if (res.error) { toast(res.error, 1); return; }
  hide('m-prod'); toast('Production enregistrée !');
  loadDecls(); applyTrs(res.trs);
}

// ── ARRÊT ──
function openStopModal() {
  hide('m-stop-det');
  selStopType = null;
  const grid = document.getElementById('stop-grid');
  grid.innerHTML = '';
  (lists.arrets || []).forEach(type => {
    const btn = document.createElement('button');
    btn.className = 'stop-btn'; btn.textContent = type;
    btn.onclick = () => {
      document.querySelectorAll('.stop-btn').forEach(b => b.classList.remove('sel'));
      btn.classList.add('sel');
      selStopType = type;
      document.getElementById('btn-stop-next').disabled = false;
    };
    grid.appendChild(btn);
  });
  document.getElementById('btn-stop-next').disabled = true;
  show('m-stop-type');
}

async function openStopDetails() {
  if (!selStopType) return;
  hide('m-stop-type');
  document.getElementById('stop-type-lbl').textContent = selStopType;
  const last = await api('/api/last_heure_fin').catch(() => ({}));
  document.getElementById('s-hdebut').value = last.last_heure_fin || nowHM();
  document.getElementById('s-hfin').value = nowHM();
  document.getElementById('s-of').value = '';
  document.getElementById('s-comment').value = '';
  document.getElementById('s-hors-trs').checked = false;
  show('m-stop-det');
}

async function submitStop() {
  const data = {
    'Type': selStopType,
    'OF': v('s-of'), 'Heure Debut': v('s-hdebut'), 'Heure Fin': v('s-hfin'),
    'Commentaire': v('s-comment'),
    'Prévu/Hors TRS': document.getElementById('s-hors-trs').checked ? 'OUI' : '',
  };
  const res = await api('/api/declaration', 'POST', data);
  if (res.error) { toast(res.error, 1); return; }
  hide('m-stop-det'); toast('Arrêt enregistré !');
  loadDecls(); applyTrs(res.trs);
}

// ── FIN DE POSTE ──
function openFinPoste() { show('m-fin'); }

async function submitFin() {
  const nett = v('fin-nett');
  if (nett) {
    await api('/api/declaration', 'POST', {
      'Type': 'Nettoyage Fin de Poste', 'Nettoyage Fin de Poste': nett,
      'Commentaire': v('fin-comment'), 'Heure Debut': nowHM(), 'Heure Fin': nowHM(),
    });
  }
  await api('/api/logout', 'POST');
  hide('m-fin'); showLogin(); toast('Poste clôturé. À bientôt !');
}

// ── ADMIN ──
function openAdmin() {
  document.getElementById('adm-pw').value = '';
  document.getElementById('adm-content').style.display = 'none';
  show('m-admin');
}
async function adminAuth() {
  const pw = v('adm-pw');
  const res = await api('/api/config', 'POST', {supervisor_pw: pw, db_path: _cfg_cached});
  if (res.error) { toast('Mot de passe incorrect', 1); return; }
  document.getElementById('adm-content').style.display = '';
  const cfg = await api('/api/config');
  document.getElementById('adm-db').textContent = cfg.db_path || 'Non configuré';
  document.getElementById('adm-html').textContent = cfg.html_output || 'Non configuré';
}
let _cfg_cached = '';
api('/api/config').then(c => { _cfg_cached = c.db_path || ''; });

async function selectDb() {
  const res = await api('/api/select_db', 'POST');
  if (res.ok) {
    document.getElementById('adm-db').textContent = res.path;
    _cfg_cached = res.path;
    lists = await api('/api/lists').catch(() => lists);
    fillLogin(); toast('Fichier Excel configuré !');
  }
}
async function selectHtml() {
  const res = await api('/api/select_html', 'POST');
  if (res.ok) { document.getElementById('adm-html').textContent = res.path; toast('HTML configuré !'); }
}
async function changeAdmPw() {
  const pw = v('adm-pw'), np1 = v('adm-np1'), np2 = v('adm-np2');
  if (!np1) { toast('Nouveau mot de passe vide', 1); return; }
  if (np1 !== np2) { toast('Les mots de passe ne correspondent pas', 1); return; }
  const res = await api('/api/config', 'POST', {supervisor_pw: pw, new_supervisor_pw: np1});
  if (res.error) { toast(res.error, 1); return; }
  toast('Mot de passe mis à jour !');
}

// ── UTILITAIRES ──
function nowHM() {
  const d = new Date();
  return pad(d.getHours()) + ':' + pad(d.getMinutes());
}
function v(id) { const e = document.getElementById(id); return e ? e.value.trim() : ''; }
function sel(id, opts) {
  const e = document.getElementById(id); if (!e) return; e.innerHTML = '';
  opts.forEach(o => { const op = document.createElement('option'); op.value = o; op.textContent = o; e.appendChild(op); });
}
function show(id) { document.getElementById(id).classList.remove('hidden'); }
function hide(id) { document.getElementById(id).classList.add('hidden'); }
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function toast(msg, err) {
  const e = document.createElement('div');
  e.className = 'toast' + (err ? ' err' : ''); e.textContent = msg;
  document.body.appendChild(e);
  setTimeout(() => e.remove(), 3000);
}
async function api(url, method, body) {
  const opts = {method: method||'GET', headers:{'Content-Type':'application/json'}};
  if (body) opts.body = JSON.stringify(body);
  const r = await fetch(url, opts);
  return r.json();
}
// Fermer modal en cliquant sur overlay
document.querySelectorAll('.overlay').forEach(o =>
  o.addEventListener('click', e => { if (e.target === o) o.classList.add('hidden'); }));

init();
</script>
</body></html>"""

# ──────────────────────────────────────────────────────────────────────────────
# DÉMARRAGE
# ──────────────────────────────────────────────────────────────────────────────
def start_flask():
    flask_app.run(port=PORT, debug=False, use_reloader=False, threaded=True)

if __name__ == "__main__":
    load_config()
    load_lists()
    t = threading.Thread(target=start_flask, daemon=True)
    t.start()
    time.sleep(0.8)
    webview.create_window(
        f"KPI ORC v{VERSION}",
        f"http://localhost:{PORT}",
        width=1280, height=820,
        min_size=(960, 640),
    )
    webview.start()
