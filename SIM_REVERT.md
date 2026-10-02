# Guide de suppression du simulateur d'horloge

Ce fichier liste **exactement** ce qui est lié à la simulation et doit être retiré,
versus les corrections qui doivent rester en production.

---

## À SUPPRIMER (sim uniquement)

### Python `main.py`

| Quoi | Localisation | Remplacement |
|---|---|---|
| Variable `_sim_offset_s = 0` | Après les globals | Supprimer la ligne |
| Fonction `_now()` | Juste après `_sim_offset_s` | Supprimer la fonction |
| Route `api_set_sim_time` | `@flask_app.route('/api/set_sim_time', ...)` | Supprimer l'endpoint entier |
| `api_end_prod` : `end_dt = _now()` | Dans `api_end_prod` | Remettre `datetime.datetime.now()` |
| `t_start` : `t["start"] = _now()` | Dans `t_start` | Remettre `datetime.datetime.now()` |
| `t_stop` : `end = end_time or _now()` | Dans `t_stop` | Remettre `end_time or datetime.datetime.now()` |
| `t_get` : `_now() - t["start"]` | Dans `t_get` | Remettre `datetime.datetime.now() - t["start"]` |
| `tl_open` : `"start": _now()` | Dans `tl_open` | Remettre `datetime.datetime.now()` |
| `tl_close` : `end_time or _now()` | Dans `tl_close` | Remettre `end_time or datetime.datetime.now()` |
| `tl_close_all` : `now = _now()` | Dans `tl_close_all` | Remettre `datetime.datetime.now()` |
| `_force_fin_poste_server` : `now = _now()` | Dans `_force_fin_poste_server` | Remettre `datetime.datetime.now()` |

### `api_set_of_start` (à vérifier)
Le bloc de recalcul de date ajouté dans ce commit :
```python
_now_ref = _now()
candidate = _now_ref.replace(hour=dt.hour, ...)
if (candidate - _now_ref).total_seconds() > 3600:
    candidate -= datetime.timedelta(days=1)
_S["of_start"] = candidate
```
→ Remettre simplement `_S["of_start"] = dt`

*(Note : cette logique est utile pour les postes de nuit même en prod. À évaluer.)*

### HTML/JS dans `main.py`

| Quoi | Identifier par |
|---|---|
| Widget simulateur dans Settings | Section avec `id="sim-debug-section"` ou bordure jaune tiretée, contient `sim-date-input`, `sim-time-input` |
| Badge sim dans le header | `id="sim-badge"` et `id="sim-badge-time"` |
| Fonction `_simUpdateBadge()` | `function _simUpdateBadge(` |
| Fonction `applySimTime()` | `async function applySimTime(` |
| Fonction `simAdvance()` | `async function simAdvance(` |
| Fonction `resetSimTime()` | `async function resetSimTime(` |

---

## À GARDER (corrections de prod déclenchées par les tests sim)

| Correction | Pourquoi garder |
|---|---|
| `_deg_s = 0.0` initialisé avant le `if` dans `api_end_prod` | Évite un NameError si `of_s_brut <= 0` (cas extrême) |
| JS try/catch autour de `r.json()` dans `confirmEndProd` | Robustesse : évite le freeze d'écran si le serveur retourne une erreur |
| `api_set_of_start` : suppression de `_S["shift_start"] = dt` | Empêche l'écrasement de la date d'index au démarrage d'OF |
| Bouton Déconnexion caché pendant le poste | Empêche la déconnexion accidentelle en cours de poste |
| Backfill kit (col 16) dans `_backfill_of_for_events` | Fix "Lot de 2" — bug prod réel |
| Couleurs timeline dans Détail OF | Fix bug prod réel |
| JS try/catch `confirmEndProd` | Robustesse générale |

---

## Commits concernés

```
9db272c  Sim : shift_start lock, logout button, simulateur (partiel à garder)
6da2dd0  Kit backfill + timeline colors (GARDER entièrement)
9ba1633  end_dt = _now(), _deg_s fix, JS try/catch (partiel)
45d8298  t_start/tl_open etc. → _now() (SUPPRIMER entièrement)
5f50e76  Sim date fix api_set_sim_time (SUPPRIMER avec l'endpoint)
1230f6d  api_set_of_start date recalc (à évaluer)
c4b473c  _force_fin_poste_server → _now() (SUPPRIMER)
```
