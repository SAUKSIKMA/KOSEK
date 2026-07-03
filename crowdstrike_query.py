"""
Module d'execution des requetes CrowdStrike Falcon via l'API REST (EU-1).

Adaptation 1:1 de sentinelquery.py : MEMES fonctions, MEMES signatures,
MEMES formats de retour (colonnes "liste" serialisees en JSON, dates ISO
8601, memes noms de colonnes que l'export CSV Sentinel), pour rester
compatible avec generate_cosec.py / anonymizer.py / reformulate.py SANS
LES MODIFIER. Seule la source change : l'API Cases + Alerts de CrowdStrike
remplace les tables SecurityIncident / SecurityAlert de Log Analytics.

Scopes API requis : Cases: Read + Alerts: Read (lecture seule).
Identifiants : variables d'environnement CLIENT_ID / CLIENT_SECRET.

Correspondances de concepts (cf lecture complete de sentinelquery.py) :
  SecurityIncident            -> Case CrowdStrike (/cases/)
  SecurityAlert (join)        -> Alertes liees, retrouvees dans evidence
  Owner.assignedTo (1re attr) -> sla.timers.acknowledgement.time_completed
  Tag "attente client"        -> pause SLA native (goal - (due - completed))
  Seuils SLA du workbook      -> reproduits a l'identique (cf SEUILS_SLA)
  Heures ouvrees lun-ven 8-18 -> reproduites (cf minutes_ouvrees), avec
                                 DECALAGE_UTC pour l'heure locale
  Label "KpiError"            -> PAS d'equivalent vu sur les cases : nuance
                                 non transposee (documente, cf note MTTR)
  Classification (TP/FP/BP)   -> PAS de champ natif vu : best-effort sur la
                                 liste "fields" de la case, sinon
                                 "Indéterminé" (meme valeur par defaut que
                                 CLASSIFICATION_QUERY_TEMPLATE)

Le parametre workspace_id des fonctions est CONSERVE pour la compatibilite
de signature avec sentinelquery.py mais IGNORE (pas de notion de workspace
chez CrowdStrike : le tenant est porte par le couple CLIENT_ID/SECRET).
Idem tenant_id.
"""

import os
import re
import json
import time
import urllib.parse
import urllib.request
from collections import Counter
from datetime import datetime, timezone, timedelta

# =============================== CONFIG ======================================
BASE_URL = "https://api.eu-1.crowdstrike.com"   # region EU-1
PROXY = ""        # ex: "http://proxy:8080" ; "" = proxy systeme/HTTPS_PROXY
TIMEOUT = 15

# Heures ouvrees : lun-ven 8h-18h (regle RESOLUTION_TIMES_QUERY_TEMPLATE).
HO_DEBUT, HO_FIN = 8, 18
# Fenetre 8h-18h exprimee en heure LOCALE : les horodatages API sont UTC.
# 0 = UTC strict (comportement exact du KQL) ; 2 = ete France ; 1 = hiver.
DECALAGE_UTC = 2

# Seuils SLA du workbook (minutes) -- repris de SLA_BREACHES_QUERY_TEMPLATE.
# "Critical" (severity >= 80, sans equivalent Sentinel) est aligne sur High.
SEUILS_SLA = {
    "MTTA": {"Critical": 30, "High": 30, "Medium": 45, "Low": 120},
    "MTTR": {"Critical": 1440, "High": 1440, "Medium": 2880, "Low": 2880},
}

MAX_CASES = 10000
MAX_ALERTES = 10000
# =============================================================================

# Identique a sentinelquery.py : colonnes serialisees en JSON pour
# parse_json_array() de generate_cosec.py.
LIST_COLUMNS = [
    "Tactics", "Techniques", "AlertSources",
    "Accounts", "Hosts", "IPs", "SecurityGroups",
    "URLs", "Files", "Processes", "CloudApps", "Mailboxes",
]

CHAMP = {
    "severity":      "severity",
    "name":          "name",
    "status":        "status",
    "creation":      "created_timestamp",
    "maj":           "updated_timestamp",
    "reference":     "reference_id",
    "ack_completed": "sla.timers.acknowledgement.time_completed",
    "ack_status":    "sla.timers.acknowledgement.status",
    "res_goal":      "sla.timers.resolution.duration_seconds",
    "res_started":   "sla.timers.resolution.time_started",
    "res_due":       "sla.timers.resolution.time_due",
    "res_completed": "sla.timers.resolution.time_completed",
    "res_status":    "sla.timers.resolution.status",
}


# ---------------------------------------------------------------------------
# Cache de jeton (equivalent du cache de credential de sentinelquery.py :
# une seule authentification par processus, renouvelee a l'approche de
# l'expiration -- le jeton OAuth2 CrowdStrike vit ~30 minutes).
# ---------------------------------------------------------------------------
_token_cache = {"token": None, "expire": 0.0}


def _opener():
    if PROXY:
        return urllib.request.build_opener(
            urllib.request.ProxyHandler({"http": PROXY, "https": PROXY}))
    return urllib.request.build_opener()


def _http(method, path, token=None, data=None, content_type=None):
    req = urllib.request.Request(BASE_URL + path, data=data, method=method)
    req.add_header("Accept", "application/json")
    if content_type:
        req.add_header("Content-Type", content_type)
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    try:
        with _opener().open(req, timeout=TIMEOUT) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        raise RuntimeError(
            f"Erreur HTTP {e.code} sur {path} : "
            f"{e.read().decode('utf-8', 'replace')}") from None
    except (urllib.error.URLError, TimeoutError) as e:
        raise RuntimeError(
            f"Connexion impossible vers {BASE_URL} ({getattr(e, 'reason', e)}). "
            f"Verifier VPN/reseau/proxy.") from None


def _get_token() -> str:
    """Jeton OAuth2 partage pour tout le processus (cf cache ci-dessus)."""
    if _token_cache["token"] and time.time() < _token_cache["expire"]:
        return _token_cache["token"]
    cid = os.environ.get("CLIENT_ID")
    secret = os.environ.get("CLIENT_SECRET")
    if not cid or not secret:
        raise RuntimeError("Definir les variables d'environnement CLIENT_ID et CLIENT_SECRET.")
    payload = urllib.parse.urlencode({
        "client_id": cid, "client_secret": secret,
        "grant_type": "client_credentials",
    }).encode("utf-8")
    resp = _http("POST", "/oauth2/token", data=payload,
                 content_type="application/x-www-form-urlencoded")
    token = resp.get("access_token")
    if not token:
        raise RuntimeError(f"Pas de jeton recu : {resp}")
    _token_cache["token"] = token
    _token_cache["expire"] = time.time() + int(resp.get("expires_in", 1800)) - 60
    return token


# --------------------------- utilitaires -----------------------------------

def _month_bounds(year: int, month: int) -> tuple:
    """Identique a sentinelquery.py : (debut, fin) UTC, fin EXCLUSIVE."""
    start = datetime(year, month, 1, tzinfo=timezone.utc)
    end = datetime(year + 1, 1, 1, tzinfo=timezone.utc) if month == 12 \
        else datetime(year, month + 1, 1, tzinfo=timezone.utc)
    return start, end


def _get_path(obj, chemin):
    cur = obj
    for cle in chemin.split("."):
        if isinstance(cur, dict) and cle in cur:
            cur = cur[cle]
        else:
            return None
    return cur


def _parse_ts(v):
    """ISO 8601 (y compris nanosecondes CrowdStrike) ou epoch -> datetime UTC."""
    if v in (None, "", 0):
        return None
    if isinstance(v, (int, float)):
        return datetime.fromtimestamp(v / 1000 if v > 1e12 else v, tz=timezone.utc)
    s = re.sub(r"(\.\d{6})\d+", r"\1", str(v).strip()).replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def _iso(dt):
    return dt.isoformat() if dt else ""


def minutes_ouvrees(debut, fin) -> int:
    """Minutes en heures ouvrees (lun-ven, HO_DEBUT-HO_FIN heure locale via
    DECALAGE_UTC) -- transposition du mv-expand horaire du KQL
    (WorkingMinutesMTTA / MTTR / MTTC)."""
    if not debut or not fin or fin <= debut:
        return 0
    total = 0.0
    heure = debut.replace(minute=0, second=0, microsecond=0)
    while heure < fin:
        h_fin = heure + timedelta(hours=1)
        locale = heure + timedelta(hours=DECALAGE_UTC)
        if locale.weekday() < 5 and HO_DEBUT <= locale.hour < HO_FIN:
            seg_debut = max(heure, debut)
            seg_fin = min(h_fin, fin)
            if seg_fin > seg_debut:
                total += (seg_fin - seg_debut).total_seconds() / 60
        heure = h_fin
    return int(total)


def _label_severite(n) -> str:
    """severity (int CrowdStrike) -> label texte facon Sentinel."""
    try:
        n = int(n)
    except (TypeError, ValueError):
        return str(n or "")
    if n >= 80:
        return "Critical"
    if n >= 60:
        return "High"
    if n >= 40:
        return "Medium"
    if n >= 20:
        return "Low"
    return "Informational"


def _fql_mois(champ_date, start, end) -> str:
    d = start.strftime("%Y-%m-%dT%H:%M:%SZ")
    f = end.strftime("%Y-%m-%dT%H:%M:%SZ")
    return f"{champ_date}:>='{d}'+{champ_date}:<'{f}'"


# --------------------------- collecte brute ---------------------------------

def _query_ids(path, fql, page=100, plafond=MAX_CASES):
    token = _get_token()
    ids, offset = [], 0
    while len(ids) < plafond:
        params = urllib.parse.urlencode({
            "filter": fql, "sort": "created_timestamp.asc",
            "limit": page, "offset": offset,
        })
        lot = _http("GET", f"{path}?{params}", token=token).get("resources", [])
        ids.extend(lot)
        if len(lot) < page:
            break
        offset += len(lot)
    return ids


def _fetch_cases_mois(year, month) -> list[dict]:
    start, end = _month_bounds(year, month)
    ids = _query_ids("/cases/queries/cases/v1",
                     _fql_mois("created_timestamp", start, end))
    token = _get_token()
    cases = []
    for i in range(0, len(ids), 1000):
        body = json.dumps({"ids": ids[i:i + 1000]}).encode("utf-8")
        cases.extend(_http("POST", "/cases/entities/cases/v2", token=token,
                           data=body, content_type="application/json"
                           ).get("resources", []))
    return cases


def _fetch_alertes(ids_composites) -> list[dict]:
    if not ids_composites:
        return []
    token = _get_token()
    alertes = []
    for i in range(0, len(ids_composites), 1000):
        body = json.dumps({"composite_ids": ids_composites[i:i + 1000]}).encode("utf-8")
        alertes.extend(_http("POST", "/alerts/entities/alerts/v2", token=token,
                             data=body, content_type="application/json"
                             ).get("resources", []))
    return alertes


def _fetch_alertes_mois(year, month) -> list[dict]:
    start, end = _month_bounds(year, month)
    ids = _query_ids("/alerts/queries/alerts/v2",
                     _fql_mois("created_timestamp", start, end),
                     page=1000, plafond=MAX_ALERTES)
    return _fetch_alertes(ids)


# ---------------------------------------------------------------------------
# Lien case -> alertes (equivalent du join SecurityIncident.AlertIds ->
# SecurityAlert). Les IDs d'alertes lies a une case sont recherches
# recursivement dans la case (typiquement sous "evidence") : toute chaine
# ressemblant a un composite_id d'alerte (motif ':ind:' des detections EDR,
# ou cle contenant 'alert' avec une valeur longue) est retenue.
# ---------------------------------------------------------------------------

def _collecter_ids_alertes(obj, sous_cle_alerte=False) -> set:
    trouves = set()
    if isinstance(obj, dict):
        for k, v in obj.items():
            trouves |= _collecter_ids_alertes(
                v, sous_cle_alerte or ("alert" in str(k).lower()))
    elif isinstance(obj, list):
        for v in obj:
            trouves |= _collecter_ids_alertes(v, sous_cle_alerte)
    elif isinstance(obj, str):
        if ":ind:" in obj or (sous_cle_alerte and len(obj) >= 20):
            trouves.add(obj)
    return trouves


# --------------------------- extraction d'entites ---------------------------

def _entites_depuis_alertes(alertes) -> dict:
    """Reconstruit les 9 colonnes d'entites du COSEC a partir des champs
    des alertes CrowdStrike (equivalent du mv-expand Entities du KQL).

    Correspondances :
      account           -> user_name
      host              -> hostname / device.hostname
      ip                -> local_ip / device.local_ip / external_ip
      file              -> filepath\\filename [sha256] (meme format que le
                           strcat Directory\\Name [hash] du KQL)
      process           -> cmdline
      url / cloud-app / mailbox / security-group : pas de champ generique
        equivalent sur une alerte CrowdStrike -> ensembles vides (colonnes
        conservees pour la compatibilite du format).
    """
    ens = {c: set() for c in ("Accounts", "Hosts", "IPs", "SecurityGroups",
                              "URLs", "Files", "Processes", "CloudApps",
                              "Mailboxes")}
    sources, tactiques, techniques = set(), set(), set()
    for a in alertes:
        u = a.get("user_name")
        if u:
            ens["Accounts"].add(str(u))
        h = a.get("hostname") or _get_path(a, "device.hostname")
        if h:
            ens["Hosts"].add(str(h))
        for champ_ip in ("local_ip", "external_ip"):
            ip = a.get(champ_ip) or _get_path(a, f"device.{champ_ip}")
            if ip:
                ens["IPs"].add(str(ip))
        nom_f = a.get("filename")
        if nom_f:
            rep = a.get("filepath") or ""
            h256 = a.get("sha256") or ""
            ens["Files"].add(f"{rep}\\{nom_f} [{h256}]" if (rep or h256) else str(nom_f))
        cmd = a.get("cmdline")
        if cmd:
            ens["Processes"].add(str(cmd))
        if a.get("product"):
            sources.add(str(a["product"]))
        if a.get("tactic"):
            tactiques.add(str(a["tactic"]))
        if a.get("technique"):
            techniques.add(str(a["technique"]))
    return {
        **{k: sorted(v) for k, v in ens.items()},
        "AlertSources": sorted(sources),
        "Tactics": sorted(tactiques),
        "Techniques": sorted(techniques),
    }


def _classification_case(case) -> tuple:
    """Best-effort : cherche une classification de cloture dans la liste
    'fields' de la case (pas de champ natif TruePositive/FalsePositive vu
    sur les cases CrowdStrike). Retourne (classification, raison,
    commentaire) -- "Indéterminé" par defaut, comme le iff(isempty(...))
    de CLASSIFICATION_QUERY_TEMPLATE."""
    classification, raison, commentaire = "Indéterminé", "", ""
    for f in (case.get("fields") or []):
        if not isinstance(f, dict):
            continue
        nom = str(f.get("name") or f.get("display_name") or "").lower()
        valeur = f.get("value") or f.get("values") or ""
        if isinstance(valeur, list):
            valeur = ", ".join(str(v) for v in valeur)
        valeur = str(valeur)
        if not valeur:
            continue
        if any(m in nom for m in ("classif", "disposition", "verdict")):
            classification = valeur
        elif any(m in nom for m in ("raison", "reason")):
            raison = valeur
        elif any(m in nom for m in ("comment", "commentaire", "note")):
            commentaire = valeur
    return classification, raison, commentaire


def _normalize_value(col_name: str, value) -> str:
    """Identique dans l'esprit a sentinelquery._normalize_value : colonnes
    listes -> chaine JSON ; datetime -> ISO 8601 ; reste -> str."""
    if value is None:
        return "[]" if col_name in LIST_COLUMNS else ""
    if col_name in LIST_COLUMNS:
        if isinstance(value, str):
            return value
        if isinstance(value, list):
            return json.dumps([str(v) for v in value])
        return json.dumps([str(value)])
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


# ---------------------------------------------------------------------------
# 1. fetch_cosec_incidents -- detail des cases CLOTUREES du mois
# ---------------------------------------------------------------------------

def fetch_cosec_incidents(workspace_id: str, year: int, month: int,
                          tenant_id: str = None) -> list[dict]:
    """
    Equivalent CrowdStrike de la requete COSEC : une ligne par case
    CLOTUREE creee dans le mois calendaire, MEMES colonnes que l'export
    CSV Sentinel (cf projection de COSEC_QUERY_TEMPLATE), colonnes listes
    serialisees en JSON.

    Nuances transposees / non transposees (cf notes de sentinelquery.py) :
      - "Status == Closed" -> status de la case == closed. CONSERVE.
      - "Classification non vide et != undetermined" : PAS de champ natif
        -> le filtre est ASSOUPLI a "cloturee" (toutes les cases closes
        sont retournees) ; la colonne Classification est remplie en
        best-effort depuis 'fields' (sinon "Indéterminé"). L'exception
        Title contains malware/software devient donc sans objet.
      - Occurrences (alertsCount) -> nombre d'alertes liees retrouvees.
      - Entites -> reconstruites depuis les alertes liees (cf
        _entites_depuis_alertes) ; ensembles vides si aucune alerte liee
        n'est retrouvable dans la case.
    """
    lignes = []
    for case in _fetch_cases_mois(year, month):
        statut = str(_get_path(case, CHAMP["status"]) or "").lower()
        if statut != "closed":
            continue

        ids_alertes = sorted(_collecter_ids_alertes(case))
        alertes = _fetch_alertes(ids_alertes) if ids_alertes else []
        entites = _entites_depuis_alertes(alertes)
        classification, raison, commentaire = _classification_case(case)

        creation = _parse_ts(_get_path(case, CHAMP["creation"]))
        cloture = _parse_ts(_get_path(case, CHAMP["res_completed"])) \
            or _parse_ts(_get_path(case, CHAMP["maj"]))

        ligne = {
            "IncidentName": case.get("id", ""),
            "Title": _get_path(case, CHAMP["name"]) or "",
            "Severity": _label_severite(_get_path(case, CHAMP["severity"])),
            "Occurrences": len(alertes) if alertes else "",
            "AlertSources": entites["AlertSources"],
            "Tactics": entites["Tactics"],
            "Techniques": entites["Techniques"],
            "Accounts": entites["Accounts"],
            "Hosts": entites["Hosts"],
            "IPs": entites["IPs"],
            "SecurityGroups": entites["SecurityGroups"],
            "URLs": entites["URLs"],
            "Files": entites["Files"],
            "Processes": entites["Processes"],
            "CloudApps": entites["CloudApps"],
            "Mailboxes": entites["Mailboxes"],
            "Classification": classification,
            "ClassificationReason": raison,
            "ClassificationComment": commentaire,
            "CreatedTime": creation,
            "ClosedTime": cloture,
        }
        lignes.append({k: _normalize_value(k, v) for k, v in ligne.items()})

    lignes.sort(key=lambda l: l["CreatedTime"], reverse=True)  # CreatedTime desc
    return lignes


# ---------------------------------------------------------------------------
# 2. fetch_typology_history -- nb de cases distinctes par (Title, Month)
# ---------------------------------------------------------------------------

def fetch_typology_history(workspace_id: str, year: int, month: int,
                           tenant_id: str = None) -> list[dict]:
    """Tous statuts, toutes severites (comme TYPOLOGY_QUERY_TEMPLATE).
    AlertSources agregees depuis les alertes liees quand disponibles."""
    mois_str = f"{year:04d}-{month:02d}"
    compteur, sources = Counter(), {}
    for case in _fetch_cases_mois(year, month):
        titre = _get_path(case, CHAMP["name"]) or "(sans nom)"
        compteur[titre] += 1
        ids = sorted(_collecter_ids_alertes(case))
        if ids:
            for a in _fetch_alertes(ids):
                if a.get("product"):
                    sources.setdefault(titre, set()).add(str(a["product"]))
    lignes = [{
        "Month": mois_str,
        "Title": titre,
        "AlertSources": json.dumps(sorted(sources.get(titre, set()))),
        "IncidentCount": n,
    } for titre, n in compteur.items()]
    lignes.sort(key=lambda l: -l["IncidentCount"])
    return lignes


# ---------------------------------------------------------------------------
# 3 & 4. Repartitions gravite / classification
# ---------------------------------------------------------------------------

def fetch_severity_breakdown(workspace_id: str, year: int, month: int,
                             tenant_id: str = None) -> list[dict]:
    """Nb de cases par gravite, tous statuts (cf SEVERITY_QUERY_TEMPLATE).
    Le filtre anti-bruit sur le titre Sentinel n'a pas d'equivalent ici."""
    compteur = Counter(
        _label_severite(_get_path(c, CHAMP["severity"]))
        for c in _fetch_cases_mois(year, month))
    return sorted(({"Severity": s, "IncidentCount": n} for s, n in compteur.items()),
                  key=lambda l: -l["IncidentCount"])


def fetch_classification_breakdown(workspace_id: str, year: int, month: int,
                                   tenant_id: str = None) -> list[dict]:
    """Nb de cases par classification de cloture -- best-effort sur
    'fields', "Indéterminé" par defaut (cf CLASSIFICATION_QUERY_TEMPLATE)."""
    compteur = Counter(
        _classification_case(c)[0] for c in _fetch_cases_mois(year, month))
    return sorted(({"Classification": s, "IncidentCount": n} for s, n in compteur.items()),
                  key=lambda l: -l["IncidentCount"])


# ---------------------------------------------------------------------------
# 5. fetch_resolution_times -- moyennes MTTA / MTTR / MTTC (heures)
# ---------------------------------------------------------------------------

def _metriques_case(case) -> dict:
    """Calcule MTTA/MTTR/MTTC effectifs (minutes) d'une case, avec les
    nuances du workbook :
      - MTTA (creation -> acquittement SLA) : High/Critical = calendaire ;
        Medium/Low = heures ouvrees.
      - MTTR (acquittement -> resolution), NET des pauses SLA (equivalent
        "attente client") : High/Critical = calendaire - pause ;
        Medium/Low = heures ouvrees - pause.
      - MTTC (creation -> resolution) : memes regles que MTTR.
      - Nuance "KpiError" NON transposee (pas de labels vus sur les cases).
    Retourne aussi les datetimes utiles et le label de gravite.
    """
    severite = _label_severite(_get_path(case, CHAMP["severity"]))
    haut = severite in ("High", "Critical")
    creation = _parse_ts(_get_path(case, CHAMP["creation"]))
    attribution = _parse_ts(_get_path(case, CHAMP["ack_completed"]))
    resolution = _parse_ts(_get_path(case, CHAMP["res_completed"]))

    # Pause SLA (minutes) : temps horloge - temps actif du timer resolution.
    goal = _get_path(case, CHAMP["res_goal"])
    due = _get_path(case, CHAMP["res_due"])
    started = _get_path(case, CHAMP["res_started"])
    completed = _get_path(case, CHAMP["res_completed"])
    pause_min = 0.0
    if all(isinstance(x, (int, float)) and x for x in (goal, due, started, completed)):
        actif_s = goal - (due - completed)
        pause_min = max(0.0, ((completed - started) - actif_s) / 60)

    mtta = mttr = mttc = None
    if creation and attribution and attribution >= creation:
        mtta = ((attribution - creation).total_seconds() / 60) if haut \
            else float(minutes_ouvrees(creation, attribution))
    if attribution and resolution and resolution >= attribution:
        brut = ((resolution - attribution).total_seconds() / 60) if haut \
            else float(minutes_ouvrees(attribution, resolution))
        mttr = max(0.0, brut - pause_min)
    if creation and resolution and resolution >= creation:
        brut = ((resolution - creation).total_seconds() / 60) if haut \
            else float(minutes_ouvrees(creation, resolution))
        mttc = max(0.0, brut - pause_min)

    return {"severite": severite, "creation": creation,
            "attribution": attribution, "resolution": resolution,
            "mtta": mtta, "mttr": mttr, "mttc": mttc,
            "statut": str(_get_path(case, CHAMP["status"]) or "").lower(),
            "titre": _get_path(case, CHAMP["name"]) or "",
            "reference": _get_path(case, CHAMP["reference"]) or case.get("id", "")}


def fetch_resolution_times(workspace_id: str, year: int, month: int,
                           tenant_id: str = None) -> dict:
    """Retourne {"MTTA": float|None, "MTTR": float|None, "MTTC": float|None}
    en HEURES -- meme contrat exact que sentinelquery.fetch_resolution_times.
    Populations identiques a l'original : MTTA sur les cases ATTRIBUEES du
    mois (tous statuts) ; MTTR/MTTC sur les cases resolues."""
    mtta_v, mttr_v, mttc_v = [], [], []
    for case in _fetch_cases_mois(year, month):
        m = _metriques_case(case)
        if m["mtta"] is not None:
            mtta_v.append(m["mtta"])
        if m["mttr"] is not None:
            mttr_v.append(m["mttr"])
        if m["mttc"] is not None:
            mttc_v.append(m["mttc"])
    moy = lambda v: (sum(v) / len(v) / 60.0) if v else None
    return {"MTTA": moy(mtta_v), "MTTR": moy(mttr_v), "MTTC": moy(mttc_v)}


# ---------------------------------------------------------------------------
# 6. fetch_sla_breaches -- liste nominative des depassements
# ---------------------------------------------------------------------------

def fetch_sla_breaches(workspace_id: str, year: int, month: int,
                       tenant_id: str = None) -> list[dict]:
    """Memes colonnes que l'original : TypeSLA ("MTTA"|"MTTR"),
    IncidentNumber, Severity, Title, CreatedTime, AttributionTime,
    ClosedTime (ISO ; "" pour une ligne MTTA non encore cloturee).
    Seuils identiques au workbook (cf SEUILS_SLA). Comme dans l'original :
    MTTA evalue sur TOUTES les cases attribuees du mois, cloturees ou non ;
    MTTR uniquement sur les cases resolues."""
    lignes = []
    for case in _fetch_cases_mois(year, month):
        m = _metriques_case(case)
        seuil_a = SEUILS_SLA["MTTA"].get(m["severite"])
        if m["mtta"] is not None and seuil_a is not None and m["mtta"] > seuil_a:
            lignes.append({
                "TypeSLA": "MTTA", "IncidentNumber": m["reference"],
                "Severity": m["severite"], "Title": m["titre"],
                "CreatedTime": _iso(m["creation"]),
                "AttributionTime": _iso(m["attribution"]),
                "ClosedTime": _iso(m["resolution"]) if m["statut"] == "closed" else "",
            })
        seuil_r = SEUILS_SLA["MTTR"].get(m["severite"])
        if m["mttr"] is not None and seuil_r is not None and m["mttr"] > seuil_r:
            lignes.append({
                "TypeSLA": "MTTR", "IncidentNumber": m["reference"],
                "Severity": m["severite"], "Title": m["titre"],
                "CreatedTime": _iso(m["creation"]),
                "AttributionTime": _iso(m["attribution"]),
                "ClosedTime": _iso(m["resolution"]),
            })
    lignes.sort(key=lambda l: (l["Severity"], l["TypeSLA"], l["CreatedTime"]))
    return lignes


# ---------------------------------------------------------------------------
# 7. fetch_workspace_name -- bandeau cosmetique
# ---------------------------------------------------------------------------

def fetch_workspace_name(workspace_id: str, tenant_id: str = None) -> str:
    """Pas de notion de workspace chez CrowdStrike : retourne le CID du
    tenant (identifiant client), lu sur une case recente -- purement
    cosmetique, "" en cas d'echec (meme contrat non-bloquant que
    l'original)."""
    try:
        ids = _query_ids("/cases/queries/cases/v1",
                         "created_timestamp:>'now-90d'", page=1, plafond=1)
        if not ids:
            return ""
        token = _get_token()
        body = json.dumps({"ids": ids}).encode("utf-8")
        cases = _http("POST", "/cases/entities/cases/v2", token=token,
                      data=body, content_type="application/json").get("resources", [])
        return cases[0].get("cid", "") if cases else ""
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# 8. fetch_mitre_tactics_stats -- couverture MITRE ATT&CK du mois
# ---------------------------------------------------------------------------

def fetch_mitre_tactics_stats(workspace_id: str, year: int, month: int,
                              tenant_id: str = None) -> list[dict]:
    """Memes colonnes que l'original : Tactic, IncidentCount, LastIncident-
    Time. Semantique adaptee : compte des ALERTES du mois par tactique
    (champ 'tactic'), tous statuts / severites -- meme esprit "vue de
    couverture sans filtre" que MITRE_TACTICS_QUERY_TEMPLATE."""
    compteur, derniere = Counter(), {}
    for a in _fetch_alertes_mois(year, month):
        t = a.get("tactic")
        if not t:
            continue
        compteur[t] += 1
        ts = _parse_ts(a.get("created_timestamp"))
        if ts and (t not in derniere or ts > derniere[t]):
            derniere[t] = ts
    lignes = [{"Tactic": t, "IncidentCount": str(n),
               "LastIncidentTime": _iso(derniere.get(t))}
              for t, n in compteur.most_common()]
    return lignes


# ---------------------------------------------------------------------------
# 9 & 10. Briques sans equivalent avec les scopes actuels -- meme contrat
# "optionnel, a degrader cote appelant" que l'original.
# ---------------------------------------------------------------------------

def fetch_active_rules_by_tactic(workspace_id: str, tenant_id: str = None) -> dict:
    """Nb de regles de correlation ACTIVES par tactique : necessite le
    scope 'Correlation Rules: Read', non attribue au client API. Leve une
    exception claire (comme l'original avec azure-mgmt-securityinsight
    absent) : a l'appelant d'afficher "N/A" sans bloquer le rapport."""
    raise RuntimeError(
        "Nombre de regles par tactique : necessite le scope API "
        "'Correlation Rules: Read' (non attribue). Ajouter ce scope au "
        "client API dans la console Falcon, ou laisser 'N/A' dans le rapport.")


def fetch_log_ingestion_costs(workspace_id: str, year: int, month: int,
                              price_per_gb: float, tenant_id: str = None) -> list[dict]:
    """Cout d'ingestion par table : donnee Log Analytics (table Usage) sans
    equivalent expose par les API Cases/Alerts. Un equivalent NGSIEM
    exigerait le scope NGSIEM Read+Write (refuse par choix de securite).
    Meme contrat : exception claire, a degrader cote appelant."""
    raise RuntimeError(
        "Cout d'ingestion : pas d'equivalent CrowdStrike avec les scopes "
        "actuels (exigerait NGSIEM Read+Write). Laisser cette slide vide "
        "ou l'alimenter depuis la console Data Connections.")


# ---------------------------------------------------------------------------
# Demo : genere le jeu de donnees complet du mois precedent
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    auj = datetime.now(timezone.utc)
    an, mo = (auj.year - 1, 12) if auj.month == 1 else (auj.year, auj.month - 1)
    print(f"[*] Jeu de donnees CrowdStrike pour {an}-{mo:02d}\n")

    jeux = {}
    jeux["cosec_incidents"] = fetch_cosec_incidents(None, an, mo)
    print(f"1. cosec_incidents          : {len(jeux['cosec_incidents'])} ligne(s)")
    jeux["typology_history"] = fetch_typology_history(None, an, mo)
    print(f"2. typology_history         : {len(jeux['typology_history'])} ligne(s)")
    jeux["severity_breakdown"] = fetch_severity_breakdown(None, an, mo)
    print(f"3. severity_breakdown       : {jeux['severity_breakdown']}")
    jeux["classification_breakdown"] = fetch_classification_breakdown(None, an, mo)
    print(f"4. classification_breakdown : {jeux['classification_breakdown']}")
    jeux["resolution_times"] = fetch_resolution_times(None, an, mo)
    print(f"5. resolution_times (h)     : {jeux['resolution_times']}")
    jeux["sla_breaches"] = fetch_sla_breaches(None, an, mo)
    print(f"6. sla_breaches             : {len(jeux['sla_breaches'])} ligne(s)")
    jeux["workspace_name"] = fetch_workspace_name(None)
    print(f"7. workspace_name (cid)     : {jeux['workspace_name']}")
    jeux["mitre_tactics_stats"] = fetch_mitre_tactics_stats(None, an, mo)
    print(f"8. mitre_tactics_stats      : {len(jeux['mitre_tactics_stats'])} tactique(s)")
    for nom in ("active_rules_by_tactic", "log_ingestion_costs"):
        print(f"9/10. {nom} : non disponible (cf docstring)")

    with open("jeu_de_donnees.json", "w", encoding="utf-8") as f:
        json.dump(jeux, f, ensure_ascii=False, indent=2, default=str)
    print("\n[*] Jeu de donnees complet ecrit dans jeu_de_donnees.json")
