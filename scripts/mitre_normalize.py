"""
Normalisation des statistiques par tactique MITRE ATT&CK pour la slide
"Dispositif de surveillance" (5e slide du template, ajoutee le 28/06/2026).

Combine 2 sources independantes :
  - sentinel_query.fetch_mitre_tactics_stats() : nombre d'incidents et
    date du dernier incident par tactique (requete KQL, mois cible).
  - sentinel_query.fetch_active_rules_by_tactic() : nombre de regles
    analytics ACTIVEES par tactique, BONUS optionnel (appel ARM, pas KQL,
    cf docstring de cette fonction).

Les deux sources renvoient leurs cles au format brut PascalCase Microsoft
(ex: "PrivilegeEscalation", "CommandAndControl") -- TACTIC_LABELS fait le
pont avec les libelles affiches dans les encadres du template (ex:
"Privilege escalation", "Command and control").
"""

from datetime import datetime, timezone, timedelta

# UTC+2 (CEST, heure d'ete Paris) -- coherent avec format_date() de
# generate_cosec.py et avec sla_normalize._parse_iso, pour que les dates
# affichees sur cette slide correspondent a la meme heure locale que
# celles affichees ailleurs dans le rapport.
_PARIS_OFFSET = timedelta(hours=2)

# Correspondance forme brute Microsoft (PascalCase, cf AttackTactic dans
# azure-mgmt-securityinsight, et AdditionalData.tactics sur SecurityIncident)
# -> libelle affiche dans les encadres du template (verifie le 28/06/2026
# en inspectant template_slide.pptx, slide "Dispositif de surveillance").
#
# Le template ne couvre que 12 des 14 tactiques MITRE ATT&CK Enterprise :
# "LateralMovement" et "Reconnaissance" n'ont PAS d'encadre correspondant
# (absentes du template tel que fourni) -- elles sont neanmoins mappees
# ici (vers un libelle plausible) plutot qu'omises, pour que les
# incidents/regles qui les concernent soient au moins identifiables dans
# les avertissements logues par generate_cosec.py (cf mitre_slide.
# fill_dispositif_surveillance_slide, valeur de retour "tactiques sans
# encadre correspondant"), au lieu de disparaitre silencieusement.
TACTIC_LABELS = {
    "InitialAccess": "Initial access",
    "Execution": "Execution",
    "Persistence": "Persistence",
    "PrivilegeEscalation": "Privilege escalation",
    "DefenseEvasion": "Defense evasion",
    "CredentialAccess": "Credential access",
    "Discovery": "Discovery",
    "LateralMovement": "Lateral movement",      # pas d'encadre dans le template
    "Collection": "Collection",
    "CommandAndControl": "Command and control",
    "Exfiltration": "Exfiltration",
    "Impact": "Impact",
    "ResourceDevelopment": "Resource Development",
    "Reconnaissance": "Reconnaissance",          # pas d'encadre dans le template
}


def _parse_iso(value: str):
    """
    Parse une date ISO 8601 (telle que renvoyee par
    sentinel_query._normalize_value) en objet datetime NAIF, converti en
    heure de Paris -- meme logique que sla_normalize._parse_iso.

    Retourne None si la valeur est vide.
    """
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value)
    except ValueError:
        return None
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc) + _PARIS_OFFSET
        dt = dt.replace(tzinfo=None)
    return dt


def normalize_tactic_label(raw_tactic: str) -> str:
    """
    Convertit une tactique brute (forme PascalCase Microsoft) en libelle
    affiche (cf TACTIC_LABELS). Une tactique non reconnue (ex: une
    nouvelle tactique MITRE future, absente de TACTIC_LABELS) est
    conservee sous son nom brut tel quel plutot que silencieusement
    ignoree.
    """
    return TACTIC_LABELS.get(raw_tactic, raw_tactic)


def build_tactic_stats(rows: list[dict]) -> dict:
    """
    Convertit les lignes brutes de sentinel_query.fetch_mitre_tactics_
    stats() (Tactic au format PascalCase Microsoft, IncidentCount,
    LastIncidentTime en chaine ISO 8601) en dict cle par LIBELLE AFFICHE,
    pret a etre consomme par mitre_slide.fill_dispositif_surveillance_
    slide().

    Retourne {libelle: {"incident_count": int, "last_incident": datetime
    naif heure de Paris, ou None}}.

    Si plusieurs tactiques brutes se rabattent sur le meme libelle affiche
    (cas non attendu en pratique, TACTIC_LABELS etant une bijection), les
    comptes sont additionnes et la date la plus recente est conservee --
    par robustesse, plutot que de laisser l'une ecraser silencieusement
    l'autre.
    """
    result = {}
    for row in rows:
        raw_tactic = row.get("Tactic", "")
        if not raw_tactic:
            continue
        label = normalize_tactic_label(raw_tactic)

        try:
            count = int(row.get("IncidentCount", 0) or 0)
        except (TypeError, ValueError):
            count = 0

        last_incident = _parse_iso(row.get("LastIncidentTime", ""))

        bucket = result.setdefault(label, {"incident_count": 0, "last_incident": None})
        bucket["incident_count"] += count
        if last_incident is not None and (bucket["last_incident"] is None
                                           or last_incident > bucket["last_incident"]):
            bucket["last_incident"] = last_incident

    return result


def build_rule_counts(raw_counts: dict) -> dict:
    """
    Convertit le dict brut retourne par sentinel_query.
    fetch_active_rules_by_tactic() (cle = tactique au format PascalCase
    Microsoft) en dict cle par LIBELLE AFFICHE, pret a etre consomme par
    mitre_slide.fill_dispositif_surveillance_slide().

    Meme logique de repli que normalize_tactic_label() pour une cle non
    reconnue, et meme logique d'addition que build_tactic_stats() en cas
    de collision de libelle.
    """
    result = {}
    for raw_tactic, count in (raw_counts or {}).items():
        label = normalize_tactic_label(raw_tactic)
        result[label] = result.get(label, 0) + count
    return result
