"""
Normalisation et agregation des resultats bruts des 3 requetes Sentinel de
la slide "Etat de la surveillance" (gravite, categorie de cloture,
MTTA/MTTR/MTTC) en UNE ligne destinee a l'onglet Surveillance de l'Excel
historique (cf excel_history.write_surveillance_history).
"""

# Severity est un enum Sentinel fixe (pas de variation de tenant a tenant),
# donc on traduit nous-meme en francais pour l'affichage sur la slide.
SEVERITY_LABELS = {
    "High": "Élevée",
    "Medium": "Moyenne",
    "Low": "Faible",
    "Informational": "Informationnelle",
}
SEVERITY_ORDER = ["High", "Medium", "Low", "Informational"]


def _classification_key(label) -> str:
    """
    Normalise une valeur Classification brute vers une des 4 cles fixes de
    l'onglet Surveillance.

    Correction du 23/06/2026 : le champ Classification de SecurityIncident
    est un enum Microsoft FIXE, toujours en anglais (TruePositive /
    FalsePositive / BenignPositive / Undetermined), quel que soit le tenant
    ou la langue de l'UI Azure -- l'hypothese precedente ("deja en francais
    dans ce tenant") etait fausse et faisait retomber TruePositive et
    FalsePositive dans Undetermined (seul BenignPositive passait, par
    coincidence, via le fallback startswith("benign")). On teste donc
    desormais les deux formes (anglais natif + francais, au cas ou un
    export/tenant les traduit), en comparant sans espaces pour couvrir les
    deux variantes avec les memes prefixes.

    Toute valeur non reconnue retombe sur "Undetermined" plutot que de
    faire echouer le pipeline (mieux vaut une categorie "Indeterminee"
    legerement surestimee qu'un crash mensuel).
    """
    text = (label or "").strip().lower().replace(" ", "")
    if text.startswith("vraipositif") or text.startswith("truepositive"):
        return "TruePositive"
    if text.startswith("fauxpositif") or text.startswith("falsepositive"):
        return "FalsePositive"
    if text.startswith("positifb") or text.startswith("benignpositive") or text.startswith("benign"):
        return "BenignPositive"
    return "Undetermined"


def build_surveillance_row(year: int, month: int, severity_rows: list,
                            classification_rows: list, resolution_times: dict) -> dict:
    """
    Assemble les 3 resultats de requete (fetch_severity_breakdown,
    fetch_classification_breakdown, fetch_resolution_times) en une ligne
    unique pour l'onglet Surveillance.

    severity_rows       : [{"Severity": "High", "IncidentCount": 10}, ...]
    classification_rows : [{"Classification": "Faux positif", "IncidentCount": 17}, ...]
    resolution_times    : {"MTTA": float|None, "MTTR": float|None, "MTTC": float|None}

    Retourne un dict {Month, Total, High, Medium, Low, Informational,
    TruePositive, FalsePositive, BenignPositive, Undetermined, MTTA, MTTR, MTTC}
    pret pour excel_history.write_surveillance_history().
    """
    month_str = f"{year:04d}-{month:02d}"

    severity_counts = {key: 0 for key in SEVERITY_ORDER}
    for row in severity_rows:
        sev = row.get("Severity")
        if sev in severity_counts:
            severity_counts[sev] += int(row.get("IncidentCount", 0) or 0)
        # Une severite hors enum standard ne devrait pas arriver (Severity
        # est un champ ferme cote Sentinel) -- on l'ignore silencieusement
        # plutot que de fausser le total avec une cle inattendue.

    classification_counts = {"TruePositive": 0, "FalsePositive": 0, "BenignPositive": 0, "Undetermined": 0}
    for row in classification_rows:
        key = _classification_key(row.get("Classification"))
        classification_counts[key] += int(row.get("IncidentCount", 0) or 0)

    total = sum(severity_counts.values())

    return {
        "Month": month_str,
        "Total": total,
        "High": severity_counts["High"],
        "Medium": severity_counts["Medium"],
        "Low": severity_counts["Low"],
        "Informational": severity_counts["Informational"],
        "TruePositive": classification_counts["TruePositive"],
        "FalsePositive": classification_counts["FalsePositive"],
        "BenignPositive": classification_counts["BenignPositive"],
        "Undetermined": classification_counts["Undetermined"],
        "MTTA": resolution_times.get("MTTA") if resolution_times.get("MTTA") is not None else 0.0,
        "MTTR": resolution_times.get("MTTR") if resolution_times.get("MTTR") is not None else 0.0,
        "MTTC": resolution_times.get("MTTC") if resolution_times.get("MTTC") is not None else 0.0,
    }
