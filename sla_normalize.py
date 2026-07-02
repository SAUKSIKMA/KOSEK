"""
Normalisation des incidents en depassement de SLA (MTTA/MTTR), issus de
sentinel_query.fetch_sla_breaches(), avant ecriture dans l'onglet "SLA"
de l'historique Excel (cf excel_history.write_sla_history).
"""

from datetime import datetime, timezone, timedelta

_SEVERITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "Informational": 3}
_TYPE_SLA_ORDER = {"MTTA": 0, "MTTR": 1}

# UTC+2 (CEST, heure d'ete Paris) -- coherent avec format_date() de
# generate_cosec.py, pour que les dates affichees sur la slide SLA
# correspondent a la meme heure locale que celles affichees sur les
# slides de detail par incident.
_PARIS_OFFSET = timedelta(hours=2)


def _parse_iso(value: str):
    """
    Parse une date ISO 8601 (telle que renvoyee par
    sentinel_query._normalize_value) en objet datetime NAIF, converti en
    heure de Paris.

    openpyxl n'accepte pas les datetimes "aware" (avec tzinfo) dans une
    cellule Excel -- on convertit donc explicitement en UTC+2 puis on
    retire le tzinfo, plutot que de laisser une valeur UTC brute qui
    serait incoherente avec les heures affichees ailleurs dans le rapport
    (cf format_date() dans generate_cosec.py, meme logique).

    Retourne None si la valeur est vide -- cas attendu pour ClosedTime
    sur une ligne MTTA dont l'incident n'est pas encore clôturé (cf note
    dans sentinel_query.SLA_BREACHES_QUERY_TEMPLATE).
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


def build_sla_rows(rows: list[dict], year: int, month: int) -> list[dict]:
    """
    Convertit les lignes brutes de fetch_sla_breaches() (TypeSLA,
    IncidentNumber, Severity, Title, CreatedTime, AttributionTime,
    ClosedTime -- dates en chaines ISO) en lignes pretes pour
    excel_history.write_sla_history().

    Ajoute la colonne Month : le mois CIBLE explicite du rapport
    (--year/--month), pas un mois deduit des dates des incidents -- ces
    derniers ne servent qu'a l'affichage (un incident attribue en fin de
    mois N peut tres bien etre clôture en debut de mois N+1, ca ne change
    pas le mois de rattachement du depassement, qui est celui de la
    requete executee).

    Retourne une liste de dicts {Month, TypeSLA, IncidentNumber, Severity,
    Title, CreatedTime, AttributionTime, ClosedTime} (dates en objets
    datetime ou None), triee par gravite puis type de SLA puis date de
    creation.
    """
    month_str = f"{year:04d}-{month:02d}"
    result = []
    for row in rows:
        result.append({
            "Month": month_str,
            "TypeSLA": row.get("TypeSLA", ""),
            "IncidentNumber": row.get("IncidentNumber", ""),
            "Severity": row.get("Severity", ""),
            "Title": row.get("Title", ""),
            "CreatedTime": _parse_iso(row.get("CreatedTime", "")),
            "AttributionTime": _parse_iso(row.get("AttributionTime", "")),
            "ClosedTime": _parse_iso(row.get("ClosedTime", "")),
        })

    result.sort(key=lambda r: (
        _SEVERITY_ORDER.get(r["Severity"], 99),
        _TYPE_SLA_ORDER.get(r["TypeSLA"], 99),
        r["CreatedTime"] or datetime.min,
    ))
    return result
