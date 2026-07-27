"""
Normalisation des typologies d'incidents Sentinel.

Sentinel ajoute automatiquement un suffixe de comptage d'entites au Title
d'un incident lors du regroupement d'alertes (ex: "... involving one
user", "... involving multiple accounts"). Ce suffixe varie selon le
nombre d'entites correlees a l'incident, mais ne change pas la typologie
de la menace -- il faut donc le retirer avant d'agreger par typologie,
sous peine d'eclater une meme typologie en plusieurs lignes.
"""

import json
import re

# Suffixe final uniquement (ancre par $) : ne touche pas a un "involving ..."
# present au milieu du titre et porteur de sens, ex:
# "Multi-stage incident involving Privilege escalation involving multiple users"
#  -> seul le DERNIER "involving multiple users" est un suffixe de comptage.
_SUFFIX_RE = re.compile(r"\s+involving\s+(?:one|multiple|no)\s+\S+\s*$", re.IGNORECASE)

# Caracteres invisibles que Microsoft inclut parfois en fin de titre
# (zero-width space, zero-width joiner/non-joiner, BOM) -- purement
# cosmetique mais a nettoyer pour eviter des doublons "invisibles".
_TRAILING_INVISIBLE_RE = re.compile(r"[\s\u200b\u200c\u200d\ufeff]+$")


def normalize_typology(title: str) -> str:
    """Retire le suffixe de comptage d'entites et les caracteres invisibles finaux."""
    if not title:
        return title
    result = _SUFFIX_RE.sub("", title)
    result = _TRAILING_INVISIBLE_RE.sub("", result)
    return result.strip()


def aggregate_typology_rows(rows: list[dict]) -> list[dict]:
    """
    Re-agrege les lignes issues de fetch_typology_history() (Month, Title,
    AlertSources, IncidentCount) en regroupant par typologie normalisee
    au sein d'un meme mois.

    - IncidentCount : somme des comptes de toutes les variantes regroupees.
    - AlertSources   : union dedupliquee (triee) des sources, re-serialisee
      en JSON (compatible parse_json_array de generate_cosec.py).

    Retourne une liste de dicts triee par mois puis par IncidentCount
    decroissant.
    """
    buckets = {}  # (month, typologie) -> {"count": int, "sources": set}

    for row in rows:
        month = row.get("Month", "")
        typology = normalize_typology(row.get("Title", ""))

        try:
            count = int(row.get("IncidentCount", 0) or 0)
        except (TypeError, ValueError):
            count = 0

        sources_raw = row.get("AlertSources", "[]")
        try:
            sources = set(json.loads(sources_raw)) if sources_raw else set()
        except (TypeError, ValueError):
            sources = set()

        bucket = buckets.setdefault((month, typology), {"count": 0, "sources": set()})
        bucket["count"] += count
        bucket["sources"] |= sources

    result = [
        {
            "Month": month,
            "Title": typology,
            "AlertSources": json.dumps(sorted(data["sources"]), ensure_ascii=False),
            "IncidentCount": data["count"],
        }
        for (month, typology), data in buckets.items()
    ]
    result.sort(key=lambda r: (r["Month"], -r["IncidentCount"]))
    return result
