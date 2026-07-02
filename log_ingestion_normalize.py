"""
Normalisation des donnees de cout d'ingestion des logs pour la slide
"Plan de collecte" (6e slide du template, ajoutee le 28/06/2026).

Regroupe les lignes brutes de sentinel_query.fetch_log_ingestion_costs()
(une ligne par table individuelle) par categorie (LogType), avec le total
de chaque categorie -- reproduit la vue arborescente du workbook source
(categorie repliable, ex: "Azure Active Directory (2)", contenant ses
tables individuelles, ex: "SigninLogs", "AuditLogs").
"""

import math

_SIZE_UNITS = ["B", "kB", "MB", "GB", "TB"]


def format_size(value_gb) -> str:
    """
    Convertit une taille exprimee en Go (telle que produite par
    sentinel_query.LOG_INGESTION_QUERY_TEMPLATE) en chaine lisible avec
    l'unite adaptee (B/kB/MB/GB/TB, base 1024) -- equivalent du
    formatage "octets" automatique de la colonne de grille des Workbooks
    Azure Monitor (cf capture d'ecran fournie), reproduit ici car cette
    mise en forme est une fonctionnalite du PORTAIL, pas une donnee
    renvoyee telle quelle par la requete KQL (qui ne calcule QUE le Go).
    """
    if value_gb is None:
        return "N/A"
    value_bytes = float(value_gb) * (1024 ** 3)
    if value_bytes <= 0:
        return "0 B"
    unit_index = 0
    value = value_bytes
    while value >= 1024 and unit_index < len(_SIZE_UNITS) - 1:
        value /= 1024
        unit_index += 1
    return f"{value:.2f} {_SIZE_UNITS[unit_index]}"


def format_cost(value) -> str:
    """Formate un cout en euros (2 decimales)."""
    if value is None:
        return "N/A"
    return f"{float(value):.2f} €"


def heat_fraction(value_gb: float, min_value_gb: float, max_value_gb: float) -> float:
    """
    Position (0.0 a 1.0) de value_gb sur une echelle LOGARITHMIQUE entre
    min_value_gb (0.0, vert) et max_value_gb (1.0, rouge) -- utilisee a
    la fois pour la longueur de la barre et sa couleur (cf
    log_ingestion_slide._heat_color), de sorte que la PLUS GROSSE table
    du tableau affiche soit toujours a 1.0 (rouge, barre pleine largeur)
    et la PLUS PETITE a 0.0 (vert, barre quasi nulle).

    Echelle logarithmique plutot que lineaire : les volumes de logs
    s'etalent typiquement sur plusieurs ordres de grandeur (du ko au Go)
    -- une echelle lineaire rendrait toutes les tables sauf la plus
    grosse visuellement quasi invisibles.

    min_value_gb/max_value_gb doivent etre les bornes reelles du JEU DE
    DONNEES AFFICHE (et non une constante absolue) : un plancher fixe
    (ex: 1 ko) ecraserait artificiellement vers le haut de l'echelle
    toutes les tables de taille "normale" (Mo-Go), qui sont pourtant
    plusieurs ordres de grandeur au-dessus d'1 ko -- ce qui empecherait
    de distinguer visuellement une petite table d'une grosse au sein
    d'un MEME rapport. Recalculer les bornes a partir du jeu de donnees
    affiche garantit un etalement visuel complet (vert -> rouge) quelle
    que soit l'amplitude reelle des valeurs ce mois-ci.

    Meme principe de simplification assumee que le rendu "heatmap" du
    composant grille des Workbooks Azure Monitor, sans pretendre en
    reproduire l'algorithme exact (non documente publiquement) -- cette
    fonction est une approximation visuelle, pas un calcul certifie.
    """
    if value_gb is None or value_gb <= 0 or max_value_gb is None or max_value_gb <= 0:
        return 0.0
    if min_value_gb is None or min_value_gb <= 0 or min_value_gb >= max_value_gb:
        # Un seul niveau de grandeur present (ou bornes degenerees) :
        # rien a etaler, on retombe sur un rouge plein par defaut.
        return 1.0
    value_bytes = value_gb * (1024 ** 3)
    min_bytes = min_value_gb * (1024 ** 3)
    max_bytes = max_value_gb * (1024 ** 3)
    floor = math.log10(min_bytes)
    top = math.log10(max_bytes)
    fraction = (math.log10(max(value_bytes, min_bytes)) - floor) / (top - floor)
    return max(0.0, min(1.0, fraction))


def build_log_ingestion_groups(rows: list[dict]) -> list[dict]:
    """
    Convertit les lignes brutes de fetch_log_ingestion_costs() (LogType,
    TableName, TableSizeGB, EstimatedCost -- chaines numeriques) en
    groupes par categorie.

    Retourne une liste de dicts {category, size_gb, cost, tables} ou
    tables est une liste de {name, size_gb, cost} -- triee par taille de
    categorie decroissante, et au sein de chaque categorie par taille de
    table decroissante (meme ordre que le tri ['Estimated cost'] desc du
    workbook source, juste applique a 2 niveaux -- categorie puis table
    -- plutot qu'a une liste plate).

    Les lignes a taille nulle ou invalide sont ignorees (deja filtrees en
    amont par la requete via `where TableSizeGB > 0`, ce filtre est donc
    une securite supplementaire si la fonction est appelee avec des
    donnees d'une autre origine).
    """
    groups = {}
    for row in rows:
        category = row.get("LogType", "") or "Other"
        table_name = row.get("TableName", "")

        try:
            size_gb = float(row.get("TableSizeGB", 0) or 0)
        except (TypeError, ValueError):
            size_gb = 0.0
        try:
            cost = float(row.get("EstimatedCost", 0) or 0)
        except (TypeError, ValueError):
            cost = 0.0

        if size_gb <= 0:
            continue

        bucket = groups.setdefault(category, {"category": category, "size_gb": 0.0, "cost": 0.0, "tables": []})
        bucket["size_gb"] += size_gb
        bucket["cost"] += cost
        bucket["tables"].append({"name": table_name, "size_gb": size_gb, "cost": cost})

    result = list(groups.values())
    for g in result:
        g["tables"].sort(key=lambda t: t["size_gb"], reverse=True)
    result.sort(key=lambda g: g["size_gb"], reverse=True)
    return result


def select_groups_for_display(groups: list[dict], max_rows: int):
    """
    Selectionne les premieres categories qui tiennent dans un budget de
    `max_rows` lignes de tableau (1 ligne d'en-tete de categorie + 1
    ligne par table de cette categorie + 1 ligne "Sous-total" -- cf
    decision du 29/06/2026, log_ingestion_slide.fill_log_ingestion_slide),
    SANS JAMAIS scinder une categorie entre lignes visibles et masquees --
    evite un en-tete de categorie sans aucune table affichee dessous, et
    evite tout double comptage dans le total de synthese.

    La toute premiere categorie est TOUJOURS incluse meme si elle depasse
    seule le budget (mieux vaut une slide legerement trop pleine qu'une
    slide vide) -- cf condition `and visible`.

    Retourne (groupes_visibles, ligne_de_synthese_ou_None) : la ligne de
    synthese agrege les categories masquees ({category, size_gb, cost,
    tables: []}, tables vide car on n'affiche pas son detail), a but
    d'affichage uniquement (cf log_ingestion_slide.
    fill_log_ingestion_slide).
    """
    visible = []
    used_rows = 0
    for g in groups:
        rows_needed = 2 + len(g["tables"])
        if used_rows + rows_needed > max_rows and visible:
            break
        visible.append(g)
        used_rows += rows_needed

    hidden = groups[len(visible):]
    if not hidden:
        return visible, None

    summary = {
        "category": f"+ {len(hidden)} autre(s) catégorie(s) de logs",
        "size_gb": sum(g["size_gb"] for g in hidden),
        "cost": sum(g["cost"] for g in hidden),
        "tables": [],
    }
    return visible, summary
