"""
Module d'execution de la requete KQL COSEC via Azure Monitor Query.

Remplace la lecture du CSV exporte manuellement : interroge directement
le workspace Log Analytics et retourne des lignes au MEME format que
celles produites par un export CSV Sentinel, pour rester compatible
avec generate_cosec.py / anonymizer.py / reformulate.py sans les modifier.

Point cle : les colonnes "liste" (Accounts, Hosts, IPs, ...) sont
serialisees en JSON ici, car parse_json_array() (dans generate_cosec.py)
attend des chaines JSON, pas des listes Python natives.
"""

import json
from datetime import datetime, timezone

from azure.identity import InteractiveBrowserCredential
from azure.monitor.query import LogsQueryClient, LogsQueryStatus
from azure.mgmt.subscription import SubscriptionClient
from azure.mgmt.resourcegraph import ResourceGraphClient
from azure.mgmt.resourcegraph.models import QueryRequest


# Colonnes qui contiennent des listes (make_set / make_set_if cote KQL) et
# qui doivent donc etre re-serialisees en JSON pour matcher le format CSV.
LIST_COLUMNS = [
    "Tactics", "Techniques", "AlertSources",
    "Accounts", "Hosts", "IPs", "SecurityGroups",
    "URLs", "Files", "Processes", "CloudApps", "Mailboxes",
]


# ---------------------------------------------------------------------------
# Cache de credential (demande du 29/06/2026)
# ---------------------------------------------------------------------------
#
# Chaque fonction fetch_* (et _resolve_workspace_resource) instanciait
# auparavant son PROPRE InteractiveBrowserCredential. Le cache de jetons
# d'azure-identity est interne a l'INSTANCE de credential, pas partage entre
# instances : creer un nouveau InteractiveBrowserCredential a chaque appel
# force donc une nouvelle authentification interactive (selection du compte
# MSP) a chaque requete Sentinel/Resource Graph/ARM du pipeline, meme au sein
# d'une seule execution du script -- jusqu'a une dizaine de fois.
#
# On met donc en cache UNE SEULE instance de credential par tenant_id (cle =
# tenant_id, ou None si non precise) pour toute la duree du processus
# Python : seul le tout premier appel declenche la popup d'authentification ;
# les appels suivants -- y compris vers une ressource/un scope different,
# ex. Log Analytics puis Resource Graph/ARM -- reutilisent la MEME instance
# et son cache de jetons, azure-identity renouvelant silencieusement
# (refresh token) tant que le compte deja authentifie reste valide pour le
# nouveau scope demande.
_credential_cache: dict = {}


def _get_credential(tenant_id: str = None) -> InteractiveBrowserCredential:
    """Retourne le InteractiveBrowserCredential partage pour ce tenant_id,
    en le creant -- et en declenchant l'authentification interactive --
    la PREMIERE fois seulement (cf commentaire ci-dessus). Les appels
    suivants avec le meme tenant_id (y compris None) reutilisent l'instance
    deja authentifiee, sans nouvelle popup."""
    if tenant_id not in _credential_cache:
        credential_kwargs = {}
        if tenant_id:
            credential_kwargs["tenant_id"] = tenant_id
        _credential_cache[tenant_id] = InteractiveBrowserCredential(**credential_kwargs)
    return _credential_cache[tenant_id]


def _month_bounds(year: int, month: int) -> tuple:
    """Retourne (debut, fin) du mois donne en UTC ; fin exclusive (1er du mois suivant)."""
    start = datetime(year, month, 1, tzinfo=timezone.utc)
    end = datetime(year + 1, 1, 1, tzinfo=timezone.utc) if month == 12 \
        else datetime(year, month + 1, 1, tzinfo=timezone.utc)
    return start, end


# Modifiee le 21/06/2026 : le filtre temporel est passe de "TimeGenerated >
# ago(30d)" (fenetre glissante depuis l'execution du script) a un mois
# calendaire explicite sur CreatedTime, pour que les slides de detail
# couvrent EXACTEMENT la meme periode que la slide d'evolution par
# typologie (cf fetch_typology_history). Le filtre sur SecurityAlert n'a
# qu'une borne basse (>= start) et pas de borne haute : un incident cree
# dans le mois cible peut continuer a recevoir des alertes apres la fin du
# mois (avant cloture), il faut donc les inclure dans le join.
#
# Modifiee le 23/06/2026 : le filtre "Severity == High" est remplace par un
# filtre sur la classification de cloture. Les slides de detail ("Focus sur
# incident") couvrent desormais tous les incidents CLOTURES MANUELLEMENT,
# c'est-a-dire dotes d'une classification explicite (TruePositive /
# FalsePositive / BenignPositive), toutes severites confondues -- et non
# plus les seuls incidents High. Un incident encore ouvert (Classification
# vide) ou classe "Undetermined" n'est pas considere comme traite
# manuellement et reste exclu. Classification est un enum Microsoft
# toujours en anglais, quel que soit le tenant (cf le bug du 23/06/2026 sur
# l'onglet Surveillance, ou l'hypothese inverse avait cause un mauvais
# classement) -- la comparaison se fait donc sur la forme anglaise, en
# minuscules pour rester robuste a une eventuelle variation de casse.
#
# Complete le 23/06/2026 : exception ajoutee pour les incidents dont le
# Title contient "software" ou "malware" (sous-chaine, insensible a la
# casse via l'operateur KQL "contains") -- ces incidents sont inclus MEME
# s'ils ont ete clotures automatiquement avec une classification
# "Undetermined" (cf demande utilisateur : ce sont des incidents jugees
# pertinents a presenter independamment du traitement manuel ou non).
# Status == "Closed" reste exige dans tous les cas (y compris pour cette
# exception) : on ne remonte pas un incident encore ouvert.
COSEC_QUERY_TEMPLATE = """
SecurityIncident
| where CreatedTime >= datetime({start}) and CreatedTime < datetime({end})
| where Status == "Closed"
| extend Classification = tostring(Classification)
| where (isnotempty(Classification) and tolower(Classification) != "undetermined")
     or (Title contains "malware" or Title contains "software")
| extend AD = todynamic(AdditionalData)
| extend
    Tactics      = AD.tactics,
    Techniques   = AD.techniques,
    AlertSources = AD.alertProductNames,
    Occurrences  = toint(AD.alertsCount)
| mv-expand AlertIds
| extend AlertId = tostring(AlertIds)
| join kind=leftouter (
    SecurityAlert
    | where TimeGenerated >= datetime({start})
    | mv-expand todynamic(Entities)
    | extend EntityType = tostring(Entities.Type)
    | extend EntityValue = case(
        EntityType == "account",           tostring(Entities.DisplayName),
        EntityType == "security-group",    tostring(Entities.DistinguishedName),
        EntityType == "mailbox",           tostring(Entities.MailboxPrimaryAddress),
        EntityType == "ip",                tostring(Entities.Address),
        EntityType == "url",               tostring(Entities.Url),
        EntityType == "host",              tostring(Entities.HostName),
        EntityType == "file",              strcat(
                                               tostring(Entities.Directory), "\\\\",
                                               tostring(Entities.Name), " [",
                                               tostring(Entities.FileHashes[0].Value), "]"
                                           ),
        EntityType == "process",           tostring(Entities.CommandLine),
        EntityType == "cloud-application", tostring(Entities.Name),
        ""
    )
    | where isnotempty(EntityValue)
    | summarize
        Accounts       = make_set_if(EntityValue, EntityType == "account"),
        Hosts          = make_set_if(EntityValue, EntityType == "host"),
        IPs            = make_set_if(EntityValue, EntityType == "ip"),
        SecurityGroups = make_set_if(EntityValue, EntityType == "security-group"),
        URLs           = make_set_if(EntityValue, EntityType == "url"),
        Files          = make_set_if(EntityValue, EntityType == "file"),
        Processes      = make_set_if(EntityValue, EntityType == "process"),
        CloudApps      = make_set_if(EntityValue, EntityType == "cloud-application"),
        Mailboxes      = make_set_if(EntityValue, EntityType == "mailbox")
        by SystemAlertId
) on $left.AlertId == $right.SystemAlertId
| summarize
    Title                 = any(Title),
    Severity              = any(Severity),
    Classification        = any(Classification),
    ClassificationReason  = any(ClassificationReason),
    ClassificationComment = any(ClassificationComment),
    CreatedTime           = min(CreatedTime),
    ClosedTime            = max(ClosedTime),
    Occurrences           = any(Occurrences),
    Tactics               = any(Tactics),
    Techniques            = any(Techniques),
    AlertSources          = any(AlertSources),
    Accounts              = make_set(tostring(Accounts)),
    Hosts                 = make_set(tostring(Hosts)),
    IPs                   = make_set(tostring(IPs)),
    SecurityGroups        = make_set(tostring(SecurityGroups)),
    URLs                  = make_set(tostring(URLs)),
    Files                 = make_set(tostring(Files)),
    Processes             = make_set(tostring(Processes)),
    CloudApps             = make_set(tostring(CloudApps)),
    Mailboxes             = make_set(tostring(Mailboxes))
    by IncidentName, ProviderIncidentId
| project
    IncidentName,
    Title,
    Severity,
    Occurrences,
    AlertSources,
    Tactics,
    Techniques,
    Accounts,
    Hosts,
    IPs,
    SecurityGroups,
    URLs,
    Files,
    Processes,
    CloudApps,
    Mailboxes,
    Classification,
    ClassificationReason,
    ClassificationComment,
    CreatedTime,
    ClosedTime
| sort by CreatedTime desc
"""


def _normalize_value(col_name: str, value) -> str:
    """
    Normalise une valeur de cellule au format attendu par parse_json_array()
    et par les helpers de generate_cosec.py (qui manipulent des str).

    Comportement reel observe du SDK azure-monitor-query (verifie via
    diagnose_columns.py sur un workspace de test) : les colonnes `dynamic`
    (make_set, make_set_if, tostring(...)) sont deja renvoyees comme des
    chaines JSON serialisees (str), pas comme des listes Python -- y
    compris la double-imbrication produite par make_set(tostring(...))
    (ex: Accounts -> '["[\\"user@domain.com\\"]"]'). C'est exactement le
    motif que parse_json_array() sait deja deplier, donc aucune
    transformation supplementaire n'est necessaire pour ces colonnes.

    - Colonnes listes (LIST_COLUMNS) : deja une str JSON -> passee telle
      quelle. Filet de securite si jamais une valeur arrive nulle/list.
    - CreatedTime / ClosedTime : datetime (avec tzinfo isodate ou standard)
      -> converti en chaine ISO 8601, deja geree par format_date().
    - Autres colonnes (Occurrences en int, etc.) : converties en str, ou
      "" si None.
    """
    if value is None:
        return "[]" if col_name in LIST_COLUMNS else ""

    if col_name in LIST_COLUMNS:
        if isinstance(value, str):
            return value
        if isinstance(value, list):
            # Filet de securite si le SDK change un jour de comportement
            # et renvoie une liste Python deja deserialisee.
            return json.dumps([str(v) for v in value])
        return json.dumps([str(value)])

    if hasattr(value, "isoformat"):  # datetime (tzinfo isodate ou standard)
        return value.isoformat()

    return str(value)


def _row_to_dict(columns: list, row) -> dict:
    """Convertit une ligne de table LogsQuery (liste positionnelle) en dict normalise."""
    return {
        col: _normalize_value(col, row[i])
        for i, col in enumerate(columns)
    }


def fetch_cosec_incidents(workspace_id: str, year: int, month: int,
                           tenant_id: str = None) -> list[dict]:
    """
    Execute la requete COSEC sur le workspace donne pour le mois calendaire
    donne, et retourne une liste de dicts, un par incident, au meme format
    que les lignes lues depuis un CSV Sentinel.

    workspace_id : GUID du workspace Log Analytics (pas le nom affiche)
    year, month  : mois calendaire cible -- DOIT correspondre au
                   year/month passe a fetch_typology_history() pour que
                   les slides de detail et la slide d'evolution couvrent
                   exactement la meme periode.
    tenant_id    : optionnel, force le tenant si le compte MSP a acces a
                   plusieurs tenants via Lighthouse (sinon resolution auto)
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = COSEC_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [_row_to_dict(columns, row) for row in table.rows]


# ---------------------------------------------------------------------------
# Historique des typologies d'incidents (slide "evolution par typologie")
# ---------------------------------------------------------------------------

# Validee avec l'utilisateur le 21/06/2026 :
#   - tous statuts, toutes severites (pas de filtre Status/Severity)
#   - IncidentCount = nombre d'INCIDENTS DISTINCTS portant ce titre
#     (dcount(IncidentName), car SecurityIncident contient une ligne par
#     mise a jour, pas une ligne par incident -- meme piege que dans
#     COSEC_QUERY)
#   - une ligne par (Title, Month) pour permettre le suivi mois par mois
TYPOLOGY_QUERY_TEMPLATE = """
SecurityIncident
| where CreatedTime >= datetime({start}) and CreatedTime < datetime({end})
| extend AD = todynamic(AdditionalData)
| extend AlertSources = AD.alertProductNames
| extend Month = format_datetime(startofmonth(CreatedTime), 'yyyy-MM')
| mv-expand AlertSources
| extend AlertSource = tostring(AlertSources)
| summarize
    IncidentCount = dcount(IncidentName),
    AlertSources  = make_set(AlertSource)
    by Title, Month
| project Month, Title, AlertSources, IncidentCount
| sort by Month asc, IncidentCount desc
"""


def fetch_typology_history(workspace_id: str, year: int, month: int,
                            tenant_id: str = None) -> list[dict]:
    """
    Recupere l'historique des typologies d'incidents (tous statuts, toutes
    severites) pour le mois donne, agrege par (Title, Month).

    Retourne une liste de dicts : Month ("yyyy-MM"), Title, AlertSources
    (chaine JSON, compatible parse_json_array), IncidentCount (int, nombre
    d'incidents distincts).

    Point d'attention sur le timespan transmis au SDK : SecurityIncident
    est mis a jour en continu (changement de statut, classification...),
    et la colonne par defaut filtree par le SDK (TimeGenerated) correspond
    a la date de la DERNIERE mise a jour de la ligne, pas a sa creation.
    Un incident cree en mai mais cloture en juin a donc un TimeGenerated
    de juin. On transmet donc un timespan large (debut du mois -> maintenant)
    au SDK pour ne rater aucune ligne, et c'est le filtre explicite sur
    CreatedTime dans le KQL qui borne reellement le mois.
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = TYPOLOGY_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete typologie sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [_row_to_dict(columns, row) for row in table.rows]


# ---------------------------------------------------------------------------
# Slide "Etat de la surveillance" (repartition gravite / cloture / MTTA-MTTR-MTTC)
# ---------------------------------------------------------------------------
#
# Validee avec l'utilisateur le 22/06/2026 :
#   - Repartition par gravite et par categorie de cloture : tous statuts,
#     meme perimetre temporel (CreatedTime dans le mois calendaire) que la
#     requete typologie, pour rester sur EXACTEMENT la meme periode.
#   - Categorie de cloture : Classification SEULE (4 categories fixes :
#     Vrai positif / Faux positif / Positif benin / Indetermine), pas de
#     detail par ClassificationReason. Les incidents sans Classification
#     (ex: encore ouverts) sont rattaches a "Indetermine".
#   - MTTA/MTTR/MTTC : requete unique adaptee des 3 requetes du workbook
#     existant (Top 3 MTTA / Top 3 MTTR / Top 3 MTTC), en conservant
#     EXACTEMENT la meme nuance de calcul :
#       * tag "attente client" deduit du MTTR/MTTC (pas du MTTA)
#       * High -> duree calendaire ; Medium/Low -> duree en heures ouvrees
#         (lun-ven 8h-19h, cf WorkingMinutes)
#       * label KpiError -> conserve la duree brute calendaire (incident
#         exclu du calcul de performance SOC)
#     Seule difference avec le workbook : les parametres interactifs du
#     workbook ({Period}, {Severity}, {Tactics}, {Owner}, {Product}) sont
#     retires (pas de filtre -> perimetre complet du mois), car ce rapport
#     mensuel n'est pas un workbook filtrable. Le filtre societe sur le
#     titre de bruit ("Data log source is not sending logs") est conserve.

# Pour le decompte par gravite/cloture, une ligne SecurityIncident par mise
# a jour -> on deduplique par IncidentNumber via arg_max(TimeGenerated, *)
# AVANT de compter, donc un simple count() par la suite est deja correct
# (pas besoin de dcount ici, contrairement a TYPOLOGY_QUERY_TEMPLATE qui
# compte par Title sans dedupliquer ligne par ligne au prealable).
SEVERITY_QUERY_TEMPLATE = """
SecurityIncident
| where CreatedTime >= datetime({start}) and CreatedTime < datetime({end})
| where Title <> "Data log source is not sending logs"
| summarize arg_max(TimeGenerated, *) by IncidentNumber
| summarize IncidentCount = count() by Severity
| project Severity, IncidentCount
| sort by IncidentCount desc
"""

CLASSIFICATION_QUERY_TEMPLATE = """
SecurityIncident
| where CreatedTime >= datetime({start}) and CreatedTime < datetime({end})
| where Title <> "Data log source is not sending logs"
| summarize arg_max(TimeGenerated, *) by IncidentNumber
| extend Classification = iff(isempty(Classification), "Indéterminé", tostring(Classification))
| summarize IncidentCount = count() by Classification
| project Classification, IncidentCount
| sort by IncidentCount desc
"""

# Liste de tags "attente client" -- identique au workbook (Top 3 MTTR / MTTC).
_RESOLUTION_TAG_LIST = (
    "dynamic(['attenteclient', 'attente', 'attente client', 'attente-client',"
    "'attente_client','attenteclients','attente-clients','attente_clients', "
    "'en attente client'])"
)

RESOLUTION_TIMES_QUERY_TEMPLATE = """
let TagList = """ + _RESOLUTION_TAG_LIST + """;
let MonthStart = datetime({start});
let MonthEnd = datetime({end});

// Premiere attribution reelle -- cf Top 3 MTTA pour le detail du parsing
// de Owner (filet contre {{"objectId":null,...}} qui passe un filtre naif).
let FirstAttribution = SecurityIncident
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| summarize AttributionTime = min(TimeGenerated) by IncidentNumber;

// Etat final des incidents attribues, clôtures -- sert au bornage des tags
// pour MTTR (perimetre [AttributionTime, ClosedTime]) et MTTC (perimetre
// [CreatedTime, ClosedTime]).
let FinalStateClosed = SecurityIncident
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| summarize arg_max(TimeGenerated, *) by IncidentNumber
| where Status == "Closed";

// Transitions de tag "attente client" -- identique au workbook.
let Transitions = SecurityIncident
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| extend HasTag = Labels has_any (TagList)
| project IncidentNumber, TimeGenerated, HasTag
| order by IncidentNumber asc, TimeGenerated asc
| serialize
| extend PrevHasTag = prev(HasTag)
| extend PrevIncident = prev(IncidentNumber)
| where IncidentNumber == PrevIncident
| where HasTag != PrevHasTag;

let TagStarts = Transitions | where HasTag == true  | project IncidentNumber, TagStartTime = TimeGenerated;
let TagEnds   = Transitions | where HasTag == false | project IncidentNumber, TagEndTime = TimeGenerated;

// Population de base du mois (tous statuts, incidents attribues) -- sert
// au MTTA. Le filtre AssignedTo non vide est applique ligne par ligne
// AVANT le dedup (comme dans le workbook) : on garde le dernier snapshot
// qui avait deja une attribution, pas le dernier snapshot tout court.
let FilteredIncidents = SecurityIncident
| where CreatedTime >= MonthStart and CreatedTime < MonthEnd
| where Title <> "Data log source is not sending logs"
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| summarize arg_max(TimeGenerated, *) by IncidentNumber;

// Sous-ensemble clôture du mois -- sert au MTTR/MTTC.
let FilteredIncidentsClosed = FilteredIncidents | where Status == "Closed";

// --- MTTA : duree de prise en charge (CreatedTime -> AttributionTime) ---
let WorkingMinutesMTTA = FilteredIncidents
| join kind=inner FirstAttribution on IncidentNumber
| mv-expand Hour = range(bin(CreatedTime, 1h), bin(AttributionTime, 1h), 1h) to typeof(datetime)
| where dayofweek(Hour) between (1d .. 5d)
| where hourofday(Hour) between (8 .. 18)
| extend HourStart = iff(Hour < CreatedTime, CreatedTime, Hour)
| extend HourEnd   = iff(Hour + 1h > AttributionTime, AttributionTime, Hour + 1h)
| extend Working   = iff(HourEnd > HourStart, datetime_diff('minute', HourEnd, HourStart), 0)
| summarize MTTA_Working_minutes = sum(Working) by IncidentNumber;

let AvgMTTA = toscalar(
    FilteredIncidents
    | join kind=inner FirstAttribution on IncidentNumber
    | join kind=leftouter WorkingMinutesMTTA on IncidentNumber
    | extend MTTA_minutes = datetime_diff('minute', AttributionTime, CreatedTime)
    | where MTTA_minutes >= 0
    | extend MTTA_effective = case(Severity == "High", MTTA_minutes, MTTA_Working_minutes)
    | extend TimeToAcknowledge = MTTA_effective / 60.0
    | where TimeToAcknowledge >= 0
    | summarize avg(TimeToAcknowledge)
);

// --- MTTR : duree de resolution (AttributionTime -> ClosedTime), nette
//     de l'attente client, bornee a [AttributionTime, ClosedTime] ---
let WaitDurationsMTTR = TagStarts
| join kind=leftouter (TagEnds) on IncidentNumber
| summarize TagEndTime = minif(TagEndTime, TagEndTime > TagStartTime) by IncidentNumber, TagStartTime
| join kind=leftouter FinalStateClosed on IncidentNumber
| join kind=leftouter FirstAttribution on IncidentNumber
| extend TagEndTime = iff(isempty(TagEndTime), ClosedTime, TagEndTime)
| extend
    TagStartTime = iff(TagStartTime < AttributionTime, AttributionTime, TagStartTime),
    TagEndTime   = iff(TagEndTime   > ClosedTime,      ClosedTime,      TagEndTime)
| where TagEndTime > TagStartTime
| summarize WaitDuration_minutes = sum(datetime_diff('minute', TagEndTime, TagStartTime)) by IncidentNumber;

let WorkingMinutesMTTR = FilteredIncidentsClosed
| join kind=inner FirstAttribution on IncidentNumber
| mv-expand Hour = range(bin(AttributionTime, 1h), bin(ClosedTime, 1h), 1h) to typeof(datetime)
| where dayofweek(Hour) between (1d .. 5d)
| where hourofday(Hour) between (8 .. 18)
| extend HourStart = iff(Hour < AttributionTime, AttributionTime, Hour)
| extend HourEnd   = iff(Hour + 1h > ClosedTime, ClosedTime, Hour + 1h)
| extend Working   = iff(HourEnd > HourStart, datetime_diff('minute', HourEnd, HourStart), 0)
| summarize MTTR_Working_minutes = sum(Working) by IncidentNumber;

let AvgMTTR = toscalar(
    FilteredIncidentsClosed
    | join kind=inner FirstAttribution on IncidentNumber
    | join kind=leftouter WaitDurationsMTTR on IncidentNumber
    | join kind=leftouter WorkingMinutesMTTR on IncidentNumber
    | extend WaitDuration_minutes = iff(isempty(WaitDuration_minutes), 0, WaitDuration_minutes)
    | extend
        MTTA_minutes = datetime_diff('minute', AttributionTime, CreatedTime),
        MTTR_minutes = datetime_diff('minute', ClosedTime, AttributionTime)
    | extend MTTR_net_minutes = MTTR_minutes - WaitDuration_minutes
    | where MTTA_minutes >= 0 and MTTR_net_minutes >= 0
    | extend MTTR_effective = case(Severity == "High", MTTR_net_minutes, MTTR_Working_minutes - WaitDuration_minutes)
    | extend TimeToResolve = iff(Labels has "KpiError", MTTR_minutes / 60.0, MTTR_effective / 60.0)
    | where TimeToResolve >= 0
    | summarize avg(TimeToResolve)
);

// --- MTTC : duree totale de cloture (CreatedTime -> ClosedTime), nette
//     de l'attente client, bornee a [CreatedTime, ClosedTime] ---
let WaitDurationsMTTC = TagStarts
| join kind=leftouter (TagEnds) on IncidentNumber
| summarize TagEndTime = minif(TagEndTime, TagEndTime > TagStartTime) by IncidentNumber, TagStartTime
| join kind=leftouter FinalStateClosed on IncidentNumber
| extend TagEndTime = iff(isempty(TagEndTime), ClosedTime, TagEndTime)
| extend
    TagStartTime = iff(TagStartTime < CreatedTime, CreatedTime, TagStartTime),
    TagEndTime   = iff(TagEndTime   > ClosedTime,  ClosedTime,  TagEndTime)
| where TagEndTime > TagStartTime
| summarize WaitDuration_minutes = sum(datetime_diff('minute', TagEndTime, TagStartTime)) by IncidentNumber;

let WorkingMinutesMTTC = FilteredIncidentsClosed
| mv-expand Hour = range(bin(CreatedTime, 1h), bin(ClosedTime, 1h), 1h) to typeof(datetime)
| where dayofweek(Hour) between (1d .. 5d)
| where hourofday(Hour) between (8 .. 18)
| extend HourStart = iff(Hour < CreatedTime, CreatedTime, Hour)
| extend HourEnd   = iff(Hour + 1h > ClosedTime, ClosedTime, Hour + 1h)
| extend Working   = iff(HourEnd > HourStart, datetime_diff('minute', HourEnd, HourStart), 0)
| summarize MTTC_Working_minutes = sum(Working) by IncidentNumber;

let AvgMTTC = toscalar(
    FilteredIncidentsClosed
    | join kind=leftouter WaitDurationsMTTC on IncidentNumber
    | join kind=leftouter WorkingMinutesMTTC on IncidentNumber
    | extend WaitDuration_minutes = iff(isempty(WaitDuration_minutes), 0, WaitDuration_minutes)
    | extend MTTC_minutes = datetime_diff('minute', ClosedTime, CreatedTime)
    | extend MTTC_net_minutes = MTTC_minutes - WaitDuration_minutes
    | where MTTC_minutes >= 0 and MTTC_net_minutes >= 0
    | extend MTTC_effective = case(Severity == "High", MTTC_net_minutes, MTTC_Working_minutes - WaitDuration_minutes)
    | extend TimeToClosure = iff(Labels has "KpiError", MTTC_minutes / 60.0, MTTC_effective / 60.0)
    | where TimeToClosure >= 0
    | summarize avg(TimeToClosure)
);

print AvgMTTA = AvgMTTA, AvgMTTR = AvgMTTR, AvgMTTC = AvgMTTC
"""


def fetch_severity_breakdown(workspace_id: str, year: int, month: int,
                              tenant_id: str = None) -> list[dict]:
    """
    Retourne le nombre d'incidents distincts par gravite (Severity) pour le
    mois donne (tous statuts). Liste de dicts {Severity, IncidentCount}.
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = SEVERITY_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete gravite sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [dict(zip(columns, row)) for row in table.rows]


def fetch_classification_breakdown(workspace_id: str, year: int, month: int,
                                    tenant_id: str = None) -> list[dict]:
    """
    Retourne le nombre d'incidents distincts par categorie de cloture
    (Classification seule -- 4 categories fixes, "Indéterminé" si vide)
    pour le mois donne. Liste de dicts {Classification, IncidentCount}.
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = CLASSIFICATION_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete cloture sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [dict(zip(columns, row)) for row in table.rows]


def fetch_resolution_times(workspace_id: str, year: int, month: int,
                            tenant_id: str = None) -> dict:
    """
    Retourne le MTTA/MTTR/MTTC moyen (en heures) pour le mois donne, avec
    la meme nuance de calcul que le workbook (tag attente client, heures
    ouvrees pour Medium/Low, KpiError) mais sans filtre interactif.

    Retourne {"MTTA": float|None, "MTTR": float|None, "MTTC": float|None}.
    None si aucun incident ne remplit les conditions du calcul (ex: aucun
    incident attribue dans le mois).
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = RESOLUTION_TIMES_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return {"MTTA": None, "MTTR": None, "MTTC": None}
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete MTTA/MTTR/MTTC sur le workspace {workspace_id} : {response}")

    row = dict(zip(table.columns, table.rows[0]))
    return {
        "MTTA": float(row["AvgMTTA"]) if row.get("AvgMTTA") is not None else None,
        "MTTR": float(row["AvgMTTR"]) if row.get("AvgMTTR") is not None else None,
        "MTTC": float(row["AvgMTTC"]) if row.get("AvgMTTC") is not None else None,
    }


# ---------------------------------------------------------------------------
# Slide "Dépassement des SLA" (liste nominative des incidents en
# depassement MTTA et/ou MTTR)
# ---------------------------------------------------------------------------
#
# Adaptee le 23/06/2026 depuis la requete de detection de depassements SLA
# du workbook existant. Nuances IMPERATIVEMENT conservees a l'identique :
#   - logique du tag "attente client" (TagList, Transitions, TagStarts/
#     TagEnds, bornage des plages dans le perimetre de la metrique) ;
#   - logique HNO/NO : High -> duree calendaire ; Medium/Low -> duree en
#     heures ouvrees (lun-ven 8h-19h, cf mv-expand + dayofweek/hourofday) ;
#   - seuils SLA par gravite (MTTA 30/45/120 min, MTTR 1440/2880/2880 min).
#
# Difference VOLONTAIRE avec le workbook (cf echange du 23/06/2026) : le
# workbook ne considere que les incidents CLOTURES (FinalState filtre sur
# Status == "Closed") pour les DEUX metriques. Ici, seul le depassement
# MTTR exige une cloture (necessaire pour calculer ClosedTime - Attribution-
# Time). Le depassement MTTA est evalue sur TOUS les incidents attribues du
# mois, clotures ou non -- un incident encore ouvert dont la prise en charge
# a deja depasse le seuil est une information utile au mois courant, pas
# seulement un constat a posteriori. C'est ce qui explique que la colonne
# Cloture puisse etre vide sur une ligne MTTA (incident pas encore cloture)
# alors qu'elle est toujours renseignee sur une ligne MTTR.
#
# Pas de nuance "Labels has 'KpiError'" ici : cette regle est specifique au
# calcul des MOYENNES (RESOLUTION_TIMES_QUERY_TEMPLATE ci-dessus) et n'existe
# pas dans la requete de detection de depassements du workbook -- on ne l'a
# donc pas reintroduite, conformement a la consigne de garder EXACTEMENT la
# nuance d'origine.
SLA_BREACHES_QUERY_TEMPLATE = """
let TagList = """ + _RESOLUTION_TAG_LIST + """;
let MonthStart = datetime({start});
let MonthEnd = datetime({end});

// Premiere attribution reelle -- identique a RESOLUTION_TIMES_QUERY_TEMPLATE.
let FirstAttribution = SecurityIncident
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| summarize AttributionTime = min(TimeGenerated) by IncidentNumber;

// Etat final des incidents attribues, clôtures -- sert uniquement au
// bornage des plages d'attente client pour le MTTR (perimetre
// [AttributionTime, ClosedTime]), comme dans RESOLUTION_TIMES_QUERY_TEMPLATE.
let FinalStateClosed = SecurityIncident
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| summarize arg_max(TimeGenerated, *) by IncidentNumber
| where Status == "Closed";

// Transitions de tag "attente client" -- identique au workbook.
let Transitions = SecurityIncident
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| extend HasTag = Labels has_any (TagList)
| project IncidentNumber, TimeGenerated, HasTag
| order by IncidentNumber asc, TimeGenerated asc
| serialize
| extend PrevHasTag = prev(HasTag)
| extend PrevIncident = prev(IncidentNumber)
| where IncidentNumber == PrevIncident
| where HasTag != PrevHasTag;

let TagStarts = Transitions | where HasTag == true  | project IncidentNumber, TagStartTime = TimeGenerated;
let TagEnds   = Transitions | where HasTag == false | project IncidentNumber, TagEndTime = TimeGenerated;

// Population MTTA : tous statuts, incidents attribues, crees dans le mois
// cible (cf note ci-dessus : on ne restreint pas aux incidents clôtures).
let FilteredIncidents = SecurityIncident
| where CreatedTime >= MonthStart and CreatedTime < MonthEnd
| where Title <> "Data log source is not sending logs"
| extend OwnerParsed = todynamic(Owner)
| extend AssignedTo = tostring(OwnerParsed.assignedTo)
| where isnotempty(AssignedTo) and AssignedTo != "null"
| summarize arg_max(TimeGenerated, *) by IncidentNumber;

// Sous-ensemble clôture du mois -- seule population valable pour le MTTR.
let FilteredIncidentsClosed = FilteredIncidents | where Status == "Closed";

// --- MTTA effectif (meme calcul que RESOLUTION_TIMES_QUERY_TEMPLATE) ---
let WorkingMinutesMTTA = FilteredIncidents
| join kind=inner FirstAttribution on IncidentNumber
| mv-expand Hour = range(bin(CreatedTime, 1h), bin(AttributionTime, 1h), 1h) to typeof(datetime)
| where dayofweek(Hour) between (1d .. 5d)
| where hourofday(Hour) between (8 .. 18)
| extend HourStart = iff(Hour < CreatedTime, CreatedTime, Hour)
| extend HourEnd   = iff(Hour + 1h > AttributionTime, AttributionTime, Hour + 1h)
| extend Working   = iff(HourEnd > HourStart, datetime_diff('minute', HourEnd, HourStart), 0)
| summarize MTTA_Working_minutes = sum(Working) by IncidentNumber;

let MTTABreaches = FilteredIncidents
| join kind=inner FirstAttribution on IncidentNumber
| join kind=leftouter WorkingMinutesMTTA on IncidentNumber
| extend MTTA_minutes = datetime_diff('minute', AttributionTime, CreatedTime)
| where MTTA_minutes >= 0
| extend MTTA_effective = case(Severity == "High", MTTA_minutes, MTTA_Working_minutes)
| extend SLA_MTTA_minutes = case(
    Severity == "High",   30,
    Severity == "Medium", 45,
    Severity == "Low",    120,
    int(null))
| where isnotnull(SLA_MTTA_minutes) and MTTA_effective > SLA_MTTA_minutes
| project
    TypeSLA         = "MTTA",
    IncidentNumber,
    Severity,
    Title,
    CreatedTime,
    AttributionTime,
    ClosedTime      = iff(Status == "Closed", ClosedTime, datetime(null));

// --- MTTR effectif (meme calcul que RESOLUTION_TIMES_QUERY_TEMPLATE,
//     attente client nette, perimetre [AttributionTime, ClosedTime]) ---
let WaitDurationsMTTR = TagStarts
| join kind=leftouter (TagEnds) on IncidentNumber
| summarize TagEndTime = minif(TagEndTime, TagEndTime > TagStartTime) by IncidentNumber, TagStartTime
| join kind=leftouter FinalStateClosed on IncidentNumber
| join kind=leftouter FirstAttribution on IncidentNumber
| extend TagEndTime = iff(isempty(TagEndTime), ClosedTime, TagEndTime)
| extend
    TagStartTime = iff(TagStartTime < AttributionTime, AttributionTime, TagStartTime),
    TagEndTime   = iff(TagEndTime   > ClosedTime,      ClosedTime,      TagEndTime)
| where TagEndTime > TagStartTime
| summarize WaitDuration_minutes = sum(datetime_diff('minute', TagEndTime, TagStartTime)) by IncidentNumber;

let WorkingMinutesMTTR = FilteredIncidentsClosed
| join kind=inner FirstAttribution on IncidentNumber
| mv-expand Hour = range(bin(AttributionTime, 1h), bin(ClosedTime, 1h), 1h) to typeof(datetime)
| where dayofweek(Hour) between (1d .. 5d)
| where hourofday(Hour) between (8 .. 18)
| extend HourStart = iff(Hour < AttributionTime, AttributionTime, Hour)
| extend HourEnd   = iff(Hour + 1h > ClosedTime, ClosedTime, Hour + 1h)
| extend Working   = iff(HourEnd > HourStart, datetime_diff('minute', HourEnd, HourStart), 0)
| summarize MTTR_Working_minutes = sum(Working) by IncidentNumber;

let MTTRBreaches = FilteredIncidentsClosed
| join kind=inner FirstAttribution on IncidentNumber
| join kind=leftouter WaitDurationsMTTR on IncidentNumber
| join kind=leftouter WorkingMinutesMTTR on IncidentNumber
| extend WaitDuration_minutes = iff(isempty(WaitDuration_minutes), 0, WaitDuration_minutes)
| extend MTTR_minutes = datetime_diff('minute', ClosedTime, AttributionTime)
| extend MTTR_net_minutes = MTTR_minutes - WaitDuration_minutes
| where MTTR_net_minutes >= 0
| extend MTTR_effective = case(Severity == "High", MTTR_net_minutes, MTTR_Working_minutes - WaitDuration_minutes)
| extend SLA_MTTR_minutes = case(
    Severity == "High",   1440,
    Severity == "Medium", 2880,
    Severity == "Low",    2880,
    int(null))
| where isnotnull(SLA_MTTR_minutes) and MTTR_effective > SLA_MTTR_minutes
| project
    TypeSLA = "MTTR",
    IncidentNumber,
    Severity,
    Title,
    CreatedTime,
    AttributionTime,
    ClosedTime;

MTTABreaches
| union MTTRBreaches
| sort by Severity asc, TypeSLA asc, CreatedTime asc
"""


def fetch_sla_breaches(workspace_id: str, year: int, month: int,
                        tenant_id: str = None) -> list[dict]:
    """
    Recupere la liste nominative des incidents en depassement de SLA
    (MTTA et/ou MTTR, un incident pouvant apparaitre dans les deux
    listes s'il depasse les deux seuils) pour le mois calendaire donne.

    Retourne une liste de dicts : TypeSLA ("MTTA"|"MTTR"), IncidentNumber,
    Severity, Title, CreatedTime, AttributionTime, ClosedTime (chaines ISO
    8601, ClosedTime == "" si l'incident MTTA n'est pas encore clôturé).
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = SLA_BREACHES_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete depassements SLA sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [_row_to_dict(columns, row) for row in table.rows]


# ---------------------------------------------------------------------------
# Resolution de la ressource ARM du workspace (bandeau "Confidentiel –
# COSEC - <client>", et BONUS "Nombre de regles" de la slide "Dispositif
# de surveillance")
# ---------------------------------------------------------------------------
#
# Ajoute le 28/06/2026. Le Workspace ID passe en parametre du script
# (--workspace-id) est le GUID interne Log Analytics (customerId), utilise
# pour interroger les donnees via LogsQueryClient -- ce N'EST PAS le nom de
# la ressource Azure (ex: "law-prd-sentinel-emh"), que l'API Log Analytics
# Query elle-meme n'expose pas. Pour retrouver ce nom, on interroge Azure
# Resource Graph (recherche transverse sur les ressources
# Microsoft.OperationalInsights/workspaces, filtree sur la propriete
# customerId), apres avoir enumere les souscriptions accessibles au compte
# authentifie -- Resource Graph exige une liste explicite de souscriptions
# a interroger, il n'y a pas de mode "toutes souscriptions accessibles"
# implicite.
#
# Factorisee le 28/06/2026 : la resolution ARM (nom + groupe de ressources
# + abonnement) est desormais isolee dans _resolve_workspace_resource(),
# reutilisee par fetch_workspace_name() (qui ne se servait jusqu'ici que
# du nom) ET par fetch_active_rules_by_tactic() ci-dessous (BONUS de la
# slide "Dispositif de surveillance"), qui a en plus besoin du groupe de
# ressources et de l'abonnement pour interroger l'API de gestion Sentinel
# (azure-mgmt-securityinsight) -- l'API Log Analytics Query ne les expose
# pas non plus.

def _resolve_workspace_resource(workspace_id: str, tenant_id: str = None) -> dict:
    """
    Retrouve la ressource ARM Azure du workspace Log Analytics dont le
    Workspace ID (customerId, GUID) est donne -- cf note ci-dessus.

    Retourne {"name", "resourceGroup", "subscriptionId"}, ou None si
    aucune ressource ne correspond (GUID invalide, ou compte sans acces a
    la souscription qui l'heberge).
    """
    credential = _get_credential(tenant_id)

    sub_client = SubscriptionClient(credential)
    subscription_ids = [sub.subscription_id for sub in sub_client.subscriptions.list()]
    if not subscription_ids:
        return None

    graph_client = ResourceGraphClient(credential)
    query = (
        "Resources"
        " | where type =~ 'microsoft.operationalinsights/workspaces'"
        f" | where properties.customerId =~ '{workspace_id}'"
        " | project name, resourceGroup, subscriptionId"
        " | limit 1"
    )
    request = QueryRequest(subscriptions=subscription_ids, query=query)
    response = graph_client.resources(request)

    if not response.data:
        return None
    data = response.data[0]
    return {
        "name": data.get("name", ""),
        "resourceGroup": data.get("resourceGroup", ""),
        "subscriptionId": data.get("subscriptionId", ""),
    }


def fetch_workspace_name(workspace_id: str, tenant_id: str = None) -> str:
    """
    Retrouve le nom de la ressource Azure du workspace Log Analytics dont
    le Workspace ID (customerId, GUID) est donne -- cf note ci-dessus.

    Retourne le nom brut (ex: "law-prd-sentinel-emh"), ou une chaine vide
    si aucune ressource ne correspond (GUID invalide, ou compte sans acces
    a la souscription qui l'heberge) -- ne leve PAS d'exception dans ce
    cas : le nom du workspace n'est qu'une information cosmetique pour le
    bandeau "Confidentiel" du rapport (cf generate_cosec.
    update_confidential_banner), pas une donnee bloquante comme les
    requetes d'incidents -- a l'appelant de logger un avertissement plutot
    que d'interrompre la generation du rapport sur cet echec.
    """
    resource = _resolve_workspace_resource(workspace_id, tenant_id=tenant_id)
    return resource["name"] if resource else ""


# ---------------------------------------------------------------------------
# Slide "Dispositif de surveillance" (couverture par tactique MITRE ATT&CK)
# ---------------------------------------------------------------------------
#
# Ajoutee le 28/06/2026. Contrairement a COSEC_QUERY_TEMPLATE (qui ne
# couvre que les incidents CLOTURES MANUELLEMENT, cf note plus haut), cette
# requete porte sur TOUS les incidents du mois calendaire SANS AUCUN AUTRE
# FILTRE (ni Status, ni Classification) -- validee avec l'utilisateur :
# l'objectif est une vue de couverture (combien d'incidents ont declenche
# chaque tactique), pas une liste d'incidents traites.
#
# Une tactique etant un array (AD.tactics, meme champ que COSEC_QUERY_
# TEMPLATE) et un incident pouvant en couvrir plusieurs a la fois, on
# mv-expand AVANT d'agreger -- meme principe que TYPOLOGY_QUERY_TEMPLATE
# (mv-expand AlertSources). IncidentCount utilise dcount(IncidentName) et
# non count(), pour la meme raison que partout ailleurs dans ce module :
# SecurityIncident contient une ligne par MISE A JOUR, pas par incident.
# LastIncidentTime (max(CreatedTime)) n'est pas affecte par cette
# duplication : CreatedTime est fixe a la creation de l'incident et ne
# varie pas entre les lignes d'un meme IncidentName.
MITRE_TACTICS_QUERY_TEMPLATE = """
SecurityIncident
| where CreatedTime >= datetime({start}) and CreatedTime < datetime({end})
| extend AD = todynamic(AdditionalData)
| extend Tactics = AD.tactics
| mv-expand Tactics
| extend Tactic = tostring(Tactics)
| where isnotempty(Tactic)
| summarize IncidentCount = dcount(IncidentName), LastIncidentTime = max(CreatedTime) by Tactic
| project Tactic, IncidentCount, LastIncidentTime
| sort by IncidentCount desc
"""


def fetch_mitre_tactics_stats(workspace_id: str, year: int, month: int,
                               tenant_id: str = None) -> list[dict]:
    """
    Recupere, pour le mois calendaire donne, le nombre d'incidents
    distincts ayant declenche chaque tactique MITRE ATT&CK (tous statuts,
    toutes severites, toutes classifications -- cf note ci-dessus) ainsi
    que la date du DERNIER incident ayant declenche cette tactique.

    Retourne une liste de dicts : Tactic (forme PascalCase Microsoft, ex:
    "PrivilegeEscalation" -- a faire passer par mitre_normalize.
    TACTIC_LABELS avant affichage), IncidentCount (chaine numerique),
    LastIncidentTime (chaine ISO 8601).

    Meme remarque que fetch_typology_history() sur le timespan transmis
    au SDK (large, du debut du mois a maintenant) : le filtre reel sur le
    mois calendaire est celui, explicite, du KQL sur CreatedTime.
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = MITRE_TACTICS_QUERY_TEMPLATE.format(start=start.isoformat(), end=end.isoformat())

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete MITRE ATT&CK sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [_row_to_dict(columns, row) for row in table.rows]


def fetch_active_rules_by_tactic(workspace_id: str, tenant_id: str = None) -> dict:
    """
    BONUS (cf demande utilisateur, slide "Dispositif de surveillance") :
    nombre de regles analytics ACTIVEES (enabled=true) par tactique MITRE
    ATT&CK, a la date d'execution du script -- l'information affichee dans
    le menu "MITRE ATT&CK" du portail Sentinel.

    Cette information n'est PAS exposee par les tables Log Analytics
    interrogeables en KQL (SecurityIncident / SecurityAlert) : les regles
    analytics (ressources ARM Microsoft.SecurityInsights/alertRules) sont
    exposees par l'API de gestion Sentinel (package azure-mgmt-
    securityinsight), pas par le moteur de requete Log Analytics. On
    reutilise donc la resolution ARM du workspace (_resolve_workspace_
    resource(), cf fetch_workspace_name) pour obtenir l'abonnement et le
    groupe de ressources necessaires a cet appel.

    Import differe (et non en tete de module) : azure-mgmt-securityinsight
    est une dependance OPTIONNELLE, requise uniquement pour ce bonus --
    son absence ne doit pas empecher le reste de ce module (requetes KQL
    "coeur" du rapport) de fonctionner.

    Une regle peut referencer PLUSIEURS tactiques (ex: ["Persistence",
    "LateralMovement"]) : elle est alors comptee une fois POUR CHAQUE
    tactique qu'elle couvre, pas une seule fois au total -- coherent avec
    l'affichage du portail (un nombre de regles PAR tactique).

    Retourne {tactique_brute: nombre_de_regles_activees} (cle au meme
    format PascalCase que fetch_mitre_tactics_stats -- a faire passer par
    mitre_normalize.build_rule_counts avant affichage).

    Leve une exception (ImportError si le package n'est pas installe, ou
    toute erreur ARM/permissions) plutot que de l'avaler silencieusement :
    cf instruction utilisateur explicite ("si non recuperable, laissons
    tomber cette partie") -- c'est a l'appelant (generate_cosec.py) de
    decider comment degrader (laisser "Nombre de règles" à "N/A" sans
    bloquer le reste du rapport), pas a cette fonction.
    """
    from azure.mgmt.securityinsight import SecurityInsights

    resource = _resolve_workspace_resource(workspace_id, tenant_id=tenant_id)
    if not resource or not resource.get("name") or not resource.get("resourceGroup") \
            or not resource.get("subscriptionId"):
        raise RuntimeError(f"Ressource ARM introuvable pour le workspace {workspace_id}.")

    credential = _get_credential(tenant_id)

    client = SecurityInsights(credential, resource["subscriptionId"])

    counts = {}
    for rule in client.alert_rules.list(resource["resourceGroup"], resource["name"]):
        if not getattr(rule, "enabled", False):
            continue
        for tactic in (getattr(rule, "tactics", None) or []):
            tactic_str = str(tactic)
            counts[tactic_str] = counts.get(tactic_str, 0) + 1
    return counts


# ---------------------------------------------------------------------------
# Slide "Plan de collecte" (coût de l'ingestion des logs par catégorie)
# ---------------------------------------------------------------------------
#
# Ajoutee le 28/06/2026. Adaptee du workbook existant fourni par
# l'utilisateur (requete "coût de l'ingestion par type de log") -- la
# table de correspondance Categories et la logique 3-branches
# (customTables / AzDiagTables / knownTables) sont reprises A L'IDENTIQUE,
# avec 2 adaptations :
#
#   1. Filtre temporel explicite sur CreatedTime -- ABSENT de la requete
#      originale du workbook, qui s'appuie sur le selecteur "Time Range"
#      du workbook (parametre d'interface, invisible dans le texte KQL
#      lui-meme, qui injecte un filtre cache sur TimeGenerated cote
#      portail). Ce script n'a pas cette couche : on ajoute donc
#      explicitement `TimeGenerated >= {start} and < {end}` sur CHACUNE
#      des 3 branches (Usage pour customTables/knownTables,
#      AzureDiagnostics pour AzDiagTables), pour bien borner le mois
#      calendaire cible -- meme principe que toutes les autres requetes
#      de ce module.
#
#   2. Le parametre workbook `{Price}` (prix au Go, configure par
#      l'utilisateur dans l'interface du workbook, pas une donnee
#      Sentinel) devient un parametre Python `{price}` substitue par
#      fetch_log_ingestion_costs() -- cf son parametre price_per_gb,
#      OBLIGATOIRE : ce tarif est specifique au contrat de chaque client
#      et ne peut pas etre devine, contrairement aux autres requetes de
#      ce module qui n'ont besoin que du workspace/de la periode.
#
# Renommage des colonnes projetees (LogType/TableName/TableSizeGB/
# EstimatedCost, sans espace ni accolade) par rapport aux noms originaux
# du workbook (['Log Type'], ['Table'], ['Table Size'], ['Estimated
# cost']) -- simple confort de lecture cote Python (coherent avec le
# reste du module, ex: IncidentCount/AlertSources), aucun changement de
# logique de calcul. ['Table Size'] reste en Go (meme unite que
# l'original : Quantity (Mo) / 1024 pour customTables/knownTables,
# _BilledSize (octets) / 1024000000 pour AzDiagTables) -- la mise en
# forme lisible (Mo/Go/Ko) est une responsabilite d'affichage, faite cote
# Python par log_ingestion_normalize.format_size(), pas par cette requete.
#
# Ajout d'un filtre `where TableSizeGB > 0` en sortie : ecarte les lignes
# a 0 (table presente dans Categories mais sans aucune ingestion
# facturable ce mois-ci) qui n'apportent rien a l'affichage.
LOG_INGESTION_QUERY_TEMPLATE = """
let Categories = datatable(Type:string,Category:string)
[
"AuditLogs" , "Azure Active Directory",
"SigninLogs" , "Azure Active Directory",
"AADNonInteractiveUserSignInLogs" , "Azure Active Directory",
"AADRiskyUsers" , "Azure Active Directory",
"AADRiskyServicePrincipals" , "Azure Active Directory",
"AADServicePrincipalRiskEvents" , "Azure Active Directory",
"ADFSSignInLogs" , "Azure Active Directory",
"NetworkAccessTraffic" , "Azure Active Directory",
"AADUserRiskEvents" , "Azure Active Directory",
"AADServicePrincipalSignInLogs" , "Azure Active Directory",
"AADManagedIdentitySignInLogs" , "Azure Active Directory",
"AADProvisioningLogs" , "Azure Active Directory",
"AZFWApplicationRule" , "Firewall",
"AZFWApplicationRuleAggregation" , "Firewall",
"AZFWDnsQuery" , "Firewall",
"AZFWIdpsSignature" , "Firewall",
"AZFWNatRule" , "Firewall",
"AZFWNatRuleAggregation" , "Firewall",
"AZFWNetworkRule" , "Firewall",
"AZFWNetworkRuleAggregation" , "Firewall",
"WindowsFirewall" , "Firewall",
"Event" , "Custom Events",
"BehaviorAnalytics" , "User Entity Behavior Analytics",
"UserPeerAnalytics" , "User Entity Behavior Analytics",
"UserAccessAnalytics" , "User Entity Behavior Analytics",
"IdentityInfo" , "User Entity Behavior Analytics",
"DeviceLogonEvents" , "Microsoft Defender for Endpoint",
"DeviceEvents" , "Microsoft Defender for Endpoint",
"DeviceNetworkInfo" , "Microsoft Defender for Endpoint",
"DeviceImageLoadEvents" , "Microsoft Defender for Endpoint",
"DeviceFileEvents" , "Microsoft Defender for Endpoint",
"DeviceInfo" , "Microsoft Defender for Endpoint",
"DeviceProcessEvents" , "Microsoft Defender for Endpoint",
"DeviceNetworkEvents" , "Microsoft Defender for Endpoint",
"DeviceRegistryEvents" , "Microsoft Defender for Endpoint",
"DeviceFileCertificateInfo" , "Microsoft Defender for Endpoint",
"EmailAttachmentInfo" , "Microsoft Defender for Office 365",
"EmailEvents" , "Microsoft Defender for Office 365",
"EmailPostDeliveryEvents" , "Microsoft Defender for Office 365",
"EmailUrlInfo" , "Microsoft Defender for Office 365",
"IdentityLogonEvents" , "Microsoft Defender for Identity",
"IdentityQueryEvents" , "Microsoft Defender for Identity",
"IdentityDirectoryEvents" , "Microsoft Defender for Identity",
"CloudAppEvents" , "Microsoft Defender for Cloud Apps",
"AlertEvidence" , "Microsoft Defender Alert Evidence",
"InsightsMetrics" , "Azure Monitor for VMs",
"VMBoundPort" , "Azure Monitor for VMs",
"VMComputer" , "Azure Monitor for VMs",
"VMConnection" , "Azure Monitor for VMs",
"VMProcess" , "Azure Monitor for VMs",
"SecurityEvent" , "Windows Security Events",
"StorageBlobLogs" , "Azure Storage",
"StorageFileLogs" , "Azure Storage",
"Syslog" , "Syslog/CEF",
"SecurityIoTRawEvent" , "IoT Logs",
"CommonSecurityLog" , "Syslog/CEF",
"ThreatIntelligenceIndicator" , "Sentinel",
"DnsEvents" , "DNS Logs",
"DnsInventory" , "DNS Logs",
"AWSCloudTrail" , "AWS Logs",
"AWSVPCFlow" , "AWS Logs",
"ConfigurationChange" , "Change Tracking",
"ConfigurationData" , "Change Tracking",
"AzureDiagnostics" , "Azure Resources",
"AzureActivity" , "Azure Resources",
"LAQueryLogs" , "Management",
"SentinelHealth" , "Sentinel",
"Perf" , "Performance",
"AzureMetrics" , "Azure Metrics",
"SecurityNestedRecommendation" , "Microsoft Defender for Cloud",
"SecurityRecommendation" , "Microsoft Defender for Cloud",
"SecurityRegulatoryCompliance" , "Microsoft Defender for Cloud",
"SecureScoreControls" , "Microsoft Defender for Cloud",
"SecurityBaseline" , "Microsoft Defender for Cloud",
"SecureScores" , "Microsoft Defender for Cloud",
"Update" , "Update Management",
"UpdateSummary" , "Update Management",
"DeviceTvmSecureConfigurationAssessment" , "Microsoft Defender Vuln. Management",
"DeviceTvmSoftwareVulnerabilities" , "Microsoft Defender Vuln. Management",
"DeviceTvmSoftwareInventory" , "Microsoft Defender Vuln. Management",
"UrlClickEvents" , "Microsoft Defender for Office 365",
"SecurityBaselineSummary" , "Microsoft Defender for Cloud",
"AZFWThreatIntel" , "Firewall",
"AWSGuardDuty" , "AWS Logs",
"Watchlist" , "Sentinel",
"HuntingBookmark" , "Sentinel",
"SentinelAudit" , "Sentinel",
"Operation" , "Log Management",
"StorageTableLogs" , "Azure Storage",
"AddonAzureBackupStorage" , "Azure Storage",
"AddonAzureBackupPolicy" , "Azure Storage",
"StorageQueueLogs" , "Azure Storage",
"AddonAzureBackupProtectedInstance" , "Azure Storage",
"AddonAzureBackupJobs" , "Azure Storage"
];
let customTables = Usage
| where TimeGenerated >= datetime({start}) and TimeGenerated < datetime({end})
| where IsBillable == true
| where DataType contains "_CL"
| summarize size = sum(Quantity)/1024 by DataType
| project LogType = "Custom Log", TableName = DataType, TableSizeGB = size, EstimatedCost = size * {price};
let AzDiagTables = AzureDiagnostics
| where TimeGenerated >= datetime({start}) and TimeGenerated < datetime({end})
| summarize TotalIngestBytes = sum(_BilledSize) by Category
| project LogType = "AzureDiagnostics", TableName = Category, TableSizeGB = TotalIngestBytes/1024000000, EstimatedCost = (TotalIngestBytes/1024000000) * {price};
let knownTables = Usage
| where TimeGenerated >= datetime({start}) and TimeGenerated < datetime({end})
| where IsBillable == true
| where DataType <> "AzureDiagnostics"
| join kind=leftouter Categories on $left.DataType == $right.Type
| summarize size = sumif(Quantity, isnotempty(Category))/1024, sizeOther = sumif(Quantity, (isempty(Category) and DataType !contains "_CL"))/1024 by Category, DataType
| project LogType = iif(isnotempty(Category), Category, "Other"), TableName = DataType, TableSizeGB = iif(isnotempty(Category), size, sizeOther), EstimatedCost = iif(isnotempty(Category), size * {price}, sizeOther * {price});
union customTables, knownTables, AzDiagTables
| where TableSizeGB > 0
| sort by EstimatedCost desc
"""


def fetch_log_ingestion_costs(workspace_id: str, year: int, month: int,
                               price_per_gb: float, tenant_id: str = None) -> list[dict]:
    """
    Recupere, pour le mois calendaire donne, le volume ingere (en Go) et
    le cout estime correspondant, par categorie de log (cf Categories
    dans LOG_INGESTION_QUERY_TEMPLATE) et par table individuelle au sein
    de chaque categorie.

    price_per_gb : tarif en €/Go a appliquer (cf note du template de
    requete ci-dessus -- OBLIGATOIRE, specifique au contrat du client,
    ne peut pas etre devine ni avoir de valeur par defaut raisonnable).

    Retourne une liste de dicts : LogType (categorie, ou "Custom Log" /
    "AzureDiagnostics" / "Other" selon la branche -- cf requete),
    TableName, TableSizeGB (chaine numerique, en Go), EstimatedCost
    (chaine numerique, en €).
    """
    credential = _get_credential(tenant_id)
    client = LogsQueryClient(credential)

    start, end = _month_bounds(year, month)
    query = LOG_INGESTION_QUERY_TEMPLATE.format(
        start=start.isoformat(), end=end.isoformat(), price=price_per_gb)

    response = client.query_workspace(
        workspace_id=workspace_id,
        query=query,
        timespan=(start, datetime.now(timezone.utc)),
    )

    if response.status == LogsQueryStatus.PARTIAL:
        print(f"⚠ Reponse partielle : {response.partial_error}")
        table = response.partial_data[0]
    elif response.status == LogsQueryStatus.SUCCESS:
        if not response.tables or not response.tables[0].rows:
            return []
        table = response.tables[0]
    else:
        raise RuntimeError(f"Echec de la requete d'ingestion des logs sur le workspace {workspace_id} : {response}")

    columns = table.columns
    return [_row_to_dict(columns, row) for row in table.rows]
