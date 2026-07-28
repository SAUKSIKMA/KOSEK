# Architecture du pipeline COSEC

Ce document décrit l'architecture technique du pipeline d'automatisation des
rapports mensuels COSEC (Comité de Sécurité). Il s'adresse à toute personne
reprenant ou faisant évoluer le code.

## 1. Vue d'ensemble

Le pipeline part d'une requête KQL live sur un workspace Microsoft Sentinel
(Azure Monitor Query) et produit en sortie :

- un fichier **PowerPoint** (`COSEC_rapport.pptx`) — le rapport mensuel,
- un fichier **Excel** (`historique_cosec.xlsx`) — l'historique multi-mois
  utilisé par certaines slides pour afficher une évolution.

Il n'y a **aucune base de données** : l'Excel fait office de mémoire
persistante entre deux exécutions mensuelles, et le pptx est régénéré en
intégralité à chaque exécution.

Point d'entrée unique : `generate_cosec.py` (CLI), qui orchestre tous les
autres modules. Le pipeline traite un client par exécution ; pour
plusieurs clients partageant le même code et le même template, voir
l'orchestrateur `run_all_clients.py` (section 9).

## 2. Schéma du pipeline

```
                         ┌─────────────────────┐
                         │   generate_cosec.py   │   <-- CLI / orchestration
                         └───────────┬───────────┘
                                     │
         ┌───────────────────────────┼───────────────────────────┐
         ▼                           ▼                           ▼
 ┌───────────────┐          ┌────────────────┐          ┌────────────────┐
 │ sentinel_query │          │ *_normalize.py │          │  *_slide.py    │
 │   (KQL live)   │ ───────▶ │ (agrégation /  │ ───────▶ │ (écriture pptx │
 │                │          │  formatage)    │          │  via python-   │
 └───────────────┘          └───────┬────────┘          │  pptx)         │
                                     │                    └────────────────┘
                                     ▼
                            ┌────────────────┐
                            │ excel_history  │   <-- historique multi-mois
                            │   (openpyxl)   │       (--update-history)
                            └────────────────┘

  Optionnel (--ai) :  generate_cosec.py ─▶ anonymizer.py ─▶ reformulate.py ─▶ Claude API
```

Chaque slide de synthèse (Évolution, Surveillance, SLA, Dispositif de
surveillance, Plan de collecte) suit le même schéma à 3 étages :
**requête Sentinel → normalisation/agrégation → remplissage pptx**, avec
ou sans passage par l'Excel historique selon la slide (cf section 5).

## 3. Modules, par responsabilité

### 3.1 Orchestration

| Fichier | Rôle |
|---|---|
| `generate_cosec.py` | Point d'entrée CLI. Enchaîne les requêtes Sentinel, la mise à jour optionnelle de l'historique Excel, le clonage des slides de détail par incident, le remplissage des slides de synthèse, leur réordonnancement, puis la sauvegarde du pptx final. Contient aussi les helpers bas niveau de manipulation du XML pptx (clonage de slide, remplacement de texte en préservant le formatage). `generate_pptx()` accepte `template_path=` et `output_path=` (défauts : `TEMPLATE_PATH`/`OUTPUT_PATH`, comme `history_excel=`) — surchargeables via `--template-path`/`--output`, ce qui permet l'usage multi-clients (cf section 9) sans dupliquer le code. |
| `run_all_clients.py` | Orchestrateur multi-clients (optionnel, ajouté le 27/07/2026). Lit `clients.json`, boucle sur chaque client déclaré et appelle `generate_pptx()` directement (import, pas de subprocess) avec un `history_excel`/`output_path` distinct par client mais un `template_path` partagé. Isole les échecs par client (un client en erreur n'interrompt pas les suivants) et journalise dans `logs/`. Voir section 9. |

### 3.2 Récupération des données (Microsoft Sentinel)

| Fichier | Rôle |
|---|---|
| `sentinel_query.py` | Unique point d'accès à Sentinel. Contient tous les templates de requêtes KQL (incidents du mois, historique de typologies, répartition gravité/clôture, MTTA/MTTR/MTTC, dépassements de SLA, statistiques MITRE ATT&CK, coût d'ingestion des logs) ainsi que la résolution de l'identité ARM du workspace (nom, groupe de ressources, abonnement) via Azure Resource Graph. Gère un cache de credential Azure partagé entre tous les appels d'une même exécution pour éviter les ré-authentifications interactives répétées. |

### 3.3 Normalisation / agrégation (logique métier pure, sans dépendance UI)

| Fichier | Rôle |
|---|---|
| `typology_normalize.py` | Retire les suffixes de comptage d'entités générés automatiquement par Sentinel dans les titres d'incidents, replie les familles de titres à identifiant/date variable sur un libellé canonique (`_FAMILY_RULES`, ex. « Purview IRM »), puis ré-agrège par typologie normalisée. |
| `surveillance_normalize.py` | Assemble les 3 requêtes de la slide « État de la surveillance » (gravité, classification de clôture, MTTA/MTTR/MTTC) en une ligne unique. Contient la logique de correspondance des valeurs `Classification` (enum Sentinel toujours en anglais). |
| `sla_normalize.py` | Convertit les dépassements de SLA bruts en lignes prêtes pour l'Excel (tri par gravité, type de SLA, date), avec conversion des dates ISO en heure de Paris. |
| `mitre_normalize.py` | Combine les statistiques d'incidents par tactique MITRE ATT&CK avec le nombre de règles analytics actives par tactique, et convertit les clés brutes Microsoft (PascalCase) en libellés affichés. |
| `log_ingestion_normalize.py` | Regroupe les volumes d'ingestion de logs par catégorie/table, calcule le positionnement sur l'échelle de couleur (heatmap) et sélectionne les catégories à afficher selon le budget de lignes disponible sur la slide. |

### 3.4 Construction des slides (python-pptx)

| Fichier | Rôle |
|---|---|
| `typology_slide.py` | Slide « Évolution des incidents par typologie ». Contient aussi plusieurs helpers réutilisés par les autres modules `*_slide.py` (clonage de slide générique, recherche de shape par nom, réordonnancement des slides) — voir section 4. |
| `surveillance_slide.py` | Slide « État de la surveillance » : panneau de texte, 2 donuts (gravité / clôture), 3 cartes KPI (MTTA/MTTR/MTTC). |
| `sla_slide.py` | Slide « Dépassement des SLA » : tableau des incidents en dépassement MTTA/MTTR pour le mois cible. |
| `mitre_slide.py` | Slide « Dispositif de surveillance » : remplissage des 12 encadrés de tactique MITRE ATT&CK déjà présents dans le template. |
| `log_ingestion_slide.py` | Slide « Plan de collecte » : tableau arborescent catégorie/table avec barre de volume et dégradé de couleur. |

### 3.5 Persistance Excel

| Fichier | Rôle |
|---|---|
| `excel_history.py` | Lecture/écriture des 3 onglets de `historique_cosec.xlsx` (Typologies, Surveillance, SLA). Toutes les écritures sont idempotentes par mois : relancer pour un mois déjà présent remplace ses lignes sans dupliquer, sans toucher aux autres mois. |

### 3.6 Reformulation IA (optionnel, `--ai`)

| Fichier | Rôle |
|---|---|
| `anonymizer.py` | Anonymisation réversible (UPN, IP, hôtes, fichiers, groupes...) avant tout envoi de texte à l'API Claude, avec table de correspondance locale et détection de fuite résiduelle. |
| `reformulate.py` | Construit le payload anonymisé d'un incident, appelle l'API Claude pour produire une description de niveau consultant, puis désanonymise le résultat. |

## 4. Conventions internes notables

- **Duplication volontaire de petites fonctions utilitaires** (`_get_shape_by_name`,
  `_set_cell_text`, `_format_dt`, parsing de date ISO en heure de Paris...)
  entre certains modules `*_slide.py` / `*_normalize.py`, plutôt qu'un import
  croisé, afin d'éviter les imports circulaires (ex: `anonymizer.py` /
  `reformulate.py` important `generate_cosec.parse_json_array`).
- **Slides de synthèse capturées par index fixe** (`prs.slides[1]` à `[5]`)
  *avant* tout réordonnancement (`move_slide_to_front`), car ce dernier
  repose sur le `slide_id` et fait glisser les index une fois invoqué.
- **Idempotence par mois** sur les 3 onglets Excel : une ré-exécution avec
  `--update-history` sur un mois déjà traité remplace ses lignes plutôt que
  de les dupliquer.
- **Tolérance aux pannes différenciée** : les échecs de requêtes
  « cosmétiques » (nom du workspace, statistiques MITRE, coût d'ingestion)
  sont avalés avec un avertissement console ; les échecs sur les données
  cœur du rapport (incidents du mois, historique surveillance/SLA en mode
  `--update-history`) interrompent le script (`sys.exit(1)`).

## 5. Les 6 slides du template (`template_slide.pptx`)

Le template contient les slides suivantes, dans cet ordre fixe en entrée
(capturées par index avant tout réordonnancement) :

| Index | Contenu | Source des données | Historique Excel ? |
|---|---|---|---|
| 0 | Détail d'un incident (« Focus sur incident ») — clonée une fois par incident du mois | `fetch_cosec_incidents` | Non |
| 1 | Évolution des incidents par typologie | `fetch_typology_history` | Oui (onglet Typologies) — affiche le dernier mois de l'historique |
| 2 | État de la surveillance (gravité / clôture / MTTA-MTTR-MTTC) | `fetch_severity_breakdown`, `fetch_classification_breakdown`, `fetch_resolution_times` | Oui (onglet Surveillance) — affiche le dernier mois de l'historique |
| 3 | Dépassement des SLA | `fetch_sla_breaches` | Oui (onglet SLA) — affiche le **mois cible** du rapport |
| 4 | Dispositif de surveillance (couverture MITRE ATT&CK) | `fetch_mitre_tactics_stats`, `fetch_active_rules_by_tactic` (bonus) | Non — affiche directement le mois cible |
| 5 | Plan de collecte (coût d'ingestion des logs) | `fetch_log_ingestion_costs` | Non — affiche directement le mois cible |

En sortie, l'ordre final des slides est : Surveillance, Évolution, SLA,
Dispositif de surveillance, Plan de collecte, puis les slides de détail par
incident (cf `generate_pptx()` dans `generate_cosec.py` pour le détail de
la logique de réordonnancement).

Chaque slide peut être désactivée individuellement via un flag CLI
`--no-<slide>-slide` (cf README.txt).

## 6. L'Excel historique (`historique_cosec.xlsx`)

Trois onglets indépendants, chacun avec sa propre granularité :

| Onglet | Granularité | Colonnes principales |
|---|---|---|
| **Typologies** | 1 ligne par (Mois, Typologie) | Mois, Typologie, Sources d'alerte, Nombre d'incidents |
| **Surveillance** | 1 ligne par Mois | Mois, Total, répartition par gravité (4), répartition par clôture (4), MTTA, MTTR, MTTC |
| **SLA** | 1 ligne par incident en dépassement | Mois, Type SLA, N°INC, Sévérité, Titre, Créé le, Attribution, Clôture |

Ces 3 onglets sont alimentés indépendamment (`write_history`,
`write_surveillance_history`, `write_sla_history`) — l'ordre d'appel entre
eux n'a aucune importance.

## 7. Dépendances Python

### 7.1 Bibliothèques tierces à installer

```bash
pip install azure-identity azure-monitor-query azure-mgmt-resourcegraph \
            azure-mgmt-subscription azure-mgmt-securityinsight \
            openpyxl python-pptx lxml anthropic
```

| Bibliothèque | Utilisée par | Rôle | Obligatoire ? |
|---|---|---|---|
| `azure-identity` | `sentinel_query.py` | Authentification interactive Azure AD (`InteractiveBrowserCredential`) | Oui |
| `azure-monitor-query` | `sentinel_query.py` | Exécution des requêtes KQL (`LogsQueryClient`) | Oui |
| `azure-mgmt-resourcegraph` | `sentinel_query.py` | Résolution du nom/groupe de ressources/abonnement du workspace (bandeau « Confidentiel », bonus MITRE) | Oui |
| `azure-mgmt-subscription` | `sentinel_query.py` | Énumération des abonnements accessibles, requise par Resource Graph | Oui |
| `azure-mgmt-securityinsight` | `sentinel_query.py` | Bonus « Nombre de règles actives » par tactique MITRE ATT&CK (slide Dispositif de surveillance) | **Non** — import différé, dégrade en `N/A` si absent |
| `openpyxl` | `excel_history.py`, `*_slide.py` (lecture) | Lecture/écriture de `historique_cosec.xlsx` | Oui |
| `python-pptx` (package `pptx`) | `generate_cosec.py`, `*_slide.py` | Génération du rapport PowerPoint | Oui |
| `lxml` | `generate_cosec.py`, `surveillance_slide.py`, `log_ingestion_slide.py`, `typology_slide.py` | Manipulation directe du XML pptx (runs, mise en page de légende, clonage de relations) | Oui |
| `anthropic` | `reformulate.py` | Appel à l'API Claude pour la reformulation des descriptions d'incidents | **Non** — uniquement requis si le flag `--ai` est utilisé |

### 7.2 Bibliothèques standard utilisées

`json`, `os`, `re`, `sys`, `copy`, `argparse`, `math`, `ipaddress`, `datetime`
— aucune installation requise (incluses avec Python).

### 7.3 Version de Python

**Python 3.9 minimum** : le code utilise la syntaxe de generics native sur
les types intégrés (`list[dict]`, `list[tuple[str, str]]`), introduite par
la PEP 585. Une version 3.10 ou 3.11 est recommandée.

### 7.4 Variable d'environnement

| Variable | Requise pour | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | `--ai` | Clé d'API Anthropic, utilisée par `reformulate.make_client()`. Le script s'arrête avec une erreur explicite si elle est absente et que `--ai` est demandé. |

## 8. Authentification Azure

L'authentification se fait via `InteractiveBrowserCredential` (popup de
connexion dans le navigateur par défaut). Une seule authentification est
déclenchée par exécution du script et par `tenant_id` distinct : l'instance
de credential est mise en cache en mémoire process (`sentinel_query.
_credential_cache`) et réutilisée pour tous les appels suivants (Log
Analytics, Resource Graph, gestion Sentinel), y compris vers des scopes
Azure différents.

Le compte utilisé doit avoir accès en lecture au workspace Log Analytics
ciblé (`--workspace-id`), ainsi qu'à l'abonnement qui l'héberge si les
fonctionnalités optionnelles (bandeau « Confidentiel », bonus « Nombre de
règles ») sont souhaitées.

> **Multi-clients** : ce cache est la raison pour laquelle
> `run_all_clients.py` (section 9) appelle `generate_pptx()` par import
> dans un seul process plutôt que via un subprocess par client — les
> clients partageant un même `tenant_id` (accès MSP/Lighthouse) ne
> déclenchent qu'une seule authentification interactive pour tout le run,
> au lieu d'une par client.

## 9. Utilisation multi-clients (ajouté le 27/07/2026)

Le pipeline génère un rapport pour **un** client par appel de
`generate_pptx()`. Pour plusieurs clients, on ne duplique pas le code : un
seul jeu de scripts et un seul `template_slide.pptx` sont partagés, seuls
`history_excel` et `output_path` varient par client (chacun ayant son
propre historique Excel, comme demandé par l'utilisateur).

### 9.1 Arborescence recommandée

```
cosec/
├── scripts/                 <- code partagé (les modules ci-dessus, inchangés)
├── template_slide.pptx      <- template unique, partagé par tous les clients
├── clients.json             <- config déclarative de tous les clients
├── clients/
│   ├── CLIENT1/
│   │   ├── historique_cosec.xlsx
│   │   └── output/          <- COSEC_CLIENT1_<AAAA>-<MM>.pptx
│   └── CLIENT2/
│       ├── historique_cosec.xlsx
│       └── output/
├── run_all_clients.py       <- orchestrateur
└── logs/                    <- un fichier de log par run (créé automatiquement)
```

### 9.2 `clients.json`

Un objet par client, avec au minimum `name`, `workspace_id` et
`history_excel`. Champs optionnels : `tenant_id`, `price_per_gb` (défaut
4.89), `output_dir`, et les flags `no_<slide>_slide` (mêmes noms que les
options CLI `--no-<slide>-slide`, en `snake_case`, valeur booléenne) pour
désactiver une slide pour un client donné sans toucher aux autres.

```json
[
  {
    "name": "CLIENT1",
    "workspace_id": "00000000-0000-0000-0000-000000000001",
    "tenant_id": "11111111-1111-1111-1111-111111111111",
    "history_excel": "clients/CLIENT1/historique_cosec.xlsx",
    "output_dir": "clients/CLIENT1/output",
    "price_per_gb": 4.89
  }
]
```

### 9.3 `run_all_clients.py`

Charge `clients.json`, puis pour chaque client : construit
`output_path = <output_dir>/COSEC_<name>_<année>-<mois>.pptx`, appelle
`generate_pptx()` avec les paramètres du client et le `template_path`
partagé, capture toute exception pour ne pas interrompre les clients
suivants, et journalise un bilan final (réussites/échecs) dans
`logs/run_<année>-<mois>_<horodatage>.log`.

Le mode mono-client d'origine (`python generate_cosec.py --workspace-id
...`) reste utilisable tel quel depuis `scripts/`, avec les nouvelles
options `--template-path`/`--output` pour pointer vers l'arborescence
ci-dessus si besoin.

### 9.4 Commandes

```bash
# Tous les clients déclarés dans clients.json
python run_all_clients.py --year 2026 --month 6 --update-history

# Un sous-ensemble de clients seulement
python run_all_clients.py --year 2026 --month 6 --only CLIENT1,CLIENT2

# Avec reformulation IA (Claude)
python run_all_clients.py --year 2026 --month 6 --update-history --ai
```

### 9.5 Limites connues

- **Exécution séquentielle uniquement** : `run_all_clients.py` ne
  parallélise pas les clients. Avec `InteractiveBrowserCredential`, lancer
  plusieurs clients en parallèle déclencherait des popups de connexion
  concurrents. Une parallélisation ne serait envisageable qu'après
  migration vers une authentification non interactive (service principal)
  par tenant.
- **Pas de gestion de secrets dédiée** : `clients.json` ne contient pas
  d'information sensible en l'état (les `workspace_id`/`tenant_id` ne sont
  pas des secrets), mais si une authentification par service principal est
  ajoutée à l'avenir, les secrets associés devront être externalisés
  (variables d'environnement ou coffre) plutôt qu'ajoutés en clair dans ce
  fichier.
