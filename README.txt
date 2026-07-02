================================================================================
 COSEC - Pipeline d'automatisation des rapports de securite mensuels
 README - Commandes et options disponibles
================================================================================

POINT D'ENTREE
--------------
Toutes les fonctionnalites du pipeline sont pilotees depuis un seul script :

    python generate_cosec.py [options]

Fichiers attendus dans le repertoire de travail :
    - template_slide.pptx       (template PowerPoint, fourni avec le projet)
    - historique_cosec.xlsx     (cree automatiquement si absent, via --update-history)

Fichier produit :
    - COSEC_rapport.pptx        (ecrase a chaque execution)


USAGE MINIMAL
-------------
    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6


OPTIONS OBLIGATOIRES
---------------------
--workspace-id <GUID>
    GUID du workspace Log Analytics cible (Azure Portal > workspace >
    Overview > Workspace ID -- PAS le nom de la ressource).

--year <AAAA>
    Annee du mois cible (ex: 2026). Utilisee pour les slides de detail par
    incident ET pour le calcul de la slide d'evolution par typologie.

--month <1-12>
    Mois cible. Meme remarque que --year : doit toujours etre fourni en
    paire coherente avec --year.


OPTIONS D'AUTHENTIFICATION
----------------------------
--tenant-id <GUID>
    (Optionnel) Force le tenant Azure AD cible. A utiliser si le compte
    connecte a acces a plusieurs tenants via Azure Lighthouse et que la
    resolution automatique echoue ou cible le mauvais tenant.
    Defaut : resolution automatique.


OPTIONS DE REFORMULATION IA (Claude API)
-------------------------------------------
--ai
    Active la reformulation automatique de la description de chaque
    incident via l'API Claude. Les donnees sensibles (comptes, IP, hotes,
    fichiers, groupes...) sont anonymisees avant tout envoi (cf
    anonymizer.py) puis desanonymisees dans le resultat.
    Necessite la variable d'environnement ANTHROPIC_API_KEY.

--debug
    Mode validation humaine : avant chaque envoi a l'API Claude, affiche le
    payload anonymise exact qui sera transmis, la table de correspondance
    (alias <-> valeur reelle, jamais envoyee), et un controle anti-fuite.
    Demande une confirmation (o/N) avant d'envoyer.
    N'a d'effet qu'en combinaison avec --ai ; si --debug est fourni sans
    --ai, --ai est active automatiquement (avec un avertissement console).


OPTIONS DE GESTION DE L'HISTORIQUE EXCEL
-------------------------------------------
--update-history
    Recupere l'historique des typologies, de l'etat de surveillance
    (gravite/cloture/MTTA-MTTR-MTTC) ET des depassements de SLA pour le
    mois --year/--month, et les integre dans le fichier Excel AVANT de
    generer le pptx. Sans cette option, le pptx est genere a partir de
    l'historique Excel existant tel quel (les slides de synthese
    afficheront alors les donnees du dernier mois deja present, qui peut
    etre different de --year/--month).
    Idempotent : relancer pour un mois deja present remplace ses lignes
    sans creer de doublon ; les autres mois sont preserves.

--history-excel <chemin>
    Chemin du fichier Excel historique (3 onglets : Typologies,
    Surveillance, SLA).
    Defaut : historique_cosec.xlsx


OPTIONS DE DESACTIVATION DES SLIDES DE SYNTHESE
---------------------------------------------------
Par defaut, les 5 slides de synthese suivantes sont TOUTES generees (en
plus des slides de detail par incident, toujours generees). Chacune peut
etre desactivee individuellement :

--no-evolution-slide
    Ne pas remplir la slide "Evolution des incidents par typologie"
    (necessite l'onglet Typologies de l'historique Excel).

--no-surveillance-slide
    Ne pas remplir la slide "Etat de la surveillance" -- repartition par
    gravite/cloture + KPI MTTA/MTTR/MTTC (necessite l'onglet Surveillance).

--no-sla-slide
    Ne pas remplir la slide "Depassement des SLA" (necessite l'onglet SLA).

--no-dispositif-slide
    Ne pas remplir la slide "Dispositif de surveillance" -- couverture par
    tactique MITRE ATT&CK. N'utilise PAS l'historique Excel (affiche
    directement le mois cible).

--no-log-ingestion-slide
    Ne pas remplir la slide "Plan de collecte" -- cout d'ingestion des logs
    par categorie. N'utilise PAS l'historique Excel (affiche directement
    le mois cible).


OPTIONS DE TARIFICATION
--------------------------
--price-per-gb <valeur>
    Prix en euros par Go ingere, utilise pour le calcul du cout estime de
    la slide "Plan de collecte".
    Defaut : 4.89 (tarif contractuel par defaut -- a adapter si le contrat
    du client differe).
    Sans effet si --no-log-ingestion-slide est utilise.


EXEMPLES
--------

1. Generation simple, a partir de l'historique Excel existant (sans mise a
   jour de l'historique, sans IA) :

    python generate_cosec.py --workspace-id 11111111-2222-3333-4444-555555555555 --year 2026 --month 6

2. Generation complete avec mise a jour de l'historique pour le mois cible :

    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 --update-history

3. Generation avec reformulation IA des descriptions d'incidents :

    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 --update-history --ai

4. Generation avec reformulation IA et validation humaine de chaque envoi :

    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 --ai --debug

5. Generation en forcant un tenant precis (acces multi-tenant Lighthouse) :

    python generate_cosec.py --workspace-id <GUID> --tenant-id <TENANT_GUID> --year 2026 --month 6

6. Generation sans les slides MITRE ATT&CK et Plan de collecte (donnees non
   disponibles ou non pertinentes pour ce client) :

    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 --no-dispositif-slide --no-log-ingestion-slide

7. Generation avec un tarif d'ingestion specifique et un fichier
   d'historique different du defaut :

    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 --price-per-gb 5.20 --history-excel historique_client_x.xlsx


NOTES
-----
- L'authentification Azure se fait via une fenetre de navigateur
  interactive (InteractiveBrowserCredential), declenchee au premier appel
  necessitant un acces a Sentinel/Resource Graph/ARM, puis reutilisee pour
  toute la duree de l'execution.
- Les echecs des fonctionnalites "cosmetiques" (nom du workspace pour le
  bandeau "Confidentiel", statistiques MITRE ATT&CK, nombre de regles
  actives, cout d'ingestion des logs) sont tolere : un avertissement est
  affiche en console et le rapport continue d'etre genere, avec les
  champs concernes a "N/A" ou la slide correspondante non renseignee.
- Les echecs sur les donnees coeur du rapport (liste des incidents du
  mois, mise a jour de l'historique surveillance/SLA en mode
  --update-history) interrompent le script immediatement (code de sortie 1).
================================================================================
