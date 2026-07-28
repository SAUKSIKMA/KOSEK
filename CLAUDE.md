# CLAUDE.md — COSEC

Directives comportementales pour Claude Code sur ce dépôt. Basées sur les
4 principes généraux de Karpathy contre le sur-engineering des agents LLM
(réflexion préalable, simplicité, changements chirurgicaux, exécution
pilotée par des critères vérifiables), déclinés avec les conventions
concrètes de ce pipeline. À fusionner avec toute directive globale
utilisateur si présente.

**Compromis assumé :** ces règles privilégient la prudence à la vitesse.
Pour une tâche triviale (typo, renommage local), utilise ton jugement —
ne demande pas de confirmation pour changer une chaîne de caractères.

## 0. Repères rapides

Pipeline de génération de rapports COSEC mensuels : requêtes KQL sur
Microsoft Sentinel → normalisation → rendu PowerPoint + historique Excel.
Pas de base de données ; l'Excel est la mémoire persistante entre deux
mois. Architecture détaillée, schéma du pipeline, et le tableau des 6
slides du template : voir `ARCHITECTURE.md` (ne pas le paraphraser
ici — le lire à la demande plutôt que de le dupliquer dans ce fichier).

Commandes courantes :

```bash
# Un seul client
python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 \
    --update-history

# Plusieurs clients (cf clients.json, ARCHITECTURE.md section 9)
python run_all_clients.py --year 2026 --month 6 --update-history

# Vérification visuelle d'un pptx généré
libreoffice --headless --convert-to pdf COSEC_rapport.pptx
```

Repères dans le code : `sentinel_query.py` (toute la donnée Sentinel),
`*_normalize.py` (logique métier pure), `*_slide.py` (écriture pptx),
`generate_cosec.py` (orchestration CLI + helpers XML bas niveau),
`excel_history.py` (les 3 onglets de `historique_cosec.xlsx`).

## 1. Réfléchir avant de coder

**Ne suppose pas. N'avance pas dans le flou. Explicite les compromis.**

Sur ce projet, le flou vient presque toujours de la donnée Sentinel, pas
du code. Avant d'écrire une requête KQL ou une normalisation :

- Vérifie la forme réelle des données retournées (`--debug`, ou une
  requête manuelle dans le portail) avant d'écrire la logique de
  normalisation — ne devine pas le schéma depuis le nom des colonnes.
- Rappelle-toi les pièges déjà documentés avant de les re-découvrir :
  `SecurityIncident` est en une ligne par mise à jour (utiliser
  `dcount(IncidentName)`, jamais `count()`) ; le filtre temporel doit
  porter sur `CreatedTime`, pas `TimeGenerated` ; `Classification` est
  toujours un enum anglais quel que soit le tenant, même si l'UI du
  portail l'affiche traduit.
- Si une valeur métier n'est pas spécifiée (seuil SLA, prix/Go, libellé
  d'une nouvelle tactique MITRE, format d'une nouvelle colonne Excel),
  demande plutôt que de choisir une valeur plausible en silence.
- Si le template `.pptx` doit être modifié ou qu'une nouvelle slide s'y
  ajoute, ouvre-le d'abord (LibreOffice) pour voir les shapes/dimensions
  réelles — ne suppose pas des dimensions standard (le template fait
  ~22×12.4 pouces, pas 13.3×7.5).
- Si une requête KQL existante (adaptée d'un workbook) doit être
  modifiée, dis explicitement quelle logique métier tu préserves
  (ex : déduction "attente client", distinction HNO/heures ouvrées) et
  laquelle tu changes — ne les modifie jamais implicitement.

## 2. Simplicité d'abord

**Le minimum de code qui résout le problème. Rien de spéculatif.**

- Pas de nouvelle couche d'abstraction, de config générique ou de
  dépendance (DB, framework web, ORM) non demandée — le pipeline reste
  volontairement sans base de données, l'Excel fait office de mémoire.
- Pas de gestion d'erreur pour des formes de réponse Sentinel jamais
  observées en pratique. Suis la tolérance aux pannes déjà en place :
  échecs "cosmétiques" (nom du workspace, stats MITRE, coût
  d'ingestion) → avertissement console et dégradation gracieuse ;
  échecs sur la donnée cœur (incidents du mois, historique en mode
  `--update-history`) → `sys.exit(1)`, ne pas avaler l'erreur.
- Une nouvelle slide ou un nouveau flag suit le pattern CLI existant
  (`--no-<slide>-slide`) plutôt qu'un nouveau système de configuration.
- Si tu écris 200 lignes pour une slide qui pourrait en tenir 50,
  réécris — mais vérifie d'abord qu'un module `*_slide.py` existant ne
  fait pas déjà 80% du travail (cf. section 3).

## 3. Changements chirurgicaux

**Ne touche qu'à ce qui est nécessaire. Ne nettoie que ton propre désordre.**

Ce dépôt a des conventions internes délibérées, documentées dans
`ARCHITECTURE.md` section 4 — ne les "corrige" pas en les prenant pour
des erreurs :

- La duplication de petits helpers (`_set_cell_text`, `_get_shape_by_name`,
  parsing de date ISO...) entre modules `*_slide.py`/`*_normalize.py` est
  **intentionnelle** (évite les imports circulaires). Ne factorise pas
  ces fonctions dans un module commun sans qu'on te le demande — et si tu
  modifies l'une des copies, dis explicitement si les autres copies
  doivent suivre ou si une divergence de comportement est voulue.
- Les slides de synthèse sont capturées par index (`prs.slides[1]` à
  `[5]`) **avant** tout `move_slide_to_front()` — ne réordonne jamais ces
  captures, et si tu ajoutes un réordonnancement, fais-le en ordre
  inverse comme l'existant.
- L'arithmétique EMU utilise `//`/`int()`, jamais `/` seul (float invalide
  pour python-pptx) — reste cohérent avec ce style dans toute nouvelle
  ligne de manipulation pptx.
- Les écritures Excel sont idempotentes par mois — un mois à zéro (0
  dépassement SLA) s'écrit explicitement, ne l'omets pas en pensant
  "simplifier".
- Si tu remarques du code mort ou une incohérence non liée à ta tâche
  (ex : une des copies dupliquées d'un helper qui a dérivé des autres),
  signale-le — ne le corrige pas au passage.

Test : chaque ligne modifiée doit se rattacher directement à la demande.

## 4. Devenir un avec la donnée

**Avant de coder une normalisation, regarde vraiment ce qu'elle normalise.**

Décliné du principe de Karpathy "become one with the data" (issu de sa
recette d'entraînement de réseaux de neurones, mais qui s'applique tel
quel ici) : la plupart des bugs de ce pipeline viennent d'une hypothèse
sur la donnée Sentinel qui ne tenait pas.

- Avant d'écrire ou modifier une fonction `*_normalize.py`, fais tourner
  la requête `sentinel_query.py` correspondante sur un vrai workspace (ou
  relis un export `--debug` récent) et regarde les valeurs brutes, pas
  seulement leur type déclaré.
- Une valeur qui a l'air d'un cas limite rare (`Classification` en
  français dans un vieux ticket, une fenêtre de rétention de 30 jours qui
  purge le mois demandé, un incident sans `Severity`) est probablement un
  cas déjà rencontré — vérifie l'historique du fichier avant de supposer
  que c'est nouveau.
- Ne généralise pas une nouvelle slide/typologie à tous les clients avant
  de l'avoir vue rendue correctement (LibreOffice) sur au moins un jeu de
  données réel.

## 5. Exécution pilotée par des critères vérifiables

**Définis le succès. Boucle jusqu'à vérification.**

Ce projet n'a pas de suite de tests automatisés formelle — la
vérification passe par des tests mock de bout en bout et un rendu visuel
LibreOffice. Transforme donc toute tâche en critère observable :

- "Ajoute une slide" → "génère un pptx avec des données mock, ouvre-le
  dans LibreOffice, compare visuellement au gabarit/à la maquette
  attendue"
- "Corrige ce bug KQL" → "reproduis l'écart entre le portail Azure et le
  résultat du SDK Python sur le même mois, explique la cause (souvent :
  filtre `TimeGenerated` implicite du portail), puis vérifie que les
  deux concordent après correction"
- "Étends `historique_cosec.xlsx`" → "relance `--update-history` deux
  fois de suite sur le même mois et vérifie qu'aucune ligne n'est
  dupliquée"
- Pour une tâche multi-étapes, énonce un plan bref avant de commencer :

  ```
  1. [Étape] → vérifier : [contrôle]
  2. [Étape] → vérifier : [contrôle]
  3. [Étape] → vérifier : [contrôle]
  ```

---

**Ces règles fonctionnent si :** les diffs ne touchent que ce qui est
demandé, les hypothèses sur la donnée Sentinel sont vérifiées avant
d'être codées, et les questions de clarification arrivent avant
l'implémentation plutôt qu'après un rendu pptx qui ne correspond pas à
l'attendu.
