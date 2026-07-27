"""
Construction de la slide "Dispositif de surveillance" (5e slide du
template, ajoutee le 28/06/2026), qui affiche pour chaque tactique MITRE
ATT&CK un encadre "Nombre de règles / Dernière exécution / Incident(s)".

Contrairement aux 3 autres slides de synthese (typology_slide.py,
surveillance_slide.py, sla_slide.py), cette slide n'a PAS de tableau ou de
graphique a construire dynamiquement : les 12 encadres existent deja dans
le template (un par tactique), chacun avec ses 3 lignes a "N/A" -- on se
contente de remplacer ces "N/A" par les valeurs calculees, exactement
comme generate_cosec.fill_zone_texte_9/fill_simple_shape le font pour les
slides de detail par incident.

Particularite du template (verifiee le 28/06/2026 en inspectant le vrai
template_slide.pptx) : les 12 encadres partagent TOUS le meme nom de
shape ("object 17" -- copies/collees depuis un meme objet d'origine), on
ne peut donc PAS les distinguer par shape.name comme TITLE_SHAPE_NAME/
SUBTITLE_SHAPE_NAME ailleurs dans le projet (typology_slide.py). On les
identifie a la place par le texte de leur 1er paragraphe (le nom de la
tactique, ex: "Persistence", "Credential access"), qui sert de cle de
correspondance avec les libelles de mitre_normalize.TACTIC_LABELS.

Structure d'un encadre (verifiee sur le template) :
  Paragraphe 0 : <Nom de la tactique>             (1 ou plusieurs runs,
                 ex: "Resource " + "Development")
  Paragraphe 1 : "Nombre de règles : " + "N/A"    (2 runs : label + valeur)
  Paragraphe 2 : "Dernière exécution : " + "N/A"  (idem)
  Paragraphe 3 : "Incident(s) : " + "N/A"         (idem)
Seul le DERNIER run de chaque paragraphe 1/2/3 (la valeur "N/A") est
remplace ; le run de label n'est jamais touche, ce qui en preserve
automatiquement la mise en forme (gras/police) -- API native python-pptx
(run.text = ...), suffisante ici car on ne modifie qu'un seul run par
paragraphe (contrairement a fill_zone_texte_9, qui reconstruit tout le
txBody car le nombre de paragraphes y est variable).
"""

TACTIC_BOX_SHAPE_NAME = "object 17"

_LABEL_NB_REGLES = "Nombre de règles"
_LABEL_DERNIERE_EXEC = "Dernière exécution"
_LABEL_INCIDENTS = "Incident(s)"

DATE_DISPLAY_FORMAT = "%d/%m/%Y"


def _paragraph_text(p) -> str:
    """Concatene le texte de tous les runs d'un paragraphe (le titre de
    tactique "Resource Development" est scinde en 2 runs "Resource " +
    "Development" dans le template -- un simple p.runs[0].text serait
    incomplet)."""
    return "".join(r.text for r in p.runs)


def _format_dt(value) -> str:
    if value is None:
        return "N/A"
    if hasattr(value, "strftime"):
        return value.strftime(DATE_DISPLAY_FORMAT)
    return str(value)


def _set_value_run(paragraph, value: str):
    """Remplace le texte du DERNIER run d'un paragraphe (la valeur "N/A"),
    en preservant la mise en forme (police/gras) de ce run -- cf docstring
    du module. Ne fait rien si le paragraphe n'a aucun run (cas inattendu,
    template modifie depuis cette ecriture)."""
    runs = paragraph.runs
    if not runs:
        return
    runs[-1].text = value


def _iter_tactic_boxes(slide):
    """Itere les encadres de tactique de la slide (cf TACTIC_BOX_SHAPE_
    NAME), en ignorant tout shape qui n'aurait pas la structure attendue
    (titre + au moins 3 paragraphes de detail) -- robustesse si le
    template evolue."""
    for shape in slide.shapes:
        if shape.name != TACTIC_BOX_SHAPE_NAME or not shape.has_text_frame:
            continue
        if len(shape.text_frame.paragraphs) < 4:
            continue
        yield shape


def fill_dispositif_surveillance_slide(slide, tactic_stats: dict = None,
                                        rules_by_tactic: dict = None) -> tuple:
    """
    Remplit les encadres de tactique MITRE ATT&CK de la 5e slide du
    template avec les valeurs calculees.

    tactic_stats    : dict {libelle_tactique: {"incident_count": int,
                      "last_incident": datetime|None}} (cf
                      mitre_normalize.build_tactic_stats). Si fourni, les
                      lignes "Incident(s)" et "Dernière exécution" de
                      CHAQUE encadre sont renseignees -- y compris a "0"/
                      "N/A" pour une tactique absente du dict (absence
                      fiable = 0 incident ce mois-ci, puisque
                      sentinel_query.MITRE_TACTICS_QUERY_TEMPLATE n'a
                      aucun filtre hormis la date). Si None (echec de la
                      requete), ces 2 lignes restent a "N/A" (valeur par
                      defaut du template) : un "N/A" honnete plutot qu'un
                      "0" qui suggererait une donnee fiable alors que la
                      requete a echoue.

    rules_by_tactic : dict {libelle_tactique: nombre_de_regles_activees}
                      (cf mitre_normalize.build_rule_counts), BONUS
                      optionnel (cf sentinel_query.fetch_active_rules_by_
                      tactic). Si None (fonctionnalite indisponible ou
                      echouee), la ligne "Nombre de règles" reste a "N/A".
                      Meme logique "absence dans le dict = 0" que pour
                      tactic_stats si le dict est fourni.

    Retourne (n_encadres_renseignes, tactiques_sans_encadre) :
      - n_encadres_renseignes : nombre d'encadres dont au moins une
        valeur a ete renseignee (a but de log uniquement, cf
        generate_cosec.py).
      - tactiques_sans_encadre : libelles presents dans tactic_stats (des
        incidents existent pour cette tactique) mais SANS encadre
        correspondant dans le template (ex: "Lateral movement",
        "Reconnaissance" -- tactiques non couvertes par les 12 encadres
        existants) -- a but de log uniquement, ces incidents ne sont
        simplement pas affiches sur cette slide.
    """
    matched_labels = set()
    n_filled = 0

    for box in _iter_tactic_boxes(slide):
        title = _paragraph_text(box.text_frame.paragraphs[0]).strip()
        if not title:
            continue

        filled_this_box = False

        if tactic_stats is not None:
            stats = tactic_stats.get(title, {"incident_count": 0, "last_incident": None})
            matched_labels.add(title)
            for p in box.text_frame.paragraphs[1:4]:
                label_text = _paragraph_text(p)
                if label_text.startswith(_LABEL_INCIDENTS):
                    _set_value_run(p, str(stats["incident_count"]))
                    filled_this_box = True
                elif label_text.startswith(_LABEL_DERNIERE_EXEC):
                    _set_value_run(p, _format_dt(stats["last_incident"]))
                    filled_this_box = True

        if rules_by_tactic is not None:
            n_rules = rules_by_tactic.get(title, 0)
            for p in box.text_frame.paragraphs[1:4]:
                if _paragraph_text(p).startswith(_LABEL_NB_REGLES):
                    _set_value_run(p, str(n_rules))
                    filled_this_box = True
                    break

        if filled_this_box:
            n_filled += 1

    unmatched = sorted(set(tactic_stats) - matched_labels) if tactic_stats is not None else []
    return n_filled, unmatched
