"""
Module de reformulation de description d'incident via un LLM.

Construit un payload complet (titre, severite, MITRE, entites, commentaire
analyste) et appelle un modele pour produire une description de niveau
consultant cybersecurite.

Deux backends, choisis EN DUR par la constante BACKEND ci-dessous (pas
d'option en ligne de commande -- decision du 03/08/2026) :

  - "claude" : API Claude distante. Le payload est anonymise avant envoi
    (anonymizer.py) puis le resultat desanonymise au retour.
  - "local"  : modele `cosec-reformulateur` (derive de mistral-small3.2)
    servi par Ollama sur 127.0.0.1. Aucune donnee ne quitte la machine :
    l'anonymisation n'est pas appliquee.
"""

import os
import json
import urllib.error
import urllib.parse
import urllib.request

from anonymizer import Anonymizer


# --- Configuration (en dur, cf docstring) -----------------------------------

BACKEND = "claude"   # "claude" (anonymise) ou "local" (sans anonymisation)

# Backend "claude"
CLAUDE_MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 1024

# Backend "local" : serveur Ollama, API /api/generate
LOCAL_URL = "http://127.0.0.1:11434/api/generate"
LOCAL_MODEL = "cosec-reformulateur"
LOCAL_TIMEOUT = 180   # secondes -- un modele 24B en local peut etre lent
 
 
_PROMPT_REDACTION = """Tu es un consultant senior en cybersecurite qui redige des \
descriptions d'incidents pour une presentation client (COSEC - Comite de Securite).

Ton role : a partir des informations techniques brutes d'un incident de securite, \
produire une description claire, synthetique et professionnelle, au niveau d'un \
consultant qui presente devant un client.

Regles de redaction :
- Francais professionnel, ton factuel et mesure, sans jargon inutile.
- 2 a 4 phrases maximum. Synthese, pas de paraphrase exhaustive.
- Structure implicite : contexte de l'incident, ce qui s'est passe, conclusion \
(impact ou resolution).
- Ne JAMAIS inventer d'information absente des donnees fournies.
- Si le commentaire de l'analyste indique un faux positif ou une activite \
legitime, le refleter clairement."""

_PROMPT_ANONYMISATION = """Regle d'anonymisation CRITIQUE :
- Les identifiants comme USER_1, IP_1, HOST_1, GROUP_1, FILE_1, etc. sont des \
ALIAS anonymises. Reutilise-les EXACTEMENT tels quels dans ta reponse. \
Ne les remplace pas, ne les renomme pas, n'invente pas de nouveaux alias."""

_PROMPT_SORTIE = """Reponds UNIQUEMENT avec la description reformulee, sans preambule ni titre."""


def build_system_prompt(anonymized: bool) -> str:
    """
    Assemble le prompt systeme. La regle d'anonymisation n'est incluse que si
    le payload est reellement anonymise : en backend local, les vraies valeurs
    sont envoyees telles quelles et parler d'alias induirait le modele en erreur.
    """
    parts = [_PROMPT_REDACTION]
    if anonymized:
        parts.append(_PROMPT_ANONYMISATION)
    parts.append(_PROMPT_SORTIE)
    return "\n\n".join(parts)
 
 
def _format_entities(entities: dict) -> str:
    """Formate le dict d'entites anonymisees en texte lisible pour le prompt."""
    lines = []
    labels = {
        "accounts": "Comptes", "hosts": "Machines", "ips": "Adresses IP",
        "groups": "Groupes", "urls": "URLs", "files": "Fichiers",
        "processes": "Processus", "apps": "Applications cloud", "mailboxes": "Boites mail",
    }
    for key, label in labels.items():
        vals = entities.get(key, [])
        if vals:
            lines.append(f"  - {label} : {', '.join(vals)}")
    return "\n".join(lines) if lines else "  (aucune entite)"
 
 
def build_payload(row: dict, anon) -> dict:
    """
    Construit le payload a partir d'une ligne d'incident.

    anon : instance Anonymizer dont la table de correspondance est deja
           construite (via build_mapping_from_row), ou None pour un envoi
           en clair (backend local, cf make_anonymizer).
    """
    from generate_cosec import parse_json_array

    def scrub(value):
        return anon.anonymize(value) if anon is not None else value

    def anon_list(col):
        return [scrub(v.strip().strip('"'))
                for v in parse_json_array(row.get(col, "")) if v.strip().strip('"')]

    tactics    = ", ".join(parse_json_array(row.get("Tactics", "")))
    techniques = ", ".join(parse_json_array(row.get("Techniques", "")))
 
    payload = {
        "title":          scrub(row.get("Title", "")),
        "severity":       row.get("Severity", ""),
        "classification": row.get("Classification", ""),
        "reason":         row.get("ClassificationReason", ""),
        "tactics":        tactics,
        "techniques":     techniques,
        "entities": {
            "accounts":  anon_list("Accounts"),
            "hosts":     anon_list("Hosts"),
            "ips":       anon_list("IPs"),
            "groups":    anon_list("SecurityGroups"),
            "urls":      anon_list("URLs"),
            "files":     anon_list("Files"),
            "processes": anon_list("Processes"),
            "apps":      anon_list("CloudApps"),
            "mailboxes": anon_list("Mailboxes"),
        },
        "analyst_comment": scrub(row.get("ClassificationComment", "") or ""),
    }
    return payload
 
 
def build_user_message(payload: dict) -> str:
    """Construit le message utilisateur envoye au modele depuis le payload."""
    return f"""Voici les informations d'un incident de securite a reformuler :
 
Typologie : {payload['title']}
Severite : {payload['severity']}
Classification : {payload['classification']} ({payload['reason']})
Tactiques MITRE ATT&CK : {payload['tactics'] or 'N/A'}
Techniques MITRE ATT&CK : {payload['techniques'] or 'N/A'}
 
Entites impliquees :
{_format_entities(payload['entities'])}
 
Commentaire de l'analyste :
{payload['analyst_comment'] or '(aucun commentaire)'}
 
Redige la description de cet incident pour la presentation client."""
 
 
def _call_claude(client, system_prompt: str, user_message: str) -> str:
    """Appelle l'API Claude et retourne le texte de la reponse."""
    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )
    return "".join(
        block.text for block in response.content if block.type == "text"
    ).strip()


def _call_local(system_prompt: str, user_message: str) -> str:
    """
    Appelle le modele local via l'API /api/generate d'Ollama (stdlib urllib,
    pas de dependance supplementaire).

    Le prompt systeme est envoye explicitement : il ecrase celui eventuellement
    embarque dans le Modelfile de `cosec-reformulateur`, pour que les deux
    backends suivent les memes regles de redaction (celles de ce fichier).
    """
    body = json.dumps({
        "model": LOCAL_MODEL,
        "system": system_prompt,
        "prompt": user_message,
        "stream": False,
        "options": {"num_predict": MAX_TOKENS},
    }).encode("utf-8")

    request = urllib.request.Request(
        LOCAL_URL, data=body, headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=LOCAL_TIMEOUT) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError as e:
        raise RuntimeError(f"Appel au modele local {LOCAL_MODEL} echoue ({LOCAL_URL}) : {e}")

    return (result.get("response") or "").strip()


def reformulate_description(row: dict, anon, client,
                            debug: bool = False) -> str:
    """
    Reformule la description d'un incident via le backend configure (BACKEND).

    row    : ligne d'incident (dict CSV)
    anon   : instance Anonymizer (sa table sera reinitialisee puis reconstruite),
             ou None si le backend n'anonymise pas (cf make_anonymizer)
    client : client anthropic.Anthropic deja instancie (backend "claude"),
             None pour le backend local
    debug  : si True, demande une validation humaine avant l'envoi -- sans
             effet sans anonymisation, l'appel etant purement local

    Retourne la description reformulee (desanonymisee si elle avait ete
    anonymisee).
    """
    # 1. Reinitialise et construit la table de correspondance pour cet incident
    if anon is not None:
        anon.reset()
        anon.build_mapping_from_row(row)

    # 2. Construit le payload (anonymise si anon est fourni)
    payload = build_payload(row, anon)
    user_message = build_user_message(payload)
    system_prompt = build_system_prompt(anonymized=anon is not None)

    # 3. Mode debug : validation humaine de l'anonymisation
    if debug and anon is not None:
        review_data = {
            "title": payload["title"],
            "analyst_comment": payload["analyst_comment"],
            "message_complet": user_message,
        }
        label = row.get("IncidentName", "") + " " + row.get("Title", "")[:40]
        if not anon.review_payload(review_data, incident_label=label):
            print("  [ABANDON] Envoi annule par l'utilisateur.")
            return row.get("ClassificationComment", "") or "N/A"

    # 4. Appel du modele
    if BACKEND == "local":
        result = _call_local(system_prompt, user_message)
    else:
        result = _call_claude(client, system_prompt, user_message)

    # 5. Desanonymisation avant retour (sans objet si rien n'a ete anonymise)
    return anon.deanonymize(result) if anon is not None else result


def make_client():
    """
    Prepare le backend configure et retourne son client.

    - "claude" : client Anthropic instancie depuis la variable d'env
                 ANTHROPIC_API_KEY.
    - "local"  : verifie que le serveur Ollama repond (echec immediat plutot
                 qu'une erreur par incident) et retourne None, _call_local
                 n'ayant besoin d'aucun objet client.
    """
    if BACKEND == "local":
        base = urllib.parse.urlsplit(LOCAL_URL)
        root = f"{base.scheme}://{base.netloc}/"
        try:
            urllib.request.urlopen(root, timeout=5).close()
        except urllib.error.URLError as e:
            raise RuntimeError(
                f"Modele local injoignable sur {root} ({e}) -- demarrer Ollama, "
                f"ou basculer BACKEND sur 'claude' dans reformulate.py."
            )
        return None

    if BACKEND != "claude":
        raise RuntimeError(
            f"BACKEND inconnu dans reformulate.py : {BACKEND!r} (attendu 'claude' ou 'local')."
        )

    import anthropic   # importe ici : inutile en backend local
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("Variable d'environnement ANTHROPIC_API_KEY non definie.")
    return anthropic.Anthropic(api_key=api_key)


def make_anonymizer():
    """
    Retourne l'Anonymizer a utiliser, ou None si le backend n'en a pas besoin.

    Le backend local tourne sur 127.0.0.1 : aucune donnee ne quitte la machine,
    l'anonymisation (et sa desanonymisation) n'a donc plus d'objet.
    """
    return None if BACKEND == "local" else Anonymizer()


def backend_description() -> str:
    """Libelle du backend actif, pour la trace console de generate_cosec.py."""
    if BACKEND == "local":
        return f"modele local {LOCAL_MODEL} — sans anonymisation"
    return f"Claude API {CLAUDE_MODEL} — payload anonymise"