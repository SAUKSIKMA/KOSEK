"""
Module de reformulation de description d'incident via Claude API.
 
Construit un payload complet anonymise (titre, severite, MITRE, entites,
commentaire analyste), appelle Claude pour produire une description de
niveau consultant cybersecurite, puis desanonymise le resultat.
"""
 
import os
import json
import anthropic
 
from anonymizer import Anonymizer
 
 
MODEL = "claude-sonnet-4-6"   # ou claude-sonnet-4-6 pour un cout/latence reduit
MAX_TOKENS = 1024
 
 
SYSTEM_PROMPT = """Tu es un consultant senior en cybersecurite qui redige des \
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
legitime, le refleter clairement.
 
Regle d'anonymisation CRITIQUE :
- Les identifiants comme USER_1, IP_1, HOST_1, GROUP_1, FILE_1, etc. sont des \
ALIAS anonymises. Reutilise-les EXACTEMENT tels quels dans ta reponse. \
Ne les remplace pas, ne les renomme pas, n'invente pas de nouveaux alias.
 
Reponds UNIQUEMENT avec la description reformulee, sans preambule ni titre."""
 
 
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
 
 
def build_payload(row: dict, anon: Anonymizer) -> dict:
    """
    Construit le payload anonymise a partir d'une ligne d'incident.
    La table de correspondance de `anon` doit deja etre construite
    (via build_mapping_from_row).
    """
    from generate_cosec import parse_json_array
 
    def anon_list(col):
        return [anon.anonymize(v.strip().strip('"'))
                for v in parse_json_array(row.get(col, "")) if v.strip().strip('"')]
 
    tactics    = ", ".join(parse_json_array(row.get("Tactics", "")))
    techniques = ", ".join(parse_json_array(row.get("Techniques", "")))
 
    payload = {
        "title":          anon.anonymize(row.get("Title", "")),
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
        "analyst_comment": anon.anonymize(row.get("ClassificationComment", "") or ""),
    }
    return payload
 
 
def build_user_message(payload: dict) -> str:
    """Construit le message utilisateur envoye a Claude depuis le payload."""
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
 
 
def reformulate_description(row: dict, anon: Anonymizer, client,
                            debug: bool = False) -> str:
    """
    Reformule la description d'un incident via Claude API.
 
    row    : ligne d'incident (dict CSV)
    anon   : instance Anonymizer (sa table sera reinitialisee puis reconstruite)
    client : client anthropic.Anthropic deja instancie
    debug  : si True, demande une validation humaine avant l'envoi
 
    Retourne la description reformulee et desanonymisee.
    """
    # 1. Reinitialise et construit la table de correspondance pour cet incident
    anon.reset()
    anon.build_mapping_from_row(row)
 
    # 2. Construit le payload anonymise
    payload = build_payload(row, anon)
    user_message = build_user_message(payload)
 
    # 3. Mode debug : validation humaine de l'anonymisation
    if debug:
        review_data = {
            "title": payload["title"],
            "analyst_comment": payload["analyst_comment"],
            "message_complet": user_message,
        }
        label = row.get("IncidentName", "") + " " + row.get("Title", "")[:40]
        if not anon.review_payload(review_data, incident_label=label):
            print("  [ABANDON] Envoi annule par l'utilisateur.")
            return row.get("ClassificationComment", "") or "N/A"
 
    # 4. Appel Claude API
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
 
    # 5. Extraction du texte
    anonymized_result = "".join(
        block.text for block in response.content if block.type == "text"
    ).strip()
 
    # 6. Desanonymisation avant retour
    return anon.deanonymize(anonymized_result)
 
 
def make_client():
    """Instancie le client Anthropic depuis la variable d'env ANTHROPIC_API_KEY."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("Variable d'environnement ANTHROPIC_API_KEY non definie.")
    return anthropic.Anthropic(api_key=api_key)