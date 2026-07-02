"""
Module d'anonymisation reversible pour appels Claude API.
 
Principe : chaque donnee sensible (UPN, IP, hostname, fichier, groupe) est
remplacee par un alias generique avant l'envoi a l'API. La table de
correspondance reste en memoire locale et sert a reinjecter les vraies
valeurs dans la reponse de Claude.
"""
 
import re
import ipaddress
 
 
class Anonymizer:
    """
    Anonymise et desanonymise un texte en se basant sur les entites
    extraites d'un incident (Accounts, Hosts, IPs, etc.).
 
    Usage :
        anon = Anonymizer()
        texte_anonymise = anon.anonymize(texte_original, row)
        # ... appel Claude API avec texte_anonymise ...
        texte_final = anon.deanonymize(reponse_claude)
    """
 
    def __init__(self):
        self.mapping = {}       # alias -> valeur reelle
        self.reverse_mapping = {}  # valeur reelle -> alias
        self._counters = {}
 
    def reset(self):
        """Reinitialise la table de correspondance (a faire entre 2 incidents)."""
        self.mapping.clear()
        self.reverse_mapping.clear()
        self._counters.clear()
 
    def _get_alias(self, value: str, category: str) -> str:
        """Retourne un alias existant ou en cree un nouveau pour cette valeur."""
        value = value.strip()
        if not value:
            return value
        if value in self.reverse_mapping:
            return self.reverse_mapping[value]
 
        self._counters[category] = self._counters.get(category, 0) + 1
        alias = f"{category}_{self._counters[category]}"
        self.mapping[alias] = value
        self.reverse_mapping[value] = alias
        return alias
 
    def build_mapping_from_row(self, row: dict):
        """
        Construit la table de correspondance a partir des entites d'une ligne
        d'incident (avant anonymisation du texte).
        """
        from generate_cosec import parse_json_array  # evite import circulaire au chargement
 
        entity_categories = {
            "Accounts": "USER",
            "Hosts": "HOST",
            "IPs": "IP",
            "SecurityGroups": "GROUP",
            "URLs": "URL",
            "Files": "FILE",
            "Processes": "PROCESS",
            "CloudApps": "APP",
            "Mailboxes": "MAILBOX",
        }
 
        for col, category in entity_categories.items():
            values = parse_json_array(row.get(col, ""))
            for v in values:
                v = v.strip().strip('"')
                if v:
                    self._get_alias(v, category)
 
    def anonymize(self, text: str) -> str:
        """
        Remplace toute occurrence connue (deja dans le mapping) par son alias.
        Detecte aussi les IPs et emails non encore mappes par regex de secours,
        ainsi que les fragments partiels d'UPN (partie locale avant le @).
        """
        if not text:
            return text
 
        result = text
 
        # Remplace les valeurs deja connues (entites de l'incident)
        # Trie par longueur decroissante pour eviter les remplacements partiels
        for value in sorted(self.reverse_mapping.keys(), key=len, reverse=True):
            if value and value in result:
                result = result.replace(value, self.reverse_mapping[value])
 
        # Matching partiel : un UPN connu (USER_x) peut apparaitre tronque
        # dans le commentaire, ex. "stephane.lazzaroni_iserba.fr#EXT#" alors
        # que l'entite complete est "...#EXT#@habiter365.onmicrosoft.com".
        # Le suffixe "#EXT#" peut lui-meme etre absent du commentaire
        # (ex: "dominique.gueret_mairie-villeurbanne.fr" sans #EXT#), donc on
        # genere plusieurs candidats par valeur : la partie locale complete,
        # et sa version sans le suffixe "#EXT#" si present.
        candidates = []  # liste de (fragment, alias)
        for value, alias in self.reverse_mapping.items():
            if "@" not in value:
                continue
            local_part = value.split("@", 1)[0]
            if len(local_part) >= 6:
                candidates.append((local_part, alias))
            # Variante sans le suffixe #EXT# (tres frequent sur les UPN invites Entra ID)
            if "#EXT#" in local_part:
                stripped = local_part.replace("#EXT#", "")
                if len(stripped) >= 6:
                    candidates.append((stripped, alias))
 
        # Tri par longueur de fragment decroissante pour eviter qu'un fragment
        # court (ex: "marthe.dupont") ne matche avant un fragment plus long
        # qui le contient (ex: "marthe.dupont.admin" ou "...#EXT#").
        candidates.sort(key=lambda c: len(c[0]), reverse=True)
 
        for fragment, alias in candidates:
            if fragment in result:
                result = result.replace(fragment, alias)
 
        # Filet de securite : emails non captures par les entites
        email_pattern = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
        for match in set(email_pattern.findall(result)):
            alias = self._get_alias(match, "USER")
            result = result.replace(match, alias)
 
        # Filet de securite : IPs non capturees
        ip_pattern = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")
        for match in set(ip_pattern.findall(result)):
            try:
                ipaddress.ip_address(match)
                alias = self._get_alias(match, "IP")
                result = result.replace(match, alias)
            except ValueError:
                continue
 
        return result
 
    def deanonymize(self, text: str) -> str:
        """Reinjecte les vraies valeurs a partir des alias presents dans le texte."""
        if not text:
            return text
        result = text
        for alias, value in self.mapping.items():
            result = result.replace(alias, value)
        return result
 
    def review_payload(self, payload: dict, incident_label: str = "") -> bool:
        """
        Mode DEBUG : affiche le payload anonymise qui sera envoye a Claude,
        liste la table de correspondance, et demande une validation humaine.
 
        Retourne True si l'utilisateur valide l'envoi, False sinon.
 
        payload : dict des champs anonymises (ex: {"title": ..., "comment": ...})
        """
        SEP = "=" * 70
        print()
        print(SEP)
        print(f"  VALIDATION ANONYMISATION  {('— ' + incident_label) if incident_label else ''}")
        print(SEP)
 
        # 1. Affiche le payload exact qui partira sur le reseau
        print("\n  [PAYLOAD ENVOYE A CLAUDE API]\n")
        for key, value in payload.items():
            print(f"  > {key} :")
            for line in str(value).splitlines() or [""]:
                print(f"      {line}")
            print()
 
        # 2. Affiche la table de correspondance (reste locale, JAMAIS envoyee)
        print("  [TABLE DE CORRESPONDANCE — locale, non envoyee]\n")
        if self.mapping:
            for alias, real in self.mapping.items():
                print(f"      {alias:<12} -> {real}")
        else:
            print("      (vide)")
        print()
 
        # 3. Controle de fuite : verifie qu'aucune vraie valeur ne subsiste
        leaks = self._detect_leaks(payload)
        if leaks:
            print("  [!] ALERTE : valeurs reelles potentiellement non anonymisees :")
            for leak in leaks:
                print(f"      - {leak}")
            print()
        else:
            print("  [OK] Aucune valeur reelle connue detectee dans le payload.\n")
 
        # 4. Demande de validation
        print(SEP)
        answer = input("  Valider l'envoi a Claude ? [o/N] : ").strip().lower()
        print(SEP)
        return answer in ("o", "oui", "y", "yes")
 
    def _detect_leaks(self, payload: dict) -> list:
        """
        Verifie qu'aucune valeur reelle de la table n'apparait encore en clair
        dans le payload anonymise (controle anti-fuite).
        """
        leaks = []
        joined = " ".join(str(v) for v in payload.values())
        for real_value in self.mapping.values():
            if real_value and real_value in joined:
                leaks.append(real_value)
        return leaks