"""
Orchestrateur multi-clients COSEC.

Lit clients.json et génère le rapport PowerPoint (+ mise à jour de
l'historique Excel si --update-history) pour chaque client déclaré,
dans un seul process Python (bénéfice : le cache de credentials Azure
keyé par tenant_id, déjà présent dans generate_cosec.py, est réutilisé
d'un client à l'autre quand ils partagent le même tenant).

Usage :
    python run_all_clients.py --year 2026 --month 6
    python run_all_clients.py --year 2026 --month 6 --update-history
    python run_all_clients.py --year 2026 --month 6 --only CLIENT1,CLIENT2
"""

import argparse
import json
import logging
import os
import sys
import traceback
from datetime import datetime

# Le dossier scripts/ (code partagé) doit être sur le PYTHONPATH.
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts"))

from generate_cosec import generate_pptx  # noqa: E402  (import après sys.path)

CONFIG_PATH = "clients.json"
TEMPLATE_PATH = "template_slide.pptx"
LOG_DIR = "logs"


def load_clients(config_path: str) -> list[dict]:
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def setup_logging(year: int, month: int) -> str:
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"run_{year}-{month:02d}_{datetime.now():%Y%m%d_%H%M%S}.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )
    return log_path


def run_for_client(client: dict, year: int, month: int, update_history: bool,
                    use_ai: bool, template_path: str) -> bool:
    name = client["name"]
    history_excel = client["history_excel"]
    output_dir = client.get("output_dir", os.path.join("clients", name, "output"))
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"COSEC_{name}_{year}-{month:02d}.pptx")

    logging.info(f"--- Début génération COSEC pour {name} ---")
    try:
        generate_pptx(
            workspace_id=client["workspace_id"],
            year=year,
            month=month,
            tenant_id=client.get("tenant_id"),
            use_ai=use_ai,
            debug=False,
            update_history=update_history,
            history_excel=history_excel,
            evolution_slide=not client.get("no_evolution_slide", False),
            surveillance_slide=not client.get("no_surveillance_slide", False),
            sla_slide=not client.get("no_sla_slide", False),
            dispositif_slide=not client.get("no_dispositif_slide", False),
            log_ingestion_slide=not client.get("no_log_ingestion_slide", False),
            price_per_gb=client.get("price_per_gb", 4.89),
            template_path=template_path,
            output_path=output_path,
        )
        logging.info(f"✅ {name} : rapport généré -> {output_path}")
        return True
    except Exception as e:
        logging.error(f"❌ {name} : échec de la génération — {e}")
        logging.error(traceback.format_exc())
        return False


def main():
    parser = argparse.ArgumentParser(description="Génère les rapports COSEC pour plusieurs clients.")
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--config", default=CONFIG_PATH, help=f"Chemin du fichier clients.json (défaut : {CONFIG_PATH})")
    parser.add_argument("--template", default=TEMPLATE_PATH, help=f"Template pptx partagé (défaut : {TEMPLATE_PATH})")
    parser.add_argument("--update-history", action="store_true")
    parser.add_argument("--ai", action="store_true")
    parser.add_argument("--only", default=None,
                         help="Liste de noms de clients séparés par des virgules, pour ne traiter qu'un sous-ensemble")
    args = parser.parse_args()

    log_path = setup_logging(args.year, args.month)
    logging.info(f"Log de ce run : {log_path}")

    clients = load_clients(args.config)
    if args.only:
        wanted = {n.strip() for n in args.only.split(",")}
        clients = [c for c in clients if c["name"] in wanted]

    results = {}
    for client in clients:
        results[client["name"]] = run_for_client(
            client, args.year, args.month, args.update_history, args.ai, args.template
        )

    ok = [n for n, success in results.items() if success]
    ko = [n for n, success in results.items() if not success]
    logging.info("=" * 60)
    logging.info(f"Bilan du run {args.year}-{args.month:02d} : {len(ok)} réussi(s), {len(ko)} échec(s).")
    if ok:
        logging.info(f"  ✅ OK  : {', '.join(ok)}")
    if ko:
        logging.info(f"  ❌ KO  : {', '.join(ko)}")

    sys.exit(1 if ko else 0)


if __name__ == "__main__":
    main()
