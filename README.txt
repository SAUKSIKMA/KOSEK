Arborescence multi-clients COSEC
=================================

cosec/
├── scripts/                    <- code partagé, patché (generate_cosec.py accepte
│                                   maintenant template_path= / output_path=)
├── template_slide.pptx          <- A AJOUTER : copier ton template existant ici
├── clients.json                 <- config déclarative de tous les clients
├── clients/
│   └── CLIENT1/
│       ├── historique_cosec.xlsx   <- ton fichier existant, déjà copié ici
│       └── output/                 <- rapports générés, un par mois
├── run_all_clients.py            <- orchestrateur multi-clients
└── logs/                         <- logs de chaque run (créé automatiquement)

ETAPES AVANT LE PREMIER RUN
----------------------------
1. Copier ton template_slide.pptx existant à la racine de cosec/.
2. Renommer clients/CLIENT1/ avec le vrai nom du client, et ajuster
   clients.json en conséquence (name, history_excel, output_dir).
3. Renseigner workspace_id / tenant_id réels dans clients.json.
4. Dupliquer le bloc client dans clients.json pour chaque client
   supplémentaire (voir clients.json.example plus détaillé fourni
   précédemment pour un exemple à 2 clients).

COMMANDES
---------
Tous les clients, mois de juin 2026, avec mise à jour de l'historique :
    python run_all_clients.py --year 2026 --month 6 --update-history

Un seul client (ou un sous-ensemble), sans toucher aux autres :
    python run_all_clients.py --year 2026 --month 6 --only CLIENT1

Avec reformulation IA (Claude) :
    python run_all_clients.py --year 2026 --month 6 --update-history --ai

Un client isolé, en mode "ancien script" (toujours possible, inchangé) :
    cd scripts
    python generate_cosec.py --workspace-id <GUID> --year 2026 --month 6 \
        --history-excel ../clients/CLIENT1/historique_cosec.xlsx \
        --template-path ../template_slide.pptx \
        --output ../clients/CLIENT1/output/COSEC_CLIENT1_2026-06.pptx \
        --update-history
