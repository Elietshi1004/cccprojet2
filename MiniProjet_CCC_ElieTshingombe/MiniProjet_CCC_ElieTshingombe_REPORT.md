
# Rapport de déploiement - Mini-projet CCC

**Candidat:** Elie Tshingombe
**Date:** 2025-09-13 02:54:42
**Archive:** MiniProjet_CCC_ElieTshingombe.zip

## Structure livrée

```
MiniProjet_CCC_ElieTshingombe/
├── README.md                    # Documentation complète
├── src/                         # Code source
│   ├── main.py                  # Point d'entrée principal
│   ├── parser_mod.py            # Extraction des données
│   ├── writer.py                # Génération des documents
│   ├── rules.py                 # Règles métier
│   ├── utils.py                 # Utilitaires
│   └── config.py                # Configuration
├── tests/                       # Tests et scripts
│   ├── test_validation.py       # Script de validation
│   └── process_all.py           # Traitement par lot
├── out/                         # Résultats générés
│   ├── Processed_RAW01.docx     # Document Word RAW01
│   ├── Processed_RAW01.csv      # Export CSV RAW01
│   ├── Processed_RAW01.xlsx     # Export Excel RAW01
│   ├── Processed_RAW02.docx     # Document Word RAW02
│   ├── Processed_RAW02.csv      # Export CSV RAW02
│   ├── Processed_RAW02.xlsx     # Export Excel RAW02
│   ├── Processed_RAW03.docx     # Document Word RAW03
│   ├── Processed_RAW03.csv      # Export CSV RAW03
│   ├── Processed_RAW03.xlsx     # Export Excel RAW03
│   └── logs/                    # Logs d'exécution
└── config/                      # Configuration
    └── questions.txt            # Spécifications du projet
```
## Instructions d'utilisation

1. Extraire l'archive
2. Installer les dépendances: `pip install python-docx pandas openpyxl`
3. Exécuter: `python src/main.py`
4. Ou traiter tous les fichiers: `python tests/process_all.py`

---
*Projet réalisé dans le cadre du mini-projet CCC - Automatisation du traitement des RAW DATA*
