# Mini-projet CCC : Automatisation du traitement des RAW DATA

## 🎯 Objectif
Développer un script automatisé capable de transformer des documents bruts Word (RAW DATA) en un document Word formaté (Processed DATA) contenant des tableaux normalisés et exploitables pour les essais CEM/EMI.

## 📁 Structure du projet
```
cccprojet2/
├── README.md                 # Ce fichier
├── main.py                   # Point d'entrée principal
├── parser_mod.py             # Extraction des données depuis Word
├── writer.py                 # Génération des documents de sortie
├── rules.py                  # Règles métier (marges, conformité)
├── utils.py                  # Utilitaires (nettoyage, hash, logs)
├── questions.txt             # Spécifications du projet
├── raw/                      # Fichiers d'entrée
│   ├── raw01.docx
│   ├── raw02.docx
│   └── raw03.docx
├── out/                      # Résultats générés
│   ├── Processed_RAW01.docx
│   ├── Processed_RAW01.csv
│   ├── Processed_RAW01.xlsx
│   └── logs/
└── venv/                     # Environnement virtuel Python
```

## 🚀 Installation et exécution

### Prérequis
- Python 3.9+
- pip (gestionnaire de paquets Python)

### Installation
1. **Cloner le projet**
   ```bash
   git clone https://github.com/Elietshi1004/cccprojet2.git
   cd cccprojet2
   ```

2. **Créer l'environnement virtuel**
   ```bash
   python -m venv venv
   ```

3. **Activer l'environnement virtuel**
   - Windows : `.\venv\Scripts\Activate.ps1`
   - Linux/Mac : `source venv/bin/activate`

4. **Installer les dépendances**
   ```bash
   pip install python-docx pandas openpyxl
   ```

### Exécution
```bash
python main.py
```

## 🔧 Choix techniques

### Langage et bibliothèques
- **Python 3.9** : Langage principal
- **python-docx** : Manipulation des documents Word
- **pandas** : Traitement des données et export CSV/Excel
- **openpyxl** : Support Excel (.xlsx)
- **re** : Expressions régulières pour le parsing
- **hashlib** : Génération de hash SHA256

### Architecture
- **Modulaire** : Séparation des responsabilités
  - `main.py` : Orchestration générale
  - `parser_mod.py` : Extraction des données
  - `writer.py` : Génération des documents
  - `rules.py` : Règles métier
  - `utils.py` : Utilitaires communs

### Logique de traitement
1. **Extraction** : Parcours des tableaux Word pour identifier les configurations
2. **Association** : Liaison des données de mesures avec leurs configurations
3. **Normalisation** : Nettoyage et standardisation des données
4. **Calculs** : Application des règles métier (marges, verdicts)
5. **Export** : Génération des documents Word, CSV et Excel

## 📊 Fonctionnalités implémentées

### ✅ Extraction des données
- Détection automatique des sections (CISPR.AVG, Q-Peak, Peak)
- Extraction des paramètres de test (Sample ID, RBW, antenne, etc.)
- Conservation de toutes les lignes chiffrées
- Normalisation des unités (dBÂµV/m → dBµV/m, virgules → points)

### ✅ Logique métier
- Calcul des marges : `Marge (dB) = Limite – Mesure`
- Verdict ligne : PASS si marge ≥ 0, FAIL si marge < 0
- Formatage des fréquences (5 décimales < 10 MHz, 3 décimales sinon)
- Formatage des dB (2 décimales)

### ✅ Génération de documents
- Structure hiérarchique (Sample ID → Configuration → Tableaux)
- Colonnes normalisées selon spécifications
- Mise en forme professionnelle (couleurs, formatage)
- Signature automatique en pied de page

### ✅ Exports multiples
- Document Word formaté (.docx)
- Export CSV (.csv)
- Export Excel (.xlsx)
- Logs d'exécution

## 🎨 Format de sortie

### Document Word
- **En-tête** : "1.1 Test results"
- **Structure** : Sample ID → Configuration → Tableau de données
- **Colonnes** : Section | Frequency (MHz) | SR | Polarization | Correction (dB) | Mesure (dBµV/m) | Limite (dBµV/m) | Marge (dB) | Verdict
- **Formatage** : OK en vert, NOK en rouge, tableaux professionnels

### Exports CSV/Excel
- Même structure de données
- Format tabulaire exploitable
- Colonnes normalisées

## 🔍 Gestion des erreurs

### Robustesse implémentée
- Gestion des encodages de fichiers
- Validation des données extraites
- Logs détaillés pour le debugging
- Gestion des configurations sans données

### Logs
- Fichiers de log dans `out/logs/`
- Timestamp des exécutions
- Traçabilité des opérations

*Projet réalisé dans le cadre du mini-projet CCC - Automatisation du traitement des RAW DATA*
