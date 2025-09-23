#!/usr/bin/env python3
"""
Configuration centralisée du projet
Mini-projet CCC - Automatisation du traitement des RAW DATA

Ce module centralise tous les paramètres de configuration :
1. Paramètres généraux (nom candidat, dossiers)
2. Paramètres de traitement (RBW, antennes)
3. Mappings de normalisation des colonnes
4. Seuils et règles de conformité
5. Paramètres de formatage et couleurs

Auteur: ElieTshingombe
Date: 2025
Projet: Mini-projet CCC - Automatisation du traitement des RAW DATA
"""

# ========================================================================
# CONFIGURATION GÉNÉRALE
# ========================================================================

# Nom du candidat (à modifier avec le vrai nom)
CANDIDATE_NAME = "ElieTshingombe"

# Nom du projet
PROJECT_NAME = "Mini-projet CCC - RAW DATA Processing"

# ========================================================================
# CONFIGURATION DES DOSSIERS
# ========================================================================

# Dossier contenant les fichiers RAW d'entrée
RAW_DIR = "raw"

# Dossier principal de sortie
OUTPUT_DIR = "out"

# Dossier des logs d'exécution
LOGS_DIR = "out/logs"

# ========================================================================
# CONFIGURATION DES FICHIERS À TRAITER
# ========================================================================

# Liste des fichiers RAW à traiter automatiquement
RAW_FILES = [
    "raw01.docx",
    "raw02.docx", 
    "raw03.docx"
]

# Formats de sortie
OUTPUT_FORMATS = {
    'word': '.docx',
    'csv': '.csv',
    'xlsx': '.xlsx'
}

# Mapping des colonnes pour l'extraction
COLUMN_MAPPING = {
    'frequency': ['frequency', 'freq'],
    'cispr_avg': ['cispr.avg', 'cispr avg'],
    'peak': ['peak'],
    'q_peak': ['q-peak', 'qpeak'],
    'limit_avg': ['lim.avg', 'lim avg'],
    'limit_q_peak': ['lim.q-peak', 'lim q-peak'],
    'limit_peak': ['lim.peak', 'lim peak'],
    'polarization': ['polarization', 'pol'],
    'correction': ['correction', 'corr']
}

# Règles de formatage
FORMATTING_RULES = {
    'frequency_precision': {
        'below_10mhz': 5,  # 5 décimales pour < 10 MHz
        'above_10mhz': 3   # 3 décimales pour >= 10 MHz
    },
    'db_precision': 2,  # 2 décimales pour les dB
    'margin_threshold': 0.05  # Seuil de tolérance pour les marges (dB)
}

# Couleurs pour le formatage Word
COLORS = {
    'pass': (0, 128, 0),      # Vert pour OK
    'fail': (200, 0, 0),      # Rouge pour NOK
    'header': (68, 114, 196)  # Bleu pour les en-têtes
}

# Messages d'erreur
ERROR_MESSAGES = {
    'file_not_found': "Fichier non trouvé: {}",
    'parsing_error': "Erreur lors de l'extraction des données: {}",
    'export_error': "Erreur lors de l'export: {}",
    'no_data': "Aucune donnée trouvée pour cette configuration"
}

# Configuration des logs
LOG_CONFIG = {
    'level': 'INFO',
    'format': '%(asctime)s - %(levelname)s - %(message)s',
    'date_format': '%Y-%m-%d %H:%M:%S'
}
