#!/usr/bin/env python3
"""
Script principal pour le traitement des données RAW CEM/EMI
Transforme les documents Word bruts en documents formatés avec tableaux normalisés

Ce script orchestre l'ensemble du processus de traitement :
1. Extraction des données depuis les fichiers Word RAW
2. Application des règles métier (calcul des marges, verdicts)
3. Génération des documents de sortie (Word, CSV, Excel)

Auteur: ElieTshingombe
Date: 2025
Projet: Mini-projet CCC - Automatisation du traitement des RAW DATA
"""

import os
import logging
from datetime import datetime

from parser_mod import extract_data
from rules import process_data, compute_section_and_global
from writer import export_word, export_csv, export_xlsx, export_word_multiple_samples
from utils import init_logger


def main(input_file="raw/raw01.docx", word_out=None, csv_out=None, xlsx_out=None):
    """
    Fonction principale de traitement des données RAW CEM/EMI
    
    Args:
        input_file (str): Chemin vers le fichier Word RAW d'entrée
        word_out (str): Chemin de sortie pour le document Word formaté
        csv_out (str): Chemin de sortie pour l'export CSV
        xlsx_out (str): Chemin de sortie pour l'export Excel
    
    Returns:
        None
    """
    # Configuration du candidat (à modifier avec le vrai nom)
    candidate_name = "ElieTshingombe"  # 🔥 modifie avec ton vrai nom
    
    # Définition des chemins de sortie par défaut si non spécifiés
    if word_out is None:
        word_out = "out/Processed_RAW01.docx"
    if csv_out is None:
        csv_out = "out/Processed_RAW01.csv"
    if xlsx_out is None:
        xlsx_out = "out/Processed_RAW01.xlsx"

    # ========================================================================
    # ÉTAPE 1: INITIALISATION
    # ========================================================================
    
    # Créer les dossiers de sortie et de logs s'ils n'existent pas
    os.makedirs("out", exist_ok=True)
    logger = init_logger()

    # Log du début du traitement
    logger.info("=== Début traitement du fichier RAW ===")
    logger.info(f"Fichier : {input_file}")

    # ========================================================================
    # ÉTAPE 2: EXTRACTION DES DONNÉES
    # ========================================================================
    
    # Extraire toutes les données du fichier Word RAW
    # Retourne un dictionnaire avec tous les Sample ID et leurs configurations
    all_samples_data = extract_data(input_file)
    logger.info(f"Nombre de Sample ID trouvés : {len(all_samples_data)}")

    # ========================================================================
    # ÉTAPE 3: TRAITEMENT DES DONNÉES PAR SAMPLE ID ET CONFIGURATION
    # ========================================================================
    
    # Dictionnaires pour stocker les données traitées et les synthèses
    all_processed_data = {}  # Données traitées par Sample ID puis par Configuration
    all_summaries = {}       # Synthèses et verdicts par Sample ID puis par Configuration

    # Parcourir chaque Sample ID trouvé dans le fichier
    for sample_id, sample_data in all_samples_data.items():
        logger.info(f"Traitement du Sample ID : {sample_id}")

        # Extraire les composants des données du Sample ID
        test_params = sample_data['test_params']           # Paramètres de test (RBW, antenne, etc.)
        config_measurements = sample_data['config_measurements']  # Mesures par configuration
        configurations = sample_data['configurations']     # Liste des configurations disponibles

        logger.info(f"  - Paramètres : {len(test_params)}")
        logger.info(f"  - Configurations : {len(configurations)}")

        # Dictionnaires pour stocker les données traitées de ce Sample ID
        sample_processed_data = {}  # Données traitées par configuration
        sample_summaries = {}       # Synthèses par configuration

        # Traiter chaque configuration de ce Sample ID séparément
        for config in configurations:
            config_name = config['config_name']
            measurements = config_measurements.get(config_name, [])
            
            logger.info(f"    Configuration {config_name} : {len(measurements)} mesures")

            # ========================================================================
            # APPLICATION DES RÈGLES MÉTIER
            # ========================================================================
            
            # Appliquer les règles de traitement (nettoyage, normalisation, calculs)
            processed = process_data(measurements)
            sample_processed_data[config_name] = processed

            # Calculer la synthèse et le verdict global pour cette configuration
            summary, global_verdict = compute_section_and_global(processed)
            sample_summaries[config_name] = (summary, global_verdict)

            logger.info(f"    - Verdict global : {global_verdict}")

        # Stocker les données traitées pour ce Sample ID
        all_processed_data[sample_id] = sample_processed_data
        all_summaries[sample_id] = sample_summaries

    # ========================================================================
    # ÉTAPE 4: GÉNÉRATION DES DOCUMENTS DE SORTIE
    # ========================================================================
    
    # Générer le document Word formaté avec structure hiérarchique
    export_word_multiple_samples(
        all_samples_data, all_processed_data, all_summaries,
        word_out, candidate_name, input_file
    )
    
    # ========================================================================
    # ÉTAPE 5: EXPORT CSV ET EXCEL
    # ========================================================================
    
    # Combiner toutes les mesures de tous les Sample ID et configurations
    # pour l'export CSV/Excel (format tabulaire)
    all_measurements = []
    for sample_id, sample_processed in all_processed_data.items():
        for config_name, processed in sample_processed.items():
            for measurement in processed:
                # Ajouter les métadonnées pour identifier la source
                measurement['Sample ID'] = sample_id
                measurement['Configuration'] = config_name
                all_measurements.append(measurement)
    
    # Générer les exports CSV et Excel
    export_csv(all_measurements, csv_out)
    export_xlsx(all_measurements, xlsx_out)

    # Log de fin avec résumé des fichiers générés
    logger.info(f"Exports générés : {word_out}, {csv_out}, {xlsx_out}")
    logger.info("=== Fin traitement ===")


if __name__ == "__main__":
    main()
