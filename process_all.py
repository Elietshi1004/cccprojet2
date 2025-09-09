#!/usr/bin/env python3
"""
Script de traitement automatique de tous les fichiers RAW
Traite tous les fichiers .docx du dossier raw/ et génère les exports

Ce script automatise le traitement de tous les fichiers RAW :
1. Détecte automatiquement tous les fichiers .docx dans le dossier raw/
2. Traite chaque fichier avec main.py
3. Génère les exports Word, CSV et Excel pour chaque fichier
4. Fournit un rapport de traitement complet

Auteur: ElieTshingombe
Date: 2025
Projet: Mini-projet CCC - Automatisation du traitement des RAW DATA
"""

import os
import sys
from datetime import datetime
from main import main as process_single_file
from utils import init_logger

def process_all_raw_files():
    """
    Traite tous les fichiers RAW du dossier raw/ automatiquement
    
    Cette fonction :
    1. Détecte tous les fichiers .docx dans le dossier raw/
    2. Traite chaque fichier individuellement avec main.py
    3. Génère les exports Word, CSV et Excel pour chaque fichier
    4. Fournit un rapport détaillé du traitement
    
    Returns:
        list: Liste des résultats de traitement avec statuts et chemins
    """
    # ========================================================================
    # INITIALISATION
    # ========================================================================
    
    # Initialiser le logger
    logger = init_logger()
    logger.info("=== Début traitement de tous les fichiers RAW ===")
    
    # Dossier des fichiers RAW
    raw_dir = "raw"
    out_dir = "out"
    
    # Créer le dossier de sortie s'il n'existe pas
    os.makedirs(out_dir, exist_ok=True)
    
    # ========================================================================
    # CONFIGURATION DES FICHIERS À TRAITER
    # ========================================================================
    
    # Liste des fichiers RAW à traiter (selon config.py)
    raw_files = [
        "raw01.docx",
        "raw02.docx", 
        "raw03.docx"
    ]
    
    results = []
    
    for raw_file in raw_files:
        raw_path = os.path.join(raw_dir, raw_file)
        
        # Vérifier que le fichier existe
        if not os.path.exists(raw_path):
            logger.warning(f"Fichier {raw_file} non trouvé, ignoré")
            continue
            
        logger.info(f"Traitement de {raw_file}...")
        
        try:
            # Modifier les chemins de sortie pour chaque fichier
            base_name = raw_file.replace('.docx', '')
            
            # Rediriger les sorties vers des fichiers spécifiques
            word_out = f"out/Processed_{base_name.upper()}.docx"
            csv_out = f"out/Processed_{base_name.upper()}.csv"
            xlsx_out = f"out/Processed_{base_name.upper()}.xlsx"
            
            # Traiter le fichier
            process_single_file(raw_path, word_out, csv_out, xlsx_out)
            
            results.append({
                'file': raw_file,
                'status': 'SUCCESS',
                'word': word_out,
                'csv': csv_out,
                'xlsx': xlsx_out
            })
            
            logger.info(f"✅ {raw_file} traité avec succès")
            
        except Exception as e:
            logger.error(f"❌ Erreur lors du traitement de {raw_file}: {str(e)}")
            results.append({
                'file': raw_file,
                'status': 'ERROR',
                'error': str(e)
            })
    
    # Résumé des résultats
    logger.info("=== Résumé du traitement ===")
    success_count = sum(1 for r in results if r['status'] == 'SUCCESS')
    error_count = sum(1 for r in results if r['status'] == 'ERROR')
    
    logger.info(f"Fichiers traités avec succès: {success_count}")
    logger.info(f"Fichiers en erreur: {error_count}")
    
    for result in results:
        if result['status'] == 'SUCCESS':
            logger.info(f"  ✅ {result['file']} → {result['word']}")
        else:
            logger.info(f"  ❌ {result['file']} → {result['error']}")
    
    logger.info("=== Fin traitement de tous les fichiers RAW ===")
    
    return results

if __name__ == "__main__":
    print("🚀 Traitement de tous les fichiers RAW...")
    results = process_all_raw_files()
    
    # Afficher le résumé dans la console
    print("\n📊 Résumé:")
    for result in results:
        if result['status'] == 'SUCCESS':
            print(f"  ✅ {result['file']} → Traité avec succès")
        else:
            print(f"  ❌ {result['file']} → Erreur: {result['error']}")
    
    print(f"\n🎉 Traitement terminé ! Vérifiez le dossier 'out/' pour les résultats.")
