#!/usr/bin/env python3
"""
Module d'utilitaires communs
Contient les fonctions de nettoyage, normalisation et logging

Ce module fournit :
1. Fonctions de nettoyage des données (décimales, unités, en-têtes)
2. Gestion du logging avec fichiers de trace
3. Génération de hash SHA256 pour la traçabilité
4. Mappings de normalisation des colonnes

Auteur: ElieTshingombe
Date: 2025
Projet: Mini-projet CCC - Automatisation du traitement des RAW DATA
"""

import hashlib
from datetime import datetime
import logging
import os


def clean_decimal(value: str):
    """
    Nettoie et convertit une valeur numérique
    
    Effectue les opérations de nettoyage suivantes :
    - Remplace les virgules par des points (format français → format anglais)
    - Supprime les espaces en début/fin
    - Convertit en float si possible
    
    Args:
        value (str): Valeur à nettoyer (ex: "1,25" ou " 44.00 ")
    
    Returns:
        float ou str: Valeur convertie en float, ou valeur originale si erreur
    """
    try:
        return float(value.replace(",", ".").strip())
    except ValueError:
        return value.strip()


def normalize_unit(unit: str) -> str:
    """
    Corrige les unités mal encodées dans les documents Word
    
    Les documents Word peuvent contenir des caractères mal encodés :
    - "Âµ" au lieu de "µ" (micro)
    - "dBμV" au lieu de "dBµV" (incohérence de caractères)
    - "Lim.Avg" au lieu de "Limit Avg" (normalisation)
    - "Lim.Peak" au lieu de "Limit Peak" (normalisation)
    - "Lim.Q-Peak" au lieu de "Limit Q-Peak" (normalisation)
    - "CISPR.AVG" au lieu de "CISPR Avg" (normalisation)
    
    Args:
        unit (str): Unité à corriger (ex: "dBÂµV/m")
    
    Returns:
        str: Unité corrigée (ex: "dBµV/m")
    """
    return (
        unit.replace("Âµ", "µ")
            .replace("dBμV", "dBµV")
            .replace("Lim.Avg", "Limit Avg")
            .replace("Lim.Peak", "Limit Peak")
            .replace("Lim.Q-Peak", "Limit Q-Peak")
            .replace("CISPR.AVG", "CISPR Avg")
            .replace("Peak-Lim.Peak", "Peak-Limit Peak")
            .replace("Q-Peak-Lim.Q-Peak", "Q-Peak-Limit Q-Peak")
            .replace("CISPR.AVG-Lim.Avg", "CISPR.AVG-Limit Avg")
    )


def file_hash(path):
    """
    Génère un hash SHA256 d'un fichier pour la traçabilité
    
    Utilisé pour :
    - Signer les documents générés avec l'empreinte du fichier RAW
    - Assurer la traçabilité et l'intégrité des données
    - Permettre la vérification de l'origine des données
    
    Args:
        path (str): Chemin vers le fichier à hasher
    
    Returns:
        str: Hash SHA256 en hexadécimal (64 caractères)
    """
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def init_logger(out_dir="out/logs"):
    """
    Initialise le système de logging avec fichier de trace
    
    Crée un fichier de log horodaté pour tracer l'exécution :
    - Format: run_YYYYMMDD_HHMMSS.log
    - Niveau: INFO (informations importantes)
    - Format: [timestamp] [niveau] message
    
    Args:
        out_dir (str): Dossier de sortie des logs (défaut: "out/logs")
    
    Returns:
        logging.Logger: Objet logger configuré
    """
    os.makedirs(out_dir, exist_ok=True)
    fname = os.path.join(out_dir, f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(filename=fname, level=logging.INFO,
                        format='%(asctime)s %(levelname)s: %(message)s')
    return logging.getLogger()

HEADER_MAP = {
    "limit peak": "Limit Peak (dBµV/m)",
    "limit avg": "Limit Avg (dBµV/m)", 
    "limit q-peak": "Limit Q-Peak (dBµV/m)",
    "cispr avg": "CISPR.AVG (dBµV/m)",
    "frequency": "Frequency (MHz)",
    "detector": "Detector type",
    "comment": "Comment",
    "applied limit": "Applied limit",
    "margin": "Margin (dB)",
    "peak": "Peak (dBµV/m)",
    "q-peak": "Q-Peak (dBµV/m)",
    "cispr": "CISPR.AVG (dBµV/m)",
    "antenna position": "Antenna Position",
    "polarization": "Polarization"
}

def normalize_header(h: str) -> str:
    """ Normalise les en-têtes pour éviter KeyError """
    h_clean = h.strip().lower()
    
    # Traiter les mappings les plus spécifiques en premier (ordre de priorité)
    # IMPORTANT: Les colonnes de marge doivent être détectées AVANT les colonnes de limite
    priority_mappings = [
        ("cispr avg-limit avg", "Margin (dB)"),       # CISPR.AVG-Lim.Avg (dB) - pattern exact
        ("peak-limit peak", "Margin (dB)"),           # Peak-Lim.Peak (dB)
        ("q-peak-limit q-peak", "Margin (dB)"),       # Q-Peak-Lim.Q-Peak (dB)
        ("peak-limit", "Margin (dB)"),                # Peak-Lim.Q-Peak (dB) - pattern plus général
        ("cispr-limit", "Margin (dB)"),               # CISPR.AVG-Lim.Avg (dB) - pattern plus général
        ("q-peak-limit", "Margin (dB)"),              # Q-Peak-Lim.Q-Peak (dB) - pattern plus général
        ("limit peak", "Limit Peak (dBµV/m)"),
        ("limit avg", "Limit Avg (dBµV/m)"),
        ("limit q-peak", "Limit Q-Peak (dBµV/m)"),
        ("cispr avg", "CISPR.AVG (dBµV/m)"),
        ("q-peak", "Q-Peak (dBµV/m)"),                # Q-Peak AVANT Peak pour éviter les conflits
        ("peak", "Peak (dBµV/m)"),
        ("cispr", "CISPR.AVG (dBµV/m)"),
        ("frequency", "Frequency (MHz)"),
        ("detector", "Detector type"),
        ("comment", "Comment"),
        ("applied limit", "Applied limit"),
        ("margin", "Margin (dB)"),
        ("antenna position", "Antenna Position"),
        ("polarization", "Polarization")
    ]
    
    for key, std in priority_mappings:
        if key in h_clean:
            return std
    return h  # défaut si pas trouvé
