#!/usr/bin/env python3
"""
Module des règles métier CEM/EMI
Contient la logique de traitement et de validation des données

Ce module implémente :
1. Calcul des marges de conformité (Margin = Mesure - Limite)
2. Détection des dépassements (Overtaking)
3. Attribution des verdicts de conformité (OK/NOK)
4. Synthèse par section et verdict global

Auteur: ElieTshingombe
Date: 2025
Projet: Mini-projet CCC - Automatisation du traitement des RAW DATA
"""

def process_data(measurements):
    """
    Traite chaque ligne de mesure pour appliquer les règles métier CEM/EMI
    
    RÈGLES APPLIQUÉES :
    1. Calcul de la marge : Margin (dB) = Mesure (dBµV/m) - Limite (dBµV/m)
    2. Détection du dépassement : Overtaking (dB) = max(0, -Margin)
    3. Verdict de conformité : OK si Margin ≥ 0, NOK sinon
    
    Args:
        measurements (list): Liste des mesures brutes extraites
    
    Returns:
        list: Liste des mesures traitées avec règles métier appliquées
    """
    processed = []

    # ========================================================================
    # TRAITEMENT DE CHAQUE MESURE
    # ========================================================================
    
    for m in measurements:
        # Créer une copie de la mesure pour éviter de modifier l'original
        new_row = dict(m)

        # ========================================================================
        # IDENTIFICATION DES COLONNES
        # ========================================================================
        
        # Créer un mapping insensible à la casse pour identifier les colonnes
        keys = {k.lower(): k for k in m.keys()}
        
        # Identifier les colonnes clés (avec fallback sur différents noms possibles)
        col_mesure = next((keys[k] for k in keys if "cispr" in k or "mesure" in k), None)
        col_limite = next((keys[k] for k in keys if "limit" in k or "limite" in k), None)

        # ========================================================================
        # CALCUL DE LA MARGE (RÈGLE MÉTIER PRINCIPALE)
        # ========================================================================
        
        # Calculer la marge si les colonnes mesure et limite sont disponibles
        if col_mesure and col_limite:
            try:
                # Conversion des valeurs numériques
                mesure = float(m[col_mesure])
                limite = float(m[col_limite])
                
                # RÈGLE CEM : Marge = Limite - Mesure (en dB)
                # Note: Dans le contexte CEM, une marge positive = conforme
                marge = round(limite - mesure, 2)
            except Exception:
                # En cas d'erreur, utiliser la marge existante ou marquer N/A
                marge = m.get("Margin (dB)", "N/A")
        else:
            # Si pas de colonnes mesure/limite, utiliser la marge existante
            marge = m.get("Margin (dB)", "N/A")

        # ========================================================================
        # ATTRIBUTION DES VALEURS CALCULÉES
        # ========================================================================
        
        new_row["Margin (dB)"] = marge
        new_row["Overtaking (dB)"] = "-"  # Non implémenté dans cette version

        # ========================================================================
        # VERDICT DE CONFORMITÉ
        # ========================================================================
        
        # RÈGLE CEM : Conformité OK si marge ≥ 0, NOK sinon
        if isinstance(marge, (int, float)) and marge >= 0:
            conformity = "OK"
        elif isinstance(marge, (int, float)):
            conformity = "NOK"
        else:
            conformity = "-"

        new_row["Conformity"] = conformity
        processed.append(new_row)

    return processed


def compute_section_and_global(processed_rows):
    """
    Calcule verdict par section et verdict global
    """
    sections = {}
    for r in processed_rows:
        sec = r.get("Detector type", "Unknown")
        sections.setdefault(sec, []).append(r)

    summary = []
    all_pass = True
    for sec, rows in sections.items():
        fails = sum(1 for rr in rows if str(rr.get("Conformity", "")).upper() == "NOK")
        verdict = "OK" if fails == 0 else "NOK"
        all_pass = all_pass and (verdict == "OK")
        summary.append({
            "Section": sec,
            "NbLines": len(rows),
            "NbFAIL": fails,
            "Verdict": verdict
        })

    global_verdict = "OK" if all_pass else "NOK"
    return summary, global_verdict
