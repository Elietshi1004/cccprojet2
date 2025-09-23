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
    
    NOUVELLES RÈGLES APPLIQUÉES :
    1. Filtrage par concordance des détecteurs (cispr.avg/lim.avg ✅, peak/lim.q-peak ❌)
    2. Sélection des meilleures marges par position d'antenne et polarisation
    3. Arrondi spécial : Négatif = tronquer, Positif = arrondir supérieur
    4. Calcul de la marge : Margin (dB) = Limite - Mesure
    5. Verdict de conformité : OK si Margin ≥ 0, NOK sinon
    
    Args:
        measurements (list): Liste des mesures brutes extraites
    
    Returns:
        list: Liste des mesures traitées avec règles métier appliquées
    """
    processed = []

    # ========================================================================
    # ÉTAPE 1: FILTRAGE PAR CONCORDANCE DES DÉTECTEURS
    # ========================================================================
    
    concordant_measurements = []
    print(f"\n=== ANALYSE DE CONCORDANCE DES DÉTECTEURS ===")
    print(f"Nombre de mesures à analyser: {len(measurements)}")
    
    for i, m in enumerate(measurements):
        detector = str(m.get("Detector type", "")).lower()
        print(f"\n--- Mesure {i+1} ---")
        print(f"Détecteur: '{detector}'")
        print(f"Colonnes disponibles: {list(m.keys())}")
        
        # Vérifier la concordance en regardant les colonnes de limite disponibles
        is_concordant = False
        
        # CISPR.AVG avec Lim.Avg (concordant)
        if "cispr" in detector and "avg" in detector:
            print("  -> Type: CISPR.AVG détecté")
            # Vérifier si on a une colonne Limit Avg spécifique
            if "Limit Avg (dBµV/m)" in m.keys():
                is_concordant = True
                print("  -> ✅ CONCORDANT (CISPR.AVG + Limit Avg)")
            else:
                print("  -> ❌ NON CONCORDANT (pas de Limit Avg)")
        # Peak avec Lim.Peak (concordant)
        elif "peak" in detector and "q-peak" not in detector:
            print("  -> Type: Peak détecté")
            # Vérifier si on a une colonne Limit Peak spécifique
            if "Limit Peak (dBµV/m)" in m.keys():
                is_concordant = True
                print("  -> ✅ CONCORDANT (Peak + Limit Peak)")
            else:
                print("  -> ❌ NON CONCORDANT (pas de Limit Peak)")
        # Q-Peak avec Lim.Q-Peak (concordant)
        elif "q-peak" in detector:
            print("  -> Type: Q-Peak détecté")
            # Vérifier si on a une colonne Limit Q-Peak spécifique
            if "Limit Q-Peak (dBµV/m)" in m.keys():
                is_concordant = True
                print("  -> ✅ CONCORDANT (Q-Peak + Limit Q-Peak)")
            else:
                print("  -> ❌ NON CONCORDANT (pas de Limit Q-Peak)")
        else:
            print(f"  -> Type: '{detector}' non reconnu")
            print("  -> ❌ NON CONCORDANT (type non reconnu)")
        
        if is_concordant:
            concordant_measurements.append(m)
            print(f"  -> AJOUTÉ aux mesures concordantes")
        else:
            print(f"  -> REJETÉ (non concordant)")
    
    print(f"\n=== RÉSULTAT CONCORDANCE ===")
    
    print(f"Mesures concordantes trouvées: {len(concordant_measurements)} sur {len(measurements)}")
    
    # ========================================================================
    # ÉTAPE 2: GROUPEMENT PAR SECTION ET SÉLECTION DES MEILLEURES MARGES
    # ========================================================================
    
    print(f"\n=== GROUPEMENT PAR SECTION ===")
    
    # Grouper par section (détecteur + limite)
    grouped_by_section = {}
    for m in concordant_measurements:
        detector = m.get("Detector type", "")
        
        # Créer une clé de section basée sur le détecteur et la limite
        if "cispr" in detector.lower() and "avg" in detector.lower():
            section_key = "CISPR.AVG/Limit Avg"
        elif "peak" in detector.lower() and "q-peak" not in detector.lower():
            section_key = "Peak/Limit Peak"
        elif "q-peak" in detector.lower():
            section_key = "Q-Peak/Limit Q-Peak"
        else:
            section_key = f"{detector}/Unknown"
        
        if section_key not in grouped_by_section:
            grouped_by_section[section_key] = []
        grouped_by_section[section_key].append(m)
        print(f"Mesure ajoutée à la section: {section_key}")
    
    print(f"Sections créées: {list(grouped_by_section.keys())}")
    
    # Sélectionner la meilleure marge pour chaque section
    best_measurements = []
    print(f"\n=== SÉLECTION DES MEILLEURES MARGES PAR SECTION ===")
    
    for section_key, group in grouped_by_section.items():
        print(f"\n--- Section: {section_key} ---")
        print(f"Nombre de mesures dans cette section: {len(group)}")
        
        if not group:
            continue
            
        # Calculer les marges et trouver la meilleure
        best_measurement = None
        best_margin = float('-inf')
        
        for i, m in enumerate(group):
            print(f"  Mesure {i+1}:")
            print(f"    Frequency: {m.get('Frequency (MHz)', 'N/A')}")
            print(f"    Detector: {m.get('Detector type', 'N/A')}")
            print(f"    Mesure: {m.get('Mesure (dBµV/m)', 'N/A')}")
            
            # Utiliser la marge déjà calculée du fichier RAW
            try:
                marge_existante = m.get("Margin (dB)", None)
                
                if marge_existante is not None and marge_existante != "-":
                    # La marge est déjà calculée dans le fichier RAW
                    marge_brute = float(marge_existante)
                    print(f"    Marge existante du fichier RAW: {marge_brute}")
                else:
                    # Fallback : calculer la marge si elle n'existe pas
                    mesure = float(m.get("Mesure (dBµV/m)", 0))
                    
                    # Trouver la bonne colonne de limite selon le détecteur
                    detector = m.get("Detector type", "").lower()
                    if "cispr" in detector and "avg" in detector:
                        limite = float(m.get("Limit Avg (dBµV/m)", 0))
                        print(f"    Limite (Limit Avg): {limite}")
                    elif "peak" in detector and "q-peak" not in detector:
                        limite = float(m.get("Limit Peak (dBµV/m)", 0))
                        print(f"    Limite (Limit Peak): {limite}")
                    elif "q-peak" in detector:
                        limite = float(m.get("Limit Q-Peak (dBµV/m)", 0))
                        print(f"    Limite (Limit Q-Peak): {limite}")
                    else:
                        limite = float(m.get("Limite (dBµV/m)", 0))  # Fallback
                        print(f"    Limite (générique): {limite}")
                    
                    marge_brute = limite - mesure
                    print(f"    Marge calculée: {marge_brute}")
                
                # Appliquer l'arrondi spécial selon les nouvelles règles
                if marge_brute < 0:
                    # Négatif : tronquer (ex: -30.75 → -30)
                    marge = int(marge_brute)
                    print(f"    Marge arrondie (négatif): {marge}")
                else:
                    # Positif : arrondir supérieur (ex: 30.75 → 31)
                    marge = int(marge_brute + 0.999999) if marge_brute != int(marge_brute) else int(marge_brute)
                    print(f"    Marge arrondie (positif): {marge}")
                
                # Garder la meilleure marge (comparaison correcte)
                if marge > best_margin:
                    best_margin = marge
                    best_measurement = dict(m)
                    # Stocker la marge pour la sélection (en négatif)
                    best_measurement["Margin (dB)"] = marge
                    print(f"    -> ✅ NOUVELLE MEILLEURE MARGE: {marge}")
                else:
                    print(f"    -> ❌ Marge inférieure: {marge} <= {best_margin}")
                    
            except (ValueError, TypeError) as e:
                print(f"    -> ❌ ERREUR calcul marge: {e}")
                continue
        
        if best_measurement:
            # Conversion absolue pour l'affichage final si limite > mesure
            marge_finale = best_measurement["Margin (dB)"]
            
            # Vérifier si limite > mesure (marge négative)
            try:
                mesure = float(best_measurement.get("Mesure (dBµV/m)", 0))
                detector = best_measurement.get("Detector type", "").lower()
                
                if "cispr" in detector and "avg" in detector:
                    limite = float(best_measurement.get("Limit Avg (dBµV/m)", 0))
                elif "peak" in detector and "q-peak" not in detector:
                    limite = float(best_measurement.get("Limit Peak (dBµV/m)", 0))
                elif "q-peak" in detector:
                    limite = float(best_measurement.get("Limit Q-Peak (dBµV/m)", 0))
                else:
                    limite = float(best_measurement.get("Limite (dBµV/m)", 0))
                
                # Si limite > mesure (marge négative), convertir en absolu pour l'affichage
                if limite > mesure and marge_finale < 0:
                    marge_finale = abs(marge_finale)
                    print(f"    -> Conversion absolue: {best_measurement['Margin (dB)']} → {marge_finale}")
                
            except (ValueError, TypeError):
                pass  # Garder la marge originale si erreur
            
            best_measurement["Margin (dB)"] = marge_finale
            best_measurements.append(best_measurement)
            print(f"  -> MEILLEURE MESURE SÉLECTIONNÉE pour {section_key}: marge = {best_margin} → affichage = {marge_finale}")
        else:
            print(f"  -> AUCUNE MESURE VALIDE DANS CETTE SECTION")
    
    print(f"\n=== RÉSULTAT SÉLECTION ===")
    print(f"Meilleures mesures sélectionnées: {len(best_measurements)}")
    for i, m in enumerate(best_measurements):
        print(f"  {i+1}. {m.get('Detector type', 'N/A')} - Marge: {m.get('Margin (dB)', 'N/A')}")
    
    # ========================================================================
    # ÉTAPE 3: TRAITEMENT FINAL ET VERDICTS
    # ========================================================================
    
    for m in best_measurements:
        # Créer une copie de la mesure pour éviter de modifier l'original
        new_row = dict(m)

        # ========================================================================
        # VERDICT DE CONFORMITÉ
        # ========================================================================
        
        marge = new_row.get("Margin (dB)", 0)
        if isinstance(marge, (int, float)) and marge >= 0:
            conformity = "OK"
        elif isinstance(marge, (int, float)):
            conformity = "NOK"
        else:
            conformity = "-"

        new_row["Conformity"] = conformity
        new_row["Overtaking (dB)"] = "-"  # Non implémenté dans cette version
        
        # ========================================================================
        # PRÉSERVER LES COLONNES IMPORTANTES
        # ========================================================================
        
        # S'assurer que les colonnes importantes sont préservées
        if "SR" not in new_row:
            new_row["S R"] = m.get("S R", "-")
        if "Correction (dB)" not in new_row:
            new_row["Correction (dB)"] = m.get("Correction (dB)", "-")
        if "Polarization" not in new_row:
            new_row["Polarization"] = m.get("Polarization", "-")
        if "Frequency (MHz)" not in new_row:
            new_row["Frequency (MHz)"] = m.get("Frequency (MHz)", "-")
        if "Mesure (dBµV/m)" not in new_row:
            new_row["Mesure (dBµV/m)"] = m.get("Mesure (dBµV/m)", "-")
        if "Limite (dBµV/m)" not in new_row:
            new_row["Limite (dBµV/m)"] = m.get("Limite (dBµV/m)", "-")
        if "Detector type" not in new_row:
            new_row["Detector type"] = m.get("Detector type", "-")
        if "Section" not in new_row:
            new_row["Section"] = m.get("Section", "-")
        if "Sample ID" not in new_row:
            new_row["Sample ID"] = m.get("Sample ID", "-")
        if "Comment" not in new_row:
            new_row["Comment"] = m.get("Comment", "-")
        if "Antenna Position" not in new_row:
            new_row["Antenna Position"] = m.get("Antenna Position", "1 (X)")
        
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
