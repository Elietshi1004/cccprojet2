#!/usr/bin/env python3
"""
Test simple pour vérifier les corrections
"""

try:
    from parser_mod import extract_data
    print("Import réussi")
    
    # Test avec un seul fichier
    print("Extraction des données...")
    data = extract_data("raw/raw01.docx")
    
    print(f"Données extraites: {len(data)} Sample IDs")
    
    # Afficher le premier Sample ID
    if data:
        first_sample = list(data.keys())[0]
        print(f"Premier Sample: {first_sample}")
        
        sample_data = data[first_sample]
        config_measurements = sample_data['config_measurements']
        config_test_params = sample_data['config_test_params']
        
        print(f"Configurations: {list(config_measurements.keys())}")
        
        # Afficher les premières mesures de la première configuration
        if config_measurements:
            first_config = list(config_measurements.keys())[0]
            measurements = config_measurements[first_config]
            test_params = config_test_params[first_config]
            
            print(f"Première config: {first_config}")
            print(f"Nombre de mesures: {len(measurements)}")
            print(f"Paramètres de test: {test_params}")
            
            if measurements:
                print("Premières mesures:")
                for i, m in enumerate(measurements[:3]):
                    print(f"  {i+1}: Section='{m.get('Section')}', Freq='{m.get('Frequency (MHz)')}', Mesure='{m.get('Mesure (dBµV/m)')}'")
    else:
        print("Aucune donnée extraite")
        
except Exception as e:
    print(f"Erreur: {e}")
    import traceback
    traceback.print_exc()
