#!/usr/bin/env python3
"""
Script de debug pour vérifier l'extraction des paramètres de test
"""

from parser_mod import extract_data

def test_params_extraction():
    """Test l'extraction des paramètres de test"""
    print("=== Test d'extraction des paramètres de test ===")
    
    # Extraire les données
    data = extract_data("raw/raw01.docx")
    
    # Analyser les résultats
    for sample_id, sample_data in data.items():
        print(f"\nSample ID: {sample_id}")
        config_test_params = sample_data['config_test_params']
        
        for config_name, test_params in config_test_params.items():
            print(f"  Configuration: {config_name}")
            print(f"  Paramètres extraits: {test_params}")
            
            # Vérifier si on a des paramètres utiles
            useful_params = {k: v for k, v in test_params.items() 
                           if k not in ["Sample ID", "Configuration"] and v}
            
            if useful_params:
                print(f"  ✅ Paramètres utiles: {useful_params}")
            else:
                print(f"  ❌ Aucun paramètre utile trouvé")

if __name__ == "__main__":
    test_params_extraction()
