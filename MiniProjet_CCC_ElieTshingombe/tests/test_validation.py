#!/usr/bin/env python3
"""
Script de validation du projet CCC
Vérifie que tous les éléments requis sont présents et fonctionnels
"""

import os
import sys
from datetime import datetime

def check_file_structure():
    """Vérifie la structure des fichiers"""
    print("🔍 Vérification de la structure des fichiers...")
    
    required_files = [
        'main.py',
        'parser_mod.py', 
        'writer.py',
        'rules.py',
        'utils.py',
        'config.py',
        'README.md',
        'questions.txt'
    ]
    
    required_dirs = [
        'raw',
        'out',
        'out/logs'
    ]
    
    missing_files = []
    missing_dirs = []
    
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    for dir in required_dirs:
        if not os.path.exists(dir):
            missing_dirs.append(dir)
    
    if missing_files:
        print(f"❌ Fichiers manquants: {missing_files}")
    else:
        print("✅ Tous les fichiers requis sont présents")
    
    if missing_dirs:
        print(f"❌ Dossiers manquants: {missing_dirs}")
    else:
        print("✅ Tous les dossiers requis sont présents")
    
    return len(missing_files) == 0 and len(missing_dirs) == 0

def check_raw_files():
    """Vérifie la présence des fichiers RAW"""
    print("\n🔍 Vérification des fichiers RAW...")
    
    raw_files = ['raw01.docx', 'raw02.docx', 'raw03.docx']
    found_files = []
    
    for file in raw_files:
        path = os.path.join('raw', file)
        if os.path.exists(path):
            size = os.path.getsize(path)
            found_files.append(f"{file} ({size/1024:.1f} KB)")
        else:
            print(f"❌ Fichier RAW manquant: {file}")
    
    if len(found_files) == len(raw_files):
        print("✅ Tous les fichiers RAW sont présents")
        for file_info in found_files:
            print(f"  📄 {file_info}")
    else:
        print(f"⚠️  {len(found_files)}/{len(raw_files)} fichiers RAW trouvés")
    
    return len(found_files) == len(raw_files)

def check_imports():
    """Vérifie que les imports fonctionnent"""
    print("\n🔍 Vérification des imports...")
    
    try:
        import main
        import parser_mod
        import writer
        import rules
        import utils
        import config
        print("✅ Tous les imports fonctionnent")
        return True
    except ImportError as e:
        print(f"❌ Erreur d'import: {e}")
        return False

def check_dependencies():
    """Vérifie les dépendances Python"""
    print("\n🔍 Vérification des dépendances...")
    
    required_packages = [
        'docx',
        'pandas', 
        'openpyxl'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"❌ Packages manquants: {missing_packages}")
        print("💡 Installez avec: pip install python-docx pandas openpyxl")
        return False
    else:
        print("✅ Toutes les dépendances sont installées")
        return True

def check_output_generation():
    """Vérifie que les fichiers de sortie peuvent être générés"""
    print("\n🔍 Test de génération des fichiers de sortie...")
    
    try:
        # Test avec RAW01
        from main import main
        main("raw/raw01.docx", "out/test_output.docx", "out/test_output.csv", "out/test_output.xlsx")
        
        # Vérifier que les fichiers ont été créés
        output_files = [
            "out/test_output.docx",
            "out/test_output.csv", 
            "out/test_output.xlsx"
        ]
        
        created_files = []
        for file in output_files:
            if os.path.exists(file):
                size = os.path.getsize(file)
                created_files.append(f"{file} ({size} bytes)")
                # Nettoyer les fichiers de test
                os.remove(file)
        
        if len(created_files) == len(output_files):
            print("✅ Génération des fichiers de sortie fonctionnelle")
            for file_info in created_files:
                print(f"  📄 {file_info}")
            return True
        else:
            print("❌ Certains fichiers de sortie n'ont pas été générés")
            return False
            
    except Exception as e:
        print(f"❌ Erreur lors de la génération: {e}")
        return False

def main():
    """Fonction principale de validation"""
    print("🚀 Validation du projet CCC - RAW DATA Processing")
    print("=" * 50)
    
    checks = [
        ("Structure des fichiers", check_file_structure),
        ("Fichiers RAW", check_raw_files),
        ("Imports Python", check_imports),
        ("Dépendances", check_dependencies),
        ("Génération de sortie", check_output_generation)
    ]
    
    results = []
    
    for check_name, check_func in checks:
        try:
            result = check_func()
            results.append((check_name, result))
        except Exception as e:
            print(f"❌ Erreur lors de {check_name}: {e}")
            results.append((check_name, False))
    
    # Résumé
    print("\n" + "=" * 50)
    print("📊 RÉSUMÉ DE LA VALIDATION")
    print("=" * 50)
    
    passed = 0
    total = len(results)
    
    for check_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status} {check_name}")
        if result:
            passed += 1
    
    print(f"\nScore: {passed}/{total} ({passed/total*100:.1f}%)")
    
    if passed == total:
        print("🎉 Toutes les validations sont passées ! Le projet est prêt.")
        return True
    else:
        print("⚠️  Certaines validations ont échoué. Vérifiez les erreurs ci-dessus.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
