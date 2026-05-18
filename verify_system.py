import os
import json
from main import APP_VERSION, parse_version
import importlib

def check_integrity():
    print("=== DIAGNOSTICA PRE-LANCIO ===")
    
    # 0. Verifica Moduli Python
    dependencies = ['PyQt5', 'pandas', 'openpyxl', 'requests', 'plyer', 'pyodbc', 'fpdf']
    print("Verifica Dipendenze:")
    for lib in dependencies:
        try:
            importlib.import_module(lib)
            status = "✅ Installato"
        except ImportError:
            status = "❌ MANCANTE"
        print(f"{lib:15}: {status}")
    print("-" * 30)

    # 1. Verifica file necessari
    required_files = ['main.py', 'catalogo.db', 'version.json', 'email_utils.py']
    for f in required_files:
        status = "✅ Presente" if os.path.exists(f) else "❌ MANCANTE"
        print(f"{f:15}: {status}")

    # 2. Controllo Coerenza Versione
    if os.path.exists('version.json'):
        with open('version.json', 'r') as f:
            v_data = json.load(f)
            json_v = v_data.get('version')
            print(f"Versione main.py : {APP_VERSION}")
            print(f"Versione JSON    : {json_v}")
            if json_v != APP_VERSION:
                print("⚠️  ATTENZIONE: La versione nel file JSON non coincide con il codice!")

    # 3. Test logica versioning (parse_version è già in main.py)
    v1 = parse_version(APP_VERSION)
    v2 = parse_version("1.1.0")
    if v1 > v2:
        print(f"Logica versioning: OK ({APP_VERSION} > 1.1.0)")
    else:
        print("Logica versioning: ERRORE")

if __name__ == "__main__":
    check_integrity()