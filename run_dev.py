import os
import sys
import subprocess

# Impostiamo l'ambiente per l'anteprima
os.environ["CATALOG_DEV_MODE"] = "1"

# Comando per eseguire l'applicazione
script_path = os.path.join(os.path.dirname(__file__), "main.py")

print("--- AVVIO ANTEPRIMA SVILUPPO ---")
try:
    subprocess.run([sys.executable, script_path])
except KeyboardInterrupt:
    print("\nAnteprima chiusa.")