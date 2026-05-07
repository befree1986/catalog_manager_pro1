import os
import re
import json
import subprocess
import sys
import datetime
import shutil

# --- Configurazione ---
# Percorso al file main.py per leggere la versione
MAIN_PY_PATH = os.path.join(os.path.dirname(__file__), 'main.py')
# Percorso al file version.json da aggiornare
VERSION_JSON_PATH = os.path.join(os.path.dirname(__file__), 'version.json')
# Percorso al file dell'installer generato da Inno Setup
# Assumiamo che Inno Setup generi l'installer nella root del progetto, come configurato in setup_script.iss
INSTALLER_PATH = os.path.join(os.path.dirname(__file__), 'CatalogoManager_Setup.exe')
# Dettagli del tuo repository GitHub
GITHUB_REPO_OWNER = "befree1986"
GITHUB_REPO_NAME = "catalog_manager_pro1"
GITHUB_BRANCH = "main" # Il branch su cui vuoi pushare (es. main, master, release)

def sync_iss_version(version):
    iss_path = os.path.join(os.path.dirname(__file__), 'setup_script.iss')
    if os.path.exists(iss_path):
        with open(iss_path, 'r') as f:
            content = f.read()
        # Sostituisce la versione definita nel file .iss
        # Corretto MyAppVersion per coincidere con il file .iss
        new_content = re.sub(r'#define MyAppVersion ".*"', f'#define MyAppVersion "{version}"', content)
        with open(iss_path, 'w') as f:
            f.write(new_content)
        print("Sincronizzata versione anche nel file .iss")

def run_pyinstaller():
    """Esegue PyInstaller per generare l'eseguibile aggiornato."""
    print("Avvio build eseguibile con PyInstaller...")
    try:
        # Usiamo sys.executable per assicurarci di usare lo stesso interprete corrente
        # Aggiungiamo --noconfirm per sovrascrivere la build precedente senza chiedere
        # Usiamo --onefile e --name per far coincidere l'output con quanto atteso da Inno Setup
        subprocess.run([
            sys.executable, 
            "-m", "PyInstaller", 
            "--noconsole", 
            "--noconfirm", 
            "--onefile", 
            "--name", "CatalogoApp", 
            "main.py"
        ], check=True)
        print("Eseguibile generato con successo nella cartella 'dist'.")
    except subprocess.CalledProcessError as e:
        print(f"Errore durante l'esecuzione di PyInstaller: {e}")
        sys.exit(1)

def run_inno_setup():
    """Esegue il compilatore Inno Setup per generare l'installer aggiornato."""
    iscc_path = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    iss_path = os.path.join(os.path.dirname(__file__), 'setup_script.iss')
    if os.path.exists(iscc_path) and os.path.exists(iss_path):
        print(f"Avvio compilazione installer: {iss_path}...")
        subprocess.run([iscc_path, iss_path], check=True)
        print("Installer generato con successo.")
    else:
        print("Avviso: Compilatore Inno Setup (ISCC.exe) non trovato o setup_script.iss mancante. Compila manualmente.")

def get_app_version_from_main_py(main_py_path):
    """Legge APP_VERSION dal file main.py."""
    with open(main_py_path, 'r', encoding='utf-8') as f:
        content = f.read()
        # Regex più robusta: accetta sia ' che " e ignora spazi extra
        match = re.search(r"APP_VERSION\s*=\s*['\"]([^'\"]+)['\"]", content)
        if match:
            return match.group(1)
    raise ValueError(f"APP_VERSION non trovato in {main_py_path}")

def update_version_json(version):
    """Aggiorna o crea version.json con la nuova versione e l'URL di download."""
    # L'URL punta all'asset della release su GitHub.
    # Assicurati che il nome del file .exe sia consistente con quello generato da Inno Setup.
    download_url = f"https://github.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/download/v{version}/CatalogoManager_Setup.exe"
    
    version_data = {
        "version": version,
        "url": download_url,
        "notes": (
            "- Integrazione Microsoft Visual C++ Redistributable nell'installer.\n"
            "- Nuovo sistema di assistenza tecnica diretta via email.\n"
            "- Migliorata condivisione WhatsApp con apertura cartella automatica.\n"
            "- Fix bug grafico nella mappatura dei listini prezzi extra."
        )
    }
    
    with open(VERSION_JSON_PATH, 'w') as f:
        json.dump(version_data, f, indent=4)
    print(f"Aggiornato {VERSION_JSON_PATH} alla versione {version}.")
    print(f"URL di download impostato su: {download_url}")

def git_commit_and_push(version):
    """Aggiunge, committa e pusha version.json su GitHub."""
    try:
        subprocess.run(['git', 'add', VERSION_JSON_PATH], check=True)
        # Aggiungiamo anche main.py per salvare il cambio versione nel repository
        subprocess.run(['git', 'add', MAIN_PY_PATH], check=True)
        print(f"Aggiunto {VERSION_JSON_PATH} all'area di staging di Git.")

        # Verifica se ci sono effettivamente cambiamenti da committare
        # Ritorna 0 se non ci sono differenze, 1 se ce ne sono
        diff_check = subprocess.run(['git', 'diff', '--cached', '--quiet'])
        
        if diff_check.returncode == 0:
            print("Il file version.json è già aggiornato nel repository. Salto il commit.")
        else:
            commit_message = f"Update version.json to v{version}"
            subprocess.run(['git', 'commit', '-m', commit_message], check=True)
            print(f"Commit effettuato con messaggio: '{commit_message}'.")

        subprocess.run(['git', 'push', 'origin', GITHUB_BRANCH], check=True)
        print(f"Push effettuato sul branch {GITHUB_BRANCH} su GitHub.")

    except subprocess.CalledProcessError as e:
        print(f"Errore durante l'operazione Git: {e}")
        print(f"Il comando '{' '.join(e.cmd)}' ha restituito un codice di uscita non zero {e.returncode}.")
        sys.exit(1)
    except Exception as e:
        print(f"Si è verificato un errore inatteso durante le operazioni Git: {e}")
        sys.exit(1)

def create_github_release(version, installer_path):
    """Crea una GitHub Release e carica l'installer come asset."""
    tag_name = f"v{version}"
    release_name = f"Version {version}"
    release_notes = f"Aggiornamento alla versione {version}." # Puoi personalizzare queste note

    # Verifica se gh esiste nel sistema
    gh_path = shutil.which("gh")
    if not gh_path:
        print("Errore: GitHub CLI ('gh') non è nel PATH. Prova a riavviare VS Code.")
        sys.exit(1)

    if not os.path.exists(installer_path):
        print(f"Errore: Installer non trovato al percorso specificato: {installer_path}")
        sys.exit(1)

    try:
        # Verifica se la release esiste già
        check_release = subprocess.run([
            'gh', 'release', 'view', tag_name,
            '--repo', f"{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}"
        ], capture_output=True)

        if check_release.returncode != 0:
            # Crea la release solo se non esiste
            print(f"Creazione della release '{release_name}' con tag '{tag_name}'...")
            subprocess.run([
                'gh', 'release', 'create', tag_name,
                '--repo', f"{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}",
                '--title', release_name,
                '--notes', release_notes,
                '--target', GITHUB_BRANCH
            ], check=True)
            print(f"Release '{release_name}' creata con successo.")
        else:
            print(f"La release '{tag_name}' esiste già. Procedo con l'aggiornamento dell'asset.")

        # Carica l'asset
        print(f"Caricamento dell'asset '{os.path.basename(installer_path)}' sulla release '{tag_name}'...")
        subprocess.run([
            'gh', 'release', 'upload', tag_name, installer_path,
            '--repo', f"{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}",
            '--clobber' # Sovrascrive il file se esiste già
        ], check=True)
        print(f"Asset '{os.path.basename(installer_path)}' caricato con successo.")

    except FileNotFoundError:
        print("Errore: GitHub CLI ('gh') non trovato. Assicurati che sia installato e nel PATH.")
        print("Puoi scaricarlo da: https://cli.github.com/")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        print(f"Errore durante l'operazione GitHub CLI: {e}")
        print(f"Il comando '{' '.join(e.cmd)}' ha restituito un codice di uscita non zero {e.returncode}.")
        print("Assicurati di essere autenticato con 'gh auth login'.")
        sys.exit(1)
    except Exception as e:
        print(f"Si è verificato un errore inatteso durante le operazioni GitHub CLI: {e}")
        sys.exit(1)

if __name__ == "__main__":
    try:
        new_version = get_app_version_from_main_py(MAIN_PY_PATH)
        
        # 1. Aggiorna version.json localmente e pusha su GitHub
        update_version_json(new_version)

        # 1b. Genera l'eseguibile .exe aggiornato
        run_pyinstaller()
        
        # Sincronizza la versione nel file .iss e compila l'installer
        sync_iss_version(new_version)
        run_inno_setup()
        
        git_commit_and_push(new_version)
        
        # 2. Crea la GitHub Release e carica l'installer
        # Assicurati che l'installer sia già stato generato da Inno Setup e si trovi in INSTALLER_PATH
        create_github_release(new_version, INSTALLER_PATH)
        
        print("\nScript di automazione completato con successo.")
        print(f"La GitHub Release v{new_version} è stata creata e l'installer è stato caricato.")
        print("Il file version.json è stato aggiornato e pushato.")
    except ValueError as e:
        print(f"Errore: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Si è verificato un errore: {e}")
        sys.exit(1)