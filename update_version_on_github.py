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

def get_release_notes(version):
    """Legge le note di rilascio da RELEASE_NOTES.txt o restituisce un default."""
    RELEASE_NOTES_PATH = os.path.join(os.path.dirname(__file__), 'RELEASE_NOTES.txt')
    if os.path.exists(RELEASE_NOTES_PATH):
        try:
            with open(RELEASE_NOTES_PATH, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content:
                    return content
        except Exception as e:
            print(f"Avviso: Errore nella lettura di RELEASE_NOTES.txt: {e}")
    return f"Aggiornamento alla versione {version}."

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
        # Usiamo --onedir e --name per far coincidere l'output con quanto atteso da Inno Setup
        # Aggiungiamo --icon per impostare l'icona personalizzata corretta
        subprocess.run([
            sys.executable, 
            "-m", "PyInstaller", 
            "--noconsole", 
            "--noconfirm", 
            "--onedir", 
            "--name", "CatalogoApp", 
            "--icon", "icon.ico",
            "main.py"
        ], check=True)
        print("Eseguibile generato con successo nella cartella 'dist'.")
        
        # Copiamo anche icon.png e icon.ico nella cartella dist/CatalogoApp per assicurarci che siano disponibili a runtime
        dist_dir = os.path.join(os.path.dirname(__file__), 'dist', 'CatalogoApp')
        if os.path.exists(dist_dir):
            shutil.copy('icon.png', dist_dir)
            shutil.copy('icon.ico', dist_dir)
            print("Icone copiate nella cartella di distribuzione dist/CatalogoApp.")
        else:
            print("Avviso: cartella di distribuzione non trovata, icone non copiate.")
            
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
        "notes": get_release_notes(version)
    }
    
    with open(VERSION_JSON_PATH, 'w') as f:
        json.dump(version_data, f, indent=4)
    print(f"Aggiornato {VERSION_JSON_PATH} alla versione {version}.")
    print(f"URL di download impostato su: {download_url}")

def git_commit_and_push(version):
    """Aggiunge, committa e pusha version.json su GitHub."""
    try:
        # Aggiungiamo TUTTE le modifiche (incluso pdf_export.py, import_utils.py, ecc.)
        subprocess.run(['git', 'add', '.'], check=True)
        print(f"File preparati per Git (staging).")

        # Verifica se ci sono cambiamenti effettivi prima di committare per evitare crash
        status = subprocess.run(['git', 'status', '--porcelain'], capture_output=True, text=True)
        if status.stdout.strip():
            commit_message = f"Update version.json and main.py to v{version}"
            subprocess.run(['git', 'commit', '-m', commit_message], check=True)
            print(f"Commit effettuato: '{commit_message}'.")
        else:
            print("Nessuna modifica rilevata nel repository Git, salto il commit.")

        print(f"Invio dati a GitHub (push)...")
        subprocess.run(['git', 'push', 'origin', GITHUB_BRANCH], check=True)
        print("Push completato con successo.")
        
        # Creazione e push del tag git (essenziale per le release)
        tag_name = f"v{version}"
        print(f"Creazione tag locale {tag_name}...")
        subprocess.run(['git', 'tag', '-a', tag_name, '-m', f"Release {tag_name}"], check=False)

        print(f"Invio tag {tag_name} a GitHub...")
        push_tag_result = subprocess.run(['git', 'push', 'origin', tag_name], capture_output=True, text=True, check=False)

        if push_tag_result.returncode != 0:
            if "already exists" in push_tag_result.stderr.lower():
                print(f"Avviso: Il tag {tag_name} esiste già sul repository remoto. Salto il push del tag.")
            else:
                # Rilancia l'eccezione se è un errore diverso
                raise subprocess.CalledProcessError(
                    push_tag_result.returncode, push_tag_result.args, output=push_tag_result.stdout, stderr=push_tag_result.stderr
                )
        else:
            print(f"Tag {tag_name} pushato con successo.")

    except Exception as e:
        print(f"Si è verificato un errore inatteso durante le operazioni Git: {e}")
        sys.exit(1)
def create_github_release(version, installer_path):
    """Crea una GitHub Release e carica l'installer come asset."""
    tag_name = f"v{version}"
    release_name = f"Version {version}"
    release_notes = get_release_notes(version) # La funzione è ora definita

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

        # 2. Carichiamo l'installer su GitHub Release PRIMA di aggiornare la versione online
        create_github_release(new_version, INSTALLER_PATH)
        
        # 3. Solo ora aggiorniamo il codice e il file version.json sul repository
        git_commit_and_push(new_version)

        print("\nScript di automazione completato con successo.")
        print(f"La GitHub Release v{new_version} è stata creata e l'installer è stato caricato.")
        print("Il file version.json è stato aggiornato e pushato.")
    except ValueError as e:
        print(f"Errore: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Si è verificato un errore: {e}")
        sys.exit(1)