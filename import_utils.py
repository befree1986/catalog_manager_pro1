import sqlite3
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import pyodbc
except ImportError:
    pyodbc = None
import os
from db import DB_PATH

def read_excel_df(file_path):
    """Legge un file Excel, normalizza le colonne e restituisce un DataFrame."""
    if pd is None:
        raise ImportError("Pandas non è installato.")
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.lower()
    return df

def read_danea_xml(file_path):
    """Legge un file XML di Danea EasyFatt e restituisce un DataFrame."""
    if pd is None:
        raise ImportError("Pandas non è installato.")
    try:
        # Placeholder: per ora restituisce un DF vuoto per evitare crash
        return pd.DataFrame()
    except Exception as e:
        raise Exception(f"Errore durante la lettura dell'XML Danea: {e}")
def normalize_key(val):
    if not val:
        return ""
    val_str = str(val).strip()
    # Rimuovi il suffisso .0 se è una rappresentazione float di un intero
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    # Rimuovi estensione se presente (es. .jpg, .png)
    name, ext = os.path.splitext(val_str)
    if ext.lower() in ['.jpg', '.jpeg', '.png', '.webp', '.bmp']:
        val_str = name
    return val_str.lower().strip()

def importa_dataframe_nel_db(df, images_folder=None, progress_callback=None, price_list_map=None):
    """Importa un DataFrame nel database, gestendo la logica di conversione e ricerca immagini."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Mappa immagini per ricerca case-insensitive veloce
    image_map = {}
    if images_folder and os.path.exists(images_folder):
        try:
            for f in os.listdir(images_folder):
                name, ext = os.path.splitext(f)
                if ext.lower() in ['.jpg', '.jpeg', '.png', '.webp', '.bmp']:
                    # Chiave normalizzata: es. "1001"
                    image_map[name.lower().strip()] = os.path.join(images_folder, f)
        except OSError: pass

    try:
        total_rows = len(df)
        for i, (_, row) in enumerate(df.iterrows()):
            # Estrazione dati con valori di default
            nome = row.get('nome', 'Nuovo Prodotto')
            descrizione = row.get('descrizione', '')
            categoria = row.get('categoria', 'Generale')
            tipologia_prodotto = row.get('tipologia_prodotto', 'Generico')
            
            # Funzione di utilità interna per pulizia prezzi
            def clean_price(val):
                if pd.isna(val) or val == '': return 0.0
                try:
                    return float(str(val).replace(',', '.'))
                except:
                    return 0.0

            # Gestione Codice/SKU (i dati dalla tabella sono stringhe)
            codice = str(row.get('codice', '') or row.get('sku', '')).strip()
            if codice.lower() == 'nan': codice = ''
            
            # Normalizziamo il codice per rimuovere eventuali .0 da Excel
            codice_normalizzato = codice
            if codice_normalizzato.endswith('.0'):
                codice_normalizzato = codice_normalizzato[:-2]

            # Conversione sicura dei prezzi
            prezzo = clean_price(row.get('prezzo', 0))
            prezzo_secondario = clean_price(row.get('prezzo_secondario', 0))
            prezzo3 = clean_price(row.get('prezzo3', 0))
            prezzo4 = clean_price(row.get('prezzo4', 0))
            
            qta_min_2 = int(float(str(row.get('qta_min_2', 0) or 0).replace(',', '.')))
            qta_min_3 = int(float(str(row.get('qta_min_3', 0) or 0).replace(',', '.')))
            qta_min_4 = int(float(str(row.get('qta_min_4', 0) or 0).replace(',', '.')))

            visibile = 1
            immagine = str(row.get('immagine', '')).strip()
            if immagine.lower() == 'nan': immagine = ''

            # LOGICA IMMAGINI DA CARTELLA:
            # 1. Se l'utente ha selezionato una cartella immagini, proviamo a risolverla tramite il valore mappato in 'immagine'
            # 2. Se non lo trova o non è mappata, proviamo tramite il 'codice'
            if images_folder:
                resolved = False
                
                # Prova prima con il valore presente nel campo immagine (es. nome file o codice scritto nella colonna)
                if immagine:
                    img_key = normalize_key(immagine)
                    if img_key in image_map:
                        immagine = image_map[img_key]
                        resolved = True
                
                # Se non risolto, prova con il codice a barre / SKU
                if not resolved and codice_normalizzato:
                    code_key = normalize_key(codice_normalizzato)
                    if code_key in image_map:
                        immagine = image_map[code_key]

            c.execute('''INSERT INTO prodotti (nome, categoria, descrizione, prezzo, visibile, immagine, 
                         prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, 
                         qta_min_2, qta_min_3, qta_min_4) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                         (nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, qta_min_2, qta_min_3, qta_min_4))
            
            # Ottieni ID del prodotto appena inserito
            new_prod_id = c.lastrowid

            # --- GESTIONE LISTINI EXTRA / CAMPI EXTRA ---
            # Definiamo i campi standard da ignorare per i listini extra
            standard_fields = {
                'id', 'nome', 'categoria', 'descrizione', 'prezzo', 'visibile', 
                'immagine', 'prezzo_secondario', 'codice', 'tipologia_prodotto', 
                'prezzo3', 'prezzo4', 'qta_min_2', 'qta_min_3', 'qta_min_4', 'quantita'
            }
            
            # Uniamo i listini definiti nella mappa e qualsiasi colonna extra presente nel DataFrame
            all_extra_columns = {} # {nome_colonna_nel_df: nome_listino}
            
            # Se abbiamo un mapping esplicito dei listini, usiamo quello
            if price_list_map:
                for col_name, listino_nome in price_list_map.items():
                    if col_name in row.index:
                        all_extra_columns[col_name] = listino_nome
            
            # Aggiungiamo qualsiasi altra colonna nel DataFrame che non sia un campo standard
            for col_name in row.index:
                col_lower = col_name.lower().strip()
                if col_lower not in standard_fields and col_name not in all_extra_columns:
                    # Il nome della colonna stessa diventa il nome del listino (es. "IVA", "Accise")
                    all_extra_columns[col_name] = col_name

            if new_prod_id:
                for col_name, listino_nome in all_extra_columns.items():
                    valore = row.get(col_name, 0)
                    try:
                        prezzo_listino = float(str(valore).replace(',', '.'))
                        if prezzo_listino > 0:
                            # Trova o crea ID listino
                            c.execute('SELECT id FROM listini WHERE nome = ?', (listino_nome,))
                            res = c.fetchone()
                            if res:
                                listino_id = res[0]
                            else:
                                c.execute('INSERT OR IGNORE INTO listini (nome, descrizione) VALUES (?, ?)', (listino_nome, "Importato"))
                                c.execute('SELECT id FROM listini WHERE nome = ?', (listino_nome,))
                                res_new = c.fetchone()
                                listino_id = res_new[0] if res_new else None
                            
                            if listino_id:
                                c.execute('INSERT OR REPLACE INTO prezzi_listini (listino_id, prodotto_id, prezzo) VALUES (?, ?, ?)', (listino_id, new_prod_id, prezzo_listino))
                    except (ValueError, TypeError):
                        pass # Valore non valido, ignora

            if progress_callback:
                progress_callback(i + 1, total_rows)
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e # Rilancia l'eccezione per essere gestita dall'UI
    finally:
        conn.close()

def get_access_tables(file_path):
    """Restituisce una lista delle tabelle presenti nel file Access."""
    if pyodbc is None:
        raise ImportError("La libreria 'pyodbc' non è installata. Impossibile connettersi ad Access.")
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file_path
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        tables = [row.table_name for row in cursor.tables(tableType='TABLE')]
        conn.close()
        return tables
    except pyodbc.Error as e:
        raise Exception(f"Impossibile connettersi al database Access.\nAssicurati di avere installato 'Microsoft Access Database Engine'.\nDettagli errore: {e}")

def read_access_table(file_path, table_name):
    """Legge una specifica tabella da un file Access e restituisce un DataFrame."""
    if pyodbc is None:
        raise ImportError("La libreria 'pyodbc' non è installata.")
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file_path
    conn = pyodbc.connect(conn_str)
    try:
        cursor = conn.cursor()
        cursor.execute(f'SELECT * FROM [{table_name}]')
        rows = cursor.fetchall()
        
        if not rows:
            # Restituisce DataFrame vuoto ma con le colonne corrette se la tabella è vuota
            if cursor.description:
                cols = [c[0] for c in cursor.description]
                return pd.DataFrame(columns=cols)
            return pd.DataFrame()
            
        columns = [column[0] for column in cursor.description]
        data = [tuple(row) for row in rows]
        
        df = pd.DataFrame(data, columns=columns)
        df.columns = df.columns.str.lower()
        return df
    finally:
        conn.close()
