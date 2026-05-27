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
import xml.etree.ElementTree as ET
import logging
from db import DB_PATH

def read_excel_df(file_path):
    """Legge un file Excel, normalizza le colonne e restituisce un DataFrame."""
    if pd is None:
        raise ImportError("Pandas non è installato.")
    # dtype=str assicura che codici come "00344" non vengano letti come numeri e perdano gli zeri
    df = pd.read_excel(file_path, dtype=str)
    df.columns = df.columns.str.lower()
    return df

def read_danea_xml(file_path):
    """Legge un file XML di Danea EasyFatt e restituisce un DataFrame."""
    if pd is None:
        raise ImportError("Pandas non è installato.")
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        prodotti = []
        # Danea XML tipicamente ha i prodotti sotto <Articoli><Articolo>
        for articolo in root.findall('.//Articolo'):
            dati = {}
            for campo in articolo:
                # Normalizziamo i nomi dei campi in minuscolo per compatibilità
                tag_name = campo.tag.lower()
                dati[tag_name] = campo.text
            prodotti.append(dati)
            
        if not prodotti:
            return pd.DataFrame()
            
        df = pd.DataFrame(prodotti)
        return df
    except Exception as e:
        raise Exception(f"Errore durante la lettura dell'XML Danea: {e}")
def normalize_key(val):
    if val is None or str(val).lower() in ('', 'nan', 'none'):
        return ""
    s = str(val).strip().lower()
    # Rimuovi estensione se presente (es. .jpg, .png)
    name, ext = os.path.splitext(s)
    if ext.lower() in ['.jpg', '.jpeg', '.png', '.webp', '.bmp']:
        s = name
    # Rimuovi il suffisso .0 se è una rappresentazione float di un intero (es. "100.0")
    if s.endswith('.0'):
        s = s[:-2]
    # Rimuovi zeri iniziali per facilitare il matching (es "00123" -> "123")
    return s.strip().lstrip('0')

def get_row_value(row, key, default=""):
    """Recupera un valore da una riga Pandas in modo case-insensitive."""
    for col in row.index:
        if col.lower().strip() == key.lower().strip():
            val = row[col]
            return val if pd is not None and pd.notna(val) else default
    return default

def importa_dataframe_nel_db(df, images_folder=None, progress_callback=None, price_list_map=None):
    """Importa un DataFrame nel database, gestendo la logica di conversione e ricerca immagini."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Mappa immagini per ricerca case-insensitive veloce
    image_map = {}
    if images_folder and os.path.exists(images_folder):
        try:
            for root, dirs, files in os.walk(images_folder):
                for f in files:
                    key = normalize_key(f)
                    if key:
                        # In caso di duplicati in sottocartelle diverse, l'ultimo trovato vince
                        image_map[key] = os.path.join(root, f)
            logging.info(f"Mappa immagini importazione: trovati {len(image_map)} file (scansione ricorsiva).")
        except OSError as e: 
            logging.error(f"Errore durante la scansione della cartella immagini '{images_folder}': {e}")
            pass

    try:
        total_rows = len(df)
        for i, (_, row) in enumerate(df.iterrows()):
            # Estrazione dati con valori di default
            nome = str(get_row_value(row, 'nome') or 'Nuovo Prodotto')
            descrizione = str(get_row_value(row, 'descrizione') or '')
            categoria = str(get_row_value(row, 'categoria') or 'Generale')
            tipologia_prodotto = str(get_row_value(row, 'tipologia_prodotto') or 'Generico')
            
            # Funzione di utilità interna per pulizia prezzi
            def clean_price(val):
                if pd.isna(val) or val == '': return 0.0
                try:
                    return float(str(val).replace(',', '.'))
                except:
                    return 0.0

            # Gestione Codice/SKU (i dati dalla tabella sono stringhe)
            codice = str(get_row_value(row, 'codice') or get_row_value(row, 'sku') or '').strip()
            if codice.lower() in ('nan', 'none'): codice = ''
            
            codice_normalizzato = normalize_key(codice)

            # Conversione sicura dei prezzi
            prezzo = clean_price(get_row_value(row, 'prezzo', 0))
            prezzo_secondario = clean_price(get_row_value(row, 'prezzo_secondario', 0))
            prezzo3 = clean_price(get_row_value(row, 'prezzo3', 0))
            prezzo4 = clean_price(get_row_value(row, 'prezzo4', 0))
            
            qta_min_2 = int(float(str(get_row_value(row, 'qta_min_2', 0)).replace(',', '.')))
            qta_min_3 = int(float(str(get_row_value(row, 'qta_min_3', 0)).replace(',', '.')))
            qta_min_4 = int(float(str(get_row_value(row, 'qta_min_4', 0)).replace(',', '.')))

            visibile = 1
            immagine = str(get_row_value(row, 'immagine') or '').strip()
            if immagine.lower() in ('nan', 'none'): immagine = ''

            # LOGICA IMMAGINI DA CARTELLA:
            # 1. Se l'utente ha selezionato una cartella immagini, proviamo a risolverla tramite il valore mappato in 'immagine'
            # 2. Se non lo trova o non è mappata, proviamo tramite il 'codice'
            if images_folder:
                logging.debug(f"Processing product '{nome}' for image resolution. Initial image field: '{immagine}', SKU/Code: '{codice}'")
                resolved = False
                
                # Prova prima con il valore presente nel campo immagine (es. nome file o codice scritto nella colonna)
                if immagine:
                    img_key = normalize_key(immagine)
                    if img_key in image_map:
                        logging.debug(f"Image resolved by 'immagine' field: '{img_key}' -> '{image_map[img_key]}'")
                        immagine = image_map[img_key]
                        resolved = True
                
                # Se non risolto, prova con il codice a barre / SKU
                if not resolved and codice_normalizzato:
                    # codice_normalizzato è già normalizzato, usalo direttamente come chiave
                    if codice_normalizzato in image_map:
                        logging.debug(f"Image resolved by 'codice' field: '{codice_normalizzato}' -> '{image_map[codice_normalizzato]}'")
                        immagine = image_map[codice_normalizzato]
                        resolved = True

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
                                # Rilevamento intelligente del suffisso (es. % per IVA)
                                default_suffisso = "%" if "iva" in listino_nome.lower() else "€"
                                try:
                                    c.execute('INSERT OR IGNORE INTO listini (nome, descrizione, suffisso) VALUES (?, ?, ?)', (listino_nome, "Importato", default_suffisso))
                                except sqlite3.OperationalError:
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

def sincronizza_immagini_database(tipologia_filter, images_folder, progress_callback=None):
    """Cerca e aggiorna i percorsi delle immagini per i prodotti esistenti nel database."""
    if not images_folder or not os.path.exists(images_folder):
        return 0
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # 1. Mappa immagini presenti nella cartella
    image_map = {}
    try:
        for root, dirs, files in os.walk(images_folder):
            for f in files:
                key = normalize_key(f)
                if key:
                    image_map[key] = os.path.join(root, f)
        logging.info(f"Sincronizzazione: trovati {len(image_map)} file nella cartella selezionata (ricerca ricorsiva).")
    except OSError as e:
        logging.error(f"Errore scansione cartella sync: {e}")
        conn.close()
        return 0

    # 2. Recupera prodotti filtrati
    logging.debug(f"Sincronizzazione immagini per tipologia: '{tipologia_filter}'")
    query = "SELECT id, nome, codice, immagine FROM prodotti"
    params = []
    if tipologia_filter and tipologia_filter not in ("Tutte", "Tutti i Gruppi"):
        query += " WHERE tipologia_prodotto = ?"
        params.append(tipologia_filter)
    
    c.execute(query, params)
    prodotti = c.fetchall()
    logging.debug(f"Prodotti trovati nel DB da controllare: {len(prodotti)}")
    
    updated_count = 0
    total = len(prodotti)
    
    for i, (p_id, nome, codice, img_attuale) in enumerate(prodotti):
        resolved_path = None
        
        # Strategia di matching:
        # 1. Prova con il codice/SKU esatto
        if codice:
            code_key = normalize_key(codice)
            if code_key in image_map:
                resolved_path = image_map[code_key]
            # 1b. Prova rimuovendo gli zeri iniziali (es. SKU 00344 -> file 344.jpg)
            elif codice.lstrip('0') in image_map:
                resolved_path = image_map[codice.lstrip('0')]
        
        # 2. Prova con il Nome (fallback)
        if not resolved_path and nome:
            name_key = normalize_key(nome)
            if name_key in image_map:
                resolved_path = image_map[name_key]

        # 3. Prova con il valore salvato nel campo immagine
        if not resolved_path and img_attuale:
            img_key = normalize_key(img_attuale)
            if img_key in image_map:
                resolved_path = image_map[img_key]
        
        if resolved_path and resolved_path != img_attuale:
            logging.debug(f"Match trovato: '{nome}' -> {resolved_path}")
            c.execute("UPDATE prodotti SET immagine = ? WHERE id = ?", (resolved_path, p_id))
            updated_count += 1
        elif not resolved_path:
            logging.debug(f"Nessun match per '{nome}' (SKU: {codice}). Chiavi cercate: {normalize_key(codice)}, {normalize_key(nome)}")
            
        if progress_callback:
            progress_callback(i + 1, total)
            
    conn.commit()
    conn.close()
    return updated_count

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
