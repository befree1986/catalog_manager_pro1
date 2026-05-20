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

def read_excel_df(file_path):
    """Legge un file Excel, normalizza le colonne e restituisce un DataFrame."""
    if pd is None:
        raise ImportError("Pandas non è installato.")
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.lower()
    return df

def read_danea_xml(file_path):
    """
    Legge un file XML di Danea EasyFatt e restituisce un DataFrame.
    NOTA: La logica di parsing effettiva per il formato XML di Danea deve essere implementata qui.
    Attualmente restituisce un DataFrame vuoto.
    """
    if pd is None:
        raise ImportError("Pandas non è installato.")
    return pd.DataFrame() # Implementare qui la logica di parsing XML

def importa_dataframe_nel_db(df, images_folder=None, progress_callback=None, price_list_map=None):
    """Importa un DataFrame nel database, gestendo la logica di conversione e ricerca immagini."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    
    # Mappa immagini per ricerca case-insensitive veloce
    image_map = {}
    if images_folder and os.path.exists(images_folder):
        try:
            for f in os.listdir(images_folder):
                name, ext = os.path.splitext(f)
                if ext.lower() in ['.jpg', '.jpeg', '.png', '.webp', '.bmp']:
                    image_map[name.lower()] = os.path.join(images_folder, f)
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

            # Conversione sicura dei prezzi
            prezzo = clean_price(row.get('prezzo', 0))
            prezzo_secondario = clean_price(row.get('prezzo_secondario', 0))

            visibile = 1
            immagine = str(row.get('immagine', ''))
            if immagine.lower() == 'nan': immagine = ''

            # LOGICA IMMAGINI DA CARTELLA: Se abbiamo SKU e Cartella, cerchiamo il file
            # Migliorata: ricerca case-insensitive tramite mappa
            if images_folder and codice:
                codice_lower = codice.lower().strip()
                if codice_lower in image_map:
                    immagine = image_map[codice_lower]

            c.execute('INSERT INTO prodotti (nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
                         (nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto))
            
            # Ottieni ID del prodotto appena inserito
            new_prod_id = c.lastrowid

            # --- GESTIONE LISTINI EXTRA ---
            if price_list_map and new_prod_id:
                for col_name, listino_nome in price_list_map.items():
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
