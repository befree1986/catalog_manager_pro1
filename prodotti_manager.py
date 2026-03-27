import sqlite3
import os

def aggiungi_prodotto(nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario=0, codice="", tipologia_prodotto="Generico", prezzo3=0, prezzo4=0, qta_min_2=0, qta_min_3=0, qta_min_4=0):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('INSERT INTO prodotti (nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, qta_min_2, qta_min_3, qta_min_4) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
              (nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, qta_min_2, qta_min_3, qta_min_4))
    conn.commit()
    conn.close()

def modifica_prodotto(id, nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario=0, codice="", tipologia_prodotto="Generico", prezzo3=0, prezzo4=0, qta_min_2=0, qta_min_3=0, qta_min_4=0):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('UPDATE prodotti SET nome=?, categoria=?, descrizione=?, prezzo=?, visibile=?, immagine=?, prezzo_secondario=?, codice=?, tipologia_prodotto=?, prezzo3=?, prezzo4=?, qta_min_2=?, qta_min_3=?, qta_min_4=? WHERE id=?',
              (nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, qta_min_2, qta_min_3, qta_min_4, id))
    conn.commit()
    conn.close()

def cancella_prodotto(id):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('DELETE FROM prodotti WHERE id=?', (id,))
    conn.commit()
    conn.close()

def lista_prodotti():
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    # Seleziona esplicitamente le colonne per garantire l'ordine, indipendentemente dalla struttura fisica della tabella
    c.execute('SELECT id, nome, categoria, descrizione, prezzo, visibile, immagine, prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, qta_min_2, qta_min_3, qta_min_4 FROM prodotti')
    prodotti = c.fetchall()
    conn.close()
    return prodotti

def get_existing_skus():
    """Restituisce un set di tutti i codici SKU già presenti nel database."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('SELECT codice FROM prodotti WHERE codice IS NOT NULL AND codice != ""')
    skus = {str(row[0]).strip() for row in c.fetchall()}
    conn.close()
    return skus

def get_tipologie_prodotto():
    """Restituisce una lista unica delle tipologie di prodotto presenti nel database."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('SELECT DISTINCT tipologia_prodotto FROM prodotti WHERE tipologia_prodotto IS NOT NULL AND tipologia_prodotto != ""')
    tipologie = [row[0] for row in c.fetchall()]
    conn.close()
    return sorted(tipologie)

def rinomina_tipologia(vecchia_tipologia, nuova_tipologia):
    """Rinomina una tipologia di prodotto aggiornando tutti i prodotti associati."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('UPDATE prodotti SET tipologia_prodotto=? WHERE tipologia_prodotto=?', (nuova_tipologia, vecchia_tipologia))
    conn.commit()
    conn.close()

def cancella_tipologia(tipologia, elimina_prodotti=False):
    """Cancella una tipologia. Se elimina_prodotti è True, cancella i prodotti, altrimenti li sposta in 'Generico'."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    if elimina_prodotti:
        c.execute('DELETE FROM prodotti WHERE tipologia_prodotto=?', (tipologia,))
    else:
        c.execute('UPDATE prodotti SET tipologia_prodotto="Generico" WHERE tipologia_prodotto=?', (tipologia,))
    conn.commit()
    conn.close()

def get_counts_per_tipologia():
    """Restituisce un dizionario con il conteggio dei prodotti per ogni tipologia."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('SELECT tipologia_prodotto, COUNT(*) FROM prodotti WHERE tipologia_prodotto IS NOT NULL AND tipologia_prodotto != "" GROUP BY tipologia_prodotto')
    counts = dict(c.fetchall())
    conn.close()
    return counts

def aggiorna_tipologia_per_ids(lista_id, nuova_tipologia):
    """Aggiorna la tipologia per una lista specifica di ID prodotto."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    # Creiamo una stringa di placeholder ?,?,? in base alla lunghezza della lista
    placeholders = ','.join('?' for _ in lista_id)
    query = f'UPDATE prodotti SET tipologia_prodotto=? WHERE id IN ({placeholders})'
    # I parametri sono la nuova tipologia + tutti gli id
    params = [nuova_tipologia] + lista_id
    c.execute(query, params)
    conn.commit()
    conn.close()

def pulisci_database():
    """Rimuove prodotti senza nome o senza prezzo."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute("DELETE FROM prodotti WHERE nome IS NULL OR nome = '' OR prezzo IS NULL")
    conn.commit()
    conn.close()

def svuota_tutto():
    """Cancella TUTTI i prodotti dal database."""
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute("DELETE FROM prodotti")
    conn.commit()
    conn.close()

# --- Gestione Listini Multipli ---
def get_listini():
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('SELECT * FROM listini')
    res = c.fetchall()
    conn.close()
    return res

def crea_listino(nome, descrizione=""):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    try:
        c.execute('INSERT INTO listini (nome, descrizione) VALUES (?, ?)', (nome, descrizione))
        conn.commit()
    except sqlite3.IntegrityError:
        pass # Nome già esistente
    conn.close()

def cancella_listino(listino_id):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    # Cancella prima i prezzi associati
    c.execute('DELETE FROM prezzi_listini WHERE listino_id=?', (str(listino_id),))
    c.execute('DELETE FROM listini WHERE id=?', (str(listino_id),))
    conn.commit()
    conn.close()

def get_prezzi_listino(listino_id):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    # Ottiene i prezzi custom, se non esistono per un prodotto, si potrebbe ritornare NULL o gestire in UI
    c.execute('SELECT prodotto_id, prezzo FROM prezzi_listini WHERE listino_id=?', (str(listino_id),))
    res = dict(c.fetchall())
    conn.close()
    return res

def aggiorna_prezzo_listino(listino_id, prodotto_id, prezzo):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    # Verifica se esiste già
    c.execute('SELECT id FROM prezzi_listini WHERE listino_id=? AND prodotto_id=?', (str(listino_id), str(prodotto_id)))
    exists = c.fetchone()
    if exists:
        c.execute('UPDATE prezzi_listini SET prezzo=? WHERE id=?', (float(prezzo), exists[0]))
    else:
        c.execute('INSERT INTO prezzi_listini (listino_id, prodotto_id, prezzo) VALUES (?, ?, ?)', (str(listino_id), str(prodotto_id), float(prezzo)))
    conn.commit()
    conn.close()

# --- Gestione Cataloghi ---
def salva_catalogo_db(nome, path, note=""):
    import datetime
    data = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('INSERT INTO cataloghi_salvati (nome, data_creazione, path_file, note) VALUES (?, ?, ?, ?)', (nome, data, path, note))
    conn.commit()
    conn.close()

def cancella_catalogo_db(id):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    # First, get the file path to delete the file
    c.execute('SELECT path_file FROM cataloghi_salvati WHERE id=?', (id,))
    res = c.fetchone()
    file_path = res[0] if res else None
    
    # Delete the record from the database
    c.execute('DELETE FROM cataloghi_salvati WHERE id=?', (id,))
    conn.commit()
    conn.close()
    
    # Delete the actual file if it exists
    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
        except OSError as e:
            print(f"Errore durante l'eliminazione del file del catalogo {file_path}: {e}")

def get_cataloghi_db():
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('SELECT * FROM cataloghi_salvati ORDER BY id DESC')
    res = c.fetchall()
    conn.close()
    return res

def rinomina_catalogo_db(id, nuovo_nome):
    conn = sqlite3.connect('catalogo.db')
    c = conn.cursor()
    c.execute('UPDATE cataloghi_salvati SET nome=? WHERE id=?', (nuovo_nome, id))
    conn.commit()
    conn.close()
