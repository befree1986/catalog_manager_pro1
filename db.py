import sqlite3
import os

# Definiamo un percorso assoluto per il database basato sulla posizione dello script
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'catalogo.db')

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS prodotti (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT,
        categoria TEXT,
        descrizione TEXT,
        prezzo REAL,
        visibile INTEGER,
        immagine TEXT,
        prezzo_secondario REAL DEFAULT 0
    )''')
    
    # Migrazione per database esistenti: prova ad aggiungere la colonna se manca
    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN prezzo_secondario REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass # La colonna prezzo_secondario esiste già
    
    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN categoria TEXT")
    except sqlite3.OperationalError:
        pass # La colonna categoria esiste già
    
    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN codice TEXT")
    except sqlite3.OperationalError:
        pass # La colonna codice esiste già
    
    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN tipologia_prodotto TEXT DEFAULT 'Generico'")
    except sqlite3.OperationalError:
        pass # La colonna tipologia_prodotto esiste già
    
    # Nuovi campi richiesti (Prezzo 3, 4, Quantità)
    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN prezzo3 REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN prezzo4 REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN quantita INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    # Nuovi campi per le soglie di quantità (Tiered Pricing)
    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN qta_min_2 INTEGER DEFAULT 0")
    except sqlite3.OperationalError: pass

    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN qta_min_3 INTEGER DEFAULT 0")
    except sqlite3.OperationalError: pass

    try:
        c.execute("ALTER TABLE prodotti ADD COLUMN qta_min_4 INTEGER DEFAULT 0")
    except sqlite3.OperationalError: pass

    # Tabelle per gestione Listini Multipli
    c.execute('''CREATE TABLE IF NOT EXISTS listini (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE,
        descrizione TEXT
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS prezzi_listini (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        listino_id INTEGER,
        prodotto_id INTEGER,
        prezzo REAL,
        FOREIGN KEY(listino_id) REFERENCES listini(id),
        FOREIGN KEY(prodotto_id) REFERENCES prodotti(id)
    )''')

    # Tabella per Cataloghi Salvati
    c.execute('''CREATE TABLE IF NOT EXISTS cataloghi_salvati (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT,
        data_creazione TEXT,
        path_file TEXT,
        note TEXT
    )''')
    
    conn.commit()
    conn.close()
