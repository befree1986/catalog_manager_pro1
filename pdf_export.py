try:
    from fpdf import FPDF
except ImportError:
    FPDF = object
import sqlite3
import os
from db import DB_PATH

class CatalogoPDF(FPDF if FPDF is not object else object):
    def __init__(self, config=None):
        if FPDF is object:
            raise ImportError("La libreria 'fpdf' non è installata.")
        super().__init__()
        self.config = config or {}
        self.primary_color = self._hex_to_rgb(self.config.get('color', '#2c3e50'))
        self.on_cover = False

    def _hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def header(self):
        if self.on_cover:
            return
        r, g, b = self.primary_color
        self.set_fill_color(r, g, b)
        self.rect(0, 0, 210, 30, 'F')
        header_font_size = self.config.get('style', {}).get('header_font_size', 20)
        self.set_font('Arial', 'B', header_font_size)
        self.set_text_color(255, 255, 255)
        self.cell(0, 15, self.config.get('title', 'Catalogo Prodotti'), 0, 1, 'C')
        self.ln(10)

    def add_cover_page(self, image_path):
        self.on_cover = True
        self.add_page()
        if os.path.exists(image_path):
            try:
                self.image(image_path, x=0, y=0, w=210) # Larghezza A4
            except Exception as e:
                print(f"Errore nel caricare l'immagine di copertina: {e}")
        self.on_cover = False

    def add_full_page_image(self, image_path):
        self.add_page()
        if os.path.exists(image_path):
            self.image(image_path, x=0, y=0, w=210, h=297)
    
    def add_separator_page(self, title, subtitle=""):
        self.add_page()
        self.set_fill_color(*self.primary_color)
        self.rect(0, 100, 210, 60, 'F')
        self.set_y(115)
        self.set_font('Arial', 'B', 28)
        self.set_text_color(255, 255, 255)
        self.cell(0, 15, title, 0, 1, 'C')
        if subtitle:
            self.set_font('Arial', 'I', 16)
            self.cell(0, 10, subtitle, 0, 1, 'C')

    def footer(self):
        if self.on_cover:
            return
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128)
        company = self.config.get('company', '')
        self.cell(0, 10, f'{company} - Pagina {self.page_no()}', 0, 0, 'C')

    def add_index_page(self, index_data):
        pass # Placeholder, implementato esternamente o in sottoclasse se necessario

    def scheda_prodotto(self, nome, categoria, descrizione, prezzo, immagine, codice, tiers):
        if self.get_y() > 240:
            self.add_page()

        start_y = self.get_y()
        
        # Disegno bordo
        self.set_draw_color(200, 200, 200)
        self.rect(10, start_y, 190, 40)
        
        # Immagine
        img_path = immagine.strip() if immagine else ""
        if img_path and os.path.exists(img_path):
            try:
                # Mantieni aspect ratio nel box 35x35
                self.image(img_path, x=12, y=start_y+2, w=35, h=35)
            except:
                pass
        
        # Testi
        self.set_xy(50, start_y + 2)
        style = self.config.get('style', {})
        
        self.set_font('Arial', 'B', style.get('product_title_size', 14))
        r, g, b = self.primary_color
        self.set_text_color(r, g, b)
        self.multi_cell(110, 6, nome, 0, 'L')
        
        # Codice SKU
        if codice:
            self.set_xy(50, self.get_y())
            self.set_font('Arial', '', style.get('sku_size', 9))
            self.set_text_color(80, 80, 80)
            self.cell(0, 4, f"SKU: {codice}", 0, 1)
        
        self.set_xy(50, self.get_y())
        self.set_font('Arial', 'I', style.get('cat_size', 10))
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, categoria if categoria else "Nessuna Categoria", 0, 1)
        
        self.set_xy(50, self.get_y())
        # Descrizione
        self.set_font('Arial', '', style.get('desc_size', 12))
        self.set_text_color(0, 0, 0)
        # Tronca descrizione lunga
        desc_short = (descrizione[:75] + '...') if len(descrizione) > 75 else descrizione
        self.multi_cell(100, 6, desc_short, 0)
        
        # Prezzo (se abilitato nelle impostazioni)
        if self.config.get('show_prices', True):
            if tiers and len(tiers) > 1:
                # Disegna tabella prezzi
                table_x = 155
                table_y = start_y + 5
                self.set_xy(table_x, table_y)
                self.set_font('Arial', 'B', 9)
                self.set_text_color(0,0,0)
                self.cell(15, 6, "Q.ta", 1, 0, 'C')
                self.cell(20, 6, "Prezzo", 1, 1, 'C')
                
                self.set_font('Arial', '', 9)
                for qty, prc in tiers:
                    self.set_x(table_x)
                    self.cell(15, 6, f"{qty}+", 1, 0, 'C')
                    self.cell(20, 6, f"EUR {prc:.2f}", 1, 1, 'R')
            else:
                # Prezzo singolo
                self.set_xy(160, start_y + 10)
                self.set_font('Arial', 'B', style.get('price_size', 16))
                self.set_text_color(39, 174, 96) # Verde
                self.cell(30, 10, f"EUR {prezzo:.2f}", 0, 1, 'R')

        self.set_y(start_y + 45) # Spazio per il prossimo

    def scheda_prodotto_grid(self, nome, categoria, descrizione, prezzo, immagine, x, y, w, h, tiers):
        self.set_draw_color(220, 220, 220)
        self.rect(x, y, w, h)

        # Immagine centrata (occupa circa 55% altezza)
        img_h = h * 0.55
        img_w_max = w - 10
        img_path = immagine.strip() if immagine else ""
        
        if img_path and os.path.exists(img_path):
            try:
                # Adattamento automatico con mantenimento proporzioni
                self.image(img_path, x=x+5, y=y+5, w=img_w_max, h=img_h-5, keep_aspect_ratio=True)
            except:
                pass        

        text_y = y + img_h + 2
        style = self.config.get('style', {})
        
        # Titolo Prodotto in Griglia
        grid_title_size = style.get('grid_title_size', 10)
        self.set_font('Arial', 'B', grid_title_size)
        
        grid_title_color = style.get('grid_title_color', '#2c3e50')
        r, g, b = self._hex_to_rgb(grid_title_color)
        self.set_text_color(r, g, b)

        # Adattamento automatico del font se il nome è molto lungo
        nome_pulito = nome.replace('\n', ' ').strip()
        if len(nome_pulito) > 30:
            self.set_font('Arial', 'B', grid_title_size - 1)
            
        self.set_xy(x + 2, text_y)
        # multi_cell permette l'andata a capo automatica nel box
        self.multi_cell(w - 4, 4.5, nome_pulito, 0, 'C')
        
        if self.config.get('show_prices', True):
            price_y = self.get_y() + 1
            if tiers and len(tiers) > 1:
                # Tabella compatta per griglia
                row_h = 4
                self.set_font('Arial', '', 7)
                self.set_text_color(0,0,0)
                
                # Header
                self.set_xy(x + 5, price_y)
                self.cell(w-10, row_h, "Prezzi per quantità:", 0, 1, 'L')
                
                for qty, prc in tiers:
                    self.set_x(x + 5)
                    self.cell(15, row_h, f"{qty}+ pz", 0, 0, 'L')
                    self.cell(20, row_h, f"EUR {prc:.2f}", 0, 1, 'L')
            else:
                self.set_xy(x + 2, price_y + 2)
                self.set_font('Arial', 'B', style.get('grid_price_size', 11))

                grid_price_color = style.get('grid_price_color', '#27ae60')
                r, g, b = self._hex_to_rgb(grid_price_color)
                self.set_text_color(r, g, b)

                self.cell(w-4, 6, f"EUR {prezzo:.2f}", 0, 1, 'C')
                self.set_text_color(0, 0, 0) # Reset

def get_catalog_data(config):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Filtro Categoria
    category_filter = config.get('category_filter', 'Tutte le categorie')
    tipologia_filter = config.get('tipologia_filter', 'Tutte')
    query = 'SELECT nome, categoria, descrizione, prezzo, immagine, prezzo_secondario, codice, tipologia_prodotto, prezzo3, prezzo4, qta_min_2, qta_min_3, qta_min_4 FROM prodotti WHERE visibile=1'
    params = []
    
    if category_filter and category_filter != "Tutte le categorie":
        query += ' AND categoria = ?'
        params.append(category_filter)
    
    if tipologia_filter and tipologia_filter != "Tutte":
        query += ' AND tipologia_prodotto = ?'
        params.append(tipologia_filter)
    
    try:
        c.execute(query, tuple(params))
    except:
        # Fallback se la colonna codice non dovesse esistere (non dovrebbe succedere se il db è aggiornato)
        # Nota: il filtro per categoria nel fallback è omesso per brevità in caso di errore schema, ma idealmente andrebbe replicato
        # Aggiunto anche tipologia_prodotto per consistenza
        c.execute('SELECT nome, categoria, descrizione, prezzo, immagine, prezzo_secondario, "" as codice, "" as tipologia_prodotto, 0, 0, 0, 0, 0 FROM prodotti WHERE visibile=1')
        
    prodotti = c.fetchall()
    conn.close()
    return prodotti

def generate_pdf_content(pdf, prodotti, config, layout, dry_run=False, page_map=None):
    cover_image = config.get('cover_image')
    if cover_image and os.path.exists(cover_image):
        pdf.add_cover_page(cover_image)

    # Pagina Indice (solo se non è dry_run e abbiamo una mappa)
    # Aggiunto controllo config 'include_index'
    if not dry_run and page_map and config.get('include_index', True):
        pdf.add_page()
        pdf.set_font('Arial', 'B', 24)
        pdf.cell(0, 20, "Indice Prodotti", 0, 1, 'C')
        pdf.ln(10)
        
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, "Categorie:", 0, 1)
        pdf.set_font('Arial', '', 12)
        
        index_items = []
        
        # Se usiamo la struttura personalizzata, seguiamola per l'indice
        structure = config.get('catalog_structure', [])
        if config.get('sort_mode') == 'Struttura Personalizzata' and structure:
            # Scansiona struttura
            processed_cats = set()
            for item in structure:
                if item['type'] == 'category':
                    cat = item['value']
                    if cat in page_map:
                        index_items.append((cat, page_map[cat]))
                        processed_cats.add(cat)
            
            # Aggiungi categorie non in struttura
            remaining = sorted([c for c in page_map if c not in processed_cats])
            for c in remaining:
                index_items.append((c, page_map[c]))
        else:
             # Ordine alfabetico standard per indice
             sorted_cats = sorted(page_map.keys())
             for cat in sorted_cats:
                 index_items.append((cat, page_map[cat]))

        for cat, page in index_items:
            link = pdf.add_link()
            pdf.set_link(link, y=0.0, page=page)
            pdf.cell(160, 8, cat, 0, 0)
            pdf.cell(20, 8, str(page), 0, 1, 'R', link=link)

    pdf.add_page()
    
    # Mappa per tracciare le categorie {NomeCategoria: NumeroPagina}
    current_category_map = {}
    last_category = None
    
    if "Griglia" in layout:
        # Configurazione Colonne
        if "4" in layout:
            num_cols = 4
        elif "3" in layout:
            num_cols = 3
        else:
            num_cols = 2
        margin = 10
        page_width = 210 - (margin * 2)
        col_gap = 5
        # Calcolo larghezza colonna: (LarghezzaPagina - (Gap * (NumCol - 1))) / NumCol
        col_w = (page_width - (col_gap * (num_cols - 1))) / num_cols
        box_h = 75
        
        current_col = 0
        row_y = pdf.get_y()
        
        for row_data in prodotti:
            # Controllo nuova pagina solo all'inizio di ogni riga
            if current_col == 0:
                if row_y > 280 - box_h:
                    pdf.add_page()
                    row_y = pdf.get_y()

            # Unpack sicuro
            nome, cat, descrizione, prezzo, immagine, p2, codice, tipologia, p3, p4, q2, q3, q4 = row_data
            
            # Tracciamento Categoria per Indice
            if cat != last_category:
                if cat not in current_category_map:
                    current_category_map[cat] = pdf.page_no()
                last_category = cat

            # Costruisci Tiers
            tiers = [(1, prezzo)]
            if p2 > 0 and q2 > 0: tiers.append((q2, p2))
            if p3 > 0 and q3 > 0: tiers.append((q3, p3))
            if p4 > 0 and q4 > 0: tiers.append((q4, p4))

            x = margin + (current_col * (col_w + col_gap))
            pdf.scheda_prodotto_grid(nome, cat, descrizione, prezzo, immagine, x, row_y, col_w, box_h, tiers)
            
            current_col += 1
            if current_col >= num_cols:
                current_col = 0
                row_y += box_h + 5 # Altezza box + margine verticale
    else:
        for row_data in prodotti:
            nome, cat, descrizione, prezzo, immagine, p2, codice, tipologia, p3, p4, q2, q3, q4 = row_data
            
            # Tracciamento Categoria per Indice
            if cat != last_category:
                if cat not in current_category_map:
                    current_category_map[cat] = pdf.page_no()
                last_category = cat

            tiers = [(1, prezzo)]
            if p2 > 0 and q2 > 0: tiers.append((q2, p2))
            if p3 > 0 and q3 > 0: tiers.append((q3, p3))
            if p4 > 0 and q4 > 0: tiers.append((q4, p4))

            pdf.scheda_prodotto(nome, cat, descrizione, prezzo, immagine, codice, tiers)
            
    return current_category_map

def esporta_catalogo_pdf(filename='catalogo.pdf', config=None):
    prodotti = get_catalog_data(config)

    # --- Gestione Ordinamento ---
    sort_mode = config.get('sort_mode', 'Categoria (Personalizzato)')
    
    if sort_mode == 'Categoria (Personalizzato)':
        # Ordina i prodotti in base all'ordine personalizzato delle categorie
        category_order = config.get('category_order', [])
        if category_order:
            order_map = {cat: i for i, cat in enumerate(category_order)}
            prodotti.sort(key=lambda p: order_map.get(p[2], float('inf')))
        else:
            # Fallback alfabetico se non c'è ordine custom
            prodotti.sort(key=lambda p: (p[2] or "", p[1] or ""))
            
    elif sort_mode == 'Categoria (Alfabetico)':
        prodotti.sort(key=lambda p: (p[2] or "", p[1] or ""))
        
    elif sort_mode == 'Nome':
        prodotti.sort(key=lambda p: (p[1] or "").lower())
        
    elif sort_mode == 'Codice (SKU)':
        prodotti.sort(key=lambda p: str(p[8] or "").lower())

    layout = config.get('layout', 'Lista') if config else 'Lista'
    include_index = config.get('include_index', True)
    
    page_map_final = None
    
    # Se l'indice è disabilitato, saltiamo il passaggio di calcolo pagine (Dry Run)
    if include_index:
        # PASSAGGIO 1: Dry Run per calcolare le pagine delle categorie
        # Creiamo un PDF temporaneo
        temp_pdf = CatalogoPDF(config=config)
        # Disabilitiamo la compressione per velocità
        temp_pdf.set_compression(False)
        
        # Generiamo contenuto (senza indice) per capire dove finiscono le categorie
        cat_map = generate_pdf_content(temp_pdf, prodotti, config, layout, dry_run=True)
        
        # Calcolo offset pagine dovuto all'indice
        # Stimiamo le pagine dell'indice: 1 pagina ogni 40 categorie circa
        if cat_map:
            num_cats = len(cat_map)
            index_pages = 1 + (num_cats // 40)
            # Shiftiamo i numeri di pagina nella mappa
            page_map_final = {k: v + index_pages for k, v in cat_map.items()}
        else:
            page_map_final = {}
            
    else:
        page_map_final = None
    
    if FPDF is object:
        print("Errore: Libreria 'fpdf' non trovata.")
        return
    # PASSAGGIO 2: Generazione Reale
    pdf = CatalogoPDF(config=config)
    generate_pdf_content(pdf, prodotti, config, layout, dry_run=False, page_map=page_map_final)
    
    pdf.output(filename)
