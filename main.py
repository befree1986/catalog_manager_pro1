import sys
import os
import shutil
import re
import datetime
import json
import logging

APP_VERSION = "1.5.2" # Fix filtro immagini e crash generazione catalogo

# --- SETUP LOGGING IMMEDIATO ---
# Deve essere fatto PRIMA di caricare i moduli locali per catturare errori di importazione
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(
    filename=os.path.join(base_dir, 'debug_log.txt'),
    level=logging.DEBUG, # Temporaneamente impostato su DEBUG per una diagnostica dettagliata
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

try:
    import pandas as pd # Requires 'pip install pandas openpyxl'
except ImportError:
    pd = None
from PyQt5.QtWidgets import (QApplication, QMainWindow, QAction, QFileDialog, QMessageBox, QWidget, QVBoxLayout, QPushButton, QLabel, QHBoxLayout, QComboBox, QScrollArea, QInputDialog, QFrame, QSizePolicy, QLineEdit, QStackedWidget, QGridLayout, QTableWidget, QTableWidgetItem, QHeaderView, QFormLayout, QCheckBox, QColorDialog, QDialog, QDialogButtonBox, QProgressDialog, QGraphicsDropShadowEffect, QListWidget, QSplitter, QTabWidget, QMenu, QStatusBar, QTextEdit, QSpinBox, QGroupBox)
import webbrowser
from PyQt5.QtGui import QPixmap, QFont, QColor, QIcon, QTextDocument
from PyQt5.QtCore import Qt, QUrl, QThread, pyqtSignal, QTimer
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
import subprocess
import tempfile
try:
    from plyer import notification # Requires 'pip install plyer'
except ImportError:
    notification = None
try:
    import requests # Requires 'pip install requests'
except ImportError:
    requests = None

from prodotto_dialog import ProdottoDialog
from prodotti_manager import (lista_prodotti, aggiungi_prodotto, modifica_prodotto, 
                              cancella_prodotto, get_existing_skus, pulisci_database, 
                              svuota_tutto, get_tipologie_prodotto, rinomina_tipologia, 
                              cancella_tipologia, aggiorna_tipologia_per_ids, get_listini, 
                              crea_listino, cancella_listino, get_prezzi_listino, 
                              aggiorna_prezzo_listino, salva_catalogo_db, get_cataloghi_db, 
                              rinomina_catalogo_db, get_counts_per_tipologia, 
                              cancella_catalogo_db, get_prezzi_prodotto,
                              get_suffisso_listino, aggiorna_suffisso_listino, get_suffissi_listini)
from email_utils import invia_email as send_email_logic
from import_utils import get_access_tables, read_access_table, read_excel_df, read_danea_xml, importa_dataframe_nel_db, sincronizza_immagini_database, pyodbc
from pdf_export import esporta_catalogo_pdf, FPDF
from db import init_db, DB_PATH

def format_prezzo_custom(prezzo, suffisso):
    """Formatta in modo intelligente un prezzo personalizzato in base al suo suffisso."""
    suff_clean = suffisso.strip()
    if not suff_clean:
        return f"{prezzo:.2f}"
    if suff_clean == '%':
        return f"{prezzo:.2f} %"
    if suff_clean == '€':
        return f"€ {prezzo:.2f}"
    # Se il suffisso è un simbolo di valuta comune lo mettiamo prima se opportuno, altrimenti dopo
    if suff_clean in ('$', '£', '¥'):
        return f"{suff_clean} {prezzo:.2f}"
    return f"{prezzo:.2f} {suff_clean}"

UPDATE_URL = "https://raw.githubusercontent.com/befree1986/catalog_manager_pro1/main/version.json" 

def parse_version(v):
    """Converte una stringa versione in una tupla di interi per confronti sicuri."""
    try:
        # Estrae solo i numeri iniziali da ogni parte (es: "5b" viene trattato come 5)
        parts = []
        for part in v.split('.'):
            match = re.search(r"(\d+)", part)
            parts.append(int(match.group(1)) if match else 0)
        return tuple(parts)
    except (ValueError, AttributeError):
        return (0, 0, 0)

class DownloadThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, url, dest_path):
        super().__init__()
        self.url = url
        self.dest_path = dest_path

    def run(self):
        try:
            if not requests:
                raise ImportError("Libreria 'requests' non trovata.")
            response = requests.get(self.url, stream=True)
            response.raise_for_status()
            total_length = int(response.headers.get('content-length', 0))
            downloaded = 0
            with open(self.dest_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    downloaded += len(chunk)
                    f.write(chunk)
                    if total_length > 0:
                        self.progress.emit(int(downloaded * 100 / total_length))
            self.finished.emit(self.dest_path)
        except Exception as e:
            self.error.emit(str(e))

class ProductCard(QWidget):
    def __init__(self, p, parent_window):
        super().__init__()
        self.p = p
        self.parent_window = parent_window
        self.setObjectName("ProductCard")
        self.setAttribute(Qt.WA_StyledBackground, True)
        
        # Layout verticale per la card (Immagine sopra, testo sotto)
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Immagine
        img_path_raw = p[6]
        img_path = self.parent_window.get_valid_image_path(img_path_raw)
            
        if img_path:
            pixmap = QPixmap(img_path)
            if not pixmap.isNull():
                img_label = QLabel()
                img_label.setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                img_label.setAlignment(Qt.AlignCenter)
                # Rimuovi ereditarietà stile per l'immagine
                img_label.setStyleSheet("border: none; background-color: transparent;")
                layout.addWidget(img_label)
        else:
            lbl_no_img = QLabel("🖼️")
            lbl_no_img.setAlignment(Qt.AlignCenter)
            lbl_no_img.setStyleSheet("font-size: 40px; color: #bdc3c7;")
            layout.addWidget(lbl_no_img)
        
        # Info
        codice_html = f"<span style='color:#7f8c8d; font-size:10px;'>SKU: {p[8]}</span><br>" if len(p) > 8 and p[8] else ""
        tipologia_html = ""
        if len(p) > 9 and p[9]:
            tipologia_html = f"<p style='color:#3498db; font-size:10px; font-style:italic; margin:0;'>{p[9]}</p>"
        info_text = f"<h4 style='margin:0; color:#2c3e50;'>{p[1]}</h4>{codice_html}<p style='color:gray; font-size:11px; margin-top:2px; margin-bottom:2px;'>{p[2]}</p>{tipologia_html}"
        layout.addWidget(QLabel(info_text))
        
        # --- Generazione Testo Prezzi in Base al Filtro ---
        active_text = "👀 Tutti i Prezzi"
        if hasattr(self.parent_window, 'price_filter_combo'):
            active_text = self.parent_window.price_filter_combo.currentText()
            
        price_text = ""
        if active_text == "👀 Tutti i Prezzi":
            price_text = f"<b style='font-size:14px; color:#27ae60;'>Base: € {(p[4] or 0.0):.2f}</b>"
            # Listino 2
            if len(p) > 7 and p[7]:
                try:
                    prezzo2 = float(p[7])
                    if prezzo2 > 0:
                        price_text += f"<br><span style='color:#e67e22; font-size:11px;'>Listino 2: € {prezzo2:.2f}</span>"
                except: pass
            # Listino 3
            if len(p) > 10 and p[10]:
                try:
                    prezzo3 = float(p[10])
                    if prezzo3 > 0:
                        price_text += f"<br><span style='color:#e67e22; font-size:11px;'>Listino 3: € {prezzo3:.2f}</span>"
                except: pass
            # Listino 4
            if len(p) > 11 and p[11]:
                try:
                    prezzo4 = float(p[11])
                    if prezzo4 > 0:
                        price_text += f"<br><span style='color:#e67e22; font-size:11px;'>Listino 4: € {prezzo4:.2f}</span>"
                except: pass
            # Listini custom/dinamici da DB
            try:
                custom_prices = get_prezzi_prodotto(p[0])
                suffissi = get_suffissi_listini()
                for listino_nome, custom_p in custom_prices.items():
                    if custom_p > 0:
                        suff = suffissi.get(listino_nome, "€")
                        valore_formattato = format_prezzo_custom(custom_p, suff)
                        price_text += f"<br><span style='color:#9b59b6; font-size:11px;'>{listino_nome}: {valore_formattato}</span>"
            except Exception as e:
                print(f"Errore nel recupero dei prezzi extra per prodotto {p[0]}: {e}")
                
        elif active_text == "💰 Prezzo Base":
            price_text = f"<b style='font-size:16px; color:#27ae60;'>€ {(p[4] or 0.0):.2f}</b>"
            
        elif active_text == "💰 Listino 2":
            prezzo2 = float(p[7]) if len(p) > 7 and p[7] else 0.0
            price_text = f"<b style='font-size:16px; color:#e67e22;'>Listino 2: € {prezzo2:.2f}</b>"
            
        elif active_text == "💰 Listino 3":
            prezzo3 = float(p[10]) if len(p) > 10 and p[10] else 0.0
            price_text = f"<b style='font-size:16px; color:#e67e22;'>Listino 3: € {prezzo3:.2f}</b>"
            
        elif active_text == "💰 Listino 4":
            prezzo4 = float(p[11]) if len(p) > 11 and p[11] else 0.0
            price_text = f"<b style='font-size:16px; color:#e67e22;'>Listino 4: € {prezzo4:.2f}</b>"
            
        elif active_text.startswith("📊 Listino:"):
            listino_nome = active_text.replace("📊 Listino: ", "")
            try:
                custom_prices = get_prezzi_prodotto(p[0])
                prezzo_cust = custom_prices.get(listino_nome, 0.0)
                suffissi = get_suffissi_listini()
                suff = suffissi.get(listino_nome, "€")
                valore_formattato = format_prezzo_custom(prezzo_cust, suff)
                price_text = f"<b style='font-size:16px; color:#9b59b6;'>{listino_nome}: {valore_formattato}</b>"
            except Exception as e:
                price_text = f"<b style='font-size:16px; color:#9b59b6;'>{listino_nome}: N/D</b>"

        layout.addWidget(QLabel(price_text))
        layout.addStretch()
        
        # --- Pulsanti Azione (Modifica / Elimina) ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(5)
        btn_layout.setContentsMargins(0, 5, 0, 0)
        
        # Tasto Visibilità (Switch)
        self.btn_visibilita = QPushButton("👁️" if p[5] else "🚫")
        self.btn_visibilita.setFixedSize(32, 28)
        self.btn_visibilita.setCursor(Qt.PointingHandCursor)
        if p[5]:
            self.btn_visibilita.setStyleSheet("background-color: #2ecc71; color: white; border: none; border-radius: 4px; font-size: 14px;")
            self.btn_visibilita.setToolTip("Visibile nel catalogo (Clicca per nascondere)")
        else:
            self.btn_visibilita.setStyleSheet("background-color: #95a5a6; color: white; border: none; border-radius: 4px; font-size: 14px;")
            self.btn_visibilita.setToolTip("Nascosto dal catalogo (Clicca per rendere visibile)")
        self.btn_visibilita.clicked.connect(self.toggle_visibility)
        
        btn_modifica = QPushButton("✏️ Modifica")
        btn_modifica.setCursor(Qt.PointingHandCursor)
        btn_modifica.setStyleSheet("background-color: #3498db; color: white; border: none; padding: 6px; border-radius: 4px; font-weight: bold;")
        btn_modifica.clicked.connect(self.modifica_click)
        
        btn_elimina = QPushButton("🗑️")
        btn_elimina.setCursor(Qt.PointingHandCursor)
        btn_elimina.setFixedSize(32, 28)
        btn_elimina.setToolTip("Elimina Prodotto")
        btn_elimina.setStyleSheet("background-color: #e74c3c; color: white; border: none; padding: 6px; border-radius: 4px;")
        btn_elimina.clicked.connect(self.elimina_click)
        
        btn_layout.addWidget(self.btn_visibilita)
        btn_layout.addWidget(btn_modifica)
        btn_layout.addWidget(btn_elimina)
        
        layout.addLayout(btn_layout)
        self.setLayout(layout)

        # --- Effetto Ombra per Hover ---
        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(15)
        self.shadow.setXOffset(0)
        self.shadow.setYOffset(4)
        self.shadow.setColor(QColor(0, 0, 0, 30)) # Ombra leggera di base
        self.setGraphicsEffect(self.shadow)

    def enterEvent(self, event):
        # Aumenta l'ombra quando il mouse entra
        self.shadow.setBlurRadius(25)
        self.shadow.setYOffset(8)
        self.shadow.setColor(QColor(0, 0, 0, 60))
        super().enterEvent(event)

    def leaveEvent(self, event):
        # Ripristina l'ombra base quando il mouse esce
        self.shadow.setBlurRadius(15)
        self.shadow.setYOffset(4)
        self.shadow.setColor(QColor(0, 0, 0, 30))
        super().leaveEvent(event)

    def mouseDoubleClickEvent(self, event):
        self.parent_window.modifica_articolo(self.p)

    def toggle_visibility(self):
        # Inverte la visibilità (1 -> 0, 0 -> 1)
        new_visibility = 0 if self.p[5] else 1
        
        # Recupera i dati correnti per chiamare modifica_prodotto
        p = self.p
        # Gestione campi opzionali per evitare errori di indice
        tipologia = p[9] if len(p) > 9 else "Generico"
        codice = p[8] if len(p) > 8 else ""
        p2 = p[7] if len(p) > 7 else 0
        p3 = p[10] if len(p) > 10 else 0
        p4 = p[11] if len(p) > 11 else 0
        
        qt2 = p[12] if len(p) > 12 else 0
        qt3 = p[13] if len(p) > 13 else 0
        qt4 = p[14] if len(p) > 14 else 0
        
        modifica_prodotto(p[0], p[1], p[2], p[3], p[4], new_visibility, p[6], p2, codice, tipologia, p3, p4, qt2, qt3, qt4)
        
        # Aggiorna la dashboard e la griglia per riflettere il cambiamento
        self.parent_window.update_dashboard_data()
        self.parent_window.aggiorna_griglia_prodotti()

    def modifica_click(self):
        self.parent_window.modifica_articolo(self.p)

    def elimina_click(self):
        reply = QMessageBox.question(self.parent_window, 'Conferma Eliminazione', 
                                     f"Sei sicuro di voler eliminare il prodotto '{self.p[1]}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            cancella_prodotto(self.p[0])
            self.parent_window.update_dashboard_data()
            self.parent_window.aggiorna_griglia_prodotti()

class TipologiaCard(QWidget):
    def __init__(self, nome_tipologia, conteggio, parent_window):
        super().__init__()
        self.nome_tipologia = nome_tipologia
        self.parent_window = parent_window
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("ProductCard") # Riusiamo lo stile delle card prodotto
        self.setAttribute(Qt.WA_StyledBackground, True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(0)

        # Header con pulsante Menu
        header_layout = QHBoxLayout()
        header_layout.addStretch()
        self.menu_btn = QPushButton("⋮")
        self.menu_btn.setFixedSize(24, 24)
        self.menu_btn.setStyleSheet("QPushButton { border: none; font-weight: bold; font-size: 16px; background: transparent; color: #7f8c8d; } QPushButton:hover { background-color: #ecf0f1; border-radius: 12px; }")
        self.menu_btn.setCursor(Qt.PointingHandCursor)
        self.menu_btn.clicked.connect(self.show_menu)
        header_layout.addWidget(self.menu_btn)
        layout.addLayout(header_layout)

        # Contenuto centrato
        content_layout = QVBoxLayout()
        content_layout.setAlignment(Qt.AlignCenter)

        icon = QLabel("📁")
        icon.setStyleSheet("font-size: 48px; border: none; background: transparent;")
        icon.setAlignment(Qt.AlignCenter)

        nome_label = QLabel(f"<b>{nome_tipologia}</b>")
        nome_label.setAlignment(Qt.AlignCenter)
        nome_label.setWordWrap(True)
        nome_label.setStyleSheet("border: none; background: transparent;")

        conteggio_label = QLabel(f"{conteggio} prodotti")
        conteggio_label.setStyleSheet("color: gray; border: none; background: transparent;")
        conteggio_label.setAlignment(Qt.AlignCenter)

        content_layout.addWidget(icon)
        content_layout.addWidget(nome_label)
        content_layout.addWidget(conteggio_label)
        
        layout.addLayout(content_layout)
        layout.addStretch()

    def mousePressEvent(self, event):
        # Evita di aprire il gruppo se si clicca sul pulsante menu (gestito separatamente)
        if self.childAt(event.pos()) != self.menu_btn:
            self.parent_window.mostra_prodotti_per_tipologia(self.nome_tipologia)

    def show_menu(self):
        menu = QMenu(self)
        
        action_rename = QAction("✏️ Rinomina", self)
        action_rename.triggered.connect(self.rinomina)
        menu.addAction(action_rename)
        
        action_delete = QAction("🗑️ Elimina Gruppo", self)
        action_delete.triggered.connect(self.elimina)
        menu.addAction(action_delete)
        
        # Mostra il menu sotto il pulsante
        menu.exec_(self.menu_btn.mapToGlobal(self.menu_btn.rect().bottomLeft()))

    def rinomina(self):
        nuovo_nome, ok = QInputDialog.getText(self, "Rinomina Gruppo", f"Nuovo nome per '{self.nome_tipologia}':", text=self.nome_tipologia)
        if ok and nuovo_nome and nuovo_nome != self.nome_tipologia:
            rinomina_tipologia(self.nome_tipologia, nuovo_nome)
            # Aggiorna UI
            self.parent_window.update_dashboard_data()
            self.parent_window.aggiorna_griglia_tipologie()
            self.parent_window.refresh_tipologie_combos()

    def elimina(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Elimina Gruppo")
        msg_box.setText(f"Vuoi eliminare il gruppo '{self.nome_tipologia}'?")
        msg_box.setInformativeText("Cosa vuoi fare con i prodotti contenuti in questo gruppo?")
        msg_box.setIcon(QMessageBox.Question)
        
        btn_keep = msg_box.addButton("Sposta in 'Generico'", QMessageBox.ActionRole)
        btn_delete = msg_box.addButton("Elimina anche i Prodotti", QMessageBox.ActionRole)
        btn_cancel = msg_box.addButton(QMessageBox.Cancel)
        
        msg_box.exec_()
        
        clicked_button = msg_box.clickedButton()
        if clicked_button == btn_cancel:
            return
            
        elimina_prodotti = (clicked_button == btn_delete)
        
        cancella_tipologia(self.nome_tipologia, elimina_prodotti)
        
        # Aggiorna UI
        self.parent_window.update_dashboard_data()
        self.parent_window.aggiorna_griglia_tipologie()
        self.parent_window.refresh_tipologie_combos()

class LatestProductItem(QWidget):
    def __init__(self, p, parent_window):
        super().__init__()
        self.p = p
        self.parent_window = parent_window
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # 1. Immagine
        img_label = QLabel()
        img_label.setFixedSize(60, 60)
        img_label.setStyleSheet("border: 1px solid #e0e0e0; border-radius: 4px;")
        img_path_raw = p[6]
        img_path = self.parent_window.get_valid_image_path(img_path_raw)
        if img_path:
            pixmap = QPixmap(img_path)
            img_label.setPixmap(pixmap.scaled(60, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            img_label.setText("🖼️")
            img_label.setAlignment(Qt.AlignCenter)
            img_label.setStyleSheet("font-size: 24px; background-color: #f0f0f0; border-radius: 4px;")
        layout.addWidget(img_label)
        
        # 2. Info (Nome, SKU, Descrizione, Prezzo)
        info_layout = QVBoxLayout()
        info_layout.setSpacing(2)
        
        nome_sku_text = f"<b style='font-size:14px;'>{p[1]}</b>"
        if len(p) > 8 and p[8]:
            nome_sku_text += f" <span style='font-size:11px; color:gray;'> (SKU: {p[8]})</span>"
        info_layout.addWidget(QLabel(nome_sku_text))
        
        desc = str(p[3]) if p[3] and not isinstance(p[3], float) else ""
        desc_text = (desc[:80] + '...') if len(desc) > 80 else desc
        info_layout.addWidget(QLabel(f"<i style='color:#555;'>{desc_text}</i>"))
        
        info_layout.addWidget(QLabel(f"<b style='color:#27ae60; font-size:14px;'>€ {p[4]:.2f}</b>"))
        
        layout.addLayout(info_layout, 1) # Stretchable
        
        # 3. Pulsante Visibilità
        self.visibility_btn = QPushButton("👁️" if p[5] else "🚫")
        self.visibility_btn.setFixedSize(32, 32)
        self.visibility_btn.setCursor(Qt.PointingHandCursor)
        self.visibility_btn.setStyleSheet("background-color: #e8f8f5; border:none; border-radius:16px; font-size:16px;" if p[5] else "background-color: #fdedec; border:none; border-radius:16px; font-size:16px;")
        self.visibility_btn.setToolTip("Clicca per cambiare visibilità")
        self.visibility_btn.clicked.connect(self.toggle_visibility)
        layout.addWidget(self.visibility_btn)

    def toggle_visibility(self):
        new_visibility = 0 if self.p[5] else 1
        # Campi extra
        p = self.p
        p3 = p[10] if len(p) > 10 else 0
        p4 = p[11] if len(p) > 11 else 0
        qt2 = p[12] if len(p) > 12 else 0
        qt3 = p[13] if len(p) > 13 else 0
        qt4 = p[14] if len(p) > 14 else 0
        modifica_prodotto(self.p[0], self.p[1], self.p[2], self.p[3], self.p[4], new_visibility, self.p[6], self.p[7], self.p[8], self.p[9] if len(self.p) > 9 else "Generico", p3, p4, qt2, qt3, qt4)
        self.parent_window.update_dashboard_data()
        self.parent_window.aggiorna_griglia_prodotti()

class ColumnMappingDialog(QDialog):
    def __init__(self, excel_columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Mappatura Colonne")
        self.setMinimumWidth(500)
        self.excel_columns = excel_columns
        # Campi target che l'applicazione si aspetta
        self.target_fields = ['nome', 'codice', 'tipologia_prodotto', 'categoria', 'descrizione', 
                              'prezzo', 'prezzo_secondario', 'prezzo3', 'prezzo4', 
                              'qta_min_2', 'qta_min_3', 'qta_min_4', 'immagine']
        self.mappings = {}
        
        self.price_list_mappings = {} # {colonna_excel: nome_listino}

        layout = QVBoxLayout(self)
        
        # Toolbar per preset
        preset_layout = QHBoxLayout()
        btn_save_preset = QPushButton("💾 Salva Preset Mappatura")
        btn_save_preset.clicked.connect(self.save_preset)
        btn_load_preset = QPushButton("📂 Carica Preset")
        btn_load_preset.clicked.connect(self.load_preset)
        preset_layout.addWidget(btn_save_preset)
        preset_layout.addWidget(btn_load_preset)
        preset_layout.addStretch()
        layout.addLayout(preset_layout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        form_layout = QFormLayout(content_widget)

        info = QLabel("Associa le colonne del tuo file Excel ai campi del database.\nSe un campo non viene associato, sarà ignorato.")
        info.setStyleSheet("margin-bottom: 10px;")
        form_layout.addRow(info)

        form_layout.addRow(QLabel("<b>CAMPI PRINCIPALI:</b>"))
        
        for field in self.target_fields:
            combo = QComboBox()
            # Aggiungi opzione per ignorare e poi le colonne del file
            combo.addItems(["[ Ignora ]"] + self.excel_columns)
            
            # Heuristic per pre-selezionare la corrispondenza migliore
            best_match = self._find_best_match(field, self.excel_columns)
            if best_match:
                combo.setCurrentText(best_match)
            
            label_text = f"Campo '{field.replace('_', ' ').title()}':"
            form_layout.addRow(label_text, combo)
            self.mappings[field] = combo

        form_layout.addRow(QLabel("<hr><b>LISTINI PREZZI EXTRA (Opzionale):</b>"))
        form_layout.addRow(QLabel("Seleziona le colonne che contengono prezzi per listini specifici."))

        # Contenitore per i listini extra per evitare problemi con QFormLayout
        self.listini_container = QWidget()
        self.listini_layout = QVBoxLayout(self.listini_container)
        self.listini_layout.setContentsMargins(0, 0, 0, 0)
        self.listini_layout.setSpacing(5)
        form_layout.addRow(self.listini_container)

        # Sezione dinamica per listini
        self.form_layout = form_layout
        self.listino_rows = []
        
        btn_add_listino = QPushButton("➕ Aggiungi Mappatura Listino")
        btn_add_listino.clicked.connect(lambda checked=False: self.add_listino_row(None, None))
        
        form_layout.addRow("Nuovo Listino:", btn_add_listino)
        
        # Auto-detect listini
        self.detect_potential_price_lists()

        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.button(QDialogButtonBox.Ok).setText("Continua")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def add_listino_row(self, default_col=None, default_name=None):
        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(0,0,0,0)
        
        combo_col = QComboBox()
        combo_col.addItems(self.excel_columns)
        if default_col and default_col in self.excel_columns:
            combo_col.setCurrentText(default_col)
            
        line_name = QLineEdit()
        line_name.setPlaceholderText("Nome Listino (es. Rivenditori)")
        if default_name:
            line_name.setText(default_name)
            
        btn_del = QPushButton("❌")
        btn_del.setFixedSize(24, 24)
        btn_del.clicked.connect(lambda checked=False: self.remove_listino_row(row_widget))
        
        row_layout.addWidget(QLabel("Colonna:"))
        row_layout.addWidget(combo_col, 1)
        row_layout.addWidget(QLabel(" -> Listino:"))
        row_layout.addWidget(line_name, 1)
        row_layout.addWidget(btn_del)
        
        # Salva riferimenti per recuperare i dati dopo
        row_widget.combo_col = combo_col
        row_widget.line_name = line_name
        
        # Aggiungi al layout del contenitore
        self.listino_rows.append(row_widget)
        self.listini_layout.addWidget(row_widget)
        row_widget.show()
        
        # Forza il ridisegno del dialogo
        self.adjustSize()
        self.update()
            
    def remove_listino_row(self, widget):
        if widget:
            # Nasconde subito visivamente il widget per un feedback immediato
            widget.hide()
            # Esegue la rimozione effettiva e l'aggiustamento del layout nel ciclo successivo dell'event loop
            QTimer.singleShot(0, lambda: self._do_remove_listino_row(widget))

    def _do_remove_listino_row(self, widget):
        if widget in self.listino_rows:
            self.listino_rows.remove(widget)
        self.listini_layout.removeWidget(widget)
        widget.setParent(None)
        widget.deleteLater()
        self.adjustSize()
        self.update()


    def _find_best_match(self, field, columns):
        """Tenta di trovare la migliore corrispondenza per un campo nelle colonne date."""
        # 1. Corrispondenza esatta (case-insensitive, già gestito da read_excel_df)
        if field in columns:
            return field
        # 2. Caso speciale comune (codice/sku)
        if field == 'codice' and 'sku' in columns:
            return 'sku'
        # 3. Corrispondenza parziale
        for col in columns:
            if field in col:
                return col
        return None

    def detect_potential_price_lists(self):
        """Cerca colonne che sembrano listini ma non sono mappate."""
        for col in self.excel_columns:
            lower_col = col.lower()
            if "listino" in lower_col or "price" in lower_col or "prezzo" in lower_col:
                # Se non è già un target field standard
                if not any(col == self._find_best_match(tf, self.excel_columns) for tf in self.target_fields):
                     # Suggerisci come listino
                     suggested_name = col.title().replace("_", " ")
                     self.add_listino_row(col, suggested_name)

    def get_column_mappings(self):
        """Restituisce un dizionario di mappatura: {'target_field': 'excel_column_name'}."""
        result = {}
        for field, combo in self.mappings.items():
            selected = combo.currentText()
            if selected != "[ Ignora ]":
                result[field] = selected
        return result

    def get_price_list_mappings(self):
        """Restituisce {colonna_excel: nome_listino}"""
        mappings = {}
        for widget in self.listino_rows:
            try:
                col = widget.combo_col.currentText()
                name = widget.line_name.text().strip()
                if col and name:
                    mappings[col] = name
            except RuntimeError:
                pass
        return mappings

    def save_preset(self):
        mapping = self.get_column_mappings()
        if not mapping:
            QMessageBox.warning(self, "Attenzione", "Nessuna mappatura da salvare.")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, 'Salva Preset Mappatura', 'mappatura.json', 'JSON (*.json)')
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(mapping, f)
                QMessageBox.information(self, "Successo", "Preset salvato correttamente.")
            except Exception as e:
                QMessageBox.critical(self, "Errore", f"Impossibile salvare il file: {e}")

    def load_preset(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Carica Preset Mappatura', '', 'JSON (*.json)')
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    mapping = json.load(f)
                
                for field, excel_col in mapping.items():
                    if field in self.mappings:
                        # Trova l'indice della colonna excel nel combobox
                        index = self.mappings[field].findText(excel_col)
                        if index >= 0:
                            self.mappings[field].setCurrentIndex(index)
            except Exception as e:
                QMessageBox.critical(self, "Errore", f"Impossibile caricare il file: {e}")

class PreviewDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.existing_skus = get_existing_skus() # Carica SKU dal DB per validazione
        self.setWindowTitle("Anteprima Importazione Dati")
        self.setMinimumSize(850, 600)
        
        layout = QVBoxLayout(self)
        
        info_label = QLabel(
            f"Trovate {len(df)} righe da importare. Controlla e modifica i dati.\n"
            "🔴 Rosso: Prezzo non valido | 🟠 Arancione: SKU duplicato nel DB | 🟡 Giallo: SKU duplicato nel file\n"
            "💡 Suggerimento: Fai doppio clic sull'intestazione di una colonna per rinominarla (es. Aliquota -> IVA)."
        )
        info_label.setStyleSheet("font-style: italic; color: #2c3e50; line-height: 1.4; margin-bottom: 10px;")
        layout.addWidget(info_label)
        
        self.table = QTableWidget()
        self.populate_table(df)
        self.table.itemChanged.connect(self.validate_item) # Connetti segnale per validazione dinamica
        
        # Abilita doppio clic su intestazione colonna per rinominarla
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.rename_column)
        
        layout.addWidget(self.table)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.button(QDialogButtonBox.Ok).setText("Conferma Importazione")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        # Rendi la tabella esplicitamente modificabile dall'utente
        self.table.setEditTriggers(QTableWidget.AllEditTriggers)

    def rename_column(self, index):
        old_name = self.table.horizontalHeaderItem(index).text()
        new_name, ok = QInputDialog.getText(self, "Rinomina Colonna", f"Rinomina la colonna '{old_name}' in:", QLineEdit.Normal, old_name)
        if ok and new_name.strip():
            self.table.horizontalHeaderItem(index).setText(new_name.strip())
        self.table.setEditTriggers(QTableWidget.AllEditTriggers)

    def _validate_cell(self, item):
        """Controlla la validità di una cella. Restituisce (is_valid, error_type)."""
        price_columns = ['prezzo', 'prezzo_secondario', 'prezzo 1', 'prezzo 2']
        sku_columns = ['codice', 'sku']
        # Assicurati che l'header esista prima di accedervi
        header_item = self.table.horizontalHeaderItem(item.column())
        if not header_item:
            return True, None
            
        col_name = header_item.text().lower()

        # Validazione Prezzi
        if col_name in price_columns:
            value = item.text()
            if not value.strip():
                return True, None
            try:
                float(value.replace(',', '.'))
                return True, None
            except ValueError:
                return False, 'price_error'
        
        # Validazione SKU (Duplicati)
        if col_name in sku_columns:
            value = item.text().strip()
            if not value: return True, None
            
            # Controlla duplicato nel DB
            if value in self.existing_skus:
                return False, 'sku_duplicate_db'
            
            # Controlla duplicati nel file
            count = 0
            for i in range(self.table.rowCount()):
                cell = self.table.item(i, item.column())
                if cell and cell.text().strip() == value:
                    count += 1
            if count > 1:
                return False, 'sku_duplicate_file'

        return True, None

    def validate_item(self, item):
        """Slot che viene chiamato quando un item della tabella viene modificato."""
        is_valid, error_type = self._validate_cell(item)
        
        # Blocca i segnali per evitare ricorsioni infinite durante la modifica dello sfondo
        self.table.blockSignals(True)
        if is_valid:
            item.setBackground(QColor("white"))
            item.setToolTip("")
        else:
            if error_type == 'price_error':
                item.setBackground(QColor(255, 224, 224))  # Rosso chiaro
                item.setToolTip("Valore numerico non valido")
            elif error_type == 'sku_duplicate_db':
                item.setBackground(QColor(255, 240, 220))  # Arancione
                item.setToolTip("Questo SKU esiste già nel database!")
            elif error_type == 'sku_duplicate_file':
                item.setBackground(QColor(255, 250, 205))  # Giallo
                item.setToolTip("SKU duplicato in questo file")
        self.table.blockSignals(False)

    def get_dataframe(self):
        """Estrae i dati dalla tabella (potenzialmente modificati) e li restituisce come DataFrame pandas."""
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        headers = [self.table.horizontalHeaderItem(j).text() for j in range(cols)]
        data = []
        for i in range(rows):
            row_data = []
            for j in range(cols):
                item = self.table.item(i, j)
                row_data.append(item.text() if item else '')
            data.append(row_data)
        
        return pd.DataFrame(data, columns=headers)

    def populate_table(self, df):
        # Riempiamo il df con stringhe vuote per evitare 'nan' e problemi di visualizzazione
        df_display = df.fillna('').astype(str)
        
        self.table.setRowCount(df_display.shape[0])
        self.table.setColumnCount(df_display.shape[1])
        self.table.setHorizontalHeaderLabels(df_display.columns)
        
        # Blocca i segnali durante il popolamento iniziale per performance
        self.table.blockSignals(True)
        for i, row in enumerate(df_display.itertuples(index=False)):
            for j, value in enumerate(row):
                item = QTableWidgetItem(value)
                self.table.setItem(i, j, item)
                # Esegui la validazione iniziale
                self.validate_item(item)
        self.table.blockSignals(False)
        
        self.table.resizeColumnsToContents()

class LivePreviewCard(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("LivePreviewCard")
        self.setFixedSize(180, 220)
        self.setStyleSheet("#LivePreviewCard { background-color: white; border: 1px solid #ccc; border-radius: 5px; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(5)

        self.image_placeholder = QLabel("🖼️")
        self.image_placeholder.setAlignment(Qt.AlignCenter)
        self.image_placeholder.setStyleSheet("font-size: 60px; color: #ddd; border: none; background-color: transparent;")
        self.image_placeholder.setFixedHeight(100)
        
        self.title_label = QLabel("Nome Prodotto Esempio")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setWordWrap(True)

        self.price_label = QLabel("€ 99.99")
        self.price_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.image_placeholder)
        layout.addWidget(self.title_label)
        layout.addWidget(self.price_label)
        layout.addStretch()

    def update_style(self, style):
        title_size = style.get('grid_title_size', 10)
        title_color = style.get('grid_title_color', '#2c3e50')
        self.title_label.setStyleSheet(f"font-size: {title_size}pt; font-weight: bold; color: {title_color}; border: none; background-color: transparent;")

        price_size = style.get('grid_price_size', 11)
        price_color = style.get('grid_price_color', '#27ae60')
        self.price_label.setStyleSheet(f"font-size: {price_size}pt; font-weight: bold; color: {price_color}; border: none; background-color: transparent;")

class TemplateEditorDialog(QDialog):
    """Micro editor per personalizzare gli stili del catalogo PDF"""
    def __init__(self, current_style=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Editor Template Catalogo")
        self.setMinimumWidth(800)
        self.style = current_style or {}
        
        main_layout = QHBoxLayout(self)
        
        # --- Left Side: Editor Controls ---
        editor_widget = QWidget()
        layout = QVBoxLayout(editor_widget)
        
        info = QLabel("Modifica l'aspetto degli elementi nel PDF.\nLe modifiche si riflettono nell'anteprima a destra.")
        layout.addWidget(info)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        form_widget = QWidget()
        form_layout = QFormLayout()
        form_widget.setLayout(form_layout)
        
        # Header
        self.header_size = QSpinBox()
        self.header_size.setRange(10, 72)
        self.header_size.setValue(self.style.get('header_font_size', 20))
        form_layout.addRow("Dimensione Intestazione (Titolo):", self.header_size)

        # Lista Layout
        grp_lista = QGroupBox("Layout Lista")
        lista_layout = QFormLayout(grp_lista)
        
        self.prod_title = QSpinBox()
        self.prod_title.setRange(8, 36)
        self.prod_title.setValue(self.style.get('product_title_size', 14))
        lista_layout.addRow("Titolo Prodotto:", self.prod_title)
        
        self.desc_size = QSpinBox()
        self.desc_size.setRange(6, 24)
        self.desc_size.setValue(self.style.get('desc_size', 12))
        lista_layout.addRow("Descrizione:", self.desc_size)
        
        self.price_size = QSpinBox()
        self.price_size.setRange(8, 36)
        self.price_size.setValue(self.style.get('price_size', 16))
        lista_layout.addRow("Prezzo:", self.price_size)
        
        form_layout.addRow(grp_lista)
        
        # Griglia Layout
        grp_grid = QGroupBox("Layout Griglia (Compatto)")
        grid_layout = QFormLayout(grp_grid)
        
        self.grid_title = QSpinBox()
        self.grid_title.setRange(6, 24)
        self.grid_title.setValue(self.style.get('grid_title_size', 10))
        grid_layout.addRow("Dimensione Titolo:", self.grid_title)

        self.grid_title_color_btn = QPushButton(self.style.get('grid_title_color', '#2c3e50'))
        self.grid_title_color_btn.clicked.connect(lambda: self.pick_color(self.grid_title_color_btn, 'grid_title_color'))
        grid_layout.addRow("Colore Titolo:", self.grid_title_color_btn)
        
        self.grid_price = QSpinBox()
        self.grid_price.setRange(6, 24)
        self.grid_price.setValue(self.style.get('grid_price_size', 11))
        grid_layout.addRow("Dimensione Prezzo:", self.grid_price)

        self.grid_price_color_btn = QPushButton(self.style.get('grid_price_color', '#27ae60'))
        self.grid_price_color_btn.clicked.connect(lambda: self.pick_color(self.grid_price_color_btn, 'grid_price_color'))
        grid_layout.addRow("Colore Prezzo:", self.grid_price_color_btn)
        
        form_layout.addRow(grp_grid)
        
        scroll.setWidget(form_widget)
        layout.addWidget(scroll)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        # --- Right Side: Live Preview ---
        preview_container = QWidget()
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setAlignment(Qt.AlignCenter)
        preview_label = QLabel("<b>Anteprima Live (Griglia)</b>")
        preview_label.setAlignment(Qt.AlignCenter)
        
        self.live_preview = LivePreviewCard(self)
        
        preview_layout.addWidget(preview_label)
        preview_layout.addWidget(self.live_preview)
        
        main_layout.addWidget(editor_widget, 1)
        main_layout.addWidget(preview_container)

        # Connect signals
        self.header_size.valueChanged.connect(self.update_preview)
        self.prod_title.valueChanged.connect(self.update_preview)
        self.desc_size.valueChanged.connect(self.update_preview)
        self.price_size.valueChanged.connect(self.update_preview)
        self.grid_title.valueChanged.connect(self.update_preview)
        self.grid_price.valueChanged.connect(self.update_preview)
        
        self.update_preview() # Initial call

    def pick_color(self, btn, style_key):
        initial_color = QColor(self.style.get(style_key, '#000000'))
        color = QColorDialog.getColor(initial_color, self, "Scegli Colore")
        if color.isValid():
            hex_color = color.name()
            self.style[style_key] = hex_color
            self.update_preview()

    def update_preview(self):
        current_style = self.get_style(read_controls=True)
        self.live_preview.update_style(current_style)
        
        # Update color button styles without opening a dialog
        for btn, key in [(self.grid_title_color_btn, 'grid_title_color'), (self.grid_price_color_btn, 'grid_price_color')]:
            color_hex = current_style.get(key, '#000000')
            color = QColor(color_hex)
            btn.setText(color_hex)
            btn.setStyleSheet(f"background-color: {color_hex}; color: {'white' if color.lightness() < 128 else 'black'}; font-weight: bold;")

    def get_style(self, read_controls=False):
        # Legge i valori dai controlli e li mette nel dizionario self.style
        self.style['header_font_size'] = self.header_size.value()
        self.style['product_title_size'] = self.prod_title.value()
        self.style['desc_size'] = self.desc_size.value()
        self.style['price_size'] = self.price_size.value()
        self.style['grid_title_size'] = self.grid_title.value()
        self.style['grid_price_size'] = self.grid_price.value()
        
        # I colori sono già in self.style grazie a pick_color
        # Assicuriamo valori di default se non sono mai stati scelti
        if 'grid_title_color' not in self.style:
            self.style['grid_title_color'] = '#2c3e50'
        if 'grid_price_color' not in self.style:
            self.style['grid_price_color'] = '#27ae60'
            
        return self.style

class SupportDialog(QDialog):
    """Dialogo per l'invio di segnalazioni ed errori all'assistenza."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Contatta Assistenza / Segnala Errore")
        self.setMinimumWidth(500)
        layout = QVBoxLayout(self)
        
        header = QLabel("<b>Invia un messaggio all'assistenza tecnica</b>")
        header.setStyleSheet("font-size: 14px; margin-bottom: 5px;")
        layout.addWidget(header)
        
        info = QLabel("Descrivi il problema o il suggerimento. Le informazioni verranno inviate in modo sicuro.")
        info.setWordWrap(True)
        info.setStyleSheet("color: #555; margin-bottom: 10px;")
        layout.addWidget(info)
        
        self.messaggio = QTextEdit()
        self.messaggio.setPlaceholderText("Scrivi qui il tuo messaggio...")
        layout.addWidget(self.messaggio)
        
        self.chk_logs = QCheckBox("Includi file di log (debug_log.txt) per diagnosi tecnica")
        self.chk_logs.setChecked(True)
        layout.addWidget(self.chk_logs)
        
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.button(QDialogButtonBox.Ok).setText("Invia Messaggio")
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_data(self):
        return self.messaggio.toPlainText(), self.chk_logs.isChecked()

class CatalogoMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Controlla se siamo in modalità sviluppo tramite variabile d'ambiente
        self.dev_mode = os.environ.get("CATALOG_DEV_MODE") == "1"
        if self.dev_mode:
            print(f"DEBUG: Avvio v{APP_VERSION} in modalità Anteprima Sviluppo")

        if not self._check_db_writable():
            logging.error("Avvio interrotto: Il database non è scrivibile. Verificare i permessi della cartella di installazione.")
            QMessageBox.critical(None, "Errore Permessi", 
                                 "Impossibile scrivere nel database (catalogo.db).\n\n"
                                 "Assicurati che l'applicazione non sia in una cartella protetta "
                                 "(come Programmi senza permessi admin) o che il file non sia bloccato.")
            sys.exit(1)

        if not self.dev_mode:
            self.backup_database() # Salta backup lento in dev
        else:
            print("DEBUG: Backup automatico saltato.")

        init_db()  # Inizializza il database all'avvio
        self.setWindowTitle('Dashboard Catalogo')
        self.resize(1024, 768)

        # Imposta l'icona dell'applicazione
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        # File per salvare l'ordine custom delle categorie
        self.category_order_file = os.path.join(self.base_dir, "category_order.json")
        self.catalog_structure_file = os.path.join(self.base_dir, "catalog_structure.json")

        self.auto_save_config = self._load_auto_save_config()
        self._configure_auto_save_timer()

        # Impostazioni predefinite Catalogo
        self.catalog_settings = {
            'title': 'Catalogo Prodotti',
            'company': 'La Mia Azienda',
            'color': '#2c3e50',
            'show_prices': True,
            'layout': 'Lista',
            'category_filter': 'Tutte le categorie',
            'cover_image': None,
            'tipologia_filter': 'Tutte',
            'include_index': True,
            'sort_mode': 'Struttura Personalizzata', # Default changed
            'style': {}, # Configurazione micro editor
            'structure': [], # Lista ordinata di elementi (Categorie, Pagine Extra)
            'show_price_base': True,
            'show_price2': False,
            'show_price3': False,
            'show_price4': False,
            'show_custom_listini': []
        }

        self.init_ui()
        # Caricamento iniziale dati
        self.switch_page(0) 

    def salva_suffisso_listino_corrente(self):
        """Salva il simbolo/suffisso (es. € o %) per il listino selezionato."""
        if self.current_listino_id is None:
            QMessageBox.warning(self, "Attenzione", "Nessun listino selezionato per salvare il suffisso.")
            return
        
        suffisso = self.input_suffisso.text().strip()
        aggiorna_suffisso_listino(self.current_listino_id, suffisso)
        QMessageBox.information(self, "Successo", "Suffisso listino aggiornato.")
        # Aggiorna la combo dei filtri prezzo per riflettere il nuovo suffisso
        self.refresh_price_filter_combo()

    def sincronizza_immagini_correnti(self):
        """Sincronizza le immagini per il gruppo di prodotti attualmente visualizzato."""
        tipologia = getattr(self, 'current_product_group_filter', None)
        if not tipologia:
            QMessageBox.warning(self, "Attenzione", "Seleziona prima un gruppo di prodotti dalla griglia principale.")
            return
            
        folder = QFileDialog.getExistingDirectory(self, f"Seleziona cartella immagini per il gruppo: {tipologia}")
        if not folder:
            return
            
        progress = QProgressDialog(f"Sincronizzazione immagini per '{tipologia}'...", "Annulla", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.show()
        
        def update_progress(curr, total):
            progress.setMaximum(total)
            progress.setValue(curr)
            QApplication.processEvents()
            
        try:
            count = sincronizza_immagini_database(tipologia, folder, update_progress)
            QMessageBox.information(self, "Sincronizzazione Completata", 
                                    f"Operazione riuscita!\n\nSono state aggiornate le immagini per {count} prodotti nel gruppo '{tipologia}'.")
            self.aggiorna_griglia_prodotti()
        except Exception as e:
            logging.error(f"Errore sincronizzazione immagini: {e}")
            QMessageBox.critical(self, "Errore", f"Si è verificato un errore durante la sincronizzazione:\n{e}")
        finally:
            progress.close()

    def filter_listini_table(self):
        if self.current_listino_id is not None:
            self.load_prezzi_table(self.current_listino_id)

    def load_listini_page_logic(self):
        # Carica sidebar
        self.listini_list_widget.clear()
        listini = get_listini() # [(id, nome, desc), ...]
        if not listini:
            crea_listino("Listino Base", "Listino predefinito")
            listini = get_listini()
        
        # Popola i filtri
        self.listini_category_combo.blockSignals(True)
        self.listini_tipologia_combo.blockSignals(True)
        
        self.listini_category_combo.clear()
        self.listini_category_combo.addItem("Tutte le Categorie")
        self.listini_tipologia_combo.clear()
        self.listini_tipologia_combo.addItem("Tutte le Tipologie")
        
        all_products = lista_prodotti()
        categories = sorted(list(set(p[2] for p in all_products if p[2])))
        tipologie = get_tipologie_prodotto() # Use existing function
        
        self.listini_category_combo.addItems(categories)
        self.listini_tipologia_combo.addItems(tipologie)
        
        self.listini_category_combo.blockSignals(False)
        self.listini_tipologia_combo.blockSignals(False)
        
        # Aggiungi listini alla lista
        from PyQt5.QtWidgets import QListWidgetItem
        for l in listini:
            list_item = QListWidgetItem(l[1])
            list_item.setData(Qt.UserRole, l[0]) # ID
            list_item.setSizeHint(list_item.sizeHint()) 
            self.listini_list_widget.addItem(list_item)
            
        # Se c'era un listino selezionato, riselezionalo o seleziona il primo
        if self.listini_list_widget.count() > 0:
            self.listini_list_widget.setCurrentRow(0)
            self.on_listino_selected(self.listini_list_widget.item(0))

    def on_listino_selected(self, item):
        self.current_listino_id = item.data(Qt.UserRole)
        self.lbl_listino_corrente.setText(f"Modifica Prezzi: {item.text()}")
        self.load_prezzi_table(self.current_listino_id)

    def load_prezzi_table(self, listino_id):
        if listino_id is None: return
        
        all_prodotti = lista_prodotti()
        
        # Filtra prodotti
        search_text = self.listini_search_input.text().lower() if hasattr(self, 'listini_search_input') else ""
        category_filter = self.listini_category_combo.currentText() if hasattr(self, 'listini_category_combo') else "Tutte le Categorie"
        tipologia_filter = self.listini_tipologia_combo.currentText() if hasattr(self, 'listini_tipologia_combo') else "Tutte le Tipologie"
        
        prodotti_filtrati = []
        for p in all_prodotti:
            name_match = search_text in p[1].lower()
            cat_match = category_filter == "Tutte le Categorie" or p[2] == category_filter
            tipo_match = tipologia_filter == "Tutte le Tipologie" or (len(p) > 9 and p[9] == tipologia_filter)

            if name_match and cat_match and tipo_match:
                prodotti_filtrati.append(p)
        
        prodotti = prodotti_filtrati
        prezzi_custom = get_prezzi_listino(listino_id) # {prod_id: prezzo}
        
        self.listini_table.blockSignals(True) # Evita trigger durante il caricamento
        
        self.listini_table.setRowCount(len(prodotti))
        for i, p in enumerate(prodotti):
            # p: id, nome, cat, desc, prezzo_base...
            
            # 0. Immagine
            img_path_raw = p[6]
            img_path = self.get_valid_image_path(img_path_raw)
            img_item = QTableWidgetItem()
            if img_path:
                 pixmap = QPixmap(img_path)
                 if not pixmap.isNull():
                     # Scaliamo per entrare nella riga
                     icon = QIcon(pixmap)
                     img_item.setIcon(icon)
            self.listini_table.setItem(i, 0, img_item)

            # 1. Nome
            self.listini_table.setItem(i, 1, QTableWidgetItem(p[1]))
            self.listini_table.item(i, 1).setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable) # Readonly

            # 2. Codice
            codice = p[8] if len(p) > 8 else ""
            self.listini_table.setItem(i, 2, QTableWidgetItem(str(codice)))
            self.listini_table.item(i, 2).setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable) # Readonly

            # 3. Quantità (Visualizzazione Soglie)
            # Mostra breve riassunto es: ">10, >50"
            soglie = []
            if len(p) > 12 and p[12]: soglie.append(f">{p[12]}")
            if len(p) > 13 and p[13]: soglie.append(f">{p[13]}")
            if len(p) > 14 and p[14]: soglie.append(f">{p[14]}")
            qt_text = ", ".join(soglie) if soglie else "-"
            
            self.listini_table.setItem(i, 3, QTableWidgetItem(qt_text))
            self.listini_table.item(i, 3).setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            # 4. Prezzo Base
            self.listini_table.setItem(i, 4, QTableWidgetItem(f"{p[4]:.2f}")) 
            self.listini_table.item(i, 4).setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable) # Readonly
            
            # 5. Prezzo Listino (Editabile)
            # Se esiste un prezzo custom usa quello, altrimenti usa il prezzo base come placeholder o valore
            prezzo_val = prezzi_custom.get(p[0])
            
            display_val = ""
            if prezzo_val is not None:
                display_val = f"{prezzo_val:.2f}"
            else:
                # Opzionale: precompila con prezzo base o lascia vuoto
                display_val = f"{p[4]:.2f}"

            item_price = QTableWidgetItem(display_val)
            item_price.setData(Qt.UserRole, p[0]) # ID Prodotto
            # Colora sfondo per indicare editabilità
            item_price.setBackground(QColor("#eaf2f8"))
            self.listini_table.setItem(i, 5, item_price)

        self.listini_table.blockSignals(False)

    def crea_nuovo_listino(self):
        nome, ok = QInputDialog.getText(self, "Nuovo Listino", "Nome del nuovo listino:")
        if ok and nome:
            crea_listino(nome)
            self.load_listini_page_logic()

    def elimina_listino_corrente(self):
        if self.current_listino_id is None:
            QMessageBox.warning(self, "Attenzione", "Nessun listino selezionato.")
            return
            
        reply = QMessageBox.question(self, "Elimina Listino", "Sei sicuro di voler eliminare questo listino e tutti i prezzi associati?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            cancella_listino(self.current_listino_id)
            self.current_listino_id = None
            self.load_listini_page_logic()

    def salva_prezzi_listino_corrente(self):
        if self.current_listino_id is None: return
        rows = self.listini_table.rowCount()
        count = 0
        for i in range(rows):
            item = self.listini_table.item(i, 5) # Colonna Prezzo Listino
            if not item: continue
            prod_id = item.data(Qt.UserRole)
            try:
                txt = item.text().strip().replace(',', '.')
                val = float(txt)
                aggiorna_prezzo_listino(self.current_listino_id, prod_id, val)
                count += 1
            except:
                pass
        QMessageBox.information(self, "Salvataggio", f"Prezzi aggiornati per {count} prodotti.")

    def setup_cataloghi_page(self):
        layout = QVBoxLayout(self.page_cataloghi)
        title = QLabel("Gestione Cataloghi Salvati")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
        layout.addWidget(title)

        # Tabella Cataloghi
        self.cataloghi_table = QTableWidget()
        self.cataloghi_table.setColumnCount(3)
        self.cataloghi_table.setHorizontalHeaderLabels(["Nome Catalogo", "Data Creazione", "Note"])
        self.cataloghi_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.cataloghi_table)
        
        btns_layout = QHBoxLayout()
        btn_gen = QPushButton("🎨 Genera Nuovo Catalogo (PDF)")
        btn_gen.setStyleSheet("background-color: #27ae60; color: white; padding: 10px; font-weight: bold;")
        btn_gen.clicked.connect(self.apri_finestra_generazione_catalogo)
        
        btn_rinomina = QPushButton("✏️ Rinomina Selezionato")
        btn_rinomina.clicked.connect(self.rinomina_catalogo_ui)
        
        btn_elimina_cat = QPushButton("🗑️ Elimina Selezionato")
        btn_elimina_cat.setStyleSheet("background-color: #e74c3c; color: white; padding: 10px;")
        btn_elimina_cat.clicked.connect(self.elimina_catalogo_ui)
        
        btns_layout.addWidget(btn_gen)
        btns_layout.addWidget(btn_rinomina)
        btns_layout.addWidget(btn_elimina_cat)
        layout.addLayout(btns_layout)

    def load_cataloghi_list(self):
        cataloghi = get_cataloghi_db() # [(id, nome, data, path, note), ...]
        self.cataloghi_table.setRowCount(len(cataloghi))
        for i, c in enumerate(cataloghi):
            self.cataloghi_table.setItem(i, 0, QTableWidgetItem(c[1]))
            self.cataloghi_table.setItem(i, 1, QTableWidgetItem(c[2]))
            self.cataloghi_table.setItem(i, 2, QTableWidgetItem(c[4]))
            
            # Memorizza ID
            self.cataloghi_table.item(i, 0).setData(Qt.UserRole, c[0])

    def rinomina_catalogo_ui(self):
        row = self.cataloghi_table.currentRow()
        if row < 0: return
        
        cat_id = self.cataloghi_table.item(row, 0).data(Qt.UserRole)
        old_name = self.cataloghi_table.item(row, 0).text()
        new_name, ok = QInputDialog.getText(self, "Rinomina", "Nuovo nome:", text=old_name)
        if ok and new_name:
            rinomina_catalogo_db(cat_id, new_name)
            self.load_cataloghi_list()

    def elimina_catalogo_ui(self):
        row = self.cataloghi_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Attenzione", "Nessun catalogo selezionato.")
            return
        
        cat_id = self.cataloghi_table.item(row, 0).data(Qt.UserRole)
        cat_name = self.cataloghi_table.item(row, 0).text()

        reply = QMessageBox.question(self, "Conferma Eliminazione",
                                     f"Sei sicuro di voler eliminare il catalogo '{cat_name}'?\nQuesta operazione cancellerà anche il file PDF associato e non può essere annullata.",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                cancella_catalogo_db(cat_id)
                QMessageBox.information(self, "Successo", "Catalogo eliminato con successo.")
                self.load_cataloghi_list()
                self.update_dashboard_data()
            except Exception as e:
                QMessageBox.critical(self, "Errore", f"Impossibile eliminare il catalogo: {e}")

    def apri_finestra_generazione_catalogo(self):
        # Usiamo un dialog per le impostazioni di generazione che prima erano nella pagina
        dialog = QDialog(self)
        dialog.setWindowTitle("Impostazioni Generazione Catalogo")
        dialog.setMinimumWidth(500)
        layout = QVBoxLayout(dialog)

        # --- GESTIONE PROFILI ---
        profile_layout = QHBoxLayout()
        profile_layout.addWidget(QLabel("<b>Profilo:</b>"))
        
        self.profile_combo = QComboBox()
        self.profiles_file = os.path.join(self.base_dir, "catalog_profiles.json")
        self.profiles = self.load_profiles()
        self.profile_combo.addItems(sorted(self.profiles.keys()))
        if "Default" not in self.profiles:
            self.profile_combo.addItem("Default")
            
        self.profile_combo.currentTextChanged.connect(self.load_profile_settings)
        profile_layout.addWidget(self.profile_combo, 1)
        
        btn_save_profile = QPushButton("💾 Salva Profilo")
        btn_save_profile.clicked.connect(self.save_current_profile)
        profile_layout.addWidget(btn_save_profile)
        
        layout.addLayout(profile_layout)
        layout.addWidget(QLabel("<hr>"))
        
        # ... Qui copiamo il form di configurazione che era nella pagina ...
        form_layout = QFormLayout()
        
        self.cat_title_input = QLineEdit(self.catalog_settings['title'])
        self.cat_company_input = QLineEdit(self.catalog_settings['company'])
        self.cat_color_btn = QPushButton(self.catalog_settings['color'])
        self.cat_color_btn.setStyleSheet(f"background-color: {self.catalog_settings['color']}; color: white; font-weight: bold;")
        self.cat_color_btn.clicked.connect(self.choose_color)
        
        self.cat_show_prices_cb = QCheckBox("Mostra Prezzi nel Catalogo")
        self.cat_show_prices_cb.setChecked(self.catalog_settings['show_prices'])
        
        self.cat_layout_combo = QComboBox()
        self.cat_layout_combo.addItems(["Lista (1 colonna)", "Griglia (2 colonne)", "Griglia (3 colonne)", "Griglia (4 colonne)"])
        self.cat_layout_combo.setCurrentText(self.catalog_settings['layout'])

        # Filtro Categoria
        self.cat_category_combo = QComboBox()
        self.cat_category_combo.addItem("Tutte le categorie")
        # Popola categorie immediatamente
        prodotti_tutti = lista_prodotti()
        categorie_esistenti = sorted(list(set([p[2] for p in prodotti_tutti if p[2]])))
        self.cat_category_combo.addItems(categorie_esistenti)

        # Filtro Tipologia
        self.cat_tipologia_combo = QComboBox()
        self.cat_tipologia_combo.addItem("Tutte")
        # Popola tipologie immediatamente
        tipologie_esistenti = get_tipologie_prodotto()
        self.cat_tipologia_combo.addItems(tipologie_esistenti)

        # Scelta Font Catalogo
        self.cat_font_combo = QComboBox()
        self.cat_font_combo.addItems(["Arial", "Courier", "Times"])

        # Font TTF Esterno
        self.cat_custom_font_layout = QHBoxLayout()
        self.cat_custom_font_btn = QPushButton("Scegli TTF...")
        self.cat_custom_font_btn.clicked.connect(self.choose_custom_font)
        self.cat_custom_font_label = QLabel(os.path.basename(self.catalog_settings.get('custom_font_path', '')) or "Nessun TTF scelto")
        self.cat_custom_font_label.setStyleSheet("font-style: italic; color: gray;")
        self.cat_custom_font_layout.addWidget(self.cat_custom_font_btn)
        self.cat_custom_font_layout.addWidget(self.cat_custom_font_label, 1)

        # Filtro Immagine Catalogo
        self.cat_img_filter_combo = QComboBox()
        self.cat_img_filter_combo.addItems(["Tutti i prodotti", "Solo prodotti con immagine"])
        self.cat_img_filter_combo.setCurrentText(self.catalog_settings.get('img_filter', 'Tutti i prodotti'))

        # Filtro Visibilità Catalogo
        self.cat_visibility_combo = QComboBox()
        self.cat_visibility_combo.addItems(["Tutti i prodotti", "Solo Visibili", "Solo Nascosti"])
        self.cat_visibility_combo.setCurrentText(self.catalog_settings.get('visibility_filter', 'Solo Visibili'))

        # Nuove Opzioni (Indice e Ordinamento)
        self.cat_include_index_cb = QCheckBox("Includi pagina Indice")
        self.cat_include_index_cb.setChecked(self.catalog_settings.get('include_index', True))

        self.cat_sort_combo = QComboBox()
        self.cat_sort_combo.addItems(["Struttura Personalizzata", "Categoria (Alfabetico)", "Nome", "Codice (SKU)"])
        current_sort = self.catalog_settings.get('sort_mode', "Struttura Personalizzata")
        self.cat_sort_combo.setCurrentText(current_sort)

        # Cover Image
        cover_layout = QHBoxLayout()
        self.cat_cover_btn = QPushButton("Scegli Immagine...")
        self.cat_cover_btn.clicked.connect(self.choose_cover_image)
        self.cat_cover_label = QLabel("Nessuna immagine selezionata.")
        self.cat_cover_label.setStyleSheet("font-style: italic; color: gray;")
        cover_layout.addWidget(self.cat_cover_btn)
        cover_layout.addWidget(self.cat_cover_label, 1)

        form_layout.addRow("Titolo Catalogo:", self.cat_title_input)
        form_layout.addRow("Nome Azienda:", self.cat_company_input)
        form_layout.addRow("Colore Principale:", self.cat_color_btn)
        form_layout.addRow("", self.cat_show_prices_cb)
        
        # Checkboxes dei listini da mostrare nel PDF
        form_layout.addRow("Carattere (Font):", self.cat_font_combo)
        form_layout.addRow("Font TTF Esterno:", self.cat_custom_font_layout)

        self.prices_group = QGroupBox("Listini da mostrare nel PDF")
        prices_group_layout = QVBoxLayout(self.prices_group)
        
        self.cb_price_base = QCheckBox("Prezzo Base")
        self.cb_price_base.setChecked(self.catalog_settings.get('show_price_base', True))
        prices_group_layout.addWidget(self.cb_price_base)
        
        self.cb_price2 = QCheckBox("Listino 2")
        self.cb_price2.setChecked(self.catalog_settings.get('show_price2', False))
        prices_group_layout.addWidget(self.cb_price2)
        
        self.cb_price3 = QCheckBox("Listino 3")
        self.cb_price3.setChecked(self.catalog_settings.get('show_price3', False))
        prices_group_layout.addWidget(self.cb_price3)
        
        self.cb_price4 = QCheckBox("Listino 4")
        self.cb_price4.setChecked(self.catalog_settings.get('show_price4', False))
        prices_group_layout.addWidget(self.cb_price4)
        
        # Listini custom/extra dinamici dal database
        self.custom_price_checkboxes = {}
        listini_db = get_listini() # [(id, nome, desc), ...]
        for item in listini_db:
            nome_listino = item[1]
            cb = QCheckBox(f"Listino Extra: {nome_listino}")
            cb.setChecked(nome_listino in self.catalog_settings.get('show_custom_listini', []))
            prices_group_layout.addWidget(cb)
            self.custom_price_checkboxes[nome_listino] = cb
            
        form_layout.addRow("", self.prices_group)
        
        # Connetti la checkbox principale "Mostra Prezzi nel Catalogo" per abilitare/disabilitare il gruppo
        def toggle_prices_group(state):
            self.prices_group.setEnabled(state == Qt.Checked)
        self.cat_show_prices_cb.stateChanged.connect(toggle_prices_group)
        # Esegui trigger iniziale
        self.prices_group.setEnabled(self.cat_show_prices_cb.isChecked())

        form_layout.addRow("Layout:", self.cat_layout_combo)
        form_layout.addRow("Ordinamento:", self.cat_sort_combo)
        form_layout.addRow("", self.cat_include_index_cb)
        form_layout.addRow("Filtra Categoria:", self.cat_category_combo)
        form_layout.addRow("Filtra Tipologia:", self.cat_tipologia_combo)
        form_layout.addRow("Filtro Immagini:", self.cat_img_filter_combo)
        form_layout.addRow("Filtro Visibilità:", self.cat_visibility_combo)
        form_layout.addRow("Copertina PDF:", cover_layout)
        
        # Pulsante Micro Editor Template
        btn_template = QPushButton("🎨 Personalizza Template (Font/Stili)")
        btn_template.clicked.connect(self.open_template_editor)
        form_layout.addRow("", btn_template)
        
        layout.addLayout(form_layout)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("Genera PDF e Salva")
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        
        btn_print = QPushButton("🖨️ Stampa")
        btn_print.clicked.connect(self.stampa_catalogo)
        
        btns_layout = QHBoxLayout()
        btns_layout.addWidget(btn_print)
        btns_layout.addWidget(btn_box)
        layout.addLayout(btns_layout)
        
        # Carica impostazioni iniziali del profilo selezionato
        self.load_profile_settings(self.profile_combo.currentText())
        
        if dialog.exec_():
            # Genera
            self.esporta_pdf()
            # Ricarica lista cataloghi dopo generazione
            self.load_cataloghi_list()

    def open_template_editor(self):
        current_style = self.catalog_settings.get('style', {})
        editor = TemplateEditorDialog(current_style, self)
        if editor.exec_():
            self.catalog_settings['style'] = editor.get_style()

    def setup_impostazioni_page(self):
        # Tabbed settings: Generali, Tipologie
        layout = QVBoxLayout(self.page_impostazioni)
        title = QLabel("Impostazioni Generali")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
        layout.addWidget(title)
        
        tabs = QTabWidget()
        
        # Tab Tipologie
        tab_tipologie = QWidget()
        self.setup_tipologie_tab(tab_tipologie)
        tabs.addTab(tab_tipologie, "Gestione Tipologie")

        # Tab Struttura Catalogo (Nuovo)
        tab_cat_order = QWidget()
        self.setup_catalog_structure_tab(tab_cat_order)
        tabs.addTab(tab_cat_order, "Struttura & Pagine Extra")
        
        # Nuova Tab: Generali
        tab_generali = QWidget()
        self.setup_generali_tab(tab_generali)
        tabs.addTab(tab_generali, "Generali")
        
        layout.addWidget(tabs)

    def setup_generali_tab(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        
        # GroupBox per Auto-salvataggio
        autosave_group = QGroupBox("Salvataggio Automatico")
        autosave_layout = QFormLayout(autosave_group)
        
        self.auto_save_checkbox = QCheckBox("Abilita salvataggio automatico")
        self.auto_save_checkbox.setChecked(self.auto_save_config['enabled'])
        self.auto_save_checkbox.stateChanged.connect(self.update_auto_save_ui_to_config)
        autosave_layout.addRow("Stato:", self.auto_save_checkbox)
        
        self.auto_save_interval_spinbox = QSpinBox()
        self.auto_save_interval_spinbox.setRange(1, 60) # Da 1 a 60 minuti
        self.auto_save_interval_spinbox.setSuffix(" minuti")
        self.auto_save_interval_spinbox.setValue(self.auto_save_config['interval_minutes'])
        self.auto_save_interval_spinbox.valueChanged.connect(self.update_auto_save_ui_to_config)
        autosave_layout.addRow("Intervallo:", self.auto_save_interval_spinbox)
        
        layout.addWidget(autosave_group)

        # GroupBox per Assistenza
        support_group = QGroupBox("Supporto Tecnico")
        support_layout = QVBoxLayout(support_group)
        
        support_label = QLabel("Hai riscontrato un errore o vuoi suggerire una funzionalità? Inviaci un messaggio diretto.")
        support_label.setWordWrap(True)
        support_layout.addWidget(support_label)
        
        btn_support = QPushButton("🆘 Invia Segnalazione / Richiedi Assistenza")
        btn_support.setFixedHeight(40)
        btn_support.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold; border-radius: 5px;")
        btn_support.clicked.connect(self.apri_supporto_dialog)
        support_layout.addWidget(btn_support)
        
        layout.addWidget(support_group)
        layout.addStretch()

    def apri_supporto_dialog(self):
        """Apre un form per inviare messaggi di assistenza allo sviluppatore."""
        dialog = SupportDialog(self)
        if dialog.exec_():
            testo, include_logs = dialog.get_data()
            if not testo.strip(): return
            log_path = 'debug_log.txt' if (include_logs and os.path.exists('debug_log.txt')) else None
            try:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                send_email_logic("g.marino787@gmail.com", f"Supporto App v{APP_VERSION}", testo, log_path)
                QApplication.restoreOverrideCursor()
                QMessageBox.information(self, "Inviato", "Il tuo messaggio è stato inviato all'assistenza.")
            except Exception as e:
                QApplication.restoreOverrideCursor()
                logging.error(f"Errore invio mail supporto: {e}")
                QMessageBox.critical(self, "Errore", f"Impossibile inviare la mail: {e}")

    def update_auto_save_ui_to_config(self):
        self.auto_save_config['enabled'] = self.auto_save_checkbox.isChecked()
        self.auto_save_config['interval_minutes'] = self.auto_save_interval_spinbox.value()
        self._save_auto_save_config()
        self._configure_auto_save_timer() # Riconfigura il timer immediatamente

    def _load_auto_save_config(self):
        """Carica la configurazione di salvataggio automatico da file o usa default."""
        config_file = os.path.join(self.base_dir, "autosave_config.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {'enabled': False, 'interval_minutes': 5}

    def _save_auto_save_config(self):
        """Salva la configurazione corrente su file."""
        try:
            config_file = os.path.join(self.base_dir, "autosave_config.json")
            with open(config_file, 'w') as f:
                json.dump(self.auto_save_config, f, indent=4)
        except Exception as e:
            print(f"Errore durante il salvataggio della configurazione: {e}")

    def _configure_auto_save_timer(self):
        """Configura e avvia/ferma il timer in base alle impostazioni."""
        if not hasattr(self, 'auto_save_timer'):
            self.auto_save_timer = QTimer()
            self.auto_save_timer.timeout.connect(self.auto_save_task)
        
        self.auto_save_timer.stop()
        if self.auto_save_config.get('enabled', False):
            interval = self.auto_save_config.get('interval_minutes', 5)
            self.auto_save_timer.start(interval * 60 * 1000)

    def auto_save_task(self):
        """Esegue il backup automatico del database."""
        autosave_dir = 'autosave'
        if not os.path.exists(autosave_dir):
            os.makedirs(autosave_dir)
        
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        dest = os.path.join(autosave_dir, f'autosave_{timestamp}.db')
        try:
            shutil.copy2(DB_PATH, dest)
        except:
            pass

    def setup_tipologie_tab(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        info = QLabel("Qui puoi rinominare le tipologie esistenti o unirle.\nInserendo il nome di una tipologia già esistente come 'Nuovo Nome', i prodotti verranno spostati in quella tipologia.")
        info.setStyleSheet("color: #555;")
        layout.addWidget(info)
        
        form_layout = QFormLayout()
        
        self.gest_tipo_combo = QComboBox()
        self.gest_tipo_nuovo_nome = QLineEdit()
        self.gest_tipo_btn = QPushButton("Rinomina / Unisci")
        self.gest_tipo_btn.setStyleSheet("background-color: #3498db; color: white; padding: 8px; border-radius: 4px; font-weight: bold;")
        self.gest_tipo_btn.clicked.connect(self.rinomina_tipologia_action)
        
        form_layout.addRow("Seleziona Tipologia da Modificare:", self.gest_tipo_combo)
        form_layout.addRow("Nuovo Nome:", self.gest_tipo_nuovo_nome)
        form_layout.addRow("", self.gest_tipo_btn)
        
        layout.addLayout(form_layout)
        layout.addStretch()

    def setup_catalog_structure_tab(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        
        toolbar = QHBoxLayout()
        btn_add_img = QPushButton("➕ Aggiungi Pagina Immagine")
        btn_add_img.clicked.connect(self.add_structure_image)
        btn_add_header = QPushButton("➕ Aggiungi Pagina Intestazione")
        btn_add_header.clicked.connect(self.add_structure_header)
        btn_refresh = QPushButton("🔄 Ricarica Categorie")
        btn_refresh.clicked.connect(self.populate_catalog_structure_list)
        
        toolbar.addWidget(btn_add_img)
        toolbar.addWidget(btn_add_header)
        toolbar.addWidget(btn_refresh)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        info = QLabel("Trascina gli elementi per ordinare il catalogo. Le categorie non presenti verranno aggiunte in fondo.")
        layout.addWidget(info)

        self.structure_list = QListWidget()
        self.structure_list.setDragDropMode(QListWidget.InternalMove)
        self.structure_list.setStyleSheet("QListWidget::item { padding: 10px; border-bottom: 1px solid #eee; }")
        self.structure_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.structure_list.customContextMenuRequested.connect(self.structure_context_menu)
        layout.addWidget(self.structure_list)

        btn_save = QPushButton("💾 Salva Struttura")
        btn_save.clicked.connect(self.save_catalog_structure)
        layout.addWidget(btn_save)

    def add_structure_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Seleziona Immagine", "", "Images (*.png *.jpg)")
        if path:
            self.add_structure_item({'type': 'image', 'content': path}, f"🖼️ PAGINA IMMAGINE: {os.path.basename(path)}")

    def add_structure_header(self):
        text, ok = QInputDialog.getText(self, "Intestazione", "Testo Intestazione:")
        if ok and text:
            subtitle, ok2 = QInputDialog.getText(self, "Sottotitolo", "Sottotitolo (Opzionale):")
            self.add_structure_item({'type': 'header', 'title': text, 'subtitle': subtitle}, f"📑 SEZIONE: {text}")

    def add_structure_item(self, data, label):
        from PyQt5.QtWidgets import QListWidgetItem
        item = QListWidgetItem(label)
        item.setData(Qt.UserRole, data)
        
        # Color coding
        if data['type'] == 'category':
            item.setBackground(QColor("#e8f8f5")) # Verde chiaro
        elif data['type'] == 'image':
            item.setBackground(QColor("#fef9e7")) # Giallo chiaro
        elif data['type'] == 'header':
            item.setBackground(QColor("#ebf5fb")) # Blu chiaro
            
        self.structure_list.addItem(item)

    def structure_context_menu(self, pos):
        item = self.structure_list.itemAt(pos)
        if not item: return
        
        menu = QMenu()
        del_action = QAction("Elimina", self)
        del_action.triggered.connect(lambda: self.structure_list.takeItem(self.structure_list.row(item)))
        menu.addAction(del_action)
        menu.exec_(self.structure_list.mapToGlobal(pos))

    def load_catalog_structure(self):
        if os.path.exists(self.catalog_structure_file):
            try:
                with open(self.catalog_structure_file, 'r') as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_catalog_structure(self):
        structure = []
        for i in range(self.structure_list.count()):
            item = self.structure_list.item(i)
            data = item.data(Qt.UserRole)
            structure.append(data)
            
        try:
            with open(self.catalog_structure_file, 'w') as f:
                json.dump(structure, f, indent=4)
            QMessageBox.information(self, "Successo", "Struttura salvata.")
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Impossibile salvare l'ordine: {e}")

    def populate_catalog_structure_list(self):
        self.structure_list.clear()
        all_cats = sorted(list(set(p[2] for p in lista_prodotti() if p[2])))
        saved_structure = self.load_catalog_structure()
        
        # 1. Aggiungi elementi salvati (se ancora validi)
        processed_cats = set()
        for item_data in saved_structure:
            if item_data['type'] == 'category':
                if item_data['value'] in all_cats:
                    self.add_structure_item(item_data, f"📦 CATEGORIA: {item_data['value']}")
                    processed_cats.add(item_data['value'])
            elif item_data['type'] in ['image', 'header']:
                label = f"🖼️ PAGINA IMMAGINE" if item_data['type'] == 'image' else f"📑 SEZIONE: {item_data.get('title','')}"
                self.add_structure_item(item_data, label)

        # 2. Aggiungi le categorie rimanenti in fondo
        for cat in all_cats:
            if cat not in processed_cats:
                self.add_structure_item({'type': 'category', 'value': cat}, f"📦 CATEGORIA: {cat}")

    def choose_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            hex_color = color.name()
            self.cat_color_btn.setStyleSheet(f"background-color: {hex_color}; color: white; font-weight: bold;")
            self.cat_color_btn.setText(hex_color)
            self.catalog_settings['color'] = hex_color

    def choose_cover_image(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Seleziona Immagine Copertina', '', 'Immagini (*.png *.jpg *.jpeg)')
        if path:
            self.catalog_settings['cover_image'] = path
            self.cat_cover_label.setText(os.path.basename(path))
            self.cat_cover_label.setStyleSheet("") # Resetta lo stile

    def update_catalog_settings_from_ui(self):
        self.catalog_settings['title'] = self.cat_title_input.text()
        self.catalog_settings['company'] = self.cat_company_input.text()
        # Colore aggiornato in choose_color
        self.catalog_settings['show_prices'] = self.cat_show_prices_cb.isChecked()
        self.catalog_settings['layout'] = self.cat_layout_combo.currentText()
        self.catalog_settings['category_filter'] = self.cat_category_combo.currentText()
        self.catalog_settings['tipologia_filter'] = self.cat_tipologia_combo.currentText()
        self.catalog_settings['font'] = self.cat_font_combo.currentText()
        self.catalog_settings['img_filter'] = self.cat_img_filter_combo.currentText()
        self.catalog_settings['visibility_filter'] = self.cat_visibility_combo.currentText()
        self.catalog_settings['font'] = self.cat_font_combo.currentText()
        self.catalog_settings['img_filter'] = self.cat_img_filter_combo.currentText()
        self.catalog_settings['visibility_filter'] = self.cat_visibility_combo.currentText()
        self.catalog_settings['include_index'] = self.cat_include_index_cb.isChecked()
        self.catalog_settings['sort_mode'] = self.cat_sort_combo.currentText()
        
        # Salviamo la visibilità dei listini prezzi
        if hasattr(self, 'cb_price_base'):
            self.catalog_settings['show_price_base'] = self.cb_price_base.isChecked()
        if hasattr(self, 'cb_price2'):
            self.catalog_settings['show_price2'] = self.cb_price2.isChecked()
        if hasattr(self, 'cb_price3'):
            self.catalog_settings['show_price3'] = self.cb_price3.isChecked()
        if hasattr(self, 'cb_price4'):
            self.catalog_settings['show_price4'] = self.cb_price4.isChecked()
            
        selected_custom = []
        if hasattr(self, 'custom_price_checkboxes'):
            for name, cb in self.custom_price_checkboxes.items():
                if cb.isChecked():
                    selected_custom.append(name)
        self.catalog_settings['show_custom_listini'] = selected_custom

    def load_profiles(self):
        if os.path.exists(self.profiles_file):
            try:
                with open(self.profiles_file, 'r') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_current_profile(self):
        profile_name, ok = QInputDialog.getText(self, "Salva Profilo", "Nome del profilo (sovrascrive se esistente):", text=self.profile_combo.currentText())
        if ok and profile_name:
            self.update_catalog_settings_from_ui()
            self.profiles[profile_name] = self.catalog_settings
            try:
                with open(self.profiles_file, 'w') as f:
                    json.dump(self.profiles, f, indent=4)
                
                # Aggiorna combo se nuovo
                if self.profile_combo.findText(profile_name) == -1:
                    self.profile_combo.addItem(profile_name)
                self.profile_combo.setCurrentText(profile_name)
                QMessageBox.information(self, "Successo", f"Profilo '{profile_name}' salvato.")
            except Exception as e:
                QMessageBox.critical(self, "Errore", f"Impossibile salvare il profilo: {e}")

    def load_profile_settings(self, profile_name):
        if profile_name in self.profiles:
            settings = self.profiles[profile_name]
            self.cat_title_input.setText(settings.get('title', ''))
            self.cat_company_input.setText(settings.get('company', ''))
            self.cat_show_prices_cb.setChecked(settings.get('show_prices', True))
            self.cat_layout_combo.setCurrentText(settings.get('layout', 'Lista'))
            self.cat_category_combo.setCurrentText(settings.get('category_filter', 'Tutte le categorie'))
            self.cat_tipologia_combo.setCurrentText(settings.get('tipologia_filter', 'Tutte'))
            self.cat_font_combo.setCurrentText(settings.get('font', 'Arial'))
            if hasattr(self, 'cat_custom_font_label'):
                self.cat_custom_font_label.setText(os.path.basename(settings.get('custom_font_path', '')) or "Nessun TTF scelto")
            self.cat_img_filter_combo.setCurrentText(settings.get('img_filter', 'Tutti i prodotti'))
            self.cat_visibility_combo.setCurrentText(settings.get('visibility_filter', 'Solo Visibili'))
            self.cat_include_index_cb.setChecked(settings.get('include_index', True))
            self.cat_img_filter_combo.setCurrentText(settings.get('img_filter', 'Tutti i prodotti'))
            self.cat_visibility_combo.setCurrentText(settings.get('visibility_filter', 'Solo Visibili'))
            self.cat_sort_combo.setCurrentText(settings.get('sort_mode', 'Categoria (Personalizzato)'))
            
            # Carica la visibilità dei listini prezzi
            if hasattr(self, 'cb_price_base'):
                self.cb_price_base.setChecked(settings.get('show_price_base', True))
            if hasattr(self, 'cb_price2'):
                self.cb_price2.setChecked(settings.get('show_price2', False))
            if hasattr(self, 'cb_price3'):
                self.cb_price3.setChecked(settings.get('show_price3', False))
            if hasattr(self, 'cb_price4'):
                self.cb_price4.setChecked(settings.get('show_price4', False))
                
            if hasattr(self, 'custom_price_checkboxes'):
                show_custom = settings.get('show_custom_listini', [])
                for name, cb in self.custom_price_checkboxes.items():
                    cb.setChecked(name in show_custom)

    def rinomina_tipologia_action(self):
        vecchia = self.gest_tipo_combo.currentText()
        nuova = self.gest_tipo_nuovo_nome.text().strip()
        if vecchia and nuova:
            rinomina_tipologia(vecchia, nuova)
            QMessageBox.information(self, "Operazione Completata", f"Tipologia '{vecchia}' rinominata in '{nuova}'.")
            self.refresh_tipologie_combos()
            self.gest_tipo_nuovo_nome.clear()
        else:
            QMessageBox.warning(self, "Attenzione", "Seleziona una tipologia e inserisci un nuovo nome.")

    def check_for_updates(self):
        """Controlla la disponibilità di nuovi aggiornamenti verificando un file JSON remoto."""
        try:
            if not requests:
                QMessageBox.warning(self, "Errore Aggiornamento", 
                                    "La libreria 'requests' non è installata. Impossibile controllare gli aggiornamenti.\n"
                                    "Per abilitare questa funzionalità, installa 'requests' tramite pip.")
                return
            
            # Carica le note di rilascio dal file locale per il confronto
            local_version_data = {}
            version_json_path = os.path.join(self.base_dir, 'version.json')
            if os.path.exists(version_json_path):
                try:
                    with open(version_json_path, 'r', encoding='utf-8') as f:
                        local_version_data = json.load(f)
                except Exception as e:
                    logging.warning(f"Impossibile leggere il file version.json locale: {e}")
            
            # Usa le note locali come fallback se non ci sono note remote
            current_notes = local_version_data.get('notes', 'Nessuna nota di rilascio disponibile.')

            # Effettua la richiesta al server remoto
            response = requests.get(UPDATE_URL, timeout=5)
            data = response.json()
            latest_version = data.get('version')
            download_url = data.get('url')
            notes = data.get('notes', '')

            # Confronto versioni corretto (numerico)
            if parse_version(latest_version) > parse_version(APP_VERSION):
                # 1. Notifica l'utente tramite sistema
                if notification:
                    try:
                        notification.notify(
                            title='Aggiornamento Disponibile',
                            message=f'La versione {latest_version} è in fase di download automatico.\n\nNovità:\n{notes}',
                            app_name='Catalogo Manager Pro',
                            timeout=10
                        )
                    except Exception as notif_e:
                        logging.warning(f"Errore durante la notifica plyer: {notif_e}")
                
                # 2. Avvia il download automaticamente in background
                if download_url:
                    self.pending_update_version = latest_version
                    self.pending_update_notes = notes
                    self.start_auto_update(download_url)
                else:
                    QMessageBox.warning(self, "Errore Aggiornamento", "URL di download non trovato nel file di configurazione degli aggiornamenti.")
            else:
                QMessageBox.information(self, "Aggiornato", 
                    f"Hai già l'ultima versione ({APP_VERSION}).")
                        
        except requests.exceptions.RequestException as re:
            QMessageBox.warning(self, "Errore Aggiornamento", 
                f"Impossibile connettersi al server di aggiornamento.\nControlla la tua connessione internet.\n\nDettaglio: {re}")
        except json.JSONDecodeError:
            QMessageBox.warning(self, "Errore Aggiornamento", "Impossibile leggere il file di configurazione degli aggiornamenti (JSON non valido).")
        except Exception as e:
            QMessageBox.critical(self, "Errore Aggiornamento", 
                f"Si è verificato un errore inatteso durante il controllo aggiornamenti.\n\nDettaglio: {e}")


    def start_auto_update(self, url):
        # Crea percorso temporaneo per l'installer
        temp_dir = tempfile.gettempdir()
        installer_name = f"catalogo_update_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.exe"
        dest_path = os.path.join(temp_dir, installer_name)
        
        # UI Download
        self.update_progress = QProgressDialog("Scaricamento aggiornamento in corso...", "Annulla", 0, 100, self)
        self.update_progress.setWindowModality(Qt.WindowModal)
        self.update_progress.show()
        
        self.download_thread = DownloadThread(url, dest_path)
        self.download_thread.progress.connect(self.update_progress.setValue)
        self.download_thread.error.connect(lambda e: QMessageBox.critical(self, "Errore Download", e))
        self.download_thread.finished.connect(self.install_update)
        self.download_thread.start()

    def install_update(self, file_path):
        self.update_progress.close()

        new_version = getattr(self, 'pending_update_version', 'Sconosciuta')
        notes = getattr(self, 'pending_update_notes', 'Nessuna nota disponibile.')

        message = (f"L'aggiornamento alla versione {new_version} è stato scaricato.\n\n"
                   f"<b>Novità:</b>\n{notes}\n\n"
                   "L'applicazione verrà chiusa per avviare l'installazione. Continuare?")
        
        reply = QMessageBox.question(self, "Download Completato", message,
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                # Avvia l'installer in modo indipendente
                subprocess.Popen([file_path], shell=True)
                # Chiudi l'app corrente
                QApplication.quit()
            except Exception as e:
                QMessageBox.critical(self, "Errore Installazione", f"Impossibile avviare l'installer: {e}")

    def create_metric_card(self, icon, text, value_label):
        card = QWidget()
        card.setProperty("class", "MetricCard")
        layout = QVBoxLayout(card)
        
        icon_lbl = QLabel(icon)
        icon_lbl.setProperty("class", "MetricIcon")
        text_lbl = QLabel(text)
        text_lbl.setProperty("class", "MetricText")
        value_label.setProperty("class", "MetricValue")

        layout.addWidget(icon_lbl)
        layout.addWidget(value_label)
        layout.addWidget(text_lbl)
        return card

    def update_dashboard_data(self):
        prodotti = lista_prodotti()
        
        # Aggiorna metriche
        self.prodotti_totali_lbl.setText(str(len(prodotti)))
        prodotti_visibili = sum(1 for p in prodotti if p[5]) # p[5] è 'visibile'
        self.prodotti_visibili_lbl.setText(str(prodotti_visibili))
        tipologie_count = len(set(p[9] for p in prodotti if len(p) > 9 and p[9]))
        self.tipologie_lbl.setText(str(tipologie_count))

        # Aggiorna tabella cataloghi dashboard
        cataloghi = get_cataloghi_db() # [(id, nome, data, path, note), ...]
        # Prendi solo i primi 10
        cataloghi = cataloghi[:10]
        
        self.cataloghi_dashboard_table.setRowCount(len(cataloghi))
        for i, c in enumerate(cataloghi):
            self.cataloghi_dashboard_table.setItem(i, 0, QTableWidgetItem(c[1]))
            self.cataloghi_dashboard_table.setItem(i, 1, QTableWidgetItem(c[2]))
            # Salva ID e PATH come data
            self.cataloghi_dashboard_table.item(i, 0).setData(Qt.UserRole, c[0])
            self.cataloghi_dashboard_table.item(i, 0).setData(Qt.UserRole + 1, c[3]) # Path

    def show_about_dialog(self):
        """Mostra un dialogo con le informazioni sulla versione e il copyright."""
        # Leggi il publisher da license.txt
        publisher = "BeFree" # Default
        license_path = os.path.join(self.base_dir, 'license.txt')
        if os.path.exists(license_path):
            try:
                with open(license_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    match = re.search(r"Marino G\. concede all'utente una licenza personale", content)
                    if match:
                        publisher = "Marino G."
            except Exception as e:
                logging.warning(f"Impossibile leggere license.txt per il publisher: {e}")

        QMessageBox.about(self, "Informazioni su Catalogo Manager Pro",
                          f"<b>Catalogo Manager Pro</b><br>"
                          f"Versione: {APP_VERSION}<br>"
                          f"Sviluppato da: {publisher}<br>"
                          f"Copyright © {datetime.datetime.now().year} {publisher}. Tutti i diritti riservati.<br><br>"
                          f"Per maggiori informazioni, visita <a href='{UPDATE_URL}'>GitHub</a>.")
    def apri_catalogo_selezionato(self):
        row = self.cataloghi_dashboard_table.currentRow()
        if row >= 0:
            path = self.cataloghi_dashboard_table.item(row, 0).data(Qt.UserRole + 1)
            if path and os.path.exists(path):
                os.startfile(path)
            else:
                QMessageBox.warning(self, "Errore", "File non trovato.")

    def email_catalogo_selezionato(self):
        row = self.cataloghi_dashboard_table.currentRow()
        if row >= 0:
            path = self.cataloghi_dashboard_table.item(row, 0).data(Qt.UserRole + 1)
            if path and os.path.exists(path):
                destinatario, ok = QInputDialog.getText(self, 'Destinatario', 'Email destinatario:')
                if ok and destinatario:
                    try:
                        QApplication.setOverrideCursor(Qt.WaitCursor)
                        send_email_logic(destinatario, 'Catalogo Prodotti', 'Ecco il catalogo in allegato.', path)
                        QMessageBox.information(self, 'Inviato', 'Email inviata con successo.')
                    except Exception as e:
                        QMessageBox.critical(self, 'Errore Invio', f"Impossibile inviare l'email:\n{e}")
                    finally:
                        QApplication.restoreOverrideCursor()
            else:
                QMessageBox.warning(self, "Errore", "File non trovato.")

    def whatsapp_catalogo_selezionato(self):
        """Apre WhatsApp Web o Desktop per inviare un messaggio rapido al cliente."""
        row = self.cataloghi_dashboard_table.currentRow()
        if row >= 0:
            cat_name = self.cataloghi_dashboard_table.item(row, 0).text()
            numero, ok = QInputDialog.getText(self, 'Condividi via WhatsApp', 
                                            'Inserisci il numero del destinatario\n(con prefisso internazionale, es: 393331234567):')
            if ok and numero:
                # Pulisce il numero da eventuali spazi o caratteri speciali
                clean_number = "".join(filter(str.isdigit, numero))
                messaggio = f"Ciao! Ti invio il catalogo '{cat_name}' creato con Catalogo Manager Pro."
                url = f"https://api.whatsapp.com/send?phone={clean_number}&text={messaggio}"
                
                # Apriamo la chat
                webbrowser.open(url)
                
                # Apriamo la cartella contenente il file per facilitare il drag-and-drop
                path = self.cataloghi_dashboard_table.item(row, 0).data(Qt.UserRole + 1)
                if path and os.path.exists(path):
                    subprocess.Popen(f'explorer /select,"{os.path.abspath(path)}"')
                
                QMessageBox.information(self, 'WhatsApp', "Si è aperta la chat di WhatsApp e la cartella del file.\n\nTrascina il file PDF evidenziato direttamente nella chat per inviarlo.")

    def aggiorna_griglia_prodotti(self):
        # Selettore Tipologia Sidebar
        tipologia_filtro = getattr(self, 'current_product_group_filter', None)
        
        # Pulisci griglia
        for i in reversed(range(self.grid_layout.count())): 
            self.grid_layout.itemAt(i).widget().setParent(None)
            
        prodotti = lista_prodotti()
        filtro = self.search_input.text().lower()
        
        img_filter = self.img_filter_combo.currentText() if hasattr(self, 'img_filter_combo') else "🖼️ Tutte le Immagini"
        vis_filter = self.visibility_grid_filter_combo.currentText() if hasattr(self, 'visibility_grid_filter_combo') else "👁️ Tutte le Visibilità"
        
        # Aggiungiamo un margine interno per evitare sovrapposizioni con la sidebar
        self.grid_container.setContentsMargins(15, 15, 15, 15)
        
        row = 0
        col = 0
        # Calcolo dinamico colonne o fisso a 3 ma con spazio garantito
        max_cols = 3 
        
        # Memorizziamo i prodotti filtrati correnti per la modifica massiva
        self.prodotti_filtrati_ids = []
        
        for p in prodotti:
            # p: id, nome, categoria, desc, prezzo, visibile, img, prezzo2, codice, tipologia
            # Filtro ricerca su Nome e Categoria
            if filtro and (filtro not in p[1].lower() and (len(p) <= 2 or not p[2] or filtro not in p[2].lower())):
                continue
            
            if tipologia_filtro and (len(p) <= 9 or p[9] != tipologia_filtro):
                continue
                
            # Filtro Immagine
            has_img = bool(self.get_valid_image_path(p[6]))
            if img_filter == "✅ Solo con Immagine" and not has_img:
                continue
            if img_filter == "❌ Solo senza Immagine" and has_img:
                continue

            # Filtro Visibilità
            is_visible = bool(p[5])
            if vis_filter == "✅ Solo Visibili" and not is_visible:
                continue
            if vis_filter == "🚫 Solo Nascosti" and is_visible:
                continue

            self.prodotti_filtrati_ids.append(p[0])
            
            card = ProductCard(p, self)
            self.grid_layout.addWidget(card, row, col)
            col += 1
            if col >= max_cols:
                col = 0
                row += 1

    def refresh_categories_combo(self):
        if not hasattr(self, 'cat_category_combo'):
            return
        try:
            current_selection = self.cat_category_combo.currentText()
            self.cat_category_combo.clear()
            self.cat_category_combo.addItem("Tutte le categorie")
            prodotti = lista_prodotti()
            categorie = sorted(list(set([p[2] for p in prodotti if p[2]])))
            self.cat_category_combo.addItems(categorie)
            self.cat_category_combo.setCurrentText(current_selection)
        except RuntimeError:
            # Il widget è stato cancellato (dialogo chiuso)
            self.cat_category_combo = None

    def refresh_tipologie_combos(self):
        tipologie = get_tipologie_prodotto()

        # Aggiorna la combo del dialogo di generazione catalogo, se esiste
        try:
            if hasattr(self, 'cat_tipologia_combo') and self.cat_tipologia_combo and not self.cat_tipologia_combo.isHidden():
                current_cat_selection = self.cat_tipologia_combo.currentText()
                self.cat_tipologia_combo.clear()
                self.cat_tipologia_combo.addItem("Tutte")
                self.cat_tipologia_combo.addItems(tipologie)
                self.cat_tipologia_combo.setCurrentText(current_cat_selection)
        except RuntimeError:
            self.cat_tipologia_combo = None

        # Aggiorna anche la combo della pagina gestione tipologie
        try:
            if hasattr(self, 'gest_tipo_combo') and self.gest_tipo_combo and not self.gest_tipo_combo.isHidden():
                current_gest_selection = self.gest_tipo_combo.currentText()
                self.gest_tipo_combo.clear()
                self.gest_tipo_combo.addItems(tipologie)
                self.gest_tipo_combo.setCurrentText(current_gest_selection)
        except (RuntimeError, SystemError):
            pass

    def modifica_tipologia_massiva(self):
        if not hasattr(self, 'prodotti_filtrati_ids') or not self.prodotti_filtrati_ids:
            QMessageBox.warning(self, "Attenzione", "Nessun prodotto attualmente visualizzato/filtrato.")
            return
            
        nuova_tipologia, ok = QInputDialog.getText(self, "Assegna Tipologia", f"Inserisci la nuova tipologia da assegnare a {len(self.prodotti_filtrati_ids)} prodotti:")
        if ok and nuova_tipologia:
            aggiorna_tipologia_per_ids(self.prodotti_filtrati_ids, nuova_tipologia)
            QMessageBox.information(self, "Successo", "Tipologia aggiornata!")
            self.update_dashboard_data()
            self.aggiorna_griglia_prodotti()
            self.refresh_tipologie_combos()

    def nuovo_articolo(self):
        tipologie = get_tipologie_prodotto()
        dialog = ProdottoDialog(self, tipologie_esistenti=tipologie)
        if dialog.exec_():
            nome = dialog.nome.text()
            categoria = dialog.categoria.text()
            codice = dialog.codice.text()
            tipologia = dialog.tipologia_prodotto.currentText().strip() or 'Generico'
            descrizione = dialog.descrizione.toPlainText()
            # Gestione sicura della conversione float
            try:
                prezzo = float(dialog.prezzo.text().replace(',', '.')) if dialog.prezzo.text() else 0.0
            except ValueError:
                prezzo = 0.0
            try:
                prezzo2 = float(dialog.prezzo_secondario.text().replace(',', '.')) if dialog.prezzo_secondario.text() else 0.0
            except ValueError:
                prezzo2 = 0.0
            try: p3 = float(dialog.prezzo3.text().replace(',', '.')) if dialog.prezzo3.text() else 0.0
            except: p3 = 0.0
            try: p4 = float(dialog.prezzo4.text().replace(',', '.')) if dialog.prezzo4.text() else 0.0
            except: p4 = 0.0
            try: qt2 = int(dialog.qt2.text()) if dialog.qt2.text() else 0
            except: qt2 = 0
            try: qt3 = int(dialog.qt3.text()) if dialog.qt3.text() else 0
            except: qt3 = 0
            try: qt4 = int(dialog.qt4.text()) if dialog.qt4.text() else 0
            except: qt4 = 0
            visibile = 1 if dialog.visibile.isChecked() else 0
            immagine = dialog.immagine_path
            aggiungi_prodotto(nome, categoria, descrizione, prezzo, visibile, immagine, prezzo2, codice, tipologia, p3, p4, qt2, qt3, qt4)
            QMessageBox.information(self, 'Successo', 'Articolo aggiunto!')
            self.update_dashboard_data()
            self.aggiorna_griglia_prodotti()
            self.refresh_tipologie_combos()

    def modifica_articolo(self, p):
        # p è la tupla del prodotto (id, nome, categoria, desc, prezzo, visibile, img, prezzo2, codice, tipologia)
        tipologie = get_tipologie_prodotto()
        dialog = ProdottoDialog(self, prodotto=p, tipologie_esistenti=tipologie)
        if dialog.exec_():
            if dialog.delete_requested:
                # L'utente ha cliccato "Elimina"
                reply = QMessageBox.question(self, 'Conferma Eliminazione', 
                                             f"Sei sicuro di voler eliminare il prodotto '{p[1]}'?",
                                             QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    cancella_prodotto(p[0])
                    QMessageBox.information(self, 'Successo', 'Articolo eliminato!')
                    self.update_dashboard_data()
                    self.aggiorna_griglia_prodotti()
                    self.refresh_tipologie_combos()
            else:
                # L'utente ha cliccato "OK", procedi con la modifica
                nome = dialog.nome.text()
                categoria = dialog.categoria.text()
                codice = dialog.codice.text()
                tipologia = dialog.tipologia_prodotto.currentText().strip() or 'Generico'
                descrizione = dialog.descrizione.toPlainText()
                # Gestione sicura della conversione float
                try:
                    prezzo = float(dialog.prezzo.text().replace(',', '.')) if dialog.prezzo.text() else 0.0
                except ValueError:
                    prezzo = 0.0
                try:
                    prezzo2 = float(dialog.prezzo_secondario.text().replace(',', '.')) if dialog.prezzo_secondario.text() else 0.0
                except ValueError:
                    prezzo2 = 0.0
                try: p3 = float(dialog.prezzo3.text().replace(',', '.')) if dialog.prezzo3.text() else 0.0
                except: p3 = 0.0
                try: p4 = float(dialog.prezzo4.text().replace(',', '.')) if dialog.prezzo4.text() else 0.0
                except: p4 = 0.0
                try: qt2 = int(dialog.qt2.text()) if dialog.qt2.text() else 0
                except: qt2 = 0
                try: qt3 = int(dialog.qt3.text()) if dialog.qt3.text() else 0
                except: qt3 = 0
                try: qt4 = int(dialog.qt4.text()) if dialog.qt4.text() else 0
                except: qt4 = 0
                visibile = 1 if dialog.visibile.isChecked() else 0
                immagine = dialog.immagine_path
                modifica_prodotto(p[0], nome, categoria, descrizione, prezzo, visibile, immagine, prezzo2, codice, tipologia, p3, p4, qt2, qt3, qt4)
                QMessageBox.information(self, 'Successo', 'Articolo modificato!')
                self.update_dashboard_data()
                self.aggiorna_griglia_prodotti()
                self.refresh_tipologie_combos()

    def importa_excel(self):
        if pd is None:
            QMessageBox.critical(self, "Errore Libreria", 
                                 "La libreria 'pandas' non è installata. Esegui: pip install pandas openpyxl")
            return
        file_path, _ = QFileDialog.getOpenFileName(self, 'Seleziona file Excel', '', 'Excel (*.xlsx *.xls)')
        if file_path:
            self._process_import(file_path, read_excel_df)

    def importa_danea(self):
        """Gestisce l'importazione specifica per il formato XML di Danea."""
        file_path, _ = QFileDialog.getOpenFileName(self, 'Seleziona file XML Danea', '', 'Danea XML (*.xml)')
        if file_path:
            self._process_import(file_path, read_danea_xml)

    def importa_access(self):
        if pyodbc is None:
            QMessageBox.critical(self, "Errore Libreria", 
                                 "La libreria 'pyodbc' non è installata.\n"
                                 "Esegui: pip install pyodbc")
            return
        file_path, _ = QFileDialog.getOpenFileName(self, 'Seleziona file Access', '', 'Access Database (*.mdb *.accdb)')
        if file_path:
            try:
                tables = get_access_tables(file_path)
                if not tables:
                    QMessageBox.warning(self, "Errore", "Nessuna tabella trovata.")
                    return
                
                table_name = tables[0]
                if len(tables) > 1:
                    item, ok = QInputDialog.getItem(self, "Seleziona Tabella", "Tabelle trovate:", tables, 0, False)
                    if ok and item:
                        table_name = item
                    else:
                        return
                
                # Usa lambda per passare entrambi gli argomenti a read_access_table
                self._process_import(file_path, lambda f: read_access_table(f, table_name))
            except Exception as e:
                QMessageBox.critical(self, "Errore Access", str(e))

    def _process_import(self, file_path, read_func):
            try:
                df_original = read_func(file_path)
                excel_columns = list(df_original.columns)

                # 2. Mostra il dialogo di mappatura per far scegliere all'utente
                mapping_dialog = ColumnMappingDialog(excel_columns, self)
                if mapping_dialog.exec_():
                    mapping = mapping_dialog.get_column_mappings()
                    price_list_map = mapping_dialog.get_price_list_mappings()

                    # Se non è stata mappata nessuna colonna, esci
                    if not mapping:
                        QMessageBox.warning(self, "Mappatura Vuota", "Nessuna colonna è stata mappata. Importazione annullata.")
                        return

                    # 3. Crea un nuovo DataFrame basato sulla mappatura
                    # Costruisci un DataFrame per la preview che include sia i campi principali mappati
                    # sia le colonne dei listini extra con i loro nomi originali.
                    df_for_preview = pd.DataFrame()
                    
                    # Aggiungi le colonne mappate ai campi principali
                    for target_col, source_col_excel in mapping.items():
                        if source_col_excel in df_original.columns:
                            df_for_preview[target_col] = df_original[source_col_excel]
                    
                    # Aggiungi le colonne dei listini extra (cruciale per importa_dataframe_nel_db)
                    for col_name_excel in price_list_map.keys():
                        if col_name_excel not in df_for_preview.columns and col_name_excel in df_original.columns:
                            df_for_preview[col_name_excel] = df_original[col_name_excel]

                    # 4. Mostra il dialogo di anteprima con il DataFrame mappato e validabile
                    preview_dialog = PreviewDialog(df_for_preview, self)
                    if preview_dialog.exec_():
                        # L'utente ha confermato, ottieni i dati (potenzialmente modificati)
                        edited_df = preview_dialog.get_dataframe()

                        # Chiedi all'utente il nome del gruppo/tipologia
                        tipologie_esistenti = get_tipologie_prodotto()
                        input_dialog = QInputDialog(self)
                        input_dialog.setLabelText("Assegna i prodotti importati a un gruppo (nuovo o esistente):")
                        input_dialog.setWindowTitle("Assegna Gruppo Prodotti")
                        input_dialog.setComboBoxItems([""] + tipologie_esistenti)
                        input_dialog.setComboBoxEditable(True)
                        
                        if input_dialog.exec_():
                            nome_gruppo = input_dialog.textValue().strip()
                            if not nome_gruppo:
                                QMessageBox.warning(self, "Nome non valido", "Il nome del gruppo non può essere vuoto. Importazione annullata.")
                                return
                            # Sovrascrivi la colonna 'tipologia_prodotto' con il nome del gruppo scelto
                            edited_df['tipologia_prodotto'] = nome_gruppo 
                        else:
                            return # L'utente ha annullato
                        
                        # 5. Chiedi la cartella immagini (opzionale)
                        images_folder = None
                        reply = QMessageBox.question(self, 'Caricamento Immagini', 
                                                     "Vuoi selezionare una cartella contenente le immagini?\n(I file devono chiamarsi come il CODICE SKU del prodotto)",
                                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                        if reply == QMessageBox.Yes:
                            images_folder = QFileDialog.getExistingDirectory(self, "Seleziona Cartella Immagini")
                        
                        # 6. Esegui l'importazione effettiva con barra di avanzamento
                        total_rows = len(edited_df)
                        progress = QProgressDialog("Importazione dati in corso...", "Annulla", 0, total_rows, self)
                        progress.setWindowModality(Qt.WindowModal)
                        progress.show()

                        def progress_callback(current, total):
                            progress.setValue(current)
                            QApplication.processEvents() # Necessario per aggiornare l'UI e rilevare il "cancel"
                            if progress.wasCanceled():
                                raise InterruptedError("Importazione annullata dall'utente.")

                        try:
                            importa_dataframe_nel_db(edited_df, images_folder, progress_callback, price_list_map)
                            progress.setValue(total_rows) # Assicura che la barra arrivi al 100%
                            QMessageBox.information(self, 'Importazione', 'Importazione completata con successo!')
                            self.update_dashboard_data()
                            self.aggiorna_griglia_prodotti()
                        except InterruptedError as e:
                            QMessageBox.warning(self, 'Operazione Annullata', str(e))

            except Exception as e:
                QMessageBox.critical(self, 'Errore Importazione', f"Errore durante la lettura del file o l'importazione:\n{str(e)}")

    def esporta_pdf(self):
        if FPDF is object:
            QMessageBox.critical(self, "Errore Libreria", 
                                 "La libreria 'fpdf' non è installata.\n"
                                 "Esegui: pip install fpdf")
            return
        self.update_catalog_settings_from_ui() # Salva le impostazioni correnti
        # Aggiungi l'ordine personalizzato delle categorie alla configurazione
        self.catalog_settings['catalog_structure'] = self.load_catalog_structure()

        file_path, _ = QFileDialog.getSaveFileName(self, 'Salva PDF', 'catalogo.pdf', 'PDF (*.pdf)')
        if file_path:
            try:
                # Creiamo un dialogo di progresso per la generazione del PDF
                progress = QProgressDialog("Generazione Catalogo PDF in corso...", "Annulla", 0, 100, self)
                progress.setWindowModality(Qt.WindowModal)
                progress.show()
                QApplication.processEvents()
                
                # Passiamo il callback per aggiornare la barra
                def pdf_progress(curr, total):
                    progress.setValue(int((curr / total) * 100))
                    QApplication.processEvents()

                esporta_catalogo_pdf(file_path, self.catalog_settings)
                # Salva record nel DB
                salva_catalogo_db(self.catalog_settings['title'], file_path, f"Generato il {datetime.datetime.now()}")
                QMessageBox.information(self, 'Esportazione', 'Catalogo esportato in PDF!')
            except PermissionError:
                QMessageBox.critical(self, 'Errore Permessi', 'Impossibile salvare il file PDF. Assicurati che non sia già aperto in un altro programma.')
            except Exception as e:
                import traceback
                traceback.print_exc()
                QMessageBox.critical(self, 'Errore', f"Errore durante la generazione del catalogo:\n{str(e)}")

    def stampa_catalogo(self):
        self.update_catalog_settings_from_ui()
        
        # Configura Stampante
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() != QDialog.Accepted:
            return

        # Filtra Dati
        category_filter = self.catalog_settings.get('category_filter', 'Tutte le categorie')
        tipologia_filter = self.catalog_settings.get('tipologia_filter', 'Tutte')
        img_filter = self.catalog_settings.get('img_filter', 'Tutti i prodotti')
        vis_filter = self.catalog_settings.get('visibility_filter', 'Solo Visibili')
        prodotti = lista_prodotti()
        
        filtered = []
        for p in prodotti:
            # Filtro Visibilità
            is_visible = bool(p[5])
            if vis_filter == "Solo Visibili" and not is_visible: continue
            if vis_filter == "Solo Nascosti" and is_visible: continue

            # p[2] = categoria
            if category_filter != 'Tutte le categorie' and p[2] != category_filter: continue
            # p[9] = tipologia
            tipo = p[9] if len(p) > 9 else 'Generico'
            if tipologia_filter != 'Tutte' and tipo != tipologia_filter: continue
            # Filtro immagine
            if img_filter == "Solo prodotti con immagine" and not self.get_valid_image_path(p[6]):
                continue
            filtered.append(p)
        # Genera HTML per la stampa
        html = "<html><head><style>"
        html += "body { font-family: Arial, sans-serif; }"
        html += f"h1 {{ text-align: center; color: {self.catalog_settings.get('color', '#000')}; }}"
        html += "td { padding: 10px; }"
        html += "</style></head><body>"
        
        html += f"<h1>{self.catalog_settings.get('title', 'Catalogo')}</h1>"
        html += f"<h3 style='text-align:center; color:gray;'>{self.catalog_settings.get('company', '')}</h3><hr><br>"

        layout_mode = self.catalog_settings.get('layout', 'Lista')
        
        if "Griglia" in layout_mode:
            cols = 3 if "3" in layout_mode else 2
            html += "<table width='100%' cellspacing='0' cellpadding='5'>"
            for i, p in enumerate(filtered):
                if i % cols == 0: html += "<tr>"
                
                img_path = self.get_valid_image_path(p[6])
                img_tag = f"<img src='{img_path}' width='150' height='150'>" if img_path else ""
                
                html += f"<td width='{100/cols}%' align='center' style='border:1px solid #ddd;'>"
                html += f"<div>{img_tag}</div>"
                html += f"<h3>{p[1]}</h3>"
                if p[8]: html += f"<p style='color:gray; font-size:small;'>{p[8]}</p>"
                
                # Gestione Tabella Prezzi a Scaglioni
                tiers = [(1, p[4])]
                # Indici basati su lista_prodotti: 7=p2, 10=p3, 11=p4, 12=qt2, 13=qt3, 14=qt4
                if len(p) > 12 and p[7] > 0 and p[12] > 0: tiers.append((p[12], p[7]))
                if len(p) > 13 and p[10] > 0 and p[13] > 0: tiers.append((p[13], p[10]))
                if len(p) > 14 and p[11] > 0 and p[14] > 0: tiers.append((p[14], p[11]))

                if self.catalog_settings['show_prices']:
                    if len(tiers) > 1:
                        html += "<table width='100%' style='border-collapse:collapse; font-size:10px; margin-top:5px;'>"
                        html += "<tr style='background-color:#eee;'><th>Q.tà</th><th>Prezzo</th></tr>"
                        for qty, prc in tiers:
                            html += f"<tr><td align='center' style='border:1px solid #ddd;'>{qty}+</td><td align='center' style='border:1px solid #ddd;'>€ {prc:.2f}</td></tr>"
                        html += "</table>"
                    else:
                        html += f"<h4 style='color:green;'>€ {p[4]:.2f}</h4>"
                
                html += "</td>"
                
                if i % cols == (cols - 1): html += "</tr>"
            if len(filtered) % cols != 0: html += "</tr>"
            html += "</table>"
        else:
            # Layout Lista
            html += "<table width='100%'>"
            for p in filtered:
                img_path = self.get_valid_image_path(p[6])
                img_tag = f"<img src='{img_path}' width='100' height='100'>" if img_path else ""
                html += f"<tr><td width='120'>{img_tag}</td>"
                html += f"<td><h3>{p[1]}</h3><p><i>{p[2]}</i></p><p>{p[3]}</p></td>"
                
                if self.catalog_settings['show_prices']:
                    html += "<td width='140' align='right'>"
                    
                    tiers = [(1, p[4])]
                    if len(p) > 12 and p[7] > 0 and p[12] > 0: tiers.append((p[12], p[7]))
                    if len(p) > 13 and p[10] > 0 and p[13] > 0: tiers.append((p[13], p[10]))
                    if len(p) > 14 and p[11] > 0 and p[14] > 0: tiers.append((p[14], p[11]))

                    if len(tiers) > 1:
                        html += "<table style='border-collapse:collapse; font-size:11px;'>"
                        html += "<tr style='background-color:#eee;'><th>Q.tà</th><th>Prezzo</th></tr>"
                        for qty, prc in tiers:
                            html += f"<tr><td align='center' style='border:1px solid #ddd; padding:2px;'>{qty}+</td><td align='right' style='border:1px solid #ddd; padding:2px;'>€ {prc:.2f}</td></tr>"
                        html += "</table>"
                    else:
                        html += f"<h3 style='color:green;'>€ {p[4]:.2f}</h3>"
                    
                    html += "</td>"

                html += "</tr><tr><td colspan='3'><hr></td></tr>"
            html += "</table>"

        html += "</body></html>"
        
        doc = QTextDocument()
        doc.setHtml(html)
        doc.print_(printer)
        QMessageBox.information(self, "Stampa", "Documento inviato alla stampante.")

    def invia_email(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Seleziona PDF da inviare', '', 'PDF (*.pdf)')
        if file_path:
            destinatario, ok = QInputDialog.getText(self, 'Destinatario', 'Email destinatario:')
            if ok and destinatario:
                send_email_logic(destinatario, 'Catalogo Prodotti', 'In allegato il catalogo.', file_path)
                QMessageBox.information(self, 'Invio', 'Email inviata!')

if __name__ == '__main__':
    logging.info(f"--- Tentativo di avvio applicazione v{APP_VERSION} ---")
    app = QApplication(sys.argv)
    try:
        window = CatalogoMainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.exception("Errore fatale all'avvio:")
        sys.exit(1)
