# c:\Users\g_mar\Documents\lavoro\catalogManager1\prodotto_dialog.py
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QFormLayout, QLineEdit, 
                             QComboBox, QTextEdit, QCheckBox, QPushButton, 
                             QHBoxLayout, QFileDialog, QLabel, QDialogButtonBox, 
                             QGroupBox, QGridLayout)
from PyQt5.QtCore import Qt
import os

class ProdottoDialog(QDialog):
    def __init__(self, parent=None, prodotto=None, tipologie_esistenti=None):
        super().__init__(parent)
        self.setWindowTitle("Dettaglio Prodotto")
        self.setMinimumWidth(600)
        self.prodotto = prodotto
        self.delete_requested = False
        self.immagine_path = ""
        
        layout = QVBoxLayout(self)
        
        # Form Layout
        form_layout = QFormLayout()
        
        self.nome = QLineEdit()
        self.codice = QLineEdit()
        self.categoria = QLineEdit() 
        
        self.tipologia_prodotto = QComboBox()
        self.tipologia_prodotto.setEditable(True)
        if tipologie_esistenti:
            self.tipologia_prodotto.addItems(tipologie_esistenti)
            
        self.descrizione = QTextEdit()
        self.descrizione.setMaximumHeight(100)
        
        self.prezzo = QLineEdit()
        self.prezzo.setPlaceholderText("0.00")
        
        self.visibile = QCheckBox("Visibile nel catalogo")
        self.visibile.setChecked(True)
        
        form_layout.addRow("Nome Prodotto:", self.nome)
        form_layout.addRow("Codice (SKU):", self.codice)
        form_layout.addRow("Categoria (Excel):", self.categoria)
        form_layout.addRow("Gruppo/Tipologia:", self.tipologia_prodotto)
        form_layout.addRow("Descrizione:", self.descrizione)
        form_layout.addRow("Prezzo Base (€):", self.prezzo)
        form_layout.addRow("", self.visibile)
        
        layout.addLayout(form_layout)
        
        # Prezzi a scaglioni (Listini)
        prices_group = QGroupBox("Prezzi a Scaglioni / Listini Extra")
        prices_layout = QGridLayout()
        
        self.qt2 = QLineEdit()
        self.qt2.setPlaceholderText("Q.tà")
        self.prezzo_secondario = QLineEdit() # Prezzo 2
        self.prezzo_secondario.setPlaceholderText("€")
        
        self.qt3 = QLineEdit()
        self.qt3.setPlaceholderText("Q.tà")
        self.prezzo3 = QLineEdit()
        self.prezzo3.setPlaceholderText("€")
        
        self.qt4 = QLineEdit()
        self.qt4.setPlaceholderText("Q.tà")
        self.prezzo4 = QLineEdit()
        self.prezzo4.setPlaceholderText("€")
        
        prices_layout.addWidget(QLabel("Scaglione 1 (Base):"), 0, 0)
        prices_layout.addWidget(QLabel("1+"), 0, 1)
        prices_layout.addWidget(QLabel("(Vedi sopra)"), 0, 2)
        
        prices_layout.addWidget(QLabel("Scaglione 2:"), 1, 0)
        prices_layout.addWidget(self.qt2, 1, 1)
        prices_layout.addWidget(self.prezzo_secondario, 1, 2)
        
        prices_layout.addWidget(QLabel("Scaglione 3:"), 2, 0)
        prices_layout.addWidget(self.qt3, 2, 1)
        prices_layout.addWidget(self.prezzo3, 2, 2)
        
        prices_layout.addWidget(QLabel("Scaglione 4:"), 3, 0)
        prices_layout.addWidget(self.qt4, 3, 1)
        prices_layout.addWidget(self.prezzo4, 3, 2)
        
        prices_group.setLayout(prices_layout)
        layout.addWidget(prices_group)
        
        # Immagine
        img_layout = QHBoxLayout()
        self.btn_img = QPushButton("Scegli Immagine")
        self.btn_img.clicked.connect(self.scegli_immagine)
        self.lbl_img = QLabel("Nessuna immagine")
        img_layout.addWidget(self.btn_img)
        img_layout.addWidget(self.lbl_img)
        layout.addLayout(img_layout)
        
        # Bottoni
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        
        if prodotto:
            # Add Delete button
            btn_delete = btn_box.addButton("Elimina Prodotto", QDialogButtonBox.DestructiveRole)
            btn_delete.clicked.connect(self.request_delete)
        
        layout.addWidget(btn_box)
        
        # Popola campi se modifica
        if prodotto:
            self.populate_fields()

    def populate_fields(self):
        p = self.prodotto
        try:
            self.nome.setText(str(p[1] or ""))
            self.categoria.setText(str(p[2] or ""))
            self.descrizione.setPlainText(str(p[3] or ""))
            self.prezzo.setText(str(p[4] or 0))
            self.visibile.setChecked(bool(p[5]))
            self.immagine_path = str(p[6] or "")
            if self.immagine_path:
                self.lbl_img.setText(os.path.basename(self.immagine_path))
                
            if len(p) > 7: self.prezzo_secondario.setText(str(p[7] or 0))
            if len(p) > 8: self.codice.setText(str(p[8] or ""))
            if len(p) > 9: self.tipologia_prodotto.setCurrentText(str(p[9] or "Generico"))
            
            if len(p) > 10: self.prezzo3.setText(str(p[10] or 0))
            if len(p) > 11: self.prezzo4.setText(str(p[11] or 0))
            if len(p) > 12: self.qt2.setText(str(p[12] or 0))
            if len(p) > 13: self.qt3.setText(str(p[13] or 0))
            if len(p) > 14: self.qt4.setText(str(p[14] or 0))

        except Exception as e:
            print(f"Errore popolamento dati: {e}")

    def scegli_immagine(self):
        path, _ = QFileDialog.getOpenFileName(self, "Seleziona Immagine", "", "Immagini (*.png *.jpg *.jpeg)")
        if path:
            self.immagine_path = path
            self.lbl_img.setText(os.path.basename(path))

    def request_delete(self):
        self.delete_requested = True
        self.accept()