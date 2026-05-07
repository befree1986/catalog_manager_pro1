import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout
from main import ProductCard

# Mock della finestra principale per testare i componenti isolati
class MockWindow:
    def get_valid_image_path(self, path): return ""
    def update_dashboard_data(self): pass
    def aggiorna_griglia_prodotti(self): pass
    def modifica_articolo(self, p): print(f"Modifica cliccata per: {p[1]}")

def preview_product_card():
    app = QApplication(sys.argv)
    
    # Dati di test (Simuliamo una riga del database)
    # id, nome, categoria, desc, prezzo, visibile, img, prezzo2, codice, tipologia
    p = [1, "Prodotto Anteprima", "Elettronica", 
         "Questa è una descrizione di prova per vedere come viene renderizzata la card.", 
         129.50, 1, "", 110.00, "SKU-PREVIEW-01", "Smartphone"]
    
    window = QWidget()
    window.setWindowTitle("Anteprima Card Prodotto")
    window.setStyleSheet("background-color: #f0f0f0;")
    layout = QVBoxLayout(window)
    
    # Creiamo la card usando il mock come parent
    card = ProductCard(p, MockWindow())
    layout.addWidget(card)
    
    window.resize(300, 400)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    preview_product_card()