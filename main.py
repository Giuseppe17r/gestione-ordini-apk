from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
import pandas as pd
import os

class OrdineApp(BoxLayout):
    def _init_(self, **kwargs):
        super()._init_(orientation='vertical', **kwargs)
        btn = Button(text="Genera ordini fornitore")
        btn.bind(on_press=self.genera_ordini)
        self.add_widget(btn)

    def genera_ordini(self, instance):
        cartella = "/storage/emulated/0/ORDINI"
        file_listino = os.path.join(cartella, "FILE DEFINITIVO CONFRONTO LISTINI.xlsx")
        file_ordine = os.path.join(cartella, "ordine.xlsx")

        fornitori = {
            "MARR": {"prezzo": "€ F1", "nome": "MARR (F1)"},
            "DAC": {"prezzo": "€ F2", "nome": "DAC (F2)"},
            "LA CERVESE": {"prezzo": "€ F3", "nome": "LA CERVESE (F3)"},
            "EUROCATERING": {"prezzo": "€ F4", "nome": "EUROCATERING (F4)"},
            "RICCI": {"prezzo": "€ F5", "nome": "RICCI (F5)"}
        }

        df_listino = pd.read_excel(file_listino, engine="openpyxl")
        df_ordine = pd.read_excel(file_ordine, engine="openpyxl")
        ordini = {f: [] for f in fornitori}

        for _, riga in df_ordine.iterrows():
            nome_articolo = str(riga["Articolo"]).strip().lower()
            quantità = riga["Quantità"]
            filtro = df_listino["Nome Generico Prodotto"].str.strip().str.lower() == nome_articolo
            prodotto = df_listino[filtro]

            if prodotto.empty:
                continue

            prodotto = prodotto.iloc[0]
            prezzi = {f: prodotto[v["prezzo"]] for f, v in fornitori.items() if pd.notna(prodotto[v["prezzo"]])}
            if not prezzi:
                continue

            fornitore_best = min(prezzi, key=prezzi.get)
            prezzo_unitario = prezzi[fornitore_best]
            nome_specifico = prodotto[fornitori[fornitore_best]["nome"]]

            ordini[fornitore_best].append({
                "Cod. Interno": prodotto["Cod. Interno"],
                "Nome Generico": prodotto["Nome Generico Prodotto"],
                "Nome Specifico": nome_specifico,
                "Quantità": quantità,
                "Prezzo unitario": prezzo_unitario,
                "Totale": prezzo_unitario * quantità
            })

        for f, righe in ordini.items():
            if righe:
                df_out = pd.DataFrame(righe)
                path_output = os.path.join(cartella, f"ordine_{f}.xlsx")
                df_out.to_excel(path_output, index=False, engine="openpyxl")

class GestioneOrdiniApp(App):
    def build(self):
        return OrdineApp()

if __name__ == '__main__':
    GestioneOrdiniApp().run()