from flask import Flask, render_template
import subprocess
import os
import pandas as pd

# Nome del tuo file con il codice ORIGINALE
SCRIPT_NAME = "analisi_online_retail.py"   # <-- cambia se il file ha un altro nome
EXCEL_NAME = "Analisi_Online_Retail.xlsx"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(BASE_DIR, SCRIPT_NAME)
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_NAME)

# static_folder='.' e static_url_path='' perché i PNG sono nella stessa cartella
app = Flask(__name__, static_folder='.', static_url_path='')

def run_original_script_if_needed():
    """
    Se l'Excel dell'analisi non esiste, lancia il tuo script ORIGINALE.
    NON modifichiamo il codice originale: lo eseguiamo solo.
    """
    if not os.path.exists(EXCEL_PATH):
        subprocess.run(["python", SCRIPT_PATH], check=True)


@app.route("/")
def index():
    # 1) Assicuriamoci che l'analisi sia stata eseguita
    run_original_script_if_needed()

    # 2) Leggiamo i risultati dall'Excel creato dal tuo codice
    xls = pd.ExcelFile(EXCEL_PATH)

    # Vendite settimanali
    df_week = pd.read_excel(xls, "Vendite_settimanali", index_col=0)
    weekly_preview = [
        (str(idx), float(val))
        for idx, val in df_week["TotalSales"].tail(10).items()
    ]

    # Vendite mensili
    df_month = pd.read_excel(xls, "Vendite_mensili", index_col=0)
    monthly_preview = [
        (str(idx), float(val))
        for idx, val in df_month["TotalSales"].tail(10).items()
    ]

    # Elasticità domanda
    df_el = pd.read_excel(xls, "Elasticita_domanda", index_col=0)
    elasticity_rows = []
    for code, row in df_el.iterrows():
        elasticity_rows.append({
            "StockCode": str(code),
            "Elasticita_media": None if pd.isna(row["Elasticità media"]) else float(row["Elasticità media"]),
            "Valutazione": row["Valutazione"],
        })

    stats = {
        "weekly_preview": weekly_preview,
        "monthly_preview": monthly_preview,
        "elasticity": elasticity_rows,
    }

    return render_template("index.html", stats=stats)


if __name__ == "__main__":
    app.run(debug=True)
