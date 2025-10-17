from flask import Flask, render_template, request, jsonify
import pandas as pd
import json
import os
from datetime import datetime

app = Flask(__name__)

# === Cargar listas ===
with open("listas.json", "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

EXCEL_FILE = "Registro Muestreo Recepcion Poscosecha.xlsx"

def save_to_excel_dict(datos_dict):
    """Guarda el diccionario como nueva fila en el Excel (crea archivo si no existe)."""
    df_nuevo = pd.DataFrame([datos_dict])
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE, engine="openpyxl")
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
    df_final.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

@app.route("/")
def index():
    return render_template("index.html", config=CONFIG)

@app.route("/enviar", methods=["POST"])
def enviar():
    # Acepta form-data normal (desde fetch con FormData también)
    f = request.form

    ahora = datetime.now()
    semana = ahora.isocalendar()[1]

    # Construir dict con campos (usar claves en mayúsculas o las que uses)
    datos = {
        "FECHA_REGISTRO": ahora.strftime("%Y-%m-%d %H:%M:%S"),
        "SEMANA": semana,
        "RESPONSABLE": f.get("responsable", ""),
        "FINCA": f.get("finca", ""),
        "BLOQUE": f.get("bloque", "") if f.get("finca") == "Malchigui" else "",
        "CULTIVO": f.get("cultivo", ""),
        "VARIEDAD": f.get("variedad", ""),
        "TOTAL_TALLOS": f.get("tallos_total", ""),
        "PUNTO_CORTE_NORMAL": f.get("pc_normal", ""),
        "PUNTO_CORTE_ABIERTO": f.get("pc_abierto", ""),
        "PUNTO_CORTE_CERRADO": f.get("pc_cerrado", ""),
        "PUNTO_CORTE_PODA": f.get("pc_poda", ""),
        "DESCABEZADOS": f.get("descabezados", ""),
        "ROTOS": f.get("rotos", ""),
        "DESHIDRATADOS": f.get("deshidratados", ""),
        "SENESCENTES": f.get("senescentes", ""),
        "PINK": f.get("pink", ""),
        "ALTERNARIA": f.get("alt", ""),
        "BOTRYTIS": f.get("bot", ""),
        "POLVOSO": f.get("polvoso", ""),
        "ROYA": f.get("roya", ""),
        "VELLOSO": f.get("velloso", ""),
        "OTROS_ENFERMEDADES": f.get("otros_enf", ""),
        "TRIPS": f.get("trips", ""),
        "AFIDOS": f.get("afidos", ""),
        "CHINCHES": f.get("chinches", ""),
        "ABEJAS": f.get("abejas", ""),
        "ACAROS": f.get("acaros", ""),
        "MINADOR": f.get("minador", ""),
        "BABOSAS": f.get("babosa", ""),
        "CARACOLES": f.get("caracoles", ""),
        "MOSCA_BLANCA": f.get("moscablanca", ""),
        "COCHINILLAS": f.get("cochinillas", ""),
        "MARIPOSAS": f.get("mariposas", ""),
        "ESCARABAJOS": f.get("escarabajos", "")
    }

    try:
        save_to_excel_dict(datos)
    except Exception as e:
        # si falla, devolvemos error JSON con mensaje para que JS muestre
        return jsonify({"ok": False, "error": str(e)}), 500

    # Respuesta JSON de éxito
    return jsonify({"ok": True})
    
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

