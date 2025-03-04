from flask import Flask, request, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__)

# Directorios para archivos subidos y procesados
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["PROCESSED_FOLDER"] = PROCESSED_FOLDER

# Función para procesar la columna 'PARCEL #'
def process_parcel(parcel):
    if isinstance(parcel, str):
        parcel = parcel.strip()
        if any(char.isalpha() for char in parcel):
            return ""
        if "-" in parcel:
            segments = parcel.split("-")
            segments = [seg.ljust(3, "0")[:3] for seg in segments]
            result = "-".join(segments)
            while len(result.replace("-", "")) < 12:
                result += "0"
            return result[:15]  
    elif not isinstance(parcel, (int, float)) or pd.isna(parcel) or parcel == 0:
        return ""
    else:
        parcel_str = f"{int(parcel):012d}"
        return f"{parcel_str[:3]}-{parcel_str[3:6]}-{parcel_str[6:9]}-{parcel_str[9:12]}"
    return ""

# Ruta para subir y procesar el archivo
@app.route("/upload", methods=["POST"])
def upload_file():
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400
    
    file = request.files["archivo"]
    
    if file.filename == "":
        return jsonify({"error": "Nombre de archivo inválido"}), 400

    filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(filepath)
    
    try:
        df = pd.read_excel(filepath, engine="openpyxl")

        if "PARCEL #" not in df.columns:
            return jsonify({"error": "El archivo no tiene la columna 'PARCEL #'"}), 400

        df["FIXED PARCEL"] = df["PARCEL #"].apply(process_parcel)

        processed_filepath = os.path.join(app.config["PROCESSED_FOLDER"], f"processed_{file.filename}")
        df.to_excel(processed_filepath, index=False, engine="openpyxl")

        return send_file(processed_filepath, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Iniciar el servidor Flask
if __name__ == "__main__":
    app.run(debug=True)
