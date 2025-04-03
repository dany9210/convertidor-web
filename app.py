from flask import Flask, render_template, request, send_file
import os
import uuid
from convertidores.convertir_scr import convertir_scr
from convertidores.convertir_scc import convertir_scc
from convertidores.convertir_reporte1 import convertir_reporte1

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        archivo = request.files["archivo"]
        tipo_reporte = request.form["tipo_reporte"]

        if archivo.filename == "":
            return "❌ No seleccionaste ningún archivo."

        # Guardar archivo subido
        nombre_unico = str(uuid.uuid4())
        ruta_txt = os.path.join(UPLOAD_FOLDER, f"{nombre_unico}.txt")
        ruta_excel = os.path.join(UPLOAD_FOLDER, f"{nombre_unico}.xlsx")
        archivo.save(ruta_txt)

        try:
            if tipo_reporte == "SCR":
                convertir_scr(ruta_txt, ruta_excel)
            elif tipo_reporte == "SCC":
                convertir_scc(ruta_txt, ruta_excel)
            elif tipo_reporte == "REPORTE1":
                convertir_reporte1(ruta_txt, ruta_excel)
            else:
                return "❌ Tipo de reporte desconocido."

            return send_file(ruta_excel, as_attachment=True)

        except Exception as e:
            return f"❌ Error al procesar el archivo: {e}"

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)