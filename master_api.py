from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from Master_core import gerar_certificados_mp
from pathlib import Path
import os
import tempfile
from io import BytesIO
import zipfile

app = Flask(__name__)

# ✅ Libera acesso do GitHub Pages (ajuste o domínio do seu site abaixo)
CORS(app, resources={r"/*": {"origins": [
    "https://vytorvilar.github.io",
    "https://vytorvilar.github.io/sst",
    "http://localhost:5500"
]}})

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "online",
        "mensagem": "API de geração de certificados está ativa 💥"
    }), 200


@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        csv_file = request.files.get("csv")
        modelos_files = request.files.getlist("modelos")
        tipo = request.form.get("tipo", "PDF")

        if not csv_file or not modelos_files:
            return jsonify({
                "status": "erro",
                "mensagem": "Envie um arquivo CSV e pelo menos um modelo DOCX."
            }), 400

        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / csv_file.filename
            csv_file.save(csv_path)

            modelos_paths = []
            for m in modelos_files:
                modelo_path = Path(tmpdir) / m.filename
                m.save(modelo_path)
                modelos_paths.append(str(modelo_path))

            resultado = gerar_certificados_mp(str(csv_path), modelos_paths, tipo)

            if resultado["status"] != "ok":
                return jsonify(resultado), 400

            # 🔹 Compacta todos os certificados gerados
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for caminho in resultado["arquivos"]:
                    zipf.write(caminho, Path(caminho).name)
            zip_buffer.seek(0)

            # 🔹 Retorna o ZIP como download direto
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name="certificados.zip",
                mimetype="application/zip"
            )

    except Exception as e:
        return jsonify({
            "status": "erro",
            "mensagem": f"Ocorreu um erro no servidor: {str(e)}"
        }), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
