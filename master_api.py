from flask import Flask, request, jsonify
from Master_core import gerar_certificados_mp
from pathlib import Path
import os
import tempfile

app = Flask(__name__)

@app.route("/")
def home():
    return jsonify({
        "status": "online",
        "mensagem": "API de geração de certificados está ativa 🚀"
    })

@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        # Recebe arquivo CSV e lista de modelos
        csv_file = request.files.get("csv")
        modelos_files = request.files.getlist("modelos")
        tipo = request.form.get("tipo", "PDF")

        if not csv_file or not modelos_files:
            return jsonify({"status": "erro", "mensagem": "Envie um CSV e ao menos um modelo DOCX"}), 400

        # Cria diretório temporário
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / csv_file.filename
            csv_file.save(csv_path)

            modelos_paths = []
            for m in modelos_files:
                m_path = Path(tmpdir) / m.filename
                m.save(m_path)
                modelos_paths.append(str(m_path))

            resultado = gerar_certificados_mp(str(csv_path), modelos_paths, tipo)
            return jsonify(resultado)

    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
