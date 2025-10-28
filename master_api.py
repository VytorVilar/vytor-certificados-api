from flask import Flask, request, jsonify
from Master_core import gerar_certificados_mp
from pathlib import Path
import os

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
        dados = request.json
        csv = dados["csv"]
        modelos = dados["modelos"]
        tipo = dados.get("tipo", "PDF")

        resultado = gerar_certificados_mp(csv, modelos, tipo)
        return jsonify(resultado)
    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
