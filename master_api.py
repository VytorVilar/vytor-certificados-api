from flask import Flask, request, jsonify
from Master import gerar_certificados_mp
from pathlib import Path
import os

app = Flask(__name__)

@app.route("/")
def home():
    return jsonify({
        "status": "online",
        "mensagem": "API do Gerador de Certificados está ativa."
    })

@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        dados = request.json
        csv = Path(dados["csv"])
        modelos = [Path(m) for m in dados["modelos"]]
        tipo = dados.get("tipo", "PDF")

        gerar_certificados_mp(csv, modelos, tipo)

        return jsonify({"status": "ok", "mensagem": "Certificados gerados com sucesso!"})
    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
