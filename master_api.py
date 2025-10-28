from flask import Flask, request, jsonify
from flask_cors import CORS
from Master_core import gerar_certificados_mp
from pathlib import Path
import os
import tempfile

app = Flask(__name__)

# ✅ Permitir chamadas do seu GitHub Pages (ajuste o domínio se mudar)
CORS(app, resources={r"/*": {"origins": ["https://vytorvilar.github.io", "http://localhost:5500"]}})

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "online",
        "mensagem": "API de geração de certificados está ativa
