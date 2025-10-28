import pandas as pd
from docxtpl import DocxTemplate
from pathlib import Path
import os
import datetime

def gerar_certificados_mp(csv_path, modelos, tipo="PDF"):
    """
    Gera certificados ou PPPs a partir de um CSV e um ou mais modelos DOCX.
    """
    try:
        csv_path = Path(csv_path)
        if not csv_path.exists():
            raise FileNotFoundError(f"Arquivo CSV não encontrado: {csv_path}")

        # Pasta de saída
        data = datetime.date.today().isoformat()
        pasta_saida = Path("saida") / data
        pasta_saida.mkdir(parents=True, exist_ok=True)

        df = pd.read_csv(csv_path, encoding="utf-8", sep=";|,", engine="python")

        resultados = []
        for _, linha in df.iterrows():
            dados = linha.to_dict()
            for modelo in modelos:
                modelo_path = Path(modelo)
                if not modelo_path.exists():
                    raise FileNotFoundError(f"Modelo não encontrado: {modelo_path}")

                doc = DocxTemplate(modelo_path)
                doc.render(dados)

                nome_saida = f"{dados.get('Nome', 'Sem_Nome')}.docx"
                arquivo_saida = pasta_saida / nome_saida
                doc.save(arquivo_saida)
                resultados.append(str(arquivo_saida))

        return {
            "status": "ok",
            "mensagem": f"{len(resultados)} arquivos gerados com sucesso.",
            "arquivos": resultados
        }

    except Exception as e:
        return {"status": "erro", "mensagem": str(e)}
