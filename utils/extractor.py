"""
extractor.py
Responsabilidade: extrair e formatar dados vindos de upload (CSV ou JSON)
para um dicionário Python que será passado ao docxtpl.
"""
import pandas as pd
import json
import io


def extrair_de_csv(arquivo_bytes: bytes) -> dict:
    """
    Lê um CSV com colunas: nome, valor, categoria
    Retorna dicionário com lista de itens e totais.
    """
    df = pd.read_csv(io.BytesIO(arquivo_bytes))
    df.columns = [c.strip().lower() for c in df.columns]

    itens = []
    for _, row in df.iterrows():
        itens.append({
            "nome": str(row.get("nome", "")),
            "valor": float(row.get("valor", 0)),
            "valor_fmt": f"R$ {float(row.get('valor', 0)):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "categoria": str(row.get("categoria", "")),
        })

    total = sum(i["valor"] for i in itens)
    return {
        "itens": itens,
        "total": total,
        "total_fmt": f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        "qtd_itens": len(itens),
    }


def extrair_de_json(arquivo_bytes: bytes) -> dict:
    """
    Lê um JSON com campos livres.
    Retorna o próprio dicionário (o JSON já é o contexto).
    """
    return json.loads(arquivo_bytes.decode("utf-8"))
