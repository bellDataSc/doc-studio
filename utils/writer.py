"""
writer.py
Responsabilidade: orquestrar a geração do documento Word.
Recebe o template path, o contexto (dicionário) e opcionalmente gráficos,
e retorna os bytes do .docx gerado.
"""
import io
import tempfile
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm


def renderizar(template_path: str, contexto: dict, graficos: dict = None) -> bytes:
    """
    Renderiza um template .docx com o contexto Jinja2.

    Args:
        template_path: caminho para o arquivo .docx template
        contexto: dicionário com variáveis Jinja2
        graficos: dict {nome_alt: bytes_png} para injetar imagens

    Returns:
        bytes do documento gerado
    """
    doc = DocxTemplate(template_path)

    # Injeta gráficos como InlineImage se houver
    if graficos:
        for chave, img_bytes in graficos.items():
            # Salva PNG em arquivo temporário (InlineImage precisa de path ou file-like)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(img_bytes)
            tmp.close()
            contexto[chave] = InlineImage(doc, tmp.name, width=Cm(14))

    doc.render(contexto)

    saida = io.BytesIO()
    doc.save(saida)
    saida.seek(0)

    # Limpeza dos temporários
    if graficos:
        for chave in graficos:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass

    return saida.read()
