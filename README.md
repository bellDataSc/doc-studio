# Doc Studio — Gerador e Editor de Documentos Word

Projeto de estudo para praticar docxtpl, Jinja2, python-docx e Streamlit.


## O que estou praticando...

- `docxtpl` → criar templates Word com `{{ variavel }}` e `{%tr for row in lista %}`
- `Jinja2` → condicionais `{% if %}`, filtros, loops
- `python-docx` → inspecionar estrutura de um .docx
- `streamlit` → `st.file_uploader`, `st.download_button`, `st.form`
- `matplotlib` → gerar gráfico e embutir no Word



## Estrutura do projeto

```
doc-studio/
├── app.py                  ← App principal Streamlit
├── requirements.txt        ← Dependências
├── templates/
│   ├── relatorio.docx      ← Template com variáveis Jinja2
│   └── contrato.docx       ← Template com loop de itens
├── utils/
│   ├── extractor.py        ← Lê dados de upload (CSV/JSON)
│   └── writer.py           ← Renderiza template com docxtpl
└── README.md
```

## Como rodar no GitHub Codespace

1. Abra o repositório no GitHub e clique em **Code > Codespaces > New codespace**
2. No terminal do Codespace, instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Rode o app:
   ```bash
   streamlit run app.py
   ```
