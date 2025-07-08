import streamlit as st
from docx import Document
from docx.shared import RGBColor
import zipfile
import os
import re
import tempfile

def processar_docx(uploaded_files, texto_referencia):
    arquivos_processados = []

    for uploaded_file in uploaded_files:
        # Como uploaded_file é um UploadedFile do Streamlit, 
        # podemos passar o buffer diretamente para Document
        doc = Document(uploaded_file)
        alterado = False

        for p in doc.paragraphs:
            if texto_referencia in p.text and '{' in p.text and '}' in p.text:
                # Procura o primeiro trecho entre chaves
                match = re.search(r'\{(.+?)\}', p.text)
                if match:
                    dentro_chave = match.group(1)
                    antes = p.text.split('{')[0]
                    depois = p.text.split('}')[1] if '}' in p.text else ''

                    p.clear()
                    p.add_run(antes)

                    run_vermelho = p.add_run(dentro_chave)
                    run_vermelho.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

                    p.add_run(depois)
                    alterado = True

        if alterado:
            # Salva em arquivo temporário para depois zipar
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(tmp_file.name)
            arquivos_processados.append((uploaded_file.name, tmp_file.name))

    return arquivos_processados

st.title("📝 Editor DOCX: Destacar texto entre chaves")

st.markdown("""
Faça upload de um ou mais arquivos `.docx`. 
Informe o texto-chave que deve aparecer no parágrafo onde será destacado o texto entre `{}`.
""")

uploaded_files = st.file_uploader("📁 Selecione arquivos DOCX", type="docx", accept_multiple_files=True)
texto_referencia = st.text_input("🔍 Texto-chave", value="será na data:")

if st.button("✅ Processar arquivos"):
    if not uploaded_files:
        st.warning("Por favor, envie pelo menos um arquivo `.docx`.")
    else:
        resultado = processar_docx(uploaded_files, texto_referencia)
        if resultado:
            # Criar zip para download
            zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for nome_arquivo, caminho_temp in resultado:
                    zipf.write(caminho_temp, arcname=nome_arquivo)
            st.success(f"{len(resultado)} arquivo(s) processado(s) com sucesso!")
            with open(zip_path, "rb") as f:
                st.download_button("📥 Baixar arquivos editados (.zip)", f.read(), file_name="documentos_editados.zip")
        else:
            st.info("Nenhum arquivo continha o texto-chave com texto entre chaves `{}` para editar.")
