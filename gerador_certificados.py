import streamlit as st
import pandas as pd
import io
from pptx import Presentation

st.set_page_config(page_title="Emissor de Certificados", page_icon="📜", layout="wide")
st.title("📜 Emissor de Certificados — NAQH/NSP")
caminho_modelo = "modelo_certificado.pptx"

def tratar_valor(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ["nan", "none", ""]:
        return ""
    return str(valor).strip()

def preencher_2_slides(caminho_modelo, dados):
    prs = Presentation(caminho_modelo)
    
    # Slide 1 — Nome do participante
    slide1 = prs.slides[0]
    for forma in slide1.shapes:
        if forma.has_text_frame:
            for paragrafo in forma.text_frame.paragraphs:
                for run in paragrafo.runs:
                    if "{{NOME}}" in run.text:
                        run.text = run.text.replace("{{NOME}}", dados["{{NOME}}"])
    
    # Slide 2 — Tabela de temas
    slide2 = prs.slides[1]
    for forma in slide2.shapes:
        if forma.has_text_frame:
            for paragrafo in forma.text_frame.paragraphs:
                for chave, valor in dados.items():
                    if chave in paragrafo.text:
                        paragrafo.text = paragrafo.text.replace(chave, valor)
        if forma.has_table:
            tabela = forma.table
            for linha in tabela.rows:
                for celula in linha.cells:
                    for paragrafo in celula.text_frame.paragraphs:
                        for chave, valor in dados.items():
                            if chave in paragrafo.text:
                                paragrafo.text = paragrafo.text.replace(chave, valor)
    return prs

# === Interface: onde você envia os dados ===
arquivo_planilha = st.file_uploader("📂 Planilha com os participantes (.xlsx)", type=["xlsx"])

if arquivo_planilha:
    df = pd.read_excel(arquivo_planilha)
    
    if "NOME" not in df.columns:
        st.error("❌ A planilha precisa ter a coluna 'NOME'")
    else:
        df = df.dropna(subset=["NOME"])
        df["NOME"] = df["NOME"].astype(str).str.strip()
        df = df[df["NOME"] != ""]
        
        st.success(f"✅ {len(df)} participantes carregados")
        st.divider()

        for idx, linha in df.iterrows():
            nome = tratar_valor(linha["NOME"])
            st.markdown(f"**{idx+1}. {nome}**")

            dados = {
                "{{NOME}}": nome,
                "{{DATA_1}}": tratar_valor(linha.get("DATA_1", "")),
                "{{TEMA_1}}": tratar_valor(linha.get("TEMA_1", "")),
                "{{CH_1}}": tratar_valor(linha.get("CH_1", "")),
                "{{DATA_2}}": tratar_valor(linha.get("DATA_2", "")),
                "{{TEMA_2}}": tratar_valor(linha.get("TEMA_2", "")),
                "{{CH_2}}": tratar_valor(linha.get("CH_2", "")),
                "{{DATA_3}}": tratar_valor(linha.get("DATA_3", "")),
                "{{TEMA_3}}": tratar_valor(linha.get("TEMA_3", "")),
                "{{CH_3}}": tratar_valor(linha.get("CH_3", "")),
                "{{DATA_4}}": tratar_valor(linha.get("DATA_4", "")),
                "{{TEMA_4}}": tratar_valor(linha.get("TEMA_4", "")),
                "{{CH_4}}": tratar_valor(linha.get("CH_4", "")),
                "{{DATA_5}}": tratar_valor(linha.get("DATA_5", "")),
                "{{TEMA_5}}": tratar_valor(linha.get("TEMA_5", "")),
                "{{CH_5}}": tratar_valor(linha.get("CH_5", "")),
                "{{DATA_6}}": tratar_valor(linha.get("DATA_6", "")),
                "{{TEMA_6}}": tratar_valor(linha.get("TEMA_6", "")),
                "{{CH_6}}": tratar_valor(linha.get("CH_6", "")),
                "{{DATA_7}}": tratar_valor(linha.get("DATA_7", "")),
                "{{TEMA_7}}": tratar_valor(linha.get("TEMA_7", "")),
                "{{CH_7}}": tratar_valor(linha.get("CH_7", "")),
                "{{CARGA_TOTAL}}": tratar_valor(linha.get("CARGA_TOTAL", "")),
            }

            certificado = preencher_2_slides(caminho_modelo, dados)
            saida = io.BytesIO()
            certificado.save(saida)
            saida.seek(0)

            st.download_button(
                label="📄 Baixar Certificado",
                data=saida,
                file_name=f"Certificado - {nome}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key=f"cert_{idx}"
            )
            st.divider()
else:
    st.info("⬆️ Faça upload da planilha .xlsx para começar")

