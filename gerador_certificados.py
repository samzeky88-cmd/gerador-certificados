import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Gerador de Certificados — Coordenadores", layout="wide")
st.title("📜 Gerador de Certificados — Coordenadores")

# Mensagem bíblica
st.markdown(
    "<div style='text-align: center; font-size: 18px; color: #1a4d80; font-style: italic; margin: 10px 0 25px 0;'>"
    "Consagre seus planos ao Senhor e eles serão bem-sucedidos ✨"
    "</div>",
    unsafe_allow_html=True
)

# --- 1. Carregar a planilha ---
arquivo_excel = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"])

if arquivo_excel:
    df = pd.read_excel(arquivo_excel, sheet_name="CORDENADORES", engine="openpyxl")
    df.columns = df.columns.str.strip()
    
    nomes = sorted(df["NOME"].dropna().unique())
    nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes)

    if nome_escolhido:
        dados_pessoa = df[df["NOME"] == nome_escolhido].reset_index(drop=True)
        contexto = {"NOME": nome_escolhido}
        
        # Preencher do 1º ao 11º — com os nomes do seu modelo!
        for i in range(11):
            if i < len(dados_pessoa):
                linha = dados_pessoa.iloc[i]
                # Data
                data_val = linha["DATA"]
                if hasattr(data_val, 'strftime'):
                    contexto[f"Data_{i+1}"] = data_val.strftime("%d/%m/%Y")
                else:
                    contexto[f"Data_{i+1}"] = str(data_val)
                # Tema
                contexto[f"Tema_{i+1}"] = str(linha["TEMA"])
                # Carga horária formatada com H
                ch_valor = float(linha["CH"])
                if ch_valor.is_integer():
                    contexto[f"Ch_{i+1}"] = f"{int(ch_valor)}H"
                else:
                    contexto[f"Ch_{i+1}"] = f"{ch_valor}H"
            else:
                # Linhas sem dados → vazias
                contexto[f"Data_{i+1}"] = ""
                contexto[f"Tema_{i+1}"] = ""
                contexto[f"Ch_{i+1}"] = ""
        
        # Cálculo da carga total
        ch_total = 0.0
        for i in range(11):
            ch_texto = contexto[f"Ch_{i+1}"]
            if ch_texto:
                ch_total += float(ch_texto.replace("H", ""))
        
        if ch_total.is_integer():
            contexto["CARGA_TOTAL"] = f"{int(ch_total)}H"
        else:
            contexto["CARGA_TOTAL"] = f"{ch_total}H"

        # Resumo na tela
        st.subheader("📋 Resumo do participante")
        st.write(f"**Nome:** {nome_escolhido}")
        st.write(f"**Quantidade de treinamentos:** {len(dados_pessoa)}")
        st.write(f"**Carga Horária Total:** {contexto['CARGA_TOTAL']}")

        # --- 2. Carregar modelo ---
        st.subheader("📄 Modelo do certificado (.docx)")
        modelo = st.file_uploader("Escolha o modelo do certificado", type=["docx"])

        if modelo:
            doc = DocxTemplate(modelo)
            doc.render(contexto)
            
            saida = io.BytesIO()
            doc.save(saida)
            saida.seek(0)
            
            st.success("✅ Certificado gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Certificado",
                data=saida,
                file_name=f"Certificado_{nome_escolhido.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
