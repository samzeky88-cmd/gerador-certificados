import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Gerador de Certificados — Coordenadores", layout="wide")
st.title("📜 Gerador de Certificados — Coordenadores")

# --- 1. Carregar a planilha ---
arquivo_excel = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"])

if arquivo_excel:
    df = pd.read_excel(arquivo_excel, sheet_name="CORDENADORES")
    
    # Limpar nomes das colunas (remover espaços)
    df.columns = df.columns.str.strip()
    
    # Lista de nomes únicos
    nomes = sorted(df["NOME"].dropna().unique())
    nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes)

    if nome_escolhido:
        # Filtrar só os dados da pessoa escolhida
        dados_pessoa = df[df["NOME"] == nome_escolhido].reset_index(drop=True)
        
        # Preparar os dados para o modelo
        contexto = {"NOME": nome_escolhido}
        
        # Preencher do 1º ao 11º treinamento
        for i in range(11):
            if i < len(dados_pessoa):
                linha = dados_pessoa.iloc[i]
                contexto[f"DATA_{i+1}"] = linha["DATA"].strftime("%d/%m/%Y") if hasattr(linha["DATA"], 'strftime') else str(linha["DATA"])
                contexto[f"TEMA_{i+1}"] = str(linha["TEMA"])
                contexto[f"CH_{i+1}"] = float(linha["CH"])
            else:
                # Se não tiver mais treinamentos, deixa em branco
                contexto[f"DATA_{i+1}"] = ""
                contexto[f"TEMA_{i+1}"] = ""
                contexto[f"CH_{i+1}"] = 0
        
        # Calcular carga total
        ch_valores = [contexto[f"CH_{i+1}"] for i in range(11)]
        contexto["CARGA_TOTAL"] = sum(ch_valores)

        # Mostrar resumo
        st.subheader("📋 Resumo do participante")
        st.write(f"**Nome:** {nome_escolhido}")
        st.write(f"**Quantidade de treinamentos:** {len(dados_pessoa)}")
        st.write(f"**Carga Horária Total:** {contexto['CARGA_TOTAL']} horas")

        # --- 2. Carregar e preencher o modelo do certificado ---
        st.subheader("📄 Modelo do certificado")
        modelo = st.file_uploader("Escolha o modelo (.docx)", type=["docx"])

        if modelo:
            doc = DocxTemplate(modelo)
            doc.render(contexto)
            
            # Salvar em memória para download
            arquivo_saida = io.BytesIO()
            doc.save(arquivo_saida)
            arquivo_saida.seek(0)
            
            st.success("✅ Certificado gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Certificado",
                data=arquivo_saida,
                file_name=f"Certificado_{nome_escolhido.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
