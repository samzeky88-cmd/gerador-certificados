import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Gerador de Certificados — Coordenadores", layout="wide")
st.title("📜 Gerador de Certificados — Coordenadores")

# --- 1. Carregar a planilha ---
arquivo_excel = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"])

if arquivo_excel:
    # Leitura com motor openpyxl
    df = pd.read_excel(arquivo_excel, sheet_name="CORDENADORES", engine="openpyxl")
    
    # Limpar nomes das colunas
    df.columns = df.columns.str.strip()
    
    # Lista de nomes únicos ordenada
    nomes = sorted(df["NOME"].dropna().unique())
    nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes)

    if nome_escolhido:
        # Filtrar dados da pessoa selecionada
        dados_pessoa = df[df["NOME"] == nome_escolhido].reset_index(drop=True)
        
        # Preparar dicionário para o modelo
        contexto = {"NOME": nome_escolhido}
        
        # Preencher do 1º ao 11º treinamento
        for i in range(11):
            if i < len(dados_pessoa):
                linha = dados_pessoa.iloc[i]
                # Formatar data corretamente
                data_val = linha["DATA"]
                if hasattr(data_val, 'strftime'):
                    contexto[f"DATA_{i+1}"] = data_val.strftime("%d/%m/%Y")
                else:
                    contexto[f"DATA_{i+1}"] = str(data_val)
                # Demais campos
                contexto[f"TEMA_{i+1}"] = str(linha["TEMA"])
                contexto[f"CH_{i+1}"] = float(linha["CH"])
            else:
                # Deixar em branco se não houver mais treinamentos
                contexto[f"DATA_{i+1}"] = ""
                contexto[f"TEMA_{i+1}"] = ""
                contexto[f"CH_{i+1}"] = 0.0
        
        # Calcular carga horária total
        ch_total = sum(contexto[f"CH_{i+1}"] for i in range(11))
        contexto["CARGA_TOTAL"] = ch_total

        # Exibir resumo na tela
        st.subheader("📋 Resumo do participante")
        st.write(f"**Nome:** {nome_escolhido}")
        st.write(f"**Quantidade de treinamentos:** {len(dados_pessoa)}")
        st.write(f"**Carga Horária Total:** {ch_total} horas")

        # --- 2. Carregar modelo do certificado ---
        st.subheader("📄 Modelo do certificado (.docx)")
        modelo = st.file_uploader("Escolha o modelo do certificado", type=["docx"])

        if modelo:
            doc = DocxTemplate(modelo)
            doc.render(contexto)
            
            # Preparar para download
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
