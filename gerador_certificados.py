import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Sistema de Certificados — Coordenadores", layout="wide")

# --- MENU DE ABAS ---
aba_certificados, aba_ranking = st.tabs(["📜 Gerar Certificados", "📊 Ranking de Participação"])

# ==================================================
# ABA 1 — GERAR CERTIFICADOS (exatamente como já está)
# ==================================================
with aba_certificados:
    st.title("📜 Gerador de Certificados — Coordenadores")
    
    # Mensagem bíblica
    st.markdown(
        "<div style='text-align: center; font-size: 18px; color: #1a4d80; font-style: italic; margin: 10px 0 25px 0;'>"
        "Consagre seus planos ao Senhor e eles serão bem-sucedidos ✨"
        "</div>",
        unsafe_allow_html=True
    )

    arquivo_excel = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"], key="cert_planilha")

    if arquivo_excel:
        df = pd.read_excel(arquivo_excel, sheet_name="CORDENADORES", engine="openpyxl")
        df.columns = df.columns.str.strip()
        
        nomes = sorted(df["NOME"].dropna().unique())
        nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes, key="cert_nome")

        if nome_escolhido:
            dados_pessoa = df[df["NOME"] == nome_escolhido].reset_index(drop=True)
            contexto = {"NOME": nome_escolhido}
            
            # Preencher do 1º ao 11º — nomes do seu modelo: Data_, Tema_, Ch_
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
                    # Carga horária com H
                    ch_valor = float(linha["CH"])
                    if ch_valor.is_integer():
                        contexto[f"Ch_{i+1}"] = f"{int(ch_valor)}H"
                    else:
                        contexto[f"Ch_{i+1}"] = f"{ch_valor}H"
                else:
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

            st.subheader("📋 Resumo do participante")
            st.write(f"**Nome:** {nome_escolhido}")
            st.write(f"**Quantidade de treinamentos:** {len(dados_pessoa)}")
            st.write(f"**Carga Horária Total:** {contexto['CARGA_TOTAL']}")

            st.subheader("📄 Modelo do certificado (.docx)")
            modelo = st.file_uploader("Escolha o modelo do certificado", type=["docx"], key="cert_modelo")

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
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="cert_download"
                )

# ==================================================
# ABA 2 — RANKING DE PARTICIPAÇÃO (NOVO — separado!)
# ==================================================
with aba_ranking:
    st.title("📊 Ranking de Participação nos Treinamentos")
    
    arquivo_ranking = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"], key="rank_planilha")

    if arquivo_ranking:
        df_rank = pd.read_excel(arquivo_ranking, sheet_name="CORDENADORES", engine="openpyxl")
        df_rank.columns = df_rank.columns.str.strip()
        
        # Calcular métricas por pessoa
        ranking = df_rank.groupby("NOME").agg(
            QUANTIDADE_TREINAMENTOS=("NOME", "count"),
            CARGA_HORARIA_TOTAL=("CH", "sum")
        ).reset_index()
        
        # Ordenar por quantidade (maior → menor) e depois por carga horária
        ranking = ranking.sort_values(
            by=["QUANTIDADE_TREINAMENTOS", "CARGA_HORARIA_TOTAL"],
            ascending=[False, False]
        ).reset_index(drop=True)
        
        # Adicionar coluna de posição
        ranking.insert(0, "POSIÇÃO", range(1, len(ranking) + 1))
        
        # Formatar carga horária
        ranking["CARGA_HORARIA_TOTAL"] = ranking["CARGA_HORARIA_TOTAL"].apply(
            lambda x: f"{int(x)}H" if x.is_integer() else f"{x}H"
        )
        
        # Exibir tabela bonita
        st.dataframe(ranking, use_container_width=True, hide_index=True)
        
        # Resumo geral
        st.subheader("📈 Resumo Geral")
        st.write(f"**Total de participantes:** {len(ranking)}")
        st.write(f"**Total de treinamentos realizados:** {ranking['QUANTIDADE_TREINAMENTOS'].sum()}")
        
        # Destaque Top 3
        st.subheader("🏆 Destaques Top 3")
        for pos in range(min(3, len(ranking))):
            emoji = "🥇" if pos == 0 else "🥈" if pos == 1 else "🥉"
            linha = ranking.iloc[pos]
            st.write(f"{emoji} **{linha['POSIÇÃO']}º Lugar:** {linha['NOME']} — {linha['QUANTIDADE_TREINAMENTOS']} treinamentos | {linha['CARGA_HORARIA_TOTAL']}")
