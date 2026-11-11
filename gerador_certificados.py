import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import unicodedata

st.set_page_config(page_title="Sistema de Certificados — Coordenadores", layout="wide")

MAX_TREINAMENTOS = 14

# ==================================================
# 🧠 FUNÇÕES AUXILIARES
# ==================================================
def normalizar_nome(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = texto.lower()
    texto = texto.strip()
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def remover_acentos(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII').lower()

def corrigir_texto(texto):
    if not isinstance(texto, str):
        try:
            texto = str(texto)
        except:
            return ""
    if texto.strip() == "" or texto.strip().upper() == "NAN":
        return ""
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto


# --- ABAS DO SISTEMA ---
aba_certificados, aba_ranking, aba_controle = st.tabs([
    "📜 Gerar Certificados", 
    "📊 Ranking de Participação",
    "✅ Controle de Impressão"
])

# ==================================================
# ABA 1 — GERAR CERTIFICADOS
# ==================================================
with aba_certificados:
    st.title("📜 Gerador de Certificados — Coordenadores")
    
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
        
        df["NOME_NORMALIZADO"] = df["NOME"].apply(normalizar_nome)
        mapa_nomes = df.drop_duplicates(subset=["NOME_NORMALIZADO"])[["NOME_NORMALIZADO", "NOME"]]
        mapa_nomes = mapa_nomes.set_index("NOME_NORMALIZADO")["NOME"].to_dict()
        df["NOME_CORRIGIDO"] = df["NOME_NORMALIZADO"].map(mapa_nomes)
        
        if "STATUS" in df.columns:
            status_por_nome = df.groupby("NOME_CORRIGIDO")["STATUS"].apply(
                lambda x: any(str(val).strip().upper() == "IMPRESSO" for val in x.dropna())
            ).reset_index()
            status_por_nome.columns = ["NOME", "JA_IMPRESSO"]
            impressos_nomes = status_por_nome[status_por_nome["JA_IMPRESSO"] == True]["NOME"].unique()
        else:
            impressos_nomes = []
        
        todos_nomes_unicos = df["NOME_CORRIGIDO"].dropna().unique()
        todos_nomes = sorted(todos_nomes_unicos, key=lambda x: remover_acentos(str(x)))
        
        mostrar_so_pendentes = st.checkbox("👁️ Mostrar apenas pendentes", value=False)
        
        if mostrar_so_pendentes:
            nomes_filtrados = [n for n in todos_nomes if n not in impressos_nomes]
            st.info(f"📋 {len(nomes_filtrados)} pendentes de {len(todos_nomes)} total")
        else:
            nomes_filtrados = todos_nomes
        
        nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes_filtrados, key="cert_nome")

        if nome_escolhido:
            dados_pessoa = df[df["NOME_CORRIGIDO"] == nome_escolhido].reset_index(drop=True)
            contexto = {"NOME": nome_escolhido}
            
            st.info(f"✅ {nome_escolhido} — {len(dados_pessoa)} treinamento(s) incluídos")
            
            for i in range(MAX_TREINAMENTOS):
                if i < len(dados_pessoa):
                    linha = dados_pessoa.iloc[i]
                    data_val = linha["DATA"]
                    if pd.isna(data_val):
                        contexto[f"Data_{i+1}"] = ""
                    elif hasattr(data_val, 'strftime'):
                        contexto[f"Data_{i+1}"] = data_val.strftime("%d/%m/%Y")
                    else:
                        contexto[f"Data_{i+1}"] = str(data_val)
                    
                    tema_bruto = linha["TEMA"]
                    if pd.isna(tema_bruto):
                        contexto[f"Tema_{i+1}"] = ""
                    else:
                        contexto[f"Tema_{i+1}"] = corrigir_texto(str(tema_bruto))
                    
                    ch_valor = linha["CH"]
                    if pd.isna(ch_valor):
                        contexto[f"Ch_{i+1}"] = ""
                    else:
                        ch_valor = float(ch_valor)
                        if ch_valor.is_integer():
                            contexto[f"Ch_{i+1}"] = f"{int(ch_valor)}H"
                        else:
                            contexto[f"Ch_{i+1}"] = f"{ch_valor:.2f}H".rstrip('0').rstrip('.') + 'H'
                else:
                    contexto[f"Data_{i+1}"] = ""
                    contexto[f"Tema_{i+1}"] = ""
                    contexto[f"Ch_{i+1}"] = ""
            
            ch_total = 0.0
            for i in range(MAX_TREINAMENTOS):
                ch_texto = contexto[f"Ch_{i+1}"]
                if ch_texto and ch_texto != "":
                    try:
                        ch_total += float(ch_texto.replace("H", ""))
                    except:
                        pass
            
            if ch_total.is_integer():
                contexto["CARGA_TOTAL"] = f"{int(ch_total)}H"
            else:
                contexto["CARGA_TOTAL"] = f"{ch_total:.2f}H".rstrip('0').rstrip('.') + 'H'

            st.subheader("📋 Resumo do participante")
            st.write(f"**Nome:** {nome_escolhido}")
            st.write(f"**Treinamentos:** {len(dados_pessoa)}")
            st.write(f"**Carga Horária Total:** {contexto['CARGA_TOTAL']}")
            
            with st.expander("📝 Ver lista de treinamentos"):
                for i in range(len(dados_pessoa)):
                    data = contexto[f'Data_{i+1}']
                    tema = contexto[f'Tema_{i+1}']
                    ch = contexto[f'Ch_{i+1}']
                    if tema:
                        st.write(f"{i+1}. {data} — {tema} — {ch}")

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
                st.info("💡 Após imprimir, coloque 'IMPRESSO' em UMA das linhas dessa pessoa na coluna STATUS da planilha e reenvie!")

# ==================================================
# ABA 2 — RANKING
# ==================================================
with aba_ranking:
    st.title("📊 Ranking de Participação nos Treinamentos")
    
    arquivo_ranking = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"], key="rank_planilha")

    if arquivo_ranking:
        df_rank = pd.read_excel(arquivo_ranking, sheet_name="CORDENADORES", engine="openpyxl")
        df_rank.columns = df_rank.columns.str.strip()
        
        df_rank["NOME_NORMALIZADO"] = df_rank["NOME"].apply(normalizar_nome)
        mapa_nomes_rank = df_rank.drop_duplicates(subset=["NOME_NORMALIZADO"])[["NOME_NORMALIZADO", "NOME"]]
        mapa_nomes_rank = mapa_nomes_rank.set_index("NOME_NORMALIZADO")["NOME"].to_dict()
        df_rank["NOME_CORRIGIDO"] = df_rank["NOME_NORMALIZADO"].map(mapa_nomes_rank)
        
        ranking = df_rank.groupby("NOME_CORRIGIDO").agg(
            QUANTIDADE_TREINAMENTOS=("NOME_CORRIGIDO", "count"),
            CARGA_HORARIA_TOTAL=("CH", "sum")
        ).reset_index()
        ranking = ranking.rename(columns={"NOME_CORRIGIDO": "NOME"})
        
        ranking = ranking.sort_values(
            by=["QUANTIDADE_TREINAMENTOS", "CARGA_HORARIA_TOTAL"],
            ascending=[False, False]
        ).reset_index(drop=True)
        
        ranking.insert(0, "POSIÇÃO", range(1, len(ranking) + 1))
        
        def formatar_ch(valor):
            if pd.isna(valor):
                return "0H"
            if valor.is_integer():
                return f"{int(valor)}H"
            return f"{valor:.2f}H".rstrip('0').rstrip('.') + 'H'
        
        ranking["CARGA_HORARIA_TOTAL"] = ranking["CARGA_HORARIA_TOTAL"].apply(formatar_ch)
        
        st.dataframe(ranking, use_container_width=True, hide_index=True)
        
        st.subheader("📈 Resumo Geral")
        st.write(f"**Total de participantes únicos:** {len(ranking)}")
        st.write(f"**Total de treinamentos realizados:** {ranking['QUANTIDADE_TREINAMENTOS'].sum()}")
        
        st.subheader("🏆 Destaques — Top 10")
        emojis = ["🥇", "🥈", "🥉"] + ["🏅"] * 7
        for pos in range(min(10, len(ranking))):
            emoji = emojis[pos]
            linha = ranking.iloc[pos]
            st.write(f"{emoji} **{linha['POSIÇÃO']}º Lugar:** {linha['NOME']} — {linha['QUANTIDADE_TREINAMENTOS']} treinamentos | {linha['CARGA_HORARIA_TOTAL']}")

# ==================================================
# ABA 3 — CONTROLE DE IMPRESSÃO (COM ORDEM ALFABÉTICA CORRIGIDA)
# ==================================================
with aba_controle:
    st.title("✅ Controle de Impressão de Certificados")
    
    arquivo_ctrl = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"], key="ctrl_planilha")
    
    if arquivo_ctrl:
        df_ctrl = pd.read_excel(arquivo_ctrl, sheet_name="CORDENADORES", engine="openpyxl")
        df_ctrl.columns = df_ctrl.columns.str.strip()
        
        df_ctrl["NOME_NORMALIZADO"] = df_ctrl["NOME"].apply(normalizar_nome)
        mapa_nomes_ctrl = df_ctrl.drop_duplicates(subset=["NOME_NORMALIZADO"])[["NOME_NORMALIZADO", "NOME"]]
        mapa_nomes_ctrl = mapa_nomes_ctrl.set_index("NOME_NORMALIZADO")["NOME"].to_dict()
        df_ctrl["NOME_CORRIGIDO"] = df_ctrl["NOME_NORMALIZADO"].map(mapa_nomes_ctrl)
        
        resumo_pessoas = df_ctrl.groupby("NOME_CORRIGIDO").agg(
            TREINAMENTOS=("NOME_CORRIGIDO", "count"),
            CARGA_HORAS=("CH", "sum")
        ).reset_index()
        resumo_pessoas = resumo_pessoas.rename(columns={"NOME_CORRIGIDO": "NOME"})
        
        if "STATUS" in df_ctrl.columns:
            status_por_nome = df_ctrl.groupby("NOME_CORRIGIDO")["STATUS"].apply(
                lambda x: any(str(val).strip().upper() == "IMPRESSO" for val in x.dropna())
            ).reset_index()
            status_por_nome.columns = ["NOME_CORRIGIDO", "JA_IMPRESSO"]
            
            resumo_pessoas = pd.merge(
                resumo_pessoas, 
                status_por_nome, 
                left_on="NOME", 
                right_on="NOME_CORRIGIDO", 
                how="left"
            )
            resumo_pessoas["STATUS"] = resumo_pessoas["JA_IMPRESSO"].apply(
                lambda x: "✅ IMPRESSO" if x else "⬜ PENDENTE"
            )
            resumo_pessoas = resumo_pessoas.drop(columns=["NOME_CORRIGIDO", "JA_IMPRESSO"])
            impressos = resumo_pessoas[resumo_pessoas["STATUS"] == "✅ IMPRESSO"]["NOME"].unique()
        else:
            resumo_pessoas["STATUS"] = "⬜ PENDENTE"
            impressos = []
        
        # ✅ ORDENAÇÃO ALFABÉTICA GARANTIDA — A a Z, ignorando acentos
        resumo_pessoas["chave_ordem"] = resumo_pessoas["NOME"].apply(remover_acentos)
        resumo_pessoas = resumo_pessoas.sort_values(by="chave_ordem", ascending=True)
        resumo_pessoas = resumo_pessoas.drop(columns=["chave_ordem"])
        
        def formatar_ch_ctrl(valor):
            if pd.isna(valor):
                return "0H"
            if valor.is_integer():
                return f"{int(valor)}H"
            return f"{valor:.2f}H".rstrip('0').rstrip('.') + 'H'
        
        resumo_pessoas["CARGA_HORAS"] = resumo_pessoas["CARGA_HORAS"].apply(formatar_ch_ctrl)
        
        total = len(resumo_pessoas)
        qtd_impressos = len(impressos)
        qtd_pendentes = total - qtd_impressos
        
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            st.metric(label="📋 TOTAL GERAL", value=f"{total}")
        with col2:
            st.metric(label="✅ IMPRESSOS", value=f"{qtd_impressos}")
        with col3:
            st.metric(label="⬜ PENDENTES", value=f"{qtd_pendentes}")
        
        if total > 0:
            porcentagem = int((qtd_impressos / total) * 100)
            st.progress(porcentagem / 100, text=f"Concluído: {porcentagem}% — {qtd_impressos} de {total}")
        
        st.markdown("---")
        st.dataframe(resumo_pessoas[["STATUS", "NOME", "TREINAMENTOS", "CARGA_HORAS"]], 
                     use_container_width=True, hide_index=True)
        
        st.info("""
        📖 **Legenda:**
        - ✅ IMPRESSO = Já foi emitido
        - ⬜ PENDENTE = Ainda precisa ser emitido
        
        ✅ Nomes em ordem alfabética de A a Z!
        """)
