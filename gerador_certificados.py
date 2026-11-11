import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import unicodedata

st.set_page_config(page_title="Sistema de Certificados — Coordenadores", layout="wide")

MAX_TREINAMENTOS = 14

# ==================================================
# 🧠 FUNÇÕES AUXILIARES — CORRIGIDAS
# ==================================================
def remover_acentos(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII').lower()

def corrigir_texto(texto):
    # ✅ Proteção contra valores vazios ou inválidos
    if not isinstance(texto, str):
        try:
            texto = str(texto)
        except:
            return ""
    if texto.strip() == "" or texto.strip().upper() == "NAN":
        return ""
    
    texto = re.sub(r'\s+', ' ', texto).strip()
    
    palavras_maiuscula = {
        "protocolo", "protocolos", "gestão", "qualidade", "saúde", "segurança",
        "higiene", "hospital", "unidade", "município", "estado", "brasil",
        "programa", "curso", "treinamento", "capacitação", "desenvolvimento",
        "liderança", "coordenadoria", "secretaria", "prefeitura", "diretoria",
        "comissão", "equipe", "serviço", "atenção", "assistência", "controle",
        "avaliação", "planejamento", "organização", "acompanhamento", "análise",
        "resultado", "processo", "melhoria", "eficiência", "eficácia",
        "sistema", "norma", "regulamento", "diretriz", "política", "estratégia"
    }
    
    palavras_minusculas = {
        "a", "o", "e", "de", "da", "do", "das", "dos", "em", "para", "com",
        "sem", "sob", "entre", "por", "pelo", "pela", "nos", "no", "na",
        "os", "as", "um", "uma", "como", "que", "se", "ou", "ao", "à",
        "aos", "às", "é", "e", "mas", "porém", "contudo", "então", "logo"
    }
    
    frases = re.split(r'([.!?])', texto)
    texto_corrigida = ""  # ✅ Nome correto da variável
    
    for i in range(0, len(frases), 2):
        frase = frases[i].strip()
        if not frase:
            continue
        pontuacao = frases[i+1] if i+1 < len(frases) else ""
        
        palavras = frase.split()
        frase_corrigida = []
        
        for idx, palavra in enumerate(palavras):
            p_min = palavra.lower()
            if idx == 0 or p_min in palavras_maiuscula:
                palavra_corrigida = p_min.capitalize()
            elif p_min in palavras_minusculas:
                palavra_corrigida = p_min
            else:
                palavra_corrigida = p_min.capitalize()
            frase_corrigida.append(palavra_corrigida)
        
        texto_corrigida += " ".join(frase_corrigida) + pontuacao
    
    return texto_corrigida  # ✅ Retorna a variável correta


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
        
        todos_nomes = df["NOME"].dropna().unique()
        
        if "STATUS" in df.columns:
            impressos_nomes = df[df["STATUS"].astype(str).str.strip().str.upper() == "IMPRESSO"]["NOME"].unique()
        else:
            impressos_nomes = []
        
        nomes_ordenados = sorted(todos_nomes, key=remover_acentos)
        
        mostrar_so_pendentes = st.checkbox("👁️ Mostrar apenas pendentes", value=False)
        
        if mostrar_so_pendentes:
            nomes_filtrados = [n for n in nomes_ordenados if n not in impressos_nomes]
            st.info(f"📋 {len(nomes_filtrados)} pendentes de {len(nomes_ordenados)} total")
        else:
            nomes_filtrados = nomes_ordenados
        
        nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes_filtrados, key="cert_nome")

        if nome_escolhido:
            dados_pessoa = df[df["NOME"] == nome_escolhido].reset_index(drop=True)
            contexto = {"NOME": nome_escolhido}
            
            st.info(f"✅ {nome_escolhido} — {len(dados_pessoa)} treinamento(s) incluídos")
            
            for i in range(MAX_TREINAMENTOS):
                if i < len(dados_pessoa):
                    linha = dados_pessoa.iloc[i]
                    
                    # Data
                    data_val = linha["DATA"]
                    if pd.isna(data_val):
                        contexto[f"Data_{i+1}"] = ""
                    elif hasattr(data_val, 'strftime'):
                        contexto[f"Data_{i+1}"] = data_val.strftime("%d/%m/%Y")
                    else:
                        contexto[f"Data_{i+1}"] = str(data_val)
                    
                    # ✅ TEMA com proteção
                    tema_bruto = linha["TEMA"]
                    if pd.isna(tema_bruto):
                        contexto[f"Tema_{i+1}"] = ""
                    else:
                        contexto[f"Tema_{i+1}"] = corrigir_texto(str(tema_bruto))
                    
                    # ✅ Carga horária com proteção
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
            
            # Cálculo carga total
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
                st.info("💡 Após imprimir, coloque 'IMPRESSO' na coluna STATUS dessa pessoa na planilha e reenvie para atualizar!")

# ==================================================
# ABA 2 — RANKING
# ==================================================
with aba_ranking:
    st.title("📊 Ranking de Participação nos Treinamentos")
    
    arquivo_ranking = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"], key="rank_planilha")

    if arquivo_ranking:
        df_rank = pd.read_excel(arquivo_ranking, sheet_name="CORDENADORES", engine="openpyxl")
        df_rank.columns = df_rank.columns.str.strip()
        
        ranking = df_rank.groupby("NOME").agg(
            QUANTIDADE_TREINAMENTOS=("NOME", "count"),
            CARGA_HORARIA_TOTAL=("CH", "sum")
        ).reset_index()
        
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
            else:
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
# ABA 3 — CONTROLE DE IMPRESSÃO
# ==================================================
with aba_controle:
    st.title("✅ Controle de Impressão de Certificados")
    
    arquivo_ctrl = st.file_uploader("📂 Escolha a planilha (.xlsx)", type=["xlsx"], key="ctrl_planilha")
    
    if arquivo_ctrl:
        df_ctrl = pd.read_excel(arquivo_ctrl, sheet_name="CORDENADORES", engine="openpyxl")
        df_ctrl.columns = df_ctrl.columns.str.strip()
        
        resumo_pessoas = df_ctrl.groupby("NOME").agg(
            TREINAMENTOS=("NOME", "count"),
            CARGA_HORAS=("CH", "sum")
        ).reset_index()
        
        if "STATUS" in df_ctrl.columns:
            impressos = df_ctrl[df_ctrl["STATUS"].astype(str).str.strip().str.upper() == "IMPRESSO"]["NOME"].unique()
            resumo_pessoas["STATUS"] = resumo_pessoas["NOME"].apply(
                lambda x: "✅ IMPRESSO" if x in impressos else "⬜ PENDENTE"
            )
        else:
            resumo_pessoas["STATUS"] = "⬜ PENDENTE"
            impressos = []
        
        resumo_pessoas = resumo_pessoas.sort_values(by="NOME", key=lambda x: x.apply(remover_acentos))
        
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
        
        def formatar_ch_ctrl(valor):
            if pd.isna(valor):
                return "0H"
            if valor.is_integer():
                return f"{int(valor)}H"
            else:
                return f"{valor:.2f}H".rstrip('0').rstrip('.') + 'H'
        
        resumo_pessoas["CARGA_HORAS"] = resumo_pessoas["CARGA_HORAS"].apply(formatar_ch_ctrl)
        
        st.subheader("📋 Lista de Participantes")
        st.dataframe(
            resumo_pessoas[["STATUS", "NOME", "TREINAMENTOS", "CARGA_HORAS"]],
            use_container_width=True,
            hide_index=True
        )
        
        st.info("""
        📖 **LEGENDA:**
        - ✅ **IMPRESSO** = Coluna STATUS preenchida com "IMPRESSO" na planilha
        - ⬜ **PENDENTE** = Célula vazia na coluna STATUS
        
        📝 **Como atualizar:**
        1. Na sua planilha → coloque **IMPRESSO** na coluna STATUS da pessoa que já imprimiu
        2. Salve o arquivo
        3. Faça upload novamente aqui no Streamlit → os números atualizam automaticamente! ✨
        """)
