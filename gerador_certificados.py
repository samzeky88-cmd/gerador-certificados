import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re

st.set_page_config(page_title="Sistema de Certificados — Coordenadores", layout="wide")

MAX_TREINAMENTOS = 14

# ==================================================
# 🧠 FUNÇÃO DE CORREÇÃO AUTOMÁTICA
# ==================================================
def corrigir_texto(texto):
    """Corrige maiúsculas indevidas, espaços e pontuação automaticamente"""
    if not isinstance(texto, str) or texto.strip() == "":
        return texto
    
    # Remove espaços duplicados e nas pontas
    texto = re.sub(r'\s+', ' ', texto).strip()
    
    # Lista de palavras que SEMPRE iniciam com maiúscula
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
    
    # Lista de palavras que SEMPRE ficam minúsculas no meio da frase
    palavras_minusculas = {
        "a", "o", "e", "de", "da", "do", "das", "dos", "em", "para", "com",
        "sem", "sob", "entre", "por", "pelo", "pela", "nos", "no", "na",
        "os", "as", "um", "uma", "como", "que", "se", "ou", "ao", "à",
        "aos", "às", "é", "e", "mas", "porém", "contudo", "então", "logo"
    }
    
    # Divide em frases e processa cada uma
    frases = re.split(r'([.!?])', texto)
    texto_corrigido = ""
    
    for i in range(0, len(frases), 2):
        frase = frases[i].strip()
        if not frase:
            continue
        pontuacao = frases[i+1] if i+1 < len(frases) else ""
        
        # Divide em palavras
        palavras = frase.split()
        frase_corrigida = []
        
        for idx, palavra in enumerate(palavras):
            p_min = palavra.lower()
            # Primeira palavra da frase ou palavra da lista de maiúsculas
            if idx == 0 or p_min in palavras_maiuscula:
                palavra_corrigida = p_min.capitalize()
            elif p_min in palavras_minusculas:
                palavra_corrigida = p_min
            else:
                # Outras palavras: primeira letra maiúscula, resto minúscula
                palavra_corrigida = p_min.capitalize()
            
            frase_corrigida.append(palavra_corrigida)
        
        texto_corrigido += " ".join(frase_corrigida) + pontuacao
    
    return texto_corrigido


# --- MENU DE ABAS ---
aba_certificados, aba_ranking = st.tabs(["📜 Gerar Certificados", "📊 Ranking de Participação"])

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
        
        nomes = sorted(df["NOME"].dropna().unique())
        nome_escolhido = st.selectbox("👤 Selecione o nome do coordenador", nomes, key="cert_nome")

        if nome_escolhido:
            dados_pessoa = df[df["NOME"] == nome_escolhido].reset_index(drop=True)
            contexto = {"NOME": nome_escolhido}
            
            for i in range(MAX_TREINAMENTOS):
                if i < len(dados_pessoa):
                    linha = dados_pessoa.iloc[i]
                    
                    # Data
                    data_val = linha["DATA"]
                    if hasattr(data_val, 'strftime'):
                        contexto[f"Data_{i+1}"] = data_val.strftime("%d/%m/%Y")
                    else:
                        contexto[f"Data_{i+1}"] = str(data_val)
                    
                    # ✅ TEMA CORRIGIDO AUTOMATICAMENTE
                    tema_bruto = str(linha["TEMA"])
                    contexto[f"Tema_{i+1}"] = corrigir_texto(tema_bruto)
                    
                    # Carga horária formatada
                    ch_valor = float(linha["CH"])
                    if ch_valor.is_integer():
                        contexto[f"Ch_{i+1}"] = f"{int(ch_valor)}H"
                    else:
                        contexto[f"Ch_{i+1}"] = f"{ch_valor:.2f}H".rstrip('0').rstrip('.') + 'H'
                else:
                    contexto[f"Data_{i+1}"] = ""
                    contexto[f"Tema_{i+1}"] = ""
                    contexto[f"Ch_{i+1}"] = ""
            
            # Cálculo da carga total
            ch_total = 0.0
            for i in range(MAX_TREINAMENTOS):
                ch_texto = contexto[f"Ch_{i+1}"]
                if ch_texto:
                    ch_total += float(ch_texto.replace("H", ""))
            
            if ch_total.is_integer():
                contexto["CARGA_TOTAL"] = f"{int(ch_total)}H"
            else:
                contexto["CARGA_TOTAL"] = f"{ch_total:.2f}H".rstrip('0').rstrip('.') + 'H'

            st.subheader("📋 Resumo do participante")
            st.write(f"**Nome:** {nome_escolhido}")
            st.write(f"**Quantidade de treinamentos:** {len(dados_pessoa)}")
            st.write(f"**Carga Horária Total:** {contexto['CARGA_TOTAL']}")
            
            # Prévia da correção
            with st.expander("🔍 Ver temas corrigidos"):
                for i in range(MAX_TREINAMENTOS):
                    tema = contexto[f"Tema_{i+1}"]
                    if tema:
                        st.write(f"{i+1}. {tema}")

            st.subheader("📄 Modelo do certificado (.docx)")
            modelo = st.file_uploader("Escolha o modelo do certificado", type=["docx"], key="cert_modelo")

            if modelo:
                doc = DocxTemplate(modelo)
                doc.render(contexto)
                
                saida = io.BytesIO()
                doc.save(saida)
                saida.seek(0)
                
                st.success("✅ Certificado gerado com correção automática aplicada!")
                st.download_button(
                    label="📥 Baixar Certificado",
                    data=saida,
                    file_name=f"Certificado_{nome_escolhido.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="cert_download"
                )

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
            if valor.is_integer():
                return f"{int(valor)}H"
            else:
                return f"{valor:.2f}H".rstrip('0').rstrip('.') + 'H'
        
        ranking["CARGA_HORARIA_TOTAL"] = ranking["CARGA_HORARIA_TOTAL"].apply(formatar_ch)
        
        st.dataframe(ranking, use_container_width=True, hide_index=True)
        
        st.subheader("📈 Resumo Geral")
        st.write(f"**Total de participantes:** {len(ranking)}")
        st.write(f"**Total de treinamentos realizados:** {ranking['QUANTIDADE_TREINAMENTOS'].sum()}")
        
        st.subheader("🏆 Destaques — Top 10")
        emojis = ["🥇", "🥈", "🥉"] + ["🏅"] * 7
        for pos in range(min(10, len(ranking))):
            emoji = emojis[pos]
            linha = ranking.iloc[pos]
            st.write(f"{emoji} **{linha['POSIÇÃO']}º Lugar:** {linha['NOME']} — {linha['QUANTIDADE_TREINAMENTOS']} treinamentos | {linha['CARGA_HORARIA_TOTAL']}")
