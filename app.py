# Site do relatório: o usuário só vê esta página; o código fica no servidor.
import hmac
import os
import streamlit as st
from pipeline_web import listar_clientes, gerar_relatorio, tem_metas
from scripts.opcoes_mapa import GRUPOS, INDICADORES, INDICADORES_POR_GRUPO

st.set_page_config(page_title="Relatório de Qualidade Operacional", page_icon="📊", layout="centered")

# Chave da IA (Gemini) vem dos Secrets e é repassada aos scripts por variável de ambiente.
try:
    if st.secrets.get("GEMINI_API_KEY"):
        os.environ["GEMINI_API_KEY"] = str(st.secrets["GEMINI_API_KEY"])
except Exception:
    pass

# ---- acesso por senha (defina APP_PASSWORD nos Secrets) ----
senha_correta = st.secrets.get("APP_PASSWORD", "")
if not senha_correta:
    st.error("Configure APP_PASSWORD nos Secrets do serviço antes de usar.")
    st.stop()
if not st.session_state.get("ok"):
    senha = st.text_input("Senha de acesso", type="password")
    if senha:
        if hmac.compare_digest(senha, senha_correta):
            st.session_state["ok"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()

st.title("📊 Relatório de Qualidade Operacional")
clientes = listar_clientes()
if not clientes:
    st.warning("Nenhum cliente cadastrado (pasta clientes/ vazia).")
    st.stop()

cliente = st.selectbox("Cliente", clientes)
arquivo = st.file_uploader("Mapa analítico (.csv)", type=["csv"])
enviar = st.checkbox("Enviar o PDF para o Telegram", value=True)

# ---- opcional: comparar com as metas operacionais (metas_operacionais.xlsx na pasta do cliente) ----
metas_ok = tem_metas(cliente)
quer_metas = st.checkbox(
    "Comparar com metas operacionais",
    value=False,
    disabled=not metas_ok,
    help="Usa a planilha metas_operacionais.xlsx da pasta do cliente. Desmarcado, o relatório mostra só o realizado.",
)
comparar_metas = None
if not metas_ok:
    st.caption("Este cliente não tem `metas_operacionais.xlsx` na pasta — a comparação com metas está indisponível.")
elif quer_metas:
    st.caption("Escolha o que considerar (pode marcar um ou os dois):")
    usa_op = st.checkbox("Metas da operação (comparação na 1ª página do processo)", value=True)
    usa_eq = st.checkbox("Metas e parâmetros dos equipamentos (comparação na tabela de equipamentos)", value=True)
    if usa_op or usa_eq:
        comparar_metas = {"operacao": usa_op, "equipamentos": usa_eq}
    else:
        st.warning("Nenhuma opção marcada: o relatório sairá sem comparação com metas.")

# ---- opcional: indicador(es) dos mapas por grupo de operação ----
opcoes_mapa = {}
with st.expander("Opções de mapa (opcional)"):
    st.caption(
        "Escolha, por grupo de operação, qual indicador aparece nos mapas. "
        "Com **1 indicador**: 1 mapa por talhão. Com **2 indicadores**: 2 mapas por talhão "
        "(1 talhão por página, mapas lado a lado). Sem escolher nada, o mapa é automático "
        "(como sempre foi). O grupo de cada processo vem da coluna **grupo** da planilha de "
        "de-para do cliente (aba Processos)."
    )
    for chave, nome in GRUPOS.items():
        escolha = st.multiselect(
            nome,
            options=INDICADORES_POR_GRUPO[chave],
            format_func=lambda k: f"{INDICADORES[k][1]} ({INDICADORES[k][2]})",
            max_selections=2,
            key=f"mapa_{chave}",
            placeholder="Automático",
        )
        if escolha:
            opcoes_mapa[chave] = escolha

if st.button("Gerar relatório", type="primary", disabled=arquivo is None):
    linhas = []
    area_log = st.empty()

    def log(msg):
        linhas.append(msg)

    with st.spinner("Gerando o relatório... (pode levar alguns minutos)"):
        ok, msg, pdf = gerar_relatorio(arquivo.getvalue(), cliente, enviar, log, opcoes_mapa, comparar_metas)
    (st.success if ok else st.error)(msg)
    if pdf:
        st.download_button("Baixar PDF", data=pdf[1], file_name=pdf[0], mime="application/pdf")
    with st.expander("Detalhes da execução"):
        st.code("\n".join(linhas[-200:]) or "(sem saída)")
