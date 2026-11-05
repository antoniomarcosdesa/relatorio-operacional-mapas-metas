# Site do relatório: o usuário só vê esta página; o código fica no servidor.
import hmac
import os
import streamlit as st
from pipeline_web import listar_clientes, gerar_relatorio

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

if st.button("Gerar relatório", type="primary", disabled=arquivo is None):
    linhas = []
    area_log = st.empty()

    def log(msg):
        linhas.append(msg)

    with st.spinner("Gerando o relatório... (pode levar alguns minutos)"):
        ok, msg, pdf = gerar_relatorio(arquivo.getvalue(), cliente, enviar, log)
    (st.success if ok else st.error)(msg)
    if pdf:
        st.download_button("Baixar PDF", data=pdf[1], file_name=pdf[0], mime="application/pdf")
    with st.expander("Detalhes da execução"):
        st.code("\n".join(linhas[-200:]) or "(sem saída)")
