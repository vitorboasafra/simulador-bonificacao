import streamlit as st
import json
import os

USER_DB = "usuarios.json"

def carregar_usuarios():
    if os.path.exists(USER_DB):
        with open(USER_DB, "r", encoding="utf-8") as f: 
            return json.load(f)
    return {"admin": "admin"} # Usuário mestre de segurança

def salvar_usuario(usuario, senha):
    users = carregar_usuarios()
    users[usuario] = senha
    with open(USER_DB, "w", encoding="utf-8") as f: 
        json.dump(users, f, indent=4)

def mostrar_login():
    st.markdown('<br><br>', unsafe_allow_html=True)
    _, col_center, _ = st.columns([0.15, 0.7, 0.15])

    with col_center:
        st.markdown('<div class="main-card">', unsafe_allow_html=True)
        
        if os.path.exists("logo.png"):
            st.image("logo.png", width=240)
        else:
            st.markdown("<h1 style='text-align: center; color: #8BC34A;'>🌾 BOA SAFRA</h1>", unsafe_allow_html=True)

        st.markdown('<p class="sub-header">Simulador de Bonificação</p>', unsafe_allow_html=True)
        
        aba_login, aba_cad = st.tabs(["🔐 Acessar", "📝 Novo Registro"])

        with aba_login:
            usuario = st.text_input("Usuário", placeholder="Digite seu usuário", key="log_u")
            senha = st.text_input("Senha", type="password", placeholder="Digite sua senha", key="log_p")
            if st.button("🔓 ENTRAR"):
                users = carregar_usuarios()
                if usuario in users and users[usuario] == senha:
                    st.session_state.logado = True
                    st.session_state.user = usuario
                    st.rerun()
                else:
                    st.error("❌ Usuário ou senha incorretos.")

        with aba_cad:
            novo_user = st.text_input("Novo Usuário", placeholder="Ex: vitor.matias", key="cad_u")
            nova_senha = st.text_input("Nova Senha", type="password", placeholder="Mínimo 4 caracteres", key="cad_p")
            if st.button("✅ CADASTRAR"):
                if novo_user and len(nova_senha) >= 4:
                    salvar_usuario(novo_user, nova_senha)
                    st.success("✅ Conta criada! Volte na aba 'Acessar' para entrar.")
                else:
                    st.warning("⚠️ Preencha corretamente (senha mínima 4 caracteres).")

        st.markdown('<p class="footer-text">Desenvolvido por Vitor Matias</p>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
