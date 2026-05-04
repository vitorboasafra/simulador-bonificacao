import streamlit as st
from login import mostrar_login
from painel import mostrar_painel

# 1. CONFIGURAÇÃO DA PÁGINA (Sempre a primeira linha)
st.set_page_config(page_title="Simulador de Bonificação", page_icon="🌾", layout="centered")

# 2. APLICAR CSS GLOBAL
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    header { visibility: hidden; display: none; }
    .stAppHeader { display: none; }
    footer { visibility: hidden; }
    .stApp { background: linear-gradient(135deg, #F0F2F6 0%, #E8EAEF 100%); font-family: 'Inter', sans-serif; }
    .main-card { background-color: #FFFFFF; padding: 30px 40px; border-radius: 20px; box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08); color: #2D3436; border: 1px solid #E8EAEF; margin-top: 20px; }
    h1, h2, h3 { font-family: 'Inter', sans-serif !important; font-weight: 700 !important; color: #1A1A1A !important; }
    div.stButton > button { width: 100%; border-radius: 12px !important; background-color: #8BC34A !important; color: white !important; height: 52px !important; font-weight: 700 !important; border: none !important; margin-top: 10px; box-shadow: 0 4px 12px rgba(139, 195, 74, 0.3); transition: 0.3s;}
    div.stButton > button:hover { background-color: #7CB342 !important; box-shadow: 0 6px 20px rgba(139, 195, 74, 0.4) !important;}
    .footer-text { margin-top: 40px; font-size: 13px; color: #8B92A1; font-weight: 500; border-top: 1px solid #E8EAEF; padding-top: 20px; text-align: center; }
    .sub-header { font-size: 22px; font-weight: 600; color: #7A8C8E; margin-bottom: 30px; text-align: center; }
    
    /* Inputs */
    .stTextInput > div > div > input, .stSelectbox > div > div > div, .stMultiSelect > div > div > div, .stNumberInput > div > div > input {
        background-color: #F8F9FA !important; border-radius: 12px !important; border: 1.5px solid #E0E3E8 !important; font-size: 16px !important;
    }
    .stTextInput > div > div > input:focus, .stSelectbox > div > div > div:focus, .stMultiSelect > div > div > div:focus {
        border-color: #8BC34A !important; box-shadow: 0 0 0 3px rgba(139, 195, 74, 0.1) !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. ESTADOS DE SESSÃO
if 'logado' not in st.session_state: st.session_state.logado = False
if 'user' not in st.session_state: st.session_state.user = ""

# 4. LÓGICA DE NAVEGAÇÃO
if not st.session_state.logado:
    mostrar_login()
else:
    mostrar_painel()
