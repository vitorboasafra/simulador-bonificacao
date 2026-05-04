import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime

def gerar_relatorio_pdf(dados):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(15, 15, 15)

    # Cabeçalho
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(139, 195, 74)
    pdf.cell(0, 15, "RELATÓRIO DE BONIFICAÇÃO", ln=True, align='C')
    pdf.ln(2)
    pdf.set_draw_color(224, 227, 232)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(5)

    # Info Geral
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(45, 52, 54)
    pdf.cell(40, 8, "Produtor (Nome):", 0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, dados['produtor_nome'], ln=True)
    
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(40, 8, "Operador:", 0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, dados['usuario'], ln=True)
    
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(40, 8, "Data Emissao:", 0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, datetime.now().strftime("%d/%m/%Y %H:%M"), ln=True)
    pdf.ln(5)

    # Detalhamento Analítico
    for item in dados['detalhes']:
        # Alterado de 'Inscricao' para 'Codigo'
        pdf.set_fill_color(241, 248, 233)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(0, 10, f" Centro: {item['centro']} | Codigo: {item['inscricao']} | Material: {item['cultivar']}", ln=True, fill=True)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(0, 5, f" Notas Fiscais: {item['nfs']}")
        pdf.set_text_color(45, 52, 54)
        
        pdf.set_font("Arial", '', 9)
        pdf.cell(45, 7, f" Peso Base: {item['peso_base']} kg", 0)
        pdf.cell(45, 7, f" Bonif.: {item['perc']}%", 0)
        pdf.cell(45, 7, f" Peso Bonif.: {item['peso_bonif']} kg", ln=True)
        
        pdf.cell(45, 7, f" Qtd. Sacas: {item['sacas']}", 0)
        pdf.cell(45, 7, f" Vlr Saca: R$ {item['vlr_saca']}", 0)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(45, 7, f" Subtotal: R$ {item['subtotal']}", ln=True)
        pdf.ln(3)

    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.set_fill_color(240, 248, 240)
    pdf.cell(0, 15, f"VALOR TOTAL GERAL: R$ {dados['total_geral']}", ln=True, align='C', fill=True)

    return pdf.output(dest='S').encode('latin-1')

def mostrar_painel():
    st.markdown("""<style>[data-testid="collapsedControl"] {display: none;}</style>""", unsafe_allow_html=True)
    st.markdown('<div class="main-card">', unsafe_allow_html=True)
    
    col_user, col_out = st.columns([0.8, 0.2])
    with col_user: 
        st.markdown(f"👤 Operador: **{st.session_state.user}**")
    with col_out: 
        if st.button("Sair"): 
            st.session_state.logado = False
            st.rerun()

    st.markdown("<h2 style='text-align: center;'>🌾 Simulador de Bonificação</h2><br>", unsafe_allow_html=True)
    
    try:
        df = pd.read_excel('zmm113.xlsx')
        
        def limpar_inscricao(valor):
            try:
                return str(int(float(valor)))
            except:
                return str(valor)
                
        df['Fornecedor'] = df['Fornecedor'].apply(limpar_inscricao)

        produtor_nome = st.selectbox("1. Selecione o Fornecedor", df['Nome (Fornecedor)'].unique(), index=None, placeholder="Busque pelo nome...")
        
        if produtor_nome:
            df_prod = df[df['Nome (Fornecedor)'] == produtor_nome]
            centros = st.multiselect("2. Selecione os Centros", df_prod['Centro'].unique())
            
            if centros:
                df_centros = df_prod[df_prod['Centro'].isin(centros)]
                cultivares = st.multiselect("3. Selecione as Cultivares", df_centros['Descrição (Material)'].unique())
                
                if cultivares:
                    df_final = df_centros[df_centros['Descrição (Material)'].isin(cultivares)]
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown('<div style="background: #F1F8E9; border-left: 4px solid #8BC34A; padding: 15px; border-radius: 8px;">', unsafe_allow_html=True)
                    # Nomeclatura alterada para 'Código'
                    st.markdown("<h4 style='color: #558B2F; margin:0;'>⚖️ Resumo por Centro e Código</h4>", unsafe_allow_html=True)
                    
                    for c in centros:
                        df_c = df_final[df_final['Centro'] == c]
                        if not df_c.empty:
                            st.markdown(f"<p style='margin: 10px 0 0 0;'>📍 <strong>Centro {c}:</strong></p>", unsafe_allow_html=True)
                            for mat in df_c['Descrição (Material)'].unique():
                                df_m = df_c[df_c['Descrição (Material)'] == mat]
                                for insc in df_m['Fornecedor'].unique():
                                    p_final = df_m[df_m['Fornecedor'] == insc]['Peso Liquido Kg'].sum()
                                    # Alterado de 'Inscrição' para 'Código'
                                    st.markdown(f"<p style='margin:0; font-size: 14px; padding-left: 20px;'>• {mat} (Código: {insc}): <strong>{p_final:,.2f} kg</strong></p>".replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
                    
                    st.markdown('</div>', unsafe_allow_html=True)

                    st.markdown("<br>#### ⚙️ Parâmetros de Bonificação", unsafe_allow_html=True)
                    modo_global = st.toggle("Aplicar os mesmos valores para todos os códigos/materiais", value=True)
                    
                    configs = {}
                    combinacoes = df_final[['Centro', 'Descrição (Material)', 'Fornecedor']].drop_duplicates().values
                    
                    if modo_global:
                        col1, col2 = st.columns(2)
                        p_g = col1.number_input("Porcentagem (%)", value=10.0, step=0.1)
                        v_g = col2.number_input("Valor da Saca (R$)", value=0.0, step=1.0)
                        for c, m, f in combinacoes:
                            configs[(c, m, f)] = {'perc': p_g, 'val': v_g}
                    else:
                        for c, m, f in combinacoes:
                            # Alterado de 'Insc' para 'Cod'
                            st.markdown(f"**Centro {c} | {m} | Cod: {f}**")
                            col1, col2 = st.columns(2)
                            p = col1.number_input(f"% Bonif.", value=10.0, step=0.1, key=f"p_{c}_{m}_{f}")
                            v = col2.number_input(f"R$ Saca", value=0.0, step=1.0, key=f"v_{c}_{m}_{f}")
                            configs[(c, m, f)] = {'perc': p, 'val': v}

                    if st.button("CALCULAR", use_container_width=True):
                        total_geral = 0
                        detalhes_pdf = []
                        
                        for (c_id, m_id, f_id), val in configs.items():
                            df_item = df_final[(df_final['Centro'] == c_id) & 
                                               (df_final['Descrição (Material)'] == m_id) & 
                                               (df_final['Fornecedor'] == f_id)]
                            
                            peso_b = df_item['Peso Liquido Kg'].sum()
                            nfs_sujas = df_item['Nº.da NF do Produtor'].dropna().unique()
                            nfs_limpas = [str(int(float(nf))) if str(nf).replace('.','',1).isdigit() else str(nf) for nf in nfs_sujas]
                            nfs = ", ".join(nfs_limpas)
                            
                            p_bonif = peso_b * (val['perc'] / 100)
                            sacas = p_bonif / 60
                            subtotal = sacas * val['val']
                            total_geral += subtotal
                            
                            detalhes_pdf.append({
                                "centro": c_id,
                                "cultivar": m_id,
                                "inscricao": f_id, # Internamente mantido como 'inscricao', mas exibido como 'Codigo' no PDF
                                "nfs": nfs if nfs else "N/A",
                                "peso_base": f"{peso_b:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                "perc": f"{val['perc']}",
                                "peso_bonif": f"{p_bonif:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                "sacas": f"{sacas:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                "vlr_saca": f"{val['val']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                "subtotal": f"{subtotal:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            })
                        
                        st.success(f"### Total Geral: R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                        
                        pdf_bytes = gerar_relatorio_pdf({
                            "usuario": st.session_state.user,
                            "produtor_nome": produtor_nome,
                            "total_geral": f"{total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                            "detalhes": detalhes_pdf
                        })
                        
                        # Nome do botão alterado para 'Baixar PDF'
                        st.download_button("📥 Baixar PDF", pdf_bytes, f"Bonif_{produtor_nome[:10]}.pdf", "application/pdf", use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
    st.markdown('</div>', unsafe_allow_html=True)
