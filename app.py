import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import os
import openpyxl

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS (ESTILO)
# ==========================================
st.set_page_config(page_title="Movimentações - Headcount", layout="wide", initial_sidebar_state="collapsed")

# MÁGICA DO VISUAL: CSS injetado para pintar os fundos e botões
st.markdown("""
<style>
/* 1. Remove os espaços vazios do topo para aproveitar a tela toda (Sem rolagem) */
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 1rem !important;
}

/* 2. Pinta a caixa de SAÍDA de Vermelho Pastel */
.element-container:has(.box-saida) + .element-container div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #ffebee !important;
    border: 1px solid #ef9a9a !important;
}

/* 3. Pinta a caixa de ENTRADA de Verde Pastel */
.element-container:has(.box-entrada) + .element-container div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #e8f5e9 !important;
    border: 1px solid #a5d6a7 !important;
}

/* 4. Pinta o Botão "Confirmar Movimentação" de Verde Forte */
button[kind="primary"] {
    background-color: #2e7d32 !important;
    color: white !important;
    border-color: #2e7d32 !important;
    font-weight: bold !important;
}
button[kind="primary"]:hover {
    background-color: #1b5e20 !important;
    border-color: #1b5e20 !important;
}

/* 5. Pinta o Botão "Não encontrou o posto?" de Laranja/Dourado */
.element-container:has(.btn-posto-marker) + .element-container button {
    background-color: #ff9800 !important;
    color: white !important;
    border: none !important;
    font-weight: bold;
}
.element-container:has(.btn-posto-marker) + .element-container button:hover {
    background-color: #f57c00 !important;
}
</style>
""", unsafe_allow_html=True)

USUARIOS_PERMITIDOS = {
    "admin": "admin123",
    "rh.agricola": "cana2026",
    "analista": "senha123"
}

# ==========================================
# 2. BANCO DE DADOS
# ==========================================
def conectar_banco():
    conn = sqlite3.connect('headcount_v3.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS movimentacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_sistema TEXT, data_registro TEXT, requisitante TEXT,
            unidade_saida TEXT, cc_saida TEXT, subprocesso_saida TEXT,
            gestor_saida TEXT, posto_saida TEXT, cargo_saida TEXT, qtd_saida INTEGER,
            unidade_entrada TEXT, cc_entrada TEXT, subprocesso_entrada TEXT,
            gestor_entrada TEXT, posto_entrada TEXT, cargo_entrada TEXT, qtd_entrada INTEGER
        )
    ''')
    conn.commit()
    return conn

# ==========================================
# 3. LER EXCEL (COM CACHE)
# ==========================================
@st.cache_data
def carregar_dados_excel():
    arquivo_excel = 'parametros.xlsx'
    if not os.path.exists(arquivo_excel):
        st.warning(f"Arquivo '{arquivo_excel}' não encontrado!")
        return pd.DataFrame(columns=['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante'])

    try:
        df = pd.read_excel(arquivo_excel, dtype=str)
        df = df.iloc[:, :7]
        df.columns = ['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante']
        df = df.fillna("") 
        return df
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")
        return pd.DataFrame(columns=['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante'])

df_parametros = carregar_dados_excel()

# ==========================================
# 4. CONTROLE DE SESSÃO
# ==========================================
if 'usuario_logado' not in st.session_state:
    st.session_state.usuario_logado = None
if 'pagina' not in st.session_state:
    st.session_state.pagina = 'login'

def fazer_logout():
    st.session_state.usuario_logado = None
    st.session_state.pagina = 'login'

# ==========================================
# 5. MODAL: SOLICITAR NOVO POSTO
# ==========================================
@st.dialog("Cadastro Posto faltante")
def modal_solicitar_posto():
    st.write("Preencha os dados abaixo.")
    
    und_p = st.selectbox("Unidade:", sorted([x for x in df_parametros['unidade'].unique() if x]), key="p_und")
    df_cc_p = df_parametros[df_parametros['unidade'] == und_p]
    
    cc_p = st.selectbox("Centro de Custo:", sorted([x for x in df_cc_p['cc'].unique() if x]), key="p_cc")
    df_sub_p = df_cc_p[df_cc_p['cc'] == cc_p]
    
    sub_p = st.selectbox("Subprocesso:", sorted([x for x in df_sub_p['sub'].unique() if x]), key="p_sub")
    df_gestor_p = df_sub_p[df_sub_p['sub'] == sub_p]
    
    gestor_p = st.selectbox("Gestor:", sorted([x for x in df_gestor_p['gestor'].unique() if x]), key="p_gestor")
    
    cargo_p = st.selectbox("Qual Cargo deve pertencer a esse posto?:", sorted([x for x in df_parametros['cargo'].unique() if x]), key="p_cargo")

    if st.button("ENVIAR SOLICITAÇÃO AO RH", type="primary"):
        arquivo_solicitacoes = "solicitacoes_postos.xlsx"
        try:
            if os.path.exists(arquivo_solicitacoes):
                wb = openpyxl.load_workbook(arquivo_solicitacoes)
                ws = wb.active
            else:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Solicitações de Postos"
                ws.append(["Data Solicitação", "Usuário", "Unidade", "Centro de Custo", "Subprocesso", "Gestor", "Cargo"])
                
            data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            ws.append([data_atual, st.session_state.usuario_logado, und_p, cc_p, sub_p, gestor_p, cargo_p])
            wb.save(arquivo_solicitacoes)
            
            st.success("Sua solicitação foi enviada com sucesso!")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao salvar solicitação.\nErro: {e}")

# ==========================================
# 6. TELAS DO APLICATIVO
# ==========================================

# --- TELA DE LOGIN ---
if st.session_state.usuario_logado is None:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.write("")
        st.write("")
        try:
            st.image("logo.png", width=150)
        except:
            pass
        st.markdown("<h2 style='text-align: center;'>Movimentações<br>HeadCount</h2>", unsafe_allow_html=True)
        
        with st.form("form_login"):
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            submit = st.form_submit_button("ACESSAR SISTEMA", use_container_width=True)
            
            if submit:
                if usuario in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[usuario] == senha:
                    st.session_state.usuario_logado = usuario
                    st.session_state.pagina = 'registro'
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")

# --- TELAS INTERNAS ---
else:
    col_titulo, col_user, col_btn1, col_btn2 = st.columns([4, 2, 2, 1])
    with col_titulo:
        st.markdown("### Sistema de Movimentações")
    with col_user:
        st.write(f"<br>👤 **Logado como:** {st.session_state.usuario_logado}", unsafe_allow_html=True)
    with col_btn1:
        st.write("<br>", unsafe_allow_html=True)
        if st.session_state.pagina == 'registro':
            if st.button("Ver Histórico (Consultas)", use_container_width=True):
                st.session_state.pagina = 'consulta'
                st.rerun()
        else:
            if st.button("Nova Movimentação", use_container_width=True):
                st.session_state.pagina = 'registro'
                st.rerun()
    with col_btn2:
        st.write("<br>", unsafe_allow_html=True)
        if st.button("Sair", use_container_width=True):
            fazer_logout()
            st.rerun()

    st.divider()

    # --- TELA PRINCIPAL (REGISTRO) ---
    if st.session_state.pagina == 'registro':
        
        lista_req = sorted([x for x in df_parametros['requisitante'].unique() if x])
        requisitante = st.selectbox("Quem solicitou a troca?", options=[""] + lista_req)

        st.write("") 

        # Divisão 50/50 da tela para as duas caixas
        col_saida, col_entrada = st.columns(2, gap="large")

        # ==== LADO ESQUERDO: SAÍDA ====
        with col_saida:
            # Marcador invisível para o CSS pintar o fundo de Vermelho
            st.markdown("<div class='box-saida'></div>", unsafe_allow_html=True)
            
            with st.container(border=True):
                st.markdown("<h4 style='text-align: center; color: #c62828;'>VAGA DE SAÍDA (RETIRADA)</h4>", unsafe_allow_html=True)
                
                # DIVISÃO INTERNA: Coloca as caixas lado a lado para economizar altura!
                s_c1, s_c2 = st.columns(2)
                
                with s_c1:
                    s_und = st.selectbox("Unidade (Saída):", options=[""] + sorted([x for x in df_parametros['unidade'].unique() if x]))
                    df_s_cc = df_parametros[df_parametros['unidade'] == s_und] if s_und else df_parametros
                    
                    s_cc = st.selectbox("Centro de Custo (Saída):", options=[""] + sorted([x for x in df_s_cc['cc'].unique() if x]))
                    df_s_sub = df_s_cc[df_s_cc['cc'] == s_cc] if s_cc else df_s_cc
                    
                    s_sub = st.selectbox("Subprocesso (Saída):", options=[""] + sorted([x for x in df_s_sub['sub'].unique() if x]))
                    
                    s_qtd = st.number_input("Quantidade (Saída):", min_value=1, value=1, step=1)

                with s_c2:
                    df_s_gestor = df_s_sub[df_s_sub['sub'] == s_sub] if s_sub else df_s_sub
                    s_gestor = st.selectbox("Gestor (Saída):", options=[""] + sorted([x for x in df_s_gestor['gestor'].unique() if x]))
                    
                    df_s_posto = df_s_gestor[df_s_gestor['gestor'] == s_gestor] if s_gestor else df_s_gestor
                    s_posto = st.selectbox("Posto (Saída):", options=[""] + sorted([x for x in df_s_posto['posto'].unique() if x]))
                    
                    df_s_cargo = df_s_posto[df_s_posto['posto'] == s_posto] if s_posto else df_s_posto
                    s_cargo = st.selectbox("Cargo (Saída):", options=[""] + sorted([x for x in df_s_cargo['cargo'].unique() if x]))

        # ==== LADO DIREITO: ENTRADA ====
        with col_entrada:
            # Marcador invisível para o CSS pintar o fundo de Verde
            st.markdown("<div class='box-entrada'></div>", unsafe_allow_html=True)
            
            with st.container(border=True):
                st.markdown("<h4 style='text-align: center; color: #2e7d32;'>VAGA DE ENTRADA (NOVA)</h4>", unsafe_allow_html=True)
                
                # DIVISÃO INTERNA
                e_c1, e_c2 = st.columns(2)
                
                with e_c1:
                    e_und = st.selectbox("Unidade (Entrada):", options=[""] + sorted([x for x in df_parametros['unidade'].unique() if x]))
                    df_e_cc = df_parametros[df_parametros['unidade'] == e_und] if e_und else df_parametros
                    
                    e_cc = st.selectbox("Centro de Custo (Entrada):", options=[""] + sorted([x for x in df_e_cc['cc'].unique() if x]))
                    df_e_sub = df_e_cc[df_e_cc['cc'] == e_cc] if e_cc else df_e_cc
                    
                    e_sub = st.selectbox("Subprocesso (Entrada):", options=[""] + sorted([x for x in df_e_sub['sub'].unique() if x]))
                    
                    e_qtd = st.number_input("Quantidade (Entrada):", min_value=1, value=1, step=1)

                with e_c2:
                    df_e_gestor = df_e_sub[df_e_sub['sub'] == e_sub] if e_sub else df_e_sub
                    e_gestor = st.selectbox("Gestor (Entrada):", options=[""] + sorted([x for x in df_e_gestor['gestor'].unique() if x]))
                    
                    df_e_posto = df_e_gestor[df_e_gestor['gestor'] == e_gestor] if e_gestor else df_e_gestor
                    e_posto = st.selectbox("Posto (Entrada):", options=[""] + sorted([x for x in df_e_posto['posto'].unique() if x]))
                    
                    df_e_cargo = df_e_posto[df_e_posto['posto'] == e_posto] if e_posto else df_e_posto
                    e_cargo = st.selectbox("Cargo (Entrada):", options=[""] + sorted([x for x in df_e_cargo['cargo'].unique() if x]))

                # BOTÃO DE SOLICITAR POSTO NOVO (Fica embaixo das colunas dentro do quadro verde)
                st.write("")
                st.markdown("<div class='btn-posto-marker'></div>", unsafe_allow_html=True)
                if st.button("Não encontrou o posto? Clique aqui para solicitar", use_container_width=True):
                    modal_solicitar_posto()

        st.write("")
        
        # ==== BOTÃO SALVAR (CÓDIGO TYPE="PRIMARY" o CSS PINTA DE VERDE) ====
        if st.button("CONFIRMAR MOVIMENTAÇÃO", type="primary", use_container_width=True):
            if not requisitante:
                st.warning("O campo Requisitante é obrigatório.")
            elif not all([s_und, s_cc, s_sub, s_gestor, s_posto, s_cargo, e_und, e_cc, e_sub, e_gestor, e_posto, e_cargo]):
                st.warning("Preencha todas as caixas de Saída e Entrada antes de salvar.")
            else:
                conn = conectar_banco()
                cursor = conn.cursor()
                data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                cursor.execute('''
                    INSERT INTO movimentacoes (
                        usuario_sistema, data_registro, requisitante,
                        unidade_saida, cc_saida, subprocesso_saida, gestor_saida, posto_saida, cargo_saida, qtd_saida,
                        unidade_entrada, cc_entrada, subprocesso_entrada, gestor_entrada, posto_entrada, cargo_entrada, qtd_entrada
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    st.session_state.usuario_logado, data_atual, requisitante,
                    s_und, s_cc, s_sub, s_gestor, s_posto, s_cargo, s_qtd,
                    e_und, e_cc, e_sub, e_gestor, e_posto, e_cargo, e_qtd
                ))
                conn.commit()

                try:
                    df_espelho = pd.read_sql_query("SELECT * FROM movimentacoes", conn)
                    df_espelho.to_excel('base_powerbi.xlsx', index=False, engine='openpyxl')
                except Exception as e:
                    st.error(f"Erro ao gerar espelho Excel: {e}")

                conn.close()
                st.success("Movimentação registrada com sucesso!")

    # --- TELA DE CONSULTA ---
    elif st.session_state.pagina == 'consulta':
        conn = conectar_banco()
        # SELEÇÃO ATUALIZADA: Puxando o CC de Saída e CC de Entrada
        df_historico = pd.read_sql_query(f'''
            SELECT id, data_registro, requisitante, 
                   cc_saida, qtd_saida, cargo_saida, 
                   cc_entrada, qtd_entrada, cargo_entrada 
            FROM movimentacoes 
            WHERE usuario_sistema = '{st.session_state.usuario_logado}' 
            ORDER BY id DESC
        ''', conn)
        conn.close()
        
        # Renomeando as colunas para o grid ficar mais legível
        df_historico.columns = ["ID", "Data", "Requisitante", "CC Saída", "Qtd Saída", "Cargo Saída", "CC Entrada", "Qtd Entrada", "Cargo Entrada"]

        total = len(df_historico)
        ultima = df_historico['Data'].iloc[0] if total > 0 else "-"

        col_metric1, col_metric2 = st.columns(2)
        col_metric1.metric("TOTAL REGISTRADO", total)
        col_metric2.metric("ÚLTIMA MOVIMENTAÇÃO", ultima)

        st.write("### Suas Movimentações Cadastradas")
        st.dataframe(df_historico, use_container_width=True, hide_index=True)
