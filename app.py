import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import os
import openpyxl
import time
import base64

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E DESIGN GLOBAL
# ==========================================
st.set_page_config(page_title="Movimentações - Headcount", layout="wide", initial_sidebar_state="collapsed")

# Injeção de CSS para o fundo da página, sombras e layout compacto
st.markdown("""
<style>
/* Fundo suave para a página inteira sair do branco básico */
.stApp {
    background-color: #f4f7f6;
}

/* Espaçamento do cabeçalho otimizado */
.block-container {
    padding-top: 3rem !important; 
    padding-bottom: 2rem !important;
    max-width: 95% !important;
}

/* Deixa os campos mais juntinhos para caber tudo na tela */
div[data-testid="stVerticalBlock"] {
    gap: 0.1rem !important;
}

/* Estilo para as caixas ganharem uma sombra e ficarem elegantes */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: white;
    border-radius: 12px !important;
    box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.05) !important;
    border: none !important;
}

/* PINTURA DA CAIXA SAÍDA (Vermelho Pastel) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(.bg-saida) {
    background-color: #fff5f5 !important;
    border: 1px solid #ffcdd2 !important;
}

/* PINTURA DA CAIXA ENTRADA (Verde Pastel) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(.bg-entrada) {
    background-color: #f1f8e9 !important;
    border: 1px solid #c8e6c9 !important;
}

/* ESTILO DO BOTÃO CONFIRMAR (VERDE) */
div[data-testid="stElementContainer"]:has(.btn-confirmar) + div[data-testid="stElementContainer"] button {
    background-color: #2e7d32 !important;
    color: white !important;
    font-size: 18px !important;
    font-weight: bold !important;
    height: 55px !important;
    border-radius: 8px !important;
    border: none !important;
    box-shadow: 0px 4px 6px rgba(46, 125, 50, 0.3) !important;
}
div[data-testid="stElementContainer"]:has(.btn-confirmar) + div[data-testid="stElementContainer"] button:hover {
    background-color: #1b5e20 !important;
}

/* ESTILO DO BOTÃO POSTO FALTANTE (LARANJA) */
div[data-testid="stElementContainer"]:has(.btn-posto) + div[data-testid="stElementContainer"] button {
    background-color: #ff9800 !important;
    color: white !important;
    font-weight: bold !important;
    border: none !important;
    border-radius: 6px !important;
}
div[data-testid="stElementContainer"]:has(.btn-posto) + div[data-testid="stElementContainer"] button:hover {
    background-color: #e68a00 !important;
}
</style>
""", unsafe_allow_html=True)

USUARIOS_PERMITIDOS = {
    "admin": "admin123",
    "rh.agricola": "cana2026",
    "analista": "senha123"
}

# ==========================================
# 2. BANCO DE DADOS E LER EXCEL
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

@st.cache_data
def carregar_dados_excel():
    arquivo_excel = 'parametros.xlsx'
    if not os.path.exists(arquivo_excel):
        return pd.DataFrame(columns=['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante'])
    try:
        df = pd.read_excel(arquivo_excel, dtype=str)
        df = df.iloc[:, :7]
        df.columns = ['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante']
        df = df.fillna("") 
        return df
    except:
        return pd.DataFrame(columns=['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante'])

df_parametros = carregar_dados_excel()

def renderizar_logo(tamanho=180):
    if os.path.exists("logo.png"):
        with open("logo.png", "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode()
            st.markdown(f'<div style="text-align: center;"><img src="data:image/png;base64,{encoded}" width="{tamanho}"></div>', unsafe_allow_html=True)

# ==========================================
# 3. SESSÃO E LOGIN
# ==========================================
if 'usuario_logado' not in st.session_state:
    st.session_state.usuario_logado = None
if 'pagina' not in st.session_state:
    st.session_state.pagina = 'login'

def fazer_logout():
    st.session_state.usuario_logado = None
    st.session_state.pagina = 'login'

# ==========================================
# 4. MODAL: CADASTRAR POSTO FALTANTE
# ==========================================
@st.dialog("Cadastro Posto faltante")
def modal_solicitar_posto():
    st.markdown("Preencha os dados abaixo.")
    
    # O index=None garante que a caixa nasça totalmente em branco!
    und_p = st.selectbox("Unidade:", options=sorted([x for x in df_parametros['unidade'].unique() if x]), index=None, placeholder="Selecione...")
    
    # Prevenção: só cria o filtro de CC se a Unidade já foi preenchida
    df_cc_p = df_parametros[df_parametros['unidade'] == und_p] if und_p else pd.DataFrame(columns=df_parametros.columns)
    cc_p = st.selectbox("Centro de Custo:", options=sorted([x for x in df_cc_p['cc'].unique() if x]), index=None, placeholder="Selecione...")
    
    df_sub_p = df_cc_p[df_cc_p['cc'] == cc_p] if cc_p else pd.DataFrame(columns=df_parametros.columns)
    sub_p = st.selectbox("Subprocesso:", options=sorted([x for x in df_sub_p['sub'].unique() if x]), index=None, placeholder="Selecione...")
    
    df_gestor_p = df_sub_p[df_sub_p['sub'] == sub_p] if sub_p else pd.DataFrame(columns=df_parametros.columns)
    gestor_p = st.selectbox("Gestor:", options=sorted([x for x in df_gestor_p['gestor'].unique() if x]), index=None, placeholder="Selecione...")
    
    cargo_p = st.selectbox("Qual Cargo deve pertencer a esse posto?:", options=sorted([x for x in df_parametros['cargo'].unique() if x]), index=None, placeholder="Selecione...")

    st.write("")
    if st.button("ENVIAR SOLICITAÇÃO AO RH", use_container_width=True):
        if not all([und_p, cc_p, sub_p, gestor_p, cargo_p]):
            st.error("Por favor, preencha todos os campos antes de enviar.")
        else:
            arquivo_solicitacoes = "solicitacoes_postos.xlsx"
            try:
                if os.path.exists(arquivo_solicitacoes):
                    wb = openpyxl.load_workbook(arquivo_solicitacoes)
                    ws = wb.active
                else:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.append(["Data Solicitação", "Usuário", "Unidade", "Centro de Custo", "Subprocesso", "Gestor", "Cargo"])
                    
                data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                ws.append([data_atual, st.session_state.usuario_logado, und_p, cc_p, sub_p, gestor_p, cargo_p])
                wb.save(arquivo_solicitacoes)
                
                # Exibe a mensagem, espera 1.5 seg pra pessoa ler, e fecha o modal
                st.success("✅ Solicitação salva com sucesso!")
                time.sleep(1.5)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar solicitação. Feche o arquivo se estiver aberto no seu PC.\nErro: {e}")

# ==========================================
# 5. TELAS DO APLICATIVO
# ==========================================

# --- TELA DE LOGIN ---
if st.session_state.usuario_logado is None:
    st.write("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.5, 1.2, 1.5]) 
    with col2:
        with st.container(border=True):
            st.write("<br>", unsafe_allow_html=True)
            renderizar_logo(160) 
            st.markdown("<h3 style='text-align: center; color: #1f538d;'>Movimentações<br>HeadCount</h3>", unsafe_allow_html=True)
            
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            
            st.write("<br>", unsafe_allow_html=True)
            if st.button("ACESSAR SISTEMA", type="primary", use_container_width=True):
                if usuario in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[usuario] == senha:
                    st.session_state.usuario_logado = usuario
                    st.session_state.pagina = 'registro'
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")

# --- TELAS INTERNAS ---
else:
    # CABEÇALHO
    col_titulo, col_user, col_btn1, col_btn2 = st.columns([4, 2.5, 2, 1])
    with col_titulo:
        st.markdown("<h3 style='color: #1f538d; margin-top: -10px;'>Sistema de Movimentações</h3>", unsafe_allow_html=True)
    with col_user:
        st.write(f"👤 Logado como: **{st.session_state.usuario_logado}**")
    with col_btn1:
        if st.session_state.pagina == 'registro':
            if st.button("Ver Histórico (Consultas)", use_container_width=True):
                st.session_state.pagina = 'consulta'
                st.rerun()
        else:
            if st.button("Nova Movimentação", use_container_width=True):
                st.session_state.pagina = 'registro'
                st.rerun()
    with col_btn2:
        if st.button("Sair", use_container_width=True):
            fazer_logout()
            st.rerun()

    st.divider()

    # --- TELA PRINCIPAL (REGISTRO) ---
    if st.session_state.pagina == 'registro':
        
        lista_req = sorted([x for x in df_parametros['requisitante'].unique() if x])
        
        # Inicia vazio também para forçar a pessoa a pesquisar
        requisitante = st.selectbox("Quem solicitou a troca? (Digite para pesquisar)", options=lista_req, index=None, placeholder="Selecione o requisitante...")

        st.write("") 

        # Divisão 50/50
        col_saida, col_entrada = st.columns(2, gap="large")

        # ==== LADO ESQUERDO: SAÍDA ====
        with col_saida:
            with st.container(border=True):
                st.markdown("<div class='bg-saida'></div>", unsafe_allow_html=True)
                st.markdown("<h4 style='text-align: center; color: #c62828;'>VAGA DE SAÍDA (RETIRADA)</h4>", unsafe_allow_html=True)
                
                # index=None para os campos começarem zerados
                s_und = st.selectbox("Unidade (Saída):", options=sorted([x for x in df_parametros['unidade'].unique() if x]), index=None, placeholder="")
                df_s_cc = df_parametros[df_parametros['unidade'] == s_und] if s_und else pd.DataFrame(columns=df_parametros.columns)
                
                s_cc = st.selectbox("Centro de Custo (Saída):", options=sorted([x for x in df_s_cc['cc'].unique() if x]), index=None, placeholder="")
                df_s_sub = df_s_cc[df_s_cc['cc'] == s_cc] if s_cc else pd.DataFrame(columns=df_parametros.columns)
                
                s_sub = st.selectbox("Subprocesso (Saída):", options=sorted([x for x in df_s_sub['sub'].unique() if x]), index=None, placeholder="")
                df_s_gestor = df_s_sub[df_s_sub['sub'] == s_sub] if s_sub else pd.DataFrame(columns=df_parametros.columns)
                
                s_gestor = st.selectbox("Gestor (Saída):", options=sorted([x for x in df_s_gestor['gestor'].unique() if x]), index=None, placeholder="")
                df_s_posto = df_s_gestor[df_s_gestor['gestor'] == s_gestor] if s_gestor else pd.DataFrame(columns=df_parametros.columns)
                
                s_posto = st.selectbox("Posto (Saída):", options=sorted([x for x in df_s_posto['posto'].unique() if x]), index=None, placeholder="")
                df_s_cargo = df_s_posto[df_s_posto['posto'] == s_posto] if s_posto else pd.DataFrame(columns=df_parametros.columns)
                
                s_cargo = st.selectbox("Cargo (Saída):", options=sorted([x for x in df_s_cargo['cargo'].unique() if x]), index=None, placeholder="")
                
                s_qtd = st.number_input("Quantidade (Saída):", min_value=1, value=1, step=1)

        # ==== LADO DIREITO: ENTRADA ====
        with col_entrada:
            with st.container(border=True):
                st.markdown("<div class='bg-entrada'></div>", unsafe_allow_html=True)
                st.markdown("<h4 style='text-align: center; color: #2e7d32;'>VAGA DE ENTRADA (NOVA)</h4>", unsafe_allow_html=True)
                
                e_und = st.selectbox("Unidade (Entrada):", options=sorted([x for x in df_parametros['unidade'].unique() if x]), index=None, placeholder="")
                df_e_cc = df_parametros[df_parametros['unidade'] == e_und] if e_und else pd.DataFrame(columns=df_parametros.columns)
                
                e_cc = st.selectbox("Centro de Custo (Entrada):", options=sorted([x for x in df_e_cc['cc'].unique() if x]), index=None, placeholder="")
                df_e_sub = df_e_cc[df_e_cc['cc'] == e_cc] if e_cc else pd.DataFrame(columns=df_parametros.columns)
                
                e_sub = st.selectbox("Subprocesso (Entrada):", options=sorted([x for x in df_e_sub['sub'].unique() if x]), index=None, placeholder="")
                df_e_gestor = df_e_sub[df_e_sub['sub'] == e_sub] if e_sub else pd.DataFrame(columns=df_parametros.columns)
                
                e_gestor = st.selectbox("Gestor (Entrada):", options=sorted([x for x in df_e_gestor['gestor'].unique() if x]), index=None, placeholder="")
                df_e_posto = df_e_gestor[df_e_gestor['gestor'] == e_gestor] if e_gestor else pd.DataFrame(columns=df_parametros.columns)
                
                e_posto = st.selectbox("Posto (Entrada):", options=sorted([x for x in df_e_posto['posto'].unique() if x]), index=None, placeholder="")
                df_e_cargo = df_e_posto[df_e_posto['posto'] == e_posto] if e_posto else pd.DataFrame(columns=df_parametros.columns)
                
                e_cargo = st.selectbox("Cargo (Entrada):", options=sorted([x for x in df_e_cargo['cargo'].unique() if x]), index=None, placeholder="")
                
                e_qtd = st.number_input("Quantidade (Entrada):", min_value=1, value=1, step=1)
                
                st.write("")
                st.markdown("<div class='btn-posto'></div>", unsafe_allow_html=True)
                if st.button("Não encontrou o posto? Clique aqui para solicitar", use_container_width=True):
                    modal_solicitar_posto()

        st.write("")
        
        # ==== BOTÃO SALVAR ====
        st.markdown("<div class='btn-confirmar'></div>", unsafe_allow_html=True)
        if st.button("CONFIRMAR MOVIMENTAÇÃO", use_container_width=True):
            if not requisitante:
                st.warning("⚠️ O campo Requisitante é obrigatório.")
            elif not all([s_und, s_cc, s_sub, s_gestor, s_posto, s_cargo, e_und, e_cc, e_sub, e_gestor, e_posto, e_cargo]):
                st.warning("⚠️ Preencha todas as caixas de Saída e Entrada antes de salvar.")
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
                st.success("✅ Movimentação registrada com sucesso!")
                time.sleep(1.5) # Dá tempo de ver a mensagem de sucesso
                st.rerun()      # Limpa os campos da tela principal

    # --- TELA DE CONSULTA ---
    elif st.session_state.pagina == 'consulta':
        conn = conectar_banco()
        df_historico = pd.read_sql_query(f'''
            SELECT id, data_registro, requisitante, 
                   cc_saida, qtd_saida, cargo_saida, 
                   cc_entrada, qtd_entrada, cargo_entrada 
            FROM movimentacoes 
            WHERE usuario_sistema = '{st.session_state.usuario_logado}' 
            ORDER BY id DESC
        ''', conn)
        conn.close()
        
        df_historico.columns = ["ID", "Data", "Requisitante", "CC Saída", "Qtd Saída", "Cargo Saída", "CC Entrada", "Qtd Entrada", "Cargo Entrada"]

        total = len(df_historico)
        ultima = df_historico['Data'].iloc[0] if total > 0 else "-"

        col_metric1, col_metric2 = st.columns(2)
        col_metric1.metric("TOTAL REGISTRADO", total)
        col_metric2.metric("ÚLTIMA MOVIMENTAÇÃO", ultima)

        st.markdown("#### Suas Movimentações Cadastradas")
        st.dataframe(df_historico, use_container_width=True, hide_index=True)
