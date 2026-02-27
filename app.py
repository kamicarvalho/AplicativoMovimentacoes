import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, timedelta, timezone
import os
import openpyxl
import base64
import time
from supabase import create_client, Client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA
# ==========================================
st.set_page_config(page_title="Movimentações - Headcount", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# 2. CONEXÃO COM O SUPABASE E FUSO
# ==========================================
fuso_br = timezone(timedelta(hours=-3))

@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_connection()

# ==========================================
# 3. O MOTOR JAVASCRIPT (Versão Segura - Anti-travamento)
# ==========================================
js_code = """
<script>
const parentDoc = window.parent.document;

function aplicarEstilos() {
    const header = parentDoc.querySelector('header');
    if(header) header.style.display = 'none';

    const mainBlock = parentDoc.querySelector('.block-container');
    if(mainBlock) {
        mainBlock.style.paddingTop = '1.5rem';
        mainBlock.style.paddingBottom = '1rem';
        mainBlock.style.maxWidth = '98%';
    }

    const blocos = parentDoc.querySelectorAll('div[data-testid="stVerticalBlock"]');
    blocos.forEach(b => b.style.gap = '0.1rem');

    const botoes = parentDoc.querySelectorAll('button');
    botoes.forEach(btn => {
        const txt = btn.innerText.trim();
        if(txt === '✅ CONFIRMAR MOVIMENTAÇÃO') {
            btn.style.backgroundColor = '#2e7d32'; btn.style.color = 'white'; btn.style.border = 'none'; btn.style.fontWeight = 'bold';
        } else if(txt === '📄 Ver Histórico (Consultas)' || txt === 'Nova Movimentação' || txt === 'ACESSAR SISTEMA' || txt === 'ENVIAR SOLICITAÇÃO') {
            btn.style.backgroundColor = '#1976d2'; btn.style.color = 'white'; btn.style.border = 'none'; btn.style.fontWeight = 'bold';
        } else if(txt === 'Sair') {
            btn.style.backgroundColor = '#d32f2f'; btn.style.color = 'white'; btn.style.border = 'none'; btn.style.fontWeight = 'bold';
        } else if(txt.includes('Não encontrou o posto?')) {
            btn.style.backgroundColor = '#ff9800'; btn.style.color = 'white'; btn.style.border = 'none'; btn.style.fontWeight = 'bold';
        }
    });
}

// Verifica e aplica as cores a cada 500ms
setInterval(aplicarEstilos, 500);
</script>
"""
components.html(js_code, height=0, width=0)

# ==========================================
# 4. PUXANDO USUÁRIOS DO COFRE DE SEGREDOS
# ==========================================
try:
    USUARIOS_PERMITIDOS = dict(st.secrets["usuarios"])
except KeyError:
    st.error("Erro Crítico: A seção [usuarios] não foi encontrada nos Secrets do Streamlit.")
    USUARIOS_PERMITIDOS = {}

# ==========================================
# 5. LER EXCEL (PARÂMETROS LOCAIS)
# ==========================================
@st.cache_data
def carregar_dados_excel():
    arquivo_excel = 'parametros.xlsx'
    if not os.path.exists(arquivo_excel):
        return pd.DataFrame(columns=['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante'])
    try:
        df = pd.read_excel(arquivo_excel, dtype=str).iloc[:, :7]
        df.columns = ['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante']
        return df.fillna("") 
    except:
        return pd.DataFrame(columns=['unidade', 'cc', 'sub', 'gestor', 'posto', 'cargo', 'requisitante'])

df_parametros = carregar_dados_excel()

def renderizar_logo(tamanho=180):
    if os.path.exists("logo.png"):
        with open("logo.png", "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
            st.markdown(f'<div style="text-align: center; margin-bottom: 10px;"><img src="data:image/png;base64,{encoded}" width="{tamanho}"></div>', unsafe_allow_html=True)

# ==========================================
# 6. CONTROLE DE SESSÃO E MEMÓRIA
# ==========================================
if 'usuario_logado' not in st.session_state:
    st.session_state.usuario_logado = None
if 'pagina' not in st.session_state:
    st.session_state.pagina = 'login'
if 'sucesso_movimentacao' not in st.session_state:
    st.session_state.sucesso_movimentacao = False
if 'form_key' not in st.session_state:
    st.session_state.form_key = 0 

def fazer_logout():
    st.session_state.usuario_logado = None
    st.session_state.pagina = 'login'

# ==========================================
# 7. MODAL: CADASTRAR POSTO FALTANTE E E-MAIL
# ==========================================
@st.dialog("Cadastro Posto faltante")
def modal_solicitar_posto():
    st.markdown("<p style='color: black;'>Preencha os dados abaixo.</p>", unsafe_allow_html=True)
    
    und_p = st.selectbox("Unidade:", options=sorted([x for x in df_parametros['unidade'].unique() if x]), index=None)
    df_cc_p = df_parametros[df_parametros['unidade'] == und_p] if und_p else pd.DataFrame(columns=df_parametros.columns)
    cc_p = st.selectbox("Centro de Custo:", options=sorted([x for x in df_cc_p['cc'].unique() if x]), index=None)
    df_sub_p = df_cc_p[df_cc_p['cc'] == cc_p] if cc_p else pd.DataFrame(columns=df_parametros.columns)
    sub_p = st.selectbox("Subprocesso:", options=sorted([x for x in df_sub_p['sub'].unique() if x]), index=None)
    df_gestor_p = df_sub_p[df_sub_p['sub'] == sub_p] if sub_p else pd.DataFrame(columns=df_parametros.columns)
    gestor_p = st.selectbox("Gestor:", options=sorted([x for x in df_gestor_p['gestor'].unique() if x]), index=None)
    cargo_p = st.selectbox("Qual Cargo deve pertencer a esse posto?:", options=sorted([x for x in df_parametros['cargo'].unique() if x]), index=None)

    st.write("")
    if st.button("ENVIAR SOLICITAÇÃO", use_container_width=True):
        if not all([und_p, cc_p, sub_p, gestor_p, cargo_p]):
            st.error("Por favor, preencha todos os campos antes de enviar.")
        else:
            with st.spinner("Salvando no banco e enviando e-mail para o RH..."):
                try:
                    data_atual = datetime.now(fuso_br).isoformat()
                    
                    # 1. SALVA NO SUPABASE
                    dados_solicitacao = {
                        "data_solicitacao": data_atual,
                        "usuario": st.session_state.usuario_logado,
                        "unidade": und_p,
                        "centro_custo": cc_p,
                        "subprocesso": sub_p,
                        "gestor": gestor_p,
                        "cargo": cargo_p
                    }
                    supabase.table("solicitacoes_postos").insert(dados_solicitacao).execute()
                    
                    # 2. DISPARA O E-MAIL
                    erro_real_do_email = ""
                    try:
                        remetente = st.secrets["EMAIL_REMETENTE"]
                        senha = st.secrets["SENHA_REMETENTE"].replace(" ", "")
                        destinatario = st.secrets["EMAIL_RH"] 
                        servidor_smtp = st.secrets["SERVIDOR_SMTP"]
                        
                        msg = MIMEMultipart()
                        msg['From'] = remetente
                        msg['To'] = destinatario
                        msg['Subject'] = "🚨 Nova Solicitação de Posto Faltante - Headcount"
                        
                        corpo_email = f"""
Olá equipe do RH,

Uma nova solicitação de criação de posto foi registrada no sistema de Movimentações de Headcount.

Detalhes da Solicitação:
--------------------------------------------------
- Solicitante: {st.session_state.usuario_logado}
- Unidade: {und_p}
- Centro de Custo: {cc_p}
- Subprocesso: {sub_p}
- Gestor Associado: {gestor_p}
- Cargo que ocupará o posto: {cargo_p}
--------------------------------------------------

Por favor, providencie o cadastro do posto no sistema oficial para que a movimentação possa ser concluída.

Mensagem automática do Sistema de Headcount.
"""
                        msg.attach(MIMEText(corpo_email, 'plain'))
                        
                        lista_destinatarios = [email.strip() for email in destinatario.split(',')]
                        
                        server = smtplib.SMTP(servidor_smtp, 587)
                        server.starttls()
                        server.login(remetente, senha)
                        server.sendmail(remetente, lista_destinatarios, msg.as_string())
                        server.quit()
                        email_sucesso = True
                    except Exception as email_err:
                        erro_real_do_email = str(email_err)
                        email_sucesso = False

                    # 3. FEEDBACK FINAL
                    if email_sucesso:
                        st.success("✅ Solicitação salva e E-mail enviado com sucesso!")
                        time.sleep(2)
                    else:
                        st.warning(f"⚠️ Salvo no Supabase, mas falhou ao enviar e-mail. ERRO: {erro_real_do_email}")
                        time.sleep(6)
                    
                    st.rerun()

                except Exception as e:
                    st.error(f"Erro ao salvar na nuvem (Supabase): {e}")

# ==========================================
# 8. TELAS DO APLICATIVO
# ==========================================

# --- TELA DE LOGIN ---
if st.session_state.usuario_logado is None:
    st.write("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.5, 1.2, 1.5]) 
    with col2:
        with st.container(border=True):
            st.write("<br>", unsafe_allow_html=True)
            renderizar_logo(180)
            st.markdown("<h3 style='text-align: center; color: black;'>Movimentações<br>HeadCount</h3>", unsafe_allow_html=True)
            st.write("<br>", unsafe_allow_html=True)
            
            with st.form("form_login", clear_on_submit=False):
                usuario = st.text_input("Usuário")
                senha = st.text_input("Senha", type="password")
                st.write("<br>", unsafe_allow_html=True)
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
    col_titulo, col_user, col_btn1, col_btn2 = st.columns([4, 2, 2, 1.5])
    
    with col_titulo:
        st.markdown("<h2 style='color: black; margin-top: -15px;'>Sistema de Movimentações</h2>", unsafe_allow_html=True)
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
        
        if st.session_state.sucesso_movimentacao:
            st.success("✅ Movimentação registrada com sucesso!")
            st.session_state.sucesso_movimentacao = False 
        
        fk = st.session_state.form_key 
        lista_req = sorted([x for x in df_parametros['requisitante'].unique() if x])
        requisitante = st.selectbox("Quem solicitou a troca? (Pode digitar para pesquisar)", options=lista_req, index=None, placeholder="Selecione o requisitante...", key=f"req_{fk}")

        st.write("") 
        col_saida, col_entrada = st.columns(2, gap="large")

        # ==== LADO ESQUERDO: SAÍDA ====
        with col_saida:
            with st.container(border=True):
                st.markdown("""
                <div style="background-color: #fff5f5; border: 2px solid #ffcdd2; border-radius: 8px; padding: 12px; margin-bottom: 15px;">
                    <h4 style="text-align: center; color: #b71c1c; margin: 0;">VAGA DE SAÍDA (RETIRADA)</h4>
                </div>
                """, unsafe_allow_html=True)
                
                s_und = st.selectbox("Unidade (Saída):", options=sorted([x for x in df_parametros['unidade'].unique() if x]), index=None, key=f"s_und_{fk}")
                df_s_cc = df_parametros[df_parametros['unidade'] == s_und] if s_und else pd.DataFrame(columns=df_parametros.columns)
                s_cc = st.selectbox("Centro de Custo (Saída):", options=sorted([x for x in df_s_cc['cc'].unique() if x]), index=None, key=f"s_cc_{fk}")
                df_s_sub = df_s_cc[df_s_cc['cc'] == s_cc] if s_cc else pd.DataFrame(columns=df_parametros.columns)
                s_sub = st.selectbox("Subprocesso (Saída):", options=sorted([x for x in df_s_sub['sub'].unique() if x]), index=None, key=f"s_sub_{fk}")
                df_s_gestor = df_s_sub[df_s_sub['sub'] == s_sub] if s_sub else pd.DataFrame(columns=df_parametros.columns)
                s_gestor = st.selectbox("Gestor (Saída):", options=sorted([x for x in df_s_gestor['gestor'].unique() if x]), index=None, key=f"s_gestor_{fk}")
                df_s_posto = df_s_gestor[df_s_gestor['gestor'] == s_gestor] if s_gestor else pd.DataFrame(columns=df_parametros.columns)
                s_posto = st.selectbox("Posto (Saída):", options=sorted([x for x in df_s_posto['posto'].unique() if x]), index=None, key=f"s_posto_{fk}")
                df_s_cargo = df_s_posto[df_s_posto['posto'] == s_posto] if s_posto else pd.DataFrame(columns=df_parametros.columns)
                s_cargo = st.selectbox("Cargo (Saída):", options=sorted([x for x in df_s_cargo['cargo'].unique() if x]), index=None, key=f"s_cargo_{fk}")
                s_qtd = st.number_input("Quantidade (Saída):", min_value=1, value=1, step=1, key=f"s_qtd_{fk}")

        # ==== LADO DIREITO: ENTRADA ====
        with col_entrada:
            with st.container(border=True):
                st.markdown("""
                <div style="background-color: #f1f8e9; border: 2px solid #c8e6c9; border-radius: 8px; padding: 12px; margin-bottom: 15px;">
                    <h4 style="text-align: center; color: #1b5e20; margin: 0;">VAGA DE ENTRADA (NOVA)</h4>
                </div>
                """, unsafe_allow_html=True)
                
                e_und = st.selectbox("Unidade (Entrada):", options=sorted([x for x in df_parametros['unidade'].unique() if x]), index=None, key=f"e_und_{fk}")
                df_e_cc = df_parametros[df_parametros['unidade'] == e_und] if e_und else pd.DataFrame(columns=df_parametros.columns)
                e_cc = st.selectbox("Centro de Custo (Entrada):", options=sorted([x for x in df_e_cc['cc'].unique() if x]), index=None, key=f"e_cc_{fk}")
                df_e_sub = df_e_cc[df_e_cc['cc'] == e_cc] if e_cc else pd.DataFrame(columns=df_parametros.columns)
                e_sub = st.selectbox("Subprocesso (Entrada):", options=sorted([x for x in df_e_sub['sub'].unique() if x]), index=None, key=f"e_sub_{fk}")
                df_e_gestor = df_e_sub[df_e_sub['sub'] == e_sub] if e_sub else pd.DataFrame(columns=df_parametros.columns)
                e_gestor = st.selectbox("Gestor (Entrada):", options=sorted([x for x in df_e_gestor['gestor'].unique() if x]), index=None, key=f"e_gestor_{fk}")
                df_e_posto = df_e_gestor[df_e_gestor['gestor'] == e_gestor] if e_gestor else pd.DataFrame(columns=df_parametros.columns)
                e_posto = st.selectbox("Posto (Entrada):", options=sorted([x for x in df_e_posto['posto'].unique() if x]), index=None, key=f"e_posto_{fk}")
                df_e_cargo = df_e_posto[df_e_posto['posto'] == e_posto] if e_posto else pd.DataFrame(columns=df_parametros.columns)
                e_cargo = st.selectbox("Cargo (Entrada):", options=sorted([x for x in df_e_cargo['cargo'].unique() if x]), index=None, key=f"e_cargo_{fk}")
                e_qtd = st.number_input("Quantidade (Entrada):", min_value=1, value=1, step=1, key=f"e_qtd_{fk}")
                
                st.write("")
                if st.button("Não encontrou o posto? Clique aqui para solicitar", use_container_width=True):
                    modal_solicitar_posto()

        st.write("")
        
        # ==== BOTÃO SALVAR (NO SUPABASE) ====
        if st.button("✅ CONFIRMAR MOVIMENTAÇÃO", use_container_width=True):
            if not requisitante:
                st.warning("⚠️ O campo Requisitante é obrigatório.")
            elif not all([s_und, s_cc, s_sub, s_gestor, s_posto, s_cargo, e_und, e_cc, e_sub, e_gestor, e_posto, e_cargo]):
                st.warning("⚠️ Preencha todas as caixas de Saída e Entrada antes de salvar.")
            else:
                try:
                    data_atual = datetime.now(fuso_br).isoformat()
                    
                    dados_movimentacao = {
                        "usuario_sistema": st.session_state.usuario_logado,
                        "data_registro": data_atual,
                        "requisitante": requisitante,
                        "unidade_saida": s_und, "cc_saida": s_cc, "subprocesso_saida": s_sub,
                        "gestor_saida": s_gestor, "posto_saida": s_posto, "cargo_saida": s_cargo, "qtd_saida": s_qtd,
                        "unidade_entrada": e_und, "cc_entrada": e_cc, "subprocesso_entrada": e_sub,
                        "gestor_entrada": e_gestor, "posto_entrada": e_posto, "cargo_entrada": e_cargo, "qtd_entrada": e_qtd
                    }
                    
                    supabase.table("movimentacoes").insert(dados_movimentacao).execute()
                    
                    st.session_state.sucesso_movimentacao = True
                    st.session_state.form_key += 1 
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar no Supabase: {e}")

    # --- TELA DE CONSULTA (DO SUPABASE) ---
    elif st.session_state.pagina == 'consulta':
        try:
            resposta = supabase.table("movimentacoes").select("*").eq("usuario_sistema", st.session_state.usuario_logado).order("id", desc=True).execute()
            df_historico = pd.DataFrame(resposta.data)
            
            if not df_historico.empty:
                df_historico = df_historico[['id', 'data_registro', 'usuario_sistema', 'requisitante', 'cc_saida', 'qtd_saida', 'cargo_saida', 'cc_entrada', 'qtd_entrada', 'cargo_entrada']]
                df_historico.columns = ["ID", "Data", "Usuário", "Requisitante", "CC Saída", "Qtd Saída", "Cargo Saída", "CC Entrada", "Qtd Entrada", "Cargo Entrada"]
                df_historico['Data'] = pd.to_datetime(df_historico['Data']).dt.strftime('%d/%m/%Y %H:%M')

                total = len(df_historico)
                ultima = df_historico['Data'].iloc[0]
            else:
                df_historico = pd.DataFrame(columns=["ID", "Data", "Usuário", "Requisitante", "CC Saída", "Qtd Saída", "Cargo Saída", "CC Entrada", "Qtd Entrada", "Cargo Entrada"])
                total = 0
                ultima = "-"

            col_metric1, col_metric2 = st.columns(2)
            col_metric1.metric("TOTAL REGISTRADO", total)
            col_metric2.metric("ÚLTIMA MOVIMENTAÇÃO", ultima)

            st.markdown("#### Suas Movimentações Cadastradas")
            
            def colorir_tabela(coluna):
                if coluna.name in ["CC Saída", "Qtd Saída", "Cargo Saída"]:
                    return ['background-color: #ffebee; color: #b71c1c'] * len(coluna)
                elif coluna.name in ["CC Entrada", "Qtd Entrada", "Cargo Entrada"]:
                    return ['background-color: #e8f5e9; color: #1b5e20'] * len(coluna)
                else:
                    return [''] * len(coluna)
                    
            if total > 0:
                df_estilizado = df_historico.style.apply(colorir_tabela)
                st.dataframe(df_estilizado, use_container_width=True, hide_index=True)
            else:
                st.info("Você ainda não possui movimentações registradas.")
                
        except Exception as e:
            st.error(f"Erro ao puxar dados do banco de dados: {e}")
