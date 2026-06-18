import streamlit as st
import pandas as pd
import sqlite3
import io

# Configuração da página
st.set_page_config(page_title="Gestão Comercial Pro", layout="wide")

@st.cache_resource
def get_db():
    return sqlite3.connect('sistema_comercial.db', check_same_thread=False)

def init_db():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS pedidos 
                          (data_entrega TEXT, cliente TEXT, nf TEXT, rota TEXT, valor REAL)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS devolucao 
                          (rota TEXT, cliente TEXT, produto TEXT, motivo TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS bpm 
                          (numero TEXT, cliente TEXT, motivo TEXT, resolvido TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS visitas 
                          (nome TEXT, local TEXT, cnpj TEXT, observacao TEXT)''')
        conn.commit()

init_db()

st.title("📦 Sistema de Gestão Comercial")

aba_pedidos, aba_devolucao, aba_bpm, aba_visitas, aba_relatorio = st.tabs([
    "Pedidos", "Devoluções", "BPM (Chamados)", "Visitas", "Relatórios & Exportação"
])

# LÓGICA DE PEDIDOS
with aba_pedidos:
    st.header("Cadastro de Pedidos")
    with st.form("form_ped", clear_on_submit=True):
        col1, col2 = st.columns(2)
        data = col1.date_input("Data de Entrega")
        cliente = col2.text_input("Nome do Cliente")
        nf = col1.text_input("Nota Fiscal (NF)")
        rota = col2.text_input("Rota")
        valor = col1.number_input("Valor (R$)", min_value=0.0)
        if st.form_submit_button("Salvar Pedido"):
            with get_db() as conn:
                conn.execute("INSERT INTO pedidos VALUES (?, ?, ?, ?, ?)", (str(data), cliente, nf, rota, valor))
                conn.commit()
            st.success("Pedido registrado!")

# LÓGICA DE DEVOLUÇÃO
with aba_devolucao:
    st.header("Registro de Devolução")
    with st.form("form_dev", clear_on_submit=True):
        rota_dev = st.text_input("Rota da Devolução")
        cliente_dev = st.text_input("Cliente da Devolução")
        prod_dev = st.text_input("Produto")
        motivo_dev = st.text_area("Motivo da Devolução")
        if st.form_submit_button("Salvar Devolução"):
            with get_db() as conn:
                conn.execute("INSERT INTO devolucao VALUES (?, ?, ?, ?)", (rota_dev, cliente_dev, prod_dev, motivo_dev))
                conn.commit()
            st.success("Devolução registrada!")

# LÓGICA DE BPM
with aba_bpm:
    st.header("Gestão de BPM (Chamados)")
    with st.form("form_bpm", clear_on_submit=True):
        num_bpm = st.text_input("Número do Processo")
        cliente_bpm = st.text_input("Cliente do Chamado")
        motivo_bpm = st.text_area("Motivo do Chamado")
        resolvido = st.selectbox("Status", ["Pendente", "Resolvido"])
        if st.form_submit_button("Salvar BPM"):
            with get_db() as conn:
                conn.execute("INSERT INTO bpm VALUES (?, ?, ?, ?)", (num_bpm, cliente_bpm, motivo_bpm, resolvido))
                conn.commit()
            st.success("BPM registrado!")

# LÓGICA DE VISITAS
with aba_visitas:
    st.header("Registro de Visitas")
    with st.form("form_visitas", clear_on_submit=True):
        nome = st.text_input("Nome do Visitante")
        local = st.text_input("Local da Visita")
        cnpj = st.text_input("CNPJ")
        obs = st.text_area("Observação")
        if st.form_submit_button("Salvar Visita"):
            with get_db() as conn:
                conn.execute("INSERT INTO visitas VALUES (?, ?, ?, ?)", (nome, local, cnpj, obs))
                conn.commit()
            st.success("Visita registrada!")

# LÓGICA DE EXPORTAÇÃO
with aba_relatorio:
    st.header("Relatório Geral")
    data_inicial = st.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
    data_final = st.date_input("Data Final", value=pd.to_datetime("2026-12-31"))

    if st.button("Gerar Excel Organizado"):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for tb in ['pedidos', 'devolucao', 'bpm', 'visitas']:
                df = pd.read_sql(f"SELECT * FROM {tb}", get_db())
                if 'data_entrega' in df.columns:
                    df['data_entrega'] = pd.to_datetime(df['data_entrega'])
                    df = df[(df['data_entrega'].dt.date >= data_inicial) & (df['data_entrega'].dt.date <= data_final)]
                
                df.to_excel(writer, sheet_name=tb.capitalize(), index=False)
                # Formatação rápida
                workbook = writer.book
                worksheet = writer.sheets[tb.capitalize()]
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
        
        buffer.seek(0) # ESSENCIAL: Volta o ponteiro para o início para permitir o download
        st.download_button(
            label="Baixar Excel Formatado",
            data=buffer,
            file_name="relatorio_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )