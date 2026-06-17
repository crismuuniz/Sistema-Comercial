import streamlit as st
import pandas as pd
import sqlite3

# Configuração da página
st.set_page_config(page_title="Gestão Comercial Pro", layout="wide")

# Conexão com Banco de Dados
def get_db():
    conn = sqlite3.connect('sistema_comercial.db')
    return conn

# Criar tabelas se não existirem
def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS pedidos 
                      (data_entrega TEXT, cliente TEXT, nf TEXT, rota TEXT, valor REAL)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS devolucao 
                      (rota TEXT, cliente TEXT, produto TEXT, motivo TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS bpm 
                      (numero TEXT, cliente TEXT, motivo TEXT, resolvido TEXT)''')
    conn.commit()
    conn.close()

init_db()

st.title("📦 Sistema de Gestão Comercial")

# Criando as abas
aba_pedidos, aba_devolucao, aba_bpm, aba_relatorio = st.tabs([
    "Pedidos", "Devoluções", "BPM (Chamados)", "Relatórios & Exportação"
])

# LÓGICA DE PEDIDOS
with aba_pedidos:
    st.header("Cadastro de Pedidos")
    col1, col2 = st.columns(2)
    data = col1.date_input("Data de Entrega")
    cliente = col2.text_input("Nome do Cliente")
    nf = col1.text_input("Nota Fiscal (NF)")
    rota = col2.text_input("Rota")
    valor = col1.number_input("Valor (R$)", min_value=0.0)
    
    if st.button("Salvar Pedido"):
        conn = get_db()
        conn.execute("INSERT INTO pedidos VALUES (?, ?, ?, ?, ?)", (str(data), cliente, nf, rota, valor))
        conn.commit()
        st.success("Pedido registrado!")

# LÓGICA DE DEVOLUÇÃO
with aba_devolucao:
    st.header("Registro de Devolução")
    rota_dev = st.text_input("Rota da Devolução")
    cliente_dev = st.text_input("Cliente da Devolução")
    prod_dev = st.text_input("Produto")
    motivo_dev = st.text_area("Motivo da Devolução")
    
    if st.button("Salvar Devolução"):
        conn = get_db()
        conn.execute("INSERT INTO devolucao VALUES (?, ?, ?, ?)", (rota_dev, cliente_dev, prod_dev, motivo_dev))
        conn.commit()
        st.success("Devolução registrada!")

# LÓGICA DE BPM
with aba_bpm:
    st.header("Gestão de BPM (Chamados)")
    num_bpm = st.text_input("Número do Processo")
    cliente_bpm = st.text_input("Cliente do Chamado")
    motivo_bpm = st.text_area("Motivo do Chamado")
    resolvido = st.selectbox("Status", ["Pendente", "Resolvido"])
    
    if st.button("Salvar BPM"):
        conn = get_db()
        conn.execute("INSERT INTO bpm VALUES (?, ?, ?, ?)", (num_bpm, cliente_bpm, motivo_bpm, resolvido))
        conn.commit()
        st.success("BPM registrado!")

# LÓGICA DE EXPORTAÇÃO
# LÓGICA DE EXPORTAÇÃO COM FILTRO
with aba_relatorio:
    st.header("Relatório Geral com Filtro")
    
    # Criar colunas para selecionar o período
    col_inicio, col_fim = st.columns(2)
    data_inicial = col_inicio.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
    data_final = col_fim.date_input("Data Final", value=pd.to_datetime("2026-12-31"))

    if st.button("Gerar Excel Filtrado"):
        with pd.ExcelWriter("relatorio_filtrado.xlsx") as writer:
            # Lista de tabelas para exportar
            for tb in ['pedidos', 'devolucao', 'bpm']:
                # Carrega a tabela
                df = pd.read_sql(f"SELECT * FROM {tb}", get_db())
                
                # Tenta converter colunas de data (ajuste conforme o nome da sua coluna de data)
                if 'data_entrega' in df.columns:
                    df['data_entrega'] = pd.to_datetime(df['data_entrega'])
                    # Filtra o DataFrame
                    mask = (df['data_entrega'].dt.date >= data_inicial) & (df['data_entrega'].dt.date <= data_final)
                    df = df.loc[mask]
                
                df.to_excel(writer, sheet_name=tb.capitalize(), index=False)
        
        st.success("Relatório 'relatorio_filtrado.xlsx' gerado!")
        with open("relatorio_filtrado.xlsx", "rb") as f:
            st.download_button("Baixar Excel Filtrado", f, "relatorio_filtrado.xlsx")