import streamlit as st
import pandas as pd
import sqlite3
from contextlib import closing

# Configuração da página
st.set_page_config(page_title="Gestão Comercial Pro", layout="wide")

# Conexão otimizada com cache
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
    col1, col2 = st.columns(2)
    data = col1.date_input("Data de Entrega")
    cliente = col2.text_input("Nome do Cliente", key="cli_ped")
    nf = col1.text_input("Nota Fiscal (NF)")
    rota = col2.text_input("Rota")
    valor = col1.number_input("Valor (R$)", min_value=0.0)
    
    if st.button("Salvar Pedido"):
        with get_db() as conn:
            conn.execute("INSERT INTO pedidos VALUES (?, ?, ?, ?, ?)", 
                         (str(data), cliente, nf, rota, valor))
            conn.commit()
        st.success("Pedido registrado!")

# LÓGICA DE DEVOLUÇÃO
with aba_devolucao:
    st.header("Registro de Devolução")
    rota_dev = st.text_input("Rota da Devolução")
    cliente_dev = st.text_input("Cliente da Devolução", key="cli_dev")
    prod_dev = st.text_input("Produto")
    motivo_dev = st.text_area("Motivo da Devolução")
    
    if st.button("Salvar Devolução"):
        with get_db() as conn:
            conn.execute("INSERT INTO devolucao VALUES (?, ?, ?, ?)", 
                         (rota_dev, cliente_dev, prod_dev, motivo_dev))
            conn.commit()
        st.success("Devolução registrada!")

# LÓGICA DE BPM
with aba_bpm:
    st.header("Gestão de BPM (Chamados)")
    num_bpm = st.text_input("Número do Processo")
    cliente_bpm = st.text_input("Cliente do Chamado", key="cli_bpm")
    motivo_bpm = st.text_area("Motivo do Chamado")
    resolvido = st.selectbox("Status", ["Pendente", "Resolvido"])
    
    if st.button("Salvar BPM"):
        with get_db() as conn:
            conn.execute("INSERT INTO bpm VALUES (?, ?, ?, ?)", 
                         (num_bpm, cliente_bpm, motivo_bpm, resolvido))
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
        
        submitted = st.form_submit_button("Salvar Visita")
        if submitted:
            with get_db() as conn:
                conn.execute("INSERT INTO visitas VALUES (?, ?, ?, ?)", 
                             (nome, local, cnpj, obs))
                conn.commit()
            st.success("Visita registrada com sucesso!")

# LÓGICA DE EXPORTAÇÃO
              
with aba_relatorio:
    st.header("Relatório Geral com Filtro")
    col_inicio, col_fim = st.columns(2)
    data_inicial = col_inicio.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
    data_final = col_fim.date_input("Data Final", value=pd.to_datetime("2026-12-31"))

    if st.button("Gerar Excel Organizado"):
        import io
        buffer = io.BytesIO()
        
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for tb in ['pedidos', 'devolucao', 'bpm' , 'visitas']:
                df = pd.read_sql(f"SELECT * FROM {tb}", get_db())
                
                # Filtragem de data (se aplicável)
                if 'data_entrega' in df.columns:
                    df['data_entrega'] = pd.to_datetime(df['data_entrega'])
                    mask = (df['data_entrega'].dt.date >= data_inicial) & (df['data_entrega'].dt.date <= data_final)
                    df = df.loc[mask]
                
                # Escreve o DataFrame
                df.to_excel(writer, sheet_name=tb.capitalize(), index=False)
                
                # Acessa o objeto workbook e worksheet do XlsxWriter
                workbook = writer.book
                worksheet = writer.sheets[tb.capitalize()]
                
                # Formato para o cabeçalho (Negrito, fundo cinza claro, bordas)
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#D3D3D3',
                    'border': 1
                })
                
                # Aplica formatação aos cabeçalhos
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    
                # Ajusta a largura das colunas automaticamente
                for i, col in enumerate(df.columns):
                    column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, column_len)
        
        st.download_button(
            label="Baixar Excel Formatado",
            data=buffer.getvalue(),
            file_name="relatorio_profissional.xlsx",
            mime="application/vnd.ms-excel"
        )
        st.success("Relatório gerado com sucesso!")