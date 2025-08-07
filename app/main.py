import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Caminho do arquivo Excel
ARQUIVO_EXCEL = os.path.join("app", "dados", "controle_financeiro.xlsx")

# Colunas padrão da planilha
COLUNAS = [
    "Data", "Cliente", "Produto", "Valor Total", "Tipo de Pagamento",
    "Nº Parcelas", "Parcela Atual", "Valor Parcela", "Status Parcela",
    "Data Pagamento"
]

def carregar_dados():
    """Carrega os dados da planilha Excel, cria se não existir ou estiver corrompido."""
    os.makedirs(os.path.dirname(ARQUIVO_EXCEL), exist_ok=True)

    if not os.path.exists(ARQUIVO_EXCEL):
        df_vazio = pd.DataFrame(columns=COLUNAS)
        with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl') as writer:
            df_vazio.to_excel(writer, sheet_name="Vendas", index=False)
        return df_vazio

    try:
        xls = pd.ExcelFile(ARQUIVO_EXCEL)
        if "Vendas" not in xls.sheet_names:
            raise ValueError("Aba 'Vendas' não encontrada no arquivo Excel.")
        df = pd.read_excel(xls, sheet_name="Vendas")
        return df

    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}. Criando arquivo novo.")
        df_vazio = pd.DataFrame(columns=COLUNAS)
        with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl') as writer:
            df_vazio.to_excel(writer, sheet_name="Vendas", index=False)
        return df_vazio

def salvar_dados(df):
    """Salva o DataFrame na planilha Excel."""
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name="Vendas", index=False)

def registrar_venda():
    """Form para registrar uma nova venda."""
    st.subheader("Registrar Nova Venda")
    with st.form("form_venda"):
        cliente = st.text_input("Nome do Cliente")
        produto = st.text_input("Produto")
        valor_total = st.number_input("Valor Total", min_value=0.0, format="%.2f")
        tipo_pagamento = st.selectbox("Tipo de Pagamento", ["À Vista", "Fiado / Parcelado"])
        num_parcelas = st.number_input("Número de Parcelas", min_value=1, value=1)
        submitted = st.form_submit_button("Registrar")

        if submitted:
            if not cliente.strip() or not produto.strip() or valor_total <= 0:
                st.error("Preencha todos os campos corretamente.")
                return

            df = carregar_dados()
            data = datetime.today().strftime('%d/%m/%Y')
            valor_parcela = round(valor_total / num_parcelas, 2)

            for i in range(1, int(num_parcelas) + 1):
                nova_linha = {
                    "Data": data,
                    "Cliente": cliente.strip(),
                    "Produto": produto.strip(),
                    "Valor Total": valor_total,
                    "Tipo de Pagamento": tipo_pagamento,
                    "Nº Parcelas": num_parcelas,
                    "Parcela Atual": i,
                    "Valor Parcela": valor_parcela,
                    "Status Parcela": "Pago" if tipo_pagamento == "À Vista" else "Pendente",
                    "Data Pagamento": data if tipo_pagamento == "À Vista" else "00/00/0000"
                }
                df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)

            salvar_dados(df)
            st.success("Venda registrada com sucesso!")

def atualizar_pagamento():
    """Form para atualizar o status de parcelas pendentes."""
    st.subheader("Atualizar Pagamento")
    df = carregar_dados()

    if df.empty or "Cliente" not in df.columns:
        st.info("Nenhum dado disponível para atualizar.")
        return

    clientes = df["Cliente"].dropna().unique()
    if len(clientes) == 0:
        st.info("Nenhum cliente encontrado.")
        return

    cliente = st.selectbox("Selecionar Cliente", clientes)
    parcelas = df[(df["Cliente"] == cliente) & (df["Status Parcela"] == "Pendente")]

    if parcelas.empty:
        st.info("Nenhuma parcela pendente para este cliente.")
        return

    parcela_escolhida = st.selectbox("Selecionar Parcela", parcelas["Parcela Atual"].astype(str))

    if st.button("Marcar como Paga"):
        idx = parcelas[parcelas["Parcela Atual"].astype(str) == parcela_escolhida].index[0]
        df.at[idx, "Status Parcela"] = "Pago"
        df.at[idx, "Data Pagamento"] = datetime.today().strftime('%d/%m/%Y')
        salvar_dados(df)
        st.success("Parcela atualizada como paga.")

def mostrar_saldos():
    """Exibe o saldo total pendente por cliente."""
    st.subheader("Saldo dos Clientes")
    df = carregar_dados()

    if df.empty:
        st.info("Nenhum dado disponível para mostrar.")
        return

    if not all(col in df.columns for col in ["Status Parcela", "Cliente", "Valor Parcela"]):
        st.error("Arquivo de dados está incompleto.")
        return

    pendentes = df[df["Status Parcela"] == "Pendente"]
    saldo = pendentes.groupby("Cliente")["Valor Parcela"].sum()

    if saldo.empty:
        st.info("Nenhum valor pendente.")
    else:
        st.dataframe(saldo.reset_index().rename(columns={"Valor Parcela": "Total Devido"}))

def mostrar_todas_vendas():
    """Exibe a tabela completa de todas as vendas registradas."""
    st.subheader("Todas as Vendas Registradas")
    df = carregar_dados()
    if df.empty:
        st.info("Nenhuma venda registrada ainda.")
    else:
        st.dataframe(df)

modulo = st.sidebar.radio("Escolha o setor de trabalho", ["Vendas", "Clientes"])

if modulo == "Vendas":
    acao_vendas = st.sidebar.selectbox("Ações de Vendas", [
        "Registrar Venda",
        "Atualizar Pagamento",
        "Ver Saldo de Clientes",
        "Ver Todas as Vendas"
    ])
    if acao_vendas == "Registrar Venda":
        registrar_venda()
    elif acao_vendas == "Atualizar Pagamento":
        atualizar_pagamento()
    elif acao_vendas == "Ver Saldo de Clientes":
        mostrar_saldos()
    elif acao_vendas == "Ver Todas as Vendas":
        mostrar_todas_vendas()

elif modulo == "Clientes":
    acao_clientes = st.sidebar.selectbox("Ações de Clientes", [
        "Registrar Cliente",
        "Atualizar Cliente",
        "Ver Todos os Clientes"
    ])
    if acao_clientes == "Registrar Cliente":
        st.info("Funcionalidade de registro de cliente ainda não implementada.")
    elif acao_clientes == "Atualizar Cliente":
        st.info("Funcionalidade de atualização de cliente ainda não implementada.")
    elif acao_clientes == "Ver Todos os Clientes":
        st.info("Funcionalidade de visualização de todos os clientes ainda não implementada.")