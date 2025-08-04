
import streamlit as st
import pandas as pd
from datetime import datetime
import os

ARQUIVO_EXCEL = "controle_financeiro_loja.xlsx"

def carregar_dados():
    if not os.path.exists(ARQUIVO_EXCEL):
        with pd.ExcelWriter(ARQUIVO_EXCEL) as writer:
            pd.DataFrame(columns=["Data", "Cliente", "Produto", "Valor Total", "Tipo de Pagamento",
                                  "Nº Parcelas", "Parcela Atual", "Valor Parcela", "Status Parcela",
                                  "Data Pagamento"]).to_excel(writer, sheet_name="Vendas", index=False)
    return pd.read_excel(ARQUIVO_EXCEL, sheet_name="Vendas")

def salvar_dados(df):
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name="Vendas", index=False)

def registrar_venda():
    st.subheader("Registrar Nova Venda")
    with st.form("form_venda"):
        cliente = st.text_input("Nome do Cliente")
        produto = st.text_input("Produto")
        valor_total = st.number_input("Valor Total", min_value=0.0)
        tipo_pagamento = st.selectbox("Tipo de Pagamento", ["À Vista", "Fiado / Parcelado"])
        num_parcelas = st.number_input("Número de Parcelas", min_value=1, value=1)
        submitted = st.form_submit_button("Registrar")

        if submitted:
            df = carregar_dados()
            data = datetime.today().strftime('%d/%m/%Y')
            valor_parcela = round(valor_total / num_parcelas, 2)
            for i in range(1, int(num_parcelas)+1):
                nova_linha = {
                    "Data": data,
                    "Cliente": cliente,
                    "Produto": produto,
                    "Valor Total": valor_total,
                    "Tipo de Pagamento": tipo_pagamento,
                    "Nº Parcelas": num_parcelas,
                    "Parcela Atual": i,
                    "Valor Parcela": valor_parcela,
                    "Status Parcela": "Pago" if tipo_pagamento == "À Vista" else "Pendente",
                    "Data Pagamento": data if tipo_pagamento == "À Vista" else ""
                }
                df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
            salvar_dados(df)
            st.success("Venda registrada com sucesso!")

def atualizar_pagamento():
    st.subheader("Atualizar Pagamento")
    df = carregar_dados()
    cliente = st.selectbox("Selecionar Cliente", df["Cliente"].unique())
    parcelas = df[(df["Cliente"] == cliente) & (df["Status Parcela"] == "Pendente")]
    if not parcelas.empty:
        parcela_escolhida = st.selectbox("Selecionar Parcela", parcelas["Parcela Atual"].astype(str))
        if st.button("Marcar como Paga"):
            idx = parcelas[parcelas["Parcela Atual"].astype(str) == parcela_escolhida].index[0]
            df.at[idx, "Status Parcela"] = "Pago"
            df.at[idx, "Data Pagamento"] = datetime.today().strftime('%d/%m/%Y')
            salvar_dados(df)
            st.success("Parcela atualizada como paga.")
    else:
        st.info("Nenhuma parcela pendente para este cliente.")

def mostrar_saldos():
    st.subheader("Saldo dos Clientes")
    df = carregar_dados()
    pendentes = df[df["Status Parcela"] == "Pendente"]
    saldo = pendentes.groupby("Cliente")["Valor Parcela"].sum()
    if not saldo.empty:
        st.dataframe(saldo.reset_index().rename(columns={"Valor Parcela": "Total Devido"}))
    else:
        st.info("Nenhum valor pendente.")

st.title("💄 Controle Financeiro - Loja de Cosméticos")

menu = st.sidebar.selectbox("Menu", ["Registrar Venda", "Atualizar Pagamento", "Ver Saldo de Clientes"])

if menu == "Registrar Venda":
    registrar_venda()
elif menu == "Atualizar Pagamento":
    atualizar_pagamento()
elif menu == "Ver Saldo de Clientes":
    mostrar_saldos()
