import pandas as pd
from datetime import datetime
import os

ARQUIVO_EXCEL = "controle_financeiro_loja.xlsx"

def criar_estrutura_excel():
    if os.path.exists(ARQUIVO_EXCEL):
        return
    
    vendas_cols = [
        "Data", "Cliente", "Produto", "Valor Total", "Tipo de Pagamento",
        "Nº Parcelas", "Parcela Atual", "Valor Parcela", "Status Parcela", "Data Pagamento"
    ]
    df_vendas = pd.DataFrame(columns=vendas_cols)
    
    clientes_cols = ["Cliente", "Telefone", "Email", "Total Devido", "Total Pago", "Saldo Atual"]
    df_clientes = pd.DataFrame(columns=clientes_cols)
    
    resumo_cols = ["Mês/Ano", "Total Vendas", "Total Recebido", "Total A Receber", "Saldo Caixa"]
    df_resumo = pd.DataFrame(columns=resumo_cols)
    
    with pd.ExcelWriter(ARQUIVO_EXCEL) as writer:
        df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
        df_clientes.to_excel(writer, sheet_name="Clientes", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo Financeiro", index=False)
    print("Arquivo Excel criado com sucesso!")


def registrar_venda():
    df_vendas = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Vendas")
    
    data = datetime.today().strftime('%d/%m/%Y')
    cliente = input("Nome do cliente: ")
    produto = input("Produto: ")
    valor_total = float(input("Valor total: "))
    tipo_pagamento = input("Tipo de pagamento (avista/fiado/parcelado): ").lower()
    
    if tipo_pagamento == "avista":
        nova_linha = {
            "Data": data, "Cliente": cliente, "Produto": produto,
            "Valor Total": valor_total, "Tipo de Pagamento": "À Vista",
            "Nº Parcelas": 1, "Parcela Atual": 1,
            "Valor Parcela": valor_total, "Status Parcela": "Pago", "Data Pagamento": data
        }
        df_vendas = df_vendas.append(nova_linha, ignore_index=True)
    
    elif tipo_pagamento in ["fiado", "parcelado"]:
        num_parcelas = int(input("Número de parcelas: "))
        valor_parcela = round(valor_total / num_parcelas, 2)
        for i in range(1, num_parcelas + 1):
            nova_linha = {
                "Data": data, "Cliente": cliente, "Produto": produto,
                "Valor Total": valor_total, "Tipo de Pagamento": "Fiado / Parcelado",
                "Nº Parcelas": num_parcelas, "Parcela Atual": i,
                "Valor Parcela": valor_parcela, "Status Parcela": "Pendente", "Data Pagamento": ""
            }
            df_vendas = df_vendas.append(nova_linha, ignore_index=True)

    else:
        print("Tipo de pagamento inválido!")
        return

    with pd.ExcelWriter(ARQUIVO_EXCEL, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
    print("Venda registrada com sucesso!")


def atualizar_pagamento():
    df_vendas = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Vendas")
    cliente = input("Nome do cliente: ")
    
    pendentes = df_vendas[(df_vendas["Cliente"] == cliente) & (df_vendas["Status Parcela"] == "Pendente")]
    if pendentes.empty:
        print("Nenhuma parcela pendente encontrada.")
        return

    print(pendentes[["Parcela Atual", "Valor Parcela"]])
    parcela = int(input("Qual parcela foi paga? "))
    idx = pendentes[pendentes["Parcela Atual"] == parcela].index

    if not idx.empty:
        df_vendas.at[idx[0], "Status Parcela"] = "Pago"
        df_vendas.at[idx[0], "Data Pagamento"] = datetime.today().strftime('%d/%m/%Y')
        with pd.ExcelWriter(ARQUIVO_EXCEL, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
        print("Parcela atualizada como paga.")
    else:
        print("Parcela não encontrada.")


def saldo_clientes():
    df_vendas = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Vendas")
    pendentes = df_vendas[df_vendas["Status Parcela"] == "Pendente"]
    total = pendentes.groupby("Cliente")["Valor Parcela"].sum()
    print("\n--- Saldo a receber por cliente ---")
    print(total)


# Menu interativo
def menu():
    criar_estrutura_excel()
    while True:
        print("\n1. Registrar nova venda")
        print("2. Atualizar pagamento")
        print("3. Mostrar saldo dos clientes")
        print("0. Sair")
        opcao = input("Escolha: ")
        if opcao == "1":
            registrar_venda()
        elif opcao == "2":
            atualizar_pagamento()
        elif opcao == "3":
            saldo_clientes()
        elif opcao == "0":
            break
        else:
            print("Opção inválida!")

if __name__ == "__main__":
    menu()
