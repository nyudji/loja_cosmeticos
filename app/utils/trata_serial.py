# app/gera_serial.py

import pandas as pd
from typing import List

# Função para gerar SKU descritivo (Mantida, mas ajustada para 'row' no lugar de 'df')
def gerar_sku_descritivo(row):
    colecao = str(row["Coleção"]).upper().replace(" ", "")[:3]
    categoria = str(row["Categoria"]).upper().replace(" ", "")[:6]
    nome = str(row["Nome"]).upper().split(" ")[0][:4]
    volume = str(row["Volume"]).upper().replace(" ", "").replace("ML", "ML") if pd.notna(row["Volume"]) else ""
    
    return f"{colecao}-{categoria}-{nome}-{volume}"

def gerar_codigos(df: pd.DataFrame, colunas_primeiras: List[str]) -> pd.DataFrame:
    """
    Gera as colunas 'Serial Produto' e 'Cod Produto' e reordena o DataFrame.
    """
    df_vendas = df.copy()

    # Remover linhas vazias (Ajuste o nome da coluna conforme o Excel)
    df_vendas = df_vendas.dropna(subset=["Produto"])
    df_vendas = df_vendas.reset_index(drop=True)

    # 1. Criar coluna Serial Produto
    df_vendas["Serial Produto"] = ["NAT-" + str(i+1).zfill(5) for i in range(len(df_vendas))]

    # 2. Criar coluna Cod Produto
    df_vendas["Cod Produto"] = df_vendas.apply(gerar_sku_descritivo, axis=1)

    # 3. Reordenar as Colunas
    todas_colunas = df_vendas.columns.tolist()
    
    # Remove as colunas iniciais da lista 'todas_colunas' para evitar duplicação
    for col in colunas_primeiras:
        if col in todas_colunas:
            todas_colunas.remove(col)

    # Cria a lista final de colunas: (Novas) + (Restantes)
    ordem_final = colunas_primeiras + todas_colunas

    # Aplica a nova ordem ao DataFrame
    return df_vendas[ordem_final]