import pandas as pd
from pathlib import Path
import os # Necessário para o dir atual

# 1. Encontra o diretório onde o script atual (gera_serial.py) está
#    __file__ é uma variável mágica que contém o caminho do script
script_dir = Path(__file__).parent 

# 2. Constrói o caminho completo a partir do diretório do script:
#    script_dir / 'dados' / 'controle_financeiro.xlsx'
caminho_absoluto = script_dir / '..' / 'dados' / 'natura.xlsx'

try:
    # O pandas aceita o objeto Path diretamente!
    vendas_df = pd.read_excel(caminho_absoluto, sheet_name='Vendas')
    print("Arquivo lido com sucesso usando pathlib!")
except FileNotFoundError:
    # Mostra o caminho que ele tentou usar para facilitar o debug
    print(f"Erro: O arquivo não foi encontrado em '{caminho_absoluto}'.")
except Exception as e:
    print(f"Ocorreu um erro ao ler o arquivo: {e}")
vendas_df

df_vendas = vendas_df.copy()

#Criar coluna SKU_UNICO (serial único)
df_vendas = df_vendas.reset_index(drop=True)
df_vendas["Serial Produto"] = ["NAT-" + str(i+1).zfill(5) for i in range(len(df_vendas))]

#Função para gerar SKU descritivo
def gerar_sku_descritivo(row):
    colecao = str(row["Coleção"]).upper().replace(" ", "")[:3]  # abrevia coleção
    categoria = str(row["Categoria"]).upper().replace(" ", "")[:6]  # abrevia categoria
    nome = str(row["Nome"]).upper().split(" ")[0][:4]  # pega parte inicial da fragrância
    volume = str(row["Volume"]).upper().replace(" ", "").replace("ML", "ML") if pd.notna(row["Volume"]) else ""
    
    return f"{colecao}-{categoria}-{nome}-{volume}"

#Criar coluna SKU_DESCRITIVO
df_vendas["Cod Produto"] = df_vendas.apply(gerar_sku_descritivo, axis=1)

# Lista de todas as colunas existentes
todas_colunas = df_vendas.columns.tolist()

# Defina as colunas que você quer no início (a nova ordem)
colunas_primeiras = ["Serial Produto", "Cod Produto"]

# Remova as colunas iniciais da lista 'todas_colunas' para evitar duplicação
for col in colunas_primeiras:
    todas_colunas.remove(col)

# Crie a lista final de colunas: (Novas) + (Restantes)
ordem_final = colunas_primeiras + todas_colunas

# Aplique a nova ordem ao DataFrame
df_vendas = df_vendas[ordem_final]

#Visualizar primeiras linhas
df_vendas.head()

# Caminho do arquivo de SAÍDA (também .xlsx)
caminho_arquivo_saida = script_dir / '..' / 'dados' / 'natura_final.xlsx'

# --- Salvando o Arquivo Limpo ---
try:
    # MUDANÇA CRUCIAL: Usando df.to_excel()
    # index=False evita salvar o índice numérico padrão do pandas
    df_vendas.to_excel(caminho_arquivo_saida, index=False)
    print(f"\n✅ Arquivo Excel limpo salvo com sucesso em: {caminho_arquivo_saida}")
except Exception as e:
    print(f"Erro ao salvar o arquivo: {e}")