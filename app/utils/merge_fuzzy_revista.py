import pandas as pd
import numpy as np
from pathlib import Path 

# ====================================================================
# 1. Configuração de Caminhos
# ====================================================================

# 1.1. Obtém o caminho do diretório ATUAL onde o script está
script_dir = Path(__file__).resolve().parent

# 1.2. Define o caminho para a pasta 'dados/' no diretório pai
caminho_dados = script_dir.parent / 'dados'

# --- 1. Variáveis de Caminho Completos ---

# O NOME DO SEU ARQUIVO BASE (MUDANÇA: AGORA É XLSX)
# Usa o caminho completo: dados/BD_Loja.xlsx
PATH_BASE_PRODUTOS = caminho_dados / 'BD_Loja.xlsx'

# O RELATÓRIO COM O MATCH JÁ FEITO (MANTÉM CSV, pois foi gerado assim)
# Usa o caminho completo: dados/relatorio_match_produtos_final_corrigido.csv
PATH_RELATORIO_MATCH = caminho_dados / 'relatorio_match_produtos_final_corrigido.csv'

# O NOME DO NOVO ARQUIVO DA BASE DE PRODUTOS ATUALIZADA (MUDANÇA: AGORA É XLSX)
# Usa o caminho completo: dados/BD_Loja_Produtos_COMCOD_revista.xlsx
PATH_BASE_ATUALIZADA = caminho_dados / 'BD_Loja_Produtos_COMCOD_revista.xlsx'

# NOME DA ABA DENTRO DO ARQUIVO XLSX QUE CONTÉM OS PRODUTOS
SHEET_NAME_BASE = 'Produtos' 


# --- 2. Carregar os Dados ---
try:
    # MUDANÇA: Carrega a Base de Produtos usando o objeto Path
    df_base = pd.read_excel(
        PATH_BASE_PRODUTOS, # Usa a variável Path
        sheet_name=SHEET_NAME_BASE,
        # Seleciona as colunas necessárias para evitar problemas de formatação
        usecols=['COD', 'Marca', 'Coleção', 'Categoria', 'Produto', 'Nome', 'Unidade', 'Volume', 'Tipo', 'Descrição', 'Preço Custo', 'Preço Venda']
    )

    # Carrega o Relatório de Match (mantido como CSV)
    # Usa a variável Path
    df_match = pd.read_csv(
        PATH_RELATORIO_MATCH, # Usa a variável Path
        encoding='utf-8',
        on_bad_lines='skip'
    )

    print("Arquivos carregados com sucesso da pasta 'dados/'.")
    # ... Restante do código
    
except FileNotFoundError as e:
    print("--------------------------------------------------------------------------------")
    print("ERRO: Um ou mais arquivos não foram encontrados no caminho:")
    print(f"Caminho da Base: {PATH_BASE_PRODUTOS.resolve()}")
    print(f"Caminho do Relatório de Match: {PATH_RELATORIO_MATCH.resolve()}")
    print("--------------------------------------------------------------------------------")
    print(f"Detalhe do erro: {e}")
    exit()
except FileNotFoundError as e:
    print(f"Erro ao carregar o arquivo: {e}. Verifique se os nomes dos arquivos estão corretos.")
    exit()
except Exception as e:
    print(f"Ocorreu um erro na leitura dos arquivos. Erro: {e}")
    exit()

# Preenchimento de NaN para evitar erros no merge e na comparação
df_base['Produto'] = df_base['Produto'].fillna('')
df_match['Produto na Base (Match)'] = df_match['Produto na Base (Match)'].fillna('')

# --- 3. Preparar o DataFrame de Match para o Merge ---
# Seleciona apenas as colunas necessárias para o merge: Produto (base) e o Código (revista)
df_codigos_revista = df_match[[
    'Código da Revista',
    'Produto na Base (Match)'
]].copy()

# Renomeia a coluna de match para que possamos fazer o merge com 'Produto' da base
df_codigos_revista = df_codigos_revista.rename(
    columns={'Produto na Base (Match)': 'Produto'}
)

# Remove duplicatas, garantindo que cada produto na base tenha apenas um código de revista
df_codigos_revista = df_codigos_revista.drop_duplicates(subset=['Produto'], keep='first')


# --- 4. Realizar o Merge para Atualizar a Base ---
df_base_atualizada = pd.merge(
    df_base,
    df_codigos_revista,
    on='Produto',
    how='left'
)

# --- 5. Lógica de Atualização (Transferir o Código da Revista para o COD da Base) ---

# Garante que a coluna 'COD' (da base) possa ser comparada e atualizada.
df_base_atualizada['COD'] = df_base_atualizada['COD'].astype(object)

# Atualiza a coluna 'COD' APENAS onde ela está vazia/NaN.
df_base_atualizada['COD'] = np.where(
    pd.isna(df_base_atualizada['COD']) | (df_base_atualizada['COD'].astype(str).str.strip() == ''),
    df_base_atualizada['Código da Revista'],
    df_base_atualizada['COD']
)


# --- 6. Finalização e Salvamento ---

# Remove a coluna temporária 'Código da Revista'
df_base_atualizada = df_base_atualizada.drop(columns=['Código da Revista'])

# Garante que a coluna 'COD' não tenha NaNs
df_base_atualizada['COD'] = df_base_atualizada['COD'].fillna('')

# MUDANÇA: Salvar a Base de Dados de Produtos ATUALIZADA como XLSX
df_base_atualizada.to_excel(PATH_BASE_ATUALIZADA, index=False)

print(f"--- Processamento Concluído ---")
print(f"Base de dados atualizada salva em: {PATH_BASE_ATUALIZADA}")
print("\nPrimeiras linhas da base ATUALIZADA (colunas COD, Produto e Marca):")
print(df_base_atualizada[['COD', 'Produto', 'Marca']].head(20))