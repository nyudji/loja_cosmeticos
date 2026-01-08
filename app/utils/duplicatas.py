import pandas as pd
import numpy as np # Importado para usar o valor de célula vazia (NaN)

# 1. Defina o nome do arquivo de entrada e de saída
ARQUIVO_ENTRADA = 'app/dados/BD_Loja.xlsx'
# Nome indicando que os códigos repetidos foram anulados, mantendo todas as linhas
ARQUIVO_SAIDA = 'app/dados/BD_Loja_CODIGOS_REPETIDOS_ANULADOS.xlsx'

# 2. Defina o nome da aba (planilha) dentro do arquivo Excel
# Se a aba for a primeira, use 0. Se souber o nome, mude para o nome exato (ex: 'Produtos').
NOME_DA_ABA = 0 

# 3. Defina o nome da coluna de códigos
COLUNA_CHAVE = 'COD'

try:
    # 4. Carrega O ARQUIVO EXCEL INTEIRO (TODAS AS COLUNAS)
    df = pd.read_excel(ARQUIVO_ENTRADA, sheet_name=NOME_DA_ABA)
    
    linhas_antes = len(df)
    
    # 5. Cria uma máscara que marca TODAS as linhas que possuem um código duplicado
    # keep=False: Marca como True TODAS as ocorrências de um código que se repete 2 ou mais vezes.
    mascara_duplicada = df.duplicated(subset=[COLUNA_CHAVE], keep=False)
    
    linhas_duplicadas_anuladas = mascara_duplicada.sum()
    
    # 6. ANULA o valor na coluna 'COD' para TODAS as linhas identificadas
    # Se o código aparecia nas linhas 10 e 80, AMBAS as linhas (10 e 80) terão o código anulado (np.nan).
    df.loc[mascara_duplicada, COLUNA_CHAVE] = np.nan 
    
    # 7. Salva o DataFrame (com todas as 777 linhas) em um novo arquivo Excel
    # index=False evita a coluna de índice do Pandas.
    df.to_excel(ARQUIVO_SAIDA, index=False)
    
    # 8. Relatório Final
    print("------------------------------------------")
    print("✅ Processo Concluído! Estrutura de Linhas Preservada.")
    print(f"Total de linhas na planilha: {linhas_antes}")
    print(f"Ocorrências (original + repetição) ANULADAS na coluna '{COLUNA_CHAVE}': {linhas_duplicadas_anuladas}")
    print(f"O arquivo final **{ARQUIVO_SAIDA}** contém **{linhas_antes} linhas** (nenhuma linha foi excluída).")
    print("------------------------------------------")

except FileNotFoundError:
    print(f"❌ Erro: O arquivo '{ARQUIVO_ENTRADA}' não foi encontrado. Verifique o caminho.")
except KeyError:
    print(f"❌ Erro: A coluna '{COLUNA_CHAVE}' não foi encontrada na planilha '{NOME_DA_ABA}'. Verifique o cabeçalho.")
except Exception as e:
    print(f"❌ Ocorreu um erro inesperado: {e}")