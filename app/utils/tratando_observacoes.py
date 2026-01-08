import pandas as pd
from pathlib import Path
import re 

# --- Configuração do Caminho (Adaptado ao seu projeto 'app' e 'dados') ---
script_dir = Path(__file__).parent 
# Caminho do arquivo de ENTRADA (agora com a extensão .xlsx)
caminho_arquivo_entrada = script_dir / '..' / 'dados' / 'natura.xlsx' 
# Caminho do arquivo de SAÍDA (também .xlsx)
caminho_arquivo_saida = script_dir / '..' / 'dados' / 'natura_limpo_obs.xlsx'

try:
    # MUDANÇA CRUCIAL: Usando pd.read_excel()
    # ATENÇÃO: Se o nome da planilha não for 'Vendas', mude o sheet_name
    df = pd.read_excel(
        caminho_arquivo_entrada, 
        sheet_name='Vendas' 
    )
    print("Arquivo Excel (natura.xlsx) lido com sucesso.")

except FileNotFoundError:
    print(f"Erro: O arquivo não foi encontrado em '{caminho_arquivo_entrada}'.")
    exit()
except Exception as e:
    print(f"Ocorreu um erro ao ler o arquivo: {e}")
    exit()


# --- Limpeza da Coluna 'Observações' ---
coluna_obs = "Observações"

# 1. Pré-tratamento: Converte a coluna para string e preenche NaNs com string vazia
df[coluna_obs] = df[coluna_obs].astype(str).fillna('')

print(f"Iniciando limpeza na coluna '{coluna_obs}'...")

# 2. Limpeza Principal: Substitui barras por |
# Troca a barra normal (/) por |
df[coluna_obs] = df[coluna_obs].str.replace('/', '|', regex=False)

# Troca a barra invertida (\) por | (usando '\\' para escapar a barra)
df[coluna_obs] = df[coluna_obs].str.replace('\\', '|', regex=False) 


# --- Tratamentos Extras para Padronização ---

# 3. Remover espaços em excesso que podem surgir da troca (ex: " | ")
df[coluna_obs] = df[coluna_obs].str.replace(' | ', '|', regex=False)
df[coluna_obs] = df[coluna_obs].str.replace(' |', '|', regex=False)
df[coluna_obs] = df[coluna_obs].str.replace('| ', '|', regex=False)

# 4. Remover múltiplos separadores seguidos (Ex: 'A||B' vira 'A|B')
df[coluna_obs] = df[coluna_obs].apply(lambda x: re.sub(r'\|+', '|', x))

# 5. Remove espaços vazios no início/fim e converte para maiúsculas (padronização)
df[coluna_obs] = df[coluna_obs].str.strip().str.upper()


# --- Salvando o Arquivo Limpo ---
try:
    # MUDANÇA CRUCIAL: Usando df.to_excel()
    # index=False evita salvar o índice numérico padrão do pandas
    df.to_excel(caminho_arquivo_saida, index=False)
    print(f"\n✅ Arquivo Excel limpo salvo com sucesso em: {caminho_arquivo_saida}")
except Exception as e:
    print(f"Erro ao salvar o arquivo: {e}")