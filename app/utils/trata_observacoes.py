
import pandas as pd
import re

class TratamentoObservacoes:
    """
    Classe responsável por limpar e padronizar a coluna 'Observações' 
    de um DataFrame, substituindo barras por '|' e garantindo consistência.
    """
    
    def __init__(self, coluna_observacoes="Observações"):
        # Define o nome da coluna a ser tratada
        self.coluna_obs = coluna_observacoes

    def limpar_observacoes(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Aplica a lógica de limpeza na coluna de observações do DataFrame.
        """
        # Cria uma cópia para evitar o SettingWithCopyWarning no Pandas
        df_processado = df.copy()

        # 1. Pré-tratamento: Converte a coluna para string e preenche NaNs com string vazia
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].astype(str).fillna('')

        print(f"-> Iniciando limpeza na coluna '{self.coluna_obs}'...")

        # 2. Limpeza Principal: Substitui barras por |
        # Troca a barra normal (/) por |
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].str.replace('/', '|', regex=False)

        # Troca a barra invertida (\) por | (usando '\\' para escapar a barra)
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].str.replace('\\', '|', regex=False) 


        # --- Tratamentos Extras para Padronização ---

        # 3. Remover espaços em excesso que podem surgir da troca (ex: " | ")
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].str.replace(' | ', '|', regex=False)
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].str.replace(' |', '|', regex=False)
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].str.replace('| ', '|', regex=False)


        # 4. Remover múltiplos separadores seguidos (Ex: 'A||B' vira 'A|B')
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].apply(lambda x: re.sub(r'\|+', '|', x))

        # 5. Remove espaços vazios no início/fim e converte para maiúsculas (padronização para análise)
        df_processado[self.coluna_obs] = df_processado[self.coluna_obs].str.strip().str.upper()

        print("-> Limpeza de observações concluída.")
        
        return df_processado

# O bloco abaixo garante que a classe não execute nada se o arquivo for importado.
if __name__ == "__main__":
    print("Este módulo não deve ser executado diretamente.")