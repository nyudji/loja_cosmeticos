import pandas as pd
from pathlib import Path

# Importa os módulos que você criou
from trata_observacoes import TratamentoObservacoes
from trata_serial import gerar_codigos # Importa apenas a função

# --- Configuração de Caminhos ---
# __file__ aponta para 'app/tratamento_final.py'. .parent leva para 'app'.
script_dir = Path(__file__).parent 

# Caminho de entrada (subindo um nível e entrando em 'dados')
caminho_arquivo_entrada = script_dir / '..' / 'dados' / 'natura.xlsx' 

# Caminho de saída para o resultado final
caminho_arquivo_saida = script_dir / '..' / 'dados' / 'natura_FINAL.xlsx'

def executar_tratamento_completo():
    
    # --- ETAPA 1: LEITURA DO ARQUIVO ---
    try:
        print(f"Lendo o arquivo: {caminho_arquivo_entrada}")
        df = pd.read_excel(caminho_arquivo_entrada, sheet_name='Vendas')
        print("Arquivo lido com sucesso.")
    except Exception as e:
        print(f"ERRO: Não foi possível ler o arquivo. {e}")
        return # Encerra a função em caso de erro

    # --- ETAPA 2: TRATAMENTO DAS OBSERVAÇÕES ---
    # Instancia a classe e chama o método de limpeza
    tratador_obs = TratamentoObservacoes()
    df_limpo = tratador_obs.limpar_observacoes(df)

    # --- ETAPA 3: GERAÇÃO DE CÓDIGOS E REORDENAÇÃO ---
    colunas_novas = ["Serial Produto", "Cod Produto"]
    df_final = gerar_codigos(df_limpo, colunas_novas)

    print("Validação final: Verificando as 5 primeiras linhas com novos códigos.")
    print(df_final.head())

    # --- ETAPA 4: SALVAMENTO DO ARQUIVO ---
    try:
        df_final.to_excel(caminho_arquivo_saida, index=False)
        print(f"\n✅ SUCESSO! Arquivo final salvo em: {caminho_arquivo_saida}")
    except Exception as e:
        print(f"ERRO ao salvar o arquivo: {e}")

if __name__ == "__main__":
    # É aqui que o seu script principal será executado
    executar_tratamento_completo()