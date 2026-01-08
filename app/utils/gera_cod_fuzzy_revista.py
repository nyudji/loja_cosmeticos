import pandas as pd
import re
from difflib import SequenceMatcher
from pathlib import Path
import re


# ====================================================================
# 1. Configuração de Caminhos
# ====================================================================

# 1.1. Obtém o caminho do diretório ATUAL onde o script está
# NOTA: __file__ só funciona se o script for executado como um módulo ou arquivo
script_dir = Path(__file__).resolve().parent

# 1.2. Define a pasta 'dados' como sendo um nível acima do script, e depois dentro de 'dados'
# Exemplo: se o script estiver em 'app/utils/script.py', o caminho 'dados/' será 'app/dados/'
caminho_dados = script_dir.parent / 'dados'

# Constrói os caminhos completos para os arquivos de ENTRADA
caminho_revista = caminho_dados / 'revista_produtos_codigos_limpos_v2.csv'
caminho_base = caminho_dados / 'BD_Loja.xlsx'

# Constrói o caminho completo para o arquivo de SAÍDA
output_file = caminho_dados / 'relatorio_match_produtos_final_corrigido.csv'


# Nomes dos arquivos (Certifique-se de que estão na mesma pasta do script)
# Estamos usando os arquivos CSV que você enviou.
# 2. Carregar os Dados
try:
    # Arquivo da Revista (Delimitador: ;)
    # Usa o caminho Path do objeto 'caminho_revista'
    df_revista = pd.read_csv(
        caminho_revista, # MUDANÇA: Usa a variável Path
        sep=';',
        encoding='utf-8',
        on_bad_lines='skip' 
    )

    # Usa o caminho Path do objeto 'caminho_base'
    df_base = pd.read_excel(
        caminho_base, # MUDANÇA: Usa a variável Path
        sheet_name='Produtos' # É crucial especificar o nome da aba (sheet)
    )
    
    # Preenchimento de NaN (Not a Number) para garantir que a coluna 'Produto' 
    df_revista['Produto'] = df_revista['Produto'].fillna('')
    df_base['Produto'] = df_base['Produto'].fillna('')

except FileNotFoundError as e:
    print(f"Erro ao carregar o arquivo: {e}. Verifique se os nomes dos arquivos estão corretos e se estão na mesma pasta.")
    exit()
except Exception as e:
    print(f"Ocorreu um erro na leitura dos arquivos. Verifique o delimitador ou codificação. Erro: {e}")
    exit()

# 2. Função de Limpeza de Texto (CORRIGIDA - PRESERVA DIFERENCIAIS)
def limpar_nome_seletiva_corrigida(nome):
    """
    Limpa e padroniza o nome do produto, removendo APENAS ruídos.
    MANTÉM TODOS OS TERMOS que podem diferenciar o produto (tipo, volume, gênero, etc.).
    """
    if not nome:
        return ""
    nome = str(nome).lower()
    
    # Lista de palavras TOTALMENTE irrelevantes, que NÃO contêm informações de produto:
    palavras_irrelevantes = ['unidade', 'caixa', 'presente', 'com', 'laço', 
                              'do', 'da', 'de', 'para', 'e', 'o', 'a', 'cm', 'natura', 'mini',
                              'embalagem', 'especial', 'kit']
    
    # 1. Remove pontuações e caracteres especiais
    nome = re.sub(r'[^\w\s]', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    
    # 2. Filtra palavras irrelevantes
    palavras = nome.split()
    palavras_filtradas = [p for p in palavras if p not in palavras_irrelevantes and len(p) > 1]
    
    return ' '.join(palavras_filtradas)

# 3. Aplicar a limpeza
df_revista['Produto_Limpo'] = df_revista['Produto'].apply(limpar_nome_seletiva_corrigida)
df_base['Produto_Limpo'] = df_base['Produto'].apply(limpar_nome_seletiva_corrigida)
# Remove linhas vazias na base após a limpeza
df_base = df_base[df_base['Produto_Limpo'] != ''].reset_index(drop=True)

# 4. Configuração de Match 
SCORE_MINIMO_RATIO = 0.75 # 75% de similaridade mínima

def ratio_match(a, b):
    """Calcula o score de similaridade entre duas strings (0.0 a 1.0)"""
    return SequenceMatcher(None, a, b).ratio()

def encontrar_melhor_match(nome_revista_limpo, df_base_limpa):
    """
    Busca o melhor match na lista de nomes da base usando difflib e o score mínimo.
    """
    melhor_score = 0
    melhor_idx = None
    
    for idx in df_base_limpa.index:
        nome_base_limpo = df_base_limpa.loc[idx, 'Produto_Limpo']
        
        # Sequencematcher para calcular a similaridade
        score = ratio_match(nome_revista_limpo, nome_base_limpo)
        
        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx
            
    if melhor_score >= SCORE_MINIMO_RATIO:
        return melhor_idx, round(melhor_score * 100) # Retorna o score em %
    
    return None, 0

# 5. Realizar o Match e Coletar as Informações
resultados = []

for index, row_revista in df_revista.iterrows():
    
    nome_revista_limpo = row_revista['Produto_Limpo']
    
    if not nome_revista_limpo:
        continue

    # Tenta encontrar o match 
    idx_match, score = encontrar_melhor_match(nome_revista_limpo, df_base)
    
    if idx_match is not None:
        # Recuperar a linha completa do produto encontrado na base
        row_base = df_base.loc[idx_match]
        
        # Coletar os dados para o resultado
        preco_venda = row_base['Preço Venda'] if 'Preço Venda' in row_base else None
        codigo_base = row_base['COD'] if 'COD' in row_base else None

        resultados.append({
            'Código da Revista': row_revista['Código'],
            'Produto na Revista': row_revista['Produto'],
            'Página da Revista': row_revista['Página'],
            'Produto na Base (Match)': row_base['Produto'],
            'Código na Base (COD)': codigo_base,
            'Preço Venda na Base': preco_venda,
            'Score de Similaridade (%)': score
        })

# 6. Criar e Salvar o DataFrame de Resultados
df_resultados = pd.DataFrame(resultados)
df_resultados.to_csv(output_file, index=False, encoding='utf-8')

print(f"--- Processamento Concluído ---")
print(f"Estratégia: Limpeza MÍNIMA + Score Mínimo de {SCORE_MINIMO_RATIO*100}%")
print(f"Total de produtos na revista: {len(df_revista)}")
print(f"Total de produtos encontrados na base: {len(df_resultados)}\n")
print(f"Relatório salvo em: {output_file}")
print("\nPrimeiros Resultados:")

# Exibe as 10 primeiras linhas do resultado
colunas_exibir = [
    'Código da Revista', 
    'Produto na Revista', 
    'Produto na Base (Match)', 
    'Código na Base (COD)', 
    'Preço Venda na Base',
    'Score de Similaridade (%)'
]
print(df_resultados[colunas_exibir].sort_values(by='Score de Similaridade (%)', ascending=False).head(10).to_string(index=False))