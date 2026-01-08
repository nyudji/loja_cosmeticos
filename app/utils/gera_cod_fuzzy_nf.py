import pandas as pd
from rapidfuzz import process, fuzz
from pathlib import Path
import re
import numpy as np
import glob

# --- CONFIGURAÇÃO DE CAMINHOS ---
script_dir = Path(__file__).resolve().parent
caminho_dir_nf = script_dir.parent / 'dados' /  'nf' /  'excel'
caminho_produtos = script_dir.parent / 'dados' / 'BD_Loja2.xlsx'

# --- DICIONÁRIO DE ABREVIAÇÕES (Mantido e Crucial) ---
DICIONARIO_NF = {
    #   Termos abreviados na NF (Chave) -> Termo completo na Base (Valor)
    r'\bSAB BAR\b': 'SABONETE EM BARRA',
    r'\bSAB LIO\b': 'SABONETE LIQUIDO',
    r'\bSAB LIQ\b': 'SABONETE LIQUIDO',
    r'\bAMEIX BAU\b': 'AMEIXA E FLOR DE BAUNILHA',
    r'\bCR CORP\b': 'CREME CORPORAL',
    r'\bDES PER\b': 'DESODORANTE CORPORAL',
    r'\bDESODORANTE COLÔNIA\b': 'COLONIA',
    r'\bDESODORANTE COLÔNIA\b': 'COLONIA',
    r'\bCOLÔNIA\b': 'COLONIA',
    r'\bDES PER\b': 'DESODORANTE CORPORAL',
    r'\bDES ROLLON\b': 'DESODORANTE ANTITRANSPIRANTE ROLL ON',
    r'\bDES SPRAY\b': 'DESODORANTE SPRAY',
    r'\bEDP\b': 'DEO PARFUM',
    r'\bSH\b': 'SHAMPOO',
    r'\bCOND\b': 'CONDICIONADOR',
    r'\bHID\b': 'HIDRATANTE',
    r'\bCORP\b': 'CORPORAL',
    r'\bCR\b': 'CREME',
    r'\bSAB BAR\b': 'SABONETE EM BARRA',
    r'\bSAB BARRA\b': 'SABONETE EM BARRA',
    r'\bDES ROLLON\b': 'DESODORANTE  ANTITRANSPIRANTE ROLL ON',
    r'\bDES SPRAY\b': 'DESODORANTE SPRAY',
    r'\bDES\b': 'DESODORANTE',
    r'\bSH\b': 'SHAMPOO',
    r'\bCOND\b': 'CONDICIONADOR',
    r'\bEDP\b': 'DEO PARFUM',
    r'\bCOL\b': 'COLÔNIA',
    r'\bAGUA COL\b': 'ÁGUA DE COLÔNIA',
    r'\bMMBB\b': 'MAMÃE E BEBÊ',
    r'\bSR N\b': 'SENHOR N',
    r'\bMASC\b': 'MASCULINO',
    r'\bFEM\b': 'FEMININO',
    r'\bFL\b': 'FLOR DE',
    r'\bAMT\b': 'AMAZÔNIA',
    r'\bPRG\b': 'PERFUMADO',
    r'\bRF\b': 'REFIL',
    r'\bPER\b': 'PERFUMADO',
    r'\s+': ' '
    # Adicione aqui todas as abreviações que você identifica na sua NF
}
# --- FUNÇÕES DE LIMPEZA E TRATAMENTO ---

def extrair_cod_natura_nf(texto):
    '''Extrai o código do produto natura do texto da NF.'''
    if pd.isna(texto): return None
    texto = str(texto).upper().strip()
    match = re.search(r'^\s*\*?(\d+)-', texto)
    if match: return match.group(1) 
    return None

def limpar(texto):
    '''Limpa a base original (BD_Loja).'''
    if pd.isna(texto): return ""
    texto = str(texto).upper()
    texto = re.sub(r'^\d+\s*', '', texto)
    texto = re.sub(r'[^A-Z0-9 ÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

def limpar_nf(texto):
    '''Limpa a descrição da NF, remove códigos e aplica dicionário.'''
    if pd.isna(texto): return ""
        
    texto = str(texto).upper()
    
    # 1. Pré-limpeza: Remove '*', código (4-6 dígitos) e hífen, e VPN. (CRUCIAL)
    texto = re.sub(r'^\s*\*\s*', '', texto).strip()
    texto = re.sub(r'^\d{4,6}-', '', texto).strip()
    texto = re.sub(r' BC R\$[^)]+| ICMS-ST[^)]+| FCI.+| NV| VPN\d*', '', texto).strip()
    
    # 2. Expansão de abreviações (CRUCIAL)
    for padrao_regex, substituicao in DICIONARIO_NF.items():
        texto = re.sub(padrao_regex, substituicao, texto)
        
    # 3. Limpeza final
    texto = re.sub(r'[^A-Z0-9 ÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ./,-]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    
    return texto

# FUNÇÃO reordenar_descricao FOI REMOVIDA

def melhor_match(nome_nf):
    '''Encontra o produto da base com o produto da nf usando Fuzzy Matching.'''
    match, score, _ = process.extractOne(
        nome_nf,
        produtos['Produto'],
        scorer=fuzz.token_sort_ratio # Este scorer é robusto à ordem das palavras
    )
    return pd.Series([match, score])

# --- PROCESSO PRINCIPAL: CARREGAMENTO ---

# 1. Carrega a base de produtos UMA VEZ e a limpa
produtos = pd.read_excel(caminho_produtos)
produtos['Produto'] = produtos['Produto'].apply(limpar)

# 2. LER E CONCATENAR TODAS AS NFs (nf*.xlsx)
lista_dfs_nf = []
arquivos_nf = list(caminho_dir_nf.glob('nf*.xlsx'))

if not arquivos_nf:
    print(f"ERRO: Nenhum arquivo 'nf*.xlsx' encontrado no diretório: {caminho_dir_nf}")
    exit()

for caminho_nf in arquivos_nf:
    print(f"Lendo o arquivo: {caminho_nf.name}")
    try:
        df_temp = pd.read_excel(caminho_nf)
        lista_dfs_nf.append(df_temp)
    except Exception as e:
        print(f"AVISO: Não foi possível ler o arquivo {caminho_nf.name}. Erro: {e}")

if lista_dfs_nf:
    nf_completa = pd.concat(lista_dfs_nf, ignore_index=True)
    print(f"Total de linhas lidas em todas as NFs: {len(nf_completa)}")
else:
    print("ERRO: Nenhuma NF foi lida com sucesso. Encerrando o script.")
    exit()

# --- PROCESSO PRINCIPAL: LIMPEZA E MATCHING ---

# 3. Extrai o código natura da NF (código da nota fiscal)
nf_completa['Cod Natura'] = nf_completa['DESCRIÇÃO'].apply(extrair_cod_natura_nf)

# 4. Aplica apenas a limpeza robusta
nf_completa['DESCRIÇÃO LIMPA'] = nf_completa['DESCRIÇÃO'].apply(limpar_nf)

# *** IMPORTANTE: A coluna para matching É IGUAL à coluna limpa (sem reordenação) ***
nf_completa['DESCRIÇÃO PARA MATCH'] = nf_completa['DESCRIÇÃO LIMPA'] 

# 5. Aplica o Fuzzy Matching usando a DESCRIÇÃO LIMPA
nf_completa[['Produto', 'Similaridade']] = nf_completa['DESCRIÇÃO PARA MATCH'].apply(melhor_match)

# 6. Filtra e salva resultados
relacionados = nf_completa[nf_completa['Similaridade'] > 62].copy() 
relacionados.to_excel(script_dir.parent / 'dados' / 'nf_relacionados_multi_final_sem_reordenar.xlsx', index=False)


## --- MERGE E ATUALIZAÇÃO DA BASE ---

# 7. Configuração do Merge
FILE_BASE_ATUALIZADA = script_dir.parent / 'dados' / 'BD_Loja_Produtos_COMCOD_NF.xlsx'
produtos['Produto'] = produtos['Produto'].str.title()
relacionados['Produto'] = relacionados['Produto'].str.title()

# Prioriza o melhor match
relacionados_merge = relacionados.copy()
relacionados_merge = relacionados_merge.sort_values(by='Similaridade', ascending=False).drop_duplicates(subset=['Produto'], keep='first')


# 8. Realiza o Merge
produtos_atualizado = pd.merge(
    produtos.copy(), 
    relacionados_merge[['Produto', 'Cod Natura']],
    on='Produto',
    how='left',
    suffixes=('_base', '_nf') 
)

# 9. Lógica de Atualização (Substitui o COD vazio/NaN na base pelo Cod Natura encontrado)
if 'COD_base' not in produtos_atualizado.columns:
     produtos_atualizado = produtos_atualizado.rename(columns={'COD': 'COD_base'})
     
produtos_atualizado['COD_base'] = produtos_atualizado['COD_base'].astype(object)

produtos_atualizado['COD_base'] = np.where(
    pd.isna(produtos_atualizado['COD_base']) | (produtos_atualizado['COD_base'].astype(str).str.strip() == ''),
    produtos_atualizado['Cod Natura'],
    produtos_atualizado['COD_base']
)


# 10. Finalização e Salvamento
produtos_atualizado = produtos_atualizado.rename(columns={'COD_base': 'COD'})
# Remove as colunas temporárias. Nota: 'DESCRIÇÃO PARA MATCH' e 'DESCRIÇÃO LIMPA' agora podem ser as mesmas.
cols_to_drop = [col for col in ['Cod Natura', 'Similaridade', 'DESCRIÇÃO LIMPA', 'DESCRIÇÃO PARA MATCH'] if col in produtos_atualizado.columns]
produtos_atualizado = produtos_atualizado.drop(columns=cols_to_drop)

produtos_atualizado['COD'] = produtos_atualizado['COD'].fillna('')

produtos_atualizado.to_excel(FILE_BASE_ATUALIZADA, index=False)

print(f"\n--- PROCESSO CONCLUÍDO ---")
print(f"Base de dados atualizada salva em: {FILE_BASE_ATUALIZADA}")
print("O Fuzzy Matching agora usa a DESCRIÇÃO LIMPA original (sem reordenar a marca).")