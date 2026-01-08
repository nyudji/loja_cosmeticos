import pandas as pd
from rapidfuzz import process, fuzz
from pathlib import Path
import re
import numpy as np
import glob
from datetime import datetime
import uuid

# --- CONFIGURAÇÃO DE CAMINHOS ---
script_dir = Path(__file__).resolve().parent
caminho_dir_nf = script_dir.parent / 'dados' /  'nf' /  'excel'
# Arquivo principal que contém as abas 'Produtos' e 'Movimento'
caminho_bd_loja = script_dir.parent / 'dados' / 'BD_Loja2.xlsx' 
# Arquivo de saída
caminho_saida = script_dir.parent / 'dados' / 'BD_Loja_Atualizado.xlsx'

# --- DICIONÁRIO DE ABREVIAÇÕES (Mantido) ---
DICIONARIO_NF = {
    r'\bSAB BAR\b': 'SABONETE EM BARRA',
    r'\bSAB LIO\b': 'SABONETE LIQUIDO',
    r'\bSAB LIQ\b': 'SABONETE LIQUIDO',
    r'\bAMEIX BAU\b': 'AMEIXA E FLOR DE BAUNILHA',
    r'\bCR CORP\b': 'CREME CORPORAL',
    r'\bDES PER\b': 'DESODORANTE CORPORAL',
    r'\bDESODORANTE COLÔNIA\b': 'COLONIA',
    r'\bCOLÔNIA\b': 'COLONIA',
    r'\bDES ROLLON\b': 'DESODORANTE ANTITRANSPIRANTE ROLL ON',
    r'\bDES SPRAY\b': 'DESODORANTE SPRAY',
    r'\bEDP\b': 'DEO PARFUM',
    r'\bSH\b': 'SHAMPOO',
    r'\bCOND\b': 'CONDICIONADOR',
    r'\bHID\b': 'HIDRATANTE',
    r'\bCORP\b': 'CORPORAL',
    r'\bCR\b': 'CREME',
    r'\bSAB BARRA\b': 'SABONETE EM BARRA',
    r'\bDES\b': 'DESODORANTE',
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
}

# --- FUNÇÕES AUXILIARES ---

def gerar_id_venda():
    '''Gera um ID curto aleatório tipo E-27B7FB'''
    return f"E-{str(uuid.uuid4())[:6].upper()}"

def extrair_cod_natura_nf(texto):
    if pd.isna(texto): return None
    texto = str(texto).upper().strip()
    match = re.search(r'^\s*\*?(\d+)-', texto)
    if match: return match.group(1) 
    return None

def limpar(texto):
    if pd.isna(texto): return ""
    texto = str(texto).upper()
    texto = re.sub(r'^\d+\s*', '', texto)
    texto = re.sub(r'[^A-Z0-9 ÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

def limpar_nf(texto):
    if pd.isna(texto): return ""
    texto = str(texto).upper()
    texto = re.sub(r'^\s*\*\s*', '', texto).strip()
    texto = re.sub(r'^\d{4,6}-', '', texto).strip()
    texto = re.sub(r' BC R\$[^)]+| ICMS-ST[^)]+| FCI.+| NV| VPN\d*', '', texto).strip()
    for padrao_regex, substituicao in DICIONARIO_NF.items():
        texto = re.sub(padrao_regex, substituicao, texto)
    texto = re.sub(r'[^A-Z0-9 ÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ./,-]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

def melhor_match(nome_nf, lista_produtos):
    match, score, _ = process.extractOne(
        nome_nf,
        lista_produtos,
        scorer=fuzz.token_sort_ratio
    )
    return pd.Series([match, score])

# --- 1. CARREGAMENTO DA BASE DE DADOS (PRODUTOS E MOVIMENTO) ---

print("Carregando bases de dados...")
xls = pd.ExcelFile(caminho_bd_loja)

# Carrega Produtos
try:
    nome_aba_prod = 'Produtos' if 'Produtos' in xls.sheet_names else xls.sheet_names[0]
    produtos = pd.read_excel(xls, sheet_name=nome_aba_prod)
    produtos['Produto'] = produtos['Produto'].apply(limpar)
except Exception as e:
    print(f"Erro ao ler aba de produtos: {e}")
    exit()

# Carrega Movimento
colunas_movimento = ['Data', 'ID do Prod', 'Produto', 'Cliente', 'Tipo de Movimento', 
                     'Quantidade', 'Preço Custo Total', 'Preço Venda Total', 
                     'Observações', 'Status', 'Data Previsão', 'Forma de Pagamento', 'ID_Venda']

if 'Movimento' in xls.sheet_names:
    movimento_antigo = pd.read_excel(xls, sheet_name='Movimento')
else:
    movimento_antigo = pd.DataFrame(columns=colunas_movimento)

# --- 2. LER E PROCESSAR NFs ---

lista_dfs_nf = []
arquivos_nf = list(caminho_dir_nf.glob('nf*.xlsx'))

if not arquivos_nf:
    print(f"ERRO: Nenhum arquivo 'nf*.xlsx' encontrado em: {caminho_dir_nf}")
    exit()

for caminho_nf in arquivos_nf:
    print(f"Lendo NF: {caminho_nf.name}")
    try:
        df_temp = pd.read_excel(caminho_nf)
        lista_dfs_nf.append(df_temp)
    except Exception as e:
        print(f"AVISO: Erro ao ler {caminho_nf.name}: {e}")

if lista_dfs_nf:
    nf_completa = pd.concat(lista_dfs_nf, ignore_index=True)
else:
    exit()

# --- 3. LIMPEZA E MATCHING ---

print("Realizando Fuzzy Matching...")
nf_completa['Cod Natura'] = nf_completa['DESCRIÇÃO'].apply(extrair_cod_natura_nf)
nf_completa['DESCRIÇÃO LIMPA'] = nf_completa['DESCRIÇÃO'].apply(limpar_nf)
nf_completa[['Produto_Match', 'Similaridade']] = nf_completa['DESCRIÇÃO LIMPA'].apply(lambda x: melhor_match(x, produtos['Produto']))

try:
    # Identificar colunas Quantidade e Valor
    col_qtd = next((c for c in nf_completa.columns if 'QUANT' in c.upper() or 'QTD' in c.upper()), None)
    col_vlr = next((c for c in nf_completa.columns if 'VALOR TOTAL' in c.upper() or 'V. TOTAL' in c.upper() or 'TOTAL' in c.upper()), None)
    
    if not col_qtd or not col_vlr:
        raise ValueError("Colunas de Quantidade ou Valor não encontradas na NF.")
        
    relacionados = nf_completa[nf_completa['Similaridade'] > 62].copy()
    relacionados = relacionados.rename(columns={col_qtd: 'QTD_NF', col_vlr: 'VLR_NF'})
    
    # *** CRIA A COLUNA JÁ FORMATADA 'NATBRA-XXXX' ***
    relacionados['Cod_Formatado'] = 'NATBRA-' + relacionados['Cod Natura'].astype(str)
    
except Exception as e:
    print(f"Erro ao processar colunas da NF: {e}")
    exit()

# --- 4. CRIAR NOVAS ENTRADAS PARA O MOVIMENTO ---

print(f"Gerando {len(relacionados)} novas entradas de estoque...")

novas_entradas = pd.DataFrame()

novas_entradas['Data'] = datetime.now().date()
# Usando a coluna formatada com NATBRA-
novas_entradas['ID do Prod'] = relacionados['Cod_Formatado']
novas_entradas['Produto'] = relacionados['Produto_Match'].str.title()
novas_entradas['Cliente'] = 'NATURA'
novas_entradas['Tipo de Movimento'] = 'ENTRADA'
novas_entradas['Quantidade'] = relacionados['QTD_NF']
novas_entradas['Preço Custo Total'] = relacionados['VLR_NF']
novas_entradas['Preço Venda Total'] = '' 
novas_entradas['Observações'] = 'Importação Automática NF'
novas_entradas['Status'] = 'PAGO'
novas_entradas['Data Previsão'] = ''
novas_entradas['Forma de Pagamento'] = 'Boleto/Pix'
novas_entradas['ID_Venda'] = [gerar_id_venda() for _ in range(len(relacionados))]

# Concatena com o movimento antigo
movimento_atualizado = pd.concat([movimento_antigo, novas_entradas], ignore_index=True)

# --- 5. ATUALIZAÇÃO DA BASE DE PRODUTOS ---

# Filtra os melhores matches
relacionados_unique = relacionados.sort_values(by='Similaridade', ascending=False).drop_duplicates(subset=['Produto_Match'], keep='first')

# Faz o merge trazendo o Cod_Formatado (NATBRA-...)
produtos_atualizado = pd.merge(
    produtos, 
    relacionados_unique[['Produto_Match', 'Cod_Formatado']], 
    left_on='Produto', 
    right_on='Produto_Match', 
    how='left'
)

# Lógica para preencher o COD na aba Produtos
if 'COD' not in produtos_atualizado.columns:
    produtos_atualizado['COD'] = produtos_atualizado['Cod_Formatado']
else:
    # Preenche onde é NaN
    produtos_atualizado['COD'] = produtos_atualizado['COD'].fillna(produtos_atualizado['Cod_Formatado'])
    # Preenche onde é string vazia
    produtos_atualizado['COD'] = produtos_atualizado.apply(
        lambda row: row['Cod_Formatado'] if (pd.isna(row['COD']) or str(row['COD']).strip() == '') else row['COD'], axis=1
    )

# Limpeza
cols_drop = ['Produto_Match', 'Cod_Formatado', 'Similaridade']
produtos_atualizado = produtos_atualizado.drop(columns=[c for c in cols_drop if c in produtos_atualizado.columns])

# --- 6. SALVAMENTO FINAL ---

print(f"Salvando arquivo consolidado em: {caminho_saida}")

with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
    produtos_atualizado.to_excel(writer, sheet_name='Produtos', index=False)
    movimento_atualizado.to_excel(writer, sheet_name='Movimento', index=False)

print("Processo concluído com sucesso!")