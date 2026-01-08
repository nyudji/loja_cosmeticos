import nodriver as uc
import pandas as pd
import asyncio
import re
import random
import datetime
from bs4 import BeautifulSoup
from urllib.parse import unquote, quote_plus

# ---- caminhos ----
# Ajuste o nome da coluna de saída para evitar sobrescrever a coluna 'COD' se já existir
CAMINHO_ARQUIVO = "app/dados/BD_Loja.xlsx"
ARQUIVO_SAIDA = "app/dados/BD_Loja_com_codigos_google.xlsx"

# ---- limites e pausas (ajustáveis) ----
PAUSE_MIN = 3.5             # pausa curta mínima entre buscas (s)
PAUSA_MAX = 5.5             # pausa curta máxima entre buscas (s)
PAUSA_POS_BUSCA_MIN = 4.0   # espera após carregar a página (s)
PAUSA_POS_BUSCA_MAX = 5.0

# ---- utilitários ----
def col_name(df, *cands):
    """Detecta o nome correto da coluna de produto."""
    for c in cands:
        if c in df.columns:
            return c
    return None

async def safe_stop(browser):
    """Tenta parar/fechar o browser de forma resiliente."""
    try:
        if hasattr(browser, "stop"):
            await browser.stop()
        elif hasattr(browser, "close"):
            await browser.close()
    except Exception:
        pass

async def buscar_codigos():
    print(f"⌛ Iniciando busca de códigos Natura via Google...")

    # 1. Leitura e Preparação
    df_all = pd.read_excel(CAMINHO_ARQUIVO)
    
    # Identificar a coluna de produto (usando lógica robusta)
    col_produto = col_name(df_all, "Produto", "produto", "Nome", "name")
    if not col_produto:
        raise RuntimeError("Coluna de produto ('Produto', 'Nome', etc.) não encontrada.")
    
    # Garante que a coluna de saída exista
    col_saida = "COD"
    if col_saida not in df_all.columns:
        df_all[col_saida] = ""

    # Filtra produtos para processar (aqueles que ainda não têm o código)
    mask = df_all[col_saida].isna() | (df_all[col_saida].astype(str).str.strip() == "")
    df_proc = df_all[mask].copy().reset_index()
    
    total = len(df_proc)
    print(f"🔍 Total de itens a processar (sem {col_saida}): {total}")
    if total == 0:
        print("Nada para fazer — todos os itens já têm o código de saída.")
        return

    produtos = df_proc[col_produto].tolist()
    indices_orig = df_proc["index"].tolist() # Mapeamento para o DataFrame original
    
    # 2. Iniciar Browser
    try:
        # Tenta conectar primeiro, se não conseguir, inicia
        browser = await uc.connect(host="http://localhost:9222")
        print("🌐 Conectado a uma sessão existente do Chrome.")
    except Exception:
        # Inicia um novo browser
        browser = await uc.start(headless=False)
        print("🆕 Iniciado um novo browser.")
        
    await asyncio.sleep(2.5)
    
    resultados = {} # map index_original -> codigo

    # 3. Loop de Busca
    try:
        for i, produto in enumerate(produtos):
            idx_orig = indices_orig[i]
            codigo = ""
            termo = str(produto).strip()
            
            if not termo:
                print(f"[{i+1}/{total}] [{idx_orig}] Termo vazio -> pulando.")
                resultados[idx_orig] = ""
                continue
            
            print(f"[{i+1}/{total}] [{idx_orig}] Buscando: '{termo}'")
            
            try:
                # Construção da query de busca
                query_encoded = quote_plus(f"site:natura.com.br {termo}") # Adiciona "natura" para focar
                url_busca = f"https://www.google.com/search?q={query_encoded}"
                
                page = await browser.get(url_busca)
                await page.wait_for("body")
                
                await asyncio.sleep(random.uniform(PAUSA_POS_BUSCA_MIN, PAUSA_POS_BUSCA_MAX))

                html = await page.get_content()
                soup = BeautifulSoup(html, "html.parser")

                # Pega todos os links (tags 'a')
                links = [a.get("href") for a in soup.find_all("a", href=True)]
                
                # Descodifica e filtra links que contenham a estrutura de produto da Natura
                # /p/ indica página de produto
                links_filtrados = [unquote(l) for l in links if "/natura.com.br/p/" in l]

                # Tenta extrair o código do primeiro link de produto encontrado
                if links_filtrados:
                    link = links_filtrados[0]
                    # Busca o padrão do código NATBRA-XXXX
                    match = re.search(r"NATBRA-\d+", link)
                    if match:
                        codigo = match.group(0)

                # Se não encontrou no primeiro filtro, tenta uma busca mais ampla (caminho secundário)
                if not codigo:
                    links_secundarios = [unquote(l) for l in links if "natura.com.br" in l]
                    for l in links_secundarios:
                        match = re.search(r"NATBRA-\d+", l)
                        if match:
                            codigo = match.group(0)
                            break
                            
                resultados[idx_orig] = codigo or ""
                print(f"    → {codigo or 'NÃO ENCONTRADO'}")

            except Exception as e:
                resultados[idx_orig] = ""
                print(f"❌ Erro na busca por '{termo}': {e}")
            
            # Pausa aleatória entre as buscas
            await asyncio.sleep(random.uniform(PAUSE_MIN, PAUSA_MAX))

    finally:
        # 4. Finalização
        await safe_stop(browser)

    # 5. Salvar Resultados
    for idx, cod in resultados.items():
        if cod: # Atualiza apenas se encontrou um código
            df_all.at[idx, col_saida] = cod

    df_all.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"\n✅ Finalizado. Arquivo salvo em: {ARQUIVO_SAIDA}")

if __name__ == "__main__":
    asyncio.run(buscar_codigos())