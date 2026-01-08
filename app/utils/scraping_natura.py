import nodriver as uc
import pandas as pd
import asyncio
import re
import random
import datetime
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
from fuzzywuzzy import fuzz, process

# ---- caminhos ----
CAMINHO_ARQUIVO = "app/dados/BD_Loja.xlsx"
ARQUIVO_SAIDA = "app/dados/BD_Loja_com_codigos.xlsx"

# ---- user-agents ----
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/122 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Chrome/120 Safari/537.36"
]

# ---- limites e pausas (ajustáveis) ----
SIMILARIDADE_MINIMA = 54
PAUSE_EVERY = 60            # pausa longa a cada N buscas
PAUSE_LONG_MIN = 60         # pausa longa mínima (s)
PAUSE_LONG_MAX = 180        # pausa longa máxima (s)
PAUSA_MIN = 5               # pausa curta mínima entre buscas (s)
PAUSA_MAX = 6               # pausa curta máxima entre buscas (s)
PAUSA_POS_BUSCA_MIN = 2.8   # espera após abrir a página do produto (s)
PAUSA_POS_BUSCA_MAX = 3.3

# ---- utilitários ----
def col_name(df, *cands):
    for c in cands:
        if c in df.columns:
            return c
    return None

def montar_termo_busca(row, col_colecao, col_nome, col_volume):
    parts = []
    for col in (col_colecao, col_nome, col_volume):
        if col and pd.notna(row.get(col, "")):
            parts.append(str(row[col]).strip())
    termo = " ".join(p for p in parts if p)
    return re.sub(r"\s+", " ", termo).strip()

def eh_refil(row, col_produto, col_nome):
    texts = []
    if col_produto and pd.notna(row.get(col_produto, "")):
        texts.append(str(row[col_produto]).lower())
    if col_nome and pd.notna(row.get(col_nome, "")):
        texts.append(str(row[col_nome]).lower())
    joined = " ".join(texts)
    return "refil" in joined

async def extrair_codigo(html):
    m = re.search(r"NATBRA-\d+", html)
    return m.group(0) if m else ""

async def safe_stop(browser):
    """Tenta parar/fechar o browser de forma resiliente."""
    try:
        await asyncio.sleep(1)
        if hasattr(browser, "stop"):
            try:
                await browser.stop()
                return
            except Exception:
                pass
        if hasattr(browser, "close"):
            try:
                await browser.close()
                return
            except Exception:
                pass
        # última tentativa sem await se existir
        f = getattr(browser, "stop", None)
        if callable(f):
            try:
                f()
            except Exception:
                pass
    except Exception:
        pass

# ---- busca individual (com detecção Access Denied / refresh / retry) ----
async def buscar_codigo_natura(browser, termo_busca, nome_produto_completo, refil=False):
    try:
        query = quote_plus(termo_busca)
        url_busca = f"https://www.natura.com.br/s/produtos?busca={query}"

        page = await browser.get(url_busca)
        await page.wait_for("body")
        await asyncio.sleep(random.uniform(8, 12))  # espera inicial

        produtos_html = []
        for _ in range(15):
            html = await page.get_content()

            # detectar bloqueio
            if ("Access Denied" in html) or ("You don't have permission" in html):
                print("\n🚫 Detectado bloqueio da Natura (Access Denied). Tentando refresh...\n")
                # 1) tentar reload simples
                try:
                    await page.reload()
                    await asyncio.sleep(random.uniform(4, 7))
                    html_retry = await page.get_content()
                    if ("Access Denied" not in html_retry) and ("You don't have permission" not in html_retry):
                        print("🔄 Refresh removeu o bloqueio! Continuando...\n")
                        html = html_retry
                    else:
                        raise Exception("refresh_failed")
                except Exception:
                    # 2) fallback: pausa longa + reconectar (sem fechar Chrome)
                    wait_time = random.uniform(180, 220)
                    print(f"⏸ Pausa longa {int(wait_time)}s (bloqueio). Aguardando...\n")
                    await asyncio.sleep(wait_time)

                    # mudar user-agent (aplicado nas próximas requisições)
                    user_agent = random.choice(USER_AGENTS)
                    print(f"🌐 Novo User-Agent sugerido: {user_agent} (aplicado ao reconectar)\n")

                    # tentar reconectar ao Chrome remoto (não fecha o Chrome)
                    try:
                        browser = await uc.connect(host="http://localhost:9222")
                        print("🔄 Reconectado ao Chrome remoto com sucesso.")
                    except Exception as e:
                        print(f"⚠️ Falha ao reconectar: {e} (continuando com o browser atual)")

                    # refazer a busca recursivamente com a nova sessão/estado
                    return await buscar_codigo_natura(browser, termo_busca, nome_produto_completo, refil)

            # parse dos produtos
            soup = BeautifulSoup(html, "html.parser")
            produtos_html = soup.find_all(
                "h4",
                class_="text-wrap text-ellipsis line-clamp-2 text-body-2 md:text-body-1"
            )
            if produtos_html:
                break
            await asyncio.sleep(1.5)
        await asyncio.sleep(1.5)
        if not produtos_html:
            print("    ⚠️ Nenhum produto encontrado na página.")
            return ""

        # filtrar irrelevantes (kits/refil/banner)
        await asyncio.sleep(1.5)
        produtos_validos = []
        for p in produtos_html:
            nome_item = p.get_text(strip=True).lower()
            if any(x in nome_item for x in ["cookies", "privacidade", "banner", "promo"]):
                continue
            if not refil and any(x in nome_item for x in ["kit", "refil", "miniatura"]):
                continue
            if refil and "refil" not in nome_item:
                continue
            produtos_validos.append(p)

        if not produtos_validos:
            print("    ⚠️ Nenhum produto válido após filtro.")
            return ""

        # fuzzy matching entre o nome completo da planilha e títulos retornados
        nomes_site = [p.get_text(strip=True) for p in produtos_validos]
        melhor = process.extractOne(nome_produto_completo, nomes_site, scorer=fuzz.token_sort_ratio)
        if not melhor:
            return ""
        melhor_match, score, _ = (melhor[0], melhor[1], (melhor[2] if len(melhor) > 2 else 0))
        print(f"    → Produto (planilha): '{nome_produto_completo}'")
        print(f"    → Melhor correspondência: '{melhor_match}' (similaridade {score}%)")
        await asyncio.sleep(random.uniform(2, 3))

        if score < SIMILARIDADE_MINIMA:
            print(f"    ⚠️ Similaridade {score}% menor que {SIMILARIDADE_MINIMA}% → ignorando.")
            return ""

        # localizar o elemento correspondente e navegar
        escolhido = next((p for p in produtos_validos if p.get_text(strip=True) == melhor_match), None)
        if not escolhido:
            return ""

        # tentar clicar; se não for possível, pegar o href do pai <a> e navegar
        try:
            await escolhido.click()
        except Exception:
            parent = escolhido.find_parent("a")
            if parent and parent.get("href"):
                href = parent.get("href")
                if href.startswith("/"):
                    href = "https://www.natura.com.br" + href
                await page.get(href)
            else:
                return ""

        await page.wait_for("body")
        await asyncio.sleep(random.uniform(PAUSA_POS_BUSCA_MIN, PAUSA_POS_BUSCA_MAX))

        html_prod = await page.get_content()
        await asyncio.sleep(random.uniform(5.2, 7.3))
        codigo = await extrair_codigo(html_prod)
        return codigo or ""

    except Exception as e:
        print(f"⚠️ Erro em buscar_codigo_natura('{termo_busca}'): {e}")
        await asyncio.sleep(2)
        return ""

# ---- rotina principal (processa somente linhas sem COD) ----
async def buscar_codigos():
    global searches_since_pause
    searches_since_pause = 0

    # lê arquivo original
    df_all = pd.read_excel(CAMINHO_ARQUIVO)

    # garante que a coluna COD exista
    if "COD" not in df_all.columns:
        df_all["COD"] = ""

    # filtra só os que estão sem COD (NaN ou string vazia)
    mask = df_all["COD"].isna() | (df_all["COD"].astype(str).str.strip() == "")
    df_proc = df_all[mask].copy().reset_index()  # reset_index para manter índice original em 'index'

    total = len(df_proc)
    print(f"🔍 Total de itens sem COD a processar: {total}")
    if total == 0:
        print("Nada para fazer — todos os itens já têm COD.")
        return

    # detectar nomes de colunas
    col_colecao = col_name(df_proc, "Coleção", "Colecao")
    col_nome = col_name(df_proc, "Nome", "nome", "Product", "produto")
    col_volume = col_name(df_proc, "Volume", "volume")
    col_produto = col_name(df_proc, "Produto", "produto")

    if not col_nome or not col_colecao or not col_volume:
        raise RuntimeError("Colunas necessárias não encontradas (Coleção, Nome, Volume).")

    # montar termos e flags
    df_proc["termo_busca"] = df_proc.apply(lambda r: montar_termo_busca(r, col_colecao, col_nome, col_volume), axis=1)
    df_proc["is_refil"] = df_proc.apply(lambda r: eh_refil(r, col_produto, col_nome), axis=1)

    termos = df_proc["termo_busca"].tolist()
    refis = df_proc["is_refil"].tolist()
    nomes_produtos = df_proc[col_produto].astype(str).tolist()
    indices_orig = df_proc["index"].tolist()  # índices no df_all

    # conectar ao Chrome remoto (tente connect, se não existir, usa start)
    try:
        browser = await uc.connect(host="http://localhost:9222")
    except Exception:
        browser = await uc.start()
    await asyncio.sleep(2.5)

    resultados = {}  # map index_original -> codigo

    try:
        for i, termo in enumerate(termos):
            idx_orig = indices_orig[i]
            nome_completo = nomes_produtos[i]
            refil_flag = refis[i]

            if not termo or termo.strip() == "":
                print(f"[{idx_orig}] termo vazio -> pulando")
                resultados[idx_orig] = ""
                continue
            print(f'Produto da planilha: "{nome_completo}"')
            print(f"[{i}/{total}] [{idx_orig}] Buscando: '{termo}'  (refil={refil_flag})")
            # pequena espera antes da requisição para variar comportamento humano
            await asyncio.sleep(random.uniform(4.2,4.8))
            codigo = await buscar_codigo_natura(browser, termo, nome_completo, refil=refil_flag)
            resultados[idx_orig] = codigo or ""
            print(f"    → {resultados[idx_orig] or 'NÃO ENCONTRADO'}")

            searches_since_pause += 1
            if searches_since_pause >= PAUSE_EVERY:
                wait_time = random.uniform(PAUSE_LONG_MIN, PAUSE_LONG_MAX)
                hora = datetime.datetime.now().strftime("%H:%M:%S")
                print(f"\n⏸ Pausa longa iniciada às {hora} — aguardando {int(wait_time)}s...\n")
                await asyncio.sleep(wait_time)
                searches_since_pause = 0

            await asyncio.sleep(random.uniform(PAUSA_MIN, PAUSA_MAX))

    finally:
        # garantir fechamento seguro do browser
        await safe_stop(browser)

    # aplicar resultados no df_all
    for idx, cod in resultados.items():
        df_all.at[idx, "COD"] = cod

    # salvar
    df_all.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"\n✅ Finalizado. Arquivo salvo em: {ARQUIVO_SAIDA}")

if __name__ == "__main__":
    asyncio.run(buscar_codigos())
