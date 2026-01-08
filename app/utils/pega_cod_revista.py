import fitz
import re
import pandas as pd
import unicodedata
import os
from pathlib import Path
# ====================================================================
# 1. Configuração de Caminhos
# ====================================================================

# 1.1. Obtém o caminho do diretório ATUAL onde o script está
script_dir = Path(__file__).resolve().parent

# 1.2. Constrói o caminho para 'dados/' (assumindo que 'dados/' está no diretório pai do script)
caminho_dados = script_dir.parent / 'dados'

# Os caminhos agora apontam para dentro da pasta 'dados'
pdf_path = caminho_dados / "revista" / "revista.pdf"
output_csv = caminho_dados / "revista_produtos_codigos_limpos_v2.csv"

def limpar_nome(nome: str) -> str:
    # Normaliza a string para lidar com diferentes representações de caracteres
    nome = unicodedata.normalize("NFKC", nome)
    # Remove caracteres de controle
    nome = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', nome)
    # Substitui aspas e pontos de lista por espaço
    nome = re.sub(r'[\"“”‘’•\xa0]', ' ', nome)
    # Remove números soltos ou no final de palavras (que não são unidades de medida)
    nome = re.sub(r'(?<=[A-Za-zÀ-ÿ])\d+\b', '', nome)
    # Remove caracteres que não são letras, números, espaços ou pontuação específica
    nome = re.sub(r"[^A-Za-zÀ-ÿ0-9\s\-\./,&%()]", "", nome)
    # Garante espaço antes da unidade de medida
    nome = re.sub(r'(\d+)\s*(ml|g|l)\b', r' \1 \2', nome, flags=re.IGNORECASE)
    # Substitui múltiplos espaços por um único
    nome = re.sub(r'\s{2,}', ' ', nome).strip()
    # Remove ponto final no final se houver
    nome = re.sub(r'\s+\.$', '', nome)
    # Capitaliza as palavras
    nome = nome.title()
    return nome

def extrair_produtos(pdf_path: str):
    rows = []
    with fitz.open(pdf_path) as pdf:
        for page_index in range(len(pdf)):
            page = pdf[page_index]
            texto = page.get_text("text")
            linhas = [l.strip() for l in texto.splitlines() if l.strip()]

            # Detecta o título principal da página para prefixar o produto
            titulo_candidates = [l for l in linhas if re.match(r"^Natura\s+[A-ZÁÉÍÓÚÂÊÔÃÕ][a-zA-ZÀ-ÿ\s]+$", l)]
            prefixo_pagina = titulo_candidates[0] if titulo_candidates else "Natura"

            for i, linha in enumerate(linhas):
                cod_match = re.search(r"\((\d{3,6})\)", linha)
                if not cod_match:
                    continue
                codigo = cod_match.group(1)
                nome_raw = ""
                is_refil = False

                # 1. Verifica se a própria linha do código é um Refil
                if re.search(r"\b(REFIL|REFILL)\b", linha, re.IGNORECASE):
                    is_refil = True

                # 2. Tenta encontrar o nome do produto nas linhas anteriores
                for j in range(i-1, max(-1, i-7), -1):
                    anterior = linhas[j]
                    # Ignora bullets ou linhas que são apenas números/preços
                    if anterior.startswith("•") or re.match(r"^\d", anterior) or re.search(r"R\$\s*\d", anterior):
                        continue
                    
                    # Se encontrar uma unidade de medida, assume que é a linha do nome (ou parte dele)
                    if re.search(r"\d+\s*(ml|g|l)\b", anterior.lower()):
                        nome_raw = anterior
                        # Tenta incluir a linha anterior se não for uma descrição de fragrância (tipo: Amadeirado aromático)
                        # e se não for uma linha de preço ou outra informação irrelevante.
                        if j > 0 and not linhas[j-1].startswith("•") and not re.search(r"^\d|R\$\s*\d", linhas[j-1]):
                            # Evita adicionar descrições de fragrância como parte do nome principal, a menos que seja o único contexto
                            if not re.match(r"^[A-Z][a-z]+do\s+[a-z]+$", linhas[j-1]): # Ex: Amadeirado aromático
                                nome_raw = linhas[j-1] + " " + nome_raw
                        break
                
                # 3. Se ainda não encontrou, tenta buscar o nome principal nas 3 linhas anteriores
                if not nome_raw:
                    cand = []
                    # Varre de 1 até 3 linhas antes
                    for j in range(i-1, max(-1, i-4), -1):
                        if not linhas[j].startswith("•") and re.search(r"[A-Za-zÀ-ÿ]", linhas[j]):
                            # Filtra linhas de descrição de fragrância, a menos que seja a única coisa que temos
                            if not re.match(r"^[A-Z][a-z]+do\s+[a-z]+$", linhas[j]):
                                cand.insert(0, linhas[j])
                    nome_raw = " ".join(cand).strip()

                if nome_raw:
                    # Limpa o nome capturado
                    nome_limpo = limpar_nome(nome_raw)
                    
                    # Garante que o nome começa com o prefixo da Natura se não for detectado
                    if not nome_limpo.lower().startswith("natura"):
                        nome_limpo = f"{prefixo_pagina} {nome_limpo}"
                    
                    # Adiciona "Refil" ao nome se for um refil
                    if is_refil and "Refil" not in nome_limpo:
                        # Tenta encontrar a posição da mL/g para inserir Refil antes dela
                        unit_match = re.search(r'(\d+\s*(ML|G|L))', nome_limpo, re.IGNORECASE)
                        if unit_match:
                            # Insere 'Refil' antes da unidade de medida
                            nome_limpo = nome_limpo.replace(unit_match.group(1), f"Refil {unit_match.group(1)}")
                        else:
                            # Adiciona 'Refil' no final se a unidade de medida não for encontrada
                            nome_limpo = f"{nome_limpo} Refil"


                    rows.append({
                        "Produto": nome_limpo,
                        "Código": codigo,
                        "Página": page_index + 1
                    })

    return pd.DataFrame(rows)

if __name__ == "__main__":
    df = extrair_produtos(pdf_path)
    # Remove duplicatas baseadas na combinação Produto, Código e Página
    df = df.drop_duplicates(subset=["Produto", "Código", "Página"]).reset_index(drop=True)
    
    # Salva o resultado em CSV
    df.to_csv(output_csv, index=False, encoding="utf-8-sig", sep=';')
    
    # Exibe o resultado (apenas as primeiras 50 linhas)
    print(f"Extraídos {len(df)} registros. CSV salvo em: {os.path.abspath(output_csv)}")
    print(df.head(50).to_string(index=False))