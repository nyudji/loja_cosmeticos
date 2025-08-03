import easyocr
import pandas as pd
import os

# Caminho da imagem escaneada
imagem = 'anotacao.jpg'  # Substitua pelo nome da sua imagem

# Inicializa o OCR em português
reader = easyocr.Reader(['pt'])

# Extrai texto da imagem
resultado = reader.readtext(imagem, detail=0)

# Exibe resultado bruto
print("Texto extraído:")
for linha in resultado:
    print(linha)

# Processa para tentar separar em colunas (Produto, Quantidade, Preço)
produtos = []
for linha in resultado:
    partes = linha.split()
    
    # Ajuste conforme o seu padrão de anotação
    if len(partes) >= 2:
        nome = ' '.join(partes[:-1])
        preco = partes[-1]
        produtos.append([nome, preco])
    else:
        produtos.append([linha, ""])

# Cria DataFrame
df = pd.DataFrame(produtos, columns=["Produto", "Preço"])

# Salva em Excel
df.to_excel("produtos_extraidos.xlsx", index=False)
print("Arquivo Excel salvo como produtos_extraidos.xlsx")