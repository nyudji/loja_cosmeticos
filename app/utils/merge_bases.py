import pandas as pd

# === 1️⃣ Ler os dois arquivos ===
base = pd.read_excel("app/dados/BD_Loja.xlsx", sheet_name="Produtos")
atualizado = pd.read_excel("app/dados/BD_Loja_Produtos_COMCOD_NF.xlsx", sheet_name="Sheet1")

# === 2️⃣ Garantir que os nomes dos produtos sejam strings para comparação segura ===
base["Produto"] = base["Produto"].astype(str).str.strip()
atualizado["Produto"] = atualizado["Produto"].astype(str).str.strip()

# === 3️⃣ Criar um dicionário {produto: codigo} a partir do arquivo atualizado ===
mapa_codigos = dict(zip(atualizado["Produto"], atualizado["COD"]))

# === 4️⃣ Preencher apenas os códigos vazios no arquivo base ===
def preencher_codigo(row):
    if pd.isna(row["COD"]) or str(row["COD"]).strip() == "":
        return mapa_codigos.get(row["Produto"], row["COD"])
    return row["COD"]

base["COD"] = base.apply(preencher_codigo, axis=1)

# === 5️⃣ Salvar novo arquivo ===
base.to_excel("app/dados/BD_Loja_Completado.xlsx", index=False)

print("✅ Códigos preenchidos com sucesso! Arquivo salvo como BD_Loja_Completado.xlsx")
