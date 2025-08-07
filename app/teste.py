import pandas as pd

ARQUIVO_EXCEL = "app/dados/controle_financeiro.xlsx"
xls = pd.ExcelFile(ARQUIVO_EXCEL)
print("Abas encontradas:", xls.sheet_names)