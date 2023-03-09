import pandas as pd

# Lê a planilha
df = pd.read_excel('AlertaAutomacao.xlsx')

# Exibe as primeiras linhas de todas as colunas
print(df.head())