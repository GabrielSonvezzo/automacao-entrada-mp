import pandas as pd
arquivo = r"C:\Projeto_Notas\Analise Embarque MP - 2025.xlsx"
df = pd.read_excel(arquivo)
print("As colunas que eu encontrei no seu Excel são:")
print(list(df.columns))