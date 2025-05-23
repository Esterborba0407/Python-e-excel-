import pandas as pd
import os

# Garanti que a pasta data está no projeto
os.makedirs("data", exist_ok=True)

# Dados de exemplo
dados = {
    "Produto": ["Notebook", "Mouse", "Teclado", "Monitor", "Cabo HDMI", "Notebook", "Mouse"],
    "Categoria": ["Informática", "Periféricos", "Periféricos", "Informática", "Acessórios", "Informática", "Periféricos"],
    "Quantidade": [5, 10, 7, 3, 15, 2, 8],
    "Valor Unitário": [3500, 50, 120, 800, 25, 3400, 55]
}

# Criei o DataFrame
df = pd.DataFrame(dados)

# Caminho da planilha a ser criada
caminho_arquivo = "data/base_dados.xlsx"

# Salvei a planilha
df.to_excel(caminho_arquivo, index=False)

print(f"✅ Planilha base criada com sucesso em: {caminho_arquivo}")
