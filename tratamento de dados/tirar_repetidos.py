import pandas as pd

# Carregar as tabelas
primeira_tabela = pd.read_csv(r"C:\Users\Rodocs\Documents\projetos\RPA-Python\tratamento_de_dados\tirar_numeros_repetidos\primeiratabela.csv")
segunda_tabela = pd.read_csv(r"C:\Users\Rodocs\Documents\projetos\RPA-Python\tratamento_de_dados\tirar_numeros_repetidos\segundatabela.csv")

# Supondo que a coluna de números em ambas as tabelas seja chamada de 'numero'
# Filtrar os números que não estão na segunda tabela
filtrada_tabela = primeira_tabela[~primeira_tabela['numero'].isin(segunda_tabela['numero'])]

# Salvar a tabela filtrada em um novo arquivo CSV
filtrada_tabela.to_csv(r'C:\Users\Rodocs\Documents\projetos\RPA-Python\tratamento_de_dados\tirar_numeros_repetidos\tabela_filtrada.csv', index=False)

print("A tabela filtrada foi salva como 'tabela_filtrada.csv'")
