import pandas as pd

# Carregar dados do CSV
df = pd.read_csv(r"C:\Users\Rodocs\Documents\projetos\RPA-Python\origem.csv")

# Função para verificar se um número de telefone começa com "11"
def verifica_telefone(numero):
    return str(numero).startswith("11")

# Criar um DataFrame com os números que começam com "11"
df_com11 = df[df['numero'].apply(verifica_telefone)]

# Criar um DataFrame sem os números que começam com "11"
df_sem11 = df[~df['numero'].apply(verifica_telefone)]

# Salvar os números que começam com "11" em um novo arquivo CSV
output_path_com11 = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\origemcom11.csv"
df_com11.to_csv(output_path_com11, index=False)

# Salvar a tabela original sem os números que começam com "11" em um novo arquivo CSV
output_path_sem11 = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\origemsem11.csv"
df_sem11.to_csv(output_path_sem11, index=False)

print(f"Números que começam com '11' foram salvos em {output_path_com11}")
print(f"Tabela original sem os números que começam com '11' foi salva em {output_path_sem11}")
