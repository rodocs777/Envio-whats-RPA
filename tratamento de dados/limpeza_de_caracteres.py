import pandas as pd
import re

# Caminho para o arquivo CSV
caminho_arquivo = r'C:\Users\Rodocs\Documents\projetos\RPA-Python\numeros_telefone.csv'

try:
    # Ler a tabela de números de telefone
    df = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1')
    print("Arquivo CSV lido com sucesso.")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_arquivo} não foi encontrado.")
    exit(1)
except pd.errors.EmptyDataError:
    print(f"Erro: O arquivo {caminho_arquivo} está vazio.")
    exit(1)
except pd.errors.ParserError:
    print(f"Erro: O arquivo {caminho_arquivo} contém dados malformados.")
    exit(1)
except Exception as e:
    print(f"Erro desconhecido: {e}")
    exit(1)

# Função para remover caracteres indesejados
def remove_caracteres_indesejados(numero):
    # Remover tudo que não for dígito
    return re.sub(r'\D', '', numero)

# Aplicar a função à coluna de números de telefone
df['numero_limpo'] = df['numero'].astype(str).apply(remove_caracteres_indesejados)

# Exibir a tabela resultante
print(df)

# Caminho para salvar o arquivo CSV resultante
caminho_arquivo_limpo = r'C:\Users\Rodocs\Documents\projetos\RPA-Python\numeros_telefone_limpo.csv'

# Salvar a tabela resultante
df.to_csv(caminho_arquivo_limpo, index=False)
print(f"Tabela limpa salva em: {caminho_arquivo_limpo}")
