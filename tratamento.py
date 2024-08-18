import pandas as pd
import re

# Caminho para o arquivo CSV original
caminho_arquivo_original = "origem.csv"
# Caminho para o novo arquivo CSV resultante
caminho_arquivo_resultante = "origem_caracteres_sem11_9_wame5511.csv"

try:
    # Ler a tabela de números de telefone
    df = pd.read_csv(caminho_arquivo_original, encoding='ISO-8859-1')
    print("Arquivo CSV lido com sucesso.")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_arquivo_original} não foi encontrado.")
    exit(1)
except pd.errors.EmptyDataError:
    print(f"Erro: O arquivo {caminho_arquivo_original} está vazio.")
    exit(1)
except pd.errors.ParserError:
    print(f"Erro: O arquivo {caminho_arquivo_original} contém dados malformados.")
    exit(1)
except Exception as e:
    print(f"Erro desconhecido: {e}")
    exit(1)

# Função para remover caracteres indesejados
def remove_caracteres_indesejados(numero):
    # Remover tudo que não for dígito
    return re.sub(r'\D', '', numero)

# Função para remover o prefixo "11" do número de telefone
def remove_prefixo_11(numero):
    if numero.startswith("11"):
        return numero[2:]
    return numero

# Função para verificar se o número começa com "9"
def verifica_comeca_com_9(numero):
    return numero.startswith("9")

# Função para verificar se o número tem 9 caracteres
def verifica_nove_caracteres(numero):
    return len(numero) == 9

# Função para adicionar a URL base na frente do número de telefone
def adicionar_url_base(numero):
    return f"https://wa.me/5511{numero}"

# Limpar caracteres indesejados
df['numero'] = df['numero'].astype(str).apply(remove_caracteres_indesejados)

# Remover o prefixo "11"
df['numero'] = df['numero'].apply(remove_prefixo_11)

# Filtrar apenas números que começam com "9"
df = df[df['numero'].apply(verifica_comeca_com_9)]

# Filtrar apenas números que tenham 9 caracteres
df = df[df['numero'].apply(verifica_nove_caracteres)]

# Remover duplicatas
df = df.drop_duplicates(subset='numero')

# Adicionar a URL base
df['numero'] = df['numero'].apply(adicionar_url_base)

# Selecionar apenas a coluna numero
df_resultante = df[['numero']]

# Adicionar cabeçalho "numero" ao DataFrame resultante
df_resultante.columns = ['numero']

# Exibir a tabela resultante
print(df_resultante)

# Salvar a tabela resultante
df_resultante.to_csv(caminho_arquivo_resultante, index=False)
print(f"Tabela resultante salva em: {caminho_arquivo_resultante}")
