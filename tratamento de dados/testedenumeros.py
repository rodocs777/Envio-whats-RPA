# _*_ coding: utf-8 _*_

import pyautogui
import time
import pandas as pd

time.sleep(1.5)

# Caminho absoluto para o arquivo de texto
caminho_arquivo = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\clientes.csv"
# Caminho absoluto para a imagem "numero_nao_encontrado.png"
caminho_imagem_nao_encontrado = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\numero_nao_encontrado.png"

# Caminho absoluto para o arquivo de texto
caminho_arquivo = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\clientes.csv"
caminho_imagem_iniciar_conversa = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\versao_2_whatsnew\iniciar_conversa.png"

try:
    # Ler a tabela de clientes
    tabela = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1')
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

# Ajustar o intervalo entre os comandos do PyAutoGUI
pyautogui.PAUSE = 0.3

# Verificar se a coluna "codigo" existe na tabela
if "codigo" not in tabela.columns:
    print("Erro: A coluna 'codigo' não foi encontrada no arquivo CSV.")
    exit(1)

# Exibir a tabela para verificação
print(tabela)

# Lista para armazenar os números que não foram encontrados
numeros_nao_encontrados = []

# Passo 4: inserir um WhatsApp
for linha in tabela.index:
    try:
        # Pegar da tabela o valor do campo que queremos preencher
        codigo = tabela.loc[linha, "codigo"]

        # Aba de pesquisa
        pyautogui.hotkey('ctrl', 'l')

        # Preencher o campo
        pyautogui.write(str(codigo))
        
        # Pressionar no botão para buscar o contato
        pyautogui.press("enter")
        time.sleep(2)
        
        # Verificar se o botão iniciar conversa está presente na tela
        pyautogui.locateCenterOnScreen(caminho_imagem_iniciar_conversa)
        pyautogui.click(caminho_imagem_iniciar_conversa)
        time.sleep(2)

        pyautogui.click(x=687, y=413) #inicar conversa

        time.sleep(10)

        # Verificar se o número não foi encontrado
        if pyautogui.locateOnScreen(caminho_imagem_nao_encontrado, confidence=0.9):
            print(f"Número não encontrado: {codigo}")
            numeros_nao_encontrados.append(linha)

    except pyautogui.FailSafeException:
        print("Erro: O script foi interrompido pelo usuário movendo o mouse para o canto da tela.")
        exit(1)
    except Exception as e:
        print(f"Erro durante a automação na linha {linha}: {e}")
        continue

# Remover as linhas dos números não encontrados da tabela original
if numeros_nao_encontrados:
    df_nao_encontrados = tabela.loc[numeros_nao_encontrados]
    tabela = tabela.drop(numeros_nao_encontrados)

    # Salvar os números não encontrados em um arquivo CSV
    df_nao_encontrados.to_csv(r"C:\Users\Rodocs\Documents\projetos\RPA-Python\numeros_nao_encontrados.csv", index=False)
    print(f"Números não encontrados foram salvos em {r'C:/Users/Rodocs/Documents/projetos/RPA-Python/numeros_nao_encontrados.csv'}")

    # Salvar a tabela original atualizada sem os números não encontrados
    tabela.to_csv(caminho_arquivo, index=False)
    print(f"Tabela original atualizada foi salva em {caminho_arquivo}")
else:
    print("Todos os números foram encontrados com sucesso.")

