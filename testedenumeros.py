# _*_ coding: utf-8 _*_

import pyautogui
import time
import pandas as pd
import signal
import logging

# Configuração de logging
logging.basicConfig(filename='C:/Users/Rodocs/Documents/projetos/RPA-Python/testenumeros_log.log', filemode='w', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

time.sleep(3)

# Caminho absoluto para o arquivo de texto
caminho_arquivo = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\clientes.csv"
# Caminho absoluto para a imagem "numero_nao_encontrado.png"
caminho_imagem_nao_encontrado = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\testedenumeros\numero_nao_encontrado.png"
# Caminho absoluto para a imagem "iniciar_conversa.png"
caminho_imagem_iniciar_conversa = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\frames\iniciar_conversa.png"
caminho_usar_whatsweb = r"C:\Users\Rodocs\Documents\projetos\RPA-Python\frames\usar_whatsweb.png"

# Função para salvar o progresso
def salvar_progresso():
    global tabela, numeros_nao_encontrados
    tabela.to_csv(caminho_arquivo, index=False)
    if numeros_nao_encontrados:
        df_nao_encontrados = tabela.loc[numeros_nao_encontrados]
        df_nao_encontrados.to_csv(r"C:\Users\Rodocs\Documents\projetos\RPA-Python\numeros_nao_encontrados.csv", index=False)
    logging.info("Progresso salvo.")

# Tratador de sinal para salvar ao interromper a automação
def tratar_interrupcao(sig, frame):
    logging.info("Interrupção detectada. Salvando progresso...")
    salvar_progresso()
    exit(0)

# Registrar o tratador de sinal
signal.signal(signal.SIGINT, tratar_interrupcao)
signal.signal(signal.SIGTERM, tratar_interrupcao)

try:
    # Ler a tabela de clientes
    tabela = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1')
    logging.info("Arquivo CSV lido com sucesso.")
except FileNotFoundError:
    logging.error(f"Erro: O arquivo {caminho_arquivo} não foi encontrado.")
    exit(1)
except pd.errors.EmptyDataError:
    logging.error(f"Erro: O arquivo {caminho_arquivo} está vazio.")
    exit(1)
except pd.errors.ParserError:
    logging.error(f"Erro: O arquivo {caminho_arquivo} contém dados malformados.")
    exit(1)
except Exception as e:
    logging.error(f"Erro desconhecido ao ler o arquivo CSV: {e}")
    exit(1)

# Ajustar o intervalo entre os comandos do PyAutoGUI
pyautogui.PAUSE = 0.3

# Verificar se a coluna "numero" existe na tabela
if "numero" not in tabela.columns:
    logging.error("Erro: A coluna 'numero' não foi encontrada no arquivo CSV.")
    exit(1)

# Exibir a tabela para verificação
logging.info(f"Tabela lida:\n{tabela}")

# Lista para armazenar os números que não foram encontrados
numeros_nao_encontrados = []

# Passo 4: inserir um WhatsApp
for linha in tabela.index:
    try:
        # Pegar da tabela o valor do campo que queremos preencher
        numero = tabela.loc[linha, "numero"]

        # Exibir no terminal o código do cliente atual
        logging.info(f"Lendo o código do cliente: {numero}")

        # Aba de pesquisa
        pyautogui.hotkey('ctrl', 'l')

        # Preencher o campo
        pyautogui.write(str(numero))
        
        # Pressionar no botão para buscar o contato
        pyautogui.press("enter")
        time.sleep(1)
        
        # Verificar se o botão iniciar conversa está presente na tela
        botao_iniciar = pyautogui.locateCenterOnScreen(caminho_imagem_iniciar_conversa)
        if botao_iniciar:
            pyautogui.click(botao_iniciar)
            time.sleep(1)
            pyautogui.locateCenterOnScreen(caminho_usar_whatsweb, confidence=0.9)
            pyautogui.click(caminho_usar_whatsweb)
            time.sleep(16) #####################################################
        else:
            logging.warning(f"Erro: Botão 'iniciar conversa' não encontrado para o código {numero}")
            numeros_nao_encontrados.append(linha)
            continue

        # Verificar se o número não foi encontrado
        if pyautogui.locateOnScreen(caminho_imagem_nao_encontrado, confidence=0.9):
            logging.warning(f"Número não encontrado: {numero}")
            numeros_nao_encontrados.append(linha)

    except pyautogui.FailSafeException:
        logging.error("Erro: O script foi interrompido pelo usuário movendo o mouse para o canto da tela.")
        salvar_progresso()
        exit(1)
    except Exception as e:
        logging.error(f"Erro durante a automação na linha {linha}: {e}")
        continue

# Remover as linhas dos números não encontrados da tabela original
if numeros_nao_encontrados:
    df_nao_encontrados = tabela.loc[numeros_nao_encontrados]
    tabela = tabela.drop(numeros_nao_encontrados)

    # Salvar os números não encontrados em um arquivo CSV
    df_nao_encontrados.to_csv(r"C:\Users\Rodocs\Documents\projetos\RPA-Python\numeros_nao_encontrados.csv", index=False)
    logging.info(f"Números não encontrados foram salvos em {r'C:/Users/Rodocs/Documents/projetos/RPA-Python/numeros_nao_encontrados.csv'}")

    # Salvar a tabela original atualizada sem os números não encontrados
    tabela.to_csv(caminho_arquivo, index=False)
    logging.info(f"Tabela original atualizada foi salva em {caminho_arquivo}")
else:
    logging.info("Todos os números foram encontrados com sucesso.")
