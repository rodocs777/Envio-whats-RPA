# _*_ coding: utf-8 _*_

import pandas
 
# Cria o dir "Logs" se ele não existir
os.makedirs("C:/automacao/Logs", exist_ok=True)

# Obter a data e hora atual e formatar como string
data_hora_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

######## Caminho absoluto para o arquivo de log, incluindo a data e hora atual no nome do arquivo
caminho_log = f'C:/automacao/logs/log_{data_hora_atual}.log'

# Configuração de logging
logging.basicConfig(filename=caminho_log, filemode='w', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
time.sleep(1.5)

######## Caminho variáveis para os arquivos
caminho_texto = r"C:\automacao\texto\texto.txt"
caminho_imagem = r'C:\automacao\imagem'

# Caminho absoluto para os arquivos
caminho_csv = r"C:\automacao\tabela\clientes.csv"
caminho_imagem_iniciar_conversa = r"C:\automacao\frames\iniciar_conversa.png"
caminho_imagem_caiu_whats = r"C:\automacao\frames\caiu_whats.png"
caminho_digitesuamensagem = r"C:\automacao\frames\digitesuamensagem.png"
caminho_usar_whatsweb = r"C:\automacao\frames\usar_whatsweb.png"
caminho_campo_mensagem = r"C:\automacao\frames\campo_mensagem.png"

# Ler o conteúdo do arquivo de texto
with open(caminho_texto, 'r', encoding='utf-8') as arquivo:
    conteudo = arquivo.read()

try:
    # Ler a tabela de clientes
    tabela = pd.read_csv(caminho_csv, encoding='ISO-8859-1')
    logging.info("Arquivo CSV lido com sucesso.")
except FileNotFoundError as e:
    logging.error(f"Erro: O arquivo {caminho_csv} não foi encontrado - {e}")
    exit(1)
except pd.errors.EmptyDataError as e:
    logging.error(f"Erro: O arquivo {caminho_csv} está vazio - {e}")
    exit(1)
except pd.errors.ParserError as e:
    logging.error(f"Erro: O arquivo {caminho_csv} contém dados malformados - {e}")
    exit(1)
except Exception as e:
    logging.error(f"Erro desconhecido ao ler o arquivo CSV: {e}")
    exit(1)

# Ajustar o intervalo entre os comandos do PyAutoGUI
pyautogui.PAUSE = 0.3

# Verificar se a coluna "codigos" existe na tabela
if "codigos" not in tabela.columns:
    logging.error("Erro: A coluna 'codigos' não foi encontrada no arquivo CSV.")
    exit(1)

# Exibir a tabela para verificação
logging.info(f"Tabela lida:\n{tabela}")


# Passo 4: inserir um WhatsApp
for linha in tabela.index:
    try:

        # Pegar da tabela o valor do campo que queremos preencher
        codigos = tabela.loc[linha, "codigos"]

        # Exibir no terminal o código do cliente atual
        logging.info(f"Lendo o código do cliente: {codigos}")

        # Aba de pesquisa
        pyautogui.hotkey('ctrl', 'l')

        # Preencher o campo
        pyautogui.write(str(codigos))
        
        # Pressionar no botão para buscar o contato
        pyautogui.press("enter")
        time.sleep(1.5)
        
        # Botão iniciar conversa está presente na tela
        pyautogui.locateOnScreen(caminho_imagem_iniciar_conversa, confidence=0.9)
        pyautogui.click(caminho_imagem_iniciar_conversa)
        time.sleep(1.5)

        pyautogui.locateCenterOnScreen(caminho_usar_whatsweb, confidence=0.9)
        pyautogui.click(caminho_usar_whatsweb)
        time.sleep(10)

        pyautogui.locateOnScreen(caminho_digitesuamensagem, confidence=0.9)
        pyautogui.click(caminho_digitesuamensagem)
        # Copiar o conteúdo para a área de transferência
        pyperclip.copy(conteudo)
        # Colar a mensagem na barra de conversa
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1.0)
        
        # Anexar imagens
        pyautogui.hotkey('win', 'e')
        time.sleep(1.0)
        pyautogui.hotkey('ctrl', 'l')
        pyperclip.copy(caminho_imagem)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(1.0)
        pyautogui.hotkey('ctrl', 'a')  # seleciona tudo
        pyautogui.hotkey('ctrl', 'c')  # Copia imagens
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1.0)
    
        pyautogui.locateCenterOnScreen(caminho_campo_mensagem, confidence=0.9)
        pyautogui.click(caminho_campo_mensagem)
        
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2.0)
        pyautogui.press('enter')
        time.sleep(2.0)


    except pyautogui.FailSafeException:
        logging.error("Erro: O script foi interrompido pelo usuário movendo o mouse para o canto da tela.")
        exit(1)
    except Exception as e:
        logging.error(f"A mensgem pode não ter sido enviada: {codigos} VERIFIQUE {e}")
        continue
    
    # Informar no log que a tabela foi finalizada
logging.info("Fim da tabela.")
