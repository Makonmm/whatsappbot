from urllib.parse import quote
from time import sleep
import os
import webbrowser
import openpyxl
import keyboard
import pyautogui

current_directory = os.path.dirname(os.path.abspath(__file__))
clientes_file = os.path.join(current_directory, 'clientes.xlsx')

# Abrir o navegador e esperar a autenticação no WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
print("Aguarde enquanto o WhatsApp Web é carregado...")
sleep(15)  # Espera 30 segundos para autenticação no WhatsApp Web

# Carregar o arquivo Excel
try:
    work_book = openpyxl.load_workbook(clientes_file)
    pagina_clientes = work_book['Sheet1']
except FileNotFoundError:
    print("Arquivo 'clientes.xlsx' não encontrado. Certifique-se de que o arquivo está no diretório correto.")
    exit()

# Iterar sobre as linhas do arquivo Excel
for linha in pagina_clientes.iter_rows(min_row=2):
    name = linha[0].value
    tel = linha[1].value
    venc = linha[2].value

    # Construir a mensagem
    msg = f"Olá, {name}, seu boleto vencerá no dia: {venc.strftime(
        '%d/%m/%Y')}. Favor fazer o pagamento pelo link: https://www.link.com"
    try:
        # Abrir o link do WhatsApp com a mensagem
        link_msg_whatsapp = f"https://web.whatsapp.com/send?phone={
            tel}"
        # Abrir a aba do WhatsApp Web apenas uma vez
        webbrowser.open(link_msg_whatsapp)
        sleep(10)  # Esperar 10 segundos para abrir o WhatsApp Web e carregar a página

        # Enviar a mensagem e pressionar Enter três vezes
        for _ in range(3):
            keyboard.write(msg)  # Digitar a mensagem
            # Pressionar Enter para enviar a mensagem
            sleep(2)
            keyboard.press_and_release('enter')
            sleep(2)  # Esperar 2 segundos após pressionar "Enter"

        # Fechar a guia do navegador após enviar todas as mensagens
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)  # Esperar 2 segundos para fechar a guia antes de abrir a próxima

    except Exception as exc:
        print(f"Erro ao enviar mensagem para {name}: {exc}")
        # Registrar o erro em um arquivo CSV
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{name}, {tel}, {exc}\n')
