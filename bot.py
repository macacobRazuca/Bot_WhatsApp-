import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep

#Abrindo o planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Planilha1']

#Pegando as informaçoões nescessarias
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value

#Mensagem personalizada
    mensagem = f'Olá, {nome}, tudo bem? Estamos passando para avisar sobre a promoção dos nosos produtos...'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        #Abrindo o WhatsApp
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        pyautogui.hotkey('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        #Separando os erros do projeto
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
