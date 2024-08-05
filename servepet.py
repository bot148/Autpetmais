import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(8)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('Guitest.xlsx.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá {nome}, A Loja Pet Mais começou o mês de agosto com promoção !\n\n Compre qualquer ração úmida no site com 10% off usando o cupom UMIDA10\n\n Alimente seu pet com o melhor, pagando menos.\n\n Ingredientes selecionados, sabores irresistíveis e nutrição balanceada.\n\n Não perca essa oportunidade de mimar seu amigo de quatro patas com produtos de alta qualidade por um preço especial.\n\n Aproveite agora!\n\n https://www.lojapetmais.com.br/'



    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
    # clicar em anexo
        anexo = pyautogui.locateCenterOnScreen('anexo4.png')
        sleep(2)
        pyautogui.click(anexo.x, anexo.y)
     # clicou em fotos e vídeos
        pyautogui.click(771,526,duration=1)
    # escolheu o documento
        pyautogui.click(589,497,duration=1)
    # abrir documento
        pyautogui.click(850,675,duration=1)
    # enviar documento
        pyautogui.click(1845,932,duration=1) 
    # fechar aba anterior
        pyautogui.click(310,24, duration=0.5)
    # clique de mizericórdia
        pyautogui.click(1054, 253, duration=0.5) 


    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
           arquivo.write(f'{nome},{telefone}{os.linesep}')
        # fechar aba anterior
        pyautogui.click(310,24, duration=0.5)
        # clique de mizericórdia
        pyautogui.click(1054, 253, duration=0.5) 
