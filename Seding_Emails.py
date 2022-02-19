import time
import pyautogui
import pyperclip
import time
import pyautogui
import pyperclip
import pandas as pd
# Anrir windons, chromeer, nova guia, pyperclip, colar na guia de pesquisa, esperar 5 seg, clicar na tabela, apertar 3 botoes, download.

#Abrindo Windons

pyautogui.press('winleft')
pyautogui.write('chrome')
pyautogui.press('enter')

time.sleep(3)


link = "https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym"

pyperclip.copy(link)

time.sleep(3)

pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(5)

#Fazendo Dowloand do arquivo

pyautogui.click(x=1344, y=453)
time.sleep(3)
pyautogui.click(x=1653, y=190)
time.sleep(3)
pyautogui.click(x=1429, y=705)



df = pd.read_excel(r'Vendas - Dez.xlsx')

faturamento = df['Valor Final'].sum()
quantidade = df['Quantidade'].sum()

#Enviando Email
pyautogui.press('winleft')
pyautogui.write('chrome')
time.sleep(2)

pyautogui.press('enter')

time.sleep(6)



email = 'https://mail.google.com/'

pyperclip.copy(email)

pyautogui.hotkey('ctrl', 'v')

pyautogui.press('enter')


time.sleep(8)


emails =  ['g70@gmail.com','gab70@gmail.com']
for cada in emails:
    
    #apertando botão para escrever novo arquivo
    pyautogui.click(x = 114, y = 214)
    time.sleep(4)
    pyautogui.write(cada)
    pyautogui.press('tab')
    pyautogui.press('tab')
    msg = 'Relatorio de vendas'
    pyperclip.copy(msg)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    texto = f'''
    Prezados, bom dia.
    O faturamento de ontem foi de: R$ {faturamento: .2f}
    A quantidade de produtos foi: {quantidade: ,}
    
    Abs
    Gabriela 
    
    '''
    pyperclip.copy(texto)
                    
    pyautogui.hotkey('ctrl', 'v')                
    
    # apertando no botão de enviar 
    
    pyautogui.click(x=1158, y=997)
    
    time.sleep(3)


#----------------Encontrando posições------------

time.sleep(5)
pyautogui.position()
pyautogui.click(x=1158, y=997)
    
