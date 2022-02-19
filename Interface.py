from PySimpleGUI import PySimpleGUI as sg
#Layout

#definidno opções de datas
sg.theme('DarkRed1')



#Criando layout de datas recentes ou antigas
layout3 = [
    [sg.Text("Deseja selecionar dados recentes ou arquivos anteriores ao ano atual?",size=(60, 2), font=("Helvetica", 15), text_color='white')],
    [sg.Text("")],
    [sg.Button('Recentes'),sg.Button('Antigos')]
    
]

janela3 = sg.Window('Tel'), layout3
ru = sg.Window('Seleção Arquivos', layout3).read(close=True)
print(ru[0])
     #Aqui colocamos um if para executar a função de acordo com o botão selecionado


#Criando layout seleção do ano
layout = [
    [sg.Text("Selecione o ano:",size=(30, 1), font=("Helvetica", 15), text_color='white')],
    [sg.Text("")],
    [sg.Button('2021'),sg.Button('2020'),sg.Button('2019'),sg.Button('2018'),sg.Button('2017'),sg.Button('2016') ],
    [sg.Button('2015'),sg.Button('2014'),sg.Button('2013'),sg.Button('2012'),sg.Button('2011'),sg.Button('2010') ],
    [sg.Button('2009'),sg.Button('2008'),sg.Button('2007'),sg.Button('2006'),sg.Button('2005'),sg.Button('2004') ]
    
    


]
janela = sg.Window('T'), layout
event = sg.Window('Seleção do Ano', layout).read(close=True)
print(event[0])


#Criando layout seleção do mês
layout2 =[ 
    [sg.Text("Selecione o mês:",size=(15, 1), font=("Helvetica", 15), text_color='white')],
    [sg.Text("")],
    [sg.Button('Janeiro'),sg.Button('Fevereiro'),sg.Button('Março'),sg.Button('Abril')],
    [sg.Button('Maio'),sg.Button('Junho'),sg.Button('Julho'),sg.Button('Agosto') ],
    [sg.Button('Setembro'),sg.Button('Outubro'),sg.Button('Novembro'),sg.Button('Dezembro') ]

]
#Janela

janela2 = sg.Window('Tel'), layout2
values = sg.Window('Seleção do Mês', layout2).read(close=True)
print(values[0])
