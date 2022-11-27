import PySimpleGUI as sg
import pandas as pd

#Tema do GUI
sg.theme('DarkTeal9')

#Puxando banco de dados pelo pandas
Exel_file = 'Banco_de_dados.xlsx'
df = pd.read_excel(Exel_file)

#Layout pelo GUI
layout = [
    [sg.Text('Por favor preencha as cedulas abaixo')],
    [sg.Text('Nome', size=(15, 1)), sg.InputText(key='Nome')],
    [sg.Text('Idade', size=(15, 1)), sg.InputText(key='Idade')],
    [sg.Text('Gênero(M/F)', size=(15, 1)), sg.InputText(key='Genero')],
    [
        sg.Text('Qual a sua cor favorita?', size=(20, 1)),
        sg.Combo(['Verde', 'Amarelo', 'Azul', 'Vermelho'], key='CorFavorita'),
    ],
    [
         sg.Text('Escolha sua linguagem', size=(20, 1)),
    ],
        [
            sg.Checkbox('Inglês', key='Ingles'),
            sg.Checkbox('Portugues', key='Portugues'),
            sg.Checkbox('Espanhol', key='Espanhol'),
        ],
    
    [sg.Submit('Enviar'), sg.Button('Limpar'), sg.Exit('Sair')],
]

#Iniciar
Janela = sg.Window('Entrada de dados', layout)

def clear_input():
    for key in valor:
        Janela[key]("")
    return None

while True:
    evento, valor = Janela.read()
    if evento == 'Limpar':
        clear_input()
        sg.popup('Formulario limpo!')
    #pop up de confirmação de envio
    if evento == 'Enviar':
        df = df.append(valor, ignore_index=True)
        df.to_excel(Exel_file, index=False)
        sg.popup('Arquivo Salvo!')
        clear_input()
    if evento == 'Sair':
        Janela.close()
    Janela.close()