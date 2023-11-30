#!/bin/python3
#this is a simple GUI to represent some program that you have in your company or something like that
import PySimpleGUI as sg

sg.theme('Reddit')

product_categories = ["Eletrônicos", "Móveis", "Roupas", "Brinquedos", "Comida", "Bebidas", 
                      "Cosméticos", "Livros", "Esportes", "Jardinagem"]

layout = [
    [sg.Text('Cliente',size=(6,0)),sg.Input(key='1',size=(20,0))],
    [sg.Text('Produto',size=(6,0)),sg.Input(size=(20,0),key='2')],
    [sg.Text('Quantidade'),sg.Input(key='3',size=(3,0))],
    [sg.Text('Categoria do Produto'), sg.Combo(product_categories, key='4')],
    [sg.Button('Salvar')]

]

window = sg.Window('Cadastro de Produtos',layout)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    elif event == 'Salvar':
        sg.popup('Produto Cadastrado')
        window['1'].update('')
        window['2'].update('')
        window['3'].update('')
        window['4'].update('')