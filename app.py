#!/bin/python3
import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('tables.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    # name
    pyautogui.click(766,430,duration=1)
    pyautogui.write(linha[0].value)
    # product
    pyautogui.click(766,458,duration=1)
    pyautogui.write(linha[1].value)
    # amount
    pyautogui.click(773,482,duration=1)
    pyautogui.write(str(linha[2].value))
    # category
    pyautogui.click(855,506,duration=1)
    pyautogui.write(linha[3].value)
    # save
    pyautogui.click(716,536,duration=1)
    pyautogui.click(790,500,duration=1)