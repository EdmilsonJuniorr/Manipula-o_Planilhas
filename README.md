#Ler dados da planilha
#Inserir cada celula de cada linha em um campo
import openpyxl

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row = 2):
    # ['Murilo barros' , 'cadeira', '454' , 'Esportes']
    pyautogui.click(1111,111, duration = 1.0)
    pyautogui.write(linha[0].value)
    pyautogui.click(1111,111, duration = 1.0)
    pyautogui.write(linha[1].value)
    pyautogui.click(1111,111, duration = 1.0)
    pyautogui.write(linha[2].value)
    pyautogui.click()# dentro a coordenada do salvar
    pyautogui.click()# dentro a coordenada do ok

