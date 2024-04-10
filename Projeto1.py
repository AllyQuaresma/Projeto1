import pyautogui
import pyperclip
import time
import pandas as pd
import openpyxl
import os
import glob


#Passo 1: abrir uma pagina na web
pyautogui.PAUSE = 2
pyautogui.press('win')
pyautogui.write("msedge")
pyautogui.press('enter')
#Passo 2: Navegar no sistema e encontrar a base de dados (pasta do exportar)
time.sleep(3)
pyautogui.hotkey("ctrl","t")
pyperclip.copy('https://fatecspgov-my.sharepoint.com/:f:/r/personal/alexandre_hashimoto_fatec_sp_gov_br/Documents/algoritmo/aulas%20de%20python%20gravadas/aula1/Aula%201?csf=1&web=1&e=AA2tOF ')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)
pyautogui.click(x=204, y=374)
pyautogui.click(x=204, y=374)
pyautogui.click(x=204, y=374)
#Passo 3: Exportar/fazer Download da base de dados
time.sleep(5)
pyautogui.click(x=1351, y=711)
pyautogui.click(x=34, y=173)
time.sleep(3)
pyautogui.click(x=89, y=387)
time.sleep(3)
pyautogui.click(x=356, y=557)
time.sleep(3)

#Passo 4: importar a base de dados para o python
diretorio_downloads = r"C:\Users\aliss\Downloads"
padrao_nome_arquivo = "vendas*.xlsx"
# Lista todos os arquivos que correspondem ao padrão
arquivos = glob.glob(os.path.join(diretorio_downloads, padrao_nome_arquivo))

# Se houver pelo menos um arquivo correspondente
if arquivos:
    # Seleciona o arquivo mais recente
    arquivo_mais_recente = max(arquivos, key=os.path.getctime)
    
    # Importa o arquivo
    tabela = pd.read_excel(arquivo_mais_recente)
else:
    print("Nenhum arquivo correspondente encontrado.")


#Passo 5 Calcular os indicadores
faturamento =tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


#passo 6: enviar um email para a diretoria com o relatorio
#abrir o e-mail
pyautogui.hotkey('ctrl','t')
pyperclip.copy("https://outlook.office.com/mail/?ui=pt-BR&rs=BR")#caminho de email pessoal
pyautogui.hotkey("ctrl","v")
pyautogui.press('enter')
#clica em novo e-mail
time.sleep(5)
pyautogui.click(x=129, y=214)
#digita o destinatario
time.sleep(2)
pyperclip.copy("quaresma15298@gmail.com")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
#escreve o Assunto
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
#escreve o corpo do e-mail
texto = f"""
Prezados, Bom dia!

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos vendidos foi de: {quantidade:,.0f} 

Att.,
Alisson Quaresma."""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
#envia o email
pyautogui.hotkey("ctrl","enter")

