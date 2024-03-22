import pyautogui
import subprocess
import time
import pandas as pd

url = "https://forms.gle/wehgxzQoGoWAeAT29"

subprocess.run(r"C:\Program Files\Google\Chrome\Application/chrome")

time.sleep(1)
pyautogui.write(url)
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)

# Pressioando TAB para ir para o INPUT
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
time.sleep(1)
    
tabela = pd.read_excel('Lista_de_Contatos.xlsx')
for linha in tabela.index:
    nome = tabela.loc[linha, 'Nome']
    empresa = tabela.loc[linha, 'Empresa']
    profissao = tabela.loc[linha, 'Profissao']

    pyautogui.write(nome)
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.write(empresa)
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.write(profissao)
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('tab')
    pyautogui.press('enter')
    







