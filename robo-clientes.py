from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pyautogui
import time
from selenium.common.exceptions import NoSuchElementException
import subprocess
import pygetwindow as gw

#Header
service = Service()
options = webdriver.ChromeOptions()
navegador = webdriver.Chrome(service= service, options= options)
subprocess.Popen(['notepad.exe'])
url = f'https://casadosdados.com.br/solucao/cnpj/lamarck-administracao-de-condominios-ltda-32235405000103'
navegador.get(url)

def switch_to_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        window.activate()  # Ativar a janela
    except IndexError:
        print(f"Janela '{window_title}' não encontrada.")
notepad_title = "Sem título - Bloco de Notas" 
browser_title = "recursos humanos( “@outlook.com”) and brazil site:instagram.com/ - Pesquisa Google - Google Chrome"  
switch_to_window(browser_title)
while True:
    try:        
        navegador.execute_script("document.getElementById('pnnext').click();")
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(2)
        switch_to_window(notepad_title)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(1)
        switch_to_window(browser_title)
        time.sleep(4)

    except NoSuchElementException:
        print("Botão 'Próxima' não encontrado. Fim da navegação.")
        break
