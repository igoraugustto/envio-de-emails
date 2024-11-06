from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pyautogui
import time
from selenium.common.exceptions import NoSuchElementException
import subprocess
import pygetwindow as gw
import pandas as pd
import os
#Header
service = Service()
options = webdriver.ChromeOptions()
navegador = webdriver.Chrome(service= service, options= options)
url = f'https://casadosdados.com.br/solucao/cnpj/pesquisa-avancada?id=NJK5V8DIbcIX_tdt7-GFDdfVEkjiYEymfyEx5OvadCI='
navegador.get(url)
time.sleep(5)


navegador.execute_script ('document.querySelector("#__nuxt > div > div.top-footer > section > div:nth-child(7) > div > div > a.button.is-success.is-medium").click();')
time.sleep(5)
elements = navegador.find_elements(By.CSS_SELECTOR, ".fas.fa-external-link-alt")


# Itera sobre os elementos e clica em cada um
for element in elements:
    navegador.execute_script("arguments[0].scrollIntoView(true);", element)
    time.sleep(2)
    element.click()

    nova_guia = navegador.window_handles[-1]

    time.sleep(2)
    navegador.switch_to.window(nova_guia)
    time.sleep(2)
    navegador.close()
    time.sleep(2)
    navegador.switch_to.frame("google_ads_iframe_/21715141650,21807930814/desktop_under_0")
    navegador.switch_to.window(navegador.window_handles[0])
