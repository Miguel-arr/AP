import sys
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from email_utils import codsia_outlook
from chrome_utils import login_manglar, Buscar_codigos
from tkinter import messagebox
import config

def main():
    # Configurar el WebDriver
    service = Service(config.webdriver_path)
    driver = webdriver.Chrome(service=service)

    # Leer los códigos del correo electrónico
    codigos = codsia_outlook()

    if codigos:
        messagebox.showinfo("Códigos Encontrados", f'Codigos extraídos: {codigos}')
    else:
        messagebox.showinfo("Sin Códigos", 'No se encontraron códigos en los correos no leídos de Manglar.')
        driver.quit()
        sys.exit()

    # Llamar a la función de login
    login_manglar(driver, config.Correo_m, config.Contra_m)

    # Llamar a la función para buscar y ingresar los códigos
    Buscar_codigos(driver, codigos)

    # Cerrar el navegador
    #driver.quit()

if __name__ == "__main__":
    main()
