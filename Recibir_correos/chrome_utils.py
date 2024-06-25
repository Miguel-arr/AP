import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

def login_manglar(driver, correo, contrasena):
    # Abrir la página de login de Manglar
    driver.get("https://gis.manglar.com/login#")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "inputEmail")))

    # Ingresar el correo electrónico
    email_input = driver.find_element(By.NAME, "inputEmail")
    email_input.send_keys(correo)

    # Ingresar la contraseña
    password_input = driver.find_element(By.NAME, "inputPassword")
    password_input.send_keys(contrasena)

    # Hacer clic en el botón de login
    login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
    login_button.click()

    try:
        time.sleep(10)
        # Hacer clic en el boton maps
        maps_button = driver.find_element(By.LINK_TEXT, "Maps")
        maps_button.click()
    except Exception:
        print("No se pudo encontrar el elemento 'a.nav-link.active[href='/maps']' dentro del tiempo de espera.")

def Buscar_codigos(driver, codigos):
    # Esperar a que el campo de búsqueda esté presente y visible
    buscar_campo = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input.form-control.font-md.border-0"))
    )

    for codigo in codigos:
        # Limpiar el campo de búsqueda usando JavaScript
        driver.execute_script("arguments[0].value = '';", buscar_campo)
        time.sleep(1)  # Esperar un poco para asegurarse de que el campo esté limpio

        if buscar_campo.get_attribute('value') == '':
            for caracter in codigo:
                buscar_campo.send_keys(caracter)
                time.sleep(0.5)
            time.sleep(3)
            buscar_campo.send_keys(Keys.ENTER)
            time.sleep(10)

             # Buscar y seleccionar el "Replant (Plan)" con la fecha más reciente
             # Seleccionar el campo de "Replant (Plan)"
            try:
                replant_checkbox = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, ".//span[contains(text(), 'Replant (Plan)')]//ancestor::div[@class='d-flex'][1]//div[@class='d-flex ml-auto mr-1']/button[@class='btn btn-sm text-secondary']//i[@class='fas fa-chevron-right']"))
                )
                if not replant_checkbox.is_selected():
                    replant_checkbox.click()
                print("Seleccionado el campo 'Replant (Plan)' para el código: " + codigo)
            except Exception as e:
                print(f"No se pudo seleccionar el campo 'Replant (Plan)' para el código {codigo}: {e}")
                
            # Hacer clic en el botón para abrir otra ventana
            try:
                boton_abrir_ventana = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "i.text-info.fas.fa-chart-pie"))
                )
                if not boton_abrir_ventana.is_selected():
                    boton_abrir_ventana.click()
                print("Botón para abrir otra ventana clickeado.")
            except Exception as e:
                print(f"No se pudo hacer clic en el botón para abrir otra ventana: {e}")
            buscar_campo.send_keys(Keys.DELETE)
            time.sleep(5)
            print("Se buscó el código: " + codigo)
        else:
            print("El campo de búsqueda no se limpió correctamente.")
