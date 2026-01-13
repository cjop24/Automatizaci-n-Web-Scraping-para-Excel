import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Configuración de Chrome optimizada para GitHub Actions
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
# User-Agent para evitar ser detectado como bot básico
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    # Cargar credenciales desde Secrets
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 30) # Aumentamos el tiempo de espera a 30s

    try:
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # PASO 1: Login
        user_input = wait.until(EC.presence_of_element_located((By.ID, "user")))
        user_input.send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        
        # Intento de click robusto en el botón INGRESAR
        try:
            # Buscamos el botón por su texto interno para mayor seguridad
            btn_ingresar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'INGRESAR')]]")))
            driver.execute_script("arguments[0].click();", btn_ingresar)
            print("Click en INGRESAR exitoso.")
        except Exception as e:
            print(f"Error al hacer click en el botón: {e}")
            driver.save_screenshot("error_login.png") # Captura de pantalla para depurar
            raise

        # Esperar a que cargue la página interna (validamos que el login fue exitoso)
        time.sleep(7)

        # PASO 2: Lectura de Excel
        file_path = "Reclamos.xlsx"
        if not os.path.exists(file_path):
            print(f"Error: El archivo {file_path} no se encuentra en la raíz.")
            return

        df = pd.read_excel(file_path, engine='openpyxl')

        # Procesamiento por filas
        for index, row in df.iterrows():
            # Columna F es índice 5 (pqr_nurc)
            pqr_nurc = str(row.iloc[5]).strip()
            
            # Detener si encuentra celda vacía
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                print(f"Fin de datos en la fila {index + 2}.")
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Procesando: {url_reclamo}")
            
            driver.get(url_reclamo)
            
            try:
                # PASO 2.3: Extraer Seguimiento (Selector ID solicitado)
                seguimiento_elem = wait.until(EC.visibility_of_element_located((By.ID, "main_table_wrapper")))
                
                # PASO 3: Guardar en columna DG (índice 110)
                df.iat[index, 110] = seguimiento_elem.text
                
            except TimeoutException:
                print(f"Tiempo excedido esperando seguimiento de NURC: {pqr_nurc}")
                df.iat[index, 110] = "ERROR: No cargó la tabla"
            
            # Delay para evitar sobrecarga (2 segundos)
            time.sleep(2)

        # PASO 3.2: Guardar archivo final
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Archivo Reclamos_scraping.xlsx generado con éxito.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
