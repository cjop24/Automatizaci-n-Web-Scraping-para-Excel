import os
import pandas as pd
import time
from selenium import webdriver
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
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    # Cargar credenciales desde Secrets de GitHub
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 30)

    try:
        # --- PASO 1: Login ---
        print("Abriendo página de login...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # Ingresar credenciales
        print("Ingresando credenciales...")
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        # CLICK MEDIANTE JAVASCRIPT (Solución para el error de Timeout)
        print("Intentando click en INGRESAR vía JS...")
        time.sleep(2) # Pausa técnica para renderizado
        
        boton_js = """
        var botones = document.querySelectorAll('button');
        for (var i = 0; i < botones.length; i++) {
            if (botones[i].textContent.includes('INGRESAR')) {
                botones[i].click();
                return true;
            }
        }
        return false;
        """
        success = driver.execute_script(boton_js)
        
        if not success:
            print("No se encontró el botón vía JS, intentando click tradicional...")
            driver.find_element(By.CSS_SELECTOR, "button").click()

        # Esperar a que el sistema procese el ingreso
        print("Esperando carga del dashboard...")
        time.sleep(10)

        # --- PASO 2: Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        if not os.path.exists(file_path):
            print(f"ERROR: No se encontró el archivo {file_path}")
            return

        df = pd.read_excel(file_path, engine='openpyxl')

        # Iterar sobre las filas (Columna F = índice 5)
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip()
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                print(f"Llegamos al final de los datos en la fila {index + 1}")
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Extrayendo NURC {pqr_nurc}...")
            
            driver.get(url_reclamo)
            
            try:
                # PASO 2.3: Extraer Seguimiento (ID: main_table_wrapper)
                seguimiento_elem = wait.until(EC.visibility_of_element_located((By.ID, "main_table_wrapper")))
                
                # PASO 3: Guardar en columna DG (índice 110)
                df.iat[index, 110] = seguimiento_elem.text
                
            except TimeoutException:
                print(f"Aviso: Timeout en NURC {pqr_nurc}. Saltando...")
                df.iat[index, 110] = "No se pudo extraer información"
            
            # Respetar tiempo del servidor
            time.sleep(3)

        # Guardar archivo final
        output_file = "Reclamos_scraping.xlsx"
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Proceso finalizado. Archivo {output_file} creado.")

    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        driver.save_screenshot("debug_error.png") # Captura para ver qué falló
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
