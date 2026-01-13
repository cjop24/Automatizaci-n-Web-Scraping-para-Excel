import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración de Chrome optimizada para GitHub Actions
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 45)

    try:
        # --- PASO 1: Login ---
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # Esperar a que los campos de login estén presentes
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        # Esperar a que el botón de ingresar sea visible antes del click JS
        # Usamos el selector de clase que vimos en la foto 1
        btn_selector = "button.mat-flat-button"
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, btn_selector)))
        
        print("Enviando formulario...")
        driver.execute_script(f"document.querySelector('{btn_selector}').click();")
        time.sleep(12) # Tiempo de espera para carga del dashboard

        # --- PASO 2: Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        # Mantenemos dtype=str para evitar redondeos
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        # Asegurar columna DG (índice 110)
        while df.shape[1] <= 110:
            df[f"Columna_Seguimiento_{df.shape[1]}"] = ""
        
        target_col = 110

        # --- PASO 3: Bucle de Extracción ---
        for index, row in df.iterrows():
            # Limpieza de NURC para asegurar que sea el ID correcto
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Abriendo NURC: {pqr_nurc}")
            driver.get(url_reclamo)
            
            try:
                # SELECTOR PROFUNDO PROPORCIONADO POR EL USUARIO
                selector_profundo = "#contenido > div > div:nth-child(1) > div > mat-card:nth-child(3) > app-follow > mat-card-content"
                
                # Esperamos a que el selector esté en el DOM y sea visible
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selector_profundo)))
                
                # Pausa extra para renderizado dinámico de contenido (Angular)
                time.sleep(10)
                
                # Extracción forzada mediante JavaScript para capturar texto oculto/dinámico
                script_extrayente = f"return document.querySelector('{selector_profundo}').innerText;"
                texto_extraido = driver.execute_script(script_extrayente)

                if texto_extraido:
                    df.iat[index, target_col] = texto_extraido.strip()
                    print(f"-> EXITO: {pqr_nurc} ({len(texto_extraido)} caracteres)")
                else:
                    print(f"-> AVISO: {pqr_nurc} el selector existe pero no tiene texto interno.")
                    df.iat[index, target_col] = "Contenedor encontrado vacío"

            except Exception as e:
                print(f"-> ERROR en NURC {pqr_nurc}: El selector no apareció a tiempo.")
                driver.save_screenshot(f"debug_{pqr_nurc}.png")
                df.iat[index, target_col] = "Error: Selector no localizado"
            
            time.sleep(3)

        # Guardar archivo final
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Proceso finalizado. Archivo Reclamos_scraping.xlsx generado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
