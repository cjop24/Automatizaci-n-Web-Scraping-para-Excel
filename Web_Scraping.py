import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración de Chrome Headless
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

def run_scraper():
    # Recuperar credenciales desde Secrets de GitHub
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- PASO 1: Login ---
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        
        # Click en el botón Ingresar
        btn_selector = "body > app-root > mat-drawer-container > mat-drawer-content > app-login > div > div:nth-child(2) > mat-card > mat-card-actions > div > button"
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, btn_selector))).click()
        
        # Espera de seguridad para carga de dashboard
        time.sleep(5)

        # --- PASO 2: Lectura de Excel (.xlsx) ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl')

        # Iterar registros (Columna F es índice 5)
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip()
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            driver.get(url_reclamo)
            
            try:
                # --- PASO 2.3: Extraer Seguimiento ---
                seguimiento_elem = wait.until(EC.visibility_of_element_located((By.ID, "main_table_wrapper")))
                
                # --- PASO 3: Guardar en Columna DG (índice 110) ---
                df.iat[index, 110] = seguimiento_elem.text
                print(f"Procesado NURC: {pqr_nurc}")
                
            except Exception as e:
                print(f"No se pudo extraer información para {pqr_nurc}: {e}")
            
            # Respetar el servidor con una pausa corta
            time.sleep(2)

        # Guardar archivo final
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Proceso completado. Archivo Reclamos_scraping.xlsx generado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
