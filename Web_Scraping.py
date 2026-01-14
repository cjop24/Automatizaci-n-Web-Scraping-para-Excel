import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración de Chrome optimizada
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
    wait = WebDriverWait(driver, 30)

    try:
        # --- LOGIN ---
        print("Iniciando sesión...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("document.querySelectorAll('button').forEach(b => { if(b.innerText.includes('INGRESAR')) b.click(); });")
        time.sleep(10)

        # --- EXCEL ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        target_col = 110 # Columna DG

        # --- PRUEBA DE SELECTOR ESPECÍFICO ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Probando selector en NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # Selector proporcionado por el usuario (Información en app-status)
                selector_test = "#contenido > div > div:nth-child(1) > div > mat-card:nth-child(1) > app-status > div > div:nth-child(3) > div:nth-child(8)"
                
                # Esperar a que el elemento esté presente
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector_test)))
                time.sleep(5) # Pausa breve para renderizado
                
                # Extracción vía JavaScript para máxima compatibilidad
                dato_prueba = driver.execute_script(f"return document.querySelector('{selector_test}').innerText;")
                
                if dato_prueba:
                    df.iat[index, target_col] = dato_prueba.strip()
                    print(f"-> EXITO DE PRUEBA: Dato encontrado: {dato_prueba.strip()}")
                else:
                    df.iat[index, target_col] = "Elemento encontrado pero sin texto"
                    print("-> AVISO: Elemento sin texto.")

            except Exception as e:
                print(f"-> ERROR: No se pudo localizar el selector de prueba.")
                driver.save_screenshot(f"test_error_{pqr_nurc}.png")
                df.iat[index, target_col] = "Selector de prueba no encontrado"
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Prueba finalizada.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
