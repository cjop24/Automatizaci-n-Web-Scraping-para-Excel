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
    wait = WebDriverWait(driver, 45)

    try:
        # --- PASO 1: Login Robusto ---
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        # Intentar click en el botón de ingresar
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'INGRESAR')]"))).click()
        except:
            driver.execute_script("document.querySelectorAll('button').forEach(b => { if(b.innerText.includes('INGRESAR')) b.click(); });")
        
        time.sleep(10)

        # --- PASO 2: Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        # Columna DG (índice 110)
        while df.shape[1] <= 110:
            df[f"Seguimiento_Extraido_{df.shape[1]}"] = ""
        
        target_col = 110

        # --- PASO 3: Bucle de Extracción ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Procesando NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # ESPERA POR LA TABLA ESPECÍFICA (identificada en tu HTML)
                wait.until(EC.presence_of_element_located((By.ID, "main_table")))
                time.sleep(8) # Tiempo para que Angular cargue las filas (ng-star-inserted)

                # SCRIPT DE EXTRACCIÓN PERSONALIZADO PARA TU ESTRUCTURA
                # Extraemos la fecha, el comentario y la descripción de cada fila
                script_extraccion = """
                let filas = document.querySelectorAll('#main_table tbody tr.ng-star-inserted');
                let resultado = "";
                filas.forEach((fila) => {
                    let cols = fila.querySelectorAll('td');
                    if(cols.length >= 4) {
                        let fecha = cols[0].innerText.trim();
                        let comentario = cols[2].innerText.trim();
                        let descripcion = cols[3].innerText.trim();
                        resultado += `[${fecha}] ${comentario}: ${descripcion}\\n---\\n`;
                    }
                });
                return resultado;
                """
                
                texto_final = driver.execute_script(script_extraccion)

                if texto_final and len(texto_final) > 10:
                    df.iat[index, target_col] = texto_final.strip()
                    print(f"-> EXITO: {len(texto_final)} caracteres capturados.")
                else:
                    # Si no hay filas, puede que diga "No hay datos"
                    df.iat[index, target_col] = "Sin registros de seguimiento en la tabla."
                    print("-> AVISO: Tabla vacía.")

            except Exception as e:
                print(f"-> ERROR en {pqr_nurc}: No se encontró la tabla de seguimiento.")
                driver.save_screenshot(f"debug_{pqr_nurc}.png")
                df.iat[index, target_col] = "Error: No se localizó la tabla de seguimiento."
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Proceso finalizado con éxito.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
