import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración para GitHub Actions
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 60) # Aumentamos el tiempo de espera a 60s

    try:
        # --- PASO 1: Login ---
        print("Accediendo a la plataforma...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # Esperar a que los campos de texto existan
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        print("Esperando botón de ingreso...")
        # Esperamos a que el botón aparezca en el DOM antes de intentar el click por JS
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "button")))
        
        # Click robusto buscando el texto "INGRESAR"
        driver.execute_script("""
            let buttons = Array.from(document.querySelectorAll('button'));
            let loginBtn = buttons.find(b => b.innerText.includes('INGRESAR') || b.className.includes('mat-flat-button'));
            if (loginBtn) {
                loginBtn.click();
            } else {
                throw new Error('Botón de ingreso no encontrado');
            }
        """)
        
        print("Login enviado, esperando carga del dashboard...")
        time.sleep(15)

        # --- PASO 2: Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        
        # Asegurar columna DG (índice 110)
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111):
                df[f"Col_Temp_{i}"] = ""
        df.columns.values[110] = col_name

        # --- PASO 3: Extracción ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Abriendo NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # Esperar al contenedor que vimos en tu foto
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-list-seguimientos")))
                time.sleep(12) # Pausa para que Angular pinte los datos

                # Extracción forzada de la tabla
                script_ext = """
                let rows = document.querySelectorAll('app-list-seguimientos table tbody tr');
                let result = [];
                rows.forEach(r => {
                    if (r.innerText.trim() && !r.innerText.includes('No hay datos')) {
                        result.push(r.innerText.replace(/\\t/g, ' | ').trim());
                    }
                });
                return result.join('\\n---\\n');
                """
                
                texto = driver.execute_script(script_ext)
                df.at[index, col_name] = texto if texto else "Tabla encontrada pero vacía"
                print(f"-> Datos extraídos para {pqr_nurc}")

            except Exception as e:
                print(f"-> No se pudo extraer datos para {pqr_nurc}")
                df.at[index, col_name] = "Sin seguimiento visible"
                driver.save_screenshot(f"error_{pqr_nurc}.png")
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Proceso finalizado.")

    except Exception as e:
        print(f"Error fatal: {e}")
        driver.save_screenshot("debug_error.png")
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
