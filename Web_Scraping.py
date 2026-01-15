import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
    wait = WebDriverWait(driver, 50)

    try:
        # --- LOGIN ---
        print("Iniciando sesión...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        
        driver.execute_script("""
            let btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR'));
            if(btn) btn.click();
        """)
        
        # VERIFICACIÓN DE LOGIN EXITOSO
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "mat-sidenav-container")))
            print("Login exitoso.")
        except:
            print("Error: El login falló o la página no cargó el dashboard.")
            driver.save_screenshot("error_login.png")
            return

        # --- EXCEL ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111): df[f"C_{i}"] = ""
        df.columns.values[110] = col_name

        # --- EXTRACCIÓN ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': continue
                
            print(f"Abriendo NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            # Forzar zoom out para ver todo el contenido
            driver.execute_script("document.body.style.zoom='70%'")
            
            try:
                # Esperamos a cualquier componente que indique que la página cargó
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "mat-card")))
                time.sleep(10)
                
                # SCRIPT DINÁMICO: Busca la tabla por etiqueta O por cualquier tabla que exista
                script_js = """
                let tabla = document.querySelector('app-list-seguimientos table') || document.querySelector('table');
                if (!tabla) return "";
                
                let logs = [];
                let filas = tabla.querySelectorAll('tbody tr');
                filas.forEach(f => {
                    let c = f.querySelectorAll('td');
                    if (c.length >= 4 && !f.innerText.includes('No hay datos')) {
                        let desc = c[3].querySelector('div') ? c[3].querySelector('div').innerText : c[3].innerText;
                        logs.push(`[${c[0].innerText.trim()}]: ${desc.trim()}`);
                    }
                });
                return logs.join('\\n---\\n');
                """
                
                resultado = ""
                for _ in range(10): # Reintentos internos
                    resultado = driver.execute_script(script_js)
                    if resultado: break
                    time.sleep(2)
                
                df.at[index, col_name] = resultado if resultado else "Sin registros visibles"
                print(f"-> Captura finalizada.")

            except Exception as e:
                print(f"-> No se detectó la tabla para {pqr_nurc}")
                df.at[index, col_name] = "Componente no cargó"
                driver.save_screenshot(f"fallo_{pqr_nurc}.png")
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Fin.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
