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
    wait = WebDriverWait(driver, 40)

    try:
        # --- LOGIN ---
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("document.querySelectorAll('button').forEach(b => { if(b.innerText.includes('INGRESAR')) b.click(); });")
        time.sleep(12)

        # --- EXCEL (Corrección de Límites) ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        
        # SOLUCIÓN AL INDEX ERROR: Creamos la columna DG explícitamente si falta
        col_name = "Seguimiento_Extraido"
        if len(df.columns) <= 110:
            # Rellenamos con columnas vacías hasta llegar a la 111 (índice 110)
            for i in range(len(df.columns), 111):
                df[f"Col_Aux_{i}"] = ""
        
        # Usamos el nombre de la columna en lugar del índice numérico para evitar errores
        df.columns.values[110] = col_name

        # --- EXTRACCIÓN ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Navegando a NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # En lugar de un selector largo, esperamos a que el contenido principal cargue
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "mat-card")))
                time.sleep(10)
                
                # Intentamos extraer TODO el texto de los seguimientos buscando la tabla por ID
                # Esto es mucho más seguro que nth-child
                script_js = """
                let tabla = document.getElementById('main_table');
                if (tabla) return tabla.innerText;
                let follow = document.querySelector('app-follow');
                if (follow) return follow.innerText;
                return "No se visualiza contenido de seguimiento";
                """
                
                resultado = driver.execute_script(script_js)
                df.at[index, col_name] = resultado.strip() if resultado else "Vacio"
                print(f"-> EXITO: Datos capturados para {pqr_nurc}")

            except Exception as e:
                print(f"-> ERROR: No se detectó el contenido en el tiempo previsto.")
                driver.save_screenshot(f"error_nurc_{pqr_nurc}.png")
                df.at[index, col_name] = "Error de carga en la página"
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Proceso terminado exitosamente.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
