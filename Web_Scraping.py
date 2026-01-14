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
    # Recordatorio: El usuario y password se manejan vía Secrets como solicitaste
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 45)

    try:
        # --- LOGIN ---
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("document.querySelectorAll('button').forEach(b => { if(b.innerText.includes('INGRESAR')) b.click(); });")
        time.sleep(12)

        # --- PREPARACIÓN DE EXCEL ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        
        # Aseguramos que la columna DG (índice 110) exista
        col_name = "Seguimiento_Extraido"
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111):
                df[f"Col_Aux_{i}"] = ""
        df.columns.values[110] = col_name

        # --- BUCLE DE EXTRACCIÓN ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Extrayendo seguimiento del NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # 1. Esperamos al componente que viste en tu inspección (app-list-seguimientos)
                # Basado en tu foto, este es el contenedor padre real.
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-list-seguimientos")))
                
                # 2. IMPORTANTE: Scroll hasta el elemento para forzar a Angular a cargar los datos
                # A veces, si el elemento no está en el área de visión, la tabla no se puebla.
                elemento_seguimiento = driver.find_element(By.TAG_NAME, "app-list-seguimientos")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_seguimiento)
                
                # 3. Pausa generosa para que el componente "dibuje" las filas
                time.sleep(12)
                
                # 4. SCRIPT DE EXTRACCIÓN (Directo a la tabla señalada en azul)
                # Extraemos Fecha, Usuario, Comentario y Descripción.
                script_js = """
                let tabla = document.querySelector('app-list-seguimientos table');
                if (!tabla) return "TABLA_NO_DETECTADA";
                
                let filas = tabla.querySelectorAll('tbody tr');
                let logs = [];
                filas.forEach(f => {
                    let c = f.querySelectorAll('td');
                    if(c.length >= 4 && !f.innerText.includes('No hay datos')) {
                        let fch = c[0].innerText.trim();
                        let usr = c[1].innerText.trim();
                        let com = c[2].innerText.trim();
                        let dsc = c[3].innerText.trim();
                        logs.push(`[${fch}] ${usr} - ${com}: ${dsc}`);
                    }
                });
                return logs.length > 0 ? logs.join('\\n---\\n') : "Sin registros en la tabla";
                """
                
                resultado = driver.execute_script(script_js)
                df.at[index, col_name] = resultado
                print(f"-> EXITO: Datos capturados para {pqr_nurc}")

            except Exception as e:
                print(f"-> AVISO: No se cargó la tabla para {pqr_nurc}")
                driver.save_screenshot(f"fallo_{pqr_nurc}.png")
                df.at[index, col_name] = "Error de visualización o tabla no encontrada"
            
            time.sleep(2)

        # Guardar resultados
        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Archivo Reclamos_scraping.xlsx generado con los seguimientos.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
