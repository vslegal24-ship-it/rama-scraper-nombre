"""
PRUEBA — Solo 5 registros para validar que el scraper funciona
"""
import json, os, re, time, shutil, warnings
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

warnings.filterwarnings("ignore")

EXCEL_INPUT = "piloto_numero_de_radicacion.xlsx"
URL = "https://consultaprocesos.ramajudicial.gov.co/Procesos/NombreRazonSocial"

# ── Leer solo los primeros 5 ──────────────────────────────────
df = pd.read_excel(EXCEL_INPUT, dtype=str).fillna("-")
df.columns = df.columns.str.strip()
muestra = df.head(5)

print("=" * 60)
print("REGISTROS A BUSCAR:")
for i, row in muestra.iterrows():
    print(f"  {i+1}. {row['Nombre Cliente']} | {row['Departamento']} | Juzgado {row['Numero Del Juzgado']} {row.get('Tipo De \nJuzgado', row.get('Tipo De Juzgado',''))}")
print("=" * 60)

# ── Chrome ────────────────────────────────────────────────────
options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1400,900")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

chrome_bin = shutil.which("chromium") or shutil.which("chromium-browser") or "/usr/bin/chromium"
chromedriver_bin = shutil.which("chromedriver") or "/usr/bin/chromedriver"
options.binary_location = chrome_bin
service = Service(chromedriver_bin)

print(f"\nChrome en: {chrome_bin}")
print(f"Chromedriver en: {chromedriver_bin}")

driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 25)

resultados = []

for i, row in muestra.iterrows():
    nombre = str(row["Nombre Cliente"]).strip()
    print(f"\n[{i+1}/5] Buscando: {nombre}")

    try:
        driver.get(URL)
        time.sleep(3)

        # Escribir nombre
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text'][1]")))
        campo.clear()
        campo.send_keys(nombre)
        time.sleep(0.5)

        # Click consultar
        boton = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(@class,'success') or contains(text(),'Consultar')]")
        ))
        driver.execute_script("arguments[0].click();", boton)
        time.sleep(5)

        # Leer tabla
        filas = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
        procesos_encontrados = []
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) >= 2:
                radicado = celdas[0].text.strip()
                despacho = celdas[1].text.strip()
                if re.match(r'^\d{15,}', radicado):
                    procesos_encontrados.append(f"{radicado} | {despacho}")

        if procesos_encontrados:
            print(f"  ✓ {len(procesos_encontrados)} proceso(s) encontrado(s):")
            for p in procesos_encontrados:
                print(f"    → {p}")
            resultados.append({"nombre": nombre, "estado": "OK", "procesos": procesos_encontrados})
        else:
            # Verificar si hay mensaje de error o captcha
            page_text = driver.find_element(By.TAG_NAME, "body").text[:300]
            print(f"  ✗ Sin resultados")
            print(f"  Texto de página: {page_text[:200]}")
            driver.save_screenshot(f"debug_{i+1}_{nombre[:15].replace(' ','_')}.png")
            resultados.append({"nombre": nombre, "estado": "SIN RESULTADOS", "procesos": []})

    except Exception as e:
        print(f"  ✗ ERROR: {str(e)[:150]}")
        driver.save_screenshot(f"error_{i+1}.png")
        resultados.append({"nombre": nombre, "estado": f"ERROR: {str(e)[:80]}", "procesos": []})

    time.sleep(2)

driver.quit()

# ── Resumen final ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("RESUMEN DE PRUEBA")
print("=" * 60)
ok = sum(1 for r in resultados if r["estado"] == "OK")
print(f"✓ Encontrados   : {ok}/5")
print(f"✗ Sin resultado : {5 - ok}/5")
print()
for r in resultados:
    estado = "✓" if r["estado"] == "OK" else "✗"
    print(f"  {estado} {r['nombre'][:40]}")
    for p in r["procesos"]:
        print(f"      {p}")

print("\n✓ Prueba completada")
