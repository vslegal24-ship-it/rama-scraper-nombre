"""
DIAGNÓSTICO — Rama Judicial NombreRazonSocial
Ejecutar PRIMERO para verificar que los selectores funcionen correctamente.
Busca solo el primer registro y muestra todo lo que encuentra.
"""
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, json, re

URL = "https://consultaprocesos.ramajudicial.gov.co/Procesos/NombreRazonSocial"

# ── Chrome ────────────────────────────────────────────────────
options = Options()
# Cambiar a False para ver el navegador mientras depura:
HEADLESS = True
if HEADLESS:
    options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1400,900")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

# Ajustar ruta según el sistema
# Linux (servidor):
# options.binary_location = "/usr/bin/chromium-browser"
# service = Service("/usr/bin/chromedriver")

# Windows:
# service = Service("C:/chromedriver/chromedriver.exe")

# Mac:
# service = Service("/usr/local/bin/chromedriver")

service = Service()  # Usa PATH automáticamente
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 25)

print("Abriendo página...")
driver.get(URL)
time.sleep(3)

print(f"Título: {driver.title}")
driver.save_screenshot("diagnostico_01_inicio.png")
print("Screenshot: diagnostico_01_inicio.png")

# Listar todos los inputs
print("\n=== INPUTS ===")
inputs = driver.find_elements(By.TAG_NAME, "input")
for i, inp in enumerate(inputs):
    print(f"  [{i}] type={inp.get_attribute('type')} "
          f"placeholder='{inp.get_attribute('placeholder')}' "
          f"id='{inp.get_attribute('id')}' "
          f"class='{(inp.get_attribute('class') or '')[:60]}'")

# Listar botones
print("\n=== BOTONES ===")
for b in driver.find_elements(By.TAG_NAME, "button"):
    print(f"  '{b.text[:50]}' class='{(b.get_attribute('class') or '')[:60]}'")

# Listar v-select
print("\n=== V-SELECTS ===")
for i, v in enumerate(driver.find_elements(By.CSS_SELECTOR, ".v-select, .v-autocomplete")):
    print(f"  [{i}] '{v.text[:60]}'")

# ── Intentar búsqueda de prueba ───────────────────────────────
NOMBRE_PRUEBA = "MIGUEL LEMUS DE AVILA"
print(f"\nBuscando: {NOMBRE_PRUEBA}")

try:
    # Primer input de texto = campo nombre
    campo = inputs[0] if inputs else None
    if campo:
        campo.clear()
        campo.send_keys(NOMBRE_PRUEBA)
        time.sleep(0.5)
        print("✓ Nombre escrito")
    
    # Botón consultar
    boton = None
    for b in driver.find_elements(By.TAG_NAME, "button"):
        cls = b.get_attribute("class") or ""
        txt = b.text.upper()
        if "success" in cls or "CONSULTAR" in txt or "BUSCAR" in txt:
            boton = b
            break
    
    if boton:
        driver.execute_script("arguments[0].click();", boton)
        print("✓ Click en consultar")
        time.sleep(5)
        driver.save_screenshot("diagnostico_02_resultados.png")
        print("Screenshot: diagnostico_02_resultados.png")
    else:
        print("✗ No se encontró botón de búsqueda")

    # Leer tabla
    print("\n=== TABLA DE RESULTADOS ===")
    filas = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
    print(f"Filas encontradas: {len(filas)}")
    for i, fila in enumerate(filas[:5]):
        celdas = fila.find_elements(By.TAG_NAME, "td")
        print(f"  [{i}] ({len(celdas)} celdas) | " + " | ".join(c.text[:30] for c in celdas[:5]))

    # HTML completo de la tabla para análisis
    try:
        tabla = driver.find_element(By.CSS_SELECTOR, ".v-data-table")
        with open("diagnostico_tabla.html", "w", encoding="utf-8") as f:
            f.write(tabla.get_attribute("outerHTML"))
        print("HTML de tabla guardado en diagnostico_tabla.html")
    except:
        pass

except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

driver.quit()
print("\n✓ Diagnóstico completado")
