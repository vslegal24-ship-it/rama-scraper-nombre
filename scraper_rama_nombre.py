"""
VS Legal — Rama Judicial Scraper por Nombre/Razón Social
- Lee base de datos Excel con clientes
- Busca cada uno en https://consultaprocesos.ramajudicial.gov.co/Procesos/NombreRazonSocial
- Cruza resultados con Juzgado + Tipo para identificar el proceso correcto
- Exporta Excel con radicados encontrados
"""
import json, os, re, time, sys, warnings
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

warnings.filterwarnings("ignore")

# ── Configuración ─────────────────────────────────────────────
EXCEL_INPUT  = "piloto_numero_de_radicacion.xlsx"
EXCEL_OUTPUT = "resultado_radicados.xlsx"
RESULTADOS_DIR = "resultados_nombre"
URL = "https://consultaprocesos.ramajudicial.gov.co/Procesos/NombreRazonSocial"

# Mapa de departamentos para el selector de la Rama
# La Rama usa nombres con tildes y formato específico
DEPTOS_MAP = {
    "Cundinamarca":   "CUNDINAMARCA",
    "Santander":      "SANTANDER",
    "Tolima":         "TOLIMA",
    "Antioquia":      "ANTIOQUIA",
    "Cordoba":        "CÓRDOBA",
    "Bolivar":        "BOLÍVAR",
    "Huila":          "HUILA",
    "Valle Del Cauca":"VALLE DEL CAUCA",
    "Atlantico":      "ATLÁNTICO",
    "Magdalena":      "MAGDALENA",
    "Caldas":         "CALDAS",
    "Quindio":        "QUINDÍO",
    "Caqueta":        "CAQUETÁ",
    "Cesar":          "CESAR",
    "Boyaca":         "BOYACÁ",
    "Sucre":          "SUCRE",
    "Guajira":        "LA GUAJIRA",
    "Risaralda":      "RISARALDA",
    "Narino":         "NARIÑO",
    "Putumayo":       "PUTUMAYO",
    "Meta":           "META",
    "Norte De Santander": "NORTE DE SANTANDER",
    "Cauca":          "CAUCA",
    "Choco":          "CHOCÓ",
}

def normalizar(texto):
    """Normaliza texto para comparación"""
    if not texto or texto == "-":
        return ""
    return str(texto).upper().strip()

def coincide_juzgado(resultado_texto, numero_juzgado, tipo_juzgado):
    """Verifica si un resultado de la Rama corresponde al juzgado de la base"""
    r = normalizar(resultado_texto)
    num = normalizar(numero_juzgado)
    tipo = normalizar(tipo_juzgado)

    if not r:
        return False

    # Checar número de juzgado
    num_ok = False
    if num and num != "-":
        # Formatear como 3 dígitos con ceros
        try:
            n = int(num)
            num_ok = (f"{n:03d}" in r) or (str(n) in r)
        except:
            num_ok = num in r

    # Checar tipo de juzgado
    tipo_ok = False
    if tipo and tipo != "-":
        tipo_keywords = {
            "CIVIL MUNICIPAL":   ["CIVIL MUNICIPAL"],
            "PEQUENAS CAUSAS":   ["PEQUEÑAS CAUSAS", "PEQUEÑAS CAUSAS CIVILES"],
            "PROMISCUO":         ["PROMISCUO"],
            "CIVIL CIRCUITO":    ["CIVIL DEL CIRCUITO"],
            "FAMILIA":           ["FAMILIA"],
            "LABORAL":           ["LABORAL"],
        }
        for k, variants in tipo_keywords.items():
            if k in tipo:
                for v in variants:
                    if v in r:
                        tipo_ok = True
                        break

    return num_ok or tipo_ok

# ── Leer base de datos ────────────────────────────────────────
print(f"Leyendo {EXCEL_INPUT}...")
df = pd.read_excel(EXCEL_INPUT, dtype=str).fillna("-")
df.columns = df.columns.str.strip()
print(f"✓ {len(df)} registros cargados")

os.makedirs(RESULTADOS_DIR, exist_ok=True)

# ── Configurar Chrome ─────────────────────────────────────────
options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1400,900")
options.add_argument("--disable-extensions")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
# Rutas para Docker/Railway (Debian slim)
import shutil
chrome_bin = shutil.which("chromium") or shutil.which("chromium-browser") or "/usr/bin/chromium"
chromedriver_bin = shutil.which("chromedriver") or "/usr/bin/chromedriver"
options.binary_location = chrome_bin

service = Service(chromedriver_bin)
driver  = webdriver.Chrome(service=service, options=options)
wait    = WebDriverWait(driver, 25)

def buscar_por_nombre(nombre, departamento=None, ciudad=None):
    """
    Busca un nombre en la Rama Judicial.
    Retorna lista de procesos encontrados: [{radicado, despacho, partes, tipo}, ...]
    """
    procesos = []
    try:
        driver.get(URL)
        time.sleep(2)

        # ── Campo nombre ────────────────────────────────────────
        # El formulario tiene campos: Nombre/Razón Social, Tipo Persona, Departamento, Ciudad
        try:
            campo_nombre = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@type='text'][1]")
            ))
            campo_nombre.clear()
            campo_nombre.send_keys(nombre)
            time.sleep(0.5)
        except Exception as e:
            print(f"    ✗ No se pudo escribir nombre: {e}")
            return procesos

        # ── Selector Departamento ────────────────────────────────
        if departamento and departamento != "-":
            depto_rama = DEPTOS_MAP.get(departamento, departamento.upper())
            try:
                # Intentar con v-select (Vue.js autocomplete)
                selects = driver.find_elements(By.CSS_SELECTOR, ".v-select__slot input, .v-autocomplete input")
                for sel in selects[1:3]:  # Skip primer input (nombre), buscar depto
                    try:
                        sel.clear()
                        sel.send_keys(depto_rama[:5])  # Primeras letras para filtrar
                        time.sleep(1)
                        # Click primera opción del dropdown
                        opciones = driver.find_elements(By.CSS_SELECTOR, ".v-list-item__title, .menuable__content__active .v-list-item")
                        for op in opciones:
                            if depto_rama[:5] in op.text.upper():
                                driver.execute_script("arguments[0].click();", op)
                                time.sleep(0.8)
                                break
                        break
                    except:
                        continue
            except:
                pass

        # ── Botón Consultar ─────────────────────────────────────
        try:
            boton = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'success') or contains(text(),'Consultar') or contains(text(),'Buscar')]")
            ))
            driver.execute_script("arguments[0].click();", boton)
            time.sleep(4)
        except Exception as e:
            print(f"    ✗ No se encontró botón consultar: {e}")
            return procesos

        # ── Leer tabla de resultados ────────────────────────────
        try:
            filas = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
            for fila in filas:
                celdas = fila.find_elements(By.TAG_NAME, "td")
                if len(celdas) < 2:
                    continue
                proceso = {
                    "radicado":  celdas[0].text.strip() if len(celdas) > 0 else "",
                    "despacho":  celdas[1].text.strip() if len(celdas) > 1 else "",
                    "partes":    celdas[2].text.strip() if len(celdas) > 2 else "",
                    "fecha":     celdas[3].text.strip() if len(celdas) > 3 else "",
                    "tipo":      celdas[4].text.strip() if len(celdas) > 4 else "",
                }
                if re.match(r'^\d{20,}$', proceso["radicado"]):
                    procesos.append(proceso)

        except:
            # Intentar estructura alternativa de la tabla
            try:
                rows = driver.find_elements(By.CSS_SELECTOR, ".v-data-table tbody tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if not cells:
                        continue
                    texto_fila = " | ".join(c.text.strip() for c in cells)
                    # Buscar radicado de 23 dígitos en el texto
                    match = re.search(r'\d{23}', texto_fila)
                    if match:
                        procesos.append({
                            "radicado": match.group(0),
                            "despacho": cells[1].text.strip() if len(cells) > 1 else "",
                            "partes":   cells[2].text.strip() if len(cells) > 2 else "",
                            "fecha":    cells[3].text.strip() if len(cells) > 3 else "",
                            "tipo":     cells[4].text.strip() if len(cells) > 4 else "",
                        })
            except:
                pass

        # ── Verificar si hay mensaje "sin resultados" ───────────
        if not procesos:
            try:
                msg = driver.find_element(By.XPATH, "//*[contains(text(),'No se encontraron') or contains(text(),'no encontr') or contains(text(),'Sin resultados')]")
                print(f"    ℹ Sin resultados para: {nombre}")
            except:
                # Tomar screenshot para debug
                driver.save_screenshot(f"{RESULTADOS_DIR}/debug_{nombre[:20].replace(' ','_')}.png")

    except Exception as e:
        print(f"    ✗ Error en búsqueda: {str(e)[:100]}")

    return procesos


# ── Columnas a añadir al DataFrame ─────────────────────────────
col_radicado = "Número De Radicado"
df["Radicado_Encontrado"] = ""
df["Despacho_Encontrado"] = ""
df["Estado_Busqueda"] = ""
df["Todos_Procesos"] = ""

# ── Loop principal ────────────────────────────────────────────
total = len(df)
encontrados = 0
no_encontrados = 0
multiples = 0

for idx, row in df.iterrows():
    nombre     = str(row.get("Nombre Cliente", "")).strip()
    depto      = str(row.get("Departamento", "")).strip()
    ciudad_juz = str(row.get("Ciudad Del Juzgado", "")).strip()
    num_juz    = str(row.get("Numero Del Juzgado", "")).strip()
    tipo_juz   = str(row.get("Tipo De\nJuzgado", row.get("Tipo De Juzgado", ""))).strip()
    radicado_actual = str(row.get(col_radicado, "")).strip()

    print(f"[{idx+1}/{total}] {nombre[:40]}")

    # Si ya tiene radicado válido, saltar
    if re.match(r'^\d{20,}$', radicado_actual):
        df.at[idx, "Estado_Busqueda"] = "Ya tenía radicado"
        continue

    # Archivo cache
    cache_key = re.sub(r'[^a-zA-Z0-9]', '_', nombre[:40])
    cache_file = f"{RESULTADOS_DIR}/{cache_key}.json"

    # Usar cache si existe
    if os.path.exists(cache_file):
        with open(cache_file, "r", encoding="utf-8") as f:
            procesos = json.load(f)
        print(f"  (desde caché: {len(procesos)} procesos)")
    else:
        procesos = buscar_por_nombre(nombre, depto, ciudad_juz)
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(procesos, f, ensure_ascii=False, indent=2)
        time.sleep(2)

    if not procesos:
        df.at[idx, "Estado_Busqueda"] = "No encontrado"
        no_encontrados += 1
        print(f"  ✗ Sin procesos")
        continue

    # Guardar todos los procesos como referencia
    todos_str = " | ".join(f"{p['radicado']} ({p['despacho'][:30]})" for p in procesos[:5])
    df.at[idx, "Todos_Procesos"] = todos_str

    if len(procesos) == 1:
        # Solo uno → usar directamente
        df.at[idx, "Radicado_Encontrado"] = procesos[0]["radicado"]
        df.at[idx, "Despacho_Encontrado"] = procesos[0]["despacho"]
        df.at[idx, "Estado_Busqueda"] = "Encontrado (único)"
        encontrados += 1
        print(f"  ✓ Radicado único: {procesos[0]['radicado']}")
    else:
        # Múltiples → intentar cruzar con juzgado
        mejor = None
        for p in procesos:
            if coincide_juzgado(p["despacho"], num_juz, tipo_juz):
                mejor = p
                break

        if mejor:
            df.at[idx, "Radicado_Encontrado"] = mejor["radicado"]
            df.at[idx, "Despacho_Encontrado"] = mejor["despacho"]
            df.at[idx, "Estado_Busqueda"] = f"Encontrado ({len(procesos)} resultados, cruzado)"
            encontrados += 1
            print(f"  ✓ Cruzado: {mejor['radicado']} | {mejor['despacho'][:40]}")
        else:
            # No se pudo cruzar — poner el primero pero marcar como revisión manual
            df.at[idx, "Radicado_Encontrado"] = procesos[0]["radicado"]
            df.at[idx, "Despacho_Encontrado"] = procesos[0]["despacho"]
            df.at[idx, "Estado_Busqueda"] = f"REVISAR ({len(procesos)} resultados)"
            multiples += 1
            print(f"  ⚠ Múltiples sin cruce claro: {len(procesos)} procesos")

driver.quit()

# ── Exportar Excel ────────────────────────────────────────────
print(f"\nExportando a {EXCEL_OUTPUT}...")

# Actualizar la columna original de radicado con los encontrados
for idx, row in df.iterrows():
    if row["Radicado_Encontrado"] and row["Radicado_Encontrado"] != "":
        if df.at[idx, col_radicado] == "-" or not re.match(r'^\d{20,}$', str(df.at[idx, col_radicado])):
            df.at[idx, col_radicado] = row["Radicado_Encontrado"]

df.to_excel(EXCEL_OUTPUT, index=False)

# ── Dar formato al Excel ─────────────────────────────────────
wb = load_workbook(EXCEL_OUTPUT)
ws = wb.active

# Estilos
verde_ok    = PatternFill("solid", fgColor="C6EFCE")
amarillo_rev = PatternFill("solid", fgColor="FFEB9C")
rojo_no     = PatternFill("solid", fgColor="FFC7CE")
gris        = PatternFill("solid", fgColor="D9D9D9")
header_fill = PatternFill("solid", fgColor="003e25")
header_font = Font(bold=True, color="FFFFFF", name="Calibri")

# Cabecera
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font

# Columna Estado_Busqueda — buscar índice
estado_col = None
for i, cell in enumerate(ws[1], 1):
    if cell.value == "Estado_Busqueda":
        estado_col = i
        break

if estado_col:
    for row in ws.iter_rows(min_row=2):
        estado = str(row[estado_col-1].value or "")
        if "No encontrado" in estado:
            fill = rojo_no
        elif "REVISAR" in estado:
            fill = amarillo_rev
        elif "Encontrado" in estado:
            fill = verde_ok
        elif "Ya tenía" in estado:
            fill = gris
        else:
            continue
        for cell in row:
            cell.fill = fill

# Ancho de columnas
for col in ws.columns:
    max_len = max((len(str(c.value or "")) for c in col), default=10)
    ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

ws.freeze_panes = "A2"

wb.save(EXCEL_OUTPUT)

# ── Resumen ─────────────────────────────────────────────────
print(f"""
═══════════════════════════════════════
RESUMEN FINAL
═══════════════════════════════════════
Total procesados : {total}
✓ Encontrados    : {encontrados}
⚠ Revisar manual : {multiples}
✗ No encontrados : {no_encontrados}
Archivo salida   : {EXCEL_OUTPUT}
═══════════════════════════════════════
""")
