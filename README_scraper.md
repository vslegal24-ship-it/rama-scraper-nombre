# VS Legal — Scraper Rama Judicial por Nombre
## Versión 1.0 — Búsqueda por NombreRazonSocial

---

## ARCHIVOS

| Archivo | Descripción |
|---|---|
| `scraper_rama_nombre.py` | Script principal — procesa los 545 registros |
| `diagnostico_rama.py` | Prueba rápida para verificar que los selectores funcionan |
| `piloto_numero_de_radicacion.xlsx` | Base de datos de entrada |
| `resultado_radicados.xlsx` | **Salida** — Excel con radicados encontrados |

---

## INSTALACIÓN

### 1. Dependencias Python
```bash
pip install selenium openpyxl pandas
```

### 2. ChromeDriver

**Linux (Ubuntu/Debian):**
```bash
apt-get install -y chromium-browser chromium-chromedriver
```

**Windows:**
- Descargar desde: https://chromedriver.chromium.org/downloads
- Usar la versión que coincida con tu Chrome instalado
- Agregar al PATH o indicar ruta en el script

**Mac:**
```bash
brew install chromedriver
```

### 3. Ajustar ruta de Chrome en el script

En `scraper_rama_nombre.py`, buscar estas líneas y ajustar según el OS:

```python
# Linux:
options.binary_location = "/usr/bin/chromium-browser"
service = Service("/usr/bin/chromedriver")

# Windows (ejemplo):
# options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"
# service = Service("C:/chromedriver/chromedriver.exe")
```

---

## USO

### Paso 1 — Ejecutar diagnóstico (primera vez)
```bash
python diagnostico_rama.py
```
Esto genera screenshots y muestra qué selectores funcionan. Si los resultados
de la tabla aparecen correctamente, el script principal funcionará.

### Paso 2 — Ejecutar scraper completo
```bash
python scraper_rama_nombre.py
```

El script:
1. Lee `piloto_numero_de_radicacion.xlsx`
2. Por cada registro sin radicado válido, busca el nombre en la Rama
3. Cruza resultados con Número y Tipo de Juzgado para identificar el proceso correcto
4. Guarda caché JSON en `resultados_nombre/` (evita repetir búsquedas)
5. Exporta `resultado_radicados.xlsx` con colores:
   - 🟢 Verde: Encontrado correctamente
   - 🟡 Amarillo: Múltiples resultados — revisar manualmente
   - 🔴 Rojo: No encontrado
   - ⬜ Gris: Ya tenía radicado

---

## LÓGICA DE CRUCE

Cuando la Rama devuelve **múltiples procesos** para el mismo nombre:

1. Se revisa el campo `Despacho` de cada resultado
2. Se compara con `Numero Del Juzgado` y `Tipo De Juzgado` de la base
3. Si hay coincidencia (ej: "JUZGADO 012 CIVIL MUNICIPAL" ↔ num=012, tipo=Civil Municipal) → se asigna ese radicado
4. Si no hay coincidencia clara → se marca como **REVISAR** y se pone el primer resultado

La columna `Todos_Procesos` siempre muestra todos los radicados encontrados para referencia.

---

## TIEMPOS ESTIMADOS

- ~3-4 segundos por registro
- 545 registros × 3.5s ≈ **~32 minutos** en total
- Con caché: los registros ya consultados se saltan automáticamente
- Se puede interrumpir y reanudar sin perder progreso

---

## NOTAS

- El script usa el **apellido + nombre** tal como está en la base. Si hay tildes
  o variaciones de nombre puede no encontrar. En ese caso revisar manualmente.
- Los 35 registros con `Ciudad Del Juzgado = "-"` se buscarán sin filtro de ciudad.
- La búsqueda NO filtra por ciudad del juzgado (solo por nombre y opcionalmente
  departamento del demandado). El cruce posterior usa el juzgado para identificar.
