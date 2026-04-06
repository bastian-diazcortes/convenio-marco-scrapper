# ============================
# Proyecto - Mercado Público - PROCOMERCE
# Hecho por: Bastián Díaz Cortés
# Lee IDs + slug (complemento de link) desde Excel, entra a cada ficha de producto, recorre regiones y extrae el precio mínimo.
# Exporta a Excel: ID | Nombre producto | Región | Precio mínimo | URL
# =============================


#Librerías utilizadas
import re #Expresiones regulares para buscar patrones en texto (ID,precios)
import time #PAra implementar espera por polling (wait_until)
import pandas as pd #Leer y escribir Excel y manejar df.
from pathlib import Path #Para rutas de archivos de forma robusta
from playwright.sync_api import sync_playwright,TimeoutError as PlaywrightTimeoutError #automatiza el nagerrador, TimeoutError se usa para manejar esperar que expiran.

#Ajustas a archivo de entrada
BASE_DIR = Path(__file__).resolve().parent #Carpeta donde está este script
INPUT_XLSX = BASE_DIR / "ids_nombrelink.xlsx" #Nombre del archivo de entrada
OUTPUT_XLSX = BASE_DIR / "precios_por_region.xlsx" #Nombre del archivo de salida

BASE_URL = "https://conveniomarco2.mercadopublico.cl/ferreteria2/" #URL del convenio y categoría

TARGET_SUPPLIER = "COMERCIALIZADORA DE PRODUCTOS Y SERVICIOS PROCOMERCE SPA" #se define proveedor
# ==========================
# Funciones auxiliares
# ==========================

def build_product_url(slug_or_url:str) -> str:
    """
    Recibe un slug (ejemplo: 'pintura-alto-trafico...') o una URL completa
    Devuelve la URL completa del producto
    """
    s = str(slug_or_url).strip() #Convierte a string la url y limpia espacios
    if s.startswith("http://") or s.startswith("https://"): #Si es URL completa, la retorna tal cual
        return s
    return BASE_URL.rstrip("/")+"/"+s.lstrip("/") #Si es slug, se concatena con la base, generalmente será así.

def to_int_price(text:str):
    """
    Convierte precios escritos como: '$66.183' a 66183 (int).
    """
    if not text:
        return None
    digits = re.sub(r"[^\d]","",text) #Elimina todo lo que no sea digito, es decir, quita $, espacios, etc.
    return int(digits) if digits else None

def extract_min_price_from_page_text(t:str):
    """
    Busca en el texto de la página el bloque de precios  "Desde YYY1 hasta YYY2" y devuelve el precio mínimo 
    """
    #Está en 2 formatos posibles 1) Desde X - Y , 2) Desde X hasta Y
    m = re.search(
        r"Desde\s*\$?\s*([\d\.\,]+)\s*-\s*\$?\s*([\d\.\,]+)|Desde\s*\$?\s*([\d\.\,]+)\s*Hasta\s*\$?\s*([\d\.\,]+)",
        t,
        flags=re.IGNORECASE
    )
    if not m:
        return None
    desde = m.group(1) or m.group(3)
    return to_int_price(desde)

def wait_until(predicate, timeout_ms=25000, step_ms=250):
    """
    Espera 'activa' por polling hasta predicate() sea True.
    predicate: función que retorna True/False (ej: "ya hay opciones de región)
    timeout_ms: tiempo máximo total.
    step_ms: intervalo entre intentos

    wait_for_function a veces falla si el elemento no es realmente un select o si el DOM cambia. Plling en Python es más controlable y robusto
    """
    end = time.time() + timeout_ms / 1000 #Momento límite en segundos.
    last = None #Guarda el último resultado, sirve para debug
    while time.time() < end:
        try:
            last = predicate()
            if last:
                return True
        except Exception:
            #Si el DOM cambia justo en el momento de consultar puede lanzar error, lo ignoramos y reintentamos
            pass
        time.sleep(step_ms / 1000)
    
    #Si termina el tiempo, lanza TimeoutError para que main() lo capture y registro
    raise PlaywrightTimeoutError(f"Timeout esperando condición ({timeout_ms}ms). Último={last}")


# ==========================
# Detección y manejo del selector de región
# ==========================

def get_region_controller(page):
    """
    Detecta cómo se selecciona la región en la ficha. 
    En algunas páginas puede ser un select nativo. En otras puede ser un dropdown tipo combobox
    Devuelve un diccionario
        {'mode':'select' o 'combo','handle':locator}
    Existe para que el scraper funcione con ambas variantes sin romperse
    """
    # Intento 1: <select> asociado a label REGIÓN
    region_select = page.locator(
        "xpath=//label[contains(translate(., 'región', 'REGIÓN'),'REGIÓN')]/following::select[1]"
    )
    if region_select.count() > 0:
        return {"mode": "select", "handle": region_select}

    # Intento 2: algún select visible si no hay label claro
    any_select = page.locator("select:visible").first
    if any_select.count() > 0:
        return {"mode": "select", "handle": any_select}

    # Intento 3: combobox “Región” (dropdown custom)
    # (Si no trae 'name', usamos el primer combobox visible como fallback)
    combo = page.get_by_role("combobox", name=re.compile(r"Regi[oó]n", re.IGNORECASE))
    if combo.count() > 0:
        return {"mode": "combo", "handle": combo}
    
    #Ultimo fallback, primer combobox visible
    combo_any = page.get_by_role("combobox").first
    return {"mode": "combo", "handle": combo_any}


def get_regions(page, controller):
    """
    Extrae la lista de regiones disponibles según el tipo de control detectado.
    Returna lista de tuplas
    """
    if controller["mode"] == "select": #Caso A: Select nativo
        region_select = controller["handle"]
        region_select.wait_for(state="visible", timeout=20000)

        def has_valid_options(): #Esperar a que el select tenga opciones válidas (no solo "Elegir...")
            opts = region_select.locator("option")
            valid = 0
            for i in range(opts.count()):
                val = (opts.nth(i).get_attribute("value") or "").strip()
                txt = (opts.nth(i).inner_text() or "").strip()
                if val and txt and "Elegir" not in txt:
                    valid += 1
            return valid >= 2

        wait_until(has_valid_options, timeout_ms=30000)

        opts = region_select.locator("option") #Construir la lista (value,label)
        region_items = []
        for i in range(opts.count()):
            val = (opts.nth(i).get_attribute("value") or "").strip()
            label = (opts.nth(i).inner_text() or "").strip()
            if not val or not label or "Elegir" in label:
                continue
            region_items.append((val, label))

        return region_items

    # Caso B: dropdown custom (combobox/listbox)
    region_combo = controller["handle"]
    region_combo.wait_for(state="visible", timeout=20000)

    # Abrir dropdown
    region_combo.click()

    # El listado suele ser listbox; si no, igual intentamos con opciones role=option visibles
    listbox = page.get_by_role("listbox")
    if listbox.count() > 0:
        listbox.wait_for(state="visible", timeout=20000)

        def has_options():
            return listbox.get_by_role("option").count() >= 2

        wait_until(has_options, timeout_ms=30000)
        opts = listbox.get_by_role("option")
    else:
        # fallback: cualquier opción visible
        def has_options_any():
            return page.get_by_role("option").count() >= 2

        wait_until(has_options_any, timeout_ms=30000)
        opts = page.get_by_role("option")

    region_items = []
    for i in range(opts.count()):
        label = (opts.nth(i).inner_text() or "").strip()
        if not label or "Elegir" in label:
            continue
        region_items.append((label, label))

    # Cerrar dropdown (clic afuera) para evitar que tape cosas
    page.keyboard.press("Escape")
    return region_items


def select_region(page, controller, value, label):
    """
     Selecciona una región, dependiendo del tipo de control:

    - select: region_select.select_option(value)
    - combo : click en combobox y luego click en opción por texto (label)

    Por qué existe:
      Abstrae la selección, para que el loop por regiones sea el mismo para ambos casos.
    """
    if controller["mode"] == "select":
        controller["handle"].select_option(value)
        return

    # combo
    region_combo = controller["handle"]
    region_combo.click()
    # Selecciona por texto visible
    page.get_by_role("option", name=label).click()
    
    #Ordenar por precio ascendente para rank real
    offers_sorted = sorted(offers,key=lambda x: x[1])
    for idx,(name,_) in enumerate(offers_sorted,start=1):
        if normalize_name(name) == target_norm:
            return idx
    return None



# ==========================
# Detección y manejo del selector de región
# ==========================

def scrape_one_product(page,url:str,fallback_id=None):
    """
    Abre una ficha de producto, detecta regiones, recorre cada región
    y extrae el precio mínimo.

    Retorna:
      Lista de dicts (una fila por región):
        {
          ID, Nombre producto, Región, Precio mínimo, URL
        }
    """
    page.goto(url, wait_until="networkidle") #networkidle espera a que el sitio termine sus llamadas de red
    page.wait_for_timeout(600) #pequeña pausa extra

    #Nombre del producto
    try:
        product_name = page.locator("h1").first.inner_text().strip()
    except Exception:
        product_name = None
    
    #ID: si no lo detecta usa el de excel. Si falla se usa fallback_id desde el Excel
    body_text = page.locator("body").inner_text()
    m = re.search(r"\bID\b\s*[\r\n]*\s*(\d{6,})",body_text)
    product_id = m.group(1) if m else str(fallback_id) if fallback_id else None

    controller = get_region_controller(page)
    region_items = get_regions(page,controller)

    if not region_items:
        raise RuntimeError("No se cargaron regiones (sin opciones).")
    
    rows = []
    for value,region_label in region_items: #Recorre todas las regiones disponibles
        select_region(page,controller,value,region_label) #Seleccionar región

        page.wait_for_timeout(1300) #Espera a que el bloque de precios se actualice para esa región
        t = page.locator("body").inner_text() #Extrae el precio mínimo desde el texto completo de la página
        precio_min = extract_min_price_from_page_text(t)

        rows.append({
            "ID":product_id,
            "Nombre producto":product_name,
            "Región":region_label,
            "AUX":product_name+region_label,
            "Precio mínimo":precio_min,
            "URL":url,
        }) #Guardar la fila resultado
    return rows

# ==========================
# Programa principal
# ==========================

def main():
    """
    Lee Excel de entrada, para cada fila: construye URL, scrapea producto por región, guarda todo en un Excel final
    """
    df_in = pd.read_excel(INPUT_XLSX)
    #Validación
    required = {"ID","Nombre link"}
    if not required.issubset(df_in.columns):
        raise ValueError(f"Tu excel debe tener columnas: {required}. Encontré: {list(df_in.columns)}")
        
    all_rows = []

    with sync_playwright() as p: #inicia playwright y un navegador Chronium
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_extra_http_headers({"Accept-Language":"es-CL,es;q=0.9"})

        for _, r in df_in.iterrows():
            pid = r["ID"]
            url = build_product_url(r["Nombre link"])
            try:
                rows = scrape_one_product(page,url,fallback_id=pid) or []
                all_rows.extend(rows)
            except Exception as e:
                all_rows.append({
                    "ID":pid,
                    "Nombre producto":None,
                    "Región": None,
                    "AUX": None,
                    "Precio mínimo":None,
                    "URL": url,
                    "Error":str(e)
                })
        
        browser.close()

    df_out = pd.DataFrame(all_rows)

    "Orden final"
    base_cols = ["ID","Nombre producto","Región","Precio mínimo","URL"]
    extra_cols = [c for c in df_out.columns if c not in base_cols]
    df_out = df_out[base_cols+extra_cols]

    #Hojas aparte
    df_spes = df_out[df_out["Nombre producto"].str.contains("SPES",na=False)]
    df_cataphote = df_out[df_out["Nombre producto"].str.contains("CATAPHOTE",na=False)]
    df_lorenzini = df_out[df_out["Nombre producto"].str.contains("LORENZINI",na=False)]

    with pd.ExcelWriter(OUTPUT_XLSX) as writer:
        df_out.to_excel(writer,sheet_name="General",index=False)
        df_spes.to_excel(writer,sheet_name="SPES",index=False)
        df_cataphote.to_excel(writer,sheet_name="Cataphote",index=False)
        df_lorenzini.to_excel(writer,sheet_name="Lorenzini",index=False)
    
    print(f"OK -> {OUTPUT_XLSX} (filas: {len(df_out)})")

if __name__ == "__main__":
    main()