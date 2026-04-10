# Convenio Marco Scraper (Mercado Público / ChileCompra)
Script en **Python** que automatiza la extracción de **precios mínimos por región** desde la **Tienda Convenio Marco** (Mercado Público).

A partir de un archivo Excel con IDs y el “slug” (complemento de link) del producto, el programa:
- Abre la ficha del producto en la Tienda Convenio Marco
- Recorre las **regiones**
- Extrae el **precio mínimo** por región (campo “Desde”)
- Exporta los resultados a un Excel

---
## Requisitos
- Python 3.10+ (recomendado)
- Dependencias Python: `pandas`, `openpyxl`, `playwright`
- Navegadores de Playwright instalados
---
## Instalación
1) Clona el repositorio:
```bash
git clone https://github.com/bastian-diazcortes/convenio-marco-scrapper/blob/main/ids_nombrelink.xlsx
```
2) Instala librerias y navegadores de Playwright
```
pip install -r requirements.txt
playwright install
```
---
## Archivo de entrada
El script espera un Excel llamado:
- ids_nombrelink.xlsx

Debe contener estas columnas:

|Columna | Descripción |
| :--- | :---: |
| ID | ID del producto |
| Nombre link | Slug del producto o URL Completa |


Por ejemplo:

|ID | Nombre link |
| :--- | :---: |
| 2178668 | pintura-alto-trafico-spes-amarillo-acrilica-acuosa-reflectante-tineta-unidad |

Si pones una URL completa en Nombre link, el script la usará tal cual.

---
## Ejecución
Ejecutar el archivo proyecto.py, sin tener los archivos de entrada y salida abiertos.

---
## Salida

El script genera un Excel:
- precios_por_region.xlsx

Las columnas del output son las siguientes:
- ID producto
- Nombre producto
- Región
- Precio mínimo
- URL

Además, puede generar hojas adicionales para las distintas marcas filtradas por la marca, que en este caso se trabaja con SPES, Cataphote y Lorenzini.

---
## Configuración rápida
Dentro del archivo se puede modificar:
- BASE_URL: ruta base del convenio/categoría
- INPUT_XLSX: nombre/ruta del Excel de entrada
- OUTPUT_XLSX: nombre/ruta del Excel de salida

---
## Notas de rendimiento
- El tiempo depende del número de productos/regiones y de la carga del sitio.
- Para acelerar se recomienda:
  - Bloquear imágenes/fuentes.
  - Evitar wait_until="networkidle" cuando no sea necesario.
  - Paralelizar por producto (múltiples workers).
---
## Uso responsable
Este proyecto automatiza navegación web. Úsalo con responsabilidad:
- Evita ejecutar con demasiada concurrencia.
- Respeta políticas del sitio y términos de uso aplicables.
- Considera pausas entre solicitudes si el sitio responde lento.

---
## Autor
Bastián Díaz Cortés
