# Convenio Marco Scraper (Mercado Público / ChileCompra)
Scraper en Python (Playwright) para extraer precio mínimo por región desde Tienda Convenio Marco de Mercado Público/ChileCompra
- Abre la ficha del producto en la Tienda Convenio Marco.
- Recorre las **regiones**
- Extrae el **precio mínimo** por región (campo "Desde")
- Exporta los resultados a un Excel a partir de un DataFrame.

---
## Requisitos
- Python 3.10+ (recomendado).
- Dependencias Python: 'pandas', 'openpyxl', 'playwright'.
- Navegadores de Playwright instalados.

## Instalación
