# Guardia MULTIPAR · v2

Lee el **Excel nuevo de la guardia** (hojas TURNO, FRANCOS, SUBIDAS, BAJADAS, CAMPANA…) y lo muestra en el celular.
Independiente de la cuenta de Google de la empresa: el archivo se lee en el propio teléfono o computadora y no se sube a ningún lado.

## Cómo usarlo
Abrir `index.html` (GitHub Pages) → "Elegir el Excel de la guardia (.xlsx)".

## Qué hay
- `index.html` — página (lector + vista de la guardia). Se genera con `herramientas/armar-pagina.js`.
- `assets/guardia-reader.js` — el lector. Ubica cada sección por sus rótulos de la columna A y los encabezados, no por posiciones fijas. Funciona con celdas combinadas (dentro de la misma fila) y con los dos formatos de puerto/hora: texto `SL 01/09:50` o tres celdas puerto | FECHA | HORA.
- `excel/PROYECTO_EXCEL_GUARDIA_v5h.xlsx` — propuesta de Excel con desplegables de puertos y fecha/hora en celdas separadas.
- `herramientas/` — pruebas (`node probar.js archivo.xlsx`) y el script que arma el v5h.
- `docs/` — cómo se lee el Excel.

## Pendiente
- Conexión en vivo con el Google Sheet (Apps Script) cuando el Excel nuevo sea oficial (~11/11/2026).
- Herramienta de Proyección de despachos.
- Reemplazar la guardia vieja (`guardia_proyecto`) cuando esta esté validada.
