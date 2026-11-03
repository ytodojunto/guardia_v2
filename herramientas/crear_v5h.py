"""Arma PROYECTO_EXCEL_GUARDIA_v5h.xlsx a partir del v5j.

Cambios (todo lo demás queda igual):
  1. LISTAS: lista de puertos limpia (un nombre por puerto, con Z.C. y "?" como valores).
  2. TURNO, bloques de viajes: PUERTO E INICIO, DESEMBARCO 1 y DESEMBARCO 2 pasan de un texto
     ("SL 01/09:50") a tres celdas: puerto (desplegable) | FECHA | HORA. Usan las mismas 3 columnas
     que ya ocupaba cada campo.
  3. Desplegable de puertos en esas celdas y en DESDE/HASTA de las programaciones.
  4. FRANCOS: desplegable de prácticos.
Los desplegables nuevos son sugerencias: no bloquean otros valores.
"""
import re, sys, datetime as dt
from copy import copy
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

ORIGEN, DESTINO = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(ORIGEN)

PUERTOS = sorted(['ZONA COMUN', 'PUNTA PEREIRA (UY)', 'NUEVA PALMIRA (UY)', 'KM 171 PG', 'DEL GUAZU', 'IBICUY',
                  'MAZARUCA', 'KM 238 PA', 'SAN PEDRO', 'RAMALLO', 'SAN NICOLAS', 'VILLA CONSTITUCION',
                  'GENERAL LAGOS', 'VGG', 'ROSARIO', 'SAN LORENZO', 'SAN MARTIN', 'TIMBUES', 'DIAMANTE',
                  'SANTA FE', 'PARANA', 'ESCOBAR', 'CAMPANA', 'ZARATE', 'LAS PALMAS', 'DELTA DOCK', 'ATUCHA']) + ['A CONFIRMAR']
CODIGO_A_NOMBRE = {'SL': 'SAN LORENZO', 'S.L.': 'SAN LORENZO', 'ESC': 'ESCOBAR', 'ZC': 'ZONA COMUN', 'Z.C.': 'ZONA COMUN',
                   'Z.C': 'ZONA COMUN', 'CPNA': 'CAMPANA', 'NPAL': 'NUEVA PALMIRA (UY)', 'SN': 'SAN NICOLAS', 'S.N.': 'SAN NICOLAS',
                   'ROS': 'ROSARIO', 'TIM': 'TIMBUES', 'ZTE': 'ZARATE', 'RLLO': 'RAMALLO', 'SP': 'SAN PEDRO'}
LISTA_PUERTOS = "LISTAS!$A$2:$A$%d" % (1 + len(PUERTOS))
LISTA_PRACTICOS = "LISTAS!$B$2:$B$26"

# ── 1. LISTAS ──
ls = wb['LISTAS']
ls['A1'] = 'PUERTOS (sugerencias)'
for i in range(2, 60):
    ls.cell(i, 1).value = None
for i, p in enumerate(PUERTOS, start=2):
    ls.cell(i, 1).value = p
    if ls.cell(i, 2).has_style:               # mismo formato que la columna de prácticos
        ls.cell(i, 1)._style = copy(ls.cell(i, 2)._style)

# ── helpers de validaciones ──
def sacar_validaciones(ws, criterio):
    ws.data_validations.dataValidation = [d for d in ws.data_validations.dataValidation if not criterio(d)]

def lista(formula, rango, aviso=False, msg=None):
    d = DataValidation(type='list', formula1=formula, allow_blank=True)
    d.showErrorMessage = bool(aviso)
    if aviso:
        d.errorStyle = 'warning'; d.error = msg or 'Ese valor no está en la lista.'
    for parte in rango.split():
        d.add(parte)
    return d

# ── 2/3. TURNO: desembarcos en 3 celdas ──
ws = wb['TURNO']
ref = ws['D2'].value                                       # fecha de la guardia
ref = ref.date() if isinstance(ref, dt.datetime) else ref

def resolver(dia, hh, mm):
    mejor = None
    for dm in (-1, 0, 1):
        m = ref.month + dm; y = ref.year + (m - 1) // 12; m = (m - 1) % 12 + 1
        try: c = dt.date(y, m, dia)
        except ValueError: continue
        if mejor is None or abs((c - ref).days) < abs((mejor - ref).days): mejor = c
    return dt.datetime(mejor.year, mejor.month, mejor.day), dt.time(hh, mm)

RX = re.compile(r'^\s*([A-Za-zÁÉÍÓÚÑ.\s]+?)\s*:?\s*(\d{1,2})/(\d{1,2}):(\d{2})\s*$')
BLOQUES = [(54, 55, 66), (68, 69, 80), (82, 83, 94), (96, 97, 108), (110, 111, 122)]   # (encabezado, 1ª fila, última)
TRIOS = [(20, 21, 22), (23, 24, 25), (26, 27, 28)]                                   # T,U,V · W,X,Y · Z,AA,AB
sin_convertir = []

for hdr, r1, r2 in BLOQUES:
    for cp, cf, ch in TRIOS:
        assert str(ws.cell(hdr, cp).value or '').startswith(('PUERTO', 'DESEMBARCO')), (hdr, cp, ws.cell(hdr, cp).value)
        for c, txt in ((cf, 'FECHA'), (ch, 'HORA')):
            cel = ws.cell(hdr, c); cel._style = copy(ws.cell(hdr, cp)._style); cel.value = txt
        for r in range(r1, r2 + 1):
            p = ws.cell(r, cp)
            for c, fmt in ((cf, 'dd/mm'), (ch, 'hh:mm')):
                cel = ws.cell(r, c); cel._style = copy(p._style); cel.number_format = fmt
            v = p.value
            if v is None or str(v).strip() == '':
                continue
            m = RX.match(str(v))
            if m:
                cod = re.sub(r'\s+', '', m.group(1)).upper()
                nombre = CODIGO_A_NOMBRE.get(cod) or CODIGO_A_NOMBRE.get(m.group(1).strip().upper())
                if not nombre:
                    sin_convertir.append((p.coordinate, v)); continue
                f, h = resolver(int(m.group(2)), int(m.group(3)), int(m.group(4)))
                p.value = nombre; ws.cell(r, cf).value = f; ws.cell(r, ch).value = h
            elif not re.search(r'\d', str(v)):
                p.value = CODIGO_A_NOMBRE.get(str(v).strip().upper(), str(v).strip())   # solo destino, sin hora
            else:
                sin_convertir.append((p.coordinate, v))

# anchos: que entren "NUEVA PALMIRA" y "PUERTO E INICIO". En el v5j la columna A está definida
# para las 32 columnas a la vez y pisa a las demás: se deja cada columna con su ancho, sin superposiciones.
ANCHOS = {c: 8 for c in range(1, 33)}
ANCHOS.update({c: 5 for c in range(1, 8)})        # SERVICIO (A:B) y BUQUE (C:G)
ANCHOS.update({c: 5 for c in range(8, 16)})       # PRACTICO 1 (H:K) y PRACTICO 2 (L:O)
ANCHOS.update({c: 4.5 for c in (16, 17, 18, 19)}) # CALADO (P:Q) y CANAL (R:S)
ANCHOS.update({20: 15, 23: 15, 26: 15})           # PUERTO E INICIO / DESEMBARCOS (entra "NUEVA PALMIRA (UY)")
ANCHOS.update({c: 6 for c in (21, 22, 24, 25, 27, 28)})   # FECHA (dd/mm) y HORA (hh:mm)
for c, w in ANCHOS.items():
    d = ws.column_dimensions[L(c)]
    d.min = c; d.max = c; d.width = w

# validaciones TURNO: puertos (sugerencia), fecha y hora (con advertencia)
rng_p = ' '.join(f'{L(c)}{r1}:{L(c)}{r2}' for _, r1, r2 in BLOQUES for c in (20, 23, 26))
rng_f = ' '.join(f'{L(c)}{r1}:{L(c)}{r2}' for _, r1, r2 in BLOQUES for c in (21, 24, 27))
rng_h = ' '.join(f'{L(c)}{r1}:{L(c)}{r2}' for _, r1, r2 in BLOQUES for c in (22, 25, 28))
ws.add_data_validation(lista(LISTA_PUERTOS, rng_p))
dvf = DataValidation(type='date', operator='greaterThanOrEqual', formula1='43831', allow_blank=True,
                     showErrorMessage=True, errorStyle='warning', error='Poné el día y el mes, por ejemplo 02/11.')
[dvf.add(x) for x in rng_f.split()]; ws.add_data_validation(dvf)
dvh = DataValidation(type='time', operator='between', formula1='0', formula2='0.999988426', allow_blank=True,
                     showErrorMessage=True, errorStyle='warning', error='Poné la hora como hh:mm, por ejemplo 09:50.')
[dvh.add(x) for x in rng_h.split()]; ws.add_data_validation(dvh)

# ── programaciones: DESDE/HASTA con la lista nueva y datos de ejemplo normalizados ──
NORMAL = {'S.L.': 'SAN LORENZO', 'S.N.': 'SAN NICOLAS', 'Z.C': 'ZONA COMUN', 'Z.C.': 'ZONA COMUN', 'CPNA': 'CAMPANA', 'NPAL': 'NUEVA PALMIRA (UY)'}
for nombre in ('SUBIDAS', 'BAJADAS', 'CAMPANA'):
    s = wb[nombre]
    viejas = [d for d in s.data_validations.dataValidation if d.formula1 and 'LISTAS!$A' in d.formula1]
    rangos = ' '.join(str(d.sqref) for d in viejas) or 'K5:L19'
    sacar_validaciones(s, lambda d: d.formula1 and 'LISTAS!$A' in d.formula1)
    s.add_data_validation(lista(LISTA_PUERTOS, rangos))
    for row in s.iter_rows(min_row=5, max_row=s.max_row, min_col=11, max_col=12):
        for c in row:
            if isinstance(c.value, str) and c.value.strip() in NORMAL:
                c.value = NORMAL[c.value.strip()]

# ── FRANCOS: desplegable de prácticos ──
fr = wb['FRANCOS']
fr.add_data_validation(lista(LISTA_PRACTICOS, 'C3:C15 C19:C20'))

# ── reparaciones heredadas del v5j (Excel las marca como "contenido a reparar") ──
# Vista de Subidas/Bajadas/Campana: solo hay columnas congeladas (xSplit), pero el archivo traía además
# selecciones de paneles que no existen (bottomLeft/bottomRight). Se deja solo la del panel que existe.
for nombre in ('SUBIDAS', 'BAJADAS', 'CAMPANA'):
    sv = wb[nombre].sheet_view
    if sv.pane is not None and not sv.pane.ySplit:
        sv.selection = [s for s in sv.selection if s.pane == sv.pane.activePane]

# Las listas que apuntan a otra hoja se guardan SIN el "=" inicial (así lo pide el formato xlsx;
# Excel lo tolera pero Google Sheets no toma el desplegable).
for _ws in wb:
    for _d in _ws.data_validations.dataValidation:
        if _d.type == 'list' and _d.formula1 and _d.formula1.startswith('='):
            _d.formula1 = _d.formula1[1:]

wb.save(DESTINO)

# Logo del encabezado: se vuelve a poner el dibujo tal como lo guardó Excel (v5i), que no se repara.
if len(sys.argv) > 3:
    import zipfile, os
    orig = zipfile.ZipFile(sys.argv[3])
    PARTES = ('xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png')
    tmp = DESTINO + '.tmp'
    with zipfile.ZipFile(DESTINO) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, orig.read(item.filename) if item.filename in PARTES else zin.read(item.filename))
    os.replace(tmp, DESTINO)

print('guardado', DESTINO)
print('no convertidos:', sin_convertir)
