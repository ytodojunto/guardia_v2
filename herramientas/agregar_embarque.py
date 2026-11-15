"""Arma PROYECTO_EXCEL_GUARDIA_v5i.xlsx a partir del v5h.
Agrega, en los 5 bloques de viajes de la hoja TURNO, la columna opcional EMBARQUE PRACTICO 2 (puerto | FECHA | HORA)
a la derecha de OBSERVACIONES (columnas AD, AE, AF), con el mismo formato y desplegables que PUERTO E INICIO.
Uso: python3 agregar_embarque.py origen.xlsx destino.xlsx"""
import sys
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter as L
wb = openpyxl.load_workbook(sys.argv[1]); ws = wb['TURNO']
bloques = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'SERVICIO']
orig = {'AD': 'T', 'AE': 'U', 'AF': 'V'}                      # de dónde se copia el formato
titulos = {'AD': 'EMBARQUE PRACTICO 2', 'AE': 'FECHA', 'AF': 'HORA'}
def copiar(dst, src):
    dst._style = copy(src._style)
for h in bloques:
    for col, t in titulos.items():
        c = ws[f'{col}{h}']; c.value = t; copiar(c, ws[f'{orig[col]}{h}'])
    for r in range(h + 1, h + 13):
        for col in titulos:
            copiar(ws[f'{col}{r}'], ws[f'{orig[col]}{r}'])
    # desplegables: se agregan las celdas nuevas a los mismos rangos de T (puerto) y U (fecha)
    ini, fin = h + 1, h + 12
    for dv in ws.data_validations.dataValidation:
        s = str(dv.sqref)
        if f'T{ini}:T' in s and dv.type == 'list' and 'LISTAS' in (dv.formula1 or ''):
            dv.add(f'AD{ini}:AD{fin}')
        if f'U{ini}:U' in s and dv.type == 'date':
            dv.add(f'AE{ini}:AE{fin}')
ws.column_dimensions['AD'].width = 20; ws.column_dimensions['AE'].width = 11; ws.column_dimensions['AF'].width = 8
wb.save(sys.argv[2]); print('ok', bloques)
