"""PROYECTO_EXCEL_GUARDIA_v5k.xlsx = v5i + ajustes por los avisos "No válido" de la planilla real:
  1. LISTAS: se agrega PUNTA ALVEAR y las listas de puertos admiten hasta la fila 40.
  2. SUBIDAS / BAJADAS / CAMPANA: las columnas HORA (H y J) aceptan hora, AM, PM, NOCHE o FOND (la regla "hora entre 0 y 23:59"
     rechazaba el texto; se quita) y quedan solo como advertencia.
Uso: python3 ajustes_v5k.py origen.xlsx destino.xlsx"""
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ls = wb['LISTAS']
puertos = [ls.cell(i, 1).value for i in range(2, 60) if ls.cell(i, 1).value]
fin = [p for p in puertos if p == 'A CONFIRMAR']
puertos = sorted(set(p for p in puertos if p != 'A CONFIRMAR') | {'PUNTA ALVEAR'}) + fin
for i in range(2, 60): ls.cell(i, 1).value = None
for i, p in enumerate(puertos, start=2):
    ls.cell(i, 1).value = p
    ls.cell(i, 1)._style = ls.cell(2, 1)._style
for ws in wb.worksheets:
    for dv in ws.data_validations.dataValidation:
        if dv.formula1 and 'LISTAS!$A$2:$A$29' in dv.formula1:
            dv.formula1 = dv.formula1.replace('$A$29', '$A$40')
for n in ('SUBIDAS', 'BAJADAS', 'CAMPANA'):
    ws = wb[n]; keep = []
    for dv in ws.data_validations.dataValidation:
        s = str(dv.sqref)
        if dv.type == 'time' and ('H5' in s or 'J5' in s): continue            # la regla que rechazaba AM / PM / FOND
        if dv.type == 'custom' and dv.formula1 and 'ISNUMBER' in dv.formula1:
            dv.formula1 = dv.formula1.replace('H5="PM")', 'H5="PM",H5="NOCHE",H5="FOND")').replace('J5="PM")', 'J5="PM",J5="NOCHE",J5="FOND")')
            dv.error = 'Poné una hora (ej. 14:00) o AM, PM, NOCHE, FOND.'; dv.errorTitle = 'Hora'
        keep.append(dv)
    ws.data_validations.dataValidation = keep
wb.save(sys.argv[2]); print('ok', len(puertos), 'puertos')
