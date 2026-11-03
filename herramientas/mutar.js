const ExcelJS=require('exceljs'),G=require('../assets/guardia-reader.js');
const src=process.argv[2];
(async()=>{
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(src);
  const t=wb.getWorksheet('TURNO');
  t.getCell('K7').value='25:70';                 // hora inválida
  t.getCell('I8').value='ayer';                  // fecha inválida
  t.getCell('Z69').value='ESCOBAR: 20:00';       // sin día
  t.getCell('W69').value='SAN LORENZO: 03/03:30 - 05:00';
  t.getCell('T70').value='XYZ 01/09:00';         // puerto desconocido
  t.getCell('P70').value='ocho coma uno';        // calado inválido
  t.getCell('R70').value='ABC';                  // canal fuera de lista
  t.getCell('B8').value='?';                     // marca rara
  t.getCell('C9').value='Pérez';                 // fuera de catálogo
  t.getCell('C54').value='PRACT 2';              // roto el encabezado de a Rosario? (fila 54: PRACTICO 1 está en H54)
  t.getCell('L54').value='PRACT 2';
  const f=wb.getWorksheet('FRANCOS'); f.getCell('C3').value='MARTINEZ'; // franco y en turno
  const s=wb.getWorksheet('SUBIDAS'); s.getCell('H5').value='tarde'; s.getCell('F5').value='9,x';
  wb.removeWorksheet('TELEFONOS');
  const d=G.parseGuardia(wb);
  console.log('viaje aRosario:',d.viajes.aRosario.length,'(esperado 0: encabezado roto)');
  d.avisos.forEach(a=>console.log(`[${a.nivel}] ${a.hoja}${a.celda?' '+a.celda:''} (${a.tipo}): ${a.mensaje}`));
})().catch(e=>{console.error('CRASH',e);process.exit(1)});
