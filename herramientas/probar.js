// Uso: node probar.js ruta/al/archivo.xlsx [--json]
const fs = require('fs');
const ExcelJS = require('exceljs');
const G = require('../assets/guardia-reader.js');

(async () => {
  const ruta = process.argv[2];
  if (!ruta) { console.error('Falta la ruta del xlsx'); process.exit(1); }
  const d = await G.leerGuardia(ExcelJS, fs.readFileSync(ruta), require('jszip'));
  if (process.argv.includes('--json')) { console.log(JSON.stringify(d, null, 1)); return; }

  console.log('Hojas:', d.hojas.join(' | '));
  console.log('Guardia:', JSON.stringify(d.guardia));
  console.log('Catálogo:', d.catalogo.length, 'prácticos');
  console.log('Leyenda:', d.leyendaColores.map(l => `${l.id}=${l.hex}`).join(' '));
  console.log('\nTURNO (' + d.turno.length + '):');
  d.turno.forEach(t => console.log(' ', t.pos, t.marca || '-', t.practico, t.entrada, t.estado || '', t.observacion || ''));
  console.log('GUARDIA ROSARIO (' + d.guardiaRosario.length + ')');
  for (const [k, v] of Object.entries(d.viajes)) {
    console.log('\nVIAJES ' + k + ' (' + v.length + ')');
    v.forEach(x => console.log(' ', x.servicio, x.buque, '|', x.practico1, '/', x.practico2, '| calado', x.calado, x.canal || '',
      '| ini', x.inicio && (x.inicio.puerto + ' ' + x.inicio.cuando),
      '| d1', x.desembarco1 && (x.desembarco1.puerto + ' ' + x.desembarco1.cuando),
      '| d2', x.desembarco2 && (x.desembarco2.puerto + ' ' + x.desembarco2.cuando),
      JSON.stringify(x.colores), x.ejemplo ? '[EJEMPLO]' : ''));
  }
  console.log('\nFRANCOS: en franco', d.francos.enFranco.length, '| fuera de turno', d.francos.fueraDeTurno.length, '| licencias', d.francos.licencias.length);
  d.francos.enFranco.slice(0, 3).forEach(f => console.log(' ', f.practico, f.salida, '→', f.ingreso, f.observacion));
  for (const [k, p] of Object.entries(d.programacion)) {
    console.log('\nPROGRAMACIÓN ' + k + ' publicada', p.publicado, '| buques', p.buques.length, '| coef prom planilla', p.coefPromedioPlanilla, 'calculado', p.coefPromedio);
    p.buques.forEach(b => console.log(' ', b.marca || '-', b.servicio, b.buque, 'coef', b.coef, '| recalada', JSON.stringify(b.recalada), '| eta', JSON.stringify(b.eta),
      '|', b.desdeCod, '→', b.hastaCod, '| aut', b.autorizado, '|', b.estado || ''));
  }
  console.log('\nServicios cruzados (programación + viaje):',
    Object.entries(d.servicios).filter(([, s]) => s.viajes.length && s.programacion.length).map(([k]) => k).join(', '));
  console.log('\nAVISOS (' + d.avisos.length + '):');
  d.avisos.forEach(a => console.log(` [${a.nivel}] ${a.hoja}${a.celda ? ' ' + a.celda : ''} (${a.tipo}): ${a.mensaje}`));
})().catch(e => { console.error('ERROR', e); process.exit(1); });
