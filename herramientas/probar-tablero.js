// Prueba el tablero completo con un Apps Script simulado.  Uso: node probar-tablero.js archivo.xlsx
const {JSDOM}=require('jsdom'),fs=require('fs'),path=require('path');
const R=path.join(__dirname,'..'); const buf=fs.readFileSync(process.argv[2]);
const html=fs.readFileSync(R+'/index.html','utf8').replace(/<script src="[^"]+"><\/script>/g,'');
const dom=new JSDOM(html,{runScripts:'dangerously',pretendToBeVisual:true,url:'https://x.github.io/guardia_v2/'});
const w=dom.window; w.ExcelJS=require('exceljs'); w.JSZip=require('jszip'); w.GUARDIA_API_URL='https://script.example/exec';
w.fetch=async()=>({json:async()=>({nombre:'G.xlsx',modificado:new Date().toISOString(),consultado:new Date().toISOString(),xlsx:buf.toString('base64')})});
w.eval(fs.readFileSync(R+'/assets/guardia-reader.js','utf8')); w.eval(fs.readFileSync(R+'/assets/data-loader.js','utf8'));
w.sessionStorage.setItem('mp_clave','x'); w.document.dispatchEvent(new w.CustomEvent('mp:auth-ok'));
setTimeout(()=>{
  const d=w.document, n=(id,sel)=>d.querySelectorAll('#'+id+' '+sel).length;
  console.log('turno',n('listaTurno','.prac-row'),'| francos',n('listaFrancos','.franco-row'),'| previstos',n('listaPrevistos','.buque-row'),'| buques',d.querySelectorAll('#listaBuquesDesp .desp-card').length,'| novedades',n('listaNov','.obs-item'),'| remises cards',d.querySelectorAll('#listaRemises .card').length);
  console.log('banner:',d.getElementById('loadingBanner').className,'|',d.getElementById('loadingBanner').textContent.slice(0,120));
  console.log('header:',d.querySelector('.hdr-fecha').textContent);
  console.log('obs turno 1:',d.querySelector('#listaTurno .prac-detalle').textContent);
  console.log('obs buque:',[...d.querySelectorAll('#listaBuquesDesp .buque-obs')].map(e=>e.textContent).join(' | ')||'(ninguna)');
  console.log('tel:',d.querySelectorAll('#sec-info .tel-row').length);
  process.exit(0);
},4000);
