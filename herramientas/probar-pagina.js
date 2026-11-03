const {JSDOM}=require('jsdom'),fs=require('fs');
const html=fs.readFileSync('../lector.html','utf8').replace(/<script src="[^"]+"><\/script>/g,'');
const dom=new JSDOM(html,{runScripts:'dangerously',pretendToBeVisual:true});
const w=dom.window; w.ExcelJS=require('exceljs'); w.JSZip=require('jszip');
const buf=fs.readFileSync(process.argv[2]);
const input=w.document.getElementById('archivo');
Object.defineProperty(input,'files',{value:[{name:'v5i.xlsx',arrayBuffer:async()=>buf.buffer.slice(buf.byteOffset,buf.byteOffset+buf.byteLength)}]});
input.dispatchEvent(new w.Event('change'));
setTimeout(()=>{
  const s=w.document.getElementById('salida');
  console.log('secciones h2:',[...s.querySelectorAll('h2')].map(e=>e.textContent).join(' | '));
  console.log('filas:',s.querySelectorAll('.fila').length,'| errores de página:',s.textContent.includes('No pude leer')?'SI':'no');
  console.log(s.textContent.slice(0,420));
  process.exit(0);
},2500);
