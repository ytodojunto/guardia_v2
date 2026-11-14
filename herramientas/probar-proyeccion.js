// Prueba la página de Proyección leyendo el Excel como si viniera de Drive.  Uso: node probar-proyeccion.js archivo.xlsx
const {JSDOM}=require('jsdom'),fs=require('fs'),path=require('path');
const R=path.join(__dirname,'..'); const buf=fs.readFileSync(process.argv[2]);
let html=fs.readFileSync(R+'/proyeccion.html','utf8');
html=html.replace(/<script src="assets\/guardia-reader.js[^"]*"><\/script>/,()=>'<script>'+fs.readFileSync(R+'/assets/guardia-reader.js','utf8')+'</script>')
         .replace(/<script src="assets\/config.js[^"]*"><\/script>/,()=>'<script>'+fs.readFileSync(R+'/assets/config.js','utf8')+'</script>')
         .replace(/<script src="https[^"]+"><\/script>/g,'');
const dom=new JSDOM(html,{runScripts:'dangerously',pretendToBeVisual:true,url:'https://x.github.io/guardia_v2/proyeccion.html',
  beforeParse(w){ w.ExcelJS=require('exceljs'); w.JSZip=require('jszip'); w.sessionStorage.setItem('mp_clave','x');
    w.fetch=async()=>({json:async()=>({nombre:'G.xlsx',modificado:new Date().toISOString(),xlsx:buf.toString('base64')})}); }});
setTimeout(()=>{
  const d=dom.window.document, t=id=>(d.getElementById(id)||{}).textContent||'';
  console.log('estado:',t('srcst').slice(0,140));
  console.log('mensaje:',t('srcmsg'));
  console.log('panorama:',t('pan').replace(/\s+/g,' ').slice(0,300));
  console.log('barcos (cards):',d.querySelectorAll('#res .card').length,'| filas prácticos:',d.querySelectorAll('#inner .prow').length);
  console.log('avisos:',[...d.querySelectorAll('#avisos li')].map(l=>l.textContent.slice(0,140)));
  process.exit(0);
},5000);
