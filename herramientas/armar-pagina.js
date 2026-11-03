const fs=require('fs');
const lector=fs.readFileSync('../assets/guardia-reader.js','utf8');
if(lector.includes('</script>')) throw new Error('el lector contiene </script>');
const html=fs.readFileSync('plantilla.html','utf8').replace('/*__LECTOR__*/',()=>lector);
fs.writeFileSync('../lector.html',html);
console.log('lector-guardia.html',Math.round(html.length/1024)+' KB');
