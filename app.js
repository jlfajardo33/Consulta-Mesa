let invitados=[];
fetch('Invitados - Mesa.xlsx').then(r=>r.arrayBuffer()).then(b=>{
 const wb=XLSX.read(b); const ws=wb.Sheets[wb.SheetNames[0]];
 const data=XLSX.utils.sheet_to_json(ws);
 console.log(data);
 invitados=data;
});
const inp=document.getElementById('buscar');
inp.addEventListener('input',()=>{
 let q=inp.value.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
 let res=document.getElementById('resultados');res.innerHTML='';
 if(q.length<3)return;
 invitados.filter(x=>String(x.Invitados||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().includes(q))
 .forEach(x=>{
  let d=document.createElement('div');d.className='item';d.textContent=x.Invitados;
  d.onclick=()=>document.getElementById('mesa').innerHTML='Tu mesa es: '+x.Mesa;
  res.appendChild(d);
 });
});
