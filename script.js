/* ========== Util & Progres ========== */
const $=(s,r=document)=>r.querySelector(s),$$=(s,r=document)=>[...r.querySelectorAll(s)];
const KEYS=[['m1','Pengenalan Excel'],['m2','Workbook & Worksheet'],['m3','Row, Column, Cell, Range'],['m4','Input & edit data'],['m5','Operator matematika'],['fSUM','Rumus SUM'],['fAVERAGE','Rumus AVERAGE'],['fMAX','Rumus MAX'],['fMIN','Rumus MIN'],['fCOUNT','Rumus COUNT'],['sim','Simulasi rumus'],['kuis','Kuis'],['eval','Evaluasi akhir']];
let P=JSON.parse(localStorage.getItem('xlP')||'{}');
function mark(k){P[k]=true;localStorage.setItem('xlP',JSON.stringify(P));renderProg()}
function renderProg(){
  const n=KEYS.filter(([k])=>P[k]).length,pct=Math.round(n/KEYS.length*100);
  $('#pbMini').style.width=$('#pbBig').style.width=pct+'%';
  $('#pctMini').textContent=pct+'%';$('#pctBig').textContent=`${pct}% selesai (${n} dari ${KEYS.length} langkah)`;
  $('#checklist').innerHTML=KEYS.map(([k,t])=>`<li class="${P[k]?'y':''}">${P[k]?'✅':'⬜'} ${t}</li>`).join('');
  $$('.chap').forEach(c=>c.classList.toggle('finished',!!P[c.dataset.k]));
}
$('#resetAll').onclick=()=>{if(confirm('Hapus semua progres?')){P={};localStorage.removeItem('xlP');renderProg()}};
$$('.chap .done').forEach(b=>b.onclick=()=>mark(b.closest('.chap').dataset.k));

/* ========== Navigasi ========== */
const pages=$$('.page');
$('#nav').innerHTML=pages.map(p=>`<button data-go="${p.id}">${p.dataset.t}</button>`).join('');
document.addEventListener('click',e=>{const g=e.target.closest('[data-go]');if(g)go(g.dataset.go)});
function go(id){pages.forEach(p=>p.classList.toggle('active',p.id===id));$$('#nav button').forEach(b=>b.classList.toggle('on',b.dataset.go===id));scrollTo(0,0)}

/* ========== Mesin rumus mini ========== */
const nums=a=>a.flat(Infinity).filter(x=>typeof x==='number');
const FN={
  SUM:(...a)=>nums(a).reduce((x,y)=>x+y,0),
  AVERAGE:(...a)=>{const n=nums(a);return n.length?n.reduce((x,y)=>x+y,0)/n.length:'#DIV/0!'},
  MAX:(...a)=>Math.max(...nums(a)),
  MIN:(...a)=>Math.min(...nums(a)),
  COUNT:(...a)=>nums(a).length
};
function calc(raw,get){
  let s=raw.slice(1).toUpperCase();
  s=s.replace(/([A-Z])(\d+):([A-Z])(\d+)/g,(m,c1,r1,c2,r2)=>{
    const o=[];
    for(let c=c1.charCodeAt(0);c<=c2.charCodeAt(0);c++)for(let r=+r1;r<=+r2;r++){const v=get(String.fromCharCode(c)+r);o.push(typeof v==='number'?v:null)}
    return JSON.stringify(o)});
  s=s.replace(/\b[A-Z]\d+\b/g,a=>{const v=get(a);return typeof v==='number'?v:0});
  const ids=s.match(/[A-Z]+/g)||[];
  if(ids.some(i=>!FN[i])||!/^[\w\s+\-*/^().,\[\]]*$/.test(s))throw 0;
  const r=Function(...Object.keys(FN),'return '+s.replace(/\^/g,'**'))(...Object.values(FN));
  return typeof r==='number'&&!isFinite(r)?'#ERR':r;
}
const fmt=v=>v==null?'':typeof v==='number'?Math.round(v*100)/100:v;

function Sheet(el,rows,cols,init,editable,onChange){
  const M={...init};
  const get=(a,d=0)=>{
    const r=M[a];if(r==null||r==='')return null;
    if(typeof r==='number')return r;
    if(r[0]==='='){if(d>20)return '#REF!';try{return calc(r,x=>get(x,d+1))}catch{return '#ERR'}}
    return isNaN(r)?r:+r;
  };
  el.innerHTML='<table class="sheet"><tr><th></th>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr>'+
    rows.map(r=>`<tr><th>${r}</th>${cols.map(c=>`<td><input data-a="${c+r}" ${editable(c+r)?'':'readonly class="ro"'}></td>`).join('')}</tr>`).join('')+'</table>';
  const refresh=()=>$$('input',el).forEach(i=>i.value=fmt(get(i.dataset.a)));
  el.addEventListener('focusin',e=>{const i=e.target;if(i.dataset.a&&!i.readOnly)i.value=M[i.dataset.a]??''});
  el.addEventListener('focusout',e=>{const i=e.target;if(!i.dataset.a)return;if(!i.readOnly)M[i.dataset.a]=i.value;refresh();onChange&&onChange(i.dataset.a)});
  el.addEventListener('keydown',e=>{if(e.key==='Enter')e.target.blur()});
  refresh();
  return{get,M,refresh};
}

/* ========== Data contoh: tabel nilai siswa ========== */
const DATA={A1:'Nama',B1:'Tugas',C1:'UTS',D1:'UAS',
  A2:'Andi',B2:80,C2:75,D2:90, A3:'Bela',B3:85,C3:80,D3:88, A4:'Citra',B4:90,C4:85,D4:92,
  A5:'Dika',B5:70,C5:65,D5:78, A6:'Eka',B6:95,C6:90,D6:96};

/* ========== Penjelasan interaktif ========== */
let mode='cell',anc=null;
const demoEl=$('#demo');
demoEl.innerHTML='<table class="dm"><tr><td class="hd"></td>'+'ABCD'.split('').map(c=>`<td class="hd">${c}</td>`).join('')+'</tr>'+
  [1,2,3,4,5,6].map(r=>`<tr><td class="hd">${r}</td>${'ABCD'.split('').map(c=>`<td data-a="${c+r}">${DATA[c+r]}</td>`).join('')}</tr>`).join('')+'</table>';
const parse=a=>[a.charCodeAt(0),+a.slice(1)];
function selectRange(a,b){
  const [c1,r1]=parse(a),[c2,r2]=parse(b);
  const cmin=Math.min(c1,c2),cmax=Math.max(c1,c2),rmin=Math.min(r1,r2),rmax=Math.max(r1,r2);
  $$('#demo td[data-a]').forEach(td=>{const [c,r]=parse(td.dataset.a);td.classList.toggle('sel',c>=cmin&&c<=cmax&&r>=rmin&&r<=rmax)});
  const s=String.fromCharCode(cmin)+rmin,e=String.fromCharCode(cmax)+rmax;
  return{s,e,n:(cmax-cmin+1)*(rmax-rmin+1)};
}
demoEl.onclick=e=>{
  const td=e.target.closest('td[data-a]');if(!td)return;const a=td.dataset.a,info=$('#demoInfo');
  if(mode==='cell'){selectRange(a,a);info.innerHTML=`Cell <b>${a}</b> = <b>${DATA[a]}</b>. Kolom ${a[0]}, baris ${a.slice(1)}.`;mark('m3')}
  else if(!anc){anc=a;selectRange(a,a);info.textContent=`Cell awal ${a}. Sekarang klik cell akhir.`}
  else{const r=selectRange(anc,a);anc=null;info.innerHTML=`Range <b>${r.s}:${r.e}</b> berisi <b>${r.n}</b> cell. Contoh: <code>=SUM(${r.s}:${r.e})</code>`;mark('m3')}
};
$$('[data-mode]').forEach(b=>b.onclick=()=>{mode=b.dataset.mode;anc=null;$$('[data-mode]').forEach(x=>x.classList.toggle('on',x===b));$$('#demo .sel').forEach(t=>t.classList.remove('sel'));$('#demoInfo').textContent=mode==='cell'?'Klik salah satu cell.':'Klik cell awal, lalu cell akhir.'});

/* ========== Rumus ========== */
const VALS=[80,90,75,85,100];
const FORM={
  SUM:['Menjumlahkan semua angka dalam range.','=SUM(range)'],
  AVERAGE:['Mencari nilai rata-rata.','=AVERAGE(range)'],
  MAX:['Mencari nilai tertinggi.','=MAX(range)'],
  MIN:['Mencari nilai terendah.','=MIN(range)'],
  COUNT:['Menghitung banyaknya cell yang berisi angka.','=COUNT(range)']
};
$('#ftabs').innerHTML=Object.keys(FORM).map(k=>`<button class="btn" data-f="${k}">${k}</button>`).join('');
function showF(k,m=true){
  const [fn,syn]=FORM[k];
  $$('[data-f]').forEach(b=>b.classList.toggle('on',b.dataset.f===k));
  $('#fbox').innerHTML=`<h3>${k}</h3><p><b>Fungsi:</b> ${fn}</p><p><b>Sintaks:</b> <code>${syn}</code></p>
    <p><b>Contoh data (B2:B6):</b></p><div class="nums">${VALS.map(v=>`<span>${v}</span>`).join('')}</div>
    <p><b>Contoh rumus:</b> <code>=${k}(B2:B6)</code></p><p><b>Hasil:</b></p><div class="res">${fmt(FN[k](VALS))}</div>`;
  if(m)mark('f'+k);
}
$('#ftabs').onclick=e=>{const b=e.target.closest('[data-f]');if(b)showF(b.dataset.f)};
showF('SUM',false);

/* ========== Simulasi ========== */
let sim;
function initSim(){
  sim=Sheet($('#simSheet'),[1,2,3,4,5,6,7,8,9],'ABCDE'.split(''),DATA,()=>true,a=>{
    const raw=sim.M[a];
    if(typeof raw==='string'&&raw[0]==='='){
      const v=sim.get(a);
      $('#simInfo').innerHTML=typeof v==='number'?`<b>${a}</b>: <code>${raw}</code> → hasil <b>${fmt(v)}</b> ✅`:`Rumus di ${a} belum benar (${v}). Cek penulisan dan range-nya.`;
      if(typeof v==='number')mark('sim');
    }
  });
}
initSim();
$('#simReset').onclick=()=>{initSim();$('#simInfo').textContent='Tabel dikembalikan.'};

/* ========== Kuis ========== */
const QZ=[
  {q:'Satu file Excel disebut…',o:['Worksheet','Workbook','Range','Cell'],a:1,e:'Workbook adalah file; worksheet adalah lembar di dalamnya.'},
  {q:'Perpotongan kolom C dan baris 4 bernama…',o:['4C','C4','C:4','CC4'],a:1,e:'Penulisan cell: huruf kolom lalu nomor baris, jadi C4.'},
  {q:'B2:B6 disebut…',o:['Range','Row','Workbook','Rumus'],a:0,e:'Range adalah kumpulan cell yang berurutan.'},
  {q:'Rumus di Excel selalu diawali tanda…',o:['#','@','=','$'],a:2,e:'Tanda = memberi tahu Excel bahwa isi cell adalah rumus.'},
  {q:'Rumus untuk menjumlahkan B2 sampai B6 adalah…',o:['=ADD(B2:B6)','=SUM(B2:B6)','=TOTAL(B2,B6)','=PLUS(B2:B6)'],a:1,e:'SUM menjumlahkan semua angka pada range.'},
  {q:'Nilai 80, 90, 70. Hasil =AVERAGE-nya adalah…',o:['70','80','90','240'],a:1,e:'(80+90+70)/3 = 80.'},
  {q:'Hasil =COUNT(B2:B6) jika B2:B6 berisi 5 angka…',o:['0','1','5','Jumlah nilainya'],a:2,e:'COUNT menghitung banyak cell berisi angka, bukan menjumlahkan.'},
  {q:'Ketik rumus untuk mencari nilai TERENDAH pada B2:B6 (lengkap, diawali =):',t:'=MIN(B2:B6)',e:'MIN mencari nilai terkecil. Jawaban: =MIN(B2:B6).'}
];
const answered={};let score=0;
$('#quiz').innerHTML=QZ.map((x,i)=>`<div class="card q" data-i="${i}"><b>${i+1}. ${x.q}</b>`+
  (x.t?`<div><input type="text" placeholder="=..."> <button class="btn chk">Periksa</button></div>`:x.o.map((o,j)=>`<button class="btn opt" data-j="${j}">${o}</button>`).join(''))+
  `<div class="fb"></div></div>`).join('');
function finishQ(i,ok,btn){
  if(answered[i])return;answered[i]=true;if(ok)score++;
  const c=$(`.q[data-i="${i}"]`);$('.fb',c).innerHTML=(ok?'✅ Benar! ':'❌ Belum tepat. ')+QZ[i].e;
  if(btn)btn.classList.add(ok?'right':'wrong');
  if(!ok&&!QZ[i].t)$$('.opt',c)[QZ[i].a].classList.add('right');
  $$('button,input',c).forEach(x=>x.disabled=true);
  const n=Object.keys(answered).length;
  $('#quizScore').innerHTML=`Terjawab ${n}/${QZ.length} • Skor: <b>${Math.round(score/QZ.length*100)}</b>`;
  if(n===QZ.length)mark('kuis');
}
$('#quiz').onclick=e=>{
  const c=e.target.closest('.q');if(!c)return;const i=+c.dataset.i;
  if(e.target.classList.contains('opt'))finishQ(i,+e.target.dataset.j===QZ[i].a,e.target);
  if(e.target.classList.contains('chk')){const v=$('input',c).value.replace(/\s/g,'').toUpperCase();finishQ(i,v===QZ[i].t,null)}
};

/* ========== Evaluasi akhir ========== */
const EV={...DATA,E1:'Rata-rata',A8:'UAS tertinggi',A9:'UAS terendah',A10:'Jumlah siswa',A11:'Total Tugas'};
const evEditable=a=>/^E[2-6]$/.test(a)||/^B(8|9|10|11)$/.test(a);
const ev=Sheet($('#evalSheet'),[1,2,3,4,5,6,7,8,9,10,11],'ABCDE'.split(''),EV,evEditable);
$('#evalCheck').onclick=()=>{
  const exp={};
  for(let r=2;r<=6;r++)exp['E'+r]=(DATA['B'+r]+DATA['C'+r]+DATA['D'+r])/3;
  const D=[2,3,4,5,6].map(r=>DATA['D'+r]),B=[2,3,4,5,6].map(r=>DATA['B'+r]);
  exp.B8=Math.max(...D);exp.B9=Math.min(...D);exp.B10=B.length;exp.B11=B.reduce((x,y)=>x+y,0);
  const hint={B8:'MAX',B9:'MIN',B10:'COUNT',B11:'SUM'};let ok=0,msg=[];
  Object.keys(exp).forEach(a=>{
    const raw=String(ev.M[a]??''),isF=raw[0]==='=',v=ev.get(a),good=isF&&typeof v==='number'&&Math.abs(v-exp[a])<0.01;
    $(`[data-a="${a}"]`,$('#evalSheet')).className=good?'ok':'bad';
    if(good)ok++;else msg.push(`❌ ${a}: ${!raw?'belum diisi':!isF?'gunakan rumus, bukan angka manual':'hasil belum tepat'}${hint[a]?` (petunjuk: ${hint[a]})`:''}`);
  });
  $('#evalInfo').innerHTML=`Benar <b>${ok}/${Object.keys(exp).length}</b> • Skor: <b>${Math.round(ok/Object.keys(exp).length*100)}</b><br>`+(msg.join('<br>')||'🎉 Hebat! Semua rumus tepat.');
  if(ok===Object.keys(exp).length)mark('eval');
};

renderProg();
