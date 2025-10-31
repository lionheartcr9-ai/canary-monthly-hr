// Helpers
const dlg = (msg)=>{ const d=document.getElementById('dlg'); document.getElementById('dlgMsg').textContent=msg; d.showModal(); };

const state = {
  bio: null, // {rows: []}
  man: null,
  merged: [],
};

function readFile(file){
  return new Promise((resolve,reject)=>{
    const reader = new FileReader();
    reader.onload = (e)=>{
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
        resolve(rows);
      }catch(err){ reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizeRow(r){
  // Expect Arabic headers: الكود | الاسم | غ | ر
  const code = String(r['الكود'] ?? r['code'] ?? r['Code'] ?? "").trim();
  const name = String(r['الاسم'] ?? r['اسم الموظف'] ?? r['name'] ?? "").trim();
  const g = parseFloat(String(r['غ'] ?? r['غياب'] ?? 0).toString().replace(',','.')) || 0;
  const rdays = parseFloat(String(r['ر'] ?? r['اجازة'] ?? 0).toString().replace(',','.')) || 0;
  return {code, name, g, r: rdays};
}

function keyOf(obj){ return `${obj.code}__${obj.name}`; }

function compare(){
  if(!state.bio || !state.man){ return; }
  const insightDiff = Math.max(1, parseInt(document.getElementById('insightDiff').value)||2);

  const bmap = new Map();
  state.bio.forEach(r=>{
    const n = normalizeRow(r);
    if(n.code||n.name) bmap.set(keyOf(n), n);
  });

  const mmap = new Map();
  state.man.forEach(r=>{
    const n = normalizeRow(r);
    if(n.code||n.name) mmap.set(keyOf(n), n);
  });

  const keys = new Set([...bmap.keys(), ...mmap.keys()]);
  const rows = [];

  keys.forEach((k, idx)=>{
    const b = bmap.get(k);
    const m = mmap.get(k);

    const out = {
      index: idx+1,
      b_code: b?.code ?? "", b_name: b?.name ?? "", b_g: b?.g ?? "", b_r: b?.r ?? "",
      m_code: m?.code ?? "", m_name: m?.name ?? "", m_g: m?.g ?? "", m_r: m?.r ?? "",
      res_g: "", res_r: "", note: ""
    };

    if(!b || !m){
      out.res_g = "بيانات ناقصة"; out.res_r = "بيانات ناقصة";
      out.note = "بيانات ناقصة في أحد الملفين.";
      rows.push(out); return;
    }

    // Results for G
    if(b.g === m.g){
      out.res_g = "مطابق غ";
    } else {
      out.res_g = "مخالف غ";
      if(b.g > m.g){
        out.note += "تحقق من الإدخال اليدوي غ؛ قد تكون قيمة اليوم خاطئة. ";
      } else {
        const diff = (m.g - b.g).toFixed(2);
        out.note += `بعد التأكد يتم عمل استيفاء غ بالفارق (${diff}). `;
      }
      if(Math.abs(b.g - m.g) >= insightDiff){
        out.note += `⚠ فرق كبير في غ (≥ ${insightDiff} يوم). `;
      }
    }

    // Results for R
    if((b.r ?? 0) === (m.r ?? 0)){
      out.res_r = "مطابق ر";
    } else {
      out.res_r = "مخالف ر";
      if((b.r ?? 0) > (m.r ?? 0)){
        out.note += "تحقق من الإدخال اليدوي ر؛ قد تكون R لم تُسجل. ";
      } else {
        const diffR = ((m.r ?? 0) - (b.r ?? 0)).toFixed(2);
        out.note += `بعد التأكد يتم عمل ر بالفارق (${diffR}). `;
      }
      if(Math.abs((b.r ?? 0) - (m.r ?? 0)) >= insightDiff){
        out.note += `⚠ فرق كبير في ر (≥ ${insightDiff} يوم). `;
      }
    }

    rows.push(out);
  });

  // Sort by code asc (biometric code), numeric if possible
  rows.sort((a,b)=>{
    const na = parseFloat(a.b_code)||0;
    const nb = parseFloat(b.b_code)||0;
    if(na!==nb) return na-nb;
    return String(a.b_code).localeCompare(String(b.b_code));
  });

  state.merged = rows;
  render();
}

function render(){
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";
  let ok=0,bad=0,miss=0;

  const q = (document.getElementById('searchBox').value||"").trim();
  const re = q? new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'i'): null;

  let i=0;
  for(const r of state.merged){
    const textline = `${r.b_code} ${r.b_name} ${r.m_code} ${r.m_name}`;
    if(re && !re.test(textline)) continue;
    i++;

    let cls="";
    if(r.res_g==="بيانات ناقصة" || r.res_r==="بيانات ناقصة"){ cls="miss"; miss++; }
    else if(r.res_g.startsWith("مخالف") || r.res_r.startsWith("مخالف")){ cls="bad"; bad++; }
    else { cls="ok"; ok++; }

    const tr = document.createElement('tr');
    tr.className = cls;
    tr.innerHTML = `
      <td class="center">${i}</td>
      <td class="center">${r.b_code}</td>
      <td>${r.b_name}</td>
      <td class="center">${r.b_g}</td>
      <td class="center">${r.b_r}</td>
      <td class="center">${r.m_code}</td>
      <td>${r.m_name}</td>
      <td class="center">${r.m_g}</td>
      <td class="center">${r.m_r}</td>
      <td class="center result">${r.res_g}</td>
      <td class="center result">${r.res_r}</td>
      <td>${r.note}</td>
    `;
    tbody.appendChild(tr);
  }

  document.getElementById('countOk').textContent = `مطابق ${ok}`;
  document.getElementById('countBad').textContent = `مخالف ${bad}`;
  document.getElementById('countMiss').textContent = `ناقص ${miss}`;
  document.getElementById('countLoad').textContent = `تم تحميل: ${(state.bio?.length||0)} / ${(state.man?.length||0)} 👥`;

  document.getElementById('btnExport').disabled = state.merged.length===0;
}

async function onPick(which, input, stat){
  try{
    const f = input.files[0];
    if(!f) return;
    const rows = await readFile(f);
    if(which==="bio") state.bio = rows;
    else state.man = rows;
    stat.textContent = `تم تحميل: ${rows.length} صف`;
    compare();
  }catch(err){
    console.error(err);
    if(typeof XLSX === 'undefined'){
      dlg("تعذّر تحميل مكتبة XLSX من الإنترنت. جرب فتح الصفحة عبر Vercel/GitHub Pages ثم أعد التحميل.");
    }else{
      dlg("تعذّر قراءة هذا الملف. تأكد أن صف العناوين يحتوي: الكود | الاسم | غ | ر.");
    }
  }
}

function exportXLSX(){
  const data = [["م","الكود (بصمة)","الاسم (بصمة)","غ (بصمة)","ر (بصمة)","الكود (يدوي)","الاسم (يدوي)","غ (يدوي)","ر (يدوي)","نتيجة غ","نتيجة ر","الملاحظة"]];
  for(const r of state.merged){
    data.push([r.index, r.b_code, r.b_name, r.b_g, r.b_r, r.m_code, r.m_name, r.m_g, r.m_r, r.res_g, r.res_r, r.note]);
  }
  const ws = XLSX.utils.aoa_to_sheet(data);
  const range = XLSX.utils.decode_range(ws['!ref']);

  for(let C=0; C<=11; C++){
    const addr = XLSX.utils.encode_cell({r:0,c:C});
    ws[addr].s = { fill:{fgColor:{rgb:"103A6B"}}, font:{bold:true,color:{rgb:"FFFFFF"}}, alignment:{horizontal:"center",vertical:"center"} };
  }

  for(let R=1; R<=range.e.r; R++){
    const resG = ws[XLSX.utils.encode_cell({r:R,c:9})]?.v || "";
    const resR = ws[XLSX.utils.encode_cell({r:R,c:10})]?.v || "";
    let fill = {fgColor:{rgb:"0E3523"}};
    if(resG.includes("بيانات ناقصة") || resR.includes("بيانات ناقصة")) fill = {fgColor:{rgb:"2A2E38"}};
    else if(resG.includes("مخالف") || resR.includes("مخالف")) fill = {fgColor:{rgb:"3A0F15"}};
    [9,10].forEach(c=>{
      const cell = XLSX.utils.encode_cell({r:R,c});
      if(ws[cell]) ws[cell].s = {fill, font:{color:{rgb:"FFFFFF"}}, alignment:{horizontal:"center"}};
    });
    [0,1,3,4,5,7,8].forEach(c=>{
      const cell = XLSX.utils.encode_cell({r:R,c});
      if(ws[cell]) ws[cell].s = {alignment:{horizontal:"center"}};
    });
  }

  ws['!cols'] = [
    {wch:4},{wch:10},{wch:26},{wch:8},{wch:8},{wch:10},{wch:26},{wch:8},{wch:8},{wch:10},{wch:10},{wch:40}
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "نتائج المطابقة");
  XLSX.writeFile(wb, "canary_monthly_compare.xlsx");
}

window.addEventListener('DOMContentLoaded', ()=>{
  const fileBio = document.getElementById('fileBio');
  const fileMan = document.getElementById('fileMan');
  document.getElementById('btnExport').addEventListener('click', exportXLSX);
  document.getElementById('searchBox').addEventListener('input', render);
  document.getElementById('insightDiff').addEventListener('change', render);

  fileBio.addEventListener('change', ()=>onPick('bio', fileBio, document.getElementById('statBio')));
  fileMan.addEventListener('change', ()=>onPick('man', fileMan, document.getElementById('statMan')));
});
