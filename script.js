/* ===================== أدوات مساعدة ===================== */
const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);

const bioFile = $('#bioFile');
const manFile = $('#manFile');
const bioHint = $('#bioHint');
const manHint = $('#manHint');
const tbody = $('#tbody');
const rowsCnt = $('#rowsCnt');
const okCnt = $('#okCnt'), badCnt = $('#badCnt'), missCnt = $('#missCnt'), totalCnt = $('#totalCnt');
const dlBtn = $('#dlBtn');
const q = $('#q');

let bio = [];   // [{code,name,g,r}]
let man = [];   // [{code,name,g,r}]
let merged = []; // صفوف العرض النهائي
let currentFilter = 'all';

/* تطبيع النص العربي لتقليل فروق الهجاء */
function normalizeName(s=''){
  return String(s)
    .replace(/\s+/g,' ')
    .replace(/[إأآا]/g,'ا')
    .replace(/ى/g,'ي')
    .replace(/ؤ/g,'و')
    .replace(/ئ/g,'ي')
    .replace(/ة/g,'ه')
    .trim();
}
function n2(x){ return isFinite(+x) ? +x : 0 }

/* قراءة CSV بسيطة */
async function readCSV(file){
  const text = await file.text();
  const lines = text.replace(/\r/g,'').split('\n').filter(Boolean);
  const rows = [];
  for (const ln of lines){
    const parts = ln.split(',').map(c=>c.trim());
    rows.push(parts);
  }
  return rows;
}

/* قراءة أي ملف (XLSX أو CSV) وإرجاع صفوف مصفوفة */
async function readAny(file){
  const ext = (file.name.split('.').pop()||'').toLowerCase();
  if (ext === 'csv'){
    return await readCSV(file);
  }
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, {cellDates:false, cellText:false});
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:''});
}

/* محاولات لاكتشاف الأعمدة */
function toObjects(rows){
  if (!rows.length) return [];
  // ابحث عن أول أربعة أعمدة (كود/اسم/غ/ر) إن وُجدت رؤوس
  const head = rows[0].map(x=>String(x).trim());
  const findIdx = (cands)=> head.findIndex(h => cands.some(c=>h.includes(c)));
  let iCode = findIdx(['الكود','الكود (بصمة)','الكود (يدوي)','code']);
  let iName = findIdx(['الاسم','الاسم (بصمة)','الاسم (يدوي)','name']);
  let iG = findIdx(['غ','غياب','G']);
  let iR = findIdx(['ر','راحة','R']);

  // إن لم نجد رؤوس، نفترض الترتيب: كود, اسم, غ, ر
  const startRow = (iCode>-1||iName>-1||iG>-1||iR>-1) ? 1 : 0;
  if (startRow===0){ iCode=0; iName=1; iG=2; iR=3; }

  const out=[];
  for (let r=startRow;r<rows.length;r++){
    const row = rows[r]||[];
    const code = String(row[iCode] ?? '').trim();
    const name = String(row[iName] ?? '').trim();
    const g = n2(row[iG]);
    const rr = n2(row[iR]);
    if (!code && !name && !g && !rr) continue;
    out.push({code, name, g, r: rr});
  }
  return out;
}

/* دمج ومقارنة وفق القواعد المطلوبة */
function mergeCompare(){
  const mapBio = new Map();
  for (const b of bio){
    const key = `${b.code}|${normalizeName(b.name)}`;
    mapBio.set(key, b);
  }

  const keys = new Set();
  const out = [];

  // مر عبر اليدوي أولاً (نضمن وجود صف لكل موظف موجود يدوياً)
  for (const m of man){
    const key = `${m.code}|${normalizeName(m.name)}`;
    keys.add(key);
    const b = mapBio.get(key) || null;

    // نقص بيانات؟
    const miss =
      !b || !b.code || !b.name || !Number.isFinite(+b.g) || !Number.isFinite(+b.r) ||
      !m.code || !m.name || !Number.isFinite(+m.g) || !Number.isFinite(+m.r);

    let resG='ناقص', resR='ناقص', note='';
    if (!miss){
      // المقارنة الدقيقة
      if (b.g === m.g){ resG='مطابق'; }
      else if (b.g > m.g){ resG='مخالف'; note = 'يتم التأكد من صحة الادخال اليدوي غ'; }
      else if (b.g < m.g){ resG='مخالف'; note = `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق ${m.g - b.g}`; }

      if (b.r === m.r){ resR='مطابق'; }
      else if (b.r > m.r){ 
        resR='مخالف'; 
        note = note ? note+' — ' : '';
        note += 'يتم التأكد من صحة الادخال اليدوي ر';
      }
      else if (b.r < m.r){
        resR='مخالف';
        note = note ? note+' — ' : '';
        note += `بعد التأكد يتم عمل ر بالفارق ${m.r - b.r}`;
      }
    }

    out.push({
      code_b: b?.code ?? '', name_b: b?.name ?? '', g_b: b?.g ?? 0, r_b: b?.r ?? 0,
      code_m: m.code, name_m: m.name, g_m: m.g, r_m: m.r,
      res_g: resG, res_r: resR,
      note: miss ? 'بيانات ناقصة' : (note || (resG==='مطابق' && resR==='مطابق' ? '' : 'فرق في غ/ر'))
    });
  }

  // أي عنصر في البصمة غير موجود في اليدوي (نضيفه كسطر ناقص)
  for (const b of bio){
    const key = `${b.code}|${normalizeName(b.name)}`;
    if (keys.has(key)) continue;
    out.push({
      code_b: b.code, name_b: b.name, g_b: b.g, r_b: b.r,
      code_m: '', name_m: '', g_m: 0, r_m: 0,
      res_g: 'ناقص', res_r: 'ناقص', note:'بيانات ناقصة'
    });
  }

  merged = out;
}

/* رسم الجدول والإحصاءات */
function render(){
  // فلترة وبحث
  const term = normalizeName(q.value||'');
  const filtered = merged.filter((r)=>{
    const inSearch = !term ||
      normalizeName(r.name_b).includes(term) ||
      normalizeName(r.name_m).includes(term) ||
      String(r.code_b).includes(term) || String(r.code_m).includes(term);

    const byFilter = (currentFilter==='all') ||
      (currentFilter==='ok' && r.res_g==='مطابق' && r.res_r==='مطابق') ||
      (currentFilter==='bad' && (r.res_g==='مخالف' || r.res_r==='مخالف')) ||
      (currentFilter==='miss' && (r.res_g==='ناقص' || r.res_r==='ناقص'));

    return inSearch && byFilter;
  });

  // إحصاءات عامة على كامل البيانات
  const allOk = merged.filter(r=>r.res_g==='مطابق' && r.res_r==='مطابق').length;
  const allBad = merged.filter(r=>r.res_g==='مخالف' || r.res_r==='مخالف').length;
  const allMiss = merged.filter(r=>r.res_g==='ناقص' || r.res_r==='ناقص').length;

  okCnt.textContent = allOk;
  badCnt.textContent = allBad;
  missCnt.textContent = allMiss;
  totalCnt.textContent = merged.length;

  rowsCnt.textContent = filtered.length;

  // بناء الصفوف
  const frag = document.createDocumentFragment();
  filtered.forEach((r,idx)=>{
    const tr = document.createElement('tr');

    function td(val, cls=''){ const t=document.createElement('td'); t.textContent=(val??''); if(cls) t.className=cls; return t; }
    function statusCls(s){ return s==='مطابق'?'ok-cell':(s==='مخالف'?'bad-cell':'miss-cell') }

    tr.appendChild(td(idx+1));                           // م
    tr.appendChild(td(r.code_b));
    tr.appendChild(td(r.name_b));
    tr.appendChild(td(r.g_b));
    tr.appendChild(td(r.r_b));
    tr.appendChild(td(r.code_m));
    tr.appendChild(td(r.name_m));
    tr.appendChild(td(r.g_m));
    tr.appendChild(td(r.r_m));
    tr.appendChild(td(r.res_g, statusCls(r.res_g)));     // نتيجة غ
    tr.appendChild(td(r.res_r, statusCls(r.res_r)));     // نتيجة ر
    tr.appendChild(td(r.note||''));                      // الملاحظة

    frag.appendChild(tr);
  });

  tbody.innerHTML='';
  tbody.appendChild(frag);

  // تفعيل/تعطيل التصدير
  dlBtn.disabled = merged.length===0;
}

/* تصدير XLSX بالتنسيق والترتيب المطلوب */
function exportXLSX(){
  const rows = [
    ['م','الكود (بصمة)','الاسم (بصمة)','غ (بصمة)','ر (بصمة)','الكود (يدوي)','الاسم (يدوي)','غ (يدوي)','ر (يدوي)','نتيجة غ','نتيجة ر','الملاحظة']
  ];

  merged.forEach((r,i)=>{
    rows.push([
      i+1, r.code_b, r.name_b, r.g_b, r.r_b,
      r.code_m, r.name_m, r.g_m, r.r_m,
      r.res_g, r.res_r, r.note||''
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'نتيجة المطابقة');
  XLSX.writeFile(wb, 'canary_monthly_compare.xlsx');
}

/* ===================== الأحداث ===================== */
async function handleUpload(which, file, hintEl){
  if(!file) return;
  hintEl.textContent = '…جارِ القراءة';

  try{
    const raw = await readAny(file);
    const objs = toObjects(raw);
    const niceName = `${file.name} — ${objs.length} صفًا`;
    hintEl.textContent = `تم رفع: ${niceName}`;

    if (which==='bio') bio = objs;
    else man = objs;

    if (bio.length || man.length){
      mergeCompare();
      render();
    }
  }catch(err){
    console.error(err);
    hintEl.textContent = 'فشل التحميل.';
    alert('تعذّر قراءة الملف.\n' + err.message);
  }
}

bioFile.addEventListener('change', e => handleUpload('bio', e.target.files?.[0], bioHint));
manFile.addEventListener('change', e => handleUpload('man', e.target.files?.[0], manHint));

q.addEventListener('input', render);

$$('.chip').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    $$('.chip').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    render();
  });
});

dlBtn.addEventListener('click', exportXLSX);

/* بداية نظيفة */
render();
