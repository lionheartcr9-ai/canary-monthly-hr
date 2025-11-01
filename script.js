/* =========================
   HR Monthly Matcher (XLSX)
   — يعمل على رؤوس أعمدة ثابتة فقط —
   الكود | الاسم | غ | ر
   ========================= */

// عناصر DOM
const bioFile  = document.getElementById('bioFile');
const manFile  = document.getElementById('manFile');
const bioHint  = document.getElementById('bioHint');
const manHint  = document.getElementById('manHint');
const q        = document.getElementById('q');
const btnExport= document.getElementById('btnExport');
const tbody    = document.getElementById('tbody');

const cAll = document.getElementById('cAll');
const cOk  = document.getElementById('cOk');
const cBad = document.getElementById('cBad');
const cMiss= document.getElementById('cMiss');

// الحالة
let bioRows = [];     // البصمة (مصفوفة كائنات {الكود, الاسم, غ, ر})
let manRows = [];     // اليدوي  ( ")
let merged  = [];     // نتيجة الدمج/المقارنة

// ————— أدوات مساعدة —————

// تثبيت قبول XLSX فقط
function ensureXlsx(file) {
  const ok = file && /\.xlsx$/i.test(file.name);
  if (!ok) throw new Error('الملف يجب أن يكون بصيغة .xlsx');
  return file;
}

// قراءة أول ورقة من XLSX إلى مصفوفة كائنات
async function readXlsx(file) {
  ensureXlsx(file);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error('لا توجد ورقة عمل في الملف');

  // التحويل إلى JSON مع اعتبار الصف الأول رؤوس أعمدة
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

  // التحقق من الرؤوس المطلوبة بالضبط
  const required = ['الكود', 'الاسم', 'غ', 'ر'];
  const header = rows.length ? Object.keys(rows[0]) : [];
  const hasAll = required.every(k => header.includes(k));
  if (!hasAll) {
    throw new Error('يجب أن تكون عناوين الأعمدة بالضبط: الكود | الاسم | غ | ر');
  }

  // تنظيف وتحويل القيم الرقمية
  return rows.map(r => ({
    الكود : String(r['الكود']).trim(),
    الاسم : String(r['الاسم']).trim(),
    غ     : toNumber(r['غ']),
    ر     : toNumber(r['ر']),
  }));
}

function toNumber(v){
  if (typeof v === 'number') return v;
  if (typeof v === 'string'){
    const t = v.replace(/[^\d.\-]/g,'').trim();
    if (t === '') return 0;
    const n = Number(t);
    return isFinite(n) ? n : 0;
  }
  return 0;
}

// تسهيل مقارنة الأسماء (إزالة فراغات متكررة + توحيد بعض الحروف الشائعة)
function normName(s){
  return s
    .replace(/[اأإآ]/g,'ا')
    .replace(/ى/g,'ي')
    .replace(/ة/g,'ه')
    .replace(/\s+/g,' ')
    .trim();
}

// مفتاح المطابقة (يشترط تطابق الكود + الاسم)
function makeKey(code, name){
  return `${String(code).trim()}|${normName(String(name))}`;
}

// مقارنة غ/ر وإنتاج نتيجة وملاحظة
function compareValues(gb, rb, gm, rm){
  // نواتج غ
  let resG = '';
  let resR = '';
  let noteG = '';
  let noteR = '';

  if (gb === gm) {
    resG = 'مطابق';
  } else if (gb > gm) {
    resG = 'مخالف';
    noteG = 'يتم التأكد من صحة الادخال اليدوي غ';
  } else { // gb < gm
    resG = 'مخالف';
    noteG = `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق [${(gm - gb)}]`;
  }

  // نواتج ر
  if (rb === rm) {
    resR = 'مطابق';
  } else if (rb > rm) {
    resR = 'مخالف';
    noteR = 'يتم التأكد من صحة الادخال اليدوي ر';
  } else { // rb < rm
    resR = 'مخالف';
    noteR = `بعد التأكد من الادخال يتم عمل ر بالفارق [${(rm - rb)}]`;
  }

  // الملاحظة المجمّعة
  let note = '';
  if (resG === 'مطابق' && resR === 'مطابق') {
    note = '';
  } else {
    note = [noteG, noteR].filter(Boolean).join(' — ');
  }

  return { resG, resR, note };
}

// دمج وبناء النتيجة
function mergeData(){
  const mapBio = new Map();
  bioRows.forEach(r => mapBio.set(makeKey(r.الكود, r.الاسم), r));

  const mapMan = new Map();
  manRows.forEach(r => mapMan.set(makeKey(r.الكود, r.الاسم), r));

  const keys = new Set([...mapBio.keys(), ...mapMan.keys()]);
  const out = [];

  for (const k of keys){
    const b = mapBio.get(k);
    const m = mapMan.get(k);

    // استخراج الكود/الاسم من المفتاح
    const [codeKey, nameKey] = k.split('|');

    if (!b || !m){
      out.push({
        codeB: b?.الكود || '',
        nameB: b?.الاسم || '',
        gB:    b?.غ ?? '',
        rB:    b?.ر ?? '',
        codeM: m?.الكود || '',
        nameM: m?.الاسم || '',
        gM:    m?.غ ?? '',
        rM:    m?.ر ?? '',
        resG:  'ناقص',
        resR:  'ناقص',
        note:  'بيانات ناقصة'
      });
      continue;
    }

    const { resG, resR, note } = compareValues(b.غ, b.ر, m.غ, m.ر);
    out.push({
      codeB: b.الكود, nameB: b.الاسم, gB: b.غ, rB: b.ر,
      codeM: m.الكود, nameM: m.الاسم, gM: m.غ, rM: m.ر,
      resG, resR, note
    });
  }

  // الترتيب تصاعدياً حسب الكود (البصمة ثم اليدوي احتياطاً)
  out.sort((a,b)=>{
    const x = Number(a.codeB || a.codeM) || 0;
    const y = Number(b.codeB || b.codeM) || 0;
    return x - y;
  });

  merged = out;
}

// عرض النتائج في الجدول + الإحصاءات + التصفية
function render(){
  const term = q.value.trim();
  const rows = merged.filter(r=>{
    if (!term) return true;
    const t = term.toLowerCase();
    return (
      String(r.codeB).includes(term) ||
      String(r.codeM).includes(term) ||
      String(r.nameB).toLowerCase().includes(t) ||
      String(r.nameM).toLowerCase().includes(t)
    );
  });

  // إحصاءات
  cAll.textContent  = `${rows.length} إجمالي`;
  cOk.textContent   = `${rows.filter(r=> r.resG==='مطابق' && r.resR==='مطابق').length} مطابق`;
  cBad.textContent  = `${rows.filter(r=> r.resG==='مخالف' || r.resR==='مخالف').length} مخالف`;
  cMiss.textContent = `${rows.filter(r=> r.resG==='ناقص'   || r.resR==='ناقص').length} ناقص/غير مكتمل`;

  // بناء الصفوف
  tbody.innerHTML = '';
  rows.forEach((r,idx)=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx+1}</td>

      <td>${safe(r.codeB)}</td>
      <td class="wide">${safe(r.nameB)}</td>
      <td>${safe(r.gB)}</td>
      <td>${safe(r.rB)}</td>

      <td>${safe(r.codeM)}</td>
      <td class="wide">${safe(r.nameM)}</td>
      <td>${safe(r.gM)}</td>
      <td>${safe(r.rM)}</td>

      <td>${r.resG}</td>
      <td>${r.resR}</td>
      <td class="wider">${safe(r.note)}</td>
    `;
    tbody.appendChild(tr);
  });

  btnExport.disabled = rows.length === 0;
}

function safe(v){ return (v===undefined || v===null) ? '' : v; }

// تصدير XLSX (بدون أي تنسيق/تلوين)
function exportXlsx(){
  const aoa = [];
  aoa.push([
    'م',
    'الكود (بصمة)','الاسم (بصمة)','غ (بصمة)','ر (بصمة)',
    'الكود (يدوي)','الاسم (يدوي)','غ (يدوي)','ر (يدوي)',
    'نتيجة غ','نتيجة ر','الملاحظة'
  ]);

  merged.forEach((r,i)=>{
    aoa.push([
      i+1,
      r.codeB, r.nameB, r.gB, r.rB,
      r.codeM, r.nameM, r.gM, r.rM,
      r.resG,  r.resR,  r.note
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'نتيجة المطابقة');
  XLSX.writeFile(wb, 'canary_monthly_compare.xlsx');
}

// ————— الأحداث —————
bioFile.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  bioHint.textContent = '...جاري القراءة';
  try{
    bioRows = await readXlsx(f);
    bioHint.textContent = `تم رفع: ${f.name} — ${bioRows.length} صفًا`;
    if (manRows.length) { mergeData(); render(); }
  }catch(err){
    bioRows = [];
    bioHint.textContent = 'فشل التحميل.';
    alert(err.message);
  }
});

manFile.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  manHint.textContent = '...جاري القراءة';
  try{
    manRows = await readXlsx(f);
    manHint.textContent = `تم رفع: ${f.name} — ${manRows.length} صفًا`;
    if (bioRows.length) { mergeData(); render(); }
  }catch(err){
    manRows = [];
    manHint.textContent = 'فشل التحميل.';
    alert(err.message);
  }
});

q.addEventListener('input', render);
btnExport.addEventListener('click', exportXlsx);
