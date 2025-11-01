/* ============ أدوات مساعدة ============ */
const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);

const bioFile = $('#bioFile');
const manFile = $('#manFile');
const bioHint = $('#bioHint');
const manHint = $('#manHint');
const tbody = $('#tbody');
const q = $('#q');
const dlBtn = $('#dlBtn');

const pills = $$('.pill');
const nTotal = $('#nTotal');
const nOk = $('#nOk');
const nBad = $('#nBad');
const nMiss = $('#nMiss');

let bio = [];   // [{code,name,g,r}]
let man = [];
let merged = []; // صفوف العرض + التصدير
let activeFilter = 'all';
let query = '';

/* قراءة أي من XLSX/CSV */
async function readExcel(file){
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, {type:'array', cellDates:true});
  const sh = wb.Sheets[wb.SheetNames[0]];
  let rows = XLSX.utils.sheet_to_json(sh, {header:1, defval:'', raw:true});

  // ابحث عن الأعمدة (كود/اسم/غ/ر) بشكل مرن
  // يدعم الترتيب الشائع لديك (B=الكود, C=الاسم, D=غ, E=ر) وغير ذلك
  const header = rows[0].map(x=>String(x).trim());
  // جرّب أكثر من احتمال
  const idx = {
    code: header.findIndex(h => /كود|الكو?د|code/i.test(h)) !== -1 ? header.findIndex(h => /كود|الكو?د|code/i.test(h)) : 0,
    name: header.findIndex(h => /الاسم|name/i.test(h)) !== -1 ? header.findIndex(h => /الاسم|name/i.test(h)) : 1,
    g   : header.findIndex(h => /^غ|غياب|غيابات|g$/i.test(h)) !== -1 ? header.findIndex(h => /^غ|غياب|غيابات|g$/i.test(h)) : 2,
    r   : header.findIndex(h => /^ر|راتب|r$/i.test(h)) !== -1 ? header.findIndex(h => /^ر|راتب|r$/i.test(h)) : 3,
  };

  // إن كانت الصفوف بدون عناوين (كما في ملفاتك) نتجاهل الصف الأول ونقرأ أعمدة B..E
  const start = /كود|الاسم|غ|ر|code|name|g|r/i.test(header.join('')) ? 1 : 0;

  const out = [];
  rows.slice(start).forEach(row=>{
    const code = String(row[idx.code] ?? '').toString().trim();
    const name = String(row[idx.name] ?? '').toString().trim();
    const g = Number(row[idx.g] ?? 0) || 0;
    const r = Number(row[idx.r] ?? 0) || 0;
    if (!code && !name) return;
    out.push({ code, name, g, r });
  });
  return out;
}

/* تطبيع الاسم والكود */
function normalizeRow(x){
  return {
    code: String(x.code ?? '').replace(/[^\d]/g,'').trim(),
    name: String(x.name ?? '').replace(/\s+/g,' ').trim(),
    g: Number(x.g ?? 0) || 0,
    r: Number(x.r ?? 0) || 0,
  };
}

/* دمج ومقارنة */
function buildMerged(){
  const mapMan = new Map(man.map(m => [m.code, m]));
  const codes = new Set([...bio.map(b=>b.code), ...man.map(m=>m.code)]);

  merged = [];
  for (const code of codes){
    const b = bio.find(x=>x.code===code) || null;
    const m = mapMan.get(code) || null;

    const b_g = b?.g ?? null, b_r = b?.r ?? null;
    const m_g = m?.g ?? null, m_r = m?.r ?? null;

    // الحالات
    const gStatus = (b_g==null || m_g==null) ? 'miss' : (b_g===m_g ? 'ok' : 'bad');
    const rStatus = (b_r==null || m_r==null) ? 'miss' : (b_r===m_r ? 'ok' : 'bad');

    // الملاحظة: تُظهر فقط عند "مخالف"
    const noteParts = [];
    if (gStatus==='bad') noteParts.push('فرق في غ');
    if (rStatus==='bad') noteParts.push('فرق في ر');
    const note = noteParts.join(' + ');

    merged.push({
      code_b: b?.code ?? '',
      name_b: b?.name ?? '',
      g_b: b_g ?? 0,
      r_b: b_r ?? 0,
      code_m: m?.code ?? '',
      name_m: m?.name ?? '',
      g_m: m_g ?? 0,
      r_m: m_r ?? 0,
      g_status: gStatus,   // ok | bad | miss
      r_status: rStatus,
      note,
    });
  }

  // فرز بالكود تصاعدي (أرقام)
  merged.sort((a,b)=> Number(a.code_b || a.code_m) - Number(b.code_b || b.code_m));
}

/* عدّادات */
function updateCounters(list){
  nTotal.textContent = list.length;
  nOk.textContent = list.filter(x=>x.g_status==='ok' && x.r_status==='ok').length;
  nBad.textContent = list.filter(x=>x.g_status==='bad' || x.r_status==='bad').length;
  nMiss.textContent = list.filter(x=>x.g_status==='miss' || x.r_status==='miss').length;
}

/* تصفية حسب البحث والفلاتر */
function getView(){
  let list = merged;

  if (query){
    const qlower = query.toLowerCase();
    list = list.filter(r=>{
      return String(r.code_b).includes(query) ||
             String(r.code_m).includes(query) ||
             r.name_b.toLowerCase().includes(qlower) ||
             r.name_m.toLowerCase().includes(qlower);
    });
  }

  if (activeFilter==='ok'){
    list = list.filter(x=>x.g_status==='ok' && x.r_status==='ok');
  }else if (activeFilter==='bad'){
    list = list.filter(x=>x.g_status==='bad' || x.r_status==='bad');
  }else if (activeFilter==='miss'){
    list = list.filter(x=>x.g_status==='miss' || x.r_status==='miss');
  }
  return list;
}

/* عرض الجدول */
function render(){
  const view = getView();
  updateCounters(view);
  const rowsHtml = view.map((r,idx)=>{
    const gBadge = `<span class="badge ${r.g_status}">${label(r.g_status)}</span>`;
    const rBadge = `<span class="badge ${r.r_status}">${label(r.r_status)}</span>`;
    return `
      <tr>
        <td>${idx+1}</td>
        <td>${safe(r.code_b)}</td>
        <td>${safe(r.name_b)}</td>
        <td>${safe(r.g_b)}</td>
        <td>${safe(r.r_b)}</td>
        <td>${safe(r.code_m)}</td>
        <td>${safe(r.name_m)}</td>
        <td>${safe(r.g_m)}</td>
        <td>${safe(r.r_m)}</td>
        <td>${gBadge}</td>
        <td>${rBadge}</td>
        <td>${r.note || ''}</td>
      </tr>
    `;
  }).join('');
  tbody.innerHTML = rowsHtml || `<tr><td colspan="12">لا توجد بيانات لعرضها…</td></tr>`;
}

function label(st){
  if (st==='ok') return 'مطابق';
  if (st==='bad') return 'مخالف';
  return 'ناقص';
}
const safe = v => (v==null ? '' : String(v));

/* أحداث الفلاتر */
pills.forEach(p=>{
  p.addEventListener('click', ()=>{
    pills.forEach(x=>x.classList.remove('active'));
    p.classList.add('active');
    activeFilter = p.dataset.filter;
    render();
  });
});
q.addEventListener('input', ()=>{
  query = q.value.trim();
  render();
});

/* رفع الملفات */
bioFile.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  bioHint.textContent = "…جاري القراءة";
  try{
    const rows = await readExcel(f);
    bio = rows.map(normalizeRow).filter(x=>x.code || x.name);
    bioHint.textContent = `تم رفع: ${f.name} — ${bio.length} صفًا`;
    buildMerged(); render();
  }catch(err){
    alert("تعذّر قراءة ملف البصمة.\n" + err.message);
    bioHint.textContent = "فشل التحميل.";
  }
});

manFile.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  manHint.textContent = "…جاري القراءة";
  try{
    const rows = await readExcel(f);
    man = rows.map(normalizeRow).filter(x=>x.code || x.name);
    manHint.textContent = `تم رفع: ${f.name} — ${man.length} صفًا`;
    buildMerged(); render();
  }catch(err){
    alert("تعذّر قراءة الملف اليدوي.\n" + err.message);
    manHint.textContent = "فشل التحميل.";
  }
});

/* تصدير XLSX بنفس ترتيب الأعمدة المطلوب */
dlBtn.addEventListener('click', ()=>{
  const view = getView();
  const header = [
    "م","الكود (بصمة)","الاسم (بصمة)","غ (بصمة)","ر (بصمة)",
    "الكود (يدوي)","الاسم (يدوي)","غ (يدوي)","ر (يدوي)",
    "نتيجة غ","نتيجة ر","الملاحظة"
  ];

  const rows = view.map((r,i)=>[
    i+1, r.code_b, r.name_b, r.g_b, r.r_b,
    r.code_m, r.name_m, r.g_m, r.r_m,
    label(r.g_status), label(r.r_status), r.note || ""
  ]);

  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "نتائج المطابقة");
  XLSX.writeFile(wb, "canary_monthly_compare.xlsx");
});
