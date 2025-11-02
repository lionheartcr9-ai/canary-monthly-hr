/* ========= DOM / Globals ========= */
let fpData = null, manualData = null, fullResults = [];

const fpInput        = document.getElementById('fpFile');
const manualInput    = document.getElementById('manualFile');
const fpNameSpan     = document.getElementById('fpName');
const manualNameSpan = document.getElementById('manualName');

const startBtn       = document.getElementById('startCompare');
const downloadBtn    = document.getElementById('downloadXlsx');

const statAllBtn     = document.getElementById('statAll');
const statMatchBtn   = document.getElementById('statMatch');
const statDiffBtn    = document.getElementById('statDiff');
const statMissingBtn = document.getElementById('statMissing');

const searchBox      = document.getElementById('searchBox');
const resultBody     = document.getElementById('resultBody');

/* ========= Utils ========= */
async function readXlsx(file){
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, {type:'array'});
  const sh  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sh, {defval:""});
}

function updateStartState(){ startBtn.disabled = !(fpData && manualData); }

const roundValue = (num)=>{
  if (num === "" || num === null || isNaN(num)) return 0;
  return Math.round((parseFloat(num) + Number.EPSILON) * 100) / 100;
};

// Arabic normalization
function normalizeArabic(str){
  if(!str) return "";
  return String(str)
    .replace(/[ًٌٍَُِّْـ]/g,"")
    .replace(/[\u200F\u200E]/g,"")
    .replace(/[إأآا]/g,"ا").replace(/ى/g,"ي").replace(/ة/g,"ه")
    .replace(/\s+/g," ").trim();
}

function tokenizeName(name){
  return normalizeArabic(name).replace(/[^ء-ي\s]/g,"").split(" ").filter(Boolean);
}

function diceSimilarity(A,B){
  const a=new Set(A), b=new Set(B); let inter=0;
  a.forEach(t=>{ if(b.has(t)) inter++; });
  const den=a.size+b.size; return den? (2*inter)/den : 0;
}

function namesClose(a,b,threshold=0.60){
  const ta=tokenizeName(a), tb=tokenizeName(b);
  if(!ta.length || !tb.length) return false;
  return diceSimilarity(ta,tb) >= threshold;
}

// mapRow with 31->30 for manual G
function mapRow(row, isManual=false){
  let g = roundValue(row["غ"] ?? row["غ (بصمة)"] ?? row["غ (يدوي)"] ?? 0);
  const r = roundValue(row["ر"] ?? row["ر (بصمة)"] ?? row["ر (يدوي)"] ?? 0);
  if(isManual && g === 31) g = 30;            // 👈 القاعدة الجديدة
  return {
    code: String(row["الكود"] ?? row["الكود (بصمة)"] ?? row["الكود (يدوي)"] ?? "").trim(),
    name: String(row["الاسم"] ?? row["الاسم (بصمة)"] ?? row["الاسم (يدوي)"] ?? "").trim(),
    g, r
  };
}

/* ========= Build / Render ========= */
function buildRow(idx, rec){
  const tr=document.createElement('tr');

  function td(text, cls){
    const cell=document.createElement('td');
    if (cls) cell.className=cls;
    cell.textContent=(text==null)?"":text;
    return cell;
  }

  tr.appendChild(td(idx+1));
  tr.appendChild(td(rec.code_fp));
  tr.appendChild(td(rec.name_fp));
  tr.appendChild(td(rec.g_fp));
  tr.appendChild(td(rec.r_fp));
  tr.appendChild(td(rec.code_m));
  tr.appendChild(td(rec.name_m));
  tr.appendChild(td(rec.g_m));
  tr.appendChild(td(rec.r_m));
  // تلوين نتيجتي غ/ر
  const clsG = rec.res_g==="مطابق" ? "status-match" : rec.res_g==="مخالف" ? "status-diff" : "status-missing";
  const clsR = rec.res_r==="مطابق" ? "status-match" : rec.res_r==="مخالف" ? "status-diff" : "status-missing";
  tr.appendChild(td(rec.res_g, clsG));
  tr.appendChild(td(rec.res_r, clsR));
  tr.appendChild(td(rec.note || "")); // لا نكتب "مطابق" كملاحظة
  return tr;
}

function renderTable(list){
  resultBody.innerHTML="";
  list.forEach((r,i)=> resultBody.appendChild(buildRow(i,r)));
}

function updateStats(){
  const all = fullResults.length;
  const match   = fullResults.filter(r=> r.res_g==="مطابق" && r.res_r==="مطابق").length;
  const diff    = fullResults.filter(r=> r.res_g==="مخالف" || r.res_r==="مخالف").length;
  const missing = fullResults.filter(r=> r.res_g==="ناقص" && r.res_r==="ناقص").length;

  statAllBtn.textContent     = `الكل ${all}`;
  statMatchBtn.textContent   = `مطابق ${match}`;
  statDiffBtn.textContent    = `مخالف ${diff}`;
  statMissingBtn.textContent = `ناقص/غير مكتمل ${missing}`;

  downloadBtn.disabled = !all;
}

function applySearchAndFilter(base){
  const q = normalizeArabic(searchBox.value);
  if(!q) return base;
  return base.filter(r =>
    normalizeArabic(r.name_fp).includes(q) ||
    normalizeArabic(r.name_m).includes(q) ||
    String(r.code_fp).includes(q) ||
    String(r.code_m).includes(q)
  );
}

function filterResults(kind){
  let list = fullResults.slice();
  if(kind==="match")   list = list.filter(r=> r.res_g==="مطابق" && r.res_r==="مطابق");
  if(kind==="diff")    list = list.filter(r=> r.res_g==="مخالف" || r.res_r==="مخالف");
  if(kind==="missing") list = list.filter(r=> r.res_g==="ناقص" && r.res_r==="ناقص");
  list = applySearchAndFilter(list);
  renderTable(list);
}

/* ========= Core Compare ========= */
function compareRecords(fpRows, manualRows){
  const fp = fpRows.map(r=>mapRow(r,false));
  const mn = manualRows.map(r=>mapRow(r,true)); // 👈 manual=true لتطبيق 31→30

  // فهرسة اليدوي حسب الكود
  const byCode=new Map();
  mn.forEach(m=>{
    if(!byCode.has(m.code)) byCode.set(m.code,[]);
    byCode.get(m.code).push(m);
  });

  const results=[];
  for(const f of fp){
    let resG="ناقص", resR="ناقص", note="";
    let mMatch=null;

    const sameCode = byCode.get(f.code) || [];
    if(sameCode.length){
      mMatch = sameCode.find(m => normalizeArabic(m.name)===normalizeArabic(f.name));
      if(!mMatch){
        // مرونة الاسم
        mMatch = sameCode.find(m => namesClose(f.name,m.name,0.60)) || null;
        if(mMatch && normalizeArabic(mMatch.name)!==normalizeArabic(f.name)){
          note = "ⓘ تم اعتماد التطبيع المرن للاسم (الكود متطابق)";
        }
      }
    }

    if(!mMatch){
      results.push({
        code_fp:f.code, name_fp:f.name, g_fp:f.g, r_fp:f.r,
        code_m:"", name_m:"", g_m:"", r_m:"",
        res_g:"ناقص", res_r:"ناقص",
        note:"بيانات ناقصة أو غير موجودة في الكشف اليدوي"
      });
      continue;
    }

    // مقارنة غ
    if (f.g === mMatch.g) {
      resG="مطابق";
    } else if (f.g > mMatch.g) {
      resG="مخالف"; note ||= "يتم التأكد من صحة الادخال اليدوي غ";
    } else {
      resG="مخالف"; note ||= `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق ${(mMatch.g - f.g).toFixed(1)}`;
    }
    // مقارنة ر
    if (f.r === mMatch.r) {
      resR="مطابق";
    } else if (f.r > mMatch.r) {
      resR="مخالف"; note ||= "يتم التأكد من صحة الادخال اليدوي ر";
    } else {
      resR="مخالف"; note ||= `بعد التأكد من الادخال يتم عمل ر بالفارق ${(mMatch.r - f.r).toFixed(1)}`;
    }

    // لا نكتب «مطابق» في الملاحظات؛ تبقى فارغة إلا إذا عندنا ملاحظة فعلية
    results.push({
      code_fp:f.code, name_fp:f.name, g_fp:f.g, r_fp:f.r,
      code_m:mMatch.code, name_m:mMatch.name, g_m:mMatch.g, r_m:mMatch.r,
      res_g:resG, res_r:resR,
      note
    });
  }

  // فرز حسب الكود تصاعدي (رقميًا إن أمكن)
  results.sort((a,b)=> Number(a.code_fp) - Number(b.code_fp));
  return results;
}

/* ========= Events ========= */
fpInput.addEventListener('change', async ()=>{
  fpData=null;
  if(fpInput.files?.[0]){
    fpNameSpan.textContent = fpInput.files[0].name;
    fpData = await readXlsx(fpInput.files[0]);
  }else{ fpNameSpan.textContent="— لم يتم اختيار ملف بعد"; }
  updateStartState();
});

manualInput.addEventListener('change', async ()=>{
  manualData=null;
  if(manualInput.files?.[0]){
    manualNameSpan.textContent = manualInput.files[0].name;
    manualData = await readXlsx(manualInput.files[0]);
  }else{ manualNameSpan.textContent="— لم يتم اختيار ملف بعد"; }
  updateStartState();
});

startBtn.addEventListener('click', ()=>{
  if(!(fpData && manualData)){
    alert("رجاءً اختر ملفي البصمة واليدوي (XLSX) أولًا.");
    return;
  }
  fullResults = compareRecords(fpData, manualData);
  updateStats();
  filterResults("all");
});

statAllBtn.addEventListener('click',   ()=>filterResults("all"));
statMatchBtn.addEventListener('click', ()=>filterResults("match"));
statDiffBtn.addEventListener('click',  ()=>filterResults("diff"));
statMissingBtn.addEventListener('click', ()=>filterResults("missing"));

searchBox.addEventListener('input', ()=> filterResults("all"));

// تنزيل النتائج XLSX
downloadBtn.addEventListener('click', ()=>{
  if(!fullResults.length) return;

  const rows = fullResults.map((r,i)=>({
    "م": i+1,
    "الكود (بصمة)": r.code_fp,
    "الاسم (بصمة)": r.name_fp,
    "غ (بصمة)": r.g_fp,
    "ر (بصمة)": r.r_fp,
    "الكود (يدوي)": r.code_m,
    "الاسم (يدوي)": r.name_m,
    "غ (يدوي)": r.g_m,
    "ر (يدوي)": r.r_m,
    "نتيجة غ": r.res_g,
    "نتيجة ر": r.res_r,
    // لا نضع «مطابق» في الملاحظة إذا فارغة
    "الملاحظة": r.note || ""
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "نتيجة المطابقة");
  XLSX.writeFile(wb, "canary_monthly_result.xlsx");
});
