/* =======================
   Globals + DOM bindings
======================= */
let fpData = null;        // بيانات البصمة (صفوف JSON)
let manualData = null;    // بيانات اليدوي
let fullResults = [];     // النتيجة الكاملة قبل الفلترة

const fpInput          = document.getElementById('fpFile');
const manualInput      = document.getElementById('manualFile');
const fpNameSpan       = document.getElementById('fpName');
const manualNameSpan   = document.getElementById('manualName');

const startBtn         = document.getElementById('startCompare');
const downloadBtn      = document.getElementById('downloadXlsx');

const statAllBtn       = document.getElementById('statAll');
const statMatchBtn     = document.getElementById('statMatch');
const statDiffBtn      = document.getElementById('statDiff');
const statMissingBtn   = document.getElementById('statMissing');

const searchBox        = document.getElementById('searchBox');
const resultBody       = document.getElementById('resultBody');

/* =======================
   Helpers
======================= */

// قراءة أول ورقة من ملف XLSX كـ JSON
async function readXlsx(file) {
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, { type: 'array' });
  const sh  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sh, { defval: "" }); // لا نترك undefined
}

// تحديث حالة زر "بدء المطابقة"
function updateStartState() {
  startBtn.disabled = !(fpData && manualData);
}

// تقريب القيم العشرية لمرتبتين
const roundValue = (num) => {
  if (num === "" || num === null || isNaN(num)) return 0;
  return Math.round((parseFloat(num) + Number.EPSILON) * 100) / 100;
};

// تطبيع بسيط للنص العربي (إزالة تشكيل/مدود/مسافات زائدة + توحيد الألف/الياء/الهاء…)
function normalizeArabic(str) {
  if (!str) return "";
  return String(str)
    .replace(/[ًٌٍَُِّْـ]/g, "")          // التشكيل والمد
    .replace(/[\u200F\u200E]/g, "")      // علامات اتجاه
    .replace(/[إأآا]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/\s+/g, " ")
    .trim();
}

// تقسيم الاسم إلى حروف عربية فقط للمقارنة المرنة
function tokenizeName(name) {
  return normalizeArabic(name)
    .replace(/[^ء-ي\s]/g, "")
    .split(" ")
    .filter(Boolean);
}

// درجة تشابه بسيطة (Dice coefficient) بين قائمتين من الكلمات
function diceSimilarity(tokensA, tokensB) {
  const setA = new Set(tokensA);
  const setB = new Set(tokensB);
  let inter = 0;
  setA.forEach(t => { if (setB.has(t)) inter++; });
  const denom = setA.size + setB.size;
  return denom ? (2 * inter) / denom : 0;
}

// تطبيع مرن للاسم: يعتبر "عثمان عبده مسعد سعيد" ≈ "عثمان عبده مسعد الفلاحي"
function namesClose(a, b, threshold = 0.60) {
  const ta = tokenizeName(a);
  const tb = tokenizeName(b);
  if (!ta.length || !tb.length) return false;
  return diceSimilarity(ta, tb) >= threshold;
}

// تأمين هيكل الصف (خرائط الأعمدة العربية المعتمدة)
function mapRow(row) {
  return {
    code: String(row["الكود"] ?? row["الكود (بصمة)"] ?? row["الكود (يدوي)"] ?? "").trim(),
    name: String(row["الاسم"] ?? row["الاسم (بصمة)"] ?? row["الاسم (يدوي)"] ?? "").trim(),
    g: roundValue(row["غ"] ?? row["غ (بصمة)"] ?? row["غ (يدوي)"] ?? 0),
    r: roundValue(row["ر"] ?? row["ر (بصمة)"] ?? row["ر (يدوي)"] ?? 0),
  };
}

// بناء صف العرض للجدول
function buildRow(idx, rec) {
  const tr = document.createElement('tr');
  const cells = [
    idx + 1,
    rec.code_fp, rec.name_fp, rec.g_fp, rec.r_fp,
    rec.code_m,  rec.name_m,  rec.g_m, rec.r_m,
    rec.res_g,   rec.res_r,   rec.note
  ];
  cells.forEach(v => {
    const td = document.createElement('td');
    td.textContent = (v === undefined || v === null) ? "" : v;
    tr.appendChild(td);
  });
  return tr;
}

// رسم الجدول مع فلترة اختيارية
function renderTable(list) {
  resultBody.innerHTML = "";
  list.forEach((rec, i) => resultBody.appendChild(buildRow(i, rec)));
}

// تحديث العدادات وتفعيل زر تنزيل
function updateStats() {
  const all = fullResults.length;
  const match = fullResults.filter(r => r.res_g === "مطابق" && r.res_r === "مطابق").length;
  const diff = fullResults.filter(r => r.res_g === "مخالف" || r.res_r === "مخالف").length;
  const missing = fullResults.filter(r => r.res_g === "ناقص" && r.res_r === "ناقص").length;

  statAllBtn.textContent     = `الكل ${all}`;
  statMatchBtn.textContent   = `مطابق ${match}`;
  statDiffBtn.textContent    = `مخالف ${diff}`;
  statMissingBtn.textContent = `ناقص/غير مكتمل ${missing}`;

  downloadBtn.disabled = all === 0;
}

// فلترة حسب نوع
function filterResults(type) {
  let filtered = fullResults.slice();
  if (type === "match") {
    filtered = filtered.filter(r => r.res_g === "مطابق" && r.res_r === "مطابق");
  } else if (type === "diff") {
    filtered = filtered.filter(r => r.res_g === "مخالف" || r.res_r === "مخالف");
  } else if (type === "missing") {
    filtered = filtered.filter(r => r.res_g === "ناقص" && r.res_r === "ناقص");
  }
  // تطبيق بحث إن وجد
  const q = normalizeArabic(searchBox.value);
  if (q) {
    filtered = filtered.filter(r =>
      normalizeArabic(r.name_fp).includes(q) ||
      normalizeArabic(r.name_m).includes(q) ||
      String(r.code_fp).includes(q) ||
      String(r.code_m).includes(q)
    );
  }
  renderTable(filtered);
}

/* =======================
   Core Compare
======================= */

// مقارنة رئيسية مع التطبيع المرن + أولوية الكود
function compareRecords(fpRows, manualRows) {
  // خرائط
  const fp = fpRows.map(mapRow);
  const mn = manualRows.map(mapRow);

  // فهرس اليدوي حسب الكود (قد يكون الكود مُكرر؛ نخزن قائمة)
  const byCode = new Map();
  mn.forEach(m => {
    if (!byCode.has(m.code)) byCode.set(m.code, []);
    byCode.get(m.code).push(m);
  });

  const results = [];

  for (const f of fp) {
    let resG = "ناقص", resR = "ناقص", note = "";
    let mMatch = null;

    // 1) نحاول مطابقة الكود أولًا
    const sameCode = byCode.get(f.code) || [];

    // 2) داخل نفس الكود: نتحقق من الاسم (مرن)
    if (sameCode.length) {
      // الأفضل: اسم متطابق تمامًا، وإلا الأقرب مرونة
      mMatch = sameCode.find(m => normalizeArabic(m.name) === normalizeArabic(f.name));
      if (!mMatch) {
        mMatch = sameCode
          .map(m => ({ m, score: namesClose(f.name, m.name, 0.60) ? 1 : 0 }))
          .filter(x => x.score > 0)
          .map(x => x.m)[0] || null;
        if (mMatch && normalizeArabic(mMatch.name) !== normalizeArabic(f.name)) {
          // ملاحظة التطبيع المرن
          note = "ⓘ تم اعتماد التطبيع المرن للاسم (الكود متطابق)";
        }
      }
    }

    if (!mMatch) {
      // لا يوجد في اليدوي بنفس الكود → بيانات ناقصة
      results.push({
        code_fp: f.code, name_fp: f.name, g_fp: f.g, r_fp: f.r,
        code_m: "", name_m: "", g_m: "", r_m: "",
        res_g: "ناقص", res_r: "ناقص",
        note: "بيانات ناقصة أو غير موجودة في الكشف اليدوي"
      });
      continue;
    }

    // تقريب مسبق تم في mapRow، نقارن الآن
    if (f.g === mMatch.g) {
      resG = "مطابق";
    } else if (f.g > mMatch.g) {
      resG = "مخالف";
      note ||= "يتم التأكد من صحة الادخال اليدوي غ";
    } else {
      resG = "مخالف";
      note ||= `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق ${(mMatch.g - f.g).toFixed(1)}`;
    }

    if (f.r === mMatch.r) {
      resR = "مطابق";
    } else if (f.r > mMatch.r) {
      resR = "مخالف";
      note ||= "يتم التأكد من صحة الادخال اليدوي ر";
    } else {
      resR = "مخالف";
      note ||= `بعد التأكد من الادخال يتم عمل ر بالفارق ${(mMatch.r - f.r).toFixed(1)}`;
    }

    results.push({
      code_fp: f.code, name_fp: f.name, g_fp: f.g, r_fp: f.r,
      code_m: mMatch.code, name_m: mMatch.name, g_m: mMatch.g, r_m: mMatch.r,
      res_g: resG, res_r: resR,
      note: note || "مطابق"
    });
  }

  // فرز تصاعدي حسب الكود (رقميًا إن أمكن)
  results.sort((a, b) => Number(a.code_fp) - Number(b.code_fp));
  return results;
}

/* =======================
   Wire events
======================= */

fpInput.addEventListener('change', async () => {
  fpData = null;
  if (fpInput.files && fpInput.files[0]) {
    fpNameSpan.textContent = fpInput.files[0].name;
    fpData = await readXlsx(fpInput.files[0]);
  } else {
    fpNameSpan.textContent = "— لم يتم اختيار ملف بعد";
  }
  updateStartState();
});

manualInput.addEventListener('change', async () => {
  manualData = null;
  if (manualInput.files && manualInput.files[0]) {
    manualNameSpan.textContent = manualInput.files[0].name;
    manualData = await readXlsx(manualInput.files[0]);
  } else {
    manualNameSpan.textContent = "— لم يتم اختيار ملف بعد";
  }
  updateStartState();
});

startBtn.addEventListener('click', () => {
  if (!(fpData && manualData)) {
    alert("رجاءً اختر ملفي البصمة واليدوي (XLSX) أولًا.");
    return;
  }
  fullResults = compareRecords(fpData, manualData);
  updateStats();
  filterResults("all");
});

statAllBtn.addEventListener('click',   () => filterResults("all"));
statMatchBtn.addEventListener('click', () => filterResults("match"));
statDiffBtn.addEventListener('click',  () => filterResults("diff"));
statMissingBtn.addEventListener('click', () => filterResults("missing"));

searchBox.addEventListener('input', () => {
  // نعيد تطبيق آخر نوع فلترة نشِط لو أردت؛ هنا نعرض كل النتائج مع البحث
  filterResults("all");
});

// تنزيل النتائج XLSX
downloadBtn.addEventListener('click', () => {
  if (!fullResults.length) return;

  const rows = fullResults.map((r, i) => ({
    "م": i + 1,
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
    "الملاحظة": r.note
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "نتيجة المطابقة");
  XLSX.writeFile(wb, "canary_monthly_result.xlsx");
});
