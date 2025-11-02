/* ================================
   Canary Monthly HR — script.js (V7)
   - XLSX only
   - Arabic-first UI (RTL)
   - Code-first matching + soft name normalization
   - Robust rounding for G/R
   - Search + filters + XLSX export
=================================== */

// ------- عناصر DOM أساسية -------
const fpInput = document.getElementById("fpFile");
const manualInput = document.getElementById("manualFile");
const startBtn = document.getElementById("startBtn");
const downloadBtn = document.getElementById("downloadBtn");
const statsTotal = document.getElementById("stat-total");
const statsMatch = document.getElementById("stat-match");
const statsMismatch = document.getElementById("stat-mismatch");
const statsMissing = document.getElementById("stat-missing");
const filterAll = document.getElementById("filter-all");
const filterOk = document.getElementById("filter-ok");
const filterBad = document.getElementById("filter-bad");
const filterMiss = document.getElementById("filter-miss");
const tableBody = document.getElementById("result-body");
const searchInput = document.getElementById("search");

// الحالة الداخلية
let fpRows = [];      // بصمة
let mnRows = [];      // يدوي
let results = [];     // نتيجة المعروضة
let rawResults = [];  // نتيجة كاملة بدون فلترة

// ====== أدوات قراءة XLSX ======
async function readXlsx(file) {
  if (!file) return [];
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return rows.map(normalizeRow);
}

// نعتمد العناوين: الكود | الاسم | غ | ر
function normalizeRow(r) {
  const code = r["الكود"] ?? r["code"] ?? r["Code"] ?? r["ID"] ?? r["id"] ?? "";
  const name = r["الاسم"] ?? r["name"] ?? r["Name"] ?? "";
  const g = r["غ"] ?? r["G"] ?? r["Abs"] ?? r["abs"] ?? "";
  const rdays = r["ر"] ?? r["R"] ?? r["Leave"] ?? r["leave"] ?? "";

  return {
    الكود: safeCode(code),
    الاسم: String(name).trim(),
    غ: toNumberSafe(g),
    ر: toNumberSafe(rdays),
  };
}

// الكود كسلسلة digits
function safeCode(val) {
  const s = String(val).trim();
  return s.replace(/[^\d]/g, "");
}

// رقم آمن + تقريب
function toNumberSafe(v) {
  if (v === "" || v === null || v === undefined) return 0;
  const n = Number(String(v).replace(",", "."));
  if (isNaN(n)) return 0;
  return round2(n);
}

// ✅ تقريب بدرجتين (يعالج 6.500000625 = 6.5)
function round2(num) {
  return Math.round((parseFloat(num) + Number.EPSILON) * 100) / 100;
}

// ====== تطبيع عربي للأسماء ======
function normalizeArabic(str) {
  if (!str) return "";
  let s = String(str).trim();

  // إزالة التشكيل والرموز
  s = s
    .replace(/[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]/g, "")
    .replace(/[\u0640\-\_\.\،\,\؛\;\:\/\(\)\[\]\{\}\«\»\"\']/g, " ")
    .replace(/\s+/g, " ");

  // توحيد أحرف
  s = s
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/ؤ/g, "و")
    .replace(/ئ/g, "ي");

  return s.trim();
}

function tokenizeName(str) {
  return normalizeArabic(str)
    .split(" ")
    .filter(Boolean);
}

// شبه-تشابه بالاعتماد على بداية الاسم + تقاطع الكلمات
function nameSimilarity(a, b) {
  const A = tokenizeName(a);
  const B = tokenizeName(b);
  if (A.length === 0 || B.length === 0) return 0;

  const firstBoost = A[0] === B[0] ? 0.3 : 0;

  const setA = new Set(A);
  const setB = new Set(B);
  let inter = 0;
  for (const w of setA) if (setB.has(w)) inter++;

  const union = new Set([...A, ...B]).size;
  const jacc = union ? inter / union : 0;

  let secondBoost = 0;
  if (A.length > 1 && B.length > 1 && A[1] === B[1]) secondBoost = 0.2;

  return Math.min(1, jacc + firstBoost + secondBoost);
}

// ====== المقارنة الرئيسية ======
function compareRecords(fp, mn) {
  // فهرس سريع لليدوي حسب الكود
  const manualByCode = new Map();
  mn.forEach((row) => {
    const code = row["الكود"];
    if (!manualByCode.has(code)) manualByCode.set(code, []);
    manualByCode.get(code).push(row);
  });

  const out = [];

  // نمر على ملف البصمة
  fp.forEach((f) => {
    const code = f["الكود"];
    const nameF = f["الاسم"];

    // لا يوجد هذا الكود في اليدوي → ناقص
    if (!manualByCode.has(code)) {
      out.push(makeRow(f, null, "ناقص", "ناقص", "بيانات ناقصة أو غير موجودة في الكشف اليدوي"));
      return;
    }

    // لو وجد أكثر من اسم لنفس الكود، نختار الأعلى تشابهًا
    const candidates = manualByCode.get(code);
    let best = candidates[0];
    let bestScore = nameSimilarity(nameF, candidates[0]["الاسم"]);
    for (let i = 1; i < candidates.length; i++) {
      const s = nameSimilarity(nameF, candidates[i]["الاسم"]);
      if (s > bestScore) {
        best = candidates[i];
        bestScore = s;
      }
    }

    // مقارنة غ/ر — الكود متطابق مهما اختلف الاسم
    const gF = round2(f["غ"]);
    const rF = round2(f["ر"]);
    const gM = round2(best["غ"]);
    const rM = round2(best["ر"]);

    let resG = "";
    let resR = "";
    let note = "";

    // غ
    if (gF === gM) {
      resG = "مطابق";
    } else if (gF > gM) {
      resG = "مخالف";
      note = addNote(note, "يتم التأكد من صحة الادخال اليدوي غ");
    } else {
      resG = "مخالف";
      note = addNote(note, `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق ${round2(gM - gF)}`);
    }

    // ر
    if (rF === rM) {
      resR = "مطابق";
    } else if (rF > rM) {
      resR = "مخالف";
      note = addNote(note, "يتم التأكد من صحة الادخال اليدوي ر");
    } else if (rF < rM) {
      resR = "مخالف";
      note = addNote(note, `بعد التأكد من الادخال يتم عمل ر بالفارق ${round2(rM - rF)}`);
    }

    // اختلاف اسم واضح مع كود صحيح → لا نرفض؛ فقط وسم مرن
    if (bestScore < 0.6) {
      note = addNote(note, `مرن: اختلاف اسم (بصمة: ${nameF} | يدوي: ${best["الاسم"]})`);
    }

    out.push({
      م: 0, // يُملأ لاحقًا
      "الكود (بصمة)": code,
      "الاسم (بصمة)": nameF,
      "غ (بصمة)": gF,
      "ر (بصمة)": rF,
      "الكود (يدوي)": best["الكود"],
      "الاسم (يدوي)": best["الاسم"],
      "غ (يدوي)": gM,
      "ر (يدوي)": rM,
      "نتيجة غ": resG,
      "نتيجة ر": resR,
      الملاحظة: note || "مطابق",
    });
  });

  // ترتيب تصاعدي حسب الكود + ترقيم عمود م
  out.sort((a, b) => Number(a["الكود (بصمة)"]) - Number(b["الكود (بصمة)"]));
  out.forEach((r, i) => (r["م"] = i + 1));

  return out;
}

function addNote(oldNote, extra) {
  if (!oldNote) return extra;
  return `${oldNote} • ${extra}`;
}

function makeRow(fpRow, mnRow, resG, resR, note) {
  return {
    م: 0,
    "الكود (بصمة)": fpRow ? fpRow["الكود"] : "",
    "الاسم (بصمة)": fpRow ? fpRow["الاسم"] : "",
    "غ (بصمة)": fpRow ? round2(fpRow["غ"]) : 0,
    "ر (بصمة)": fpRow ? round2(fpRow["ر"]) : 0,
    "الكود (يدوي)": mnRow ? mnRow["الكود"] : "",
    "الاسم (يدوي)": mnRow ? mnRow["الاسم"] : "",
    "غ (يدوي)": mnRow ? round2(mnRow["غ"]) : 0,
    "ر (يدوي)": mnRow ? round2(mnRow["ر"]) : 0,
    "نتيجة غ": resG,
    "نتيجة ر": resR,
    الملاحظة: note,
  };
}

// ====== عرض النتائج ======
function renderStats(list) {
  const total = list.length;
  const ok = list.filter((r) => r["نتيجة غ"] === "مطابق" && r["نتيجة ر"] === "مطابق").length;
  const bad = list.filter((r) => r["نتيجة غ"] === "مخالف" || r["نتيجة ر"] === "مخالف").length;
  const miss = list.filter((r) => r["نتيجة غ"] === "ناقص" || r["نتيجة ر"] === "ناقص").length;

  statsTotal.textContent = total;
  statsMatch.textContent = ok;
  statsMismatch.textContent = bad;
  statsMissing.textContent = miss;
}

function renderTable(list) {
  tableBody.innerHTML = "";
  const frag = document.createDocumentFragment();
  list.forEach((r) => {
    const tr = document.createElement("tr");
    const clsG = r["نتيجة غ"] === "مطابق" ? "badge green" : r["نتيجة غ"] === "مخالف" ? "badge red" : "badge gray";
    const clsR = r["نتيجة ر"] === "مطابق" ? "badge green" : r["نتيجة ر"] === "مخالف" ? "badge red" : "badge gray";

    tr.innerHTML = `
      <td>${r["م"]}</td>
      <td>${r["الكود (بصمة)"]}</td>
      <td>${r["الاسم (بصمة)"]}</td>
      <td>${r["غ (بصمة)"]}</td>
      <td>${r["ر (بصمة)"]}</td>
      <td>${r["الكود (يدوي)"]}</td>
      <td>${r["الاسم (يدوي)"]}</td>
      <td>${r["غ (يدوي)"]}</td>
      <td>${r["ر (يدوي)"]}</td>
      <td><span class="${clsG}">${r["نتيجة غ"]}</span></td>
      <td><span class="${clsR}">${r["نتيجة ر"]}</span></td>
      <td>${r["الملاحظة"] || ""}</td>
    `;
    frag.appendChild(tr);
  });
  tableBody.appendChild(frag);
  renderStats(list);
}

// ====== فلاتر + بحث ======
function applyFilter(type) {
  let list = [...rawResults];
  if (type === "ok") {
    list = list.filter((r) => r["نتيجة غ"] === "مطابق" && r["نتيجة ر"] === "مطابق");
  } else if (type === "bad") {
    list = list.filter((r) => r["نتيجة غ"] === "مخالف" || r["نتيجة ر"] === "مخالف");
  } else if (type === "miss") {
    list = list.filter((r) => r["نتيجة غ"] === "ناقص" || r["نتيجة ر"] === "ناقص");
  }
  results = list;
  applySearch();
}

function applySearch() {
  const q = normalizeArabic(searchInput?.value || "");
  if (!q) {
    renderTable(results);
    return;
  }
  const filtered = results.filter((r) => {
    const code = String(r["الكود (بصمة)"]);
    const name1 = normalizeArabic(r["الاسم (بصمة)"]);
    const name2 = normalizeArabic(r["الاسم (يدوي)"]);
    return code.includes(q) || name1.includes(q) || name2.includes(q);
  });
  renderTable(filtered);
}

// ====== تصدير XLSX ======
function downloadXlsx(list) {
  if (!list.length) return;
  const data = list.map((r) => ({
    "م": r["م"],
    "الكود (بصمة)": r["الكود (بصمة)"],
    "الاسم (بصمة)": r["الاسم (بصمة)"],
    "غ (بصمة)": r["غ (بصمة)"],
    "ر (بصمة)": r["ر (بصمة)"],
    "الكود (يدوي)": r["الكود (يدوي)"],
    "الاسم (يدوي)": r["الاسم (يدوي)"],
    "غ (يدوي)": r["غ (يدوي)"],
    "ر (يدوي)": r["ر (يدوي)"],
    "نتيجة غ": r["نتيجة غ"],
    "نتيجة ر": r["نتيجة ر"],
    "الملاحظة": r["الملاحظة"],
  }));
  const ws = XLSX.utils.json_to_sheet(data, { header: Object.keys(data[0] || {}) });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "canary_monthly_result");
  XLSX.writeFile(wb, "canary_monthly_result.xlsx");
}

// ====== ربط الأحداث ======
fpInput?.addEventListener("change", async (e) => {
  fpRows = await readXlsx(e.target.files[0]);
});
manualInput?.addEventListener("change", async (e) => {
  mnRows = await readXlsx(e.target.files[0]);
});

startBtn?.addEventListener("click", () => {
  if (!fpRows.length || !mnRows.length) {
    alert("رجاءً اختر ملفي البصمة واليدوي (XLSX) أولاً.");
    return;
  }
  rawResults = compareRecords(fpRows, mnRows);
  results = [...rawResults];
  renderTable(results);
});

downloadBtn?.addEventListener("click", () => {
  if (!rawResults.length) return;
  downloadXlsx(results.length ? results : rawResults);
});

// فلاتر
filterAll?.addEventListener("click", () => applyFilter("all"));
filterOk?.addEventListener("click", () => applyFilter("ok"));
filterBad?.addEventListener("click", () => applyFilter("bad"));
filterMiss?.addEventListener("click", () => applyFilter("miss"));

// بحث مباشر
searchInput?.addEventListener("input", applySearch);
