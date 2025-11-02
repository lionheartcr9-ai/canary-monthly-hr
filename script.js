// ✅ دالة لتقريب القيم العشرية بدقة (مثلاً 6.500000625 → 6.5)
const roundValue = (num) => {
  if (num === "" || num === null || isNaN(num)) return 0;
  return Math.round((parseFloat(num) + Number.EPSILON) * 100) / 100;
};

// ✅ تحميل ملف Excel وإرجاع بياناته كمصفوفة كائنات
async function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet);
      resolve(rows);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ✅ الدالة الرئيسية لمقارنة ملفي البصمة واليدوي
function compareRecords(fingerprint, manual) {
  const results = [];

  fingerprint.forEach((f) => {
    // إيجاد السطر المطابق بالكود والاسم
    const match = manual.find(
      (m) => String(m["الكود"]) === String(f["الكود"]) && m["الاسم"] === f["الاسم"]
    );

    // إذا لم يوجد تطابق → بيانات ناقصة
    if (!match) {
      results.push({
        الكود: f["الكود"],
        الاسم: f["الاسم"],
        "غ (بصمة)": f["غ"],
        "ر (بصمة)": f["ر"],
        "غ (يدوي)": "",
        "ر (يدوي)": "",
        "نتيجة غ": "ناقص",
        "نتيجة ر": "ناقص",
        الملاحظة: "بيانات ناقصة أو غير موجودة في الكشف اليدوي",
      });
      return;
    }

    // تقريب القيم لتفادي فروق الكسور العشرية
    const g_f = roundValue(f["غ"]);
    const g_m = roundValue(match["غ"]);
    const r_f = roundValue(f["ر"]);
    const r_m = roundValue(match["ر"]);

    let resultG = "";
    let resultR = "";
    let note = "";

    // ✅ مقارنة عدد الغياب (غ)
    if (g_f === g_m) {
      resultG = "مطابق";
    } else if (g_f > g_m) {
      resultG = "مخالف";
      note = "يتم التأكد من صحة الادخال اليدوي غ";
    } else if (g_f < g_m) {
      resultG = "مخالف";
      note = `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق ${(g_m - g_f).toFixed(1)}`;
    }

    // ✅ مقارنة عدد الإجازات (ر)
    if (r_f === r_m) {
      resultR = "مطابق";
    } else if (r_f > r_m) {
      resultR = "مخالف";
      if (!note) note = "يتم التأكد من صحة الادخال اليدوي ر";
    } else if (r_f < r_m) {
      resultR = "مخالف";
      if (!note) note = `بعد التأكد من الادخال يتم عمل ر بالفارق ${(r_m - r_f).toFixed(1)}`;
    }

    // ✅ إضافة النتيجة النهائية
    results.push({
      الكود: f["الكود"],
      الاسم: f["الاسم"],
      "غ (بصمة)": g_f,
      "ر (بصمة)": r_f,
      "غ (يدوي)": g_m,
      "ر (يدوي)": r_m,
      "نتيجة غ": resultG,
      "نتيجة ر": resultR,
      الملاحظة: note || "مطابق",
    });
  });

  return results;
}

// ✅ عرض النتائج في الجدول داخل الصفحة
function displayResults(results) {
  const tableBody = document.getElementById("resultsTable");
  tableBody.innerHTML = "";

  results.forEach((r) => {
    const row = document.createElement("tr");

    // ألوان الخلفية حسب النتيجة
    const colorG =
      r["نتيجة غ"] === "مطابق" ? "#004d00" : r["نتيجة غ"] === "ناقص" ? "#666600" : "#660000";
    const colorR =
      r["نتيجة ر"] === "مطابق" ? "#004d00" : r["نتيجة ر"] === "ناقص" ? "#666600" : "#660000";

    row.innerHTML = `
      <td>${r["الكود"]}</td>
      <td style="max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${r["الاسم"]}</td>
      <td>${r["غ (بصمة)"]}</td>
      <td>${r["غ (يدوي)"]}</td>
      <td style="background:${colorG};color:white">${r["نتيجة غ"]}</td>
      <td>${r["ر (بصمة)"]}</td>
      <td>${r["ر (يدوي)"]}</td>
      <td style="background:${colorR};color:white">${r["نتيجة ر"]}</td>
      <td style="max-width:300px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${r["الملاحظة"]}</td>
    `;
    tableBody.appendChild(row);
  });
}

// ✅ تصدير النتائج إلى ملف Excel
function exportToExcel(results) {
  const worksheet = XLSX.utils.json_to_sheet(results);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "النتائج الشهرية");
  XLSX.writeFile(workbook, "kanari_monthly_result.xlsx");
}

// ✅ تحميل الملفات وتشغيل المقارنة
document.getElementById("compareBtn").addEventListener("click", async () => {
  const fingerprintFile = document.getElementById("fingerFile").files[0];
  const manualFile = document.getElementById("manualFile").files[0];

  if (!fingerprintFile || !manualFile) {
    alert("الرجاء اختيار الملفين أولاً (البصمة واليدوي)");
    return;
  }

  const fingerprint = await readExcel(fingerprintFile);
  const manual = await readExcel(manualFile);
  const results = compareRecords(fingerprint, manual);

  displayResults(results);

  // حساب الإجماليات
  const total = results.length;
  const matching = results.filter((r) => r["نتيجة غ"] === "مطابق" && r["نتيجة ر"] === "مطابق").length;
  const diff = results.filter(
    (r) => r["نتيجة غ"] === "مخالف" || r["نتيجة ر"] === "مخالف"
  ).length;
  const missing = results.filter(
    (r) => r["نتيجة غ"] === "ناقص" || r["نتيجة ر"] === "ناقص"
  ).length;

  document.getElementById("summary").innerHTML = `
    <b>الإجمالي:</b> ${total} &nbsp; | &nbsp;
    <b>مطابق:</b> ${matching} &nbsp; | &nbsp;
    <b>مخالف:</b> ${diff} &nbsp; | &nbsp;
    <b>ناقص:</b> ${missing}
  `;

  // حفظ النتيجة العامة لاستخدامها في التصدير
  window.lastResults = results;
});

// ✅ زر تصدير النتيجة إلى ملف Excel
document.getElementById("exportBtn").addEventListener("click", () => {
  if (!window.lastResults || window.lastResults.length === 0) {
    alert("لا توجد نتائج لتصديرها، الرجاء إجراء المقارنة أولاً.");
    return;
  }
  exportToExcel(window.lastResults);
});
