/* ===============================
   أدوات مساعدة
================================*/

// تقريب دقيق لرقمين عشريّين (يعالج 6.500000625 → 6.5)
const round2 = (v) => {
  if (v === "" || v === null || v === undefined) return 0;
  const n = parseFloat(v);
  if (Number.isNaN(n)) return 0;
  return Math.round((n + Number.EPSILON) * 100) / 100;
};

// مقارنة أرقام مع سماحية عشرية صغيرة (لتمييز 6.5 و 6.500000625 أنهم متساويان)
const nearlyEqual = (a, b, tol = 0.01) => Math.abs(round2(a) - round2(b)) <= tol;

// تطبيع نص عربي لإزالة الإشكالات الشائعة (مسافات/تشكيل/ألف-همزات/ياء-ألف مقصورة/هاء-تاء مربوطة…)
const normalizeArabic = (s = "") =>
  String(s)
    .trim()
    // حذف التشكيل
    .replace(/[\u064B-\u065F]/g, "")
    // توحيد الألف بهمزاتها
    .replace(/[إأآٱ]/g, "ا")
    // ياء/ألف مقصورة
    .replace(/ى/g, "ي")
    // هاء/تاء مربوطة
    .replace(/ة/g, "ه")
    // همزة منفردة
    .replace(/ؤ|ئ/g, "ء")
    // مسافات مكررة
    .replace(/\s+/g, " ")
    .toLowerCase();

// تطابق مرن للأسماء: مساواة بعد التطبيع أو احتواء أو تشابه بالتوكنات
const softNameMatch = (a, b) => {
  const A = normalizeArabic(a);
  const B = normalizeArabic(b);
  if (!A || !B) return false;
  if (A === B) return true;
  if (A.includes(B) || B.includes(A)) return true;

  // تشابه بالتوكنات (تقاطع الكلمات)
  const ta = new Set(A.split(" ").filter(Boolean));
  const tb = new Set(B.split(" ").filter(Boolean));
  const inter = [...ta].filter((w) => tb.has(w)).length;
  const union = new Set([...ta, ...tb]).size;
  const jaccard = inter / (union || 1);
  return jaccard >= 0.7;
};

/* ===============================
   التحضير: فهرسة الكشف اليدوي بالكود
================================*/
const indexByCode = (rows) => {
  const map = new Map();
  rows.forEach((r) => {
    const code = String(r["الكود"]).trim();
    if (!map.has(code)) map.set(code, []);
    map.get(code).push(r);
  });
  return map;
};

/* ===============================
   المقارنة الرئيسية (كود ← اسم مرن)
================================*/
function compareRecords(fingerprintRows, manualRows) {
  const manualByCode = indexByCode(manualRows);
  const results = [];

  fingerprintRows.forEach((f) => {
    const code = String(f["الكود"]).trim();
    const nameF = String(f["الاسم"] || "").trim();

    // ابحث حسب الكود فقط
    const manualCandidates = manualByCode.get(code) || [];

    // لا يوجد أي صف بنفس الكود في اليدوي ⇒ بيانات ناقصة
    if (manualCandidates.length === 0) {
      results.push({
        "م": 0, // سيُعاد ملؤه بعد الترتيب
        "الكود (بصمة)": code,
        "الاسم (بصمة)": nameF,
        "غ (بصمة)": round2(f["غ"]),
        "ر (بصمة)": round2(f["ر"]),
        "الكود (يدوي)": "",
        "الاسم (يدوي)": "",
        "غ (يدوي)": "",
        "ر (يدوي)": "",
        "نتيجة غ": "ناقص",
        "نتيجة ر": "ناقص",
        "الملاحظة": "بيانات ناقصة",
      });
      return;
    }

    // اختر أفضل مرشح بالاسم (إن وُجد) وإلا خذ أول صف للكود
    let best = manualCandidates[0];
    let usedSoftMatch = false;

    // لو وُجد مرشح يطابق الاسم مرنًا نختاره
    const bySoft = manualCandidates.find((m) => softNameMatch(nameF, m["الاسم"]));
    if (bySoft) {
      best = bySoft;
      // إن الاسم بعد التطبيع ليس مطابقًا تمامًا للنص الأصلي، أشر بأنه "مرن"
      usedSoftMatch = normalizeArabic(nameF) !== normalizeArabic(best["الاسم"]);
    } else if (manualCandidates.length > 1) {
      // أكثر من صف بنفس الكود والاسم لم يطابق مرنًا: نواصل على أول صف ونشير "مرن"
      usedSoftMatch = true;
    }

    // القيم الرقمية (مع التقريب وسماحية)
    const gF = round2(f["غ"]);
    const rF = round2(f["ر"]);
    const gM = round2(best["غ"]);
    const rM = round2(best["ر"]);

    let resultG = "مطابق";
    let resultR = "مطابق";
    let note = "";

    // مقارنة غياب (غ)
    if (!nearlyEqual(gF, gM)) {
      if (gF > gM) {
        resultG = "مخالف";
        note = "يتم التأكد من صحة الادخال اليدوي غ";
      } else {
        resultG = "مخالف";
        const diff = round2(gM - gF);
        note = `بعد التأكد من الادخال يتم عمل استيفاء غ بالفارق ${diff}`;
      }
    }

    // مقارنة (ر)
    if (!nearlyEqual(rF, rM)) {
      if (resultR === "مطابق") {
        if (rF > rM) {
          resultR = "مخالف";
          if (!note) note = "يتم التأكد من صحة الادخال اليدوي ر";
        } else {
          resultR = "مخالف";
          const diff = round2(rM - rF);
          if (!note) note = `بعد التأكد من الادخال يتم عمل ر بالفارق ${diff}`;
        }
      } else {
        // موجودة ملاحظة مسبقًا من غ؛ فقط عدّل النتيجة
        if (rF > rM) resultR = "مخالف";
        else resultR = "مخالف";
      }
    }

    // وسم التطبيع المرن (قصير وملاصق لأي ملاحظة موجودة)
    if (usedSoftMatch) {
      note = note ? `${note} • مرن` : "مرن";
    }
    if (!note) note = "مطابق";

    results.push({
      "م": 0, // سيملأ لاحقًا بعد الفرز
      "الكود (بصمة)": code,
      "الاسم (بصمة)": nameF,
      "غ (بصمة)": gF,
      "ر (بصمة)": rF,
      "الكود (يدوي)": String(best["الكود"]).trim(),
      "الاسم (يدوي)": String(best["الاسم"] || "").trim(),
      "غ (يدوي)": gM,
      "ر (يدوي)": rM,
      "نتيجة غ": resultG,
      "نتيجة ر": resultR,
      "الملاحظة": note,
    });
  });

  // فرز تصاعدي بالكود (رقميًا إن أمكن)
  results.sort((a, b) => {
    const na = Number(a["الكود (بصمة)"]);
    const nb = Number(b["الكود (بصمة)"]);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
    return String(a["الكود (بصمة)"]).localeCompare(String(b["الكود (بصمة)"]));
  });

  // تعبئة عمود الترقيم "م"
  results.forEach((row, i) => (row["م"] = i + 1));

  return results;
}

/* ملاحظة:
   - لا يوجد أي رفض للسطر إذا اختلف الاسم مع تطابق الكود؛ فقط نُشير بـ "مرن".
   - لو ظهرت ملاحظة أخرى بسبب (غ/ر)، ستظهر بهذا الشكل: "… • مرن".
*/
