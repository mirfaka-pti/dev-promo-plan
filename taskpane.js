/* global Office, Excel, OfficeRuntime */

// =============================================================================
// KONFIGURASI — SESUAIKAN
// =============================================================================

// ID unik yang harus ada di file "Promo Plan" asli. Ini BUKAN nilai bawaan Excel,
// jadi Anda perlu menempelkannya sendiri di file Promo Plan (lihat cara di bawah).
const EXPECTED_FILE_ID = "PROMO-PLAN-2026"; // <-- ganti sesuai kebutuhan Anda

// Nama custom document property yang dipakai untuk menyimpan ID di atas.
const FILE_ID_PROPERTY_KEY = "PromoPlanFileId";

// Nama sheet log
const LOG_SHEET_NAME = "Automation Log";

/*
  CARA MENEMPEL FILE_ID KE FILE PROMO PLAN ASLI (cukup sekali per file):
  Jalankan kode berikut sekali lewat console taskpane (F12) saat file itu terbuka,
  atau taruh sementara di tombol lain lalu klik sekali:

    Excel.run(async (context) => {
      context.workbook.properties.custom.add("PromoPlanFileId", "PROMO-PLAN-2026");
      await context.sync();
    });

  Setelah itu, custom property ini ikut tersimpan di dalam file .xlsx-nya sendiri
  (bukan di add-in), jadi permanen selama file itu ada.
*/

// =============================================================================

Office.onReady(async (info) => {
  try {
    const isExcel = info.host === Office.HostType.Excel;
    if (!isExcel) {
      showError("Add-in ini hanya didukung di aplikasi Excel.");
      return;
    }

    const isPromoPlan = await checkIsPromoPlanFile();
    const userName = await getCurrentUserName();
    if (!isPromoPlan) {
      showError(
        `File yang sedang dibuka bukan file Promo Plan yang dikenali (ID "${FILE_ID_PROPERTY_KEY}" tidak cocok atau tidak ditemukan). User: "${userName}"`
      );
      return;
    }

    showMain();
    document.getElementById("createLogBtn").addEventListener("click", createLog);
  } catch (err) {
    showError("Terjadi error saat memuat add-in: " + err.message);
  }
});

// ===================== Validasi =====================

async function checkIsPromoPlanFile() {
  let matched = false;
  await Excel.run(async (context) => {
    const prop = context.workbook.properties.custom.getItemOrNullObject(FILE_ID_PROPERTY_KEY);
    prop.load("value");
    await context.sync();

    if (!prop.isNullObject) {
      matched = String(prop.value) === EXPECTED_FILE_ID;
    }
  });
  return matched;
}

// ===================== UI state helpers =====================

function showError(message) {
  document.getElementById("checkingView").style.display = "none";
  document.getElementById("mainView").style.display = "none";
  document.getElementById("errorView").style.display = "block";
  document.getElementById("errorDetail").textContent = message;
}

function showMain() {
  document.getElementById("checkingView").style.display = "none";
  document.getElementById("errorView").style.display = "none";
  document.getElementById("mainView").style.display = "block";
}

function setResultMessage(text, isError) {
  const el = document.getElementById("resultMsg");
  el.textContent = text;
  el.className = isError ? "error" : "ok";
}

// ===================== Toast =====================

let toastTimer = null;

function showToast(message, type = "success", durationMs = 3000) {
  const toast = document.getElementById("toast");
  if (!toast) return;

  const iconSvg =
    type === "success"
      ? '<svg class="toast-icon" viewBox="0 0 16 16" fill="none"><path d="M6.5 11.5L3 8l1-1 2.5 2.5L12 4l1 1-6.5 6.5z" fill="currentColor"/></svg>'
      : '<svg class="toast-icon" viewBox="0 0 16 16" fill="none"><path d="M8 1a7 7 0 100 14A7 7 0 008 1zm.75 10.5h-1.5v-1.5h1.5v1.5zm0-3h-1.5V4.5h1.5v4z" fill="currentColor"/></svg>';

  toast.innerHTML = iconSvg + `<span>${message}</span>`;
  toast.className = `toast show ${type}`;

  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("show");
  }, durationMs);
}

// ===================== createLog() =====================

async function createLog() {
  const btn = document.getElementById("createLogBtn");
  btn.disabled = true;
  setResultMessage("Menambahkan log...", false);

  try {
    const userName = await getCurrentUserName();
    const now = new Date();

    const executionId = formatExecutionId(now);       // contoh: 260127-211305
    const executionStart = formatExecutionStart(now);  // contoh: 27/01/2026 21:13:05

    const rowValues = [
      executionId,
      executionStart,
      "Hello World",   // Sheet Category/Function
      "Hello world 2",  // Process Name
      userName,         // User
      "On going",       // Status
    ];

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(LOG_SHEET_NAME);
      const usedRange = sheet.getUsedRangeOrNullObject(true);
      usedRange.load("rowIndex, rowCount");
      await context.sync();

      const nextRowIndex = usedRange.isNullObject ? 0 : usedRange.rowIndex + usedRange.rowCount;
      const targetRange = sheet.getRangeByIndexes(nextRowIndex, 0, 1, rowValues.length);
      targetRange.values = [rowValues];
      await context.sync();
    });

    setResultMessage(`Log berhasil ditambahkan.\nExecution ID: ${executionId}`, false);
    showToast("Berhasil ditambahkan", "success");
  } catch (err) {
    setResultMessage("Gagal menambahkan log: " + err.message, true);
    showToast("Gagal menambahkan log", "error");
  } finally {
    btn.disabled = false;
  }
}

// ===================== Format tanggal =====================

function pad2(n) {
  return String(n).padStart(2, "0");
}

// yymmdd-hhmmss, contoh: 260127-211305
function formatExecutionId(date) {
  const yy = pad2(date.getFullYear() % 100);
  const mm = pad2(date.getMonth() + 1);
  const dd = pad2(date.getDate());
  const hh = pad2(date.getHours());
  const mi = pad2(date.getMinutes());
  const ss = pad2(date.getSeconds());
  return `${yy}${mm}${dd}-${hh}${mi}${ss}`;
}

// dd/mm/yyyy hh:mm:ss, contoh: 27/01/2026 21:13:05
function formatExecutionStart(date) {
  const dd = pad2(date.getDate());
  const mm = pad2(date.getMonth() + 1);
  const yyyy = date.getFullYear();
  const hh = pad2(date.getHours());
  const mi = pad2(date.getMinutes());
  const ss = pad2(date.getSeconds());
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

// ===================== Nama user (dengan fallback) =====================

const ROAMING_KEY_USERNAME = "automationLogUserName";

async function getCurrentUserName() {
  // 1) Coba SSO. Ini hanya akan berhasil kalau manifest sudah dikonfigurasi
  //    <WebApplicationInfo> + App Registration Azure AD sudah dibuat.
  //    Kalau belum, ini akan gagal secara diam-diam dan lanjut ke fallback.
  try {
    if (window.OfficeRuntime && OfficeRuntime.auth) {
      const tokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
      const claims = parseJwt(tokenEncoded);
      const name = claims.name || claims.preferred_username;
      if (name) return name;
    }
  } catch (e) {
    // SSO tidak tersedia/tidak dikonfigurasi -> lanjut ke fallback di bawah
  }

  // 2) Fallback: nama yang pernah diisi manual sebelumnya, tersimpan permanen di file ini
  const saved = Office.context.roamingSettings.get(ROAMING_KEY_USERNAME);
  if (saved) return saved;

  // 3) Fallback terakhir: minta user isi sekali, lalu simpan supaya tidak ditanya lagi
  const typed = window.prompt(
    "Nama Anda belum terdeteksi otomatis. Isi nama untuk kolom 'User':"
  );
  const finalName = (typed && typed.trim()) || "Unknown User";
  Office.context.roamingSettings.set(ROAMING_KEY_USERNAME, finalName);
  Office.context.roamingSettings.saveAsync();
  return finalName;
}

function parseJwt(token) {
  const base64Url = token.split(".")[1];
  const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
  const jsonPayload = decodeURIComponent(
    atob(base64)
      .split("")
      .map((c) => "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2))
      .join("")
  );
  return JSON.parse(jsonPayload);
}
