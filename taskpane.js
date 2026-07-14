/* global Office, Excel, OfficeRuntime */

// =============================================================================
// KONFIGURASI — SESUAIKAN
// =============================================================================

const EXPECTED_FILE_ID = "PROMO-PLAN-2026"; // ID unik file Promo Plan
const FILE_ID_PROPERTY_KEY = "PromoPlanFileId";
const LOG_SHEET_NAME = "Automation Log";

// =============================================================================
// INISIALISASI
// =============================================================================

Office.onReady(async (info) => {
  try {
    if (info.host !== Office.HostType.Excel) {
      showError("Add-in ini hanya didukung di aplikasi Excel.", {
        fileId: null,
        userName: null,
      });
      return;
    }

    const { matched, actualFileId } = await checkIsPromoPlanFile();
    const userName = await getCurrentUserName();

    if (!matched) {
      showError(
        `File yang sedang dibuka bukan file Promo Plan yang dikenali (ID "${FILE_ID_PROPERTY_KEY}" tidak cocok atau tidak ditemukan).`,
        { fileId: actualFileId, userName }
      );
      return;
    }

    showMain();
    document.getElementById("createLogBtn").addEventListener("click", createLog);
  } catch (err) {
    let userName = null;
    try {
      userName = await getCurrentUserName();
    } catch (e) {}
    showError("Terjadi error saat memuat add-in: " + err.message, {
      fileId: null,
      userName,
    });
  }
});

// =============================================================================
// VALIDASI FILE PROMO PLAN
// =============================================================================

async function checkIsPromoPlanFile() {
  let matched = false;
  let actualFileId = null;

  await Excel.run(async (context) => {
    const props = context.workbook.properties;
    props.load("custom");
    await context.sync();

    if (props.custom) {
      const prop = props.custom.getItemOrNullObject(FILE_ID_PROPERTY_KEY);
      prop.load("value");
      await context.sync();

      if (!prop.isNullObject) {
        actualFileId = String(prop.value);
        matched = actualFileId === EXPECTED_FILE_ID;
      }
    }

    // fallback: gunakan URL file aktif jika custom property tidak ada
    if (!actualFileId) {
      actualFileId = Office.context.document.url || "(tidak tersedia)";
    }
  });

  return { matched, actualFileId };
}

// =============================================================================
// UI STATE HELPERS
// =============================================================================

function showError(message, meta = {}) {
  document.getElementById("checkingView").style.display = "none";
  document.getElementById("mainView").style.display = "none";
  document.getElementById("errorView").style.display = "block";
  document.getElementById("errorDetail").textContent = message;

  const { fileId, userName } = meta;
  const metaEl = document.getElementById("errorMeta");
  const fileIdText = fileId ? fileId : "(tidak ditemukan)";
  const userNameText = userName ? userName : "(tidak diketahui)";

  metaEl.innerHTML = `
    File ID terdeteksi: <b>${escapeHtml(fileIdText)}</b><br/>
    User: <b>${escapeHtml(userNameText)}</b>
  `;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
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

// =============================================================================
// TOAST
// =============================================================================

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

// =============================================================================
// CREATE LOG
// =============================================================================

async function createLog() {
  const btn = document.getElementById("createLogBtn");
  btn.disabled = true;
  setResultMessage("Menambahkan log...", false);

  try {
    const userName = await getCurrentUserName();
    const now = new Date();

    const executionId = formatExecutionId(now);
    const executionStart = formatExecutionStart(now);

    const rowValues = [
      executionId,
      executionStart,
      "Hello World",
      "Hello world 2",
      userName,
      "On going",
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

// =============================================================================
// FORMAT TANGGAL
// =============================================================================

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatExecutionId(date) {
  const yy = pad2(date.getFullYear() % 100);
  const mm = pad2(date.getMonth() + 1);
  const dd = pad2(date.getDate());
  const hh = pad2(date.getHours());
  const mi = pad2(date.getMinutes());
  const ss = pad2(date.getSeconds());
  return `${yy}${mm}${dd}-${hh}${mi}${ss}`;
}

function formatExecutionStart(date) {
  const dd = pad2(date.getDate());
  const mm = pad2(date.getMonth() + 1);
  const yyyy = date.getFullYear();
  const hh = pad2(date.getHours());
  const mi = pad2(date.getMinutes());
  const ss = pad2(date.getSeconds());
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

// =============================================================================
// NAMA USER (SSO + FALLBACK)
// =============================================================================

const ROAMING_KEY_USERNAME = "automationLogUserName";

async function getCurrentUserName() {
  try {
    if (Office.context && Office.context.userProfile && Office.context.userProfile.displayName) {
      return Office.context.userProfile.displayName;
    }
  } catch (e) {}

  try {
    if (window.OfficeRuntime && OfficeRuntime.auth) {
      const tokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
      const claims = parseJwt(tokenEncoded);
      const name = claims.name || claims.preferred_username;
      if (name) return name;
    }
  } catch (e) {}

  const saved = Office.context.roamingSettings.get(ROAMING_KEY_USERNAME);
  if (saved) return saved;

  const typed = window.prompt("Nama Anda belum terdeteksi otomatis. Isi nama untuk kolom 'User':");
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
