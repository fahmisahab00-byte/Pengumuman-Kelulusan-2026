/* ============================================================
   KONFIGURASI — GANTI URL INI DENGAN URL GOOGLE APPS SCRIPT ANDA
   ============================================================ */
const CONFIG = {
  // Ganti dengan URL Web App dari Google Apps Script Anda
  APPS_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbyZxrkhKCSF8hMK1pgyMWtFcjvzo2Q-95gUTBHumRZ1EKJGPnVnJXPOsmN0rqgKVSXRdA/exec",

  // Kredensial login (bisa diganti sesuai kebutuhan)
  ADMIN_USERNAME: "admin",
  ADMIN_PASSWORD: "12345",
  SISWA_PASSWORD: "123456",

  // Info sekolah (untuk SKL PDF)
  SCHOOL_NAME: "SMP NEGERI 9 BATANG",
  SCHOOL_ADDRESS: "Jalan R.E. Martadinata Gg. Kakap Merah Karangsari Karangasem Utara Batang Kode Pos 51216",
  KEPALA_SEKOLAH: "Mas’ud, S.Pd.",
  NIP_KS: "NIP. 19660510 199702 1 002",
  TAHUN_AJARAN: "2025/2026",
};

/* ============================================================
   STATE
   ============================================================ */
let allStudents = [];       // semua data dari Spreadsheet
let filteredStudents = [];  // data setelah pencarian
let importData = [];        // data dari Excel yang akan diimport

/* Template SKL custom */
let sklTemplate = {
  file: null,          // File object Word asli
  arrayBuffer: null,   // ArrayBuffer untuk pemrosesan docx
  htmlPreview: null,   // HTML string hasil konversi mammoth.js
  name: "",            // nama file
  size: "",            // ukuran file
};

/* ============================================================
   NAVIGASI HALAMAN
   ============================================================ */
function showPage(id) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  const target = document.getElementById(id);
  if (target) {
    target.classList.add("active");
    window.scrollTo(0, 0);
  }
}

function showSection(id, el) {
  document.querySelectorAll(".admin-section").forEach(s => s.classList.remove("active"));
  document.querySelectorAll(".sidebar-link").forEach(l => l.classList.remove("active"));
  const sec = document.getElementById(id);
  if (sec) sec.classList.add("active");
  if (el) el.classList.add("active");
  if (id === "sec-stats") updateStats();
  if (id === "sec-skl") refreshSKLStatus();
}

/* ============================================================
   UTILITY
   ============================================================ */
function togglePw(inputId, btn) {
  const input = document.getElementById(inputId);
  if (!input) return;
  const isHidden = input.type === "password";
  input.type = isHidden ? "text" : "password";
  btn.innerHTML = isHidden ? '<i class="bi bi-eye-slash"></i>' : '<i class="bi bi-eye"></i>';
}

function showAlert(elId, msg, type = "danger") {
  const el = document.getElementById(elId);
  if (!el) return;
  el.className = `custom-alert alert-${type}`;
  el.innerHTML = `<i class="bi bi-${type === "danger" ? "exclamation-triangle-fill" : type === "success" ? "check-circle-fill" : "info-circle-fill"}"></i> ${msg}`;
  el.classList.remove("d-none");
  setTimeout(() => el.classList.add("d-none"), 5000);
}

function setLoading(btnTextId, spinnerId, isLoading, text = "") {
  const btnText = document.getElementById(btnTextId);
  const spinner = document.getElementById(spinnerId);
  if (!btnText || !spinner) return;
  if (isLoading) {
    btnText.textContent = "Memproses...";
    spinner.classList.remove("d-none");
  } else {
    btnText.textContent = text;
    spinner.classList.add("d-none");
  }
}

function formatTanggal(tgl) {
  if (!tgl) return "-";
  // Coba parse berbagai format tanggal
  const d = new Date(tgl);
  if (!isNaN(d)) {
    const bulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
    return `${d.getDate()} ${bulan[d.getMonth()]} ${d.getFullYear()}`;
  }
  return tgl;
}

function cleanNumber(val) {
  if (val === null || val === undefined || val === "") return "-";
  const num = parseFloat(String(val).replace(",", "."));
  return isNaN(num) ? val : num.toFixed(2);
}

/* ============================================================
   LOGIN SISWA
   ============================================================ */
async function loginSiswa() {
  const nisn = document.getElementById("siswa-nisn").value.trim();
  const pw   = document.getElementById("siswa-pw").value.trim();
  const alertEl = "siswa-alert";

  if (!nisn) { showAlert(alertEl, "NISN tidak boleh kosong."); return; }
  if (!pw)   { showAlert(alertEl, "Password tidak boleh kosong."); return; }
  if (pw !== CONFIG.SISWA_PASSWORD) {
    showAlert(alertEl, "Password salah. Password default: 123456");
    return;
  }

  setLoading("siswa-btn-text", "siswa-spinner", true);

  try {
    const siswa = await fetchSiswaNISN(nisn);
    if (!siswa) {
      showAlert(alertEl, `Data dengan NISN <strong>${nisn}</strong> tidak ditemukan. Pastikan NISN benar.`);
    } else {
      tampilkanHasilSiswa(siswa);
    }
  } catch (err) {
    showAlert(alertEl, `Gagal terhubung ke server. Pastikan URL Apps Script sudah dikonfigurasi. <br><small>${err.message}</small>`);
  }

  setLoading("siswa-btn-text", "siswa-spinner", false, "Cek Kelulusan");
}

/* ============================================================
   LOGIN ADMIN
   ============================================================ */
function loginAdmin() {
  const user = document.getElementById("admin-user").value.trim();
  const pw   = document.getElementById("admin-pw").value.trim();
  const alertEl = "admin-alert";

  if (!user) { showAlert(alertEl, "Username tidak boleh kosong."); return; }
  if (!pw)   { showAlert(alertEl, "Password tidak boleh kosong."); return; }

  if (user !== CONFIG.ADMIN_USERNAME || pw !== CONFIG.ADMIN_PASSWORD) {
    showAlert(alertEl, "Username atau password salah.");
    return;
  }

  showPage("page-dashboard");
  loadAllData();
}

/* ============================================================
   FETCH DATA DARI GOOGLE SHEETS (via Apps Script)
   ============================================================ */
async function fetchSiswaNISN(nisn) {
  const url = `${CONFIG.APPS_SCRIPT_URL}?action=getByNISN&nisn=${encodeURIComponent(nisn)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json.status === "error") throw new Error(json.message || "Error dari server");
  return json.data || null;
}

async function fetchAllData() {
  const url = `${CONFIG.APPS_SCRIPT_URL}?action=getAll`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json.status === "error") throw new Error(json.message || "Error dari server");
  return json.data || [];
}

async function postData(rows) {
  const res = await fetch(CONFIG.APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action: "addRows", rows }),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json.status === "error") throw new Error(json.message || "Error saat menyimpan data");
  return json;
}

/* ============================================================
   TAMPILKAN HASIL SISWA
   ============================================================ */
function tampilkanHasilSiswa(s) {
  const isLulus = (s.Status || "").toUpperCase() === "LULUS";
  const ttl = `${s.Tempat_Lahir || "-"}, ${formatTanggal(s.Tanggal_Lahir)}`;
  const nilaiRata = cleanNumber(s.Nilai_Rata);
  const nilaiAda  = nilaiRata !== "-";

  const html = `
    <div class="hasil-card">
      <div class="hasil-header">
        <img src="https://i.imgur.com/SoMVyZx.png" alt="Logo" class="hasil-school-logo"
             onerror="this.src='https://via.placeholder.com/64/1a4f91/ffffff?text=L'"/>
        <h4>SMP NEGERI 9 BATANG</h4>
        <h2>Pengumuman Kelulusan</h2>
        <p style="color:rgba(255,255,255,.75);font-size:.85rem;margin-top:4px">Tahun Ajaran ${CONFIG.TAHUN_AJARAN}</p>
        <div class="status-badge ${isLulus ? "lulus" : "tidak"}">
          <i class="bi bi-${isLulus ? "patch-check-fill" : "x-circle-fill"}"></i>
          ${isLulus ? "DINYATAKAN LULUS" : "TIDAK LULUS"}
        </div>
      </div>
      <div class="hasil-body">
        <div class="info-row">
          <i class="bi bi-person-fill info-icon"></i>
          <div><div class="info-label">Nama Lengkap</div><div class="info-value">${s.Nama || "-"}</div></div>
        </div>
        ${s.NIS ? `
        <div class="info-row">
          <i class="bi bi-person-badge info-icon"></i>
          <div><div class="info-label">NIS</div><div class="info-value">${s.NIS}</div></div>
        </div>` : ""}
        <div class="info-row">
          <i class="bi bi-credit-card info-icon"></i>
          <div><div class="info-label">NISN</div><div class="info-value">${s.NISN || "-"}</div></div>
        </div>
        <div class="info-row">
          <i class="bi bi-geo-alt-fill info-icon"></i>
          <div><div class="info-label">Tempat, Tanggal Lahir</div><div class="info-value">${ttl}</div></div>
        </div>
        ${nilaiAda ? `
        <div class="nilai-box">
          <div class="nilai-num">${nilaiRata}</div>
          <div class="nilai-label">Nilai Rata-rata</div>
        </div>` : ""}
        ${!isLulus ? `
        <div class="custom-alert alert-danger" style="display:flex;margin-top:8px">
          <i class="bi bi-exclamation-triangle-fill"></i>
          Hubungi pihak sekolah untuk informasi lebih lanjut.
        </div>` : ""}
      </div>
    </div>`;

  document.getElementById("hasil-content").innerHTML = html;
  showPage("page-hasil-siswa");
}

/* ============================================================
   LOAD ALL DATA (ADMIN)
   ============================================================ */
async function loadAllData() {
  const loading = document.getElementById("data-loading");
  const alertEl = "data-alert";
  const tbody   = document.getElementById("tbl-body");
  const empty   = document.getElementById("table-empty");
  const summary = document.getElementById("tbl-summary");

  loading.classList.remove("d-none");
  empty.classList.add("d-none");
  summary.classList.add("d-none");
  tbody.innerHTML = "";

  try {
    allStudents = await fetchAllData();
    filteredStudents = [...allStudents];
    renderTable(filteredStudents);
    updateStats();
  } catch (err) {
    showAlert(alertEl, `Gagal memuat data: ${err.message}. Pastikan URL Apps Script sudah diatur di CONFIG.`);
    empty.classList.remove("d-none");
  }

  loading.classList.add("d-none");
}

function renderTable(data) {
  const tbody = document.getElementById("tbl-body");
  const empty = document.getElementById("table-empty");
  const summary = document.getElementById("tbl-summary");

  tbody.innerHTML = "";

  if (!data.length) {
    empty.classList.remove("d-none");
    summary.classList.add("d-none");
    return;
  }

  empty.classList.add("d-none");

  data.forEach((s, i) => {
    const isLulus = (s.Status || "").toUpperCase() === "LULUS";
    const ttl     = `${s.Tempat_Lahir || "-"}, ${formatTanggal(s.Tanggal_Lahir)}`;
    const nilai   = cleanNumber(s.Nilai_Rata);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><strong>${s.Nama || "-"}</strong></td>
      <td>${s.NIS  || "-"}</td>
      <td>${s.NISN || "-"}</td>
      <td>${ttl}</td>
      <td>${nilai !== "-" ? nilai : '<span style="color:var(--gray-400)">–</span>'}</td>
      <td><span class="badge-status ${isLulus ? "badge-lulus" : "badge-tidak"}">${(s.Status || "–").toUpperCase()}</span></td>`;
    tbody.appendChild(tr);
  });

  const lulusCount = data.filter(s => (s.Status||"").toUpperCase() === "LULUS").length;
  summary.innerHTML = `
    <span>Total: <strong>${data.length}</strong> siswa</span>
    <span>Lulus: <strong style="color:var(--green)">${lulusCount}</strong></span>
    <span>Tidak Lulus: <strong style="color:var(--red)">${data.length - lulusCount}</strong></span>`;
  summary.classList.remove("d-none");
}

function filterTable() {
  const q = document.getElementById("search-input").value.toLowerCase();
  filteredStudents = allStudents.filter(s =>
    (s.Nama  || "").toLowerCase().includes(q) ||
    (s.NISN  || "").toLowerCase().includes(q)
  );
  renderTable(filteredStudents);
}

function updateStats() {
  const total = allStudents.length;
  const lulus = allStudents.filter(s => (s.Status||"").toUpperCase() === "LULUS").length;
  const tidak = total - lulus;
  const avg   = total
    ? (allStudents.reduce((acc, s) => acc + (parseFloat(s.Nilai_Rata) || 0), 0) / total).toFixed(2)
    : "–";
  const pct   = total ? Math.round((lulus / total) * 100) : 0;

  document.getElementById("stat-total").textContent = total || "–";
  document.getElementById("stat-lulus").textContent = lulus || "–";
  document.getElementById("stat-tidak").textContent = tidak || "–";
  document.getElementById("stat-avg").textContent   = total ? avg : "–";
  document.getElementById("pct-lulus").textContent  = `${pct}%`;
  document.getElementById("stat-bar-fill").style.width = `${pct}%`;
}

/* ============================================================
   IMPORT EXCEL (SheetJS)
   ============================================================ */
function previewExcel(input) {
  const file = input.files[0];
  if (!file) return;

  document.getElementById("file-name-label").textContent = file.name;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      importData = rows.map(r => ({
        Nama:         r["Nama"] || r["nama"] || "",
        NISN:         String(r["NISN"] || r["nisn"] || ""),
        Tempat_Lahir: r["Tempat_Lahir"] || r["Tempat Lahir"] || r["tempat_lahir"] || "",
        Tanggal_Lahir: formatTanggalFromExcel(r["Tanggal_Lahir"] || r["Tanggal Lahir"] || r["tanggal_lahir"] || ""),
        Nilai_Rata:   r["Nilai_Rata"] || r["Nilai Rata"] || r["nilai_rata"] || "",
        Status:       (r["Status"] || r["status"] || "").toUpperCase(),
      }));

      // Preview 10 baris pertama
      const preview = importData.slice(0, 10);
      const headers = ["Nama", "NISN", "Tempat_Lahir", "Tanggal_Lahir", "Nilai_Rata", "Status"];
      let tableHTML = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr></thead><tbody>`;
      preview.forEach(r => {
        tableHTML += `<tr>${headers.map(h => `<td>${r[h] || "–"}</td>`).join("")}</tr>`;
      });
      tableHTML += "</tbody>";

      document.getElementById("preview-table").innerHTML = tableHTML;
      document.getElementById("preview-count").textContent = `${importData.length} baris data`;
      document.getElementById("preview-table-wrap").classList.remove("d-none");
    } catch (err) {
      showAlert("import-alert", `Gagal membaca file Excel: ${err.message}`);
    }
  };
  reader.readAsBinaryString(file);
}

function formatTanggalFromExcel(val) {
  if (!val) return "";
  if (val instanceof Date) return val.toISOString().split("T")[0];
  return String(val);
}

async function importToSheets() {
  if (!importData.length) { showAlert("import-alert", "Tidak ada data untuk diimport."); return; }

  document.getElementById("import-loading").classList.remove("d-none");
  document.getElementById("import-alert").classList.add("d-none");

  try {
    const result = await postData(importData);
    showAlert("import-alert", `✅ Berhasil mengirim <strong>${importData.length}</strong> data ke Google Spreadsheet.`, "success");
    importData = [];
    document.getElementById("preview-table-wrap").classList.add("d-none");
    document.getElementById("file-name-label").textContent = "Belum ada file dipilih";
    document.getElementById("excel-file").value = "";
    // Refresh data tabel admin
    loadAllData();
  } catch (err) {
    showAlert("import-alert", `Gagal mengirim data: ${err.message}`);
  }

  document.getElementById("import-loading").classList.add("d-none");
}

function downloadTemplate() {
  const ws = XLSX.utils.aoa_to_sheet([
    ["Nama", "NISN", "Tempat_Lahir", "Tanggal_Lahir", "Nilai_Rata", "Status"],
    ["Ahmad Fauzi",    "0123456789", "Batang", "2010-05-12", "85.50", "LULUS"],
    ["Budi Santoso",   "0123456790", "Pekalongan", "2010-03-20", "72.30", "LULUS"],
    ["Citra Dewi",     "0123456791", "Semarang", "2010-07-08", "60.00", "TIDAK"],
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "DataSiswa");
  XLSX.writeFile(wb, "template_data_siswa_smpn9batang.xlsx");
}

/* ============================================================
   TEMPLATE SKL CUSTOM
   ============================================================ */

function uploadTemplateSKL(input) {
  const file = input.files[0];
  if (!file) return;

  const isDocx = file.name.toLowerCase().endsWith(".docx") ||
    file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  if (!isDocx) {
    showAlert("skl-upload-alert", "File harus berformat Word (.docx).");
    return;
  }
  if (file.size > 10 * 1024 * 1024) {
    showAlert("skl-upload-alert", "Ukuran file maksimal 10MB.");
    return;
  }
  document.getElementById("skl-file-label").textContent = `Membaca file: ${file.name}…`;

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const arrayBuffer = e.target.result;

      // Simpan ke sessionStorage SEBELUM mammoth menggunakan buffer
      // (mammoth bisa men-detach buffer pada beberapa browser)
      let base64 = "";
      try {
        base64 = arrayBufferToBase64(arrayBuffer);
        sessionStorage.setItem("skl_template_base64", base64);
        sessionStorage.setItem("skl_template_name", file.name);
        sessionStorage.setItem("skl_template_size", formatFileSize(file.size));
      } catch (storageErr) {
        console.warn("sessionStorage penuh atau tidak tersedia:", storageErr);
      }

      // Clone buffer sebelum diberikan ke mammoth agar tidak detached
      const bufferForMammoth = arrayBuffer.slice(0);
      const result = await mammoth.convertToHtml({ arrayBuffer: bufferForMammoth });

      // Simpan buffer ASLI (fresh clone dari base64 agar selalu valid)
      sklTemplate.file = file;
      sklTemplate.arrayBuffer = base64 ? base64ToArrayBuffer(base64) : arrayBuffer.slice(0);
      sklTemplate.htmlPreview = result.value;
      sklTemplate.name = file.name;
      sklTemplate.size = formatFileSize(file.size);

      try {
        sessionStorage.setItem("skl_template_html", sklTemplate.htmlPreview);
      } catch (e) {}

      document.getElementById("skl-file-label").textContent = `✅ ${file.name} berhasil dibaca`;
      showAlert("skl-upload-alert", `Template Word <strong>${file.name}</strong> berhasil diupload! Siswa dapat mendownload SKL sebagai PDF.`, "success");
      refreshSKLStatus();
      tampilkanReviewTemplate();
    } catch (err) {
      showAlert("skl-upload-alert", `Gagal membaca file Word: ${err.message}`);
      document.getElementById("skl-file-label").textContent = "Gagal membaca file";
    }
  };
  reader.readAsArrayBuffer(file);
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(2) + " MB";
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function tampilkanReviewTemplate() {
  const wrap       = document.getElementById("skl-review-wrap");
  const previewDiv = document.getElementById("skl-word-preview");
  if (!wrap || !previewDiv) return;

  const html = sklTemplate.htmlPreview || sessionStorage.getItem("skl_template_html");
  if (!html) return;

  previewDiv.innerHTML = `<div class="word-preview-content">${html}</div>`;
  wrap.classList.remove("d-none");
}

function refreshSKLStatus() {
  const name    = sklTemplate.name    || sessionStorage.getItem("skl_template_name")    || "";
  const size    = sklTemplate.size    || sessionStorage.getItem("skl_template_size")    || "";
  const hasData = !!(sklTemplate.arrayBuffer || sessionStorage.getItem("skl_template_base64"));
  const card    = document.getElementById("skl-status-card");
  const icon    = document.getElementById("skl-status-icon");
  const nameEl  = document.getElementById("skl-status-name");
  const sizeEl  = document.getElementById("skl-status-size");
  const actions = document.getElementById("skl-status-actions");
  if (hasData && name) {
    if (card)    card.classList.add("has-template");
    if (icon)    icon.innerHTML = '<i class="bi bi-file-earmark-word-fill"></i>';
    if (nameEl)  nameEl.textContent = name;
    if (sizeEl)  sizeEl.textContent = size;
    if (actions) actions.innerHTML = `
      <button class="btn-preview-siswa" onclick="tampilkanReviewTemplate()"><i class="bi bi-eye"></i> Lihat Template</button>
      <button class="btn-hapus-template" onclick="hapusTemplate()"><i class="bi bi-trash3-fill"></i> Hapus</button>`;
    tampilkanReviewTemplate();
  } else {
    if (card)    card.classList.remove("has-template");
    if (icon)    icon.innerHTML = '<i class="bi bi-file-earmark-word"></i>';
    if (nameEl)  nameEl.textContent = "Belum ada template — menggunakan template default sistem";
    if (sizeEl)  sizeEl.textContent = "";
    if (actions) actions.innerHTML = "";
    const wrap = document.getElementById("skl-review-wrap");
    if (wrap) wrap.classList.add("d-none");
  }
}

function hapusTemplate() {
  if (!confirm("Yakin ingin menghapus template SKL? Sistem akan kembali menggunakan template default.")) return;
  sklTemplate = { file: null, arrayBuffer: null, htmlPreview: null, name: "", size: "" };
  sessionStorage.removeItem("skl_template_base64");
  sessionStorage.removeItem("skl_template_name");
  sessionStorage.removeItem("skl_template_size");
  sessionStorage.removeItem("skl_template_html");
  document.getElementById("skl-file-label").textContent = "Belum ada file dipilih";
  document.getElementById("skl-pdf-file").value = "";
  refreshSKLStatus();
  const al = document.getElementById("skl-upload-alert");
  if (al) {
    al.className = "custom-alert alert-info";
    al.innerHTML = '<i class="bi bi-info-circle-fill"></i> Template berhasil dihapus. Sistem menggunakan template default.';
    al.classList.remove("d-none");
    setTimeout(() => al.classList.add("d-none"), 4000);
  }
}

/* ============================================================
   MODAL PREVIEW SKL
   ============================================================ */
function showPreviewModal() {
  document.getElementById("modal-preview").classList.remove("d-none");
}

function closeModalPreview(e) {
  if (!e || e.target === document.getElementById("modal-preview")) {
    document.getElementById("modal-preview").classList.add("d-none");
  }
}

function previewSKLModal() {
  // Ambil TTL, pisahkan tempat dan tanggal lahir
  const ttlRaw = (document.getElementById("prev-ttl") || {}).value || "Batang, 12 Mei 2010";
  const ttlParts = ttlRaw.split(",").map(x => x.trim());
  const siswaData = {
    Nama:          (document.getElementById("prev-nama") || {}).value  || "Ahmad Fauzi",
    NIS:           (document.getElementById("prev-nis")  || {}).value  || "1234",
    NISN:          (document.getElementById("prev-nisn") || {}).value  || "0123456789",
    Tempat_Lahir:  ttlParts[0] || "Batang",
    Tanggal_Lahir: ttlParts.slice(1).join(",").trim() || "12 Mei 2010",
    Nilai_Rata:    "85.50",
    Status:        "LULUS",
  };
  const hasTemplate = !!(sklTemplate.arrayBuffer || sessionStorage.getItem("skl_template_base64"));
  closeModalPreview();
  if (hasTemplate) {
    generateSKLDariTemplate(siswaData, true);
  } else {
    generateSKL(siswaData, true);
  }
}

/* ============================================================
   GENERATE SKL — DENGAN TEMPLATE CUSTOM (WORD .docx → PDF)
   Pendekatan: baca template Word, replace placeholder,
   render HTML via mammoth.js, lalu export PDF via html2pdf.js
   ============================================================ */
async function generateSKLDariTemplate(s, previewOnly = false) {
  // Selalu buat fresh ArrayBuffer dari base64 untuk menghindari detached buffer
  let arrayBuffer = null;
  const base64 = sessionStorage.getItem("skl_template_base64");
  if (base64) {
    arrayBuffer = base64ToArrayBuffer(base64);
  } else if (sklTemplate.arrayBuffer) {
    arrayBuffer = sklTemplate.arrayBuffer.slice(0);
  }
  if (!arrayBuffer) { generateSKL(s, previewOnly); return; }

  try {
    const today = new Date();
    const bln   = ["Januari","Februari","Maret","April","Mei","Juni",
                   "Juli","Agustus","September","Oktober","November","Desember"];
    const tglCetak = `${today.getDate()} ${bln[today.getMonth()]} ${today.getFullYear()}`;
    const ttl      = `${s.Tempat_Lahir || "-"}, ${formatTanggal(s.Tanggal_Lahir)}`;

    // Map placeholder → nilai pengganti
    const values = {
      "{{NAMA}}":    s.Nama   || "-",
      "{{NIS}}":     s.NIS    || "-",
      "{{NISN}}":    s.NISN   || "-",
      "{{TTL}}":     ttl,
      "{{NILAI}}":   cleanNumber(s.Nilai_Rata) !== "-" ? cleanNumber(s.Nilai_Rata) : "-",
      "{{STATUS}}":  (s.Status || "-").toUpperCase(),
      "{{TANGGAL}}": tglCetak,
    };

    // 1. Replace placeholder di DOCX
    const uint8  = new Uint8Array(arrayBuffer);
    const result = await replaceInDocx(uint8, values);

    // 2. Konversi DOCX yang sudah diisi → HTML via mammoth.js
    // Gunakan slice(0) untuk clone agar buffer tidak detached
    const mammothResult = await mammoth.convertToHtml({ arrayBuffer: result instanceof ArrayBuffer ? result.slice(0) : result });
    const htmlContent   = mammothResult.value;

    // 3. Bangun HTML cetak yang sempurna — layout 1:1 dengan Word
    const safeName = (s.Nama || "siswa").replace(/[^a-z0-9]/gi, "_");
    const fileName = `SKL_${safeName}_${s.NISN || s.NIS || ""}`;

    // CSS untuk meniru layout Word: margin standar A4 (2.5cm kiri-kanan, 2cm atas-bawah)
    // @page memastikan margin PDF persis seperti Word ketika di-print
    const printCss = `
      @page {
        size: A4 portrait;
        margin: 2cm 2.5cm;
      }
      * { box-sizing: border-box; }
      body {
        font-family: 'Times New Roman', Times, serif;
        font-size: 12pt;
        line-height: 1.5;
        color: #000;
        background: #fff;
        margin: 0;
        padding: 0;
      }
      p { margin: 0 0 6pt 0; }
      table { border-collapse: collapse; width: 100%; margin: 6pt 0; }
      td, th { padding: 3pt 6pt; font-size: 11pt; vertical-align: top; }
      h1, h2, h3 { text-align: center; margin: 4pt 0; }
      strong, b { font-weight: bold; }
      u { text-decoration: underline; }
      img { max-width: 100%; height: auto; }
      /* Mammoth menghasilkan class ini untuk alignment */
      .docx-center, p[style*="text-align: center"] { text-align: center !important; }
      .docx-right,  p[style*="text-align: right"]  { text-align: right !important;  }
      /* Jangan tampilkan elemen non-cetak */
      .no-print { display: none !important; }
    `;

    // HTML untuk ditampilkan di window baru (preview + print ke PDF)
    const printHtml = `<!DOCTYPE html>
<html lang="id"><head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>SKL - ${s.Nama || "Siswa"}</title>
  <style>${printCss}</style>
  <style media="screen">
    /* Tampilan screen: tunjukkan kertas A4 */
    html { background: #888; }
    body {
      max-width: 21cm;
      margin: 1cm auto;
      padding: 2cm 2.5cm;
      background: #fff;
      box-shadow: 0 0 12px rgba(0,0,0,.4);
      min-height: 29.7cm;
    }
    .toolbar {
      position: fixed; top: 0; left: 0; right: 0;
      background: #1a4f91; color: #fff;
      padding: 10px 20px;
      display: flex; align-items: center; gap: 12px;
      z-index: 999; font-family: sans-serif; font-size: 14px;
      box-shadow: 0 2px 8px rgba(0,0,0,.3);
    }
    .toolbar button {
      background: #fff; color: #1a4f91; border: none;
      padding: 7px 18px; border-radius: 5px; font-weight: bold;
      cursor: pointer; font-size: 14px;
    }
    .toolbar button:hover { background: #e0ecff; }
    .toolbar .hint { font-size: 12px; opacity: .85; }
    body { margin-top: 60px; }
  </style>
</head><body>
  <div class="toolbar no-print">
    <button onclick="window.print()">⬇ Simpan / Print PDF</button>
    <span class="hint">Pilih "Save as PDF" di dialog print browser Anda &nbsp;|&nbsp; Nama file: ${fileName}.pdf</span>
  </div>
  ${htmlContent}
</body></html>`;

    if (previewOnly) {
      // Mode preview — buka dengan banner preview saja
      const previewHtml = printHtml.replace(
        '<button onclick="window.print()">⬇ Simpan / Print PDF</button>',
        '<span style="font-weight:bold">👁 MODE PREVIEW</span>'
      ).replace(
        `Nama file: ${fileName}.pdf`,
        "Ini adalah tampilan preview — data siswa sudah terisi"
      );
      const win = window.open("", "_blank");
      if (win) { win.document.write(previewHtml); win.document.close(); }
    } else {
      // Mode download — buka window baru, langsung tampilkan dengan tombol print
      const win = window.open("", "_blank");
      if (win) {
        win.document.write(printHtml);
        win.document.close();
        // Fokus ke window baru agar mudah di-print
        win.focus();
      } else {
        alert("Pop-up diblokir browser!\nIzinkan pop-up untuk situs ini, lalu coba lagi.\n\nAtau: buka Pengaturan Browser → Izin → Pop-up → Izinkan situs ini.");
      }
    }

  } catch (err) {
    console.error("Gagal generate dari template Word:", err);
    alert("Gagal memproses template Word. Menggunakan template default.\n\nError: " + err.message);
    generateSKL(s, previewOnly);
  }
}

/* ============================================================
   REPLACE PLACEHOLDER DI DALAM DOCX (ZIP/XML)
   Menggunakan DecompressionStream / manual ZIP parsing
   Fallback: replace langsung di binary string XML
   ============================================================ */
async function replaceInDocx(uint8, values) {
  // Konversi docx ke string untuk mencari & replace di XML
  // DOCX = ZIP, berisi word/document.xml
  // Kita akan coba unzip secara manual (ZIP local file header)

  try {
    // Gunakan JSZip jika tersedia (di-load via CDN)
    if (typeof JSZip !== "undefined") {
      return await replaceDocxWithJSZip(uint8, values);
    }
    // Fallback: replace langsung di raw bytes (teks XML tidak ter-encode)
    return replaceDocxRaw(uint8, values);
  } catch (e) {
    return replaceDocxRaw(uint8, values);
  }
}

/* Fallback: replace placeholder langsung di raw XML bytes */
function replaceDocxRaw(uint8, values) {
  // Decode sebagai binary string
  let str = "";
  for (let i = 0; i < uint8.length; i++) str += String.fromCharCode(uint8[i]);

  // Replace setiap placeholder di raw XML
  Object.entries(values).forEach(([ph, val]) => {
    // Escape karakter XML
    const safeVal = String(val)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
    // Replace versi langsung
    str = str.split(ph).join(safeVal);
    // Replace versi XML-encoded (kadang { dan } menjadi entitas)
    str = str.split(ph.replace(/\{/g, "&#123;").replace(/\}/g, "&#125;")).join(safeVal);
  });

  // Konversi kembali ke Uint8Array
  const result = new Uint8Array(str.length);
  for (let i = 0; i < str.length; i++) result[i] = str.charCodeAt(i) & 0xFF;
  return result.buffer;
}

/* Replace menggunakan JSZip */
async function replaceDocxWithJSZip(uint8, values) {
  const zip = await JSZip.loadAsync(uint8);
  const xmlFile = zip.file("word/document.xml");
  if (!xmlFile) throw new Error("word/document.xml tidak ditemukan");

  let xml = await xmlFile.async("string");

  // Placeholder bisa tersebar di beberapa <w:r> runs dalam satu <w:p>
  // Strategi: gabungkan teks dalam paragraf, replace, lalu tulis kembali
  // Simple approach: replace langsung (works jika placeholder tidak terpecah)
  Object.entries(values).forEach(([ph, val]) => {
    const safeVal = String(val)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
    xml = xml.split(ph).join(safeVal);
  });

  zip.file("word/document.xml", xml);
  const out = await zip.generateAsync({ type: "arraybuffer" });
  return out;
}

/* ============================================================
   GENERATE SKL PDF — Format Resmi SMP Negeri 9 Batang
   Sesuai template KEPUTUSAN_KEPALA: kop, dasar hukum,
   identitas siswa, pernyataan status, tanda tangan.
   ============================================================ */
function generateSKL(s, previewOnly = false) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

  const W  = 210;
  const ml = 25;
  const mr = 18;
  const cw = W - ml - mr;

  const bln = ["Januari","Februari","Maret","April","Mei","Juni",
               "Juli","Agustus","September","Oktober","November","Desember"];
  const today    = new Date();
  const tglCetak = `${today.getDate()} ${bln[today.getMonth()]} ${today.getFullYear()}`;
  const NOMOR_SK = "R/108/400.3.11.1/VI/2025";

  // ── Logo base64 ──────────────────────────────────────────────
  const LOGO1_B64 = "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";
  const LOGO2_B64 = "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";

  // ════════════════════════════════════════════════════════════
  // 1. KOP SURAT
  // ════════════════════════════════════════════════════════════
  const logoW = 18, logoH = 17; // mm

  // Logo kiri (Pemda Batang)
  try {
    doc.addImage("data:image/png;base64," + LOGO1_B64, "PNG", ml, 5, logoW, logoH);
  } catch(e) {
    doc.setFillColor(150,0,0); doc.circle(ml + logoW/2, 5 + logoH/2, logoW/2, "F");
  }

  // Logo kanan (Kemdikbud)
  try {
    doc.addImage("data:image/png;base64," + LOGO2_B64, "PNG", W - mr - logoW, 5, logoW - 1, logoH);
  } catch(e) {
    doc.setFillColor(26,79,145); doc.circle(W - mr - logoW/2, 5 + logoH/2, logoW/2, "F");
  }

  // Teks kop — tengah
  doc.setTextColor(0);
  doc.setFont("helvetica","bold"); doc.setFontSize(10);
  doc.text("PEMERINTAH KABUPATEN BATANG", W/2, 9, { align:"center" });
  doc.setFontSize(15);
  doc.text("SMP NEGERI 9 BATANG", W/2, 16, { align:"center" });
  doc.setFont("helvetica","normal"); doc.setFontSize(8);
  doc.text("Jalan Tampangsono Batang Telepon (0285) 391296 Kode Pos 51214", W/2, 21, { align:"center" });
  doc.text("Laman : smpn9batang@yahoo.co.id  ~  Blognet : smpn9batang.wordpress.com", W/2, 25, { align:"center" });

  // Garis kop tebal + tipis
  const kopY = 28;
  doc.setDrawColor(0); doc.setLineWidth(1.0);
  doc.line(ml, kopY, W - mr, kopY);
  doc.setLineWidth(0.3);
  doc.line(ml, kopY + 1.2, W - mr, kopY + 1.2);

  // ════════════════════════════════════════════════════════════
  // 2. JUDUL SURAT
  // ════════════════════════════════════════════════════════════
  let y = kopY + 7;
  doc.setFont("helvetica","bold"); doc.setFontSize(12); doc.setTextColor(0);
  doc.text("KEPUTUSAN KEPALA SMP NEGERI 9 BATANG", W/2, y, { align:"center" }); y += 5.5;
  doc.setFontSize(11);
  doc.text(`No. ${NOMOR_SK}`, W/2, y, { align:"center" }); y += 5;
  doc.text("Tentang", W/2, y, { align:"center" }); y += 5;
  doc.setFontSize(12);
  doc.text("KELULUSAN PESERTA DIDIK SMP NEGERI 9 BATANG", W/2, y, { align:"center" }); y += 5.5;
  doc.text(`TAHUN AJARAN ${CONFIG.TAHUN_AJARAN}`, W/2, y, { align:"center" }); y += 7;

  // ════════════════════════════════════════════════════════════
  // 3. TABEL DASAR HUKUM
  // ════════════════════════════════════════════════════════════
  doc.setFont("helvetica","normal"); doc.setFontSize(10.5); doc.setTextColor(0);

  const dasar = [
    "Peraturan Menteri Pendidikan dan Kebudayaan Nomor 21 Tahun 2022 tentang Standar Penilaian Pendidikan pada Pendidikan Anak Usia Dini, Jenjang Pendidikan Dasar, dan Jenjang Pendidikan Menengah;",
    "Peraturan Menteri Pendidikan, Kebudayaan, Riset dan Teknologi Nomor 58 Tahun 2024 tentang Ijazah Jenjang Pendidikan Dasar dan Pendidikan Menengah;",
    "Permendikbudristek nomor 12 tahun 2024 dan Panduan Pembelajaran Dan Asesmen Edisi Revisi Tahun 2024 tentang penentuan kelulusan;",
    "Hasil Rapat Koordinasi Musyawarah Kerja Kepala Sekolah Kabupaten Batang tanggal 14 Mei 2025;",
    `Peraturan Akademik ${CONFIG.SCHOOL_NAME} Tahun Ajaran ${CONFIG.TAHUN_AJARAN};`,
    `Hasil Keputusan Rapat Pleno Dewan Guru ${CONFIG.SCHOOL_NAME} tanggal 2 Juni 2025.`,
  ];

  const cLabel = ml;         // x kolom "Dasar"
  const cSep   = ml + 13;    // x titik dua
  const cNo    = ml + 17;    // x nomor (1. 2. dst)
  const cIsi   = ml + 23;    // x isi teks
  const maxIsi = W - mr - cIsi;
  const lh     = 4.8;        // line height

  // Label "Dasar :"
  doc.setFont("helvetica","normal");
  doc.text("Dasar", cLabel, y);
  doc.text(":", cSep, y);

  dasar.forEach((teks, i) => {
    doc.text(`${i+1}.`, cNo, y);
    const lines = doc.splitTextToSize(teks, maxIsi);
    doc.text(lines, cIsi, y);
    y += lines.length * lh + 1;
  });

  y += 3;

  // ════════════════════════════════════════════════════════════
  // 4. MEMUTUSKAN + IDENTITAS
  // ════════════════════════════════════════════════════════════
  doc.setFont("helvetica","normal"); doc.setFontSize(11);
  doc.text("Memutuskan", W/2, y, { align:"center" }); y += 6;
  doc.text("Menetapkan : Bahwa peserta didik di bawah ini :", ml, y); y += 6;

  const ttl = `${s.Tempat_Lahir || "-"}, ${formatTanggal(s.Tanggal_Lahir)}`;
  const idCLabel = ml;
  const idCSep   = ml + 55;
  const idCVal   = ml + 60;
  const idRows = [
    ["Nama",                    s.Nama  || "-",  true ],
    ["Tempat dan Tanggal Lahir", ttl,             false],
    ["NIS",                     s.NIS   || "-",  false],
    ["NISN",                    s.NISN  || "-",  false],
  ];

  idRows.forEach(([label, val, bold]) => {
    doc.setFont("helvetica","normal"); doc.setFontSize(11); doc.setTextColor(0);
    doc.text(label, idCLabel, y);
    doc.text(":", idCSep, y);
    if (bold) doc.setFont("helvetica","bold");
    doc.text(String(val), idCVal, y);
    y += 6;
  });

  y += 3;

  // ════════════════════════════════════════════════════════════
  // 5. PERNYATAAN STATUS
  // ════════════════════════════════════════════════════════════
  const isLulus = (s.Status || "").toUpperCase() === "LULUS";
  const statusTeks = isLulus ? "LULUS" : "TIDAK LULUS";

  doc.setFont("helvetica","normal"); doc.setFontSize(11); doc.setTextColor(0);
  doc.text("Sesuai dengan keputusan dewan guru, siswa dengan identitas tersebut dinyatakan :", ml, y);
  y += 7;

  doc.setFont("helvetica","bold"); doc.setFontSize(18);
  doc.text(statusTeks, W/2, y, { align:"center" });
  y += 8;

  doc.setFont("helvetica","normal"); doc.setFontSize(11);
  doc.text(`dari Satuan Pendidikan ${CONFIG.SCHOOL_NAME} Tahun Ajaran ${CONFIG.TAHUN_AJARAN}`, ml, y);
  y += 6;
  doc.text("Demikian surat keputusan ini dibuat agar dapat digunakan sebagaimana mestinya.", ml, y);
  y += 8;

  // ════════════════════════════════════════════════════════════
  // 6. TANDA TANGAN
  // ════════════════════════════════════════════════════════════
  const ttdX = W - mr - 58;
  doc.setFont("helvetica","normal"); doc.setFontSize(11); doc.setTextColor(0);
  doc.text(`Batang, 2 Juni 2025`, ttdX, y);
  y += 5;
  doc.text("Kepala Sekolah,", ttdX, y);
  y += 22; // ruang tanda tangan

  doc.setFont("helvetica","bold"); doc.setFontSize(11);
  doc.text(CONFIG.KEPALA_SEKOLAH, ttdX, y);
  y += 5;
  doc.setFont("helvetica","normal"); doc.setFontSize(10);
  doc.text(CONFIG.NIP_KS, ttdX, y);

  // ════════════════════════════════════════════════════════════
  // 7. SIMPAN / PREVIEW
  // ════════════════════════════════════════════════════════════
  const safeName = (s.Nama || "siswa").replace(/[^a-z0-9]/gi, "_");
  if (previewOnly) {
    window.open(URL.createObjectURL(doc.output("blob")), "_blank");
  } else {
    doc.save(`SKL_${safeName}_${s.NISN || s.NIS || ""}.pdf`);
  }
}

/* ============================================================
   WRAPPER generateSKL — cek template custom dulu
   ============================================================ */
function cetakSKL(s) {
  // Cek template dari sessionStorage (lebih andal, tidak terhapus saat navigasi halaman)
  const hasTemplate = !!(sessionStorage.getItem("skl_template_base64") || sklTemplate.arrayBuffer);
  if (hasTemplate) {
    generateSKLDariTemplate(s, false);
  } else {
    generateSKL(s, false);
  }
}

/* ============================================================
   ENTER KEY SUPPORT
   ============================================================ */
document.addEventListener("keydown", (e) => {
  if (e.key !== "Enter") return;
  const active = document.querySelector(".page.active");
  if (!active) return;
  if (active.id === "page-login-siswa") loginSiswa();
  if (active.id === "page-login-admin") loginAdmin();
});
