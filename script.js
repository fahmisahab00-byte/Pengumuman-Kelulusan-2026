/* ============================================================
   KONFIGURASI — GANTI URL INI DENGAN URL GOOGLE APPS SCRIPT ANDA
   ============================================================ */
const CONFIG = {
  // Ganti dengan URL Web App dari Google Apps Script Anda
  APPS_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbweDAdlFDjwn58iT24oxEXEYwMVeV3wHS2S_8KQMzGeZAURfBUh1kgPKxe8bTa4fRd6RQ/exec",

  // Kredensial login (bisa diganti sesuai kebutuhan)
  ADMIN_USERNAME: "admin",
  ADMIN_PASSWORD: "12345",
  SISWA_PASSWORD: "123456",

  // Info sekolah (untuk SKL PDF)
  SCHOOL_NAME: "SMP NEGERI 9 BATANG",
  SCHOOL_ADDRESS: "Jl. Tentara Pelajar No.9, Batang, Jawa Tengah",
  KEPALA_SEKOLAH: "Nama Kepala Sekolah, S.Pd., M.M.",
  NIP_KS: "NIP. 19XXXXXXXXXXXXXXXXX",
  TAHUN_AJARAN: "2025/2026",
};

/* ============================================================
   STATE
   ============================================================ */
let allStudents = [];       // semua data dari Spreadsheet
let filteredStudents = [];  // data setelah pencarian
let importData = [];        // data dari Excel yang akan diimport

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

function showSection(id) {
  document.querySelectorAll(".admin-section").forEach(s => s.classList.remove("active"));
  document.querySelectorAll(".sidebar-link").forEach(l => l.classList.remove("active"));
  const sec = document.getElementById(id);
  if (sec) sec.classList.add("active");
  event.currentTarget.classList.add("active");
  // Jika section statistik, update stats
  if (id === "sec-stats") updateStats();
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
        <div class="info-row">
          <i class="bi bi-credit-card info-icon"></i>
          <div><div class="info-label">NISN</div><div class="info-value">${s.NISN || "-"}</div></div>
        </div>
        <div class="info-row">
          <i class="bi bi-geo-alt-fill info-icon"></i>
          <div><div class="info-label">Tempat, Tanggal Lahir</div><div class="info-value">${ttl}</div></div>
        </div>
        <div class="nilai-box">
          <div class="nilai-num">${cleanNumber(s.Nilai_Rata)}</div>
          <div class="nilai-label">Nilai Rata-rata</div>
        </div>
        ${isLulus ? `
        </button>` : `
        <div class="custom-alert alert-danger" style="display:flex;margin-top:8px">
          <i class="bi bi-exclamation-triangle-fill"></i> 
          Hubungi pihak sekolah untuk informasi lebih lanjut.
        </div>`}
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
    const ttl = `${s.Tempat_Lahir || "-"}, ${formatTanggal(s.Tanggal_Lahir)}`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><strong>${s.Nama || "-"}</strong></td>
      <td>${s.NISN || "-"}</td>
      <td>${ttl}</td>
      <td>${cleanNumber(s.Nilai_Rata)}</td>
      <td><span class="badge-status ${isLulus ? "badge-lulus" : "badge-tidak"}">${(s.Status || "–").toUpperCase()}</span></td>
      <td>
        
      </td>`;
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
   GENERATE SKL PDF (jsPDF)
   ============================================================ */
function generateSKL(s) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const W = 210, ml = 25, mr = 25, cw = W - ml - mr; // page width, margins, content width

  // ----- HEADER / KOP SURAT -----
  // Garis atas tebal
  doc.setLineWidth(1.2);
  doc.setDrawColor(26, 79, 145);
  doc.line(ml, 18, W - mr, 18);

  // Logo placeholder (circle)
  doc.setFillColor(26, 79, 145);
  doc.circle(ml + 14, 30, 13, "F");
  doc.setFontSize(7); doc.setTextColor(255,255,255); doc.setFont("helvetica","bold");
  doc.text("LOGO", ml + 14, 30, { align:"center" });
  doc.text("SEKOLAH", ml + 14, 33.5, { align:"center" });

  // Teks kop
  doc.setTextColor(26, 79, 145);
  doc.setFont("helvetica","bold"); doc.setFontSize(13);
  doc.text("PEMERINTAH KABUPATEN BATANG", W / 2, 22, { align:"center" });
  doc.setFontSize(11); doc.text("DINAS PENDIDIKAN DAN KEBUDAYAAN", W / 2, 27.5, { align:"center" });
  doc.setFontSize(16); doc.text(CONFIG.SCHOOL_NAME, W / 2, 34.5, { align:"center" });
  doc.setFontSize(8.5); doc.setFont("helvetica","normal"); doc.setTextColor(80,80,80);
  doc.text(CONFIG.SCHOOL_ADDRESS, W / 2, 39.5, { align:"center" });

  // Garis bawah kop
  doc.setLineWidth(0.4); doc.setDrawColor(26,79,145);
  doc.line(ml, 43, W - mr, 43);
  doc.setLineWidth(1.5); doc.line(ml, 44.5, W - mr, 44.5);

  // ----- JUDUL -----
  doc.setFont("helvetica","bold"); doc.setFontSize(15); doc.setTextColor(0);
  doc.text("SURAT KETERANGAN LULUS", W / 2, 56, { align:"center" });
  doc.setFontSize(10); doc.setFont("helvetica","normal"); doc.setTextColor(80,80,80);
  doc.text(`Nomor: 422/${(s.NISN||"000").slice(-4)}/SMP09/2026`, W / 2, 62, { align:"center" });

  // ----- TUBUH -----
  const y0 = 72;
  doc.setFontSize(11); doc.setTextColor(30);
  doc.setFont("helvetica","normal");

  const intro = `Yang bertanda tangan di bawah ini, Kepala ${CONFIG.SCHOOL_NAME}, dengan ini menerangkan bahwa peserta didik yang tersebut di bawah ini:`;
  const lines = doc.splitTextToSize(intro, cw);
  doc.text(lines, ml, y0);

  // ----- TABEL IDENTITAS -----
  let ty = y0 + (lines.length * 5.5) + 4;
  const fields = [
    ["Nama Lengkap",       s.Nama || "-"],
    ["NISN",               s.NISN || "-"],
    ["Tempat, Tgl Lahir",  `${s.Tempat_Lahir || "-"}, ${formatTanggal(s.Tanggal_Lahir)}`],
    ["Nilai Rata-rata",    cleanNumber(s.Nilai_Rata)],
    ["Tahun Ajaran",       CONFIG.TAHUN_AJARAN],
  ];

  doc.setFillColor(240, 246, 255);
  doc.roundedRect(ml, ty - 4, cw, fields.length * 9 + 5, 3, 3, "F");

  fields.forEach(([label, val]) => {
    doc.setFont("helvetica","bold"); doc.setFontSize(10.5); doc.setTextColor(40,40,40);
    doc.text(label, ml + 6, ty);
    doc.setFont("helvetica","normal"); doc.setTextColor(0);
    doc.text(":", ml + 54, ty);
    doc.text(String(val), ml + 58, ty);
    ty += 9;
  });

  // ----- PERNYATAAN LULUS -----
  ty += 6;
  doc.setFont("helvetica","normal"); doc.setFontSize(11); doc.setTextColor(30);
  const stmt = `Telah dinyatakan LULUS dari Satuan Pendidikan SMP Negeri 9 Batang berdasarkan hasil Rapat Dewan Guru dan sesuai dengan Peraturan yang berlaku.`;
  const stmtLines = doc.splitTextToSize(stmt, cw);
  doc.text(stmtLines, ml, ty);
  ty += stmtLines.length * 5.5;

  // Kotak STATUS LULUS
  ty += 6;
  doc.setFillColor(22, 163, 74);
  doc.roundedRect((W - 100) / 2, ty, 100, 14, 4, 4, "F");
  doc.setTextColor(255); doc.setFont("helvetica","bold"); doc.setFontSize(13);
  doc.text("✓  DINYATAKAN LULUS", W / 2, ty + 9.5, { align:"center" });

  // ----- TANGGAL & TTD -----
  ty += 24;
  const today = new Date();
  const bln = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
  const tglStr = `Batang, ${today.getDate()} ${bln[today.getMonth()]} ${today.getFullYear()}`;
  doc.setTextColor(30); doc.setFont("helvetica","normal"); doc.setFontSize(10.5);
  doc.text(tglStr, W - mr, ty, { align:"right" });

  ty += 6;
  doc.text("Kepala Sekolah,", W - mr, ty, { align:"right" });
  ty += 28; // ruang TTD
  doc.setFont("helvetica","bold");
  doc.text(CONFIG.KEPALA_SEKOLAH, W - mr, ty, { align:"right" });
  ty += 5;
  doc.setFont("helvetica","normal");
  doc.text(CONFIG.NIP_KS, W - mr, ty, { align:"right" });

  // ----- GARIS BAWAH -----
  doc.setLineWidth(0.4); doc.setDrawColor(200);
  doc.line(ml, 285, W - mr, 285);
  doc.setFontSize(7.5); doc.setTextColor(150);
  doc.text("Dokumen ini dicetak secara elektronik dan sah tanpa tanda tangan basah.", W / 2, 288, { align:"center" });

  // ----- SAVE -----
  const safeName = (s.Nama || "siswa").replace(/[^a-z0-9]/gi, "_");
  doc.save(`SKL_${safeName}_${s.NISN || ""}.pdf`);
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
