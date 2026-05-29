// =============================================================
// DRIVER MONITORING EYEWAKE - WEBSITE SCRIPT
// File ini menangani Firebase, form driver, dashboard realtime,
// visual garis pulse, histori microsleep, alarm, dan export Excel.
// =============================================================

// ------------------ KONFIGURASI FIREBASE ------------------
// Menyimpan konfigurasi Firebase yang sama dengan proyek ESP32.
const firebaseConfig = {
  // API key Firebase web app.
  apiKey: "AIzaSyCxGUL6WfZzzn3W_phwlMn6kNChyo7lmDA",
  // URL Realtime Database yang dipakai ESP32.
  databaseURL: "https://driver-monitoring-6e4df-default-rtdb.asia-southeast1.firebasedatabase.app/",
};

// Menginisialisasi Firebase app pada browser.
firebase.initializeApp(firebaseConfig);

// Membuat referensi Realtime Database agar mudah dipakai di bawah.
const db = firebase.database();

// ------------------ STATE APLIKASI ------------------
// Menyimpan data monitoring terakhir dari Firebase.
let currentData = {};

// Menyimpan status terakhir agar alarm hanya bunyi saat transisi ke Microsleep.
let lastStatus = "Normal";

// Menyimpan daftar histori microsleep yang tampil di tabel dan Excel.
let microsleepData = [];

// Menandai listener monitoring agar tidak dibuat berulang.
let monitoringListenerStarted = false;

// Menandai listener histori agar tidak dibuat berulang.
let historyListenerStarted = false;

// Menandai listener pulse_live agar tidak dibuat berulang.
let pulseListenerStarted = false;

// Menyimpan gender aktif sesuai value yang dibaca ESP32: pria/perempuan.
let activeGender = "pria";

// Menyimpan nama driver aktif dari form.
let activeName = "";

// Menyimpan waktu terakhir website menerima update dari alat ESP32.
let lastDeviceUpdateAt = 0;

// Menyimpan status koneksi browser ke Firebase Realtime Database.
let firebaseBrowserConnected = false;

// Batas waktu maksimal tanpa update sebelum alat dianggap tidak terhubung.
const DEVICE_STALE_MS = 10000;

// Menyimpan data rekaman Mode Pengujian untuk diunduh sebagai Excel.
let testData = [];

// Menyimpan interval timer Mode Pengujian.
let testTimer = null;

// Menyimpan jumlah detik pengujian berjalan.
let testSeconds = 0;

// Menandai apakah panel Mode Pengujian sedang terbuka.
let isTestOpen = false;

// Menandai apakah Mode Pengujian sedang merekam data.
let isTestRunning = false;

// ------------------ ELEMENT HTML ------------------
// Mengambil elemen jam realtime.
const clockEl = document.getElementById("clock");

// Mengambil pembungkus status koneksi alat.
const deviceStatusEl = document.getElementById("deviceStatus");

// Mengambil teks utama status koneksi alat.
const deviceStatusTextEl = document.getElementById("deviceStatusText");

// Mengambil detail kecil status koneksi alat.
const deviceStatusDetailEl = document.getElementById("deviceStatusDetail");

// Mengambil kartu setup driver.
const setupCardEl = document.getElementById("setupCard");

// Mengambil area dashboard monitoring.
const monitoringAreaEl = document.getElementById("monitoringArea");

// Mengambil form driver.
const driverFormEl = document.getElementById("driverForm");

// Mengambil input nama driver.
const driverNameEl = document.getElementById("driverName");

// Mengambil select jenis kelamin.
const driverGenderEl = document.getElementById("driverGender");

// Mengambil teks catatan rule.
const ruleNoteEl = document.getElementById("ruleNote");

// Mengambil teks nama driver aktif.
const activeDriverNameEl = document.getElementById("activeDriverName");

// Mengambil teks jenis kelamin aktif.
const activeDriverGenderEl = document.getElementById("activeDriverGender");

// Mengambil teks rule BPM aktif.
const activeRuleEl = document.getElementById("activeRule");

// Mengambil tombol ganti driver.
const changeDriverButtonEl = document.getElementById("changeDriverButton");

// Mengambil kartu BPM agar warnanya bisa berubah saat BPM masuk kondisi microsleep.
const bpmCardEl = document.getElementById("bpmCard");

// Mengambil elemen nilai BPM.
const bpmEl = document.getElementById("bpm");

// Mengambil elemen info pulse.
const pulseInfoEl = document.getElementById("pulseInfo");

// Mengambil garis pulse merah.
const pulseLineEl = document.getElementById("pulseLine");

// Mengambil kartu kemiringan agar warnanya bisa berubah saat kepala menunduk/melewati batas.
const tiltCardEl = document.getElementById("tiltCard");

// Mengambil elemen kemiringan.
const tiltEl = document.getElementById("tilt");

// Mengambil kartu mata agar warnanya bisa berubah saat mata tertutup.
const eyeCardEl = document.getElementById("eyeCard");

// Mengambil elemen mata.
const eyeEl = document.getElementById("eye");

// Mengambil elemen status.
const statusEl = document.getElementById("status");

// Mengambil kartu status.
const statusCardEl = document.getElementById("statusCard");

// Mengambil elemen durasi realtime.
const liveDurationEl = document.getElementById("liveDuration");

// Mengambil body tabel histori microsleep.
const microTableBodyEl = document.getElementById("microTableBody");

// Mengambil tombol download Excel.
const downloadMicrosleepButtonEl = document.getElementById("downloadMicrosleepButton");

// Mengambil tombol hapus histori.
const clearMicrosleepButtonEl = document.getElementById("clearMicrosleepButton");

// Mengambil popup alert.
const popupEl = document.getElementById("popup");

// Mengambil audio alarm.
const alarmEl = document.getElementById("alarm");

// Mengambil tombol buka/tutup Mode Pengujian.
const btnTestEl = document.getElementById("btnTest");

// Mengambil panel Mode Pengujian.
const testModeEl = document.getElementById("testMode");

// Mengambil input nama pada Mode Pengujian.
const testNameEl = document.getElementById("nama");

// Mengambil tombol Start Mode Pengujian.
const startTestButtonEl = document.getElementById("startTestButton");

// Mengambil tombol Stop Mode Pengujian.
const stopTestButtonEl = document.getElementById("stopTestButton");

// Mengambil elemen timer Mode Pengujian.
const timerEl = document.getElementById("timer");

// ------------------ HELPER RULE ------------------
// Fungsi ini mengubah value gender menjadi label tampilan Indonesia.
function getGenderLabel(gender) {
  // Mengembalikan label perempuan jika gender adalah perempuan.
  if (gender === "perempuan") return "Perempuan";
  // Mengembalikan label laki-laki untuk value pria/default.
  return "Laki-laki";
}

// Fungsi ini mengambil threshold BPM dari jenis kelamin.
function getBpmThreshold(gender) {
  // Threshold perempuan sesuai rule penelitian dari ESP32 adalah < 63.
  if (gender === "perempuan") return 63;
  // Threshold laki-laki/pria sesuai rule penelitian dari ESP32 adalah < 65.
  return 65;
}

// Fungsi ini membuat teks rule BPM yang mudah dibaca.
function getRuleText(gender) {
  // Mengambil threshold berdasarkan gender.
  const threshold = getBpmThreshold(gender);
  // Mengembalikan teks rule untuk ditampilkan di dashboard.
  return `Microsleep jika BPM < ${threshold}`;
}

// Fungsi ini memastikan nilai kosong tampil sebagai strip.
function safeText(value, fallback = "-") {
  // Mengembalikan fallback jika value undefined, null, atau string kosong.
  if (value === undefined || value === null || value === "") return fallback;
  // Mengembalikan value asli jika tersedia.
  return value;
}

// Fungsi ini memformat BPM dari Firebase.
function formatBpm(value) {
  // Menampilkan No Signal jika ESP32 mengirim false.
  if (value === false) return "No Signal";
  // Menampilkan strip jika tidak ada data.
  if (value === undefined || value === null || value === "") return "-";
  // Mengembalikan angka/text BPM jika tersedia.
  return value;
}

// Fungsi ini memformat durasi microsleep dari beberapa kemungkinan field Firebase.
function formatDuration(data) {
  // Menggunakan teks durasi dari ESP32 jika tersedia.
  if (data && data.microsleep_duration_text) return data.microsleep_duration_text;
  // Menggunakan detik jika field seconds tersedia.
  if (data && typeof data.microsleep_duration_sec === "number") return `${data.microsleep_duration_sec} detik`;
  // Mengubah millisecond menjadi detik jika hanya field ms yang tersedia.
  if (data && typeof data.microsleep_duration_ms === "number") return `${Math.round(data.microsleep_duration_ms / 1000)} detik`;
  // Mengembalikan nol detik untuk default.
  return "0 detik";
}

// Fungsi ini mengubah teks apa pun menjadi huruf kecil agar pengecekan status lebih aman.
function normalizeText(value) {
  // Mengubah undefined/null menjadi string kosong agar tidak error saat diproses.
  if (value === undefined || value === null) return "";
  // Mengubah value menjadi string, menghapus spasi tepi, lalu membuat huruf kecil.
  return String(value).trim().toLowerCase();
}

// Fungsi ini mengecek apakah suatu status sensor bernilai Microsleep.
function isMicrosleepStatus(value) {
  // Mengembalikan true jika teks status sama dengan microsleep.
  return normalizeText(value) === "microsleep";
}

// Fungsi ini mengambil angka valid dari data Firebase.
function toNumberOrNull(value) {
  // Mengembalikan null jika value adalah false karena ESP32 memakai false saat BPM belum valid.
  if (value === false) return null;
  // Mengubah value menjadi angka.
  const numberValue = Number(value);
  // Mengembalikan null jika hasilnya bukan angka valid.
  if (Number.isNaN(numberValue)) return null;
  // Mengembalikan angka valid.
  return numberValue;
}

// Fungsi ini menyalakan atau mematikan warna merah pada kartu sensor.
function setSensorCardAlert(cardElement, isAlert) {
  // Menghentikan fungsi jika elemen kartu tidak ditemukan.
  if (!cardElement) return;
  // Menambahkan class sensor-alert saat sensor masuk kondisi bahaya.
  cardElement.classList.toggle("sensor-alert", Boolean(isAlert));
}

// Fungsi ini membuat teks waktu update terakhir yang mudah dibaca.
function getLastSeenText() {
  // Jika belum pernah ada data alat, kembalikan keterangan menunggu data.
  if (!lastDeviceUpdateAt) return "Menunggu data dari ESP32.";
  // Menghitung selisih detik dari update terakhir.
  const secondsAgo = Math.max(0, Math.round((Date.now() - lastDeviceUpdateAt) / 1000));
  // Mengembalikan teks update terakhir dalam detik.
  return `Update terakhir ${secondsAgo} detik lalu.`;
}

// ------------------ UI DASAR ------------------
// Fungsi ini memperbarui jam realtime di header.
function updateClock() {
  // Mengisi elemen jam dengan tanggal dan waktu lokal browser.
  clockEl.textContent = new Date().toLocaleString("id-ID");
}

// Fungsi ini menampilkan popup sementara di pojok kanan atas.
function showPopup(message, isDanger = true) {
  // Mengisi teks popup.
  popupEl.textContent = message;
  // Mengubah warna popup sesuai jenis pesan.
  popupEl.style.background = isDanger ? "linear-gradient(135deg, #ef4444, #991b1b)" : "linear-gradient(135deg, #2563eb, #0891b2)";
  // Menampilkan popup.
  popupEl.style.display = "block";
  // Menyembunyikan popup setelah tiga detik.
  setTimeout(() => {
    // Menutup popup.
    popupEl.style.display = "none";
  }, 3000);
}

// Fungsi ini mengubah tampilan catatan rule pada form.
function updateRuleNote() {
  // Mengambil gender yang dipilih user.
  const gender = driverGenderEl.value;
  // Mengisi catatan rule sesuai gender.
  ruleNoteEl.textContent = `${getGenderLabel(gender)}: ${getRuleText(gender)}.`;
}

// Fungsi ini menampilkan dashboard dan menyembunyikan form setup.
function showMonitoringArea() {
  // Menyembunyikan kartu setup.
  setupCardEl.classList.add("hidden");
  // Menampilkan area monitoring.
  monitoringAreaEl.classList.remove("hidden");
}

// Fungsi ini menampilkan form setup untuk mengganti driver.
function showSetupArea() {
  // Menampilkan kartu setup.
  setupCardEl.classList.remove("hidden");
  // Menyembunyikan area monitoring.
  monitoringAreaEl.classList.add("hidden");
}

// Fungsi ini memperbarui strip informasi driver aktif.
function updateDriverStrip() {
  // Mengisi nama driver aktif.
  activeDriverNameEl.textContent = activeName || "-";
  // Mengisi jenis kelamin aktif.
  activeDriverGenderEl.textContent = getGenderLabel(activeGender);
  // Mengisi rule aktif berdasarkan gender.
  activeRuleEl.textContent = getRuleText(activeGender);
}

// Fungsi ini memperbarui indikator hijau/merah koneksi alat.
function updateDeviceStatus() {
  // Mengecek apakah browser masih terhubung ke Firebase.
  if (!firebaseBrowserConnected) {
    // Mengubah kartu status menjadi mode merah.
    deviceStatusEl.className = "device-status disconnected";
    // Mengisi teks utama ketika Firebase tidak terhubung.
    deviceStatusTextEl.textContent = "Alat Tidak Terhubung";
    // Mengisi detail bahwa koneksi Firebase/browser bermasalah.
    deviceStatusDetailEl.textContent = "Firebase offline atau internet terputus.";
    // Menghentikan fungsi setelah status merah diperbarui.
    return;
  }

  // Mengecek apakah data alat masih segar berdasarkan update terakhir.
  const deviceFresh = lastDeviceUpdateAt > 0 && Date.now() - lastDeviceUpdateAt <= DEVICE_STALE_MS;

  // Mengubah kartu status menjadi hijau jika data alat masih masuk.
  deviceStatusEl.className = `device-status ${deviceFresh ? "connected" : "disconnected"}`;

  // Mengisi teks utama sesuai kondisi alat.
  deviceStatusTextEl.textContent = deviceFresh ? "Alat Terhubung" : "Alat Tidak Terhubung";

  // Mengisi detail update terakhir dari ESP32.
  deviceStatusDetailEl.textContent = deviceFresh ? getLastSeenText() : `Tidak ada data baru. ${getLastSeenText()}`;
}

// Fungsi ini menandai bahwa data terbaru dari alat baru saja diterima.
function markDeviceDataReceived() {
  // Menyimpan waktu browser saat snapshot Firebase berubah.
  lastDeviceUpdateAt = Date.now();
  // Memperbarui status koneksi alat.
  updateDeviceStatus();
}

// ------------------ FIREBASE SETTINGS DRIVER ------------------
// Fungsi ini menyimpan nama dan gender ke Firebase sebelum monitoring dimulai.
async function saveDriverSettings(name, gender) {
  // Menyimpan nama dan gender ke variabel global.
  activeName = name;
  // Menyimpan gender aktif ke variabel global.
  activeGender = gender;
  // Menulis nama driver ke path yang dibaca ESP32.
  await db.ref("settings/driver/name").set(name);
  // Menulis jenis kelamin ke path yang dibaca ESP32.
  await db.ref("settings/driver/gender").set(gender);
  // Memastikan threshold laki-laki tetap 65 sesuai rule.
  await db.ref("settings/rules/bpm/male_threshold").set(65);
  // Memastikan threshold perempuan tetap 63 sesuai rule.
  await db.ref("settings/rules/bpm/female_threshold").set(63);
  // Menyimpan nama ke localStorage agar input otomatis terisi saat reload.
  localStorage.setItem("eyewake_driver_name", name);
  // Menyimpan gender ke localStorage agar input otomatis terisi saat reload.
  localStorage.setItem("eyewake_driver_gender", gender);
}

// Fungsi ini menjalankan proses mulai monitoring dari form.
async function handleDriverSubmit(event) {
  // Mencegah form melakukan reload halaman.
  event.preventDefault();
  // Mengambil nama driver dari input.
  const name = driverNameEl.value.trim();
  // Mengambil gender driver dari select.
  const gender = driverGenderEl.value;
  // Mengecek nama kosong.
  if (!name) {
    // Menampilkan pesan jika nama belum diisi.
    showPopup("Nama driver wajib diisi.");
    // Menghentikan fungsi.
    return;
  }
  // Mencoba menyimpan data ke Firebase.
  try {
    // Menunggu proses simpan settings selesai.
    await saveDriverSettings(name, gender);
    // Memperbarui strip driver.
    updateDriverStrip();
    // Menampilkan dashboard.
    showMonitoringArea();
    // Menyalakan semua listener realtime.
    startRealtimeListeners();
    // Menampilkan pesan sukses.
    showPopup("Monitoring dimulai.", false);
  } catch (error) {
    // Menampilkan error di console untuk debugging.
    console.error(error);
    // Menampilkan pesan gagal kepada user.
    showPopup("Gagal menyimpan data driver ke Firebase.");
  }
}

// Fungsi ini mengisi form dari localStorage bila ada.
function loadSavedDriverForm() {
  // Mengambil nama lama dari localStorage.
  const savedName = localStorage.getItem("eyewake_driver_name") || "";
  // Mengambil gender lama dari localStorage.
  const savedGender = localStorage.getItem("eyewake_driver_gender") || "pria";
  // Mengisi input nama.
  driverNameEl.value = savedName;
  // Mengisi pilihan gender.
  driverGenderEl.value = savedGender;
  // Memperbarui catatan rule sesuai gender tersimpan.
  updateRuleNote();
}

// ------------------ LISTENER REALTIME ------------------
// Fungsi ini menyalakan semua listener Firebase yang dibutuhkan website.
function startRealtimeListeners() {
  // Menyalakan listener data monitoring.
  startMonitoringListener();
  // Menyalakan listener data pulse_live.
  startPulseLiveListener();
  // Menyalakan listener histori microsleep.
  startMicrosleepHistoryListener();
}

// Fungsi ini menyalakan pengecekan status koneksi Firebase dan alat ESP32.
function startConnectionStatusWatcher() {
  // Mendengarkan status koneksi browser ke Firebase Realtime Database.
  db.ref(".info/connected").on("value", (snapshot) => {
    // Menyimpan status koneksi Firebase ke variabel global.
    firebaseBrowserConnected = snapshot.val() === true;
    // Memperbarui indikator hijau/merah setelah status Firebase berubah.
    updateDeviceStatus();
  });

  // Mendengarkan timestamp monitoring agar indikator bisa aktif meskipun dashboard belum dibuka.
  db.ref("monitoring/firebase_timestamp").on("value", (snapshot) => {
    // Mengambil timestamp dari ESP32.
    const timestamp = snapshot.val();
    // Menandai data alat diterima jika timestamp tersedia.
    if (timestamp !== null && timestamp !== undefined && timestamp !== "") markDeviceDataReceived();
  });

  // Mengecek ulang setiap detik agar indikator otomatis merah saat alat berhenti update.
  setInterval(updateDeviceStatus, 1000);
}

// Fungsi ini mendengarkan data realtime dari path /monitoring.
function startMonitoringListener() {
  // Menghentikan fungsi jika listener sudah aktif.
  if (monitoringListenerStarted) return;
  // Menandai listener monitoring sudah aktif.
  monitoringListenerStarted = true;
  // Membaca path monitoring secara realtime.
  db.ref("monitoring").on("value", (snapshot) => {
    // Mengambil isi snapshot Firebase.
    const data = snapshot.val();
    // Menangani kondisi data kosong.
    if (!data) {
      // Mengisi status No Data.
      statusEl.textContent = "No Data";
      // Menghentikan callback.
      return;
    }
    // Menandai bahwa data baru dari alat sudah diterima.
    markDeviceDataReceived();
    // Menyimpan data terakhir ke variabel global.
    currentData = data;
    // Merender data monitoring ke dashboard.
    renderMonitoring(data);
  });
}

// Fungsi ini mendengarkan data visual denyut dari path /pulse_live.
function startPulseLiveListener() {
  // Menghentikan fungsi jika listener sudah aktif.
  if (pulseListenerStarted) return;
  // Menandai listener pulse sudah aktif.
  pulseListenerStarted = true;
  // Membaca path pulse_live yang dikirim ESP32 untuk visualizer.
  db.ref("pulse_live").on("value", (snapshot) => {
    // Mengambil data pulse_live.
    const pulse = snapshot.val();
    // Menghentikan callback jika data kosong.
    if (!pulse) return;
    // Merender garis pulse berdasarkan data pulse_live.
    renderPulseLine(pulse);
  });
}

// Fungsi ini mendengarkan histori microsleep dari path /history_microsleep.
function startMicrosleepHistoryListener() {
  // Menghentikan fungsi jika listener sudah aktif.
  if (historyListenerStarted) return;
  // Menandai listener histori sudah aktif.
  historyListenerStarted = true;
  // Membaca 30 data terakhir agar tabel tidak terlalu berat.
  db.ref("history_microsleep").limitToLast(30).on("value", (snapshot) => {
    // Mengosongkan array sebelum diisi ulang.
    const rows = [];
    // Melakukan loop setiap child histori.
    snapshot.forEach((child) => {
      // Mengambil isi data child.
      const data = child.val();
      // Menambahkan key Firebase ke data.
      data.key = child.key;
      // Menampilkan hanya data dengan status Microsleep.
      if (data.status === "Microsleep") rows.push(data);
    });
    // Mengurutkan data terbaru ke paling atas.
    rows.sort((a, b) => String(b.firebase_timestamp || b.key).localeCompare(String(a.firebase_timestamp || a.key)));
    // Menyimpan data untuk export Excel.
    microsleepData = rows;
    // Merender tabel histori microsleep.
    renderMicrosleepTable(rows);
  });
}

// ------------------ RENDER DASHBOARD ------------------
// Fungsi ini menampilkan data monitoring realtime ke kartu dashboard.
function renderMonitoring(data) {
  // Mengambil status sistem dari Firebase.
  const status = data.status || "-";
  // Mengisi nilai BPM.
  bpmEl.textContent = formatBpm(data.bpm);
  // Mengisi nilai kemiringan kepala.
  tiltEl.textContent = safeText(data.kemiringan);
  // Mengisi kondisi mata.
  eyeEl.textContent = safeText(data.mata);
  // Mengisi status sistem.
  statusEl.textContent = status;
  // Mengambil gender dari monitoring jika ESP32 sudah mengirimnya.
  activeGender = data.driver_gender || activeGender;
  // Mengambil nama driver dari monitoring jika ESP32 sudah mengirimnya.
  activeName = data.driver_name || activeName;
  // Mengambil threshold dari monitoring atau memakai threshold berdasarkan gender.
  const threshold = data.bpm_rule_threshold || getBpmThreshold(activeGender);
  // Mengambil angka BPM agar bisa dicek terhadap threshold.
  const bpmNumber = toNumberOrNull(data.bpm);
  // Mengambil angka kemiringan agar bisa dicek terhadap batas 10 derajat.
  const tiltNumber = toNumberOrNull(data.kemiringan);
  // Mengecek apakah sensor BPM mendeteksi kondisi microsleep berdasarkan status ESP32 atau threshold web.
  const bpmAlert = isMicrosleepStatus(data.status_pulse) || isMicrosleepStatus(data.status_pulse_system) || (bpmNumber !== null && bpmNumber > 0 && bpmNumber < Number(threshold));
  // Mengecek apakah sensor kemiringan melewati batas rule kepala, yaitu minimal 10 derajat.
  const tiltAlert = isMicrosleepStatus(data.status_head) || (tiltNumber !== null && tiltNumber >= 10);
  // Mengecek apakah sensor mata membaca kondisi tertutup atau status eye microsleep.
  const eyeAlert = isMicrosleepStatus(data.status_eye) || normalizeText(data.mata).includes("tertutup");
  // Mengubah kotak BPM menjadi merah jika sensor BPM masuk kondisi microsleep.
  setSensorCardAlert(bpmCardEl, bpmAlert);
  // Mengubah kotak kemiringan menjadi merah jika kemiringan kepala minimal 10 derajat.
  setSensorCardAlert(tiltCardEl, tiltAlert);
  // Mengubah kotak mata menjadi merah jika mata tertutup.
  setSensorCardAlert(eyeCardEl, eyeAlert);
  // Memperbarui nama driver aktif.
  activeDriverNameEl.textContent = activeName || "-";
  // Memperbarui label jenis kelamin aktif.
  activeDriverGenderEl.textContent = getGenderLabel(activeGender);
  // Memperbarui rule BPM aktif.
  activeRuleEl.textContent = `Microsleep jika BPM < ${threshold}`;
  // Mengisi durasi realtime jika microsleep aktif.
  liveDurationEl.textContent = `Durasi: ${formatDuration(data)}`;
  // Mengatur kelas status jika microsleep aktif.
  const isMicrosleep = status === "Microsleep";
  // Mengubah warna status menjadi mode alert.
  statusCardEl.classList.toggle("alert", isMicrosleep);
  // Mengubah background body menjadi mode alert.
  document.body.classList.toggle("alert-mode", isMicrosleep);
  // Mengatur warna tulisan status.
  statusEl.className = `metric-value ${isMicrosleep ? "danger-text" : "normal-text"}`;
  // Memakai pulse_visual dari monitoring sebagai fallback visualizer.
  renderPulseLine({
    visual: data.pulse_visual,
    signal_present: data.pulse_signal_present,
    beat_event: data.pulse_state,
    pulse_valid: data.pulse_valid,
    pulse_no_signal: data.pulse_no_signal,
  });
  // Membunyikan alarm hanya saat status berubah dari bukan Microsleep ke Microsleep.
  if (isMicrosleep && lastStatus !== "Microsleep") {
    // Menampilkan popup microsleep.
    showPopup("🚨 MICROSLEEP TERDETEKSI!");
    // Memutar alarm dengan catch karena browser bisa memblokir autoplay.
    alarmEl.play().catch(() => {});
  }
  // Menyimpan status terakhir untuk deteksi transisi alarm.
  lastStatus = status;
}

// Fungsi ini menggerakkan bar vertikal merah di kotak BPM.
function renderPulseLine(pulse) {
  // Mengambil nilai visual 0 sampai 100 dari ESP32.
  const visual = Number(pulse.visual ?? 0);
  // Membatasi visual agar tetap berada di rentang 0 sampai 100.
  const clampedVisual = Math.max(0, Math.min(100, visual));
  // Mengubah nilai visual menjadi tinggi bar dalam pixel.
  const barHeight = 8 + clampedVisual * 0.58;
  // Mengecek apakah sinyal pulse dianggap ada.
  const signalPresent = pulse.signal_present === 1 || pulse.signal_present === true;
  // Mengecek apakah pulse valid menurut ESP32.
  const pulseValid = pulse.pulse_valid === 1 || pulse.pulse_valid === true;
  // Mengecek apakah ESP32 mendeteksi no signal.
  const noSignal = pulse.pulse_no_signal === 1 || pulse.pulse_no_signal === true;
  // Mengubah tinggi bar supaya naik-turun seperti volume meter.
  pulseLineEl.style.height = `${barHeight}px`;
  // Membuat bar lebih redup jika sinyal tidak ada.
  pulseLineEl.style.opacity = signalPresent ? "1" : "0.35";
  // Mengisi keterangan pulse di bawah BPM.
  pulseInfoEl.textContent = noSignal ? "Pulse: No Signal" : pulseValid ? "Pulse: Valid" : signalPresent ? "Pulse: Terdeteksi" : "Menunggu sinyal pulse";
  // Menambahkan animasi singkat jika beat_event masuk.
  if (pulse.beat_event === 1 || pulse.beat_event === true) {
    // Menghapus class beat agar animasi bisa diputar ulang.
    pulseLineEl.classList.remove("beat");
    // Memaksa browser membaca ulang layout sebelum class dipasang lagi.
    void pulseLineEl.offsetWidth;
    // Menambahkan class beat untuk animasi.
    pulseLineEl.classList.add("beat");
  }
}

// ------------------ RENDER HISTORI MICROSLEEP ------------------
// Fungsi ini merender tabel histori microsleep.
function renderMicrosleepTable(rows) {
  // Mengecek apakah tidak ada data microsleep.
  if (!rows.length) {
    // Menampilkan baris kosong.
    microTableBodyEl.innerHTML = `<tr><td colspan="10" class="empty-row">Belum ada histori microsleep.</td></tr>`;
    // Menghentikan fungsi.
    return;
  }
  // Membuat HTML baris dari data microsleep.
  microTableBodyEl.innerHTML = rows.map((row) => {
    // Mengambil durasi microsleep dari row.
    const duration = formatDuration(row);
    // Mengambil gender row dengan fallback dari driver aktif.
    const gender = row.driver_gender || activeGender;
    // Mengambil threshold row dengan fallback rule gender.
    const threshold = row.bpm_rule_threshold || getBpmThreshold(gender);
    // Mengembalikan template baris tabel.
    return `
      <tr>
        <td>${safeText(row.datetime)}</td>
        <td>${safeText(row.driver_name)}</td>
        <td>${getGenderLabel(gender)}</td>
        <td>${formatBpm(row.bpm)}</td>
        <td>BPM &lt; ${threshold}</td>
        <td>${safeText(row.kemiringan)}</td>
        <td>${safeText(row.mata)}</td>
        <td>${duration}</td>
        <td>${safeText(row.status)}</td>
        <td><button class="row-delete-button" type="button" onclick="deleteMicrosleepItem('${row.key}')">Hapus</button></td>
      </tr>
    `;
  }).join("");
}

// Fungsi ini menghapus satu data histori microsleep berdasarkan key Firebase.
function deleteMicrosleepItem(key) {
  // Menghentikan fungsi jika key kosong.
  if (!key) return;
  // Menampilkan konfirmasi sebelum menghapus data.
  if (!confirm("Hapus data microsleep ini?")) return;
  // Menghapus data dari path history_microsleep.
  db.ref(`history_microsleep/${key}`).remove();
}

// Fungsi ini menghapus semua histori microsleep yang tampil.
function clearMicrosleepHistory() {
  // Menampilkan konfirmasi agar data tidak terhapus tanpa sengaja.
  if (!confirm("Hapus semua histori microsleep?")) return;
  // Menghapus seluruh path history_microsleep dari Firebase.
  db.ref("history_microsleep").remove();
}

// Fungsi ini membuat file Excel dari histori microsleep saja.
function downloadMicrosleepExcel() {
  // Mengecek apakah data kosong.
  if (!microsleepData.length) {
    // Menampilkan pesan jika belum ada data.
    showPopup("Belum ada histori microsleep untuk di-download.", false);
    // Menghentikan fungsi.
    return;
  }
  // Membuat data yang lebih rapi untuk Excel.
  const rows = microsleepData.map((row) => ({
    // Kolom nama driver.
    Nama: safeText(row.driver_name),
    // Kolom jenis kelamin.
    "Jenis Kelamin": getGenderLabel(row.driver_gender || activeGender),
    // Kolom waktu kejadian.
    Waktu: safeText(row.datetime),
    // Kolom BPM.
    BPM: formatBpm(row.bpm),
    // Kolom rule BPM.
    "Rule BPM": `BPM < ${row.bpm_rule_threshold || getBpmThreshold(row.driver_gender || activeGender)}`,
    // Kolom kemiringan.
    Kemiringan: safeText(row.kemiringan),
    // Kolom mata.
    Mata: safeText(row.mata),
    // Kolom durasi microsleep baru.
    "Durasi Microsleep": formatDuration(row),
    // Kolom status.
    Status: safeText(row.status),
  }));
  // Membuat worksheet dari array JSON.
  const worksheet = XLSX.utils.json_to_sheet(rows);
  // Membuat workbook kosong.
  const workbook = XLSX.utils.book_new();
  // Menambahkan worksheet ke workbook.
  XLSX.utils.book_append_sheet(workbook, worksheet, "Microsleep");
  // Menulis file Excel ke komputer user.
  XLSX.writeFile(workbook, "histori_microsleep.xlsx");
}

// ------------------ MODE PENGUJIAN ------------------
// Fungsi ini mengubah angka detik menjadi format MM:SS.
function formatTimerValue(totalSeconds) {
  // Menghitung menit dari total detik.
  const minutes = Math.floor(totalSeconds / 60);
  // Mengambil sisa detik setelah menit dihitung.
  const seconds = totalSeconds % 60;
  // Mengembalikan format dua digit agar tampil seperti 00:00.
  return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

// Fungsi ini memperbarui tampilan timer Mode Pengujian.
function timerDisplay() {
  // Mengisi teks timer berdasarkan jumlah detik yang sedang berjalan.
  timerEl.textContent = formatTimerValue(testSeconds);
}

// Fungsi ini membuka atau menutup panel Mode Pengujian.
function toggleTest() {
  // Membalik status panel terbuka atau tertutup.
  isTestOpen = !isTestOpen;
  // Menampilkan panel jika terbuka dan menyembunyikan jika tertutup.
  testModeEl.classList.toggle("hidden", !isTestOpen);
  // Mengubah teks tombol sesuai kondisi panel.
  btnTestEl.textContent = isTestOpen ? "⬅️ Kembali" : "🧪 Mode Pengujian";
}

// Fungsi ini mengambil nama untuk file pengujian.
function getTestName() {
  // Mengambil nama dari input Mode Pengujian terlebih dahulu.
  const inputName = testNameEl.value.trim();
  // Jika input Mode Pengujian kosong, gunakan nama driver aktif.
  return inputName || activeName || driverNameEl.value.trim();
}

// Fungsi ini menyimpan satu snapshot data realtime ke array Mode Pengujian.
function pushTestSnapshot(name) {
  // Mengambil data monitoring terakhir dengan fallback object kosong.
  const data = currentData || {};
  // Menambahkan data pengujian sesuai struktur kode awal, ditambah durasi microsleep.
  testData.push({
    // Menyimpan nama pengujian atau nama driver.
    nama: name,
    // Menyimpan waktu rekam dari browser.
    waktu: new Date().toLocaleString("id-ID"),
    // Menyimpan BPM realtime.
    bpm: formatBpm(data.bpm),
    // Menyimpan kemiringan kepala realtime.
    kemiringan: safeText(data.kemiringan),
    // Menyimpan kondisi mata realtime.
    mata: safeText(data.mata),
    // Menyimpan status sistem realtime.
    status: safeText(data.status),
    // Menyimpan durasi microsleep realtime jika ada.
    durasi_microsleep: formatDuration(data),
    // Menyimpan status pulse per sensor jika tersedia.
    status_pulse: safeText(data.status_pulse),
    // Menyimpan status kepala per sensor jika tersedia.
    status_head: safeText(data.status_head),
    // Menyimpan status mata per sensor jika tersedia.
    status_eye: safeText(data.status_eye),
  });
}

// Fungsi ini memulai Mode Pengujian dan merekam data setiap detik.
function startTest() {
  // Mengambil nama pengujian.
  const name = getTestName();
  // Memastikan nama tersedia sebelum rekaman dimulai.
  if (!name) {
    // Menampilkan pesan jika nama belum diisi.
    showPopup("Isi nama terlebih dahulu!");
    // Menghentikan fungsi.
    return;
  }
  // Mengisi input Mode Pengujian agar nama terlihat oleh user.
  testNameEl.value = name;
  // Menghentikan timer lama jika Start ditekan ulang.
  if (testTimer) clearInterval(testTimer);
  // Mengosongkan data pengujian sebelumnya.
  testData = [];
  // Mengatur ulang detik ke nol.
  testSeconds = 0;
  // Menandai Mode Pengujian sedang berjalan.
  isTestRunning = true;
  // Memperbarui tampilan timer ke 00:00.
  timerDisplay();
  // Menyimpan snapshot awal saat tombol Start ditekan.
  pushTestSnapshot(name);
  // Menampilkan pesan mulai.
  showPopup("Mode Pengujian dimulai.", false);
  // Membuat interval untuk merekam data setiap satu detik.
  testTimer = setInterval(() => {
    // Menambah hitungan detik.
    testSeconds += 1;
    // Memperbarui tampilan timer.
    timerDisplay();
    // Menyimpan snapshot data realtime ke array.
    pushTestSnapshot(name);
  }, 1000);
}

// Fungsi ini menghentikan Mode Pengujian dan mengunduh Excel.
function stopTest() {
  // Mengambil nama pengujian.
  const name = getTestName();
  // Menghentikan timer jika sedang berjalan.
  if (testTimer) clearInterval(testTimer);
  // Menghapus referensi timer agar tidak dobel.
  testTimer = null;
  // Menandai Mode Pengujian sudah berhenti.
  isTestRunning = false;
  // Mengecek apakah ada data yang dapat diunduh.
  if (!testData.length) {
    // Menampilkan pesan jika belum ada data.
    showPopup("Belum ada data pengujian untuk di-download.", false);
    // Menghentikan fungsi.
    return;
  }
  // Membuat worksheet Excel dari data pengujian.
  const worksheet = XLSX.utils.json_to_sheet(testData);
  // Membuat workbook kosong.
  const workbook = XLSX.utils.book_new();
  // Menambahkan worksheet ke workbook.
  XLSX.utils.book_append_sheet(workbook, worksheet, "Pengujian");
  // Mengambil tanggal hari ini untuk nama file.
  const today = new Date().toISOString().slice(0, 10);
  // Membersihkan nama file dari karakter yang tidak aman.
  const safeName = (name || "driver").replace(/[^a-z0-9_-]/gi, "_");
  // Mengunduh file Excel seperti fitur awal.
  XLSX.writeFile(workbook, `Pengujian_${safeName}_${today}.xlsx`);
  // Menampilkan pesan berhasil.
  showPopup("Download pengujian berhasil.", false);
}

// ------------------ EVENT LISTENER HALAMAN ------------------
// Menjalankan update jam pertama kali.
updateClock();

// Menjalankan update jam setiap satu detik.
setInterval(updateClock, 1000);

// Menyalakan indikator status koneksi alat di bagian atas halaman.
startConnectionStatusWatcher();

// Mengisi form dari data tersimpan browser.
loadSavedDriverForm();

// Memasang listener perubahan gender untuk memperbarui teks rule.
driverGenderEl.addEventListener("change", updateRuleNote);

// Memasang listener submit form driver.
driverFormEl.addEventListener("submit", handleDriverSubmit);

// Memasang listener tombol ganti driver.
changeDriverButtonEl.addEventListener("click", showSetupArea);

// Memasang listener tombol download Excel.
downloadMicrosleepButtonEl.addEventListener("click", downloadMicrosleepExcel);

// Memasang listener tombol hapus histori microsleep.
clearMicrosleepButtonEl.addEventListener("click", clearMicrosleepHistory);

// Memasang listener tombol buka/tutup Mode Pengujian.
btnTestEl.addEventListener("click", toggleTest);

// Memasang listener tombol Start Mode Pengujian.
startTestButtonEl.addEventListener("click", startTest);

// Memasang listener tombol Stop Mode Pengujian.
stopTestButtonEl.addEventListener("click", stopTest);

// Membuka fungsi hapus per baris ke global scope karena dipanggil dari HTML template.
window.deleteMicrosleepItem = deleteMicrosleepItem;
