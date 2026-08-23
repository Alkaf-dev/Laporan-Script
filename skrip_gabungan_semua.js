/*************************************************************
 * SKRIP GABUNGAN (MUTABA'AH + LAPORAN) — SATU KODE UNTUK KEDUA FILE
 *
 * CARA PAKAI (SANGAT MUDAH):
 *  1. Paste SELURUH kode ini di Apps Script file MUTABA'AH
 *     (Ctrl+A -> Delete -> paste), lalu Save.
 *  2. Paste SELURUH kode ini JUGA di Apps Script file LAPORAN BULAN
 *     (Ctrl+A -> Delete -> paste), lalu Save.
 *  3. Tidak perlu pilih mana yang ke mana — saat spreadsheet dibuka,
 *     skrip otomatis mendeteksi jenis filenya:
 *       - File Mutaba'ah -> menu: Sinkronisasi Laporan / Tanggal Otomatis / Pekan Tools
 *       - File Laporan   -> menu: Mutabaah > Perbaiki Semua
 *
 * PENGAMAN: fitur khas Mutaba'ah (validasi edit, sapu bersih, gerbang bulan)
 * otomatis NONAKTIF bila berada di file Laporan, jadi tidak akan mengganggu
 * data/tampilan laporan sama sekali. Demi aman, jangan menjalankan fungsi
 * manual dari editor selain lewat menu yang tersedia di tiap file.
 *
 * [v9.1-GABUNGAN] Sisi Mutaba'ah kini memuat mesin audit nama lengkap dari
 * skrip_mutabaah.js v9.1: menu 'Perbaiki Nama Santri (Audit Kelas)',
 * auto-perbaikan nama di awal Sinkronisasi, penomoran robust batch (cepat),
 * serta writeToLaporan anti-nomor-dobel + trim Grade di sisi Laporan.
 *************************************************************/

/* Deteksi jenis file:
   File LAPORAN memiliki tab kelas (7A..12) / MasterData / Rekap / Persentase Total,
   sedangkan file MUTABA'AH hanya berisi sheet halaqah ustadz. */
function adalahFileLaporan_(ss) {
  var ciri = ['7A', '7B', '8', '9A', '9B', '10A', '10B', '11A', '11B', '12',
              'MasterData', 'Rekap Absensi SMP', 'Rekap Absensi SMA', 'Persentase Total'];
  for (var i = 0; i < ciri.length; i++) {
    if (ss.getSheetByName(ciri[i])) return true;
  }
  return false;
}

// ============== MENU OTOMATIS SESUAI JENIS FILE ==============
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (adalahFileLaporan_(ss)) {
    // ----- Menu untuk file LAPORAN -----
    SpreadsheetApp.getUi().createMenu('Mutabaah')
      .addItem('Perbaiki Semua', 'perbaikiSemua')
      .addItem('Sembunyikan Error Rumus (#VALUE! dll.) jadi 0/0%', 'jalankanSembunyikanErrorRumus')
      .addItem('Sembunyikan Error Rumus - Sheet Aktif Saja', 'jalankanSembunyikanErrorRumusAktif')
      .addToUi();
    return;
  }
  // ----- Menu untuk file MUTABA'AH -----
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Sinkronisasi Laporan')
    .addItem('Sync Sekarang (Semua Halaqah)', 'syncSemuaHalaqah')
    .addItem('Pasang Trigger Otomatis', 'pasangTriggerOtomatis')
    .addSeparator()
    .addItem('Perbaiki Nama Santri (Audit Kelas)', 'perbaikiNamaMutabaah')
    .addSeparator()
    .addItem('Samakan Nama Halaqah dgn Bulan Sebelumnya', 'samakanNamaHalaqah')
    .addToUi();
  ui.createMenu('📅 Tanggal Otomatis')
    .addItem('Isi Tanggal Otomatis', 'isiTanggalOtomatis')
    .addItem('Isi Nama Bulan Otomatis', 'isiBulanOtomatis')
    .addToUi();
  // Nama bulan diisi otomatis tiap file dibuka (silent, tanpa alert).
  try { isiBulanOtomatisSilent(); } catch (e) {}
  ui.createMenu('Pekan Tools')
    .addItem('Cek Status Bulan Sebelumnya', 'cekStatusBulanSebelumnya')
    .addItem('Bersihkan Data Tidak Valid Sekarang', 'jalankanSapuBersihManual')
    .addItem('Keras Segarkan Gerbang (Cek Bulan Lalu)', 'jalankanRefreshGerbang')
    .addSeparator()
    .addItem('Aktifkan Pembersihan Otomatis Tiap 5 Menit', 'pasangJadwalSapuBersihOtomatis')
    .addSeparator()
    .addItem('Perbaiki Referensi Rumus Error (#REF!) Sekarang', 'jalankanPerbaikiRefErrorManual')
    .addSeparator()
    .addItem('Sembunyikan Error Rumus (#VALUE! dll.) jadi 0/0%', 'jalankanSembunyikanErrorRumus')
    .addSeparator()
    .addItem('🧹 Hapus Semua Rentang Dilindungi', 'jalankanHapusSemuaProteksi')
    .addSeparator()
    .addItem('🔍 Diagnostik: Cek Segmen Sheet Aktif', 'tesSegmen')
    .addItem('📋 Cek Kolom Belum Terisi (Sheet Aktif)', 'jalankanDiagnosisKosong')
    .addToUi();
}
/**
 * =========================================================
* SCRIPT GABUNGAN MUTABA'AH (VERSI DIPERBAIKI v5 - BLOK AKURAT + FIX LABEL BULAN)
 * =========================================================
 * 1. SINKRONISASI -> kirim ringkasan ke Master Laporan (debounce 30 dtk)
 * 2. TANGGAL OTOMATIS -> isi header hari/tanggal tiap pekan
 * 3. VALIDASI LINTAS FILE -> BLOKIR + hapus isian yang keliru disertai alert,
 *    DENGAN CEK YANG AKURAT (tidak menghapus bila sudah lengkap).
 * 4. PERBAIKI #REF! (MANUAL)
 *
 * PERUBAHAN PENTING v5 (FIX "SHEET ALKAF DIANGGAP BELUM LENGKAP"):
 * - isNamaTidakValid() sekarang membandingkan nama TANPA spasi. Baris footer
 *   "I Z I N", "S A K I T", "A L P A", "T D K  SETOR" (memakai spasi antar
 *   huruf) dulu tidak cocok dengan kata kunci "IZIN", sehingga dihitung sebagai
 *   baris siswa yang tidak punya data -> semua sheet (mis. Alkaf) dinilai
 *   "belum lengkap" padahal datanya sudah penuh. Kini label itu dikenali dan
 *   dikecualikan, sehingga bulan yang lengkap kembali dianggap LENGKAP.
 *
 * PERUBAHAN PENTING v4 (kebijakan yang Anda minta):
 * - onEdit() MENGHAPUS HANYA SEL yang diisi di posisi yang belum boleh
 *   (mis. mengisi Pekan sebelum prasyarat yang belum selesai), lalu menampilkan
 *   alert (toast). TIDAK PERNAH menghapus semua data sekaligus.
 * - refreshGerbangCache() menetapkan allowed:false HANYA kalau bulan lalu
 *   SUNGGUH belum lengkap (dicek akurat). Kalau sudah lengkap -> allowed:true.
 *   Kalau file bulan lalu tak ketemu / gagal dibuka / sheet tak ada -> JANGAN
 *   mengunci (dijaga agar tidak menghapus data yang sah).
 * - [KUNCI AKURASI] cek bulan lalu memakai pekan TERAKHIR yang benar-benar
 *   berisi data (bukan Pekan 5 template), jadi bulan yang memang sudah lengkap
 *   TIDAK kena hapus lagi. Inilah yang dulu bikin "sudah lengkap malah terhapus"
 *   di Agustus.
 * - Gerbang lintas-bulan di onEdit hanya berlaku bila cache masih SEGAR (sesuai
 *   jeda trigger berkala), supaya tidak menghapus isian yang sebenarnya sah
 *   karena cache basi.
 * - sapuBersihOtomatis() (tiap 5 menit) tetap TIDAK menghapus massal; ia hanya
 *   menghitung + memberi peringatan. Penghapusan besar hanya lewat menu MANUAL
 *   'Bersihkan Data Tidak Valid Sekarang' bila benar-benar dikehendaki.
 *
 * PENTING - CARA PAKAI:
 * 1. Buka Apps Script (Extensions > Apps Script), hapus semua kode lama,
 *    tempel seluruh isi file ini, lalu Simpan.
 * 2. Jalankan ONCE menu: '🔄 Sinkronisasi Laporan > Pasang Trigger Otomatis'.
 * 3. Hapus trigger lama bernama onEditValidasi / onEditInstallable pada
 *    menu Pengelolaan > Triggers bila masih muncul.
 * 4. Fitur 'Perbaiki #REF!' hanya dijalankan manual lewat menu.
 * 5. Menutup bulan sebelumnya: isi dulu bulan itu sampai lengkap, lalu
 *    jalankan menu 'Keras Segarkan Gerbang (Cek Bulan Lalu)'.
 */

// ============== KONFIGURASI ==============
const EXCLUDED_SHEETS = ["Rekap", "Panduan"];
const DATA_START_ROW = 5;
const NAME_COL = "B";
const EXCLUDED_NAME_VALUES = ["IZIN", "SAKIT", "ALPA", "TDK SETOR"];
const PREREQ_COLS = ["D", "E", "F", "G", "H"];
const SYNC_DEBOUNCE_MS = 30000;

// Batas isian: hanya area DI ATAS baris berwarna hitam yang wajib diisi.
// JML_KOLOM_SCAN_BATAS = lebar pita kolom yang dipindai untuk mendeteksi hitam.
const WARNA_HITAM = ["#000000", "#000"];
const JML_KOLOM_SCAN_BATAS = 20;
const CACHE_BARIS_BATAS_DTK = 300;

// [OPTIMASI] Isi dengan nama sheet bila ingin HANYA sheet itu yang divalidasi
// (mis. "Adlan"). Kosongkan ("") untuk semua sheet. Mempercepat validasi.
const TRACK_ONLY_SHEET = "";

const TEMPLATE_WEEKS = [
  { name: "Pekan 1", cols: ["J","K","L","M","N","O","P","Q","R","S","T","U","W","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BE","BF"] },
  { name: "Pekan 2", cols: ["BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BX","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ","DA","DB","DC","DD","DF","DG"] },
  { name: "Pekan 3", cols: ["DL","DM","DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DY","EA","EB","EC","ED","EE","EF","EG","EH","EI","EJ","EK","EL","EM","EN","EO","EP","EQ","ER","ES","ET","EU","EV","EW","EX","EY","EZ","FA","FB","FC","FD","FE","FG","FH"] },
  { name: "Pekan 4", cols: ["FM","FN","FO","FP","FQ","FR","FS","FT","FU","FV","FW","FX","FZ","GB","GC","GD","GE","GF","GG","GH","GI","GJ","GK","GL","GM","GN","GO","GP","GQ","GR","GS","GT","GU","GV","GW","GX","GY","GZ","HA","HB","HC","HD","HE","HF","HH","HI"] },
  { name: "Pekan 5", cols: ["HN","HO","HP","HQ","HR","HS","HT","HU","HV","HW","HX","HY","IA","IC","ID","IE","IF","IG","IH","II","IJ","IK","IL","IM","IN","IO","IP","IQ","IR","IS","IT","IU","IV","IW","IX","IY","IZ","JA","JB","JC","JD","JE","JF","JG","JI","JJ"] }
];

const DAFTAR_URUTAN_BULAN = [
  "juli", "agustus", "september", "oktober", "november",
  "januari", "februari", "maret", "april", "mei", "juni"
];

const MSG_PREREQ = "LENGKAPI TERLEBIH DAHULU TARGET HAFALAN BULANAN!!!";
const MSG_PREV_WEEK = "LENGKAPI TERLEBIH DAHULU KOLOM MUTABA'AH PEKAN SEBELUMNYA!!!";

const TANGGAL_ROW = 4;
const NAMA_HARI = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
const TANDA_HARI_KOSONG = "-";
const BULAN_KE_ANGKA = {
  "januari":1,"februari":2,"maret":3,"april":4,"mei":5,"juni":6,
  "juli":7,"agustus":8,"september":9,"oktober":10,"november":11,"desember":12
};

const SYNC_CONFIG = {
  HEADER_ROW: 3, SUBHEADER_ROW: 4, DATA_START_ROW: 5, ROWS_PER_STUDENT: 3,
  COL_NO: 1, COL_NAMA: 2, COL_KELAS: 3, COL_TARGET_DARI: 4, COL_TARGET_SAMPAI: 5,
  COL_TARGET_TOTAL: 6, COL_GRADE: 7,
  LAPORAN_DATA_START_ROW: 4,
  LAPORAN_COLS: {
    NO: 1, NAMA: 2, GRADE: 3, TARGET_BULANAN: 4, TARGET_TERCAPAI: 5,
    PRESENTASE: 6, KETERANGAN: 7, IZIN: 8, SAKIT: 9, ALPA: 10,
    TIDAK_SETOR: 11, JUMLAH_HAFALAN: 12, CATATAN: 13
  }
};

// ============== HELPER ==============
function columnLetterToNumber(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) column = column * 26 + (letter.charCodeAt(i) - 64);
  return column;
}

function kapitalisasi(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

// Apakah sheet diperbolehkan untuk ditrack/divalidasi (lihat TRACK_ONLY_SHEET).
function sheetDitrack(namaSheet) {
  return TRACK_ONLY_SHEET === "" || namaSheet === TRACK_ONLY_SHEET;
}

// Cari baris batas: baris pertama pada pita kolom [kelom Awal, ..] yang latarnya
// berwarna hitam. Return nomor baris, atau -1 bila tidak ditemukan.
function cariBarisBatas(sheet) {
  const cache = CacheService.getScriptCache();
  const key = 'batas_' + sheet.getParent().getId() + '_' + sheet.getName();
  const c = cache.get(key);
  if (c !== null) { try { return Number(c); } catch (e) {} }

  const fromCol = columnLetterToNumber(NAME_COL);
  const maxRows = sheet.getMaxRows();
  const jmlBaris = Math.min(maxRows - DATA_START_ROW + 1, 200);
  if (jmlBaris < 1) return -1;
  const bgs = sheet.getRange(DATA_START_ROW, fromCol, jmlBaris, JML_KOLOM_SCAN_BATAS).getBackgroundColors();
  let hasil = -1;
  outer:
  for (let r = 0; r < jmlBaris; r++) {
    for (let c = 0; c < JML_KOLOM_SCAN_BATAS; c++) {
      const warna = String(bgs[r][c]).toUpperCase();
      if (WARNA_HITAM.indexOf(warna) === -1) continue outer;
    }
    hasil = DATA_START_ROW + r;
    break;
  }
  try { cache.put(key, String(hasil), CACHE_BARIS_BATAS_DTK); } catch (e) {}
  return hasil;
}

// Baris data terakhir yang wajib diisi = baris tepat sebelum baris hitam,
// atau getLastRow bila baris hitam tak terdeteksi (fallback aman).
function getBarisDataTerakhir(sheet) {
  const batas = cariBarisBatas(sheet);
  return batas > 0 ? batas - 1 : sheet.getLastRow();
}

function getSegmentsForSheet(sheet) {
  const maxCol = sheet.getMaxColumns();
  const segments = [{ type: "prereq", label: "Target Hafalan Bulanan", cols: PREREQ_COLS }];
  for (let i = 0; i < TEMPLATE_WEEKS.length; i++) {
    const week = TEMPLATE_WEEKS[i];
    if (Math.max.apply(null, week.cols.map(columnLetterToNumber)) > maxCol) break;
    segments.push({ type: "week", label: week.name, cols: week.cols });
  }
  return segments;
}

function isCellFilled(sheet, row, colNum) {
  const cell = sheet.getRange(row, colNum);
  if (cell.isPartOfMerge()) {
    const merged = cell.getMergedRanges();
    if (merged.length > 0) {
      const v = merged[0].getCell(1, 1).getValue();
      return v !== "" && v !== null;
    }
  }
  const v = cell.getValue();
  return v !== "" && v !== null;
}

// [v5] Membandingkan nama TANPA spasi: baris label seperti "I Z I N", "S A K I T",
// "A L P A", "T D K  SETOR" (memakai spasi antar huruf) tetap dikenali sebagai
// label yang harus dikecualikan. Sebelumnya "I Z I N" tidak cocok dengan kata
// kunci "IZIN", sehingga baris footer itu dihitung sebagai siswa tanpa data dan
// membuat sheet (mis. Alkaf) dinilai "belum lengkap" padahal sudah lengkap.
function isNamaTidakValid(namaRaw) {
  if (namaRaw === "" || namaRaw === null) return true;
  const normalized = String(namaRaw).replace(/[\s\u00A0]+/g, "").toUpperCase();
  if (normalized === "") return true;
  return EXCLUDED_NAME_VALUES.some(function (k) {
    return normalized.indexOf(String(k).replace(/[\s\u00A0]+/g, "").toUpperCase()) !== -1;
  });
}

function deteksiBulanDariJudul(judul) {
  const lower = judul.toLowerCase();
  for (let i = 0; i < DAFTAR_URUTAN_BULAN.length; i++) {
    if (new RegExp("\\b" + DAFTAR_URUTAN_BULAN[i] + "\\b", "i").test(lower)) {
      return { index: i, kata: DAFTAR_URUTAN_BULAN[i] };
    }
  }
  return null;
}

function namaBulanSebelumnya(i) { return i <= 0 ? null : DAFTAR_URUTAN_BULAN[i - 1]; }

function cariFileBulanSebelumnya(infoBulan) {
  const fileIni = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const namaBulanPrev = namaBulanSebelumnya(infoBulan.index);
  if (!namaBulanPrev) return { file: null, judulCari: null };
  const judulCari = fileIni.getName().replace(
    new RegExp("\\b" + infoBulan.kata + "\\b", "i"), kapitalisasi(namaBulanPrev)
  );
  let fileTarget = null;
  const folders = fileIni.getParents();
  if (folders.hasNext()) {
    const c = folders.next().getFilesByName(judulCari);
    if (c.hasNext()) fileTarget = c.next();
  }
  if (!fileTarget) {
    const c2 = DriveApp.getFilesByName(judulCari);
    if (c2.hasNext()) fileTarget = c2.next();
  }
  return { file: fileTarget, judulCari: judulCari };
}

function bacaKolomDenganMerge(sheet, startRow, jml, colNum) {
  const range = sheet.getRange(startRow, colNum, jml, 1);
  const values = range.getValues();
  const merges = range.getMergedRanges();
  if (merges.length === 0) return values.map(function (r) { return r[0]; });
  const hasil = [];
  for (let i = 0; i < jml; i++) {
    const row = startRow + i;
    let v = values[i][0];
    if (v === "" || v === null) {
      for (let m = 0; m < merges.length; m++) {
        const mr = merges[m];
        if (row >= mr.getRow() && row < mr.getRow() + mr.getNumRows()) { v = mr.getCell(1, 1).getValue(); break; }
      }
    }
    hasil.push(v);
  }
  return hasil;
}

function bacaSegmentDenganMerge(sheet, startRow, jml, colNumbers) {
  const colMin = Math.min.apply(null, colNumbers);
  const colMax = Math.max.apply(null, colNumbers);
  const range = sheet.getRange(startRow, colMin, jml, colMax - colMin + 1);
  const values = range.getValues();
  const merges = range.getMergedRanges();
  return function (rowOffset) {
    const row = startRow + rowOffset;
    return colNumbers.every(function (colNum) {
      let v = values[rowOffset][colNum - colMin];
      if ((v === "" || v === null) && merges.length > 0) {
        for (let m = 0; m < merges.length; m++) {
          const mr = merges[m];
          if (row >= mr.getRow() && row < mr.getRow() + mr.getNumRows() &&
              colNum >= mr.getColumn() && colNum < mr.getColumn() + mr.getNumColumns()) {
            v = mr.getCell(1, 1).getValue(); break;
          }
        }
      }
      return v !== "" && v !== null;
    });
  };
}

// [v3] Sebuah kumpulan kolom dianggap "terisi" bila SETIAP siswa punya
// MINIMAL SATU sel terisi (merge-aware). Tidak lagi mengharuskan SEMUA kolom
// terisi, sehingga satu-dua sel kosong (siswa berhalangan, hari libur) tidak
// mengunci segmen berikutnya.
// [KETAT] Sebuah segmen dianggap "lengkap" hanya bila SEMUA kolom di dalamnya
// terisi pada SETIAP baris siswa (merge-aware). Satu kolom kosong saja pada
// segmen itu -> segmen dinilai belum lengkap -> segmen berikutnya terkunci.
function apakahKolumTerisiMinimalSatu(sheet, cols) {
  const start = DATA_START_ROW;
  const lastRow = getBarisDataTerakhir(sheet);
  if (lastRow < start) return true;
  const n = lastRow - start + 1;
  const nameCol = columnLetterToNumber(NAME_COL);
  const cMin = Math.min.apply(null, cols);
  const cMax = Math.max.apply(null, cols);
  const nameVals = sheet.getRange(start, nameCol, n, 1).getValues().map(function (r) { return r[0]; });
  const range = sheet.getRange(start, cMin, n, cMax - cMin + 1);
  const block = range.getValues();
  const merges = range.getMergedRanges();
  for (let i = 0; i < n; i++) {
    if (isNamaTidakValid(nameVals[i])) continue;
    const row = start + i;
    for (let k = 0; k < cols.length; k++) {
      const c = cols[k];
      let v = block[i][c - cMin];
      if ((v === "" || v === null) && merges.length > 0) {
        for (let m = 0; m < merges.length; m++) {
          const mr = merges[m];
          if (row >= mr.getRow() && row < mr.getRow() + mr.getNumRows() &&
              c >= mr.getColumn() && c < mr.getColumn() + mr.getNumColumns()) {
            v = mr.getCell(1, 1).getValue(); break;
          }
        }
      }
      if (v === "" || v === null) return false;
    }
  }
  return true;
}

function apakahSegmentLengkap(sheet, segment) {
  return apakahKolumTerisiMinimalSatu(sheet, segment.cols.map(columnLetterToNumber));
}

// [v3] Bulan dianggap "lengkap" bila PEKAN TERAKHIR YANG SUNGGUH BERISI DATA
// sudah terisi minimal satu sel per siswa. Kalau bulan pendek dan Pekan 5 tidak
// dipakai, yang diperiksa adalah pekan terakhir yang benar-benar diisi, bukan
// Pekan 5 di template. Kalau tidak ada data sama sekali -> dianggap belum lengkap.
function apakahSheetLengkap(sheet) {
  const segsWeek = getSegmentsForSheet(sheet).filter(function (s) { return s.type === "week"; });
  if (segsWeek.length === 0) return true;
  const lastRow = getBarisDataTerakhir(sheet);
  const start = DATA_START_ROW;
  if (lastRow < start) return true;
  const n = lastRow - start + 1;

  let lastUsed = -1;
  for (let idx = segsWeek.length - 1; idx >= 0; idx--) {
    const cols = segsWeek[idx].cols.map(columnLetterToNumber);
    const cMin = Math.min.apply(null, cols);
    const cMax = Math.max.apply(null, cols);
    const bl = sheet.getRange(start, cMin, n, cMax - cMin + 1).getValues();
    let ada = false;
    outer:
    for (let r = 0; r < n; r++) {
      for (let cc = 0; cc < bl[r].length; cc++) {
        const v = bl[r][cc];
        if (v !== "" && v !== null) { ada = true; break outer; }
      }
    }
    if (ada) { lastUsed = idx; break; }
  }

  if (lastUsed === -1) return false;
  return apakahKolumTerisiMinimalSatu(sheet, segsWeek[lastUsed].cols.map(columnLetterToNumber));
}

// ============== GERBANG BULAN SEBELUMNYA (PER SHEET, CACHE 60 dtk) ==============
function cekGerbangUntukSheet(namaSheet) {
  const cache = CacheService.getScriptCache();
  const key = 'gerbang_' + SpreadsheetApp.getActiveSpreadsheet().getId() + '_' + namaSheet;
  const cached = cache.get(key);
  if (cached !== null) return JSON.parse(cached);
  const hasil = hitungGerbangUntukSheet(namaSheet);
  try { cache.put(key, JSON.stringify(hasil), 60); } catch (e) {}
  return hasil;
}

// [v3] INFORMATIF: hasilnya TIDAK pernah dipakai untuk menghapus/mengunci data.
// allowed selalu true; status & warning disediakan untuk menu diagnostik.
function hitungGerbangUntukSheet(namaSheet) {
  const ssIni = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ssIni.getName());
  if (!infoBulan || infoBulan.index === 0) {
    return { allowed: true, status: "bulan pertama (tidak ada cek bulan lalu)" };
  }
  const hasil = cariFileBulanSebelumnya(infoBulan);
  if (!hasil.file) {
    return { allowed: true, status: "file bulan lalu tak ditemukan", warning: 'Tidak ditemukan file dengan judul "' + hasil.judulCari + '". Validasi lintas-bulan tidak aktif.' };
  }
  let ssPrev;
  try { ssPrev = SpreadsheetApp.openById(hasil.file.getId()); }
  catch (err) { return { allowed: true, status: "gagal buka file", warning: "Gagal buka file bulan lalu: " + err.message }; }
  const sheetPrev = ssPrev.getSheetByName(namaSheet);
  if (!sheetPrev) {
    return { allowed: true, status: "sheet tak ada di bulan lalu", warning: 'Sheet "' + namaSheet + '" tidak ada di file bulan lalu. Cek nama sheet sama persis.' };
  }
  let lengkap;
  try { lengkap = apakahSheetLengkap(sheetPrev); }
  catch (err) { return { allowed: true, status: "cek gagal", warning: "Gagal cek sheet \"" + namaSheet + "\" bulan lalu: " + err.message }; }
  return lengkap
    ? { allowed: true, status: "lengkap" }
    : { allowed: true, status: "belum-lengkap", warning: 'Bulan lalu belum lengkap. Silakan lengkapi sheet "' + namaSheet + '" bulan ' + kapitalisasi(namaBulanSebelumnya(infoBulan.index)).toUpperCase() + ' (data TIDAK dihapus otomatis).' };
}

// ============== VALIDASI + SINKRONISASI (SATU HANDLER) ==============
// Simple trigger (dipanggil otomatis tiap edit). Simple trigger DILARANG akses
// Drive, jadi gerbang lintas-bulan di sini hanya berjalan bila ada cache yang
// SEGAR (diisi trigger berkala / menu 'Keras Segarkan Gerbang'). Aturan urutan
// pekan dalam sheet tetap selalu dicek.
function onEdit(e) { prosesValidasi(e, false); }

// Installable edit trigger (dibuat oleh pasangTriggerOtomatis). Boleh akses
// Drive, jadi status bulan lalu dihitung LANGSUNG tiap edit (akurat/real-time),
// tidak lagi tergantung cache basi. Ini yang membuat pengisian Agustus-Syahrul
// benar-benar diblokir selama Juli-Syahrul belum lengkap.
function onEditOtomatis(e) { prosesValidasi(e, true); }

function prosesValidasi(e, withDrive) {
  if (adalahFileLaporan_(SpreadsheetApp.getActiveSpreadsheet())) return; // pengaman: nonaktif di file Laporan
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) return;
  // [OPTIMASI] Skip langsung bila sheet ini tidak di-track.
  if (!sheetDitrack(sheet.getName())) return;

  const startRow = range.getRow();
  const endRow = startRow + range.getNumRows() - 1;
  const startCol = range.getColumn();
  const endCol = startCol + range.getNumColumns() - 1;
  const batasBawah = getBarisDataTerakhir(sheet) + 5;
  if (endRow < DATA_START_ROW || startRow > batasBawah) return;

  const segments = getSegmentsForSheet(sheet);
  if (segments.length === 0) return;

  // ---- Gerbang lintas-bulan ----
  let gerbang = null;
  if (withDrive) {
    try { gerbang = refreshGerbangLive(sheet.getName()); } catch (er) { gerbang = null; }
  } else if (gerbangMasihSegar()) {
    gerbang = bacaGerbangCache(sheet.getName());
  }
  if (gerbang !== null && !gerbang.allowed) {
    range.clearContent();
    sheet.getRange(startRow, startCol).setNote(gerbang.message || gerbang.warning || "Bulan lalu belum lengkap.");
    toggPeringatan(sheet, gerbang.message || gerbang.warning || "Bulan lalu belum lengkap.");
    return;
  }

  // ---- ATURAN POSISI (dalam sheet) ----
  // limit = indeks segmen pertama yang BELUM terisi (boleh diisi segmen itu
  // sendiri, tapi TIDAK boleh mengisi segmen setelahnya).
  let limit = -1;
  for (let i = 0; i < segments.length; i++) {
    if (!apakahSegmentLengkap(sheet, segments[i])) { limit = i; break; }
  }

  // Segmen mana tempat sel pertama yang diedit berada.
  let idxEdit = -1;
  for (let i = 0; i < segments.length; i++) {
    const colNums = segments[i].cols.map(columnLetterToNumber);
    if (colNums.some(function (c) { return c >= startCol && c <= endCol; })) { idxEdit = i; break; }
  }
  // Kolom di luar segmen (mis. No, Kelas, kolom bantu/jumlah lain) TIDAK
  // divalidasi urutannya: selalu diizinkan agar isian sah tidak pernah terhapus.
  // Hanya isian pada kolom segmen (target/prasyarat & pekan) yang diatur.
  if (idxEdit === -1) return;

  // Hapus hanya sel yang baru diisi ini + alert. Bukan menghapus semua data.
  if (limit !== -1 && idxEdit > limit) {
    const penyebab = segments[limit];
    const pesan = penyebab.type === "prereq" ? MSG_PREREQ : MSG_PREV_WEEK;
    range.clearContent();
    sheet.getRange(startRow, startCol).setNote(pesan);
    toggPeringatan(sheet, pesan);
    return;
  }
}

// [v4] refreshGerbangCache menetapkan allowed:false HANYA bila bulan lalu
// sungguh belum lengkap (dicek akurat lewat pekan-terakhir-berisi-data). Kalau
// sudah lengkap -> allowed:true (tidak dihapus). Kalau file bulan lalu tak
// ketemu / gagal dibuka / sheet tak ada -> tetap allowed:true (tidak mengunci),
// supaya tidak menghapus data yang sah karena kendala teknis, bukan karena data.
function refreshGerbangCache(hanyaSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const key = 'gerbang_' + ss.getId();
  const namaSheets = ss.getSheets().map(function (sh) { return sh.getName(); })
    .filter(function (n) { return EXCLUDED_SHEETS.indexOf(n) === -1; });
  // [OPTIMASI] Kalau hanyaSheet diberikan (dari onEdit), hanya sheet itu yang
  // dihitung - tidak membuka & memeriksa seluruh file bulan lalu tiap edit.
  const daftar = hanyaSheet
    ? namaSheets.filter(function (n) { return n === hanyaSheet; })
    : namaSheets;
  if (daftar.length === 0) return bacaMapGerbang() || {};
  const map = bacaMapGerbang() || {};

  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan || infoBulan.index === 0) {
    daftar.forEach(function (n) { map[n] = { allowed: true }; });
    map.__ts__ = Date.now();
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(map));
    return map;
  }

  const namaPrev = kapitalisasi(namaBulanSebelumnya(infoBulan.index)).toUpperCase();
  const hasil = cariFileBulanSebelumnya(infoBulan);
  let ssPrev = null;
  let peringatan = null;
  if (!hasil.file) {
    peringatan = "Belum ditemukan file bulan sebelumnya (cari: \"" + hasil.judulCari + "\"). Validasi lintas-bulan nonaktif.";
  } else {
    try { ssPrev = SpreadsheetApp.openById(hasil.file.getId()); }
    catch (err) { peringatan = "Gagal buka file bulan lalu: " + err.message + ". Validasi lintas-bulan nonaktif."; }
  }

  if (peringatan) {
    daftar.forEach(function (n) { map[n] = { allowed: true, warning: peringatan }; });
  } else {
    daftar.forEach(function (n) {
      const sp = ssPrev.getSheetByName(n);
      if (!sp) {
        map[n] = { allowed: true, warning: "Sheet \"" + n + "\" tidak ada di file bulan lalu. Cek nama sheet sama persis." };
        return;
      }
      let lengkap;
      try { lengkap = apakahSheetLengkap(sp); }
      catch (err) {
        map[n] = { allowed: true, warning: "Gagal cek sheet \"" + n + "\": " + err.message };
        return;
      }
      // KEBIJAKAN: bulan lalu yang BENAR-BENAR belum lengkap -> allowed:false
      // (isian di bulan ini yang menyasar posisi itu akan diblokir onEdit).
      // Kalau sudah lengkap -> allowed:true (tidak dihapus). Berkat cek
      // pekan-terakhir-berisi-data, bulan yang lengkap TIDAK lagi kena blokir.
      map[n] = lengkap
        ? { allowed: true }
        : { allowed: false, message: "LENGKAPI TERLEBIH DAHULU DATA SHEET \"" + n + "\" BULAN " + namaPrev + "!!!" };
    });
  }

  map.__ts__ = Date.now();
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(map));
  return map;
}

function bacaGerbangCache(namaSheet) {
  const raw = PropertiesService.getScriptProperties()
    .getProperty('gerbang_' + SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!raw) return null;
  try {
    const map = JSON.parse(raw);
    return map.hasOwnProperty(namaSheet) ? map[namaSheet] : null;
  } catch (e) {
    return null;
  }
}

function bacaMapGerbang() {
  const raw = PropertiesService.getScriptProperties()
    .getProperty('gerbang_' + SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

// Cache dianggap segar bila < 6 menit. Trigger berkala merefresh tiap 5 menit,
// jadi pada kondisi normal cache segar. Kalau belum ada refresh sama sekali
// (cache basi), onEdit TIDAK memblokir apa pun supaya tidak menghapus isian
// yang sebenarnya sah karena cache yang tidak akurat.
function gerbangMasihSegar() {
  const m = bacaMapGerbang();
  if (!m || !m.__ts__) return false;
  return (Date.now() - m.__ts__) < 360000;
}

// Dipakai oleh trigger INSTALLABLE (punya akses Drive): menghitung status bulan
// lalu SECARA LANGSUNG untuk satu sheet, lalu menyimpannya 3 menit di cache
// ringan supaya tidak membuka file bulan lalu untuk tiap sel yang diketik.
function refreshGerbangLive(namaSheet) {
  const cache = CacheService.getScriptCache();
  const ck = 'gerbangLive_' + SpreadsheetApp.getActiveSpreadsheet().getId() + '_' + namaSheet;
  const c = cache.get(ck);
  if (c !== null) {
    try { return JSON.parse(c); } catch (e) {}
  }
  const map = refreshGerbangCache(namaSheet);
  const hasil = (map && map.hasOwnProperty(namaSheet)) ? map[namaSheet] : { allowed: true };
  try { cache.put(ck, JSON.stringify(hasil), 180); } catch (e) {}
  return hasil;
}

function toggPeringatan(sheet, pesan) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(pesan, "Pekan Tools", 8);
  } catch (e) {}
}

function isHalaqahSheet(sheet) {
  return String(sheet.getRange(SYNC_CONFIG.HEADER_ROW, SYNC_CONFIG.COL_NO).getValue()).trim() === 'No';
}

function bolehSyncLagi(sheetName) {
  const cache = CacheService.getScriptCache();
  const key = 'sync_' + SpreadsheetApp.getActiveSpreadsheet().getId() + '_' + sheetName;
  const last = cache.get(key);
  const now = Date.now();
  if (last !== null && (now - Number(last)) < SYNC_DEBOUNCE_MS) return false;
  cache.put(key, String(now), Math.ceil(SYNC_DEBOUNCE_MS / 1000));
  return true;
}

function pasangTriggerOtomatis() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditValidasi' || fn === 'onEditInstallable' || fn === 'onEditAuto' || fn === 'onEditOtomatis') ScriptApp.deleteTrigger(t);
  });
  // Installable EDIT trigger: boleh akses Drive, sehingga gerbang lintas-bulan
  // (cek bulan lalu) berjalan LANGSUNG dan akurat setiap edit.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onEditOtomatis").forSpreadsheet(ss).onEdit().create();
  installSapuBersihOtomatis();
  // bersihkan cache lama yang mungkin masih berisi allowed:false dari versi
  // sebelumnya, supaya onEdit tidak pernah membaca status kunci yang basi.
  try { refreshGerbangCache(); } catch (e) {}
  SpreadsheetApp.getUi().alert('Selesai. Validasi otomatis terpasang: (1) trigger edit berizin Drive untuk cek bulan lalu secara langsung, (2) pembersihan 5 menit (hanya peringatan).');
}

function installSapuBersihOtomatis() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "sapuBersihOtomatis") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("sapuBersihOtomatis").timeBased().everyMinutes(5).create();
}

function cekStatusBulanSebelumnya() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan) { SpreadsheetApp.getUi().alert("Nama file tak mengandung bulan yang dikenali. Validasi lintas-bulan tidak aktif."); return; }
  if (infoBulan.index === 0) { SpreadsheetApp.getUi().alert("Bulan: " + kapitalisasi(DAFTAR_URUTAN_BULAN[0]) + " (bulan pertama tahun ajaran) - tidak ada cek ke bulan lalu."); return; }
  const hasil = cariFileBulanSebelumnya(infoBulan);
  if (!hasil.file) { SpreadsheetApp.getUi().alert("File bulan lalu \"" + hasil.judulCari + "\" tidak ketemu di folder sama. Sheet TIDAK dikunci otomatis - cek nama file/folder."); return; }
  let ssPrev;
  try { ssPrev = SpreadsheetApp.openById(hasil.file.getId()); }
  catch (err) { SpreadsheetApp.getUi().alert("File bulan lalu ketemu tapi gagal dibuka: " + err.message); return; }
  const baris = [];
  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (EXCLUDED_SHEETS.indexOf(n) !== -1) return;
    const sp = ssPrev.getSheetByName(n);
    if (!sp) { baris.push("- " + n + ": tak ada di file bulan lalu -"); return; }
    try { baris.push("- " + n + ": " + (apakahSheetLengkap(sp) ? "LENGKAP" : "BELUM LENGKAP")); }
    catch (e) { baris.push("- " + n + ": gagal cek (" + e.message + ")"); }
  });
  SpreadsheetApp.getUi().alert("Bulan: " + kapitalisasi(DAFTAR_URUTAN_BULAN[infoBulan.index]) + ".\nFile bulan lalu: \"" + hasil.file.getName() + "\".\n\nStatus tiap sheet:\n\n" + baris.join("\n"));
}

// ============== SAPU BERSIH (HANYA MANUAL YANG MENGHAPUS) ==============
function segmenBerisi(sheet, row, segment) {
  return segment.cols.map(columnLetterToNumber).some(function (c) { return isCellFilled(sheet, row, c); });
}

function sapuBersihDalamFile(dryRun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let total = 0;
  ss.getSheets().forEach(function (sheet) {
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) return;
    if (!sheetDitrack(sheet.getName())) return;
    const lastRow = getBarisDataTerakhir(sheet);
    if (lastRow < DATA_START_ROW) return;
    const segs = getSegmentsForSheet(sheet);
    for (let i = 1; i < segs.length; i++) {
      const prev = segs[i - 1];
      if (apakahSegmentLengkap(sheet, prev)) continue;
      const seg = segs[i];
      if (dryRun) { total++; continue; }
      const cols = seg.cols.map(columnLetterToNumber);
      const cMin = Math.min.apply(null, cols);
      const cMax = Math.max.apply(null, cols);
      const rows = lastRow - DATA_START_ROW + 1;
      sheet.getRange(DATA_START_ROW, cMin, rows, cMax - cMin + 1).clearContent();
      sheet.getRange(DATA_START_ROW, cols[0]).setNote(
        "Data \"" + seg.label + "\" dihapus manual karena \"" + prev.label + "\" belum terisi."
      );
      total++;
    }
  });
  return total;
}

function sapuBersihLintasFile(dryRun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const key = 'gerbang_' + ss.getId();
  let map = null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(key);
    if (raw) map = JSON.parse(raw);
  } catch (e) { map = null; }
  if (!map) { try { map = refreshGerbangCache(); } catch (e) { return 0; } }

  let total = 0;
  ss.getSheets().forEach(function (sheet) {
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) return;
    if (!sheetDitrack(sheet.getName())) return;
    // [v3] refreshGerbangCache selalu menetapkan allowed:true, jadi bagian ini
    // praktis tak akan pernah aktif. Dipertahankan demi pembersihan manual yang
    // benar-benar disengaja bila suatu saat kebijakan diubah.
    const g = map[sheet.getName()];
    if (!g || g.allowed) return;
    const lastRow = getBarisDataTerakhir(sheet);
    if (lastRow < DATA_START_ROW) return;
    const segs = getSegmentsForSheet(sheet);
    if (dryRun) { total += segs.length; return; }
    segs.forEach(function (seg) {
      const cols = seg.cols.map(columnLetterToNumber);
      const cMin = Math.min.apply(null, cols);
      const cMax = Math.max.apply(null, cols);
      const rows = lastRow - DATA_START_ROW + 1;
      sheet.getRange(DATA_START_ROW, cMin, rows, cMax - cMin + 1).clearContent();
      total++;
    });
  });
  return total;
}

function jalankanSapuBersihManual() {
  try {
    const total = sapuBersihDalamFile(false) + sapuBersihLintasFile(false);
    SpreadsheetApp.getUi().alert(total > 0 ? "Selesai. " + total + " data terisi sebelum waktunya dibersihkan." : "Selesai. Tidak ada data yang perlu dibersihkan.");
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

// [v3] Trigger berkala 5 menit TIDAK lagi menghapus data. Ia hanya menghitung
// dan menampilkan toast peringatan. Penghapusan nyata hanya lewat menu manual.
function sapuBersihOtomatis() {
  if (adalahFileLaporan_(SpreadsheetApp.getActiveSpreadsheet())) return; // pengaman: nonaktif di file Laporan
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { refreshGerbangCache(); } catch (e) {}
  let potensi = 0;
  try { potensi = sapuBersihDalamFile(true) + sapuBersihLintasFile(true); } catch (e) {}
  if (potensi > 0) {
    try {
      ss.toast(potensi + " segmen terisi sebelum waktunya terdeteksi (TIDAK dihapus otomatis). Lengkapi segmen sebelumnya, atau bersihkan manual lewat menu 'Bersihkan Data Tidak Valid Sekarang'.", "Pekan Tools", 12);
    } catch (e) {}
  }
}

function jalankanRefreshGerbang() {
  try {
    refreshGerbangCache();
    SpreadsheetApp.getUi().alert("Cache gerbang berhasil diperbarui.");
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

function pasangJadwalSapuBersihOtomatis() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "sapuBersihOtomatis") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("sapuBersihOtomatis").timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert("Berhasil! Tiap 5 menit sistem memeriksa data yang terisi sebelum waktunya dan memberi peringatan (tanpa menghapus).");
}

// ============== TANGGAL OTOMATIS ==============
function ambilTahunUntukBulan(judul, infoBulan) {
  const m = judul.match(/(\d{4})\s*-\s*(\d{4})/);
  if (m) { const a = parseInt(m[1], 10), b = parseInt(m[2], 10); return infoBulan.index <= 4 ? a : b; }
  const s = judul.match(/(\d{4})/);
  return s ? parseInt(s[1], 10) : new Date().getFullYear();
}

function hitungKelompokPekan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan) return null;
  const bulanAngka = BULAN_KE_ANGKA[infoBulan.kata];
  const tahun = ambilTahunUntukBulan(ss.getName(), infoBulan);
  const jumlahHari = new Date(tahun, bulanAngka, 0).getDate();
  const kelompok = [];
  let p = [];
  for (let d = 1; d <= jumlahHari; d++) {
    const dt = new Date(tahun, bulanAngka - 1, d);
    const hari = dt.getDay();
    if (hari === 0) continue;
    if (hari === 1 && p.length > 0) { kelompok.push(p); p = []; }
    p.push({ tanggal: d, hari: hari });
  }
  if (p.length > 0) kelompok.push(p);
  return kelompok;
}

// ============== BULAN OTOMATIS (ISI NAMA BULAN DI BARIS JUDUL) ==============
// Nama bulan diambil dari NAMA FILE (mis. "Mutaba'ah Agustus 2026-2027" ->
// "Agustus"), lalu ditulis ke sel label "Bulan : " yang MASIH KOSONG di baris
// judul (baris 1) tiap sheet halaqah. Otomatis dijalankan tiap file dibuka
// (onOpen) dan saat menu 'Isi Tanggal Otomatis'. Sel yang SUDAH berisi nilai
// (mis. "Bulan : 2") TIDAK disentuh, supaya tidak menimpa data yang sah.
function isiBulanOtomatisSilent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan) return 0;
  const bulan = kapitalisasi(infoBulan.kata);
  let jumlah = 0;
  ss.getSheets().forEach(function (sheet) {
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const judulBaris = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < lastCol; c++) {
      const v = String(judulBaris[c]).trim();
      if (/^bulan\s*:?\s*$/i.test(v)) {
        sheet.getRange(1, c + 1).setValue("Bulan : " + bulan);
        jumlah++;
      }
    }
  });
  return jumlah;
}

function isiBulanOtomatis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan) {
    SpreadsheetApp.getUi().alert("Nama file tidak memuat nama bulan yang dikenali (mis. 'Agustus'). Tidak ada yang diisi.");
    return;
  }
  const jumlah = isiBulanOtomatisSilent();
  const pesan = jumlah > 0
    ? "Selesai. Nama bulan '" + kapitalisasi(infoBulan.kata) + "' diisi otomatis di " + jumlah + " sel (baris judul tiap sheet)."
    : "Selesai. Tidak ada sel label 'Bulan :' kosong yang ditemukan.";
  try { SpreadsheetApp.getUi().alert(pesan); }
  catch (e) { try { ss.toast(pesan, "Bulan Otomatis", 8); } catch (e2) {} }
}

function isiTanggalOtomatis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan) { SpreadsheetApp.getUi().alert("Nama file tidak memuat nama bulan yang dikenali."); return; }
  const kelompok = hitungKelompokPekan();
  if (!kelompok) { SpreadsheetApp.getUi().alert("Gagal menghitung pekan."); return; }
  if (kelompok.length > TEMPLATE_WEEKS.length) {
    SpreadsheetApp.getUi().alert("Bulan ini punya " + kelompok.length + " pekan, tapi template hanya " + TEMPLATE_WEEKS.length + ". Pekan berikutnya tidak diisi.");
  }
  // Isi juga label "Bulan :" yang masih kosong di baris judul tiap sheet.
  const jumlahBulan = isiBulanOtomatisSilent();
  let jumlahSheet = 0;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sheet) {
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) return;
    const lastRow = Math.max(getBarisDataTerakhir(sheet), DATA_START_ROW);
    for (let p = 0; p < TEMPLATE_WEEKS.length; p++) {
      const week = TEMPLATE_WEEKS[p];
      const kolPertama = columnLetterToNumber(week.cols[0]);
      const tanggalPekan = p < kelompok.length ? kelompok[p] : [];
      for (let k = 0; k < 6; k++) {
        const kolTgl = kolPertama + (2 * k);
        if (k < tanggalPekan.length) {
          sheet.getRange(TANGGAL_ROW, kolTgl).setValue(NAMA_HARI[tanggalPekan[k].hari] + ", " + tanggalPekan[k].tanggal);
        } else {
          sheet.getRange(TANGGAL_ROW, kolTgl).setValue(TANDA_HARI_KOSONG);
          if (lastRow >= DATA_START_ROW) {
            sheet.getRange(DATA_START_ROW, kolTgl, lastRow - DATA_START_ROW + 1, 2).setValue(TANDA_HARI_KOSONG);
          }
        }
      }
    }
    jumlahSheet++;
  });
  SpreadsheetApp.getUi().alert("Selesai! " + jumlahSheet + " sheet diisi tanggalnya" + (jumlahBulan > 0 ? ", dan nama bulan '" + kapitalisasi(infoBulan.kata) + "' diisi di " + jumlahBulan + " sel." : "."));
}

// ============== SAMAKAN NAMA HALAQAH DENGAN BULAN SEBELUMNYA ==============
// Mengambil nama halaqah dari file bulan sebelumnya (mis. Juli) lalu mencocokkan
// ke sheet yang SAMA di file aktif (Agustus). Setiap sheet dibuka berdasar nama
// guru/yang identik; diambil kolom A1 sheet lama, ekstrak nama setelah " : ", dan
// ditulis ulang ke A1 sheet aktif dengan format konsisten "Nama Halaqah : <nama>".
function ekstrakNamaHalaqah(teks) {
  const t = String(teks).trim();
  if (!t) return "";
  const idx = t.indexOf(":");
  return idx === -1 ? t : t.slice(idx + 1).trim();
}

function samakanNamaHalaqah() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoBulan = deteksiBulanDariJudul(ss.getName());
  if (!infoBulan || infoBulan.index === 0) {
    SpreadsheetApp.getUi().alert("Bulan pertama tahun ajaran, tidak ada file sebelumnya untuk disamakan.");
    return;
  }
  const hasil = cariFileBulanSebelumnya(infoBulan);
  if (!hasil.file) {
    SpreadsheetApp.getUi().alert('File bulan sebelumnya "' + hasil.judulCari + '" tidak ditemukan di folder yang sama.');
    return;
  }
  let ssPrev;
  try { ssPrev = SpreadsheetApp.openById(hasil.file.getId()); }
  catch (err) { SpreadsheetApp.getUi().alert("Gagal membuka file bulan sebelumnya: " + err.message); return; }

  const diubah = [];
  const dilewati = [];
  ss.getSheets().forEach(function (sh) {
    if (EXCLUDED_SHEETS.indexOf(sh.getName()) !== -1) return;
    const shPrev = ssPrev.getSheetByName(sh.getName());
    if (!shPrev) { dilewati.push(sh.getName() + " (sheet tidak ada di file lama)"); return; }
    const namaPrev = ekstrakNamaHalaqah(shPrev.getRange(1, 1).getValue());
    if (!namaPrev) { dilewati.push(sh.getName() + " (nama halaqah kosong di file lama)"); return; }
    sh.getRange(1, 1).setValue("Nama Halaqah : " + namaPrev);
    diubah.push(sh.getName() + " -> " + namaPrev);
  });

  let pesan = "Selesai menyamakan " + diubah.length + " halaqah dengan file bulan '" +
    kapitalisasi(namaBulanSebelumnya(infoBulan.index)).toUpperCase() + "'.\n" +
    (diubah.length ? "\n---\n" + diubah.join("\n") : "");
  if (dilewati.length) pesan += "\n\nDilewati:\n" + dilewati.join("\n");
  try { SpreadsheetApp.getUi().alert(pesan); }
  catch (e) { try { ss.toast(pesan, "Sinkronisasi Laporan", 15); } catch (e2) {} }
}

// ============== PERBAIKAN NAMA SANTRI (HASIL AUDIT KELAS) [v8/v9.1] ==============
// Sheet KELAS di file Laporan = acuan ejaan. Struktur Mutaba'ah: 1 santri =
// blok 3 baris mulai baris 5 (SYNC_CONFIG.DATA_START_ROW), nama di kolom B
// baris pertama blok. Idempoten: nama sudah baku / baris sudah terhapus -> SKIP.
const RENAME_MUTABAAH = [
  ['Alkaf',   'Ibrohim bin Donald New',              'Ibrahim Bin Donald Arthur Muhammad'],
  ['Munawar', 'Fathi Dzahabi.S',                     'Fathi Dzahabi S.'],
  ['Farhan',  'Abdurrohman Afif',                    'Abdurrahman Afif'],
  ['Azzam',   'Muhamad Bahtiar Almer Tajusa',        'Muhammad Bachtiar Almer Tajusa'],
  ['Azzam',   'Achmad Fadhil Al zam',                'Achmad Fadhil Al Zam'],
  ['Mundzir', 'Arqam Wadud Affandi',                 'Arqam Wadud'],
  ['Mundzir', 'Muhammad Bin Donald Arthur Muhammad', 'Muhammad bin Donald'],
  ['Adlan',   'Anfhal Zhafir Putra Alfi',            'Anfaal Zhafirputra Alfi'],
  ['Daud',    'Muhammad Baariq Alfaqih Yusup',       'Muhamad Baariq Al Faqih Yusup']
];
const HAPUS_MUTABAAH = [
  ['Ibrahim', 'Cholid'],
  ['Farhan',  'Abdillah Arrafif'],
  ['Adlan',   'Genta Lilo Abimanyu'],
  ['Alwan',   'Ganendra Farzan Pratama']
];
const GANTI_TAB_MUTABAAH = [['Dani', 'Daud'], ['Nadzief', 'Nadzif']];

function perbaikiNamaMutabaah(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  // 0) samakan nama TAB dengan file Laporan (prasyarat pairing sinkronisasi)
  GANTI_TAB_MUTABAAH.forEach(function (p) {
    const lama = ss.getSheetByName(p[0]);
    const baru = ss.getSheetByName(p[1]);
    if (!lama) { log.push('Tab "' + p[0] + '": ' + (baru ? 'sudah "' + p[1] + '" (SKIP)' : 'tidak ada (SKIP)')); return; }
    if (baru) { log.push('Tab "' + p[1] + '" sudah ada - "' + p[0] + '" TIDAK diubah (cek manual!)'); return; }
    lama.setName(p[1]);
    log.push('Tab "' + p[0] + '" -> "' + p[1] + '" OK');
  });
  // 1) rename nilai nama santri
  RENAME_MUTABAAH.forEach(function (it) {
    const sh = ss.getSheetByName(it[0]);
    if (!sh) { log.push(it[0] + ': sheet tidak ada!'); return; }
    const r = cariBarisNamaMutabaah_(sh, it[1]);
    if (!r) {
      log.push(it[0] + ': "' + it[1] + '" -> ' + (cariBarisNamaMutabaah_(sh, it[2]) ? 'SKIP (sudah baku)' : 'TIDAK KETEMU!'));
      return;
    }
    sh.getRange(r, SYNC_CONFIG.COL_NAMA).setValue(it[2]);
    log.push(it[0] + ': "' + it[1] + '" -> "' + it[2] + '" OK');
  });
  // 2) hapus santri tak dikenal / duplikat (utuh 3 baris per blok)
  HAPUS_MUTABAAH.forEach(function (it) {
    const sh = ss.getSheetByName(it[0]);
    if (!sh) { log.push(it[0] + ': sheet tidak ada!'); return; }
    const r = cariBarisNamaMutabaah_(sh, it[1]);
    if (!r) { log.push(it[0] + ': hapus "' + it[1] + '" SKIP'); return; }
    const jml = Math.min(SYNC_CONFIG.ROWS_PER_STUDENT, sh.getMaxRows() - r + 1);
    sh.deleteRows(r, jml);
    rapikanNoMutabaah_(sh);
    log.push(it[0] + ': hapus "' + it[1] + '" (' + jml + ' baris) OK');
  });
  // 3) [v9] rapikan kolom No SEMUA sheet halaqah (bukan hanya yang ada hapusan):
  // menyembuhkan nomor dobel/lompat & nomor pada baris label tanpa menunggu hapus.
  ss.getSheets().forEach(function (sh) {
    if (isHalaqahSheet(sh)) rapikanNoMutabaah_(sh);
  });
  log.push('Penomoran semua sheet halaqah dirapikan (v9).');
  Logger.log('PERBAIKAN NAMA MUTABAAH:\n' + log.join('\n'));
  const pesan = log.length ? log.join('\n') : 'Tidak ada perubahan.';
  try {
    if (silent === true) ss.toast(pesan, 'Perbaikan Nama Santri', 10);
    else SpreadsheetApp.getUi().alert('Hasil Perbaikan Nama Santri:\n\n' + pesan);
  } catch (e) {}
  return pesan;
}

function cariBarisNamaMutabaah_(sheet, nama) {
  // [v9.1] Pindai tiap baris dengan SATU pembacaan batch kolom Nama:
  // blok santri yang tidak rata kelipatan 3 tetap ketemu, tanpa ribuan
  // panggilan API per sel (penyebab eksekusi lama di v9).
  const last = Math.max(getBarisDataTerakhir(sheet), sheet.getLastRow());
  const jml = last - SYNC_CONFIG.DATA_START_ROW + 1;
  if (jml < 1) return 0;
  const namaArr = sheet.getRange(SYNC_CONFIG.DATA_START_ROW, SYNC_CONFIG.COL_NAMA, jml, 1).getValues();
  for (let i = 0; i < namaArr.length; i++) {
    if (String(namaArr[i][0]).trim() === nama) return SYNC_CONFIG.DATA_START_ROW + i;
  }
  return 0;
}

function rapikanNoMutabaah_(sheet) {
  // [v9.1] Pindai TIAP baris (bukan langkah 3): nomor dobel/lompat akibat blok
  // yang bergeser hilang, dan nomor yang menempel di baris label
  // (I Z I N / S A K I T / A L P A / T D K SETOR) ikut dikosongkan.
  // Dikerjakan BATCH: 1x baca kolom No+Nama, maksimal 1x tulis kolom No,
  // sehingga eksekusi hitungan detik (v9 per-sel membuatnya ber menit).
  const last = Math.max(getBarisDataTerakhir(sheet), sheet.getLastRow());
  const jml = last - SYNC_CONFIG.DATA_START_ROW + 1;
  if (jml < 1) return;
  const blok = sheet.getRange(SYNC_CONFIG.DATA_START_ROW, 1, jml, Math.max(2, SYNC_CONFIG.COL_NO, SYNC_CONFIG.COL_NAMA)).getValues();
  let no = 0, adaPerubahan = false;
  for (let i = 0; i < jml; i++) {
    const nama = String(blok[i][SYNC_CONFIG.COL_NAMA - 1]).trim();
    if (!nama) continue;
    let target;
    if (isNamaTidakValid(nama)) target = "";
    else { no++; target = no; }
    const sekarang = blok[i][SYNC_CONFIG.COL_NO - 1];
    const sama = (sekarang === null || sekarang === undefined) ? (target === "") : String(sekarang) === String(target);
    if (!sama) {
      blok[i][SYNC_CONFIG.COL_NO - 1] = target;
      adaPerubahan = true;
    }
  }
  if (adaPerubahan) {
    sheet.getRange(SYNC_CONFIG.DATA_START_ROW, SYNC_CONFIG.COL_NO, jml, 1).setValues(
      blok.map(function (baris) { return [baris[SYNC_CONFIG.COL_NO - 1]]; })
    );
  }
}

// ============== SINKRONISASI KE MASTER LAPORAN ==============
function cariFileLaporanBulanIni() {
  const fileIni = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const judulLaporan = fileIni.getName().replace(/Mutaba['’]?ah/i, "Laporan Bulan");
  let fileTarget = null;
  const folders = fileIni.getParents();
  if (folders.hasNext()) {
    const c = folders.next().getFilesByName(judulLaporan);
    if (c.hasNext()) fileTarget = c.next();
  }
  if (!fileTarget) {
    const c2 = DriveApp.getFilesByName(judulLaporan);
    if (c2.hasNext()) fileTarget = c2.next();
  }
  return { file: fileTarget, judulLaporan: judulLaporan };
}

function syncSemuaHalaqah() {
  // [v9.1-gabungan] Bakukan nama santri dulu supaya yang terkirim ke Laporan
  // identik dengan kelas (ejaan baku, jumlah baris benar).
  try { perbaikiNamaMutabaah(true); } catch (e) { Logger.log('perbaikiNamaMutabaah gagal: ' + e); }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = cariFileLaporanBulanIni();
  if (!hasil.file) { SpreadsheetApp.getUi().alert("File Laporan tidak ditemukan (cari: \"" + hasil.judulLaporan + "\")."); return; }
  const ssLaporan = SpreadsheetApp.openById(hasil.file.getId());
  let jumlah = 0;
  ss.getSheets().forEach(function (sheet) {
    if (isHalaqahSheet(sheet)) { syncHalaqah(sheet.getName(), ssLaporan); jumlah++; }
  });
  SpreadsheetApp.getUi().alert('Sinkronisasi selesai untuk ' + jumlah + ' halaqah.');
}

function syncHalaqah(namaHalaqah, ssLaporanOpt) {
  const ssSumber = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSumber = ssSumber.getSheetByName(namaHalaqah);
  if (!sheetSumber) return;
  let ssLaporan = ssLaporanOpt;
  if (!ssLaporan) {
    const hasil = cariFileLaporanBulanIni();
    if (!hasil.file) return;
    ssLaporan = SpreadsheetApp.openById(hasil.file.getId());
  }
  const sheetLaporan = ssLaporan.getSheetByName(namaHalaqah);
  if (!sheetLaporan) return;

  const lastCol = sheetSumber.getLastColumn();
  const lastRow = getBarisDataTerakhir(sheetSumber);
  const judulRowVals = sheetSumber.getRange(SYNC_CONFIG.HEADER_ROW - 1, 1, 1, lastCol).getDisplayValues()[0];
  const headerRowVals = sheetSumber.getRange(SYNC_CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const subHeaderVals = sheetSumber.getRange(SYNC_CONFIG.SUBHEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];

  const colsJumlah = findColumns(subHeaderVals, 'Jumlah');
  // 'Jumlah Hafalan Keseluruhan' ada di baris judul (baris 2); fallback ke baris header.
  let colsTotalHafalan = findColumns(judulRowVals, 'Jumlah Hafalan Keseluruhan');
  if (!colsTotalHafalan.length) colsTotalHafalan = findColumns(headerRowVals, 'Jumlah Hafalan Keseluruhan');
  // Izin/Sakit/Alpa/Tidak Setor diambil HANYA dari blok 'Akumulasi Absen' (baris judul),
  // supaya tidak dobel dengan label yang sama di blok per-pekan pada baris header.
  const colsAkumulasiAbsen = findColumns(judulRowVals, 'Akumulasi Absen');
  let colsIzin = [], colsSakit = [], colsAlpa = [], colsTidakSetor = [];
  if (colsAkumulasiAbsen.length === 1) {
    const startA = colsAkumulasiAbsen[0];
    const span = Math.min(lastCol - startA + 1, 8);
    const subAbsen = sheetSumber.getRange(SYNC_CONFIG.HEADER_ROW, startA, 1, span).getDisplayValues()[0];
    subAbsen.forEach(function (v, i) {
      const t = String(v).trim();
      if (t === 'Izin') colsIzin.push(startA + i);
      else if (t === 'Sakit') colsSakit.push(startA + i);
      else if (t === 'Alpa') colsAlpa.push(startA + i);
      else if (t === 'Tidak Setor') colsTidakSetor.push(startA + i);
    });
  }
  if (!colsIzin.length || !colsSakit.length || !colsAlpa.length || !colsTidakSetor.length) {
    // fallback: perilaku lama bila blok akumulasi tidak lengkap ditemukan
    colsIzin = findColumns(headerRowVals, 'Izin');
    colsSakit = findColumns(headerRowVals, 'Sakit');
    colsAlpa = findColumns(headerRowVals, 'Alpa');
    colsTidakSetor = findColumns(headerRowVals, 'Tidak Setor');
  }
  const colsCatatan = findColumns(headerRowVals, 'Catatan Perkembangan');

  const data = sheetSumber.getRange(1, 1, lastRow, lastCol).getValues();
  const results = [];
  for (let r = SYNC_CONFIG.DATA_START_ROW - 1; r < lastRow; r += SYNC_CONFIG.ROWS_PER_STUDENT) {
    const row = data[r];
    const nama = row[SYNC_CONFIG.COL_NAMA - 1];
    if (!nama || isNamaTidakValid(nama)) continue;
    const no = row[SYNC_CONFIG.COL_NO - 1];
    const grade = row[SYNC_CONFIG.COL_GRADE - 1];
    const targetBulanan = Number(row[SYNC_CONFIG.COL_TARGET_TOTAL - 1]) || 0;

    let targetTercapai = 0;
    colsJumlah.forEach(function (c) { targetTercapai += Number(row[c - 1]) || 0; });
    let izin = 0, sakit = 0, alpa = 0, tidakSetor = 0;
    colsIzin.forEach(function (c) { izin += Number(row[c - 1]) || 0; });
    colsSakit.forEach(function (c) { sakit += Number(row[c - 1]) || 0; });
    colsAlpa.forEach(function (c) { alpa += Number(row[c - 1]) || 0; });
    colsTidakSetor.forEach(function (c) { tidakSetor += Number(row[c - 1]) || 0; });

    let jumHafalan = '';
    for (let c = colsTotalHafalan.length - 1; c >= 0; c--) { const v = row[colsTotalHafalan[c] - 1]; if (v !== "" && v !== null) { jumHafalan = v; break; } }
    const catatan = colsCatatan.map(function (c) { return row[c - 1]; }).filter(function (v) { return v; }).join(' | ');
    const presentase = targetBulanan > 0 ? (targetTercapai / targetBulanan) : 0;

    results.push({ no: no, nama: nama, grade: grade, targetBulanan: targetBulanan, targetTercapai: targetTercapai, presentase: presentase, izin: izin, sakit: sakit, alpa: alpa, tidakSetor: tidakSetor, jumHafalan: jumHafalan, catatan: catatan });
  }

  writeToLaporan(sheetLaporan, results);
}

function findColumns(headerArr, text) {
  const cols = [];
  headerArr.forEach(function (v, i) { if (String(v).trim() === text) cols.push(i + 1); });
  return cols;
}

function writeToLaporan(sheetLaporan, results) {
  const startRow = SYNC_CONFIG.LAPORAN_DATA_START_ROW;
  const C = SYNC_CONFIG.LAPORAN_COLS;
  const lastColL = Math.max.apply(null, Object.keys(C).map(function (k) { return C[k]; }));
  const existing = sheetLaporan.getLastRow();
  const rowsToClear = Math.max(existing - startRow + 1, results.length);
  // Kolom Keterangan (G / kolom 7) TIDAK ikut dikosongkan: isinya rumus
  // kategori otomatis milik file Laporan (formatHalaqah_), bukan data sync.
  if (rowsToClear > 0) {
    sheetLaporan.getRange(startRow, 1, rowsToClear, C.KETERANGAN - 1).clearContent();
    sheetLaporan.getRange(startRow, C.IZIN, rowsToClear, lastColL - C.IZIN + 1).clearContent();
  }

  results.forEach(function (s, idx) {
    const r = startRow + idx;
    // [v9.1-gabungan] Nomor dari posisi baris tujuan (idx+1), bukan nomor sumber,
    // supaya sheet Laporan tidak mewarisi nomor dobel/lompat milik sumber.
    sheetLaporan.getRange(r, C.NO).setValue(idx + 1);
    sheetLaporan.getRange(r, C.NAMA).setValue(s.nama);
    // [v9.1-gabungan] Trim nilai Grade agar varian ber-spasi seperti "C " tidak mengganggu rumus.
    sheetLaporan.getRange(r, C.GRADE).setValue(String(s.grade == null ? '' : s.grade).trim());
    sheetLaporan.getRange(r, C.TARGET_BULANAN).setValue(s.targetBulanan);
    sheetLaporan.getRange(r, C.TARGET_TERCAPAI).setValue(s.targetTercapai);
    sheetLaporan.getRange(r, C.PRESENTASE).setValue(s.presentase);
    sheetLaporan.getRange(r, C.IZIN).setValue(s.izin);
    sheetLaporan.getRange(r, C.SAKIT).setValue(s.sakit);
    sheetLaporan.getRange(r, C.ALPA).setValue(s.alpa);
    sheetLaporan.getRange(r, C.TIDAK_SETOR).setValue(s.tidakSetor);
    sheetLaporan.getRange(r, C.JUMLAH_HAFALAN).setValue(s.jumHafalan);
    sheetLaporan.getRange(r, C.CATATAN).setValue(s.catatan);
  });
}

// ============== PERBAIKI REFERENSI RUMUS RUSAK (#REF!) ==============
// [v9.2] Auto-perbaikan HANYA untuk fungsi agregat variadic (SUM, AVERAGE,
// COUNT, COUNTA, MAX, MIN) yang hasilnya tetap sah bila satu argumen dibuang.
// Fungsi berargumen-posisi (VLOOKUP, IF, INDEX, dll.) TIDAK dibersihkan
// otomatis, karena membuang argumen tengah menggeser argumen lain -> nilai
// SALAH tanpa tanda error. Semua kasus lain jatuh ke jalur "cek manual".
const FUNGSI_AMAN_REF = ["SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN"];

function buangReferensiRusakDariFormula(formula) {
  if (!/#ref!/i.test(formula)) return null;
  // Kenali fungsi paling luar, mis. "=SUM(" atau "=ArrayFormula(sum("
  const m = formula.match(/^=(?:ARRAYFORMULA\()?([A-Za-z][A-Za-z0-9_.]*)\(/i);
  if (!m || FUNGSI_AMAN_REF.indexOf(m[1].toUpperCase()) === -1) return null;
  let hasil = formula;
  hasil = hasil.replace(/;\s*#ref!/gi, '').replace(/#ref!\s*;/gi, '')
               .replace(/,\s*#ref!/gi, '').replace(/#ref!\s*,/gi, '');
  if (/#ref!/i.test(hasil)) return null;
  return hasil;
}

function tambahCatatanRef_(sel, teks) {
  // [v9.2] Jangan menimpa catatan lama; catatan baru disisipkan DI ATASnya.
  const lama = sel.getNote();
  sel.setNote(lama ? teks + "\n---\n" + lama : teks);
}

function scanPerbaikiSheet(sheet) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { diperbaiki: 0, manual: 0 };
  const formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let diperbaiki = 0, manual = 0;
  for (let r = 0; r < lastRow; r++) for (let c = 0; c < lastCol; c++) {
    const formula = formulas[r][c];
    if (!formula) continue;
    const val = values[r][c];
    if (typeof val !== 'string' || val.indexOf('#REF!') === -1) continue;
    const sel = sheet.getRange(r + 1, c + 1);
    const bersih = buangReferensiRusakDariFormula(formula);
    if (bersih) { sel.setFormula(bersih); tambahCatatanRef_(sel, "Diperbaiki: #REF! dibuang dari rumus.\n\nBahasa:\n" + formula); diperbaiki++; }
    else { tambahCatatanRef_(sel, "#REF! tidak bisa dibuang otomatis (fungsi bukan agregat aman / pola tak dikenali), cek manual:\n" + formula); manual++; }
  }
  return { diperbaiki: diperbaiki, manual: manual };
}

function perbaikiSemuaRefError() {
  let td = 0, tm = 0;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    if (EXCLUDED_SHEETS.indexOf(sh.getName()) !== -1) return;
    const res = scanPerbaikiSheet(sh);
    td += res.diperbaiki; tm += res.manual;
  });
  return { diperbaiki: td, manual: tm };
}

function jalankanPerbaikiRefErrorManual() {
  try {
    const h = perbaikiSemuaRefError();
    if (h.diperbaiki === 0 && h.manual === 0) { SpreadsheetApp.getUi().alert("Selesai. Tidak ada rumus error (#REF!)."); return; }
    let pesan = "Selesai.\n";
    if (h.diperbaiki > 0) pesan += h.diperbaiki + " rumus diperbaiki.\n";
    if (h.manual > 0) pesan += h.manual + " rumus perlu dicek manual.\n";
    SpreadsheetApp.getUi().alert(pesan);
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

// ============== SEMBUNYIKAN ERROR RUMUS (#VALUE! DLL.) [v9.4-BATCH] ==============
// Sel ber-rumus yang menampilkan teks error (#VALUE!, #DIV/0!, #N/A, #NUM!,
// #NAME?, #NULL!) dibungkus IFERROR dengan fallback angka 0, sehingga tampilan
// otomatis ikut format sel: kolom persen -> 0%, kolom angka -> 0. Idempoten:
// rumus yang sudah dibungkus IFERROR dilewati. #REF! SENGAJA tidak disentuh di
// sini - tetap lewat jalur 'Perbaiki Referensi Rumus Error' agar akar masalah
// referensi terlihat dan bisa diperbaiki manual.
// [v9.4] Dipercepat: nilai + rumus + peta merge dibaca SEKALI per sheet, lalu
// penulisan dikemas menjadi MAKSIMAL SATU panggilan setFormulas per baris yang
// berubah (dulu: 2-3 panggilan API per sel -> eksekusi menit-an).
const DAFTAR_ERROR_TAMPILAN = ["#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#NULL!", "#ERROR!", "#CYCLE?"];

function apakahNilaiErrorTampilan_(v) {
  if (typeof v !== 'string') return false;
  const t = v.trim().toUpperCase();
  if (t.charAt(0) !== '#') return false;
  for (let i = 0; i < DAFTAR_ERROR_TAMPILAN.length; i++) {
    if (t.indexOf(DAFTAR_ERROR_TAMPILAN[i]) !== -1) return true;
  }
  return false;
}

function bungkusFormulaAman_(formula) {
  // Return rumus baru, atau null bila sudah aman / tak layak dibungkus.
  const isi = String(formula).replace(/^\s*=/, '');
  if (/^\s*(?:ARRAYFORMULA\s*\(\s*)?IFERROR\s*\(/i.test(isi)) return null;
  const pemisah = isi.indexOf(';') !== -1 ? ';' : (isi.indexOf(',') !== -1 ? ',' : ';');
  if (/^\s*ARRAYFORMULA\s*\(/i.test(isi)) {
    // Bungkus DI DALAM array supaya elemen array lain tidak ikut hilang.
    const dalam = isi.replace(/^\s*ARRAYFORMULA\s*\(/i, '').replace(/\)\s*$/, '');
    return '=ARRAYFORMULA(IFERROR(' + dalam + pemisah + '0))';
  }
  return '=IFERROR(' + isi + pemisah + '0)';
}

function sembunyikanErrorDiSheet_(sheet) {
  // [v9.5] LAZY READ: baca NILAI saja dulu; jika tidak ada sel ber-error,
  // selesai tanpa getFormulas()/getMergedRanges()/penulisan sama sekali.
  const range = sheet.getDataRange();
  const lastRow = range.getLastRow(), lastCol = range.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { dibungkus: 0 };
  const values = range.getValues();

  // Pra-scan murah: kumpulkan kandidat sel yang menampilkan error.
  const kandidat = [];
  for (let r = 0; r < lastRow; r++) {
    const vals = values[r];
    for (let c = 0; c < lastCol; c++) {
      if (apakahNilaiErrorTampilan_(vals[c])) kandidat.push([r, c]);
    }
  }
  if (!kandidat.length) return { dibungkus: 0 };

  // Ada kandidat: baru baca rumus + peta merge sekali untuk seluruh area
  // (sel non-jangkar merge tidak boleh ditulis).
  const formulas = range.getFormulas();
  const nonJangkar = {};
  range.getMergedRanges().forEach(function (mr) {
    const r0 = mr.getRow(), c0 = mr.getColumn();
    for (let r = r0; r < r0 + mr.getNumRows(); r++) {
      for (let c = c0; c < c0 + mr.getNumColumns(); c++) {
        if (r !== r0 || c !== c0) nonJangkar[r + ',' + c] = true;
      }
    }
  });

  // Kumpulkan perubahan per baris.
  const barisUbah = {}; // r -> { ubah: {kolom: rumusBaru}, kolom: [kolom,...] }
  let total = 0;
  for (let i = 0; i < kandidat.length; i++) {
    const r = kandidat[i][0] + 1, c = kandidat[i][1] + 1;
    const f = formulas[r - 1][c - 1];
    if (!f) continue;
    const baru = bungkusFormulaAman_(f);
    if (!baru || nonJangkar[r + ',' + c]) continue;
    const rec = barisUbah[r] || (barisUbah[r] = { ubah: {}, kolom: [] });
    rec.ubah[c] = baru;
    rec.kolom.push(c);
    total++;
  }
  if (!total) return { dibungkus: 0 };

  // Tulis batch per baris: satu segmen berdempet yang memuat rumus baru +
  // padding aman (rumus asli / nilai beku). Segmen dipotong sebelum kolom
  // ber-teks berisiko (=,+,-,@ tanpa rumus) supaya tidak tersentuh sama sekali.
  Object.keys(barisUbah).forEach(function (k) {
    const r = Number(k), rec = barisUbah[k];
    const vals = values[r - 1], frms = formulas[r - 1];
    const buf = [];
    let segStart = -1;
    const flush = function (endC) {
      if (!buf.length) return;
      sheet.getRange(r, segStart, 1, endC - segStart + 1).setFormulas([buf]);
      buf.length = 0;
    };
    const minC = rec.kolom[0], maxC = rec.kolom[rec.kolom.length - 1];
    for (let c = minC; c <= maxC; c++) {
      if (rec.ubah[c]) {
        if (segStart === -1) segStart = c;
        buf.push(rec.ubah[c]);
        continue;
      }
      if (segStart === -1) continue;
      const fAsli = frms[c - 1];
      if (fAsli) { buf.push(fAsli); continue; }
      if (bisakahDibekukan_(vals[c - 1])) { buf.push(bekukanNilai_(vals[c - 1])); continue; }
      flush(c - 1);
      segStart = -1;
    }
    flush(maxC);
  });
  return { dibungkus: total };
}

function bisakahDibekukan_(v) {
  return !(typeof v === 'string' && /^[=+@-]/.test(v));
}

function bekukanNilai_(v) {
  // Pad dalam pengiriman batch: sel tanpa rumus dikirim ulang sebagai literal
  // agar isinya tidak berubah; string kosong = sel kosong (aman).
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return String(v).replace('.', ',');
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    try { return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch (e) { return ''; }
  }
  return String(v);
}

function sembunyikanErrorRumusSemua() {
  let total = 0;
  const ringkas = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    if (EXCLUDED_SHEETS.indexOf(sh.getName()) !== -1) return;
    const h = sembunyikanErrorDiSheet_(sh);
    if (h.dibungkus) { total += h.dibungkus; ringkas.push(sh.getName() + ': ' + h.dibungkus + ' sel'); }
  });
  return { total: total, ringkas: ringkas };
}

function jalankanSembunyikanErrorRumus() {
  try {
    const h = sembunyikanErrorRumusSemua();
    const pesan = h.total === 0
      ? "Selesai (v9.5). Tidak ada sel rumus yang sedang menampilkan error."
      : "Selesai (v9.5). " + h.total + " sel rumus error dibungkus IFERROR (tampil 0 / 0%):\n\n" + h.ringkas.join("\n");
    SpreadsheetApp.getUi().alert(pesan);
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

function jalankanSembunyikanErrorRumusAktif() {
  // [v9.4] Perbaikan cepat untuk SATU sheet aktif saja.
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) {
      SpreadsheetApp.getUi().alert('Sheet "' + sheet.getName() + '" dikecualikan dari pemindaian.');
      return;
    }
    const h = sembunyikanErrorDiSheet_(sheet);
    SpreadsheetApp.getUi().alert(h.dibungkus === 0
      ? 'Selesai (v9.5). Tidak ada sel rumus yang sedang menampilkan error di sheet "' + sheet.getName() + '".'
      : 'Selesai (v9.5). ' + h.dibungkus + ' sel rumus error di sheet "' + sheet.getName() + '" dibungkus IFERROR (tampil 0 / 0%).');
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

// ============== DIAGNOSTIK ==============
// Baca blok sekaligus + peta nilai dari merge (sekali getValues/getMergedRanges).
// Lookup per sel jadi O(1), bukan loop semua merge per sel kosong.
function bacaBlok(sheet, start, jmlBaris, colMin, jmlKolom) {
  const range = sheet.getRange(start, colMin, jmlBaris, jmlKolom);
  const block = range.getValues();
  const merges = range.getMergedRanges();
  const peta = {};
  for (let m = 0; m < merges.length; m++) {
    const mr = merges[m];
    const v = mr.getCell(1, 1).getValue();
    const r0 = mr.getRow() - start;
    const c0 = mr.getColumn() - colMin;
    for (let rr = 0; rr < mr.getNumRows(); rr++)
      for (let cc = 0; cc < mr.getNumColumns(); cc++)
        peta[(r0 + rr) + "," + (c0 + cc)] = v;
  }
  return { block: block, peta: peta };
}

function nilaiSel(out, r, c) {
  const k = r + "," + c;
  return out.peta.hasOwnProperty(k) ? out.peta[k] : out.block[r][c];
}

// Cari sel kosong per siswa per segmen pada satu sheet. Memakai aturan yang
// sama dengan validasi (semua kolom harus terisi, merge-aware) — DIOPTIMASI.
function diagnosaKosongSheet(namaSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(namaSheet);
  if (!sheet) return [];
  const segments = getSegmentsForSheet(sheet);
  const lastRow = getBarisDataTerakhir(sheet);
  const start = DATA_START_ROW;
  if (lastRow < start) return [];
  const n = lastRow - start + 1;
  const nameCol = columnLetterToNumber(NAME_COL);
  const namaAwal = bacaBlok(sheet, start, n, nameCol, 1);
  const hasil = [];
  segments.forEach(function (seg) {
    const colNums = seg.cols.map(columnLetterToNumber);
    const cMin = Math.min.apply(null, colNums);
    const cMax = Math.max.apply(null, colNums);
    const out = bacaBlok(sheet, start, n, cMin, cMax - cMin + 1);
    for (let i = 0; i < n; i++) {
      const nama = nilaiSel(namaAwal, i, 0);
      if (isNamaTidakValid(nama)) continue;
      const rowNum = start + i;
      const kosong = [];
      for (let k = 0; k < colNums.length; k++) {
        const v = nilaiSel(out, i, colNums[k] - cMin);
        if (v === "" || v === null) kosong.push(seg.cols[k]);
      }
      if (kosong.length === 0) continue;
      hasil.push({
        segmen: seg.label,
        nama: String(nama).replace(/\s+/g, " ").trim(),
        baris: rowNum,
        kosong: kosong
      });
    }
  });
  return hasil;
}

function jalankanDiagnosisKosong() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const namaSheet = sheet.getName();
    const data = diagnosaKosongSheet(namaSheet);
    const outName = "Diagnosis_" + namaSheet;
    let out = ss.getSheetByName(outName);
    if (!out) out = ss.insertSheet(outName);
    else out.clear();
    out.getRange(1, 1).setValue("Segmen");
    out.getRange(1, 2).setValue("Siswa (baris)");
    out.getRange(1, 3).setValue("Kolom belum terisi");
    if (data.length) {
      const vals = data.map(function (d) {
        return [d.segmen, d.nama + " (baris " + d.baris + ")", d.kosong.join(", ")];
      });
      out.getRange(2, 1, vals.length, 3).setValues(vals);
    }
    out.getRange(1, 1, 1, 3).setFontWeight("bold");
    out.setColumnWidth(1, 140);
    out.setColumnWidth(2, 180);
    out.setColumnWidth(3, 300);
    SpreadsheetApp.getUi().alert(
      "Selesai. Hasil ditulis ke sheet \"" + outName + "\".\nKolom C = kolom yang masih kosong per baris siswa.\n(Semua baris yang ditulis berarti masih belum lengkap; baris cil yang TIDAK muncul sudah terisi penuh.)"
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message);
  }
}

function tesSegmen() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) { SpreadsheetApp.getUi().alert("Sheet " + sheet.getName() + " berada di EXCLUDED_SHEETS (tidak divalidasi)."); return; }
    const segs = getSegmentsForSheet(sheet);
    let pesan = "Sheet: " + sheet.getName() + "\n\n";
    segs.forEach(function (seg) { pesan += seg.label + ": " + (apakahSegmentLengkap(sheet, seg) ? "TERISI" : "BELUM TERISI") + "\n"; });
    const g = cekGerbangUntukSheet(sheet.getName());
    pesan += "\nBulan lalu: " + (g.status || "-");
    if (g.warning) pesan += "\n\n⚠️ PERINGATAN: " + g.warning;
    SpreadsheetApp.getUi().alert(pesan);
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

// ============== HAPUS SEMUA RENTANG / SHEET DILINDUNGI ==============
function hapusSemuaProteksi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalRange = 0, totalSheet = 0;
  ss.getSheets().forEach(function (sh) {
    const protRanges = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protRanges.forEach(function (p) {
      const nama = p.getDescription() || sh.getName() + "!" + p.getRange().getA1Notation();
      p.remove();
      totalRange++;
    });
    const protSheets = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protSheets.forEach(function (p) {
      p.remove();
      totalSheet++;
    });
  });
  return { totalRange: totalRange, totalSheet: totalSheet };
}

function jalankanHapusSemuaProteksi() {
  try {
    const hasil = hapusSemuaProteksi();
    const pesan = hasil.totalRange === 0 && hasil.totalSheet === 0
      ? "Selesai. Tidak ada rentang/sheet yang dilindungi di spreadsheet ini."
      : "Selesai. Dihapus: " + hasil.totalRange + " rentang dilindungi, " + hasil.totalSheet + " sheet dilindungi.";
    SpreadsheetApp.getUi().alert(pesan);
  } catch (err) { SpreadsheetApp.getUi().alert("TERJADI ERROR:\n\n" + err.message); }
}

/*************************************************************
 * PERBAIKAN MUTABA'AH — Laporan Bulan Juli 2026/2027
 * Digenerate otomatis dari audit dua spreadsheet:
 *  - Sumber nama/kelas : Data Santri (Perpulangan) 2026/2027
 *  - Sumber nilai      : 15 sheet halaqah (ustadz)
 *
 * CATATAN LOCALE:
 *  Formula ditulis dengan pemisah argumen yang DIDETEKSI OTOMATIS
 *  ("," utk locale en, ";" utk locale lain spt Indonesia). Deteksi
 *  empiris: sel uji =SUM(1<sep>2) di MasterData AD2000.
 *
 * CARA PAKAI:
 *  1. Buka spreadsheet Laporan Bulan Juli -> Ekstensi > Apps Script
 *  2. HAPUS SEMUA file .gs lama / pastikan hanya ADA SATU file berisi kode ini
 *     (Ctrl+A -> Delete -> tempel seluruh file ini), Save
 *  3. Pilih fungsi perbaikiSemua -> Run -> izinkan akses
 *  4. Cek log eksekusi (tanda obat/Executions) - ada verifikasi otomatis
 *  5. Kembali ke spreadsheet, cek sheet 7A..12, Rekap, Persentase Total
 *
 * Aman dijalankan berulang (semua sheet dibangun ulang).
 *************************************************************/

var HALAQAH = ['Alkaf','Nanda','Ibrahim','Husnu','Bagus','Munawar','Farhan','Syahrul','Azzam','Daud','Mundzir','Adlan','Alwan','Azhar','Nadzif'];
var KELAS_LAMA = ["Kelas 7","Kelas 8","Kelas 9","Kelas 10","Kelas 10 New","Kelas 11"];
var TABS = ["7A","7B","8","9A","9B","10A","10B","11A","11B","12"];
var BULAN = 'Juli';
var SEP = ','; // pemisah argumen formula; dideteksi otomatis saat run

// {t: tab kelas, n: nama kanonik, s: sheet halaqah (null = belum ada data),
//  k: kunci lookup persis seperti di sheet halaqah, sc: true = lookup scoped ke sheet ustadz}
var SANTRI = [{"t":"7A","n":"Abdullah Azzam","s":"Adlan","k":"Abdullah Azzam","sc":true},{"t":"7A","n":"Alfyzhar Fahrizsyah Abbad","s":"Adlan","k":"Alfyzhar Fahrizsyah Abbad","sc":false},{"t":"7A","n":"Ayyash Aydin Athallah","s":"Adlan","k":"Ayyash Aydin Athallah","sc":false},{"t":"7A","n":"Ezio Umara Al Ayyubi","s":"Adlan","k":"Ezio Umara Al Ayyubi","sc":false},{"t":"7A","n":"Ganendra Farzan Pratama","s":"Adlan","k":"Ganendra Farzan Pratama","sc":true},{"t":"7A","n":"Kenzo Abrizam Ar Raafi Sadono","s":"Azhar","k":"Kenzo Abrizam Ar Raafi Sadono","sc":false},{"t":"7A","n":"Khalifi Aris Andriyanto","s":"Azhar","k":"Khalifi Aris Andriyanto","sc":false},{"t":"7A","n":"Muhammad","s":"Azhar","k":"Muhammad","sc":false},{"t":"7A","n":"Muhammad Al Fatih","s":"Azhar","k":"Muhammad Al Fatih","sc":true},{"t":"7A","n":"Muhammad Alfatih","s":"Azhar","k":"Muhammad Alfatih","sc":true},{"t":"7A","n":"Muhammad Sarfaraz Rafisqy","s":"Nadzif","k":"Muhammad Sarfaraz Rafisqy","sc":false},{"t":"7A","n":"Syabil Zhafran Rafani","s":"Nadzif","k":"Syabil Zhafran Rafani","sc":false},{"t":"7A","n":"Syafiq Bin Abdul Gofur","s":"Nadzif","k":"Syafiq Bin Abdul Gofur","sc":false},{"t":"7A","n":"Umar Qaisar Alfarouq Wilman","s":"Nadzif","k":"Umar Qaisar Alfarouq Wilman","sc":false},{"t":"7A","n":"Zhafran Zhafify El Musyaffa","s":"Azhar","k":"Zhafran Zhafify El Musyaffa","sc":false},{"t":"7B","n":"Ahmad Ayyas Addzaki","s":"Adlan","k":"Ahmad Ayyas Addzaki","sc":false},{"t":"7B","n":"Arroyan Muhammad Ramadan","s":"Adlan","k":"Arroyan Muhammad Ramadan","sc":false},{"t":"7B","n":"Arya Bima Sena","s":"Adlan","k":"Arya Bima Sena","sc":false},{"t":"7B","n":"Asyraf Ahero Ramadhan","s":"Adlan","k":"Asyraf Ahero Ramadhan","sc":false},{"t":"7B","n":"Haufanza Ahmad Febyandima","s":"Azhar","k":"Haufanza Ahmad Febyandima","sc":false},{"t":"7B","n":"Irsyad Khairuddin","s":"Azhar","k":"Irsyad Khairuddin","sc":false},{"t":"7B","n":"Laksmana Alkairo Aimery Oktavianto","s":"Azhar","k":"Laksmana Alkairo Aimery Oktavianto","sc":false},{"t":"7B","n":"Muhammad Fathir Alghifari","s":"Azhar","k":"Muhammad Fathir Alghifari","sc":false},{"t":"7B","n":"Muhammad Harits Fadli","s":"Azhar","k":"Muhammad Harits Fadli","sc":false},{"t":"7B","n":"Muhammad Nabil Fauzan","s":"Nadzif","k":"Muhammad Nabil Fauzan","sc":false},{"t":"7B","n":"Muhammad Naufal Azmi","s":"Nadzif","k":"Muhammad Naufal Azmi","sc":false},{"t":"7B","n":"Muhammad Salman Al Fatih","s":"Nadzif","k":"Muhammad Salman Al Fatih","sc":false},{"t":"7B","n":"Rayyan Sani Alfarizi","s":"Nadzif","k":"Rayyan Sani Alfarizi","sc":false},{"t":"7B","n":"Shabri Siraj Qalbi","s":"Nadzif","k":"Shabri Siraj Qalbi","sc":false},{"t":"7B","n":"Utsman Nasution","s":"Nadzif","k":"Utsman Nasution","sc":false},{"t":"8","n":"Ahmad Maulana Muklis","s":"Syahrul","k":"Ahmad Maulana Muklis","sc":false},{"t":"8","n":"Anfaal Zhafirputra Alfi","s":"Adlan","k":"Anfaal Zhafirputra Alfi","sc":false},{"t":"8","n":"Armand Satria Erlangga","s":"Syahrul","k":"Armand Satria Erlangga","sc":false},{"t":"8","n":"Atha Rizki Prasetia","s":"Syahrul","k":"Atha Rizki Prasetia","sc":false},{"t":"8","n":"Devano Ilham Baihaqi","s":"Daud","k":"Devano Ilham Baihaqi","sc":false},{"t":"8","n":"Hanif Datin Asqalani","s":"Syahrul","k":"Hanif Datin Asqalani","sc":false},{"t":"8","n":"Masagus Muhammad Taufiqurrahman","s":"Husnu","k":"Masagus Muhammad Taufiqurrahman","sc":false},{"t":"8","n":"Muhammad Abdurrahman Alfatih","s":"Azzam","k":"Muhammad Abdurrahman Alfatih","sc":false},{"t":"8","n":"Muhammad Adzka Alfatih","s":"Syahrul","k":"Muhammad Adzka Alfatih","sc":false},{"t":"8","n":"Muhamad Baariq Al Faqih Yusup","s":"Daud","k":"Muhamad Baariq Al Faqih Yusup","sc":false},{"t":"8","n":"Muhammad Fabio Reandra","s":"Daud","k":"Muhammad Fabio Reandra","sc":false},{"t":"8","n":"Muhammad Faiz","s":"Husnu","k":"Muhammad Faiz","sc":false},{"t":"8","n":"Muhammad Faiz Hidayat","s":"Syahrul","k":"Muhammad Faiz Hidayat","sc":false},{"t":"8","n":"Muhammad Fathin Ammar","s":"Mundzir","k":"Muhammad Fathin Ammar","sc":false},{"t":"8","n":"Prayata Syandana Sangadi Putra","s":"Adlan","k":"Prayata Syandana Sangadi Putra","sc":false},{"t":"8","n":"Rakha Fayyaz Gunawan","s":"Husnu","k":"Rakha Fayyaz Gunawan","sc":false},{"t":"8","n":"Raza Abyan Al Fatih","s":"Syahrul","k":"Raza Abyan Al Fatih","sc":false},{"t":"8","n":"Salman Narendra Yudhistira","s":"Daud","k":"Salman Narendra Yudhistira","sc":false},{"t":"8","n":"Umar Luthfi Taqiyuddin","s":"Husnu","k":"Umar Luthfi Taqiyuddin","sc":false},{"t":"8","n":"Zul Hilmi","s":"Daud","k":"Zul Hilmi","sc":false},{"t":"9A","n":"Casilas Ibnu Ammar","s":"Daud","k":"Casilas Ibnu Ammar","sc":false},{"t":"9A","n":"Fidai Syahwal Seftyan","s":"Ibrahim","k":"Fidai Syahwal Seftyan","sc":false},{"t":"9A","n":"Hadziq Farrais Fanzuri","s":"Ibrahim","k":"Hadziq Farrais Fanzuri","sc":false},{"t":"9A","n":"Ibnu Sa'ad Fauzi","s":"Azzam","k":"Ibnu Sa'ad Fauzi","sc":false},{"t":"9A","n":"Ibrahim Sahl","s":"Daud","k":"Ibrahim Sahl","sc":false},{"t":"9A","n":"Ilham Nur Rahman","s":"Munawar","k":"Ilham Nur Rahman","sc":false},{"t":"9A","n":"Iqbal Fadhlillah","s":"Syahrul","k":"Iqbal Fadhlillah","sc":false},{"t":"9A","n":"Khalid Abdurrahman","s":"Farhan","k":"Khalid Abdurrahman","sc":false},{"t":"9A","n":"Mohammad Damar Adjiepradipta Nandipinto","s":"Mundzir","k":"Mohammad Damar Adjiepradipta Nandipinto","sc":false},{"t":"9A","n":"Muhammad Albanrizky","s":"Ibrahim","k":"Muhammad Albanrizky","sc":false},{"t":"9A","n":"Muhammad Fadhlan Al Banna","s":"Ibrahim","k":"Muhammad Fadhlan Al Banna","sc":false},{"t":"9A","n":"Muhammad Gibral Arasta","s":"Azzam","k":"Muhammad Gibral Arasta","sc":false},{"t":"9A","n":"Muhammad Rayhan Albaihaqi","s":"Mundzir","k":"Muhammad Rayhan Albaihaqi","sc":false},{"t":"9A","n":"Muhammad Zaidan El Islamy","s":"Mundzir","k":"Muhammad Zaidan El Islamy","sc":false},{"t":"9A","n":"Raditya Warman Syah","s":"Syahrul","k":"Raditya Warman Syah","sc":false},{"t":"9A","n":"Raiyandi Haikal Budiman","s":"Nanda","k":"Raiyandi Haikal Budiman","sc":false},{"t":"9A","n":"Sabiel Kaikhaalish Herdiansyah","s":"Nanda","k":"Sabiel Kaikhaalish Herdiansyah","sc":false},{"t":"9B","n":"Abbiyu Musyafa Suparno","s":"Husnu","k":"Abbiyu Musyafa Suparno","sc":false},{"t":"9B","n":"Abdul Falah Mubarak","s":"Nanda","k":"Abdul Falah Mubarak","sc":false},{"t":"9B","n":"Achmad Fadhil Al Zam","s":"Azzam","k":"Achmad Fadhil Al Zam","sc":false},{"t":"9B","n":"Adzil Rufiantoro","s":"Ibrahim","k":"Adzil Rufiantoro","sc":false},{"t":"9B","n":"Ahmad Fauzan Hasibuan","s":"Ibrahim","k":"Ahmad Fauzan Hasibuan","sc":false},{"t":"9B","n":"Alaric Dimitri Alfarabi","s":"Syahrul","k":"Alaric Dimitri Alfarabi","sc":false},{"t":"9B","n":"Alfatih Nararya Yulando","s":"Daud","k":"Alfatih Nararya Yulando","sc":false},{"t":"9B","n":"Aqeela Tsaqif Utomo","s":"Mundzir","k":"Aqeela Tsaqif Utomo","sc":false},{"t":"9B","n":"Azka Alfarabi","s":"Azzam","k":"Azka Alfarabi","sc":false},{"t":"9B","n":"Fairuz","s":"Munawar","k":"Fairuz","sc":false},{"t":"9B","n":"Hanif Abbad","s":"Husnu","k":"Hanif Abbad","sc":false},{"t":"9B","n":"Ibrahim Danish Anggara","s":"Ibrahim","k":"Ibrahim Danish Anggara","sc":false},{"t":"9B","n":"Jaris Ghaly Mikail","s":"Ibrahim","k":"Jaris Ghaly Mikail","sc":false},{"t":"9B","n":"Muhammad Gibran Vito Fathoni","s":"Mundzir","k":"Muhammad Gibran Vito Fathoni","sc":false},{"t":"9B","n":"Muhammad Hafiz Nur Iman","s":"Farhan","k":"Muhammad Hafiz Nur Iman","sc":false},{"t":"9B","n":"Muhammad Rasyeed Dwitama","s":"Bagus","k":"Muhammad Rasyeed Dwitama","sc":false},{"t":"9B","n":"Raasyid Yusuf Al Faatih","s":"Mundzir","k":"Raasyid Yusuf Al Faatih","sc":false},{"t":"9B","n":"Zidan Pratama Ramadhan","s":"Husnu","k":"Zidan Pratama Ramadhan","sc":false},{"t":"10A","n":"Abdurrahman Afif","s":"Farhan","k":"Abdurrahman Afif","sc":false},{"t":"10A","n":"Azzam Al Ghifary","s":"Munawar","k":"Azzam Al Ghifary","sc":false},{"t":"10A","n":"Ghaisan Raqilla Manan","s":"Bagus","k":"Ghaisan Raqilla Manan","sc":false},{"t":"10A","n":"Haidar Mahya Sabilillah","s":"Nanda","k":"Haidar Mahya Sabilillah","sc":false},{"t":"10A","n":"Hilmi Romzi Nagib","s":"Bagus","k":"Hilmi Romzi Nagib","sc":false},{"t":"10A","n":"Hiraki Prayata Ananggadipa","s":"Bagus","k":"Hiraki Prayata Ananggadipa","sc":false},{"t":"10A","n":"Ibrahim Omar","s":"Azzam","k":"Ibrahim Omar","sc":false},{"t":"10A","n":"Muhammad Fakhril Fakih","s":"Mundzir","k":"Muhammad Fakhril Fakih","sc":false},{"t":"10A","n":"Muhammad Rhafif Assidik","s":"Munawar","k":"Muhammad Rhafif Assidik","sc":false},{"t":"10A","n":"Nabigh Habibie Ghany","s":"Nanda","k":"Nabigh Habibie Ghany","sc":false},{"t":"10A","n":"Reza Khoirul Azzam","s":"Farhan","k":"Reza Khoirul Azzam","sc":false},{"t":"10A","n":"Satria Rupawan","s":"Alkaf","k":"Satria Rupawan","sc":false},{"t":"10A","n":"Syaelendra Ebrahim Kamil","s":"Alkaf","k":"Syaelendra Ebrahim Kamil","sc":false},{"t":"10A","n":"Vittorio Veneto Setiyantara","s":"Alkaf","k":"Vittorio Veneto Setiyantara","sc":false},{"t":"10B","n":"Ahmad Alfathussholeh","s":"Alwan","k":"Ahmad Alfathussholeh","sc":false},{"t":"10B","n":"Akmal Luqman Syamlan","s":"Alwan","k":"Akmal Luqman Syamlan","sc":false},{"t":"10B","n":"Arqam Wadud","s":"Mundzir","k":"Arqam Wadud","sc":false},{"t":"10B","n":"Fajri Ahmad","s":"Azzam","k":"Fajri Ahmad","sc":false},{"t":"10B","n":"Fatih Akbar Dhiyaa Ul Haq","s":"Alwan","k":"Fatih Akbar Dhiyaa Ul Haq","sc":false},{"t":"10B","n":"Muhammad Bachtiar Almer Tajusa","s":"Azzam","k":"Muhammad Bachtiar Almer Tajusa","sc":false},{"t":"10B","n":"Muhammad bin Donald","s":"Mundzir","k":"Muhammad bin Donald","sc":false},{"t":"10B","n":"Muhammad Fadhli","s":"Alwan","k":"Muhammad Fadhli","sc":false},{"t":"10B","n":"Muhammad Faqih Alzawawi","s":"Alwan","k":"Muhammad Faqih Alzawawi","sc":false},{"t":"10B","n":"Muhammad Qaishar Rasendrya","s":"Alwan","k":"Muhammad Qaishar Rasendrya","sc":false},{"t":"10B","n":"Rafi Ardhani","s":"Alwan","k":"Rafi Ardhani","sc":false},{"t":"10B","n":"Yusuf Putra Rinadi","s":"Alwan","k":"Yusuf Putra Rinadi","sc":false},{"t":"11A","n":"Abdul Hafidz As Syauqi","s":"Ibrahim","k":"Abdul Hafidz As Syauqi","sc":false},{"t":"11A","n":"Abdurrahman","s":"Mundzir","k":"Abdurrahman","sc":false},{"t":"11A","n":"Albar Abdul Malik","s":"Alkaf","k":"Albar Abdul Malik","sc":false},{"t":"11A","n":"Amr Faiz","s":"Husnu","k":"Amr Faiz","sc":false},{"t":"11A","n":"Azzam Siddiq Mutahar","s":"Azzam","k":"Azzam Siddiq Mutahar","sc":false},{"t":"11A","n":"Baraka Ramadhan","s":"Munawar","k":"Baraka Ramadhan","sc":false},{"t":"11A","n":"Barra Adivian","s":"Bagus","k":"Barra Adivian","sc":false},{"t":"11A","n":"Faiz Ghazali Raindra Wisnumurti","s":"Munawar","k":"Faiz Ghazali Raindra Wisnumurti","sc":false},{"t":"11A","n":"Fawwaz Romzi Nagib","s":"Bagus","k":"Fawwaz Romzi Nagib","sc":false},{"t":"11A","n":"Imam Sandy Bachtiar","s":"Bagus","k":"Imam Sandy Bachtiar","sc":false},{"t":"11A","n":"Izzamnuddin Al Qassam","s":"Farhan","k":"Izzamnuddin Al Qassam","sc":false},{"t":"11A","n":"Luthfi Novriasyah","s":"Bagus","k":"Luthfi Novriasyah","sc":false},{"t":"11A","n":"Muhammad Hibban Syakir","s":"Azzam","k":"Muhammad Hibban Syakir","sc":false},{"t":"11A","n":"Muhammad Kahfi Achyarudin","s":"Nanda","k":"Muhammad Kahfi Achyarudin","sc":false},{"t":"11A","n":"Muhammad Naufal Jamil","s":"Azzam","k":"Muhammad Naufal Jamil","sc":false},{"t":"11A","n":"Rafif Ardinata","s":"Nanda","k":"Rafif Ardinata","sc":false},{"t":"11A","n":"Rakky Achmad Baihaqi","s":"Syahrul","k":"Rakky Achmad Baihaqi","sc":false},{"t":"11A","n":"Zhafran Atha Razin Hadiny","s":"Ibrahim","k":"Zhafran Atha Razin Hadiny","sc":false},{"t":"11B","n":"Abdullah Azzam","s":"Mundzir","k":"Abdullah Azzam","sc":true},{"t":"11B","n":"Arsyad Faqih Alhisyami","s":"Farhan","k":"Arsyad Faqih Alhisyami","sc":false},{"t":"11B","n":"Avrijal","s":"Farhan","k":"Avrijal","sc":false},{"t":"11B","n":"Azki Bagas Daneshwara","s":"Daud","k":"Azki Bagas Daneshwara","sc":false},{"t":"11B","n":"Fadhil Saadi Zahid Azkha Anatha","s":"Ibrahim","k":"Fadhil Saadi Zahid Azkha Anatha","sc":false},{"t":"11B","n":"Fathi Dzahabi S.","s":"Munawar","k":"Fathi Dzahabi S.","sc":false},{"t":"11B","n":"Fauzan Amali","s":"Nanda","k":"Fauzan Amali","sc":false},{"t":"11B","n":"Luthfan Ihtisyamuddin Efendi","s":"Azzam","k":"Luthfan Ihtisyamuddin Efendi","sc":false},{"t":"11B","n":"Luthfi Anhar","s":"Munawar","k":"Luthfi Anhar","sc":false},{"t":"11B","n":"Luthfian Azzam Fadhnazzir","s":"Farhan","k":"Luthfian Azzam Fadhnazzir","sc":false},{"t":"11B","n":"Muhammad Garuda Pratama","s":"Munawar","k":"Muhammad Garuda Pratama","sc":false},{"t":"11B","n":"Muhammad Mustafid Ilmi","s":"Farhan","k":"Muhammad Mustafid Ilmi","sc":false},{"t":"11B","n":"Muhammad Rafka Fathircha","s":"Daud","k":"Muhammad Rafka Fathircha","sc":false},{"t":"11B","n":"Muhammad Rifky Himawan Widitama","s":"Alkaf","k":"Muhammad Rifky Himawan Widitama","sc":false},{"t":"11B","n":"Raffa Hitipeuw","s":"Bagus","k":"Raffa Hitipeuw","sc":false},{"t":"11B","n":"Salafy Abdullah Yusuf","s":"Alkaf","k":"Salafy Abdullah Yusuf","sc":false},{"t":"11B","n":"Syateer Syafiq Sungkar","s":"Alwan","k":"Syateer Syafiq Sungkar","sc":false},{"t":"11B","n":"Tsany Al Fachrizy","s":"Nanda","k":"Tsany Al Fachrizy","sc":false},{"t":"12","n":"Abduh Hatim","s":"Bagus","k":"Abduh Hatim","sc":false},{"t":"12","n":"Abyan Eshan","s":"Farhan","k":"Abyan Eshan","sc":false},{"t":"12","n":"Azka Farid Arsyad","s":"Husnu","k":"Azka Farid Arsyad","sc":false},{"t":"12","n":"Fadgham Khairul Hafizh","s":"Munawar","k":"Fadgham Khairul Hafizh","sc":false},{"t":"12","n":"Galang Ramadhan","s":"Munawar","k":"Galang Ramadhan","sc":false},{"t":"12","n":"Hammam","s":"Alkaf","k":"Hammam","sc":false},{"t":"12","n":"Hendrico Altaf Husein Dasa","s":"Farhan","k":"Hendrico Altaf Husein Dasa","sc":false},{"t":"12","n":"Ibrahim Bin Donald Arthur Muhammad","s":"Alkaf","k":"Ibrahim Bin Donald Arthur Muhammad","sc":false},{"t":"12","n":"Jusuf Fathan Nuradly","s":"Alkaf","k":"Jusuf Fathan Nuradly","sc":false},{"t":"12","n":"Muhammad Akram Almair","s":"Alkaf","k":"Muhammad Akram Almair","sc":false},{"t":"12","n":"Muhammad Hafidz Aditya Zaini","s":"Munawar","k":"Muhammad Hafidz Aditya Zaini","sc":false},{"t":"12","n":"Muhammad Ilyas Abudussalam","s":"Bagus","k":"Muhammad Ilyas Abudussalam","sc":false},{"t":"12","n":"Muhammad Taqiyuddin","s":"Nanda","k":"Muhammad Taqiyuddin","sc":false},{"t":"12","n":"Nayaka Danendra Al Thafah","s":"Nanda","k":"Nayaka Danendra Al Thafah","sc":false},{"t":"12","n":"Perdana Muhammad Wildanumukhaladun","s":"Alkaf","k":"Perdana Muhammad Wildanumukhaladun","sc":false},{"t":"12","n":"Rafi Uddin Hannan","s":"Bagus","k":"Rafi Uddin Hannan","sc":false},{"t":"12","n":"Rayyan Ghibran Ananta","s":"Husnu","k":"Rayyan Ghibran Ananta","sc":false},{"t":"12","n":"Turky Husein Hatim","s":"Nanda","k":"Turky Husein Hatim","sc":false},{"t":"12","n":"Yazid Habibi Tambunan","s":"Alkaf","k":"Yazid Habibi Tambunan","sc":false}];


function perbaikiSemua() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  perbaikiDataHalaqah_(ss); // WAJIB pertama: samakan ejaan halaqah sebelum VLOOKUP dibangun
  hapusKelasLama_(ss);
  hapusBarisLabelHalaqah_(ss);
  var md = buildMasterData_(ss);
  SEP = deteksiSep_(md);
  var loc = '?';
  try { loc = ss.getSpreadsheetLocale(); } catch (e) {}
  Logger.log('Locale spreadsheet: ' + loc + ' | pemisah argumen formula terpakai: "' + SEP + '"');
  buildSemuaKelas_(ss);
  buildRekap_(ss, 'Rekap Absensi SMP', ['7A', '7B', '8', '9A', '9B']);
  buildRekap_(ss, 'Rekap Absensi SMA', ['10A', '10B', '11A', '11B', '12']);
  buildPersentaseTotal_(ss);
  buildRekapAbsensi_(ss);
  formatHalaqah_(ss);
  SpreadsheetApp.flush();
  verifikasi_(ss, loc);
  try { ss.toast('Selesai! Sheet kelas, rekap absensi, dan persentase sudah dibangun ulang.', 'Mutabaah', 10); } catch (e) {}
}

/* Deteksi pemisah argumen yang diterima locale spreadsheet ini.
   Mencoba =SUM(1,2) lalu =SUM(1;2) di sel uji; yang hasilnya 3 dipakai. */
function deteksiSep_(sh) {
  var cands = [',', ';'];
  var cell = sh.getRange(2000, 30); // AD2000 - jauh dari area data
  for (var i = 0; i < cands.length; i++) {
    try { cell.setFormula('=SUM(1' + cands[i] + '2)'); } catch (e) { continue; }
    SpreadsheetApp.flush();
    var v = null;
    try { v = cell.getValue(); } catch (e) { v = null; }
    cell.clearContent();
    if (v === 3) return cands[i];
  }
  return ',';
}

function verifikasi_(ss, loc) {
  var sh = ss.getSheetByName(TABS[0]);
  if (!sh) return;
  var f = '', v = null;
  try { f = sh.getRange('C3').getFormula(); } catch (e) {}
  try { v = sh.getRange('C3').getValue(); } catch (e) { v = '(exception saat baca nilai)'; }
  Logger.log('Verifikasi ' + TABS[0] + '!C3 formula: ' + f);
  Logger.log('Verifikasi ' + TABS[0] + '!C3 nilai   : ' + v);
  if (String(v) === '#ERROR!') {
    Logger.log('PERINGATAN: masih #ERROR! walau separator "' + SEP + '" (locale ' + loc + '). Salin seluruh isi log ini dan laporkan.');
  }
  Logger.log('TOTAL SANTRI di SANTRI[]: ' + SANTRI.length + ' (harus 166)');
  var pt = ss.getSheetByName('Persentase Total');
  if (pt) {
    try {
      SpreadsheetApp.flush();
      Logger.log('Verifikasi Persentase Total -> Tercapai=' + pt.getRange('C7').getValue()
        + ' | TidakTercapai=' + pt.getRange('D7').getValue()
        + ' | TanpaData=' + pt.getRange('E7').getValue()
        + ' | JumlahSantri=' + pt.getRange('F7').getValue()
        + ' | %Tercapai=' + pt.getRange('G7').getValue());
      Logger.log('(harapannya 18 / 99 / 49 / 166 / sekitar 10,8%)');
    } catch (e) { Logger.log('Verifikasi Persentase Total gagal dibaca: ' + e); }
  }
}


/* ---------- Ganti Ustadz (sekali jalan) ----------
   Ustadz Dani keluar, digantikan Daud (seluruh santri ikut).
   Menyamakan nama tab halaqah "Dani" -> "Daud" di file Laporan ini
   DAN di spreadsheet Mutaba'ah Juli.
   Aman dijalankan berulang: jika tab "Dani" sudah tidak ada, dilewati. */
var FILE_MUTABAAH = '1c0WwiYJR9S6ykNBnewl_CdrFhdRIIL8NBdu2Y9cUdC8';

function gantiUstadzDaniDaud() {
  var hasil = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  hasil.push(renameDani_(ss, 'Laporan'));
  try {
    hasil.push(renameDani_(SpreadsheetApp.openById(FILE_MUTABAAH), "Mutaba'ah"));
  } catch (e) {
    hasil.push("Mutaba'ah: GAGAL (" + e + ") - jalankan ulang dan izinkan akses");
  }
  Logger.log(hasil.join('\n'));
  try { ss.toast(hasil.join(' | '), 'Ganti Ustadz: Dani -> Daud', 15); } catch (e) {}
}

function renameDani_(ss, label) {
  var sh = ss.getSheetByName('Dani');
  if (!sh) {
    if (ss.getSheetByName('Daud')) return label + ': tab sudah "Daud" (dilewati)';
    return label + ': tab "Dani" TIDAK ditemukan!';
  }
  if (ss.getSheetByName('Daud')) return label + ': tab "Daud" sudah ada - "Dani" tidak diubah (cek manual)';
  sh.setName('Daud');
  return label + ': tab "Dani" -> "Daud" sukses';
}
/* ---------- util ---------- */

function getOrCreate_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function resetSheet_(sh) {
  sh.clear();
  try { sh.clearFormats(); } catch (e) {}
  try { sh.setConditionalFormatRules([]); } catch (e) {}
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
}

/* ---------- Styling (tema biru profesional, Times New Roman) ---------- */

var FONT = 'Times New Roman';
var C_TITLE_BG = '#1F4E79', C_TITLE_FG = '#FFFFFF';
var C_HDR_BG = '#CFE2F3', C_HDR_FG = '#1F4E79';
var C_BAND = '#F3F7FB';
var C_TOTAL_BG = '#FFF2CC';
var C_BORDER = '#B7B7B7';
var C_SIDE_TTL_BG = '#38761D', C_SIDE_HDR_BG = '#D9EAD3';
var C_OK_BG = '#D9EAD3', C_MID_BG = '#FFF2CC', C_BAD_BG = '#F4CCCC';

function judul_(rg, size) {
  rg.merge().setBackground(C_TITLE_BG).setFontColor(C_TITLE_FG).setFontSize(size || 14)
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function header_(rg, size) {
  rg.setBackground(C_HDR_BG).setFontColor(C_HDR_FG).setFontSize(size || 11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function border_(rg) {
  try { rg.setBorder(true, true, true, true, true, true, C_BORDER, SpreadsheetApp.BorderStyle.SOLID); } catch (e) {}
}

function zebra_(sh, row0, n, cols) {
  for (var i = 1; i < n; i += 2) sh.getRange(row0 + i, 1, 1, cols).setBackground(C_BAND);
}

function formatKelas_(sh, n) {
  sh.getRange(1, 1, Math.min(sh.getMaxRows(), 100), 19).setFontFamily(FONT);
  // Judul dipecah agar tidak melintasi batas freeze kolom A:B
  var nm = sh.getName();
  sh.getRange(1, 1, 1, 2).clearContent().merge();
  sh.getRange(1, 1).setValue('Kelas ' + nm);
  judul_(sh.getRange(1, 1, 1, 2), 12);
  sh.getRange(1, 3).setValue('Presentase Capaian Hafalan Santri Kelas ' + nm + ' Bulan : ' + BULAN);
  judul_(sh.getRange(1, 3, 1, 9));
  sh.getRange(1, 13, 1, 7).merge().setBackground(C_SIDE_TTL_BG).setFontColor(C_TITLE_FG)
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  header_(sh.getRange(2, 1, 1, 11));
  sh.getRange(2, 13, 1, 7).setBackground(C_SIDE_HDR_BG).setFontColor(C_HDR_FG).setFontSize(11)
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 44);
  if (n > 0) {
    var body = sh.getRange(3, 1, n, 11);
    body.setFontSize(10).setVerticalAlignment('middle');
    zebra_(sh, 3, n, 11);
    sh.getRange(3, 1, n, 1).setHorizontalAlignment('center');
    sh.getRange(3, 2, n, 1).setHorizontalAlignment('left');
    sh.getRange(3, 3, n, 9).setHorizontalAlignment('center');
    border_(sh.getRange(2, 1, n + 1, 11));
    var side = sh.getRange(3, 13, 4, 7);
    side.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(3, 13, 4, 1).setFontWeight('bold');
    zebraSide_(sh);
    sh.getRange(7, 13, 1, 7).setBackground(C_TOTAL_BG).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);
    border_(sh.getRange(2, 13, 5, 7));
    var rngDf = 'D3:G' + (n + 2); // D..G: Target, Tercapai, Presentase, Keterangan ikut berwarna
    // Warna mengikuti logika Keterangan persis: cukup ISNUMBER(F) + ambang 80%.
    // Tanpa syarat "data sah" (D/E>0) agar baris bernilai 0 pun ikut diwarnai.
    var rules = [
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($F3)' + SEP + '$F3*100>=80)').setBackground(C_OK_BG).setRanges([sh.getRange(rngDf)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($F3)' + SEP + '$F3*100<80)').setBackground(C_BAD_BG).setRanges([sh.getRange(rngDf)]).build()
    ];
    sh.setConditionalFormatRules(rules);
  }
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 230);
  sh.setColumnWidth(3, 70);
  for (var c = 4; c <= 11; c++) sh.setColumnWidth(c, 95);
  sh.setColumnWidth(12, 25);
  for (var c2 = 13; c2 <= 19; c2++) sh.setColumnWidth(c2, 95);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
}

function zebraSide_(sh) {
  sh.getRange(4, 13, 1, 7).setBackground(C_BAND);
  sh.getRange(6, 13, 1, 7).setBackground(C_BAND);
}

/* Perapian KOSMETIK 15 sheet halaqah ustadz.
   Tampilan: font, warna, border, lebar kolom, format angka, perataan.
   Satu-satunya isi yang ditulis skrip ini: kolom Keterangan (G) diisi
   kategori otomatis "Tercapai"/"Tidak Tercapai" dari Presentase (F)
   + warna hijau/merah, sama seperti sheet kelas. Kolom lain tidak
   disentuh agar data sinkron dan VLOOKUP MasterData tetap valid. */
function formatHalaqah_(ss) {
  HALAQAH.forEach(function (nm) {
    var sh = ss.getSheetByName(nm);
    if (!sh) return;
    try {
      var last = Math.max(sh.getLastRow(), 3);
      if (last > 200) last = 200;
      var LC = 13; // A..M (M = Catatan Perkembangan)
      sh.getRange(1, 1, last, LC).setFontFamily(FONT);
      judul_(sh.getRange('A1:M1'), 12);
      sh.setRowHeight(1, 30);
      // Header kolom M (Catatan Perkembangan): nilai + merge M2:M3 lalu gaya seragam
      sh.getRange('M2').setValue('Catatan Perkembangan');
      try { sh.getRange('M2:M3').merge(); } catch (e) {}
      header_(sh.getRange('A2:M3'), 10);
      sh.setRowHeight(2, 22);
      sh.setRowHeight(3, 22);
      if (last >= 4) {
        sh.getRange(4, 1, last - 3, LC).setBackground(null); // hapus sisa warna zebra lama
        zebra_(sh, 4, last - 3, LC);
        sh.getRange(4, 1, last - 3, LC).setFontSize(10).setVerticalAlignment('middle');
        sh.getRange(4, 1, last - 3, 1).setHorizontalAlignment('center');
        sh.getRange(4, 2, last - 3, 1).setHorizontalAlignment('left');
        sh.getRange(4, 3, last - 3, 9).setHorizontalAlignment('center');
        sh.getRange(4, 6, last - 3, 1).setNumberFormat('0.0%');
        sh.getRange(4, 12, last - 3, 2).setHorizontalAlignment('center').setWrap(true);
        border_(sh.getRange(2, 1, last - 1, LC));
        // Hapus total sisa template lama DI BAWAH baris terakhir agar tabel
        // berakhir tepat di nama santri terakhir (tidak ada tabel nyangkut).
        try {
          var maxR = sh.getMaxRows();
          if (maxR > last) sh.getRange(last + 1, 1, maxR - last, LC).breakApart().clearFormat().clearContent();
        } catch (e) {}
        // Kolom Keterangan (G): kategori otomatis dari Presentase (F),
        // sama seperti sheet kelas (>=80% -> "Tercapai", <80% -> "Tidak Tercapai").
        var ketRows = [];
        for (var kr = 4; kr <= last; kr++) {
          ketRows.push(['=IF(ISNUMBER($F' + kr + ')' + SEP + 'IF($F' + kr + '*100>=80' + SEP + '"Tercapai"' + SEP + '"Tidak Tercapai")' + SEP + '"")']);
        }
        sh.getRange(4, 7, last - 3, 1).setFormulas(ketRows);
        // Warna hijau/merah pada kolom Keterangan berbasis F.
        // Aturan ket lama buatan skrip ini dibuang dulu agar tidak menumpuk saat run ulang.
        var gRng = sh.getRange(4, 7, last - 3, 1);
        var oldRules = [];
        try { oldRules = sh.getConditionalFormatRules(); } catch (e2) {}
        var keptRules = oldRules.filter(function (r0) {
          var f0 = '';
          try { f0 = r0.getWhenFormulaSatisfied() || ''; } catch (e3) {}
          return f0.indexOf('$F4*100') === -1;
        });
        keptRules.push(
          SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($F4)' + SEP + '$F4*100>=80)').setBackground(C_OK_BG).setRanges([gRng]).build(),
          SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($F4)' + SEP + '$F4*100<80)').setBackground(C_BAD_BG).setRanges([gRng]).build()
        );
        sh.setConditionalFormatRules(keptRules);
      }
      sh.setColumnWidth(1, 42);
      sh.setColumnWidth(2, 180);
      sh.setColumnWidth(3, 62);
      sh.setColumnWidth(4, 82);
      sh.setColumnWidth(5, 82);
      sh.setColumnWidth(6, 105);
      sh.setColumnWidth(7, 115);
      for (var c = 8; c <= 11; c++) sh.setColumnWidth(c, 58);
      sh.setColumnWidth(12, 130);
      sh.setColumnWidth(13, 300);
      try {
        var maxC = sh.getMaxColumns();
        if (maxC > LC) sh.getRange(1, LC + 1, Math.min(sh.getMaxRows(), last), maxC - LC).clearContent().clearFormat();
      } catch (e) {}
      sh.setFrozenRows(3);
    } catch (e) { Logger.log('Format halaqah ' + nm + ': ' + e); }
  });
  Logger.log('Perapian sheet halaqah selesai.');
}

function formatRekapBlock_(sh, startCol, last, tr) {
  var w = [38, 190, 60, 65, 65, 65, 65, 70];
  for (var i = 0; i < 8; i++) sh.setColumnWidth(startCol + i, w[i]);
  sh.setColumnWidth(startCol + 8, 20);
  sh.setRowHeight(2, 40);
  judul_(sh.getRange(1, startCol, 1, 8), 11);
  header_(sh.getRange(2, startCol, 1, 8), 10);
  var n = last - 2;
  var body = sh.getRange(3, startCol, n, 8);
  body.setFontSize(10).setVerticalAlignment('middle');
  for (var j = 1; j < n; j += 2) sh.getRange(3 + j, startCol, 1, 8).setBackground(C_BAND);
  sh.getRange(3, startCol, n, 1).setHorizontalAlignment('center');
  sh.getRange(3, startCol + 1, n, 1).setHorizontalAlignment('left');
  sh.getRange(3, startCol + 2, n, 6).setHorizontalAlignment('center');
  sh.getRange(tr, startCol, 1, 8).setBackground(C_TOTAL_BG).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);
  border_(sh.getRange(2, startCol, tr - 1, 8));
}

function formatPersentase_(sh) {
  sh.getRange(1, 1, 7, 7).setFontFamily(FONT);
  judul_(sh.getRange(1, 1, 1, 7));
  header_(sh.getRange(2, 1, 1, 7));
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 44);
  var body = sh.getRange(3, 1, 5, 7);
  body.setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('center');
  zebra_(sh, 3, 4, 7);
  sh.getRange(7, 1, 1, 7).setBackground(C_TOTAL_BG).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);
  border_(sh.getRange(2, 1, 6, 7));
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 70);
  for (var c = 3; c <= 6; c++) sh.setColumnWidth(c, 140);
  sh.setColumnWidth(7, 110);
  var colG = sh.getRange('G3:G7');
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.75).setBackground(C_OK_BG).setRanges([colG]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.5).setBackground(C_MID_BG).setRanges([colG]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5).setBackground(C_BAD_BG).setRanges([colG]).build()
  ];
  sh.setConditionalFormatRules(rules);
}

function colLetter_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function hapusKelasLama_(ss) {
  KELAS_LAMA.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh) { try { ss.deleteSheet(sh); } catch (e) {} }
  });
}

/* Hapus baris sampah berlabel ("S A K I T", "I Z I N", "A L P A", "T D K SETOR")
   yang menempel di akhir tiap sheet halaqah, agar tabel berakhir tepat di
   nama santri terakhir. Pencocokan dinormalisasi: huruf besar + buang semua
   karakter non-huruf, jadi "S A K I T" -> "SAKIT". Baris dihapus dari bawah
   ke atas supaya indeks tidak bergeser. Data di atasnya tidak terpengaruh. */
function hapusBarisLabelHalaqah_(ss) {
  var LABEL = { 'IZIN': 1, 'SAKIT': 1, 'ALPA': 1, 'TDKSETOR': 1 };
  var total = 0;
  HALAQAH.forEach(function (nm) {
    var sh = ss.getSheetByName(nm);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 4) return;
    var vals = sh.getRange(4, 2, last - 3, 1).getValues(); // kolom B mulai baris 4
    var rows = [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0]).toUpperCase().replace(/[^A-Z]/g, '');
      if (v && LABEL[v]) rows.push(4 + i);
    }
    for (var j = rows.length - 1; j >= 0; j--) {
      try { sh.deleteRows(rows[j], 1); total++; } catch (e) { Logger.log(nm + ': gagal hapus baris ' + rows[j] + ' - ' + e); }
    }
    if (rows.length) Logger.log(nm + ': dihapus ' + rows.length + ' baris label.');
  });
  Logger.log('Total baris label sampah dihapus: ' + total);
}

/* ---------- Perbaikan data halaqah & kelas (hasil audit silang 170 vs 166) ----------
   Sheet KELAS = acuan ejaan. Tahap ini:
   1) me-rename 9 nama di sheet halaqah agar identik dengan kunci di SANTRI
   2) menghapus 4 baris halaqah yang tak dikenali kelas / duplikat antar ustadz
   3) membakukan 3 ejaan di sheet kelas (Afif, Bachtiar, Abdullah)
   4) menomori ulang kolom No pada sheet halaqah yang barisnya berkurang
   Idempoten: jika nama sudah baku / baris sudah terhapus -> dilewati (SKIP). */
var RENAME_HALAQAH = [
  ['Alkaf',   'Ibrohim bin Donald New',              'Ibrahim Bin Donald Arthur Muhammad'],
  ['Munawar', 'Fathi Dzahabi.S',                     'Fathi Dzahabi S.'],
  ['Farhan',  'Abdurrohman Afif',                    'Abdurrahman Afif'],
  ['Azzam',   'Muhamad Bahtiar Almer Tajusa',        'Muhammad Bachtiar Almer Tajusa'],
  ['Azzam',   'Achmad Fadhil Al zam',                'Achmad Fadhil Al Zam'],
  ['Mundzir', 'Arqam Wadud Affandi',                 'Arqam Wadud'],
  ['Mundzir', 'Muhammad Bin Donald Arthur Muhammad', 'Muhammad bin Donald'],
  ['Adlan',   'Anfhal Zhafir Putra Alfi',            'Anfaal Zhafirputra Alfi'],
  ['Daud',    'Muhammad Baariq Alfaqih Yusup',       'Muhamad Baariq Al Faqih Yusup']
];
var HAPUS_HALAQAH = [
  ['Ibrahim', 'Cholid'],
  ['Farhan',  'Abdillah Arrafif'],
  ['Adlan',   'Genta Lilo Abimanyu'],
  ['Alwan',   'Ganendra Farzan Pratama']
];
var BAKU_KELAS = [
  ['10A', 'Abdurrahman afif',             'Abdurrahman Afif'],
  ['10B', 'Muhamad Bchtiar Almer Tajusa', 'Muhammad Bachtiar Almer Tajusa'],
  ['11B', 'Salafy Abdulah Yusuf',         'Salafy Abdullah Yusuf']
];

function perbaikiDataHalaqah_(ss) {
  var log = [];
  RENAME_HALAQAH.forEach(function (it) {
    var r = cariBarisNama_(ss, it[0], it[1]);
    if (!r) {
      log.push(it[0] + ': "' + it[1] + '" -> ' + (cariBarisNama_(ss, it[0], it[2]) ? 'SKIP (sudah baku)' : 'TIDAK KETEMU!'));
      return;
    }
    ss.getSheetByName(it[0]).getRange(r, 2).setValue(it[2]);
    log.push(it[0] + ': "' + it[1] + '" -> "' + it[2] + '" OK');
  });
  HAPUS_HALAQAH.forEach(function (it) {
    var sh = ss.getSheetByName(it[0]);
    var r = cariBarisNama_(ss, it[0], it[1]);
    if (!r || !sh) { log.push(it[0] + ': hapus "' + it[1] + '" SKIP'); return; }
    sh.deleteRow(r);
    rapikanNo_(sh);
    log.push(it[0] + ': hapus "' + it[1] + '" OK');
  });
  BAKU_KELAS.forEach(function (it) {
    var sh = ss.getSheetByName(it[0]);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 3) return;
    var vals = sh.getRange(3, 2, last - 2, 1).getValues(); // kolom B mulai baris 3
    for (var i = 0; i < vals.length; i++) {
      var nm = String(vals[i][0]).trim();
      if (nm === it[1]) { sh.getRange(3 + i, 2).setValue(it[2]); log.push(it[0] + ': "' + it[1] + '" -> "' + it[2] + '" OK (kelas)'); return; }
      if (nm === it[2]) { log.push(it[0] + ': "' + it[2] + '" SKIP (sudah baku)'); return; }
    }
    log.push(it[0] + ': "' + it[1] + '" TIDAK KETEMU di kelas!');
  });
  Logger.log('PERBAIKAN DATA HALAQAH/KELAS:\n' + log.join('\n'));
}

function cariBarisNama_(ss, sheetName, nama) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 4) return null;
  var vals = sh.getRange(4, 2, last - 3, 1).getValues(); // data nama mulai baris 4
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === nama) return 4 + i;
  }
  return null;
}

function rapikanNo_(sh) {
  var last = sh.getLastRow();
  if (last < 4) return;
  var vals = sh.getRange(4, 2, last - 3, 1).getValues();
  var no = 0;
  for (var i = 0; i < vals.length; i++) {
    if (!String(vals[i][0]).trim()) continue;
    no++;
    sh.getRange(4 + i, 1).setValue(no);
  }
}

/* ---------- MasterData ---------- */

function buildMasterData_(ss) {
  var sh = getOrCreate_(ss, 'MasterData');
  resetSheet_(sh);
  // [PENTING] Rumus array menumpuk 15 sheet x 99 baris = 1485 baris hasil.
  // Grid sheet WAJIB cukup panjang; kalau lebih pendek, rumus menjadi
  // #VALUE! dan seluruh nilai di sheet kelas ikut kosong (pernah terjadi).
  var needRows = HALAQAH.length * 99 + 2;
  if (sh.getMaxRows() < needRows) {
    sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
  }
  // Kolom hasil: Nama, Grade, TargetHal, JumlahHal, Presentase, Ket, Izin, Sakit, Alpa, TidakSetor
  var f = '={' + HALAQAH.map(function (h) { return "'" + h + "'!B2:K100"; }).join(';') + '}';
  sh.getRange(1, 1).setFormula(f);
  SpreadsheetApp.flush();
  var cekA1 = '';
  try { cekA1 = String(sh.getRange('A1').getValue()); } catch (e) { cekA1 = '(exception)'; }
  if (cekA1.charAt(0) === '#') Logger.log('PERINGATAN: MasterData!A1 = ' + cekA1 + ' | grid ' + sh.getMaxRows() + ' baris. Jalankan diagnosaMasterData.');
  try { sh.hideSheet(); } catch (e) {}
  return sh;
}

/* ---------- Diagnosa MasterData (sekali jalan, RINGAN) ----------
   Tahap 1 (instan): laporkan tinggi grid MasterData. Kalau lebih pendek dari
   kebutuhan spill (15 x 99 + 2 = 1487 baris), akar masalah #VALUE! pasti
   grid yang pendek -> tidak perlu tes rumus apa pun; cukup jalankan
   perbaikiSemua yang kini otomatis memperluas grid.
   Tahap 2 (hanya bila grid cukup): uji rumus PER-SHEET (spill 99 baris)
   di blok baris terpisah tanpa membersihkan seluruh sheet, dengan anggaran
   waktu 3,5 menit agar tidak mungkin kena batas eksekusi 6 menit. */
function diagnosaMasterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var needRows = HALAQAH.length * 99 + 2;
  var hasil = [];
  var md = ss.getSheetByName('MasterData');
  var maxMd = md ? md.getMaxRows() : 0;
  hasil.push('Tahap 1 - Grid MasterData:');
  hasil.push('- maxRows = ' + maxMd + ' | butuh minimal ' + needRows);
  if (!md) { hasil.push('- MasterData TIDAK ADA! Jalankan perbaikiSemua.'); tampilkanHasil_(ss, hasil); return; }
  if (maxMd < needRows) {
    hasil.push('- KESIMPULAN: grid TERLALU PENDEK -> ini penyebab #VALUE!.');
    hasil.push('- Obat: jalankan perbaikiSemua (sudah otomatis memperluas grid). Tes rumus dilewati.');
    tampilkanHasil_(ss, hasil);
    return;
  }
  hasil.push('- Grid CUKUP. Lanjut tahap 2: uji per sheet...');
  // ---- Tahap 2: uji per-sheet, ringan & berbatas waktu ----
  var namaDiag = 'DiagMaster';
  var old = ss.getSheetByName(namaDiag);
  if (old) ss.deleteSheet(old); // sisa run sebelumnya yang timeout
  var sh = ss.insertSheet(namaDiag);
  var blokBaris = 110; // spill maksimum satu sheet = 99 baris + jarak
  if (sh.getMaxRows() < HALAQAH.length * blokBaris + 10) {
    sh.insertRowsAfter(sh.getMaxRows(), HALAQAH.length * blokBaris + 10 - sh.getMaxRows());
  }
  var mulai = Date.now();
  var BATAS_MS = 210000; // 3,5 menit
  var adaGagal = false;
  for (var i = 0; i < HALAQAH.length; i++) {
    if (Date.now() - mulai > BATAS_MS) {
      hasil.push('... dihentikan tepat waktu (' + (HALAQAH.length - i) + ' sheet belum dites).');
      break;
    }
    var r0 = i * blokBaris + 1;
    try {
      sh.getRange(r0, 1).setFormula("={'" + HALAQAH[i] + "'!B2:K100}");
      SpreadsheetApp.flush();
      var v = '';
      try { v = String(sh.getRange(r0, 1).getValue()); } catch (e2) {}
      var gagal = v.charAt(0) === '#';
      if (gagal) { adaGagal = true; hasil.push(HALAQAH[i] + ' -> GAGAL (' + v + ')'); }
      else { hasil.push(HALAQAH[i] + ' -> OK'); }
    } catch (e) {
      adaGagal = true;
      hasil.push(HALAQAH[i] + ' -> GAGAL (' + e + ')');
    }
  }
  ss.deleteSheet(sh);
  if (!adaGagal && hasil[hasil.length - 1].indexOf('belum dites') === -1) {
    hasil.push('KESIMPULAN: semua sheet sehat sendiri-sendiri. Jalankan perbaikiSemua lalu cek MasterData lagi.');
  }
  tampilkanHasil_(ss, hasil);
}

function tampilkanHasil_(ss, hasil) {
  Logger.log(hasil.join('\n'));
  var pesan = hasil.join('\n');
  try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { try { ss.toast(pesan.substring(0, 240), 'Diagnosa', 30); } catch (e2) {} }
}

/* ---------- Sheet Kelas ---------- */

function buildSemuaKelas_(ss) {
  TABS.forEach(function (tab) {
    var sh = getOrCreate_(ss, tab);
    resetSheet_(sh);
    var anggota = SANTRI.filter(function (x) { return x.t === tab; });

    // Baris 1: judul utama + judul blok samping (di-merge & diwarnai saat styling)
    sh.getRange(1, 1).setValue('Presentase Capaian Hafalan Santri Kelas ' + tab + ' Bulan : ' + BULAN).setFontWeight('bold');
    sh.getRange(1, 13).setValue('Rata-rata Target Tercapai Per Grade');

    // Baris 2: header tabel utama + header blok samping (7 kolom M..S)
    var hdr = ['No', 'Nama Santri', 'Grade', 'Target Bulanan (Hal)', 'Target Tercapai (Hal)', 'Presentase', 'Keterangan', 'Sakit', 'Izin', 'Alpa', 'Tidak Setor'];
    hdr.forEach(function (h, i) { sh.getRange(2, 1 + i).setValue(h).setFontWeight('bold'); });
    var sideH2 = ['Grade', 'Jumlah Santri', 'Jumlah Tercapai', 'Presentase', 'Jumlah Tidak Tercapai', 'Presentase', 'Tanpa Data'];
    sideH2.forEach(function (h, i) { sh.getRange(2, 13 + i).setValue(h).setFontWeight('bold'); });

    // Baris 3+: santri
    var n = anggota.length;
    if (n > 0) {
      var noVals = [], namaVals = [], fRows = [];
      anggota.forEach(function (s, i) {
        noVals.push([i + 1]);
        namaVals.push([s.n]);
        if (!s.s) { fRows.push(['', '', '', '', '', '', '', '', '']); return; }
        var ref = s.sc ? ("'" + s.s + "'!$B:$K") : 'MasterData!$A:$J';
        function v(col) { return '=IFERROR(VLOOKUP("' + s.k + '"' + SEP + ref + SEP + col + SEP + 'FALSE)' + SEP + '"")'; }
        // Urutan kolom C..K: Grade, Target, Tercapai, Presentase, Ket, Sakit, Izin, Alpa, TidakSetor
        // Keterangan = kategori otomatis dari Presentase: >=80% "Tercapai", <80% "Tidak Tercapai"
        var rN = 3 + i;
        var ketF = '=IF(ISNUMBER($F' + rN + ')' + SEP + 'IF($F' + rN + '*100>=80' + SEP + '"Tercapai"' + SEP + '"Tidak Tercapai")' + SEP + '"")';
        fRows.push([v(2), v(3), v(4), v(5), ketF, v(8), v(7), v(9), v(10)]);
      });
      sh.getRange(3, 1, n, 1).setValues(noVals);
      sh.getRange(3, 2, n, 1).setValues(namaVals);
      sh.getRange(3, 3, n, 9).setFormulas(fRows);
    }

    // Blok rekap per grade (baris 3-6) + rata-rata (baris 7)
    var grades = ['A', 'B', 'C', 'D'];
    for (var g = 0; g < 4; g++) {
      var r = 3 + g;
      sh.getRange(r, 13).setValue(grades[g]);
      // N = headcount seluruh anggota grade | O = tercapai (data sah & % ustadz >= 100%)
      // P/R = persentase dengan penyebut data sah (N-S) | S = tanpa data
      // sahF & headcount sengaja tanpa koma argumen -> aman utk semua locale
      var sahF = '(TRIM($C$3:$C$100&"")=M' + r + ')*ISNUMBER($F$3:$F$100)*((($D$3:$D$100>0)*ISNUMBER($D$3:$D$100))+(($E$3:$E$100>0)*ISNUMBER($E$3:$E$100))>0)';
      sh.getRange(r, 14).setFormula('=SUMPRODUCT(TRIM($C$3:$C$100&"")=M' + r + ')');
      sh.getRange(r, 15).setFormula('=SUMPRODUCT(' + sahF + '*($F$3:$F$100>=1))');
      sh.getRange(r, 16).setFormula('=IFERROR(O' + r + '/(N' + r + '-S' + r + ')' + SEP + '0)');
      sh.getRange(r, 17).setFormula('=N' + r + '-O' + r + '-S' + r);
      sh.getRange(r, 18).setFormula('=IFERROR(Q' + r + '/(N' + r + '-S' + r + ')' + SEP + '0)');
      sh.getRange(r, 19).setFormula('=N' + r + '-SUMPRODUCT(' + sahF + ')');
    }
    sh.getRange(7, 13).setValue('Rata-rata').setFontWeight('bold');
    sh.getRange(7, 14).setFormula('=SUM(N3:N6)');
    sh.getRange(7, 15).setFormula('=SUM(O3:O6)');
    sh.getRange(7, 16).setFormula('=IFERROR(O7/(N7-S7)' + SEP + '0)');
    sh.getRange(7, 17).setFormula('=SUM(Q3:Q6)');
    sh.getRange(7, 18).setFormula('=IFERROR(Q7/(N7-S7)' + SEP + '0)');
    sh.getRange(7, 19).setFormula('=SUM(S3:S6)');

    // Format
    sh.getRange('F3:F100').setNumberFormat('0.0%');
    sh.getRange('P3:P7').setNumberFormat('0.0%');
    sh.getRange('R3:R7').setNumberFormat('0.0%');
    formatKelas_(sh, n);
  });
}

/* ---------- Rekap Absensi ---------- */

function buildRekap_(ss, sheetName, tabs) {
  var sh = getOrCreate_(ss, sheetName);
  resetSheet_(sh);
  // Pastikan grid cukup lebar utk semua blok (8 kolom isi + 1 spacer per blok).
  // Penulisan nilai auto-meluaskan grid, tapi setColumnWidth tidak -> harus dijamin eksplisit.
  var needCols = tabs.length * 9;
  if (sh.getMaxColumns() < needCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }
  var header = ['No', 'Nama', 'Grade', 'Absensi Sakit', 'Izin', 'Alpa', 'Tidak Setor', 'Total'];

  tabs.forEach(function (tab, bi) {
    var startCol = 1 + bi * 9; // A, J, S, AB, AK
    var L = function (off) { return colLetter_(startCol + off); };

    sh.getRange(1, startCol).setValue(sheetName.replace('Rekap Absensi ', 'Rekap Absensi ') + ' — Kelas ' + tab).setFontWeight('bold');
    header.forEach(function (h, i) { sh.getRange(2, startCol + i).setValue(h).setFontWeight('bold'); });

    var anggota = SANTRI.filter(function (x) { return x.t === tab; });
    if (!anggota.length) return;

    var noVals = [], namaVals = [], fRows = [];
    anggota.forEach(function (s, i) {
      noVals.push([i + 1]);
      namaVals.push([s.n]);
      var q = "'" + tab + "'";
      var look = function (col) { return '=IFERROR(VLOOKUP($' + L(1) + (3 + i) + SEP + q + '!$B$3:$K$100' + SEP + col + SEP + 'FALSE)' + SEP + '"")'; };
      fRows.push([look(2), look(7), look(8), look(9), look(10), '']);
    });
    var last = 2 + anggota.length;
    sh.getRange(3, startCol, anggota.length, 1).setValues(noVals);
    sh.getRange(3, startCol + 1, anggota.length, 1).setValues(namaVals);
    // Kolom: Grade, Sakit, Izin, Alpa, TidakSetor lalu Total manual
    sh.getRange(3, startCol + 2, anggota.length, 5).setFormulas(fRows.map(function (r) { return [r[0], r[1], r[2], r[3], r[4]]; }));
    for (var i = 0; i < anggota.length; i++) {
      var r = 3 + i;
      sh.getRange(r, startCol + 7).setFormula('=IF(' + L(3) + r + '=""' + SEP + '""' + SEP + 'SUM(' + L(3) + r + ':' + L(6) + r + '))');
    }

    // Baris TOTAL
    var tr = last + 1;
    sh.getRange(tr, startCol).setValue('TOTAL').setFontWeight('bold');
    for (var c = 3; c <= 7; c++) {
      sh.getRange(tr, startCol + c).setFormula('=SUM(' + L(c) + '3:' + L(c) + last + ')');
    }
    formatRekapBlock_(sh, startCol, last, tr);
  });
  sh.setFrozenRows(2);
}

/* ---------- Persentase Total ---------- */

function buildPersentaseTotal_(ss) {
  var sh = getOrCreate_(ss, 'Persentase Total');
  resetSheet_(sh);

  sh.getRange(1, 1).setValue('Persentase Capaian Hafalan Seluruh Kelas — Bulan ' + BULAN).setFontWeight('bold');
  // 7 kolom: kolom "Tanpa Data" ikut dihitung agar Jumlah Santri = SELURUH santri (166),
  // bukan hanya yang bersisa (117) -> % Tercapai memakai penyebut semua santri.
  var hdr = ['No', 'Grade', 'Tercapai (' + BULAN + ')', 'Tidak Tercapai (' + BULAN + ')', 'Tanpa Data (' + BULAN + ')', 'Jumlah Santri', '% Tercapai'];
  hdr.forEach(function (h, i) { sh.getRange(2, 1 + i).setValue(h).setFontWeight('bold'); });

  var grades = ['A', 'B', 'C', 'D'];
  for (var g = 0; g < 4; g++) {
    var r = 3 + g; // baris di sheet kelas untuk grade yg sama
    sh.getRange(r, 1).setValue(g + 1);
    sh.getRange(r, 2).setValue(grades[g]);
    // Blok samping tiap sheet kelas: O=Tercapai | Q=Tidak Tercapai | S=Tanpa Data
    var sumO = TABS.map(function (t) { return "'" + t + "'!O" + r; }).join('+');
    var sumQ = TABS.map(function (t) { return "'" + t + "'!Q" + r; }).join('+');
    var sumS = TABS.map(function (t) { return "'" + t + "'!S" + r; }).join('+');
    sh.getRange(r, 3).setFormula('=' + sumO);
    sh.getRange(r, 4).setFormula('=' + sumQ);
    sh.getRange(r, 5).setFormula('=' + sumS);
    sh.getRange(r, 6).setFormula('=C' + r + '+D' + r + '+E' + r);
    sh.getRange(r, 7).setFormula('=IFERROR(C' + r + '/F' + r + SEP + '0)');
  }
  sh.getRange(7, 1).setValue('');
  sh.getRange(7, 2).setValue('TOTAL').setFontWeight('bold');
  ['C', 'D', 'E'].forEach(function (L) {
    sh.getRange(7, L.charCodeAt(0) - 64).setFormula('=SUM(' + L + '3:' + L + '6)');
  });
  sh.getRange('F7').setFormula('=C7+D7+E7');
  sh.getRange('G7').setFormula('=IFERROR(C7/F7' + SEP + '0)');
  sh.getRange('C3:F7').setNumberFormat('0');
  sh.getRange('G3:G7').setNumberFormat('0.0%');
  formatPersentase_(sh);
}

/* ---------- Rekap Absensi Keseluruhan ----------
   Daftar santri dengan total absen (Sakit+Izin+Alpa+Tidak Setor) >= AMBANG_ABSEN.
   Dibangun ulang otomatis tiap Run -> ikut terbuat saat spreadsheet diduplikasi utk bulan depan. */
var AMBANG_ABSEN = 3;

function buildRekapAbsensi_(ss) {
  var sh = getOrCreate_(ss, 'Rekap Absensi Keseluruhan');
  resetSheet_(sh);
  sh.getRange(1, 1).setValue('Rekap Absensi & Ketidakhadiran Santri - Bulan ' + BULAN).setFontWeight('bold');
  var hdr = ['No', 'Kelas', 'Nama Santri', 'Grade', 'Sakit', 'Izin', 'Alpa', 'Tidak Setor', 'Total', 'Keterangan'];
  hdr.forEach(function (h, i) { sh.getRange(2, 1 + i).setValue(h).setFontWeight('bold'); });
  var out = [];
  TABS.forEach(function (t) {
    var cs = ss.getSheetByName(t);
    if (!cs) return;
    var last = cs.getLastRow();
    if (last < 3) return;
    var vals = cs.getRange(3, 2, last - 2, 10).getValues(); // B..K
    vals.forEach(function (v) {
      var nama = String(v[0]).trim();
      if (!nama) return;
      var sk = Number(v[6]) || 0;
      var iz = Number(v[7]) || 0;
      var al = Number(v[8]) || 0;
      var ts = Number(v[9]) || 0;
      var tot = sk + iz + al + ts;
      if (tot < AMBANG_ABSEN) return;
      var ket = [];
      if (sk + iz + al > 0) ket.push('Tidak hadir ' + (sk + iz + al) + 'x (sakit ' + sk + ', izin ' + iz + ', alpa ' + al + ')');
      if (ts > 0) ket.push('Tidak setor ' + ts + 'x');
      out.push([t, nama, String(v[1]), sk, iz, al, ts, tot, ket.join('; ')]);
    });
  });
  if (out.length) sh.getRange(3, 1, out.length, 10).setValues(out.map(function (r, i) { return [i + 1].concat(r); }));
  formatRekapAbsensi_(sh, out.length);
  Logger.log('Rekap Absensi Keseluruhan: ' + out.length + ' santri (total absen >= ' + AMBANG_ABSEN + ').');
}

function formatRekapAbsensi_(sh, n) {
  var last = Math.max(n + 2, 3);
  sh.getRange(1, 1, last, 10).setFontFamily(FONT);
  judul_(sh.getRange(1, 1, 1, 10), 14);
  header_(sh.getRange(2, 1, 1, 10));
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 44);
  if (n > 0) {
    var body = sh.getRange(3, 1, n, 10);
    body.setFontSize(10).setVerticalAlignment('middle');
    zebra_(sh, 3, n, 10);
    sh.getRange(3, 1, n, 2).setHorizontalAlignment('center');
    sh.getRange(3, 3, n, 1).setHorizontalAlignment('left');
    sh.getRange(3, 4, n, 6).setHorizontalAlignment('center');
    sh.getRange(3, 10, n, 1).setHorizontalAlignment('left').setWrap(true);
    sh.getRange(3, 9, n, 1).setFontWeight('bold');
    border_(sh.getRange(2, 1, n + 1, 10));
    try {
      var rngT = sh.getRange(3, 9, n, 1);
      sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(5).setBackground(C_BAD_BG).setRanges([rngT]).build()]);
    } catch (e) {}
  }
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 60);
  sh.setColumnWidth(3, 230);
  sh.setColumnWidth(4, 62);
  for (var c = 5; c <= 9; c++) sh.setColumnWidth(c, 65);
  sh.setColumnWidth(10, 320);
  sh.setFrozenRows(2);
}

