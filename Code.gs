/**
 * PintarSD - Backend Lisensi & Bank Soal (Google Apps Script)
 * Cara pakai: Extensions > Apps Script pada Google Sheet, tempel file ini,
 * jalankan setupWorksheetDB() sekali, lalu Deploy > Web app
 * (Execute as: Me, Who has access: Anyone). Salin URL /exec ke index.html (API_URL).
 */
const SH_L = 'Lisensi';
const SH_Q = 'Bank_Soal';

function setupWorksheetDB() {
  const ss = SpreadsheetApp.getActive();

  let l = ss.getSheetByName(SH_L) || ss.insertSheet(SH_L);
  if (l.getLastRow() === 0) {
    l.appendRow(['Kode_Lisensi', 'Nama_Pembeli', 'Paket_Akses', 'Device_1', 'Device_2', 'Status', 'Tanggal_Aktivasi']);
    l.getRange(2, 1, 1000, 5).setNumberFormat('@');
    l.getRange(2, 1, 2, 6).setValues([
      ['PINTAR-DEMO-001', 'Pembeli Contoh A', 'PKT-MTK1', '', '', 'AKTIF'],
      ['PINTAR-DEMO-002', 'Pembeli Contoh A', 'PKT-IND1', '', '', 'AKTIF']
    ]);
    l.setFrozenRows(1);
  }

  let q = ss.getSheetByName(SH_Q) || ss.insertSheet(SH_Q);
  if (q.getLastRow() === 0) {
    q.appendRow(['ID_Paket', 'Kelas', 'Mapel', 'Judul_Paket', 'Tipe_Soal', 'Pertanyaan', 'Data_Soal_JSON', 'Kunci_Jawaban', 'Petunjuk']);
    q.getRange(2, 1, 1000, 9).setNumberFormat('@');
    q.getRange(2, 1, 7, 9).setValues([
      ['PKT-MTK1', '1', 'Matematika', 'Berhitung Dasar', 'multiple_choice', 'Berapa jumlah apel di gambar?', '{"gambar":"img/apel3.png","alt":"tiga buah apel","opsi":["2","3","4"]}', '3', 'Hitung apelnya satu per satu.'],
      ['PKT-MTK1', '1', 'Matematika', 'Berhitung Dasar', 'multiple_choice', 'Mana yang berbentuk lingkaran?', '{"opsi":[{"gambar":"img/lingkaran.png","alt":"lingkaran","id":"lingkaran"},{"gambar":"img/kotak.png","alt":"kotak","id":"kotak"},{"gambar":"img/segitiga.png","alt":"segitiga","id":"segitiga"}]}', 'lingkaran', 'Lingkaran bentuknya bulat dan tidak punya sudut.'],
      ['PKT-MTK1', '1', 'Matematika', 'Berhitung Dasar', 'multiple_choice', '2 + 3 = ...', '{"opsi":["4","5","6","7"]}', '5', 'Hitung pakai jari: 2, lalu tambah 3.'],
      ['PKT-MTK1', '1', 'Matematika', 'Berhitung Dasar', 'tap_match', 'Pasangkan angka dengan jumlah apel.', '{"kiri":["1","2","3"],"kanan":["🍎","🍎🍎","🍎🍎🍎"]}', 'urut', 'Hitung dulu apel di sebelah kanan.'],
      ['PKT-MTK1', '1', 'Matematika', 'Berhitung Dasar', 'sort_order', 'Urutkan angka dari yang paling kecil.', '{"item":["5","2","4","1"]}', '1,2,4,5', 'Mulai dari angka yang paling kecil.'],
      ['PKT-IND1', '1', 'Bahasa Indonesia', 'Mengenal Kata', 'tap_match', 'Pasangkan gambar dengan namanya.', '{"kiri":["🐱","🐶","🐟"],"kanan":["Kucing","Anjing","Ikan"]}', 'urut', 'Lihat gambarnya, lalu baca namanya.'],
      ['PKT-IND1', '1', 'Bahasa Indonesia', 'Mengenal Kata', 'sort_order', 'Urutkan kegiatan di pagi hari.', '{"item":["Berangkat sekolah","Bangun tidur","Mandi","Sarapan"]}', 'Bangun tidur,Mandi,Sarapan,Berangkat sekolah', 'Apa yang kamu lakukan pertama kali setelah bangun?']
    ]);
    q.setFrozenRows(1);
  }
  SpreadsheetApp.getUi && SpreadsheetApp.getActive().toast('Database PintarSD siap.');
}

function doGet() {
  return out({ ok: true, msg: 'PintarSD API aktif' });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return out({ ok: false, msg: 'Server sedang sibuk. Coba lagi sebentar.' }); }
  try {
    const req = JSON.parse(e.postData.contents);
    const kode = String(req.kode || '').trim().toUpperCase();
    const dev = String(req.deviceId || '').trim();
    if (!kode || !dev) return out({ ok: false, msg: 'Kode atau perangkat tidak valid.' });

    const sh = SpreadsheetApp.getActive().getSheetByName(SH_L);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() !== kode) continue;
      if (String(rows[i][5]).trim().toUpperCase() !== 'AKTIF') return out({ ok: false, msg: 'Kode ini diblokir. Hubungi penjual.' });

      const d1 = String(rows[i][3]).trim(), d2 = String(rows[i][4]).trim();
      if (d1 !== dev && d2 !== dev) {
        if (!d1) sh.getRange(i + 1, 4).setValue(dev);
        else if (!d2) sh.getRange(i + 1, 5).setValue(dev);
        else return out({ ok: false, msg: 'Batas perangkat penuh 2/2. Kode ini sudah dipakai di 2 perangkat lain.' });
      }
      if (!rows[i][6]) sh.getRange(i + 1, 7).setValue(new Date());
      return out({ ok: true, pembeli: rows[i][1], pakets: getPakets(String(rows[i][2])) });
    }
    return out({ ok: false, msg: 'Kode tidak ditemukan. Periksa penulisannya.' });
  } catch (err) {
    return out({ ok: false, msg: 'Terjadi kesalahan: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function getPakets(idsCsv) {
  const ids = idsCsv.split(',').map(s => s.trim()).filter(Boolean);
  const rows = SpreadsheetApp.getActive().getSheetByName(SH_Q).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i], id = String(r[0]).trim();
    if (!id || ids.indexOf(id) < 0) continue;
    if (!map[id]) map[id] = { id: id, kelas: String(r[1]), mapel: String(r[2]), judul: String(r[3]), soal: [] };
    let data = {};
    try { data = JSON.parse(r[6]); } catch (e) {}
    map[id].soal.push({ tipe: String(r[4]).trim(), pertanyaan: String(r[5]), data: data, kunci: String(r[7]), petunjuk: String(r[8]) });
  }
  return Object.keys(map).map(k => map[k]);
}
