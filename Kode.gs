// =================================================================
// 1. ROUTING UTAMA
// =================================================================
function doGet(e) {
  checkAndCreateTables();

  if (e && e.parameter && e.parameter.callback) {
    return handleJsonpRequest(e);
  }

  return HtmlService
    .createHtmlOutput('<!doctype html><meta charset="utf-8"><title>API Pramuka</title><body style="font-family:system-ui;margin:32px"><h1>API Pramuka aktif</h1><p>Frontend aplikasi dijalankan dari GitHub Pages.</p></body>')
    .setTitle('API Pramuka SMKN 2 Marabahan');
}

function handleJsonpRequest(e) {
  const requestId = String(e.parameter.requestId || '');
  const callback = sanitizeCallbackName(e.parameter.callback);
  let result;
  let ok = true;
  let message = '';

  try {
    const payload = JSON.parse(e.parameter.payload || '{}');
    result = runApiAction(payload.action, payload.args || []);
  } catch (error) {
    ok = false;
    message = error.message;
    result = { status: false, message: error.message };
  }

  const response = {
    source: 'pramuka-api',
    requestId: requestId,
    ok: ok,
    message: message,
    result: result
  };

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(response) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function sanitizeCallbackName(value) {
  const callback = String(value || '').trim();
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
    throw new Error('Callback JSONP tidak valid.');
  }
  return callback;
}

function doPost(e) {
  checkAndCreateTables();

  const requestId = e && e.parameter ? String(e.parameter.requestId || '') : '';
  let result;
  let ok = true;
  let message = '';

  try {
    const payload = JSON.parse(e.parameter.payload || '{}');
    result = runApiAction(payload.action, payload.args || []);
  } catch (error) {
    ok = false;
    message = error.message;
    result = { status: false, message: error.message };
  }

  const response = {
    source: 'pramuka-api',
    requestId: requestId,
    ok: ok,
    message: message,
    result: result
  };

  return HtmlService
    .createHtmlOutput('<!doctype html><meta charset="utf-8"><script>window.parent.postMessage(' + JSON.stringify(response) + ', "*");</script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function runApiAction(action, args) {
  const allowed = {
    loginUser: loginUser,
    updateProfilPembina: updateProfilPembina,
    getSiswaByPembina: getSiswaByPembina,
    tambahSiswa: tambahSiswa,
    updateSiswa: updateSiswa,
    hapusSiswa: hapusSiswa,
    simpanPresensi: simpanPresensi,
    getPresensiByTanggal: getPresensiByTanggal,
    simpanNilai: simpanNilai,
    simpanNilaiBatch: simpanNilaiBatch,
    getNilaiByNisList: getNilaiByNisList,
    getNilaiByNis: getNilaiByNis,
    getRekapData: getRekapData
  };

  if (!allowed[action]) {
    throw new Error('Aksi API tidak dikenal: ' + action);
  }

  return allowed[action].apply(null, args);
}

// =================================================================
// 2. INISIALISASI DATABASE OTOMATIS
// =================================================================
function checkAndCreateTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = {
    'Pembina': ['ID_Pembina', 'Nama_Pembina', 'Username', 'Password'],
    'Siswa': ['NIS', 'Nama_Siswa', 'Sangga', 'ID_Pembina'],
    'Presensi': ['Tanggal', 'NIS', 'Status', 'ID_Pembina'],
    'Nilai': ['NIS', 'N2_Ketaqwaan', 'N3_Disiplin', 'N4_Sikap', 'N5_Keaktifan', 'N6_GotongRoyong', 'N7_Kepemimpinan', 'N8_Keterampilan', 'N9_Partisipasi']
  };

  for (let sheetName in sheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheets[sheetName]);
      // Membekukan baris pertama (header)
      sheet.setFrozenRows(1); 
      // Menebalkan font header
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setFontWeight("bold");
    }
  }
  
  // Masukkan data dummy Pembina HANYA jika sheet Pembina kosong (selain header)
  let sheetPembina = ss.getSheetByName('Pembina');
  if (sheetPembina.getLastRow() <= 1) {
    sheetPembina.appendRow(['P01', 'Kak Budi', 'budi', '12345']);
    sheetPembina.appendRow(['P02', 'Kak Siti', 'siti', '12345']);
  }
}

// =================================================================
// 3. OTENTIKASI
// =================================================================
function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pembina');
    
    if (!sheet) {
      return { status: false, message: 'Database Pembina tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    
    // Loop mulai dari baris ke-2
    for (let i = 1; i < data.length; i++) {
      let dbUser = String(data[i][2]).trim();
      let dbPass = String(data[i][3]).trim();
      let inputUser = String(username).trim();
      let inputPass = String(password).trim();
      
      if (dbUser === inputUser && dbPass === inputPass) {
        return { status: true, idPembina: data[i][0], nama: data[i][1], username: data[i][2] };
      }
    }
    return { status: false, message: 'Username atau Password salah!' };
  } catch (error) {
    return { status: false, message: 'Terjadi kesalahan sistem: ' + error.message };
  }
}

// =================================================================
// 4. DATA MASTER SISWA
// =================================================================
function getSiswaByPembina(idPembina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  if(!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let result = [];
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(idPembina)) {
      result.push({ nis: data[i][0], nama: data[i][1], sangga: data[i][2] });
    }
  }
  return result;
}

function tambahSiswa(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Siswa');
    const nis = String(data.nis || '').trim();
    const nama = String(data.nama || '').trim();
    const sangga = String(data.sangga || '').trim();

    if (!nis || !nama || !sangga) {
      return { status: false, message: "NIS, nama, dan sangga wajib diisi." };
    }

    const dataSiswa = sheet.getDataRange().getValues();
    
    for(let i = 1; i < dataSiswa.length; i++) {
      if(String(dataSiswa[i][0]) === nis) {
         return { status: false, message: "NIS sudah terdaftar!" };
      }
    }
    
    sheet.appendRow([nis, nama, sangga, data.idPembina]);
    return { status: true, message: 'Siswa berhasil ditambahkan' };
  } catch (e) { 
    return { status: false, message: e.message }; 
  }
}

function updateProfilPembina(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pembina');
    const idPembina = String(data.idPembina || '').trim();
    const nama = String(data.nama || '').trim();
    const username = String(data.username || '').trim();
    const password = String(data.password || '').trim();

    if (!idPembina || !nama || !username) {
      return { status: false, message: 'Nama dan username wajib diisi.' };
    }

    const rows = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < rows.length; i++) {
      const rowId = String(rows[i][0]).trim();
      const rowUsername = String(rows[i][2]).trim();
      if (rowUsername === username && rowId !== idPembina) {
        return { status: false, message: 'Username sudah digunakan pembina lain.' };
      }
      if (rowId === idPembina) {
        targetRow = i + 1;
      }
    }

    if (targetRow === -1) {
      return { status: false, message: 'Profil pembina tidak ditemukan.' };
    }

    const currentPassword = String(rows[targetRow - 1][3] || '').trim();
    sheet.getRange(targetRow, 2, 1, 3).setValues([[nama, username, password || currentPassword]]);
    return { status: true, idPembina: idPembina, nama: nama, username: username, message: 'Profil pembina berhasil diperbarui.' };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function updateSiswa(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Siswa');
    const oldNis = String(data.oldNis || '').trim();
    const nis = String(data.nis || '').trim();
    const nama = String(data.nama || '').trim();
    const sangga = String(data.sangga || '').trim();
    const idPembina = String(data.idPembina || '').trim();

    if (!oldNis || !nis || !nama || !sangga || !idPembina) {
      return { status: false, message: "Data edit siswa belum lengkap." };
    }

    const rows = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < rows.length; i++) {
      const rowNis = String(rows[i][0]).trim();
      const rowPembina = String(rows[i][3]).trim();
      if (rowNis === nis && rowNis !== oldNis) {
        return { status: false, message: "NIS baru sudah dipakai siswa lain." };
      }
      if (rowNis === oldNis && rowPembina === idPembina) {
        targetRow = i + 1;
      }
    }

    if (targetRow === -1) {
      return { status: false, message: "Data siswa tidak ditemukan." };
    }

    sheet.getRange(targetRow, 1, 1, 4).setValues([[nis, nama, sangga, idPembina]]);

    if (nis !== oldNis) {
      updateNisReference(ss.getSheetByName('Presensi'), oldNis, nis, idPembina, 2, 4);
      updateNisReference(ss.getSheetByName('Nilai'), oldNis, nis, null, 1, null);
    }

    return { status: true, message: "Data siswa berhasil diperbarui." };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function hapusSiswa(nis, idPembina) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nisKey = String(nis || '').trim();
    const pembinaKey = String(idPembina || '').trim();

    if (!nisKey || !pembinaKey) {
      return { status: false, message: "Data siswa yang akan dihapus belum lengkap." };
    }

    deleteRowsByMatch(ss.getSheetByName('Siswa'), nisKey, pembinaKey, 1, 4);
    deleteRowsByMatch(ss.getSheetByName('Presensi'), nisKey, pembinaKey, 2, 4);
    deleteRowsByMatch(ss.getSheetByName('Nilai'), nisKey, null, 1, null);

    return { status: true, message: "Siswa dan data terkait berhasil dihapus." };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

// =================================================================
// 5. TRANSAKSI PRESENSI & NILAI
// =================================================================
function simpanPresensi(tanggal, presensiList, idPembina) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Presensi');

    if (!tanggal) {
      return { status: false, message: 'Tanggal latihan wajib diisi.' };
    }

    if (!Array.isArray(presensiList) || presensiList.length === 0) {
      return { status: false, message: 'Tidak ada data presensi untuk disimpan.' };
    }

    const allData = sheet.getDataRange().getValues();
    const rowMap = {};

    for (let index = 1; index < allData.length; index++) {
      const row = allData[index];
      const key = [
        normalizeDateKey(row[0]),
        String(row[1]).trim(),
        String(row[3]).trim()
      ].join('|');
      rowMap[key] = index;
    }

    presensiList.forEach(p => {
      const rowData = [tanggal, p.nis, p.status, idPembina];
      const key = [
        normalizeDateKey(tanggal),
        String(p.nis).trim(),
        String(idPembina).trim()
      ].join('|');

      if (typeof rowMap[key] === 'number') {
        allData[rowMap[key]] = rowData;
      } else {
        rowMap[key] = allData.length;
        allData.push(rowData);
      }
    });

    sheet.getRange(1, 1, allData.length, 4).setValues(allData);
    return { status: true, message: presensiList.length + ' data presensi tanggal ' + tanggal + ' berhasil disimpan atau diperbarui.' };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function getPresensiByTanggal(tanggal, idPembina) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Presensi');
    if (!tanggal || !sheet || sheet.getLastRow() <= 1) {
      return { status: true, presensi: {} };
    }

    const targetDate = normalizeDateKey(tanggal);
    const targetPembina = String(idPembina || '').trim();
    const rows = sheet.getDataRange().getValues();
    let presensi = {};

    for (let i = 1; i < rows.length; i++) {
      const sameDate = normalizeDateKey(rows[i][0]) === targetDate;
      const samePembina = String(rows[i][3]).trim() === targetPembina;
      if (sameDate && samePembina) {
        presensi[String(rows[i][1]).trim()] = String(rows[i][2]).trim().toUpperCase();
      }
    }

    return { status: true, presensi: presensi };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function simpanNilai(dataNilai) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nilai');
    const allData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    const aspek = ['n2', 'n3', 'n4', 'n5', 'n6', 'n7', 'n8', 'n9'];
    let nilaiBersih = {};

    if (!dataNilai.nis) {
      return { status: false, message: 'Pilih siswa terlebih dahulu.' };
    }

    for (let i = 0; i < aspek.length; i++) {
      const key = aspek[i];
      const rawNilai = dataNilai[key] === '' || dataNilai[key] === null || typeof dataNilai[key] === 'undefined' ? 0 : dataNilai[key];
      const nilai = Number(rawNilai);
      if (isNaN(nilai) || nilai < 0 || nilai > 100) {
        return { status: false, message: 'Nilai ' + key.toUpperCase() + ' harus berada pada rentang 0 sampai 100.' };
      }
      nilaiBersih[key] = nilai;
    }
    
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(dataNilai.nis)) { rowIndex = i + 1; break; }
    }
    
    const rowData = [
      dataNilai.nis, nilaiBersih.n2, nilaiBersih.n3, nilaiBersih.n4, 
      nilaiBersih.n5, nilaiBersih.n6, nilaiBersih.n7, nilaiBersih.n8, nilaiBersih.n9
    ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { status: true, message: 'Nilai berhasil disimpan.' };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function simpanNilaiBatch(nilaiList) {
  try {
    if (!Array.isArray(nilaiList) || nilaiList.length === 0) {
      return { status: false, message: 'Tidak ada data nilai untuk disimpan.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nilai');
    const allData = sheet.getDataRange().getValues();
    const aspek = ['n2', 'n3', 'n4', 'n5', 'n6', 'n7', 'n8', 'n9'];
    const rowByNis = {};

    for (let i = 1; i < allData.length; i++) {
      rowByNis[String(allData[i][0])] = i;
    }

    for (let i = 0; i < nilaiList.length; i++) {
      const dataNilai = nilaiList[i] || {};
      if (!dataNilai.nis) {
        return { status: false, message: 'Baris ' + (i + 1) + ': NIS siswa tidak ditemukan.' };
      }

      const nilaiBersih = {};
      for (let j = 0; j < aspek.length; j++) {
        const key = aspek[j];
        const rawNilai = dataNilai[key] === '' || dataNilai[key] === null || typeof dataNilai[key] === 'undefined' ? 0 : dataNilai[key];
        const nilai = Number(rawNilai);
        if (isNaN(nilai) || nilai < 0 || nilai > 100) {
          return { status: false, message: 'Baris ' + (i + 1) + ': Nilai ' + key.toUpperCase() + ' harus berada pada rentang 0 sampai 100.' };
        }
        nilaiBersih[key] = nilai;
      }

      const rowData = [
        dataNilai.nis, nilaiBersih.n2, nilaiBersih.n3, nilaiBersih.n4,
        nilaiBersih.n5, nilaiBersih.n6, nilaiBersih.n7, nilaiBersih.n8, nilaiBersih.n9
      ];
      const existingIndex = rowByNis[String(dataNilai.nis)];
      if (typeof existingIndex === 'number') {
        allData[existingIndex] = rowData;
      } else {
        rowByNis[String(dataNilai.nis)] = allData.length;
        allData.push(rowData);
      }
    }

    sheet.getRange(1, 1, allData.length, 9).setValues(allData);
    return { status: true, message: nilaiList.length + ' data nilai berhasil disimpan.' };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function getNilaiByNisList(nisList) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nilai');
    let result = {};
    if (!sheet || sheet.getLastRow() <= 1 || !Array.isArray(nisList)) {
      return { status: true, nilai: result };
    }

    const wanted = {};
    nisList.forEach(nis => wanted[String(nis).trim()] = true);

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const nis = String(data[i][0]).trim();
      if (wanted[nis]) {
        result[nis] = {
          n2: data[i][1] === '' ? '' : data[i][1],
          n3: data[i][2] === '' ? '' : data[i][2],
          n4: data[i][3] === '' ? '' : data[i][3],
          n5: data[i][4] === '' ? '' : data[i][4],
          n6: data[i][5] === '' ? '' : data[i][5],
          n7: data[i][6] === '' ? '' : data[i][6],
          n8: data[i][7] === '' ? '' : data[i][7],
          n9: data[i][8] === '' ? '' : data[i][8]
        };
      }
    }

    return { status: true, nilai: result };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

function getNilaiByNis(nis) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nilai');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: true, nilai: null };
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(nis)) {
        return {
          status: true,
          nilai: {
            n2: data[i][1] === '' ? '' : data[i][1],
            n3: data[i][2] === '' ? '' : data[i][2],
            n4: data[i][3] === '' ? '' : data[i][3],
            n5: data[i][4] === '' ? '' : data[i][4],
            n6: data[i][5] === '' ? '' : data[i][5],
            n7: data[i][6] === '' ? '' : data[i][6],
            n8: data[i][7] === '' ? '' : data[i][7],
            n9: data[i][8] === '' ? '' : data[i][8]
          }
        };
      }
    }

    return { status: true, nilai: null };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

// =================================================================
// 6. REKAPITULASI
// =================================================================
function getRekapData(idPembina) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const getSheetData = (sheetName) => {
      let sh = ss.getSheetByName(sheetName);
      if(!sh || sh.getLastRow() <= 1) return [];
      return sh.getDataRange().getValues();
    };

    const siswaData = getSheetData('Siswa');
    const presensiData = getSheetData('Presensi');
    const nilaiData = getSheetData('Nilai');

    let rekap = [];
    let dates = new Set();
    
    if(presensiData.length > 1) {
      for(let i = 1; i < presensiData.length; i++){
        if(String(presensiData[i][3]) === String(idPembina)) {
           dates.add(normalizeDateKey(presensiData[i][0])); 
        }
      }
    }
    let totalPertemuan = dates.size;

    if(siswaData.length > 1) {
      for (let i = 1; i < siswaData.length; i++) {
        if (String(siswaData[i][3]) === String(idPembina)) {
          let nis = siswaData[i][0];
          let nama = siswaData[i][1];
          let sangga = siswaData[i][2];
          
          let hadir = 0, izin = 0, sakit = 0, alpa = 0;
          let statusPerTanggal = {};
          
          if(presensiData.length > 1) {
            for(let p = 1; p < presensiData.length; p++) {
              if(String(presensiData[p][1]) === String(nis) && String(presensiData[p][3]) === String(idPembina)) {
                statusPerTanggal[normalizeDateKey(presensiData[p][0])] = String(presensiData[p][2]).toUpperCase();
              }
            }
          }

          Object.keys(statusPerTanggal).forEach(tgl => {
            let status = statusPerTanggal[tgl];
            if(status === 'H') hadir++;
            else if(status === 'I') izin++;
            else if(status === 'S') sakit++;
            else alpa++;
          });
          
          let persentaseHadir = totalPertemuan === 0 ? 0 : (hadir / totalPertemuan) * 100;
          let n = { n2:0, n3:0, n4:0, n5:0, n6:0, n7:0, n8:0, n9:0 };
          
          if(nilaiData.length > 1) {
            for(let k = 1; k < nilaiData.length; k++) {
              if(String(nilaiData[k][0]) === String(nis)) {
                n = {
                  n2: Number(nilaiData[k][1]) || 0, n3: Number(nilaiData[k][2]) || 0, 
                  n4: Number(nilaiData[k][3]) || 0, n5: Number(nilaiData[k][4]) || 0,
                  n6: Number(nilaiData[k][5]) || 0, n7: Number(nilaiData[k][6]) || 0, 
                  n8: Number(nilaiData[k][7]) || 0, n9: Number(nilaiData[k][8]) || 0
                };
                break;
              }
            }
          }

          let totalSkor = (persentaseHadir * 0.10) + (n.n2 * 0.05) + (n.n3 * 0.10) + 
                          (n.n4 * 0.15) + (n.n5 * 0.10) + (n.n6 * 0.10) + 
                          (n.n7 * 0.10) + (n.n8 * 0.20) + (n.n9 * 0.10);

          rekap.push({
            nis: nis, nama: nama, sangga: sangga,
            presensi: { h: hadir, i: izin, s: sakit, a: alpa, skor: persentaseHadir },
            nilai: n, skorAkhir: totalSkor.toFixed(2)
          });
        }
      }
    }
    
    return { status: true, rekap: rekap, totalPertemuan: totalPertemuan };
  } catch (error) {
    return { status: false, message: error.message };
  }
}

function normalizeDateKey(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).trim();
}

function updateNisReference(sheet, oldNis, newNis, idPembina, nisCol, pembinaCol) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const sameNis = String(data[i][nisCol - 1]).trim() === oldNis;
    const samePembina = !pembinaCol || String(data[i][pembinaCol - 1]).trim() === String(idPembina).trim();
    if (sameNis && samePembina) {
      sheet.getRange(i + 1, nisCol).setValue(newNis);
    }
  }
}

function deleteRowsByMatch(sheet, nis, idPembina, nisCol, pembinaCol) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const sameNis = String(data[i][nisCol - 1]).trim() === nis;
    const samePembina = !pembinaCol || String(data[i][pembinaCol - 1]).trim() === String(idPembina).trim();
    if (sameNis && samePembina) {
      sheet.deleteRow(i + 1);
    }
  }
}
