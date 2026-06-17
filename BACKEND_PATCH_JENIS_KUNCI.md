# Patch Backend Apps Script: Pisah Kunci Cowo dan Cewe

Tujuan patch ini:
- Kunci `Cowo 01` dan `Cewe 01` dianggap berbeda.
- Sheet `DATA_KUNCI` menjadi 200 data: Cowo 01-100 dan Cewe 01-100.
- Kolom `Jenis Kunci` muncul di `DATA_KUNCI`, `LOG_GYM`, `REKAP_HARIAN`, dan rekap bulanan.
- Teks `Cewe` otomatis merah di spreadsheet.

## 1. Ganti konstanta header

Ganti `MONTHLY_BLOCK_WIDTH`:

```js
const MONTHLY_BLOCK_WIDTH = 12;
```

Ganti `LOG_HEADERS`:

```js
const LOG_HEADERS = [
  'No',
  'Waktu Lengkap',
  'Tanggal',
  'Jam',
  'Nama',
  'No Kunci',
  'Status',
  'Admin',
  'Jenis Kunci'
];
```

Ganti `KEY_HEADERS`:

```js
const KEY_HEADERS = [
  'No Kunci',
  'Status',
  'Dipakai Oleh',
  'Jam Masuk',
  'Update Terakhir',
  'Jenis Kunci'
];
```

Ganti komentar struktur `REKAP_HARIAN` dan `DAILY_HEADERS`:

```js
// Struktur REKAP_HARIAN lo:
// A = Tanggal
// B = No
// C = Nama
// D = No Kunci
// E = kosong
// F = Jam Masuk
// G = Admin Masuk
// H = Jam Keluar
// I = Sudah Keluar
// J = Waktu Keluar Lengkap
// K = Admin Keluar
// L = Jenis Kunci
const DAILY_HEADERS = [
  'Tanggal',
  'No',
  'Nama',
  'No Kunci',
  '',
  'Jam Masuk',
  'Admin Masuk',
  'Jam Keluar',
  'Sudah Keluar',
  'Waktu Keluar Lengkap',
  'Admin Keluar',
  'Jenis Kunci'
];
```

Tambahkan konstanta ini di bawah `MAX_KEY_NUMBER`:

```js
const KEY_TYPES = ['Cowo', 'Cewe'];
```

## 2. Ganti fungsi `setupHeader_`

```js
function setupHeader_(sheet, headers) {
  const neededCols = headers.length;

  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, neededCols).setValues([headers]);
  styleHeader_(sheet, neededCols);
}
```

## 3. Ganti fungsi `seedKeys_`

```js
function seedKeys_(sheet, maxKey) {
  setupHeader_(sheet, KEY_HEADERS);

  const existingKeys = new Set();
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, KEY_HEADERS.length).getValues();

    values.forEach(function (row, index) {
      const rowIndex = index + 2;
      const key = normalizeKeyNumber_(row[0]);
      let keyType = normalizeKeyType_(row[5]);

      // Data lama yang belum punya jenis kunci dianggap Cowo.
      if (key && !cleanText_(row[5])) {
        keyType = 'Cowo';
        sheet.getRange(rowIndex, 6).setValue(keyType);
      }

      if (key) {
        existingKeys.add(makeKeyIdentity_(keyType, key));
      }
    });
  }

  const rows = [];

  KEY_TYPES.forEach(function (keyType) {
    for (let i = 1; i <= maxKey; i++) {
      const key = String(i).padStart(2, '0');
      const identity = makeKeyIdentity_(keyType, key);

      if (!existingKeys.has(identity)) {
        rows.push([key, 'Kosong', '', '', '', keyType]);
      }
    }
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, KEY_HEADERS.length).setValues(rows);
  }

  styleKeyTypeColumn_(sheet, 6);
}
```

## 4. Ganti fungsi `normalizePayload_`, lalu tambahkan helper jenis kunci

```js
function normalizePayload_(params) {
  const admin = cleanText_(params.admin);
  const customerName = cleanText_(params.customerName || params.nama || params.namaPelanggan);
  const keyNumber = normalizeKeyNumber_(params.keyNumber || params.noKunci || params.nomorKunci);
  const keyType = normalizeKeyType_(params.keyType || params.jenisKunci || params.gender || params.tipeKunci);
  const status = normalizeStatus_(params.status);

  if (!admin) {
    throw new Error('Nama admin/pegawai wajib diisi.');
  }

  if (!keyNumber) {
    throw new Error('Nomor kunci wajib diisi.');
  }

  if (!keyType) {
    throw new Error('Jenis kunci wajib dipilih: Cowo atau Cewe.');
  }

  if (!status) {
    throw new Error('Status tidak valid.');
  }

  if (status === 'Masuk' && !customerName) {
    throw new Error('Nama pelanggan wajib diisi untuk check-in.');
  }

  return {
    admin: admin,
    customerName: customerName,
    keyNumber: keyNumber,
    keyType: keyType,
    status: status,
    timestamp: new Date()
  };
}

function normalizeKeyType_(value) {
  const raw = String(value || '').trim().toLowerCase();

  // Biar request lama yang belum ngirim keyType tetap masuk ke Cowo.
  if (!raw) return 'Cowo';

  if (['cowo', 'cowok', 'pria', 'laki', 'laki-laki', 'male'].indexOf(raw) !== -1) return 'Cowo';
  if (['cewe', 'cewek', 'wanita', 'perempuan', 'female'].indexOf(raw) !== -1) return 'Cewe';

  return '';
}

function makeKeyIdentity_(keyType, keyNumber) {
  return normalizeKeyType_(keyType) + '__' + normalizeKeyNumber_(keyNumber);
}
```

## 5. Di `saveLogFromParams_`, ganti pemanggilan record kunci

Cari:

```js
const currentKey = getKeyRecord_(keySheet, payload.keyNumber);
```

Ganti jadi:

```js
const currentKey = getKeyRecord_(keySheet, payload.keyNumber, payload.keyType);
```

Di bagian `checkoutPayload`, tambahkan `keyType`:

```js
const checkoutPayload = {
  admin: payload.admin,
  customerName: currentKey.customerName || payload.customerName || 'Tanpa Nama',
  keyNumber: payload.keyNumber,
  keyType: payload.keyType,
  status: 'Keluar',
  timestamp: payload.timestamp,
  previousCheckInTime: currentKey.checkInTime || ''
};
```

## 6. Ganti `appendLog_`

```js
function appendLog_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();
  const no = Math.max(sheet.getLastRow(), 1);

  sheet.appendRow([
    no,
    formatDateTime_(timestamp),
    formatDate_(timestamp),
    formatTime_(timestamp),
    payload.customerName,
    payload.keyNumber,
    payload.status,
    payload.admin,
    payload.keyType || 'Cowo'
  ]);

  styleKeyTypeCell_(sheet, sheet.getLastRow(), 9, payload.keyType);
}
```

## 7. Ganti fungsi `appendDailyCheckIn_`

```js
function appendDailyCheckIn_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();

  const tanggal = formatDate_(timestamp);
  const jamMasuk = formatTime_(timestamp);
  const nomorHarian = getNextDailyNumber_(sheet, tanggal);

  const rowValues = [
    tanggal,
    nomorHarian,
    payload.customerName,
    payload.keyNumber,
    '',
    jamMasuk,
    payload.admin,
    '',
    false,
    '',
    '',
    payload.keyType || 'Cowo'
  ];

  sheet.appendRow(rowValues);

  const rowIndex = sheet.getLastRow();

  sheet.getRange(rowIndex, 9).insertCheckboxes();
  sheet.getRange(rowIndex, 9).setValue(false);
  styleKeyTypeCell_(sheet, rowIndex, 12, payload.keyType);

  safeSyncMonthly_();
}
```

## 8. Ganti `markDailyCheckout_` bagian pencarian baris

Cari:

```js
const rowIndex = findOpenDailyRow_(sheet, payload.keyNumber);
```

Ganti jadi:

```js
const rowIndex = findOpenDailyRow_(sheet, payload.keyNumber, payload.keyType);
```

## 9. Ganti `appendRecoveryCheckout_`

```js
function appendRecoveryCheckout_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();

  const tanggal = formatDate_(timestamp);
  const jamKeluar = formatTime_(timestamp);
  const waktuKeluarLengkap = formatDateTime_(timestamp);
  const nomorHarian = getNextDailyNumber_(sheet, tanggal);

  const previousCheckInTime = cleanText_(payload.previousCheckInTime);
  const jamMasuk = extractTimeFromDateTimeText_(previousCheckInTime);

  const rowValues = [
    tanggal,
    nomorHarian,
    payload.customerName,
    payload.keyNumber,
    '',
    jamMasuk || '-',
    'RECOVERY',
    jamKeluar,
    true,
    waktuKeluarLengkap,
    payload.admin,
    payload.keyType || 'Cowo'
  ];

  sheet.appendRow(rowValues);

  const rowIndex = sheet.getLastRow();

  sheet.getRange(rowIndex, 9).insertCheckboxes();
  sheet.getRange(rowIndex, 9).setValue(true);
  styleKeyTypeCell_(sheet, rowIndex, 12, payload.keyType);

  safeSyncMonthly_();
}
```

## 10. Ganti `findOpenDailyRow_`

```js
function findOpenDailyRow_(sheet, keyNumber, keyType) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, DAILY_HEADERS.length).getValues();
  const targetIdentity = makeKeyIdentity_(keyType, keyNumber);

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];

    const rowKunci = normalizeKeyNumber_(row[3]) || cleanText_(row[3]);
    const rowType = normalizeKeyType_(row[11]);
    const rowIdentity = makeKeyIdentity_(rowType, rowKunci);
    const sudahKeluar = isChecked_(row[8]);

    if (rowIdentity === targetIdentity && !sudahKeluar) {
      return i + 2;
    }
  }

  return null;
}
```

## 11. Ganti fungsi DATA_KUNCI

### `updateKey_`

```js
function updateKey_(sheet, payload) {
  const record = getKeyRecord_(sheet, payload.keyNumber, payload.keyType);
  const rowIndex = record.rowIndex || appendKeyRow_(sheet, payload.keyNumber, payload.keyType);
  const nowText = formatDateTime_(new Date());

  if (payload.status === 'Masuk') {
    sheet.getRange(rowIndex, 1, 1, KEY_HEADERS.length).setValues([[
      payload.keyNumber,
      'Dipakai',
      payload.customerName,
      formatDateTime_(payload.timestamp || new Date()),
      nowText,
      payload.keyType || 'Cowo'
    ]]);
  } else {
    sheet.getRange(rowIndex, 1, 1, KEY_HEADERS.length).setValues([[
      payload.keyNumber,
      'Kosong',
      '',
      '',
      nowText,
      payload.keyType || 'Cowo'
    ]]);
  }

  styleKeyTypeCell_(sheet, rowIndex, 6, payload.keyType);
}
```

### `appendKeyRow_`

```js
function appendKeyRow_(sheet, keyNumber, keyType) {
  const rowIndex = sheet.getLastRow() + 1;

  sheet.getRange(rowIndex, 1, 1, KEY_HEADERS.length).setValues([[
    keyNumber,
    'Kosong',
    '',
    '',
    '',
    keyType || 'Cowo'
  ]]);

  styleKeyTypeCell_(sheet, rowIndex, 6, keyType);

  return rowIndex;
}
```

### `getKeyRecord_`

```js
function getKeyRecord_(sheet, keyNumber, keyType) {
  const lastRow = sheet.getLastRow();
  const targetIdentity = makeKeyIdentity_(keyType, keyNumber);

  if (lastRow < 2) {
    return {
      rowIndex: null,
      keyNumber: keyNumber,
      keyType: keyType || 'Cowo',
      status: 'Kosong',
      customerName: '',
      checkInTime: '',
      updatedAt: ''
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, KEY_HEADERS.length).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowKey = normalizeKeyNumber_(row[0]);
    const rowType = normalizeKeyType_(row[5]);
    const rowIdentity = makeKeyIdentity_(rowType, rowKey);

    if (rowIdentity === targetIdentity) {
      return {
        rowIndex: i + 2,
        keyNumber: rowKey,
        keyType: rowType,
        status: cleanText_(row[1]) || 'Kosong',
        customerName: cleanText_(row[2]),
        checkInTime: stringifyCell_(row[3]),
        updatedAt: stringifyCell_(row[4])
      };
    }
  }

  return {
    rowIndex: null,
    keyNumber: keyNumber,
    keyType: keyType || 'Cowo',
    status: 'Kosong',
    customerName: '',
    checkInTime: '',
    updatedAt: ''
  };
}
```

### `getKeys_`

```js
function getKeys_() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, SHEET_KEYS);

  setupHeader_(sheet, KEY_HEADERS);
  seedKeys_(sheet, MAX_KEY_NUMBER);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, KEY_HEADERS.length).getValues();

  return values
    .filter(function (row) {
      return cleanText_(row[0]);
    })
    .map(function (row) {
      return {
        keyNumber: normalizeKeyNumber_(row[0]),
        status: cleanText_(row[1]) || 'Kosong',
        customerName: cleanText_(row[2]),
        checkInTime: stringifyCell_(row[3]),
        updatedAt: stringifyCell_(row[4]),
        keyType: normalizeKeyType_(row[5])
      };
    });
}
```

## 12. Ganti `getLogs_` mapping

Di bagian `.map(function (row) { return { ... } })`, pakai versi ini:

```js
return {
  no: row[0],
  waktuLengkap: stringifyCell_(row[1]),
  tanggal: stringifyCell_(row[2]),
  jam: stringifyCell_(row[3]),
  nama: cleanText_(row[4]),
  noKunci: normalizeKeyNumber_(row[5]) || cleanText_(row[5]),
  status: cleanText_(row[6]),
  admin: cleanText_(row[7]),
  keyType: normalizeKeyType_(row[8])
};
```

## 13. Ganti `getDailyRecap_` mapping

Di bagian `.map(function (row) { return { ... } })`, pakai versi ini:

```js
return {
  tanggal: stringifyDateOnly_(row[0]),
  no: row[1],
  nama: cleanText_(row[2]),
  noKunci: normalizeKeyNumber_(row[3]) || cleanText_(row[3]),
  jamMasuk: stringifyTimeOnly_(row[5]),
  adminMasuk: cleanText_(row[6]),
  jamKeluar: stringifyTimeOnly_(row[7]),
  sudahKeluar: isChecked_(row[8]),
  waktuKeluarLengkap: stringifyDateTimeSafe_(row[9]),
  adminKeluar: cleanText_(row[10]),
  keyType: normalizeKeyType_(row[11])
};
```

## 14. Di `syncSemuaRekapHarianKeBulanan()`, tambah index dan data jenis kunci

Di object `idx`, tambahkan:

```js
keyType: findHeaderIndex_(sourceHeaders, ['jenis kunci', 'gender', 'tipe kunci'])
```

Lalu di `normalizedRow`, tambahkan item terakhir:

```js
normalizeKeyType_(getRawRowValue_(row, idx.keyType))
```

Sehingga `normalizedRow` menjadi 12 kolom:

```js
const normalizedRow = [
  formatDateOnlySafe_(getRawRowValue_(row, idx.tanggal)),
  getRawRowValue_(row, idx.no),
  nama,
  getRawRowValue_(row, idx.noKunci),
  '',
  stringifyTimeOnly_(getRawRowValue_(row, idx.jamMasuk)),
  getRawRowValue_(row, idx.adminMasuk),
  stringifyTimeOnly_(getRawRowValue_(row, idx.jamKeluar)),
  isChecked_(getRawRowValue_(row, idx.sudahKeluar)),
  stringifyDateTimeSafe_(getRawRowValue_(row, idx.waktuKeluarLengkap)),
  getRawRowValue_(row, idx.adminKeluar),
  normalizeKeyType_(getRawRowValue_(row, idx.keyType))
];
```

## 15. Ganti `writeOneRowToMonthlyBlock_`

```js
function writeOneRowToMonthlyBlock_(sheet, startCol, rowIndex, rowValues) {
  const range = sheet.getRange(rowIndex, startCol, 1, MONTHLY_BLOCK_WIDTH);

  // Paksa text supaya Jam Masuk/Jam Keluar tidak berubah jadi 12/30/1899.
  range.setNumberFormat('@');
  range.setValues([rowValues]);

  const checkboxCol = startCol + 8;
  const checkboxCell = sheet.getRange(rowIndex, checkboxCol);

  checkboxCell.insertCheckboxes();
  checkboxCell.setValue(rowValues[8] === true);

  styleKeyTypeCell_(sheet, rowIndex, startCol + 11, rowValues[11]);
}
```

## 16. Ganti `setMonthlyBlockColumnWidths_`

```js
function setMonthlyBlockColumnWidths_(sheet, startCol) {
  const widths = [
    110, // Tanggal
    55,  // No
    220, // Nama
    90,  // No Kunci
    35,  // Kosong
    120, // Jam Masuk
    120, // Admin Masuk
    120, // Jam Keluar
    120, // Sudah Keluar
    180, // Waktu Keluar Lengkap
    120, // Admin Keluar
    110  // Jenis Kunci
  ];

  widths.forEach(function (width, index) {
    sheet.setColumnWidth(startCol + index, width);
  });
}
```

## 17. Tambahkan helper warna di bagian GENERAL UTILITIES

```js
function styleKeyTypeCell_(sheet, rowIndex, colIndex, keyType) {
  const normalized = normalizeKeyType_(keyType);
  const color = normalized === 'Cewe' ? '#dc2626' : '#2563eb';

  sheet.getRange(rowIndex, colIndex)
    .setFontColor(color)
    .setFontWeight('bold');
}

function styleKeyTypeColumn_(sheet, colIndex) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const values = range.getValues();

  const colors = values.map(function (row) {
    return [normalizeKeyType_(row[0]) === 'Cewe' ? '#dc2626' : '#2563eb'];
  });

  const weights = values.map(function () {
    return ['bold'];
  });

  range.setFontColors(colors);
  range.setFontWeights(weights);
}
```

## 18. Setelah update backend

Jalankan fungsi manual di Apps Script:

```js
setupGymSheets()
setupRekapBulananDariHarian()
syncSemuaRekapHarianKeBulanan()
```

Setelah itu, `DATA_KUNCI` akan berisi:
- `Cowo 01` sampai `Cowo 100`
- `Cewe 01` sampai `Cewe 100`

Catatan: Data lama yang belum punya kolom jenis kunci otomatis dianggap `Cowo`.
