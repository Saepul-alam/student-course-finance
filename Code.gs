/**
 * LMS + Buku Tabungan - Google Apps Script
 * Manajemen Siswa, Kelas, Absensi, Progres, dan Buku Tabungan Terintegrasi
 */

// ==================== CONSTANTS ====================
const CONFIG = {
  STUDENTS_SHEET: 'Murid',
  CLASSES_SHEET: 'Kelas',
  ATTENDANCE_SHEET: 'Absensi',
  PROGRESS_SHEET: 'Progres',
  SAVINGS_SHEET: 'Tabungan',
  TRANSACTIONS_SHEET: 'Transaksi',
  SETTINGS_SHEET: 'Settings',
  USERS_SHEET: 'Users',
  ACTIVITY_SHEET: 'Log Aktivitas'
};

// ==================== ON OPEN MENU ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LMS & Tabungan')
    .addItem('🏠 Dashboard', 'showDashboard')
    .addItem('🌐 Buka Web App', 'showOpenWebApp')
    .addSeparator()
    .addSubMenu(ui.createMenu('Murid')
      .addItem('📋 Daftar Murid', 'showStudentList'))
    .addSubMenu(ui.createMenu('Kelas')
      .addItem('📋 Daftar Kelas', 'showClassList'))
    .addSubMenu(ui.createMenu('User')
      .addItem('👥 Daftar User', 'showUserList'))
    .addSubMenu(ui.createMenu('Log')
      .addItem('📜 Riwayat Aktivitas', 'showActivityLog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Absensi')
      .addItem('📝 Absensi Hari Ini', 'showAttendance')
      .addItem('📊 Riwayat Absensi', 'showAttendanceHistory'))
    .addSubMenu(ui.createMenu('Progres')
      .addItem('📊 Riwayat Progres', 'showProgressHistory'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Tabungan')
      .addItem('📊 Riwayat Transaksi', 'showTransactionHistory')
      .addItem('🏦 Cek Saldo', 'showBalanceCheck'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Laporan Otomatis')
      .addItem('▶️ Pasang Trigger Bulanan', 'installMonthlyReportTrigger')
      .addItem('⏹️ Hentikan Trigger', 'removeMonthlyReportTrigger')
      .addItem('🔍 Cek Status Trigger', 'showTriggerStatus')
      .addItem('🧪 Test Kirim Sekarang', 'showTestReport'))
    .addSubMenu(ui.createMenu('Pengaturan WhatsApp')
      .addItem('⚙️ Konfigurasi Fonnte', 'showWhatsAppConfig')
      .addItem('🧪 Test Kirim WA', 'showWhatsAppTest'))
    .addSeparator()
    .addItem('⚙️ Setup Awal', 'showSetup')
    .addItem('ℹ️ Bantuan', 'showHelp')
    .addToUi();
}

/**
 * Dialog kecil berisi link ke web app (dipakai beberapa menu).
 */
function showOpenWebApp() {
  const url = getWebAppUrl();
  const html = `
    <div style="padding:20px; font-family:Arial; text-align:center;">
      <h2 style="margin-top:0;">🌐 LMS & Tabungan Web App</h2>
      <p style="color:#666;">Semua fitur input (murid, kelas, absensi, transaksi,<br>
      laporan, WhatsApp) ada di web app:</p>
      <p><a href="${url}" target="_blank" style="display:inline-block; padding:12px 24px; background:#1a73e8; color:white; text-decoration:none; border-radius:8px; font-weight:bold;">🚀 Buka Web App</a></p>
      <p style="font-size:0.8rem; color:#999; margin-top:15px; word-break:break-all;">${url}</p>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(280), 'Buka Web App');
}

/**
 * Ambil URL web app. Jika belum ada deployment, beri petunjuk.
 */
function getWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    return url || '(Belum di-deploy — buka Apps Script editor → Deploy → New deployment → Web app)';
  } catch (e) {
    return '(Belum di-deploy — buka Apps Script editor → Deploy → New deployment → Web app)';
  }
}

// ==================== SETUP SHEETS ====================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create Murid sheet
  if (!ss.getSheetByName(CONFIG.STUDENTS_SHEET)) {
    const studentsSheet = ss.insertSheet(CONFIG.STUDENTS_SHEET);
    studentsSheet.getRange(1, 1, 1, 9).setValues([
      ['ID Murid', 'Nama', 'Kelas ID', 'Tanggal Lahir', 'Email Orang Tua', 'No HP', 'Status', 'Tanggal Daftar', 'ID Tabungan']
    ]);
    studentsSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  }

  // Create Kelas sheet
  if (!ss.getSheetByName(CONFIG.CLASSES_SHEET)) {
    const classesSheet = ss.insertSheet(CONFIG.CLASSES_SHEET);
    classesSheet.getRange(1, 1, 1, 7).setValues([
      ['ID Kelas', 'Nama Kelas', 'Guru', 'Jadwal', 'Kapasitas', 'Biaya Bulanan', 'Status']
    ]);
    classesSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#34A853').setFontColor('white');
  }

  // Create Absensi sheet
  if (!ss.getSheetByName(CONFIG.ATTENDANCE_SHEET)) {
    const attendanceSheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET);
    attendanceSheet.getRange(1, 1, 1, 7).setValues([
      ['ID Absensi', 'ID Murid', 'Tanggal', 'Status', 'Catatan', 'Waktu Check-in', 'Waktu Check-out']
    ]);
    attendanceSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#FBBC05').setFontColor('black');
  }

  // Create Progres sheet
  if (!ss.getSheetByName(CONFIG.PROGRESS_SHEET)) {
    const progressSheet = ss.insertSheet(CONFIG.PROGRESS_SHEET);
    progressSheet.getRange(1, 1, 1, 8).setValues([
      ['ID Progres', 'ID Murid', 'Mata Pelajaran', 'Topik', 'Nilai', 'Deskripsi', 'Tanggal', 'Guru']
    ]);
    progressSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#9C27B0').setFontColor('white');
  }

  // Create Tabungan sheet
  if (!ss.getSheetByName(CONFIG.SAVINGS_SHEET)) {
    const savingsSheet = ss.insertSheet(CONFIG.SAVINGS_SHEET);
    savingsSheet.getRange(1, 1, 1, 8).setValues([
      ['ID Tabungan', 'ID Murid', 'Saldo Awal', 'Total Setoran', 'Total Penarikan', 'Saldo Saat Ini', 'Tanggal Buka', 'Status']
    ]);
    savingsSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#0F9D58').setFontColor('white');
  }

  // Create Transaksi sheet
  if (!ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET)) {
    const transactionsSheet = ss.insertSheet(CONFIG.TRANSACTIONS_SHEET);
    transactionsSheet.getRange(1, 1, 1, 9).setValues([
      ['ID Transaksi', 'ID Tabungan', 'ID Murid', 'Jenis', 'Jumlah', 'Saldo Setelah', 'Catatan', 'Tanggal', 'Waktu']
    ]);
    transactionsSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#FF6D01').setFontColor('white');
  }

  // Create Settings sheet
  if (!ss.getSheetByName(CONFIG.SETTINGS_SHEET)) {
    const settingsSheet = ss.insertSheet(CONFIG.SETTINGS_SHEET);
    settingsSheet.getRange(1, 1, 4, 2).setValues([
      ['Parameter', 'Nilai'],
      ['Nama Sekolah', ''],
      ['Guru Pengampu', ''],
      ['Tahun Ajaran', '']
    ]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#607D8B').setFontColor('white');
  }

  // Create Users sheet (daftar email yang boleh login web app)
  if (!ss.getSheetByName(CONFIG.USERS_SHEET)) {
    const usersSheet = ss.insertSheet(CONFIG.USERS_SHEET);
    usersSheet.getRange(1, 1, 1, 5).setValues([
      ['Email', 'Nama', 'Peran', 'Status', 'Dibuat']
    ]);
    usersSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#795548').setFontColor('white');

    // Seed awal: isi dgn email di ALLOWED_USERS sebagai Admin agar sistem tidak terkunci
    const seed = ALLOWED_USERS
      .filter(e => e && e.indexOf('@') !== -1)
      .map(e => [e, '', 'Admin', 'Aktif', new Date()]);
    if (seed.length > 0) {
      usersSheet.getRange(2, 1, seed.length, 5).setValues(seed);
    }
  }

  // Create Log Aktivitas sheet (audit trail: siapa, apa, kapan)
  if (!ss.getSheetByName(ACTIVITY_LOG_SHEET)) {
    const actSheet = ss.insertSheet(ACTIVITY_LOG_SHEET);
    actSheet.getRange(1, 1, 1, 5).setValues([
      ['Waktu', 'User', 'Aksi', 'Detail', 'Target']
    ]);
    actSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#455A64').setFontColor('white');
  }

  // Freeze header rows
  [CONFIG.STUDENTS_SHEET, CONFIG.CLASSES_SHEET, CONFIG.ATTENDANCE_SHEET,
   CONFIG.PROGRESS_SHEET, CONFIG.SAVINGS_SHEET, CONFIG.TRANSACTIONS_SHEET,
   CONFIG.USERS_SHEET, ACTIVITY_LOG_SHEET
  ].forEach(sheetName => {
    ss.getSheetByName(sheetName).setFrozenRows(1);
  });

  return '✅ Setup selesai! Semua sheet siap digunakan.';
}

function showSetup() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '⚙️ Setup Awal Sistem',
    'Script akan membuat semua sheet (Murid, Kelas, Absensi, Progres, Tabungan, Transaksi, Settings).\n\nSheet yang sudah ada TIDAK akan diubah. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const message = setupSheets();
    ui.alert(message);
  }
}

// ==================== STUDENT MANAGEMENT ====================
function addStudent(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  const students = sheet.getDataRange().getValues();

  // Write-through: invalidasi cache dulu sebelum tulis
  invalidateDataCache();

  // Generate ID Murid
  const studentId = genId('S');

  // Generate ID Tabungan
  const savingsId = genId('TAB');

  // Get class info if class ID provided
  let className = '';
  if (data.kelasId) {
    const classes = ss.getSheetByName(CONFIG.CLASSES_SHEET).getDataRange().getValues();
    for (let i = 1; i < classes.length; i++) {
      if (classes[i][0] == data.kelasId) {
        className = classes[i][1];
        break;
      }
    }
  }

  const newRow = [
    studentId,
    data.nama,
    data.kelasId || '',
    data.tanggalLahir || '',
    data.email || '',
    data.noHP || '',
    data.status || 'Aktif',
    new Date(),
    savingsId
  ];

  sheet.appendRow(newRow);

  // Auto-create savings account
  createSavingsAccount(savingsId, studentId, data.nama, 0);

  logActivity('Tambah Murid', 'ID ' + studentId + ' · Rekening ' + savingsId, data.nama);
  return { success: true, studentId, savingsId, message: 'Murid berhasil ditambahkan!' };
}

function getStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const classes = getCachedClasses();
  const classMap = {};
  classes.forEach(c => classMap[c.id] = c.nama);

  return data.slice(1).map(row => ({
    id: row[0],
    nama: row[1],
    kelasId: row[2],
    kelasNama: classMap[row[2]] || '-',
    tanggalLahir: row[3],
    email: row[4],
    noHP: row[5],
    status: row[6],
    tanggalDaftar: row[7],
    savingsId: row[8]
  }));
}

function updateStudent(id, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STUDENTS_SHEET);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == id) {
      values[i][1] = data.nama || values[i][1];
      values[i][2] = data.kelasId || values[i][2];
      values[i][3] = data.tanggalLahir || values[i][3];
      values[i][4] = data.email || values[i][4];
      values[i][5] = data.noHP || values[i][5];
      values[i][6] = data.status || values[i][6];
      sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
      invalidateDataCache();
      logActivity('Edit Murid', 'Perubahan data disimpan', id);
      return { success: true, message: 'Data murid berhasil diperbarui!' };
    }
  }
  return { success: false, message: 'Murid tidak ditemukan!' };
}

function showStudentList() {
  const students = getCachedStudents();
  const balanceMap = getBalanceMap();
  let html = '<div style="padding: 15px;">';
  html += '<h2>Daftar Murid</h2>';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 15px;">';
  html += '<tr style="background:#4285F4; color:white;">';
  html += '<th>No</th><th>Nama</th><th>Kelas</th><th>Status</th><th>Saldo</th><th>Aksi</th>';
  html += '</tr>';

  students.forEach((s, i) => {
    const balance = balanceMap[s.savingsId] || 0;
    html += `<tr style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px; font-weight:bold;">${s.nama}</td>
      <td style="padding:8px;">${s.kelasNama}</td>
      <td style="padding:8px; color:${s.status === 'Aktif' ? 'green' : 'red'};">${s.status}</td>
      <td style="padding:8px;">Rp ${formatCurrency(balance)}</td>
      <td style="padding:8px;">
        <button onclick="viewStudent('${s.id}')" style="margin:2px; padding:4px 8px; cursor:pointer;">👁️ Lihat</button>
      </td>
    </tr>`;
  });

  html += '</table></div>';
  html += '<div id="detail-container"></div>';
  html += '<script>function viewStudent(id){google.script.run.withSuccessHandler(function(d){showStudentDetailDialog(d);}).withFailureHandler(function(e){alert("Error: " + e.message);}).getStudent(id);}<\/script>';
  html += '<script>function showStudentDetailDialog(d){'
    + 'var rows="";'
    + 'rows += "<tr><td>ID Murid</td><td>"+d.id+"</td></tr>";'
    + 'rows += "<tr><td>Nama</td><td>"+d.nama+"</td></tr>";'
    + 'rows += "<tr><td>Kelas</td><td>"+d.kelasNama+"</td></tr>";'
    + 'rows += "<tr><td>Tanggal Lahir</td><td>"+(d.tanggalLahir||"-")+"</td></tr>";'
    + 'rows += "<tr><td>Email Orang Tua</td><td>"+(d.email||"-")+"</td></tr>";'
    + 'rows += "<tr><td>No HP</td><td>"+(d.noHP||"-")+"</td></tr>";'
    + 'rows += "<tr><td>Status</td><td>"+d.status+"</td></tr>";'
    + 'rows += "<tr><td>ID Tabungan</td><td>"+(d.savingsId||"-")+"</td></tr>";'
    + 'rows += "<tr><td>Saldo Tabungan</td><td><b>Rp "+Number(d.saldo||0).toLocaleString("id-ID")+"</b></td></tr>";'
    + 'var txRows="";'
    + '(d.transactions||[]).slice(0,10).forEach(function(tx){'
    + '  txRows += "<tr><td>"+new Date(tx.tanggal).toLocaleDateString("id-ID")+"</td><td>"+tx.jenis+"</td><td style=\"text-align:right\">Rp "+Number(tx.jumlah).toLocaleString("id-ID")+"</td><td style=\"text-align:right\">Rp "+Number(tx.saldoSetelah).toLocaleString("id-ID")+"</td></tr>";'
    + '});'
    + 'var out="<div style=\"padding:15px; font-family:Arial\">"'
    + '+"<h2 style=\"margin-top:0\">"+d.nama+"</h2>"'
    + '+"<table style=\"border-collapse:collapse; width:100%; margin-bottom:15px\">"+rows+"</table>"'
    + '+"<h3>Riwayat Transaksi Terakhir</h3>"'
    + '+((d.transactions||[]).length? "<table style=\"border-collapse:collapse; width:100%\"><tr style=\"background:#eee\"><th>Tanggal</th><th>Jenis</th><th>Jumlah</th><th>Saldo</th></tr>"+txRows+"</table>" : "<p style=\"color:#999\">Belum ada transaksi</p>")'
    + '+"</div>";'
    + 'document.getElementById("detail-container").innerHTML = out;'
    + '}<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(800).setHeight(500), 'Daftar Murid');
}

function getStudent(id) {
  const students = getCachedStudents();
  const student = students.find(s => s.id === id);
  if (!student) return null;

  student.saldo = getBalanceMap()[student.savingsId] || 0;
  student.transactions = getTransactionsByStudent(id);
  return student;
}

function getBalance(savingsId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SAVINGS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == savingsId) {
      return data[i][5] || 0;
    }
  }
  return 0;
}

// ==================== CLASS MANAGEMENT ====================
function addClass(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  const classes = sheet.getDataRange().getValues();

  const classId = genId('K');

  const newRow = [
    classId,
    data.nama,
    data.guru || '',
    data.jadwal || '',
    data.kapasitas || 30,
    data.biaya || 0,
    data.status || 'Aktif'
  ];

  sheet.appendRow(newRow);
  invalidateDataCache();
  logActivity('Tambah Kelas', 'ID ' + classId, data.nama);
  return { success: true, classId, message: 'Kelas berhasil ditambahkan!' };
}

function getClasses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CLASSES_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(row => ({
    id: row[0],
    nama: row[1],
    guru: row[2],
    jadwal: row[3],
    kapasitas: row[4],
    biaya: row[5],
    status: row[6]
  }));
}

function showClassList() {
  const classes = getCachedClasses();
  let html = '<div style="padding: 15px;">';
  html += '<h2>Daftar Kelas</h2>';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 15px;">';
  html += '<tr style="background:#34A853; color:white;">';
  html += '<th>No</th><th>Nama Kelas</th><th>Guru</th><th>Jadwal</th><th>Kapasitas</th><th>Biaya</th><th>Aksi</th>';
  html += '</tr>';

  classes.forEach((c, i) => {
    const studentCount = getStudentCountInClass(c.id);
    html += `<tr style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px; font-weight:bold;">${c.nama}</td>
      <td style="padding:8px;">${c.guru || '-'}</td>
      <td style="padding:8px;">${c.jadwal || '-'}</td>
      <td style="padding:8px;">${studentCount}/${c.kapasitas}</td>
      <td style="padding:8px;">Rp ${formatCurrency(c.biaya)}</td>
      <td style="padding:8px;">
        <button onclick="deleteClass('${c.id}')" style="margin:2px; padding:4px 8px; cursor:pointer; color:red;">Hapus</button>
      </td>
    </tr>`;
  });

  html += '</table></div>';
  html += '<script>function deleteClass(id){if(confirm("Hapus kelas ini?")){google.script.run.deleteClass(id);google.script.host.close();}}<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(400), 'Daftar Kelas');
}

function showUserList() {
  const users = getCachedUsers();
  let html = '<div style="padding: 15px;">';
  html += '<h2>Daftar User (Akses Web App)</h2>';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 15px;">';
  html += '<tr style="background:#795548; color:white;">';
  html += '<th>No</th><th>Email</th><th>Nama</th><th>Peran</th><th>Status</th>';
  html += '</tr>';

  users.forEach((u, i) => {
    html += `<tr style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px;">${u.email}</td>
      <td style="padding:8px;">${u.nama || '-'}</td>
      <td style="padding:8px;">${u.peran}</td>
      <td style="padding:8px; color:${u.status === 'Aktif' ? 'green' : 'red'};">${u.status}</td>
    </tr>`;
  });

  html += '</table>';
  html += '<p style="margin-top:15px; color:#666; font-size:13px;">💡 Tambah/edit/hapus user lewat web app → menu 👥 Users (khusus Admin).</p>';
  html += '</div>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(400), 'Daftar User');
}

// ==================== USERS (LOGIN SHEET) ====================
/**
 * Ambil semua user dari sheet Users.
 */
function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0])
    .map(r => ({
      email: String(r[0]).trim(),
      nama: r[1],
      peran: r[2] || 'Guru',
      status: r[3] || 'Aktif',
      dibuat: r[4]
    }));
}

/**
 * Tambah user baru. Khusus Admin.
 */
function addUser(data) {
  if (!isAdminUser()) return { success: false, message: 'Hanya Admin yang bisa menambah user!' };

  const email = (data.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) {
    return { success: false, message: 'Email tidak valid!' };
  }

  const users = getCachedUsers();
  if (users.some(u => (u.email || '').toLowerCase() === email)) {
    return { success: false, message: 'Email sudah terdaftar!' };
  }

  const peran = (data.peran === 'Admin') ? 'Admin' : 'Guru';
  const status = (data.status === 'Nonaktif') ? 'Nonaktif' : 'Aktif';

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET)
    .appendRow([email, data.nama || '', peran, status, new Date()]);
  invalidateDataCache();

  logActivity('Tambah User', 'Peran ' + peran + ', status ' + status, email);
  return { success: true, message: `User ${email} berhasil ditambahkan!` };
}

/**
 * Update user (cari by email lama). Khusus Admin.
 * Jika email diubah, cukup tulis ulang kolom Email di baris yang sama.
 */
function updateUser(oldEmail, data) {
  if (!isAdminUser()) return { success: false, message: 'Hanya Admin yang bisa mengubah user!' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const target = String(oldEmail || '').trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === target) {
      if (data.email) values[i][0] = String(data.email).trim().toLowerCase();
      if (data.nama !== undefined) values[i][1] = data.nama;
      if (data.peran) values[i][2] = (data.peran === 'Admin') ? 'Admin' : 'Guru';
      if (data.status) values[i][3] = (data.status === 'Nonaktif') ? 'Nonaktif' : 'Aktif';
      sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
      invalidateDataCache();
      logActivity('Edit User', 'Data user diperbarui', oldEmail);
      return { success: true, message: 'User berhasil diperbarui!' };
    }
  }
  return { success: false, message: 'User tidak ditemukan!' };
}

/**
 * Hapus user by email. Khusus Admin.
 * Tolak hapus diri sendiri & tolak menghapus Admin terakhir
 * (mencegah tidak ada yang bisa mengelola user).
 */
function deleteUser(email) {
  if (!isAdminUser()) return { success: false, message: 'Hanya Admin yang bisa menghapus user!' };

  const me = getCurrentUser().trim().toLowerCase();
  const target = String(email || '').trim().toLowerCase();
  if (target === me) {
    return { success: false, message: 'Tidak bisa menghapus akun sendiri!' };
  }

  const users = getCachedUsers();
  const targetUser = users.find(u => (u.email || '').toLowerCase() === target);
  if (targetUser && targetUser.peran === 'Admin') {
    const activeAdmins = users.filter(u =>
      u.peran === 'Admin' && (u.status || '').toLowerCase() === 'aktif');
    if (activeAdmins.length <= 1) {
      return { success: false, message: 'Tidak bisa menghapus Admin terakhir!' };
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === target) {
      sheet.deleteRow(i + 1);
      invalidateDataCache();
      logActivity('Hapus User', 'Akses login dicabut', target);
      return { success: true, message: 'User berhasil dihapus!' };
    }
  }
  return { success: false, message: 'User tidak ditemukan!' };
}

// ==================== ACTIVITY LOG ====================
const ACTIVITY_LOG_SHEET = 'Log Aktivitas';

/**
 * Catat aktivitas user (siapa, apa, kapan). Dipanggil dari semua aksi tulis.
 * Gagal logging TIDAK boleh menggagalkan aksi utama — dibungkus try/catch.
 */
function logActivity(aksi, detail, target) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ACTIVITY_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(ACTIVITY_LOG_SHEET);
      sheet.getRange(1, 1, 1, 5).setValues([['Waktu', 'User', 'Aksi', 'Detail', 'Target']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#455A64').setFontColor('white');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      getCurrentUser() || '(tidak diketahui)',
      aksi,
      detail || '',
      target || ''
    ]);
  } catch (e) {
    // Logging gagal (mis. kuota) — abaikan, jangan ganggu aksi user
  }
}

/**
 * Ambil riwayat aktivitas (terbaru dulu, dibatasi `limit`).
 */
function getActivityLog(limit) {
  const n = limit || 100;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVITY_LOG_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0])
    .slice(-n)
    .reverse()
    .map(r => ({
      waktu: r[0],
      user: r[1],
      aksi: r[2],
      detail: r[3],
      target: r[4]
    }));
}

function showActivityLog() {
  const logs = getActivityLog(100);
  if (logs.length === 0) {
    SpreadsheetApp.getUi().alert('Belum ada aktivitas tercatat.');
    return;
  }

  let html = '<div style="padding: 15px;">';
  html += '<h2>Riwayat Aktivitas (100 terakhir)</h2>';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 15px;">';
  html += '<tr style="background:#455A64; color:white;">';
  html += '<th>Waktu</th><th>User</th><th>Aksi</th><th>Detail</th><th>Target</th>';
  html += '</tr>';

  logs.forEach(l => {
    html += `<tr style="border-bottom:1px solid #ddd;">
      <td style="padding:6px; font-size:12px;">${new Date(l.waktu).toLocaleString('id-ID')}</td>
      <td style="padding:6px; font-size:12px;">${l.user}</td>
      <td style="padding:6px; font-size:12px; font-weight:bold;">${l.aksi}</td>
      <td style="padding:6px; font-size:12px;">${l.detail || '-'}</td>
      <td style="padding:6px; font-size:12px;">${l.target || '-'}</td>
    </tr>`;
  });

  html += '</table></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(850).setHeight(500), 'Riwayat Aktivitas');
}

function getStudentCountInClass(classId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == classId && data[i][6] === 'Aktif') count++;
  }
  return count;
}

function deleteClass(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  const data = classesSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      classesSheet.deleteRow(i + 1);
      invalidateDataCache();
      logActivity('Hapus Kelas', 'Data kelas dihapus', id);
      return { success: true, message: 'Kelas berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Kelas tidak ditemukan!' };
}

function deleteStudent(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      invalidateDataCache();
      logActivity('Hapus Murid', 'Data murid dihapus', id);
      return { success: true, message: 'Murid berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Murid tidak ditemukan!' };
}

// ==================== ATTENDANCE ====================
function recordAttendance(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);

  // Check if already recorded today
  const today = new Date().toDateString();
  const attendanceData = sheet.getDataRange().getValues();
  for (let i = 1; i < attendanceData.length; i++) {
    if (attendanceData[i][1] == data.studentId &&
        new Date(attendanceData[i][2]).toDateString() === today) {
      return { success: false, message: 'Absensi sudah tercatat hari ini!' };
    }
  }

  const id = genId('ABS');
  const status = data.status || 'Hadir';
  const checkIn = data.checkIn || new Date();
  const checkOut = data.status === 'Sakit' || data.status === 'Izin' ? '' : '';

  sheet.appendRow([
    id,
    data.studentId,
    data.tanggal || new Date(),
    status,
    data.catatan || '',
    checkIn,
    checkOut
  ]);

  logActivity('Absensi', status + (data.catatan ? ' · ' + data.catatan : ''), data.studentId);

  // Notifikasi email/WA TIDAK dikirim di sini (asinkron via antrean) —
  // simpan absensi selesai < 1 detik; notifikasi menyusul lewat trigger.
  if (data.notifyParent || data.notifyWhatsApp) {
    queueNotification('attendance', {
      studentId: data.studentId,
      status: status,
      catatan: data.catatan || '',
      notifyParent: !!data.notifyParent,
      notifyWhatsApp: !!data.notifyWhatsApp
    });
    return { success: true, message: 'Absensi berhasil dicatat! Notifikasi sedang dikirim ke orang tua...' };
  }

  return { success: true, message: 'Absensi berhasil dicatat!' };
}

function getAttendanceToday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const today = new Date().toDateString();

  return data.slice(1).filter(row =>
    new Date(row[2]).toDateString() === today
  ).map(row => ({
    id: row[0],
    studentId: row[1],
    tanggal: row[2],
    status: row[3],
    catatan: row[4]
  }));
}

function showAttendance() {
  const students = getCachedStudents();
  const today = new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

  let html = '<div style="padding: 15px;">';
  html += `<h2>Absensi - ${today}</h2>`;
  html += '<div id="attendance-form">';
  html += '<label style="display:block; margin:10px 0 5px;">Pilih Murid:</label>';
  html += '<select id="student-select" style="width:100%; padding:8px; margin-bottom:10px;">';
  students.filter(s => s.status === 'Aktif').forEach(s => {
    html += `<option value="${s.id}">${s.nama} (${s.kelasNama})</option>`;
  });
  html += '</select>';
  html += '<label style="display:block; margin:10px 0 5px;">Status:</label>';
  html += '<select id="status-select" style="width:100%; padding:8px; margin-bottom:10px;">';
  html += '<option value="Hadir">Hadir</option>';
  html += '<option value="Sakit">Sakit</option>';
  html += '<option value="Izin">Izin</option>';
  html += '<option value="Alpha">Alpha</option>';
  html += '</select>';
  html += '<label style="display:block; margin:10px 0 5px;">Catatan (opsional):</label>';
  html += '<textarea id="catatan" rows="2" style="width:100%; padding:8px; margin-bottom:10px;"></textarea>';
  html += '<button onclick="submitAttendance()" style="padding:10px 20px; background:#4285F4; color:white; border:none; cursor:pointer; font-size:14px;">✅ Simpan Absensi</button>';
  html += '</div>';
  html += '<div id="attendance-result"></div>';
  html += '</div>';

  html += '<script>';
  html += 'function submitAttendance(){';
  html += '  const studentId = document.getElementById("student-select").value;';
  html += '  const status = document.getElementById("status-select").value;';
  html += '  const catatan = document.getElementById("catatan").value;';
  html += '  google.script.run.withSuccessHandler(function(res){';
  html += '    document.getElementById("attendance-result").innerHTML = "<p style=\'color:green; margin-top:10px;\'>'+res.message+ '</p>";';
  html += '  }).withFailureHandler(function(err){';
  html += '    document.getElementById("attendance-result").innerHTML = "<p style=\'color:red; margin-top:10px;\'>Error: "+err.message+"</p>";';
  html += '  }).recordAttendance({studentId:studentId, status:status, catatan:catatan});';
  html += '}';
  html += '<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(550), 'Absensi Hari Ini');
}

function showAttendanceHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('Belum ada data absensi!');
    return;
  }

  let html = '<div style="padding: 15px;">';
  html += '<h2>Riwayat Absensi</h2>';
  html += '<input type="text" id="search-absensi" placeholder="Cari nama murid..." ';
  html += 'style="padding:8px; width:300px; margin-bottom:10px;" onkeyup="filterAttendance()">';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 10px;" id="attendance-table">';
  html += '<tr style="background:#FBBC05;">';
  html += '<th>No</th><th>Nama Murid</th><th>Tanggal</th><th>Status</th><th>Catatan</th></tr>';

  data.slice(1).reverse().forEach((row, i) => {
    const nama = studentMap[row[1]] || '-';
    html += `<tr class="attendance-row" style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px;">${nama}</td>
      <td style="padding:8px;">${new Date(row[2]).toLocaleDateString('id-ID')}</td>
      <td style="padding:8px; color:${getStatusColor(row[3])};">${row[3]}</td>
      <td style="padding:8px;">${row[4] || '-'}</td>
    </tr>`;
  });

  html += '</table></div>';
  html += '<script>function getStatusColor(s){if(s==="Hadir")return "green";if(s==="Sakit"||s==="Izin")return "orange";return "red";}<\/script>';
  html += '<script>function filterAttendance(){const q=document.getElementById("search-absensi").value.toLowerCase();document.querySelectorAll(".attendance-row").forEach(row=>{const text=row.textContent.toLowerCase();row.style.display=text.includes(q)?"":"none";});}<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500), 'Riwayat Absensi');
}

// ==================== PROGRESS ====================
function addProgress(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PROGRESS_SHEET);

  const id = genId('PRG');

  sheet.appendRow([
    id,
    data.studentId,
    data.mapel || '',
    data.topik || '',
    data.nilai || 0,
    data.deskripsi || '',
    new Date(),
    data.guru || ''
  ]);

  logActivity('Tambah Progres', (data.mapel || '-') + ' · Nilai ' + (data.nilai || 0), data.studentId);
  return { success: true, message: 'Progres berhasil ditambahkan!' };
}

function getProgressByStudent(studentId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PROGRESS_SHEET);
  const data = sheet.getDataRange().getValues();

  return data.slice(1).filter(row => row[1] == studentId)
    .map(row => ({
      id: row[0],
      studentId: row[1],
      mapel: row[2],
      topik: row[3],
      nilai: row[4],
      deskripsi: row[5],
      tanggal: row[6],
      guru: row[7]
    })).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

function showProgressHistory() {
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PROGRESS_SHEET);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('Belum ada data progres!');
    return;
  }

  let html = '<div style="padding: 15px;">';
  html += '<h2>Riwayat Progres</h2>';
  html += '<input type="text" id="search-progres" placeholder="Cari nama murid..." ';
  html += 'style="padding:8px; width:300px; margin-bottom:10px;" onkeyup="filterProgress()">';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 10px;" id="progress-table">';
  html += '<tr style="background:#9C27B0; color:white;">';
  html += '<th>No</th><th>Nama Murid</th><th>Mata Pelajaran</th><th>Topik</th><th>Nilai</th><th>Deskripsi</th><th>Tanggal</th></tr>';

  data.slice(1).reverse().forEach((row, i) => {
    const nama = studentMap[row[1]] || '-';
    html += `<tr class="progress-row" style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px;">${nama}</td>
      <td style="padding:8px;">${row[2] || '-'}</td>
      <td style="padding:8px;">${row[3] || '-'}</td>
      <td style="padding:8px; font-weight:bold;">${row[4] || 0}</td>
      <td style="padding:8px;">${row[5] || '-'}</td>
      <td style="padding:8px;">${new Date(row[6]).toLocaleDateString('id-ID')}</td>
    </tr>`;
  });

  html += '</table></div>';
  html += '<script>function filterProgress(){const q=document.getElementById("search-progres").value.toLowerCase();document.querySelectorAll(".progress-row").forEach(row=>{const text=row.textContent.toLowerCase();row.style.display=text.includes(q)?"":"none";});}<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500), 'Riwayat Progres');
}

// ==================== SAVINGS / TABUNGAN ====================
function createSavingsAccount(id, studentId, studentName, initialBalance) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SAVINGS_SHEET);

  sheet.appendRow([
    id,
    studentId,
    initialBalance || 0,
    0,
    0,
    initialBalance || 0,
    new Date(),
    'Aktif'
  ]);

  return { success: true, savingsId: id };
}

function addTransaction(data) {
  // Lock: cegah saldo rusak jika 2 transaksi masuk bersamaan
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionsSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    const savingsSheet = ss.getSheetByName(CONFIG.SAVINGS_SHEET);

  // Validate savings account
  const savingsData = savingsSheet.getDataRange().getValues();
  let currentBalance = 0;
  let savingsRow = -1;

  for (let i = 1; i < savingsData.length; i++) {
    if (savingsData[i][0] == data.savingsId) {
      currentBalance = savingsData[i][5] || 0;
      savingsRow = i + 1;
      break;
    }
  }

  if (savingsRow === -1) {
    return { success: false, message: 'Rekening tabungan tidak ditemukan!' };
  }

  // Calculate new balance
  const amount = parseInt(data.jumlah) || 0;
  let newBalance = currentBalance;

  if (data.jenis === 'Setoran') {
    newBalance += amount;
  } else if (data.jenis === 'Penarikan') {
    if (amount > currentBalance) {
      return { success: false, message: 'Saldo tidak cukup!' };
    }
    newBalance -= amount;
  }

  // Create transaction ID
  const transactionId = genId('TRX');

  // Add transaction
  transactionsSheet.appendRow([
    transactionId,
    data.savingsId,
    data.studentId,
    data.jenis,
    amount,
    newBalance,
    data.catatan || '',
    new Date(),
    new Date()
  ]);

  // Update totals & saldo di sheet Tabungan
  // Kolom: C=Saldo Awal, D=Total Setoran, E=Total Penarikan, F=Saldo Saat Ini
  // (Saldo Awal tidak boleh diubah oleh transaksi!)
  invalidateDataCache();
  if (data.jenis === 'Setoran') {
    const currentDeposit = Number(savingsData[savingsRow - 1][3]) || 0;
    savingsSheet.getRange(savingsRow, 4).setValue(currentDeposit + amount);
  } else {
    const currentWithdrawal = Number(savingsData[savingsRow - 1][4]) || 0;
    savingsSheet.getRange(savingsRow, 5).setValue(currentWithdrawal + amount);
  }
  savingsSheet.getRange(savingsRow, 6).setValue(newBalance);

  logActivity('Transaksi ' + data.jenis, 'Rp ' + formatCurrency(amount) + ' · Saldo baru Rp ' + formatCurrency(newBalance), data.studentId);

  // Notifikasi TIDAK dikirim di sini (asinkron via antrean) —
  // transaksi selesai < 1 detik; email/WA menyusul lewat trigger.
  if (data.notifyParent || data.notifyWhatsApp) {
    queueNotification('transaction', {
      studentId: data.studentId,
      jenis: data.jenis,
      jumlah: amount,
      saldo: newBalance,
      catatan: data.catatan || '',
      notifyParent: !!data.notifyParent,
      notifyWhatsApp: !!data.notifyWhatsApp
    });
    return {
      success: true,
      message: `${data.jenis} berhasil! Saldo saat ini: Rp ${formatCurrency(newBalance)}. Notifikasi sedang dikirim ke orang tua...`,
      newBalance
    };
  }

  return {
    success: true,
    message: `${data.jenis} berhasil! Saldo saat ini: Rp ${formatCurrency(newBalance)}`,
    newBalance
  };
  } finally {
    lock.releaseLock();
  }
}

// ==================== NOTIFICATION QUEUE (ASINKRON) ====================
// Email/WA butuh 1-3 detik per kirim. Agar input user tidak menunggu,
// notifikasi masuk antrean lalu diproses trigger tiap 1 menit.
// Trigger dibuat hanya saat ada antrean & auto-hapus saat kosong (hemat kuota).
const NOTIF_QKEY = 'notifq';

function queueNotification(type, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const q = cacheGet(NOTIF_QKEY) || [];
    q.push({ type: type, payload: payload });
    cachePut(NOTIF_QKEY, q);
  } finally {
    lock.releaseLock();
  }
  ensureNotifTrigger();
}

function ensureNotifTrigger() {
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'processNotificationQueue');
  if (!exists) {
    ScriptApp.newTrigger('processNotificationQueue').timeBased().everyMinutes(1).create();
  }
}

function deleteNotifTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processNotificationQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * Dipanggil trigger tiap 1 menit selama ada antrean.
 * Gagal kirim dicatat ke Log Laporan, tidak mengulang (hindari dobel-kirim).
 */
function processNotificationQueue() {
  const q = cacheGet(NOTIF_QKEY) || [];
  if (q.length === 0) {
    deleteNotifTrigger();
    return;
  }

  q.forEach(item => {
    try {
      sendNotificationByType(item.type, item.payload);
    } catch (e) {
      try { logReportRun('NOTIF GAGAL', item.type + ': ' + e.message); } catch (e2) {}
    }
  });

  CacheService.getScriptCache().remove(NOTIF_QKEY);
  deleteNotifTrigger();
}

function sendNotificationByType(type, p) {
  if (type === 'attendance') {
    if (p.notifyParent) sendAbsensiNotification(p.studentId, p.status, p.catatan);
    if (p.notifyWhatsApp) sendAbsensiWhatsApp(p.studentId, p.status, p.catatan);
  } else if (type === 'transaction') {
    if (p.notifyParent) sendTabunganNotification(p.studentId, p.jenis, p.jumlah, p.saldo, p.catatan);
    if (p.notifyWhatsApp) sendTabunganWhatsApp(p.studentId, p.jenis, p.jumlah, p.saldo, p.catatan);
  }
}

function getTransactionsByStudent(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const savingsSheet = ss.getSheetByName(CONFIG.SAVINGS_SHEET);
  const transactionsSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);

  // Get savings ID for student
  const savingsData = savingsSheet.getDataRange().getValues();
  let savingsId = null;

  for (let i = 1; i < savingsData.length; i++) {
    if (savingsData[i][1] == studentId) {
      savingsId = savingsData[i][0];
      break;
    }
  }

  if (!savingsId) return [];

  // Get transactions
  const transactionsData = transactionsSheet.getDataRange().getValues();
  return transactionsData.slice(1).filter(row => row[1] == savingsId)
    .map(row => ({
      id: row[0],
      savingsId: row[1],
      studentId: row[2],
      jenis: row[3],
      jumlah: row[4],
      saldoSetelah: row[5],
      catatan: row[6],
      tanggal: row[7],
      waktu: row[8]
    })).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

function showTransactionHistory() {
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('Belum ada transaksi!');
    return;
  }

  let html = '<div style="padding: 15px;">';
  html += '<h2>Riwayat Transaksi Tabungan</h2>';
  html += '<input type="text" id="search-transaksi" placeholder="Cari nama murid..." ';
  html += 'style="padding:8px; width:300px; margin-bottom:10px;" onkeyup="filterTransactions()">';
  html += '<table style="width:100%; border-collapse: collapse; margin-top: 10px;" id="transaction-table">';
  html += '<tr style="background:#FF6D01; color:white;">';
  html += '<th>No</th><th>Nama Murid</th><th>Jenis</th><th>Jumlah</th><th>Saldo Setelah</th><th>Catatan</th><th>Tanggal</th></tr>';

  data.slice(1).reverse().forEach((row, i) => {
    const nama = studentMap[row[2]] || '-';
    const jenisColor = row[3] === 'Setoran' ? 'green' : 'red';
    html += `<tr class="transaction-row" style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px;">${nama}</td>
      <td style="padding:8px; color:${jenisColor}; font-weight:bold;">${row[3]}</td>
      <td style="padding:8px; text-align:right;">+${formatCurrency(row[4])}</td>
      <td style="padding:8px; text-align:right;">${formatCurrency(row[5])}</td>
      <td style="padding:8px;">${row[6] || '-'}</td>
      <td style="padding:8px;">${new Date(row[7]).toLocaleDateString('id-ID')}</td>
    </tr>`;
  });

  html += '</table></div>';
  html += '<script>function formatCurrency(val){return Number(val).toLocaleString("id-ID");}<\/script>';
  html += '<script>function filterTransactions(){const q=document.getElementById("search-transaksi").value.toLowerCase();document.querySelectorAll(".transaction-row").forEach(row=>{const text=row.textContent.toLowerCase();row.style.display=text.includes(q)?"":"none";});}<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500), 'Riwayat Transaksi');
}

function showBalanceCheck() {
  const students = getCachedStudents();
  const balanceMap = getBalanceMap();
  let html = '<div style="padding: 15px;">';
  html += '<h2>Cek Saldo Tabungan</h2>';

  html += '<table style="width:100%; border-collapse: collapse; margin-top: 15px;">';
  html += '<tr style="background:#0F9D58; color:white;">';
  html += '<th>No</th><th>Nama Murid</th><th>Rekening</th><th>Saldo</th><th>Status</th></tr>';

  students.forEach((s, i) => {
    const balance = balanceMap[s.savingsId] || 0;
    html += `<tr style="border-bottom:1px solid #ddd;">
      <td style="padding:8px;">${i + 1}</td>
      <td style="padding:8px; font-weight:bold;">${s.nama}</td>
      <td style="padding:8px;">${s.savingsId}</td>
      <td style="padding:8px; text-align:right; font-size:14px;">💰 Rp ${formatCurrency(balance)}</td>
      <td style="padding:8px; color:${balance > 0 ? 'green' : 'gray'};">${balance > 0 ? 'Punya Saldo' : 'Belum Ada Setoran'}</td>
    </tr>`;
  });

  html += '</table></div>';

  html += '<script>function formatCurrency(val){return Number(val).toLocaleString("id-ID");}<\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(500).setHeight(400), 'Cek Saldo Tabungan');
}

// ==================== DASHBOARD ====================
function showDashboard() {
  const students = getCachedStudents();
  const classes = getCachedClasses();
  const todayAttendance = getAttendanceToday();
  const totalTransactions = getTotalTransactions();
  const totalDeposits = getTotalDeposits();
  const totalWithdrawals = getTotalWithdrawals();

  // Get recent activity
  const recentTransactions = getRecentTransactions(5);

  let html = '<div style="padding: 20px; font-family: Arial, sans-serif;">';
  html += '<h1 style="color:#1a73e8; margin-bottom:5px;">📚 LMS & Buku Tabungan</h1>';
  html += `<p style="color:#666; margin-top:0;">Dashboard - ${new Date().toLocaleDateString("id-ID", {weekday:"long", year:"numeric", month:"long", day:"numeric"})}</p>`;

  // Summary cards
  html += '<div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:15px; margin:20px 0;">';

  html += `<div style="background:#e3f2fd; padding:20px; border-radius:8px;">
    <div style="font-size:28px; color:#1976d2;">${students.length}</div>
    <div style="font-size:14px; color:#666;">Total Murid</div>
  </div>`;

  html += `<div style="background:#e8f5e9; padding:20px; border-radius:8px;">
    <div style="font-size:28px; color:#388e3c;">${classes.length}</div>
    <div style="font-size:14px; color:#666;">Total Kelas</div>
  </div>`;

  html += `<div style="background:#fff3e0; padding:20px; border-radius:8px;">
    <div style="font-size:28px; color:#f57c00;">${todayAttendance.length}</div>
    <div style="font-size:14px; color:#666;">Absensi Hari Ini</div>
  </div>`;

  html += `<div style="background:#fce4ec; padding:20px; border-radius:8px;">
    <div style="font-size:28px; color:#c2185b;">${totalTransactions}</div>
    <div style="font-size:14px; color:#666;">Total Transaksi</div>
  </div>`;

  html += '</div>';

  // Quick actions
  html += '<h3 style="color:#333; margin-top:20px;">⚡ Aksi Cepat</h3>';
  html += '<div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:10px;">';
  html += '<button onclick="showAttendance()" style="padding:12px; background:#4285F4; color:white; border:none; cursor:pointer; border-radius:5px;">📝 Absensi Hari Ini</button>';
  html += '<button onclick="openWebApp()" style="padding:12px; background:#0F9D58; color:white; border:none; cursor:pointer; border-radius:5px;">💰 Transaksi Tabungan</button>';
  html += '<button onclick="showStudentList()" style="padding:12px; background:#9C27B0; color:white; border:none; cursor:pointer; border-radius:5px;">📋 Daftar Murid</button>';
  html += '</div>';

  // Tabungan summary
  html += '<h3 style="color:#333; margin-top:20px;">💰 Ringkasan Tabungan</h3>';
  html += '<div style="background:#f5f5f5; padding:15px; border-radius:8px;">';
  html += `<div style="display:flex; justify-content:space-between; margin:10px 0;">
    <span>Total Setoran:</span>
    <span style="color:green; font-weight:bold;">+ Rp ${formatCurrency(totalDeposits)}</span>
  </div>`;
  html += `<div style="display:flex; justify-content:space-between; margin:10px 0;">
    <span>Total Penarikan:</span>
    <span style="color:red; font-weight:bold;">- Rp ${formatCurrency(totalWithdrawals)}</span>
  </div>`;
  html += `<div style="display:flex; justify-content:space-between; margin:10px 0; padding-top:10px; border-top:1px solid #ddd;">
    <span style="font-size:16px;">Total Saldo Semua Murid:</span>
    <span style="font-size:16px; color:#1976d2; font-weight:bold;">Rp ${formatCurrency(totalDeposits - totalWithdrawals)}</span>
  </div>`;
  html += '</div>';

  // Recent transactions
  html += '<h3 style="color:#333; margin-top:20px;">🕐 Transaksi Terbaru</h3>';
  if (recentTransactions.length > 0) {
    html += '<table style="width:100%; border-collapse: collapse;">';
    html += '<tr style="background:#FF6D01; color:white;">';
    html += '<th>Nama</th><th>Jenis</th><th>Jumlah</th><th>Waktu</th></tr>';
    recentTransactions.forEach(tx => {
      const warna = tx.jenis === 'Setoran' ? 'green' : 'red';
      const simbol = tx.jenis === 'Setoran' ? '+' : '-';
      html += `<tr style="border-bottom:1px solid #ddd;">
        <td style="padding:8px;">${tx.nama}</td>
        <td style="padding:8px; color:${warna}; font-weight:bold;">${tx.jenis}</td>
        <td style="padding:8px; text-align:right;">${simbol} Rp ${formatCurrency(tx.jumlah)}</td>
        <td style="padding:8px; font-size:12px;">${tx.waktu}</td>
      </tr>`;
    });
    html += '</table>';
  } else {
    html += '<p style="color:#999; font-style:italic;">Belum ada transaksi</p>';
  }

  html += '</div>';

  html += '<script>function showAttendance(){google.script.run.showAttendance();google.script.host.close();}<\/script>';
  html += '<script>function showStudentList(){google.script.run.showStudentList();google.script.host.close();}<\/script>';
  html += '<script>function openWebApp(){google.script.run.withSuccessHandler(function(u){window.open(u, "_blank");}).getWebAppUrl();}<\\/script>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(800).setHeight(600), 'Dashboard LMS & Tabungan');
}

function getTotalTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  return data.length - 1;
}

function getTotalDeposits() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Setoran') {
      total += data[i][4] || 0;
    }
  }
  return total;
}

function getTotalWithdrawals() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Penarikan') {
      total += data[i][4] || 0;
    }
  }
  return total;
}

function getRecentTransactions(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transactionsSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const studentsSheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);

  const transactions = transactionsSheet.getDataRange().getValues();
  const students = studentsSheet.getDataRange().getValues();
  const studentMap = {};
  students.slice(1).forEach(s => studentMap[s[0]] = s[1]);

  // Sort by waktu terbaru dulu (bukan urutan baris di sheet), lalu ambil `limit`
  return transactions.slice(1)
    .sort((a, b) => new Date(b[8]) - new Date(a[8]))
    .slice(0, limit)
    .reverse()
    .map(row => ({
    nama: studentMap[row[2]] || '-',
    jenis: row[3],
    jumlah: row[4],
    waktu: new Date(row[8]).toLocaleTimeString('id-ID', {hour:'2-digit', minute:'2-digit'})
  }));
}

// ==================== HELP ====================
function showHelp() {
  const helpText = `
📚 PANDUAN LMS & Buku Tabungan

1. SETUP AWAL
   - Pilih "Setup Awal" untuk membuat semua sheet
   - Isi Settings dengan nama sekolah

2. MANAJEMEN MURID
   - Tambah Kelas terlebih dahulu
   - Tambah Murid dan pilih kelas
   - Setiap murid otomatis dapat rekening tabungan

3. ABSENSI
   - Pilih murid dan status (Hadir/Sakit/Izin/Alpha)
   - Absensi otomatis tercatat hari ini

4. PROGRES
   - Input nilai dan topik pembelajaran
   - Riwayat tersimpan per murid

5. TABUNGAN
   - Setoran: Menambah saldo
   - Penarikan: Mengurangi saldo (cek saldo cukup)
   - Saldo otomatis terupdate

⚙️ Tips:
   - Gunakan menu "LMS & Tabungan" di spreadsheet
   - Semua data terintegrasi dengan ID Murid
  `;

  SpreadsheetApp.getUi().alert(helpText);
}

// ==================== HELPERS ====================
/**
 * Generate ID unik. Date.now() saja bisa bentrok jika 2 aksi terjadi
 * di milidetik yang sama, jadi ditambah suffix acak.
 */
function genId(prefix) {
  const time = Date.now().toString(36).toUpperCase();
  const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return prefix + '-' + time + rand;
}

function formatCurrency(amount) {
  return Number(amount).toLocaleString('id-ID');
}

// ==================== CACHE LAYER ====================
// Data murid/kelas dibaca terus-menerus dari banyak fungsi. CacheService menyimpan
// hasil baca sheet selama 2 menit → panggilan berikutnya tidak perlu baca sheet lagi.
// Cache otomatis dibuang (invalidate) setiap kali data diubah (add/update/delete).
const CACHE_TTL = 120; // detik

function cacheGet(key) {
  const cached = CacheService.getScriptCache().get(key);
  return cached ? JSON.parse(cached) : null;
}

function cachePut(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), CACHE_TTL);
  } catch (e) {
    // Cache penuh atau value terlalu besar — abaikan, jangan sampai gagalkan request
  }
}

function invalidateDataCache() {
  CacheService.getScriptCache().removeAll([
    'students', 'classes', 'savings_accounts', 'users_v5',
    'savings_list', 'transactions_all'
  ]);
}

function getCachedStudents() {
  let students = cacheGet('students');
  if (!students) {
    students = getStudents();
    cachePut('students', students);
  }
  return students;
}

function getCachedClasses() {
  let classes = cacheGet('classes');
  if (!classes) {
    classes = getClasses();
    cachePut('classes', classes);
  }
  return classes;
}

function getCachedUsers() {
  // TTL 5 menit (key khusus) — daftar user jarang berubah
  let users = CacheService.getScriptCache().get('users_v5');
  if (!users) {
    users = JSON.stringify(getUsers());
    try {
      CacheService.getScriptCache().put('users_v5', users, 300);
    } catch (e) { /* cache penuh — abaikan */ }
  }
  return JSON.parse(users);
}

/**
 * Ambil SEMUA saldo sekaligus (1x baca sheet Tabungan).
 * Menggantikan pola N+1: getBalance() per murid = baca sheet berulang-ulang.
 */
function getBalanceMap() {
  let balances = cacheGet('savings_accounts');
  if (!balances) {
    const data = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.SAVINGS_SHEET).getDataRange().getValues();
    balances = {};
    data.slice(1).forEach(row => {
      if (row[0]) balances[row[0]] = Number(row[5]) || 0;
    });
    cachePut('savings_accounts', balances);
  }
  return balances;
}

// ==================== SCHEDULED MONTHLY REPORTS ====================
const REPORTS_FOLDER = 'LMS & Tabungan - Arsip Laporan';

/**
 * Install trigger bulanan. Trigger berjalan setiap hari tanggal 1 jam 8 pagi,
 * lalu fungsi sendMonthlyReports otomatis membuat laporan bulan SEBELUMNYA.
 */
function installMonthlyReportTrigger() {
  // Hapus trigger lama dulu biar tidak dobel
  removeMonthlyReportTrigger();
  
  ScriptApp.newTrigger('sendMonthlyReports')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();
  
  return { 
    success: true, 
    message: '✅ Trigger bulanan terpasang! Laporan otomatis dikirim setiap tanggal 1 jam 08:00 (waktu zona spreadsheet).' 
  };
}

function removeMonthlyReportTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyReports') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return { success: true, message: `🗑️ ${removed} trigger bulanan dihapus.` };
}

function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const monthly = triggers.find(t => t.getHandlerFunction() === 'sendMonthlyReports');
  
  if (monthly) {
    return { 
      active: true, 
      message: '✅ Trigger aktif — laporan dikirim otomatis setiap tanggal 1, jam 08:00.' 
    };
  }
  return { 
    active: false, 
    message: '⏸️ Tidak ada trigger aktif. Pasang dulu untuk laporan otomatis.' 
  };
}

/**
 * Fungsi utama yang dipanggil trigger. Membuat laporan untuk bulan sebelumnya
 * dan mengirimnya ke semua penerima di Settings.
 */
function sendMonthlyReports() {
  const settings = getSettings();
  const emails = (settings['Email Penerima Laporan'] || '')
    .split(',')
    .map(e => e.trim())
    .filter(e => e);
  
  if (emails.length === 0) {
    logReportRun('DILEWATI', 'Tidak ada email penerima di Settings (kolom "Email Penerima Laporan").');
    return { success: false, message: 'Tidak ada email penerima laporan di Settings!' };
  }
  
  // Rentang bulan lalu
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 1, 1, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0); // eksklusif
  const monthLabel = start.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
  
  try {
    // 1. Generate konten laporan
    const savingsHtml = generateSavingsPDFContentRange(getSavingsAccounts(), getTransactionsInRange(start, end), monthLabel);
    const attendanceHtml = generateAttendancePDFContentRange(getAttendanceInRange(start, end), monthLabel);
    
    // 2. Convert ke PDF via Google Docs
    const pdfFiles = [];
    pdfFiles.push(createPdfFromHtml(savingsHtml, `Laporan Tabungan - ${monthLabel}.pdf`));
    pdfFiles.push(createPdfFromHtml(attendanceHtml, `Laporan Absensi - ${monthLabel}.pdf`));
    
    // 3. Arsipkan ke Drive
    const folder = getOrCreateReportsFolder();
    pdfFiles.forEach(f => folder.addFile(f));
    
    // 4. Kirim email dengan PDF terlampir
    const schoolName = settings['Nama Sekolah'] || 'LMS & Tabungan';
    const attachments = pdfFiles.map(f => f.getAs('application/pdf'));
    
    MailApp.sendEmail({
      to: emails.join(','),
      subject: `📊 Laporan Bulanan ${schoolName} - ${monthLabel}`,
      htmlBody: generateMonthlyReportEmailBody(monthLabel, pdfFiles, schoolName),
      attachments: attachments,
      name: schoolName
    });
    
    logReportRun('BERHASIL', `Laporan ${monthLabel} dikirim ke ${emails.join(', ')}.`);
    return { success: true, message: `✅ Laporan ${monthLabel} berhasil dikirim ke ${emails.length} penerima!` };
    
  } catch (e) {
    logReportRun('GAGAL', e.message);
    return { success: false, message: 'Gagal: ' + e.message };
  }
}

/**
 * Test run manual — pakai rentang bulan ini (biar ada datanya untuk testing).
 */
function testMonthlyReportNow() {
  const settings = getSettings();
  const emails = (settings['Email Penerima Laporan'] || '')
    .split(',').map(e => e.trim()).filter(e => e);
  
  if (emails.length === 0) {
    return { success: false, message: 'Isi dulu "Email Penerima Laporan" di sheet Settings (baris baru).' };
  }
  
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
  const end = new Date(); // sampai sekarang
  const monthLabel = 'Test - ' + now.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
  
  try {
    const savingsHtml = generateSavingsPDFContentRange(getSavingsAccounts(), getTransactionsInRange(start, end), monthLabel);
    const attendanceHtml = generateAttendancePDFContentRange(getAttendanceInRange(start, end), monthLabel);
    
    const pdfFiles = [];
    pdfFiles.push(createPdfFromHtml(savingsHtml, `Laporan Tabungan - ${monthLabel}.pdf`));
    pdfFiles.push(createPdfFromHtml(attendanceHtml, `Laporan Absensi - ${monthLabel}.pdf`));
    
    const schoolName = settings['Nama Sekolah'] || 'LMS & Tabungan';
    MailApp.sendEmail({
      to: emails.join(','),
      subject: `🧪 [TEST] Laporan ${schoolName} - ${monthLabel}`,
      htmlBody: generateMonthlyReportEmailBody(monthLabel, pdfFiles, schoolName),
      attachments: pdfFiles.map(f => f.getAs('application/pdf')),
      name: schoolName
    });
    
    return { success: true, message: `✅ Email test terkirim ke ${emails.join(', ')}! Cek inbox.` };
  } catch (e) {
    return { success: false, message: 'Gagal: ' + e.message };
  }
}

// ---------- Helper: filter data per rentang tanggal ----------
function getTransactionsInRange(start, end) {
  return getAllTransactions().filter(t => {
    const d = new Date(t.tanggal);
    return d >= start && d < end;
  });
}

function getAttendanceInRange(start, end) {
  return getAttendanceHistory().filter(a => {
    const d = new Date(a.tanggal);
    return d >= start && d < end;
  });
}

// ---------- Data agregat utk halaman Statistik (1 panggilan) ----------
function getStatsData() {
  const now = new Date();
  const year = now.getFullYear();

  // 12 bulan terakhir (indeks 0 = 11 bulan lalu, indeks 11 = bulan ini)
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(year, now.getMonth() - i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth() });
  }
  const rangeStart = new Date(months[0].year, months[0].month, 1);
  const rangeEnd = new Date(year, now.getMonth() + 1, 1);

  const transactions = getTransactionsInRange(rangeStart, rangeEnd);
  const attendance = getAttendanceInRange(rangeStart, rangeEnd);
  const classes = getCachedClasses();
  const students = getCachedStudents();

  // Tren setor/tarik per bulan
  const monthly = months.map(m => {
    const ms = new Date(m.year, m.month, 1);
    const me = new Date(m.year, m.month + 1, 1);
    let deposit = 0, withdraw = 0;
    transactions.forEach(t => {
      const d = new Date(t.tanggal);
      if (d >= ms && d < me) {
        if (t.jenis === 'Setoran') deposit += Number(t.jumlah);
        else if (t.jenis === 'Penarikan') withdraw += Number(t.jumlah);
      }
    });
    return {
      label: ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'][m.month],
      deposit: deposit,
      withdraw: withdraw
    };
  });

  // Absensi per kelas (12 bulan terakhir)
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.kelasId);
  const byClass = classes.map(c => ({ id: c.id, nama: c.nama, Hadir: 0, Sakit: 0, Izin: 0, Alpha: 0 }));
  const classMap = {};
  byClass.forEach(c => classMap[c.id] = c);
  attendance.forEach(a => {
    const c = classMap[studentMap[a.studentId]];
    if (c && c[a.status] !== undefined) c[a.status]++;
  });

  return {
    monthly: monthly,
    byClass: byClass,
    totalAttendance: attendance.length,
    totalTransactions: transactions.length
  };
}

// ---------- Helper: HTML konten laporan dengan rentang ----------
function generateSavingsPDFContentRange(accounts, transactions, monthLabel) {
  const settings = getSettings();
  const totalBalance = accounts.reduce((sum, a) => sum + Number(a.balance), 0);
  const totalDeposits = transactions.filter(t => t.jenis === 'Setoran').reduce((sum, t) => sum + Number(t.jumlah), 0);
  const totalWithdrawals = transactions.filter(t => t.jenis === 'Penarikan').reduce((sum, t) => sum + Number(t.jumlah), 0);
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Laporan Tabungan</title>
  <style>
    body { font-family: 'Arial', sans-serif; padding: 40px; color: #333; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0F9D58; padding-bottom: 20px; }
    .header h1 { color: #0F9D58; margin-bottom: 5px; }
    .summary { display: flex; gap: 20px; margin: 20px 0; }
    .summary-card { background: #f0f7ff; padding: 15px; border-radius: 8px; flex: 1; }
    .summary-card .value { font-size: 1.5rem; font-weight: bold; color: #1a73e8; }
    .summary-card.success { background: #e8f5e9; }
    .summary-card.success .value { color: #0F9D58; }
    .summary-card.danger { background: #ffebee; }
    .summary-card.danger .value { color: #d93025; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th { background: #0F9D58; color: white; padding: 10px; text-align: left; }
    td { padding: 10px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background: #f5f5f5; }
    .text-right { text-align: right; }
    .fw-bold { font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <h1>💰 Laporan Tabungan Bulanan</h1>
    <p>${settings['Nama Sekolah'] || 'LMS & Tabungan'} | Periode: ${monthLabel}</p>
  </div>
  
  <div class="summary">
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Total Rekening</div>
      <div class="value">${accounts.length}</div>
    </div>
    <div class="summary-card success">
      <div style="font-size:0.9rem; color:#666;">Setoran Periode Ini</div>
      <div class="value">Rp ${totalDeposits.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-card danger">
      <div style="font-size:0.9rem; color:#666;">Penarikan Periode Ini</div>
      <div class="value">Rp ${totalWithdrawals.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Total Saldo Saat Ini</div>
      <div class="value">Rp ${totalBalance.toLocaleString('id-ID')}</div>
    </div>
  </div>
  
  <h2>Rekening per Murid</h2>
  <table>
    <thead>
      <tr><th>Nama Murid</th><th>ID Tabungan</th><th class="text-right">Saldo</th><th>Status</th></tr>
    </thead>
    <tbody>
      ${accounts.map(a => `
        <tr>
          <td>${a.studentName}</td>
          <td>${a.id}</td>
          <td class="text-right fw-bold">Rp ${Number(a.balance).toLocaleString('id-ID')}</td>
          <td>${a.status}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
  
  <h2 style="margin-top:30px;">Transaksi Periode ${monthLabel}</h2>
  ${transactions.length > 0 ? `
    <table>
      <thead>
        <tr><th>Tanggal</th><th>Nama</th><th>Jenis</th><th class="text-right">Jumlah</th><th class="text-right">Saldo</th><th>Catatan</th></tr>
      </thead>
      <tbody>
        ${transactions.map(t => `
          <tr>
            <td>${new Date(t.tanggal).toLocaleDateString('id-ID')}</td>
            <td>${t.nama}</td>
            <td>${t.jenis}</td>
            <td class="text-right" style="color:${t.jenis === 'Setoran' ? '#0F9D58' : '#d93025'};">${t.jenis === 'Setoran' ? '+' : '-'} Rp ${Number(t.jumlah).toLocaleString('id-ID')}</td>
            <td class="text-right">Rp ${Number(t.saldoSetelah).toLocaleString('id-ID')}</td>
            <td>${t.catatan || '-'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  ` : '<p style="color:#999;">Tidak ada transaksi pada periode ini.</p>'}
</body>
</html>`;
}

function generateAttendancePDFContentRange(attendanceData, monthLabel) {
  const settings = getSettings();
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Laporan Absensi</title>
  <style>
    body { font-family: 'Arial', sans-serif; padding: 40px; color: #333; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #FBBC05; padding-bottom: 20px; }
    .header h1 { color: #F4B400; margin-bottom: 5px; }
    table { width: 100%; border-collapse: collapse; }
    th { background: #FBBC05; color: #333; padding: 10px; text-align: left; }
    td { padding: 10px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background: #fff8e1; }
    .summary { display: flex; gap: 20px; margin: 20px 0; }
    .summary-card { background: #f8f9fa; padding: 15px; border-radius: 8px; flex: 1; }
    .summary-card .value { font-size: 1.5rem; font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <h1>📝 Laporan Absensi Bulanan</h1>
    <p>${settings['Nama Sekolah'] || 'LMS & Tabungan'} | Periode: ${monthLabel}</p>
  </div>
  
  <div class="summary">
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Total Hadir</div>
      <div class="value" style="color:#0F9D58;">${attendanceData.filter(a => a.status === 'Hadir').length}</div>
    </div>
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Sakit / Izin</div>
      <div class="value" style="color:#F4B400;">${attendanceData.filter(a => a.status === 'Sakit' || a.status === 'Izin').length}</div>
    </div>
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Alpha</div>
      <div class="value" style="color:#d93025;">${attendanceData.filter(a => a.status === 'Alpha').length}</div>
    </div>
  </div>
  
  ${attendanceData.length > 0 ? `
    <table>
      <thead>
        <tr><th>Tanggal</th><th>Nama Murid</th><th>Status</th><th>Catatan</th></tr>
      </thead>
      <tbody>
        ${attendanceData.map(a => `
          <tr>
            <td>${new Date(a.tanggal).toLocaleDateString('id-ID')}</td>
            <td>${a.nama}</td>
            <td>${a.status}</td>
            <td>${a.catatan || '-'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  ` : '<p style="color:#999;">Tidak ada absensi pada periode ini.</p>'}
</body>
</html>`;
}

// ---------- Helper: PDF & Drive ----------
function createPdfFromHtml(htmlContent, filename) {
  // Trik standar: buat Google Docs sementara dari HTML, lalu export sebagai PDF
  const resourceId = Utilities.base64Encode(filename + Date.now()).replace(/[^a-zA-Z0-9]/g, '').substring(0, 20);
  const doc = Document.create(filename.replace(/\.pdf$/, ''));
  const body = doc.getBody();
  
  // Google Docs tidak render CSS penuh — sisipkan konten dasar (tabel tetap terbaca)
  const text = htmlContent
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<h1[^>]*>/gi, '\n=== ')
    .replace(/<\/h1>/gi, ' ===\n')
    .replace(/<h2[^>]*>/gi, '\n--- ')
    .replace(/<\/h2>/gi, ' ---\n')
    .replace(/<tr[^>]*>/gi, '\n')
    .replace(/<\/(th|td)>/gi, ' | ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\n{3,}/g, '\n\n');
  
  body.setText(text);
  doc.saveAndClose();
  
  // Export doc sebagai PDF
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf').copyBlob();
  pdfBlob.setName(filename);
  DriveApp.getFileById(doc.getId()).setTrashed(true); // hapus doc sementara
  
  return DriveApp.createFile(pdfBlob);
}

function getOrCreateReportsFolder() {
  const it = DriveApp.getFoldersByName(REPORTS_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(REPORTS_FOLDER);
}

function logReportRun(status, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Log Laporan');
  if (!sheet) {
    sheet = ss.insertSheet('Log Laporan');
    sheet.appendRow(['Waktu', 'Status', 'Detail']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#607D8B').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), status, detail]);
}

function generateMonthlyReportEmailBody(monthLabel, pdfFiles, schoolName) {
  const fileList = pdfFiles.map(f => `<li>${f.getName()} (${Math.round(f.getSize() / 1024)} KB)</li>`).join('');
  
  return `
    <div style="font-family:Arial,sans-serif; max-width:600px; margin:0 auto;">
      <div style="background:#1a73e8; color:white; padding:20px; text-align:center; border-radius:8px 8px 0 0;">
        <h2 style="margin:0;">📊 Laporan Bulanan</h2>
        <p style="margin:5px 0 0; opacity:0.9;">${schoolName} — ${monthLabel}</p>
      </div>
      <div style="padding:20px; background:#f8f9fa; border:1px solid #e0e0e0; border-top:none; border-radius:0 0 8px 8px;">
        <p>Halo,</p>
        <p>Berikut laporan bulanan <strong>${monthLabel}</strong> yang terlampir dalam email ini:</p>
        <ul>${fileList}</ul>
        <p>Laporan juga diarsipkan otomatis di Google Drive folder <strong>"${REPORTS_FOLDER}"</strong>.</p>
        <p style="color:#666; font-size:0.9rem;">Email otomatis dari sistem LMS & Tabungan. Jangan balas email ini.</p>
      </div>
    </div>
  `;
}

/**
 * Simpan email penerima laporan ke Settings.
 */
function saveReportEmails(emailsString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  
  // Cari baris "Email Penerima Laporan" atau tambah baru
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'Email Penerima Laporan') {
      sheet.getRange(i + 1, 2).setValue(emailsString);
      return { success: true, message: 'Email penerima disimpan!' };
    }
  }
  sheet.appendRow(['Email Penerima Laporan', emailsString]);
  return { success: true, message: 'Email penerima disimpan (baris baru)!' };
}

// ==================== WHATSAPP NOTIFICATION (FONNTE) ====================
// Daftar di https://fonnte.com → dashboard → dapatkan token → scan QR WhatsApp
// Token disimpan aman di Script Properties (bukan hardcode)

function getFonnteToken() {
  return PropertiesService.getScriptProperties().getProperty('FONNTE_TOKEN');
}

function setFonnteToken(token) {
  PropertiesService.getScriptProperties().setProperty('FONNTE_TOKEN', token.trim());
  return { success: true, message: '✅ Token Fonnte tersimpan!' };
}

function isWhatsAppEnabled() {
  const settings = getSettings();
  return (settings['WhatsApp Aktif'] || 'Tidak') === 'Ya';
}

function setWhatsAppEnabled(enabled) {
  return setSettingValue('WhatsApp Aktif', enabled ? 'Ya' : 'Tidak');
}

function setSettingValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return { success: true, message: 'Pengaturan "' + key + '" disimpan!' };
    }
  }
  sheet.appendRow([key, value]);
  return { success: true, message: 'Pengaturan "' + key + '" disimpan (baris baru)!' };
}

/**
 * Kirim pesan WhatsApp via Fonnte API.
 * target: nomor HP format internasional tanpa + (contoh: 6281234567890)
 */
function sendWhatsAppMessage(target, message) {
  if (!isWhatsAppEnabled()) {
    return { success: false, message: 'Notifikasi WhatsApp tidak aktif (Settings: WhatsApp Aktif = Tidak).' };
  }
  
  const token = getFonnteToken();
  if (!token) {
    return { success: false, message: 'Token Fonnte belum diisi! Isi di menu Pengaturan WhatsApp.' };
  }
  
  // Normalisasi nomor: 08xx → 628xx, hapus karakter non-digit
  let normalized = String(target).replace(/\D/g, '');
  if (normalized.startsWith('0')) {
    normalized = '62' + normalized.substring(1);
  }
  
  const options = {
    method: 'post',
    payload: {
      'target': normalized,
      'message': message,
      'countryCode': '62'
    },
    headers: {
      'Authorization': token
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    if (code === 200) {
      return { success: true, message: 'WhatsApp terkirim ke ' + normalized };
    } else {
      return { success: false, message: 'Fonnte error ' + code + ': ' + body.substring(0, 100) };
    }
  } catch (e) {
    return { success: false, message: 'Gagal kirim WhatsApp: ' + e.message };
  }
}

/**
 * Kirim notifikasi WhatsApp ke orang tua murid.
 * Dipanggil dari recordAttendance / addTransaction saat data.notifyWhatsApp = true.
 */
function sendWhatsAppNotification(studentId, message) {
  const student = getStudent(studentId);
  if (!student || !student.noHP) {
    return { success: false, message: 'No HP orang tua tidak tersedia untuk murid ini!' };
  }
  
  return sendWhatsAppMessage(student.noHP, message);
}

function sendAbsensiWhatsApp(studentId, status, catatan) {
  const student = getStudent(studentId);
  if (!student) return { success: false, message: 'Murid tidak ditemukan!' };
  
  const settings = getSettings();
  const schoolName = settings['Nama Sekolah'] || 'LMS';
  const today = new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  
  const emoji = status === 'Hadir' ? '✅' : status === 'Alpha' ? '❌' : '📋';
  const message = 
    `${emoji} *Notifikasi Absensi ${schoolName}*

` +
    `Nama: *${student.nama}*
` +
    `Tanggal: ${today}
` +
    `Status: *${status}*
` +
    (catatan ? `Catatan: ${catatan}
` : '') +
    `\nTerima kasih.`;
  
  return sendWhatsAppNotification(studentId, message);
}

function sendTabunganWhatsApp(studentId, jenis, jumlah, newBalance, catatan) {
  const student = getStudent(studentId);
  if (!student) return { success: false, message: 'Murid tidak ditemukan!' };
  
  const settings = getSettings();
  const schoolName = settings['Nama Sekolah'] || 'LMS';
  const now = new Date().toLocaleString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  
  const emoji = jenis === 'Setoran' ? '💰' : '💸';
  const message = 
    `${emoji} *Notifikasi Tabungan ${schoolName}*

` +
    `Nama: *${student.nama}*
` +
    `Jenis: *${jenis}*
` +
    `Jumlah: Rp ${Number(jumlah).toLocaleString('id-ID')}
` +
    `Saldo Sekarang: *Rp ${Number(newBalance).toLocaleString('id-ID')}*
` +
    (catatan ? `Catatan: ${catatan}
` : '') +
    `Waktu: ${now}
` +
    `\nTerima kasih.`;
  
  return sendWhatsAppNotification(studentId, message);
}

/**
 * Test kirim WhatsApp ke nomor sendiri (dari web app / menu).
 */
function testWhatsApp(phoneNumber) {
  const message = 
    '🧪 *Test Notifikasi LMS & Tabungan*\n\n' +
    'Jika Anda menerima pesan ini, integrasi WhatsApp Fonnte sudah berfungsi! 🎉';
  
  return sendWhatsAppMessage(phoneNumber || '6281234567890', message);
}

// Web app endpoints
function getWhatsAppSettings() {
  return {
    enabled: isWhatsAppEnabled(),
    hasToken: !!getFonnteToken()
  };
}

function saveWhatsAppSettings(token, enabled) {
  const results = [];
  
  if (token && token.trim()) {
    results.push(setFonnteToken(token));
  }
  results.push(setWhatsAppEnabled(enabled));
  
  return { 
    success: true, 
    message: results.map(r => r.message).join(' ') 
  };
}

// ==================== SAVINGS TREND (GRAPH DATA) ====================
/**
 * Data tren saldo total semua murid per hari (default 30 hari terakhir)
 * untuk grafik di dashboard web app.
 */
function getSavingsTrend(days) {
  const numDays = parseInt(days) || 30;
  const transactions = getAllTransactions();
  
  // Hitung saldo akhir saat ini per rekening
  const accounts = getSavingsAccounts();
  let currentTotal = accounts.reduce((sum, a) => sum + Number(a.balance || 0), 0);
  
  // Build daily delta map (tanggal -> total delta hari itu)
  const dailyDelta = {};
  transactions.forEach(t => {
    const d = new Date(t.tanggal);
    d.setHours(0, 0, 0, 0);
    const key = d.getTime();
    const amount = Number(t.jumlah) || 0;
    dailyDelta[key] = (dailyDelta[key] || 0) + (t.jenis === 'Setoran' ? amount : -amount);
  });
  
  // Hitung mundur dari hari ini
  const trend = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 0; i < numDays; i++) {
    const day = new Date(today.getTime() - i * 24 * 60 * 60 * 1000);
    const dayKey = day.getTime();
    
    trend.unshift({
      date: day.toISOString().split('T')[0],
      label: day.toLocaleDateString('id-ID', { day: 'numeric', month: 'short' }),
      total: currentTotal
    });
    
    // Mundur: saldo hari ini dikurangi delta hari ini = saldo kemarin
    if (dailyDelta[dayKey]) {
      currentTotal -= dailyDelta[dayKey];
    }
  }
  
  return {
    days: numDays,
    data: trend,
    startTotal: trend.length > 0 ? trend[0].total : 0,
    endTotal: trend.length > 0 ? trend[trend.length - 1].total : 0,
    change: trend.length > 0 ? trend[trend.length - 1].total - trend[0].total : 0
  };
}

// ==================== REPORTS ENTRY POINTS ====================
// Web app endpoints untuk laporan terjadwal
function getReportSettings() {
  const status = getTriggerStatus();
  const settings = getSettings();
  return {
    triggerActive: status.active,
    triggerMessage: status.message,
    emails: settings['Email Penerima Laporan'] || ''
  };
}

function saveReportSettings(emailsString) {
  return saveReportEmails(emailsString);
}

function triggerTestReport() {
  return testMonthlyReportNow();
}

function triggerInstallMonthly() {
  return installMonthlyReportTrigger();
}

function triggerRemoveMonthly() {
  return removeMonthlyReportTrigger();
}

// ==================== AUTHENTICATION ====================
// Sumber kebenaran user: sheet "Users" (email terdaftar & Status=Aktif).
// ALLOWED_USERS hanya menjadi seed awal Admin saat setupSheets dijalankan —
// setelah itu kelola user lewat web app (menu 👥 Users), bukan lewat kode.
const ALLOWED_USERS = ['email@sekolah.com', 'admin@sekolah.com']; // ← seed awal, ganti dgn email kamu

/**
 * doGet — selalu menyajikan 'index' (aplikasi + overlay login dalam satu file).
 * Status otorisasi di-inject ke template: jika belum terotorisasi, index
 * menampilkan layar login di atas aplikasi.
 */
function doGet(request) {
  const loggedIn = isLoggedIn();

  const template = HtmlService.createTemplateFromFile('index');
  template.loggedIn = loggedIn;

  return template.evaluate()
    .setTitle('LMS & Tabungan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}

/**
 * Email user aktif dari sheet Users (cache 2 menit).
 * Fallback: jika sheet Users kosong/belum ada, pakai ALLOWED_USERS
 * agar sistem tidak pernah terkunci total.
 */
function getActiveUserEmails() {
  const users = getCachedUsers();
  const emails = (users || [])
    .filter(u => (u.status || '').toLowerCase() === 'aktif')
    .map(u => (u.email || '').trim().toLowerCase())
    .filter(e => e);
  if (emails.length === 0) {
    return ALLOWED_USERS.map(u => u.toLowerCase());
  }
  return emails;
}

function checkAuth(email) {
  if (!email) return false;
  return getActiveUserEmails().indexOf(email.trim().toLowerCase()) !== -1;
}

function isLoggedIn() {
  const userEmail = getCurrentUser();
  if (!userEmail) return false;
  return checkAuth(userEmail);
}

/**
 * Guard admin: hanya peran "Admin" yang boleh kelola user.
 * Jika belum ada satu pun Admin aktif di sheet, semua user terotorisasi
 * dianggap admin (mencegah kondisi tidak ada yang bisa mengelola).
 */
function isAdminUser() {
  const email = getCurrentUser();
  if (!email || !isLoggedIn()) return false;
  const users = getCachedUsers();
  const activeAdmins = (users || []).filter(u =>
    u.peran === 'Admin' && (u.status || '').toLowerCase() === 'aktif');
  if (activeAdmins.length === 0) return true;
  return activeAdmins.some(u => (u.email || '').trim().toLowerCase() === email.trim().toLowerCase());
}

function verifyLogin(email) {
  if (!email) {
    return { success: false, message: 'Email tidak boleh kosong!' };
  }
  
  if (checkAuth(email)) {
    return { success: true, message: 'Login berhasil!' };
  }
  
  return { 
    success: false, 
    message: 'Email "' + email + '" tidak terdaftar. Hubungi admin untuk akses.' 
  };
}

// ==================== AUTHENTICATION ====================

// ==================== MENU HELPERS (SPREADSHEET UI) ====================
function showTriggerStatus() {
  const status = getTriggerStatus();
  SpreadsheetApp.getUi().alert(status.message);
}

function showTestReport() {
  const result = testMonthlyReportNow();
  SpreadsheetApp.getUi().alert(result.message);
}

// ==================== PDF EXPORT ====================
function exportPDF(type, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let content = '';
  let filename = '';
  
  if (type === 'murid' && id) {
    const student = getStudent(id);
    if (!student) return { success: false, message: 'Murid tidak ditemukan!' };
    
    const transactions = getTransactionsByStudent(id);
    const attendance = getAttendanceByStudent(id);
    const progress = getProgressByStudent(id);
    
    content = generateStudentPDFContent(student, transactions, attendance, progress);
    filename = `Laporan_Murid_${student.nama.replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().split('T')[0]}.pdf`;
  } else if (type === 'absensi') {
    const attendanceData = getAttendanceHistory();
    content = generateAttendancePDFContent(attendanceData);
    filename = `Laporan_Absensi_${new Date().toISOString().split('T')[0]}.pdf`;
  } else if (type === 'tabungan') {
    const accounts = getSavingsAccounts();
    const transactions = getAllTransactions();
    content = generateSavingsPDFContent(accounts, transactions);
    filename = `Laporan_Tabungan_${new Date().toISOString().split('T')[0]}.pdf`;
  } else if (type === 'kelas' && id) {
    const classInfo = getClassById(id);
    if (!classInfo) return { success: false, message: 'Kelas tidak ditemukan!' };
    
    const students = getStudentsByClass(id);
    const attendance = getAttendanceByClass(id);
    content = generateClassPDFContent(classInfo, students, attendance);
    filename = `Laporan_Kelas_${classInfo.nama.replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().split('T')[0]}.pdf`;
  }
  
  if (content) {
    const htmlOutput = HtmlService.createHtmlOutput(content)
      .setTitle('Laporan')
      .setWidth(800)
      .setHeight(1000);
    
    return {
      success: true,
      message: 'Laporan berhasil dibuat!',
      filename: filename,
      content: content
    };
  }
  
  return { success: false, message: 'Gagal membuat laporan!' };
}

function generateStudentPDFContent(student, transactions, attendance, progress) {
  const settings = getSettings();
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Laporan Murid</title>
  <style>
    body { font-family: 'Arial', sans-serif; padding: 40px; color: #333; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #1a73e8; padding-bottom: 20px; }
    .header h1 { color: #1a73e8; margin-bottom: 5px; }
    .header p { color: #666; }
    .section { margin-bottom: 25px; }
    .section h2 { color: #333; border-left: 4px solid #1a73e8; padding-left: 10px; margin-bottom: 15px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th { background: #1a73e8; color: white; padding: 8px; text-align: left; }
    td { padding: 8px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background: #f5f5f5; }
    .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin: 15px 0; }
    .info-item { background: #f8f9fa; padding: 10px; border-radius: 5px; }
    .info-item strong { display: block; color: #666; font-size: 0.85rem; }
    .info-item span { font-size: 1rem; }
    @media print { body { padding: 20px; } }
  </style>
</head>
<body>
  <div class="header">
    <h1>📚 Laporan Murid</h1>
    <p>${settings.namaSekolah || 'LMS & Tabungan'} | ${new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
  </div>
  
  <div class="section">
    <h2>👤 Data Murid</h2>
    <div class="info-grid">
      <div class="info-item"><strong>ID Murid</strong><span>${student.id}</span></div>
      <div class="info-item"><strong>Nama</strong><span>${student.nama}</span></div>
      <div class="info-item"><strong>Kelas</strong><span>${student.kelasNama}</span></div>
      <div class="info-item"><strong>Tanggal Lahir</strong><span>${student.tanggalLahir || '-'}</span></div>
      <div class="info-item"><strong>Email Orang Tua</strong><span>${student.email || '-'}</span></div>
      <div class="info-item"><strong>No HP</strong><span>${student.noHP || '-'}</span></div>
      <div class="info-item"><strong>Status</strong><span>${student.status}</span></div>
      <div class="info-item"><strong>ID Tabungan</strong><span>${student.savingsId}</span></div>
    </div>
  </div>
  
  <div class="section">
    <h2>💰 Riwayat Tabungan</h2>
    ${transactions.length > 0 ? `
      <table>
        <thead>
          <tr><th>Tanggal</th><th>Jenis</th><th>Jumlah</th><th>Saldo</th><th>Catatan</th></tr>
        </thead>
        <tbody>
          ${transactions.slice(0, 20).map(tx => `
            <tr>
              <td>${new Date(tx.tanggal).toLocaleDateString('id-ID')}</td>
              <td>${tx.jenis}</td>
              <td style="text-align:right;">Rp ${Number(tx.jumlah).toLocaleString('id-ID')}</td>
              <td style="text-align:right;">Rp ${Number(tx.saldoSetelah).toLocaleString('id-ID')}</td>
              <td>${tx.catatan || '-'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    ` : '<p style="color:#999;">Belum ada transaksi</p>'}
  </div>
  
  <div class="section">
    <h2>📝 Riwayat Absensi</h2>
    ${attendance.length > 0 ? `
      <table>
        <thead>
          <tr><th>Tanggal</th><th>Status</th><th>Catatan</th></tr>
        </thead>
        <tbody>
          ${attendance.slice(0, 20).map(a => `
            <tr>
              <td>${new Date(a.tanggal).toLocaleDateString('id-ID')}</td>
              <td>${a.status}</td>
              <td>${a.catatan || '-'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    ` : '<p style="color:#999;">Belum ada absensi</p>'}
  </div>
  
  <div class="section">
    <h2>📈 Riwayat Progres</h2>
    ${progress.length > 0 ? `
      <table>
        <thead>
          <tr><th>Tanggal</th><th>Mata Pelajaran</th><th>Topik</th><th>Nilai</th><th>Deskripsi</th></tr>
        </thead>
        <tbody>
          ${progress.slice(0, 20).map(p => `
            <tr>
              <td>${new Date(p.tanggal).toLocaleDateString('id-ID')}</td>
              <td>${p.mapel}</td>
              <td>${p.topik || '-'}</td>
              <td>${p.nilai}</td>
              <td>${p.deskripsi || '-'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    ` : '<p style="color:#999;">Belum ada progres</p>'}
  </div>
</body>
</html>`;
}

function generateAttendancePDFContent(attendanceData) {
  const settings = getSettings();
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Laporan Absensi</title>
  <style>
    body { font-family: 'Arial', sans-serif; padding: 40px; color: #333; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #FBBC05; padding-bottom: 20px; }
    .header h1 { color: #FBBC05; margin-bottom: 5px; }
    table { width: 100%; border-collapse: collapse; }
    th { background: #FBBC05; color: #333; padding: 10px; text-align: left; }
    td { padding: 10px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background: #fff8e1; }
    .summary { display: flex; gap: 20px; margin: 20px 0; }
    .summary-card { background: #f8f9fa; padding: 15px; border-radius: 8px; flex: 1; }
    .summary-card .value { font-size: 1.5rem; font-weight: bold; }
    @media print { body { padding: 20px; } }
  </style>
</head>
<body>
  <div class="header">
    <h1>📝 Laporan Absensi</h1>
    <p>${settings.namaSekolah || 'LMS & Tabungan'} | ${new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
  </div>
  
  <div class="summary">
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Total Hadir</div>
      <div class="value" style="color:#0F9D58;">${attendanceData.filter(a => a.status === 'Hadir').length}</div>
    </div>
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Sakit / Izin</div>
      <div class="value" style="color:#F4B400;">${attendanceData.filter(a => a.status === 'Sakit' || a.status === 'Izin').length}</div>
    </div>
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Alpha</div>
      <div class="value" style="color:#d93025;">${attendanceData.filter(a => a.status === 'Alpha').length}</div>
    </div>
  </div>
  
  <table>
    <thead>
      <tr><th>Tanggal</th><th>Nama Murid</th><th>Status</th><th>Catatan</th></tr>
    </thead>
    <tbody>
      ${attendanceData.map(a => `
        <tr>
          <td>${new Date(a.tanggal).toLocaleDateString('id-ID')}</td>
          <td>${a.nama}</td>
          <td>${a.status}</td>
          <td>${a.catatan || '-'}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
</body>
</html>`;
}

function generateSavingsPDFContent(accounts, transactions) {
  const settings = getSettings();
  const totalBalance = accounts.reduce((sum, a) => sum + Number(a.balance), 0);
  const totalDeposits = transactions.filter(t => t.jenis === 'Setoran').reduce((sum, t) => sum + Number(t.jumlah), 0);
  const totalWithdrawals = transactions.filter(t => t.jenis === 'Penarikan').reduce((sum, t) => sum + Number(t.jumlah), 0);
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Laporan Tabungan</title>
  <style>
    body { font-family: 'Arial', sans-serif; padding: 40px; color: #333; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0F9D58; padding-bottom: 20px; }
    .header h1 { color: #0F9D58; margin-bottom: 5px; }
    .summary { display: flex; gap: 20px; margin: 20px 0; }
    .summary-card { background: #f0f7ff; padding: 15px; border-radius: 8px; flex: 1; }
    .summary-card .value { font-size: 1.5rem; font-weight: bold; color: #1a73e8; }
    .summary-card.success { background: #e8f5e9; }
    .summary-card.success .value { color: #0F9D58; }
    .summary-card.danger { background: #ffebee; }
    .summary-card.danger .value { color: #d93025; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th { background: #0F9D58; color: white; padding: 10px; text-align: left; }
    td { padding: 10px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background: #f5f5f5; }
    .text-right { text-align: right; }
    .fw-bold { font-weight: bold; }
    @media print { body { padding: 20px; } }
  </style>
</head>
<body>
  <div class="header">
    <h1>💰 Laporan Tabungan</h1>
    <p>${settings.namaSekolah || 'LMS & Tabungan'} | ${new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
  </div>
  
  <div class="summary">
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Total Rekening</div>
      <div class="value">${accounts.length}</div>
    </div>
    <div class="summary-card success">
      <div style="font-size:0.9rem; color:#666;">Total Setoran</div>
      <div class="value">Rp ${totalDeposits.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-card danger">
      <div style="font-size:0.9rem; color:#666;">Total Penarikan</div>
      <div class="value">Rp ${totalWithdrawals.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-card">
      <div style="font-size:0.9rem; color:#666;">Total Saldo</div>
      <div class="value">Rp ${totalBalance.toLocaleString('id-ID')}</div>
    </div>
  </div>
  
  <h2 style="margin-bottom:15px;">Rekening per Murid</h2>
  <table>
    <thead>
      <tr><th>Nama Murid</th><th>ID Tabungan</th><th class="text-right">Saldo</th><th>Status</th></tr>
    </thead>
    <tbody>
      ${accounts.map(a => `
        <tr>
          <td>${a.studentName}</td>
          <td>${a.id}</td>
          <td class="text-right fw-bold">Rp ${Number(a.balance).toLocaleString('id-ID')}</td>
          <td>${a.status}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
  
  <h2 style="margin:30px 0 15px;">Riwayat Transaksi (terakhir 50)</h2>
  <table>
    <thead>
      <tr><th>Tanggal</th><th>Nama Murid</th><th>Jenis</th><th class="text-right">Jumlah</th><th class="text-right">Saldo</th><th>Catatan</th></tr>
    </thead>
    <tbody>
      ${transactions.slice(-50).map(t => `
        <tr>
          <td>${new Date(t.tanggal).toLocaleDateString('id-ID')}</td>
          <td>${t.nama}</td>
          <td>${t.jenis}</td>
          <td class="text-right" style="color:${t.jenis === 'Setoran' ? '#0F9D58' : '#d93025'};">${t.jenis === 'Setoran' ? '+' : '-'} Rp ${Number(t.jumlah).toLocaleString('id-ID')}</td>
          <td class="text-right">Rp ${Number(t.saldoSetelah).toLocaleString('id-ID')}</td>
          <td>${t.catatan || '-'}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
</body>
</html>`;
}

function generateClassPDFContent(classInfo, students, attendance) {
  const settings = getSettings();
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Laporan Kelas</title>
  <style>
    body { font-family: 'Arial', sans-serif; padding: 40px; color: #333; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #34A853; padding-bottom: 20px; }
    .header h1 { color: #34A853; margin-bottom: 5px; }
    .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin: 20px 0; }
    .info-card { background: #f8f9fa; padding: 15px; border-radius: 8px; }
    .info-card h3 { margin-bottom: 10px; color: #34A853; }
    table { width: 100%; border-collapse: collapse; margin-top: 15px; }
    th { background: #34A853; color: white; padding: 10px; text-align: left; }
    td { padding: 10px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background: #f5f5f5; }
    @media print { body { padding: 20px; } }
  </style>
</head>
<body>
  <div class="header">
    <h1>🏫 Laporan Kelas</h1>
    <p>${settings.namaSekolah || 'LMS & Tabungan'} | ${new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
  </div>
  
  <div class="info-grid">
    <div class="info-card">
      <h3>Informasi Kelas</h3>
      <p><strong>Nama:</strong> ${classInfo.nama}</p>
      <p><strong>Guru:</strong> ${classInfo.guru || '-'}</p>
      <p><strong>Jadwal:</strong> ${classInfo.jadwal || '-'}</p>
      <p><strong>Kapasitas:</strong> ${classInfo.kapasitas} murid</p>
      <p><strong>Biaya/Bulan:</strong> ${classInfo.biaya > 0 ? 'Rp ' + classInfo.biaya.toLocaleString('id-ID') : 'Gratis'}</p>
      <p><strong>Status:</strong> ${classInfo.status}</p>
    </div>
    <div class="info-card">
      <h3>Statistik</h3>
      <p><strong>Total Murid:</strong> ${students.length}</p>
      <p><strong>Status Aktif:</strong> ${students.filter(s => s.status === 'Aktif').length}</p>
      <p><strong>Status Cuti:</strong> ${students.filter(s => s.status === 'Cuti').length}</p>
      <p><strong>Status Lulus:</strong> ${students.filter(s => s.status === 'Lulus').length}</p>
    </div>
  </div>
  
  <h2>Daftar Murid</h2>
  <table>
    <thead>
      <tr><th>No</th><th>Nama</th><th>No HP</th><th>Email</th><th>Status</th><th>Saldo Tabungan</th></tr>
    </thead>
    <tbody>
      ${(() => {
        const balanceMap = getBalanceMap();
        return students.map((s, i) => {
        const balance = balanceMap[s.savingsId] || 0;
        return `
          <tr>
            <td style="text-align:center;">${i + 1}</td>
            <td><strong>${s.nama}</strong></td>
            <td>${s.noHP || '-'}</td>
            <td>${s.email || '-'}</td>
            <td>${s.status}</td>
            <td style="text-align:right;">${balance > 0 ? 'Rp ' + balance.toLocaleString('id-ID') : '-'}</td>
          </tr>
        `;
        }).join('');
      })()}
    </tbody>
  </table>
  
  <h2 style="margin-top:30px;">Riwayat Absensi Kelas</h2>
  <table>
    <thead>
      <tr><th>Tanggal</th><th>Nama Murid</th><th>Status</th><th>Catatan</th></tr>
    </thead>
    <tbody>
      ${attendance.slice(0, 30).map(a => `
        <tr>
          <td>${new Date(a.tanggal).toLocaleDateString('id-ID')}</td>
          <td>${a.nama}</td>
          <td>${a.status}</td>
          <td>${a.catatan || '-'}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
</body>
</html>`;
}

// ==================== PDF EXPORT ====================

// ==================== NOTIFICATION ====================
function sendNotification(type, target) {
  const settings = getSettings();
  
  if (type === 'absensi') {
    if (!target.studentId || !target.studentEmail) {
      return { success: false, message: 'Data murid tidak lengkap!' };
    }
    
    const today = new Date().toLocaleDateString('id-ID', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
    
    const html = `
      <h2>📝 Notifikasi Absensi</h2>
      <p>Halo,</p>
      <p>Murid <strong>${target.studentName}</strong> telah melakukan absensi pada tanggal ${today}:</p>
      <table style="border-collapse:collapse; width:100%; margin:15px 0;">
        <tr style="background:#f0f0f0;">
          <th style="padding:8px; text-align:left;">Status</th>
          <th style="padding:8px; text-align:left;">Catatan</th>
          <th style="padding:8px; text-align:left;">Waktu</th>
        </tr>
        <tr>
          <td style="padding:8px;">${target.status}</td>
          <td style="padding:8px;">${target.catatan || '-'}</td>
          <td style="padding:8px;">${new Date().toLocaleTimeString('id-ID')}</td>
        </tr>
      </table>
      <p>Silakan cek aplikasi LMS untuk informasi lebih lanjut.</p>
      <hr>
      <p style="color:#999; font-size:0.9rem;">Notifikasi otomatis dari ${settings.namaSekolah || 'LMS & Tabungan System'}</p>
    `;
    
    try {
      MailApp.sendEmail({
        to: target.studentEmail,
        subject: `Absensi ${target.studentName} - ${today}`,
        htmlBody: html
      });
      return { success: true, message: 'Notifikasi absensi berhasil dikirim!' };
    } catch (e) {
      return { success: false, message: 'Gagal mengirim: ' + e.message };
    }
  } 
  else if (type === 'tabungan') {
    if (!target.studentId || !target.studentEmail) {
      return { success: false, message: 'Data murid tidak lengkap!' };
    }
    
    const html = `
      <h2>💰 Notifikasi Transaksi Tabungan</h2>
      <p>Halo,</p>
      <p>Terjadi transaksi pada rekening tabungan <strong>${target.studentName}</strong>:</p>
      <div style="background:#f0f7ff; padding:15px; border-radius:8px; margin:15px 0;">
        <p style="margin:5px 0;"><strong>Jenis:</strong> ${target.jenis}</p>
        <p style="margin:5px 0;"><strong>Jumlah:</strong> Rp ${Number(target.jumlah).toLocaleString('id-ID')}</p>
        <p style="margin:5px 0;"><strong>Saldo Sekarang:</strong> Rp ${Number(target.newBalance).toLocaleString('id-ID')}</p>
        <p style="margin:5px 0; color:#666;"><strong>Catatan:</strong> ${target.catatan || '-'}</p>
      </div>
      <p>Waktu transaksi: ${new Date().toLocaleString('id-ID')}</p>
      <hr>
      <p style="color:#999; font-size:0.9rem;">Notifikasi otomatis dari ${settings.namaSekolah || 'LMS & Tabungan System'}</p>
    `;
    
    try {
      MailApp.sendEmail({
        to: target.studentEmail,
        subject: `Transaksi Tabungan ${target.studentName}`,
        htmlBody: html
      });
      return { success: true, message: 'Notifikasi tabungan berhasil dikirim!' };
    } catch (e) {
      return { success: false, message: 'Gagal mengirim: ' + e.message };
    }
  }
  
  return { success: false, message: 'Tipe notifikasi tidak dikenal!' };
}

// ==================== NOTIFICATION ====================

/**
 * SEMUA data dashboard dalam SATU panggilan.
 * Menggantikan double/triple-hop (getDashboardStats → getSavingsAccounts → getStudents)
 * yang membuat web app lambat — tiap hop = 1 round-trip network + baca ulang semua sheet.
 * Di sini setiap sheet dibaca HANYA SEKALI.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Baca setiap sheet sekali saja
  const studentRows = ss.getSheetByName(CONFIG.STUDENTS_SHEET).getDataRange().getValues();
  const classRows = ss.getSheetByName(CONFIG.CLASSES_SHEET).getDataRange().getValues();
  const attRows = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET).getDataRange().getValues();
  const savingsRows = ss.getSheetByName(CONFIG.SAVINGS_SHEET).getDataRange().getValues();
  const txRows = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET).getDataRange().getValues();

  // Map kelas & murid
  const classMap = {};
  classRows.slice(1).forEach(r => { if (r[0]) classMap[r[0]] = r[1]; });
  const students = studentRows.slice(1).filter(r => r[0]).map(r => ({
    id: r[0],
    nama: r[1],
    kelasId: r[2],
    kelasNama: classMap[r[2]] || '-',
    tanggalLahir: r[3],
    email: r[4],
    noHP: r[5],
    status: r[6],
    tanggalDaftar: r[7],
    savingsId: r[8]
  }));
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);

  // Statistik
  const today = new Date().toDateString();
  const todayCount = attRows.slice(1).filter(r => r[2] && new Date(r[2]).toDateString() === today).length;
  let totalDeposits = 0, totalWithdrawals = 0;
  txRows.slice(1).forEach(r => {
    if (r[3] === 'Setoran') totalDeposits += Number(r[4]) || 0;
    if (r[3] === 'Penarikan') totalWithdrawals += Number(r[4]) || 0;
  });
  const classStats = classRows.slice(1).filter(r => r[0]).map(r => ({
    id: r[0],
    nama: r[1],
    total: students.filter(s => s.kelasId === r[0] && s.status === 'Aktif').length,
    capacity: r[4]
  }));

  const stats = {
    totalStudents: students.length,
    activeStudents: students.filter(s => s.status === 'Aktif').length,
    totalClasses: Math.max(0, classRows.length - 1),
    todayAttendance: todayCount,
    totalTransactions: Math.max(0, txRows.length - 1),
    totalDeposits,
    totalWithdrawals,
    netTotal: totalDeposits - totalWithdrawals,
    classStats,
    lastUpdated: new Date().toISOString()
  };

  // Rekening tabungan
  const savingsAccounts = savingsRows.slice(1).filter(r => r[0]).map(r => ({
    id: r[0],
    studentId: r[1],
    balance: r[5] || 0,
    status: r[7],
    studentName: studentMap[r[1]] || '-'
  }));

  // 5 transaksi terbaru (sort by waktu, bukan urutan baris)
  const recentTransactions = txRows.slice(1)
    .sort((a, b) => new Date(b[8]) - new Date(a[8]))
    .slice(0, 5)
    .map(r => ({
      nama: studentMap[r[2]] || '-',
      jenis: r[3],
      jumlah: r[4],
      waktu: new Date(r[8]).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })
    }));

  return {
    stats,
    students,          // untuk dropdown quick actions (absensi/transaksi)
    savingsAccounts,
    recentTransactions
  };
}

function getDashboardDataStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = getCachedStudents();
  const classes = getCachedClasses();
  const today = new Date().toDateString();
  
  // Count today attendance
  const attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const attData = attSheet.getDataRange().getValues();
  const todayCount = attData.slice(1).filter(r => new Date(r[2]).toDateString() === today).length;
  
  // Get transaction totals
  const txSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const txData = txSheet.getDataRange().getValues();
  let totalDeposits = 0, totalWithdrawals = 0;
  txData.slice(1).forEach(r => {
    if (r[3] === 'Setoran') totalDeposits += Number(r[4]) || 0;
    if (r[3] === 'Penarikan') totalWithdrawals += Number(r[4]) || 0;
  });
  
  // Get total students by class
  const classStats = classes.map(c => {
    const count = students.filter(s => s.kelasId === c.id && s.status === 'Aktif').length;
    return { id: c.id, nama: c.nama, total: count, capacity: c.kapasitas };
  });
  
  return {
    totalStudents: students.length,
    activeStudents: students.filter(s => s.status === 'Aktif').length,
    totalClasses: classes.length,
    todayAttendance: todayCount,
    totalTransactions: txData.length - 1,
    totalDeposits,
    totalWithdrawals,
    netTotal: totalDeposits - totalWithdrawals,
    classStats,
    lastUpdated: new Date().toISOString()
  };
}

function getSavingsAccounts() {
  // Cache 2 menit — dibuang otomatis setiap transaksi/murid berubah (invalidateDataCache)
  let accounts = cacheGet('savings_list');
  if (accounts) return accounts;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SAVINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);
  
  accounts = data.slice(1).map(row => ({
    id: row[0],
    studentId: row[1],
    balance: row[5] || 0,
    status: row[7],
    studentName: studentMap[row[1]] || '-'
  }));
  cachePut('savings_list', accounts);
  return accounts;
}


// ==================== HELPER FOR PDF/NOTIFICATION ====================
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SETTINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  const settings = {};
  data.slice(1).forEach(row => {
    settings[row[0]] = row[1];
  });
  return settings;
}

function getAttendanceByStudent(studentId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);
  
  return data.slice(1).filter(row => row[1] == studentId)
    .map(row => ({
      id: row[0],
      studentId: row[1],
      nama: studentMap[row[1]] || '-',
      tanggal: row[2],
      status: row[3],
      catatan: row[4]
    })).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

function getAttendanceByClass(classId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);
  
  const studentIdsInClass = students.filter(s => s.kelasId == classId).map(s => s.id);
  
  return data.slice(1).filter(row => studentIdsInClass.includes(row[1]))
    .map(row => ({
      id: row[0],
      studentId: row[1],
      nama: studentMap[row[1]] || '-',
      tanggal: row[2],
      status: row[3],
      catatan: row[4]
    })).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

function getAttendanceHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);
  
  return data.slice(1).map(row => ({
    id: row[0],
    studentId: row[1],
    nama: studentMap[row[1]] || '-',
    tanggal: row[2],
    status: row[3],
    catatan: row[4]
  })).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

function getAllTransactions() {
  // Cache 2 menit — dibuang otomatis oleh addTransaction (invalidateDataCache)
  let transactions = cacheGet('transactions_all');
  if (transactions) return transactions;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = getCachedStudents();
  const studentMap = {};
  students.forEach(s => studentMap[s.id] = s.nama);
  
  transactions = data.slice(1).map(row => ({
    id: row[0],
    savingsId: row[1],
    studentId: row[2],
    nama: studentMap[row[2]] || '-',
    jenis: row[3],
    jumlah: row[4],
    saldoSetelah: row[5],
    catatan: row[6],
    tanggal: row[7]
  })).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  cachePut('transactions_all', transactions);
  return transactions;
}

function getClassById(id) {
  const classes = getCachedClasses();
  return classes.find(c => c.id === id);
}

function getStudentsByClass(classId) {
  const students = getCachedStudents();
  return students.filter(s => s.kelasId == classId);
}

function downloadReport(type, id) {
  const result = exportPDF(type, id);
  if (result.success && result.content) {
    // Return HTML content untuk di-print sebagai PDF
    return { success: true, htmlContent: result.content, filename: result.filename };
  }
  return result;
}

function sendAbsensiNotification(studentId, status, catatan) {
  const student = getStudent(studentId);
  if (!student || !student.email) {
    return { success: false, message: 'Email orang tua tidak tersedia!' };
  }
  
  return sendNotification('absensi', {
    studentId: studentId,
    studentName: student.nama,
    studentEmail: student.email,
    status: status,
    catatan: catatan
  });
}

function sendTabunganNotification(studentId, jenis, jumlah, newBalance, catatan) {
  const student = getStudent(studentId);
  if (!student || !student.email) {
    return { success: false, message: 'Email orang tua tidak tersedia!' };
  }
  
  return sendNotification('tabungan', {
    studentId: studentId,
    studentName: student.nama,
    studentEmail: student.email,
    jenis: jenis,
    jumlah: jumlah,
    newBalance: newBalance,
    catatan: catatan
  });
}
