// --- KONFIGURASI GOOGLE SHEET ---
const SPREADSHEET_ID = "14qAHqpqPXRdmxXElAhg1MyLaG9yzLPmDGqHfP_ga_uA"; // Ganti dengan ID Google Sheet Anda
// Anda bisa mendapatkan ID dari URL Google Sheet Anda:
// https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_HERE/edit

// --- DOMAIN YANG DIIZINKAN UNTUK CORS ---
// Ganti dengan domain Vercel Anda yang sebenarnya
const ALLOWED_ORIGIN = "https://presensi-toptea.vercel.app";
// Jika Anda ingin mengizinkan dari semua domain (TIDAK DISARANKAN UNTUK PRODUKSI):
// const ALLOWED_ORIGIN = '*';

// --- NAMA SHEET (TAB) ---
const SHEET_NAMES = {
  USERS: "Users",
  EMPLOYEES: "Employees",
  STORES: "Stores",
  ATTENDANCE: "Attendance",
};

// --- FUNGSI UTAMA DOPOST (untuk permintaan POST) ---
function doPost(e) {
  if (e && e.parameter && e.parameter.action === "preflight") {
    return ContentService.createTextOutput("")
      .setMimeType(ContentService.MimeType.TEXT)
      .setHeaders({
        "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Max-Age": "3600",
      });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Tunggu hingga 30 detik untuk mendapatkan lock

    let result;
    let requestData = {};

    // Cek apakah e.postData.contents ada dan valid JSON
    if (e && e.postData && e.postData.contents) {
      try {
        const parsedData = JSON.parse(e.postData.contents);
        const action = parsedData.action;
        requestData = parsedData.data || {};

        switch (action) {
          case "login":
            result = login(requestData);
            break;
          case "getRealtimeData":
            result = getRealtimeData();
            break;
          case "getEmployees":
            result = getEmployees();
            break;
          case "addEmployee":
            result = addEmployee(requestData);
            break;
          case "deleteEmployee":
            result = deleteEmployee(requestData);
            break;
          case "getStores":
            result = getStores();
            break;
          case "addStore":
            result = addStore(requestData);
            break;
          case "deleteStore":
            result = deleteStore(requestData);
            break;
          case "clockIn":
            result = clockIn(requestData);
            break;
          case "clockOut":
            result = clockOut(requestData);
            break;
          case "markStatus":
            result = markStatus(requestData);
            break;
          case "getTodayStatus":
            result = getTodayStatus(requestData);
            break;
          case "getEmployeeHistory":
            result = getEmployeeHistory(requestData);
            break;
          case "getMonthlyReports":
            result = getMonthlyReports();
            break;
          default:
            result = { success: false, message: "Aksi tidak valid" };
        }
      } catch (jsonError) {
        result = {
          success: false,
          message: "Invalid JSON payload: " + jsonError.message,
        };
      }
    } else {
      result = {
        success: false,
        message: "No valid data received in POST request.",
      };
    }

    const output = ContentService.createTextOutput(
      JSON.stringify(result)
    ).setMimeType(ContentService.MimeType.JSON);

    // Set CORS headers
    output.setHeaders({
      "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    });

    return output;
  } catch (error) {
    // Tangani error dan pastikan juga mengirim header CORS
    const errorOutput = ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.message || error.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);
    errorOutput.setHeaders({
      "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    });
    return errorOutput;
  } finally {
    lock.releaseLock();
  }
}

// --- FUNGSI DOGET (untuk permintaan GET, termasuk OPTIONS preflight) ---
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "preflight") {
    return ContentService.createTextOutput("")
      .setMimeType(ContentService.MimeType.TEXT)
      .setHeaders({
        "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Max-Age": "3600",
      });
  }
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, message: "GET request not supported" })
  )
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({ "Access-Control-Allow-Origin": ALLOWED_ORIGIN });

  const output = ContentService.createTextOutput("") // Respons preflight biasanya kosong
    .setMimeType(ContentService.MimeType.TEXT) // Atau ContentService.MimeType.JSON
    .setHeaders({
      "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS", // Izinkan GET, POST, OPTIONS
      "Access-Control-Allow-Headers": "Content-Type", // Izinkan header Content-Type
      "Access-Control-Max-Age": 86400, // Opsional: cache preflight response selama 24 jam
    });
  return output;

  // Jika Anda memiliki GET request lain yang perlu ditangani, tambahkan di sini.
  // Misalnya:
  // if (e.parameter.action === 'someGetData') {
  //   // Lakukan sesuatu
  //   return ContentService.createTextOutput(JSON.stringify({ success: true, data: 'some data' }))
  //         .setMimeType(ContentService.MimeType.JSON)
  //         .setHeaders({'Access-Control-Allow-Origin': ALLOWED_ORIGIN});
  // }
  // return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'GET request not supported for this API.' }))
  //        .setMimeType(ContentService.MimeType.JSON)
  //        .setHeaders({'Access-Control-Allow-Origin': ALLOWED_ORIGIN});
}

// --- FUNGSI UTILITAS ---
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(
      `Sheet "${sheetName}" tidak ditemukan. Pastikan nama sheet sudah benar.`
    );
  }
  return sheet;
}

function getRowsData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length === 0 || values[0].length === 0) return []; // Handle empty sheet or no headers
  const headers = values[0];
  const data = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
    }
    data.push(row);
  }
  return data;
}

function appendRowToSheet(sheet, rowObject) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = [];
  for (let i = 0; i < headers.length; i++) {
    newRow.push(rowObject[headers[i]] || ""); // Handle undefined properties with empty string
  }
  sheet.appendRow(newRow);
}

function updateRowInSheet(sheet, rowIndex, rowObject) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updatedValues = [];
  // Perlu mengambil nilai yang ada di baris tersebut terlebih dahulu
  const currentRowValues = sheet
    .getRange(rowIndex + 1, 1, 1, headers.length)
    .getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    // Jika properti ada di rowObject, gunakan nilai dari rowObject
    // Jika tidak ada, gunakan nilai yang sudah ada di sheet
    updatedValues.push(
      rowObject[header] !== undefined ? rowObject[header] : currentRowValues[i]
    );
  }
  sheet
    .getRange(rowIndex + 1, 1, 1, updatedValues.length)
    .setValues([updatedValues]);
}

function deleteRowInSheet(sheet, rowIndex) {
  sheet.deleteRow(rowIndex + 1);
}

function getNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0 || sheet.getRange(1, 1).getValue() !== "id") {
    // If sheet is empty or no 'id' header
    return 1;
  }
  // Ambil hanya kolom 'id' dari baris kedua hingga akhir
  const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxId = 0;
  for (let i = 0; i < idColumn.length; i++) {
    const currentId = idColumn[i][0];
    if (typeof currentId === "number") {
      maxId = Math.max(maxId, currentId);
    }
  }
  return maxId + 1;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

// --- FUNGSI LOGIN ---
function login(requestData) {
  const { userID, password } = requestData;
  const usersSheet = getSheet(SHEET_NAMES.USERS);
  const users = getRowsData(usersSheet);

  const user = users.find(
    (u) => u.userID === userID && u.password === password
  );

  if (user) {
    return { success: true, data: user };
  } else {
    return { success: false, message: "User ID atau password salah." };
  }
}

// --- FUNGSI EMPLOYEES (ADMIN) ---
function getEmployees() {
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const employees = getRowsData(employeesSheet);
  return { success: true, data: employees };
}

function addEmployee(employeeData) {
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const usersSheet = getSheet(SHEET_NAMES.USERS);

  const newId = getNextId(employeesSheet);
  const newEmployee = {
    id: newId,
    name: employeeData.name,
    userID: employeeData.userID,
    password: employeeData.password,
    role: "employee", // Default role for new employees
    store: employeeData.store,
  };

  appendRowToSheet(employeesSheet, {
    id: newId,
    name: newEmployee.name,
    userID: newEmployee.userID,
    password: newEmployee.password,
    store: newEmployee.store,
  });
  appendRowToSheet(usersSheet, {
    id: getNextId(usersSheet),
    userID: newEmployee.userID,
    password: newEmployee.password,
    name: newEmployee.name,
    role: newEmployee.role,
    store: newEmployee.store,
  }); // Also add to Users sheet for login

  return {
    success: true,
    message: "Karyawan berhasil ditambahkan.",
    data: newEmployee,
  };
}

function deleteEmployee(requestData) {
  const { id } = requestData;
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const usersSheet = getSheet(SHEET_NAMES.USERS);

  const employees = getRowsData(employeesSheet);
  const indexToDelete = employees.findIndex((emp) => emp.id === id);

  if (indexToDelete > -1) {
    const employeeToDelete = employees[indexToDelete];
    deleteRowInSheet(employeesSheet, indexToDelete + 1); // +1 because getRowsData returns data from 2nd row

    // Also delete from Users sheet
    const users = getRowsData(usersSheet);
    const userIndexToDelete = users.findIndex(
      (u) => u.userID === employeeToDelete.userID
    );
    if (userIndexToDelete > -1) {
      deleteRowInSheet(usersSheet, userIndexToDelete + 1);
    }

    return { success: true, message: "Karyawan berhasil dihapus." };
  } else {
    return { success: false, message: "Karyawan tidak ditemukan." };
  }
}

// --- FUNGSI STORES (ADMIN) ---
function getStores() {
  const storesSheet = getSheet(SHEET_NAMES.STORES);
  const stores = getRowsData(storesSheet);
  return { success: true, data: stores };
}

function addStore(storeData) {
  const storesSheet = getSheet(SHEET_NAMES.STORES);
  const newId = getNextId(storesSheet);
  const newStore = {
    id: newId,
    name: storeData.name,
    startTime: storeData.startTime,
    endTime: storeData.endTime,
  };
  appendRowToSheet(storesSheet, newStore);
  return {
    success: true,
    message: "Toko berhasil ditambahkan.",
    data: newStore,
  };
}

function deleteStore(requestData) {
  const { id } = requestData;
  const storesSheet = getSheet(SHEET_NAMES.STORES);
  const stores = getRowsData(storesSheet);
  const indexToDelete = stores.findIndex((store) => store.id === id);

  if (indexToDelete > -1) {
    deleteRowInSheet(storesSheet, indexToDelete + 1); // +1 because getRowsData returns data from 2nd row
    return { success: true, message: "Toko berhasil dihapus." };
  } else {
    return { success: false, message: "Toko tidak ditemukan." };
  }
}

// --- FUNGSI ATTENDANCE (EMPLOYEE) ---
function clockIn(requestData) {
  const { employeeID } = requestData;
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const employees = getRowsData(employeesSheet);
  const employee = employees.find((emp) => emp.id === employeeID);

  if (!employee) {
    return { success: false, message: "Karyawan tidak ditemukan." };
  }

  const today = formatDate(new Date());
  const attendanceRecords = getRowsData(attendanceSheet);

  // Check if employee already clocked in today
  const existingRecordIndex = attendanceRecords.findIndex(
    (record) => record.employeeID === employeeID && record.date === today
  );

  if (existingRecordIndex > -1) {
    const record = attendanceRecords[existingRecordIndex];
    if (record.status === "present" && record.clockIn) {
      return { success: false, message: "Anda sudah absen masuk hari ini." };
    } else {
      // If status was Izin/Sakit/Belum Absen, update to present and set clockIn
      record.clockIn = formatTime(new Date());
      record.status = "present";
      record.reason = ""; // Clear reason if now present
      updateRowInSheet(attendanceSheet, existingRecordIndex, record); // Use existingRecordIndex directly for updateRowInSheet
      return { success: true, message: "Absen masuk dicatat." };
    }
  }

  const newId = getNextId(attendanceSheet);
  const newRecord = {
    id: newId,
    employeeID: employeeID,
    employeeName: employee.name,
    date: today,
    clockIn: formatTime(new Date()),
    clockOut: "",
    status: "present",
    reason: "",
  };
  appendRowToSheet(attendanceSheet, newRecord);
  return { success: true, message: "Absen masuk berhasil." };
}

function clockOut(requestData) {
  const { employeeID } = requestData;
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const today = formatDate(new Date());
  const attendanceRecords = getRowsData(attendanceSheet);

  const existingRecordIndex = attendanceRecords.findIndex(
    (record) => record.employeeID === employeeID && record.date === today
  );

  if (existingRecordIndex > -1) {
    const record = attendanceRecords[existingRecordIndex];
    if (record.clockOut) {
      return { success: false, message: "Anda sudah absen pulang hari ini." };
    }
    // Jika belum absen masuk, tetap izinkan absen pulang tapi kosongkan jam masuknya
    if (!record.clockIn && record.status === "present") {
      record.clockIn = ""; // Pastikan kosong jika belum ada jam masuk
    }
    record.clockOut = formatTime(new Date());
    updateRowInSheet(attendanceSheet, existingRecordIndex, record);
    return { success: true, message: "Absen pulang berhasil." };
  } else {
    // If no clock-in record for today, create one with clock-out only
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const employees = getRowsData(employeesSheet);
    const employee = employees.find((emp) => emp.id === employeeID);
    if (!employee) {
      return { success: false, message: "Karyawan tidak ditemukan." };
    }
    const newId = getNextId(attendanceSheet);
    const newRecord = {
      id: newId,
      employeeID: employeeID,
      employeeName: employee.name,
      date: today,
      clockIn: "", // No clock in
      clockOut: formatTime(new Date()),
      status: "present", // Assume present if clocking out
      reason: "",
    };
    appendRowToSheet(attendanceSheet, newRecord);
    return { success: true, message: "Absen pulang berhasil dicatat." };
  }
}

function markStatus(requestData) {
  const { employeeID, status, reason } = requestData;
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const employees = getRowsData(employeesSheet);
  const employee = employees.find((emp) => emp.id === employeeID);

  if (!employee) {
    return { success: false, message: "Karyawan tidak ditemukan." };
  }

  const today = formatDate(new Date());
  const attendanceRecords = getRowsData(attendanceSheet);

  const existingRecordIndex = attendanceRecords.findIndex(
    (record) => record.employeeID === employeeID && record.date === today
  );

  if (existingRecordIndex > -1) {
    // Update existing record
    const record = attendanceRecords[existingRecordIndex];
    record.status = status;
    record.reason = reason;
    if (status !== "present") {
      // Clear clock times if status is not present
      record.clockIn = "";
      record.clockOut = "";
    }
    updateRowInSheet(attendanceSheet, existingRecordIndex, record);
  } else {
    // Create new record
    const newId = getNextId(attendanceSheet);
    const newRecord = {
      id: newId,
      employeeID: employeeID,
      employeeName: employee.name,
      date: today,
      clockIn: "",
      clockOut: "",
      status: status,
      reason: reason,
    };
    appendRowToSheet(attendanceSheet, newRecord);
  }
  return { success: true, message: "Status berhasil diperbarui." };
}

function getTodayStatus(requestData) {
  const { employeeID } = requestData;
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const today = formatDate(new Date());
  const attendanceRecords = getRowsData(attendanceSheet);

  const todayRecord = attendanceRecords.find(
    (record) => record.employeeID === employeeID && record.date === today
  );

  if (todayRecord) {
    return { success: true, data: todayRecord };
  } else {
    // Return empty status if no record for today
    return {
      success: true,
      data: {
        clockIn: "-",
        clockOut: "-",
        status: "Belum Absen",
        reason: "",
      },
    };
  }
}

function getEmployeeHistory(requestData) {
  const { employeeID } = requestData;
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const attendanceRecords = getRowsData(attendanceSheet);

  const employeeHistory = attendanceRecords.filter(
    (record) => record.employeeID === employeeID
  );

  // Sort by date descending and get last 7 days
  employeeHistory.sort((a, b) => new Date(b.date) - new Date(a.date));
  const last7DaysHistory = employeeHistory.slice(0, 7);

  return { success: true, data: last7DaysHistory };
}

// --- FUNGSI REALTIME DATA (ADMIN) ---
function getRealtimeData() {
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const allEmployees = getRowsData(employeesSheet);
  const today = formatDate(new Date());
  const todayAttendance = getRowsData(attendanceSheet).filter(
    (record) => record.date === today
  );

  return {
    success: true,
    data: { employees: allEmployees, todayAttendance: todayAttendance },
  };
}

// --- FUNGSI REPORTS (ADMIN) ---
function getMonthlyReports() {
  const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);

  const allAttendance = getRowsData(attendanceSheet);
  const allEmployees = getRowsData(employeesSheet);

  const currentMonth = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM"
  );

  const reports = allEmployees.map((employee) => {
    let presentDays = 0;
    let permissionDays = 0;
    let sickDays = 0;

    allAttendance.forEach((att) => {
      // Pastikan att.date adalah string tanggal yang valid sebelum membuat objek Date
      if (att.date && typeof att.date === "string") {
        const attendanceMonth = Utilities.formatDate(
          new Date(att.date),
          Session.getScriptTimeZone(),
          "yyyy-MM"
        );
        if (
          att.employeeID === employee.id &&
          attendanceMonth === currentMonth
        ) {
          if (att.status === "present") {
            presentDays++;
          } else if (att.status === "permission") {
            permissionDays++;
          } else if (att.status === "sick") {
            sickDays++;
          }
        }
      }
    });

    return {
      id: employee.id,
      name: employee.name,
      store: employee.store,
      presentDays: presentDays,
      permissionDays: permissionDays,
      sickDays: sickDays,
    };
  });

  return { success: true, data: reports };
}
