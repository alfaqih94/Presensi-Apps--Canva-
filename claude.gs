// --- KONFIGURASI GOOGLE SHEET ---
const SPREADSHEET_ID = "14qAHqpqPXRdmxXElAhg1MyLaG9yzLPmDGqHfP_ga_uA"; // Ganti dengan ID Google Sheet Anda

// --- DOMAIN YANG DIIZINKAN UNTUK CORS ---
const ALLOWED_ORIGIN = "https://presensi-toptea.vercel.app";

// --- NAMA SHEET (TAB) ---
const SHEET_NAMES = {
  USERS: "Users",
  EMPLOYEES: "Employees",
  STORES: "Stores",
  ATTENDANCE: "Attendance",
};

// === FUNGSI UNTUK MENANGANI CORS ===
function setCorsHeaders() {
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers":
      "Content-Type, Authorization, X-Requested-With",
    "Access-Control-Max-Age": "86400",
  };
}

// === FUNGSI DOGET - MENANGANI PREFLIGHT REQUEST ===
function doGet(e) {
  // Menangani preflight request (OPTIONS)
  const output = ContentService.createTextOutput("").setMimeType(
    ContentService.MimeType.TEXT
  );

  // Set CORS headers
  const headers = setCorsHeaders();
  for (const [key, value] of Object.entries(headers)) {
    output.setHeader(key, value);
  }

  return output;
}

// === FUNGSI DOPOST - MENANGANI REQUEST UTAMA ===
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000); // Tunggu hingga 30 detik

    let result;
    let requestData = {};

    // Log untuk debugging
    console.log("Received POST request");
    console.log("e.postData:", e.postData);

    // Cek apakah ada data POST
    if (e && e.postData && e.postData.contents) {
      try {
        const parsedData = JSON.parse(e.postData.contents);
        console.log("Parsed data:", parsedData);

        const action = parsedData.action;
        requestData = parsedData.data || {};

        console.log("Action:", action);
        console.log("Request data:", requestData);

        // Route ke fungsi yang sesuai
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
            result = { success: false, message: "Aksi tidak valid: " + action };
        }
      } catch (jsonError) {
        console.error("JSON Parse Error:", jsonError);
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

    // Buat response dengan CORS headers
    const output = ContentService.createTextOutput(
      JSON.stringify(result)
    ).setMimeType(ContentService.MimeType.JSON);

    // Set CORS headers
    const headers = setCorsHeaders();
    for (const [key, value] of Object.entries(headers)) {
      output.setHeader(key, value);
    }

    console.log("Returning result:", result);
    return output;
  } catch (error) {
    console.error("Error in doPost:", error);

    // Return error dengan CORS headers
    const errorOutput = ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.message || error.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);

    // Set CORS headers untuk error response
    const headers = setCorsHeaders();
    for (const [key, value] of Object.entries(headers)) {
      errorOutput.setHeader(key, value);
    }

    return errorOutput;
  } finally {
    lock.releaseLock();
  }
}

// === FUNGSI UTILITAS ===
function getSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(
        `Sheet "${sheetName}" tidak ditemukan. Pastikan nama sheet sudah benar.`
      );
    }
    return sheet;
  } catch (error) {
    console.error("Error getting sheet:", error);
    throw error;
  }
}

function getRowsData(sheet) {
  try {
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length === 0 || values[0].length === 0) return [];

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
  } catch (error) {
    console.error("Error getting rows data:", error);
    throw error;
  }
}

function appendRowToSheet(sheet, rowObject) {
  try {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const newRow = [];
    for (let i = 0; i < headers.length; i++) {
      newRow.push(rowObject[headers[i]] || "");
    }
    sheet.appendRow(newRow);
  } catch (error) {
    console.error("Error appending row:", error);
    throw error;
  }
}

function updateRowInSheet(sheet, rowIndex, rowObject) {
  try {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const currentRowValues = sheet
      .getRange(rowIndex + 1, 1, 1, headers.length)
      .getValues()[0];

    const updatedValues = [];
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      updatedValues.push(
        rowObject[header] !== undefined
          ? rowObject[header]
          : currentRowValues[i]
      );
    }
    sheet
      .getRange(rowIndex + 1, 1, 1, updatedValues.length)
      .setValues([updatedValues]);
  } catch (error) {
    console.error("Error updating row:", error);
    throw error;
  }
}

function deleteRowInSheet(sheet, rowIndex) {
  try {
    sheet.deleteRow(rowIndex + 1);
  } catch (error) {
    console.error("Error deleting row:", error);
    throw error;
  }
}

function getNextId(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0 || sheet.getRange(1, 1).getValue() !== "id") {
      return 1;
    }

    const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let maxId = 0;
    for (let i = 0; i < idColumn.length; i++) {
      const currentId = idColumn[i][0];
      if (typeof currentId === "number") {
        maxId = Math.max(maxId, currentId);
      }
    }
    return maxId + 1;
  } catch (error) {
    console.error("Error getting next ID:", error);
    return 1;
  }
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

// === FUNGSI LOGIN ===
function login(requestData) {
  try {
    const { userID, password } = requestData;

    if (!userID || !password) {
      return { success: false, message: "User ID dan password harus diisi." };
    }

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
  } catch (error) {
    console.error("Error in login:", error);
    return { success: false, message: "Error saat login: " + error.message };
  }
}

// === FUNGSI EMPLOYEES (ADMIN) ===
function getEmployees() {
  try {
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const employees = getRowsData(employeesSheet);
    return { success: true, data: employees };
  } catch (error) {
    console.error("Error getting employees:", error);
    return {
      success: false,
      message: "Error mengambil data karyawan: " + error.message,
    };
  }
}

function addEmployee(employeeData) {
  try {
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const usersSheet = getSheet(SHEET_NAMES.USERS);

    const newId = getNextId(employeesSheet);
    const newEmployee = {
      id: newId,
      name: employeeData.name,
      userID: employeeData.userID,
      password: employeeData.password,
      role: "employee",
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
    });

    return {
      success: true,
      message: "Karyawan berhasil ditambahkan.",
      data: newEmployee,
    };
  } catch (error) {
    console.error("Error adding employee:", error);
    return {
      success: false,
      message: "Error menambah karyawan: " + error.message,
    };
  }
}

function deleteEmployee(requestData) {
  try {
    const { id } = requestData;
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const usersSheet = getSheet(SHEET_NAMES.USERS);

    const employees = getRowsData(employeesSheet);
    const indexToDelete = employees.findIndex((emp) => emp.id === id);

    if (indexToDelete > -1) {
      const employeeToDelete = employees[indexToDelete];
      deleteRowInSheet(employeesSheet, indexToDelete + 1);

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
  } catch (error) {
    console.error("Error deleting employee:", error);
    return {
      success: false,
      message: "Error menghapus karyawan: " + error.message,
    };
  }
}

// === FUNGSI STORES (ADMIN) ===
function getStores() {
  try {
    const storesSheet = getSheet(SHEET_NAMES.STORES);
    const stores = getRowsData(storesSheet);
    return { success: true, data: stores };
  } catch (error) {
    console.error("Error getting stores:", error);
    return {
      success: false,
      message: "Error mengambil data toko: " + error.message,
    };
  }
}

function addStore(storeData) {
  try {
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
  } catch (error) {
    console.error("Error adding store:", error);
    return { success: false, message: "Error menambah toko: " + error.message };
  }
}

function deleteStore(requestData) {
  try {
    const { id } = requestData;
    const storesSheet = getSheet(SHEET_NAMES.STORES);
    const stores = getRowsData(storesSheet);
    const indexToDelete = stores.findIndex((store) => store.id === id);

    if (indexToDelete > -1) {
      deleteRowInSheet(storesSheet, indexToDelete + 1);
      return { success: true, message: "Toko berhasil dihapus." };
    } else {
      return { success: false, message: "Toko tidak ditemukan." };
    }
  } catch (error) {
    console.error("Error deleting store:", error);
    return {
      success: false,
      message: "Error menghapus toko: " + error.message,
    };
  }
}

// === FUNGSI ATTENDANCE (EMPLOYEE) ===
function clockIn(requestData) {
  try {
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

    const existingRecordIndex = attendanceRecords.findIndex(
      (record) => record.employeeID === employeeID && record.date === today
    );

    if (existingRecordIndex > -1) {
      const record = attendanceRecords[existingRecordIndex];
      if (record.status === "present" && record.clockIn) {
        return { success: false, message: "Anda sudah absen masuk hari ini." };
      } else {
        record.clockIn = formatTime(new Date());
        record.status = "present";
        record.reason = "";
        updateRowInSheet(attendanceSheet, existingRecordIndex, record);
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
  } catch (error) {
    console.error("Error in clockIn:", error);
    return { success: false, message: "Error absen masuk: " + error.message };
  }
}

function clockOut(requestData) {
  try {
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
      if (!record.clockIn && record.status === "present") {
        record.clockIn = "";
      }
      record.clockOut = formatTime(new Date());
      updateRowInSheet(attendanceSheet, existingRecordIndex, record);
      return { success: true, message: "Absen pulang berhasil." };
    } else {
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
        clockIn: "",
        clockOut: formatTime(new Date()),
        status: "present",
        reason: "",
      };
      appendRowToSheet(attendanceSheet, newRecord);
      return { success: true, message: "Absen pulang berhasil dicatat." };
    }
  } catch (error) {
    console.error("Error in clockOut:", error);
    return { success: false, message: "Error absen pulang: " + error.message };
  }
}

function markStatus(requestData) {
  try {
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
      const record = attendanceRecords[existingRecordIndex];
      record.status = status;
      record.reason = reason;
      if (status !== "present") {
        record.clockIn = "";
        record.clockOut = "";
      }
      updateRowInSheet(attendanceSheet, existingRecordIndex, record);
    } else {
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
  } catch (error) {
    console.error("Error in markStatus:", error);
    return {
      success: false,
      message: "Error mengubah status: " + error.message,
    };
  }
}

function getTodayStatus(requestData) {
  try {
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
  } catch (error) {
    console.error("Error getting today status:", error);
    return {
      success: false,
      message: "Error mengambil status hari ini: " + error.message,
    };
  }
}

function getEmployeeHistory(requestData) {
  try {
    const { employeeID } = requestData;
    const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
    const attendanceRecords = getRowsData(attendanceSheet);

    const employeeHistory = attendanceRecords.filter(
      (record) => record.employeeID === employeeID
    );

    employeeHistory.sort((a, b) => new Date(b.date) - new Date(a.date));
    const last7DaysHistory = employeeHistory.slice(0, 7);

    return { success: true, data: last7DaysHistory };
  } catch (error) {
    console.error("Error getting employee history:", error);
    return {
      success: false,
      message: "Error mengambil history: " + error.message,
    };
  }
}

// === FUNGSI REALTIME DATA (ADMIN) ===
function getRealtimeData() {
  try {
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
  } catch (error) {
    console.error("Error getting realtime data:", error);
    return {
      success: false,
      message: "Error mengambil data realtime: " + error.message,
    };
  }
}

// === FUNGSI REPORTS (ADMIN) ===
function getMonthlyReports() {
  try {
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
  } catch (error) {
    console.error("Error getting monthly reports:", error);
    return {
      success: false,
      message: "Error mengambil laporan bulanan: " + error.message,
    };
  }
}
