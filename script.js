// Global variables to store data
let reportData = null;
let hostelData = null;
let inst1Data = null;
let inst2Data = null;
let stud1Data = null;
let stud2Data = null;
let currentReportData = null;

// File paths in static folder
const files = {
  report: "./static/Scholarship_Report.xlsx",
  inst1: "./static/Institute_I_All.xlsx",
  inst2: "./static/Institute_II_All.xlsx",
  stud1: "./static/Student_I_All.xlsx",
  stud2: "./static/Student_II_All.xlsx",
};

// Column definitions
const columns = [
  "Class",
  "Type",
  "Roll No",
  "Application No",
  "Scheme",
  "Financial Year",
  "Institute 1st Amount",
  "Institute 1st Reason",
  "Institute 2nd Amount",
  "Institute 2nd Reason",
  "Student 1st Amount",
  "Student 1st Reason",
  "Student 1st Bank",
  "Student 1st Acc",
  "Student 1st Date",
  "Student 2nd Amount",
  "Student 2nd Reason",
  "Student 2nd Bank",
  "Student 2nd Acc",
  "Student 2nd Date",
  "Hostel 1st Amount",
  "Hostel 1st Reason",
  "Hostel 1st Bank",
  "Hostel 1st Acc",
  "Hostel 1st Date",
  "Hostel 2nd Amount",
  "Hostel 2nd Reason",
  "Hostel 2nd Bank",
  "Hostel 2nd Acc",
  "Hostel 2nd Date",
];

// Initialize when page loads
window.addEventListener("DOMContentLoaded", function () {
  loadAllData();
});

// Load all Excel files
async function loadAllData() {
  try {
    showLoading(true);

    const [reportFile, inst1File, inst2File, stud1File, stud2File] =
      await Promise.all([
        fetch(files.report).then((res) => res.arrayBuffer()),
        fetch(files.inst1).then((res) => res.arrayBuffer()),
        fetch(files.inst2).then((res) => res.arrayBuffer()),
        fetch(files.stud1).then((res) => res.arrayBuffer()),
        fetch(files.stud2).then((res) => res.arrayBuffer()),
      ]);

    // Parse Excel files
    const reportWorkbook = XLSX.read(reportFile, { type: "array" });
    reportData = XLSX.utils.sheet_to_json(reportWorkbook.Sheets["Sheet1"], {
      header: 1,
      defval: "",
    });
    hostelData = XLSX.utils.sheet_to_json(reportWorkbook.Sheets["Sheet2"], {
      header: 1,
      defval: "",
    });

    const inst1Workbook = XLSX.read(inst1File, { type: "array" });
    inst1Data = XLSX.utils.sheet_to_json(
      inst1Workbook.Sheets[inst1Workbook.SheetNames[0]],
      { header: 1, defval: "" }
    );

    const inst2Workbook = XLSX.read(inst2File, { type: "array" });
    inst2Data = XLSX.utils.sheet_to_json(
      inst2Workbook.Sheets[inst2Workbook.SheetNames[0]],
      { header: 1, defval: "" }
    );

    const stud1Workbook = XLSX.read(stud1File, { type: "array" });
    stud1Data = XLSX.utils.sheet_to_json(
      stud1Workbook.Sheets[stud1Workbook.SheetNames[0]],
      { header: 1, defval: "" }
    );

    const stud2Workbook = XLSX.read(stud2File, { type: "array" });
    stud2Data = XLSX.utils.sheet_to_json(
      stud2Workbook.Sheets[stud2Workbook.SheetNames[0]],
      { header: 1, defval: "" }
    );

    showResult("✅ All data files loaded successfully!", "success");
  } catch (error) {
    console.error("Error loading data:", error);
    showResult(
      "❌ Error loading data files. Please check if all Excel files are in the static folder.",
      "error"
    );
  } finally {
    showLoading(false);
  }
}

// Clean application number
function cleanApp(value) {
  if (!value || value === "") return "";
  return String(value).trim().replace(".0", "");
}

// Get value from data array
function getValue(data, headers, rowIndex, columnName) {
  if (!data || !headers || rowIndex < 0 || rowIndex >= data.length) return "";
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) return "";
  return data[rowIndex][colIndex] || "";
}

// Find row index by enrollment number
function findRowByEnrollment(data, headers, enrollmentNo) {
  const enrollColIndex = headers.indexOf("Enrollment No");
  if (enrollColIndex === -1) return -1;

  for (let i = 1; i < data.length; i++) {
    // Start from 1 to skip header
    if (cleanApp(data[i][enrollColIndex]) === enrollmentNo) {
      return i;
    }
  }
  return -1;
}

// Find rows by application number
function findRowsByAppNo(data, headers, appNo) {
  const appColIndex = headers.indexOf("Application No");
  if (appColIndex === -1) return [];

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    // Start from 1 to skip header
    if (cleanApp(data[i][appColIndex]) === appNo) {
      rows.push(i);
    }
  }
  return rows;
}

// Main search function
function searchStudent() {
  const enrollmentNo = document.getElementById("enrollmentNo").value.trim();

  if (!enrollmentNo) {
    showResult("❌ Please enter an enrollment number", "error");
    return;
  }

  if (
    !reportData ||
    !hostelData ||
    !inst1Data ||
    !inst2Data ||
    !stud1Data ||
    !stud2Data
  ) {
    showResult("❌ Data not loaded. Please wait for files to load.", "error");
    return;
  }

  showLoading(true);

  try {
    // Find student in both sheets
    const reportHeaders = reportData[0];
    const hostelHeaders = hostelData[0];

    const studentRowIndex = findRowByEnrollment(
      reportData,
      reportHeaders,
      enrollmentNo
    );
    const hostelStudentRowIndex = findRowByEnrollment(
      hostelData,
      hostelHeaders,
      enrollmentNo
    );

    if (studentRowIndex === -1 && hostelStudentRowIndex === -1) {
      showResult(`❌ Enrollment No not found: ${enrollmentNo}`, "error");
      showLoading(false);
      return;
    }

    // Get student name
    let studentName = "";
    if (studentRowIndex !== -1) {
      studentName = getValue(
        reportData,
        reportHeaders,
        studentRowIndex,
        "Beneficiary Name"
      );
    } else if (hostelStudentRowIndex !== -1) {
      studentName = getValue(
        hostelData,
        hostelHeaders,
        hostelStudentRowIndex,
        "Beneficiary Name"
      );
    }

    showResult(`✅ Report for ${enrollmentNo} - ${studentName}`, "success");

    // Generate report data
    const rows = [];
    const colMap = { FY: 3, SY: 7, TY: 11 }; // Column indices for Application No

    // Process normal scholarship data
    if (studentRowIndex !== -1) {
      Object.entries(colMap).forEach(([cls, colIdx]) => {
        const rollNo = reportData[studentRowIndex][colIdx - 1] || "";
        const appNo = cleanApp(reportData[studentRowIndex][colIdx]);
        const scheme = reportData[studentRowIndex][colIdx + 1] || "";
        const finYear = reportData[studentRowIndex][colIdx + 2] || "";

        if (!appNo) return;

        const inst1Rows = findRowsByAppNo(inst1Data, inst1Data[0], appNo);
        const inst2Rows = findRowsByAppNo(inst2Data, inst2Data[0], appNo);
        const stud1Rows = findRowsByAppNo(stud1Data, stud1Data[0], appNo);
        const stud2Rows = findRowsByAppNo(stud2Data, stud2Data[0], appNo);

        const row = {
          Class: cls,
          Type: "Scholarship",
          "Roll No": rollNo,
          "Application No": appNo,
          Scheme: scheme,
          "Financial Year": finYear,
          "Institute 1st Amount":
            inst1Rows.length > 0
              ? getValue(
                  inst1Data,
                  inst1Data[0],
                  inst1Rows[0],
                  "Disbursed Amount(₹)"
                )
              : "",
          "Institute 1st Reason":
            inst1Rows.length > 0
              ? getValue(inst1Data, inst1Data[0], inst1Rows[0], "Reason")
              : "",
          "Institute 2nd Amount":
            inst2Rows.length > 0
              ? getValue(
                  inst2Data,
                  inst2Data[0],
                  inst2Rows[0],
                  "Disbursed Amount(₹)"
                )
              : "",
          "Institute 2nd Reason":
            inst2Rows.length > 0
              ? getValue(inst2Data, inst2Data[0], inst2Rows[0], "Reason")
              : "",
          "Student 1st Amount":
            stud1Rows.length > 0
              ? getValue(
                  stud1Data,
                  stud1Data[0],
                  stud1Rows[0],
                  "Amount Disbursed To Student(₹)"
                )
              : "",
          "Student 1st Reason":
            stud1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], stud1Rows[0], "Reason")
              : "",
          "Student 1st Bank":
            stud1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], stud1Rows[0], "Bank Name")
              : "",
          "Student 1st Acc":
            stud1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], stud1Rows[0], "Account No")
              : "",
          "Student 1st Date":
            stud1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], stud1Rows[0], "Credit Date")
              : "",
          "Student 2nd Amount":
            stud2Rows.length > 0
              ? getValue(
                  stud2Data,
                  stud2Data[0],
                  stud2Rows[0],
                  "Amount Disbursed To Student(₹)"
                )
              : "",
          "Student 2nd Reason":
            stud2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], stud2Rows[0], "Reason")
              : "",
          "Student 2nd Bank":
            stud2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], stud2Rows[0], "Bank Name")
              : "",
          "Student 2nd Acc":
            stud2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], stud2Rows[0], "Account No")
              : "",
          "Student 2nd Date":
            stud2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], stud2Rows[0], "Credit Date")
              : "",
          "Hostel 1st Amount": "",
          "Hostel 1st Reason": "",
          "Hostel 1st Bank": "",
          "Hostel 1st Acc": "",
          "Hostel 1st Date": "",
          "Hostel 2nd Amount": "",
          "Hostel 2nd Reason": "",
          "Hostel 2nd Bank": "",
          "Hostel 2nd Acc": "",
          "Hostel 2nd Date": "",
        };
        rows.push(row);
      });
    }

    // Process hostel scholarship data
    if (hostelStudentRowIndex !== -1) {
      Object.entries(colMap).forEach(([cls, colIdx]) => {
        const rollNo = hostelData[hostelStudentRowIndex][colIdx - 1] || "";
        const appNo = cleanApp(hostelData[hostelStudentRowIndex][colIdx]);
        const scheme = hostelData[hostelStudentRowIndex][colIdx + 1] || "";
        const finYear = hostelData[hostelStudentRowIndex][colIdx + 2] || "";

        if (!appNo) return;

        const h1Rows = findRowsByAppNo(stud1Data, stud1Data[0], appNo);
        const h2Rows = findRowsByAppNo(stud2Data, stud2Data[0], appNo);

        const row = {
          Class: cls,
          Type: "Hostel",
          "Roll No": rollNo,
          "Application No": appNo,
          Scheme: scheme,
          "Financial Year": finYear,
          "Institute 1st Amount": "",
          "Institute 1st Reason": "",
          "Institute 2nd Amount": "",
          "Institute 2nd Reason": "",
          "Student 1st Amount": "",
          "Student 1st Reason": "",
          "Student 1st Bank": "",
          "Student 1st Acc": "",
          "Student 1st Date": "",
          "Student 2nd Amount": "",
          "Student 2nd Reason": "",
          "Student 2nd Bank": "",
          "Student 2nd Acc": "",
          "Student 2nd Date": "",
          "Hostel 1st Amount":
            h1Rows.length > 0
              ? getValue(
                  stud1Data,
                  stud1Data[0],
                  h1Rows[0],
                  "Amount Disbursed To Student(₹)"
                )
              : "",
          "Hostel 1st Reason":
            h1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], h1Rows[0], "Reason")
              : "",
          "Hostel 1st Bank":
            h1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], h1Rows[0], "Bank Name")
              : "",
          "Hostel 1st Acc":
            h1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], h1Rows[0], "Account No")
              : "",
          "Hostel 1st Date":
            h1Rows.length > 0
              ? getValue(stud1Data, stud1Data[0], h1Rows[0], "Credit Date")
              : "",
          "Hostel 2nd Amount":
            h2Rows.length > 0
              ? getValue(
                  stud2Data,
                  stud2Data[0],
                  h2Rows[0],
                  "Amount Disbursed To Student(₹)"
                )
              : "",
          "Hostel 2nd Reason":
            h2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], h2Rows[0], "Reason")
              : "",
          "Hostel 2nd Bank":
            h2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], h2Rows[0], "Bank Name")
              : "",
          "Hostel 2nd Acc":
            h2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], h2Rows[0], "Account No")
              : "",
          "Hostel 2nd Date":
            h2Rows.length > 0
              ? getValue(stud2Data, stud2Data[0], h2Rows[0], "Credit Date")
              : "",
        };
        rows.push(row);
      });
    }

    currentReportData = rows;
    displayTable(rows);
  } catch (error) {
    console.error("Error processing data:", error);
    showResult("❌ Error processing data: " + error.message, "error");
  } finally {
    showLoading(false);
  }
}

// Display data in table
function displayTable(data) {
  const tableContainer = document.getElementById("tableContainer");
  const tableHead = document.getElementById("tableHead");
  const tableBody = document.getElementById("tableBody");
  const downloadBtn = document.getElementById("downloadBtn");

  // Clear previous data
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";

  if (data.length === 0) {
    tableContainer.style.display = "none";
    return;
  }

  // Create header
  const headerRow = document.createElement("tr");
  columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  // Create body rows
  data.forEach((row) => {
    const tr = document.createElement("tr");
    tr.className = row.Type === "Hostel" ? "hostel-row" : "scholarship-row";

    columns.forEach((column) => {
      const td = document.createElement("td");
      const value = row[column] || "";
      td.textContent = value;

      // Apply conditional formatting
      if (value.includes("Not Available") || value.includes("Pending")) {
        td.classList.add("not-available");
      } else if (value.includes("Amount Disbursed")) {
        td.classList.add("amount-disbursed");
      }

      tr.appendChild(td);
    });

    tableBody.appendChild(tr);
  });

  tableContainer.style.display = "block";
  downloadBtn.style.display = "block";
}

// Download Excel file
function downloadExcel() {
  if (!currentReportData) return;

  const enrollmentNo = document.getElementById("enrollmentNo").value.trim();
  const ws = XLSX.utils.json_to_sheet(currentReportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");

  // Apply some basic styling (limited in browser environment)
  const range = XLSX.utils.decode_range(ws["!ref"]);

  // Style header row
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
    if (!ws[cellRef]) continue;
    ws[cellRef].s = {
      fill: { fgColor: { rgb: "FFFF99" } },
      font: { bold: true },
    };
  }

  XLSX.writeFile(wb, `Scholarship_Report_${enrollmentNo}.xlsx`);
}

// Utility functions
function showLoading(show) {
  const loading = document.getElementById("loading");
  const searchBtn = document.getElementById("searchBtn");

  if (show) {
    loading.style.display = "block";
    searchBtn.disabled = true;
  } else {
    loading.style.display = "none";
    searchBtn.disabled = false;
  }
}

function showResult(message, type) {
  const resultDiv = document.getElementById("result");
  resultDiv.innerHTML = `<div class="result ${type}">${message}</div>`;
}

// Allow Enter key to trigger search
document
  .getElementById("enrollmentNo")
  .addEventListener("keypress", function (event) {
    if (event.key === "Enter") {
      searchStudent();
    }
  });
