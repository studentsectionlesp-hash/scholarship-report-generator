// script.js - Converted from your new Python script
// Requires: xlsx (already in your HTML). This file will dynamically load jsPDF + autotable for PDF export.

/////////////////////////////////////////////////////
// Globals & file paths (same as before)
/////////////////////////////////////////////////////
let reportData = null;
let hostelData = null;
let inst1Data = null;
let inst2Data = null;
let stud1Data = null;
let stud2Data = null;
let currentEnroll = null;

const files = {
  report: "./static/Scholarship_Report.xlsx",
  inst1: "./static/Institute_I_All.xlsx",
  inst2: "./static/Institute_II_All.xlsx",
  stud1: "./static/Student_I_All.xlsx",
  stud2: "./static/Student_II_All.xlsx",
};

const sch_cols = [
  "Class","Type","Roll No","Application No","Scheme","Financial Year",
  "Institute 1st Amount","Institute 1st Reason","Institute 2nd Amount","Institute 2nd Reason",
  "Student 1st Amount","Student 1st Reason","Student 1st Bank","Student 1st Acc","Student 1st Date",
  "Student 2nd Amount","Student 2nd Reason","Student 2nd Bank","Student 2nd Acc","Student 2nd Date"
];

const hostel_cols = [
  "Class","Type","Roll No","Application No","Scheme","Financial Year",
  "Hostel 1st Amount","Hostel 1st Reason","Hostel 1st Bank","Hostel 1st Acc","Hostel 1st Date",
  "Hostel 2nd Amount","Hostel 2nd Reason","Hostel 2nd Bank","Hostel 2nd Acc","Hostel 2nd Date"
];

/////////////////////////////////////////////////////
// Utility helpers
/////////////////////////////////////////////////////
function loadScript(url) {
  return new Promise((res, rej) => {
    if (document.querySelector(`script[src="${url}"]`)) return res();
    const s = document.createElement("script");
    s.src = url;
    s.onload = () => res();
    s.onerror = () => rej(new Error("Failed to load " + url));
    document.head.appendChild(s);
  });
}

function cleanApp(v) {
  if (v === undefined || v === null) return "";
  return String(v).trim().replace(".0", "");
}

function getValFromRows(rows, headers, idx, colName) {
  if (!rows || idx < 0 || idx >= rows.length) return "";
  const colIndex = headers.indexOf(colName);
  if (colIndex === -1) return "";
  return rows[idx][colIndex] === undefined ? "" : rows[idx][colIndex];
}

function findRowByEnrollment(data, headers, enrollmentNo) {
  const enrollIdx = headers.indexOf("Enrollment No");
  if (enrollIdx === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (cleanApp(data[i][enrollIdx]) === enrollmentNo) return i;
  }
  return -1;
}

function findRowsByAppNo(data, headers, appNo) {
  const appIdx = headers.indexOf("Application No");
  if (appIdx === -1) return [];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (cleanApp(data[i][appIdx]) === appNo) rows.push(i);
  }
  return rows;
}

/////////////////////////////////////////////////////
// Load all Excel files and parse to arrays (header row included)
/////////////////////////////////////////////////////
async function loadAllData() {
  try {
    showLoading(true);

    const [rBuf, i1Buf, i2Buf, s1Buf, s2Buf] = await Promise.all([
      fetch(files.report).then(r => r.arrayBuffer()),
      fetch(files.inst1).then(r => r.arrayBuffer()),
      fetch(files.inst2).then(r => r.arrayBuffer()),
      fetch(files.stud1).then(r => r.arrayBuffer()),
      fetch(files.stud2).then(r => r.arrayBuffer()),
    ]);

    const rWb = XLSX.read(rBuf, { type: "array" });
    // Sheet1 and Sheet2 from report workbook
    reportData = XLSX.utils.sheet_to_json(rWb.Sheets["Sheet1"], { header: 1, defval: "" });
    hostelData = XLSX.utils.sheet_to_json(rWb.Sheets["Sheet2"], { header: 1, defval: "" });

    const i1Wb = XLSX.read(i1Buf, { type: "array" });
    inst1Data = XLSX.utils.sheet_to_json(i1Wb.Sheets[i1Wb.SheetNames[0]], { header: 1, defval: "" });

    const i2Wb = XLSX.read(i2Buf, { type: "array" });
    inst2Data = XLSX.utils.sheet_to_json(i2Wb.Sheets[i2Wb.SheetNames[0]], { header: 1, defval: "" });

    const s1Wb = XLSX.read(s1Buf, { type: "array" });
    stud1Data = XLSX.utils.sheet_to_json(s1Wb.Sheets[s1Wb.SheetNames[0]], { header: 1, defval: "" });

    const s2Wb = XLSX.read(s2Buf, { type: "array" });
    stud2Data = XLSX.utils.sheet_to_json(s2Wb.Sheets[s2Wb.SheetNames[0]], { header: 1, defval: "" });

    showResult("✅ All data files loaded successfully!", "success");
  } catch (err) {
    console.error("Error loading data:", err);
    showResult("❌ Error loading data files. Please check static folder.", "error");
  } finally {
    showLoading(false);
  }
}

/////////////////////////////////////////////////////
// Search logic (mirrors new Python version)
/////////////////////////////////////////////////////
function buildRowsForStudent(enrollmentNo) {
  const rows = [];
  const colMap = { FY: 3, SY: 7, TY: 11 };

  // Ensure enrollment strings trimmed in report arrays
  // (we'll use cleanApp on reads)
  const rHeaders = reportData[0] || [];
  const hHeaders = hostelData[0] || [];

  const studentIdx = findRowByEnrollment(reportData, rHeaders, enrollmentNo);
  const hostelIdx = findRowByEnrollment(hostelData, hHeaders, enrollmentNo);

  if (studentIdx === -1 && hostelIdx === -1) return { rows: [], studentIdx, hostelIdx };

  // Scholarship portion
  if (studentIdx !== -1) {
    Object.entries(colMap).forEach(([cls, colIdx]) => {
      const roll = reportData[studentIdx][colIdx - 1] || "";
      const appno = cleanApp(reportData[studentIdx][colIdx]);
      const scheme = reportData[studentIdx][colIdx + 1] || "";
      const fin_year = reportData[studentIdx][colIdx + 2] || "";
      if (!appno) return;

      const inst1Rows = findRowsByAppNo(inst1Data, inst1Data[0], appno);
      const inst2Rows = findRowsByAppNo(inst2Data, inst2Data[0], appno);
      const stud1Rows = findRowsByAppNo(stud1Data, stud1Data[0], appno);
      const stud2Rows = findRowsByAppNo(stud2Data, stud2Data[0], appno);

      rows.push({
        "Class": cls, "Type": "Scholarship", "Roll No": roll, "Application No": appno,
        "Scheme": scheme, "Financial Year": fin_year,
        "Institute 1st Amount": inst1Rows.length>0 ? inst1Data[inst1Rows[0]][inst1Data[0].indexOf("Disbursed Amount(₹)")] || "" : "",
        "Institute 1st Reason": inst1Rows.length>0 ? inst1Data[inst1Rows[0]][inst1Data[0].indexOf("Reason")] || "" : "",
        "Institute 2nd Amount": inst2Rows.length>0 ? inst2Data[inst2Rows[0]][inst2Data[0].indexOf("Disbursed Amount(₹)")] || "" : "",
        "Institute 2nd Reason": inst2Rows.length>0 ? inst2Data[inst2Rows[0]][inst2Data[0].indexOf("Reason")] || "" : "",
        "Student 1st Amount": stud1Rows.length>0 ? stud1Data[stud1Rows[0]][stud1Data[0].indexOf("Amount Disbursed To Student(₹)")] || "" : "",
        "Student 1st Reason": stud1Rows.length>0 ? stud1Data[stud1Rows[0]][stud1Data[0].indexOf("Reason")] || "" : "",
        "Student 1st Bank": stud1Rows.length>0 ? stud1Data[stud1Rows[0]][stud1Data[0].indexOf("Bank Name")] || "" : "",
        "Student 1st Acc": stud1Rows.length>0 ? stud1Data[stud1Rows[0]][stud1Data[0].indexOf("Account No")] || "" : "",
        "Student 1st Date": stud1Rows.length>0 ? stud1Data[stud1Rows[0]][stud1Data[0].indexOf("Credit Date")] || "" : "",
        "Student 2nd Amount": stud2Rows.length>0 ? stud2Data[stud2Rows[0]][stud2Data[0].indexOf("Amount Disbursed To Student(₹)")] || "" : "",
        "Student 2nd Reason": stud2Rows.length>0 ? stud2Data[stud2Rows[0]][stud2Data[0].indexOf("Reason")] || "" : "",
        "Student 2nd Bank": stud2Rows.length>0 ? stud2Data[stud2Rows[0]][stud2Data[0].indexOf("Bank Name")] || "" : "",
        "Student 2nd Acc": stud2Rows.length>0 ? stud2Data[stud2Rows[0]][stud2Data[0].indexOf("Account No")] || "" : "",
        "Student 2nd Date": stud2Rows.length>0 ? stud2Data[stud2Rows[0]][stud2Data[0].indexOf("Credit Date")] || "" : "",
      });
    });
  }

  // Hostel portion
  if (hostelIdx !== -1) {
    Object.entries(colMap).forEach(([cls, colIdx]) => {
      const roll = hostelData[hostelIdx][colIdx - 1] || "";
      const appno = cleanApp(hostelData[hostelIdx][colIdx]);
      const scheme = hostelData[hostelIdx][colIdx + 1] || "";
      const fin_year = hostelData[hostelIdx][colIdx + 2] || "";
      if (!appno) return;

      const h1Rows = findRowsByAppNo(stud1Data, stud1Data[0], appno);
      const h2Rows = findRowsByAppNo(stud2Data, stud2Data[0], appno);

      rows.push({
        "Class": cls, "Type": "Hostel", "Roll No": roll, "Application No": appno,
        "Scheme": scheme, "Financial Year": fin_year,
        "Hostel 1st Amount": h1Rows.length>0 ? stud1Data[h1Rows[0]][stud1Data[0].indexOf("Amount Disbursed To Student(₹)")] || "" : "",
        "Hostel 1st Reason": h1Rows.length>0 ? stud1Data[h1Rows[0]][stud1Data[0].indexOf("Reason")] || "" : "",
        "Hostel 1st Bank": h1Rows.length>0 ? stud1Data[h1Rows[0]][stud1Data[0].indexOf("Bank Name")] || "" : "",
        "Hostel 1st Acc": h1Rows.length>0 ? stud1Data[h1Rows[0]][stud1Data[0].indexOf("Account No")] || "" : "",
        "Hostel 1st Date": h1Rows.length>0 ? stud1Data[h1Rows[0]][stud1Data[0].indexOf("Credit Date")] || "" : "",
        "Hostel 2nd Amount": h2Rows.length>0 ? stud2Data[h2Rows[0]][stud2Data[0].indexOf("Amount Disbursed To Student(₹)")] || "" : "",
        "Hostel 2nd Reason": h2Rows.length>0 ? stud2Data[h2Rows[0]][stud2Data[0].indexOf("Reason")] || "" : "",
        "Hostel 2nd Bank": h2Rows.length>0 ? stud2Data[h2Rows[0]][stud2Data[0].indexOf("Bank Name")] || "" : "",
        "Hostel 2nd Acc": h2Rows.length>0 ? stud2Data[h2Rows[0]][stud2Data[0].indexOf("Account No")] || "" : "",
        "Hostel 2nd Date": h2Rows.length>0 ? stud2Data[h2Rows[0]][stud2Data[0].indexOf("Credit Date")] || "" : "",
      });
    });
  }

  return { rows, studentIdx, hostelIdx };
}

/////////////////////////////////////////////////////
// Transpose routine (mirrors pandas set_index('Class').T)
/////////////////////////////////////////////////////
function transposeReport(rows, type) {
  if (!rows || rows.length === 0) return [];

  // pick proper columns list (fields we want)
  const useCols = type === "Scholarship" ? sch_cols : hostel_cols;

  // Build mapping: { Class1: rowObj, Class2: rowObj, ... }
  // rows contain objects where "Class" is FY/SY/TY, and keys are the useCols
  const classMap = {};
  const classesOrdered = []; // preserve order found (FY,SY,TY)
  rows.forEach(r => {
    const cls = r["Class"] || "";
    if (!cls) return;
    classMap[cls] = r;
    if (!classesOrdered.includes(cls)) classesOrdered.push(cls);
  });

  // Fields: all column names except "Class" and "Type" and "Roll No"??? In python they used df.reindex(columns=use_cols) and set_index("Class").T
  // So fields = use_cols minus "Class" (first col), then the rest become rows.
  const fields = useCols.filter(c => c !== "Class");

  // Build transposed array: first column "Field", then one column per class in classesOrdered with values
  const transposed = [];
  const header = ["Field", ...classesOrdered];
  transposed.push(header);

  fields.forEach(field => {
    const row = [field];
    classesOrdered.forEach(cls => {
      const val = (classMap[cls] && (classMap[cls][field] !== undefined ? classMap[cls][field] : "")) || "";
      row.push(val);
    });
    transposed.push(row);
  });

  // Return as array-of-arrays (header included) to make it easy for display and writing
  return transposed;
}

/////////////////////////////////////////////////////
// Display: show two tables inside #tableContainer
/////////////////////////////////////////////////////
function displayReports(transSch, transHostel) {
  const tableContainer = document.getElementById("tableContainer");
  tableContainer.innerHTML = ""; // wipe previous

  // helper to create table from array-of-arrays
  function makeTableFromAA(aa, title) {
    const wrapper = document.createElement("div");
    wrapper.style.marginTop = "18px";
    const h = document.createElement("h3");
    h.textContent = title;
    wrapper.appendChild(h);

    const table = document.createElement("table");
    table.style.width = "100%";
    table.style.borderCollapse = "collapse";
    table.style.marginTop = "6px";

    const thead = document.createElement("thead");
    const thr = document.createElement("tr");
    aa[0].forEach(hc => {
      const th = document.createElement("th");
      th.textContent = hc;
      th.style.border = "1px solid #ddd";
      th.style.padding = "6px";
      th.style.background = "#ffff99";
      th.style.fontSize = "12px";
      thr.appendChild(th);
    });
    thead.appendChild(thr);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let r = 1; r < aa.length; r++) {
      const tr = document.createElement("tr");
      aa[r].forEach((cell, ci) => {
        const td = document.createElement("td");
        td.textContent = cell === undefined || cell === null ? "" : String(cell);
        td.style.border = "1px solid #ddd";
        td.style.padding = "6px";
        td.style.fontSize = "12px";

        // conditional color hints similar to Python
        if (typeof cell === "string" && (cell.includes("Not Available") || cell.includes("Pending"))) {
          td.style.background = "#ff9999";
        } else if (typeof cell === "string" && cell.includes("Amount Disbursed")) {
          td.style.background = "#ccffcc";
        }

        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    wrapper.appendChild(table);
    return wrapper;
  }

  if (transSch && transSch.length > 1) {
    tableContainer.appendChild(makeTableFromAA(transSch, "Scholarship Report"));
  } else {
    const p = document.createElement("p");
    p.textContent = "No Scholarship data found for this enrollment.";
    tableContainer.appendChild(p);
  }

  if (transHostel && transHostel.length > 1) {
    tableContainer.appendChild(makeTableFromAA(transHostel, "Hostel Report"));
  } else {
    const p = document.createElement("p");
    p.textContent = "No Hostel data found for this enrollment.";
    tableContainer.appendChild(p);
  }

  tableContainer.style.display = "block";

  // create download buttons (Excel + PDF). Keep original downloadBtn (if exists) and add a PDF button dynamically
  const existingDownloadBtn = document.getElementById("downloadBtn");
  if (existingDownloadBtn) {
    existingDownloadBtn.style.display = "inline-block";
  } else {
    // create one for compatibility
    const b = document.createElement("button");
    b.id = "downloadBtn";
    b.textContent = "📥 Download Excel Report";
    b.className = "download-btn";
    b.style.display = "inline-block";
    b.onclick = () => downloadExcel();
    tableContainer.appendChild(b);
  }

  // PDF button - create or update
  let pdfBtn = document.getElementById("downloadPdfBtn");
  if (!pdfBtn) {
    pdfBtn = document.createElement("button");
    pdfBtn.id = "downloadPdfBtn";
    pdfBtn.textContent = "📄 Download PDF Report";
    pdfBtn.className = "download-btn";
    pdfBtn.style.marginLeft = "8px";
    pdfBtn.onclick = () => downloadPdf();
    tableContainer.appendChild(pdfBtn);
  } else {
    pdfBtn.style.display = "inline-block";
  }
}

/////////////////////////////////////////////////////
// Main search handler called by HTML's button
/////////////////////////////////////////////////////
async function searchStudent() {
  const enrollmentNo = document.getElementById("enrollmentNo").value.trim();
  currentEnroll = enrollmentNo;
  if (!enrollmentNo) {
    showResult("❌ Please enter an enrollment number", "error");
    return;
  }

  if (!reportData || !hostelData || !inst1Data || !inst2Data || !stud1Data || !stud2Data) {
    showResult("❌ Data not loaded. Please wait.", "error");
    return;
  }

  showLoading(true);
  try {
    const { rows, studentIdx, hostelIdx } = buildRowsForStudent(enrollmentNo);

    if (!rows || rows.length === 0) {
      showResult(`❌ Enrollment No not found: ${enrollmentNo}`, "error");
      document.getElementById("tableContainer").style.display = "none";
      const db = document.getElementById("downloadBtn");
      if (db) db.style.display = "none";
      const pb = document.getElementById("downloadPdfBtn");
      if (pb) pb.style.display = "none";
      return;
    }

    // separate scholarship and hostel rows
    const schRows = rows.filter(r => r.Type === "Scholarship");
    const hRows = rows.filter(r => r.Type === "Hostel");

    const transSch = transposeReport(schRows, "Scholarship"); // array-of-arrays
    const transHostel = transposeReport(hRows, "Hostel");

    displayReports(transSch, transHostel);

    // show header success
    // attempt to read name for message
    let name = "";
    if (studentIdx !== -1) {
      const rh = reportData[0];
      name = getValFromRows(reportData, rh, studentIdx, "Beneficiary Name");
    } else if (hostelIdx !== -1) {
      const hh = hostelData[0];
      name = getValFromRows(hostelData, hh, hostelIdx, "Beneficiary Name");
    }
    showResult(`✅ Report for ${enrollmentNo} - ${name}`, "success");

  } catch (err) {
    console.error("Error processing:", err);
    showResult("❌ Error processing data: " + (err.message || err), "error");
  } finally {
    showLoading(false);
  }
}

/////////////////////////////////////////////////////
// Excel export (two sheets: Scholarship, Hostel)
/////////////////////////////////////////////////////
function downloadExcel() {
  const enrollmentNo = currentEnroll || document.getElementById("enrollmentNo").value.trim();
  if (!enrollmentNo) return;

  // find the current displayed tables from DOM
  const tableContainer = document.getElementById("tableContainer");
  if (!tableContainer || tableContainer.style.display === "none") return;

  // Create workbook and sheets from the displayed HTML tables
  const wb = XLSX.utils.book_new();

  // find the two tables we created (by headings)
  const tables = tableContainer.querySelectorAll("div > table");
  // if tables not found, fallback: try any table under container
  const allTables = tableContainer.querySelectorAll("table");
  const useTables = tables.length ? tables : allTables;

  // If none found, produce empty workbook
  if (!useTables || useTables.length === 0) {
    showResult("❌ No report to export.", "error");
    return;
  }

  // Convert each table to sheet and append
  for (let i = 0; i < useTables.length; i++) {
    const tbl = useTables[i];
    const sheet = XLSX.utils.table_to_sheet(tbl);
    const name = i === 0 ? "Scholarship" : "Hostel";
    XLSX.utils.book_append_sheet(wb, sheet, name);
  }

  XLSX.writeFile(wb, `Scholarship_Report_${enrollmentNo}.xlsx`);
  showResult(`📂 Excel saved: Scholarship_Report_${enrollmentNo}.xlsx`, "success");
}

/////////////////////////////////////////////////////
// PDF export using jsPDF + autotable (dynamically load libs)
/////////////////////////////////////////////////////
async function downloadPdf() {
  const enrollmentNo = currentEnroll || document.getElementById("enrollmentNo").value.trim();
  if (!enrollmentNo) return;

  showLoading(true);
  try {
    // dynamic load jsPDF and autotable if not present
    const jspdfUrl = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    const autotableUrl = "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js";

    await loadScript(jspdfUrl);
    await loadScript(autotableUrl);

    // jsPDF is available under window.jspdf
    const { jsPDF } = window.jspdf || window.jspdf || { jsPDF: window.jsPDF || window.jspdf?.jsPDF };
    if (!jsPDF) throw new Error("jsPDF not available");

    const doc = new jsPDF({
      unit: "pt",
      format: "a4",
    });

    const tableContainer = document.getElementById("tableContainer");
    if (!tableContainer) throw new Error("No report to export");

    // Build PDF content from displayed tables
    const headings = tableContainer.querySelectorAll("h3");
    const tables = tableContainer.querySelectorAll("table");

    let cursorY = 40;
    doc.setFontSize(14);
    doc.text(`Scholarship Report - ${enrollmentNo}`, 40, cursorY);
    cursorY += 16;

    // Helper to convert HTML table to autotable data
    function tableToAuto(table) {
      const aa = [];
      const headerCells = table.querySelectorAll("thead tr th");
      const headers = Array.from(headerCells).map(h => h.textContent || "");
      aa.push(headers);
      const rows = table.querySelectorAll("tbody tr");
      rows.forEach(tr => {
        const row = Array.from(tr.querySelectorAll("td")).map(td => td.textContent || "");
        aa.push(row);
      });
      return aa;
    }

    for (let i = 0; i < tables.length; i++) {
      const title = headings[i] ? headings[i].textContent : (i===0 ? "Scholarship" : "Hostel");
      const auto = tableToAuto(tables[i]);
      // add some space before table
      if (i !== 0) {
        doc.addPage();
        cursorY = 40;
        doc.setFontSize(14);
        doc.text(`${title} - ${enrollmentNo}`, 40, cursorY);
        cursorY += 16;
      } else {
        // first table already has title printed above
        doc.setFontSize(12);
        doc.text(title, 40, cursorY);
        cursorY += 14;
      }

      // Use autotable
      // eslint-disable-next-line no-undef
      doc.autoTable({
        head: [auto[0]],
        body: auto.slice(1),
        startY: cursorY + 10,
        styles: { fontSize: 7 },
        headStyles: { fillColor: [255, 255, 153], textColor: 0, halign: "center" },
        theme: "grid",
        margin: { left: 40, right: 40 },
      });

      cursorY = doc.lastAutoTable ? doc.lastAutoTable.finalY + 10 : 40;
    }

    doc.save(`Scholarship_Report_${enrollmentNo}.pdf`);
    showResult(`📂 PDF saved: Scholarship_Report_${enrollmentNo}.pdf`, "success");
  } catch (err) {
    console.error("PDF error:", err);
    showResult("❌ Could not create PDF: " + (err.message || err), "error");
  } finally {
    showLoading(false);
  }
}

/////////////////////////////////////////////////////
// UI helpers (reuse from your HTML)
/////////////////////////////////////////////////////
function showLoading(show) {
  const loading = document.getElementById("loading");
  const searchBtn = document.getElementById("searchBtn");
  if (loading) loading.style.display = show ? "block" : "none";
  if (searchBtn) searchBtn.disabled = show;
}

function showResult(message, type) {
  const resultDiv = document.getElementById("result");
  if (!resultDiv) return;
  resultDiv.innerHTML = `<div class="result ${type}">${message}</div>`;
}

/////////////////////////////////////////////////////
// Init on DOMContentLoaded
/////////////////////////////////////////////////////
window.addEventListener("DOMContentLoaded", async function () {
  await loadAllData();

  // Allow Enter key to trigger search (if not already attached)
  const enrollmentInput = document.getElementById("enrollmentNo");
  if (enrollmentInput) {
    enrollmentInput.addEventListener("keypress", function (e) {
      if (e.key === "Enter") searchStudent();
    });
  }
});
