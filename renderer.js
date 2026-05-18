// Excel package
const XLSX = require("xlsx");

// ========================================
// MODULE 1
// STR / AE Replace
// ========================================

async function processReplace() {
  function updateProgress(value) {
    document.getElementById("progressBar").style.width = `${value}%`;
  }

  document.getElementById("progressContainer").style.display = "block";
  updateProgress(10);

  const targetFile = document.getElementById("targetFile").files[0];

  if (!targetFile) {
    alert("Please select a file");
    return;
  }

  const strInput = document.getElementById("strInput");
  const aeInput = document.getElementById("aeInput");

  const strValue = strInput.value.trim();
  const aeValue = aeInput.value.trim();

  if (!strValue || !aeValue) {
    alert("Please fill all inputs");
    return;
  }

  // ==========================
  // STR +1 LOGIC
  // ==========================
  const match = strValue.match(/(\d+)$/);

  let nextStrValue = strValue;

  updateProgress(20);

  const reader = new FileReader();

  reader.onload = async function (e) {
    const data = e.target.result;

    const workbook = XLSX.read(data, {
      type: "array",
    });

    updateProgress(40);

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const jsonData = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
    });

    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];

      // B COLUMN
      const bCell = row[1];
      if (typeof bCell === "string") {
        if (bCell.startsWith("STR-")) {
          row[1] = strValue;
        }
      }

      // AE COLUMN
      const aeCell = row[30];
      if (typeof aeCell === "string") {
        if (aeCell.includes("x")) {
          row[30] = aeValue;
        }
      }
    }

    updateProgress(60);

    const newSheet = XLSX.utils.aoa_to_sheet(jsonData);
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;

    const { dialog } = require("@electron/remote");

    const savePath = await dialog.showSaveDialog({
      defaultPath: `${nextStrValue}.csv`,
      filters: [
        {
          name: "CSV File",
          extensions: ["csv"],
        },
      ],
    });

    if (savePath.canceled) {
      document.getElementById("status").innerHTML = "Save cancelled";
      return;
    }

    XLSX.writeFile(workbook, savePath.filePath, {
      bookType: "csv",
    });

    updateProgress(100);

    setTimeout(() => {
      document.getElementById("progressContainer").style.display = "none";
      updateProgress(0);
    }, 1000);

    document.getElementById("status").innerHTML = "File Saved Successfully!";

    // reset other inputs (BUT STR now already updated)
    aeInput.value = "";
    document.getElementById("targetFile").value = "";
  };

  reader.readAsArrayBuffer(targetFile);
}
// ========================================
// AUTO FILL INPUTS FROM FILE NAME
// ========================================

window.addEventListener("DOMContentLoaded", () => {
  document.getElementById("targetFile").addEventListener("change", function () {
    const file = this.files[0];
    if (!file) return;

    const nameWithoutExt = file.name.replace(/\.[^/.]+$/, "");

    const lastNumber = Number(nameWithoutExt.split("-")[1]);
    const incremented = Number(lastNumber) + 1;

    // 🔥 IMPORTANT: input field update
    const strInput = document.getElementById("strInput");
    strInput.value = `${"STR-" + incremented}`;

    document.getElementById("aeInput").value = "xmorshed";
  });
});
// ========================================
// MODULE 2
// Ingredient Calculation
// ========================================

async function processCalculation() {
  // =====================================
  // FILE INPUTS
  // =====================================

  const file1 = document.getElementById("calcSource").files[0];

  const file2 = document.getElementById("calcTarget").files[0];

  const file3 = document.getElementById("finalFile").files[0];

  // validation
  if (!file1 || !file2 || !file3) {
    alert("Please select all 3 files");

    return;
  }

  // =====================================
  // READ FILES
  // =====================================

  const data1 = await file1.arrayBuffer();

  const data2 = await file2.arrayBuffer();

  const data3 = await file3.arrayBuffer();

  const wb1 = XLSX.read(data1, {
    type: "array",
  });

  const wb2 = XLSX.read(data2, {
    type: "array",
  });

  const wb3 = XLSX.read(data3, {
    type: "array",
  });

  // =====================================
  // SHEETS
  // =====================================

  const sheet1 = wb1.Sheets[wb1.SheetNames[0]];

  const sheet2 = wb2.Sheets[wb2.SheetNames[0]];

  const sheet3 = wb3.Sheets[wb3.SheetNames[0]];

  // =====================================
  // JSON
  // =====================================

  const json1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });

  const json2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

  let json3 = XLSX.utils.sheet_to_json(sheet3, { header: 1 });

  // =====================================
  // COPY DATA
  // =====================================

  const ingredientsRows = [];
  const byProductRows = [];
  const productRows = [];

  // =====================================
  // FILE 1
  // INGREADIENTS + BY PRODUCT
  // =====================================

  for (let row of json1) {
    const dCell = row[3];

    if (typeof dCell === "string") {
      if (dCell.includes("INGREADIENTS")) {
        ingredientsRows.push([...row]);
      } else if (dCell.includes("BY PRODUCT")) {
        byProductRows.push([...row]);
      }
    }
  }

  // =====================================
  // FILE 2
  // PRODUCT
  // =====================================

  for (let row of json2) {
    const dCell = row[3];

    if (typeof dCell === "string" && dCell.includes("PRODUCT")) {
      productRows.push([...row]);
    }
  }

  // =====================================
  // MERGE ALL ROWS
  // =====================================

  const finalRows = [...ingredientsRows, ...byProductRows, ...productRows];

  // =====================================
  // REMOVE EMPTY ROWS
  // =====================================

  json3 = json3.filter((row) => {
    return row.some((cell) => {
      return cell !== undefined && cell !== "";
    });
  });

  // =====================================
  // INSERT FROM ROW 2
  // =====================================

  // header row রেখে
  const headerRow = json3[0];

  // data
  const remainingRows = json3.slice(1);

  // new final data
  json3 = [headerRow, ...finalRows, ...remainingRows];

  // =====================================
  // E COLUMN SERIAL
  // =====================================

  let serial = 1;

  // Row 2 থেকে loop
  for (let i = 1; i < json3.length; i++) {
    // D column
    const dCell = json3[i][3];

    // যদি INGREADIENTS হয়
    if (typeof dCell === "string" && dCell.includes("INGREADIENTS")) {
      // E column = index 4
      json3[i][4] = serial;

      // next serial
      serial++;
    }
  }

  // =====================================
  // H COLUMN FIXED DECIMAL
  // =====================================

  for (let i = 0; i < json3.length; i++) {
    const hCell = json3[i][7];

    if (hCell !== undefined && !isNaN(hCell)) {
      json3[i][7] = Number(hCell).toFixed(5);
    }
  }

  // =====================================
  // CREATE NEW SHEET
  // =====================================

  const newSheet = XLSX.utils.aoa_to_sheet(json3);

  wb3.Sheets[wb3.SheetNames[0]] = newSheet;

  // =====================================
  // FILE NAME FROM B COLUMN
  // =====================================

  let outputName = "output.csv";

  for (let row of json3) {
    const bCell = row[1];

    if (typeof bCell === "string" && bCell.startsWith("STR-")) {
      outputName = `${bCell}.csv`;

      break;
    }
  }

  // =====================================
  // SAVE DIALOG
  // =====================================

  const { dialog } = require("@electron/remote");

  const savePath = await dialog.showSaveDialog({
    defaultPath: outputName,

    filters: [
      {
        name: "CSV File",
        extensions: ["csv"],
      },
    ],
  });

  // cancel
  if (savePath.canceled) {
    document.getElementById("status1").innerText = "❌ Save cancelled";

    return;
  }

  // =====================================
  // SAVE FILE
  // =====================================

  XLSX.writeFile(wb3, savePath.filePath, {
    bookType: "csv",
  });

  // =====================================
  // SUCCESS
  // =====================================

  document.getElementById("status1").innerText = "Calculation Complete!";

  // =====================================
  // RESET INPUTS
  // =====================================

  document.getElementById("calcSource").value = "";

  document.getElementById("calcTarget").value = "";

  document.getElementById("finalFile").value = "";
}

// ========================================
// PIN SYSTEM
// ========================================

function checkPin() {
  // input pin নিচ্ছে
  const pin = document.getElementById("pinInput").value;

  // correct pin
  const correctPin = "0000";

  // check করছে
  if (pin === correctPin) {
    // PIN screen hide
    document.getElementById("pinScreen").style.display = "none";

    // Main app show
    document.getElementById("mainApp").style.display = "block";
  } else {
    // Error message
    document.getElementById("pinError").innerText = "Wrong PIN";

    // INPUT CLEAR
    document.getElementById("pinInput").value = "";
  }
}
