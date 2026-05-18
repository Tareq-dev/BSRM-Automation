// Excel package
const XLSX = require("xlsx");

// ========================================
// MODULE 1
// STR / AE Replace
// ========================================

async function processReplace() {
  // Progress function
  function updateProgress(value) {
    document.getElementById("progressBar").style.width = `${value}%`;
  }
  // show progress bar
  document.getElementById("progressContainer").style.display = "block";
  // Start progress
  updateProgress(10);

  // File নিচ্ছে
  const targetFile = document.getElementById("targetFile").files[0];

  // File check
  if (!targetFile) {
    alert("Please select a file");

    return;
  }

  // STR input
  const strValue = document.getElementById("strInput").value.trim();

  // AE input
  const aeValue = document.getElementById("aeInput").value.trim();

  // Validation
  if (!strValue || !aeValue) {
    alert("Please fill all inputs");

    return;
  }

  // Progress
  updateProgress(20);

  // FileReader
  const reader = new FileReader();

  // File load
  reader.onload = async function (e) {
    // File data
    const data = e.target.result;

    // Workbook
    const workbook = XLSX.read(data, {
      type: "array",
    });

    // Progress
    updateProgress(40);

    // Sheet
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Sheet → JSON
    const jsonData = XLSX.utils.sheet_to_json(
      sheet,

      {
        header: 1,
      },
    );

    // Loop rows
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];

      // ==========================
      // B COLUMN
      // ==========================

      const bCell = row[1];

      if (typeof bCell === "string") {
        if (bCell.startsWith("STR-")) {
          row[1] = strValue;
        }
      }

      // ==========================
      // AE COLUMN
      // ==========================

      const aeCell = row[30];

      if (typeof aeCell === "string") {
        if (aeCell.includes("x")) {
          row[30] = aeValue;
        }
      }
    }

    // Progress
    updateProgress(60);

    // Progress
    updateProgress(80);

    // =================================
    // NEW SHEET
    // =================================

    const newSheet = XLSX.utils.aoa_to_sheet(jsonData);

    workbook.Sheets[workbook.SheetNames[0]] = newSheet;

    // =================================
    // SAVE LOCATION
    // =================================

    const { dialog } = require("@electron/remote");

    const savePath = await dialog.showSaveDialog({
      defaultPath: `${strValue}.csv`,

      filters: [
        {
          name: "CSV File",

          extensions: ["csv"],
        },
      ],
    });

    // যদি cancel করে
    if (savePath.canceled) {
      document.getElementById("status").innerHTML = "❌ Save cancelled";

      return;
    }

    // =================================
    // SAVE FILE
    // =================================

    XLSX.writeFile(
      workbook,

      savePath.filePath,

      {
        bookType: "csv",
      },
    );

    // Progress
    updateProgress(100);
    // hide progress bar after completion
    setTimeout(() => {
      document.getElementById("progressContainer").style.display = "none";
      updateProgress(0);
    }, 1000);
    // =================================
    // SUCCESS MESSAGE
    // =================================

    document.getElementById("status").innerHTML = `File Saved Successfully!`;

    // RESET INPUTS
    document.getElementById("strInput").value = "";
    document.getElementById("aeInput").value = "";
    document.getElementById("targetFile").value = "";
  };

  // File read
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

    document.getElementById("strInput").value = nameWithoutExt;

    document.getElementById("aeInput").value = "xmorshed";
  });
});
// ========================================
// MODULE 2
// Ingredient Calculation
// ========================================

async function processCalculation() {

  // ==============================
  // FILE INPUTS
  // ==============================

  const sourceFile =
    document.getElementById(
      "calcSource"
    ).files[0];

  const targetFile =
    document.getElementById(
      "calcTarget"
    ).files[0];



  // validation
  if (!sourceFile || !targetFile) {

    alert("Please select both files");

    return;
  }



  // ==============================
  // READ SOURCE FILE
  // ==============================

  const sourceData =
    await sourceFile.arrayBuffer();

  const sourceWorkbook =
    XLSX.read(sourceData, {
      type: "array",
    });

  const sourceSheet =
    sourceWorkbook.Sheets[
      sourceWorkbook.SheetNames[0]
    ];



  // ==============================
  // READ TARGET FILE
  // ==============================

  const targetData =
    await targetFile.arrayBuffer();

  const targetWorkbook =
    XLSX.read(targetData, {
      type: "array",
    });

  const targetSheet =
    targetWorkbook.Sheets[
      targetWorkbook.SheetNames[0]
    ];



  // ==============================
  // SHEET → JSON
  // ==============================

  const sourceJson =
    XLSX.utils.sheet_to_json(
      sourceSheet,
      { header: 1 }
    );

  const targetJson =
    XLSX.utils.sheet_to_json(
      targetSheet,
      { header: 1 }
    );



  // ==============================
  // PRODUCT ROWS COPY
  // ==============================

  const productRows = [];



  for (let row of sourceJson) {

    // D column = index 3
    const dCell = row[3];



    if (
      typeof dCell === "string" &&
      dCell.includes("PRODUCT")
    ) {

      // পুরো row copy
      productRows.push(row);
    }
  }



  // ==============================
  // FIND INGREADIENTS ROW
  // ==============================

  let insertIndex = -1;



  for (let i = 0; i < targetJson.length; i++) {

    const row = targetJson[i];

    const dCell = row[3];



    if (
      typeof dCell === "string" &&
      dCell.includes("INGREADIENTS")
    ) {

      insertIndex = i;

      break;
    }
  }



  // যদি INGREADIENTS না পায়
  if (insertIndex === -1) {

    alert("INGREADIENTS row not found");

    return;
  }



  // ==============================
  // INSERT PRODUCT ROWS
  // ==============================

  targetJson.splice(
    insertIndex,
    0,
    ...productRows
  );



  // ==============================
  // NEW SHEET
  // ==============================

  const newSheet =
    XLSX.utils.aoa_to_sheet(
      targetJson
    );

  targetWorkbook.Sheets[
    targetWorkbook.SheetNames[0]
  ] = newSheet;



  // ==============================
  // FILE NAME FROM B COLUMN
  // ==============================

  let fileName = "output.csv";



  for (let row of targetJson) {

    // B column = index 1
    const bCell = row[1];



    if (
      typeof bCell === "string" &&
      bCell.startsWith("STR-")
    ) {

      fileName = `${bCell}.csv`;

      break;
    }
  }



  // ==============================
  // SAVE DIALOG
  // ==============================

  const { dialog } =
    require("@electron/remote");



  const savePath =
    await dialog.showSaveDialog({

      defaultPath: fileName,

      filters: [
        {
          name: "CSV File",
          extensions: ["csv"],
        },
      ],
    });



  // cancel check
  if (savePath.canceled) {

    document.getElementById(
      "status1"
    ).innerText =
      "❌ Save cancelled";

    return;
  }



  // ==============================
  // SAVE FILE
  // ==============================

  XLSX.writeFile(
    targetWorkbook,
    savePath.filePath,
    {
      bookType: "csv",
    }
  );



  // ==============================
  // SUCCESS MESSAGE
  // ==============================

  document.getElementById(
    "status1"
  ).innerText =
    "✅ Calculation Complete!";
}



// ========================================
// PIN SYSTEM
// ========================================

function checkPin() {

  // input pin নিচ্ছে
  const pin =
    document.getElementById(
      "pinInput"
    ).value;


  // correct pin
  const correctPin = "0000";


  // check করছে
  if (pin === correctPin) {

    // PIN screen hide
    document.getElementById(
      "pinScreen"
    ).style.display = "none";


    // Main app show
    document.getElementById(
      "mainApp"
    ).style.display = "block";
  }

  else {

    // Error message
    document.getElementById(
      "pinError"
    ).innerText =
      "Wrong PIN";


    // INPUT CLEAR
    document.getElementById(
      "pinInput"
    ).value = "";
  }
}