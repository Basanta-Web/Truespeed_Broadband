const excelFile = "Truesped_Broadband_2026 5.xlsx";
let tableData = [];
let hotInstance; // Keep a reference to the table instance

async function loadExcel() {
  try {
    const response = await fetch(excelFile);
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Convert to Array of Arrays
    tableData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Process Dates
    for (let r = 1; r < tableData.length; r++) {
      for (let c = 0; c < tableData[r].length; c++) {
        // Refined check: ensure it's a number and likely a date range for 2010-2030
        if (typeof tableData[r][c] === "number" && tableData[r][c] > 40000 && tableData[r][c] < 60000) {
          tableData[r][c] = excelDateToJSDate(tableData[r][c]);
        }
      }
    }
    renderTable();
  } catch (err) {
    console.error("Error loading Excel:", err);
  }
}

function excelDateToJSDate(serial) {
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return date.toLocaleDateString("en-GB");
}

function renderTable() {
  const container = document.getElementById("table");

  // 1. Initialize HyperFormula
  const hfInstance = HyperFormula.buildFromArray(tableData, {
    licenseKey: 'non-commercial-and-evaluation'
  });

  // 2. Connect HyperFormula to Handsontable
  hotInstance = new Handsontable(container, {
    data: tableData,
    formulas: {
      engine: hfInstance, // This enables cell calculations
    },
    rowHeaders: true,
    colHeaders: true,
    filters: true,
    dropdownMenu: true,
    stretchH: "all",
    height: "auto",
    licenseKey: "non-commercial-and-evaluation"
  });
}

function backupExcel() {
  // Grab the current state of the data from the table instance
  const currentData = hotInstance.getData(); 
  const ws = XLSX.utils.aoa_to_sheet(currentData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Customers");
  XLSX.writeFile(wb, "Truesped_Backup_2026.xlsx");
}

loadExcel();
