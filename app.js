const excelFile = "Truesped_Broadband_2026 5.xlsx";
let tableData = [];
let tableHeaders = [];
let hotInstance;

async function loadExcel() {
  try {
    const response = await fetch(excelFile);
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Extract raw data
    let rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // Separate Headers from Data for better sorting/filtering
    if (rawData.length > 0) {
        tableHeaders = rawData.shift(); // Pulls the first row out to use as colHeaders
    }
    
    tableData = rawData;

    // Process Dates
    for (let r = 0; r < tableData.length; r++) { // Start at 0 now that headers are gone
      for (let c = 0; c < tableData[r].length; c++) {
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

// Custom Renderer for Status Badges
function statusBadgeRenderer(instance, td, row, col, prop, value, cellProperties) {
    // 1. Let Handsontable do its default text rendering first
    Handsontable.renderers.TextRenderer.apply(this, arguments);

    // 2. If the cell is empty, stop here
    if (!value) return;

    // 3. Define our badge colors based on your dark theme
    const statusColors = {
        'active':  { bg: '#064e3b', text: '#34d399' }, // Emerald Dark/Light
        'offline': { bg: '#7f1d1d', text: '#f87171' }, // Red Dark/Light
        'pending': { bg: '#713f12', text: '#facc15' }, // Yellow Dark/Light
        'overdue': { bg: '#701a75', text: '#e879f9' }  // Fuchsia Dark/Light
    };

    // Clean up the text (lowercase, remove spaces) to match our dictionary
    const safeValue = String(value).toLowerCase().trim();
    const colors = statusColors[safeValue];

    // 4. If a match is found, transform the cell HTML into a badge
    if (colors) {
        td.innerHTML = ''; // Clear default text
        
        const badge = document.createElement('span');
        badge.innerText = value.toUpperCase();
        
        // Apply modern inline styling
        badge.style.backgroundColor = colors.bg;
        badge.style.color = colors.text;
        badge.style.padding = '4px 10px';
        badge.style.borderRadius = '12px';
        badge.style.fontSize = '0.7rem';
        badge.style.fontWeight = '600';
        badge.style.letterSpacing = '0.05em';
        badge.style.display = 'inline-block';
        
        td.appendChild(badge);
        td.style.textAlign = 'center'; // Center the badge in the cell
    }
}

function renderTable() {
  const container = document.getElementById("table");

  // Find which column number holds the "Status" data
  const statusColIndex = tableHeaders.findIndex(header => 
      typeof header === 'string' && header.toLowerCase().includes('status')
  );

  const hfInstance = HyperFormula.buildFromArray(tableData, {
    licenseKey: 'non-commercial-and-evaluation'
  });

  hotInstance = new Handsontable(container, {
    data: tableData,
    colHeaders: tableHeaders, // Inject the extracted headers
    rowHeaders: true,
    formulas: { engine: hfInstance },
    
    // UI Enhancements
    columnSorting: true,      // Clicking a header now actually sorts!
    filters: true,            // Adds filter dropdowns to headers
    dropdownMenu: true,
    autoColumnSize: true,
    stretchH: "all",
    
    // Apply our custom renderer only to the Status column
    cells: function(row, col) {
        const cellProperties = {};
        if (col === statusColIndex) {
            cellProperties.renderer = statusBadgeRenderer;
        }
        return cellProperties;
    },
    
    licenseKey: "non-commercial-and-evaluation"
  });
}

function backupExcel() {
  // Re-attach the headers to the current data before exporting
  const currentData = hotInstance.getData(); 
  const dataWithHeaders = [tableHeaders, ...currentData];
  
  const ws = XLSX.utils.aoa_to_sheet(dataWithHeaders);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Customers");
  XLSX.writeFile(wb, "Truesped_Backup_2026.xlsx");
}

loadExcel();
