const excelURL =
"https://api.onedrive.com/v1.0/shares/u!aHR0cHM6Ly8xZHJ2Lm1zL3gvYy9lY2I2YWJiMDkyM2FiZDg0L0lRRDhKcENDTms4NFFLTGxOaThWdjZtYUFhVUNnRVNYYjA3V0ZXdndROU03NXZvP2U9aWIwVm5a/root/content"

let tableData = []

async function loadExcel(){

const response = await fetch(excelURL)

const buffer = await response.arrayBuffer()

const workbook = XLSX.read(buffer,{type:"array"})

const sheet = workbook.Sheets[workbook.SheetNames[0]]

tableData = XLSX.utils.sheet_to_json(sheet,{header:1})

renderTable()

}

function renderTable(){

const container = document.getElementById("table")

HyperFormula.buildFromArray(tableData)

new Handsontable(container,{

data:tableData,
rowHeaders:true,
colHeaders:true,
stretchH:"all",
height:"100%",
licenseKey:"non-commercial-and-evaluation"

})

}

function backupExcel(){

const ws = XLSX.utils.aoa_to_sheet(tableData)

const wb = XLSX.utils.book_new()

XLSX.utils.book_append_sheet(wb,ws,"Customers")

XLSX.writeFile(wb,"backup.xlsx")

}

loadExcel()