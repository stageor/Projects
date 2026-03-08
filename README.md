<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Excel Automation Demo (Browser + Excel JS API style)</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      margin: 20px;
      background: #f8f9fa;
    }
    h1 { color: #1a3c6d; }
    .container {
      max-width: 900px;
      margin: 0 auto;
    }
    button {
      padding: 10px 18px;
      margin: 6px 8px 6px 0;
      background: #0066cc;
      color: white;
      border: none;
      border-radius: 5px;
      cursor: pointer;
      font-size: 0.95rem;
    }
    button:hover { background: #0055aa; }
    button.danger { background: #c62828; }
    button.danger:hover { background: #a51c1c; }
    textarea {
      width: 100%;
      height: 140px;
      font-family: 'Consolas', 'Courier New', monospace;
      padding: 10px;
      border: 1px solid #ccc;
      border-radius: 5px;
      background: #fff;
    }
    #result {
      margin-top: 1.5rem;
      padding: 12px;
      background: #e8f4fd;
      border-left: 5px solid #0066cc;
      border-radius: 4px;
      min-height: 80px;
    }
    table.demo {
      border-collapse: collapse;
      margin: 20px 0;
      font-size: 0.95em;
    }
    table.demo th, table.demo td {
      border: 1px solid #ddd;
      padding: 10px 14px;
    }
    table.demo th {
      background: #e3f2fd;
      color: #1a3c6d;
    }
  </style>
</head>
<body>

<div class="container">

<h1>Excel-like Automation Demo</h1>

<p>Try these typical Excel automation tasks using JavaScript in the browser:</p>

<div>
  <button onclick="generateSampleData()">1. Create Sample Sales Data</button>
  <button onclick="calculateTotals()">2. Calculate Totals & VAT</button>
  <button onclick="highlightAboveAverage()">3. Highlight Above Average</button>
  <button onclick="exportToCSV()">4. Export Table → CSV</button>
  <button class="danger" onclick="clearTable()">Clear Table</button>
</div>

<table id="dataTable" class="demo">
  <thead>
    <tr>
      <th>Product</th>
      <th>Quantity</th>
      <th>Unit Price</th>
      <th>Total</th>
      <th>VAT 18%</th>
      <th>Final Amount</th>
    </tr>
  </thead>
  <tbody id="tableBody"></tbody>
  <tfoot>
    <tr>
      <td colspan="3" style="text-align:right; font-weight:bold;">Grand Total</td>
      <td id="grandTotal" style="font-weight:bold;">—</td>
      <td id="grandVAT" style="font-weight:bold;">—</td>
      <td id="grandFinal" style="font-weight:bold;">—</td>
    </tr>
  </tfoot>
</table>

<h3>Generated JavaScript code (like Office Script / Excel automation)</h3>
<textarea id="codeOutput" readonly></textarea>

<div id="result"></div>

</div>

<script>
// ────────────────────────────────────────────────
// Helper - format number with 2 decimals
// ────────────────────────────────────────────────
function formatNum(n) {
  return Number(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// ────────────────────────────────────────────────
// 1. Generate sample data (like filling data in Excel)
// ────────────────────────────────────────────────
function generateSampleData() {
  const products = [
    ["Laptop Pro", 8, 1249.99],
    ["Wireless Mouse", 45, 24.50],
    ["USB-C Hub", 19, 38.90],
    ["27\" Monitor", 6, 289.00],
    ["Mechanical Keyboard", 12, 89.95],
    ["External SSD 1TB", 14, 109.00]
  ];

  const tbody = document.getElementById("tableBody");
  tbody.innerHTML = "";

  products.forEach(([name, qty, price]) => {
    const total = qty * price;
    const vat = total * 0.18;
    const final = total + vat;

    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${name}</td>
      <td style="text-align:right">${qty}</td>
      <td style="text-align:right">$${formatNum(price)}</td>
      <td style="text-align:right">$${formatNum(total)}</td>
      <td style="text-align:right">$${formatNum(vat)}</td>
      <td style="text-align:right">$${formatNum(final)}</td>
    `;
    tbody.appendChild(row);
  });

  document.getElementById("result").innerHTML = "Sample sales data generated (6 rows)";
  showCodeExample("generateSampleData");
}

// ────────────────────────────────────────────────
// 2. Calculate totals (like SUM, SUMIF in Excel)
// ────────────────────────────────────────────────
function calculateTotals() {
  const rows = document.querySelectorAll("#tableBody tr");
  if (rows.length === 0) {
    alert("Please generate data first");
    return;
  }

  let sumTotal = 0, sumVAT = 0, sumFinal = 0;

  rows.forEach(row => {
    const cells = row.cells;
    const total  = parseFloat(cells[3].textContent.replace(/[,$]/g,"")) || 0;
    const vat    = parseFloat(cells[4].textContent.replace(/[,$]/g,"")) || 0;
    const final  = parseFloat(cells[5].textContent.replace(/[,$]/g,"")) || 0;

    sumTotal += total;
    sumVAT   += vat;
    sumFinal += final;
  });

  document.getElementById("grandTotal").textContent = "$" + formatNum(sumTotal);
  document.getElementById("grandVAT").textContent   = "$" + formatNum(sumVAT);
  document.getElementById("grandFinal").textContent = "$" + formatNum(sumFinal);

  document.getElementById("result").innerHTML = "Totals calculated successfully";
  showCodeExample("calculateTotals");
}

// ────────────────────────────────────────────────
// 3. Conditional formatting simulation
// ────────────────────────────────────────────────
function highlightAboveAverage() {
  const rows = document.querySelectorAll("#tableBody tr");
  if (rows.length === 0) return;

  // Get all final amounts
  const finals = Array.from(rows).map(row => 
    parseFloat(row.cells[5].textContent.replace(/[,$]/g,"")) || 0
  );

  const avg = finals.reduce((a,b)=>a+b,0) / finals.length;

  rows.forEach((row,i) => {
    const val = finals[i];
    row.style.background = val > avg ? "#e8f5e9" : "";
    row.style.fontWeight = val > avg ? "600" : "normal";
  });

  document.getElementById("result").innerHTML = 
    `Highlighted rows above average (avg = $${formatNum(avg)})`;
  showCodeExample("highlightAboveAverage");
}

// ────────────────────────────────────────────────
// 4. Export to CSV (very common automation task)
// ────────────────────────────────────────────────
function exportToCSV() {
  const rows = [["Product","Quantity","Unit Price","Total","VAT 18%","Final Amount"]];
  
  document.querySelectorAll("#dataTable tr").forEach(tr => {
    const row = Array.from(tr.cells).map(td => 
      `"${td.textContent.trim().replace(/"/g,'""')}"`
    );
    rows.push(row);
  });

  const csv = rows.map(r => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  
  const link = document.createElement("a");
  link.href = url;
  link.download = "sales_export_" + new Date().toISOString().slice(0,10) + ".csv";
  link.click();
  URL.revokeObjectURL(url);

  document.getElementById("result").innerHTML = "CSV file downloaded";
}

// ────────────────────────────────────────────────
function clearTable() {
  document.getElementById("tableBody").innerHTML = "";
  document.getElementById("grandTotal").textContent = "—";
  document.getElementById("grandVAT").textContent   = "—";
  document.getElementById("grandFinal").textContent = "—";
  document.getElementById("result").innerHTML = "Table cleared";
}

// ────────────────────────────────────────────────
// Show similar code you would write in Excel Office Scripts
// ────────────────────────────────────────────────
function showCodeExample(funcName) {
  const examples = {
    generateSampleData: `// Office Scripts / Excel JS API style
async function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  let data = [
    ["Laptop Pro", 8, 1249.99],
    // ... more rows
  ];
  sheet.getRange("A2:C7").setValues(data);
  sheet.getRange("D2:D7").setFormula("=B2*C2");
  // ...`,
    
    calculateTotals: `// Office Scripts example
let totalsRange = sheet.getRange("D2:D7");
let sumTotal = totalsRange.getValues()
  .reduce((sum, row) => sum + Number(row[0]), 0);
sheet.getRange("D9").setValue(sumTotal);`,

    highlightAboveAverage: `// Conditional formatting via script
let values = sheet.getRange("F2:F7").getValues().flat();
let avg = values.reduce((a,b)=>a+b)/values.length;
sheet.getRange("A2:F7")
  .getConditionalFormats()
  .add(ExcelScript.ConditionalFormatType.cellValue)
  .getCellValue().setFormat({
    font: { bold: true },
    fill: { color: "#e8f5e9" }
  }).setRule({ formula: `F2>${avg}` });`
  };

  const codeArea = document.getElementById("codeOutput");
  codeArea.value = examples[funcName] || "// Select an action to see similar Excel script code";
}
</script>

</body>
</html>
