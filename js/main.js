// main.js

// Global Variables
let numRows = 20,
  numCols = 10,
  cellData = {},
  selectedCell = null,
  isDragging = false,
  startCell = null;

// Initialize the spreadsheet when the DOM is ready
document.addEventListener("DOMContentLoaded", () => {
  initSpreadsheet();
  addToolbarListeners();
  addDragListeners();
});

// Build the initial spreadsheet table
function initSpreadsheet() {
  const table = document.getElementById("spreadsheet");
  table.innerHTML = "";

  // Create header row (first cell empty, then column letters)
  let headerRow = document.createElement("tr");
  let corner = document.createElement("th");
  headerRow.appendChild(corner);
  for (let j = 0; j < numCols; j++) {
    let th = document.createElement("th");
    th.innerText = colToLetter(j);
    headerRow.appendChild(th);
  }
  table.appendChild(headerRow);

  // Create data rows
  for (let i = 0; i < numRows; i++) {
    let row = document.createElement("tr");
    let rowHeader = document.createElement("th");
    rowHeader.innerText = i + 1;
    row.appendChild(rowHeader);
    for (let j = 0; j < numCols; j++) {
      let td = document.createElement("td");
      let cellId = colToLetter(j) + (i + 1);
      td.setAttribute("data-cell", cellId);
      td.contentEditable = true;
      td.addEventListener("focus", cellFocus);
      td.addEventListener("blur", cellBlur);
      td.addEventListener("input", cellInput);
      row.appendChild(td);
      // Initialize cell storage
      cellData[cellId] = { formula: "", value: "", format: {} };
    }
    table.appendChild(row);
  }
}

// Cell focus handler: highlight cell and load formula/value into the formula bar
function cellFocus(e) {
  selectedCell = e.target;
  const cellId = selectedCell.getAttribute("data-cell");
  document.getElementById("formula-input").value =
    cellData[cellId].formula ? "=" + cellData[cellId].formula : selectedCell.innerText;
  clearSelection();
  selectedCell.classList.add("selected");
}

function cellBlur(e) {
  // Optional: Handle any actions on blur
}

// Handle cell input: detect if a formula (starts with "=") and recalc all cells
function cellInput(e) {
  const cellId = e.target.getAttribute("data-cell");
  let content = e.target.innerText;
  if (content.startsWith("=")) {
    // Store formula without the "=" sign
    cellData[cellId].formula = content.substring(1);
  } else {
    cellData[cellId].formula = "";
    cellData[cellId].value = content;
  }
  recalcAll();
}

// Recalculate all cells (update formulas to computed values)
function recalcAll() {
  for (let cellId in cellData) {
    if (cellData[cellId].formula) {
      let computed = evaluateFormula(cellData[cellId].formula);
      cellData[cellId].value = computed;
    }
    // Update the cell's display if it exists in the DOM
    let cellElem = document.querySelector(`[data-cell="${cellId}"]`);
    if (cellElem) {
      cellElem.innerText = cellData[cellId].value;
    }
  }
}

// Evaluate a formula string
function evaluateFormula(formula) {
    formula = formula.toUpperCase();
  
    // Check for SUM function
    if (formula.startsWith("SUM(") && formula.endsWith(")")) {
      let range = formula.substring(4, formula.length - 1); // Extract range (e.g., "A1:A5")
      let values = getRangeValues(range);
      let sum = values.reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
      return sum;
    }
  
    // Check for AVERAGE function
    if (formula.startsWith("AVERAGE(") && formula.endsWith(")")) {
      let range = formula.substring(8, formula.length - 1); // Extract range (e.g., "A1:A5")
      let values = getRangeValues(range);
      let sum = values.reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
      return values.length > 0 ? sum / values.length : 0; // Avoid division by zero
    }
  
    // Simple cell reference (e.g., "A1")
    if (/^[A-Z]+\d+$/.test(formula.trim())) {
      return getCellValue(formula.trim());
    }
  
    // Replace cell references in arithmetic expressions (e.g., "A1+B2")
    let processedFormula = formula.replace(/([A-Z]+\d+)/g, function (match) {
      return getCellValue(match) || 0;
    });
  
    try {
      return eval(processedFormula);
    } catch (e) {
      return "#ERROR";
    }
  }
  

// Get all values within a given range (e.g., "A1:B2")
function getRangeValues(rangeStr) {
  let parts = rangeStr.split(":");
  if (parts.length !== 2) return [];
  let start = parts[0].trim();
  let end = parts[1].trim();
  let startCoords = cellIdToCoords(start);
  let endCoords = cellIdToCoords(end);
  let values = [];
  for (
    let i = Math.min(startCoords.row, endCoords.row);
    i <= Math.max(startCoords.row, endCoords.row);
    i++
  ) {
    for (
      let j = Math.min(startCoords.col, endCoords.col);
      j <= Math.max(startCoords.col, endCoords.col);
      j++
    ) {
      let cellId = colToLetter(j) + (i + 1);
      values.push(getCellValue(cellId));
    }
  }
  return values;
}

// Retrieve the value stored in a cell
function getCellValue(cellId) {
  if (cellData[cellId]) {
    return cellData[cellId].value || "";
  }
  return "";
}

// Convert a cell ID like "B3" to its row and column indices
function cellIdToCoords(cellId) {
  let match = cellId.match(/([A-Z]+)(\d+)/);
  if (match) {
    let colLetters = match[1];
    let row = parseInt(match[2], 10) - 1;
    let col = letterToCol(colLetters);
    return { col, row };
  }
  return { col: 0, row: 0 };
}

// Convert letters (e.g., "A", "B") to a zero-based column index
function letterToCol(letters) {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

// Convert a zero-based column index to a letter (e.g., 0 -> "A")
function colToLetter(col) {
  let letter = "";
  col++;
  while (col > 0) {
    let rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Add event listeners for the toolbar buttons and inputs
function addToolbarListeners() {
  // Apply formula/value from the formula bar to the selected cell
  document.getElementById("apply-formula-btn").addEventListener("click", () => {
    if (selectedCell) {
      let cellId = selectedCell.getAttribute("data-cell");
      let formula = document.getElementById("formula-input").value;
      if (formula.startsWith("=")) {
        cellData[cellId].formula = formula.substring(1);
      } else {
        cellData[cellId].formula = "";
        cellData[cellId].value = formula;
      }
      recalcAll();
    }
  });

  // Toggle Bold styling
  document.getElementById("bold-btn").addEventListener("click", () => {
    if (selectedCell) {
      selectedCell.style.fontWeight =
        selectedCell.style.fontWeight === "bold" ? "normal" : "bold";
    }
  });

  // Toggle Italic styling
  document.getElementById("italic-btn").addEventListener("click", () => {
    if (selectedCell) {
      selectedCell.style.fontStyle =
        selectedCell.style.fontStyle === "italic" ? "normal" : "italic";
    }
  });

  // Change Font Size
  document.getElementById("font-size-selector").addEventListener("change", (e) => {
    if (selectedCell) {
      selectedCell.style.fontSize = e.target.value;
    }
  });

  // Change Text Color
  document.getElementById("color-picker").addEventListener("change", (e) => {
    if (selectedCell) {
      selectedCell.style.color = e.target.value;
    }
  });

  // Add Row
  document.getElementById("add-row-btn").addEventListener("click", addRow);
  // Add Column
  document.getElementById("add-col-btn").addEventListener("click", addColumn);
  // Delete Row
  document.getElementById("delete-row-btn").addEventListener("click", deleteRow);
  // Delete Column
  document.getElementById("delete-col-btn").addEventListener("click", deleteColumn);

  // Save spreadsheet to localStorage
  document.getElementById("save-btn").addEventListener("click", () => {
    localStorage.setItem("spreadsheetData", JSON.stringify(cellData));
    alert("Spreadsheet saved!");
  });

  // Load spreadsheet from localStorage
  document.getElementById("load-btn").addEventListener("click", () => {
    let data = localStorage.getItem("spreadsheetData");
    if (data) {
      cellData = JSON.parse(data);
      recalcAll();
      alert("Spreadsheet loaded!");
    }
  });
}

// Add a new row to the spreadsheet
function addRow() {
  numRows++;
  const table = document.getElementById("spreadsheet");
  let row = document.createElement("tr");
  let rowHeader = document.createElement("th");
  rowHeader.innerText = numRows;
  row.appendChild(rowHeader);
  for (let j = 0; j < numCols; j++) {
    let td = document.createElement("td");
    let cellId = colToLetter(j) + numRows;
    td.setAttribute("data-cell", cellId);
    td.contentEditable = true;
    td.addEventListener("focus", cellFocus);
    td.addEventListener("blur", cellBlur);
    td.addEventListener("input", cellInput);
    row.appendChild(td);
    cellData[cellId] = { formula: "", value: "", format: {} };
  }
  table.appendChild(row);
}

// Add a new column to the spreadsheet
function addColumn() {
  numCols++;
  const table = document.getElementById("spreadsheet");
  // Update header row
  let headerRow = table.rows[0];
  let th = document.createElement("th");
  th.innerText = colToLetter(numCols - 1);
  headerRow.appendChild(th);

  // Update each data row
  for (let i = 1; i < table.rows.length; i++) {
    let row = table.rows[i];
    let cellId = colToLetter(numCols - 1) + i;
    let td = document.createElement("td");
    td.setAttribute("data-cell", cellId);
    td.contentEditable = true;
    td.addEventListener("focus", cellFocus);
    td.addEventListener("blur", cellBlur);
    td.addEventListener("input", cellInput);
    row.appendChild(td);
    cellData[cellId] = { formula: "", value: "", format: {} };
  }
}

// Delete the row of the currently selected cell
function deleteRow() {
  if (selectedCell) {
    let cellId = selectedCell.getAttribute("data-cell");
    let rowNumber = parseInt(cellId.match(/\d+/)[0]);
    const table = document.getElementById("spreadsheet");
    if (rowNumber < table.rows.length) {
      table.deleteRow(rowNumber);
      numRows--;
      rebuildTable();
    }
  }
}

// Delete the column of the currently selected cell
function deleteColumn() {
  if (selectedCell) {
    let cellId = selectedCell.getAttribute("data-cell");
    let colLetter = cellId.match(/[A-Z]+/)[0];
    let colIndex = letterToCol(colLetter);
    const table = document.getElementById("spreadsheet");
    // Delete header cell (offset by 1)
    table.rows[0].deleteCell(colIndex + 1);
    // Delete cell from each data row
    for (let i = 1; i < table.rows.length; i++) {
      table.rows[i].deleteCell(colIndex + 1);
    }
    numCols--;
    rebuildTable();
  }
}

// Rebuild the table headers and cell IDs after deletions
function rebuildTable() {
  const table = document.getElementById("spreadsheet");
  let headerRow = table.rows[0];
  for (let j = 1; j < headerRow.cells.length; j++) {
    headerRow.cells[j].innerText = colToLetter(j - 1);
  }
  for (let i = 1; i < table.rows.length; i++) {
    let row = table.rows[i];
    row.cells[0].innerText = i;
    for (let j = 1; j < row.cells.length; j++) {
      let newId = colToLetter(j - 1) + i;
      row.cells[j].setAttribute("data-cell", newId);
      if (!cellData[newId])
        cellData[newId] = { formula: "", value: "", format: {} };
    }
  }
}

// Add basic drag selection support
function addDragListeners() {
  const table = document.getElementById("spreadsheet");
  table.addEventListener("mousedown", (e) => {
    if (e.target.tagName === "TD") {
      isDragging = true;
      startCell = e.target;
      clearSelection();
      e.target.classList.add("selected");
    }
  });
  table.addEventListener("mouseover", (e) => {
    if (isDragging && e.target.tagName === "TD") {
      clearSelection();
      selectRange(startCell, e.target);
    }
  });
  document.addEventListener("mouseup", () => {
    isDragging = false;
  });
}

// Clear all selected cell highlights
function clearSelection() {
  let cells = document.querySelectorAll("td.selected");
  cells.forEach((cell) => cell.classList.remove("selected"));
}

// Highlight all cells in the rectangular area from start to end
function selectRange(start, end) {
  let startId = start.getAttribute("data-cell");
  let endId = end.getAttribute("data-cell");
  let startCoords = cellIdToCoords(startId);
  let endCoords = cellIdToCoords(endId);
  let minRow = Math.min(startCoords.row, endCoords.row);
  let maxRow = Math.max(startCoords.row, endCoords.row);
  let minCol = Math.min(startCoords.col, endCoords.col);
  let maxCol = Math.max(startCoords.col, endCoords.col);
  for (let i = minRow; i <= maxRow; i++) {
    for (let j = minCol; j <= maxCol; j++) {
      let cell = document.querySelector(
        `[data-cell="${colToLetter(j) + (i + 1)}"]`
      );
      if (cell) cell.classList.add("selected");
    }
  }
}
