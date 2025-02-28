# Google Sheets Clone

## Overview
This project is a web application that mimics the core functionalities and user interface of Google Sheets. It allows users to enter data, apply basic formatting, and compute formulas such as **SUM()** and **AVERAGE()**. The project is built using plain HTML, CSS, and JavaScript.

## 📸 Screenshots
![Screenshot 2025-02-28 153840](https://github.com/user-attachments/assets/57d8550e-73c1-4db0-948e-a72870f33ba9)

## Directory Structure
```sh
GoogleSheetsClone/ 
├── index.html
├── css/ 
│     └── style.css 
└── js/ 
     └── main.js
```


## How to Run
1. **Clone or Download the Repository:**  
   Make sure all files are in their respective directories as shown above.
2. **Open in Browser:**  
   Simply open the `index.html` file in your favorite web browser.
3. **Interact with the Spreadsheet:**  
   - Enter values or formulas in cells.
   - Use the toolbar for formatting and to add/delete rows or columns.
   - Test functions such as `=SUM(A1:A5)` or `=AVERAGE(A1:A5)`.

## Key Functions in main.js

### Spreadsheet Initialization & UI Interactions
- **`initSpreadsheet()`**  
  - **Use:** Dynamically creates the spreadsheet table with headers and cells based on preset row and column counts.
  
- **`cellFocus(e)`**  
  - **Use:** Highlights the selected cell and loads its stored formula or value into the formula bar when a cell is focused.
  
- **`cellInput(e)`**  
  - **Use:** Detects user input in a cell. If the input starts with an "=" sign, it treats it as a formula; otherwise, it stores a literal value.
  
- **`recalcAll()`**  
  - **Use:** Iterates over all cells, recalculates their values based on any stored formulas, and updates the display accordingly.

### Formula Evaluation & Mathematical Functions
- **`evaluateFormula(formula)`**  
  - **Use:** Processes a formula string and returns its computed value.  
  - **Details:**  
    - Supports basic arithmetic, cell references (e.g., `A1+B2`), and functions such as:
      - **`SUM(range)`**: Computes the sum of a range of cells.
      - **`AVERAGE(range)`**: Computes the average value of a range of cells.
  
- **`getRangeValues(rangeStr)`**  
  - **Use:** Returns an array of cell values from a specified range (e.g., `"A1:B2"`).  
  - **Purpose:** Utilized by functions like `SUM()` and `AVERAGE()` to fetch data for calculation.
  
- **`getCellValue(cellId)`**  
  - **Use:** Retrieves the value stored in a given cell (e.g., `"B3"`).

### Cell Coordinate Conversion
- **`cellIdToCoords(cellId)`**  
  - **Use:** Converts a cell identifier (like `"B3"`) into numerical row and column indices.
  
- **`letterToCol(letters)`**  
  - **Use:** Converts a column letter (e.g., `"A"`) to its corresponding zero-based index.
  
- **`colToLetter(col)`**  
  - **Use:** Converts a zero-based column index back into a column letter for display.

### Toolbar and Interaction Handlers
- **`addToolbarListeners()`**  
  - **Use:** Sets up event listeners for toolbar actions (such as bold, italic, font size, text color, adding/deleting rows and columns, saving, and loading).
  
- **`addDragListeners()`**  
  - **Use:** Enables the drag-selection of cells, mimicking Google Sheets’ behavior.

### Dynamic Table Management
- **`addRow()` & `addColumn()`**  
  - **Use:** Dynamically add rows or columns to the spreadsheet.
  
- **`deleteRow()` & `deleteColumn()`**  
  - **Use:** Remove the currently selected row or column from the spreadsheet.
  
- **`rebuildTable()`**  
  - **Use:** Rebuilds table headers and reassigns cell IDs after modifications such as row/column deletion.

## Additional Features
- **Local Storage Integration:**  
  - **Save:** Users can save the current state of the spreadsheet to their browser's local storage.
  - **Load:** Previously saved spreadsheets can be loaded back into the application.
  
- **Cell Formatting:**  
  - Toolbar buttons allow for applying styles such as bold, italic, changing font size, and text color.

## Conclusion
This Google Sheets Clone serves as a foundation for a dynamic spreadsheet web application. It demonstrates key concepts such as dynamic table generation, event handling, formula evaluation, and user interface enhancements. Feel free to extend its functionality.

---
## Contact
For any questions or feedback, feel free to reach out to [yuvrajsogarwal@gmail.com].


