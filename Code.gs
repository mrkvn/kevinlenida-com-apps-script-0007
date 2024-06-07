function onOpen() {
  createMenu();
}
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Merge");
  menu.addItem("Start", "mergeSheets");
  menu.addToUi();
}

// CONSTANTS
const MERGED_SHEET_NAME = "combined"

function mergeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let combinedSheet = ss.getSheetByName(MERGED_SHEET_NAME);
  if (!combinedSheet) {
    combinedSheet = ss.insertSheet(MERGED_SHEET_NAME);
  } else {
    combinedSheet.clear();
  }

  // Collect all unique headers in a case-insensitive manner
  let allHeaders = new Set();
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    if (sheetName === MERGED_SHEET_NAME) continue;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(header => allHeaders.add(header.toLowerCase()));
  }

  // Convert the set to an array and add the "sheet_name" column
  let allHeadersArray = Array.from(allHeaders);
  allHeadersArray.push("sheet_name");

  // Initialize the combined data with headers
  let combinedData = [allHeadersArray];

  // Loop through each sheet and collect data
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    if (sheetName === MERGED_SHEET_NAME) continue;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Create a mapping of lower-case header to column index
    let headerMap = {};
    headers.forEach((header, index) => {
      headerMap[header.toLowerCase()] = index;
    });

    // Process each row
    for (let j = 1; j < data.length; j++) {
      let row = new Array(allHeadersArray.length).fill("");
      for (let k = 0; k < allHeadersArray.length - 1; k++) {
        const header = allHeadersArray[k];
        if (headerMap.hasOwnProperty(header)) {
          row[k] = data[j][headerMap[header]];
        }
      }
      row[allHeadersArray.length - 1] = sheetName; // Add the sheet name
      combinedData.push(row);
    }
  }

  // Write the combined data to the merged sheet
  combinedSheet.getRange(1, 1, combinedData.length, combinedData[0].length).setValues(combinedData);

  // Move the combined sheet to the beginning
  ss.setActiveSheet(combinedSheet);
  ss.moveActiveSheet(1);
}
