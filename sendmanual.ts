function filterComplexCriteria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Get today's date in MM/DD/YYYY format
  var today = new Date();
  var dateString = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();

  // Construct the new sheet name
  var newSheetName = 'Send Manual ' + dateString;
  
  // Try to get the sheet with the new name or create it if it doesn't exist
  var filteredSheet = ss.getSheetByName(newSheetName);
  if (filteredSheet) {
    // Clear the existing sheet if it already exists
    filteredSheet.clear();
  } else {
    // If no such sheet exists, create a new one
    filteredSheet = ss.insertSheet(newSheetName);
  }
  
  // Prepare filtered data array
  var filteredData = [data[0]]; // Initialize with headers

  // Column indices based on your data setup; adjust these as necessary
  var bpNameIndex = 3; // BP Name column
  var sunFrogSendOrderIndex = 6; // SunFrog: Send Order column
  var artOnOrderLevelIndex1 = 8; // First ART ON ORDER LEVEL column index
  var artOnOrderLevelIndex2 = 9; // Second ART ON ORDER LEVEL column index
  var artOnBomLevelIndex = 7; // ART ON BOM LEVEL column index

  // Filter and store data
  for (var i = 1; i < data.length; i++) {
    var bpName = data[i][bpNameIndex];
    var sunFrogSendOrder = data[i][sunFrogSendOrderIndex];
    var artOnOrderLevel1 = data[i][artOnOrderLevelIndex1] || "";
    var artOnOrderLevel2 = data[i][artOnOrderLevelIndex2] || "";
    var artOnBomLevel = data[i][artOnBomLevelIndex].trim(); // Trim spaces for accurate blank check

    if ((bpName === 'Fanatics' || bpName === 'One Time Shopify Customer' || bpName === 'TS - lax.com') &&
        sunFrogSendOrder !== 'Y' &&
        !artOnOrderLevel1.includes('YCIS') &&
        !artOnOrderLevel2.includes('YCIS') &&
        artOnBomLevel !== "") { // Ensure ART ON BOM LEVEL is not blank
      filteredData.push(data[i]);
    }
  }

  // Append all filtered data at once if there is data to append
  if (filteredData.length > 1) {
    filteredSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  } else {
    Logger.log("No data to display after filtering.");
  }
}
