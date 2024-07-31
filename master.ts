function runAllFunctions() {
  filterData();
  filterComplexCriteria();
  moveDesignRows();
  filterOneTimeShopYCIS();
}

// Original functions
function filterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master");
  ss.setActiveSheet(masterSheet);
  var data = masterSheet.getDataRange().getValues();
  
  // Get today's date in MM/DD/YYYY format
  var today = new Date();
  var dateString = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();

  // Construct the new sheet name
  var newSheetName = 'Bulks ' + dateString;
  
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

  // Filter and store data
  for (var i = 1; i < data.length; i++) {
    var bpName = data[i][3]; // Assuming BP Name is in the fourth column
    if (bpName !== 'Fanatics' && bpName !== 'One Time Shopify Customer' && bpName !== 'TS - lax.com') {
      filteredData.push(data[i]);
    }
  }

  // Append all filtered data at once if there is data to append
  if (filteredData.length > 1) {
    filteredSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  } else {
    Logger.log("No data to display after filtering.");
  }
  
  ss.setActiveSheet(masterSheet); // Switch back to the Master sheet
}

function filterComplexCriteria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master");
  ss.setActiveSheet(masterSheet);
  var data = masterSheet.getDataRange().getValues();
  
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

  ss.setActiveSheet(masterSheet); // Switch back to the Master sheet
}

function moveDesignRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master");
  ss.setActiveSheet(masterSheet);
  var data = masterSheet.getDataRange().getValues();
  
  // Format the current date as MM/DD/YYYY
  var today = new Date();
  var dateString = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();
  var designSheetName = 'Design ' + dateString;
  
  // Try to get the 'Design' sheet or create a new one with the current date
  var designSheet = ss.getSheetByName(designSheetName);
  if (!designSheet) {
    designSheet = ss.insertSheet(designSheetName);
    // Set headers for the new 'Design' sheet if it's newly created
    var headers = ["Posting Date", "Sales Order Number", "BP Code", "BP Name", "Reference Number", "Sunfrog: ID", 
                   "SunFrog: Send Order", "ART ON BOM LEVEL", "ART ON ORDER LEVEL", "Shopify Side Text", 
                   "Left Side Text", "Right Side Text", "Top Left Text", "Top Right Text", "Log", "Bulk Order File Name", 
                   "Royalty Entity", "Royalty Team / Show", "Royalty Player Character"];
    designSheet.appendRow(headers);
  } else {
    // Clear the existing 'Design' sheet if it already exists
    designSheet.clear();
    // Set headers for the 'Design' sheet after clearing
    var headers = ["Posting Date", "Sales Order Number", "BP Code", "BP Name", "Reference Number", "Sunfrog: ID", 
                   "SunFrog: Send Order", "ART ON BOM LEVEL", "ART ON ORDER LEVEL", "Shopify Side Text", 
                   "Left Side Text", "Right Side Text", "Top Left Text", "Top Right Text", "Log", "Bulk Order File Name", 
                   "Royalty Entity", "Royalty Team / Show", "Royalty Player Character"];
    designSheet.appendRow(headers);
  }

  // Column indices for BP Name and ART ON BOM LEVEL
  var bpNameIndex = 3;
  var artOnBomLevelIndex = 7;
  var shopifySideTextIndex = 9;

  // Filter rows based on criteria
  for (var i = 1; i < data.length; i++) {
    var bpName = data[i][bpNameIndex];
    var artOnBomLevel = data[i][artOnBomLevelIndex].trim(); // Trim spaces for accurate check
    var shopifySideText = data[i][shopifySideTextIndex];

    // Check if BP Name is "Fanatics" or "One Time Shopify Customer", ART ON BOM LEVEL is blank, and "YCIS" is not in Shopify Side Text
    if ((bpName === 'Fanatics' || bpName === 'One Time Shopify Customer' && artOnBomLevel === "") && shopifySideText.indexOf('YCIS') === -1) {
      designSheet.appendRow(data[i]);
    }
  }
  
  // Check if there are more than just the header rows
  if (designSheet.getLastRow() <= 1) {
    Logger.log("No specific rows to move to the Design sheet.");
  }

  ss.setActiveSheet(masterSheet); // Switch back to the Master sheet
}

function filterOneTimeShopYCIS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master");
  ss.setActiveSheet(masterSheet);
  var data = masterSheet.getDataRange().getValues();
  
  // Format the current date as MM/DD/YYYY
  var today = new Date();
  var dateString = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();
  var filterSheetName = 'YCIS ' + dateString;
  
  // Try to get the filter sheet or create a new one with the current date
  var filterSheet = ss.getSheetByName(filterSheetName);
  if (!filterSheet) {
    filterSheet = ss.insertSheet(filterSheetName);
  } else {
    // Clear the existing filter sheet if it already exists
    filterSheet.clear();
  }

  // Set headers for the new filter sheet
  var headers = ["Posting Date", "Sales Order Number", "BP Code", "BP Name", "Reference Number", "Sunfrog: ID", 
                 "SunFrog: Send Order", "ART ON BOM LEVEL", "ART ON ORDER LEVEL", "ART ON ORDER LEVEL", "SKU", 
                 "Shopify Side Text", "Left Side Text", "Right Side Text", "Top Left Text", "Top Right Text", "Log", 
                 "Bulk Order File Name", "Royalty Entity", "Royalty Team / Show", "Royalty Player Character"];
  filterSheet.appendRow(headers);

  // Column indices based on your data setup; adjust as necessary
  var bpNameIndex = 3;
  var skuIndex = 9; // Assuming SKU is in the 10th column (index starts from 0)

  // Filter rows based on criteria
  for (var i = 1; i < data.length; i++) {
    var bpName = data[i][bpNameIndex];
    var sku = data[i][skuIndex] || ""; // Ensure there's no undefined error

    // Check if BP Name is "One Time Shopify Customer" and SKU contains "YCIS"
    if (bpName === 'One Time Shopify Customer' && sku.includes('YCIS')) {
      filterSheet.appendRow(data[i]);
    }
  }
  
  // Check if there are more than just the header rows
  if (filterSheet.getLastRow() <= 1) {
    Logger.log("No specific rows to move to the filter sheet.");
  }

  ss.setActiveSheet(masterSheet); // Switch back to the Master sheet
}
