function filterOneTimeShopYCIS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
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
}
