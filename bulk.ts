function filterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
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
}
