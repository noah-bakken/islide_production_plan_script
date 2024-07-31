function moveDesignRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
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
}
