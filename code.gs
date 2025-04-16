function compareAndCreateVulReports() {
  // Log the start of the process
  Logger.log('Starting the process of generating vulnerability reports...');
  
  // Replace with the actual Google Sheet ID of your fixed sheet
  const sheetId = 'YOUR_FIXED_SHEET_ID';  // Fixed Google Sheet ID
  
  // Replace with the Folder ID where you want to store the reports
  const folderId = 'YOUR_FOLDER_ID';  // Google Drive Folder ID
  
  // Log the folder ID
  Logger.log('Folder ID: ' + folderId);
  
  // Get the folder from Google Drive
  const folder = DriveApp.getFolderById(folderId);
  Logger.log('Accessed the folder: ' + folder.getName());
  
  // Open the Google Sheet using its ID
  const ss = SpreadsheetApp.openById(sheetId);
  Logger.log('Opened the Google Sheet with ID: ' + sheetId);
  
  // Fetch data from the 'Detail Data' and 'Last Week Data' tabs
  const sheetDetailData = ss.getSheetByName('Detail Data');
  const dataDetailData = sheetDetailData.getDataRange().getValues();  // Fetch all data from 'Detail Data'
  Logger.log('Fetched data from Detail Data sheet.');
  
  const sheetLastWeekData = ss.getSheetByName('Last Week Data');
  const dataLastWeekData = sheetLastWeekData.getDataRange().getValues();  // Fetch all data from 'Last Week Data'
  Logger.log('Fetched data from Last Week Data sheet.');
  
  // Get the current date
  const currentDate = new Date().toISOString().slice(0, 10); // Format: YYYY-MM-DD
  Logger.log('Current date: ' + currentDate);
  
  // List of vulnerabilities (replace with your actual list of vulnerabilities)
  const vulnerabilities = [
    "kb", "vul1", "vul2", "vul3", "vul4", "vul5", "vul6", "vul7", "vul8", "vul9",
    "vul10", "vul11", "vul12", "vul13", "vul14", "vul15", "vul16", "vul17", "vul18", "vul19",
    "vul20", "vul21", "vul22", "vul23", "vul24", "vul25", "vul26", "vul27", "vul28", "vul29",
    "vul30", "vul31", "vul32", "vul33", "vul34", "vul35", "vul36"
  ];  // Add your vulnerability keywords here
  
  // Loop through each vulnerability and create a report
  vulnerabilities.forEach(function(vul) {
    Logger.log('Processing vulnerability: ' + vul);
    
    // Step 1: Filter Detail Data by the keyword in Plugin name (this simulates your filter search)
    const filteredDetailData = dataDetailData.filter(row => row[0].toLowerCase().includes(vul.toLowerCase())); // Filter by Plugin name column (index 0)
    const filteredLastWeekData = dataLastWeekData.filter(row => row[0].toLowerCase().includes(vul.toLowerCase())); // Same for Last Week Data
    
    // Log number of findings
    Logger.log('Found ' + filteredDetailData.length + ' instances in Detail Data for ' + vul);
    Logger.log('Found ' + filteredLastWeekData.length + ' instances in Last Week Data for ' + vul);
    
    // Step 2: Create a new Google Sheet for this vulnerability
    const newVulSheet = SpreadsheetApp.create(vul + '_' + currentDate);
    Logger.log('Created a new sheet for vulnerability: ' + vul);
    
    const oldSheet = newVulSheet.insertSheet('Old');
    const newSheet = newVulSheet.insertSheet('New');
    
    // Step 3: Add headers to the new sheets (using the headers from Detail Data)
    oldSheet.appendRow(dataDetailData[0]);  // Use headers from the Detail Data
    newSheet.appendRow(dataDetailData[0]);  // Use headers from the Detail Data
    
    const oldData = [];
    const newData = [];
    
    // Step 4: Compare filtered data between Detail Data and Last Week Data
    const lastWeekUniqueIds = filteredLastWeekData.map(row => row[1]); // Assuming "Unique Identifier w Repository & Port" is in column 2 (index 1)
    
    // Step 5: Compare and categorize as old or new data
    filteredDetailData.forEach(function(row) {
      const pluginName = row[0];  // Plugin name
      const uniqueId = row[1];  // Unique Identifier w Repository & Port
      
      if (lastWeekUniqueIds.includes(uniqueId)) {
        // If it exists in Last Week Data, it's an old finding
        oldData.push(row); // Push the entire row into the old data array
      } else {
        // If it doesn't exist in Last Week Data, it's a new finding
        newData.push(row); // Push the entire row into the new data array
      }
    });
    
    // Log progress
    Logger.log('Comparison complete for vulnerability: ' + vul);
    
    // Step 6: Paste old data into the "Old" sheet
    if (oldData.length > 0) {
      oldSheet.getRange(2, 1, oldData.length, oldData[0].length).setValues(oldData);
      Logger.log('Pasted old data into the "Old" sheet.');
    }
    
    // Step 7: Paste new data into the "New" sheet
    if (newData.length > 0) {
      newSheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
      Logger.log('Pasted new data into the "New" sheet.');
    }
    
    // Step 8: Move the created sheet to the specified folder in Google Drive
    const file = DriveApp.getFileById(newVulSheet.getId());
    folder.createFile(file); // Move the file to the specific folder
    file.setTrashed(true); // Delete the original file from the root folder
    Logger.log('Moved the new sheet to the specified folder in Google Drive.');
    
    // Log the link to the newly created sheet
    Logger.log('New sheet created for ' + vul + ': ' + newVulSheet.getUrl());
  });
  
  // Log completion of the entire process
  Logger.log('Process completed. All reports generated and stored in the specified Google Drive folder.');
}
