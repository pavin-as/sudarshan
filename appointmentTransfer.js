// appointmentTransfer.js
// Script to transfer appointment data to another spreadsheet

/**
 * Transfers appointment data from the current spreadsheet to the target spreadsheet.
 * Only transfers data for the current date.
 */
function transferAppointmentData() {
  try {
    Logger.log("Starting appointment data transfer");
    
    // Get current date in format that matches the appointment sheet
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, "Asia/Kolkata", "dd/MM/yyyy");
    Logger.log("Getting appointments for date: " + formattedDate);
    
    // Source spreadsheet (current spreadsheet)
    const sourceSpreadsheet = SpreadsheetApp.getActive();
    const appointmentSheet = sourceSpreadsheet.getSheetByName("appointment");
    
    if (!appointmentSheet) {
      Logger.log("Error: 'appointment' sheet not found in the source spreadsheet");
      return;
    }
    
    // Get all data from appointment sheet
    const appointmentData = appointmentSheet.getDataRange().getValues();
    
    // Determine which column has the date (assuming it's column B, index 1)
    const dateColumnIndex = 1;
    
    // Find column indices for required data
    const mrdColumnIndex = 3;       // Column D
    const nameColumnIndex = 4;      // Column E
    const apptTimeColumnIndex = 2;  // Column C
    const ageColumnIndex = 5;       // Column F
    const planColumnIndex = 10;     // Column K
    const remarksColumnIndex = 11;  // Column L
    
    // Filter data for today's date
    const todayAppointments = appointmentData.filter((row, index) => {
      // Skip header row
      if (index === 0) return false;
      
      // Check if date matches today
      const rowDate = row[dateColumnIndex];
      let formattedRowDate = "";
      
      if (rowDate instanceof Date) {
        formattedRowDate = Utilities.formatDate(rowDate, "Asia/Kolkata", "dd/MM/yyyy");
      } else if (typeof rowDate === "string") {
        formattedRowDate = rowDate;
      }
      
      return formattedRowDate === formattedDate;
    });
    
    Logger.log(`Found ${todayAppointments.length} appointments for today`);
    
    if (todayAppointments.length === 0) {
      Logger.log("No appointments found for today. Exiting.");
      return;
    }
    
    // --- Key Fix: Handle TimeSlot Format Variations ---
    const timeToMinutes = (timeValue) => {
      if (!timeValue) return 0; // Empty/undefined
      
      let timeStr;
      // Case 1: Already a string (e.g., "9:00 AM")
      if (typeof timeValue === 'string') {
        timeStr = timeValue.trim();
      } 
      // Case 2: Date object (e.g., column contains datetime)
      else if (timeValue instanceof Date) {
        timeStr = Utilities.formatDate(timeValue, "Asia/Kolkata", "h:mm a");
      }
      // Case 3: Excel-style serial number (e.g., 0.375 for 9:00 AM)
      else if (typeof timeValue === 'number') {
        const date = new Date(Math.round((timeValue - 25569) * 86400 * 1000)); // Convert Excel serial to JS Date
        timeStr = Utilities.formatDate(date, "Asia/Kolkata", "h:mm a");
      }
      else {
        return 0; // Unsupported type
      }
      
      // Parse the time string (e.g., "9:00 AM")
      const timeParts = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
      if (!timeParts) return 0;
      
      let hours = parseInt(timeParts[1]);
      const minutes = parseInt(timeParts[2]);
      const period = timeParts[3].toUpperCase();
      
      if (period === 'PM' && hours < 12) hours += 12;
      if (period === 'AM' && hours === 12) hours = 0;
      
      return hours * 60 + minutes;
    };
    
    // Prepare and sort data
    let targetData = todayAppointments.map(row => [
      row[mrdColumnIndex],      // A: MRD
      "",                       // B: (empty)
      row[nameColumnIndex],     // C: Name
      row[ageColumnIndex],      // D: Age
      row[apptTimeColumnIndex], // E: TimeSlot (raw value)
      "",                       // F: (empty)
      "",                       // G: (empty)
      "",                       // H: (empty)
      row[remarksColumnIndex],  // I: Remarks
      "",                       // J: (empty)
      row[0],                   // K: AppointmentID
      "",                       // L: (empty)
      row[planColumnIndex]      // S: Plan of action
    ]);
    
    // Sort by time (Column E)
    targetData.sort((a, b) => timeToMinutes(a[4]) - timeToMinutes(b[4]));
  

    // Target spreadsheet
    const targetSpreadsheetId = "1jw4mTbWovgdnjpwN8e4MZKXV3ZyMmHf0N80v2kLjNnI";
    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    
    if (!targetSpreadsheet) {
      Logger.log("Error: Could not open target spreadsheet with ID: " + targetSpreadsheetId);
      return;
    }
    
       // Get the target sheet by name (instead of just the first sheet)
    const targetSheetName = "list_Sam"; // Specify your target sheet name here
    const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
    
    if (!targetSheet) {
      Logger.log(`Error: Sheet "${targetSheetName}" not found in target spreadsheet`);
      return;
    }

    // Find the first empty row in the target sheet
    const lastRow = Math.max(1, targetSheet.getLastRow());
    

       if (lastRow > 1) {
    targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn()).clearContent();
  }
    const startRow = lastRow + 1;
    // Write data to target spreadsheet
    if (targetData.length > 0) {
      targetSheet.getRange(startRow, 1, targetData.length, targetData[0].length).setValues(targetData);
      // Apply time format to Column E (Time Slot)
      const timeRange = targetSheet.getRange(startRow, 5, targetData.length, 1); // Column E
      timeRange.setNumberFormat("HH:mm"); // Force "09:00" format
      Logger.log(`Successfully transferred ${targetData.length} appointments to target spreadsheet`);
    }
    
  } catch (error) {
    Logger.log(`Error in transferAppointmentData: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * Creates a time-based trigger to run the transfer function at 6 AM daily.
 * This function needs to be run manually once to set up the trigger.
 */
function createDailyTrigger() {
  // Delete any existing triggers with the same function name
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === "transferAppointmentData") {
      ScriptApp.deleteTrigger(existingTriggers[i]);
    }
  }
  
  // Create a new trigger to run at 6 AM daily
  ScriptApp.newTrigger("transferAppointmentData")
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();
  
  Logger.log("Daily trigger created to run at 6 AM");
}

/**
 * Manual function to run the transfer immediately for testing.
 */
function manualRun() {
  transferAppointmentData();
} 