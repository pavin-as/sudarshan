// Helper: get number inside an "SPDxxxx" string
function extractNumeric(id) {
    const text = id ? id.toString() : '';
    const match = text.match(/\d+/);
    return match ? parseInt(match[0], 10) : 0;
  }
  
  function syncPatientData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('yaragoPatientMaster');
    const targetSheet = ss.getSheetByName('patientMaster');
    const timeZone = Session.getScriptTimeZone();
  
    // Is this the very first sync?
    const isFirstSync = targetSheet.getLastRow() === 1;
  
    // All source rows (skip header)
    const sourceData = sourceSheet
      .getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn())
      .getValues();
  
    // Highest numeric MRD already in patientMaster
    const lastSyncedNum = getLastSyncedNumericMRD(targetSheet);
  
    // Filter & transform new rows
    const newData = sourceData
      .filter(row => 
        isFirstSync 
          ? true 
          : extractNumeric(row[28]) > lastSyncedNum
      )
      .map(row => {
        // Format DOB (col E → index 4)
        let dobDate = '';
        if (row[4] instanceof Date) {
          dobDate = Utilities.formatDate(row[4], timeZone, 'yyyy-MM-dd');
        } else if (row[4].toString().trim()) {
          try {
            dobDate = Utilities.formatDate(new Date(row[4]), timeZone, 'yyyy-MM-dd');
          } catch(e) {
            dobDate = '';
          }
        }
  
        // Format Registration Date (col AQ → index 42)
        let regDate = '';
        if (row[42] instanceof Date) {
          regDate = Utilities.formatDate(row[42], timeZone, 'yyyy-MM-dd HH:mm:ss');
        } else if (row[42].toString().trim()) {
          try {
            regDate = Utilities.formatDate(new Date(row[42]), timeZone, 'yyyy-MM-dd HH:mm:ss');
          } catch(e) {
            regDate = '';
          }
        }
  
        return [
          row[28],                          // A: MRD No (full "SPDxxxx")
          [row[1], row[2], row[3]].join(' ').trim(), // B–D: Full Name
          dobDate,                          // E: DOB
          row[5],                           // F: Gender
          row[11],                          // L: Phone
          row[7],                           // H: Address
          '', '', '', '',                   // G–J: placeholders
          regDate                           // K: Registration Date
        ];
      });
  
    // Append if there’s anything new
    if (newData.length) {
      const startRow = isFirstSync ? 2 : targetSheet.getLastRow() + 1;
      targetSheet
        .getRange(startRow, 1, newData.length, newData[0].length)
        .setValues(newData);
    }
  }
  
  // Reads column A, extracts numbers, returns the max (or 0)
  function getLastSyncedNumericMRD(targetSheet) {
    const ids = targetSheet.getRange('A2:A').getValues().flat();
    const nums = ids.map(extractNumeric);
    return nums.length ? Math.max(...nums) : 0;
  }
  