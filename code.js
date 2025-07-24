// Main entry point: renders login page or requested page based on session and page parameter.
function doGet(e) {
  var template;
  if (e.parameter.sessionToken) {
    var username = getSessionUsername(e.parameter.sessionToken);
    if (username) {
      if (e.parameter.page) {
        var page = e.parameter.page;
        try {
          template = HtmlService.createTemplateFromFile(page);
          template.sessionToken = e.parameter.sessionToken;
          template.webAppUrl = ScriptApp.getService().getUrl();

          // Add appointmentId to template if it exists in URL parameters
          if (e.parameter.appointmentId) {
            template.appointmentId = e.parameter.appointmentId;
          }
          
          var output = template.evaluate()
            .setTitle(page.charAt(0).toUpperCase() + page.slice(1))
            .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
          return output;
        } catch (error) {
          return HtmlService.createHtmlOutput("Page not found.");
        }
      } else {
        template = HtmlService.createTemplateFromFile('Dashboard');
        template.sessionToken = e.parameter.sessionToken;
        template.webAppUrl = ScriptApp.getService().getUrl();
        
        var output = template.evaluate()
          .setTitle('Dashboard')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
        return output;
      }
    }
  }
 
  template = HtmlService.createTemplateFromFile('Login');
  template.webAppUrl = ScriptApp.getService().getUrl();
  
  return template.evaluate()
    .setTitle('Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// Utility function to include other HTML files
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

// Validate login credentials from the "Login" sheet
function validateLogin(username, password) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Login');
  if (!sheet) {
    throw new Error('Login sheet not found!');
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return true; // Credentials match
    }
  }
  return false; // Credentials do not match
}

// Create a session for the logged-in user
function createSession(username) {
  const userData = {
    username: username,
    role: getUserRole(username) // You'll need to implement this function
  };
  return createUserSession(userData);
}

// Retrieve username from the session token
function getSessionUsername(sessionToken) {
  const session = getUserSession(sessionToken);
  return session ? session.username : null;
}

// Clear the session
function clearSession(sessionToken) {
  return deleteUserSession(sessionToken);
}

/**
 * Get user session information from the session token
 * @param {string} sessionToken - The session token to verify
 * @return {Object|null} User session object or null if invalid
 */
function getUserSession(sessionToken) {
  try {
    // Get the user properties
    const userProperties = PropertiesService.getUserProperties();
    
    // Get the session data
    const sessionData = userProperties.getProperty(sessionToken);
    if (!sessionData) {
      return null;
    }

    // Parse the session data
    const session = JSON.parse(sessionData);
    
    // Check if session is expired (24 hours)
    const now = new Date().getTime();
    if (now - session.timestamp > 24 * 60 * 60 * 1000) {
      // Session expired, remove it
      userProperties.deleteProperty(sessionToken);
      return null;
    }

    // Update session timestamp
    session.timestamp = now;
    userProperties.setProperty(sessionToken, JSON.stringify(session));

    return session;
  } catch (error) {
    return null;
  }
}

/**
 * Create a new user session
 * @param {Object} userData - User data to store in session
 * @return {string} Session token
 */
function createUserSession(userData) {
  try {
    // Generate a unique session token
    const sessionToken = Utilities.getUuid();
    
    // Create session object
    const session = {
      ...userData,
      timestamp: new Date().getTime()
    };

    // Store session data
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(sessionToken, JSON.stringify(session));

    return sessionToken;
  } catch (error) {
    return null;
  }
}

/**
 * Delete a user session
 * @param {string} sessionToken - The session token to delete
 * @return {boolean} True if session was deleted successfully
 */
function deleteUserSession(sessionToken) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty(sessionToken);
    return true;
  } catch (error) {
    return false;
  }
}

/**
 * Get user role from the Login sheet
 * @param {string} username - The username to get role for
 * @return {string} User role or 'user' if not found
 */
function getUserRole(username) {
  try {
    Logger.log("getUserRole called for username: " + username);
    
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Login');
    if (!sheet) {
      Logger.log("Login sheet not found, returning 'user'");
      return 'user';
    }

    Logger.log("Getting Login sheet data...");
    const data = sheet.getDataRange().getValues();
    Logger.log("Login sheet has " + data.length + " rows");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        const role = data[i][2] || 'user';
        Logger.log("Found role for " + username + ": " + role);
        return role; // Assuming role is in the third column
      }
    }
    
    Logger.log("User " + username + " not found in Login sheet, returning 'user'");
    return 'user';
  } catch (error) {
    Logger.log("Error in getUserRole for " + username + ": " + error.toString());
    return 'user';
  }
}

/**
 * Get user access level based on role
 * @param {string} role - The user's role
 * @return {number} Access level (1-4, where 1 is highest)
 */
function getUserAccessLevel(role) {
  const roleLevels = {
    'director': 1,
    'admin': 1,
    'doctor': 2,
    'manager': 3,
    'user': 4,
    'staff': 4,
    'receptionist': 4,
    'nurse': 4
  };
  
  const normalizedRole = role ? role.toLowerCase().trim() : 'user';
  return roleLevels[normalizedRole] || 4; // Default to level 4 for unknown roles
}

/**
 * Check if user has minimum required access level
 * @param {string} sessionToken - The session token to check
 * @param {number} requiredLevel - The minimum required access level (1-4)
 * @return {boolean} True if user has sufficient access level
 */
function hasAccessLevel(sessionToken, requiredLevel) {
  try {
    const session = getUserSession(sessionToken);
    if (!session) {
      Logger.log("hasAccessLevel: Invalid session token");
      return false;
    }
    
    const userLevel = getUserAccessLevel(session.role);
    Logger.log("hasAccessLevel: User " + session.username + " has level " + userLevel + ", required: " + requiredLevel);
    
    return userLevel <= requiredLevel; // Lower number = higher access
  } catch (error) {
    Logger.log("Error in hasAccessLevel: " + error.toString());
    return false;
  }
}

/**
 * Check if user is Level 1 (Director or Admin)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user is Level 1
 */
function isLevel1(sessionToken) {
  return hasAccessLevel(sessionToken, 1);
}

/**
 * Check if user is Level 2 (Doctor)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user is Level 2
 */
function isLevel2(sessionToken) {
  return hasAccessLevel(sessionToken, 2);
}

/**
 * Check if user is Level 3 (Manager)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user is Level 3
 */
function isLevel3(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Check if user is Level 4 (General user)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user is Level 4
 */
function isLevel4(sessionToken) {
  return hasAccessLevel(sessionToken, 4);
}

/**
 * Get user's current access level
 * @param {string} sessionToken - The session token
 * @return {Object} Object containing success status and access level
 */
function getCurrentUserAccessLevel(sessionToken) {
  try {
    const session = getUserSession(sessionToken);
    if (!session) {
      return { success: false, message: "Invalid session" };
    }
    
    const level = getUserAccessLevel(session.role);
    const roleNames = {
      1: 'Director/Admin',
      2: 'Doctor', 
      3: 'Manager',
      4: 'General User'
    };
    
    return {
      success: true,
      level: level,
      roleName: roleNames[level] || 'Unknown',
      role: session.role
    };
  } catch (error) {
    Logger.log("Error in getCurrentUserAccessLevel: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Check if user has access to admin functions (Level 1 only)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user has admin access
 */
function hasAdminAccess(sessionToken) {
  return isLevel1(sessionToken);
}

/**
 * Check if user has access to doctor functions (Level 1 or 2)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user has doctor access
 */
function hasDoctorAccess(sessionToken) {
  return hasAccessLevel(sessionToken, 2);
}

/**
 * Check if user has access to manager functions (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user has manager access
 */
function hasManagerAccess(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Check if user can delete appointments (Level 1 only)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can delete appointments
 */
function canDeleteAppointments(sessionToken) {
  return isLevel1(sessionToken);
}

/**
 * Check if user can archive appointments (Level 1 or 2)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can archive appointments
 */
function canArchiveAppointments(sessionToken) {
  return hasAccessLevel(sessionToken, 2);
}

/**
 * Check if user can update patient details (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can update patient details
 */
function canUpdatePatientDetails(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Check if user can view analytics (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can view analytics
 */
function canViewAnalytics(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Check if user can access visualization dashboard (Level 1 only)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can access visualization
 */
function canAccessVisualization(sessionToken) {
  return isLevel1(sessionToken);
}

/**
 * Check if user can access settings configuration (Level 1 only)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can access settings
 */
function canAccessSettings(sessionToken) {
  return isLevel1(sessionToken);
}

/**
 * Check if user can access rescheduling analytics (Level 1 only)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can access rescheduling analytics
 */
function canAccessRescheduleAnalytics(sessionToken) {
  return isLevel1(sessionToken);
}

/**
 * Check if user can manage system settings (Level 1 only)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can manage system settings
 */
function canManageSystemSettings(sessionToken) {
  return isLevel1(sessionToken);
}

/**
 * Check if user can book appointments (All levels)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can book appointments
 */
function canBookAppointments(sessionToken) {
  return hasAccessLevel(sessionToken, 4); // All levels can book
}

/**
 * Check if user can reschedule appointments (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can reschedule appointments
 */
function canRescheduleAppointments(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Check if user can cancel appointments (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can cancel appointments
 */
function canCancelAppointments(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Check if user can confirm appointments (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if user can confirm appointments
 */
function canConfirmAppointments(sessionToken) {
  return hasAccessLevel(sessionToken, 3);
}

/**
 * Get all available roles for the system
 * @return {Array} Array of role objects with level and description
 */
function getAvailableRoles() {
  return [
    { role: 'director', level: 1, description: 'Director - Full system access' },
    { role: 'admin', level: 1, description: 'Administrator - Full system access' },
    { role: 'doctor', level: 2, description: 'Doctor - Medical and patient management' },
    { role: 'manager', level: 3, description: 'Manager - Operational management' },
    { role: 'user', level: 4, description: 'General User - Basic appointment operations' },
    { role: 'staff', level: 4, description: 'Staff - Basic appointment operations' },
    { role: 'receptionist', level: 4, description: 'Receptionist - Basic appointment operations' },
    { role: 'nurse', level: 4, description: 'Nurse - Basic appointment operations' }
  ];
}

/**
 * Retrieves appointments for the given date.
 * Expects date in "YYYY-MM-DD" format.
 * Returns an array of appointment objects.
 */
function getAppointmentsForDay(dateString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("appointment");
  if (!sheet) {
    throw new Error("Appointment sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var results = [];


  
  // Find the index of the Fixed Slot column (ensure case-insensitive comparison)
  var fixedSlotIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && typeof headers[i] === 'string' && 
        headers[i].toLowerCase() === 'fixed slot') {
      fixedSlotIndex = i;
      break;
    }
  }

    // Find the index of the PatientType column (ensure case-insensitive comparison)
  var patientTypeIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && typeof headers[i] === 'string' && 
        headers[i].toLowerCase() === 'patienttype') {
      patientTypeIndex = i;
      break;
    }
  }
  
  // Find the index of the Family Identifier column (ensure case-insensitive comparison)
  var familyIdentifierIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && typeof headers[i] === 'string' && 
        headers[i].toLowerCase() === 'family identifier') {
      familyIdentifierIndex = i;
      break;
    }
  }
  


  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var formattedDate = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (formattedDate === dateString) {
      var appointment = {};
      for (var j = 0; j < headers.length; j++) {
        appointment[headers[j]] = row[j].toString(); // Pass all values as strings
      }
      
      // Ensure FixedSlot property is set regardless of header case
      if (fixedSlotIndex !== -1) {
        appointment["FixedSlot"] = row[fixedSlotIndex].toString();
      }
      
      // Ensure PatientType property is set regardless of header case
      if (patientTypeIndex !== -1) {
        appointment["PatientType"] = row[patientTypeIndex].toString();
      }

      // Ensure Family Identifier property is set regardless of header case
      if (familyIdentifierIndex !== -1) {
        appointment["Family Identifier"] = row[familyIdentifierIndex].toString();
      }

      results.push(appointment);
    }
  }


  return results;
}

function confirmAppointment4Day(appointmentId, sessionToken) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("appointment");
    if (!sheet) {
      return { success: false, message: "Sheet not found." };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var appointmentRow = null;
    var rowIndex = -1;

    // Find the appointment row
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === appointmentId) {
        appointmentRow = data[i];
        rowIndex = i;
        break;
      }
    }

    if (!appointmentRow) {
      return { success: false, message: "Appointment not found." };
    }

    // Check for pre-existing 4-day confirmation
    var fourDayConfirmDateCol = headers.indexOf("FourDayConfirmDate");
    var fourDayConfirmByCol = headers.indexOf("FourDayConfirmBy");

    if (fourDayConfirmDateCol === -1 || fourDayConfirmByCol === -1) {
      return { success: false, message: "Confirmation columns not found." };
    }

    if (appointmentRow[fourDayConfirmDateCol] && appointmentRow[fourDayConfirmByCol]) {
      return { success: false, message: "4-Day Confirmation already done." };
    }

    var apptDate = new Date(appointmentRow[1]);
    if (isNaN(apptDate.getTime())) {
      return { success: false, message: "Invalid appointment date." };
    }

    var today = new Date();
    var sevenDaysBefore = new Date(apptDate);
    sevenDaysBefore.setDate(apptDate.getDate() - 7);

    var formattedApptDate = Utilities.formatDate(apptDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedSevenDaysBefore = Utilities.formatDate(sevenDaysBefore, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

     
    // Allow confirmation within 7 days before the appointment, including the appointment day

    if (formattedToday >= formattedSevenDaysBefore && formattedToday <= formattedApptDate) {
      var user = getSessionUsername(sessionToken);
      var todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

      sheet.getRange(rowIndex + 1, fourDayConfirmDateCol + 1).setValue(todayFormatted);
      sheet.getRange(rowIndex + 1, fourDayConfirmByCol + 1).setValue(user);

      return { 
        success: true, 
        message: "4-Day confirmation successful.", 
        confirmationDate: todayFormatted, 
        confirmedBy: user 
      };
    } else {
      return { 
        success: false, 
        message: "Confirmation is not allowed at this time. It must be within 7 days before the appointment." 
      };
    }
  } catch (e) {
    return { success: false, message: "An error occurred during confirmation." };
  }
}



function confirmAppointment1Day(appointmentId, sessionToken) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("appointment");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var appointmentRow = null;
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === appointmentId) {
        appointmentRow = data[i];
        rowIndex = i;
        break;
      }
    }

    if (!appointmentRow) {
      return { success: false, message: "Appointment not found." };
    }

    //Check for pre-existing 1-day confirmation
    if (appointmentRow[headers.indexOf("OneDayConfirmDate")] && appointmentRow[headers.indexOf("OneDayConfirmBy")]){
      return {success: false, message: "1-Day Confirmation already done."};
    }

    var apptDate = new Date(appointmentRow[1]);
    var formattedApptDate = Utilities.formatDate(apptDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var today = new Date();
    var threeDaysBefore = new Date(apptDate);
    threeDaysBefore.setDate(apptDate.getDate() - 3);
    var formattedThreeDaysBefore = Utilities.formatDate(threeDaysBefore, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Allow confirmation within 3 days before the appointment, including the appointment day
    if (formattedToday >= formattedThreeDaysBefore && formattedToday <= formattedApptDate) {
      var oneDayConfirmDateCol = headers.indexOf("OneDayConfirmDate");
      var oneDayConfirmByCol = headers.indexOf("OneDayConfirmBy");
      var user = getSessionUsername(sessionToken);
      var todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

      if (oneDayConfirmDateCol !== -1 && oneDayConfirmByCol !== -1) {
        sheet.getRange(rowIndex + 1, oneDayConfirmDateCol + 1).setValue(todayFormatted);
        sheet.getRange(rowIndex + 1, oneDayConfirmByCol + 1).setValue(user);
      } else {
        return { success: false, message: "Confirmation columns not found." };
      }

      return { success: true, message: "1-Day confirmation successful.", confirmationDate: todayFormatted, confirmedBy: user };
    } else {
      return { success: false, message: "Confirmation is not allowed at this time. It must be within 3 days before the appointment or on the appointment day." };
    }
  } catch (e) {
    return { success: false, message: "An error occurred during confirmation." };
  }
}

// ... (Your existing code: doGet, include, validateLogin, createSession, getSessionUsername, clearSession, getAppointmentsForDay, confirmAppointment4Day, confirmAppointment1Day) ...

function getAppointmentsForMonth(year, month, doctor) {
  try {
    // Convert month index to month name for logging
    const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const monthName = monthNames[month];
  
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("appointment");
    if (!sheet) {
      throw new Error("Appointment sheet not found.");
    }
    
    var data = sheet.getDataRange().getValues();
    var results = {};
    
    // Column indices (0-based)
    var appointmentDateCol = 1; // Column B (AppointmentDate)
    var mrdCol = 3; // Column D (MRDNo)
    var appointmentTypeCol = 8; // Column I (AppointmentType)
    var doctorCol = 9; // Column J (Doctor)
    
    for (var i = 1; i < data.length; i++) { // Start from row 1 to skip header
      var row = data[i];
      var appointmentDate = new Date(row[appointmentDateCol]);
      if (isNaN(appointmentDate.getTime())) {
        continue; // Skip invalid dates
      }
      
      // Filter by doctor if provided
      if (doctor && row[doctorCol] !== doctor) {
        continue;
      }
      
      if (appointmentDate.getFullYear() === year && appointmentDate.getMonth() === month) {
        var day = appointmentDate.getDate();
        if (!results[day]) {
          results[day] = { NEW: 0, OLD: 0 };
        }
        
        var mrd = row[mrdCol].toString();
        if (mrd.startsWith("N")) {
          results[day].NEW++; // Increment NEW count
        } else if (mrd.startsWith("S")) {
          results[day].OLD++; // Increment OLD count
        }
        
        var appointmentType = row[appointmentTypeCol].toString();
        if (!results[day][appointmentType]) {
          results[day][appointmentType] = 0;
        }
        results[day][appointmentType]++;
      }
    }
    
    Logger.log(`getAppointmentsForMonth: Results for ${monthName} ${year}, doctor ${doctor || 'ALL'}: ${JSON.stringify(results)}`);
    return results;
  } catch (e) {
    Logger.log(`Error in getAppointmentsForMonth for ${monthName} ${year}: ${e.toString()}`);
    return { error: e.toString() };
  }
}

function getAvailableSlotsCount() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    const doctorAvailabilitySheet = ss.getSheetByName("DoctorAvailability");
    
    if (!appointmentsSheet || !doctorAvailabilitySheet) {
      Logger.log("Required sheets not found");
      return { count: 0, trend: 0 };
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Get day of week (0 = Sunday, 1 = Monday, etc.)
    const dayOfWeek = today.getDay();
    const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const todayName = dayNames[dayOfWeek];
    
    // Get booked slots for today
    const appointmentData = appointmentsSheet.getDataRange().getValues();
    const bookedSlots = new Set();
    
    // Count booked slots for today
    for (let i = 1; i < appointmentData.length; i++) {
      const appointmentDate = new Date(appointmentData[i][1]); // Column B is appointment date
      appointmentDate.setHours(0, 0, 0, 0);
      
      if (appointmentDate.getTime() === today.getTime()) {
        const timeSlot = appointmentData[i][2]; // Column C is time slot
        bookedSlots.add(timeSlot);
      }
    }
    
    // Get doctor availability for today
    const availabilityData = doctorAvailabilitySheet.getDataRange().getValues();
    const headers = availabilityData[0];
    const dayColumnIndex = headers.indexOf(todayName);
    
    if (dayColumnIndex === -1) {
      return { count: 0, trend: 0 };
    }
    
    // Calculate total available minutes for all doctors
    let totalAvailableMinutes = 0;
    
    for (let i = 1; i < availabilityData.length; i++) {
      const doctor = availabilityData[i][0]; // Column A is doctor name
      const availabilityRanges = availabilityData[i][dayColumnIndex];
      
      if (!availabilityRanges) continue;
      
      // Parse availability ranges and calculate total minutes
      const ranges = availabilityRanges.toString().split(',').map(r => r.trim());
      
      for (const range of ranges) {
        const [startTime, endTime] = range.split('-');
        if (!startTime || !endTime) continue;
        
        const startMinutes = timeStringToMinutes(startTime);
        const endMinutes = timeStringToMinutes(endTime);
        
        if (!isNaN(startMinutes) && !isNaN(endMinutes) && endMinutes > startMinutes) {
          totalAvailableMinutes += (endMinutes - startMinutes);
        }
      }
    }
    
    // Calculate total slots based on 10-minute intervals
    const slotInterval = 10; // Same as in bookAppointment page
    const totalPossibleSlots = Math.floor(totalAvailableMinutes / slotInterval);
    
    // Subtract booked slots
    const availableSlots = totalPossibleSlots - bookedSlots.size;
    
    return {
      count: availableSlots > 0 ? availableSlots : 0,
      trend: 0 // No trend for available slots
    };
  } catch (error) {
    return { count: 0, trend: 0 };
  }
}
/*function getAvailableSlotsCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appointmentsSheet = ss.getSheetByName("appointment");
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const data = appointmentsSheet.getDataRange().getValues();
  const bookedSlots = new Set();
  
  // Count booked slots for today
  for (let i = 1; i < data.length; i++) {
    const appointmentDate = new Date(data[i][1]); // Column B is appointment date
    appointmentDate.setHours(0, 0, 0, 0);
    
    if (appointmentDate.getTime() === today.getTime()) {
      const timeSlot = data[i][2]; // Column C is time slot
      bookedSlots.add(timeSlot);
    }
  }
  
  // Get total available slots from general settings
  const generalSettingsSheet = ss.getSheetByName("generalSettings");
  const settingsData = generalSettingsSheet.getDataRange().getValues();
  const totalSlots = settingsData[1][3]; // Assuming total slots are stored in cell D2
  
  return {
    count: totalSlots - bookedSlots.size,
    trend: 0 // No trend for available slots
  };
}
*/

// Book appointment form, MRD No
function getPatientDetailsByMRD(mrdNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var patientMasterSheet = ss.getSheetByName("patientMaster");
    var tempPatientMasterSheet = ss.getSheetByName("tempPatientMaster");
    
    if (!patientMasterSheet && !tempPatientMasterSheet) {
      return { success: false, message: "Patient sheets not found." };
    }
    // Normalize the MRD number
    mrdNo = mrdNo.trim();
    
    // If it's just a number, add SPD prefix
    if (/^\d+$/.test(mrdNo)) {
        mrdNo = "SPD" + mrdNo;
    }
    // If it starts with SPD but has no numbers after, add the number
    else if (/^SPD$/i.test(mrdNo)) {
      mrdNo = "SPD0";
    }
    // If it starts with SPD and has numbers, ensure proper format
    else if (/^SPD/i.test(mrdNo)) {
     // Extract the number part
      var numPart = mrdNo.replace(/^SPD/i, '');
      if (/^\d+$/.test(numPart)) {
           mrdNo = "SPD" + numPart;
      }
    }
    
    // Helper function to calculate exact age difference
    function calculateExactAge(dob) {
      var today = new Date();
      var birthDate = new Date(dob);
      
      // Calculate difference in milliseconds
      var diffInMs = today.getTime() - birthDate.getTime();
      var diffInDays = Math.floor(diffInMs / (1000 * 60 * 60 * 24));
      
      if (diffInDays < 30) {
        // Less than 30 days - return in days
        return diffInDays + " days";
      } else if (diffInDays < 365) {
        // 30 days or more but less than a year - return in months
        var months = Math.floor(diffInDays / 30);
        return months + " months";
      } else {
        // A year or more - return in years
        var years = Math.floor(diffInDays / 365);
        return years + " years";
      }
    }
    
    // Function to search in a sheet
    function searchInSheet(sheet) {
      if (!sheet) return null;
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var currentMRD = data[i][0] ? data[i][0].toString().trim().toUpperCase() : '';
        
        if (currentMRD === mrdNo.toUpperCase()) {
          // Calculate exact age from DOB
          var dob = new Date(data[i][2]);
          var exactAge = calculateExactAge(dob);
       
             return {
            MRDNo:         data[i][0],
            Name:          data[i][1],
            DOB:           Utilities.formatDate(
                          new Date(data[i][2]),
                          Session.getScriptTimeZone(),
                          'yyyy-MM-dd'
                        ),
            Age:           exactAge,
            Gender:        data[i][3],
            Phone:         data[i][4],
            Address:       data[i][5],
            AdditionalInfo:data[i][6],
            Type:          data[i][7],
            Status:        data[i][8]
                   };
        }
      }
      return null;
    }
    
    // First check tempPatientMaster, then patientMaster
    var patient = searchInSheet(tempPatientMasterSheet) || searchInSheet(patientMasterSheet);
    
    if (patient) {
      return { success: true, patient: patient };
    }
    
    return { success: false, message: "MRD No not found." };
  } catch (error) {
     return { success: false, message: "Error retrieving patient details: " + error.toString() };
  }
}

//Book appointment, Appointment Type
function getAppointmentTypes() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("generalSettings");
    if (!sheet) {
      return { success: false, message: "generalSettings sheet not found." };
    }
    
    // Get the appointment types from cell D2
    var cellValue = sheet.getRange("D2").getValue();
    if (!cellValue) {
      return { success: false, message: "No appointment types found in cell D2." };
    }
    
    // Split the cell's value by comma, trim each type, and filter out any empty strings
    var appointmentTypes = cellValue.split(",")
      .map(function(item) { return item.trim(); })
      .filter(function(item) { return item !== ""; });
    
    return { success: true, appointmentTypes: appointmentTypes };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}



/**
 * Submit a new appointment.
 * appointmentData is an object containing:
 *   - patientType: "existingPatient" or "newPatient"
 *   - patientData: Object with patient details (MRDNo, Name, Age, Gender, Phone, Address, PatientType)
 *   - appointmentType, doctor, appointmentDate, timeSlot, planOfAction, remarks
 *   - sessionToken (for retrieving BookedBy)
 */
function submitAppointment(appointmentData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var appointmentSheet = ss.getSheetByName("appointment");
    if (!appointmentSheet) {
      throw new Error("Appointment sheet not found.");
    }

    // Check and ensure the header for column Z is "Fixed Slot"
    var headers = appointmentSheet.getRange(1, 1, 1, appointmentSheet.getLastColumn()).getValues()[0];
    
    // Check if we need to add the Fixed Slot header
    var fixedSlotColIndex = headers.indexOf("Fixed Slot");
    if (fixedSlotColIndex === -1) {
      // If Fixed Slot column doesn't exist, add it to column Z (26)
      fixedSlotColIndex = 25; // 0-based index for column Z
      if (appointmentSheet.getLastColumn() < 26) {
        // If sheet doesn't have 26 columns yet, add columns up to Z
        appointmentSheet.insertColumnsAfter(appointmentSheet.getLastColumn(), 26 - appointmentSheet.getLastColumn());
      }
      appointmentSheet.getRange(1, 26).setValue("Fixed Slot");
      headers[25] = "Fixed Slot";
    }

    // Check if we need to add the Urgency Level header in column AB (28)
    var urgencyLevelColIndex = headers.indexOf("Urgency Level");
    if (urgencyLevelColIndex === -1) {
      // If Urgency Level column doesn't exist, add it to column AB (28)
      urgencyLevelColIndex = 27; // 0-based index for column AB
      if (appointmentSheet.getLastColumn() < 28) {
        // If sheet doesn't have 28 columns yet, add columns up to AB
        appointmentSheet.insertColumnsAfter(appointmentSheet.getLastColumn(), 28 - appointmentSheet.getLastColumn());
      }
      appointmentSheet.getRange(1, 28).setValue("Urgency Level");
      headers[27] = "Urgency Level";
    }

    // Check if we need to add the Family Identifier header in column AC (29)
    var familyIdentifierColIndex = headers.indexOf("Family Identifier");
    if (familyIdentifierColIndex === -1) {
      // If Family Identifier column doesn't exist, add it to column AC (29)
      familyIdentifierColIndex = 28; // 0-based index for column AC
      if (appointmentSheet.getLastColumn() < 29) {
        // If sheet doesn't have 29 columns yet, add columns up to AC
        appointmentSheet.insertColumnsAfter(appointmentSheet.getLastColumn(), 29 - appointmentSheet.getLastColumn());
      }
      appointmentSheet.getRange(1, 29).setValue("Family Identifier");
      headers[28] = "Family Identifier";
    }

    // Generate timestamp-based appointment ID
    var now = new Date();
    var appointmentID = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    
    // For existing patients, fetch details from patientMaster
    if (appointmentData.patientType === "existingPatient") {
      var patientDetails = getPatientDetailsByMRD(appointmentData.patientData.MRDNo);
      if (!patientDetails.success) {
        throw new Error("Failed to fetch existing patient details: " + patientDetails.message);
      }
      // Use the fetched patient details
      appointmentData.patientData = patientDetails.patient;
    }
    
    // For new patients, generate MRD number and add to tempPatientMaster
    else if (appointmentData.patientType === "newPatient") {
      var newMRD = generateNewMRD();
      appointmentData.patientData.MRDNo = newMRD;
      
      // Add new patient to tempPatientMaster sheet
      var tempPatientMasterSheet = ss.getSheetByName("tempPatientMaster");
      if (!tempPatientMasterSheet) {
        throw new Error("tempPatientMaster sheet not found.");
      }
      var newPatientRow = [
        newMRD,                                    // MRD No (Column A)
        appointmentData.patientData.Name,          // Name (Column B)
        appointmentData.patientData.Age,           // Age (Column C) - Direct input from form
        appointmentData.patientData.Gender,        // Gender (Column D)
        appointmentData.patientData.Phone,         // Phone (Column E)
        appointmentData.patientData.Address,       // Address (Column F)
        appointmentData.patientData.AdditionalInfo || "", // Additional Info (Column G)
        appointmentData.patientData.Type,          // Type (Column H)
        "Active",                                  // Status (Column I)
        "",                                        // Empty Column (Column J)
        now,                                       // RegisteredOn (Column K)
        appointmentData.fixedSlot ? "Yes" : ""     // Fixed Slot (Column Z)
      ];
      
      tempPatientMasterSheet.appendRow(newPatientRow);
    }
    
    // Prepare appointment record
    var bookedBy = getSessionUsername(appointmentData.sessionToken);
    var bookingDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
       // Create an array with 29 elements (A through AC)
    var newRow = new Array(29).fill("");
    
    // Get patient type from the appointmentData
    var patientType = "";
    if (appointmentData.patientType === "existingPatient") {
      patientType = appointmentData.patientData.Type || "";
    } else {
      patientType = appointmentData.patientData.Type || "";
    }
    
    // Fill in the appointment data at the correct indices
    newRow[0] = appointmentID;                               // Appointment ID (Column A)
    newRow[1] = appointmentData.appointmentDate;             // Appointment Date (Column B)
    newRow[2] = appointmentData.timeSlot;                    // Time Slot (Column C)
    newRow[3] = appointmentData.patientData.MRDNo;           // MRD No (Column D)
    newRow[4] = appointmentData.patientData.Name;            // Name (Column E)
    newRow[5] = appointmentData.patientData.Age;             // Age (Column F)
    newRow[6] = appointmentData.patientData.Gender;          // Gender (Column G)
    newRow[7] = appointmentData.patientData.Phone;           // Phone (Column H)
    newRow[8] = appointmentData.appointmentType;             // Appointment Type (Column I)
    newRow[9] = appointmentData.doctor;                      // Doctor (Column J)
    newRow[10] = appointmentData.planOfAction;               // Plan of Action (Column K)
    newRow[11] = appointmentData.remarks;                    // Remarks (Column L)
    newRow[12] = bookedBy;                                   // Booked By (Column M)
    newRow[13] = bookingDate;                                // Booking Date (Column N)
    newRow[14] = "";                                         // FourDayConfirmDate (Column O)
    newRow[15] = "";                                         // FourDayConfirmBy (Column P)
    newRow[16] = "";                                         // OneDayConfirmDate (Column Q)
    newRow[17] = "";                                         // OneDayConfirmBy (Column R)
    newRow[18] = "";                                         // RescheduledDate (Column S)
    newRow[19] = "";                                         // RescheduledBy (Column T)
    newRow[25] = appointmentData.fixedSlot || "";            // Fixed Slot (Column Z)
    newRow[26] = patientType;                                // Patient Type (Column AA)
    newRow[27] = appointmentData.urgencyLevel || "";         // Urgency Level (Column AB)
    newRow[28] = appointmentData.familyIdentifier || "";     // Family Identifier (Column AC)
    

    
    appointmentSheet.appendRow(newRow);
    
    // Also save to doctor-specific "list_<Doctor>" sheet
    var doctorSheetName = "list_" + appointmentData.doctor.replace(/[^a-zA-Z0-9]/g, "_"); // Clean doctor name for sheet name
    var doctorSheet = ss.getSheetByName(doctorSheetName);
    
    // Create the doctor sheet if it doesn't exist
    if (!doctorSheet) {
      doctorSheet = ss.insertSheet(doctorSheetName);
      // Add headers to the new sheet
      var headers = ["MRD No", "Name", "Time Slot", "Age", "Gender", "Phone", "Appointment Type", "Plan of Action", "Remarks"];
      doctorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    // Create a simplified row for the doctor sheet with Age in column D (index 3)
    var doctorSheetRow = [
      appointmentData.patientData.MRDNo,           // MRD No (Column A)
      appointmentData.patientData.Name,            // Name (Column B)
      appointmentData.timeSlot,                    // Time Slot (Column C)
      appointmentData.patientData.Age,             // Age (Column D) - This will have the exact age format
      appointmentData.patientData.Gender,          // Gender (Column E)
      appointmentData.patientData.Phone,           // Phone (Column F)
      appointmentData.appointmentType,             // Appointment Type (Column G)
      appointmentData.planOfAction,                // Plan of Action (Column H)
      appointmentData.remarks                      // Remarks (Column I)
    ];
    
    doctorSheet.appendRow(doctorSheetRow);
    
    return { 
      success: true, 
      message: "Appointment submitted successfully." + 
        (appointmentData.patientType === "newPatient" ? 
          " New MRD Number: " + appointmentData.patientData.MRDNo : "")
    };
  } catch (e) {
    Logger.log("Error in submitAppointment: " + e.toString());
    return { success: false, message: "Error submitting appointment: " + e.toString() };
  }
}
//Book Appointment select Doctor
function getDoctors() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DoctorAvailability");
    if (!sheet) {
      return { success: false, message: "DoctorAvailability sheet not found." };
    }
    
    // Get values from column A, starting at A2
    var data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
    var doctors = [];
    for (var i = 0; i < data.length; i++) {
      var doc = data[i][0];
      if (doc && doc.toString().trim() !== "") {
        doctors.push(doc.toString().trim());
      }
    }
    
    return { success: true, doctors: doctors };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Book Appointment: Returns an array of holiday dates (formatted as "yyyy-MM-dd") from the Holidays sheet (A2:A)
function getHolidays() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Holidays");
    if (!sheet) {
      return { success: false, message: "Holidays sheet not found." };
    }
    var data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
    var holidays = data.filter(function(row) {
      return row[0];
    }).map(function(row) {
      return Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    });
    return { success: true, holidays: holidays };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}




// Available time slot

// Convert minutes to a "HH:mm" string (24‑hour format)
function formatTime(totalMinutes) {
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  return ("0" + hours).slice(-2) + ":" + ("0" + minutes).slice(-2);
}

// Convert a time string ("HH:mm") to minutes past midnight
function timeStringToMinutes(timeStr) {
  try {
    const parts = timeStr.trim().split(":");
    if (parts.length !== 2) return NaN;
    const hours = parseInt(parts[0],10), minutes = parseInt(parts[1],10);
    return isNaN(hours)||isNaN(minutes) ? NaN : hours*60 + minutes;
  } catch (e) {
    return NaN;
  }
}





  


function getDoctorAvailabilityRanges(doctor, dayOfWeek) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DoctorAvailability");
  var data = sheet.getDataRange().getValues();
  var ranges = [];
  // Assume first row is header with day names (Monday, Tuesday, etc.)
  var headers = data[0];
  var dayIndex = headers.indexOf(dayOfWeek);
  if (dayIndex === -1) throw new Error("Day not found in DoctorAvailability header.");
  
  // Find the row for the given doctor (assume doctor name is in Column A)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === doctor) {
      var cellValue = data[i][dayIndex];
      if (cellValue) {
        // Expect a string like "09:00-12:00, 13:00-17:30"
        ranges = cellValue.split(",").map(function(r) { return r.trim(); });
      }
      break;
    }
  }
  return ranges;
}

function getBookedSlots(doctor, dateString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("appointment");
  var data = sheet.getDataRange().getValues();
  var booked = [];

  //Logger.log("getBookedSlots: Total rows in appointment sheet: " + data.length);
  
  // Assume: Date is in Column B, Time Slot in Column C, Doctor in Column J; first row is header.
  for (var i = 1; i < data.length; i++) {
    var rowDate = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var rowDoctor = data[i][9] ? data[i][9].toString().trim() : "";
    
    // Format the raw time slot into "HH:mm" format
    var timeSlotRaw = data[i][2];
    var formattedTime = Utilities.formatDate(new Date(timeSlotRaw), Session.getScriptTimeZone(), "HH:mm");
    

    
    if (rowDate === dateString && rowDoctor === doctor) {
      booked.push(formattedTime);
    }
  }
  
  //Logger.log("getBookedSlots: Booked slots for doctor " + doctor + " on " + dateString + ": " + JSON.stringify(booked));
  return booked;
}

function getDoctorExceptionRanges(doctor, dateString) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DoctorAvailabilityExceptions");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var exceptions = [];
    
    for (var i = 1; i < data.length; i++) {
      // Column 0: Doctor, Column 1: Date, Column 2: From Time, Column 3: To Time
      var rowDoctor = data[i][0]?.toString().trim() || "";
      var rowDate = data[i][1] ? Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      
      if (rowDoctor === doctor && rowDate === dateString) {
        // Format times consistently as "HH:mm"
        var fromTime = formatTimeString(data[i][2]);
        var toTime = formatTimeString(data[i][3]);
        
        if (fromTime && toTime) {
          exceptions.push(fromTime + "-" + toTime);
        }
      }
    }
    return exceptions;
  } catch (e) {
    return [];
  }
}

// Helper function to standardize time strings
function formatTimeString(timeInput) {
  if (!timeInput) return "";
  
  // Handle cases where time might be a string ("9:00") or Date object
  var timeStr = typeof timeInput === 'string' ? timeInput : Utilities.formatDate(timeInput, Session.getScriptTimeZone(), "HH:mm");
  
  // Ensure 24-hour format with leading zero
  var parts = timeStr.split(":");
  if (parts.length === 2) {
    var hours = parts[0].padStart(2, '0');
    var minutes = parts[1].padStart(2, '0');
    return hours + ":" + minutes;
  }
  return "";
}


/*function getDoctorExceptionRanges(doctor, dateString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DoctorAvailabilityExceptions");
  var data = sheet.getDataRange().getValues();
  var exceptions = [];
  // Assume first row is header.
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === doctor && 
        Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd") === dateString) {
      var fromTime = data[i][2]; // From Time (Column C)
      var toTime = data[i][3];   // To Time (Column D)
      if (fromTime && toTime) {
        exceptions.push(fromTime + "-" + toTime);
      }
    }
  }
  return exceptions;
}
*/
function isSlotWithinRanges(slot, ranges) {
  if (!slot || !ranges || ranges.length === 0) return false;
  
  var slotMins = timeStringToMinutes(slot);
  
  for (var i = 0; i < ranges.length; i++) {
    var rangeParts = ranges[i].split("-");
    if (rangeParts.length !== 2) continue;
    
    var startTime = rangeParts[0].trim();
    var endTime = rangeParts[1].trim();
    
    var startMins = timeStringToMinutes(startTime);
    var endMins = timeStringToMinutes(endTime);
    
    if (slotMins >= startMins && slotMins < endMins) {
      return true;
    }
  }
  return false;
}


// Server-side function in your .gs file:
function getDoctorMonthlyData(doctor, monthKey, filterDoctor) {
  // Parse monthKey (e.g., "2025-03") to get year and month.
  var parts = monthKey.split("-");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // 0-indexed

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("appointment");
  var data = sheet.getDataRange().getValues();
  var bookedSlots = {};
  var exceptions = {};

  // Loop through appointments (assuming first row is header).
  for (var i = 1; i < data.length; i++) {
    var apptDate = new Date(data[i][1]);
    var formattedDate = Utilities.formatDate(apptDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    // Check if appointment is for the specified doctor and month.
    var appointmentDoctor = data[i][9].toString().trim();
    if (formattedDate.indexOf(monthKey) === 0) {
      // If filterDoctor is provided, only include appointments for that doctor
      if (!filterDoctor || appointmentDoctor === filterDoctor) {
        // If looking for specific doctor data, match exactly; otherwise include all
        if (!doctor || appointmentDoctor === doctor) {
          // Assuming booked slot is in column C in "HH:mm" format.
          if (!bookedSlots[formattedDate]) {
            bookedSlots[formattedDate] = [];
          }
          bookedSlots[formattedDate].push(data[i][2].toString().trim());
        }
      }
    }
  }

  // Similarly, you can fetch exception ranges from the DoctorAvailabilityExceptions sheet.
  var exceptionSheet = ss.getSheetByName("DoctorAvailabilityExceptions");
  if (exceptionSheet) {
    var exData = exceptionSheet.getDataRange().getValues();
    for (var j = 1; j < exData.length; j++) {
      var exceptionDoctor = exData[j][0].toString().trim();
      // Filter exceptions by doctor if filterDoctor is provided
      if (!filterDoctor || exceptionDoctor === filterDoctor) {
        if (!doctor || exceptionDoctor === doctor) {
          var exDate = new Date(exData[j][1]);
          var formattedExDate = Utilities.formatDate(exDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          if (formattedExDate.indexOf(monthKey) === 0) {
            if (!exceptions[formattedExDate]) {
              exceptions[formattedExDate] = [];
            }
            var fromTime = exData[j][2];
            var toTime = exData[j][3];
             if (fromTime && toTime) {
    var fromStr = Utilities.formatDate(new Date(fromTime), Session.getScriptTimeZone(), "HH:mm");
    var toStr = Utilities.formatDate(new Date(toTime), Session.getScriptTimeZone(), "HH:mm");
    exceptions[formattedExDate].push(fromStr + "-" + toStr);
  }

          }
        }
      }
    }
  }

  return { bookedSlots: bookedSlots, exceptions: exceptions };
}


/**
 * Return every doctor's weekly availability so
 * Month Overview can compute total slots per day.
 */
function getDoctorAvailabilityDataToday(doctor) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DoctorAvailability");
    if (!sheet) {
      throw new Error("DoctorAvailability sheet not found.");
    }

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];                // ["Doctor","Monday",…,"Saturday"]
    var doctors = [];

    // build one object per doctor
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var doctorName = row[0];
      
      // Filter by doctor if provided
      if (doctor && doctorName !== doctor) {
        continue;
      }
      
      var doc = { name: doctorName };
      for (var j = 1; j < headers.length; j++) {
        doc[ headers[j] ] = row[j];      // e.g. doc["Monday"] = "09:00-12:00"
      }
      doctors.push(doc);
    }

    return { success: true, doctors: doctors };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


function getDoctorAvailabilityData(doctor, selectedDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DoctorAvailability");
    if (!sheet) {
      throw new Error("DoctorAvailability sheet not found.");
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0]; // Assumes first row contains day names
    var availabilityRanges = [];

    // Determine the day of the week for the selected date
    var selectedDate = new Date(selectedDateStr);
    var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    var dayOfWeek = dayNames[selectedDate.getDay()];

    Logger.log(`Selected Date: ${selectedDateStr}`);
    Logger.log(`Day of Week: ${dayOfWeek}`);

    // Find the row for the given doctor (assuming doctor name is in Column A)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === doctor) {
        var cellValue = data[i][headers.indexOf(dayOfWeek)];
        if (cellValue) {
          // Expect a string like "10:00-12:00, 13:00-17:30"
          availabilityRanges = cellValue.toString().split(",").map(function(item) {
            return item.trim();
          });
        }
        break;
      }
    }

    // Get exception ranges for this doctor and date
    var exceptions = getDoctorExceptionRanges(doctor, selectedDateStr);

    Logger.log(`Availability ranges for ${dayOfWeek}: ${JSON.stringify(availabilityRanges)}`);
    Logger.log(`Exception ranges for ${selectedDateStr}: ${JSON.stringify(exceptions)}`);
    
    return { 
      success: true, 
      availability: availabilityRanges,
      exceptions: exceptions 
    };
  } catch (e) {
    Logger.log("Error in getDoctorAvailabilityData: " + e.toString());
    return { success: false, message: e.toString() };
  }
}


/**
 * Returns allowed time slots for a given doctor, appointment type, and appointment date,
 * factoring in day codes (e.g., "M", "Tu") and offset logic.
 */

/**
 * Returns allowed time slots for a given doctor, appointment type, and appointment date,
 * factoring in day codes (e.g., "M", "Tu") and offset logic.
 */
function getAllowedSlots(doctor, appointmentType, appointmentDateStr, slotInterval = 10) {
  try {
    Logger.log("=== START getAllowedSlots ===");
    Logger.log("Parameters - Doctor: " + doctor + 
               ", Type: " + appointmentType + 
               ", Date: " + appointmentDateStr + 
               ", Interval: " + slotInterval);

    // Validate inputs
    if (!doctor || !appointmentType || !appointmentDateStr) {
      throw new Error("Missing required parameters");
    }

    var appointmentDate = new Date(appointmentDateStr);
    if (isNaN(appointmentDate.getTime())) {
      throw new Error("Invalid date format. Expected YYYY-MM-DD");
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    appointmentDate.setHours(0, 0, 0, 0);

    // Check if appointment is today
    var isToday = appointmentDate.getTime() === today.getTime();
    Logger.log("Is appointment today: " + isToday);

    // Calculate days between today and appointment
    var diffDays = Math.floor((appointmentDate - today) / (1000 * 60 * 60 * 24));
    Logger.log("Days until appointment: " + diffDays);

    // Get day name and code
    var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    var dayOfWeek = dayNames[appointmentDate.getDay()];
    var selectedDayCode = getDayCode(dayOfWeek);
    Logger.log("Day of week: " + dayOfWeek + " (" + selectedDayCode + ")");

    // 1. Get doctor's default availability
    var availabilityResp = getDoctorAvailabilityData(doctor, appointmentDateStr);
    if (!availabilityResp.success) {
      throw new Error("Error fetching doctor's availability: " + availabilityResp.message);
    }
    
    // getDoctorAvailabilityData returns availability as an array, not an object with day keys
    var defaultIntervals = availabilityResp.availability || [];
    Logger.log("Default intervals: " + JSON.stringify(defaultIntervals));

    // Convert each range into time slots
    var defaultSlots = [];
    defaultIntervals.forEach(function(range) {
      var parts = range.split("-");
      if (parts.length < 2) return;
      
      var startTime = parts[0].trim();
      var endTime = parts[1].trim();
      var startMins = timeStringToMinutes(startTime);
      var endMins = timeStringToMinutes(endTime);
      
      Logger.log("Processing range: " + startTime + " to " + endTime);
      
      for (var t = startMins; t < endMins; t += parseInt(slotInterval)) {
        defaultSlots.push(formatTime(t));
      }
    });
    Logger.log("Generated " + defaultSlots.length + " default slots");

    // 2. Remove Booked Slots
    var booked = getBookedSlots(doctor, appointmentDateStr);
    Logger.log("Booked slots: " + JSON.stringify(booked));
    
    var availableSlots = defaultSlots.filter(function(slot) {
      return booked.indexOf(slot) === -1;
    });
    Logger.log("Available after removing booked: " + availableSlots.length + " slots");

    // 3. Remove Exception Ranges
    var exceptions = getDoctorExceptionRanges(doctor, appointmentDateStr);
    Logger.log("Exception ranges: " + JSON.stringify(exceptions));
    
    availableSlots = availableSlots.filter(function(slot) {
      var slotMins = timeStringToMinutes(slot);
      var isBlocked = false;
      
      for (var i = 0; i < exceptions.length; i++) {
        var exParts = exceptions[i].split("-");
        if (exParts.length < 2) continue;
        
        var exStart = timeStringToMinutes(exParts[0]);
        var exEnd = timeStringToMinutes(exParts[1]);
        
        if (slotMins >= exStart && slotMins < exEnd) {
          Logger.log("Blocking slot " + slot + " due to exception " + exceptions[i]);
          isBlocked = true;
          break;
        }
      }
      return !isBlocked;
    });
    Logger.log("Available after exceptions: " + availableSlots.length + " slots");

    // 4. Type-specific filtering
    appointmentType = appointmentType.toUpperCase();
    Logger.log("Applying filters for appointment type: " + appointmentType);

    // 4A. Waiting List - no slots
    if (appointmentType === "WAITING LIST") {
      Logger.log("Waiting list - no slots available");
      return { success: true, allowedSlots: [] };
    }

   
    // 4B. General Appointment filtering
    if (appointmentType === "GENERAL") {
       var disallowed = getDisallowedIntervals();
      Logger.log("General appointment disallowed intervals: " + JSON.stringify(disallowed));
      availableSlots = availableSlots.filter(function(slot) {
        var slotMins = timeStringToMinutes(slot);
        var isAllowed = true;

    
        for (var i = 0; i < disallowed.length; i++) {
          var d = disallowed[i];
                 // Check for both specific day and daily ("D") intervals
          if (d.day !== selectedDayCode && d.day !== "D") continue;
          
              // Block if booking is more than the specified days
          if (diffDays > d.days) {
            var dStart = timeStringToMinutes(d.startTime);
            var dEnd = timeStringToMinutes(d.endTime);
            
            if (slotMins >= dStart && slotMins < dEnd) {
              Logger.log("Blocking slot " + slot + " due to general disallowed interval");
              isAllowed = false;
              break;
            }
          }

        }
         return isAllowed;
      });
        Logger.log("Available after general filters: " + availableSlots.length + " slots");
      return { success: true, allowedSlots: availableSlots };
    }

     // 4C. Specialized Appointment filtering (HFA, ROP, DILATATION, etc.)
    var specialized = getSpecializedIntervals(appointmentType);
    Logger.log("Specialized intervals for " + appointmentType + ": " + JSON.stringify(specialized));

    // If no specialized intervals are defined, return all available slots
    if (specialized.length === 0) {
      Logger.log("No specialized intervals defined, returning all available slots");
      return { success: true, allowedSlots: availableSlots };
    }
    var specializedSlots = [];
    specialized.forEach(function(interval) {
        // Check for both specific day and daily ("D") intervals
      if (interval.day !== selectedDayCode && interval.day !== "D") return;
           // Ensure days is a number
      var days = typeof interval.days === 'number' ? interval.days : 0;
   
      // If days is -1, the slot is only available today
      // If days is 0, the slot is always available
      // If days is N, the slot is only available when booking N or fewer days in advance
        if (days === -1) {
        if (isToday) {
          var startMins = timeStringToMinutes(interval.startTime);
          var endMins = timeStringToMinutes(interval.endTime);
          
          for (var t = startMins; t < endMins; t += parseInt(slotInterval)) {
            specializedSlots.push(formatTime(t));
          }
        }
        } else if (days === 0 || diffDays <= days) {
        var startMins = timeStringToMinutes(interval.startTime);
        var endMins = timeStringToMinutes(interval.endTime);
        
        for (var t = startMins; t < endMins; t += parseInt(slotInterval)) {
          specializedSlots.push(formatTime(t));
        }
      }
    });
     Logger.log("Specialized slots: " + JSON.stringify(specializedSlots));
    Logger.log("Available slots before intersection: " + JSON.stringify(availableSlots));

    // Intersect available slots with specialized slots
    availableSlots = availableSlots.filter(function(slot) {
     
      var isInSpecialized = specializedSlots.indexOf(slot) !== -1;
      Logger.log("Checking slot " + slot + " - In specialized slots: " + isInSpecialized);
      return isInSpecialized;
    });
        Logger.log("Final available slots: " + JSON.stringify(availableSlots));

    // After getting booked slots, check for fixed slots
    const fixedSlots = [];
    for (const slot of availableSlots) {
      if (isSlotFixed(doctor, appointmentDateStr, slot)) {
        fixedSlots.push(slot);
      }
    }
    
    // Return both available and fixed slots
    return {
      success: true,
      allowedSlots: availableSlots,
      fixedSlots: fixedSlots
    };
    
  } catch (e) {
    Logger.log("Error in getAllowedSlots: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Helper function to convert day name to code
function getDayCode(dayName) {
  var map = {
    "Sunday": "Su",
    "Monday": "M",
    "Tuesday": "Tu",
    "Wednesday": "W",
    "Thursday": "Th",
    "Friday": "F",
    "Saturday": "Sa",
    "Daily": "D"  // Add Daily option
  };
  return map[dayName] || "";
}


//helper function to validate times:
function validateTimeInput(timeStr) {
  return /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/.test(timeStr);
}



// Helper function to format minutes as HH:mm
function formatTime(totalMinutes) {
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  return ("0" + hours).slice(-2) + ":" + ("0" + minutes).slice(-2);
}


/**
 * Parses a slot string in the format "M-09:20-09:40-0" and returns an object.
 */

function parseSlotData(slotString) {
  var parts = slotString.split("-");
  var days = 0; // Default to 0 if not specified
  
  // If there's a fourth part, try to parse it as a number
  if (parts.length >= 4) {
    var daysStr = parts[3].trim();
    // Handle special case for -1 (today only)
    if (daysStr === "-1") {
      days = -1;
    } else {
      var parsedDays = parseInt(daysStr);
      days = isNaN(parsedDays) ? 0 : parsedDays;
    }
  }
    Logger.log("Parsing interval: " + slotString + " -> days: " + days);
  return {
    // parts[0] = "M" or "Tu" or "W" ...
    day: parts[0].trim(),
    startTime: parts[1].trim(),
    endTime: parts[2].trim(),
       days: days
  };
}
/*// ADD: Helper to map day name -> code
function getDayCode(dayName) {
  // Example mapping (adapt to however you store them):
  var map = {
    "Sunday": "Su",
    "Monday": "M",
    "Tuesday": "Tu",
    "Wednesday": "W",
    "Thursday": "Th",
    "Friday": "F",
    "Saturday": "Sa"
  };
  return map[dayName] || ""; 
}
*/

/**
 * Retrieves allowed specialized intervals for the given appointment type from the "generalSettings" sheet.
 * For example, for HFA, the allowed interval might be stored in Column F.
 * Returns an array of objects using parseSlotData.
 */
function getSpecializedIntervals(appointmentType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("generalSettings");
  if (!sheet) throw new Error("generalSettings sheet not found.");
  
  var columnMapping = {
    "HFA": "F",
    "ROP": "G",
    "INJECTION": "H",
    "URGENT NEW": "I",
    "EMERGENCY": "J",
    "DILATION": "K",
    "LASER": "L",
    "REFERRALS": ["E", "I", "J"]
  };
  
  var mapping = columnMapping[appointmentType];
  var intervals = [];

    
  if (!mapping) {
    Logger.log("No mapping found for appointment type: " + appointmentType);
    return intervals;
  }
  
   function parseIntervalFromCell(cellValue) {
    if (!cellValue) return null;
    
    var parts = cellValue.toString().split("-");
    if (parts.length < 4) return null;
    
    var daysStr = parts[3].trim();
    var days;
    
    // Handle special case for -1 (today only)
    if (daysStr === "-1") {
      days = -1;
    } else {
      var parsedDays = parseInt(daysStr);
      days = isNaN(parsedDays) ? 0 : parsedDays;
    }
    
    return {
      day: parts[0].trim(),
      startTime: parts[1].trim(),
      endTime: parts[2].trim(),
      days: days
    };
  }

  if (Array.isArray(mapping)) {
    mapping.forEach(function(col) {
      var data = sheet.getRange(col + "2:" + col + "100").getValues();
      data.forEach(function(row) {
        if (row[0]) {
            var interval = parseIntervalFromCell(row[0]);
          if (interval) intervals.push(interval);
        }
      });
    });
  } else {
    var data = sheet.getRange(mapping + "2:" + mapping + "100").getValues();
    data.forEach(function(row) {
      if (row[0]) {
        var interval = parseIntervalFromCell(row[0]);
        if (interval) intervals.push(interval);
      }
    });
  }
  Logger.log("Retrieved intervals for " + appointmentType + ": " + JSON.stringify(intervals));
  return intervals;
}


function getAllDoctorLeaveDays(doctor) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DoctorLeaveDays");
    if (!sheet) {
      return { success: false, message: "DoctorLeaveDays sheet not found." };
    }

    var data = sheet.getDataRange().getValues();
    var leaveDates = [];
    
    // Start from i=1 to skip header row
    for (var i = 1; i < data.length; i++) {
      // Match doctor in Column A; date in Column B
      if (data[i][0] && data[i][0].toString().trim() === doctor && data[i][1]) {
        var dateVal = new Date(data[i][1]);
        // Only add if parse is successful
        if (!isNaN(dateVal.getTime())) {
          var formatted = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
          leaveDates.push(formatted);
        }
      }
    }

    return { success: true, leaveDates: leaveDates };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
function getDisallowedIntervals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("generalSettings");
  if (!sheet) throw new Error("generalSettings sheet not found.");
  
  var data = sheet.getRange("E2:E100").getValues();
  var intervals = [];
  
  data.forEach(function(row) {
    if (row[0]) {
      // Use parseSlotData to parse the line, e.g., "M-09:20-09:40-4"
      var slot = parseSlotData(row[0].toString());
      intervals.push(slot);
    }
  });
  
  return intervals;
}

function updateAppointmentDetails(updatedData, sessionToken) {
  try {
    Logger.log("Starting updateAppointmentDetails for appointmentId: " + updatedData.appointmentId);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("appointment");
    if (!sheet) {
      Logger.log("Sheet 'appointment' not found.");
      return { success: false, message: "Sheet not found." };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rowIndex = -1;

    // Find the appointment row
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === updatedData.appointmentId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      Logger.log("Appointment ID not found: " + updatedData.appointmentId);
      return { success: false, message: "Appointment not found." };
    }

    // Get column indices for the fields to update
    var typeCol = headers.indexOf("AppointmentType");
    var doctorCol = headers.indexOf("Doctor");
    var planCol = headers.indexOf("PlanOfAction");
    var remarksCol = headers.indexOf("Remarks");

    // Update the values in the sheet
    if (typeCol !== -1) sheet.getRange(rowIndex + 1, typeCol + 1).setValue(updatedData.type);
    if (doctorCol !== -1) sheet.getRange(rowIndex + 1, doctorCol + 1).setValue(updatedData.doctor);
    if (planCol !== -1) sheet.getRange(rowIndex + 1, planCol + 1).setValue(updatedData.plan);
    if (remarksCol !== -1) sheet.getRange(rowIndex + 1, remarksCol + 1).setValue(updatedData.remarks);

    Logger.log("Appointment details updated successfully.");
    return { success: true, message: "Appointment details updated successfully." };
  } catch (e) {
    Logger.log("Error in updateAppointmentDetails: " + e.toString());
    return { success: false, message: "An error occurred while updating appointment details: " + e.toString() };
  }
}


function rescheduleAppointment(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("appointment");
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    // Get username from session token
    const username = getSessionUsername(data.sessionToken);
    if (!username) {
      return { success: false, message: "Invalid session. Please login again." };
    }

    // Find the appointment row
    const rowIndex = values.findIndex(row => row[0] === data.appointmentId);
    if (rowIndex === -1) return { success: false, message: "Appointment not found" };

    // Update core appointment details
    const updateMap = {
      'AppointmentDate': data.newDate,
      'TimeSlot': data.newTimeSlot,
      'AppointmentType': data.appointmentType,
      'Doctor': data.doctor,
      'PlanOfAction': data.newPlanOfAction,
      'Remarks': data.newRemarks
    };

    // Update spreadsheet columns
    Object.entries(updateMap).forEach(([header, value]) => {
      const colIndex = headers.indexOf(header);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
      }
    });

    // Clear confirmation columns as fresh confirmations will be required
    const fourDayConfirmDateCol = headers.indexOf("FourDayConfirmDate");
    const fourDayConfirmByCol = headers.indexOf("FourDayConfirmBy");
    const oneDayConfirmDateCol = headers.indexOf("OneDayConfirmDate");
    const oneDayConfirmByCol = headers.indexOf("OneDayConfirmBy");
    
    if (fourDayConfirmDateCol !== -1) {
      sheet.getRange(rowIndex + 1, fourDayConfirmDateCol + 1).setValue("");
    }
    if (fourDayConfirmByCol !== -1) {
      sheet.getRange(rowIndex + 1, fourDayConfirmByCol + 1).setValue("");
    }
    if (oneDayConfirmDateCol !== -1) {
      sheet.getRange(rowIndex + 1, oneDayConfirmDateCol + 1).setValue("");
    }
    if (oneDayConfirmByCol !== -1) {
      sheet.getRange(rowIndex + 1, oneDayConfirmByCol + 1).setValue("");
    }

    // Update reschedule history (using username instead of email)
    const historyCol = headers.indexOf("RescheduleHistory");
    const history = values[rowIndex][historyCol] ? JSON.parse(values[rowIndex][historyCol]) : [];
    
    // Get the current appointment date and time before updating
    const currentDate = values[rowIndex][headers.indexOf("AppointmentDate")];
    const currentTime = values[rowIndex][headers.indexOf("TimeSlot")];
    
    history.push({
      timestamp: new Date(),
      by: username,  // Use username from Login sheet instead of email
      fromDate: currentDate,
      toDate: data.newDate,
      fromTime: currentTime,
      toTime: data.newTimeSlot,
      action: "rescheduled"
    });
    sheet.getRange(rowIndex + 1, historyCol + 1).setValue(JSON.stringify(history));

    // Update tracking columns (using username instead of email)
    const rescheduleDateCol = headers.indexOf("RescheduledDate");
    const rescheduledByCol = headers.indexOf("RescheduledBy");
    const now = new Date();
    sheet.getRange(rowIndex + 1, rescheduleDateCol + 1).setValue(now);
    sheet.getRange(rowIndex + 1, rescheduledByCol + 1).setValue(username);  // Use username from Login sheet instead of email

    return { success: true, message: "Appointment rescheduled successfully" };
  } catch (e) {
    console.error("Reschedule error:", e);
    return { 
      success: false, 
      message: "Error rescheduling appointment: " + e.toString() 
    };
  }
}
// Function to generate new MRD number
function generateNewMRD() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("appointment");
    if (!sheet) {
      throw new Error("appointment sheet not found.");
    }
    
    // Get all MRD numbers from column D
    var data = sheet.getDataRange().getValues();
    var maxMRD = 0;
    
    // Start from row 2 to skip header
    for (var i = 1; i < data.length; i++) {
      var mrd = data[i][3].toString().trim();
      if (mrd.startsWith("N")) {
        // Extract the number part after 'N'
        var num = parseInt(mrd.substring(1));
        if (!isNaN(num) && num > maxMRD) {
          maxMRD = num;
        }
      }
    }
    
    // Generate new MRD number by incrementing the max number found
    var newMRD = "N" + (maxMRD + 1).toString().padStart(4, "0");
    return newMRD;
  } catch (e) {
    Logger.log("Error in generateNewMRD: " + e.toString());
    throw e;
  }
}

// Dashboard Statistics Functions
function getTodayAppointmentsCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appointmentsSheet = ss.getSheetByName("appointment");
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const data = appointmentsSheet.getDataRange().getValues();
  let count = 0;
  let yesterdayCount = 0;
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  for (let i = 1; i < data.length; i++) {
    const appointmentDate = new Date(data[i][1]); // Column B is appointment date
    appointmentDate.setHours(0, 0, 0, 0);
    
    if (appointmentDate.getTime() === today.getTime()) {
      count++;
    } else if (appointmentDate.getTime() === yesterday.getTime()) {
      yesterdayCount++;
    }
  }
  
  return {
    count: count,
    trend: yesterdayCount > 0 ? ((count - yesterdayCount) / yesterdayCount * 100).toFixed(1) : 0
  };
}

function getPending4DayConfirmationsCount() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("appointment");
  const data  = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0,0,0,0);

  // 1. Build next 4 working days (skip Sundays)
  const workDays = [];
  let cursor = new Date(today);
  while (workDays.length < 4) {
    if (cursor.getDay() !== 0) workDays.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }

  // 2. Filter and count
  let pendingCount = 0;
  let totalCount   = 0;
  let minDistance  = Infinity;

  data.slice(1).forEach(row => {
    const apptDate = new Date(row[1]);
    apptDate.setHours(0,0,0,0);
    const fourDayConfirm = row[14]; // FourDayConfirmDate

    // Is this date in our 4-day window?
    const idx = workDays.findIndex(d => d.getTime() === apptDate.getTime());
    if (idx > -1) {
      totalCount++;
      if (!fourDayConfirm) {
        pendingCount++;
        minDistance = Math.min(minDistance, idx);
      }
    }
  });

  // 3. Trend %
  const trend = totalCount
   ? (pendingCount / totalCount) * 100
    : 0;

  // 4. Urgency
  let urgency = "success";
  if      (minDistance <= 2) urgency = "danger";
 else if (minDistance <= 3) urgency = "warning";

  return { count: pendingCount, trend, urgency };
}


function getMonthlyPatientsCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appointmentsSheet = ss.getSheetByName("appointment");
  const today = new Date();
  const firstDayOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const firstDayOfLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  
  const data = appointmentsSheet.getDataRange().getValues();
  let currentMonthCount = 0;
  let lastMonthCount = 0;
  const uniquePatients = new Set();
  const lastMonthPatients = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const appointmentDate = new Date(data[i][1]); // Column B is appointment date
    const patientMRD = data[i][3]; // Column D is MRD No
    
    if (appointmentDate >= firstDayOfMonth) {
      uniquePatients.add(patientMRD);
      currentMonthCount = uniquePatients.size;
    } else if (appointmentDate >= firstDayOfLastMonth && appointmentDate < firstDayOfMonth) {
      lastMonthPatients.add(patientMRD);
      lastMonthCount = lastMonthPatients.size;
    }
  }
  
  return {
    count: currentMonthCount,
    trend: lastMonthCount > 0 ? ((currentMonthCount - lastMonthCount) / lastMonthCount * 100).toFixed(1) : 0
  };
}


function getUpcomingAppointments() {
  try {
    Logger.log("Starting getUpcomingAppointments");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    
    if (!appointmentsSheet) {
      Logger.log("Error: Appointment sheet not found");
      return { success: false, message: "Appointment sheet not found", appointments: [] };
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const data = appointmentsSheet.getDataRange().getValues();
    Logger.log("Total rows in appointment sheet: " + data.length);
    
    const headers = data[0];
    const upcomingAppointments = [];
    
    // Get column indices
    const idCol = headers.indexOf("AppointmentID");
    const dateCol = headers.indexOf("AppointmentDate");
    const timeCol = headers.indexOf("TimeSlot");
    const patientNameCol = headers.indexOf("PatientName");
    const typeCol = headers.indexOf("AppointmentType");
    const doctorCol = headers.indexOf("Doctor");
    const fourDayConfirmDateCol = headers.indexOf("FourDayConfirmDate");
    const fourDayConfirmByCol = headers.indexOf("FourDayConfirmBy");
    const oneDayConfirmDateCol = headers.indexOf("OneDayConfirmDate");
    const oneDayConfirmByCol = headers.indexOf("OneDayConfirmBy");
    const mrdCol = headers.indexOf("MRDNo");
    const phoneCol = headers.indexOf("Phone");
    
    Logger.log("Column indices - ID: " + idCol + ", Date: " + dateCol + ", Time: " + timeCol);
    
    // Validate column indices
    if (idCol === -1 || dateCol === -1 || timeCol === -1 || patientNameCol === -1 || 
        typeCol === -1 || doctorCol === -1 || fourDayConfirmDateCol === -1 || 
        fourDayConfirmByCol === -1 || oneDayConfirmDateCol === -1 || oneDayConfirmByCol === -1) {
      Logger.log("Error: Required columns not found in sheet");
      return { 
        success: false, 
        message: "Required columns not found in sheet", 
        appointments: [] 
      };
    }
    
    for (let i = 1; i < data.length; i++) {
      try {
        const row = data[i];
        if (!row[idCol] || !row[dateCol] || !row[timeCol]) {
          Logger.log("Skipping row " + i + " due to missing required data");
          continue;
        }
        
        const appointmentDate = new Date(row[dateCol]);
        const fourDayConfirmed = row[fourDayConfirmDateCol];
        const oneDayConfirmed = row[oneDayConfirmDateCol];
        
        // Skip if both confirmations are done
        if (fourDayConfirmed && oneDayConfirmed) {
          continue;
        }
        
        if (appointmentDate >= today) {
          // Format dates as strings for serialization
          let appointmentTime;
          try {
            appointmentTime = Utilities.formatDate(new Date(row[timeCol]), Session.getScriptTimeZone(), "HH:mm");
          } catch (timeError) {
            Logger.log("Error formatting time for row " + i + ": " + timeError);
            appointmentTime = "00:00"; // Default time if there's an error
          }
          
          let formattedFourDayDate = null;
          if (fourDayConfirmed) {
            try {
              formattedFourDayDate = Utilities.formatDate(new Date(fourDayConfirmed), Session.getScriptTimeZone(), "yyyy-MM-dd");
            } catch (dateError) {
              Logger.log("Error formatting 4-day date: " + dateError);
            }
          }
          
          let formattedOneDayDate = null;
          if (oneDayConfirmed) {
            try {
              formattedOneDayDate = Utilities.formatDate(new Date(oneDayConfirmed), Session.getScriptTimeZone(), "yyyy-MM-dd");
            } catch (dateError) {
              Logger.log("Error formatting 1-day date: " + dateError);
            }
          }
          
          const appointment = {
            id: row[idCol].toString(),
            time: appointmentTime,
            date: Utilities.formatDate(appointmentDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
            patientName: row[patientNameCol] ? row[patientNameCol].toString() : '',
            mrdNo: row[mrdCol] ? row[mrdCol].toString() : '',
            phone: row[phoneCol] ? row[phoneCol].toString() : '',
            type: row[typeCol] ? row[typeCol].toString() : '',
            doctor: row[doctorCol] ? row[doctorCol].toString() : '',
            fourDayConfirmDate: formattedFourDayDate,
            fourDayConfirmBy: row[fourDayConfirmByCol] ? row[fourDayConfirmByCol].toString() : '',
            oneDayConfirmDate: formattedOneDayDate,
            oneDayConfirmBy: row[oneDayConfirmByCol] ? row[oneDayConfirmByCol].toString() : ''
          };
          
          // Validate appointment data
          if (!appointment.id || !appointment.time) {
            Logger.log("Skipping invalid appointment data at row " + i);
            continue;
          }
          
          upcomingAppointments.push(appointment);
        }
      } catch (rowError) {
        Logger.log("Error processing row " + i + ": " + rowError.toString());
      }
    }
    
    Logger.log("Found " + upcomingAppointments.length + " upcoming appointments");
    
    // Sort by date first, then by time
    upcomingAppointments.sort((a, b) => {
      // First compare dates
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      const dateDiff = dateA - dateB;
      
      // If dates are the same, compare times
      if (dateDiff === 0) {
        const timeA = a.time.split(':');
        const timeB = b.time.split(':');
        const hourDiff = parseInt(timeA[0]) - parseInt(timeB[0]);
        if (hourDiff !== 0) return hourDiff;
        return parseInt(timeA[1]) - parseInt(timeB[1]);
      }
      
      return dateDiff;
    });
    
    // Return only next 5 appointments
    const result = upcomingAppointments.slice(0, 5);
    Logger.log("Returning " + result.length + " appointments");
    
    // Create response object
    const response = {
      success: true,
      appointments: result,
      timestamp: new Date().getTime()
    };
    
    // Serialize and deserialize to ensure JSON compatibility
    const jsonString = JSON.stringify(response);
    Logger.log("JSON response (first 200 chars): " + jsonString.substring(0, 200) + (jsonString.length > 200 ? "..." : ""));
    
    // Return parsed object
    return JSON.parse(jsonString);
  } catch (error) {
    Logger.log("Error in getUpcomingAppointments: " + error.toString());
    // Return a simple error object that is guaranteed to be serializable
    return {
      success: false,
      message: "Error: " + error.toString(),
      appointments: [],
      timestamp: new Date().getTime()
    };
  }
}

function getAppointmentDistribution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appointmentsSheet = ss.getSheetByName("appointment");
  const today = new Date();
  const firstDayOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  
  const data = appointmentsSheet.getDataRange().getValues();
  const distribution = {};
  
  for (let i = 1; i < data.length; i++) {
    const appointmentDate = new Date(data[i][1]); // Column B is appointment date
    if (appointmentDate >= firstDayOfMonth) {
      const type = data[i][8]; // Column I is appointment type
      distribution[type] = (distribution[type] || 0) + 1;
    }
  }
  
  return distribution;
}

function getWeeklyTrends() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appointmentsSheet = ss.getSheetByName("appointment");
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay());
  startOfWeek.setHours(0, 0, 0, 0);
  
  const data = appointmentsSheet.getDataRange().getValues();
  const weeklyData = [0, 0, 0, 0, 0, 0, 0]; // Initialize array for 7 days
  
  for (let i = 1; i < data.length; i++) {
    const appointmentDate = new Date(data[i][1]); // Column B is appointment date
    if (appointmentDate >= startOfWeek) {
      const dayIndex = appointmentDate.getDay();
      weeklyData[dayIndex]++;
    }
  }
  
  return weeklyData;
}

/**
 * Searches appointments based on various criteria
 * @param {Object} params - Search parameters
 * @param {string} params.mrdNo - MRD number to search for
 * @param {string} params.patientName - Patient name to search for
 * @param {string} params.phone - Phone number to search for
 * @param {string} params.fourDayConfirm - 4-day confirmation status
 * @param {string} params.oneDayConfirm - 1-day confirmation status
 * @param {number} params.page - Page number for pagination
 * @param {number} params.itemsPerPage - Number of items per page
 * @return {Object} Search results with appointments and total count
 */
function searchAppointments(params) {
  try {
    Logger.log("Starting searchAppointments with params: " + JSON.stringify(params));
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("appointment");
    if (!sheet) {
      Logger.log("Appointment sheet not found");
      return {
        success: false,
        message: "Appointment sheet not found"
      };
    }

    const data = sheet.getDataRange().getValues();
    
    // Define the correct column mapping
    const columnMap = {
      AppointmentID: 0,    // Column A
      AppointmentDate: 1,  // Column B
      TimeSlot: 2,        // Column C
      MRDNo: 3,           // Column D
      PatientName: 4,     // Column E
      Age: 5,             // Column F
      Gender: 6,          // Column G
      Phone: 7,           // Column H
      AppointmentType: 8, // Column I
      Doctor: 9,          // Column J
      PlanOfAction: 10,   // Column K
      Remarks: 11,        // Column L
      BookedBy: 12,       // Column M
      BookingDate: 13,    // Column N
      FourDayConfirmDate: 14, // Column O
      FourDayConfirmBy: 15,   // Column P
      OneDayConfirmDate: 16,  // Column Q
      OneDayConfirmBy: 17,    // Column R
      FixedSlot: 25       // Column Z
    };
    
    // Filter appointments based on search criteria
    let filteredData = data.slice(1).filter(row => {
      try {
        // MRD No filter
        if (params.mrdNo && row[columnMap.MRDNo] && 
            row[columnMap.MRDNo].toString().toLowerCase().indexOf(params.mrdNo.toLowerCase()) === -1) {
          return false;
        }
        
        // Patient name filter
        if (params.patientName && row[columnMap.PatientName] && 
            row[columnMap.PatientName].toString().toLowerCase().indexOf(params.patientName.toLowerCase()) === -1) {
          return false;
        }
        
        // Phone filter
        if (params.phone && row[columnMap.Phone] && 
            row[columnMap.Phone].toString().toLowerCase().indexOf(params.phone.toLowerCase()) === -1) {
          return false;
        }
        
        // 4-day confirmation filter
        if (params.fourDayConfirm) {
          const isConfirmed = row[columnMap.FourDayConfirmDate] !== "";
          if (params.fourDayConfirm === "confirmed" && !isConfirmed) return false;
          if (params.fourDayConfirm === "pending" && isConfirmed) return false;
        }
        
        // 1-day confirmation filter
        if (params.oneDayConfirm) {
          const isConfirmed = row[columnMap.OneDayConfirmDate] !== "";
          if (params.oneDayConfirm === "confirmed" && !isConfirmed) return false;
          if (params.oneDayConfirm === "pending" && isConfirmed) return false;
        }
        
        return true;
      } catch (rowError) {
        Logger.log("Error processing row: " + JSON.stringify(row) + " - " + rowError.toString());
        return false;
      }
    });
    
    Logger.log("Filtered data count: " + filteredData.length);
    
    // Calculate pagination
    const startIndex = ((params.page || 1) - 1) * (params.itemsPerPage || 10);
    const endIndex = startIndex + (params.itemsPerPage || 10);
    const paginatedData = filteredData.slice(startIndex, endIndex);
    
    // Convert rows to appointment objects with proper formatting
    const appointments = paginatedData.map(row => {
      try {
        const appointmentDate = row[columnMap.AppointmentDate] ? 
          Utilities.formatDate(new Date(row[columnMap.AppointmentDate]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
        
        const timeSlot = row[columnMap.TimeSlot] ? 
          Utilities.formatDate(new Date(row[columnMap.TimeSlot]), Session.getScriptTimeZone(), "HH:mm") : "";

           // Check Fixed Slot value and log it
        const fixedSlotValue = row[columnMap.FixedSlot];
        Logger.log("Fixed Slot value for appointment " + row[columnMap.AppointmentID] + ": " + fixedSlotValue + " (type: " + typeof fixedSlotValue + ")");
        
        // Convert FixedSlot to a more consistent value
        const isFixedSlot = (fixedSlotValue === "Yes" || fixedSlotValue === true || fixedSlotValue === "true");

        return {
          AppointmentID: row[columnMap.AppointmentID]?.toString() || "",
          AppointmentDate: appointmentDate,
          TimeSlot: timeSlot,
          MRDNo: row[columnMap.MRDNo]?.toString() || "",
          PatientName: row[columnMap.PatientName]?.toString() || "",
          Age: row[columnMap.Age]?.toString() || "",
          Gender: row[columnMap.Gender]?.toString() || "",
          Phone: row[columnMap.Phone]?.toString() || "",
          AppointmentType: row[columnMap.AppointmentType]?.toString() || "",
          Doctor: row[columnMap.Doctor]?.toString() || "",
          PlanOfAction: row[columnMap.PlanOfAction]?.toString() || "-",
          Remarks: row[columnMap.Remarks]?.toString() || "-",
          FourDayConfirmDate: row[columnMap.FourDayConfirmDate] ? 
            Utilities.formatDate(new Date(row[columnMap.FourDayConfirmDate]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          OneDayConfirmDate: row[columnMap.OneDayConfirmDate] ? 
            Utilities.formatDate(new Date(row[columnMap.OneDayConfirmDate]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          FixedSlot: isFixedSlot
        };
      } catch (error) {
        Logger.log("Error formatting appointment: " + error.toString());
        return null;
      }
    }).filter(appointment => appointment !== null); // Remove any failed conversions
    
    Logger.log("Returning " + appointments.length + " appointments");
    
    return {
      success: true,
      appointments: appointments,
      total: filteredData.length
    };
  } catch (error) {
    Logger.log("Error in searchAppointments: " + error.toString());
    return {
      success: false,
      message: "Error searching appointments: " + error.toString()
    };
  }
}

function getAppointmentDetails(appointmentId) {
  try {
    Logger.log("=== START getAppointmentDetails ===");
    Logger.log("Input appointmentId:", appointmentId);
    Logger.log("Current user:", Session.getActiveUser().getEmail());

    if (!appointmentId) {
      Logger.log("❌ No appointment ID provided");
      return { 
        success: false, 
        message: "No appointment ID provided",
        appointment: null
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("Active spreadsheet:", ss.getName());
    
    // Check both appointment and cancel sheets
    const appointmentSheet = ss.getSheetByName("appointment");
    const cancelSheet = ss.getSheetByName("cancel");
    
    if (!appointmentSheet && !cancelSheet) {
      Logger.log("❌ Neither appointment nor cancel sheet found");
      return { 
        success: false, 
        message: "Required sheets not found",
        appointment: null
      };
    }

    // Function to search in a sheet
    function searchInSheet(sheet) {
      if (!sheet) return null;
      const data = sheet.getDataRange().getValues();
      Logger.log("Total rows in sheet:", data.length);
      const headers = data[0];
      Logger.log("Headers found:", headers);
      
      // Decode and clean the appointment ID
      const searchId = decodeURIComponent(appointmentId).trim();
      Logger.log("Cleaned searchId:", searchId);
      
      // Find the appointment row
      Logger.log("Searching for appointment in rows...");
      const rowIndex = data.findIndex((row, index) => {
        const currentId = row[0]?.toString().trim();
        Logger.log(`Row ${index}: Comparing "${currentId}" with "${searchId}"`);
        return currentId === searchId;
      });
      Logger.log("Found row index:", rowIndex);

      if (rowIndex === -1) return null;

      // Build appointment object
      Logger.log("Building appointment object from row:", rowIndex + 1);
      const appointment = {};
      headers.forEach((header, index) => {
        const value = data[rowIndex][index];
        let formattedValue;
        
        if (value instanceof Date) {
          // Format dates consistently
          formattedValue = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else if (value === null || value === undefined) {
          formattedValue = "";
        } else {
          formattedValue = value.toString();
        }
        
        appointment[header] = formattedValue;
        Logger.log(`Column ${header}: ${formattedValue}`);
      });

      return appointment;
    }

    // First check appointment sheet, then cancel sheet
    const appointment = searchInSheet(appointmentSheet) || searchInSheet(cancelSheet);

    if (!appointment) {
      Logger.log("❌ Appointment not found in either sheet");
      return { 
        success: false, 
        message: `Appointment not found for ID: ${appointmentId}`,
        errorCode: 'APPOINTMENT_NOT_FOUND',
        appointment: null
      };
    }

    Logger.log("Built appointment object:", JSON.stringify(appointment, null, 2));
    Logger.log("=== END getAppointmentDetails ===");

    const response = { 
      success: true, 
      message: "Appointment details loaded",
      appointment: appointment
    };
    
    Logger.log("Returning response:", JSON.stringify(response));
    return response;
    
  } catch (e) {
    Logger.log("❌ Error in getAppointmentDetails:", e.toString());
    Logger.log("Stack trace:", e.stack);
    return { 
      success: false, 
      message: `Server error: ${e.toString()}`,
      errorDetails: e.stack,
      appointment: null
    };
  }
}


function getRescheduleAllowedSlots(doctor, appointmentType, date) {
  try {
    Logger.log("Starting getRescheduleAllowedSlots for doctor: " + doctor + ", type: " + appointmentType + ", date: " + date); 

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("appointment");
    if (!sheet) {
      return { success: false, message: "Sheet not found." };
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[0];

    // Get column indices
    var doctorCol = headers.indexOf("Doctor");
    var typeCol = headers.indexOf("AppointmentType");
    var dateCol = headers.indexOf("AppointmentDate");
    var timeSlotCol = headers.indexOf("TimeSlot");

    if (doctorCol === -1 || typeCol === -1 || dateCol === -1 || timeSlotCol === -1) {
       Logger.log("Required columns not found in appointment sheet");
      return { success: false, message: "Required columns not found." };
    }

    // Get all booked slots for the date
    var bookedSlots = [];
    for (var i = 1; i < values.length; i++) {
      if (values[i][dateCol] && 
          Utilities.formatDate(new Date(values[i][dateCol]), Session.getScriptTimeZone(), "yyyy-MM-dd") === date) {
         bookedSlots.push(Utilities.formatDate(new Date(values[i][timeSlotCol]), Session.getScriptTimeZone(), "HH:mm"));
      }
    }

    // Get all possible slots for the doctor and appointment type
    var allSlots = getAllTimeSlots(doctor, appointmentType);
    
        // Filter out booked slots
    var allowedSlots = allSlots.filter(slot => !bookedSlots.includes(slot));

    Logger.log("Found " + allowedSlots.length + " available slots for rescheduling"); 
    return { success: true, allowedSlots: allowedSlots };
  } catch (e) {
   Logger.log("Error in getRescheduleAllowedSlots: " + e.toString());
    return { success: false, message: "An error occurred while fetching allowed slots for rescheduling: " + e.toString() };
  }
}

function getAllTimeSlots(doctor, appointmentType) {
  // Define time slots based on doctor and appointment type
  // This is a simplified version - you may want to customize this based on your needs
  var slots = [];
  var startHour = 9; // 9 AM
  var endHour = 17; // 5 PM
  var interval = 30; // 30 minutes

  for (var hour = startHour; hour < endHour; hour++) {
    for (var minute = 0; minute < 60; minute += interval) {
      var time = Utilities.formatDate(
        new Date(2000, 0, 1, hour, minute),
        Session.getScriptTimeZone(),
        "HH:mm"
      );
      slots.push(time);
    }
  }
  return slots;
}

function cancelAppointment(appointmentId, sessionToken, callbackRequired = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentSheet = ss.getSheetByName("appointment");
    const cancelSheet = ss.getSheetByName("cancel");
    
    // Get the appointment data
    const data = appointmentSheet.getDataRange().getValues();
    const headers = data[0];
    const appointmentIdIndex = headers.indexOf("AppointmentID");
    
    // Find the appointment row index in the data array (including header)
    const appointmentRowIndex = data.findIndex((row, index) => 
      index > 0 && row[appointmentIdIndex] === appointmentId
    );
    
    if (appointmentRowIndex === -1) {
      return {
        success: false,
        message: "Appointment not found"
      };
    }
    
    const appointment = data[appointmentRowIndex];
    
    // Create a copy of the appointment row to avoid modifying the original
    const cancelledAppointment = [...appointment];
    
    // Ensure the row has enough columns (extend to at least column AH - index 33)
    while (cancelledAppointment.length < 34) {
      cancelledAppointment.push('');
    }
    
    // Set columns O, P, Q, R, Z, AB, AC to null when moving to cancel sheet
    cancelledAppointment[14] = ""; // Column O
    cancelledAppointment[15] = ""; // Column P
    cancelledAppointment[16] = ""; // Column Q
    cancelledAppointment[17] = ""; // Column R
    cancelledAppointment[25] = ""; // Column Z
    cancelledAppointment[27] = ""; // Column AB
    cancelledAppointment[28] = ""; // Column AC
    
    Logger.log("Set columns O, P, Q, R, Z, AB, AC to null for cancelled appointment");
    
    // Add cancellation date and time to the correct columns to match searchCancelledAppointments expectations
    const now = new Date();
    const username = getSessionUsername(sessionToken);
    
    cancelledAppointment[21] = now; // DeletedOn - Column V (index 21)
    cancelledAppointment[22] = username || Session.getActiveUser().getEmail(); // DeletedBy - Column W (index 22)
    
    // Add callback flag to column AH (index 33)
    cancelledAppointment[33] = callbackRequired ? 'Yes' : '';
    
    // Add to cancel sheet
    cancelSheet.appendRow(cancelledAppointment);
    
    // Delete from appointment sheet using the correct row index
    appointmentSheet.deleteRow(appointmentRowIndex + 1);
    
    Logger.log("Successfully cancelled appointment with columns O, P, Q, R, Z, AB, AC set to null");
    return {
      success: true,
      message: "Appointment cancelled successfully" + (callbackRequired ? " with callback scheduled" : "")
    };
  } catch (error) {
    return {
      success: false,
      message: "Error cancelling appointment: " + error.toString()
    };
  }
}

/**
 * Search patients based on various criteria
 * @param {Object} searchParams - Search parameters
 * @param {string} searchParams.mrdNo - MRD number to search for
 * @param {string} searchParams.patientName - Patient name to search for
 * @param {string} searchParams.phone - Phone number to search for
 * @param {string} searchParams.type - Patient type to search for
 * @param {string} searchParams.status - Patient status to search for
 * @param {number} searchParams.page - Page number for pagination
 * @param {number} searchParams.itemsPerPage - Number of items per page
 * @return {Object} Search results with patients and total count
 */
function searchPatients(searchParams) {
  try {
    Logger.log("Starting searchPatients with params: " + JSON.stringify(searchParams));
    
      const ss = SpreadsheetApp.getActiveSpreadsheet();
    const patientMasterSheet = ss.getSheetByName("patientMaster");
    //const tempPatientMasterSheet = ss.getSheetByName("tempPatientMaster");
    
    if (!patientMasterSheet) {
      Logger.log("patientMaster sheet not found");
      return {
        success: false,
         message: "Patient master sheet not found"
      };
    }


    
    // Define the correct column mapping
    const columnMap = {
      MRDNo: 0,           // Column A
      Name: 1,            // Column B
      DOB: 2,             // Column C
      Gender: 3,          // Column D
      Phone: 4,           // Column E
      Address: 5,         // Column F
      AdditionalInfo: 6,  // Column G
      Type: 7,            // Column H
      Status: 8,          // Column I
      RegisteredOn: 10    // Column K
    };
    
    // Function to filter data from a sheet
    function filterSheetData(sheet) {
      if (!sheet) return [];
      const data = sheet.getDataRange().getValues();
      return data.slice(1).filter(row => {
        try {
          // MRD No filter
          if (searchParams.mrdNo && row[columnMap.MRDNo] && 
              row[columnMap.MRDNo].toString().toLowerCase().indexOf(searchParams.mrdNo.toLowerCase()) === -1) {
            return false;
          }
          
          // Patient name filter
          if (searchParams.patientName && row[columnMap.Name] && 
              row[columnMap.Name].toString().toLowerCase().indexOf(searchParams.patientName.toLowerCase()) === -1) {
            return false;
          }
          
          // Phone filter
          if (searchParams.phone && row[columnMap.Phone] && 
              row[columnMap.Phone].toString().toLowerCase().indexOf(searchParams.phone.toLowerCase()) === -1) {
            return false;
          }
          
          // Type filter
          if (searchParams.type && row[columnMap.Type] && 
              row[columnMap.Type].toString().toLowerCase() !== searchParams.type.toLowerCase()) {
            return false;
          }
          
          // Status filter
          if (searchParams.status && row[columnMap.Status] && 
              row[columnMap.Status].toString().toLowerCase() !== searchParams.status.toLowerCase()) {
            return false;
          }
          
          return true;
        } catch (rowError) {
          Logger.log("Error processing row: " + JSON.stringify(row) + " - " + rowError.toString());
          return false;
        }
      });
    }
    
   // Get filtered data only from patientMaster sheet (not from tempPatientMaster)
    let filteredData = filterSheetData(patientMasterSheet);
    
    Logger.log("Filtered data count: " + filteredData.length);
    
    // Calculate pagination
    const startIndex = ((searchParams.page || 1) - 1) * (searchParams.itemsPerPage || 10);
    const endIndex = startIndex + (searchParams.itemsPerPage || 10);
    const paginatedData = filteredData.slice(startIndex, endIndex);
    
    // Convert rows to patient objects with proper formatting
    const patients = paginatedData.map(row => {
      try {
        return {
          MRDNo: row[columnMap.MRDNo]?.toString() || "",
          Name: row[columnMap.Name]?.toString() || "",
          DOB: row[columnMap.DOB] ? 
            Utilities.formatDate(new Date(row[columnMap.DOB]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          Gender: row[columnMap.Gender]?.toString() || "",
          Phone: row[columnMap.Phone]?.toString() || "",
          Address: row[columnMap.Address]?.toString() || "",
          AdditionalInfo: row[columnMap.AdditionalInfo]?.toString() || "",
          Type: row[columnMap.Type]?.toString() || "",
          Status: row[columnMap.Status]?.toString() || "",
          RegisteredOn: row[columnMap.RegisteredOn] ? 
            Utilities.formatDate(new Date(row[columnMap.RegisteredOn]), Session.getScriptTimeZone(), "yyyy-MM-dd") : ""
        };
      } catch (error) {
        Logger.log("Error formatting patient: " + error.toString());
        return null;
      }
    }).filter(patient => patient !== null); // Remove any failed conversions
    
    Logger.log("Returning " + patients.length + " patients");
    
    return {
      success: true,
      patients: patients,
      total: filteredData.length
    };
  } catch (error) {
    Logger.log("Error in searchPatients: " + error.toString());
    return {
      success: false,
      message: "Error searching patients: " + error.toString()
    };
  }
}

/**
 * Get details of a specific patient
 * @param {string} mrdNo - MRD number of the patient
 * @return {Object} Patient details
 */
function getPatientDetails(mrdNo) {
  try {
    Logger.log("Starting getPatientDetails for MRD: " + mrdNo);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("patientMaster");
    if (!sheet) {
      Logger.log("Patient master sheet not found");
      return {
        success: false,
        message: "Patient master sheet not found"
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the patient row
    const patientRow = data.slice(1).find(row => row[0].toString() === mrdNo);
    
    if (!patientRow) {
      Logger.log("Patient not found for MRD: " + mrdNo);
      return {
        success: false,
        message: "Patient not found"
      };
    }
    
    // Helper function to calculate exact age difference
    function calculateExactAge(dob) {
      var today = new Date();
      var birthDate = new Date(dob);
      
      // Calculate difference in milliseconds
      var diffInMs = today.getTime() - birthDate.getTime();
      var diffInDays = Math.floor(diffInMs / (1000 * 60 * 60 * 24));
      
      if (diffInDays < 30) {
        // Less than 30 days - return in days
        return diffInDays + " days";
      } else if (diffInDays < 365) {
        // 30 days or more but less than a year - return in months
        var months = Math.floor(diffInDays / 30);
        return months + " months";
      } else {
        // A year or more - return in years
        var years = Math.floor(diffInDays / 365);
        return years + " years";
      }
    }
    
    // Calculate exact age from DOB
    const dob = new Date(patientRow[2]);
    const exactAge = calculateExactAge(dob);
    
    return {
      success: true,
      patient: {
        MRDNo: patientRow[0],
        Name: patientRow[1],
        DOB: Utilities.formatDate(dob, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        Age: exactAge,
        Gender: patientRow[3],
        Phone: patientRow[4],
        Address: patientRow[5],
        AdditionalInfo: patientRow[6],
        Type: patientRow[7],
        Status: patientRow[8],
        RegisteredOn: patientRow[10] ? 
          Utilities.formatDate(new Date(patientRow[10]), Session.getScriptTimeZone(), "yyyy-MM-dd") : ""
      }
    };
  } catch (error) {
    Logger.log("Error in getPatientDetails: " + error.toString());
    return {
      success: false,
      message: "Error getting patient details: " + error.toString()
    };
  }
}

/**
 * Update patient details
 * @param {Object} updatedData - Updated patient data
 * @param {string} sessionToken - Session token for authentication
 * @return {Object} Update status
 */
function updatePatientDetails(updatedData, sessionToken) {
  try {
    // Handle both naming conventions (lowercase from frontend, uppercase for consistency)
    const mrdNo = updatedData.mrdNo || updatedData.MRDNo;
    const name = updatedData.name || updatedData.Name;
    const dob = updatedData.dob || updatedData.DOB;
    const gender = updatedData.gender || updatedData.Gender;
    const phone = updatedData.phone || updatedData.Phone;
    const address = updatedData.address || updatedData.Address;
    const additionalInfo = updatedData.additionalInfo || updatedData.AdditionalInfo;
    const type = updatedData.type || updatedData.Type;
    
    Logger.log("Starting updatePatientDetails for MRD: " + mrdNo);
    
    // Verify admin access
    const session = getUserSession(sessionToken);
    if (!session || session.role !== 'admin') {
      Logger.log("Unauthorized access attempt");
      return {
        success: false,
         message: "Unauthorized access. Admin or Manager privileges required."
      };
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("patientMaster");
    if (!sheet) {
      Logger.log("Patient master sheet not found");
      return {
        success: false,
        message: "Patient master sheet not found"
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the patient row
    const rowIndex = data.findIndex(row => row[0].toString() === mrdNo);
    
    if (rowIndex === -1) {
      Logger.log("Patient not found for MRD: " + mrdNo);
      return {
        success: false,
        message: "Patient not found"
      };
    }
    
    // Update the values
    const updateRow = rowIndex + 1; // +1 for header row
    sheet.getRange(updateRow, 2).setValue(name); // Column B
    sheet.getRange(updateRow, 3).setValue(new Date(dob)); // Column C
    sheet.getRange(updateRow, 4).setValue(gender); // Column D
    sheet.getRange(updateRow, 5).setValue(phone); // Column E
    sheet.getRange(updateRow, 6).setValue(address); // Column F
    sheet.getRange(updateRow, 7).setValue(additionalInfo); // Column G
    sheet.getRange(updateRow, 8).setValue(type); // Column H
    
    Logger.log("Patient details updated successfully");
    return {
      success: true,
      message: "Patient details updated successfully"
    };
  } catch (error) {
    Logger.log("Error in updatePatientDetails: " + error.toString());
    return {
      success: false,
      message: "Error updating patient details: " + error.toString()
    };
  }
}

/**
 * Toggle patient status between Active and Inactive
 * @param {string} mrdNo - MRD number of the patient
 * @param {string} sessionToken - Session token for authentication
 * @return {Object} Toggle status
 */
function togglePatientStatus(mrdNo, sessionToken) {
  try {
    Logger.log("Starting togglePatientStatus for MRD: " + mrdNo);
    
    // Verify admin access
    const session = getUserSession(sessionToken);
    if (!session || session.role !== 'admin') {
      Logger.log("Unauthorized access attempt");
      return {
        success: false,
        message: "Unauthorized access"
      };
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("patientMaster");
    if (!sheet) {
      Logger.log("Patient master sheet not found");
      return {
        success: false,
        message: "Patient master sheet not found"
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the patient row
    const rowIndex = data.findIndex(row => row[0].toString() === mrdNo);
    
    if (rowIndex === -1) {
      Logger.log("Patient not found for MRD: " + mrdNo);
      return {
        success: false,
        message: "Patient not found"
      };
    }
    
    // Toggle status
    const currentStatus = data[rowIndex][8]; // Column I
    const newStatus = currentStatus === "Active" ? "Inactive" : "Active";
    
    sheet.getRange(rowIndex + 1, 9).setValue(newStatus); // +1 for header row, Column I
    
    Logger.log("Patient status toggled successfully");
    return {
      success: true,
      message: "Patient status toggled successfully",
      newStatus: newStatus
    };
  } catch (error) {
    Logger.log("Error in togglePatientStatus: " + error.toString());
    return {
      success: false,
      message: "Error toggling patient status: " + error.toString()
    };
  }
}

function getReschedulePageUrl(appointmentId) {
  var sessionToken = Session.getActiveUser().getEmail();
  return ScriptApp.getService().getUrl() + '?page=rescheduleAppointment&appointmentId=' + appointmentId + '&sessionToken=' + sessionToken;
}

function showReschedulePage(appointmentId) {
  var template = HtmlService.createTemplateFromFile('rescheduleAppointment');
  template.appointmentId = appointmentId;
  template.sessionToken = Session.getActiveUser().getEmail();
  template.webAppUrl = ScriptApp.getService().getUrl();
  
  return template.evaluate()
    .setTitle('Reschedule Appointment')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Check if the user is an admin (Level 1 access)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if the user is an admin
 */
function isAdmin(sessionToken) {
  return hasAdminAccess(sessionToken);
}
/**
 * Get the email address of the logged-in user
 * @param {string} sessionToken - The session token
 * @return {Object} Object containing success status and email address
 */
function getUserEmail(sessionToken) {
  try {
    const session = getUserSession(sessionToken);
    if (!session) {
      return { success: false, message: "Invalid session" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Login");
    if (!sheet) {
      return { success: false, message: "Login sheet not found" };
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === session.username) {
        return { success: true, email: data[i][4] }; // Email is in column E
      }
    }

    return { success: false, message: "User email not found" };
  } catch (error) {
    Logger.log("Error in getUserEmail: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Send day overview email
 * @param {string} date - The date to send overview for
 * @param {string} sessionToken - The session token
 * @return {Object} Object containing success status and recipient email
 */
function sendDayOverviewEmail(date, sessionToken) {
  try {
    // Initialize spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get user email
    const userEmailResponse = getUserEmail(sessionToken);
    if (!userEmailResponse.success) {
      return { success: false, message: userEmailResponse.message };
    }
    const recipientEmail = userEmailResponse.email;

    // Get appointments for the day
    const appointments = getAppointmentsForDay(date);
    if (!appointments || appointments.length === 0) {
      return { success: false, message: "No appointments found for the selected date" };
    }

        // Sort appointments by time
    appointments.sort((a, b) => {
      const timeA = a.TimeSlot ? new Date(a.TimeSlot).getTime() : 0;
      const timeB = b.TimeSlot ? new Date(b.TimeSlot).getTime() : 0;
      return timeA - timeB;
    });


    // Create email content
    const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
    const subject = `Day Overview Report - ${formattedDate}`;

    // Count appointments by type and plan of action
    const summary = {
      total: appointments.length,
      confirmed: 0,
      waitingList: 0,
      planOfAction: {} // Object to store plan of action counts
    };

    appointments.forEach(appt => {
      if (appt.AppointmentType === 'WAITING LIST') {
        summary.waitingList++;
      } else {
        summary.confirmed++;
      }
      
      // Count plan of action
      if (appt.PlanOfAction) {
        try {
          const plans = JSON.parse(appt.PlanOfAction);
          const planArray = Array.isArray(plans) ? plans : [plans];
          planArray.forEach(plan => {
            let planName = typeof plan === 'object' ? plan.name : plan.toString().trim();
            // Remove eye specifications (LE, RE, BE) from the plan name
            planName = planName.replace(/\s*\([LRB]E\)/g, '').trim();
            summary.planOfAction[planName] = (summary.planOfAction[planName] || 0) + 1;
          });
        } catch (e) {
          // If not JSON, treat as single plan
          let planName = appt.PlanOfAction.toString().trim();
          // Remove eye specifications (LE, RE, BE) from the plan name
          planName = planName.replace(/\s*\([LRB]E\)/g, '').trim();
          summary.planOfAction[planName] = (summary.planOfAction[planName] || 0) + 1;
        }
      }
    });

    // Get available slots count
    const availableSlotsResp = getAvailableSlotsCount();
    const availableSlots = availableSlotsResp.count || 0;

    // Get blocked slots count for each doctor
    const doctorsResponse = getDoctors();
    let totalBlockedSlots = 0;
    
    if (doctorsResponse && doctorsResponse.success && Array.isArray(doctorsResponse.doctors)) {
      doctorsResponse.doctors.forEach(function(doctor) {
        const blockedSlots = getBlockedSlotsCount(doctor, date);
        totalBlockedSlots += blockedSlots;
      });
    }

    // Create HTML content with print-friendly styles
    let htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          @media print {
            body { margin: 0; padding: 0; }
            .page { margin: 0; padding: 0; }
            .no-print { display: none; }
          }
          
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 0;
          }
          
          .page {
            width: 210mm;
            min-height: 297mm;
            padding: 20mm;
            margin: 0 auto;
            background: white;
            box-sizing: border-box;
          }
          
          .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid #2c5282;
          }
          
          .header h1 {
            color: #2c5282;
            margin: 0;
            font-size: 24px;
          }
          
          .header h2 {
            color: #4a5568;
            margin: 10px 0 0;
            font-size: 18px;
            font-weight: normal;
          }
          
          .summary-section {
            background-color: #f8fafc;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
          }
          
          .summary-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 20px;
            margin-bottom: 20px;
          }
          
          .summary-item {
            background: white;
            padding: 15px;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          .summary-item h3 {
            margin: 0 0 10px;
            color: #2d3748;
            font-size: 16px;
          }
          
          .summary-value {
            font-size: 24px;
            font-weight: bold;
            color: #2c5282;
          }
          
          .plan-breakdown {
            background: white;
            padding: 15px;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          .plan-breakdown h3 {
            margin: 0 0 15px;
            color: #2d3748;
            font-size: 16px;
          }
          
          .plan-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 10px;
          }
          
          .plan-item {
            display: flex;
            justify-content: space-between;
            padding: 8px;
            background: #f8fafc;
            border-radius: 4px;
          }
          
          .appointments-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 12px;
          }
          
          .appointments-table th {
            background-color: #2c5282;
            color: white;
            padding: 12px;
            text-align: left;
            border: 1px solid #e2e8f0;
          }
          
          .appointments-table td {
            padding: 10px;
            border: 1px solid #e2e8f0;
          }
          
          .appointments-table tr:nth-child(even) {
            background-color: #f8fafc;
          }
          
          .waiting-list {
            background-color: #f3e8ff !important;
          }
          
          .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e2e8f0;
            text-align: center;
            color: #718096;
            font-size: 12px;
          }
        </style>
      </head>
      <body>
        <div class="page">
          <div class="header">
            <h1>Day Overview Report</h1>
            <h2>${formattedDate}</h2>
          </div>

          <div class="summary-section">
            <div class="summary-grid">
              <div class="summary-item">
                <h3>Total Appointments</h3>
                <div class="summary-value">${summary.total}</div>
              </div>
              <div class="summary-item">
                <h3>Confirmed Appointments</h3>
                <div class="summary-value">${summary.confirmed}</div>
              </div>
              <div class="summary-item">
                <h3>Waiting List</h3>
                <div class="summary-value">${summary.waitingList}</div>
              </div>
              <div class="summary-item">
                <h3>Available Slots</h3>
                <div class="summary-value">${availableSlots}</div>
              </div>
              <div class="summary-item">
                <h3>Free Blocked Slots</h3>
                <div class="summary-value">${totalBlockedSlots}</div>
              </div>
            </div>

            <div class="plan-breakdown">
              <h3>Plan of Action Breakdown</h3>
              <div class="plan-grid">
                ${Object.entries(summary.planOfAction)
                  .map(([plan, count]) => `
                    <div class="plan-item">
                      <span>${plan}</span>
                      <span style="font-weight: bold;">${count}</span>
                    </div>
                  `).join('')}
              </div>
            </div>
          </div>

          <table class="appointments-table">
            <thead>
              <tr>
                <th>Time</th>
                <th>MRD</th>
                <th>Name</th>
                <th>Gender</th>
                <th>Age</th>
                <th>Phone</th>
                <th>Patient Type</th>
                <th>Plan</th>
                <th>Remarks</th>
              </tr>
            </thead>
            <tbody>
              ${appointments.map(appt => `
                <tr class="${appt.AppointmentType === 'WAITING LIST' ? 'waiting-list' : ''}">
                  <td>${Utilities.formatDate(new Date(appt.TimeSlot), Session.getScriptTimeZone(), "hh:mm a")}</td>
                  <td>${appt.MRDNo}</td>
                  <td>${appt.PatientName}</td>
                  <td>${appt.Gender}</td>
                  <td>${appt.Age}</td>
                  <td>${appt.Phone}</td>
                  <td>${appt.PatientType || '-'}</td>
                  <td>${appt.PlanOfAction || '-'}</td>
                  <td>${appt.Remarks || '-'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div class="footer">
            <p>This is an automated report generated from the Appointment Management System.</p>
            <p>Generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' hh:mm a")}</p>
          </div>
        </div>
      </body>
      </html>
    `;

    // Send email
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlContent
    });

    // Log the email send in audit sheet
    const auditSheet = ss.getSheetByName("EmailAudit") || ss.insertSheet("EmailAudit");
    if (auditSheet.getLastRow() === 0) {
      auditSheet.appendRow(["Timestamp", "Date", "Recipient", "Sent By", "Status"]);
    }
    auditSheet.appendRow([
      new Date(),
      date,
      recipientEmail,
      userEmailResponse.username,
      "Success"
    ]);

    return { 
      success: true, 
      message: "Email sent successfully",
      recipient: recipientEmail
    };
  } catch (error) {
    Logger.log("Error in sendDayOverviewEmail: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Search cancelled appointments based on various criteria
 * @param {Object} params - Search parameters
 * @param {string} params.mrdNo - MRD number to search for
 * @param {string} params.patientName - Patient name to search for
 * @param {string} params.phone - Phone number to search for
 * @param {number} params.page - Page number for pagination
 * @param {number} searchParams.itemsPerPage - Number of items per page
 * @return {Object} Search results with appointments and total count
 */
function searchCancelledAppointments(params) {
  try {
    Logger.log("Starting searchCancelledAppointments with params: " + JSON.stringify(params));
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cancel");
    if (!sheet) {
      Logger.log("Cancel sheet not found");
      return {
        success: false,
        message: "Cancel sheet not found"
      };
    }

    const data = sheet.getDataRange().getValues();
    
    // Define the correct column mapping
    const columnMap = {
      AppointmentID: 0,    // Column A
      AppointmentDate: 1,  // Column B
      TimeSlot: 2,        // Column C
      MRDNo: 3,           // Column D
      PatientName: 4,     // Column E
      Age: 5,             // Column F
      Gender: 6,          // Column G
      Phone: 7,           // Column H
      AppointmentType: 8, // Column I
      Doctor: 9,          // Column J
      PlanOfAction: 10,   // Column K
      Remarks: 11,        // Column L
      BookedBy: 12,       // Column M
      BookingDate: 13,    // Column N
      FourDayConfirmDate: 14, // Column O
      FourDayConfirmBy: 15,   // Column P
      OneDayConfirmDate: 16,  // Column Q
      OneDayConfirmBy: 17,    // Column R
      RescheduledDate: 18,    // Column S
      RescheduledBy: 19,      // Column T
      RescheduleHistory: 20,  // Column U
      DeletedOn: 21,          // Column V
      DeletedBy: 22,          // Column W
      ColumnAH: 33,           // Column AH
      CallbackDate: 34        // Column AI
    };
    
    // Get today's date for overdue calculation
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Filter appointments based on search criteria
    let filteredData = data.slice(1).filter(row => {
      try {
        // MRD No filter
        if (params.mrdNo && row[columnMap.MRDNo] && 
            row[columnMap.MRDNo].toString().toLowerCase().indexOf(params.mrdNo.toLowerCase()) === -1) {
          return false;
        }
        
        // Patient name filter
        if (params.patientName && row[columnMap.PatientName] && 
            row[columnMap.PatientName].toString().toLowerCase().indexOf(params.patientName.toLowerCase()) === -1) {
          return false;
        }
        
        // Phone filter
        if (params.phone && row[columnMap.Phone] && 
            row[columnMap.Phone].toString().toLowerCase().indexOf(params.phone.toLowerCase()) === -1) {
          return false;
        }
        
        // Callback date range filter
        if (params.callbackFrom || params.callbackTo) {
          const callbackDate = row[columnMap.CallbackDate];
          if (callbackDate) {
            const cbDate = new Date(callbackDate);
            if (params.callbackFrom) {
              const fromDate = new Date(params.callbackFrom);
              if (cbDate < fromDate) return false;
            }
            if (params.callbackTo) {
              const toDate = new Date(params.callbackTo);
              if (cbDate > toDate) return false;
            }
          } else if (params.callbackFrom || params.callbackTo) {
            // If filtering by callback date but appointment has no callback date, exclude it
            return false;
          }
        }
        
        // Card color filter
        if (params.cardColor) {
          const isPastelMauve = row[columnMap.ColumnAH] && row[columnMap.ColumnAH].toString().toLowerCase() === 'yes';
          const callbackDate = row[columnMap.CallbackDate];
          const isOverdue = callbackDate && new Date(callbackDate) < today;
          
          switch (params.cardColor) {
            case 'overdue':
              if (!isOverdue) return false;
              break;
            case 'pastel-mauve':
              if (!isPastelMauve) return false;
              break;
            case 'ontime':
              if (isOverdue || isPastelMauve) return false;
              break;
          }
        }
        
        return true;
      } catch (rowError) {
        Logger.log("Error processing row: " + JSON.stringify(row) + " - " + rowError.toString());
        return false;
      }
    });
    
    Logger.log("Filtered data count: " + filteredData.length);
    
    // Calculate pagination
    const startIndex = ((params.page || 1) - 1) * (params.itemsPerPage || 10);
    const endIndex = startIndex + (params.itemsPerPage || 10);
    const paginatedData = filteredData.slice(startIndex, endIndex);
    
    // Convert rows to appointment objects with proper formatting
    const appointments = paginatedData.map(row => {
      try {
        const appointmentDate = row[columnMap.AppointmentDate] ? 
          Utilities.formatDate(new Date(row[columnMap.AppointmentDate]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
        
        const timeSlot = row[columnMap.TimeSlot] ? 
          Utilities.formatDate(new Date(row[columnMap.TimeSlot]), Session.getScriptTimeZone(), "HH:mm") : "";

        // Calculate status flags
        const isPastelMauve = row[columnMap.ColumnAH] && row[columnMap.ColumnAH].toString().toLowerCase() === 'yes';
        const callbackDate = row[columnMap.CallbackDate];
        const isOverdue = callbackDate && new Date(callbackDate) < today;
        
        const callbackDateFormatted = callbackDate ? 
          Utilities.formatDate(new Date(callbackDate), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";

        return {
          AppointmentID: row[columnMap.AppointmentID]?.toString() || "",
          AppointmentDate: appointmentDate,
          TimeSlot: timeSlot,
          MRDNo: row[columnMap.MRDNo]?.toString() || "",
          PatientName: row[columnMap.PatientName]?.toString() || "",
          Age: row[columnMap.Age]?.toString() || "",
          Gender: row[columnMap.Gender]?.toString() || "",
          Phone: row[columnMap.Phone]?.toString() || "",
          AppointmentType: row[columnMap.AppointmentType]?.toString() || "",
          Doctor: row[columnMap.Doctor]?.toString() || "",
          PlanOfAction: row[columnMap.PlanOfAction]?.toString() || "-",
          Remarks: row[columnMap.Remarks]?.toString() || "-",
          DeletedOn: row[columnMap.DeletedOn] ? 
            Utilities.formatDate(new Date(row[columnMap.DeletedOn]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
          DeletedBy: row[columnMap.DeletedBy]?.toString() || "",
          RescheduleHistory: row[columnMap.RescheduleHistory]?.toString() || "",
          CallbackDate: callbackDateFormatted,
          isOverdue: isOverdue,
          isPastelMauve: isPastelMauve,
          ColumnAH: row[columnMap.ColumnAH]?.toString() || ""
        };
      } catch (error) {
        Logger.log("Error formatting appointment: " + error.toString());
        return null;
      }
    }).filter(appointment => appointment !== null);
    
    Logger.log("Returning " + appointments.length + " appointments");
    
    return {
      success: true,
      appointments: appointments,
      total: filteredData.length
    };
  } catch (error) {
    Logger.log("Error in searchCancelledAppointments: " + error.toString());
    return {
      success: false,
      message: "Error searching cancelled appointments: " + error.toString()
    };
  }
}

/**
 * Update a cancelled appointment's details
 * @param {Object} updatedData - Updated appointment data
 * @param {string} sessionToken - Session token for authentication
 * @return {Object} Update status
 */
function updateCancelledAppointmentDetails(updatedData, sessionToken) {
  try {
    Logger.log("Starting updateCancelledAppointmentDetails for appointmentId: " + updatedData.appointmentId);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("cancel");
    if (!sheet) {
      Logger.log("Sheet 'cancel' not found.");
      return { success: false, message: "Sheet not found." };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rowIndex = -1;

    // Find the appointment row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === updatedData.appointmentId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      Logger.log("Appointment ID not found: " + updatedData.appointmentId);
      return { success: false, message: "Appointment not found." };
    }

    // Get column indices for the fields to update
    const typeCol = headers.indexOf("AppointmentType");
    const doctorCol = headers.indexOf("Doctor");
    const planCol = headers.indexOf("PlanOfAction");
    const remarksCol = headers.indexOf("Remarks");

    // Update the values in the sheet
    if (typeCol !== -1) sheet.getRange(rowIndex + 1, typeCol + 1).setValue(updatedData.type);
    if (doctorCol !== -1) sheet.getRange(rowIndex + 1, doctorCol + 1).setValue(updatedData.doctor);
    if (planCol !== -1) sheet.getRange(rowIndex + 1, planCol + 1).setValue(updatedData.plan);
    if (remarksCol !== -1) sheet.getRange(rowIndex + 1, remarksCol + 1).setValue(updatedData.remarks);

    // Update callback date if provided
    if (updatedData.callbackDate !== undefined) {
      const callbackDateValue = updatedData.callbackDate ? new Date(updatedData.callbackDate) : "";
      sheet.getRange(rowIndex + 1, 35).setValue(callbackDateValue); // Column AI (35th column)
    }

    Logger.log("Cancelled appointment details updated successfully.");
    return { success: true, message: "Cancelled appointment details updated successfully." };
  } catch (e) {
    Logger.log("Error in updateCancelledAppointmentDetails: " + e.toString());
    return { success: false, message: "An error occurred while updating cancelled appointment details: " + e.toString() };
  }
}

/**
 * Update only the callback date for a cancelled appointment
 * @param {Object} updatedData - Updated appointment data with callback date
 * @param {string} sessionToken - Session token for authentication
 * @return {Object} Update status
 */
function updateCancelledAppointmentCallbackDate(updatedData, sessionToken) {
  try {
    Logger.log("Starting updateCancelledAppointmentCallbackDate for appointmentId: " + updatedData.appointmentId);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("cancel");
    if (!sheet) {
      Logger.log("Sheet 'cancel' not found.");
      return { success: false, message: "Sheet not found." };
    }

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // Find the appointment row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === updatedData.appointmentId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      Logger.log("Appointment ID not found: " + updatedData.appointmentId);
      return { success: false, message: "Appointment not found." };
    }

    // Update callback date in column AI (35th column)
    const callbackDateValue = updatedData.callbackDate ? new Date(updatedData.callbackDate) : "";
    sheet.getRange(rowIndex + 1, 35).setValue(callbackDateValue);

    Logger.log("Callback date updated successfully.");
    return { success: true, message: "Callback date updated successfully." };
  } catch (e) {
    Logger.log("Error in updateCancelledAppointmentCallbackDate: " + e.toString());
    return { success: false, message: "An error occurred while updating callback date: " + e.toString() };
  }
}

/**
 * Restore a cancelled appointment back to the appointment sheet
 * @param {string} appointmentId - ID of the appointment to restore
 * @param {Object} updatedData - Updated appointment data
 * @param {string} sessionToken - Session token for authentication
 * @return {Object} Restore status
 */
function restoreCancelledAppointment(rescheduleData) {
  try {
    Logger.log("Starting restoreCancelledAppointment for appointmentId: " + rescheduleData.appointmentId);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cancelSheet = ss.getSheetByName("cancel");
    const appointmentSheet = ss.getSheetByName("appointment");
    
    if (!cancelSheet || !appointmentSheet) {
      return { success: false, message: "Required sheets not found." };
    }

    const data = cancelSheet.getDataRange().getValues();
    const headers = data[0];
    let rowIndex = -1;

    // Find the appointment row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === rescheduleData.appointmentId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Appointment not found in cancelled list." };
    }

    // Get the row data (create a copy to avoid modifying original)
    const rowData = [...data[rowIndex]];

    // Update the row with new data
    const updateMap = {
      'AppointmentDate': rescheduleData.newDate,
      'TimeSlot': rescheduleData.newTimeSlot,
      'AppointmentType': rescheduleData.appointmentType,
      'Doctor': rescheduleData.doctor,
      'PlanOfAction': rescheduleData.newPlanOfAction,
      'Remarks': rescheduleData.newRemarks
    };

    // Update the values
    Object.entries(updateMap).forEach(([header, value]) => {
      const colIndex = headers.indexOf(header);
      if (colIndex !== -1) {
        rowData[colIndex] = value;
      }
    });

    // Set columns O, P, Q, R, Z, AB, AC, AH to null when moving to appointment sheet
    // Ensure the array is long enough to accommodate these indices (up to AH = index 33)
    while (rowData.length < 34) {
      rowData.push("");
    }
    
    rowData[14] = ""; // Column O
    rowData[15] = ""; // Column P
    rowData[16] = ""; // Column Q
    rowData[17] = ""; // Column R
    rowData[25] = ""; // Column Z
    rowData[27] = ""; // Column AB
    rowData[28] = ""; // Column AC
    rowData[33] = ""; // Column AH
    
    Logger.log("Set columns O, P, Q, R, Z, AB, AC, AH to null for restored appointment");

    // Add to reschedule history
    const historyCol = headers.indexOf("RescheduleHistory");
    if (historyCol !== -1) {
      const history = rowData[historyCol] ? JSON.parse(rowData[historyCol]) : [];
      
      // Get the original appointment date and time before updating
      const originalDate = rowData[headers.indexOf("AppointmentDate")];
      const originalTime = rowData[headers.indexOf("TimeSlot")];
      
      history.push({
        timestamp: new Date(),
        action: "restored",
        by: getSessionUsername(rescheduleData.sessionToken),
        fromDate: originalDate,
        toDate: rescheduleData.newDate,
        fromTime: originalTime,
        toTime: rescheduleData.newTimeSlot
      });
      rowData[historyCol] = JSON.stringify(history);
      
      Logger.log("Added restore history entry: " + JSON.stringify({
        timestamp: new Date(),
        action: "restored",
        by: getSessionUsername(rescheduleData.sessionToken),
        fromDate: originalDate,
        toDate: rescheduleData.newDate,
        fromTime: originalTime,
        toTime: rescheduleData.newTimeSlot
      }));
    } else {
      Logger.log("Warning: RescheduleHistory column not found in cancel sheet headers");
      Logger.log("Available headers: " + headers.join(", "));
    }

    // Move to appointment sheet
    appointmentSheet.appendRow(rowData);
    
    // Delete from cancel sheet
    cancelSheet.deleteRow(rowIndex + 1);

    Logger.log("Successfully restored appointment with columns O, P, Q, R, Z, AB, AC, AH set to null");
    Logger.log("Reschedule history added to column U: " + (rowData[historyCol] || "No history"));
    return { success: true, message: "Appointment restored successfully." };
  } catch (e) {
    Logger.log("Error in restoreCancelledAppointment: " + e.toString());
    return { success: false, message: "An error occurred while restoring appointment: " + e.toString() };
  }
}


/**
 * Archive a cancelled appointment to the archive sheet
 * @param {string} appointmentId - ID of the appointment to archive
 * @param {string} sessionToken - Session token for authentication
 * @return {Object} Archive status
 */
function archiveCancelledAppointment(appointmentId, sessionToken) {
  try {
    Logger.log("Starting archiveCancelledAppointment for appointmentId: " + appointmentId);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cancelSheet = ss.getSheetByName("cancel");
    const archiveSheet = ss.getSheetByName("Archive");
    
    if (!cancelSheet || !archiveSheet) {
      return { success: false, message: "Required sheets not found." };
    }

    // Get the full data range to ensure all columns are included
    const lastRow = cancelSheet.getLastRow();
    const lastCol = cancelSheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      return { success: false, message: "Cancel sheet is empty." };
    }

    const data = cancelSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    let rowIndex = -1;

    // Find the appointment row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === appointmentId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Appointment not found in cancelled list." };
    }

    // Get the complete row data (ensuring all columns including V and W)
    const rowData = [...data[rowIndex]]; // Create a copy of the array
    
    // Log column V and W values for verification (columns 21 and 22 in 0-based indexing)
    Logger.log("Column V value: " + (rowData[21] || "empty"));
    Logger.log("Column W value: " + (rowData[22] || "empty"));
    Logger.log("Total columns being transferred: " + rowData.length);

    // Add to reschedule history
    const historyCol = headers.indexOf("RescheduleHistory");
    if (historyCol !== -1) {
      const history = rowData[historyCol] ? JSON.parse(rowData[historyCol]) : [];
      history.push({
        timestamp: new Date(),
        action: "archived",
        by: getSessionUsername(sessionToken)
      });
      rowData[historyCol] = JSON.stringify(history);
    }

    // Ensure Archive sheet has proper headers including the new archive columns
    const archiveLastCol = archiveSheet.getLastColumn();
    const requiredColumns = Math.max(lastCol, 25); // Ensure we have at least 25 columns (up to column Y)
    
    if (archiveLastCol < requiredColumns) {
      // Extend the header row with missing headers
      const archiveHeaders = archiveSheet.getRange(1, 1, 1, Math.min(archiveLastCol, 1)).getValues()[0] || [];
      const extendedHeaders = [...headers]; // Copy original headers
      
      // Ensure we have headers for columns X and Y
      while (extendedHeaders.length < 23) extendedHeaders.push(""); // Fill up to column W
      if (extendedHeaders.length === 23) extendedHeaders.push("ArchivedDate"); // Column X
      if (extendedHeaders.length === 24) extendedHeaders.push("ArchivedBy");   // Column Y
      
      // Set the complete header row
      archiveSheet.getRange(1, 1, 1, extendedHeaders.length).setValues([extendedHeaders]);
    }

    // Extend rowData to include archived date and archived by information
    // Ensure we have at least 25 columns (up to column Y)
    while (rowData.length < 25) {
      rowData.push(""); // Fill empty columns
    }
    
    // Set Column X (index 23) - Archived Date
    rowData[23] = new Date();
    
    // Set Column Y (index 24) - Archived By  
    rowData[24] = getSessionUsername(sessionToken);
    
    // Log the final data being added for debugging
    Logger.log("Final rowData length: " + rowData.length);
    Logger.log("Column X (Archived Date): " + rowData[23]);
    Logger.log("Column Y (Archived By): " + rowData[24]);

    // Move to archive sheet - ensure all columns are transferred
    archiveSheet.appendRow(rowData);
    
    // Delete from cancel sheet
    cancelSheet.deleteRow(rowIndex + 1);

    Logger.log("Successfully archived appointment with " + rowData.length + " columns");
    return { success: true, message: "Appointment archived successfully with all columns including V and W." };
  } catch (e) {
    Logger.log("Error in archiveCancelledAppointment: " + e.toString());
    return { success: false, message: "An error occurred while archiving appointment: " + e.toString() };
  }
}

// Add this function to check if a slot is fixed
function isSlotFixed(doctor, dateString, timeSlot) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("appointment");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowDate = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const rowTime = Utilities.formatDate(new Date(data[i][2]), Session.getScriptTimeZone(), "HH:mm");
      const rowDoctor = data[i][9];
      const isFixed = data[i][25]; // Column Z (0-based index)
      
      if (rowDate === dateString && 
          rowTime === timeSlot && 
          rowDoctor === doctor && 
          isFixed === "Yes") {
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log("Error in isSlotFixed: " + e.toString());
    return false;
  }
}


/**
 * Check if the user has access to the Patient Master page (Level 1, 2, or 3)
 * @param {string} sessionToken - The session token to check
 * @return {boolean} True if the user has access (Level 1, 2, or 3)
 */
function hasAccessToPatientMaster(sessionToken) {
  return canUpdatePatientDetails(sessionToken);
}


/**
 * Get patient type options from generalSettings sheet
 * @return {Object} Object containing success status and patient type options
 */
function getPatientTypeOptions() {
  try {
    Logger.log("Getting patient type options from generalSettings sheet");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("generalSettings");
    
    if (!sheet) {
      Logger.log("generalSettings sheet not found");
      return {
        success: false,
        message: "generalSettings sheet not found"
      };
    }
    
    // Get values from column B (skip header row)
    const data = sheet.getRange("B2:B100").getValues();
    const typeOptions = [];
    
    // Process each row and add non-empty values to the options array
    for (let i = 0; i < data.length; i++) {
      const value = data[i][0];
      if (value && typeof value === 'string' && value.trim() !== "") {
        typeOptions.push(value.trim());
      }
    }
    
    Logger.log("Found " + typeOptions.length + " patient type options");
    
    return {
      success: true,
      typeOptions: typeOptions
    };
  } catch (error) {
    Logger.log("Error in getPatientTypeOptions: " + error.toString());
    return {
      success: false,
      message: "Error retrieving patient type options: " + error.toString()
    };
  }
}

/**
 * Returns the URL for the Dashboard page
 * This function helps with proper navigation between pages
 * @param {string} sessionToken - The session token to include in the URL
 * @return {string} The fully formed URL to the dashboard page
 */
function getDashboardUrl(sessionToken) {
  try {
    Logger.log("Getting dashboard URL for session: " + sessionToken);
    
    // Validate the session first
    const session = getUserSession(sessionToken);
    if (!session) {
      Logger.log("Invalid session token in getDashboardUrl");
      // If session is invalid, return URL to login page
      return ScriptApp.getService().getUrl();
    }
    
    // Construct the proper URL
    const baseUrl = ScriptApp.getService().getUrl();
    const dashboardUrl = baseUrl + "?page=Dashboard&sessionToken=" + encodeURIComponent(sessionToken);
    
    Logger.log("Dashboard URL: " + dashboardUrl);
    return dashboardUrl;
  } catch (error) {
    Logger.log("Error in getDashboardUrl: " + error.toString());
    return ScriptApp.getService().getUrl(); // Return base URL as fallback
  }
}

function getPlanOfActionOptions() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    // ← Use your actual sheet name here
    const sheet = ss.getSheetByName("PlanOfActionMaster");
    if (!sheet) return { success: false, message: "PlanOfActionMaster sheet not found" };
    const data = sheet.getDataRange().getValues();
    const options = data.slice(1)
      .filter(r => r[0])
      .map(r => ({
        name: r[0].toString().trim(),
        requiresEyeSelection: String(r[1]).toLowerCase() === "yes"
      }));
    return { success: true, options };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Deletes an appointment by ID
 * @param {string} appointmentId - The ID of the appointment to delete
 * @param {string} sessionToken - The session token for authentication
 * @return {Object} Status object indicating success or failure
 */
function deleteAppointment(appointmentId, sessionToken) {
  try {
    Logger.log("deleteAppointment: Starting deletion for appointment ID: " + appointmentId);
    
    // Validate session
    const session = getUserSession(sessionToken);
    if (!session || !session.username) {
      Logger.log("deleteAppointment: Invalid session when deleting appointment");
      return { success: false, message: "Not authorized: Invalid session" };
    }
    Logger.log("deleteAppointment: Session validated for user: " + session.username);
    
    // Get the spreadsheet - using getActive() instead of openById
    const ss = SpreadsheetApp.getActive();
    Logger.log("deleteAppointment: Active spreadsheet name: " + ss.getName());
    
    // List all sheets for debugging
    const allSheets = ss.getSheets();
    const sheetNames = allSheets.map(sheet => sheet.getName());
    Logger.log("deleteAppointment: All sheets in spreadsheet: " + JSON.stringify(sheetNames));
    
    const appointmentsSheet = ss.getSheetByName("Appointments");
    
    if (!appointmentsSheet) {
      Logger.log("deleteAppointment: ERROR - Appointments sheet not found in spreadsheet: " + ss.getName());
      
      // Try alternative sheet names
      const possibleSheetNames = ["appointments", "APPOINTMENTS", "AppointmentsData", "appointment", "Appointment"];
      Logger.log("deleteAppointment: Trying alternative sheet names...");
      
      for (const sheetName of possibleSheetNames) {
        const altSheet = ss.getSheetByName(sheetName);
        if (altSheet) {
          Logger.log("deleteAppointment: Found alternative sheet: " + sheetName);
          // Use this sheet instead
          return deleteAppointmentFromSheet(appointmentId, altSheet, session);
        }
      }
      
      return { success: false, message: "Error: Appointments sheet not found" };
    }
    
    Logger.log("deleteAppointment: Appointments sheet found, proceeding with deletion");
    return deleteAppointmentFromSheet(appointmentId, appointmentsSheet, session);
    
  } catch (error) {
    Logger.log(`deleteAppointment: Error deleting appointment: ${error.toString()}`);
    return { success: false, message: `Error: ${error.toString()}` };
  }
}

/**
 * Helper function to delete an appointment from a specific sheet
 * @param {string} appointmentId - The ID of the appointment to delete
 * @param {Sheet} appointmentsSheet - The sheet containing the appointment data
 * @param {Object} session - The user session object
 * @return {Object} Status object indicating success or failure
 */
function deleteAppointmentFromSheet(appointmentId, appointmentsSheet, session) {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Get all data
    const data = appointmentsSheet.getDataRange().getValues();
    const headers = data[0];
    Logger.log("deleteAppointmentFromSheet: Headers found: " + JSON.stringify(headers));
    
    const idColIndex = headers.indexOf("AppointmentID");
    
    if (idColIndex === -1) {
      Logger.log("deleteAppointmentFromSheet: AppointmentID column not found in headers");
      return { success: false, message: "Error: AppointmentID column not found" };
    }
    
    // Find the appointment row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] == appointmentId) {
        rowIndex = i + 1; // +1 because array is 0-indexed but sheet is 1-indexed
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log("deleteAppointmentFromSheet: Appointment ID " + appointmentId + " not found in sheet");
      return { success: false, message: "Appointment not found" };
    }
    
    Logger.log("deleteAppointmentFromSheet: Found appointment at row " + rowIndex);
    
    // Instead of deleting the row, we can move it to a "Cancelled" or "Deleted" sheet to keep history
    const cancelledSheet = ss.getSheetByName("cancel");
    
    if (cancelledSheet) {
      Logger.log("deleteAppointmentFromSheet: Found cancel sheet, copying data before deletion");
      // Get the row data to move
      const rowData = appointmentsSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      
      // Add cancellation metadata
      const cancelledByColIndex = headers.indexOf("CancelledBy");
      const cancelDateColIndex = headers.indexOf("CancelDate");
      
      if (cancelledByColIndex !== -1 && cancelDateColIndex !== -1) {
        rowData[cancelledByColIndex] = session.username;
        rowData[cancelDateColIndex] = new Date();
        Logger.log("deleteAppointmentFromSheet: Added cancellation metadata");
      } else {
        Logger.log("deleteAppointmentFromSheet: Cancellation metadata columns not found, continuing without metadata");
      }
      
      // Append to cancelled sheet
      cancelledSheet.appendRow(rowData);
      Logger.log("deleteAppointmentFromSheet: Appointment copied to cancel sheet");
    } else {
      Logger.log("deleteAppointmentFromSheet: Cancel sheet not found, skipping data archiving");
    }
    
    // Delete the row from the appointments sheet
    appointmentsSheet.deleteRow(rowIndex);
    Logger.log("deleteAppointmentFromSheet: Appointment deleted from main sheet");
    
    // Log the deletion
    Logger.log(`deleteAppointmentFromSheet: Appointment ${appointmentId} deleted by ${session.username}`);
    
    return { 
      success: true, 
      message: "Appointment deleted successfully",
      deletedBy: session.username,
      deleteDate: new Date()
    };
  } catch (error) {
    Logger.log(`deleteAppointmentFromSheet: Error: ${error.toString()}`);
    return { success: false, message: `Error: ${error.toString()}` };
  }
}

function getBlockedSlotsCount(doctor, dateString) {
  try {
    let blockedSlots = 0;
    
    // Get disallowed intervals for general appointments
    const disallowed = getDisallowedIntervals();
    disallowed.forEach(function(interval) {
      if (interval.day === "D" || interval.day === getDayCode(new Date(dateString))) {
        const startMins = timeStringToMinutes(interval.startTime);
        const endMins = timeStringToMinutes(interval.endTime);
        
        if (!isNaN(startMins) && !isNaN(endMins) && endMins > startMins) {
          // Get booked slots for this time range
          const bookedSlots = getBookedSlots(doctor, dateString);
          const bookedSlotsInRange = bookedSlots.filter(slot => {
            const slotMins = timeStringToMinutes(slot);
            return slotMins >= startMins && slotMins < endMins;
          });
          
          // Calculate total possible slots in this range
          const totalSlotsInRange = Math.floor((endMins - startMins) / 10);
          // Add only the unfilled slots to blocked count
          blockedSlots += (totalSlotsInRange - bookedSlotsInRange.length);
        }
      }
    });

    // Get specialized intervals for Urgent New and Emergency
    const urgentNewIntervals = getSpecializedIntervals("URGENT NEW");
    const emergencyIntervals = getSpecializedIntervals("EMERGENCY");
    
    // Calculate blocked slots for Urgent New
    urgentNewIntervals.forEach(function(interval) {
      if (interval.day === "D" || interval.day === getDayCode(new Date(dateString))) {
        const startMins = timeStringToMinutes(interval.startTime);
        const endMins = timeStringToMinutes(interval.endTime);
        
        if (!isNaN(startMins) && !isNaN(endMins) && endMins > startMins) {
          // Get booked slots for this time range
          const bookedSlots = getBookedSlots(doctor, dateString);
          const bookedSlotsInRange = bookedSlots.filter(slot => {
            const slotMins = timeStringToMinutes(slot);
            return slotMins >= startMins && slotMins < endMins;
          });
          
          // Calculate total possible slots in this range
          const totalSlotsInRange = Math.floor((endMins - startMins) / 10);
          // Add only the unfilled slots to blocked count
          blockedSlots += (totalSlotsInRange - bookedSlotsInRange.length);
        }
      }
    });

    // Calculate blocked slots for Emergency
    emergencyIntervals.forEach(function(interval) {
      if (interval.day === "D" || interval.day === getDayCode(new Date(dateString))) {
        const startMins = timeStringToMinutes(interval.startTime);
        const endMins = timeStringToMinutes(interval.endTime);
        
        if (!isNaN(startMins) && !isNaN(endMins) && endMins > startMins) {
          // Get booked slots for this time range
          const bookedSlots = getBookedSlots(doctor, dateString);
          const bookedSlotsInRange = bookedSlots.filter(slot => {
            const slotMins = timeStringToMinutes(slot);
            return slotMins >= startMins && slotMins < endMins;
          });
          
          // Calculate total possible slots in this range
          const totalSlotsInRange = Math.floor((endMins - startMins) / 10);
          // Add only the unfilled slots to blocked count
          blockedSlots += (totalSlotsInRange - bookedSlotsInRange.length);
        }
      }
    });
    
    return blockedSlots;
  } catch (error) {
    Logger.log("Error in getBlockedSlotsCount: " + error.toString());
    return 0;
  }
}

function sendEmailOverview() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Get all appointments for today
    const data = appointmentsSheet.getDataRange().getValues();
    const todayAppointments = [];
    const planOfActionCounts = {};
    
    for (let i = 1; i < data.length; i++) {
      const appointmentDate = new Date(data[i][1]);
      appointmentDate.setHours(0, 0, 0, 0);
      
      if (appointmentDate.getTime() === today.getTime()) {
        todayAppointments.push(data[i]);
        const planOfAction = data[i][3] || "No Plan"; // Column D is Plan of Action
        planOfActionCounts[planOfAction] = (planOfActionCounts[planOfAction] || 0) + 1;
      }
    }
    
    // Get available slots count
    const availableSlots = getAvailableSlotsCount();
    
    // Get blocked slots count for each doctor
    const doctors = getDoctors();
    let totalBlockedSlots = 0;
    
    doctors.forEach(function(doctor) {
      const blockedSlots = getBlockedSlotsCount(doctor.name, Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd"));
      totalBlockedSlots += blockedSlots;
    });
    
    // Create email body
    let emailBody = "Appointment Overview for " + Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy") + "\n\n";
    
    // Add Plan of Action breakdown
    emailBody += "Plan of Action Breakdown:\n";
    for (const [plan, count] of Object.entries(planOfActionCounts)) {
      emailBody += `${plan}: ${count}\n`;
    }
    
    // Add slot information
    emailBody += "\nSlot Information:\n";
    emailBody += `Total Available Slots: ${availableSlots.count}\n`;
    emailBody += `Total Blocked Slots: ${totalBlockedSlots}\n`;
    
    // Add appointment details
    emailBody += "\nAppointment Details:\n";
    todayAppointments.forEach(function(appointment) {
      emailBody += `Time: ${appointment[2]}, Patient: ${appointment[4]}, Plan: ${appointment[3] || "No Plan"}\n`;
    });
    
    // Send email
    const emailAddress = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: emailAddress,
      subject: "Appointment Overview - " + Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      body: emailBody
    });
    
    return { success: true };
  } catch (error) {
    Logger.log("Error in sendEmailOverview: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

function getWaitingListDistribution() {
  try {
    Logger.log("Starting getWaitingListDistribution function");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    const data = appointmentsSheet.getDataRange().getValues();
    Logger.log("Total rows in appointments sheet: " + data.length);
    
    // Log header row to verify column indices
    Logger.log("Header row: " + JSON.stringify(data[0]));
    
    // Initialize counters
    const distribution = {
      red: 0,
      yellow: 0,
      green: 0
    };
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log("Row " + (i + 1) + " data: " + JSON.stringify(row));
      
      // Check if it's a waiting list appointment
      const appointmentType = row[8]?.toString().trim().toUpperCase() || ''; // Column I (index 8) is AppointmentType
      Logger.log("Row " + (i + 1) + " appointment type: " + appointmentType);
      
      if (appointmentType === 'WAITING LIST') {
        Logger.log("Found waiting list appointment at row " + (i + 1));
        const urgencyLevel = row[27]?.toString().trim().toLowerCase() || ''; // Last column (index 27) is Urgency Level
        Logger.log("Row " + (i + 1) + " urgency level: " + urgencyLevel);
        
        switch(urgencyLevel) {
          case 'red':
            distribution.red++;
            Logger.log("Incrementing red count. New count: " + distribution.red);
            break;
          case 'yellow':
            distribution.yellow++;
            Logger.log("Incrementing yellow count. New count: " + distribution.yellow);
            break;
          case 'green':
            distribution.green++;
            Logger.log("Incrementing green count. New count: " + distribution.green);
            break;
          default:
            Logger.log("Unknown urgency level: " + urgencyLevel);
        }
      }
    }
    
    Logger.log("Final distribution: " + JSON.stringify(distribution));
    return distribution;
  } catch (error) {
    Logger.log("Error in getWaitingListDistribution: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    return { red: 0, yellow: 0, green: 0 };
  }
}
/**
 * 1) Copy appointments for the given date into slotOptimization sheet
 * 2) Call the heap-based scheduler to fill results sheets
 */
function runSlotOptimizationForDate(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName('appointment');
  var dest = ss.getSheetByName('slotOptimization');
  if (!src || !dest) throw new Error('appointment or slotOptimization sheet missing');
  // Clear and set headers
  dest.clearContents();
  dest.appendRow([
    'Date','TimeSlot','MRD No','Patient Name',
    'Age','Gender','Mobile','Procedures','Fixed','Family Group'
  ]);

  var selected = new Date(dateStr);
  selected.setHours(0,0,0,0);
  var all = src.getDataRange().getValues();
  // Copy matching rows
  for (var i = 1; i < all.length; i++) {
    var r = all[i], d = new Date(r[1]);
    d.setHours(0,0,0,0);
    if (d.getTime() === selected.getTime()) {
      dest.appendRow([
         dateStr,  
         r[2], r[3], r[4],
        r[5], r[6], r[7], r[10], r[25], r[28]
      ]);
    }
  }
  // Run scheduler (populates Optimized Schedule, Availability Blocks, etc.)
  return generateOptimizedSchedules();
}

/**
 * Return all values from the named sheet as a 2D array of strings.
 */
function getSheetData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  return sheet.getDataRange().getDisplayValues();
}

/**
 * Calculate total consultation time (in minutes) for a given date
 */
function calculateTotalConsultationTime(dateString) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    const planOfActionSheet = ss.getSheetByName("PlanOfActionMaster");
    
    if (!appointmentsSheet || !planOfActionSheet) {
      throw new Error("Required sheets not found");
    }
    
    // Load PlanOfActionMaster data
    const planData = planOfActionSheet.getDataRange().getDisplayValues();
    const planHeaders = planData[0];
    const procedureData = new Map();
    
    // Create map of procedure to duration
    const procCol = planHeaders.indexOf('Plan of action');
    const consultCol = planHeaders.indexOf('ConsultMin');
    
    for (let i = 1; i < planData.length; i++) {
      const row = planData[i];
      const procedure = row[procCol];
      if (!procedure) continue;
      procedureData.set(procedure, parseInt(row[consultCol], 10) || 10); // Default 10 mins if not specified
    }
    
    // Get appointments for the date
    const appointments = appointmentsSheet.getDataRange().getDisplayValues();
    let totalConsultTime = 0;
    
    for (let i = 1; i < appointments.length; i++) {
      const row = appointments[i];
      const apptDate = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      if (apptDate === dateString) {
        let procedures = [];
        
        // Parse Plan of Action from column K (index 10)
        if (row[10]) {
          try {
            // Try to parse as JSON array first
            const planOfActionStr = row[10].toString().trim();
            if (planOfActionStr.startsWith('[') && planOfActionStr.endsWith(']')) {
              procedures = JSON.parse(planOfActionStr);
            } else {
              // Fallback to comma-separated string
              procedures = planOfActionStr.split(',').map(p => p.trim());
            }
          } catch (parseError) {
            Logger.log("JSON parse error for Plan of Action: " + parseError.toString());
            // Fallback to string splitting
            procedures = row[10].toString().split(',').map(p => p.trim());
          }
        }
        
        let maxDuration = 0;
        
        // Process each procedure
        procedures.forEach(proc => {
          if (!proc) return;
          
          // Clean up the procedure string by removing side labels like (RE), (LE), (BE)
          const cleanProc = proc
            .toString()
            .replace(/\s*\(.*?\)\s*/g, '') // Remove anything in parentheses
            .replace(/^["']|["']$/g, '')   // Remove leading/trailing quotes
            .trim();
          
          if (cleanProc) {
            const duration = procedureData.get(cleanProc) || 10; // Default 10 mins if not found
            maxDuration = Math.max(maxDuration, duration);
          }
        });
        
        totalConsultTime += maxDuration;
      }
    }
    
    return totalConsultTime;
  } catch (e) {
    Logger.log("Error in calculateTotalConsultationTime: " + e.toString());
    return 0;
  }
}

/**
 * Calculate total available minutes for a doctor on a given date
 */
function calculateTotalAvailableMinutes(doctor, dateString) {
  try {
    // Get default availability for the day
    const dayOfWeek = new Date(dateString).toLocaleDateString('en-US', { weekday: 'long' });
    const availabilityRanges = getDoctorAvailabilityRanges(doctor, dayOfWeek);
    
    let totalMinutes = 0;
    availabilityRanges.forEach(range => {
      const [startTime, endTime] = range.split('-');
      if (startTime && endTime) {
        const startMins = timeStringToMinutes(startTime.trim());
        const endMins = timeStringToMinutes(endTime.trim());
        if (endMins > startMins) {
          totalMinutes += (endMins - startMins);
        }
      }
    });
    
    // Subtract exception ranges
    const exceptions = getDoctorExceptionRanges(doctor, dateString);
    exceptions.forEach(range => {
      const [startTime, endTime] = range.split('-');
      if (startTime && endTime) {
        const startMins = timeStringToMinutes(startTime.trim());
        const endMins = timeStringToMinutes(endTime.trim());
        if (endMins > startMins) {
          totalMinutes -= (endMins - startMins);
        }
      }
    });
    
    // Subtract lunch break (30 minutes)
    totalMinutes -= 30;
    
    return Math.max(0, totalMinutes);
  } catch (e) {
    Logger.log("Error in calculateTotalAvailableMinutes: " + e.toString());
    return 0;
  }
}

/**
 * Calculate load for a specific day
 */
function calculateDayLoad(dateString) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    const doctorsSheet = ss.getSheetByName("DoctorAvailability");
    
    if (!appointmentsSheet || !doctorsSheet) {
      throw new Error("Required sheets not found");
    }
    
    // Get all doctors
    const doctorsData = doctorsSheet.getDataRange().getDisplayValues();
    const doctors = doctorsData.slice(1).map(row => row[0]); // First column contains doctor names
    
    // Calculate total consultation time once for all appointments (not per doctor)
    const totalConsultTime = calculateTotalConsultationTime(dateString);
    
    let totalAvailableMinutes = 0;
    
    // Calculate total available minutes for all doctors
    doctors.forEach(doctor => {
      totalAvailableMinutes += calculateTotalAvailableMinutes(doctor, dateString);
    });
    
    return {
      totalConsultTime: totalConsultTime,
      totalAvailableMinutes: totalAvailableMinutes
    };
  } catch (e) {
    Logger.log("Error in calculateDayLoad: " + e.toString());
    return {
      totalConsultTime: 0,
      totalAvailableMinutes: 0
    };
  }
}

/**
 * Move one appointment from "appointment" → "Archive" sheet.
 */
function archiveAppointment(appointmentId, sessionToken) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const src  = ss.getSheetByName('appointment');
  const dest = ss.getSheetByName('Archive');
  if (!src || !dest) {
    return { success: false, message: 'Appointment or Archive sheet not found.' };
  }

  // 1. Read all rows and find the matching appointment
  const data = src.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === appointmentId) {
      rowIdx = i;
      break;
    }
  }
  if (rowIdx === -1) {
    return { success: false, message: 'Appointment not found.' };
  }

  // 2. Prepare metadata
  const archiveDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm'
  );
  const archivedBy = getSessionUsername(sessionToken);

  // 3. Append the original row to Archive
  const originalRow = data[rowIdx];
  dest.appendRow(originalRow);

  // 4. Write metadata into columns X (24) and Y (25)
  const newRowIndex = dest.getLastRow();
  dest.getRange(newRowIndex, 24).setValue(archiveDate);
  dest.getRange(newRowIndex, 25).setValue(archivedBy);

  // 5. Remove the row from the source sheet
  src.deleteRow(rowIdx + 1);

  return { success: true, message: 'Appointment archived successfully.' };
}


/**
 * Returns, for the given month (YYYY-MM), a map of
 *   doctor → { dateString → [ "HH:mm-HH:mm", … ] }
 * representing each doctor's exception blocks.
 */
function getDoctorMonthlyExceptionMap(monthKey, filterDoctor) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DoctorAvailabilityExceptions");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const map  = {};

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const [doc, rawDate, fromRaw, toRaw] = data[i];
    const doctor = doc.toString().trim();
    
    // Filter by doctor if provided
    if (filterDoctor && doctor !== filterDoctor) {
      continue;
    }
    
    const date   = Utilities.formatDate(new Date(rawDate),
                       Session.getScriptTimeZone(),
                       "yyyy-MM-dd");
    // only include this month
    if (!date.startsWith(monthKey)) continue;

    const from = formatTimeString(fromRaw);
    const to   = formatTimeString(toRaw);
    if (!from || !to) continue;

    map[doctor]           = map[doctor]           || {};
    map[doctor][date]     = map[doctor][date]     || [];
    map[doctor][date].push(`${from}-${to}`);
  }
  return map;
}

/**
 * Returns all appointments for MRD numbers that appear more than once.
 */
function getDuplicateAppointments() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('appointment');
    if (!sheet) return { success: false, message: 'Appointment sheet not found' };
    
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const mrdCol  = headers.indexOf('MRDNo');
    if (mrdCol < 0) return { success: false, message: 'MRDNo column not found' };
    
    // 1. Count occurrences
    const counts = {};
    for (let i = 1; i < data.length; i++) {
      const mrd = data[i][mrdCol];
      if (mrd) counts[mrd] = (counts[mrd] || 0) + 1;
    }
    const duplicates = Object.keys(counts).filter(m => counts[m] > 1);
    if (duplicates.length === 0) {
      return { success: false, message: 'No MRD with multiple appointments' };
    }
    
    // Helper function to safely format dates
    function safeFormatDate(dateValue, format) {
      try {
        if (!dateValue) return 'N/A';
        let date;
        if (dateValue instanceof Date) {
          date = dateValue;
        } else {
          date = new Date(dateValue);
        }
        if (isNaN(date.getTime())) return 'Invalid Date';
        return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
      } catch (e) {
        return 'Invalid Date';
      }
    }
    
    // Helper function to safely get column value
    function safeGetValue(row, columnIndex) {
      try {
        if (columnIndex < 0 || columnIndex >= row.length) return '';
        const value = row[columnIndex];
        return value !== null && value !== undefined ? value.toString() : '';
      } catch (e) {
        return '';
      }
    }
    
    // 2. Build appointment objects for duplicates
    const col = name => headers.indexOf(name);
    const appts = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (duplicates.includes(row[mrdCol])) {
        try {
          appts.push({
            AppointmentID:   safeGetValue(row, col('AppointmentID')),
            AppointmentDate: safeFormatDate(row[col('AppointmentDate')], 'yyyy-MM-dd'),
            TimeSlot:        safeFormatDate(row[col('TimeSlot')], 'hh:mm a'),
            MRDNo:           safeGetValue(row, mrdCol),
            PatientName:     safeGetValue(row, col('PatientName')),
            Gender:          safeGetValue(row, col('Gender')),
            Age:             safeGetValue(row, col('Age')),
            Phone:           safeGetValue(row, col('Phone')),
            AppointmentType: safeGetValue(row, col('AppointmentType')),
            Doctor:          safeGetValue(row, col('Doctor')),
            FixedSlot:       safeGetValue(row, col('FixedSlot')),
            PlanOfAction:    safeGetValue(row, col('PlanOfAction')),
            Remarks:         safeGetValue(row, col('Remarks'))
          });
        } catch (e) {
          // Log the error for this specific row but continue processing
    
          continue;
        }
      }
    }
    
    // 3. Sort by MRD, then by appointment date and time
    appts.sort((a, b) => {
      if (a.MRDNo !== b.MRDNo) {
        return a.MRDNo.localeCompare(b.MRDNo);
      }
      // For same MRD, sort by appointment date, then time
      if (a.AppointmentDate !== b.AppointmentDate) {
        return a.AppointmentDate.localeCompare(b.AppointmentDate);
      }
      return a.TimeSlot.localeCompare(b.TimeSlot);
    });
    
    return { success: true, appointments: appts, total: appts.length };
    
  } catch (e) {

    return { success: false, message: 'Error retrieving duplicate appointments: ' + e.toString() };
  }
}

/**
 * Updates the appointment sheet based on final selections
 * @param {string} date - The optimization date
 * @param {Array} finalSelections - Array of {mrd, time} objects
 * @return {Object} Status object
 */
function updateFinalAppointments(date, finalSelections) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentSheet = ss.getSheetByName("appointment");
    const data = appointmentSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Get column indices
    const mrdCol = headers.indexOf("MRDNo");
    const dateCol = headers.indexOf("AppointmentDate");
    const timeCol = headers.indexOf("TimeSlot");
    const typeCol = headers.indexOf("AppointmentType");
    
    if (mrdCol === -1 || dateCol === -1 || timeCol === -1 || typeCol === -1) {
      throw new Error("Required columns not found in appointment sheet");
    }
    
    // Format the date for comparison
    const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Process each row in the appointment sheet
    for (let i = 1; i < data.length; i++) {
      const rowDate = Utilities.formatDate(new Date(data[i][dateCol]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      // Only process rows for the selected date
      if (rowDate === formattedDate) {
        const mrd = data[i][mrdCol].toString();
        const isSelected = finalSelections.some(sel => sel.mrd === mrd);
        
        if (isSelected) {
          // Update time slot for selected appointments
          const selection = finalSelections.find(sel => sel.mrd === mrd);
          appointmentSheet.getRange(i + 1, timeCol + 1).setValue(selection.time);
          appointmentSheet.getRange(i + 1, typeCol + 1).setValue("GENERAL");
        } else {
          // Set to WAITING LIST for non-selected appointments
          appointmentSheet.getRange(i + 1, typeCol + 1).setValue("WAITING LIST");
          appointmentSheet.getRange(i + 1, timeCol + 1).setValue("");
        }
      }
    }

       // Clear the optimization sheets
    clearOptimizationSheets();
    
    return { success: true, message: "Appointments updated successfully" };
  } catch (e) {
    Logger.log("Error in updateFinalAppointments: " + e.toString());
    return { success: false, message: e.toString() };
  }
}


function clearOptimizationSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToClear = [
      "slotOptimization",
      "Optimized Schedule",
      "Availability Blocks",
      "Utilization Report",
      "Reschedule List"
    ];
    
    sheetsToClear.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        sheet.clearContents();
        // Add headers back to slotOptimization sheet if it exists
        if (sheetName === "slotOptimization") {
          sheet.appendRow([
            'Date','TimeSlot','MRD No','Patient Name',
            'Age','Gender','Mobile','Procedures'
          ]);
        }
      }
    });
    
    return true;
  } catch (e) {
    Logger.log("Error in clearOptimizationSheets: " + e.toString());
    return false;
  }
}

/**
 * Generate a unique family identifier for appointments
 * @param {string} appointmentId - The appointment ID to update
 * @param {boolean} isFirstMember - Whether this is the first member of the family
 * @returns {Object} Success/failure response with the generated identifier
 */
function generateFamilyIdentifier(appointmentId, isFirstMember) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("appointment");
    
    if (!sheet) {
      return { success: false, message: "Appointment sheet not found." };
    }

    // Get headers and find Family Identifier column (AC)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const familyIdentifierColIndex = headers.indexOf("Family Identifier");
    
    if (familyIdentifierColIndex === -1) {
      return { success: false, message: "Family Identifier column not found. Please ensure column AC exists." };
    }

    // Check if this is a temporary booking ID
    const isTemporaryBooking = appointmentId.startsWith("TEMP_BOOKING_");
    
    let familyIdentifier = "";

    if (isFirstMember) {
      // Generate a new family identifier
      const existingIdentifiers = [];
      
      // Collect all existing family identifiers from column AC
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const existingValue = data[i][familyIdentifierColIndex];
        if (existingValue && existingValue.toString().trim().startsWith("F")) {
          const numericPart = existingValue.toString().trim().substring(1);
          const number = parseInt(numericPart, 10);
          if (!isNaN(number)) {
            existingIdentifiers.push(number);
          }
        }
      }

      // Find the smallest unused number from 1 to 1000
      let newNumber = 1;
      while (existingIdentifiers.includes(newNumber) && newNumber <= 1000) {
        newNumber++;
      }

      if (newNumber > 1000) {
        return { success: false, message: "Maximum family identifier limit reached (1000)." };
      }

      familyIdentifier = "F" + newNumber.toString().padStart(3, "0");
      
    } else {
      // This is not the first member, so we need to find the first member's identifier
      // For now, we'll generate a new identifier and let the user manually link them
      // In a more complex implementation, you might want to add a "Family Group" field
      // that links to the first member's appointment ID
      
      // For simplicity, we'll generate a new identifier and suggest manual linking
      const existingIdentifiers = [];
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const existingValue = data[i][familyIdentifierColIndex];
        if (existingValue && existingValue.toString().trim().startsWith("F")) {
          const numericPart = existingValue.toString().trim().substring(1);
          const number = parseInt(numericPart, 10);
          if (!isNaN(number)) {
            existingIdentifiers.push(number);
          }
        }
      }

      let newNumber = 1;
      while (existingIdentifiers.includes(newNumber) && newNumber <= 1000) {
        newNumber++;
      }

      if (newNumber > 1000) {
        return { success: false, message: "Maximum family identifier limit reached (1000)." };
      }

      familyIdentifier = "F" + newNumber.toString().padStart(3, "0");
    }

    // Only write to the sheet if this is not a temporary booking
    if (!isTemporaryBooking) {
      // Find the appointment row
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === appointmentId) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex === -1) {
        return { success: false, message: "Appointment not found." };
      }

      // Write the family identifier to the appointment row
      sheet.getRange(rowIndex + 1, familyIdentifierColIndex + 1).setValue(familyIdentifier);
    }

    return { 
      success: true, 
      message: `Family identifier ${familyIdentifier} assigned successfully.`,
      familyIdentifier: familyIdentifier
    };

  } catch (error) {
    Logger.log("Error in generateFamilyIdentifier: " + error.toString());
    return { success: false, message: "Error generating family identifier: " + error.toString() };
  }
}

/**
 * Assign a manually entered family identifier to an appointment
 * @param {string} appointmentId - The appointment ID to update
 * @param {string} manualIdentifier - The manually entered family identifier (e.g. F001)
 * @returns {Object} Success/failure response
 */
function assignManualFamilyIdentifier(appointmentId, manualIdentifier, sessionToken) {
  try {
    // Validate session
    const session = getUserSession(sessionToken);
    if (!session) {
      return { success: false, message: "Invalid session" };
    }

    // Get the appointments sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    
    if (!appointmentsSheet) {
      return { success: false, message: "Appointments sheet not found" };
    }

    // Find the appointment by ID
    const data = appointmentsSheet.getDataRange().getValues();
    let appointmentRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === appointmentId) { // Column A contains AppointmentID
        appointmentRow = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (appointmentRow === -1) {
      return { success: false, message: "Appointment not found" };
    }

    // Update column AC (index 28) with the manual family identifier
    appointmentsSheet.getRange(appointmentRow, 29).setValue(manualIdentifier); // Column AC is index 28, but getRange is 1-indexed so 29
    
    return { 
      success: true, 
      message: `Family identifier ${manualIdentifier} assigned successfully`,
      familyIdentifier: manualIdentifier
    };
    
  } catch (error) {
    console.error("Error in assignManualFamilyIdentifier:", error);
    return { success: false, message: "Error assigning family identifier: " + error.toString() };
  }
}

function getExistingFamilyIdentifiers() {
  try {
    // Get the appointments sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    
    if (!appointmentsSheet) {
      return { success: false, message: "Appointments sheet not found" };
    }

    // Get all values from column AC (index 28)
    const lastRow = appointmentsSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, familyIdentifiers: [] };
    }
    
    const familyIdentifierRange = appointmentsSheet.getRange(2, 29, lastRow - 1, 1); // Column AC is index 29 (1-indexed)
    const familyIdentifierValues = familyIdentifierRange.getValues();
    
    // Extract unique, non-empty family identifiers
    const uniqueFamilyIdentifiers = new Set();
    
    familyIdentifierValues.forEach(row => {
      const value = row[0];
      if (value && typeof value === 'string' && value.trim() !== '') {
        uniqueFamilyIdentifiers.add(value.trim());
      }
    });
    
    // Convert to sorted array
    const sortedFamilyIdentifiers = Array.from(uniqueFamilyIdentifiers).sort();
    
    return { 
      success: true, 
      familyIdentifiers: sortedFamilyIdentifiers
    };
    
  } catch (error) {
    console.error("Error in getExistingFamilyIdentifiers:", error);
    return { success: false, message: "Error fetching family identifiers: " + error.toString() };
  }
}

/**
 * Get email-formatted day overview for printing
 * @param {string} date - The date to get overview for
 * @param {string} sessionToken - The session token
 * @return {Object} Object containing success status and HTML content
 */
function getEmailFormattedDayOverview(date, sessionToken) {
  try {
    // Validate session
    const session = getUserSession(sessionToken);
    if (!session) {
      return { success: false, message: "Invalid session" };
    }

    // Get appointments for the day
    const appointments = getAppointmentsForDay(date);
    if (!appointments || appointments.length === 0) {
      return { success: false, message: "No appointments found for the selected date" };
    }

    // Sort appointments by time
    appointments.sort((a, b) => {
      const timeA = a.TimeSlot ? new Date(a.TimeSlot).getTime() : 0;
      const timeB = b.TimeSlot ? new Date(b.TimeSlot).getTime() : 0;
      return timeA - timeB;
    });

    // Helper function to format plan of action for human reading
    function formatPlanOfActionForPrint(planOfAction) {
      if (!planOfAction || planOfAction === '-' || planOfAction.trim() === '') {
        return '-';
      }
      
      try {
        // Try to parse as JSON first
        const plans = JSON.parse(planOfAction);
        const planArray = Array.isArray(plans) ? plans : [plans];
        
        return planArray.map(plan => {
          if (typeof plan === 'object' && plan.name) {
            // If it's an object with name and eye properties
            let formatted = plan.name;
            if (plan.eye) {
              formatted += ` (${plan.eye})`;
            }
            return formatted;
          } else {
            // If it's just a string
            return plan.toString();
          }
        }).join(', ');
        
      } catch (e) {
        // If not JSON, split by common separators and clean up
        const plans = planOfAction.split(/[/,;]/);
        return plans.map(plan => plan.trim()).filter(plan => plan !== '').join(', ');
      }
    }

    // Create the same content as email but return HTML instead of sending
    const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");

    // Count appointments by type and plan of action
    const summary = {
      total: appointments.length,
      confirmed: 0,
      waitingList: 0,
      newPatients: 0,
      oldPatients: 0,
      planOfAction: {}
    };

    appointments.forEach(appt => {
      if (appt.AppointmentType === 'WAITING LIST') {
        summary.waitingList++;
      }
      if (appt.FourDayConfirmDate || appt.OneDayConfirmDate) {
        summary.confirmed++;
      }
      
      // Count new vs old patients based on MRD number
      const mrdNo = (appt.MRDNo || '').toString().trim().toUpperCase();
      if (mrdNo.startsWith('N')) {
        summary.newPatients++;
      } else if (mrdNo.startsWith('S') || (mrdNo && !mrdNo.startsWith('N'))) {
        summary.oldPatients++;
      }
      
      // Count plan of action occurrences (exclude "New" entries)
      if (appt.PlanOfAction && appt.PlanOfAction !== '-') {
        try {
          const plans = JSON.parse(appt.PlanOfAction);
          const planArray = Array.isArray(plans) ? plans : [plans];
          planArray.forEach(plan => {
            const planName = typeof plan === 'object' ? plan.name : plan.toString();
            // Skip if plan name contains "New" (case insensitive)
            if (!planName.toLowerCase().includes('new')) {
              summary.planOfAction[planName] = (summary.planOfAction[planName] || 0) + 1;
            }
          });
        } catch (e) {
          const plans = appt.PlanOfAction.split(/[/,;]/);
          plans.forEach(plan => {
            const cleanPlan = plan.trim();
            // Skip if plan name contains "New" (case insensitive)
            if (cleanPlan && !cleanPlan.toLowerCase().includes('new')) {
              summary.planOfAction[cleanPlan] = (summary.planOfAction[cleanPlan] || 0) + 1;
            }
          });
        }
      }
    });

    // Generate HTML content (same as email)
    let htmlContent = `
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; color: #333; line-height: 1.6; margin: 0; padding: 0; }
          .container { max-width: 800px; margin: 0 auto; padding: 20px; }
          .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #ddd; padding-bottom: 20px; }
          .header h1 { color: #2c3e50; margin: 0; }
          .header p { margin: 5px 0; color: #7f8c8d; }
          .summary { display: flex; justify-content: space-around; background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 30px; }
          .summary-item { text-align: center; }
          .summary-item h3 { margin: 0; font-size: 28px; color: #3498db; }
          .summary-item p { margin: 5px 0 0 0; font-weight: bold; color: #2c3e50; }
          .plan-summary { margin-bottom: 30px; }
          .plan-summary h2 { color: #2c3e50; border-bottom: 1px solid #ddd; padding-bottom: 10px; }
          .plan-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-top: 15px; }
          .plan-item { background-color: #f8f9fa; padding: 10px; border-radius: 5px; display: flex; justify-content: space-between; align-items: center; }
          .appointments-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          .appointments-table th, .appointments-table td { border: 1px solid #ddd; padding: 12px 8px; text-align: left; }
          .appointments-table th { background-color: #f8f9fa; font-weight: bold; color: #2c3e50; }
          .appointments-table th:first-child, .appointments-table td:first-child { width: 60px; text-align: center; font-weight: bold; }
          .appointments-table tr:nth-child(even) { background-color: #f9f9f9; }
          .appointments-table tr.waiting-list { background-color: #fff3cd; }
          .footer { margin-top: 30px; text-align: center; color: #7f8c8d; font-size: 12px; border-top: 1px solid #ddd; padding-top: 20px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>Day Overview Report</h1>
            <p>${formattedDate}</p>
          </div>

          <div class="summary">
            <div class="summary-item">
              <h3>${summary.total}</h3>
              <p>Total Appointments</p>
            </div>
            <div class="summary-item">
              <h3>${summary.confirmed}</h3>
              <p>Confirmed</p>
            </div>
            <div class="summary-item">
              <h3>${summary.waitingList}</h3>
              <p>Waiting List</p>
            </div>
            <div class="summary-item">
              <h3>${summary.newPatients}</h3>
              <p>New</p>
            </div>
            <div class="summary-item">
              <h3>${summary.oldPatients}</h3>
              <p>Old</p>
            </div>
          </div>

          ${Object.keys(summary.planOfAction).length > 0 ? `
          <div class="plan-summary">
            <h2>Plan of Action Summary</h2>
            <div class="plan-grid">
              ${Object.entries(summary.planOfAction)
                .sort((a, b) => b[1] - a[1])
                .map(([plan, count]) => `
                  <div class="plan-item">
                    <span>${plan}</span>
                    <strong>${count}</strong>
                  </div>
                `).join('')}
            </div>
          </div>
          ` : ''}

          <h2>Appointments Detail</h2>
          <table class="appointments-table">
            <thead>
              <tr>
                <th>S.No</th>
                <th>Time</th>
                <th>MRD No</th>
                <th>Patient Name</th>
                <th>Gender</th>
                <th>Age</th>
                <th>Phone</th>
                <th>Member Type</th>
                <th>Plan of Action</th>
                <th>Remarks</th>
              </tr>
            </thead>
            <tbody>
              ${appointments.map((appt, index) => `
                <tr class="${appt.AppointmentType === 'WAITING LIST' ? 'waiting-list' : ''}">
                  <td>${index + 1}</td>
                  <td>${Utilities.formatDate(new Date(appt.TimeSlot), Session.getScriptTimeZone(), "hh:mm a")}</td>
                  <td>${appt.MRDNo}</td>
                  <td>${appt.PatientName}</td>
                  <td>${appt.Gender}</td>
                  <td>${appt.Age}</td>
                  <td>${appt.Phone}</td>
                  <td>${appt.PatientType || '-'}</td>
                  <td>${formatPlanOfActionForPrint(appt.PlanOfAction)}</td>
                  <td>${appt.Remarks || '-'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div class="footer">
            <p>This is an automated report generated from the Appointment Management System.</p>
            <p>Generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' hh:mm a")}</p>
          </div>
        </div>
      </body>
      </html>
    `;

    return {
      success: true,
      htmlContent: htmlContent
    };

  } catch (error) {
    Logger.log("Error in getEmailFormattedDayOverview: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Update family identifier for an existing appointment (simple version)
 * @param {string} appointmentId - The appointment ID to update
 * @param {string} familyIdentifier - The family identifier to set
 * @param {string} sessionToken - Session token for validation
 * @returns {Object} Success/failure response
 */
function updateFamilyIdentifier(appointmentId, familyIdentifier, sessionToken) {
  try {
    // Validate session
    const session = getUserSession(sessionToken);
    if (!session) {
      return { success: false, message: "Invalid session" };
    }

    // Get the appointments sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName("appointment");
    
    if (!appointmentsSheet) {
      return { success: false, message: "Appointments sheet not found" };
    }

    // Find the appointment by ID
    const data = appointmentsSheet.getDataRange().getValues();
    let appointmentRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === appointmentId) { // Column A contains AppointmentID
        appointmentRow = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (appointmentRow === -1) {
      return { success: false, message: "Appointment not found" };
    }

    // Update column AC (index 28) with the family identifier
    // Column AC is index 28, but getRange is 1-indexed so 29
    appointmentsSheet.getRange(appointmentRow, 29).setValue(familyIdentifier);
    
    return { 
      success: true, 
      message: `Family identifier ${familyIdentifier} updated successfully`,
      familyIdentifier: familyIdentifier
    };
    
  } catch (error) {
    console.error("Error in updateFamilyIdentifier:", error);
    return { success: false, message: "Error updating family identifier: " + error.toString() };
  }
}

/**
 * Get comprehensive day timeline slots for modal display (OPTIMIZED VERSION)
 * Returns all slots (available and blocked) using same logic as load indicator
 * Performance improvements: Single sheet read, batch processing, reduced loops
 * Now includes procedure-based consultation durations from PlanOfActionMaster
 */
function getDayTimelineSlots(doctor, dateString) {
  try {
    // Get day name for availability lookup
    var appointmentDate = new Date(dateString);
    var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    var dayOfWeek = dayNames[appointmentDate.getDay()];
    
    // OPTIMIZATION: Get all sheet data in one operation
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetsData = {};
    
    // Load all required sheet data at once
    try {
      var availabilitySheet = ss.getSheetByName("DoctorAvailability");
      sheetsData.availability = availabilitySheet ? availabilitySheet.getDataRange().getValues() : [];
    } catch (e) { sheetsData.availability = []; }
    
    try {
      var appointmentSheet = ss.getSheetByName("appointment");
      sheetsData.appointments = appointmentSheet ? appointmentSheet.getDataRange().getValues() : [];
    } catch (e) { sheetsData.appointments = []; }
    
    try {
      var exceptionsSheet = ss.getSheetByName("DoctorAvailabilityExceptions");
      sheetsData.exceptions = exceptionsSheet ? exceptionsSheet.getDataRange().getValues() : [];
    } catch (e) { sheetsData.exceptions = []; }
    
    // Load PlanOfActionMaster data for procedure durations
    try {
      var planOfActionSheet = ss.getSheetByName("PlanOfActionMaster");
      sheetsData.planOfAction = planOfActionSheet ? planOfActionSheet.getDataRange().getValues() : [];
    } catch (e) { sheetsData.planOfAction = []; }
    
    // Create procedure duration map from PlanOfActionMaster
    var procedureMap = {};
    if (sheetsData.planOfAction.length > 0) {
      var planHeaders = sheetsData.planOfAction[0];
      var procCol = planHeaders.indexOf('Plan of action');
      var consultCol = planHeaders.indexOf('ConsultMin');
      var prepCol = planHeaders.indexOf('PrepMin');
      var priorityCol = planHeaders.indexOf('Priority');
      
      Logger.log("Building procedure map from PlanOfActionMaster...");
      
      for (var p = 1; p < sheetsData.planOfAction.length; p++) {
        var row = sheetsData.planOfAction[p];
        var procedure = row[procCol];
        if (procedure) {
          var procName = procedure.toString().trim().toUpperCase();
          var duration = parseInt(row[consultCol], 10) || 10;
          procedureMap[procName] = {
            duration: duration,
            prep: parseInt(row[prepCol], 10) || 0,
            priority: parseInt(row[priorityCol], 10) || 999
          };
          Logger.log("Added procedure: " + procName + " -> " + duration + " minutes");
        }
      }
      Logger.log("Total procedures loaded: " + Object.keys(procedureMap).length);
    }
    
    // Helper function to clean procedure names
    function cleanProcedureName(procedure) {
      return procedure.replace(/\s*\([^)]*\)\s*/, '').trim().toUpperCase();
    }
    
    // Helper function to calculate consultation time for procedures
    function calculateConsultTime(procedures) {
      if (!procedures || procedures.length === 0) return 10; // default
      
      var maxDuration = 0;
      var foundProcedures = [];
      
      for (var i = 0; i < procedures.length; i++) {
        var originalProc = procedures[i].toString().trim();
        var cleanProc = cleanProcedureName(originalProc);
        
        // Since procedureMap keys are already uppercase, direct match should work
        var procData = procedureMap[cleanProc];
        
        if (procData) {
          maxDuration = Math.max(maxDuration, procData.duration);
          foundProcedures.push(cleanProc + ":" + procData.duration + "min");
        } else {
          foundProcedures.push(cleanProc + ":NOT_FOUND");
        }
      }
      
      // Add debug logging
      Logger.log("Procedure calculation for [" + procedures.join(", ") + "]:");
      Logger.log("  -> Cleaned: " + foundProcedures.join(", "));
      Logger.log("  -> Final Duration: " + (maxDuration || 10) + " minutes");
      
      return maxDuration || 10; // default if no procedures found
    }
    
    // OPTIMIZATION: Get doctor availability ranges from cached data
    var availabilityRanges = [];
    if (sheetsData.availability.length > 0) {
      var headers = sheetsData.availability[0];
      var dayIndex = headers.indexOf(dayOfWeek);
      if (dayIndex !== -1) {
        for (var i = 1; i < sheetsData.availability.length; i++) {
          if (sheetsData.availability[i][0].toString().trim() === doctor) {
            var cellValue = sheetsData.availability[i][dayIndex];
            if (cellValue) {
              availabilityRanges = cellValue.split(",").map(function(r) { return r.trim(); });
            }
            break;
          }
        }
      }
    }
    
    // Generate all possible base slots from availability ranges (5-minute intervals for precision)
    var allBaseSlots = [];
    var baseInterval = 5; // 5-minute base intervals for more precision
    
    for (var r = 0; r < availabilityRanges.length; r++) {
      var parts = availabilityRanges[r].split("-");
      if (parts.length < 2) continue;
      
      var startTime = parts[0].trim();
      var endTime = parts[1].trim();
      var startMins = timeStringToMinutes(startTime);
      var endMins = timeStringToMinutes(endTime);
      
      for (var t = startMins; t < endMins; t += baseInterval) {
        allBaseSlots.push(t); // Store as minutes for easier calculation
      }
    }
    
    // Constants for lunch break
    var lunchStart = timeStringToMinutes("13:00");
    var lunchEnd = timeStringToMinutes("13:30");
    
    // OPTIMIZATION: Process appointments data in single pass to get booked slots with their durations
    var bookedSlotRanges = []; // Store as {start: mins, end: mins, formatted: "HH:mm"}
    var fixedSlots = [];
    var appointmentHeaders = sheetsData.appointments.length > 0 ? sheetsData.appointments[0] : [];
    var proceduresColIndex = appointmentHeaders.indexOf('Procedures');
    
    for (var i = 1; i < sheetsData.appointments.length; i++) {
      var rowDate = Utilities.formatDate(new Date(sheetsData.appointments[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var rowDoctor = sheetsData.appointments[i][9] ? sheetsData.appointments[i][9].toString().trim() : "";
      
      if (rowDate === dateString && rowDoctor === doctor) {
        // Get booked time slot
        var timeSlotRaw = sheetsData.appointments[i][2];
        var formattedTime = Utilities.formatDate(new Date(timeSlotRaw), Session.getScriptTimeZone(), "HH:mm");
        var slotStartMins = timeStringToMinutes(formattedTime);
        
        // Get procedures for this appointment to calculate duration
        var procedures = [];
        if (proceduresColIndex >= 0 && sheetsData.appointments[i][proceduresColIndex]) {
          var procString = sheetsData.appointments[i][proceduresColIndex].toString().trim();
          Logger.log("Raw procedure string for " + formattedTime + ": " + procString);
          
          try {
            if (procString.startsWith('[') && procString.endsWith(']')) {
              // Handle JSON array format: ["REF","DIL","OCT (RE)"]
              procedures = JSON.parse(procString);
            } else if (procString.indexOf(',') > -1) {
              // Handle comma-separated format
              procedures = procString.split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p; });
            } else if (procString) {
              // Single procedure
              procedures = [procString];
            }
            Logger.log("Parsed procedures: " + JSON.stringify(procedures));
          } catch (e) {
            Logger.log("Error parsing procedures: " + e.toString());
            procedures = [procString]; // fallback to single procedure
          }
        }
        
        // Calculate actual consultation duration
        var consultDuration = calculateConsultTime(procedures);
        var slotEndMins = slotStartMins + consultDuration;
        
        bookedSlotRanges.push({
          start: slotStartMins,
          end: slotEndMins,
          formatted: formattedTime,
          duration: consultDuration,
          procedures: procedures
        });
        
        // Check if slot is fixed
        var isFixed = sheetsData.appointments[i][25]; // Column Z (0-based index 25)
        if (isFixed === "Yes") {
          fixedSlots.push(formattedTime);
        }
      }
    }
    
    // OPTIMIZATION: Get exception ranges from cached data
    var exceptions = [];
    for (var j = 1; j < sheetsData.exceptions.length; j++) {
      var rowDoctor = sheetsData.exceptions[j][0]?.toString().trim() || "";
      var rowDate = sheetsData.exceptions[j][1] ? 
        Utilities.formatDate(new Date(sheetsData.exceptions[j][1]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      
      if (rowDoctor === doctor && rowDate === dateString) {
        var fromTime = formatTimeString(sheetsData.exceptions[j][2]);
        var toTime = formatTimeString(sheetsData.exceptions[j][3]);
        
        if (fromTime && toTime) {
          exceptions.push(fromTime + "-" + toTime);
        }
      }
    }
    
    // Build available, blocked, and booked slot lists with proper durations
    var availableSlots = [];
    var blockedSlots = [];
    var bookedSlots = [];
    
    // Process base slots in 10-minute intervals for display
    var displayInterval = 10;
    for (var s = 0; s < allBaseSlots.length; s += (displayInterval / baseInterval)) {
      var slotMins = allBaseSlots[s];
      var slotTime = formatTime(slotMins);
      var isBlocked = false;
      var isBooked = false;
      
      // Check lunch break
      if (slotMins >= lunchStart && slotMins < lunchEnd) {
        blockedSlots.push(slotTime);
        continue;
      }
      
      // Check if this slot overlaps with any booked appointment
      for (var b = 0; b < bookedSlotRanges.length; b++) {
        var booking = bookedSlotRanges[b];
        if (slotMins >= booking.start && slotMins < booking.end) {
          if (bookedSlots.indexOf(booking.formatted) === -1) {
            bookedSlots.push(booking.formatted);
          }
          isBooked = true;
          break;
        }
      }
      
      if (isBooked) continue;
      
      // Check exception ranges
      for (var e = 0; e < exceptions.length; e++) {
        var exParts = exceptions[e].split("-");
        if (exParts.length < 2) continue;
        
        var exStart = timeStringToMinutes(exParts[0]);
        var exEnd = timeStringToMinutes(exParts[1]);
        
        if (slotMins >= exStart && slotMins < exEnd) {
          isBlocked = true;
          break;
        }
      }
      
      if (isBlocked) {
        blockedSlots.push(slotTime);
      } else {
        availableSlots.push(slotTime);
      }
    }
    
    return {
      success: true,
      availableSlots: availableSlots,
      bookedSlots: bookedSlots,
      fixedSlots: fixedSlots,
      blockedSlots: blockedSlots,
      bookedDetails: bookedSlotRanges // Additional detail for debugging
    };
    
  } catch (e) {
    Logger.log("Error in getDayTimelineSlots: " + e.toString());
    return { 
      success: false, 
      message: e.toString(),
      availableSlots: [],
      bookedSlots: [],
      fixedSlots: [],
      blockedSlots: []
    };
  }
}

/**
 * Move one appointment from "appointment" → "Archive" sheet.
 */

/**
 * Get comprehensive rescheduling analytics for dashboard
 * @return {Object} Rescheduling analytics data
 */
function getReschedulingAnalytics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentSheet = ss.getSheetByName("appointment");
    const cancelSheet = ss.getSheetByName("cancel");
    const archiveSheet = ss.getSheetByName("Archive");
    const appointmentArchiveSheet = ss.getSheetByName("appointment_archive");
    
    const analytics = {
      totalReschedules: 0,
      reschedulesByReason: {},
      frequentReschedules: [],
      reschedulesByStaff: {},
      reschedulesByMonth: {},
      highRiskAppointments: [],
      averageReschedulesPerAppointment: 0,
      rescheduleTrend: {}
    };
    
    // Process appointment sheet
    if (appointmentSheet) {
      const appointmentData = appointmentSheet.getDataRange().getValues();
      const headers = appointmentData[0];
      
      // Use column U directly (index 20, since arrays are 0-indexed)
      const rescheduleHistoryCol = 20; // Column U
      
      // Safety check to ensure column U exists
      if (headers.length <= rescheduleHistoryCol) {
        Logger.log("Warning: Column U does not exist in appointment sheet. Total columns: " + headers.length);
        return analytics;
      }
      
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const rescheduledDateCol = headers.indexOf("RescheduledDate");
      const appointmentIdCol = headers.indexOf("AppointmentID");
      const patientNameCol = headers.indexOf("PatientName");
      const mrdNoCol = headers.indexOf("MRDNo");
      
      for (let i = 1; i < appointmentData.length; i++) {
        const row = appointmentData[i];
        const rescheduleHistory = row[rescheduleHistoryCol];
        const rescheduledBy = row[rescheduledByCol];
        const rescheduledDate = row[rescheduledDateCol];
        const appointmentId = row[appointmentIdCol];
        const patientName = row[patientNameCol];
        const mrdNo = row[mrdNoCol];
        
        // Debug: Log first few rows to see what we're getting
        if (i <= 3) {
          Logger.log("Row " + i + " - AppointmentID: " + appointmentId + ", RescheduleHistory: " + (rescheduleHistory ? rescheduleHistory.substring(0, 100) + "..." : "null"));
        }
        
        if (rescheduleHistory) {
          try {
            const history = JSON.parse(rescheduleHistory);
            analytics.totalReschedules += history.length;
            
            // Track by staff member
            history.forEach(entry => {
              const staff = entry.by || 'Unknown';
              analytics.reschedulesByStaff[staff] = (analytics.reschedulesByStaff[staff] || 0) + 1;
              
              // Track by month
              const date = new Date(entry.timestamp);
              const monthKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
              analytics.reschedulesByMonth[monthKey] = (analytics.reschedulesByMonth[monthKey] || 0) + 1;
            });
            
            // Track frequent reschedules
            if (history.length >= 2) {
              analytics.frequentReschedules.push({
                appointmentId: appointmentId,
                patientName: patientName,
                mrdNo: mrdNo,
                rescheduleCount: history.length,
                lastReschedule: history[history.length - 1].timestamp,
                rescheduledBy: rescheduledBy
              });
            }
            
            // High risk appointments (3+ reschedules)
            if (history.length >= 3) {
              analytics.highRiskAppointments.push({
                appointmentId: appointmentId,
                patientName: patientName,
                mrdNo: mrdNo,
                rescheduleCount: history.length,
                lastReschedule: history[history.length - 1].timestamp,
                rescheduledBy: rescheduledBy,
                history: history
              });
            }
          } catch (e) {
            Logger.log("Error parsing reschedule history for appointment " + appointmentId + ": " + e.toString());
          }
        }
      }
    }
    
    // Process cancel sheet
    if (cancelSheet) {
      const cancelData = cancelSheet.getDataRange().getValues();
      const headers = cancelData[0];
      const rescheduleHistoryCol = 20; // Column U
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const appointmentIdCol = headers.indexOf("AppointmentID");
      const patientNameCol = headers.indexOf("PatientName");
      const mrdNoCol = headers.indexOf("MRDNo");
      
      for (let i = 1; i < cancelData.length; i++) {
        const row = cancelData[i];
        const rescheduleHistory = row[rescheduleHistoryCol];
        const rescheduledBy = row[rescheduledByCol];
        const appointmentId = row[appointmentIdCol];
        const patientName = row[patientNameCol];
        const mrdNo = row[mrdNoCol];
        
        if (rescheduleHistory) {
          try {
            const history = JSON.parse(rescheduleHistory);
            analytics.totalReschedules += history.length;
            
            // Track by staff member
            history.forEach(entry => {
              const staff = entry.by || 'Unknown';
              analytics.reschedulesByStaff[staff] = (analytics.reschedulesByStaff[staff] || 0) + 1;
              
              // Track by month
              const date = new Date(entry.timestamp);
              const monthKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
              analytics.reschedulesByMonth[monthKey] = (analytics.reschedulesByMonth[monthKey] || 0) + 1;
            });
            
            // Track frequent reschedules
            if (history.length >= 2) {
              analytics.frequentReschedules.push({
                appointmentId: appointmentId,
                patientName: patientName,
                mrdNo: mrdNo,
                rescheduleCount: history.length,
                lastReschedule: history[history.length - 1].timestamp,
                rescheduledBy: rescheduledBy,
                status: 'Cancelled'
              });
            }
          } catch (e) {
            Logger.log("Error parsing reschedule history for cancelled appointment " + appointmentId + ": " + e.toString());
          }
        }
      }
    }
    
    // Process archive sheet
    if (archiveSheet) {
      const archiveData = archiveSheet.getDataRange().getValues();
      const headers = archiveData[0];
      const rescheduleHistoryCol = 20; // Column U
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const appointmentIdCol = headers.indexOf("AppointmentID");
      const patientNameCol = headers.indexOf("PatientName");
      const mrdNoCol = headers.indexOf("MRDNo");
      
      for (let i = 1; i < archiveData.length; i++) {
        const row = archiveData[i];
        const rescheduleHistory = row[rescheduleHistoryCol];
        const rescheduledBy = row[rescheduledByCol];
        const appointmentId = row[appointmentIdCol];
        const patientName = row[patientNameCol];
        const mrdNo = row[mrdNoCol];
        
        if (rescheduleHistory) {
          try {
            const history = JSON.parse(rescheduleHistory);
            analytics.totalReschedules += history.length;
            
            // Track by staff member
            history.forEach(entry => {
              const staff = entry.by || 'Unknown';
              analytics.reschedulesByStaff[staff] = (analytics.reschedulesByStaff[staff] || 0) + 1;
              
              // Track by month
              const date = new Date(entry.timestamp);
              const monthKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
              analytics.reschedulesByMonth[monthKey] = (analytics.reschedulesByMonth[monthKey] || 0) + 1;
            });
          } catch (e) {
            Logger.log("Error parsing reschedule history for archived appointment " + appointmentId + ": " + e.toString());
          }
        }
      }
    }
    
    // Process appointment_archive sheet
    if (appointmentArchiveSheet) {
      const appointmentArchiveData = appointmentArchiveSheet.getDataRange().getValues();
      const headers = appointmentArchiveData[0];
      const rescheduleHistoryCol = 20; // Column U
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const appointmentIdCol = headers.indexOf("AppointmentID");
      const patientNameCol = headers.indexOf("PatientName");
      const mrdNoCol = headers.indexOf("MRDNo");
      
      for (let i = 1; i < appointmentArchiveData.length; i++) {
        const row = appointmentArchiveData[i];
        const rescheduleHistory = row[rescheduleHistoryCol];
        const rescheduledBy = row[rescheduledByCol];
        const appointmentId = row[appointmentIdCol];
        const patientName = row[patientNameCol];
        const mrdNo = row[mrdNoCol];
        
        if (rescheduleHistory && rescheduleHistory.toString().trim() !== '') {
          try {
            const history = JSON.parse(rescheduleHistory);
            Logger.log("Appointment Archive sheet - Found reschedule history for appointment " + appointmentId + ": " + history.length + " entries");
            analytics.totalReschedules += history.length;
            
            // Track by staff member
            history.forEach(entry => {
              const staff = entry.by || 'Unknown';
              analytics.reschedulesByStaff[staff] = (analytics.reschedulesByStaff[staff] || 0) + 1;
              
              // Track by month
              const date = new Date(entry.timestamp);
              const monthKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
              analytics.reschedulesByMonth[monthKey] = (analytics.reschedulesByMonth[monthKey] || 0) + 1;
            });
            
            // Track frequent reschedules
            if (history.length >= 2) {
              analytics.frequentReschedules.push({
                appointmentId: appointmentId,
                patientName: patientName,
                mrdNo: mrdNo,
                rescheduleCount: history.length,
                lastReschedule: history[history.length - 1].timestamp,
                rescheduledBy: rescheduledBy,
                status: 'Archived'
              });
            }
            
            // High risk appointments (3+ reschedules)
            if (history.length >= 3) {
              analytics.highRiskAppointments.push({
                appointmentId: appointmentId,
                patientName: patientName,
                mrdNo: mrdNo,
                rescheduleCount: history.length,
                lastReschedule: history[history.length - 1].timestamp,
                rescheduledBy: rescheduledBy,
                history: history,
                status: 'Archived'
              });
            }
          } catch (e) {
            Logger.log("Error parsing reschedule history for archived appointment " + appointmentId + ": " + e.toString());
          }
        }
      }
    }
    
    // Calculate averages and trends
    const totalAppointments = (appointmentSheet ? appointmentSheet.getLastRow() - 1 : 0) + 
                             (cancelSheet ? cancelSheet.getLastRow() - 1 : 0) + 
                             (archiveSheet ? archiveSheet.getLastRow() - 1 : 0) +
                             (appointmentArchiveSheet ? appointmentArchiveSheet.getLastRow() - 1 : 0);
    
    analytics.averageReschedulesPerAppointment = totalAppointments > 0 ? 
      (analytics.totalReschedules / totalAppointments).toFixed(2) : 0;
    
    // Calculate current month reschedules
    const currentDate = new Date();
    const currentMonth = `${currentDate.getFullYear()}-${(currentDate.getMonth() + 1).toString().padStart(2, '0')}`;
    
    // Also try UTC month key in case of timezone issues
    const utcMonth = `${currentDate.getUTCFullYear()}-${(currentDate.getUTCMonth() + 1).toString().padStart(2, '0')}`;
    
    analytics.currentMonthReschedules = analytics.reschedulesByMonth[currentMonth] || analytics.reschedulesByMonth[utcMonth] || 0;
    
    // Debug logging
    Logger.log("Current month key: " + currentMonth);
    Logger.log("UTC month key: " + utcMonth);
    Logger.log("Available month keys: " + Object.keys(analytics.reschedulesByMonth).join(", "));
    Logger.log("Current month reschedules: " + analytics.currentMonthReschedules);
    Logger.log("Full reschedulesByMonth object: " + JSON.stringify(analytics.reschedulesByMonth));
    
    // Sort frequent reschedules by count
    analytics.frequentReschedules.sort((a, b) => b.rescheduleCount - a.rescheduleCount);
    analytics.highRiskAppointments.sort((a, b) => b.rescheduleCount - a.rescheduleCount);
    
    // Sort staff by reschedule count
    analytics.reschedulesByStaff = Object.fromEntries(
      Object.entries(analytics.reschedulesByStaff).sort(([,a], [,b]) => b - a)
    );
    
    return analytics;
  } catch (e) {
    Logger.log("Error in getReschedulingAnalytics: " + e.toString());
    return {
      totalReschedules: 0,
      reschedulesByReason: {},
      frequentReschedules: [],
      reschedulesByStaff: {},
      reschedulesByMonth: {},
      highRiskAppointments: [],
      averageReschedulesPerAppointment: 0,
      rescheduleTrend: {},
      error: e.toString()
    };
  }
}

/**
 * Get rescheduling frequency for a specific appointment ID
 * @param {string} appointmentId - The appointment ID to check
 * @return {Object} Rescheduling details for the appointment
 */
function getAppointmentRescheduleHistory(appointmentId) {
  try {
    // Validate input
    if (!appointmentId) {
      return { error: "Appointment ID is required" };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ["appointment", "cancel", "Archive", "appointment_archive"];
    let foundAppointment = null;
    
    for (const sheetName of sheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log("Sheet not found: " + sheetName);
        continue;
      }
      
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        Logger.log("Sheet " + sheetName + " has no data rows");
        continue;
      }
      
      const headers = data[0];
      const appointmentIdCol = headers.indexOf("AppointmentID");
      const rescheduleHistoryCol = 20; // Column U
      const patientNameCol = headers.indexOf("PatientName");
      const mrdNoCol = headers.indexOf("MRDNo");
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const rescheduledDateCol = headers.indexOf("RescheduledDate");
      
      // Validate column indices
      if (appointmentIdCol === -1) {
        Logger.log("AppointmentID column not found in sheet: " + sheetName);
        continue;
      }
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[appointmentIdCol] && row[appointmentIdCol].toString() === appointmentId.toString()) {
          const rescheduleHistory = row[rescheduleHistoryCol];
          let history = [];
          
          if (rescheduleHistory && rescheduleHistory !== "") {
            try {
              history = JSON.parse(rescheduleHistory);
              if (!Array.isArray(history)) {
                history = [];
                Logger.log("Reschedule history is not an array for appointment: " + appointmentId);
              } else {
                // Convert old format to new format if needed
                history = history.map(entry => {
                  if (entry.previousDate && !entry.fromDate) {
                    // Convert old format to new format
                    return {
                      timestamp: entry.timestamp,
                      by: entry.by,
                      fromDate: entry.previousDate,
                      toDate: entry.previousDate, // We don't have the new date in old format
                      fromTime: entry.previousTime,
                      toTime: entry.previousTime, // We don't have the new time in old format
                      action: entry.action || "rescheduled"
                    };
                  }
                  return entry;
                });
              }
            } catch (e) {
              Logger.log("Error parsing reschedule history for appointment " + appointmentId + ": " + e.toString());
              history = [];
            }
          }
          
          foundAppointment = {
            appointmentId: appointmentId,
            patientName: row[patientNameCol] || 'Unknown',
            mrdNo: row[mrdNoCol] || 'Unknown',
            currentSheet: sheetName,
            rescheduleCount: history.length,
            rescheduleHistory: history,
            lastRescheduledBy: row[rescheduledByCol] || 'Unknown',
            lastRescheduledDate: row[rescheduledDateCol] || null,
            isHighRisk: history.length >= 3
          };
          break;
        }
      }
      
      if (foundAppointment) break;
    }
    
    if (!foundAppointment) {
      Logger.log("Appointment not found: " + appointmentId);
      
      // Debug: Let's check all sheets to see what's available
      const debugInfo = {
        appointmentId: appointmentId,
        sheetsChecked: [],
        availableSheets: []
      };
      
      const allSheets = ss.getSheets();
      for (const sheet of allSheets) {
        const sheetName = sheet.getName();
        debugInfo.availableSheets.push(sheetName);
        
        if (sheets.includes(sheetName)) {
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          const appointmentIdCol = headers.indexOf("AppointmentID");
          
          if (appointmentIdCol !== -1) {
            let foundInThisSheet = false;
            for (let i = 1; i < data.length; i++) {
              const row = data[i];
              if (row[appointmentIdCol] && row[appointmentIdCol].toString() === appointmentId.toString()) {
                foundInThisSheet = true;
                break;
              }
            }
            debugInfo.sheetsChecked.push({
              sheetName: sheetName,
              hasAppointmentId: true,
              foundAppointment: foundInThisSheet,
              totalRows: data.length - 1
            });
          } else {
            debugInfo.sheetsChecked.push({
              sheetName: sheetName,
              hasAppointmentId: false,
              foundAppointment: false,
              totalRows: data.length - 1
            });
          }
        }
      }
      
      Logger.log("Debug info: " + JSON.stringify(debugInfo));
      return { error: "Appointment not found", debug: debugInfo };
    }
    
    return foundAppointment;
  } catch (e) {
    Logger.log("Error in getAppointmentRescheduleHistory: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Get rescheduling statistics for dashboard cards
 * @return {Object} Quick stats for dashboard
 */
function getReschedulingStats() {
  try {
    const analytics = getReschedulingAnalytics();
    
    // Calculate trend (compare current month with previous month)
    const currentDate = new Date();
    const currentMonth = `${currentDate.getFullYear()}-${(currentDate.getMonth() + 1).toString().padStart(2, '0')}`;
    const previousMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
    const previousMonthKey = `${previousMonth.getFullYear()}-${(previousMonth.getMonth() + 1).toString().padStart(2, '0')}`;
    
    const currentMonthReschedules = analytics.reschedulesByMonth[currentMonth] || 0;
    const previousMonthReschedules = analytics.reschedulesByMonth[previousMonthKey] || 0;
    
    let trend = 0;
    if (previousMonthReschedules > 0) {
      trend = ((currentMonthReschedules - previousMonthReschedules) / previousMonthReschedules) * 100;
    }
    
    return {
      totalReschedules: analytics.totalReschedules,
      highRiskAppointments: analytics.highRiskAppointments.length,
      averageReschedulesPerAppointment: analytics.averageReschedulesPerAppointment,
      currentMonthReschedules: currentMonthReschedules,
      trend: trend,
      topReschedulingStaff: Object.keys(analytics.reschedulesByStaff).slice(0, 3)
    };
  } catch (e) {
    Logger.log("Error in getReschedulingStats: " + e.toString());
    return {
      totalReschedules: 0,
      highRiskAppointments: 0,
      averageReschedulesPerAppointment: 0,
      currentMonthReschedules: 0,
      trend: 0,
      topReschedulingStaff: [],
      error: e.toString()
    };
  }
}

/**
 * Test function to debug current month reschedule calculation
 * @return {Object} Debug information about current month calculation
 */
function debugCurrentMonthReschedules() {
  try {
    const currentDate = new Date();
    const currentMonth = `${currentDate.getFullYear()}-${(currentDate.getMonth() + 1).toString().padStart(2, '0')}`;
    const utcMonth = `${currentDate.getUTCFullYear()}-${(currentDate.getUTCMonth() + 1).toString().padStart(2, '0')}`;
    
    Logger.log("Current date: " + currentDate);
    Logger.log("Current month key: " + currentMonth);
    Logger.log("UTC month key: " + utcMonth);
    
    // Test parsing a sample timestamp
    const sampleTimestamp = "2025-06-26T13:08:38.388Z";
    const sampleDate = new Date(sampleTimestamp);
    const sampleMonthKey = `${sampleDate.getFullYear()}-${(sampleDate.getMonth() + 1).toString().padStart(2, '0')}`;
    Logger.log("Sample timestamp: " + sampleTimestamp);
    Logger.log("Sample parsed date: " + sampleDate);
    Logger.log("Sample month key: " + sampleMonthKey);
    
    // Get analytics to see what month keys exist
    const analytics = getReschedulingAnalytics();
    
    return {
      currentDate: currentDate.toString(),
      currentMonthKey: currentMonth,
      utcMonthKey: utcMonth,
      sampleTimestamp: sampleTimestamp,
      sampleMonthKey: sampleMonthKey,
      availableMonthKeys: Object.keys(analytics.reschedulesByMonth),
      currentMonthReschedules: analytics.currentMonthReschedules,
      totalReschedules: analytics.totalReschedules,
      reschedulesByMonth: analytics.reschedulesByMonth
    };
  } catch (e) {
    Logger.log("Error in debugCurrentMonthReschedules: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Debug function to check sheet structure and column positions
 * @return {Object} Sheet structure information
 */
function debugSheetStructure() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentSheet = ss.getSheetByName("appointment");
    const cancelSheet = ss.getSheetByName("cancel");
    const archiveSheet = ss.getSheetByName("Archive");
    
    const debugInfo = {
      appointment: null,
      cancel: null,
      archive: null
    };
    
    // Check appointment sheet
    if (appointmentSheet) {
      const appointmentData = appointmentSheet.getDataRange().getValues();
      const headers = appointmentData[0];
      
      debugInfo.appointment = {
        totalRows: appointmentData.length,
        headers: headers,
        rescheduleHistoryCol: headers.indexOf("RescheduleHistory"),
        rescheduleHistoryColU: headers.indexOf("U"), // Check if it's column U
        sampleData: appointmentData.length > 1 ? appointmentData[1] : null,
        cellU2: appointmentSheet.getRange("U2").getValue()
      };
    }
    
    // Check cancel sheet
    if (cancelSheet) {
      const cancelData = cancelSheet.getDataRange().getValues();
      const headers = cancelData[0];
      
      debugInfo.cancel = {
        totalRows: cancelData.length,
        headers: headers,
        rescheduleHistoryCol: headers.indexOf("RescheduleHistory"),
        sampleData: cancelData.length > 1 ? cancelData[1] : null
      };
    }
    
    // Check archive sheet
    if (archiveSheet) {
      const archiveData = archiveSheet.getDataRange().getValues();
      const headers = archiveData[0];
      
      debugInfo.archive = {
        totalRows: archiveData.length,
        headers: headers,
        rescheduleHistoryCol: headers.indexOf("RescheduleHistory"),
        sampleData: archiveData.length > 1 ? archiveData[1] : null
      };
    }
    
    return debugInfo;
  } catch (e) {
    Logger.log("Error in debugSheetStructure: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Get booking analytics data for visualization
 * Returns booking counts by user (excluding doctors and admins) for a date range
 */
function getBookingAnalytics(fromDate, toDate) {
  Logger.log("=== getBookingAnalytics START ===");
  Logger.log("Input parameters - fromDate: " + fromDate + ", toDate: " + toDate);
  
  try {
    Logger.log("Getting active spreadsheet...");
    const ss = SpreadsheetApp.getActive();
    Logger.log("Active spreadsheet obtained successfully");
    
    // Cache user roles at the beginning to avoid repeated lookups
    Logger.log("Caching user roles...");
    const userRoleCache = {};
    const loginSheet = ss.getSheetByName('Login');
    if (loginSheet) {
      const loginData = loginSheet.getDataRange().getValues();
      Logger.log("Login sheet has " + loginData.length + " rows");
      
      for (let i = 1; i < loginData.length; i++) {
        const username = loginData[i][0];
        const role = loginData[i][2] || 'user';
        userRoleCache[username] = role;
        Logger.log("Cached role for " + username + ": " + role);
      }
    }
    Logger.log("User role cache created with " + Object.keys(userRoleCache).length + " users");
    
    Logger.log("Looking for appointment sheets...");
    const appointmentSheet = ss.getSheetByName("appointment");
    const archiveSheet = ss.getSheetByName("appointment_archive");
    
    Logger.log("Appointment sheet found: " + (appointmentSheet ? "YES" : "NO"));
    Logger.log("Archive sheet found: " + (archiveSheet ? "YES" : "NO"));
    
    if (!appointmentSheet && !archiveSheet) {
      Logger.log("ERROR: No appointment sheets found");
      return { success: false, message: "No appointment sheets found" };
    }
    
    const bookingData = {};
    Logger.log("Initialized bookingData object");
    
    // Process appointment sheet
    if (appointmentSheet) {
      Logger.log("=== Processing appointment sheet ===");
      Logger.log("Getting data range from appointment sheet...");
      const data = appointmentSheet.getDataRange().getValues();
      Logger.log("Appointment sheet data retrieved. Rows: " + data.length);
      
      const headers = data[0];
      Logger.log("Headers found: " + headers.join(", "));
      
      const bookedByCol = headers.indexOf("BookedBy");
      const bookingDateCol = headers.indexOf("BookingDate");
      
      Logger.log("BookedBy column index: " + bookedByCol);
      Logger.log("BookingDate column index: " + bookingDateCol);
      
      if (bookedByCol === -1 || bookingDateCol === -1) {
        Logger.log("ERROR: Required columns not found in appointment sheet");
        return { success: false, message: "Required columns not found in appointment sheet" };
      }
      
      Logger.log("Starting to process " + (data.length - 1) + " rows in appointment sheet...");
      let processedRows = 0;
      let validBookings = 0;
      let skippedRows = 0;
      
      for (let i = 1; i < data.length; i++) {
        processedRows++;
        
        // Log progress every 100 rows
        if (processedRows % 100 === 0) {
          Logger.log("Processed " + processedRows + " rows in appointment sheet...");
        }
        
        const row = data[i];
        const bookedBy = row[bookedByCol];
        const bookingDate = row[bookingDateCol];
        
        if (bookedBy && bookingDate) {
          // Check if user is not doctor or admin using cached role
          const userRole = userRoleCache[bookedBy] || 'user';
          
          if (userRole !== 'doctor' && userRole !== 'admin') {
            // Check if booking date is within range
            try {
              const bookingDateObj = new Date(bookingDate);
              const fromDateObj = new Date(fromDate);
              const toDateObj = new Date(toDate);
              
              // Normalize dates to compare only date part (remove time component)
              const bookingDateOnly = new Date(bookingDateObj.getFullYear(), bookingDateObj.getMonth(), bookingDateObj.getDate());
              const fromDateOnly = new Date(fromDateObj.getFullYear(), fromDateObj.getMonth(), fromDateObj.getDate());
              const toDateOnly = new Date(toDateObj.getFullYear(), toDateObj.getMonth(), toDateObj.getDate());
              
              if (bookingDateOnly >= fromDateOnly && bookingDateOnly <= toDateOnly) {
                if (!bookingData[bookedBy]) {
                  bookingData[bookedBy] = 0;
                }
                bookingData[bookedBy]++;
                validBookings++;
              }
            } catch (dateError) {
              Logger.log("Date parsing error for row " + i + ": " + dateError.toString());
              skippedRows++;
              continue;
            }
          } else {
            skippedRows++;
          }
        } else {
          skippedRows++;
        }
      }
      
      Logger.log("=== Appointment sheet processing complete ===");
      Logger.log("Total rows processed: " + processedRows);
      Logger.log("Valid bookings found: " + validBookings);
      Logger.log("Skipped rows: " + skippedRows);
    }
    
    // Process archive sheet
    if (archiveSheet) {
      Logger.log("=== Processing archive sheet ===");
      Logger.log("Getting data range from archive sheet...");
      const data = archiveSheet.getDataRange().getValues();
      Logger.log("Archive sheet data retrieved. Rows: " + data.length);
      
      const headers = data[0];
      Logger.log("Archive headers found: " + headers.join(", "));
      
      const bookedByCol = headers.indexOf("BookedBy");
      const bookingDateCol = headers.indexOf("BookingDate");
      
      Logger.log("Archive BookedBy column index: " + bookedByCol);
      Logger.log("Archive BookingDate column index: " + bookingDateCol);
      
      if (bookedByCol !== -1 && bookingDateCol !== -1) {
        Logger.log("Starting to process " + (data.length - 1) + " rows in archive sheet...");
        let processedRows = 0;
        let validBookings = 0;
        let skippedRows = 0;
        
        for (let i = 1; i < data.length; i++) {
          processedRows++;
          
          // Log progress every 100 rows
          if (processedRows % 100 === 0) {
            Logger.log("Processed " + processedRows + " rows in archive sheet...");
          }
          
          const row = data[i];
          const bookedBy = row[bookedByCol];
          const bookingDate = row[bookingDateCol];
          
          if (bookedBy && bookingDate) {
            // Check if user is not doctor or admin using cached role
            const userRole = userRoleCache[bookedBy] || 'user';
            
            if (userRole !== 'doctor' && userRole !== 'admin') {
                          // Check if booking date is within range
            try {
              const bookingDateObj = new Date(bookingDate);
              const fromDateObj = new Date(fromDate);
              const toDateObj = new Date(toDate);
              
              // Normalize dates to compare only date part (remove time component)
              const bookingDateOnly = new Date(bookingDateObj.getFullYear(), bookingDateObj.getMonth(), bookingDateObj.getDate());
              const fromDateOnly = new Date(fromDateObj.getFullYear(), fromDateObj.getMonth(), fromDateObj.getDate());
              const toDateOnly = new Date(toDateObj.getFullYear(), toDateObj.getMonth(), toDateObj.getDate());
              
              if (bookingDateOnly >= fromDateOnly && bookingDateOnly <= toDateOnly) {
                  if (!bookingData[bookedBy]) {
                    bookingData[bookedBy] = 0;
                  }
                  bookingData[bookedBy]++;
                  validBookings++;
                }
              } catch (dateError) {
                Logger.log("Archive date parsing error for row " + i + ": " + dateError.toString());
                skippedRows++;
                continue;
              }
            } else {
              skippedRows++;
            }
          } else {
            skippedRows++;
          }
        }
        
        Logger.log("=== Archive sheet processing complete ===");
        Logger.log("Total archive rows processed: " + processedRows);
        Logger.log("Valid archive bookings found: " + validBookings);
        Logger.log("Skipped archive rows: " + skippedRows);
      } else {
        Logger.log("Required columns not found in archive sheet");
      }
    }
    
    Logger.log("=== Finalizing results ===");
    Logger.log("Final bookingData: " + JSON.stringify(bookingData));
    
    // Convert to array format for chart
    const chartData = Object.keys(bookingData).map(user => ({
      user: user,
      count: bookingData[user]
    })).sort((a, b) => b.count - a.count);
    
    Logger.log("Chart data prepared: " + JSON.stringify(chartData));
    
    const totalBookings = Object.values(bookingData).reduce((sum, count) => sum + count, 0);
    Logger.log("Total bookings calculated: " + totalBookings);
    
    Logger.log("=== getBookingAnalytics SUCCESS ===");
    return {
      success: true,
      data: chartData,
      totalBookings: totalBookings
    };
    
  } catch (error) {
    Logger.log("=== getBookingAnalytics ERROR ===");
    Logger.log("Error in getBookingAnalytics: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    return { success: false, message: error.toString() };
  }
}

/**
 * Get rescheduling analytics data for visualization
 * Returns rescheduling counts by user (excluding doctors and admins) for a date range
 */
function getReschedulingAnalyticsByDateRange(fromDate, toDate) {
      Logger.log("=== getReschedulingAnalyticsByDateRange START ===");
  Logger.log("Input parameters - fromDate: " + fromDate + ", toDate: " + toDate);
  
  try {
    Logger.log("Getting active spreadsheet...");
    const ss = SpreadsheetApp.getActive();
    Logger.log("Active spreadsheet obtained successfully");
    
    // Cache user roles at the beginning to avoid repeated lookups
    Logger.log("Caching user roles...");
    const userRoleCache = {};
    const loginSheet = ss.getSheetByName('Login');
    if (loginSheet) {
      const loginData = loginSheet.getDataRange().getValues();
      Logger.log("Login sheet has " + loginData.length + " rows");
      
      for (let i = 1; i < loginData.length; i++) {
        const username = loginData[i][0];
        const role = loginData[i][2] || 'user';
        userRoleCache[username] = role;
        Logger.log("Cached role for " + username + ": " + role);
      }
    }
    Logger.log("User role cache created with " + Object.keys(userRoleCache).length + " users");
    
    Logger.log("Looking for appointment sheets...");
    const appointmentSheet = ss.getSheetByName("appointment");
    const archiveSheet = ss.getSheetByName("appointment_archive");
    
    Logger.log("Appointment sheet found: " + (appointmentSheet ? "YES" : "NO"));
    Logger.log("Archive sheet found: " + (archiveSheet ? "YES" : "NO"));
    
    if (!appointmentSheet && !archiveSheet) {
      Logger.log("ERROR: No appointment sheets found");
      return { success: false, message: "No appointment sheets found" };
    }
    
    const reschedulingData = {};
    Logger.log("Initialized reschedulingData object");
    
    // Process appointment sheet
    if (appointmentSheet) {
      Logger.log("=== Processing appointment sheet ===");
      Logger.log("Getting data range from appointment sheet...");
      const data = appointmentSheet.getDataRange().getValues();
      Logger.log("Appointment sheet data retrieved. Rows: " + data.length);
      
      const headers = data[0];
      Logger.log("Headers found: " + headers.join(", "));
      
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const rescheduledDateCol = headers.indexOf("RescheduledDate");
      
      Logger.log("RescheduledBy column index: " + rescheduledByCol);
      Logger.log("RescheduledDate column index: " + rescheduledDateCol);
      
      if (rescheduledByCol === -1 || rescheduledDateCol === -1) {
        Logger.log("ERROR: Required columns not found in appointment sheet");
        return { success: false, message: "Required columns not found in appointment sheet" };
      }
      
      Logger.log("Starting to process " + (data.length - 1) + " rows in appointment sheet...");
      let processedRows = 0;
      let validReschedules = 0;
      let skippedRows = 0;
      
      for (let i = 1; i < data.length; i++) {
        processedRows++;
        
        // Log progress every 100 rows
        if (processedRows % 100 === 0) {
          Logger.log("Processed " + processedRows + " rows in appointment sheet...");
        }
        
        const row = data[i];
        const rescheduledBy = row[rescheduledByCol];
        const rescheduledDate = row[rescheduledDateCol];
        
        if (rescheduledBy && rescheduledDate) {
          // Check if user is not doctor or admin using cached role
          const userRole = userRoleCache[rescheduledBy] || 'user';
          
          if (userRole !== 'doctor' && userRole !== 'admin') {
            // Check if reschedule date is within range
            try {
              const rescheduledDateObj = new Date(rescheduledDate);
              const fromDateObj = new Date(fromDate);
              const toDateObj = new Date(toDate);
              // Normalize dates to compare only date part (remove time component)
              const rescheduledDateOnly = new Date(rescheduledDateObj.getFullYear(), rescheduledDateObj.getMonth(), rescheduledDateObj.getDate());
              const fromDateOnly = new Date(fromDateObj.getFullYear(), fromDateObj.getMonth(), fromDateObj.getDate());
              const toDateOnly = new Date(toDateObj.getFullYear(), toDateObj.getMonth(), toDateObj.getDate());
              if (rescheduledDateOnly >= fromDateOnly && rescheduledDateOnly <= toDateOnly) {
                if (!reschedulingData[rescheduledBy]) {
                  reschedulingData[rescheduledBy] = 0;
                }
                reschedulingData[rescheduledBy]++;
                validReschedules++;
              }
            } catch (dateError) {
              Logger.log("Date parsing error for row " + i + ": " + dateError.toString());
              skippedRows++;
              continue;
            }
          } else {
            skippedRows++;
          }
        } else {
          skippedRows++;
        }
      }
      
      Logger.log("=== Appointment sheet processing complete ===");
      Logger.log("Total rows processed: " + processedRows);
      Logger.log("Valid reschedules found: " + validReschedules);
      Logger.log("Skipped rows: " + skippedRows);
    }
    
    // Process archive sheet
    if (archiveSheet) {
      Logger.log("=== Processing archive sheet ===");
      Logger.log("Getting data range from archive sheet...");
      const data = archiveSheet.getDataRange().getValues();
      Logger.log("Archive sheet data retrieved. Rows: " + data.length);
      
      const headers = data[0];
      Logger.log("Archive headers found: " + headers.join(", "));
      
      const rescheduledByCol = headers.indexOf("RescheduledBy");
      const rescheduledDateCol = headers.indexOf("RescheduledDate");
      
      Logger.log("Archive RescheduledBy column index: " + rescheduledByCol);
      Logger.log("Archive RescheduledDate column index: " + rescheduledDateCol);
      
      if (rescheduledByCol !== -1 && rescheduledDateCol !== -1) {
        Logger.log("Starting to process " + (data.length - 1) + " rows in archive sheet...");
        let processedRows = 0;
        let validReschedules = 0;
        let skippedRows = 0;
        
        for (let i = 1; i < data.length; i++) {
          processedRows++;
          
          // Log progress every 100 rows
          if (processedRows % 100 === 0) {
            Logger.log("Processed " + processedRows + " rows in archive sheet...");
          }
          
          const row = data[i];
          const rescheduledBy = row[rescheduledByCol];
          const rescheduledDate = row[rescheduledDateCol];
          
          if (rescheduledBy && rescheduledDate) {
            // Check if user is not doctor or admin using cached role
            const userRole = userRoleCache[rescheduledBy] || 'user';
            
            if (userRole !== 'doctor' && userRole !== 'admin') {
              // Check if reschedule date is within range
              try {
                const rescheduledDateObj = new Date(rescheduledDate);
                const fromDateObj = new Date(fromDate);
                const toDateObj = new Date(toDate);
                // Normalize dates to compare only date part (remove time component)
                const rescheduledDateOnly = new Date(rescheduledDateObj.getFullYear(), rescheduledDateObj.getMonth(), rescheduledDateObj.getDate());
                const fromDateOnly = new Date(fromDateObj.getFullYear(), fromDateObj.getMonth(), fromDateObj.getDate());
                const toDateOnly = new Date(toDateObj.getFullYear(), toDateObj.getMonth(), toDateObj.getDate());
                if (rescheduledDateOnly >= fromDateOnly && rescheduledDateOnly <= toDateOnly) {
                  if (!reschedulingData[rescheduledBy]) {
                    reschedulingData[rescheduledBy] = 0;
                  }
                  reschedulingData[rescheduledBy]++;
                  validReschedules++;
                }
              } catch (dateError) {
                Logger.log("Archive date parsing error for row " + i + ": " + dateError.toString());
                skippedRows++;
                continue;
              }
            } else {
              skippedRows++;
            }
          } else {
            skippedRows++;
          }
        }
        
        Logger.log("=== Archive sheet processing complete ===");
        Logger.log("Total archive rows processed: " + processedRows);
        Logger.log("Valid archive reschedules found: " + validReschedules);
        Logger.log("Skipped archive rows: " + skippedRows);
      } else {
        Logger.log("Required columns not found in archive sheet");
      }
    }
    
    Logger.log("=== Finalizing results ===");
    Logger.log("Final reschedulingData: " + JSON.stringify(reschedulingData));
    
    // Convert to array format for chart
    const chartData = Object.keys(reschedulingData).map(user => ({
      user: user,
      count: reschedulingData[user]
    })).sort((a, b) => b.count - a.count);
    
    Logger.log("Chart data prepared: " + JSON.stringify(chartData));
    
    const totalReschedules = Object.values(reschedulingData).reduce((sum, count) => sum + count, 0);
    Logger.log("Total reschedules calculated: " + totalReschedules);
    
    Logger.log("=== getReschedulingAnalyticsByDateRange SUCCESS ===");
    return {
      success: true,
      data: chartData,
      totalReschedules: totalReschedules
    };
    
  } catch (error) {
    Logger.log("=== getReschedulingAnalyticsByDateRange ERROR ===");
    Logger.log("Error in getReschedulingAnalyticsByDateRange: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    return { success: false, message: error.toString() };
  }
}
