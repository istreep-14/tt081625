/**
 * Bar Employee CRM - Google Sheets Add-on
 * This creates a modal popup CRM directly in Google Sheets
 * 
 * Setup Instructions:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Replace Code.gs content with this script
 * 4. Save and run the onOpen function
 * 5. Refresh your Google Sheet
 * 6. You'll see "Employee CRM" in the menu bar
 */

// Configuration
const SHEET_NAME = 'Sheet2';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Headers in order
const HEADERS = ['Emp Id', 'First Name', 'Last Name', 'Phone', 'Email', 'Position', 'Status', 'Note', 'Photo URL'];

/**
 * Runs when the spreadsheet is opened - adds the CRM menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Employee CRM')
    .addItem('Open CRM Manager', 'openCRMDialog')
    .addSeparator()
    .addItem('Initialize Sheet2', 'initializeSheet')
    .addToUi();
}

/**
 * Opens the CRM modal dialog
 */
function openCRMDialog() {
  const html = HtmlService.createTemplateFromFile('CRMDialog');
  const htmlOutput = html.evaluate()
    .setWidth(1200)
    .setHeight(700)
    .setTitle('Bar Employee CRM Manager');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bar Employee CRM Manager');
}

/**
 * Include external files (for CSS/JS in HTML template)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Persist positions list in script properties
 */
function getPositionsList() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('CRM_POSITIONS_LIST') || '';
    const positions = raw ? JSON.parse(raw) : [];
    return { success: true, positions };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function savePositionsList(list) {
  try {
    const normalized = (list || []).map(function(s){ return String(s || '').trim(); }).filter(function(s){ return s.length > 0; });
    PropertiesService.getScriptProperties().setProperty('CRM_POSITIONS_LIST', JSON.stringify(normalized));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Persist and retrieve the "Me" employee id
 */
function getMeEmployeeId() {
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('CRM_ME_EMP_ID') || '';
    return { success: true, empId: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function setMeEmployeeId(empId) {
  try {
    PropertiesService.getScriptProperties().setProperty('CRM_ME_EMP_ID', String(empId || ''));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function clearMeEmployeeId() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('CRM_ME_EMP_ID');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Initialize the sheet with headers if they don't exist
 */
function initializeSheet() {
  const sheet = getSheet();
  
  // Check if headers exist
  const existingHeaders = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = existingHeaders.some(header => header !== '');
  
  if (!hasHeaders) {
    // Set headers
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    // Format headers
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    headerRange.setBackground('#f8f9fa');
    headerRange.setFontWeight('bold');
    headerRange.setBorder(true, true, true, true, true, true);
    
    SpreadsheetApp.getUi().alert('Success', 'Sheet2 has been initialized with employee headers!', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Sheet initialized with headers');
  } else {
    // NOTE: HEADERS now includes a 'Status' column; initializeSheet will write updated headers if not present.
    // If headers already exist, we need to adjust the index of the 'Status' column
    // to match the new HEADERS array.
    const statusIndex = HEADERS.indexOf('Status');
    const existingStatusIndex = existingHeaders.indexOf('Status');
    
    if (existingStatusIndex !== -1 && statusIndex !== existingStatusIndex) {
      // If the 'Status' column is at a different index, move it to the correct position
      const newHeaders = [...HEADERS];
      const statusHeader = newHeaders.splice(statusIndex, 1)[0];
      newHeaders.splice(existingStatusIndex, 0, statusHeader);
      
      sheet.getRange(HEADER_ROW, 1, 1, newHeaders.length).setValues([newHeaders]);
      Logger.log(`Adjusted Status column position from ${existingStatusIndex} to ${statusIndex}`);
    }
    SpreadsheetApp.getUi().alert('Info', 'Sheet2 already has headers initialized.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  return sheet;
}

/**
 * Get or create the target sheet
 */
function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    Logger.log(`Created new sheet: ${SHEET_NAME}`);
  }
  
  return sheet;
}

/**
 * Get all employees from the sheet
 */
function getAllEmployees() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: true, employees: [] };
    }
    
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, HEADERS.length);
    const values = dataRange.getValues();
    
    const employees = values
      .filter(row => row[0] !== '') // Filter out empty rows (Emp Id is required)
      .map(row => ({
        empId: row[0] || '',
        firstName: row[1] || '',
        lastName: row[2] || '',
        phone: row[3] || '',
        email: row[4] || '',
        position: row[5] || '',
        status: row[6] || '',
        note: row[7] || '',
        photoUrl: row[8] || ''
      }));
    
    Logger.log(`Retrieved ${employees.length} employees`);
    return { success: true, employees: employees };
    
  } catch (error) {
    Logger.log('Error in getAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save all employees to the sheet (replaces existing data)
 */
function saveAllEmployees(employees) {
  try {
    const sheet = getSheet();
    
    // Ensure headers are present and up-to-date
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    headerRange.setBackground('#f8f9fa');
    headerRange.setFontWeight('bold');
    headerRange.setBorder(true, true, true, true, true, true);
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow >= DATA_START_ROW) {
      sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW + 1, HEADERS.length).clear();
    }
    
    if (employees && employees.length > 0) {
      // Prepare data rows
      const dataRows = employees.map(emp => [
        emp.empId || '',
        emp.firstName || '',
        emp.lastName || '',
        emp.phone || '',
        emp.email || '',
        emp.position || '',
        emp.status || '',
        emp.note || '',
        emp.photoUrl || ''
      ]);
      
      // Write data to sheet
      const range = sheet.getRange(DATA_START_ROW, 1, dataRows.length, HEADERS.length);
      range.setValues(dataRows);
      
      Logger.log(`Saved ${employees.length} employees`);
    }
    
    return { success: true, message: `Saved ${employees ? employees.length : 0} employees` };
    
  } catch (error) {
    Logger.log('Error in saveAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Add a single employee
 */
function addEmployee(employee) {
  try {
    const sheet = getSheet();
    
    // Check for duplicate Emp ID
    const existingData = getAllEmployees();
    if (existingData.success) {
      const duplicate = existingData.employees.find(emp => emp.empId === employee.empId);
      if (duplicate) {
        return { success: false, error: 'Employee ID already exists' };
      }
    }
    
    // Add to the end of the sheet
    const newRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.status || '',
      employee.note || '',
      employee.photoUrl || ''
    ];
    
    sheet.appendRow(newRow);
    
    Logger.log(`Added employee: ${employee.empId}`);
    return { success: true, message: 'Employee added successfully' };
    
  } catch (error) {
    Logger.log('Error in addEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update an existing employee
 */
function updateEmployee(employee, originalEmpId) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1);
    const empIds = dataRange.getValues().flat();
    const rowIndex = empIds.findIndex(id => id === originalEmpId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Check for duplicate Emp ID if it's being changed
    if (employee.empId !== originalEmpId) {
      const duplicate = empIds.findIndex(id => id === employee.empId);
      if (duplicate !== -1) {
        return { success: false, error: 'New Employee ID already exists' };
      }
    }
    
    // Update the row
    const targetRow = DATA_START_ROW + rowIndex;
    const updatedRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.status || '',
      employee.note || '',
      employee.photoUrl || ''
    ];
    
    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([updatedRow]);
    
    Logger.log(`Updated employee: ${originalEmpId} -> ${employee.empId}`);
    return { success: true, message: 'Employee updated successfully' };
    
  } catch (error) {
    Logger.log('Error in updateEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete an employee
 */
function deleteEmployee(empId) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1);
    const empIds = dataRange.getValues().flat();
    const rowIndex = empIds.findIndex(id => id === empId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Delete the row
    const targetRow = DATA_START_ROW + rowIndex;
    sheet.deleteRow(targetRow);
    
    Logger.log(`Deleted employee: ${empId}`);
    return { success: true, message: 'Employee deleted successfully' };
    
  } catch (error) {
    Logger.log('Error in deleteEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Upload a photo to Google Drive and return an embeddable URL
 */
function uploadEmployeePhoto(dataUrl, fileName, empId) {
  try {
    if (!dataUrl) {
      return { success: false, error: 'Missing dataUrl' };
    }
    const matches = dataUrl.match(/^data:(.+);base64,(.*)$/);
    if (!matches) {
      return { success: false, error: 'Invalid data URL' };
    }
    const contentType = matches[1];
    const base64Data = matches[2];
    const bytes = Utilities.base64Decode(base64Data);
    const tz = Session.getScriptTimeZone() || 'Etc/GMT';
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    const safeEmpId = (empId || 'employee').toString().replace(/[^a-zA-Z0-9_-]+/g, '_');
    const baseName = `${safeEmpId}_${timestamp}`;
    const finalName = fileName ? `${baseName}_${fileName}` : `${baseName}.png`;
    const blob = Utilities.newBlob(bytes, contentType, finalName);

    const folderName = 'Employee CRM Photos';
    let folderIter = DriveApp.getFoldersByName(folderName);
    let folder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(blob).setName(finalName);
    // Make viewable via link for embedding in <img>
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // In some domains, ANYONE_WITH_LINK may be restricted; ignore if fails
    }
    const id = file.getId();
    const viewUrl = `https://drive.google.com/uc?export=view&id=${id}`;
    return { success: true, url: viewUrl, id: id };
  } catch (error) {
    Logger.log('Error in uploadEmployeePhoto: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
