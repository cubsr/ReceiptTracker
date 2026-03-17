debugFlag = true;
/**
 * Creates custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage Transactions')
    .addItem('Add Manual Entries to Sheet', 'applyManualEntries')
    .addItem('Add Categories from Code', 'regenerateSheetCategories')
    .addToUi();
  if (!debugFlag) {return;}
  ui.createMenu('Receipt Tracking')
    .addItem('Initialize Spreadsheet', 'initializeSpreadsheet')
    .addItem('Add Profit Table', 'addProfitTable')
    .addItem('Add Net Results Table', 'addNetResultsTable')
    .addItem('Add Manual Entries Log (Migration)', 'addManualEntriesLogToSheet')
    .addSeparator()
    .addItem('Populate Test Data', 'populateTestData')
    .addSeparator()
    .addItem('Show Valid Categories', 'showValidCategories')
    .addItem('Test Dynamic Categories', 'testDynamicCategories')
    .addItem('Create Receipts Folder', 'createReceiptsFolderAndGetId')
    .addToUi();
  
}

// CATEGORIES, CATEGORY_ALIASES → Config
// API_KEY_USERS, RECEIPTS_FOLDER_ID → Config

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];

const MANUAL_LOG_START_COL = 9; // Column I — manual entries log (to the right of the transaction log)

/**
 * Main function - accepts photo as base64 string
 * Params:
 * date, category, amount, apiKey, photosBase64
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    
    // SECURITY CHECK: Verify API key
    if (!params.apiKey || !API_KEY_USERS[params.apiKey]) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Invalid or missing API key. Please check your shortcut configuration.'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get user name from API key
    const userName = API_KEY_USERS[params.apiKey];
    
    // Validate required fields
    if (!params.date || !params.category || !params.amount) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Missing required fields: date, category, amount'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Normalize and validate category
    const inputCategory = params.category.toLowerCase().trim();
    let category;
    
    // Check if input is an alias first
    if (CATEGORY_ALIASES[inputCategory]) {
      category = CATEGORY_ALIASES[inputCategory];
    } 
    // Check if input matches a category exactly (case-insensitive)
    else {
      const matchedCategory = CATEGORIES.find(cat => cat.toLowerCase() === inputCategory);
      if (matchedCategory) {
        category = matchedCategory;
      } else {
        // Dynamic category - use the input as-is (capitalize first letter of each word)
        category = params.category.split(' ').map(word => 
          word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
        ).join(' ');
      }
    }
    
    // Sanitize and validate amount
    const sanitizedAmount = params.amount.toString().replace(/[$,\s]/g, ''); // Remove $, commas, and spaces
    const amount = parseFloat(sanitizedAmount);
    if (isNaN(amount) || amount < 0) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Amount must be a valid positive number'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Upload photos to Google Drive if provided
    let fileLinks = [];
    if (params.photosBase64 && Array.isArray(params.photosBase64) && params.photosBase64.length > 0) {
      try {
        fileLinks = uploadPhotosToDrive(
          params.photosBase64,
          params.date,
          category,
          userName
        );
      } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Failed to upload photos: ' + error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    } else if (params.photoBase64) {
      // Backward compatibility - single photo
      try {
        const singleLink = uploadPhotoToDrive(
          params.photoBase64,
          params.date,
          category,
          userName
        );
        fileLinks = [singleLink];
      } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Failed to upload photo: ' + error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // Add to spreadsheet
    const result = addReceipt(
      params.date,
      category,
      amount,
      fileLinks,
      userName
    );
    
    // Create user-friendly success message
    const date = new Date(params.date);
    const month = MONTHS[date.getMonth()];
    const year = date.getFullYear();
    const friendlyDate = `${month} ${date.getDate()}, ${year}`;
    
    const photoCount = fileLinks.length;
    const successMessage = `✅ Receipt added successfully!\n\n` +
      `📅 Date: ${friendlyDate}\n` +
      `🏷️ Category: ${category}\n` +
      `💰 Amount: $${amount}\n` +
      `👤 Added by: ${userName}` +
      (photoCount > 0 ? `\n📎 ${photoCount} photo${photoCount > 1 ? 's' : ''} saved to Drive` : '');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: successMessage,
      fileLinks: fileLinks,
      fileLink: fileLinks.length > 0 ? fileLinks[0] : '', // Backward compatibility
      summary: {
        date: friendlyDate,
        category: category,
        amount: `$${amount}`,
        user: userName,
        photoCount: photoCount,
        hasPhotos: photoCount > 0
      }
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Something went wrong. Please try again.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Uploads multiple photos to Google Drive in organized folder structure
 */
function uploadPhotosToDrive(photosBase64, dateString, category, userName) {
  const fileLinks = [];
  
  photosBase64.forEach((base64Data, index) => {
    try {
      const fileLink = uploadPhotoToDrive(base64Data, dateString, category, userName, index);
      fileLinks.push(fileLink);
    } catch (error) {
      Logger.log(`Error uploading photo ${index + 1}: ${error.toString()}`);
      throw new Error(`Failed to upload photo ${index + 1}: ${error.toString()}`);
    }
  });
  
  return fileLinks;
}

/**
 * Uploads photo to Google Drive in organized folder structure
 */
function uploadPhotoToDrive(base64Data, dateString, category, userName, photoIndex = 0) {
  // Parse date in script timezone to avoid UTC conversion issues
  let date;
  if (dateString.includes('T')) {
    date = new Date(dateString);
  } else {
    const parts = dateString.split('-');
    date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  
  const year = date.getFullYear();
  const month = MONTHS[date.getMonth()];
  
  // Get or create root Receipts folder
  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(RECEIPTS_FOLDER_ID);
  } catch (e) {
    // If folder ID not set or invalid, create in root
    rootFolder = DriveApp.createFolder('Receipts');
    Logger.log('Created Receipts folder. ID: ' + rootFolder.getId());
  }
  
  // Get or create year folder
  const yearFolderName = year.toString();
  let yearFolder = getOrCreateFolder(rootFolder, yearFolderName);
  
  // Get or create month folder
  let monthFolder = getOrCreateFolder(yearFolder, month);
  
  // Create filename with timestamp
  const timestamp = Utilities.formatDate(date, 'America/Chicago', 'yyyy-MM-dd_HH-mm-ss');
  const filename = photoIndex > 0 
    ? `receipt_${category}_${timestamp}_${photoIndex + 1}.jpg`
    : `receipt_${category}_${timestamp}.jpg`;
  
  // Decode base64 and create file
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    'image/jpeg',
    filename
  );
  
  const file = monthFolder.createFile(blob);
  
  // Make file accessible via link
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

/**
 * Helper function to get or create a folder
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

/**
 * Add receipt to spreadsheet
 */
function addReceipt(dateString, category, amount, fileLinks, userName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Parse date in script timezone to avoid UTC conversion issues
  // Handle both 'YYYY-MM-DD' and full ISO date strings
  let date;
  if (dateString.includes('T')) {
    // Full ISO string with time
    date = new Date(dateString);
  } else {
    // Date-only string (YYYY-MM-DD) - parse as local date
    const parts = dateString.split('-');
    date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  
  const year = date.getFullYear();
  const month = MONTHS[date.getMonth()];
  
  // Get or create the year's sheet
  let sheet = ss.getSheetByName(year.toString());
  if (!sheet) {
    sheet = createYearSheet(ss, year);
  }
  
  // Check if category exists in sheet, add if new
  const categoryReady = ensureCategoryExists(sheet, category);
  if (!categoryReady) {
    // Category doesn't exist and can't be added (past year sheet)
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes(category)) {
      throw new Error(`Category "${category}" does not exist in the ${year} sheet. Past year sheets cannot have new categories added.`);
    }
  }

  // Add transaction to log
  addTransaction(sheet, date, month, category, amount, fileLinks, userName);

  // Update monthly summary table
  updateMonthlySummary(sheet, month, category, amount);

  return `Receipt logged: ${category} - $${amount} for ${month} ${year}`;
}

/**
 * Ensures a category exists in the sheet, adds it if it doesn't.
 * Returns true if the category exists (or was added), false if blocked (past year).
 */
function ensureCategoryExists(sheet, category) {
  // Only add new categories to current year and future sheets
  const sheetYear = parseInt(sheet.getName());
  const currentYear = new Date().getFullYear();
  if (!isNaN(sheetYear) && sheetYear < currentYear) {
    Logger.log(`Category "${category}" not added to past year sheet ${sheetYear}`);
    return false;
  }

  // Get current categories from header row (row 2)
  const headerRow = 2;
  const headerRange = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];

  // Check if category already exists
  const categoryExists = headers.includes(category);

  if (!categoryExists) {
    // Find the column before TOTAL (last column)
    const totalColumn = headers.indexOf('TOTAL');
    const insertColumn = totalColumn > -1 ? totalColumn : headers.length;

    // Insert new column for the category
    sheet.insertColumnBefore(insertColumn + 1);

    // Add category header in main summary
    sheet.getRange(headerRow, insertColumn + 1).setValue(category)
      .setFontWeight('bold')
      .setBackground('#34A853')
      .setFontColor('white');

    // Initialize all month rows with 0 for this category
    const firstMonthRow = 3;
    const lastMonthRow = firstMonthRow + MONTHS.length - 1;

    for (let row = firstMonthRow; row <= lastMonthRow; row++) {
      sheet.getRange(row, insertColumn + 1).setValue(0).setNumberFormat('$#,##0.00');
    }

    // Update annual total row
    const annualTotalRow = lastMonthRow + 1;
    const colLetter = columnToLetter(insertColumn + 1);
    const formula = `=SUM(${colLetter}${firstMonthRow}:${colLetter}${lastMonthRow})`;
    sheet.getRange(annualTotalRow, insertColumn + 1).setFormula(formula)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold')
      .setBackground('#FBBC04');

    // Update TOTAL column formulas to include new category
    updateTotalFormulas(sheet, insertColumn + 1);

    Logger.log(`Added new category "${category}" to sheet`);
  }
  return true;
}

/**
 * Updates TOTAL column formulas to include all categories
 */
function updateTotalFormulas(sheet, newCategoryColumn) {
  const headerRow = 2;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const totalColumn = headers.indexOf('TOTAL') + 1;
  
  if (totalColumn > 0) {
    const firstMonthRow = 3;
    const lastMonthRow = firstMonthRow + MONTHS.length;
    
    // Update monthly totals
    for (let row = firstMonthRow; row <= lastMonthRow; row++) {
      const startCol = 2; // First category column
      const endCol = totalColumn - 1; // Column before TOTAL
      const formula = `=SUM(B${row}:${String.fromCharCode(65 + endCol - 1)}${row})`;
      sheet.getRange(row, totalColumn).setFormula(formula);
    }
  }
}

/**
 * Builds the manual entries log to the right of the transaction log (cols I-M).
 * Includes date picker validation and category dropdown.
 * Called once during sheet creation; also used to add the log to existing sheets.
 */
function createManualEntriesLog(sheet, year) {
  const titleRow = findTransactionLogStartRow(sheet) - 2; // same row as TX LOG title
  const headerRow = titleRow + 1;
  const firstDataRow = titleRow + 2;

  // Title - merged across 5 manual log columns
  sheet.getRange(titleRow, MANUAL_LOG_START_COL, 1, 5).merge()
    .setValue('MANUAL ENTRIES - ' + year)
    .setFontWeight('bold').setFontSize(14);

  // Headers
  sheet.getRange(headerRow, MANUAL_LOG_START_COL, 1, 5)
    .setValues([['Date', 'Category', 'Amount', 'Note', 'Added By']])
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('white');

  const bufferRows = 200;

  // Date validation — triggers calendar picker in Sheets
  const dateRange = sheet.getRange(firstDataRow, MANUAL_LOG_START_COL, bufferRows, 1);
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  dateRange.setDataValidation(dateRule).setNumberFormat('MM/dd/yyyy');

  // Category dropdown from CATEGORIES array
  const catRange = sheet.getRange(firstDataRow, MANUAL_LOG_START_COL + 1, bufferRows, 1);
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIES, true)
    .setAllowInvalid(true)
    .build();
  catRange.setDataValidation(catRule);

  // Amount column — currency format, supports negatives
  sheet.getRange(firstDataRow, MANUAL_LOG_START_COL + 2, bufferRows, 1)
    .setNumberFormat('$#,##0.00');

  // Column widths for the manual log
  sheet.setColumnWidth(MANUAL_LOG_START_COL, 110);     // Date
  sheet.setColumnWidth(MANUAL_LOG_START_COL + 1, 130); // Category
  sheet.setColumnWidth(MANUAL_LOG_START_COL + 2, 100); // Amount
  sheet.setColumnWidth(MANUAL_LOG_START_COL + 3, 180); // Note
  sheet.setColumnWidth(MANUAL_LOG_START_COL + 4, 100); // Added By
}

/**
 * Re-applies the category dropdown validation to the manual log after
 * the CATEGORIES array changes (called by regenerateSheetCategories).
 */
function refreshManualLogValidation(sheet) {
  const firstDataRow = findTransactionLogStartRow(sheet);
  const manualStartCol = findManualLogStartCol(sheet);
  const catRange = sheet.getRange(firstDataRow, manualStartCol + 1, 200, 1);
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIES, true)
    .setAllowInvalid(true)
    .build();
  catRange.setDataValidation(catRule);
}

/**
 * Converts a 1-based column number to a letter (e.g. 1→A, 26→Z, 27→AA)
 */
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Creates a new sheet for the year with proper structure
 *
 * SHEET LAYOUT:
 * Rows 1-15:  Monthly Summary Table (title, headers, 12 months, annual total)
 * Row 16:     Blank row for spacing
 * Row 17:     Transaction Log title (col A) | Manual Entries title (col I)
 * Row 18:     TX log headers (cols A-G)     | Manual log headers  (cols I-M)
 * Row 19+:    Transaction entries            | Manual entry rows
 */
function createYearSheet(spreadsheet, year) {
  const sheet = spreadsheet.insertSheet(year.toString());

  // Create Monthly Summary Table (rows 1-15)
  createMonthlySummaryTable(sheet, year);

  // Transaction Log header (cols A-G, rows 17-18)
  sheet.getRange(17, 1).setValue('TRANSACTION LOG - ' + year)
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange(18, 1, 1, 7).setValues([[
    'Date', 'Month', 'Category', 'Amount', 'Receipt Link', 'Added By', 'Notes'
  ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');

  // Manual Entries Log (cols I-M, same rows 17-18+)
  createManualEntriesLog(sheet, year);

  // Freeze rows through transaction log header
  sheet.setFrozenRows(18);
  
  // Set column widths
  sheet.setColumnWidth(1, 100); // Date/Month
  sheet.setColumnWidth(2, 120); // Month/Category in summary
  sheet.setColumnWidth(3, 120); // Category
  sheet.setColumnWidth(4, 100); // Amount
  sheet.setColumnWidth(5, 200); // Link
  sheet.setColumnWidth(6, 100); // Added By
  sheet.setColumnWidth(7, 150); // Notes/Additional categories
  sheet.setColumnWidth(8, 100); // Additional categories
  sheet.setColumnWidth(9, 100); // Additional categories
  
  return sheet;
}

/**
 * Creates the monthly summary table with categories
 */
function createMonthlySummaryTable(sheet, year) {
  const startRow = 1; // Now at the top of the sheet!
  
  // Title
  sheet.getRange(startRow, 1).setValue('MONTHLY SPENDING BY CATEGORY - ' + year)
    .setFontWeight('bold').setFontSize(14).setBackground('#E8F0FE');
  
  // Headers
  const headers = ['Month', ...CATEGORIES, 'TOTAL'];
  sheet.getRange(startRow + 1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('white');
  
  // Month rows
  for (let i = 0; i < MONTHS.length; i++) {
    const row = startRow + 2 + i;
    sheet.getRange(row, 1).setValue(MONTHS[i]);
    
    // Initialize category columns with 0
    for (let j = 0; j < CATEGORIES.length; j++) {
      sheet.getRange(row, 2 + j).setValue(0).setNumberFormat('$#,##0.00');
    }
    
    // Total formula
    const totalFormula = `=SUM(B${row}:${String.fromCharCode(65 + CATEGORIES.length)}${row})`;
    sheet.getRange(row, 2 + CATEGORIES.length).setFormula(totalFormula)
      .setNumberFormat('$#,##0.00').setFontWeight('bold');
  }
  
  // Annual totals row
  const totalRow = startRow + 2 + MONTHS.length;
  sheet.getRange(totalRow, 1).setValue('ANNUAL TOTAL').setFontWeight('bold');
  
  for (let j = 0; j < CATEGORIES.length + 1; j++) {
    const col = 2 + j;
    const colLetter = String.fromCharCode(65 + col - 1);
    const formula = `=SUM(${colLetter}${startRow + 2}:${colLetter}${totalRow - 1})`;
    sheet.getRange(totalRow, col).setFormula(formula)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold')
      .setBackground('#FBBC04');
  }
  
  // Format the table
  const tableRange = sheet.getRange(startRow + 1, 1, MONTHS.length + 2, headers.length);
  tableRange.setBorder(true, true, true, true, true, true);
}

/**
 * Adds a profit table 2 columns to the right of the expenses table
 * Works on all year-named sheets that don't already have the table
 */
function addProfitTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const yearSheets = [];
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (/^\d{4}$/.test(sheetName)) {
      yearSheets.push(sheet);
    }
  }
  
  if (yearSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No year sheets found. Please initialize the spreadsheet first.');
    return;
  }
  
  let addedCount = 0;
  let skippedCount = 0;
  let skippedPastCount = 0;
  const currentYear = new Date().getFullYear();

  for (const sheet of yearSheets) {
    const sheetName = sheet.getName();

    if (parseInt(sheetName) < currentYear) {
      skippedPastCount++;
      continue;
    }

    if (sheetHasProfitTable(sheet)) {
      skippedCount++;
      continue;
    }

    addProfitTableToSheet(sheet, sheetName);
    addedCount++;
  }

  let message = '';
  if (addedCount > 0) {
    message += `Profit table added to ${addedCount} sheet(s).\n\n`;
  }
  if (skippedCount > 0) {
    message += `${skippedCount} sheet(s) already had the profit table.\n`;
  }
  if (skippedPastCount > 0) {
    message += `${skippedPastCount} past year sheet(s) skipped.`;
  }
  
  SpreadsheetApp.getUi().alert(
    'Profit Table Update Complete',
    message.trim(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Checks if a sheet already has a profit table
 */
function sheetHasProfitTable(sheet) {
  const expensesColumns = CATEGORIES.length + 2;
  const profitStartColumn = expensesColumns + 1;
  const cellValue = sheet.getRange(1, profitStartColumn).getValue();
  return cellValue && cellValue.toString().includes('PROFIT');
}

/**
 * Adds the profit table to a specific sheet
 */
function addProfitTableToSheet(sheet, year) {
  const expensesColumns = CATEGORIES.length + 2;
  const profitStartColumn = expensesColumns + 2;
  const profitStartColLetter = String.fromCharCode(65 + profitStartColumn - 1);
  
  const startRow = 1;
  
  sheet.getRange(startRow, profitStartColumn).setValue('MONTHLY PROFIT - ' + year)
    .setFontWeight('bold').setFontSize(14).setBackground('#E8F0FE');
  
  const profitHeaders = ['Month', 'Square', 'Other', 'TOTAL'];
  sheet.getRange(startRow + 1, profitStartColumn, 1, profitHeaders.length)
    .setValues([profitHeaders])
    .setFontWeight('bold')
    .setBackground('#9C27B0')
    .setFontColor('white');
  
  for (let i = 0; i < MONTHS.length; i++) {
    const row = startRow + 2 + i;
    sheet.getRange(row, profitStartColumn).setValue(MONTHS[i]);
    
    sheet.getRange(row, profitStartColumn + 1).setValue(0).setNumberFormat('$#,##0.00');
    sheet.getRange(row, profitStartColumn + 2).setValue(0).setNumberFormat('$#,##0.00');
    
    const totalFormula = `=SUM(${String.fromCharCode(65 + profitStartColumn)}${row}:${String.fromCharCode(65 + profitStartColumn + 1)}${row})`;
    sheet.getRange(row, profitStartColumn + 3).setFormula(totalFormula)
      .setNumberFormat('$#,##0.00').setFontWeight('bold');
  }
  
  const totalRow = startRow + 2 + MONTHS.length;
  sheet.getRange(totalRow, profitStartColumn).setValue('ANNUAL TOTAL').setFontWeight('bold');
  
  for (let j = 0; j < profitHeaders.length - 1; j++) {
    const col = profitStartColumn + j + 1;
    const colLetter = String.fromCharCode(65 + col - 1);
    const formula = `=SUM(${colLetter}${startRow + 2}:${colLetter}${totalRow - 1})`;
    sheet.getRange(totalRow, col).setFormula(formula)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold')
      .setBackground('#FBBC04');
  }
  
  const profitTableRange = sheet.getRange(startRow + 1, profitStartColumn, MONTHS.length + 2, profitHeaders.length);
  profitTableRange.setBorder(true, true, true, true, true, true);
}

/**
 * Adds a net results table that shows monthly and yearly net (profit - expenses)
 * Works on all year-named sheets that have both expenses and profit tables
 */
function addNetResultsTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const yearSheets = [];
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (/^\d{4}$/.test(sheetName)) {
      yearSheets.push(sheet);
    }
  }
  
  if (yearSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No year sheets found. Please initialize the spreadsheet first.');
    return;
  }
  
  let addedCount = 0;
  let skippedExpenses = 0;
  let skippedProfit = 0;
  let skippedAlready = 0;
  let skippedPastCount = 0;
  const currentYear = new Date().getFullYear();

  for (const sheet of yearSheets) {
    const sheetName = sheet.getName();

    if (parseInt(sheetName) < currentYear) {
      skippedPastCount++;
      continue;
    }

    if (!sheetHasExpensesTable(sheet)) {
      skippedExpenses++;
      continue;
    }

    if (!sheetHasProfitTable(sheet)) {
      skippedProfit++;
      continue;
    }

    if (sheetHasNetResultsTable(sheet)) {
      skippedAlready++;
      continue;
    }

    addNetResultsTableToSheet(sheet, sheetName);
    addedCount++;
  }

  let message = '';
  if (addedCount > 0) {
    message += `Net results table added to ${addedCount} sheet(s).\n\n`;
  }
  if (skippedExpenses > 0) {
    message += `${skippedExpenses} sheet(s) missing expenses table.\n`;
  }
  if (skippedProfit > 0) {
    message += `${skippedProfit} sheet(s) missing profit table.\n`;
  }
  if (skippedAlready > 0) {
    message += `${skippedAlready} sheet(s) already had net results table.\n`;
  }
  if (skippedPastCount > 0) {
    message += `${skippedPastCount} past year sheet(s) skipped.`;
  }
  
  SpreadsheetApp.getUi().alert(
    'Net Results Table Update Complete',
    message.trim(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Checks if a sheet has an expenses table
 */
function sheetHasExpensesTable(sheet) {
  const cellValue = sheet.getRange(1, 1).getValue();
  return cellValue && cellValue.toString().includes('SPENDING');
}

/**
 * Checks if a sheet already has a net results table
 */
function sheetHasNetResultsTable(sheet) {
  const expensesColumns = CATEGORIES.length + 2;
  const profitColumns = 4;
  const netStartColumn = expensesColumns + 1 + profitColumns + 1;
  const cellValue = sheet.getRange(1, netStartColumn).getValue();
  return cellValue && cellValue.toString().includes('NET');
}

/**
 * Adds the net results table to a specific sheet
 */
function addNetResultsTableToSheet(sheet, year) {
  const expensesColumns = CATEGORIES.length + 2;
  const profitColumns = 4;
  const profitStartColumn = expensesColumns + 2;
  const netStartColumn = profitStartColumn + profitColumns + 1;
  
  const startRow = 1;
  
  const expensesTotalCol = expensesColumns;
  const profitTotalCol = profitStartColumn + 3;
  const profitTotalColLetter = String.fromCharCode(64 + profitTotalCol);
  const expensesTotalColLetter = String.fromCharCode(64 + expensesTotalCol);
  
  sheet.getRange(startRow, netStartColumn).setValue('MONTHLY NET RESULTS - ' + year)
    .setFontWeight('bold').setFontSize(14).setBackground('#E8F0FE');
  
  const netHeaders = ['Month', 'Profit', 'Expenses', 'NET'];
  sheet.getRange(startRow + 1, netStartColumn, 1, netHeaders.length)
    .setValues([netHeaders])
    .setFontWeight('bold')
    .setBackground('#FF5722')
    .setFontColor('white');
  
  for (let i = 0; i < MONTHS.length; i++) {
    const row = startRow + 2 + i;
    sheet.getRange(row, netStartColumn).setValue(MONTHS[i]);
    
    sheet.getRange(row, netStartColumn + 1).setFormula(`=${profitTotalColLetter}${row}`)
      .setNumberFormat('$#,##0.00');
    
    sheet.getRange(row, netStartColumn + 2).setFormula(`=${expensesTotalColLetter}${row}`)
      .setNumberFormat('$#,##0.00');
    
    const netColLetter = String.fromCharCode(64 + netStartColumn + 3);
    const profitColLetter = String.fromCharCode(64 + netStartColumn + 1);
    const expenseColLetter = String.fromCharCode(64 + netStartColumn + 2);
    sheet.getRange(row, netStartColumn + 3).setFormula(`=${profitColLetter}${row}-${expenseColLetter}${row}`)
      .setNumberFormat('$#,##0.00').setFontWeight('bold');
  }
  
  const totalRow = startRow + 2 + MONTHS.length;
  sheet.getRange(totalRow, netStartColumn).setValue('YEARLY NET').setFontWeight('bold');
  
  sheet.getRange(totalRow, netStartColumn + 1).setFormula(`=${profitTotalColLetter}${totalRow}`)
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setBackground('#FBBC04');
  
  sheet.getRange(totalRow, netStartColumn + 2).setFormula(`=${expensesTotalColLetter}${totalRow}`)
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setBackground('#FBBC04');
  
  const netColLetter = String.fromCharCode(64 + netStartColumn + 3);
  const profitColLetter = String.fromCharCode(64 + netStartColumn + 1);
  const expenseColLetter = String.fromCharCode(64 + netStartColumn + 2);
  sheet.getRange(totalRow, netStartColumn + 3).setFormula(`=${profitColLetter}${totalRow}-${expenseColLetter}${totalRow}`)
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setBackground('#34A853').setFontColor('white');
  
  const netTableRange = sheet.getRange(startRow + 1, netStartColumn, MONTHS.length + 2, netHeaders.length);
  netTableRange.setBorder(true, true, true, true, true, true);
}

/**
 * Shows a dialog to insert a new transaction manually
 */
function insertTransaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = new Date().getFullYear();
  
  let sheet = ss.getSheetByName(year.toString());
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No sheet found for current year. Please initialize the spreadsheet first.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  
  const categoryList = CATEGORIES.join(', ');
  
  const response = ui.prompt(
    'Insert Transaction',
    'Enter transaction details (format: YYYY-MM-DD, Category, Amount, Your Name)\n\n' +
    'Example: 2026-01-15, Retail Shelf, 125.50, Levi\n\n' +
    `Valid categories: ${categoryList}`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const input = response.getResponseText().trim();
  const parts = input.split(',').map(s => s.trim());
  
  if (parts.length < 4) {
    ui.alert('Error', 'Please enter all 4 fields: Date, Category, Amount, Your Name', ui.ButtonSet.OK);
    return;
  }
  
  const dateStr = parts[0];
  const category = parts[1];
  const amount = parseFloat(parts[2]);
  const userName = parts[3];
  
  if (isNaN(amount) || amount < 0) {
    ui.alert('Error', 'Amount must be a valid positive number', ui.ButtonSet.OK);
    return;
  }
  
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) {
    ui.alert('Error', 'Invalid date format. Use YYYY-MM-DD', ui.ButtonSet.OK);
    return;
  }
  
  const month = MONTHS[date.getMonth()];
  const transactionYear = date.getFullYear();
  
  if (transactionYear !== parseInt(year)) {
    const confirm = ui.alert(
      'Year Mismatch',
      `This date is for year ${transactionYear}, but current sheet is for ${year}. Add anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) {
      return;
    }
  }
  
  try {
    addTransaction(sheet, date, month, category, amount, [], userName);
    updateMonthlySummary(sheet, month, category, amount);
    ui.alert('Success', `Transaction added:\n\nDate: ${dateStr}\nCategory: ${category}\nAmount: $${amount.toFixed(2)}\nBy: ${userName}`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to add transaction: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Finds the first data row of the transaction log.
 * Handles both old layout (tx log at row 17, data at 19) and
 * new layout (tx log at row 32, data at 34) by scanning for the title.
 */
function findTransactionLogStartRow(sheet) {
  for (let row = 15; row < 50; row++) {
    const val = sheet.getRange(row, 1).getValue();
    if (val && val.toString().startsWith('TRANSACTION LOG')) {
      return row + 2; // title row + header row + 1 = first data row
    }
  }
  return 19; // fallback: original layout (tx log title at row 17, data at row 19)
}

/**
 * Dynamically finds the starting column of the manual entries log by scanning
 * the title row for a cell containing "MANUAL ENTRIES". Falls back to
 * MANUAL_LOG_START_COL if not found (e.g. during initial creation).
 */
function findManualLogStartCol(sheet) {
  const titleRow = findTransactionLogStartRow(sheet) - 2;
  const lastCol = sheet.getLastColumn();
  for (let col = 2; col <= lastCol; col++) {
    const val = sheet.getRange(titleRow, col).getValue();
    if (val && val.toString().startsWith('MANUAL ENTRIES')) {
      return col;
    }
  }
  return MANUAL_LOG_START_COL; // fallback to constant if log not yet created
}

/**
 * Adds a transaction to the log
 */
function addTransaction(sheet, date, month, category, amount, fileLinks, userName) {
  // Find the next empty row in transaction log (dynamic for old/new layout)
  let lastRow = findTransactionLogStartRow(sheet);
  while (sheet.getRange(lastRow, 1).getValue() !== '') {
    lastRow++;
    if (lastRow > 10000) break; // Safety check
  }
  
  // Format date
  const dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  
  // Create display text for multiple photos
  let photoDisplay = '';
  if (fileLinks && fileLinks.length > 0) {
    if (fileLinks.length === 1) {
      photoDisplay = 'View Receipt';
    } else {
      photoDisplay = `View ${fileLinks.length} Photos`;
    }
  }
  
  // Add the transaction
  sheet.getRange(lastRow, 1, 1, 6).setValues([[
    dateFormatted,
    month,
    category,
    amount,
    photoDisplay,
    userName
  ]]);
  
  // Format amount as currency
  sheet.getRange(lastRow, 4).setNumberFormat('$#,##0.00');
  
  // Make links clickable if provided
  if (fileLinks && fileLinks.length > 0) {
    if (fileLinks.length === 1) {
      // Single photo - direct link
      sheet.getRange(lastRow, 5).setFormula(`=HYPERLINK("${fileLinks[0]}", "View Receipt")`);
    } else {
      // Multiple photos - create dropdown or comma-separated links
      const linkText = fileLinks.map((link, index) => 
        `=HYPERLINK("${link}", "Photo ${index + 1}")`
      ).join(', ');
      sheet.getRange(lastRow, 5).setFormula(linkText);
    }
  }
}

/**
 * Updates the monthly summary table
 */
function updateMonthlySummary(sheet, month, category, amount) {
  // Summary table starts at row 1, header at row 2, first month at row 3
  const firstMonthRow = 3;
  const monthIndex = MONTHS.indexOf(month);
  
  if (monthIndex === -1) {
    throw new Error('Invalid month: ' + month);
  }
  
  // Find category column dynamically
  const headerRow = 2;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const categoryIndex = headers.indexOf(category);
  
  if (categoryIndex === -1) {
    throw new Error('Category not found in sheet: ' + category);
  }
  
  const row = firstMonthRow + monthIndex;
  const col = categoryIndex + 1; // Convert to 1-based column
  
  // Get current value and add new amount
  const currentValue = sheet.getRange(row, col).getValue() || 0;
  sheet.getRange(row, col).setValue(currentValue + parseFloat(amount));
}

/**
 * Reads unapplied rows from the manual entries log (cols I-M), confirms with
 * the user, applies each to the monthly summary, and marks the rows green.
 */
function applyManualEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (!/^\d{4}$/.test(sheet.getName())) {
    ui.alert('Wrong Sheet', 'Navigate to a year sheet (e.g. "2026") first.', ui.ButtonSet.OK);
    return;
  }

  const firstDataRow = findTransactionLogStartRow(sheet);
  const manualStartCol = findManualLogStartCol(sheet);
  const sheetYear = sheet.getName();

  // Scan all manual log columns to find the last row with any data,
  // since the date column might be empty while other columns have values.
  const scanRows = Math.min(2000, sheet.getMaxRows() - firstDataRow + 1);
  const manualColData = sheet.getRange(firstDataRow, manualStartCol, scanRows, 5).getValues();
  let lastManualRow = firstDataRow - 1;
  for (let i = manualColData.length - 1; i >= 0; i--) {
    if (manualColData[i].some(cell => cell !== '')) {
      lastManualRow = firstDataRow + i;
      break;
    }
  }

  const pending = [];

  for (let row = firstDataRow; row <= lastManualRow; row++) {
    const dateVal = sheet.getRange(row, manualStartCol).getValue();
    const category = sheet.getRange(row, manualStartCol + 1).getValue();
    const amount = sheet.getRange(row, manualStartCol + 2).getValue();
    const note = sheet.getRange(row, manualStartCol + 3).getValue();

    if (!dateVal || !category || amount === '' || amount === null) continue;

    // Skip already-applied rows (marked green)
    const bg = sheet.getRange(row, manualStartCol).getBackground();
    if (bg === '#b7e1cd') continue;

    const date = new Date(dateVal);
    if (isNaN(date.getTime())) continue;

    const entryYear = date.getFullYear().toString();
    if (entryYear !== sheetYear) {
      Logger.log(`Manual entry row ${row} has date from ${entryYear}, skipped on ${sheetYear} sheet`);
      continue;
    }

    const numericAmount = parseFloat(amount);
    if (isNaN(numericAmount)) continue;

    const month = MONTHS[date.getMonth()];
    pending.push({ row, date, month, category, amount: numericAmount, note });
  }

  if (pending.length === 0) {
    ui.alert('No Entries', 'No unapplied manual entries found in the log.', ui.ButtonSet.OK);
    return;
  }

  const lines = pending.map(p => {
    const dateStr = Utilities.formatDate(p.date, Session.getScriptTimeZone(), 'MM/dd');
    const sign = p.amount >= 0 ? '+' : '';
    const noteStr = p.note ? `  (${p.note})` : '';
    return `${dateStr} / ${p.category}: ${sign}${p.amount.toFixed(2)}${noteStr}`;
  });

  const confirm = ui.alert(
    'Apply Manual Entries',
    `Apply ${pending.length} entr(ies) to the monthly summary?\n\n${lines.join('\n')}`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const mainHeaders = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (const p of pending) {
    if (!mainHeaders.includes(p.category)) {
      Logger.log(`Category "${p.category}" not in summary header, skipping row ${p.row}`);
      continue;
    }
    updateMonthlySummary(sheet, p.month, p.category, p.amount);
    // Mark row as applied with light green background
    sheet.getRange(p.row, manualStartCol, 1, 5).setBackground('#b7e1cd');
  }

  ui.alert('Done', `${pending.length} new item(s) applied to the monthly summary.`, ui.ButtonSet.OK);
}

/**
 * Rebuilds the monthly summary table headers and adjustments table headers
 * from the current CATEGORIES array, preserving existing cell values for
 * categories that are still present.
 */
function regenerateSheetCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (!/^\d{4}$/.test(sheet.getName())) {
    ui.alert('Wrong Sheet', 'Navigate to a year sheet (e.g. "2026") first.', ui.ButtonSet.OK);
    return;
  }

  // Read current sheet categories from header row — only up to the TOTAL column
  // to avoid picking up profit/net table headers that also live in row 2.
  const currentHeaders = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const totalColIdx = currentHeaders.indexOf('TOTAL'); // 0-based
  const expenseHeaders = totalColIdx !== -1 ? currentHeaders.slice(0, totalColIdx) : currentHeaders;
  const sheetCategories = expenseHeaders.filter(h => h && h !== 'Month');
  const totalCols = sheet.getLastColumn();

  // Capture main summary data keyed by [month][category]
  const savedData = {};
  for (let i = 0; i < MONTHS.length; i++) {
    const month = MONTHS[i];
    savedData[month] = {};
    for (const cat of sheetCategories) {
      const colIdx = currentHeaders.indexOf(cat);
      if (colIdx !== -1) {
        savedData[month][cat] = sheet.getRange(3 + i, colIdx + 1).getValue() || 0;
      }
    }
  }

  // Compute diff
  const surviving = CATEGORIES.filter(c => sheetCategories.includes(c));
  const dropped = sheetCategories.filter(c => !CATEGORIES.includes(c));
  const added = CATEGORIES.filter(c => !sheetCategories.includes(c));

  let confirmMsg = 'Reorganize the summary table to match the CATEGORIES array.\n\n';
  if (surviving.length) confirmMsg += `Keeping (data preserved): ${surviving.join(', ')}\n`;
  if (added.length) confirmMsg += `Adding (fresh $0.00): ${added.join(', ')}\n`;
  if (dropped.length) confirmMsg += `Removing (data will be lost): ${dropped.join(', ')}\n`;
  confirmMsg += '\nContinue?';

  const confirm = ui.alert('Regenerate Categories', confirmMsg, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  // Clear main summary area (rows 2-15, all cols up to current last col)
  sheet.getRange(2, 1, 14, totalCols).clear();

  // Rebuild header row 2
  const newHeaders = ['Month', ...CATEGORIES, 'TOTAL'];
  sheet.getRange(2, 1, 1, newHeaders.length)
    .setValues([newHeaders])
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('white');

  // Rebuild month rows 3-14
  for (let i = 0; i < MONTHS.length; i++) {
    const row = 3 + i;
    const month = MONTHS[i];
    sheet.getRange(row, 1).setValue(month);

    for (let j = 0; j < CATEGORIES.length; j++) {
      const cat = CATEGORIES[j];
      const value = surviving.includes(cat) ? (savedData[month][cat] || 0) : 0;
      sheet.getRange(row, 2 + j).setValue(value).setNumberFormat('$#,##0.00');
    }

    // TOTAL formula
    const totalCol = 2 + CATEGORIES.length;
    const lastCatLetter = columnToLetter(totalCol - 1);
    sheet.getRange(row, totalCol)
      .setFormula(`=SUM(B${row}:${lastCatLetter}${row})`)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold');
  }

  // Rebuild annual total row (row 15)
  sheet.getRange(15, 1).setValue('ANNUAL TOTAL').setFontWeight('bold');
  for (let j = 0; j <= CATEGORIES.length; j++) {
    const col = 2 + j;
    const colLetter = columnToLetter(col);
    sheet.getRange(15, col)
      .setFormula(`=SUM(${colLetter}3:${colLetter}14)`)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold')
      .setBackground('#FBBC04');
  }

  // Clear orphaned columns between end of new expenses table and wherever profit table starts
  // (or to totalCols if no profit table). This removes stale expense columns without
  // touching the profit table data we're about to relocate.
  const newTableWidth = newHeaders.length;

  // Find profit table's current column by scanning row 1 for "PROFIT"
  let oldProfitStartCol = -1;
  const row1Values = sheet.getRange(1, 1, 1, totalCols).getValues()[0];
  for (let c = 0; c < row1Values.length; c++) {
    if (row1Values[c] && row1Values[c].toString().includes('PROFIT')) {
      oldProfitStartCol = c + 1; // 1-based
      break;
    }
  }

  // Save profit table data (Square + Other, rows 3-14) before clearing
  const savedProfitData = []; // array of [square, other] per month
  let hasProfitTable = oldProfitStartCol !== -1;
  if (hasProfitTable) {
    for (let i = 0; i < MONTHS.length; i++) {
      const row = 3 + i;
      const square = sheet.getRange(row, oldProfitStartCol + 1).getValue();
      const other = sheet.getRange(row, oldProfitStartCol + 2).getValue();
      savedProfitData.push([square, other]);
    }
  }

  // Find net results table's current column (for clearing)
  let oldNetStartCol = -1;
  for (let c = 0; c < row1Values.length; c++) {
    if (row1Values[c] && row1Values[c].toString().includes('NET')) {
      oldNetStartCol = c + 1;
      break;
    }
  }

  // Clear everything to the right of the new expenses table through the old tables
  const clearFrom = newTableWidth + 1;
  const clearWidth = totalCols - newTableWidth;
  if (clearWidth > 0) {
    sheet.getRange(1, clearFrom, 15, clearWidth).clear();
  }

  // Restore borders on main summary table
  sheet.getRange(2, 1, 14, newTableWidth).setBorder(true, true, true, true, true, true);

  // Re-place profit table at its new correct column (based on new CATEGORIES.length)
  if (hasProfitTable) {
    const year = sheet.getName();
    addProfitTableToSheet(sheet, year);

    // Restore the saved Square + Other values (addProfitTableToSheet writes 0s)
    const newProfitStartCol = CATEGORIES.length + 4; // same formula as addProfitTableToSheet
    for (let i = 0; i < MONTHS.length; i++) {
      const row = 3 + i;
      sheet.getRange(row, newProfitStartCol + 1).setValue(savedProfitData[i][0]).setNumberFormat('$#,##0.00');
      sheet.getRange(row, newProfitStartCol + 2).setValue(savedProfitData[i][1]).setNumberFormat('$#,##0.00');
    }

    // Re-place net results table if it existed
    if (oldNetStartCol !== -1) {
      addNetResultsTableToSheet(sheet, year);
    }
  }

  // Refresh the category dropdown in the manual entries log
  refreshManualLogValidation(sheet);

  ui.alert('Done', `Categories regenerated.\n\nKept: ${surviving.join(', ') || 'none'}\nAdded: ${added.join(', ') || 'none'}\nRemoved: ${dropped.join(', ') || 'none'}`, ui.ButtonSet.OK);
}

/**
 * Migration: adds the manual entries log to an existing year sheet that was
 * created before this feature was added. Safe to run on any year sheet.
 */
function addManualEntriesLogToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (!/^\d{4}$/.test(sheet.getName())) {
    ui.alert('Wrong Sheet', 'Navigate to a year sheet (e.g. "2026") first.', ui.ButtonSet.OK);
    return;
  }

  // Check if manual log already exists by scanning the title row
  const titleRow = findTransactionLogStartRow(sheet) - 2;
  const lastCol = Math.max(sheet.getLastColumn(), MANUAL_LOG_START_COL);
  const titleRowValues = sheet.getRange(titleRow, 1, 1, lastCol).getValues()[0];
  const alreadyExists = titleRowValues.some(v => v && v.toString().startsWith('MANUAL ENTRIES'));
  if (alreadyExists) {
    ui.alert('Already Exists', 'This sheet already has a manual entries log.', ui.ButtonSet.OK);
    return;
  }

  createManualEntriesLog(sheet, sheet.getName());
  ui.alert('Done', 'Manual entries log added to the right of the transaction log.', ui.ButtonSet.OK);
}

/**
 * INITIALIZATION - Run this once to set up the spreadsheet
 * This creates the current year sheet with proper structure
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = new Date().getFullYear();
  
  // Check if sheet already exists
  let sheet = ss.getSheetByName(year.toString());
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Already Exists',
      `A sheet for ${year} already exists. Do you want to recreate it? (This will delete existing data)`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
      sheet = createYearSheet(ss, year);
      Logger.log(`Sheet for ${year} recreated successfully`);
    } else {
      Logger.log('Initialization cancelled');
      return;
    }
  } else {
    sheet = createYearSheet(ss, year);
    Logger.log(`Sheet for ${year} created successfully`);
  }
  
  SpreadsheetApp.getUi().alert(
    'Initialization Complete!',
    `The ${year} sheet has been created with:\n\n` +
    '✓ Transaction log header\n' +
    '✓ Monthly summary table\n' +
    '✓ All categories configured\n\n' +
    'You can now start adding receipts or run populateTestData() to add sample data.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Helper function to manually create current year sheet (legacy support)
 */
function initializeCurrentYear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = new Date().getFullYear();
  createYearSheet(ss, year);
}

/**
 * POPULATE TEST DATA - Adds realistic fake data to see how the sheet looks
 * Run this after initializing the spreadsheet to see it in action
 */
function populateTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = new Date().getFullYear();

  // Make sure the sheet exists
  let sheet = ss.getSheetByName(year.toString());
  if (!sheet) {
    sheet = createYearSheet(ss, year);
    Logger.log('Created sheet for ' + year);
  }

  // Derive users from API_KEY_USERS in Config.js
  const users = Object.values(API_KEY_USERS);

  // Build test entries dynamically from the current CATEGORIES array.
  // Spreads entries across all 12 months: every category gets at least one
  // entry per month, cycling through users.
  const testData = [];
  const months = [
    { num: '01', days: [3, 8, 15, 22, 28] },
    { num: '02', days: [3, 8, 14, 20, 27] },
    { num: '03', days: [2, 7, 12, 18, 25] },
    { num: '04', days: [1, 6, 11, 17, 24] },
    { num: '05', days: [2, 7, 13, 19, 27] },
    { num: '06', days: [3, 8, 14, 20, 27] },
    { num: '07', days: [1, 7, 12, 18, 25] },
    { num: '08', days: [2, 8, 14, 20, 28] },
    { num: '09', days: [3, 9, 15, 21, 28] },
    { num: '10', days: [1, 6, 12, 18, 25] },
    { num: '11', days: [2, 7, 13, 19, 26] },
    { num: '12', days: [1, 6, 12, 18, 24] },
  ];

  // Simple deterministic amount generator — varies by category index and month
  // so the summary table has visually distinct, non-uniform values.
  function testAmount(catIndex, monthIndex, dayIndex) {
    const base = 50 + (catIndex * 73 + monthIndex * 37 + dayIndex * 19) % 450;
    return Math.round(base * 100) / 100;
  }

  let entryIndex = 0;
  months.forEach((month, monthIndex) => {
    CATEGORIES.forEach((category, catIndex) => {
      // Use a subset of days so we don't flood the log — one entry per category per month
      const day = month.days[catIndex % month.days.length];
      const date = `${year}-${month.num}-${String(day).padStart(2, '0')}`;
      const user = users[entryIndex % users.length];
      const amount = testAmount(catIndex, monthIndex, catIndex);
      testData.push({ date, category, amount, user });
      entryIndex++;
    });
  });

  // Add all test receipts
  Logger.log('Adding test data...');
  let count = 0;

  testData.forEach(receipt => {
    try {
      const fakeLink = `https://drive.google.com/file/d/TEST_${receipt.date}_${receipt.category.replace(/\W+/g, '_')}`;
      addReceipt(receipt.date, receipt.category, receipt.amount, [fakeLink], receipt.user);
      count++;
    } catch (error) {
      Logger.log(`Error adding receipt for ${receipt.date} / ${receipt.category}: ${error.toString()}`);
    }
  });

  Logger.log(`Successfully added ${count} test receipts`);

  SpreadsheetApp.getUi().alert(
    'Test Data Added!',
    `Added ${count} test receipts across 12 months.\n\n` +
    `Categories used: ${CATEGORIES.join(', ')}\n` +
    `Users: ${users.join(', ')}\n\n` +
    'Check the transaction log and monthly summary table!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Test function - simulates a complete API call with multiple photo upload
 * Demonstrates uploading multiple photos for a single transaction
 */
function testAddReceiptWithMultiplePhotos() {
  // Create simple test images (1x1 red pixel JPEG)
  const testImageBase64 = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwA2Af/Z';
  
  try {
    const testCategory = CATEGORIES[0];
    const fileLinks = uploadPhotosToDrive(
      [testImageBase64, testImageBase64, testImageBase64], // 3 photos
      '2025-10-08',
      testCategory,
      'Test User'
    );
    Logger.log('Photos uploaded successfully: ' + fileLinks.length + ' files');

    const result = addReceipt(
      '2025-10-08',
      testCategory,
      125.75,
      fileLinks,
      'Test User'
    );
    Logger.log(result);
  } catch (error) {
    Logger.log('Test failed: ' + error.toString());
  }
}

/**
 * Test function - demonstrates category alias usage
 * Shows that simplified inputs like "products", "repair", "fuel" work
 */
function testCategoryAliases() {
  Logger.log('Testing category aliases...\n');

  // Derive test cases directly from CATEGORY_ALIASES so this stays in sync with Config.js
  Object.entries(CATEGORY_ALIASES).forEach(([input, expected]) => {
    const resolved = CATEGORY_ALIASES[input.toLowerCase()];
    const status = resolved === expected ? '✓' : '✗';
    Logger.log(`${status} "${input}" → "${resolved}" (expected: "${expected}")`);
  });

  Logger.log('\nRun showValidCategories() to see all valid inputs.');
}

/**
 * Test function - demonstrates dynamic category creation
 * Shows that new categories are automatically added to the sheet
 */
function testDynamicCategories() {
  Logger.log('Testing dynamic category creation...\n');
  
  const year = new Date().getFullYear();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(year.toString());
  
  if (!sheet) {
    sheet = createYearSheet(ss, year);
    Logger.log(`Created sheet for ${year}`);
  }
  
  const testCategories = [
    'Marketing',
    'Software Licenses', 
    'Travel Expenses',
    'Equipment Purchase',
    'Professional Services'
  ];
  
  Logger.log('Adding receipts with new categories...');
  
  testCategories.forEach((category, index) => {
    try {
      const result = addReceipt(
        `${year}-10-${String(index + 1).padStart(2, '0')}`,
        category,
        100 + (index * 25),
        [`https://drive.google.com/file/d/test_${category.replace(/\s+/g, '_')}`],
        'Test User'
      );
      Logger.log(`✓ Added "${category}": ${result}`);
    } catch (error) {
      Logger.log(`✗ Failed to add "${category}": ${error.toString()}`);
    }
  });
  
  Logger.log('\n======================');
  Logger.log('✓ DYNAMIC CATEGORY TEST COMPLETE');
  Logger.log('======================');
  Logger.log('Check the spreadsheet to see the new categories added to the monthly summary table.');
}

/**
 * Get Receipts folder ID helper
 */
function createReceiptsFolderAndGetId() {
  const folder = DriveApp.createFolder('Receipts');
  Logger.log('Receipts folder created!');
  Logger.log('Folder ID: ' + folder.getId());
  Logger.log('Update RECEIPTS_FOLDER_ID in your script with this ID');
  return folder.getId();
}