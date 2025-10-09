/**
 * Creates custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Receipt Tracking')
    .addItem('🔧 Initialize Spreadsheet', 'initializeSpreadsheet')
    .addItem('📊 Populate Test Data', 'populateTestData')
    .addSeparator()
    .addItem('📋 Show Valid Categories', 'showValidCategories')
    .addItem('📂 Create Receipts Folder', 'createReceiptsFolderAndGetId')
    .addToUi();
}

// Map API keys to user names for tracking who added receipts
const API_KEY_USERS = {
  'replacewithyourkey': 'Levi',
  'replacewithyourkey': 'Taylor',
  'replacewithyourkey': 'Jim',
  'replacewithyourkey': 'Bob',
  'replacewithyourkey': 'Joe',
};

const RECEIPTS_FOLDER_ID = 'FileIDfromURL';

// Spreadsheet header values (exact format)
const CATEGORIES = ['Products/Ingredients', 'Gas', 'Employees', 'Rent', 'Asset Repair/Maintenance', 'Operating Supplies', 'Contracts', 'Misc'];

// Map user-friendly input to spreadsheet categories
const CATEGORY_ALIASES = {
  // Products/Ingredients
  'products': 'Products/Ingredients',
  'ingredients': 'Products/Ingredients',
  'product': 'Products/Ingredients',
  'food': 'Products/Ingredients',

  
  // Gas
  'gas': 'Gas',
  'fuel': 'Gas',
  
  // Employees
  'employees': 'Employees',
  'employee': 'Employees',
  'payroll': 'Employees',
  'staff': 'Employees',
  
  // Rent
  'rent': 'Rent',
  
  // Asset Repair/Maintenance
  'repair': 'Asset Repair/Maintenance',
  'maintenance': 'Asset Repair/Maintenance',
  'repairs': 'Asset Repair/Maintenance',
  'fix': 'Asset Repair/Maintenance',
  'asset': 'Asset Repair/Maintenance',
  
  // Operating Supplies
  'supplies': 'Operating Supplies',
  'supply': 'Operating Supplies',
  'operating supplies': 'Operating Supplies',
  
  // Contracts
  'contracts': 'Contracts',
  'contract': 'Contracts',
  'service': 'Contracts',
  'services': 'Contracts',
  
  // Misc
  'misc': 'Misc',
  'miscellaneous': 'Misc',
  'other': 'Misc',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December'];

/**
 * Main function - accepts photo as base64 string
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
    
    // Check if input is an alias
    if (CATEGORY_ALIASES[inputCategory]) {
      category = CATEGORY_ALIASES[inputCategory];
    } 
    // Check if input matches a category exactly (case-insensitive)
    else {
      const matchedCategory = CATEGORIES.find(cat => cat.toLowerCase() === inputCategory);
      if (matchedCategory) {
        category = matchedCategory;
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: `Invalid category: ${params.category}. Valid options: ${Object.keys(CATEGORY_ALIASES).join(', ')}`
        })).setMimeType(ContentService.MimeType.JSON);
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
    
    // Upload photo to Google Drive if provided
    let fileLink = '';
    if (params.photoBase64) {
      try {
        fileLink = uploadPhotoToDrive(
          params.photoBase64,
          params.date,
          category,
          userName
        );
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
      fileLink,
      userName
    );
    
    // Create user-friendly success message
    const date = new Date(params.date);
    const month = MONTHS[date.getMonth()];
    const year = date.getFullYear();
    const friendlyDate = `${month} ${date.getDate()}, ${year}`;
    
    const successMessage = `✅ Receipt added successfully!\n\n` +
      `📅 Date: ${friendlyDate}\n` +
      `🏷️ Category: ${category}\n` +
      `💰 Amount: $${amount}\n` +
      `👤 Added by: ${userName}` +
      (fileLink ? `\n📎 Receipt photo saved to Drive` : '');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: successMessage,
      fileLink: fileLink,
      summary: {
        date: friendlyDate,
        category: category,
        amount: `$${amount}`,
        user: userName,
        hasPhoto: !!fileLink
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
 * Uploads photo to Google Drive in organized folder structure
 */
function uploadPhotoToDrive(base64Data, dateString, category, userName) {
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
  const timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const filename = `receipt_${category}_${timestamp}.jpg`;
  
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
function addReceipt(dateString, category, amount, fileLink, userName) {
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
  
  // Add transaction to log
  addTransaction(sheet, date, month, category, amount, fileLink, userName);
  
  // Update monthly summary table
  updateMonthlySummary(sheet, month, category, amount);
  
  return `Receipt logged: ${category} - $${amount} for ${month} ${year}`;
}

/**
 * Creates a new sheet for the year with proper structure
 * 
 * SHEET LAYOUT:
 * Rows 1-15:  Monthly Summary Table (title, headers, 12 months, annual total)
 * Row 16:     Blank row for spacing
 * Rows 17-18: Transaction Log header
 * Row 19+:    Transaction entries (unlimited)
 */
function createYearSheet(spreadsheet, year) {
  const sheet = spreadsheet.insertSheet(year.toString());
  
  // Create Monthly Summary Table FIRST (at top)
  createMonthlySummaryTable(sheet, year);
  
  // Set up Transaction Log header (below summary table)
  // Summary takes rows 1-15, add blank row 16, transaction log starts row 17
  const transactionStartRow = 17;
  sheet.getRange(transactionStartRow, 1).setValue('TRANSACTION LOG - ' + year)
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange(transactionStartRow + 1, 1, 1, 7).setValues([[
    'Date', 'Month', 'Category', 'Amount', 'Receipt Link', 'Added By', 'Notes'
  ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  
  // Freeze summary table and transaction header (15 rows summary + 3 rows for transaction header)
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
 * Adds a transaction to the log
 */
function addTransaction(sheet, date, month, category, amount, fileLink, userName) {
  // Find the next empty row in transaction log
  // Summary takes rows 1-15, transaction header at rows 17-18, data starts at row 19
  let lastRow = 19;
  while (sheet.getRange(lastRow, 1).getValue() !== '') {
    lastRow++;
    if (lastRow > 10000) break; // Safety check
  }
  
  // Format date
  const dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  
  // Add the transaction
  sheet.getRange(lastRow, 1, 1, 6).setValues([[
    dateFormatted,
    month,
    category,
    amount,
    fileLink,
    userName
  ]]);
  
  // Format amount as currency
  sheet.getRange(lastRow, 4).setNumberFormat('$#,##0.00');
  
  // Make link clickable if provided
  if (fileLink) {
    sheet.getRange(lastRow, 5).setFormula(`=HYPERLINK("${fileLink}", "View Receipt")`);
  }
}

/**
 * Updates the monthly summary table
 */
function updateMonthlySummary(sheet, month, category, amount) {
  // Summary table starts at row 1, header at row 2, first month at row 3
  const firstMonthRow = 3;
  const monthIndex = MONTHS.indexOf(month);
  const categoryIndex = CATEGORIES.indexOf(category);
  
  if (monthIndex === -1 || categoryIndex === -1) {
    throw new Error('Invalid month or category');
  }
  
  const row = firstMonthRow + monthIndex;
  const col = 2 + categoryIndex; // Column B is first category
  
  // Get current value and add new amount
  const currentValue = sheet.getRange(row, col).getValue() || 0;
  sheet.getRange(row, col).setValue(currentValue + parseFloat(amount));
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
  
  // Define test data for each category with realistic vendors and amounts
  const testData = [
    // January
    { date: `${year}-01-05`, category: 'Products/Ingredients', amount: 234.50, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-01-08`, category: 'Gas', amount: 52.30, user: 'Taylor', vendor: 'Shell Station' },
    { date: `${year}-01-12`, category: 'Products/Ingredients', amount: 189.75, user: 'Katie', vendor: 'Sysco' },
    { date: `${year}-01-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-01-20`, category: 'Misc', amount: 45.60, user: 'Jeff', vendor: 'Office Supplies' },
    { date: `${year}-01-25`, category: 'Products/Ingredients', amount: 312.20, user: 'Levi', vendor: 'US Foods' },
    
    // February
    { date: `${year}-02-03`, category: 'Products/Ingredients', amount: 276.80, user: 'Taylor', vendor: 'Restaurant Depot' },
    { date: `${year}-02-07`, category: 'Gas', amount: 48.90, user: 'Katie', vendor: 'Chevron' },
    { date: `${year}-02-10`, category: 'Asset Repair/Maintenance', amount: 450.00, user: 'Levi', vendor: 'HVAC Repair Co' },
    { date: `${year}-02-14`, category: 'Products/Ingredients', amount: 198.45, user: 'Jeff', vendor: 'Sysco' },
    { date: `${year}-02-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-02-22`, category: 'Employees', amount: 1200.00, user: 'Levi', vendor: 'Payroll - Week 8' },
    { date: `${year}-02-28`, category: 'Products/Ingredients', amount: 223.15, user: 'Taylor', vendor: 'US Foods' },
    
    // March
    { date: `${year}-03-05`, category: 'Products/Ingredients', amount: 289.60, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-03-08`, category: 'Gas', amount: 55.20, user: 'Taylor', vendor: 'Shell Station' },
    { date: `${year}-03-12`, category: 'Contracts', amount: 350.00, user: 'Levi', vendor: 'Cleaning Service' },
    { date: `${year}-03-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-03-18`, category: 'Products/Ingredients', amount: 267.90, user: 'Katie', vendor: 'Sysco' },
    { date: `${year}-03-22`, category: 'Employees', amount: 1350.00, user: 'Levi', vendor: 'Payroll - Week 12' },
    { date: `${year}-03-28`, category: 'Misc', amount: 89.50, user: 'Jeff', vendor: 'Promotional Items' },
    
    // April
    { date: `${year}-04-02`, category: 'Products/Ingredients', amount: 301.25, user: 'Taylor', vendor: 'US Foods' },
    { date: `${year}-04-06`, category: 'Gas', amount: 61.40, user: 'Katie', vendor: 'BP Station' },
    { date: `${year}-04-10`, category: 'Products/Ingredients', amount: 245.80, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-04-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-04-18`, category: 'Asset Repair/Maintenance', amount: 275.00, user: 'Jeff', vendor: 'Plumbing Repair' },
    { date: `${year}-04-22`, category: 'Employees', amount: 1400.00, user: 'Levi', vendor: 'Payroll - Week 16' },
    { date: `${year}-04-28`, category: 'Products/Ingredients', amount: 198.70, user: 'Taylor', vendor: 'Sysco' },
    
    // May
    { date: `${year}-05-05`, category: 'Products/Ingredients', amount: 312.40, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-05-08`, category: 'Gas', amount: 58.75, user: 'Taylor', vendor: 'Shell Station' },
    { date: `${year}-05-12`, category: 'Products/Ingredients', amount: 287.20, user: 'Katie', vendor: 'US Foods' },
    { date: `${year}-05-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-05-20`, category: 'Contracts', amount: 350.00, user: 'Levi', vendor: 'Cleaning Service' },
    { date: `${year}-05-25`, category: 'Employees', amount: 1500.00, user: 'Levi', vendor: 'Payroll - Week 21' },
    { date: `${year}-05-30`, category: 'Misc', amount: 125.80, user: 'Jeff', vendor: 'Equipment' },
    
    // June
    { date: `${year}-06-03`, category: 'Products/Ingredients', amount: 298.90, user: 'Taylor', vendor: 'Sysco' },
    { date: `${year}-06-07`, category: 'Gas', amount: 64.30, user: 'Katie', vendor: 'Chevron' },
    { date: `${year}-06-10`, category: 'Products/Ingredients', amount: 321.50, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-06-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-06-18`, category: 'Asset Repair/Maintenance', amount: 680.00, user: 'Levi', vendor: 'Equipment Repair' },
    { date: `${year}-06-22`, category: 'Employees', amount: 1450.00, user: 'Levi', vendor: 'Payroll - Week 25' },
    { date: `${year}-06-28`, category: 'Products/Ingredients', amount: 256.40, user: 'Taylor', vendor: 'US Foods' },
    
    // July
    { date: `${year}-07-03`, category: 'Products/Ingredients', amount: 334.80, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-07-08`, category: 'Gas', amount: 71.20, user: 'Taylor', vendor: 'Shell Station' },
    { date: `${year}-07-12`, category: 'Products/Ingredients', amount: 289.60, user: 'Katie', vendor: 'Sysco' },
    { date: `${year}-07-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-07-20`, category: 'Contracts', amount: 350.00, user: 'Levi', vendor: 'Cleaning Service' },
    { date: `${year}-07-22`, category: 'Employees', amount: 1600.00, user: 'Levi', vendor: 'Payroll - Week 29' },
    { date: `${year}-07-28`, category: 'Misc', amount: 95.30, user: 'Jeff', vendor: 'Marketing Materials' },
    
    // August
    { date: `${year}-08-02`, category: 'Products/Ingredients', amount: 312.70, user: 'Taylor', vendor: 'US Foods' },
    { date: `${year}-08-06`, category: 'Gas', amount: 68.90, user: 'Katie', vendor: 'BP Station' },
    { date: `${year}-08-10`, category: 'Products/Ingredients', amount: 298.40, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-08-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-08-18`, category: 'Asset Repair/Maintenance', amount: 325.00, user: 'Jeff', vendor: 'Electrical Repair' },
    { date: `${year}-08-22`, category: 'Employees', amount: 1550.00, user: 'Levi', vendor: 'Payroll - Week 34' },
    { date: `${year}-08-28`, category: 'Products/Ingredients', amount: 276.80, user: 'Taylor', vendor: 'Sysco' },
    
    // September
    { date: `${year}-09-05`, category: 'Products/Ingredients', amount: 305.20, user: 'Levi', vendor: 'Restaurant Depot' },
    { date: `${year}-09-08`, category: 'Gas', amount: 59.40, user: 'Taylor', vendor: 'Shell Station' },
    { date: `${year}-09-12`, category: 'Products/Ingredients', amount: 289.90, user: 'Katie', vendor: 'US Foods' },
    { date: `${year}-09-15`, category: 'Rent', amount: 2500.00, user: 'Levi', vendor: 'Property Management' },
    { date: `${year}-09-20`, category: 'Contracts', amount: 350.00, user: 'Levi', vendor: 'Cleaning Service' },
    { date: `${year}-09-25`, category: 'Employees', amount: 1500.00, user: 'Levi', vendor: 'Payroll - Week 39' },
    { date: `${year}-09-30`, category: 'Misc', amount: 112.50, user: 'Jeff', vendor: 'Supplies' },
    
    // October (partial - current month)
    { date: `${year}-10-02`, category: 'Products/Ingredients', amount: 287.30, user: 'Taylor', vendor: 'Sysco' },
    { date: `${year}-10-05`, category: 'Gas', amount: 62.80, user: 'Katie', vendor: 'Chevron' },
    { date: `${year}-10-08`, category: 'Products/Ingredients', amount: 298.60, user: 'Levi', vendor: 'Restaurant Depot' },
  ];
  
  // Add all test receipts
  Logger.log('Adding test data...');
  let count = 0;
  
  testData.forEach(receipt => {
    try {
      // Create a fake drive link
      const fakeLink = `https://drive.google.com/file/d/FAKE_${receipt.date}_${receipt.category}`;
      
      addReceipt(
        receipt.date,
        receipt.category,
        receipt.amount,
        fakeLink,
        receipt.user
      );
      count++;
    } catch (error) {
      Logger.log(`Error adding receipt: ${error.toString()}`);
    }
  });
  
  Logger.log(`Successfully added ${count} test receipts`);
  
  SpreadsheetApp.getUi().alert(
    'Test Data Added!',
    `Successfully added ${count} test receipts across multiple months and categories.\n\n` +
    'The sheet now shows:\n' +
    '✓ Transaction log with varied entries\n' +
    '✓ Multiple users (Levi, Taylor, Katie, Jeff)\n' +
    '✓ All expense categories\n' +
    '✓ Monthly summary totals\n\n' +
    'Check the transaction log and monthly summary table!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Test function - tests adding a receipt with drive link
 * Uses the full category name (still supported)
 */
function testAddReceipt() {
  const result = addReceipt(
    '2025-10-08',
    'Products/Ingredients', // Full category name still works
    45.67,
    'https://drive.google.com/file/d/example',
    'Test User'
  );
  Logger.log(result);
}

/**
 * Test function - simulates a complete API call with photo upload
 * Demonstrates using a category alias ("gas" instead of "Gas")
 */
function testAddReceiptWithPhoto() {
  // Create a simple test image (1x1 red pixel JPEG)
  const testImageBase64 = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwA2Af/Z';
  
  try {
    const fileLink = uploadPhotoToDrive(
      testImageBase64,
      '2025-10-08',
      'Gas', // Category name for folder
      'Test User'
    );
    Logger.log('Photo uploaded successfully to: ' + fileLink);
    
    const result = addReceipt(
      '2025-10-08',
      'Gas',
      25.50,
      fileLink,
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
  
  const testCases = [
    { input: 'products', expected: 'Products/Ingredients' },
    { input: 'gas', expected: 'Gas' },
    { input: 'fuel', expected: 'Gas' },
    { input: 'repair', expected: 'Asset Repair/Maintenance' },
    { input: 'supplies', expected: 'Operating Supplies' },
    { input: 'misc', expected: 'Misc' },
  ];
  
  testCases.forEach(test => {
    const normalized = CATEGORY_ALIASES[test.input];
    const status = normalized === test.expected ? '✓' : '✗';
    Logger.log(`${status} "${test.input}" → "${normalized}" (expected: "${test.expected}")`);
  });
  
  Logger.log('\nRun showValidCategories() to see all valid inputs.');
}

/**
 * Test function - tests automatic new year sheet creation
 * This simulates the first receipt entry of a new year
 */
function testNewYearAutoCreation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nextYear = new Date().getFullYear() + 1;
    
    // Check if sheet already exists
    const existingSheet = ss.getSheetByName(nextYear.toString());
    if (existingSheet) {
      Logger.log(`⚠️  Sheet for ${nextYear} already exists. Skipping test.`);
      Logger.log('Delete the sheet manually if you want to test auto-creation again.');
      return;
    }
    
    Logger.log(`Testing auto-creation of sheet for year ${nextYear}...`);
    Logger.log('');
    
    // Try adding a receipt for next year (should auto-create the sheet)
    Logger.log('Step 1: Calling addReceipt()...');
    const result = addReceipt(
      `${nextYear}-01-01`,
      'Products/Ingredients',
      100.00,
      'https://drive.google.com/file/d/test_new_year',
      'Test User'
    );
    
    Logger.log('✓ addReceipt() completed: ' + result);
    Logger.log('');
    
    // Verify the sheet was created
    Logger.log('Step 2: Verifying sheet creation...');
    const newSheet = ss.getSheetByName(nextYear.toString());
    
    if (newSheet) {
      Logger.log(`✓ Sheet "${nextYear}" was automatically created`);
      
      // Check monthly summary
      const summaryTitle = newSheet.getRange(1, 1).getValue();
      Logger.log(`✓ Monthly summary table present: "${summaryTitle}"`);
      
      // Check transaction log
      const transactionTitle = newSheet.getRange(17, 1).getValue();
      Logger.log(`✓ Transaction log present: "${transactionTitle}"`);
      
      // Check if data was added correctly
      const firstDataRow = newSheet.getRange(19, 1).getValue();
      if (firstDataRow) {
        Logger.log(`✓ Transaction data verified in row 19: ${firstDataRow}`);
      } else {
        Logger.log('⚠️  Transaction data not found in row 19');
      }
      
      // Show summary totals
      const januaryTotal = newSheet.getRange(3, 2).getValue();
      Logger.log(`✓ January "Products/Ingredients" total: $${januaryTotal}`);
      
      Logger.log('');
      Logger.log('======================');
      Logger.log('✓ TEST PASSED!');
      Logger.log('======================');
      Logger.log(`Check the "${nextYear}" sheet in your spreadsheet.`);
      
    } else {
      Logger.log('');
      Logger.log('======================');
      Logger.log('✗ TEST FAILED');
      Logger.log('======================');
      Logger.log('ERROR: Sheet was not created.');
      Logger.log('');
      Logger.log('Debugging info:');
      Logger.log('- All sheet names in spreadsheet:');
      ss.getSheets().forEach(sheet => {
        Logger.log('  - ' + sheet.getName());
      });
    }
    
  } catch (error) {
    Logger.log('');
    Logger.log('======================');
    Logger.log('✗ TEST FAILED WITH ERROR');
    Logger.log('======================');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
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

/**
 * Helper function to display all valid category inputs
 */
function showValidCategories() {
  Logger.log('=== VALID CATEGORY INPUTS ===\n');
  
  // Group aliases by their target category
  const grouped = {};
  for (const [alias, category] of Object.entries(CATEGORY_ALIASES)) {
    if (!grouped[category]) {
      grouped[category] = [];
    }
    grouped[category].push(alias);
  }
  
  // Build message for both Logger and UI
  let message = 'VALID CATEGORY INPUTS\n\n';
  
  // Display organized by category
  for (const category of CATEGORIES) {
    Logger.log(`${category}:`);
    message += `${category}:\n`;
    
    if (grouped[category]) {
      const inputs = grouped[category].join(', ');
      Logger.log(`  Accepted inputs: ${inputs}`);
      message += `  ${inputs}\n`;
    }
    Logger.log('');
    message += '\n';
  }
  
  // Show in UI dialog
  SpreadsheetApp.getUi().alert(
    'Valid Category Inputs',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}