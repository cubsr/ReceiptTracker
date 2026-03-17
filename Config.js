// ============================================================
// CONFIG.JS — Edit this file to configure the receipt tracker
// ============================================================
// NOTE: API keys and the Receipts folder ID are stored in .env
// (gitignored). Edit .env directly for those sensitive values.

// Expense categories shown in the monthly summary table.
// Edit this array, then run: Manage Transactions → Regenerate Categories from Array
const CATEGORIES = ['Retail Shelf', 'Gas', 'Rent', 'Backbar', 'Misc'];

const API_KEY_USERS = {
  'Key1': 'Levi',
  'Key2': 'Kate',
  'Key3': 'Noah',
};

const RECEIPTS_FOLDER_ID = '1wfFJ-NbbMA4b4bC64uMftmb6exjaQ876';

// Maps user-friendly input text to the category name above.
// Add aliases here when the API or manual entry needs shorthand.
const CATEGORY_ALIASES = {
  // Products for Sale
  'products': 'Retail Shelf',
  'product': 'Retail Shelf',
  'for sale': 'Retail Shelf',
  'retail shelf': 'Retail Shelf',

  // Gas
  'gas': 'Gas',
  'fuel': 'Gas',

  // Rent
  'rent': 'Rent',

  // Service Expenses (Tools/Shampoo/Similar)
  'service expenses': 'Backbar',
  'service expense': 'Backbar',
  'service': 'Backbar',
  'backbar supply': 'Backbar',
  'backbar': 'Backbar',

  // Misc
  'misc': 'Misc',
  'miscellaneous': 'Misc',
  'other': 'Misc',
};
