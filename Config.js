// ============================================================
// CONFIG.JS — Edit this file to configure the receipt tracker
// ============================================================
// NOTE: API keys and the Receipts folder ID are stored in .env
// (gitignored). Edit .env directly for those sensitive values.

// Expense categories shown in the monthly summary table.
// Edit this array, then run: Manage Transactions → Regenerate Categories from Array
const CATEGORIES = ['Insurance', 'Media marketing', 'Accounting/Attorney', 'Vehicle', 'Services', 'Property maintenance', 'Horse maintenance'];

const API_KEY_USERS = {
  'Levi-key': 'Levi',
};

const RECEIPTS_FOLDER_ID = '';

// Maps user-friendly input text to the category name above.
// Add aliases here when the API or manual entry needs shorthand.
const CATEGORY_ALIASES = {
  // Insurance
  'insurance': 'Insurance',
  'ins': 'Insurance',
  'coverage': 'Insurance',
  'policy': 'Insurance',
  'liability': 'Insurance',

  // Media / Marketing
  'media': 'Media Marketing',
  'marketing': 'Media Marketing',
  'media marketing': 'Media Marketing',
  'advertising': 'Media Marketing',
  'ads': 'Media Marketing',
  'ad': 'Media Marketing',
  'social media': 'Media Marketing',
  'promo': 'Media Marketing',
  'promotion': 'Media Marketing',

  // Accounting / Attorney
  'accounting': 'Accounting/Attorney',
  'attorney': 'Accounting/Attorney',
  'accountant': 'Accounting/Attorney',
  'lawyer': 'Accounting/Attorney',
  'legal': 'Accounting/Attorney',
  'cpa': 'Accounting/Attorney',
  'bookkeeping': 'Accounting/Attorney',
  'taxes': 'Accounting/Attorney',
  'tax': 'Accounting/Attorney',

  // Vehicle
  'vehicle': 'Vehicle',
  'vehicles': 'Vehicle',
  'car': 'Vehicle',
  'truck': 'Vehicle',
  'auto': 'Vehicle',
  'mileage': 'Vehicle',
  'fuel': 'Vehicle',
  'gas': 'Vehicle',
  'maintenance auto': 'Vehicle',

  // Services
  'services': 'Services',
  'service': 'Services',
  'contractor': 'Services',
  'subcontractor': 'Services',
  'sub': 'Services',
  'labor': 'Services',
  'outsourced': 'Services',

  // Property Maintenance
  'property maintenance': 'Property Maintenance',
  'property': 'Property Maintenance',
  'repairs': 'Property Maintenance',
  'repair': 'Property Maintenance',
  'maintenance': 'Property Maintenance',
  'janitorial': 'Property Maintenance',
  'cleaning': 'Property Maintenance',
  'landscaping': 'Property Maintenance',
  'utilities': 'Property Maintenance',

  // Horse Maintenance
  'horse maintenance': 'Horse Maintenance',
  'horse': 'Horse Maintenance',
  'horses': 'Horse Maintenance',
  'equine': 'Horse Maintenance',
  'farrier': 'Horse Maintenance',
  'feed': 'Horse Maintenance',
  'vet': 'Horse Maintenance',
  'boarding': 'Horse Maintenance',
  'stable': 'Horse Maintenance',
  'stables': 'Horse Maintenance',
};
