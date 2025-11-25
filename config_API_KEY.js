// Configuration Helper for Google Sheets Section Extractor Plugin

// =============================================================================
// GOOGLE SHEETS CONFIGURATION
// =============================================================================

// Your Google Sheets API Key (get from Google Cloud Console)
const GOOGLE_SHEETS_API_KEY = 'AIzaSyC55La9khFWuMKnBA-9ORaaDH9NYyWfncc';

// Your Spreadsheet ID (from the URL: docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit)
const SPREADSHEET_ID = '1HneZCjZcNQ2I6I3HkWRda7FE_a0Z9RTHtoNAZjPfFws';

// Sheet range to read (adjust if your data is in different columns)
const SHEET_RANGE = 'Sheet1!A:R';

// =============================================================================
// SECTION MAPPING CONFIGURATION
// =============================================================================

// Map column names to expected Figma section names
// These should match the actual section names in your current Figma file
const SECTION_MAPPING = {
  'Eats-Storefront-Ring': 'Eats-Storefront-Ring',     // Column E (index 4)
  'Partner-Rewards-Hub': 'Partner-Rewards-Hub',       // Column F (index 5)
  'Ucraft': 'Ucraft',                                 // Column G (index 6)
  'Email': 'Email',                                   // Column H (index 7)
  'LandingPage': 'LandingPage',                       // Column I (index 8)
  'Email-Module': 'Email-Module',                     // Column J (index 9)
  'Eats-Billboard': 'Eats-Billboard',                 // Column K (index 10)
  'Eats-Masthead': 'Eats-Masthead',                   // Column L (index 11)
  'Rides-Masthead': 'Rides-Masthead',                 // Column M (index 12)
  'Eats-Interstitial': 'Eats-Interstitial',           // Column N (index 13)
  'Rides-Interstitial': 'Rides-Interstitial',         // Column O (index 14)
  'Ring-NewB': 'Ring-NewB',                           // Column P (index 15)
  'Push': 'Push'                                      // Column Q (index 16)
  // Column R (Comments) is excluded from section extraction
};

// =============================================================================
// PLUGIN SETTINGS
// =============================================================================

const PLUGIN_SETTINGS = {
  // How to space sections on the generated pages (pixels)
  SECTION_SPACING: 250,
  
  // Maximum length for each part of the page name
  MAX_FILENAME_PART_LENGTH: 50,
  
  // Character to use for joining page name parts
  FILENAME_SEPARATOR: ' - ',
  
  // Whether to add creation timestamp to page names
  ADD_TIMESTAMP: false,
  
  // Layout settings
  SECTIONS_PER_ROW: 999,      // Keep all sections in one horizontal row
  PAGE_MARGIN: 100,           // Margin around the sections
  
  // Cleanup settings
  MASTER_TEMPLATES_PAGE_NAME: 'MASTER TEMPLATES'  // Page to delete after generation
};

// =============================================================================
// SETUP INSTRUCTIONS
// =============================================================================

/*

QUICK SETUP STEPS:

1. GOOGLE SHEETS API:
   - Go to https://console.cloud.google.com/
   - Create/select project
   - Enable Google Sheets API
   - Create API key
   - Update GOOGLE_SHEETS_API_KEY above

2. SPREADSHEET SETUP:
   - Make your sheet publicly readable
   - Copy spreadsheet ID from URL
   - Update SPREADSHEET_ID above
   - Ensure columns match expected structure:
     A: Account, B: Trigger, C: KeyMessage, D: Dates
     E: Eats-Storefront-Ring, F: Partner-Rewards-Hub, G: Ucraft, H: Email
     I: LandingPage, J: Email-Module, K: Eats-Billboard, L: Eats-Masthead
     M: Rides-Masthead, N: Eats-Interstitial, O: Rides-Interstitial
     P: Ring-NewB, Q: Push, R: Comments (ignored)

3. FIGMA FILE PREPARATION:
   - Copy template sections to your current file
   - Sections should be named exactly as shown in SECTION_MAPPING
   - Run plugin from the file containing the sections

4. INSTALL PLUGIN:
   - Install plugin in Figma Desktop App
   - Test with small dataset first

SECTION REQUIREMENTS:
- Current file must contain Figma "Section" layer types
- Section names must match the mapping above exactly
- Sections can contain any content (components, frames, text, images)
- Comments column (R) will be ignored for section extraction
- MASTER TEMPLATES page will be automatically deleted after generation
- All sections will be arranged horizontally in one row (side by side)

*/

// =============================================================================
// VALIDATION HELPER
// =============================================================================

function validateConfiguration() {
  const errors = [];
  
  if (!GOOGLE_SHEETS_API_KEY || GOOGLE_SHEETS_API_KEY === 'AIzaSyC55La9khFWuMKnBA-9ORaaDH9NYyWfncc') {
    errors.push('Google Sheets API key not configured');
  }
  
  if (!SPREADSHEET_ID || SPREADSHEET_ID === '1HneZCjZcNQ2I6I3HkWRda7FE_a0Z9RTHtoNAZjPfFws') {
    errors.push('Spreadsheet ID not configured');
  }
  
  if (errors.length > 0) {
    console.error('Configuration Errors:', errors);
    return false;
  }
  
  console.log('Configuration validation passed!');
  return true;
}

// Export configuration
const PLUGIN_CONFIG = {
  GOOGLE_SHEETS_API_KEY,
  SPREADSHEET_ID,
  RANGE: SHEET_RANGE,
  SECTION_MAPPING,
  PLUGIN_SETTINGS
};

// Run validation
validateConfiguration();
