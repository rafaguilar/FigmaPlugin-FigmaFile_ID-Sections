// Enhanced Figma Plugin for Google Sheets Section Extraction
// Compatible with Figma's JavaScript environment

figma.showUI(__html__, { width: 400, height: 600 });

// =============================================================================
// CONFIGURATION - Update these values with your actual source file ID
// =============================================================================

const CONFIG = {
  GOOGLE_SHEETS_API_KEY: 'AIzaSyC55La9khFWuMKnBA-9ORaaDH9NYyWfncc', // Set this from your config
  SPREADSHEET_ID: '1HneZCjZcNQ2I6I3HkWRda7FE_a0Z9RTHtoNAZjPfFws', // Set this from your config 
  RANGE: 'Sheet1!A:R', // Updated range to accommodate new columns (A-R = 18 columns)
  
  // Plugin Settings
  SECTION_SPACING: 250,
  MAX_FILENAME_PART_LENGTH: 50,
  FILENAME_SEPARATOR: ' - ',
  ADD_TIMESTAMP: false,
  SECTIONS_PER_ROW: 999, // Large number to keep all sections in one horizontal row
  PAGE_MARGIN: 100, // Margin around the sections
  MASTER_TEMPLATES_PAGE_NAME: 'MASTER TEMPLATES' // Name of page to delete after generation
};

// Column mapping (0-indexed) - Comments column (R/17) is excluded
const COLUMNS = {
  ACCOUNT: 0,
  TRIGGER: 1,
  KEY_MESSAGE: 2,
  DATES: 3,
  PUSH: 4,                      // Col 4 (E)
  RING_NEWB: 5,                 // Col 5 (F)
  RIDES_INTERSTITIAL: 6,        // Col 6 (G)
  EATS_INTERSTITIAL: 7,         // Col 7 (H)
  RIDES_MASTHEAD: 8,            // Col 8 (I)
  EATS_MASTHEAD: 9,             // Col 9 (J)
  EATS_BILLBOARD: 10,           // Col 10 (K)
  EMAIL_MODULE: 11,             // Col 11 (L)
  LANDING_PAGE: 12,             // Col 12 (M)
  EMAIL: 13,                    // Col 13 (N)
  UCRAFT: 14,                   // Col 14 (O)
  PARTNER_REWARDS_HUB: 15,      // Col 15 (P)
  EATS_STOREFRONT_RING: 16      // Col 16 (Q)
  // COMMENTS: 17 - Excluded from section creation
};

// Section name mapping (matches Figma section names)
const SECTION_NAMES = {
  [COLUMNS.PUSH]: 'Push',                                    // Col 4 (E)
  [COLUMNS.RING_NEWB]: 'Ring-NewB',                          // Col 5 (F)
  [COLUMNS.RIDES_INTERSTITIAL]: 'Rides-Interstitial',        // Col 6 (G)
  [COLUMNS.EATS_INTERSTITIAL]: 'Eats-Interstitial',          // Col 7 (H)
  [COLUMNS.RIDES_MASTHEAD]: 'Rides-Masthead',                // Col 8 (I)
  [COLUMNS.EATS_MASTHEAD]: 'Eats-Masthead',                  // Col 9 (J)
  [COLUMNS.EATS_BILLBOARD]: 'Eats-Billboard',                // Col 10 (K)
  [COLUMNS.EMAIL_MODULE]: 'Email-Module',                    // Col 11 (L)
  [COLUMNS.LANDING_PAGE]: 'LandingPage',                     // Col 12 (M)
  [COLUMNS.EMAIL]: 'Email',                                  // Col 13 (N)
  [COLUMNS.UCRAFT]: 'Ucraft',                                // Col 14 (O)
  [COLUMNS.PARTNER_REWARDS_HUB]: 'Partner-Rewards-Hub',      // Col 15 (P)
  [COLUMNS.EATS_STOREFRONT_RING]: 'Eats-Storefront-Ring'     // Col 16 (Q)
};

// =============================================================================
// MAIN PLUGIN LOGIC
// =============================================================================

figma.ui.onmessage = async function(msg) {
  try {
    if (msg.type === 'start-generation') {
      await generateFilesFromSheet(msg.apiKey, msg.spreadsheetId, msg.deleteMasterTemplates);
    }
    
    if (msg.type === 'test-connection') {
      await testGoogleSheetsConnection(msg.apiKey, msg.spreadsheetId);
    }
    
    if (msg.type === 'delete-master-templates') {
      await deleteMasterTemplatesPage();
    }
    
    if (msg.type === 'cancel') {
      figma.closePlugin();
    }
  } catch (error) {
    figma.ui.postMessage({
      type: 'error',
      message: 'Plugin error: ' + error.message
    });
  }
};

async function generateFilesFromSheet(apiKey, spreadsheetId, deleteMasterTemplates) {
  try {
    // Update config with user inputs
    CONFIG.GOOGLE_SHEETS_API_KEY = apiKey;
    CONFIG.SPREADSHEET_ID = spreadsheetId;
    
    // Step 1: Validate configuration
    figma.ui.postMessage({ type: 'status', message: 'Validating configuration...' });
    validateConfiguration();
    
    // Step 2: Load sections from current file
    figma.ui.postMessage({ type: 'status', message: 'Looking for sections in current file...' });
    const sourceSections = await loadSourceFileSections();
    
    if (sourceSections.length === 0) {
      throw new Error('No sections found in current file');
    }
    
    figma.ui.postMessage({ 
      type: 'status', 
      message: 'Found ' + sourceSections.length + ' sections in current file'
    });
    
    // SAFETY: Create a temporary working page and switch to it
    // This ensures we're not accidentally modifying Source_Template
    const tempWorkingPage = figma.createPage();
    tempWorkingPage.name = '_TEMP_WORKING_PAGE';
    figma.currentPage = tempWorkingPage;
    
    // Step 3: Fetch data from Google Sheets
    figma.ui.postMessage({ type: 'status', message: 'Fetching data from Google Sheets...' });
    const sheetData = await fetchSheetData(apiKey, spreadsheetId);
    
    if (!sheetData || sheetData.length === 0) {
      throw new Error('No data found in the specified sheet range');
    }
    
    // Step 4: Process each row (skip header row)
    const dataRows = sheetData.slice(1).filter(function(row) { return row && row.length > 0; });
    
    if (dataRows.length === 0) {
      throw new Error('No data rows found to process');
    }
    
    figma.ui.postMessage({ 
      type: 'status', 
      message: 'Found ' + dataRows.length + ' rows to process. Starting generation...'
    });
    
    let successCount = 0;
    let errorCount = 0;
    let createdPages = [];
    
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      
      try {
        figma.ui.postMessage({ 
          type: 'status', 
          message: 'Processing row ' + (i + 1) + ' of ' + dataRows.length + ': ' + generateFileName(row)
        });
        
        const newPage = await processRow(row, i + 1, sourceSections);
        if (newPage) {
          createdPages.push(newPage);
        }
        successCount++;
        
      } catch (error) {
        console.error('Error processing row ' + (i + 1) + ':', error);
        errorCount++;
        
        figma.ui.postMessage({ 
          type: 'warning', 
          message: 'Row ' + (i + 1) + ' failed: ' + error.message
        });
      }
    }
    
    // Step 5: Optional cleanup - Remove MASTER TEMPLATES page if requested
    if (deleteMasterTemplates && createdPages.length > 0) {
      figma.ui.postMessage({ type: 'status', message: 'Cleaning up template pages...' });
      await cleanupMasterTemplatesPage(createdPages[0]);
    } else if (deleteMasterTemplates) {
      figma.ui.postMessage({ 
        type: 'warning', 
        message: 'Could not cleanup MASTER TEMPLATES page - no pages were created successfully'
      });
    }
    
    // Step 6: Remove temporary working page
    figma.ui.postMessage({ type: 'status', message: 'Cleaning up temporary working page...' });
    if (tempWorkingPage) {
      if (createdPages.length > 0) {
        figma.currentPage = createdPages[0]; // Switch to first created page
        tempWorkingPage.remove(); // Remove temp page
      } else {
        // DEBUG MODE: No pages created, just remove temp page
        // First switch to another page before removing temp
        const otherPages = figma.root.children.filter(function(p) { return p.type === 'PAGE' && p !== tempWorkingPage; });
        if (otherPages.length > 0) {
          figma.currentPage = otherPages[0];
          tempWorkingPage.remove();
        }
      }
    }
    
    // Final summary
    let summary = 'Generation complete! ✅ ' + successCount + ' pages created successfully';
    if (errorCount > 0) {
      summary += ', ❌ ' + errorCount + ' failed';
    }
    if (deleteMasterTemplates && createdPages.length > 0) {
      summary += ', 🗑️ MASTER TEMPLATES page removed';
    }
    
    figma.ui.postMessage({
      type: 'success',
      message: summary
    });
    
  } catch (error) {
    throw new Error('Failed to generate files: ' + error.message);
  }
}

async function testGoogleSheetsConnection(apiKey, spreadsheetId) {
  try {
    figma.ui.postMessage({ type: 'status', message: 'Testing Google Sheets connection...' });
    
    const sheetData = await fetchSheetData(apiKey, spreadsheetId);
    
    if (sheetData && sheetData.length > 0) {
      figma.ui.postMessage({
        type: 'success',
        message: '✅ Connection successful! Found ' + sheetData.length + ' rows'
      });
    } else {
      throw new Error('No data found in sheet');
    }
  } catch (error) {
    throw new Error('Connection test failed: ' + error.message);
  }
}

function validateConfiguration() {
  const errors = [];
  
  // Note: Source file ID is no longer required since we're using current file
  // The sections should be present in the current file
  
  if (errors.length > 0) {
    throw new Error('Configuration validation failed:\n' + errors.join('\n'));
  }
}

// =============================================================================
// SOURCE FILE SECTIONS LOADING
// =============================================================================

async function loadSourceFileSections() {
  try {
    console.log('🔍 Starting to load source sections...');
    
    // Find sections ONLY in Source_Template or MASTER TEMPLATES page
    const sections = [];
    const seenNames = {}; // Track which section names we've already found
    
    function findTopLevelSections(page) {
      console.log('📄 Scanning page: ' + page.name + ' (children: ' + page.children.length + ')');
      
      // Only look at direct children of the page
      for (let i = 0; i < page.children.length; i++) {
        const node = page.children[i];
        console.log('  - Child ' + i + ': ' + node.type + ' named "' + node.name + '"');
        
        if (node.type === 'SECTION') {
          // Only add if we haven't seen this section name before (avoid duplicates)
          if (!seenNames[node.name]) {
            sections.push({
              node: node,
              name: node.name,
              width: node.width,
              height: node.height
            });
            seenNames[node.name] = true;
            console.log('    ✅ Added section: ' + node.name);
          } else {
            console.warn('    ⚠️⚠️⚠️ DUPLICATE FOUND IN SOURCE_TEMPLATE! Skipping duplicate section: ' + node.name);
            console.warn('    🚨 Your Source_Template has been corrupted with duplicates!');
            console.warn('    🪧 Please clean up Source_Template manually - delete all duplicate sections!');
          }
        }
      }
    }
    
    // Search ONLY in Source_Template or MASTER TEMPLATES pages
    let templatePage = null;
    console.log('🔍 Searching for Source_Template or MASTER TEMPLATES page...');
    
    for (let i = 0; i < figma.root.children.length; i++) {
      const page = figma.root.children[i];
      console.log('  Checking page: ' + page.name);
      
      if (page.type === 'PAGE' && (page.name === 'Source_Template' || page.name === 'MASTER TEMPLATES')) {
        templatePage = page;
        console.log('  ✅ Found template page: ' + page.name);
        findTopLevelSections(page);
        break; // Stop after finding first template page
      }
    }
    
    // If no template page found, show helpful error
    if (!templatePage) {
      throw new Error('Template page not found. Please create a page named "Source_Template" or "MASTER TEMPLATES" with your template sections.');
    }
    
    // If no sections found in template page, show helpful error
    if (sections.length === 0) {
      throw new Error('No sections found in template page. Please copy the template sections to "Source_Template" or "MASTER TEMPLATES" page first. Required sections: Eats-Storefront-Ring, Partner-Rewards-Hub, Ucraft, Email, LandingPage, Email-Module, Eats-Billboard, Eats-Masthead, Rides-Masthead, Eats-Interstitial, Rides-Interstitial, Ring-NewB, Push');
    }
    
    console.log('✅ Successfully loaded ' + sections.length + ' unique sections');
    
    // Check if there were duplicates detected in Source_Template
    const totalChildren = templatePage.children.filter(function(node) { return node.type === 'SECTION'; }).length;
    if (totalChildren > sections.length) {
      const duplicateCount = totalChildren - sections.length;
      console.error('🚨🚨🚨 DUPLICATE SECTIONS DETECTED IN SOURCE_TEMPLATE! 🚨🚨🚨');
      console.error('Total sections in template: ' + totalChildren);
      console.error('Unique sections loaded: ' + sections.length);
      console.error('Duplicate sections found: ' + duplicateCount);
      console.error('🪧 ACTION REQUIRED: Clean up your Source_Template page manually!');
      console.error('🪧 Delete all duplicate sections to avoid issues!');
      
      figma.ui.postMessage({
        type: 'warning',
        message: '⚠️ WARNING: Found ' + duplicateCount + ' duplicate sections in Source_Template! Please clean up duplicates manually.'
      });
    }
    
    // Preload all fonts from all sections
    console.log('🔤 Starting font preloading from template sections...');
    await preloadAllFontsFromSections(sections);
    console.log('✅ Font preloading complete');
    
    return sections;
    
  } catch (error) {
    throw new Error('Failed to load sections: ' + error.message);
  }
}

// =============================================================================
// GOOGLE SHEETS API
// =============================================================================

async function fetchSheetData(apiKey, spreadsheetId) {
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + CONFIG.RANGE + '?key=' + apiKey;
  
  const response = await fetch(url);
  
  if (!response.ok) {
    const errorData = await response.text();
    throw new Error('Google Sheets API error (' + response.status + '): ' + errorData);
  }
  
  const data = await response.json();
  return data.values || [];
}

// =============================================================================
// ROW PROCESSING
// =============================================================================

async function processRow(row, rowNumber, sourceSections) {
  // Step 1: Generate page name from columns B, C, D, E
  const pageName = generateFileName(row);
  
  if (!pageName || pageName.trim() === '') {
    throw new Error('Could not generate valid page name from row data');
  }
  
  // Step 2: Identify which sections to include (exclude Comments column)
  const sectionsToInclude = identifySections(row);
  
  if (sectionsToInclude.length === 0) {
    console.log('Row ' + rowNumber + ': No sections marked with "x", skipping');
    return null;
  }
  
  // 🐛 DEBUG MODE: Just log what WOULD be created - DON'T actually create anything
  console.log('═══════════════════════════════════════════════════════════');
  console.log('📄 ROW ' + rowNumber + ' - DEBUG MODE (NOT CREATING)');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('📝 Would create page: "' + pageName + '"');
  console.log('📦 Would include sections (' + sectionsToInclude.length + '):');
  for (let i = 0; i < sectionsToInclude.length; i++) {
    console.log('   ' + (i + 1) + '. ' + sectionsToInclude[i]);
  }
  console.log('═══════════════════════════════════════════════════════════');
  console.log('');
  
  // DON'T actually create pages or copy sections in DEBUG mode
  // Just return null
  return null;
}

function generateFileName(row) {
  const account = (row[COLUMNS.ACCOUNT] || '').toString().trim();
  const trigger = (row[COLUMNS.TRIGGER] || '').toString().trim();
  const keyMessage = (row[COLUMNS.KEY_MESSAGE] || '').toString().trim();
  const dates = (row[COLUMNS.DATES] || '').toString().trim();
  
  // Clean and filter parts
  const parts = [account, trigger, keyMessage, dates]
    .filter(function(part) { return part && part !== '' && part.toLowerCase() !== 'x'; })
    .map(function(part) { return part.substring(0, CONFIG.MAX_FILENAME_PART_LENGTH); })
    .map(function(part) { return part.replace(/[^\w\s-]/g, '').trim(); })
    .map(function(part) { return part.replace(/\s+/g, '_'); }); // Replace spaces with underscores
  
  if (parts.length === 0) {
    return 'Generated_Page_' + Date.now();
  }
  
  // Join with underscores
  let fileName = parts.join('_');
  
  if (CONFIG.ADD_TIMESTAMP) {
    const timestamp = new Date().toISOString().slice(0, 10);
    fileName += '_' + timestamp;
  }
  
  return fileName;
}

function identifySections(row) {
  const sections = [];
  
  console.log('\n🔍 IDENTIFYING SECTIONS FROM ROW:');
  console.log('  Row length: ' + row.length);
  console.log('  Row data: ' + JSON.stringify(row));
  
  Object.entries(SECTION_NAMES).forEach(function(entry) {
    const columnIndex = entry[0];
    const sectionName = entry[1];
    const cellValue = row[parseInt(columnIndex)];
    
    console.log('  Col ' + columnIndex + ' (' + sectionName + '): "' + cellValue + '"');
    
    if (cellValue && cellValue.toString().toLowerCase().trim() === 'x') {
      sections.push(sectionName);
      console.log('    ✅ ADDED: ' + sectionName);
    }
  });
  
  console.log('  📦 Final sections list: ' + JSON.stringify(sections));
  
  return sections;
}

// =============================================================================
// FONT LOADING HELPERS
// =============================================================================

async function preloadAllFontsFromSections(sections) {
  const uniqueFonts = {}; // Track unique font family+style combinations
  let successCount = 0;
  let warningCount = 0;
  
  // Step 1: Collect all unique fonts from all sections
  console.log('  🔍 Scanning ' + sections.length + ' sections for fonts...');
  for (let i = 0; i < sections.length; i++) {
    const section = sections[i];
    const textNodes = section.node.findAll(function(node) { return node.type === 'TEXT'; });
    
    for (let j = 0; j < textNodes.length; j++) {
      const textNode = textNodes[j];
      const fontKey = textNode.fontName.family + '|' + textNode.fontName.style;
      if (!uniqueFonts[fontKey]) {
        uniqueFonts[fontKey] = textNode.fontName;
      }
    }
  }
  
  const fontList = Object.values(uniqueFonts);
  console.log('  📝 Found ' + fontList.length + ' unique fonts to load');
  
  // Send status to UI
  figma.ui.postMessage({ 
    type: 'status', 
    message: 'Preloading ' + fontList.length + ' fonts from templates...'
  });
  
  // Step 2: Load all unique fonts
  for (let i = 0; i < fontList.length; i++) {
    const font = fontList[i];
    try {
      await figma.loadFontAsync(font);
      console.log('  ✅ Loaded: ' + font.family + ' ' + font.style);
      successCount++;
    } catch (error) {
      console.warn('  ⚠️ Could not load: ' + font.family + ' ' + font.style);
      warningCount++;
    }
  }
  
  // Summary message
  let summary = '✅ Preloaded ' + successCount + ' fonts successfully';
  if (warningCount > 0) {
    summary += ' (' + warningCount + ' warnings)';
  }
  console.log('  ' + summary);
  
  figma.ui.postMessage({ 
    type: 'status', 
    message: summary
  });
}

async function loadFontsForSection(section) {
  console.log('🔤 Loading fonts for section: ' + section.name);
  
  // Find all text nodes in the section
  const textNodes = section.findAll(function(node) { return node.type === 'TEXT'; });
  
  // Load fonts for each text node
  for (let i = 0; i < textNodes.length; i++) {
    const textNode = textNodes[i];
    try {
      await figma.loadFontAsync(textNode.fontName);
      console.log('  ✅ Loaded font: ' + textNode.fontName.family + ' ' + textNode.fontName.style);
    } catch (error) {
      console.warn('  ⚠️ Could not load font: ' + textNode.fontName.family + ' ' + textNode.fontName.style);
    }
  }
}

// =============================================================================
// SECTION COPYING LOGIC
// =============================================================================

async function copySectionsToPage(targetPage, sectionNames, sourceSections, rowNumber) {
  console.log('\n🔷🔷🔷 ROW ' + rowNumber + ' - STARTING COPY 🔷🔷🔷');
  console.log('📋 Sections requested from spreadsheet: ' + JSON.stringify(sectionNames));
  console.log('📦 Available source sections: ' + sourceSections.map(function(s) { return s.name; }).join(', '));
  
  // 🚨 CRITICAL: Verify we're NOT on Source_Template or MASTER TEMPLATES page
  if (figma.currentPage.name === 'Source_Template' || figma.currentPage.name === 'MASTER TEMPLATES') {
    console.error('❌❌❌ DANGER: Currently on template page "' + figma.currentPage.name + '"!');
    console.error('🔀 Switching to target page: "' + targetPage.name + '"');
  }
  
  figma.ui.postMessage({ 
    type: 'status', 
    message: 'Copying ' + sectionNames.length + ' sections to "' + targetPage.name + '"...'
  });
  
  // 🔒 LOCK: Set current page to target BEFORE any operations
  figma.currentPage = targetPage;
  console.log('✅ Current page locked to: "' + figma.currentPage.name + '"');
  
  let currentX = CONFIG.PAGE_MARGIN;
  let currentY = CONFIG.PAGE_MARGIN;
  let rowHeight = 0;
  let sectionsInCurrentRow = 0;
  
  for (let i = 0; i < sectionNames.length; i++) {
    const sectionName = sectionNames[i];
    
    console.log('\n  ➡️ Processing section ' + (i + 1) + '/' + sectionNames.length + ': "' + sectionName + '"');
    
    try {
      const sourceSection = sourceSections.find(function(section) {
        return section.name === sectionName;
      });
      
      console.log('  🔍 Looking for section "' + sectionName + '" in source sections...');
      
      if (!sourceSection) {
        console.error('  ❌ NOT FOUND in source sections!');
        await createPlaceholderSection(targetPage, sectionName, currentX, currentY);
        
        // Update position for next section
        const placeholderWidth = 300;
        const placeholderHeight = 200;
        rowHeight = Math.max(rowHeight, placeholderHeight);
        currentX += placeholderWidth + CONFIG.SECTION_SPACING;
        sectionsInCurrentRow++;
        
        if (sectionsInCurrentRow >= CONFIG.SECTIONS_PER_ROW) {
          currentX = CONFIG.PAGE_MARGIN;
          currentY += rowHeight + CONFIG.SECTION_SPACING;
          rowHeight = 0;
          sectionsInCurrentRow = 0;
        }
        
        continue;
      }
      
      console.log('  ✅ FOUND: "' + sourceSection.name + '" (width: ' + sourceSection.width + ', height: ' + sourceSection.height + ')');
      
      // Clone the section (fonts already preloaded)
      console.log('  📋 Cloning section...');
      const clonedSection = sourceSection.node.clone();
      clonedSection.x = currentX;
      clonedSection.y = currentY;
      
      // Add to target page - DIRECTLY append to targetPage object, don't rely on figma.currentPage
      console.log('  ➕ Appending to page "' + targetPage.name + '" at position (' + currentX + ', ' + currentY + ')...');
      targetPage.appendChild(clonedSection);
      
      console.log('  ✅ Successfully copied section: ' + sectionName);
      
      // Update position for next section
      rowHeight = Math.max(rowHeight, sourceSection.height);
      currentX += sourceSection.width + CONFIG.SECTION_SPACING;
      sectionsInCurrentRow++;
      
      // Check if we need to start a new row
      if (sectionsInCurrentRow >= CONFIG.SECTIONS_PER_ROW) {
        currentX = CONFIG.PAGE_MARGIN;
        currentY += rowHeight + CONFIG.SECTION_SPACING;
        rowHeight = 0;
        sectionsInCurrentRow = 0;
      }
      
    } catch (error) {
      console.error('  ❌❌❌ EXCEPTION while copying "' + sectionName + '": ' + error.message);
      console.error('  Stack: ' + error.stack);
      // Don't create error placeholders - just skip failed sections
      // This prevents duplicate sections from appearing when font loading fails
    }
  }
  
  console.log('\n🔷 ROW ' + rowNumber + ' - COPY COMPLETE 🔷\n');
}

// =============================================================================
// PLACEHOLDER CREATION
// =============================================================================

async function createPlaceholderSection(targetPage, sectionName, x, y) {
  // Create placeholder frame
  const frame = figma.createFrame();
  frame.name = sectionName + ' (Not Found)';
  frame.x = x;
  frame.y = y;
  frame.resize(300, 200);
  frame.fills = [{ 
    type: 'SOLID', 
    color: { r: 0.95, g: 0.95, b: 0.95 } 
  }];
  frame.strokes = [{ 
    type: 'SOLID', 
    color: { r: 0.7, g: 0.7, b: 0.7 } 
  }];
  frame.strokeWeight = 2;
  frame.dashPattern = [5, 5];
  
  // Add section name text
  await addTextToFrame(frame, sectionName, 10, 10, 16);
  await addTextToFrame(frame, '(Section not found in source file)', 10, 40, 12, { r: 0.6, g: 0.6, b: 0.6 });
  
  targetPage.appendChild(frame);
}

async function createErrorPlaceholder(targetPage, sectionName, x, y, errorMessage) {
  const frame = figma.createFrame();
  frame.name = sectionName + ' (Error)';
  frame.x = x;
  frame.y = y;
  frame.resize(300, 200);
  frame.fills = [{ 
    type: 'SOLID', 
    color: { r: 1, g: 0.9, b: 0.9 } 
  }];
  frame.strokes = [{ 
    type: 'SOLID', 
    color: { r: 0.8, g: 0.2, b: 0.2 } 
  }];
  frame.strokeWeight = 2;
  
  await addTextToFrame(frame, sectionName, 10, 10, 16);
  await addTextToFrame(frame, '❌ Error: ' + errorMessage.substring(0, 50), 10, 40, 10, { r: 0.8, g: 0.2, b: 0.2 });
  
  targetPage.appendChild(frame);
}

// =============================================================================
// PAGE CLEANUP
// =============================================================================

async function deleteMasterTemplatesPage() {
  try {
    figma.ui.postMessage({ 
      type: 'status', 
      message: 'Looking for MASTER TEMPLATES page...'
    });
    
    // Find the MASTER TEMPLATES page
    const masterTemplatesPage = figma.root.children.find(function(page) {
      return page.type === 'PAGE' && page.name === CONFIG.MASTER_TEMPLATES_PAGE_NAME;
    });
    
    if (!masterTemplatesPage) {
      figma.ui.postMessage({
        type: 'error',
        message: 'MASTER TEMPLATES page not found'
      });
      return;
    }
    
    // Safety check: Make sure we're not on the MASTER TEMPLATES page
    if (figma.currentPage === masterTemplatesPage) {
      // Switch to a different page first
      const otherPage = figma.root.children.find(function(page) {
        return page.type === 'PAGE' && page !== masterTemplatesPage;
      });
      
      if (otherPage) {
        figma.currentPage = otherPage;
        console.log('Switched to page: ' + otherPage.name + ' before deleting MASTER TEMPLATES');
      } else {
        figma.ui.postMessage({
          type: 'error',
          message: 'Cannot delete MASTER TEMPLATES page - no other pages available'
        });
        return;
      }
    }
    
    // Delete the MASTER TEMPLATES page
    masterTemplatesPage.remove();
    
    figma.ui.postMessage({ 
      type: 'success', 
      message: '✅ MASTER TEMPLATES page deleted successfully'
    });
    
    console.log('✅ Successfully deleted MASTER TEMPLATES page');
    
  } catch (error) {
    console.error('❌ Failed to delete MASTER TEMPLATES page:', error);
    figma.ui.postMessage({ 
      type: 'error', 
      message: 'Failed to delete MASTER TEMPLATES page: ' + error.message
    });
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

async function addTextToFrame(frame, text, x, y, fontSize, color) {
  fontSize = fontSize || 16;
  color = color || { r: 0, g: 0, b: 0 };
  
  try {
    // Load font (try to load a common font, fallback to default)
    try {
      await figma.loadFontAsync({ family: "Inter", style: "Regular" });
    } catch (fontError) {
      try {
        await figma.loadFontAsync({ family: "Roboto", style: "Regular" });
      } catch (fallbackError) {
        // Use default font
      }
    }
    
    const textNode = figma.createText();
    textNode.characters = text;
    textNode.x = frame.x + x;
    textNode.y = frame.y + y;
    textNode.fontSize = fontSize;
    textNode.fills = [{ type: 'SOLID', color: color }];
    
    frame.appendChild(textNode);
    return textNode;
  } catch (error) {
    console.error('Failed to add text:', error);
  }
}

// =============================================================================
// CLEANUP FUNCTIONS
// =============================================================================

async function cleanupMasterTemplatesPage(fallbackPage) {
  try {
    // Find the MASTER TEMPLATES page
    let masterTemplatesPage = null;
    
    for (let i = 0; i < figma.root.children.length; i++) {
      const page = figma.root.children[i];
      if (page.type === 'PAGE' && page.name === 'MASTER TEMPLATES') {
        masterTemplatesPage = page;
        break;
      }
    }
    
    if (masterTemplatesPage) {
      // Switch to a different page before deleting (can't delete current page)
      if (figma.currentPage === masterTemplatesPage) {
        figma.currentPage = fallbackPage;
      }
      
      // Remove the MASTER TEMPLATES page
      masterTemplatesPage.remove();
      
      console.log('✅ Successfully removed MASTER TEMPLATES page');
      figma.ui.postMessage({ 
        type: 'status', 
        message: 'Removed MASTER TEMPLATES page'
      });
    } else {
      console.log('ℹ️ No MASTER TEMPLATES page found to remove');
    }
    
  } catch (error) {
    console.error('❌ Failed to cleanup MASTER TEMPLATES page:', error);
    figma.ui.postMessage({ 
      type: 'warning', 
      message: 'Could not remove MASTER TEMPLATES page: ' + error.message
    });
  }
}

// Initialize plugin
figma.ui.postMessage({ type: 'ready' });
console.log('Google Sheets Section Extractor Plugin Loaded');
