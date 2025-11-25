# Google Sheets Section Extractor Plugin - Changelog

## Update: Horizontal Layout Configuration

### Changes Made (Latest):

#### Layout Configuration Update:
- **Changed SECTIONS_PER_ROW**: From `3` to `999`
- **Effect**: All sections now arranged in a single horizontal row (side by side)
- **Benefit**: Layout matches Source_Templates page structure
- **Maintains**: 250px spacing between sections and 100px page margins

#### Files Modified:
- `code.js` - Updated CONFIG.SECTIONS_PER_ROW from 3 to 999
- `README.md` - Updated layout documentation to reflect horizontal arrangement

---

## Update: Section Names and Column Order Revision

### Changes Made (Latest):

#### Section List Update:
**Removed sections:**
- Meta(TBC)
- YouTube(TBC)

**Final 13 sections (in order):**
1. Eats-Storefront-Ring
2. Partner-Rewards-Hub
3. Ucraft
4. Email
5. LandingPage
6. Email-Module
7. Eats-Billboard
8. Eats-Masthead
9. Rides-Masthead
10. Eats-Interstitial
11. Rides-Interstitial
12. Ring-NewB
13. Push

#### Column Structure Updated:
- **Range changed**: From A:U to A:R (18 columns total)
- **Columns A-D**: Naming metadata (Account, Trigger, KeyMessage, Dates)
- **Columns E-Q**: 13 section flags
- **Column R**: Comments (ignored)

#### Files Modified:
- `code.js` - Updated COLUMNS mapping, SECTION_NAMES mapping, and RANGE
- `config.js` - Updated SECTION_MAPPING and SHEET_RANGE
- `README.md` - Complete documentation rewrite with correct column structure
- `CHANGELOG.md` - This file

---

## Update: Expanded Column Structure for New Campaign Assets

### Changes Made:

#### Column Structure Update:
**Expanded from 13 columns to 20 columns (A-T)**

**Previous structure:**
- A-D: Naming metadata (Account, Trigger, Key message, Dates)
- E-L: 8 section types
- M: Comments

**New structure:**
- A-D: Naming metadata (Account, Trigger, KeyMessage, Dates)
- E-S: **15 section types** (expanded coverage)
- T: Comments (still ignored)

#### New Section Types Added:
1. **Ring-NewB** - New banner ring format
2. **Rides-Interstitial** - Rides-specific interstitial
3. **Eats-Interstitial** - Eats-specific interstitial
4. **Rides-Masthead** - Rides app masthead
5. **Eats-Masthead** - Eats app masthead
6. **Eats-Billboard** - Eats billboard format
7. **Email-Module** - Email module component
8. **LandingPage** - Landing page design
9. **Ucraft** - Ucraft integration
10. **Partner-Rewards-Hub** - Partner rewards hub
11. **Eats-Storefront-Ring** - Eats storefront ring
12. **Meta(TBC)** - Meta platform assets
13. **YouTube(TBC)** - YouTube platform assets

**Retained sections:**
- Push
- Email

#### Code Changes (code.js):

1. **Updated RANGE**: `Sheet1!A:N` → `Sheet1!A:U` (21 columns)

2. **Updated COLUMNS mapping**:
```javascript
const COLUMNS = {
  ACCOUNT: 0,
  TRIGGER: 1,
  KEY_MESSAGE: 2,
  DATES: 3,
  PUSH: 4,
  RING_NEWB: 5,
  RIDES_INTERSTITIAL: 6,
  EATS_INTERSTITIAL: 7,
  RIDES_MASTHEAD: 8,
  EATS_MASTHEAD: 9,
  EATS_BILLBOARD: 10,
  EMAIL_MODULE: 11,
  LANDING_PAGE: 12,
  EMAIL: 13,
  UCRAFT: 14,
  PARTNER_REWARDS_HUB: 15,
  EATS_STOREFRONT_RING: 16,
  META_TBC: 17,
  YOUTUBE_TBC: 18
  // COMMENTS: 19 - Excluded
};
```

3. **Updated SECTION_NAMES mapping**: All 15 new section names mapped to their exact Figma layer names

4. **Updated generateFileName()**: Now uses columns A, B, C, D (Account, Trigger, KeyMessage, Dates) for page naming

#### Documentation Updates:

1. **README.md**:
   - Updated sheet structure table with all 20 columns
   - Updated required section names list
   - Updated page naming logic description
   - Updated examples to reflect new structure

2. **Error messages**: Updated to list all 15 required section names

#### Backward Compatibility:
- ⚠️ **Breaking change**: Old spreadsheets with 13 columns will need to be updated
- Section finding logic remains the same (looks for "x" markers)
- Page creation and layout logic unchanged
- MASTER TEMPLATES cleanup functionality preserved

### Benefits:

1. **More comprehensive asset coverage**: 15 section types vs 8 previously
2. **Better platform separation**: Rides vs Eats specific assets
3. **Clearer naming**: More descriptive section names
4. **Maintains flexibility**: Comments column still ignored for notes
5. **Scalable structure**: Easy to add more columns in future

---

## Previous Update: Enhanced UI with Progressive States and Workflow

### Changes Made:

#### UI Changes (ui.html):
1. **Progressive Button States**: Buttons now change appearance based on completion:
   - Test Connection → "Connection Checked ✓" (green) after successful test
   - Extract Sections → "Extract Sections to Pages - Done! ✓" (green) after generation
   - Delete button disappears after deletion, Cancel becomes "Close"

2. **Smart Delete Message**: Added a warning message that appears above the delete button only after pages are generated

3. **Enhanced Styling**: Added success button styles (green) and delete message styling with smooth animations

4. **Hidden Checkbox Initially**: The auto-delete checkbox remains hidden until the delete button is clicked

5. **State Persistence**: Completed buttons remain disabled and styled to show completion

#### JavaScript Changes (ui.html):
1. **State Management Functions**:
   - `updateTestConnectionSuccess()`: Changes test button to green "Connection Checked"
   - `updateGenerationSuccess()`: Changes generate button to green "Done!" and shows delete message
   - `updateDeleteSuccess()`: Hides delete button and changes Cancel to Close

2. **Enhanced Message Handling**: Updated onmessage to detect different types of success and trigger appropriate UI updates

3. **Smart Button Management**: Updated validation and processing functions to respect completed states

#### User Experience Flow:
1. **Initial State**: Clean interface with API fields and disabled buttons
2. **After Test**: Test button turns green with "Connection Checked ✓"
3. **After Generation**: 
   - Generate button turns green with "Extract Sections to Pages - Done! ✓"
   - Delete warning message appears above delete button
4. **After Deletion**: 
   - Delete button disappears
   - Cancel button becomes "Close"
   - Interface shows completion state

### Workflow Benefits:

1. **Clear Progress Indication**: Users can see what steps they've completed
2. **Context-Aware Interface**: Delete options only appear when relevant
3. **Prevents Redundant Actions**: Completed buttons stay disabled
4. **Progressive Disclosure**: UI elements appear based on user actions and progress
5. **Visual Feedback**: Color coding (green for success, red for delete) guides user understanding

### Technical Improvements:

1. **State-Aware Processing**: Processing states respect completion status
2. **Smooth Animations**: Fade-in effects for appearing elements
3. **Robust Error Handling**: UI states reset appropriately on errors
4. **Clean Separation**: Each workflow step has distinct visual states

### Files Modified:
- `ui.html` - Complete UI overhaul with progressive states, new styling, and enhanced JavaScript
- `code.js` - Updated with new column structure and section mappings
- `README.md` - Comprehensive documentation update for new structure
