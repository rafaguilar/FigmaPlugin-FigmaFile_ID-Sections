# Google Sheets Section Extractor - Figma Plugin

🔧 **Enhanced version for extracting Figma sections based on Google Sheets data**

## 🎯 What it does

1. **Reads your Google Sheet** with campaign/content data
2. **Finds Figma sections** in the current file
3. **Creates new Figma pages** (one per row) with names based on columns A+B+C+D
4. **Copies selected sections** marked with "x" in columns E-Q to new pages
5. **Organizes sections** horizontally in a single row on each page
6. **Ignores Comments column** (column R) - used only for reference in sheets
7. **Automatically cleans up** by removing the "MASTER TEMPLATES" page after generation

## 📋 Quick Setup

### 1. Google Sheets API Setup
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable the Google Sheets API
4. Create credentials (API Key)
5. Restrict the API key to Google Sheets API only

### 2. Google Sheet Configuration
- Make sure your sheet is publicly readable
- Your sheet should have this structure:

```
| A       | B       | C          | D     | E                    | F                   | G      | H     | I           | J            | K              | L             | M              | N                 | O                  | P         | Q    | R        |
|---------|---------|------------|-------|----------------------|---------------------|--------|-------|-------------|--------------|----------------|---------------|----------------|-------------------|--------------------|-----------|------|----------|
| Account | Trigger | KeyMessage | Dates | Eats-Storefront-Ring | Partner-Rewards-Hub | Ucraft | Email | LandingPage | Email-Module | Eats-Billboard | Eats-Masthead | Rides-Masthead | Eats-Interstitial | Rides-Interstitial | Ring-NewB | Push | Comments |
| Wait    | Launch  | New offer  | May   | x                    |                     | x      |       | x           |              |                | x             |                | x                 |                    |           | x    | Note here|
| NonWait | Remind  | Last call  | June  |                      | x                   |        | x     |             | x            | x              |               | x              |                   | x                  | x         |      | Another  |
```

### 3. Figma File Setup - IMPORTANT WORKFLOW
**You need to copy your template sections to the same file where you'll run the plugin:**

1. **Open your template file** with the sections you need
2. **Copy the sections** (Eats-Storefront-Ring, Partner-Rewards-Hub, Ucraft, etc.)
3. **Open a new Figma file** or the file where you want to generate pages
4. **Paste the sections** into this file (any page is fine - recommend naming it "MASTER TEMPLATES")
5. **Run the plugin** from this file

**Required section names in current file (in order):**
- `Eats-Storefront-Ring`, `Partner-Rewards-Hub`, `Ucraft`, `Email`, `LandingPage`, `Email-Module`, `Eats-Billboard`, `Eats-Masthead`, `Rides-Masthead`, `Eats-Interstitial`, `Rides-Interstitial`, `Ring-NewB`, `Push`

### 4. Automatic Cleanup
- **MASTER TEMPLATES page deletion**: The plugin will automatically delete any page named "MASTER TEMPLATES" after generation
- **Safety checks**: Ensures current page is switched before deletion
- **Error handling**: Won't fail if the page doesn't exist

## 🛠️ Installation

1. **Download Plugin Files**: Extract this folder
2. **Install in Figma**: 
   - Open Figma Desktop App
   - Go to Plugins → Development → New Plugin
   - Choose "Link existing plugin"
   - Select your `manifest.json` file

## ⚙️ Optional Configuration

You can adjust layout settings in the `CONFIG` object in `code.js`:

```javascript
const CONFIG = {
  // Layout settings
  SECTION_SPACING: 250,        // Space between sections
  SECTIONS_PER_ROW: 999,      // Keep all sections in one horizontal row
  PAGE_MARGIN: 100            // Page margins
};
```

## 🚀 Usage

1. Open the plugin in Figma
2. Enter your Google Sheets API key
3. Enter your spreadsheet ID (from the URL)
4. **Choose whether to delete MASTER TEMPLATES page** (optional, with warning)
5. Click "Test Connection" to verify
6. Click "Extract Sections to Pages"
7. Confirm the generation (with cleanup warning if selected)
8. Wait for processing - new pages will be created with sections

### 🗑️ MASTER TEMPLATES Cleanup Options

- **Checkbox enabled**: Plugin will delete the "MASTER TEMPLATES" page after generation
- **Checkbox disabled**: Templates page will remain for future use
- **Warning dialog**: Shows before generation starts if deletion is enabled
- **Confirmation**: Clear indication in success message if page was removed

## 📊 Expected Sheet Structure & Section Mapping

```
| Column | Name                 | Purpose           | Creates Section?               |
|--------|----------------------|-------------------|--------------------------------|
| A      | Account              | Page naming       | No (metadata)                  |
| B      | Trigger              | Page naming       | No (metadata)                  |
| C      | KeyMessage           | Page naming       | No (metadata)                  |
| D      | Dates                | Page naming       | No (metadata)                  |
| E      | Eats-Storefront-Ring | Section flag      | Yes → "Eats-Storefront-Ring"   |
| F      | Partner-Rewards-Hub  | Section flag      | Yes → "Partner-Rewards-Hub"    |
| G      | Ucraft               | Section flag      | Yes → "Ucraft"                 |
| H      | Email                | Section flag      | Yes → "Email"                  |
| I      | LandingPage          | Section flag      | Yes → "LandingPage"            |
| J      | Email-Module         | Section flag      | Yes → "Email-Module"           |
| K      | Eats-Billboard       | Section flag      | Yes → "Eats-Billboard"         |
| L      | Eats-Masthead        | Section flag      | Yes → "Eats-Masthead"          |
| M      | Rides-Masthead       | Section flag      | Yes → "Rides-Masthead"         |
| N      | Eats-Interstitial    | Section flag      | Yes → "Eats-Interstitial"      |
| O      | Rides-Interstitial   | Section flag      | Yes → "Rides-Interstitial"     |
| P      | Ring-NewB            | Section flag      | Yes → "Ring-NewB"              |
| Q      | Push                 | Section flag      | Yes → "Push"                   |
| R      | Comments             | Reference only    | No (ignored)                   |
```

## 🐛 Troubleshooting

### ❌ "No sections found in source file"
- Verify your current file contains Figma Section layers
- Check that section names match exactly (case-sensitive): `Eats-Storefront-Ring`, `Partner-Rewards-Hub`, `Ucraft`, `Email`, `LandingPage`, `Email-Module`, `Eats-Billboard`, `Eats-Masthead`, `Rides-Masthead`, `Eats-Interstitial`, `Rides-Interstitial`, `Ring-NewB`, `Push`
- Ensure sections are not nested too deeply

### ❌ "Section not found in source file"
- Check section names in source file match the mapping exactly
- Verify sections are actual "Section" layer types, not frames

### ❌ "API key invalid"
- Verify your API key is correct in Google Cloud Console
- Check that Google Sheets API is enabled

### ❌ "Spreadsheet not found"
- Check spreadsheet ID is correct (from the URL)
- Verify sheet is publicly readable

## 📝 Page Naming Logic

Pages are named by concatenating columns A, B, C, D (Account, Trigger, KeyMessage, Dates):
- Empty cells are ignored
- "x" values are ignored  
- Each part is limited to 50 characters
- Parts are joined with " - "

Example: `"Wait - Launch - New offer - May"`

## 🎨 Section Layout

Sections are arranged horizontally:
- All sections in a single horizontal row (side by side)
- 250px spacing between sections (configurable)
- 100px margin around the page (configurable)
- Sections maintain their original size from source file
- Layout matches the Source_Templates page structure

## 🔄 Key Features

1. **Sections instead of Components**: Extracts Figma Section layers
2. **Current file sections**: Sections must be present in the current Figma file
3. **Comments ignored**: Column R (Comments) is not used for section creation
4. **Horizontal layout**: Sections are arranged in a single row horizontally
5. **Section preservation**: Maintains original section content and styling
6. **13 section types supported**: Covers comprehensive campaign asset types

## 📚 Resources

- [Figma Plugin API Documentation](https://www.figma.com/plugin-docs/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Figma Sections Documentation](https://help.figma.com/hc/en-us/articles/9771500257687-Organize-your-canvas-with-sections)
- [Setup Guide](./SETUP_GUIDE.md) for detailed instructions

---

**⚠️ Important:** Make sure your template sections are copied to the current file and have the exact names listed above!
