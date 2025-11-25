# 🚀 Setup Guide: Google Sheets Section Extractor

Complete step-by-step setup instructions for the Figma section extraction plugin.

## 📋 Quick Start Checklist

- [ ] Google account with access to Google Cloud Console
- [ ] Your Google Sheet with section data
- [ ] Figma file with organized sections
- [ ] Figma Desktop App installed

**Estimated setup time: 30-45 minutes**

## 🔑 Google Sheets API Setup

### Step 1: Create Google Cloud Project

1. **Go to Google Cloud Console**
   - Visit: https://console.cloud.google.com/
   - Sign in with your Google account

2. **Create or Select Project**
   - Click "Select a project" → "New Project"
   - Name it: "Figma Section Extractor"
   - Click "Create"

3. **Enable Google Sheets API**
   - Go to "APIs & Services" → "Library"
   - Search for "Google Sheets API"
   - Click on it and press "Enable"

### Step 2: Create API Key

1. **Navigate to Credentials**
   - Go to "APIs & Services" → "Credentials"
   - Click "Create Credentials" → "API Key"

2. **Restrict API Key (Important for Security)**
   - Click on your new API key to edit it
   - Under "API restrictions": Select "Restrict key" and check "Google Sheets API"
   - Under "Application restrictions": Select "HTTP referrers" and add: `https://www.figma.com/*`
   - Click "Save"

3. **Copy Your API Key**
   - Copy the API key and save it securely
   - ⚠️ **Keep this private**

### Step 3: Prepare Your Google Sheet

1. **Share Your Sheet**
   - Open your Google Sheet
   - Click "Share" → "Change to anyone with the link"
   - Set permission to "Viewer"

2. **Set Up Correct Structure**
   Your sheet should have these columns:
   ```
   A: Account | B: Trigger | C: Key message | D: Dates | E: Email | F: Push | G: Ring (En Route) | H: Ring (Message Hub) | I: Uber Interstitial | J: Landing page | K: Meta (TBC) | L: Youtube | M: Comments
   ```

3. **Mark Sections with "x"**
   - Put "x" in columns E-L for sections you want to extract
   - Column M (Comments) is for reference only - won't create sections

4. **Get Spreadsheet ID**
   - From your sheet URL: `docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit`
   - Copy the ID between `/d/` and `/edit`

## 🎨 Figma Source File Setup

### Step 1: Create Master Template File

1. **Create a New Figma File**
   - This will be your master template containing all sections
   - Name it something like "Template Sections"

2. **Add Figma Sections**
   - Use Insert → Section (or press Shift + S)
   - Create sections with these exact names:
     - `Email`
     - `Push`
     - `Ring (En Route, On Trip)`
     - `Ring (Message Hub)`
     - `Uber Interstitial`
     - `Landing page`
     - `Meta (TBC)`
     - `Youtube`

3. **Design Your Sections**
   - Add content to each section (components, frames, text, images)
   - Sections can contain anything you need
   - Make sure each section is properly sized

4. **Get File ID**
   - From your Figma URL: `https://www.figma.com/file/ABC123DEF456/Template-File`
   - Copy the file ID: `ABC123DEF456`

### Step 2: Section Organization Tips

- **Use descriptive section names** that match your spreadsheet exactly
- **Keep sections organized** on the canvas for easy management
- **Test section content** to ensure it looks good when copied
- **Consider section sizes** - they'll maintain their dimensions when copied

## 💾 Plugin Installation

### Step 1: Configure the Plugin

1. **Update config.js**
   ```javascript
   const SOURCE_FILE_ID = 'ABC123DEF456'; // Your actual file ID
   ```

2. **Update code.js CONFIG object**
   ```javascript
   const CONFIG = {
     SOURCE_FILE_ID: 'ABC123DEF456', // Your actual file ID
     SECTION_SPACING: 250,           // Space between sections
     SECTIONS_PER_ROW: 3,           // Sections per row
     PAGE_MARGIN: 100               // Page margins
   };
   ```

### Step 2: Install in Figma

1. **Open Figma Desktop App** (required)
2. **Install Plugin**
   - Go to "Plugins" → "Development" → "New Plugin"
   - Choose "Link existing plugin"
   - Select your `manifest.json` file
3. **Test Installation**
   - Plugin should open successfully

## 🧪 Testing & Usage

### Step 1: Test Setup

1. **Create Test Sheet** with 2-3 rows
2. **Mark sections** with "x" in various columns
3. **Open Plugin** in Figma
4. **Enter credentials** (API key and spreadsheet ID)
5. **Test Connection** - should show success

### Step 2: Extract Sections

1. **Click "Extract Sections to Pages"**
2. **Wait for processing** - watch status messages
3. **Check Results**:
   - New pages should be created
   - Each page named from columns B+C+D+E
   - Sections arranged in grid layout
   - Only sections marked with "x" are included

### Step 3: Verify Output

- **Page naming**: Should use data from columns B, C, D, E
- **Section extraction**: Only sections marked with "x"
- **Layout**: Sections arranged in 3-column grid
- **Content**: Sections should maintain original styling

## 🐛 Common Issues

### ❌ "Source file ID not configured"
**Solutions:**
1. Check SOURCE_FILE_ID is set in config.js
2. Verify file ID is correct from Figma URL
3. Ensure file is accessible

### ❌ "No sections found in source file"
**Solutions:**
1. Verify you're using Figma Section layers (not frames)
2. Check section names match exactly (case-sensitive)
3. Ensure sections aren't nested too deep

### ❌ "Section not found in source file"
**Solutions:**
1. Check section names in source file match mapping:
   - Column E → "Email"
   - Column F → "Push"
   - Column G → "Ring (En Route, On Trip)"
   - etc.
2. Verify section names are spelled exactly right

### ❌ "API Key Invalid"
**Solutions:**
1. Double-check API key from Google Cloud Console
2. Verify Google Sheets API is enabled
3. Check API restrictions are set correctly

## 🔧 Advanced Configuration

### Section Layout Customization
```javascript
CONFIG.SECTION_SPACING = 300;     // More space between sections
CONFIG.SECTIONS_PER_ROW = 2;      // Only 2 sections per row
CONFIG.PAGE_MARGIN = 150;         // Larger page margins
```

### Page Naming Customization
```javascript
CONFIG.FILENAME_SEPARATOR = ' | '; // Use | instead of -
CONFIG.MAX_FILENAME_PART_LENGTH = 30; // Shorter name parts
CONFIG.ADD_TIMESTAMP = true;       // Add timestamps to page names
```

## ✅ Success Checklist

By the end of setup, you should have:

- [ ] Working Google Sheets API key
- [ ] Publicly accessible Google Sheet with correct structure
- [ ] Master Figma file with properly named sections
- [ ] Installed and configured plugin
- [ ] Successful test run with section extraction
- [ ] Generated pages with organized sections

## 📞 Support

### Debug Tools
1. **Browser Console**: Press F12 in Figma to see error messages
2. **Plugin Console**: Check for specific error details
3. **Test Mode**: Use small datasets for debugging

### Key Resources
- [Figma Sections Guide](https://help.figma.com/hc/en-us/articles/9771500257687-Organize-your-canvas-with-sections)
- [Google Sheets API Docs](https://developers.google.com/sheets/api)
- [Figma Plugin API Docs](https://www.figma.com/plugin-docs/)

---

**🎉 Ready to extract sections!** Start with a small test dataset and scale up once everything works perfectly.
