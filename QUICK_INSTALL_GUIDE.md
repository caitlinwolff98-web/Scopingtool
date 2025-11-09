# Quick Installation Guide - v1.1.0

## What's New in v1.1.0

✅ **Fixed:** "Ambiguous name detected" VBA error  
✅ **Added:** Direct Power BI integration with entity scoping  
✅ **Enhanced:** Code robustness and error handling  
✅ **Created:** 4 new Power BI integration sheets  

---

## Installation (5 Minutes)

### Step 1: Open VBA Editor
1. Open your Excel workbook (or create new one)
2. Press `Alt + F11`

### Step 2: Import Modules (IN THIS ORDER)

**IMPORTANT:** Import in this exact order to avoid errors

1. **ModConfig.bas** ← Import FIRST (other modules need this)
2. **ModMain.bas**
3. **ModTabCategorization.bas**
4. **ModDataProcessing.bas**
5. **ModTableGeneration.bas**
6. **ModPowerBIIntegration.bas**

**How to Import:**
- File → Import File
- Navigate to `VBA_Modules` folder
- Select file → Open
- Repeat for each file

### Step 3: Verify
1. Debug → Compile VBAProject
2. Should show: "Compile completed successfully" (or no message = success)

### Step 4: Add Button (First Time Only)
1. Close VBA Editor
2. Insert → Button (Form Control)
3. Assign macro: `StartScopingTool`
4. Label button: "Start TGK Scoping Tool"

### Step 5: Save
1. Save as `.xlsm` (macro-enabled workbook)

---

## Usage (Same as Before)

1. Open your TGK consolidation workbook
2. Open your tool workbook with the button
3. Click "Start TGK Scoping Tool"
4. Enter consolidation workbook name
5. Categorize tabs (1-10)
6. Select column type (Consolidation recommended)
7. Wait for processing to complete
8. Review output workbook

---

## What You Get

### Original 10 Tables (Same as v1.0.0)
1. Full Input Table
2. Full Input Percentage
3. Journals Table
4. Journals Percentage
5. Full Console Table
6. Full Console Percentage
7. Discontinued Table
8. Discontinued Percentage
9. FSLi Key Table
10. Pack Number Company Table

### NEW: 4 Power BI Integration Sheets
11. **PowerBI_Metadata** - Tool info, source workbook, categories
12. **PowerBI_Scoping** - Entity scoping configuration template
13. **DAX Measures Guide** - Copy-paste DAX measures
14. **Entity Scoping Summary** - Entity totals and percentages

---

## Power BI Import (Enhanced)

### Step 1: Import Tables
1. Open Power BI Desktop
2. Get Data → Excel Workbook
3. Select output file
4. Check ALL tables (including new Power BI sheets)
5. Click "Transform Data"

### Step 2: Unpivot (Same as Before)
1. Select Full Input Table
2. Right-click "Pack" column → Unpivot Other Columns
3. Rename: Attribute → FSLi, Value → Amount
4. Repeat for other data tables

### Step 3: Create Relationships
- Pack Number Company Table ←→ Full Input Table
- Pack Number Company Table ←→ PowerBI_Scoping (NEW)
- Pack Number Company Table ←→ Entity Scoping Summary (NEW)
- FSLi Key Table ←→ Full Input Table

### Step 4: Add DAX Measures (NEW)
Open **DAX Measures Guide** sheet in Excel.  
Copy measures into Power BI:

```dax
Total Amount = SUM('Full Input Table'[Amount])

Threshold Flag = IF([Total Amount] > 300000000, "Yes", "No")

Coverage % = DIVIDE([Total Amount], CALCULATE([Total Amount], ALL('Full Input Table'[Pack])))
```

### Step 5: Create Scoping Dashboard (NEW)
**Table Visual:**
- Columns: Pack Name, Division, Total Amount, Coverage %
- Filter: Threshold Flag = "Yes"
- Sort: Total Amount (descending)

**Card Visuals:**
- Total Entities
- Scoped Entities
- Scoping %

**Bar Chart:**
- X-axis: Pack Name
- Y-axis: Total Amount
- Color: Division

---

## Common Issues

### "Can't find project or library"
**Solution:** Import ModConfig.bas first

### "Ambiguous name detected" still appears
**Solution:** 
1. Remove ALL old modules
2. Re-import in correct order

### Power BI sheets not created
**Solution:** 
1. Verify ModPowerBIIntegration.bas imported
2. Re-run tool

### "Object required" error
**Solution:** Run from StartScopingTool (not individual functions)

---

## Upgrading from v1.0.0

### Option 1: Fresh Import (Recommended)
1. Backup your current workbook
2. Remove old modules (optional)
3. Import 6 new modules (see Step 2 above)
4. Test on sample data

### Option 2: Add New Modules
1. Keep existing 4 modules (ModMain, ModTabCategorization, ModDataProcessing, ModTableGeneration)
2. Import 2 new modules:
   - ModConfig.bas (MUST import)
   - ModPowerBIIntegration.bas (MUST import)
3. Update existing modules with new versions (download from VBA_Modules folder)
4. Test

---

## Verification Checklist

After installation, verify:
- [ ] All 6 modules visible in VBA Project Explorer
- [ ] Compiles without errors (Debug → Compile)
- [ ] Tool runs without errors
- [ ] All 14 sheets created
- [ ] Power BI sheets have data
- [ ] Can import into Power BI

---

## Getting Help

1. **FIX_SUMMARY.md** - Complete problem analysis and solution
2. **CODE_IMPROVEMENTS.md** - Detailed improvements and examples
3. **DOCUMENTATION.md** - Full user guide
4. **POWERBI_INTEGRATION_GUIDE.md** - Power BI setup details

---

## What Changed (Technical)

**Fixed:**
- Removed duplicate GetTabByCategory function from ModDataProcessing
- All function calls now use: `ModTableGeneration.GetTabByCategory()`

**Added:**
- ModConfig.bas - Centralized configuration (220 lines)
- ModPowerBIIntegration.bas - Direct Power BI support (450 lines)

**Enhanced:**
- Better error handling (ModConfig.ShowError)
- Input validation (ModConfig.SafeTrim, IsValidNumber)
- Centralized constants (all modules use ModConfig)
- Comprehensive documentation

**Result:**
- No compilation errors ✅
- Enhanced Power BI workflow ✅
- Better code quality ✅
- 4 new integration sheets ✅

---

**Version:** 1.1.0  
**Install Time:** ~5 minutes  
**Compatible With:** Excel 2016+, Power BI Desktop  
**Breaking Changes:** None (backward compatible)

---

**Ready to install? Start with Step 1 above!**
