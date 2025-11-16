# Quick Start Guide - Bidvest Scoping Tool v3.1.0

**Purpose:** Get up and running with the enhanced scoping tool in 15 minutes

---

## ⚡ 5-Minute VBA Setup

### Step 1: Create Macro Workbook (2 minutes)
1. Open Excel
2. Create new workbook
3. Save as `Bidvest_Scoping_Tool.xlsm` (macro-enabled)
4. Close the workbook

### Step 2: Import VBA Modules (3 minutes)
1. Open `Bidvest_Scoping_Tool.xlsm`
2. Press `Alt + F11` (opens VBA Editor)
3. For each of the 8 files in `VBA_Modules` folder:
   - Go to `File` → `Import File...`
   - Select the `.bas` file
   - Click `Open`

**Files to import (in any order):**
- ModConfig.bas
- ModMain.bas
- ModTabCategorization.bas
- ModDataProcessing.bas
- ModTableGeneration.bas
- ModThresholdScoping.bas
- ModPowerBIIntegration.bas
- ModInteractiveDashboard.bas

### Step 3: Add Button (1 minute)
1. Close VBA Editor (or press `Alt + F11` again)
2. In Excel ribbon, click `Developer` tab
   - If you don't see Developer tab: File → Options → Customize Ribbon → Check "Developer"
3. Click `Insert` → `Button (Form Control)`
4. Draw a button on the worksheet
5. In "Assign Macro" dialog, select `StartScopingTool`
6. Click `OK`
7. Right-click button → `Edit Text` → Type "Start Scoping Tool"
8. Save the workbook

**✅ VBA Setup Complete!**

---

## 🚀 10-Minute First Run

### Step 1: Open Consolidation Workbook (30 seconds)
1. Open your TGK consolidation workbook (e.g., `Consolidation_2024.xlsx`)
2. Keep it open

### Step 2: Run the Tool (30 seconds)
1. Switch to `Bidvest_Scoping_Tool.xlsm`
2. Click "Start Scoping Tool" button
3. Read welcome message → Click `OK`
4. Enter your consolidation workbook name (e.g., `Consolidation_2024.xlsx`)
5. Click `OK`

### Step 3: Categorize Tabs (3-5 minutes)
For each tab, you'll see a dialog. Enter the category number:

**Most Common Categories:**
- `1` = Segment/Division tab (e.g., "TGK_UK", "TGK_US")
- `3` = Input Continuing Operations tab (REQUIRED - the main data tab)
- `5` = Consol tab
- `9` = Uncategorized (skip this tab)

**Example:**
- "TGK_UK" → Enter `1` → Enter division name "UK"
- "TGK_Input" → Enter `3`
- "TGK_Consol" → Enter `5`
- "Summary_Sheet" → Enter `9`

### Step 4: Select Consolidated Entity (30 seconds) 🆕
**NEW FEATURE!**

You'll see a list of all packs:
```
1. Bidvest Group Limited (BVT-001)
2. Bidvest UK Limited (BVT-UK-001)
3. Bidvest US Inc (BVT-US-001)
```

- Enter the number for your consolidated entity (e.g., `1` for BVT-001)
- Click `Yes` to confirm
- This pack will be excluded from scoping

### Step 5: Optional Threshold Scoping (2 minutes)
**Recommended for first-time users:**
- Click `Yes` to configure thresholds
- Select FSLIs (e.g., enter `1,3,5` for FSLIs 1, 3, and 5)
  - Or type names: `Revenue, Total Assets`
- Enter threshold value for each (e.g., `300000000` for $300M)
- Packs exceeding thresholds are automatically scoped in

**Alternative:**
- Click `No` to skip (you can scope manually in PowerBI later)

### Step 6: Wait for Processing (2-5 minutes)
- Watch status bar for progress
- Tool generates all tables
- Output saved as `Bidvest Scoping Tool Output.xlsx`

### Step 7: Review Output (1 minute)
1. Open `Bidvest Scoping Tool Output.xlsx`
2. Check these key sheets:
   - **Control Panel** - Overview
   - **Scoping Control Table** - For PowerBI scoping 🆕
   - **Pack Number Company Table** - See Is Consolidated flag 🆕
   - **Full Input Table** - Your main data
   - **Scoping Summary** - Recommendations

**✅ First Run Complete!**

---

## 📊 PowerBI Setup (Optional - 20 minutes)

### When to Use PowerBI
- ✅ You want interactive dashboards
- ✅ You need to manually adjust scoping decisions
- ✅ You want to visualize coverage by FSLi or Division
- ✅ You need dynamic percentage tracking

### Quick PowerBI Setup

**Step 1: Import Data (2 minutes)**
1. Open PowerBI Desktop
2. Home → Get Data → Excel Workbook
3. Select `Bidvest Scoping Tool Output.xlsx`
4. Select ALL tables (especially **Scoping Control Table** 🆕)
5. Click `Load`

**Step 2: Create Relationships (3 minutes)**
1. Click Model view (left sidebar, middle icon)
2. Drag from `Pack Number Company Table[Pack Code]` to `Scoping Control Table[Pack Code]`
3. Drag from `Pack Number Company Table[Pack Code]` to `Full Input Table[Pack Code]`
4. Drag from `FSLi Key Table[FSLi]` to `Scoping Control Table[FSLi]`

**Step 3: Create Key Measures (5 minutes)**

Go to Modeling tab → New Measure. Create these 3 essential measures:

```DAX
Total Packs = 
CALCULATE(
    DISTINCTCOUNT('Scoping Control Table'[Pack Code]),
    'Scoping Control Table'[Is Consolidated] = "No"
)
```

```DAX
Scoped In Packs = 
CALCULATE(
    DISTINCTCOUNT('Scoping Control Table'[Pack Code]),
    'Scoping Control Table'[Scoping Status] = "Scoped In",
    'Scoping Control Table'[Is Consolidated] = "No"
)
```

```DAX
Coverage % = 
VAR Scoped = CALCULATE(SUM('Scoping Control Table'[Amount]), 'Scoping Control Table'[Scoping Status] = "Scoped In", 'Scoping Control Table'[Is Consolidated] = "No")
VAR Total = CALCULATE(SUM('Scoping Control Table'[Amount]), 'Scoping Control Table'[Is Consolidated] = "No")
RETURN DIVIDE(ABS(Scoped), ABS(Total), 0)
```

**Step 4: Create Simple Dashboard (10 minutes)**

1. **Add KPI Cards:**
   - Drag [Total Packs] to canvas → Card visual
   - Drag [Scoped In Packs] to canvas → Card visual
   - Drag [Coverage %] to canvas → Card visual (format as %)

2. **Add Pack Slicer:**
   - Insert Slicer visual
   - Field: `Pack Number Company Table[Pack Name]`
   - Enable multi-select

3. **Add Scoping Table:**
   - Insert Table visual
   - Fields: Pack Name, Pack Code, FSLi, Amount, Scoping Status
   - This is where you'll change scoping! 🆕

4. **Add Coverage Chart:**
   - Insert Clustered Bar Chart
   - Axis: FSLi
   - Values: [Coverage %]

**✅ PowerBI Setup Complete!**

---

## 🎯 Manual Scoping in PowerBI (5 minutes)

### Method 1: Direct Edit (Simplest)
1. Click on your **Scoping Table** visual
2. Find a row you want to scope in
3. Click on the "Scoping Status" cell
4. Change value: `"Not Scoped"` → `"Scoped In"`
5. Watch your KPI cards update automatically! ✨

### Method 2: Use Slicers (Faster for multiple packs)
1. Use Pack Name slicer to select packs
2. Use FSLi slicer to filter FSLIs
3. View filtered rows in Scoping Table
4. Change Scoping Status for visible rows

---

## 📤 Export Results (2 minutes)

### From PowerBI to Excel
1. Click on your Scoping Table visual
2. Click `...` (More options)
3. Select `Export data`
4. Choose `Excel` format
5. Save as `Scoping_Decisions_Final.xlsx`

This file contains your manual scoping decisions for audit documentation.

---

## 🆘 Troubleshooting (2 minutes)

### Problem: "Could not find workbook"
**Solution:** Make sure consolidation workbook is open and name matches exactly (including .xlsx)

### Problem: "Required tabs are missing"
**Solution:** At least one tab must be categorized as "3" (Input Continuing Operations)

### Problem: PowerBI relationships not working
**Solution:** Use Pack Code for relationships, not Pack Name

### Problem: Consolidated entity still in analysis
**Solution:** Check that Is Consolidated = "No" filter is in your measures

### Problem: No divisions showing
**Solution:** Only tabs categorized as "1" (Segment) create divisions

---

## 📚 Full Documentation

For complete details, see:
- **POWERBI_DYNAMIC_SCOPING_GUIDE.md** - Complete workflow (500+ lines)
- **RELEASE_NOTES_V3.1.md** - What's new in v3.1
- **README.md** - Tool overview

---

## ✅ Success Checklist

After following this guide, you should have:
- [x] VBA modules imported (8 files)
- [x] Button created and working
- [x] First run completed successfully
- [x] Consolidated entity selected and flagged
- [x] Output file generated
- [x] PowerBI report created (optional)
- [x] Manual scoping tested (optional)

---

## 🎉 You're Ready!

**Time Invested:** 15-40 minutes (depending on PowerBI setup)  
**Result:** Full ISA 600 compliant scoping tool with dynamic PowerBI integration

**Next Steps:**
1. Run on your actual consolidation workbook
2. Review scoping recommendations
3. Refine scoping decisions in PowerBI
4. Export for audit documentation
5. Refer to full documentation as needed

---

**Quick Start Guide v3.1.0**  
**Last Updated:** November 2024  
**Questions?** See POWERBI_DYNAMIC_SCOPING_GUIDE.md for detailed help
