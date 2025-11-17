# ISA 600 Scoping Tool - Implementation Guide
## Get Running in 15 Minutes

**Version:** 5.0 Production Ready
**Last Updated:** November 2025
**Difficulty:** Easy (No prior VBA or Power BI experience required)
**Time Required:** 15-20 minutes

---

## 🎯 What You'll Achieve

By the end of this guide, you will have:

✅ VBA scoping tool installed and working
✅ Your consolidation workbook analyzed
✅ All FSLIs and packs extracted
✅ Power BI dashboard with interactive scoping
✅ Real-time coverage percentage tracking
✅ ISA 600 compliant audit documentation

---

## 📋 Prerequisites

Before starting, ensure you have:

- [ ] Windows computer with Excel 2016 or later
- [ ] Power BI Desktop installed ([Download free](https://powerbi.microsoft.com/desktop/))
- [ ] Your consolidation workbook (Excel file with Input Continuing tab)
- [ ] Macro security set to allow VBA (we'll configure this)
- [ ] 20 minutes of uninterrupted time

---

## 🚀 Phase 1: VBA Tool Installation (5 minutes)

### Step 1.1: Download the VBA Modules

1. Navigate to the `VBA_Modules` folder in this repository
2. You should see 8 files:
   - ModMain.bas
   - ModConfig.bas
   - ModTabCategorization.bas
   - ModDataProcessing.bas
   - ModTableGeneration.bas
   - ModPowerBIIntegration.bas
   - ModThresholdScoping.bas
   - ModInteractiveDashboard.bas

3. Download all 8 files to a folder on your computer (e.g., `C:\BidvestTool\VBA_Modules`)

### Step 1.2: Create the Tool Workbook

1. **Open Excel**
2. **Create a new blank workbook**
3. **Save it as:**
   - Name: `Bidvest_Scoping_Tool.xlsm`
   - Type: **Excel Macro-Enabled Workbook (.xlsm)**
   - Location: `C:\BidvestTool\` (or your preferred location)

### Step 1.3: Enable Macro Security

1. In Excel, go to **File → Options**
2. Click **Trust Center → Trust Center Settings**
3. Click **Macro Settings**
4. Select: **Enable all macros** (temporarily, for setup)
5. Check: **Trust access to the VBA project object model**
6. Click **OK** twice

> **Security Note:** After installation, you can change this to "Disable all macros with notification" for better security.

### Step 1.4: Open VBA Editor

1. Press **Alt + F11** (this opens the VBA Editor)
2. You should see the VBA Editor window with your workbook listed

### Step 1.5: Import the VBA Modules

For each of the 8 .bas files:

1. In VBA Editor, go to **File → Import File...**
2. Navigate to your `VBA_Modules` folder
3. Select **ModMain.bas** → Click **Open**
4. Repeat for the other 7 files:
   - ModConfig.bas
   - ModTabCategorization.bas
   - ModDataProcessing.bas
   - ModTableGeneration.bas
   - ModPowerBIIntegration.bas
   - ModThresholdScoping.bas
   - ModInteractiveDashboard.bas

5. **Verify:** In the Project Explorer (left pane), you should see 8 modules listed under "Modules"

### Step 1.6: Create a Run Button

1. Close the VBA Editor (or keep it open)
2. Back in Excel, go to **Developer** tab
   - Don't see Developer tab? **File → Options → Customize Ribbon** → Check "Developer"
3. Click **Insert → Button (Form Control)**
4. Draw a button on Sheet1 (click and drag)
5. In the "Assign Macro" dialog, select **StartScopingTool**
6. Click **OK**
7. Right-click the button → **Edit Text** → Type: "Run ISA 600 Scoping Tool"

### Step 1.7: Test the Installation

1. **Save your workbook** (Ctrl+S)
2. Click the button you just created
3. You should see a message: "Please enter the name of the workbook to analyze"
4. Click **Cancel** for now (we haven't prepared our data yet)

✅ **VBA Installation Complete!** If you saw the prompt, everything is working.

---

## 📊 Phase 2: Analyze Your Consolidation Workbook (5 minutes)

### Step 2.1: Prepare Your Consolidation Workbook

1. **Open your consolidation workbook** (e.g., `Bidvest_Consolidation_2024.xlsx`)
2. **Verify the structure:**
   - Row 6: Contains column type identifiers (e.g., "Original/Entity Currency")
   - Row 7: Contains pack names (e.g., "The Bidvest Group Limited", "UK Division")
   - Row 8: Contains pack codes (e.g., "BVT 001", "BVT 100")
   - Row 9+: Financial Statement Line Items (FSLIs) start here
   - Column B: FSLI names

3. **Identify your tabs:**
   - Which tab is "Input Continuing Operations"? (This is REQUIRED)
   - Do you have: Journals, Consol, Discontinued tabs? (Optional)
   - Which tabs are segment/division tabs? (Optional)

### Step 2.2: Run the Scoping Tool

1. **With both workbooks open** (your tool and your consolidation workbook):
   - Bidvest_Scoping_Tool.xlsm
   - Your consolidation workbook

2. **Click the "Run ISA 600 Scoping Tool" button**

3. **Enter workbook name:**
   - Type the EXACT name of your consolidation workbook (including .xlsx or .xlsm)
   - Example: `Bidvest_Consolidation_2024.xlsx`
   - Click **OK**

### Step 2.3: Categorize Tabs

You'll see a dialog asking you to categorize each worksheet. Use these numbers:

| Number | Category | When to Use |
|--------|----------|-------------|
| **3** | **Input Continuing Operations** | **Your main input tab (REQUIRED)** |
| 2 | Discontinued Operations | If you have a discontinued ops tab |
| 4 | Journals Continuing | If you have a journals tab |
| 5 | Consol Continuing | If you have a consolidated tab |
| 6 | Balance Sheet | If you have a separate BS tab |
| 7 | Income Statement | If you have a separate IS tab |
| 1 | Segment Tab | For division/segment tabs |
| 9 | Uncategorized | For tabs to ignore |

**Example:**
```
Tab: "TGK_Input_Continuing" → Enter: 3
Tab: "TGK_Journals" → Enter: 4
Tab: "TGK_Consol" → Enter: 5
Tab: "UK_Division" → Enter: 1
Tab: "Working_Papers" → Enter: 9
```

### Step 2.4: Select Consolidated Entity

The tool will ask: "Select the consolidated entity pack code"

1. You'll see a list of all pack codes (e.g., BVT 001, BVT 100, BVT 200)
2. Enter the code for **The Bidvest Group Consolidated** (usually **BVT-001** or **BVT 001**)
3. This entity will be automatically excluded from scoping calculations

### Step 2.5: Choose Currency Type

You'll be asked: "Use Consolidation Currency columns?"

- Click **Yes** (recommended) - uses consolidation currency
- Click **No** - uses entity/original currency

**Recommendation:** Click **Yes** to use consolidation currency for consistency.

### Step 2.6: Configure Threshold-Based Scoping (Optional)

The tool will ask if you want to set up automatic threshold-based scoping.

**Option 1: Yes (Recommended for first-time users)**
1. Click **Yes**
2. You'll see a list of all FSLIs
3. Enter the numbers of FSLIs you want to apply thresholds to
   - Example: Enter `1,3,5` to select FSLIs #1, #3, and #5
   - Or enter FSLI names: `Total Assets, Revenue`
4. For each selected FSLI, enter the threshold amount
   - Example: `300000000` for R300 million

**Option 2: No (Manual scoping only)**
1. Click **No**
2. You'll do all scoping manually in Power BI later

### Step 2.7: Wait for Processing

The tool will now:
- ✅ Extract all FSLIs from Column B (stops at "Notes" row)
- ✅ Identify all packs from rows 7-8
- ✅ Create all data tables
- ✅ Calculate percentage tables
- ✅ Apply threshold-based scoping (if configured)
- ✅ Generate Power BI integration tables

**Processing time:** 2-5 minutes (depending on workbook size)

**You'll see status bar updates** at the bottom of Excel:
- "Processing Input Continuing tab..."
- "Creating FSLi Key Table..."
- "Creating Percentage Tables..."

### Step 2.8: Output File Created

When complete, you'll see: "Scoping tool completed successfully!"

**Output file location:**
- Same folder as your consolidation workbook
- Name: **`Bidvest Scoping Tool Output.xlsx`**
- This file contains all the tables for Power BI

✅ **Data Extraction Complete!** Now let's visualize it in Power BI.

---

## 📊 Phase 3: Power BI Dashboard Setup (5-10 minutes)

### Step 3.1: Open Power BI Desktop

1. Launch **Power BI Desktop**
2. Click **Get Data** on the home screen
3. Select **Excel** → Click **Connect**
4. Navigate to **`Bidvest Scoping Tool Output.xlsx`**
5. Click **Open**

### Step 3.2: Select Tables to Import

You'll see a list of tables. **Select ALL of these:**

**Data Tables** (Check these):
- ✅ Full Input Table
- ✅ Full Input Percentage
- ✅ Journals Table (if exists)
- ✅ Journals Percentage (if exists)
- ✅ Full Consol Table (if exists)
- ✅ Full Consol Percentage (if exists)
- ✅ Discontinued Table (if exists)
- ✅ Discontinued Percentage (if exists)

**Reference Tables** (Check these):
- ✅ FSLi Key Table
- ✅ Pack Number Company Table

**Power BI Integration** (Check these):
- ✅ Scoping_Control_Table
- ✅ PowerBI_Scoping (if exists)
- ✅ Entity Scoping Summary (if exists)

**Do NOT import:**
- ❌ Control Panel
- ❌ DAX Measures Guide (this is text, not data)
- ❌ Interactive Dashboard (Excel-based, not for Power BI)

Click **Load** (NOT "Transform Data" - we'll transform later if needed)

### Step 3.3: Create Relationships

1. Click on the **Model view** icon (left sidebar - looks like three connected boxes)
2. You should see all your tables displayed
3. **Create these relationships** by dragging fields between tables:

**Relationship 1: Pack Code (Pack Company → Full Input)**
- Drag `Pack Code` from **Pack Number Company Table**
- Drop on `Pack Code` in **Full Input Table**
- Cardinality: One to Many (1:*)
- Cross filter direction: Both

**Relationship 2: FSLI (FSLi Key → Scoping Control)**
- Drag `FSLI` from **FSLi Key Table**
- Drop on `FSLI` in **Scoping_Control_Table**
- Cardinality: One to Many (1:*)
- Cross filter direction: Both

**Relationship 3: Pack Code (Pack Company → Scoping Control)**
- Drag `Pack Code` from **Pack Number Company Table**
- Drop on `Pack Code` in **Scoping_Control_Table**
- Cardinality: One to Many (1:*)
- Cross filter direction: Both

> **Note:** Power BI may auto-create some relationships. Verify they match the above.

### Step 3.4: Add DAX Measures

Now let's add the key measures for dynamic analysis. See **[DAX_MEASURES_LIBRARY.md](DAX_MEASURES_LIBRARY.md)** for comprehensive library.

**Quick Setup - Add These 5 Essential Measures:**

1. Click **Data view** (left sidebar - table icon)
2. Click on **Scoping_Control_Table** in the Fields pane
3. Click **New Measure** in the ribbon
4. Copy and paste each measure below:

**Measure 1: Total Packs**
```DAX
Total Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Measure 2: Scoped In Packs**
```DAX
Scoped In Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Measure 3: Coverage Percentage**
```DAX
Coverage % =
DIVIDE(
    [Scoped In Packs],
    [Total Packs],
    0
)
```

**Measure 4: Total Amount Scoped In**
```DAX
Total Amount Scoped In =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"}
)
```

**Measure 5: Coverage % by Amount**
```DAX
Coverage % by Amount =
DIVIDE(
    [Total Amount Scoped In],
    SUM(Scoping_Control_Table[Amount]),
    0
)
```

### Step 3.5: Create Your First Dashboard Page

1. Click **Report view** (left sidebar - bar chart icon)
2. You should see a blank canvas

**Add Summary Cards:**

1. Click **Card** visual (visualizations pane)
2. Drag **Total Packs** measure to the card
3. Format: Font size 28, Bold, Title "Total Packs"

4. Add 3 more cards:
   - **Scoped In Packs**
   - **Coverage %** (format as percentage)
   - **Coverage % by Amount** (format as percentage)

5. Position them across the top of your page

**Add Scoping Control Table:**

6. Click **Table** visual
7. Add these fields to the table:
   - Pack Name
   - Pack Code
   - Division
   - FSLI
   - Amount
   - **Scoping Status** ⭐
   - Is Consolidated

8. Make this table large (bottom 2/3 of page)

**Add Pack Scoping Chart:**

9. Click **Donut Chart** visual
10. Legend: `Scoping Status`
11. Values: `Pack Code` (count distinct)
12. Position on right side

✅ **Your first dashboard is complete!**

### Step 3.6: Enable Manual Scoping (Critical)

To enable real-time manual scoping:

1. Select the **Scoping Control Table** visual (click on it)
2. In **Visualizations** pane → **Format** → **General**
3. Scroll to **Advanced options**
4. Toggle **Edit mode** to **ON**

> **Detailed instructions:** See [POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md)

Now you can:
- Click in the **Scoping Status** column
- Change from "Not Scoped" to "Scoped In (Manual)"
- Watch your coverage percentages update in real-time!

### Step 3.7: Save Your Power BI File

1. Click **File → Save**
2. Name: `Bidvest_ISA600_Scoping_Dashboard.pbix`
3. Location: Same folder as your output Excel file

✅ **Power BI Dashboard Complete!**

---

## 🎯 Phase 4: Using the Tool (Ongoing)

### Daily Workflow

**Step 1: Update Data (if consolidation changes)**
1. Run VBA tool on updated consolidation workbook
2. In Power BI, click **Refresh** (toolbar)
3. All visuals update automatically

**Step 2: Manual Scoping**
1. Go to Scoping Control page
2. Filter to specific FSLI (e.g., "Inventory")
3. Select packs to scope in
4. Change Scoping Status to "Scoped In (Manual)"
5. Watch coverage % update

**Step 3: Generate Reports**
1. Filter by Division
2. Export visual to PowerPoint
3. Export data to Excel for audit files
4. Take screenshots for documentation

### Coverage Analysis by FSLI

**Create a coverage analysis visual:**

1. Add **Matrix** visual
2. Rows: `FSLI`
3. Columns: `Scoping Status`
4. Values: `Pack Code` (count distinct)
5. Add Coverage % measure to tooltips

This shows you coverage for each FSLI across all packs.

### Coverage Analysis by Division

**Create a division analysis visual:**

1. Add **Clustered Bar Chart**
2. Axis: `Division`
3. Values: `Coverage %`
4. Filter: Remove "Not Categorized"

This shows which divisions need more scoping.

---

## ✅ Verification Checklist

Use this checklist to verify everything is working:

- [ ] VBA tool runs without errors
- [ ] All 8 VBA modules imported
- [ ] Consolidation workbook opens and processes
- [ ] Output file created: `Bidvest Scoping Tool Output.xlsx`
- [ ] Output file contains 10+ tables
- [ ] FSLi Key Table has all FSLIs (no headers like "INCOME STATEMENT")
- [ ] FSLIs stop at "Notes" row (Notes section not included)
- [ ] Pack Number Company Table has all packs
- [ ] Consolidated entity marked "Is Consolidated = Yes"
- [ ] Power BI imports all tables successfully
- [ ] Relationships created correctly (no errors)
- [ ] DAX measures calculate correctly
- [ ] Total Packs count is correct (excludes consolidated entity)
- [ ] Coverage % displays as percentage (0-100%)
- [ ] Can change Scoping Status in table (edit mode works)
- [ ] Coverage % updates when scoping status changes
- [ ] Can filter by FSLI and Division
- [ ] Can export visuals to PDF/PowerPoint

**For detailed verification steps:** See [VERIFICATION_CHECKLIST.md](VERIFICATION_CHECKLIST.md)

---

## 🔧 Troubleshooting

### Issue: "Could not find workbook"

**Cause:** Workbook name doesn't match or workbook isn't open

**Solution:**
1. Verify consolidation workbook is open in Excel
2. Type EXACT name including extension (e.g., "Consolidation.xlsx")
3. Check for extra spaces in workbook name

### Issue: "No FSLIs found"

**Cause:** FSLIs not in expected location (Column B, starting Row 9)

**Solution:**
1. Verify FSLIs are in Column B
2. Verify data starts at Row 9 (after headers in rows 6-8)
3. Check that Row 6 has column type identifiers

### Issue: "INCOME STATEMENT" appearing as an FSLI

**Cause:** This should NOT happen in v5.0 (fixed)

**Solution:**
1. Verify you're using v5.0 VBA modules
2. Check ModDataProcessing.bas has `IsStatementHeader()` function
3. Re-import ModDataProcessing.bas

### Issue: FSLIs cut off / not all appearing

**Cause:** Tool stopped before "Notes" row, or Notes section included

**Solution:**
1. Verify you have a row with "NOTES" in Column B
2. This row should come AFTER your last FSLI
3. Check that FSLIs with data are not being filtered as empty

### Issue: Power BI relationships not working

**Cause:** Field names don't match or data types incompatible

**Solution:**
1. Verify Pack Code is text in both tables
2. Use Pack Code (NOT Pack Name) for relationships
3. Delete and recreate relationships manually

### Issue: Can't change Scoping Status in Power BI

**Cause:** Edit mode not enabled or using wrong visual type

**Solution:**
1. Use **Table** visual (not Matrix)
2. Enable Edit mode in Format pane
3. See detailed guide: [POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md)

### Issue: Coverage % not updating

**Cause:** Measure formula error or filter context issue

**Solution:**
1. Check DAX measure syntax (see [DAX_MEASURES_LIBRARY.md](DAX_MEASURES_LIBRARY.md))
2. Verify Scoping Status values match exactly: "Scoped In (Auto)", "Scoped In (Manual)"
3. Check that Is Consolidated filter is applied

### Issue: Consolidated entity appearing in counts

**Cause:** Is Consolidated filter not applied in measures

**Solution:**
1. Verify Is Consolidated column exists in Pack Number Company Table
2. Update measures to include: `[Is Consolidated] = "No"`
3. Check that consolidated entity was selected during VBA run

---

## 📚 Additional Resources

### Documentation

- **[README.md](README.md)** - Repository overview
- **[COMPREHENSIVE_GUIDE.md](COMPREHENSIVE_GUIDE.md)** - Detailed technical reference
- **[DAX_MEASURES_LIBRARY.md](DAX_MEASURES_LIBRARY.md)** - Complete DAX measures with examples
- **[VERIFICATION_CHECKLIST.md](VERIFICATION_CHECKLIST.md)** - Testing guide
- **[VISUALIZATION_ALTERNATIVES.md](VISUALIZATION_ALTERNATIVES.md)** - Why Power BI?
- **[POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md)** - Manual scoping setup

### VBA Code

- **VBA_Modules/README.md** - Module documentation
- All modules have extensive inline comments in v5.0

### ISA 600 Compliance

See **COMPREHENSIVE_GUIDE.md Section 7** for:
- ISA 600 revised requirements
- Component identification approach
- Scoping materiality considerations
- Audit trail documentation

---

## 🎓 Tips for Success

### Best Practices

1. **Start Small:** Test on a small consolidation first
2. **Use Threshold Scoping:** Let the tool do initial scoping automatically
3. **Validate FSLIs:** Review FSLi Key Table to ensure all line items captured
4. **Check Consolidated Entity:** Verify it's excluded from all calculations
5. **Save Often:** Save Power BI file frequently during setup
6. **Document Thresholds:** Keep record of threshold decisions for audit trail

### Common Mistakes to Avoid

1. ❌ Using Pack Name instead of Pack Code for relationships
2. ❌ Not excluding consolidated entity from calculations
3. ❌ Forgetting to refresh Power BI after updating Excel
4. ❌ Using Matrix visual instead of Table for manual scoping
5. ❌ Not enabling edit mode in Power BI
6. ❌ Including "Notes" section FSLIs (tool should auto-exclude)

### Performance Optimization

- Close unnecessary applications while processing large workbooks
- Use Consolidation Currency (faster than Entity Currency)
- Disable Excel calculation during VBA execution (automatic)
- Limit Power BI visuals per page (5-7 recommended)

---

## 📊 What's Next?

### Immediate Next Steps

1. **Run your first analysis** following this guide
2. **Create additional dashboard pages** for different analyses:
   - Page 1: Overview & Manual Scoping
   - Page 2: FSLI Coverage Analysis
   - Page 3: Division Coverage Analysis
   - Page 4: Threshold Configuration Review

2. **Explore advanced features:**
   - Drill-through to pack details
   - Bookmarks for different views
   - Export to PowerPoint for presentations

3. **Review ISA 600 compliance:**
   - See COMPREHENSIVE_GUIDE.md Section 7
   - Document scoping methodology
   - Prepare audit trail

### Advanced Topics

For advanced users, see:
- **Custom DAX measures** - DAX_MEASURES_LIBRARY.md
- **VBA customization** - VBA module inline comments
- **Power BI template creation** - POWER_BI_TEMPLATE_GUIDE.md (coming soon)
- **Integration with other systems** - Contact for custom development

---

## ❓ Need Help?

### Support Resources

1. **Documentation:** Check COMPREHENSIVE_GUIDE.md Section 8 (Troubleshooting)
2. **Verification:** Use VERIFICATION_CHECKLIST.md to diagnose issues
3. **Code Review:** VBA modules have extensive inline comments
4. **Community:** Check repository issues for similar problems

### Reporting Issues

If you encounter a problem:

1. Check this guide's Troubleshooting section
2. Review VERIFICATION_CHECKLIST.md
3. Check that you're using v5.0 modules
4. Document the issue with screenshots
5. Report on repository issues page

---

## 🎉 Success!

If you've completed this guide, you now have:

✅ Working VBA scoping tool
✅ Power BI dashboard with dynamic scoping
✅ ISA 600 compliant workflow
✅ Real-time coverage analysis
✅ Audit-ready documentation

**Congratulations!** You're ready to perform efficient, compliant ISA 600 scoping for Bidvest Group Limited consolidations.

---

**Guide Version:** 5.0
**Last Updated:** November 2025
**Estimated Time to Complete:** 15-20 minutes
**Difficulty:** Easy
**Status:** Production Ready ✅

---

**Questions or feedback?** See repository README.md for contact information.

**Looking for more detail?** See COMPREHENSIVE_GUIDE.md for complete technical documentation.
