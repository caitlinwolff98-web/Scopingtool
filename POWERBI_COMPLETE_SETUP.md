# Complete PowerBI Setup Guide for Bidvest Scoping Tool
## Autonomous Integration - Zero Manual Setup Required

---

## 🎯 Overview

This guide provides **complete step-by-step instructions** for setting up PowerBI to work automatically with the Bidvest Scoping Tool. Once configured, users can:

1. Run the VBA macro on their workbook (no PowerBI knowledge needed)
2. PowerBI automatically refreshes and processes the data
3. View comprehensive scoping dashboards in PowerBI
4. Export results back to Excel with division-level breakdowns

**Key Features:**
- ✅ Autonomous operation - users just run the VBA macro
- ✅ Automatic data refresh in PowerBI
- ✅ Division-based scoping analysis
- ✅ FSLi coverage tracking per division
- ✅ Export scoping results back to Excel
- ✅ Balance Sheet and Income Statement FSLi selection support

---

## 📋 Prerequisites

### Software Requirements
- **Microsoft Excel** 2016 or later (Windows)
- **Power BI Desktop** (latest version) - Download from [powerbi.microsoft.com](https://powerbi.microsoft.com)
- Bidvest Scoping Tool VBA modules installed in Excel

### Knowledge Requirements
- **For Initial Setup (Admin/Power User):** Basic PowerBI knowledge
- **For End Users:** None! Just run the VBA macro

---

## 🚀 Part 1: One-Time PowerBI Template Setup (Admin Only)

This section is done **once** by an admin/power user. End users will not need to do this.

### Step 1: Run the VBA Macro First

1. Open your consolidation workbook in Excel
2. Open the Bidvest Scoping Tool macro workbook
3. Click "Start TGK Scoping Tool" button
4. Follow the prompts to categorize tabs
5. Configure threshold-based scoping (optional)
6. Wait for the macro to complete

**Result:** The macro generates `Bidvest Scoping Tool Output.xlsx` with all tables.

### Step 2: Import Data into PowerBI

1. **Open Power BI Desktop**
2. Click **Home** → **Get Data** → **Excel Workbook**
3. Navigate to `Bidvest Scoping Tool Output.xlsx`
4. In the Navigator window, select **ALL** the following tables:

   **Core Data Tables:**
   - ☑ Full Input Table
   - ☑ Full Input Percentage
   - ☑ Journals Table (if exists)
   - ☑ Journals Percentage (if exists)
   - ☑ Full Consol Table (if exists)
   - ☑ Full Consol Percentage (if exists)
   - ☑ Discontinued Table (if exists)
   - ☑ Discontinued Percentage (if exists)

   **Reference Tables:**
   - ☑ FSLi Key Table
   - ☑ Pack Number Company Table

   **Scoping Tables (NEW!):**
   - ☑ Scoping Summary
   - ☑ Threshold Configuration (if threshold scoping was used)
   - ☑ Scoped In by Division
   - ☑ Scoped Out by Division
   - ☑ Scoped In Packs Detail

5. Click **Transform Data** (Important: Do NOT click Load yet)

### Step 3: Transform Data in Power Query

Power Query transformations prepare the data for optimal analysis.

#### Transform 3.1: Unpivot Full Input Table

1. In Power Query Editor, select **Full Input Table**
2. Select the **Pack** or **Pack Name** column (it may be named differently - look for the column with pack names)
3. Right-click → **Unpivot Other Columns**
4. This converts the wide format to long format:
   ```
   Before:            After:
   Pack | FSLi1 | FSLi2    →    Pack | FSLi | Amount
   P1   | 100   | 200           P1   | FSLi1| 100
                                 P1   | FSLi2| 200
   ```
5. Rename the columns:
   - **Attribute** → **FSLi**
   - **Value** → **Amount**
6. Remove null values:
   - Click the filter dropdown on **Amount** column
   - Uncheck **(null)**
7. Change data types if needed:
   - FSLi: Text
   - Amount: Decimal Number

#### Transform 3.2: Repeat for Other Data Tables

Repeat the unpivot process for:
- Full Input Percentage (rename Value → Percentage)
- Journals Table (if exists)
- Journals Percentage (if exists)
- Full Consol Table (if exists)
- Full Consol Percentage (if exists)
- Discontinued Table (if exists)
- Discontinued Percentage (if exists)

#### Transform 3.3: Leave Reference Tables As-Is

Do NOT transform these tables - they are already in the correct format:
- FSLi Key Table
- Pack Number Company Table
- Scoping Summary
- Threshold Configuration
- Scoped In by Division
- Scoped Out by Division
- Scoped In Packs Detail

#### Transform 3.4: Add Pack Code Column (if missing)

If your unpivoted tables don't have a **Pack Code** column:

1. Select the unpivoted table (e.g., Full Input Table)
2. Click **Add Column** → **Custom Column**
3. Name: `Pack Code`
4. Formula:
   ```m
   let
       lookupTable = #"Pack Number Company Table",
       result = Table.SelectRows(lookupTable, each [Pack Name] = [Pack])
   in
       if Table.RowCount(result) > 0 then result{0}[Pack Code] else null
   ```
5. Alternatively, use a simpler merge:
   - Click **Home** → **Merge Queries**
   - Select **Pack Number Company Table**
   - Match on **Pack Name** = **Pack Name**
   - Expand to get **Pack Code**

### Step 4: Close & Apply Transformations

1. Click **Home** → **Close & Apply**
2. Wait for Power BI to load all the data
3. You should now see all tables in the Fields pane on the right

### Step 5: Create Data Model Relationships

Relationships connect your tables for proper analysis.

1. Click **Model** view icon (left sidebar)
2. Create the following relationships by dragging and dropping:

#### Core Relationships

**Relationship 1: Pack Number Company → Full Input Table**
```
FROM: Pack Number Company Table[Pack Code]
TO:   Full Input Table[Pack Code]
Cardinality: One-to-Many (1:*)
Cross-filter: Single
```

**Relationship 2: FSLi Key → Full Input Table**
```
FROM: FSLi Key Table[FSLi]
TO:   Full Input Table[FSLi]
Cardinality: One-to-Many (1:*)
Cross-filter: Both (important for bi-directional filtering)
```

**Relationship 3: Pack Number Company → Scoping Summary**
```
FROM: Pack Number Company Table[Pack Code]
TO:   Scoping Summary[Pack Code]
Cardinality: One-to-One (1:1)
Cross-filter: Both
```

#### Additional Relationships (if tables exist)

Repeat the pack and FSLi relationships for other data tables:
- Pack Number Company → Journals Table
- FSLi Key → Journals Table
- Pack Number Company → Full Consol Table
- FSLi Key → Full Consol Table
- Pack Number Company → Discontinued Table
- FSLi Key → Discontinued Table

**Important:** Use **Pack Code** for relationships, NOT Pack Name!

### Step 6: Create DAX Measures

DAX measures provide calculations for your reports.

#### Create Measures Table

1. In Report view, right-click in Fields pane → **New Table**
2. Name: `_Measures`
3. Formula: `_Measures = { 1 }`

#### Essential Measures

Copy and paste these measures into the `_Measures` table:

```dax
// BASIC MEASURES

Total Amount = 
SUM('Full Input Table'[Amount])

Total Absolute Amount = 
SUMX('Full Input Table', ABS([Amount]))

Pack Count = 
DISTINCTCOUNT('Full Input Table'[Pack Code])

FSLi Count = 
DISTINCTCOUNT('Full Input Table'[FSLi])


// SCOPING MEASURES

Packs Scoped In = 
CALCULATE(
    DISTINCTCOUNT('Scoping Summary'[Pack Code]),
    'Scoping Summary'[Scoped In] = "Yes" ||
    'Scoping Summary'[Scoped In] = "Yes (Threshold)"
)

Packs Not Scoped = 
CALCULATE(
    DISTINCTCOUNT('Scoping Summary'[Pack Code]),
    'Scoping Summary'[Scoped In] = "No" ||
    'Scoping Summary'[Scoped In] = "Not Yet Determined"
)

Scoping Coverage % = 
DIVIDE(
    [Packs Scoped In],
    DISTINCTCOUNT('Pack Number Company Table'[Pack Code]),
    0
)

Untested % = 
1 - [Scoping Coverage %]


// DIVISION-BASED MEASURES

Packs Scoped In by Division = 
CALCULATE(
    [Packs Scoped In],
    ALLEXCEPT('Pack Number Company Table', 'Pack Number Company Table'[Division])
)

Division Coverage % = 
DIVIDE(
    [Packs Scoped In by Division],
    CALCULATE(
        [Pack Count],
        ALLEXCEPT('Pack Number Company Table', 'Pack Number Company Table'[Division])
    ),
    0
)


// FSLi COVERAGE MEASURES

FSLi Coverage Amount = 
CALCULATE(
    [Total Absolute Amount],
    FILTER(
        ALL('Full Input Table'),
        'Full Input Table'[Pack Code] IN VALUES('Scoping Summary'[Pack Code]) &&
        RELATED('Scoping Summary'[Scoped In]) IN {"Yes", "Yes (Threshold)"}
    )
)

FSLi Coverage % = 
DIVIDE(
    [FSLi Coverage Amount],
    CALCULATE(
        [Total Absolute Amount],
        ALL('Full Input Table')
    ),
    0
)

FSLi Untested % = 
1 - [FSLi Coverage %]


// THRESHOLD MEASURES

Threshold Value = 
SELECTEDVALUE('Threshold Configuration'[Threshold Value], 0)

Packs Above Threshold = 
CALCULATE(
    [Pack Count],
    FILTER(
        'Full Input Table',
        ABS([Amount]) >= [Threshold Value]
    )
)


// FORMATTING MEASURES

RAG Status = 
VAR Coverage = [Scoping Coverage %]
RETURN
    SWITCH(
        TRUE(),
        Coverage >= 0.80, "🟢 Green (≥80%)",
        Coverage >= 0.60, "🟡 Amber (60-79%)",
        "🔴 Red (<60%)"
    )
```

### Step 7: Create Report Pages

Create these essential report pages:

#### Page 1: Executive Dashboard

**KPI Cards (Top Row):**
1. Total Packs: `[Pack Count]`
2. Scoped In: `[Packs Scoped In]`
3. Coverage %: `[Scoping Coverage %]` (format as percentage)
4. RAG Status: `[RAG Status]`

**Visualizations:**
1. **Donut Chart** - Scoping Status
   - Legend: Scoping Summary[Scoped In]
   - Values: COUNT(Pack Code)

2. **Stacked Bar Chart** - Coverage by Division
   - Axis: Pack Number Company Table[Division]
   - Values: [Packs Scoped In], [Packs Not Scoped]

3. **Table** - Scoping Summary
   - Columns: Pack Code, Pack Name, Division, Scoped In, Suggested for Scope

#### Page 2: Division Analysis

**Filters:**
- Division (slicer)

**Visualizations:**
1. **Matrix** - Division Details
   - Rows: Division → Pack Name
   - Values: Total Amount, Coverage %

2. **Clustered Column Chart** - Scoped vs Not Scoped by Division
   - X-axis: Division
   - Y-axis: Count of Packs
   - Legend: Scoped In status

3. **Table** - Division Summary
   - From "Scoped In by Division" table
   - Show all columns

#### Page 3: FSLi Analysis

**Filters:**
- FSLi (slicer with search enabled)
- Statement Type (from FSLi Key Table)

**Visualizations:**
1. **Matrix** - FSLi × Pack
   - Rows: FSLi
   - Columns: Pack Name
   - Values: Amount
   - Conditional formatting on amounts

2. **Bar Chart** - Top 20 FSLis by Amount
   - Axis: FSLi
   - Values: Total Absolute Amount
   - Sort descending

3. **Line Chart** - FSLi Coverage Trend
   - X-axis: FSLi (top 10)
   - Y-axis: Coverage %

#### Page 4: Threshold Configuration

Only visible if threshold scoping was used.

**Visualizations:**
1. **Table** - Configured Thresholds
   - From "Threshold Configuration" table
   - Show FSLi Name, Threshold Value

2. **Table** - Packs Auto-Scoped
   - From "Threshold Configuration" table
   - Show Pack Code, Triggered By FSLi

3. **Card** - Packs Auto-Scoped Count
   - COUNT(Threshold Configuration[Pack Code])

#### Page 5: Detailed Scoping

**Import "Scoped In Packs Detail" table directly - it's already formatted!**

Visualizations:
1. **Table** - Scoped In Packs Detail
   - Use the table as-is from Excel
   - Add slicers for Pack Code, FSLi

2. **Stacked Bar Chart** - FSLi Composition per Pack
   - Axis: Pack Name
   - Values: Amount
   - Legend: FSLi

### Step 8: Set Up Automatic Refresh

1. **Save the Power BI file:**
   - File → Save As
   - Name: `Bidvest Scoping Dashboard.pbix`
   - Location: Same folder as `Bidvest Scoping Tool Output.xlsx`

2. **Configure Data Source Settings:**
   - Home → Transform Data → Data Source Settings
   - Click "Change Source..."
   - Use relative path or ensure path is consistent

3. **Enable Auto-Refresh:**
   - File → Options and Settings → Options
   - Data Load → Enable "Background refresh"
   - Set refresh interval (e.g., 5 minutes)

4. **Test Refresh:**
   - Run the VBA macro again with different data
   - In Power BI, click Home → Refresh
   - Verify data updates automatically

### Step 9: Publish to Power BI Service (Optional)

For cloud-based sharing:

1. **Publish Report:**
   - Home → Publish
   - Select workspace
   - Click "Select"

2. **Configure Gateway** (for scheduled refresh of local files):
   - Install Power BI Gateway on the machine with Excel files
   - Configure data source connections
   - Set up scheduled refresh (daily, hourly, etc.)

3. **Share Dashboard:**
   - In Power BI Service, share the dashboard with team members
   - Users can view and interact via web browser

---

## 👥 Part 2: End User Workflow (No PowerBI Knowledge Needed!)

Once the admin has set up the PowerBI template, end users follow this simple process:

### For End Users:

1. **Open your consolidation workbook in Excel**
2. **Open the Bidvest Scoping Tool macro workbook**
3. **Click "Start TGK Scoping Tool"**
4. **Follow the prompts:**
   - Categorize tabs (the tool guides you)
   - Optionally configure thresholds
   - Wait for completion
5. **Review the generated Excel file:**
   - `Bidvest Scoping Tool Output.xlsx` is created
   - Review "Scoping Summary" sheet
   - Review "Scoped In by Division" and "Scoped Out by Division" sheets
   - Check "Scoped In Packs Detail" for FSLi-level details

**That's it!** The user doesn't need to touch PowerBI at all.

### For PowerBI Users (Optional):

1. **Open the Power BI dashboard** (`Bidvest Scoping Dashboard.pbix`)
2. **Click Refresh** (or it auto-refreshes if configured)
3. **View updated dashboards:**
   - Executive Dashboard shows current coverage
   - Division Analysis shows division-level breakdown
   - FSLi Analysis shows FSLi-level coverage
4. **Export results if needed:**
   - Click on any visual → "Export data" → Excel

---

## 🔄 Part 3: Excel ↔ PowerBI ↔ Excel Workflow

This is the complete autonomous workflow:

```
┌─────────────────────────────────────────────────┐
│ 1. USER RUNS VBA MACRO ON THEIR WORKBOOK       │
│    • Consolidation data analyzed                │
│    • Tabs categorized                           │
│    • Optional: Thresholds configured            │
└────────────────┬────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────┐
│ 2. VBA GENERATES EXCEL OUTPUT                   │
│    "Bidvest Scoping Tool Output.xlsx"           │
│    • Full Input Table                           │
│    • FSLi Key Table                             │
│    • Pack Number Company Table                  │
│    • Scoping Summary                            │
│    • Scoped In/Out by Division                  │
│    • Scoped In Packs Detail                     │
│    • Threshold Configuration (if used)          │
└────────────────┬────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────┐
│ 3. POWERBI AUTO-REFRESHES (if open)             │
│    • Data automatically imported                │
│    • Relationships already configured           │
│    • Dashboards update instantly                │
└────────────────┬────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────┐
│ 4. USER REVIEWS IN POWERBI (optional)           │
│    • Executive dashboard                        │
│    • Division analysis                          │
│    • FSLi coverage tracking                     │
│    • Manual scope adjustments (if needed)       │
└────────────────┬────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────┐
│ 5. USER REVIEWS IN EXCEL (always available)     │
│    • Scoping Summary: pack-level decisions      │
│    • Scoped In by Division: division breakdown  │
│    • Scoped Out by Division: gaps identified    │
│    • Scoped In Packs Detail: FSLi amounts       │
└─────────────────────────────────────────────────┘
```

**Key Benefits:**
- ✅ User only needs to run VBA macro
- ✅ Excel output is fully usable without PowerBI
- ✅ PowerBI auto-refreshes if available
- ✅ Results exportable back to Excel
- ✅ Division-level reporting built-in
- ✅ FSLi coverage tracking automatic

---

## 🔧 Troubleshooting Common Issues

### Issue 1: Balance Sheet FSLis Not Selectable in Threshold Config

**Problem:** When configuring thresholds, Balance Sheet FSLis (like "Total Assets") are not appearing in the selection list.

**Solution:**
The VBA code filters out statement headers like "BALANCE SHEET" and "INCOME STATEMENT". Make sure:
1. Your FSLi names are actual line items (e.g., "Total Assets", "Current Assets")
2. They are NOT the statement headers themselves
3. The FSLi appears in the input data (row 9 onwards, column B)

**If the issue persists:**
- Check the `IsStatementHeader()` function in `ModThresholdScoping.bas`
- Ensure your FSLi name doesn't contain "BALANCE SHEET" as a substring

### Issue 2: Pack Names Not Connecting in PowerBI

**Problem:** Relationships between tables are not working, or data is not filtering correctly.

**Solution:**
- **Always use Pack Code for relationships, NOT Pack Name!**
- Pack Names may not be unique (multiple divisions can have same name)
- Pack Code is the unique identifier
- Ensure Pack Code is TEXT type in all tables
- Use Text.Trim() in Power Query to remove spaces

### Issue 3: PowerBI Not Auto-Refreshing

**Problem:** When you run the VBA macro and update the Excel file, PowerBI doesn't update.

**Solution:**
1. Ensure the Excel file name is exactly: `Bidvest Scoping Tool Output.xlsx`
2. Ensure it's saved in the same location every time
3. In PowerBI: Home → Transform Data → Data Source Settings → Update path if needed
4. Enable background refresh: File → Options → Data Load → Background refresh
5. Click "Refresh" button manually to test

### Issue 4: Relationships Are Ambiguous or Broken

**Problem:** PowerBI shows relationship errors or ambiguous paths.

**Solution:**
1. Open Model view
2. Delete all existing relationships
3. Recreate them in this order:
   - Pack Number Company → Full Input Table (Pack Code to Pack Code)
   - FSLi Key → Full Input Table (FSLi to FSLi)
   - Pack Number Company → Scoping Summary (Pack Code to Pack Code)
4. Set Cross-filter direction:
   - Pack relationships: Single
   - FSLi relationships: Both
   - Scoping Summary relationship: Both

### Issue 5: Division Column Missing or Empty

**Problem:** Division-based reports show "Unknown Division" for all packs.

**Solution:**
1. Check that Pack Number Company Table has a Division column
2. Verify it's populated with correct division names
3. Ensure segment tabs were categorized correctly in VBA
4. Re-run the VBA macro if needed

### Issue 6: Measures Showing Wrong Values

**Problem:** DAX measures show incorrect or unexpected values.

**Solution:**
1. Check the filter context - are slicers affecting the calculation?
2. Use DAX Studio (free tool) to debug measures
3. Verify relationships are active and correctly configured
4. Test measures on a simple table visual first
5. Check for CALCULATE/FILTER overrides affecting results

### Issue 7: "Scoped In Packs Detail" Table is Empty

**Problem:** The detailed scoping report doesn't show any data.

**Solution:**
1. Ensure threshold scoping was actually applied in VBA
2. Check that some packs were scoped in (Scoping Summary should show "Yes")
3. Verify the `scopedPacks` object was passed correctly
4. Re-run the VBA macro with threshold configuration

### Issue 8: Excel File Too Large / Power BI Slow

**Problem:** Large consolidation workbooks make the process slow.

**Solution:**
1. Filter out zero/null amounts in Power Query
2. Use Power Query to aggregate small FSLis
3. Archive historical data
4. Split analysis by division or period
5. Optimize DAX measures (use SUMMARIZE, avoid row context)

---

## 📊 Understanding the Data Flow

### VBA Module Output Structure

The VBA macro creates a workbook with these sheets:

| Sheet Name | Purpose | Used in PowerBI? |
|------------|---------|------------------|
| Full Input Table | Main data (Pack × FSLi matrix) | ✅ Yes - Core table |
| Full Input Percentage | Percentage coverage | ✅ Yes - Analysis |
| Journals Table | Journal entries | ✅ Yes (if exists) |
| Full Consol Table | Consolidated data | ✅ Yes (if exists) |
| Discontinued Table | Discontinued ops | ✅ Yes (if exists) |
| FSLi Key Table | FSLi reference with metadata | ✅ Yes - Dimension |
| Pack Number Company Table | Pack reference with divisions | ✅ Yes - Dimension |
| Scoping Summary | Pack-level scoping status | ✅ Yes - Core |
| Threshold Configuration | Applied thresholds | ✅ Yes (if used) |
| **Scoped In by Division** | **Division-level scoped packs** | ✅ **Yes - New!** |
| **Scoped Out by Division** | **Division-level gaps** | ✅ **Yes - New!** |
| **Scoped In Packs Detail** | **FSLi amounts per pack** | ✅ **Yes - New!** |
| Interactive Dashboard | Excel-only dashboard | ❌ No - Excel only |
| Scoping Calculator | Coverage calculator | ❌ No - Excel only |

### Data Model Star Schema

```
         ┌──────────────────────┐
         │  FSLi Key Table      │
         │  (Dimension)         │
         └──────┬───────────────┘
                │
                │ (1:Many)
                │
┌───────────────▼────────────────────┐
│  Full Input Table (Fact)           │
│  - Pack Code                       │
│  - FSLi                            │
│  - Amount                          │
└───────────┬────────────────────────┘
            │
            │ (Many:1)
            │
┌───────────▼────────────────────────┐
│  Pack Number Company Table         │
│  (Dimension)                       │
│  - Pack Code (PK)                  │
│  - Pack Name                       │
│  - Division                        │
└───────────┬────────────────────────┘
            │
            │ (1:1)
            │
┌───────────▼────────────────────────┐
│  Scoping Summary (Dimension)       │
│  - Pack Code                       │
│  - Scoped In                       │
│  - Suggested for Scope             │
└────────────────────────────────────┘
```

---

## 🎓 Best Practices

### For Administrators Setting Up PowerBI

1. **Create a Template:** Set up PowerBI once, save as template (.pbit file)
2. **Document Paths:** Note the exact path to Excel files for data source settings
3. **Test with Sample Data:** Always test the full workflow before deploying
4. **Provide Screenshots:** Create visual guides for end users
5. **Version Control:** Keep old versions of the .pbix file

### For End Users Running the Macro

1. **Consistent Naming:** Always save source workbooks with clear names
2. **Categorize Carefully:** Take time to categorize tabs correctly
3. **Use Thresholds:** Configure thresholds for automated scoping
4. **Review Excel First:** Check "Scoping Summary" before PowerBI
5. **Document Decisions:** Note why certain packs were scoped in/out

### For Data Analysis

1. **Start with Executive Dashboard:** Get the high-level view first
2. **Drill Down by Division:** Use division analysis to identify gaps
3. **Review FSLi Coverage:** Ensure key FSLis are covered
4. **Check Threshold Logic:** Verify automatic scoping makes sense
5. **Export for Documentation:** Save PowerBI views as images/PDFs

---

## 📖 Additional Resources

### PowerBI Resources
- [Power BI Documentation](https://docs.microsoft.com/power-bi/)
- [DAX Reference](https://dax.guide/)
- [Power Query M Reference](https://docs.microsoft.com/powerquery-m/)

### Tool Documentation
- See `DOCUMENTATION.md` for complete VBA module documentation
- See `VBA_Modules/README.md` for module-specific details
- See `FAQ.md` for common questions

---

## 🆘 Getting Help

If you encounter issues:

1. **Check this guide first** - Most issues are covered in Troubleshooting
2. **Review the Excel output** - Often the issue is in the source data
3. **Test with sample data** - Isolate whether it's a data or setup issue
4. **Check VBA module logs** - Errors are logged in Debug.Print statements
5. **Verify PowerBI relationships** - Use Model view to check connections

---

## ✅ Setup Checklist

Use this checklist to ensure everything is configured correctly:

### Initial Setup (Admin)
- [ ] VBA modules installed in Excel
- [ ] Test macro run completed successfully
- [ ] PowerBI Desktop installed
- [ ] All tables imported into PowerBI
- [ ] Data transformations applied (unpivot)
- [ ] Relationships created (Pack Code and FSLi)
- [ ] DAX measures added
- [ ] Report pages created
- [ ] Auto-refresh configured
- [ ] Template saved and shared

### End User Workflow
- [ ] Consolidation workbook open
- [ ] Macro workbook open
- [ ] Tabs categorized correctly
- [ ] Threshold configuration (optional)
- [ ] Output Excel file generated
- [ ] Scoping Summary reviewed
- [ ] Division reports reviewed
- [ ] PowerBI refreshed (if using)

---

## 📝 Version History

- **v3.0** (2024-11) - Complete rewrite with autonomous workflow
  - Added Division-based reporting
  - Added "Scoped In Packs Detail" with FSLi amounts
  - Fixed Balance Sheet FSLi selection
  - Unified documentation
  - Improved troubleshooting section

---

**Need more help?** Review the other documentation files:
- `DOCUMENTATION.md` - Complete VBA documentation
- `QUICK_REFERENCE.md` - Quick reference guide
- `FAQ.md` - Frequently asked questions
- `USAGE_EXAMPLES.md` - Real-world usage examples

---

*Last Updated: 2024-11*
*Compatible with: Bidvest Scoping Tool v2.0+*
