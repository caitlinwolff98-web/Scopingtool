# PowerBI Dynamic Scoping Guide
## Complete Guide for Manual Scoping with Bidvest Scoping Tool

---

## 🎯 Overview

This guide provides **complete step-by-step instructions** for using PowerBI to dynamically scope packs and FSLIs after running the VBA macro. The workflow supports:

1. **Automatic threshold-based scoping** in VBA (optional)
2. **Manual pack/FSLi scoping** in PowerBI with live updates
3. **Per-FSLi and per-Division analysis** showing scoped percentages
4. **Consolidated entity exclusion** from scoping calculations
5. **Bidirectional workflow** - changes in PowerBI can be tracked

---

## 📋 Key Features

### ✅ What's New
- **Consolidated Entity Selection**: VBA macro now prompts to select which pack is the consolidated entity (excluded from scoping)
- **Pack Name Relationships**: All tables now include both Pack Name and Pack Code for proper PowerBI relationships
- **Scoping Control Table**: New table enables manual scoping status updates in PowerBI
- **Division-Only from Category 1**: Only tabs categorized as "Segment Tabs" (Category 1) are treated as divisions
- **Dynamic Percentage Tracking**: Real-time calculation of scoped in/out percentages per FSLi and Division

### 🎯 Workflow
```
Excel VBA Macro
    ↓
Select Consolidated Entity (e.g., BVT-001)
    ↓
Optional: Threshold-Based Auto Scoping
    ↓
Generate Tables with Pack Name + Pack Code
    ↓
Import to PowerBI
    ↓
Manually Scope/Unscope Packs & FSLIs
    ↓
View Dynamic Coverage Analysis
    ↓
Export Results (Optional)
```

---

## 🚀 Part 1: VBA Macro Setup

### Step 1: Run the Macro

1. Open your consolidation workbook in Excel
2. Open the Bidvest Scoping Tool macro workbook
3. Click **"Start TGK Scoping Tool"** button
4. Enter the consolidation workbook name
5. Categorize tabs:
   - **Category 1 (Segment Tabs)**: These become Divisions
   - **Category 3 (Input Continuing)**: Required - primary data source
   - **Other categories**: Do NOT become divisions

### Step 2: Select Consolidated Entity

**NEW FEATURE!**

After tab categorization, you'll see a dialog listing all packs:

```
CONSOLIDATED ENTITY SELECTION

Select which pack represents the CONSOLIDATED entity.
This pack will be EXCLUDED from scoping calculations.

Available Packs:
------------------------------------------------------------
1. Bidvest Group Limited (BVT-001)
2. Bidvest UK Limited (BVT-UK-001)
3. Bidvest US Inc (BVT-US-001)
4. Bidvest Europe SA (BVT-EU-001)

Enter the number of the consolidated pack:
(Or leave blank to include all packs in scoping)
```

- **Enter the number** of the consolidated pack (e.g., `1` for BVT-001)
- The consolidated pack will be marked with `Is Consolidated = Yes`
- It will be **excluded** from threshold calculations and scoping analysis
- Click **Yes** to confirm your selection

### Step 3: Optional Threshold-Based Scoping

- Choose whether to configure automatic threshold-based scoping
- Select FSLIs for threshold analysis (e.g., "Total Assets", "Revenue")
- Enter threshold values for each FSLI
- Packs exceeding thresholds are automatically marked as "Scoped In"

### Step 4: Wait for Processing

The macro generates:
- **Full Input Table** (with Pack Name + Pack Code)
- **Pack Number Company Table** (with Is Consolidated flag)
- **Scoping Control Table** (for PowerBI dynamic scoping)
- **Scoping Summary** (recommendations)
- **Division-based reports** (Scoped In/Out by Division)
- **Other supporting tables**

Output saved as: **"Bidvest Scoping Tool Output.xlsx"**

---

## 📊 Part 2: PowerBI Setup

### Step 1: Import Data into PowerBI

1. Open **Power BI Desktop**
2. Click **Home** → **Get Data** → **Excel Workbook**
3. Navigate to `Bidvest Scoping Tool Output.xlsx`
4. Select **ALL** the following tables:

   **Core Data Tables:**
   - ☑ Full Input Table
   - ☑ Full Input Percentage
   
   **Reference Tables:**
   - ☑ Pack Number Company Table
   - ☑ FSLi Key Table
   
   **Scoping Tables:**
   - ☑ Scoping Control Table (**MOST IMPORTANT**)
   - ☑ Scoping Summary
   - ☑ Scoped In by Division
   - ☑ Scoped Out by Division
   - ☑ Threshold Configuration (if threshold scoping was used)

5. Click **Transform Data** (open Power Query Editor)

### Step 2: Create Relationships

In the **Model View** (left sidebar, middle icon):

**Primary Relationships:**
```
Pack Number Company Table[Pack Code] → Scoping Control Table[Pack Code] (Many-to-One)
Pack Number Company Table[Pack Code] → Full Input Table[Pack Code] (One-to-Many)
FSLi Key Table[FSLi] → Scoping Control Table[FSLi] (One-to-Many)
```

**Why Pack Code instead of Pack Name?**
- Pack Code is unique and consistent across all tables
- Pack Name can have variations or duplicates
- Both fields are available, but relationships use Pack Code
- You can still display Pack Name in visuals!

### Step 3: Create DAX Measures

Go to **Modeling** tab → **New Measure**

#### Measure 1: Total Packs (Excluding Consolidated)
```DAX
Total Packs = 
CALCULATE(
    DISTINCTCOUNT('Scoping Control Table'[Pack Code]),
    'Scoping Control Table'[Is Consolidated] = "No"
)
```

#### Measure 2: Scoped In Packs Count
```DAX
Scoped In Packs = 
CALCULATE(
    DISTINCTCOUNT('Scoping Control Table'[Pack Code]),
    'Scoping Control Table'[Scoping Status] = "Scoped In",
    'Scoping Control Table'[Is Consolidated] = "No"
)
```

#### Measure 3: Scoping Coverage %
```DAX
Scoping Coverage % = 
VAR ScopedTotal = 
    CALCULATE(
        SUM('Scoping Control Table'[Amount]),
        'Scoping Control Table'[Scoping Status] = "Scoped In",
        'Scoping Control Table'[Is Consolidated] = "No"
    )
VAR GrandTotal = 
    CALCULATE(
        SUM('Scoping Control Table'[Amount]),
        'Scoping Control Table'[Is Consolidated] = "No"
    )
RETURN
DIVIDE(ABS(ScopedTotal), ABS(GrandTotal), 0)
```

#### Measure 4: Scoping Coverage % by FSLi
```DAX
Coverage % by FSLi = 
VAR CurrentFSLi = SELECTEDVALUE('Scoping Control Table'[FSLi])
VAR ScopedAmount = 
    CALCULATE(
        SUM('Scoping Control Table'[Amount]),
        'Scoping Control Table'[FSLi] = CurrentFSLi,
        'Scoping Control Table'[Scoping Status] = "Scoped In",
        'Scoping Control Table'[Is Consolidated] = "No"
    )
VAR TotalAmount = 
    CALCULATE(
        SUM('Scoping Control Table'[Amount]),
        'Scoping Control Table'[FSLi] = CurrentFSLi,
        'Scoping Control Table'[Is Consolidated] = "No"
    )
RETURN
DIVIDE(ABS(ScopedAmount), ABS(TotalAmount), 0)
```

#### Measure 5: Coverage % by Division
```DAX
Coverage % by Division = 
VAR CurrentDivision = SELECTEDVALUE('Pack Number Company Table'[Division])
VAR ScopedAmount = 
    CALCULATE(
        SUM('Scoping Control Table'[Amount]),
        'Pack Number Company Table'[Division] = CurrentDivision,
        'Scoping Control Table'[Scoping Status] = "Scoped In",
        'Scoping Control Table'[Is Consolidated] = "No"
    )
VAR TotalAmount = 
    CALCULATE(
        SUM('Scoping Control Table'[Amount]),
        'Pack Number Company Table'[Division] = CurrentDivision,
        'Scoping Control Table'[Is Consolidated] = "No"
    )
RETURN
DIVIDE(ABS(ScopedAmount), ABS(TotalAmount), 0)
```

#### Measure 6: Untested %
```DAX
Untested % = 1 - [Scoping Coverage %]
```

---

## 🎨 Part 3: Create Scoping Dashboard

### Page 1: Scoping Control Panel

**Visual 1: KPI Cards** (Top row)
- Card: [Total Packs]
- Card: [Scoped In Packs]
- Card: [Scoping Coverage %] (format as percentage)
- Card: [Untested %] (format as percentage)

**Visual 2: Pack Selector** (Left sidebar)
- Type: **Slicer**
- Field: `Pack Number Company Table[Pack Name]`
- Settings: 
  - Multi-select: ✅ Enabled
  - Select All: ✅ Enabled
- Filter: `Pack Number Company Table[Is Consolidated]` = "No"

**Visual 3: FSLi Selector** (Left sidebar, below Pack)
- Type: **Slicer**
- Field: `Scoping Control Table[FSLi]`
- Settings: Multi-select enabled

**Visual 4: Division Filter** (Top row)
- Type: **Slicer**
- Field: `Pack Number Company Table[Division]`
- Settings: Multi-select enabled

**Visual 5: Scoping Status Table** (Main area)
- Type: **Table**
- Columns:
  1. `Pack Number Company Table[Pack Name]`
  2. `Pack Number Company Table[Pack Code]`
  3. `Pack Number Company Table[Division]`
  4. `Scoping Control Table[FSLi]`
  5. `Scoping Control Table[Amount]` (format as currency)
  6. `Scoping Control Table[Scoping Status]` ← **EDITABLE**
- Formatting:
  - Conditional formatting on Scoping Status:
    - "Scoped In" = Green background
    - "Not Scoped" = Gray background
    - "Scoped Out" = Red background

**Visual 6: Coverage by FSLi** (Bottom left)
- Type: **Clustered Bar Chart**
- Axis: `Scoping Control Table[FSLi]`
- Values: [Coverage % by FSLi]
- Data label: Show percentage

**Visual 7: Coverage by Division** (Bottom right)
- Type: **Donut Chart**
- Legend: `Pack Number Company Table[Division]`
- Values: [Coverage % by Division]

### Page 2: FSLi Analysis

**Visual 1: FSLi Coverage Matrix**
- Type: **Matrix**
- Rows: `Scoping Control Table[FSLi]`
- Columns: `Pack Number Company Table[Division]`
- Values: [Coverage % by FSLi]
- Formatting: Color scale (0% = Red, 100% = Green)

**Visual 2: FSLi Detail Table**
- Type: **Table**
- Columns:
  1. `Scoping Control Table[FSLi]`
  2. Total Amount (SUM of Amount, all packs)
  3. Scoped Amount (SUM of Amount where Scoped In)
  4. [Coverage % by FSLi]
  5. [Untested %]

### Page 3: Division Analysis

**Visual 1: Division Coverage Bars**
- Type: **Clustered Column Chart**
- Axis: `Pack Number Company Table[Division]`
- Values: [Coverage % by Division]

**Visual 2: Packs by Division Table**
- Type: **Table**
- Columns:
  1. `Pack Number Company Table[Division]`
  2. `Pack Number Company Table[Pack Name]`
  3. `Scoping Control Table[Scoping Status]`
  4. Total Amount per pack

---

## 🔄 Part 4: Manual Scoping Workflow

### Method 1: Edit Scoping Status Directly

1. Go to **Scoping Control Panel** page
2. Find the **Scoping Status Table** visual
3. Click on a row to select it
4. **Right-click** → **Edit**
5. Change **Scoping Status** value:
   - `"Not Scoped"` → `"Scoped In"`
   - `"Scoped In"` → `"Scoped Out"`
6. All KPIs and charts **update automatically**!

**Note:** This requires PowerBI Desktop with edit permissions. For published reports, use Method 2.

### Method 2: Use Slicers for Bulk Selection

1. Use **Pack Selector** slicer to select one or more packs
2. Use **FSLi Selector** to filter specific FSLIs
3. View the filtered **Scoping Status Table**
4. Create a calculated column or use bookmarks to mark these as "Scoped In"

### Method 3: Create Scoping Buttons (Advanced)

**Button: "Scope In Selected Packs"**

1. Insert → **Button**
2. Action → **Bookmark** → Create new bookmark after:
   - Selecting packs in slicer
   - Applying filter: Scoping Status = "Scoped In"
3. Use this to quickly scope in multiple packs

### Method 4: Use DAX to Create Virtual Scoping

Create a **calculated column** in Scoping Control Table:

```DAX
Manual Scoping Status = 
IF(
    'Scoping Control Table'[Pack Name] IN {"Pack 1", "Pack 2", "Pack 3"},
    "Scoped In",
    'Scoping Control Table'[Scoping Status]
)
```

Then use this column instead of the original Scoping Status in all visuals.

---

## 📤 Part 5: Export Results Back to Excel

### Option 1: Export Scoping Status Table

1. Click on **Scoping Status Table** visual
2. Click **... (More options)** → **Export data**
3. Choose **Excel** format
4. Save as `Scoping_Results.xlsx`
5. This file contains all scoping decisions

### Option 2: Use PowerBI Service (Published Reports)

1. Publish your report to PowerBI Service
2. Users can interact with slicers and filters
3. Use **Export to Excel** on any visual
4. Results can be imported back to source Excel workbook

### Option 3: Create Scoping Export Table

Create a new table page specifically for export:

**Visual: Scoping Export Table**
- Type: **Table**
- Columns:
  1. Pack Name
  2. Pack Code
  3. Division
  4. FSLi
  5. Amount
  6. Scoping Status (with your manual updates)
  7. Coverage %

Export this table to Excel for documentation.

---

## 🎯 Part 6: ISA 600 Compliance Reporting

### Required Analysis for ISA 600 (Revised)

1. **Component Scoping Coverage**
   - Use [Scoping Coverage %] measure
   - Target: Typically 60-80% coverage
   - Show by Division using [Coverage % by Division]

2. **FSLi Coverage by Component**
   - Use FSLi Analysis page
   - Matrix showing which FSLis are covered per division
   - Identify gaps in coverage

3. **Consolidated Entity Exclusion**
   - Consolidated pack automatically excluded from all calculations
   - Marked with `Is Consolidated = Yes` in Pack Number Company Table
   - Verify in any visual by adding Is Consolidated filter

4. **Division-Level Analysis**
   - Only Category 1 (Segment Tabs) are treated as divisions
   - Use Division Analysis page for coverage by division
   - Ensure key divisions have adequate coverage

5. **Untested Risk Assessment**
   - Use [Untested %] measure
   - Identify high-value untested FSLIs
   - Document rationale for scoping decisions

---

## 💡 Best Practices

### 1. Start with Threshold-Based Scoping
- Run VBA macro with threshold scoping for major FSLIs (Revenue, Total Assets)
- This provides a starting point based on materiality
- Then refine in PowerBI

### 2. Use Division Analysis
- Ensure each division has adequate coverage
- Don't rely solely on overall coverage %
- Check for division-specific risks

### 3. Exclude Consolidated Entity
- Always select the consolidated pack during VBA macro execution
- Verify `Is Consolidated = Yes` is set correctly
- Double-check it's excluded from coverage calculations

### 4. Document Scoping Decisions
- Export final Scoping Status Table
- Include rationale for packs scoped in/out
- Track changes over time

### 5. Refresh Data Regularly
- When source data changes, re-run VBA macro
- Refresh PowerBI data sources
- Review and update scoping decisions

---

## 🔧 Troubleshooting

### Issue: Pack Names not connecting in relationships
**Solution:** Use Pack Code for relationships, not Pack Name. Both fields are available in all tables.

### Issue: Consolidated pack appearing in scoping analysis
**Solution:** 
1. Check `Is Consolidated` column in Pack Number Company Table
2. Ensure your DAX measures filter `Is Consolidated = "No"`
3. Re-run VBA macro and correctly select consolidated entity

### Issue: Division showing "Not Categorized"
**Solution:** 
- Only tabs categorized as Category 1 (Segment Tabs) become divisions
- Other categories show "Not Categorized"
- This is correct behavior per requirements

### Issue: Scoping Status not updating
**Solution:** 
- In PowerBI Desktop, you can edit table values directly
- In PowerBI Service (published), use calculated columns or bookmarks
- Consider using Method 4 (DAX calculated column)

### Issue: Coverage % not calculating correctly
**Solution:**
1. Verify relationships are set up correctly (Pack Code based)
2. Check that `Is Consolidated = "No"` filter is applied in measures
3. Ensure amounts are numeric (not text)

---

## 📚 Additional Resources

- **POWERBI_COMPLETE_SETUP.md**: Original PowerBI setup guide
- **README.md**: Tool overview and features
- **DOCUMENTATION.md**: Complete VBA documentation
- **VBA_Modules/README.md**: Module-level documentation

---

## 🆘 Support

For questions or issues:
1. Review this guide
2. Check troubleshooting section
3. Verify VBA macro completed successfully
4. Check PowerBI relationships in Model View

---

**Last Updated:** November 2024  
**Version:** 3.1.0  
**Compatibility:** PowerBI Desktop (latest), Excel 2016+
