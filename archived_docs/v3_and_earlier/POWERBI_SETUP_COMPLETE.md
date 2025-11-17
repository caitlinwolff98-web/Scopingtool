# Enhanced Power BI Integration Guide
## Bidvest Scoping Tool - Complete Setup Instructions

This guide provides comprehensive step-by-step instructions for setting up Power BI with the Bidvest Scoping Tool output.

---

## Table of Contents

1. [Quick Setup (5 Minutes)](#quick-setup)
2. [Detailed Setup Instructions](#detailed-setup)
3. [Data Model & Relationships](#data-model-relationships)
4. [DAX Measures Library](#dax-measures-library)
5. [Dashboard Creation](#dashboard-creation)
6. [Auto-Refresh Configuration](#auto-refresh-configuration)
7. [Troubleshooting](#troubleshooting)

---

## Quick Setup

### Prerequisites
- Power BI Desktop installed (latest version)
- "Bidvest Scoping Tool Output.xlsx" file generated from the VBA macro

### 5-Minute Setup

1. **Open Power BI Desktop**
   - Launch Power BI Desktop
   - Click "Get Data" → "Excel Workbook"

2. **Select the Output File**
   - Navigate to: `Bidvest Scoping Tool Output.xlsx`
   - Select ALL tables shown in Navigator:
     - ✅ Full Input Table
     - ✅ Full Input Percentage
     - ✅ Pack Number Company Table
     - ✅ FSLi Key Table
     - ✅ Scoping Summary (NEW!)
     - ✅ Threshold Configuration (if applicable)
     - ✅ (Optional) Journals, Console, Discontinued tables
   - Click "Transform Data"

3. **Transform Data (Power Query)**
   - For "Full Input Table":
     - Right-click column header "Pack Name" → "Unpivot Other Columns"
     - Rename "Attribute" → "FSLi"
     - Rename "Value" → "Amount"
     - Filter out null values in "Amount" column
   - Repeat for other data tables (Full Input Percentage, etc.)
   - Click "Close & Apply"

4. **Create Relationships**
   - Click "Model" view
   - Create these relationships:
     - `Pack Number Company Table[Pack Code]` → `Full Input Table[Pack Code]` (Many-to-One)
     - `FSLi Key Table[FSLi]` → `Full Input Table[FSLi]` (Many-to-One)
     - `Scoping Summary[Pack Code]` → `Pack Number Company Table[Pack Code]` (One-to-One)

5. **Create Basic Measures** (see DAX section below)

6. **Build Dashboard** (see Dashboard section below)

---

## Detailed Setup Instructions

### Step 1: Import Data with Optimal Settings

```
File Location: [Same folder as source workbook]/Bidvest Scoping Tool Output.xlsx
Import Method: Excel Workbook
Tables to Import: ALL TABLES
```

**Important:** The output file is named consistently as "Bidvest Scoping Tool Output.xlsx" which allows Power BI to auto-refresh without breaking connections.

### Step 2: Transform Pack Company Table

The Pack Number Company Table establishes the foundation for all relationships.

**Power Query Steps:**
```m
// Ensure Pack Code is text type
= Table.TransformColumnTypes(Source,{{"Pack Code", type text}})

// Remove any duplicates
= Table.Distinct(#"Changed Type", {"Pack Code"})

// Trim whitespace
= Table.TransformColumns(#"Removed Duplicates",{{"Pack Code", Text.Trim}})
```

### Step 3: Transform Full Input Table (Critical!)

This transformation converts the wide format to long format for proper analysis.

**Power Query Steps:**
```m
// Step 1: Select Pack Name column
// Right-click "Pack Name" → Unpivot Other Columns

// Step 2: Rename columns
= Table.RenameColumns(#"Unpivoted Columns",{{"Attribute", "FSLi"}, {"Value", "Amount"}})

// Step 3: Add Pack Code column if not present
= Table.AddColumn(#"Renamed Columns", "Pack Code", each 
    // Extract pack code from Pack Number Company Table based on Pack Name
)

// Step 4: Filter nulls and zeros (optional)
= Table.SelectRows(#"Added Pack Code", each [Amount] <> null and [Amount] <> 0)

// Step 5: Set data types
= Table.TransformColumnTypes(#"Filtered Rows",{
    {"Pack Name", type text},
    {"FSLi", type text},
    {"Amount", type number},
    {"Pack Code", type text}
})
```

### Step 4: Transform FSLi Key Table

**Power Query Steps:**
```m
// Ensure FSLi is text and trimmed
= Table.TransformColumns(Source,{{"FSLi", Text.Trim, type text}})

// Mark as dimension table
// (Right-click table → Properties → check "Dimension Table")
```

---

## Data Model & Relationships

### Relationship Structure

```
Pack Number Company Table (Dimension)
    ↓ (One-to-Many)
Full Input Table (Fact)
    ↑ (Many-to-One)
FSLi Key Table (Dimension)

Scoping Summary (Dimension)
    ↓ (One-to-One)
Pack Number Company Table
```

### Create Relationships in Model View

1. **Pack Relationships:**
   ```
   FROM: Pack Number Company Table[Pack Code]
   TO:   Full Input Table[Pack Code]
   TYPE: One-to-Many
   CROSS-FILTER: Single
   ACTIVE: Yes
   ```

2. **FSLi Relationships:**
   ```
   FROM: FSLi Key Table[FSLi]
   TO:   Full Input Table[FSLi]
   TYPE: One-to-Many
   CROSS-FILTER: Single
   ACTIVE: Yes
   ```

3. **Scoping Relationships:**
   ```
   FROM: Scoping Summary[Pack Code]
   TO:   Pack Number Company Table[Pack Code]
   TYPE: One-to-One
   CROSS-FILTER: Both
   ACTIVE: Yes
   ```

**Fix for Pack Name Connection Issue:**
If Pack Names are not connecting properly:
1. Ensure Pack Code is used for relationships (NOT Pack Name)
2. Pack Code must be TEXT type in all tables
3. Trim whitespace: `Text.Trim([Pack Code])`
4. Check for leading/trailing spaces or hidden characters

---

## DAX Measures Library

### Basic Measures

```dax
// Total Amount
Total Amount = SUM('Full Input Table'[Amount])

// Absolute Amount (for percentages)
Total Absolute Amount = SUMX('Full Input Table', ABS([Amount]))

// Count of Packs
Pack Count = DISTINCTCOUNT('Full Input Table'[Pack Code])

// Count of FSLis
FSLi Count = DISTINCTCOUNT('Full Input Table'[FSLi])
```

### Scoping Measures

```dax
// Packs Scoped In
Packs Scoped In = 
    CALCULATE(
        DISTINCTCOUNT('Scoping Summary'[Pack Code]),
        'Scoping Summary'[Scoped In] = "Yes"
    )

// Packs Pending Review
Packs Pending Review = 
    CALCULATE(
        DISTINCTCOUNT('Scoping Summary'[Pack Code]),
        'Scoping Summary'[Scoped In] = "No"
    )

// Scoping Coverage %
Scoping Coverage % = 
    DIVIDE(
        [Packs Scoped In],
        DISTINCTCOUNT('Pack Number Company Table'[Pack Code]),
        0
    )

// Suggested for Scope Count
Suggested for Scope = 
    CALCULATE(
        DISTINCTCOUNT('Scoping Summary'[Pack Code]),
        'Scoping Summary'[Suggested for Scope] = "Yes"
    )
```

### Threshold-Based Measures

```dax
// Amount Above Threshold
Amount Above Threshold = 
    VAR ThresholdValue = 300000000 // Set your threshold
    RETURN
    CALCULATE(
        [Total Absolute Amount],
        FILTER(
            'Full Input Table',
            ABS([Amount]) >= ThresholdValue
        )
    )

// Packs Above Threshold
Packs Above Threshold = 
    VAR ThresholdValue = 300000000 // Set your threshold
    RETURN
    CALCULATE(
        DISTINCTCOUNT('Full Input Table'[Pack Code]),
        FILTER(
            'Full Input Table',
            ABS([Amount]) >= ThresholdValue
        )
    )

// Dynamic Threshold (use slicer)
Dynamic Threshold Amount = 
    VAR SelectedThreshold = VALUES('Threshold Slicer'[Threshold])
    RETURN
    CALCULATE(
        [Total Absolute Amount],
        FILTER(
            'Full Input Table',
            ABS([Amount]) >= SelectedThreshold
        )
    )
```

### Percentage Measures

```dax
// Pack % of Total
Pack % of Total = 
    DIVIDE(
        [Total Absolute Amount],
        CALCULATE(
            [Total Absolute Amount],
            ALL('Pack Number Company Table')
        ),
        0
    )

// FSLi % of Total
FSLi % of Total = 
    DIVIDE(
        [Total Absolute Amount],
        CALCULATE(
            [Total Absolute Amount],
            ALL('FSLi Key Table')
        ),
        0
    )

// Coverage % by Division
Coverage by Division = 
    DIVIDE(
        CALCULATE(
            [Packs Scoped In],
            ALLEXCEPT('Pack Number Company Table', 'Pack Number Company Table'[Division])
        ),
        CALCULATE(
            [Pack Count],
            ALLEXCEPT('Pack Number Company Table', 'Pack Number Company Table'[Division])
        ),
        0
    )
```

### Conditional Formatting Measures

```dax
// Red/Amber/Green Status
RAG Status = 
    VAR Coverage = [Scoping Coverage %]
    RETURN
    SWITCH(
        TRUE(),
        Coverage >= 0.8, "Green",
        Coverage >= 0.6, "Amber",
        "Red"
    )

// Risk Flag
Risk Flag = 
    IF(
        [Total Absolute Amount] > 500000000,
        "High Risk",
        IF(
            [Total Absolute Amount] > 100000000,
            "Medium Risk",
            "Low Risk"
        )
    )
```

---

## Dashboard Creation

### Dashboard Layout

**Page 1: Executive Summary**
- KPI Cards: Total Packs, Scoped In, Coverage %
- Pie Chart: Scoping Status (Scoped vs Pending)
- Bar Chart: Top 10 Packs by Amount
- Table: Scoping Summary

**Page 2: Pack Analysis**
- Matrix: Pack × FSLi with amounts
- Slicers: Division, Scoped In Status
- Bar Chart: Pack amounts
- Conditional formatting on amounts

**Page 3: FSLi Analysis**
- Matrix: FSLi × Pack with amounts
- Slicers: Statement Type, Is Total
- Line Chart: FSLi trends
- Table: Top FSLis by total amount

**Page 4: Threshold Analysis**
- Slicer: Threshold Value (100M, 300M, 500M)
- Card: Packs Above Threshold
- Table: Packs exceeding threshold
- Scatter Plot: Pack vs FSLi amounts

### Visual Recommendations

1. **KPI Cards**
   - Total Packs: `[Pack Count]`
   - Scoped In: `[Packs Scoped In]`
   - Coverage: `[Scoping Coverage %]`

2. **Pie Chart - Scoping Status**
   - Legend: Scoping Summary[Scoped In]
   - Values: DISTINCTCOUNT(Pack Code)

3. **Matrix - Pack × FSLi**
   - Rows: Pack Number Company Table[Pack Name]
   - Columns: FSLi Key Table[FSLi]
   - Values: [Total Amount]
   - Conditional Formatting: Background color by [Total Absolute Amount]

4. **Table - Scoping Summary**
   - Columns: Pack Code, Pack Name, Scoped In, Suggested for Scope
   - Conditional Formatting: Highlight "Yes" in green

---

## Auto-Refresh Configuration

### Option 1: File-Based Refresh (Recommended)

Since the output file is always named "Bidvest Scoping Tool Output.xlsx" and saved in the same location:

1. **Save Power BI File in Same Folder**
   - Save your .pbix file in the same directory as the Excel output
   - Use relative path if possible

2. **Set Up Scheduled Refresh**
   - File → Options → Data Load
   - Enable: "Background data refresh"
   - Set refresh interval (e.g., every hour)

3. **Refresh Options**
   - Manual: Click "Refresh" button
   - Automatic: On file open (set in Options)
   - Scheduled: Via Power BI Service (if publishing)

### Option 2: Power BI Service Refresh

1. **Publish to Power BI Service**
   - Click "Publish" → Select workspace
   
2. **Configure Gateway** (for local files)
   - Install Power BI Gateway on machine with Excel file
   - Configure data source connection
   
3. **Schedule Refresh**
   - In Power BI Service, go to dataset settings
   - Configure refresh schedule (daily, hourly, etc.)

### Option 3: One-Click Refresh

Create a simple refresh button:

1. Add a "Refresh" button to Page 1
2. Action: "Refresh"
3. Tooltip: "Click to refresh data from Excel"

---

## Troubleshooting

### Issue: Pack Names Not Connecting

**Solution:**
- Use Pack Code for relationships, NOT Pack Name
- Ensure Pack Code is TEXT type in all tables
- Check for whitespace: Use `Text.Trim([Pack Code])` in Power Query

### Issue: Relationships Not Working

**Solution:**
1. Check data types match (both text or both number)
2. Verify no duplicates in dimension tables
3. Ensure cardinality is correct (One-to-Many)
4. Check cross-filter direction

### Issue: FSLi Not Showing in Visuals

**Solution:**
- Verify FSLi Key Table is connected to Full Input Table
- Check that FSLi names match exactly (case-sensitive)
- Use Text.Trim() to remove whitespace

### Issue: Measures Showing Wrong Values

**Solution:**
- Check filter context in measures
- Verify relationships are active
- Use ALLEXCEPT() or CALCULATE() to control context

### Issue: File Not Auto-Refreshing

**Solution:**
1. Verify file path hasn't changed
2. Ensure Excel file is not open when refreshing
3. Check Power Query connection settings
4. Use "Edit Queries" → "Data Source Settings" to update path

---

## Best Practices

### Performance Optimization
1. Remove unnecessary columns in Power Query
2. Use Integer data types where possible
3. Avoid calculated columns; use measures instead
4. Disable auto date/time hierarchy
5. Create aggregation tables for large datasets

### Data Refresh Strategy
1. Close Excel file before Power BI refresh
2. Save both files in same directory
3. Use consistent naming convention
4. Schedule refreshes during off-hours
5. Test refresh before publishing

### Dashboard Design
1. Use consistent color scheme (Bidvest brand colors)
2. Add clear labels and tooltips
3. Include filters on all pages
4. Use drill-through for detailed analysis
5. Add bookmarks for different views

---

## Advanced Features

### Dynamic Threshold Slicer

Create a parameter table:
```dax
Threshold Values = 
DATATABLE(
    "Threshold", INTEGER,
    "Label", STRING,
    {
        {100000000, "100M"},
        {300000000, "300M"},
        {500000000, "500M"},
        {1000000000, "1B"}
    }
)
```

Use in visual as slicer, reference in measures:
```dax
Dynamic Scoped Packs = 
VAR SelectedThreshold = SELECTEDVALUE('Threshold Values'[Threshold], 300000000)
RETURN
CALCULATE(
    [Pack Count],
    FILTER('Full Input Table', ABS([Amount]) >= SelectedThreshold)
)
```

### What-If Analysis

Create What-If parameter for target coverage:
```
Home → Modeling → New Parameter
Name: Target Coverage
Minimum: 0
Maximum: 1
Increment: 0.05
Default: 0.8
```

Use in measure:
```dax
Packs Needed for Target = 
VAR Target = [Target Coverage Value]
VAR TotalPacks = [Pack Count]
RETURN
TotalPacks * Target
```

---

## Support

For issues or questions:
1. Review this guide thoroughly
2. Check the Scoping Summary sheet in Excel
3. Verify data in source tables
4. Test with sample data first
5. Consult Power BI documentation

---

**Version:** 2.0  
**Last Updated:** [Current Date]  
**Compatible with:** Bidvest Scoping Tool Output.xlsx (standardized naming)
