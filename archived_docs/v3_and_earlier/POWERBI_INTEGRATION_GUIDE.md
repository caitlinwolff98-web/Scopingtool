# Power BI Integration Guide for TGK Scoping Tool

## Overview

This guide provides detailed instructions for integrating the Excel output from the TGK Scoping Tool with Power BI to create interactive scoping dashboards and perform threshold-based analysis.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Data Import](#data-import)
3. [Data Transformation](#data-transformation)
4. [Data Model Setup](#data-model-setup)
5. [DAX Measures](#dax-measures)
6. [Report Creation](#report-creation)
7. [Scoping Workflows](#scoping-workflows)
8. [Advanced Features](#advanced-features)

---

## Prerequisites

### Software Requirements
- Power BI Desktop (latest version)
- Output workbook from TGK Scoping Tool
- Basic understanding of Power BI interface

### Knowledge Requirements
- Power Query basics
- DAX fundamentals (helpful but not required)
- Excel table concepts

---

## Data Import

### Step 1: Open Power BI Desktop

1. Launch Power BI Desktop
2. Close any existing reports or start with a blank report

### Step 2: Import Excel Data

1. Click **Home** → **Get Data** → **Excel Workbook**
2. Navigate to the output workbook created by the TGK Scoping Tool
3. Click **Open**

### Step 3: Select Tables

In the Navigator window, select the following tables:

**Required Tables:**
- ☑ Full Input Table
- ☑ Full Input Percentage
- ☑ Pack Number Company Table
- ☑ FSLi Key Table

**Optional Tables (if available):**
- ☑ Journals Table
- ☑ Journals Percentage
- ☑ Full Consol Table
- ☑ Full Consol Percentage
- ☑ Discontinued Table
- ☑ Discontinued Percentage

### Step 4: Transform Data

1. Click **Transform Data** (do NOT click Load yet)
2. This opens Power Query Editor

---

## Data Transformation

### Transform 1: Unpivot Full Input Table

**Purpose:** Convert wide format to long format for better analysis

1. In Power Query Editor, select **Full Input Table**
2. Right-click the **Pack** column → **Unpivot Other Columns**
3. Rename columns:
   - **Attribute** → **FSLi**
   - **Value** → **Amount**
4. Remove rows with null values:
   - Click filter on **Amount** → uncheck **(null)**

**Result:**
```
| Pack      | FSLi          | Amount    |
|-----------|---------------|-----------|
| Entity 1  | Revenue       | 1,000,000 |
| Entity 1  | Cost of Sales | 600,000   |
| Entity 2  | Revenue       | 2,000,000 |
```

### Transform 2: Clean FSLi Names

**Purpose:** Remove metadata tags for proper matching

1. Still in **Full Input Table**
2. Select **FSLi** column
3. Click **Transform** → **Replace Values**
   - Value to Find: ` (Total)`
   - Replace With: *(empty)*
4. Repeat for:
   - ` (Subtotal)`
   - ` (total)`
   - ` (subtotal)`

### Transform 3: Unpivot Full Input Percentage

1. Select **Full Input Percentage** query
2. Right-click **Pack** column → **Unpivot Other Columns**
3. Rename columns:
   - **Attribute** → **FSLi**
   - **Value** → **Percentage**
4. Clean FSLi names (same as Transform 2)

### Transform 4: Repeat for Other Tables

Repeat Transforms 1-3 for:
- Journals Table & Journals Percentage
- Full Consol Table & Full Consol Percentage
- Discontinued Table & Discontinued Percentage

### Transform 5: Pack Number Company Table

No transformation needed - load as-is

### Transform 6: FSLi Key Table

No transformation needed - load as-is

### Step 5: Close & Apply

1. Click **Home** → **Close & Apply**
2. Wait for data to load into Power BI

---

## Data Model Setup

### Step 1: Open Model View

1. Click **Model** view (icon on left sidebar)
2. You'll see all imported tables

### Step 2: Create Relationships

#### Relationship 1: Pack to Input Table

1. Drag **Pack Number Company Table[Pack Name]** 
2. Drop on **Full Input Table[Pack]**
3. Verify relationship:
   - From: Pack Number Company Table[Pack Name]
   - To: Full Input Table[Pack]
   - Cardinality: One to Many
   - Cross filter direction: Single

#### Relationship 2: Pack to Input Percentage

1. Drag **Pack Number Company Table[Pack Name]**
2. Drop on **Full Input Percentage[Pack]**
3. Configure same as above

#### Relationship 3: FSLi to Input Table

1. Drag **FSLi Key Table[FSLi]**
2. Drop on **Full Input Table[FSLi]**
3. Configure:
   - Cardinality: One to Many
   - Cross filter direction: Both

#### Relationship 4: FSLi to Input Percentage

1. Drag **FSLi Key Table[FSLi]**
2. Drop on **Full Input Percentage[FSLi]**
3. Configure same as Relationship 3

#### Additional Relationships

Repeat for other table pairs:
- Pack Number Company → Journals Table
- Pack Number Company → Full Consol Table
- Pack Number Company → Discontinued Table
- FSLi Key Table → Journals Table
- FSLi Key Table → Full Consol Table
- FSLi Key Table → Discontinued Table

### Step 3: Verify Model

Your data model should look like a star schema:
- **Center:** Pack Number Company Table & FSLi Key Table
- **Outer:** Data tables (Input, Journal, Console, Discontinued)

---

## DAX Measures

### Create Measure Table

1. Right-click in Fields pane → **New Table**
2. Name: `_Measures`
3. Formula: `_Measures = { "Scoping Measures" }`

### Measure 1: Selected Packs (for scoping)

```DAX
Selected Packs = 
CALCULATE(
    DISTINCTCOUNT('Full Input Table'[Pack]),
    ALLSELECTED('Full Input Table'[Pack])
)
```

**Purpose:** Count how many packs are currently selected/scoped in

### Measure 2: Total Amount

```DAX
Total Amount = 
SUM('Full Input Table'[Amount])
```

**Purpose:** Sum all amounts for selected packs and FSLis

### Measure 3: Total Coverage Amount

```DAX
Coverage Amount = 
VAR SelectedPacks = VALUES('Pack Number Company Table'[Pack Name])
RETURN
CALCULATE(
    SUM('Full Input Table'[Amount]),
    KEEPFILTERS(SelectedPacks)
)
```

**Purpose:** Calculate total amount for scoped-in packs

### Measure 4: Coverage Percentage

```DAX
Coverage % = 
VAR ScopedAmount = [Coverage Amount]
VAR TotalAmount = 
    CALCULATE(
        SUM('Full Input Table'[Amount]),
        ALL('Pack Number Company Table'),
        ALL('FSLi Key Table')
    )
RETURN
DIVIDE(ScopedAmount, TotalAmount, 0)
```

**Purpose:** Calculate what percentage of total is covered by scope

### Measure 5: Untested Percentage

```DAX
Untested % = 1 - [Coverage %]
```

**Purpose:** Calculate remaining untested percentage

### Measure 6: Threshold Check

```DAX
Meets Threshold = 
VAR ThresholdValue = 300000000  -- $300M default
VAR CurrentAmount = [Total Amount]
RETURN
IF(CurrentAmount > ThresholdValue, "Yes", "No")
```

**Purpose:** Check if amount meets threshold for automatic scoping

### Measure 7: Pack Coverage Count

```DAX
Packs Scoped In = 
CALCULATE(
    DISTINCTCOUNT('Full Input Table'[Pack]),
    'Full Input Table'[Amount] > 0
)
```

**Purpose:** Count distinct packs in current scope

### Measure 8: FSLi Coverage Count

```DAX
FSLis Scoped In = 
CALCULATE(
    DISTINCTCOUNT('Full Input Table'[FSLi]),
    'Full Input Table'[Amount] > 0
)
```

**Purpose:** Count distinct FSLis in current scope

### Measure 9: Average Percentage

```DAX
Average Coverage % = 
AVERAGE('Full Input Percentage'[Percentage])
```

**Purpose:** Calculate average percentage coverage

### Measure 10: Threshold Parameter (Dynamic)

```DAX
Threshold Value = 
VAR DefaultThreshold = 300000000
VAR SelectedThreshold = SELECTEDVALUE('Threshold Table'[Threshold], DefaultThreshold)
RETURN
SelectedThreshold
```

**Purpose:** Allow dynamic threshold selection

### Create Threshold Table (for dynamic threshold)

1. **Home** → **Enter Data**
2. Create table with one column:
   ```
   | Threshold     |
   |---------------|
   | 100,000,000   |
   | 200,000,000   |
   | 300,000,000   |
   | 500,000,000   |
   | 1,000,000,000 |
   ```
3. Name: `Threshold Table`
4. Click **Load**

---

## Report Creation

### Report Page 1: Coverage Dashboard

#### Visual 1: Coverage Card

- **Type:** Card
- **Field:** `_Measures[Coverage %]`
- **Format:** Percentage, 2 decimals
- **Title:** "Total Coverage %"

#### Visual 2: Untested Card

- **Type:** Card
- **Field:** `_Measures[Untested %]`
- **Format:** Percentage, 2 decimals
- **Title:** "Untested %"

#### Visual 3: Packs Scoped In Card

- **Type:** Card
- **Field:** `_Measures[Packs Scoped In]`
- **Title:** "Packs in Scope"

#### Visual 4: Coverage Gauge

- **Type:** Gauge
- **Value:** `_Measures[Coverage %]`
- **Target:** Set to 80% or desired coverage target
- **Title:** "Coverage Progress"

#### Visual 5: Pack Selector

- **Type:** Slicer
- **Field:** `Pack Number Company Table[Pack Name]`
- **Settings:** 
  - Slicer type: List
  - Multi-select: ☑ On
  - Select All: ☑ On

#### Visual 6: Division Filter

- **Type:** Slicer
- **Field:** `Pack Number Company Table[Division]`
- **Settings:** Dropdown style

#### Visual 7: FSLi Selector

- **Type:** Slicer
- **Field:** `FSLi Key Table[FSLi]`
- **Settings:**
  - Slicer type: List
  - Multi-select: ☑ On
  - Search: ☑ On

#### Visual 8: Details Table

- **Type:** Table
- **Columns:**
  1. `Pack Number Company Table[Pack Name]`
  2. `Pack Number Company Table[Division]`
  3. `Full Input Table[FSLi]`
  4. `_Measures[Total Amount]` (Format: Currency)
  5. `_Measures[Average Coverage %]` (Format: Percentage)

### Report Page 2: Threshold Scoping

#### Visual 1: Threshold Selector

- **Type:** Slicer
- **Field:** `Threshold Table[Threshold]`
- **Format:** Currency
- **Title:** "Select Threshold"

#### Visual 2: FSLi for Threshold

- **Type:** Slicer
- **Field:** `FSLi Key Table[FSLi]`
- **Settings:** Single select
- **Title:** "Select FSLi to Apply Threshold"

#### Visual 3: Packs Meeting Threshold

- **Type:** Table
- **Columns:**
  1. `Pack Number Company Table[Pack Name]`
  2. `Full Input Table[FSLi]`
  3. `Full Input Table[Amount]`
  4. `_Measures[Meets Threshold]`
- **Filter:** `_Measures[Meets Threshold]` = "Yes"
- **Title:** "Packs Meeting Threshold"

#### Visual 4: Threshold Impact Card

- **Type:** Card
- **Field:** `_Measures[Packs Scoped In]`
- **Title:** "Packs Scoped by Threshold"

#### Visual 5: Threshold Coverage

- **Type:** Donut Chart
- **Legend:** Create measure:
  ```DAX
  Scope Status = 
  IF([Total Amount] > [Threshold Value], "Scoped In", "Not Scoped")
  ```
- **Values:** `_Measures[Total Amount]`

### Report Page 3: Manual Selection

#### Visual 1: Pack Picker

- **Type:** Slicer (Chiclet style if available)
- **Field:** `Pack Number Company Table[Pack Name]`
- **Settings:** Multi-select enabled

#### Visual 2: FSLi Picker

- **Type:** Slicer (Chiclet style if available)
- **Field:** `FSLi Key Table[FSLi]`
- **Settings:** Multi-select enabled

#### Visual 3: Current Selection Table

- **Type:** Table
- **Columns:**
  1. `Pack Number Company Table[Pack Name]`
  2. `Full Input Table[FSLi]`
  3. `Full Input Table[Amount]`
  4. `Full Input Percentage[Percentage]`

#### Visual 4: Selection Summary

- **Type:** Multi-row Card
- **Fields:**
  - `_Measures[Packs Scoped In]`
  - `_Measures[FSLis Scoped In]`
  - `_Measures[Coverage %]`
  - `_Measures[Untested %]`

### Report Page 4: Division Analysis

#### Visual 1: Division Slicer

- **Type:** Slicer (Dropdown)
- **Field:** `Pack Number Company Table[Division]`

#### Visual 2: Packs by Division

- **Type:** Bar Chart
- **Axis:** `Pack Number Company Table[Division]`
- **Values:** `DISTINCTCOUNT(Pack Number Company Table[Pack Name])`
- **Title:** "Pack Count by Division"

#### Visual 3: Coverage by Division

- **Type:** Clustered Column Chart
- **Axis:** `Pack Number Company Table[Division]`
- **Values:** 
  - `_Measures[Coverage Amount]`
  - Total Amount (from all packs)
- **Title:** "Amount by Division"

#### Visual 4: Division Details

- **Type:** Matrix
- **Rows:** 
  - `Pack Number Company Table[Division]`
  - `Pack Number Company Table[Pack Name]`
- **Values:**
  - `_Measures[Total Amount]`
  - `_Measures[Coverage %]`

---

## Scoping Workflows

### Workflow 1: Threshold-Based Automatic Scoping

**Objective:** Automatically scope in packs that meet specific FSLi thresholds

**Steps:**

1. **Navigate to "Threshold Scoping" page**

2. **Select FSLi for threshold**
   - Click FSLi slicer
   - Select one FSLi (e.g., "Net Revenue")

3. **Set threshold amount**
   - Use threshold slicer
   - Select value (e.g., $300,000,000)

4. **Review packs meeting threshold**
   - Check "Packs Meeting Threshold" table
   - Verify amounts exceed threshold

5. **Apply scope**
   - Note pack names meeting threshold
   - These packs will be automatically included in scope

6. **Check coverage impact**
   - Review "Threshold Impact Card"
   - Check coverage percentage increase

7. **Export scoped packs**
   - Right-click table → Export Data
   - Save for documentation

### Workflow 2: Manual Pack and FSLi Selection

**Objective:** Manually select specific packs and FSLis for scoping

**Steps:**

1. **Navigate to "Manual Selection" page**

2. **Select specific packs**
   - Use Pack Picker slicer
   - Click to select/deselect packs
   - Multi-select enabled

3. **Select specific FSLis**
   - Use FSLi Picker slicer
   - Choose FSLis of interest
   - Can select multiple

4. **Review current selection**
   - Check "Current Selection Table"
   - Verify amounts and percentages

5. **Monitor coverage**
   - Watch "Selection Summary" card
   - Track coverage % as you select

6. **Refine selection**
   - Add/remove packs or FSLis
   - Optimize for target coverage

7. **Document selection**
   - Export table
   - Screenshot for records

### Workflow 3: Complete Pack Scoping

**Objective:** Include all FSLis for selected packs

**Steps:**

1. **Navigate to "Coverage Dashboard"**

2. **Select packs**
   - Use Pack Selector slicer
   - Select one or more packs

3. **Leave FSLi filter clear**
   - Don't select any specific FSLis
   - This includes ALL FSLis for selected packs

4. **Review details**
   - Check Details Table
   - Verify all FSLis are included

5. **Check coverage**
   - Review Coverage Card
   - Verify percentage increase

6. **Compare by division**
   - Use Division Filter
   - Analyze coverage by segment

### Workflow 4: Hybrid Approach

**Objective:** Combine threshold and manual selection

**Steps:**

1. **Start with threshold scoping**
   - Select key FSLis (Revenue, Total Assets)
   - Set appropriate thresholds
   - Note automatically scoped packs

2. **Switch to manual selection**
   - Navigate to Manual Selection page
   - Pre-selected packs from threshold visible

3. **Add specific FSLis**
   - For packs not meeting threshold
   - Select specific FSLis of interest
   - E.g., select "Inventory" for certain packs

4. **Review combined coverage**
   - Check total Coverage %
   - Verify meets audit requirements

5. **Document methodology**
   - Export both threshold and manual selections
   - Note rationale for each

---

## Advanced Features

### Feature 1: Bookmark-Based Scoping Scenarios

**Purpose:** Save different scoping scenarios

**Setup:**

1. **Create Scenario 1:**
   - Select packs and FSLis
   - Click **View** → **Bookmarks** → **Add**
   - Name: "Scenario 1 - High Revenue"

2. **Create Scenario 2:**
   - Change selections
   - Add another bookmark
   - Name: "Scenario 2 - High Assets"

3. **Create Scenario 3:**
   - Different selection
   - Add bookmark
   - Name: "Scenario 3 - Manual Selection"

4. **Add Bookmark Navigator:**
   - Insert buttons for each bookmark
   - Configure button actions to navigate bookmarks

### Feature 2: Dynamic Threshold with What-If Parameter

**Purpose:** Allow real-time threshold adjustment

**Setup:**

1. **Create What-If Parameter:**
   - **Modeling** → **New Parameter**
   - Name: "Threshold"
   - Data type: Whole number
   - Minimum: 0
   - Maximum: 1,000,000,000
   - Increment: 10,000,000
   - Default: 300,000,000
   - Add slicer: ☑ Yes

2. **Update Measures:**
   ```DAX
   Dynamic Threshold Check = 
   VAR ThresholdValue = SELECTEDVALUE('Threshold'[Threshold Value], 300000000)
   VAR CurrentAmount = [Total Amount]
   RETURN
   IF(CurrentAmount > ThresholdValue, "Yes", "No")
   ```

3. **Add to Report:**
   - Place threshold slicer prominently
   - Link to threshold visuals

### Feature 3: Coverage History Tracking

**Purpose:** Track coverage changes over time

**Setup:**

1. **Create History Table:**
   ```DAX
   Coverage History = 
   DATATABLE(
       "Date", DATETIME,
       "Coverage %", DOUBLE,
       "Packs Scoped", INTEGER,
       "Scenario", STRING,
       {
           {"2024-01-01", 0.65, 10, "Initial"},
           {"2024-01-15", 0.72, 12, "After Threshold"},
           {"2024-02-01", 0.85, 15, "Final"}
       }
   )
   ```

2. **Create Line Chart:**
   - Axis: Date
   - Values: Coverage %
   - Legend: Scenario

3. **Update Manually:**
   - Edit table in Power Query
   - Add new rows as scenarios are finalized

### Feature 4: FSLi Importance Weighting

**Purpose:** Weight certain FSLis higher for coverage calculation

**Setup:**

1. **Create FSLi Weight Table:**
   ```DAX
   FSLi Weights = 
   DATATABLE(
       "FSLi", STRING,
       "Weight", DOUBLE,
       {
           {"Revenue", 2.0},
           {"Total Assets", 2.0},
           {"Net Profit", 1.5},
           {"Inventory", 1.0}
       }
   )
   ```

2. **Create Weighted Coverage Measure:**
   ```DAX
   Weighted Coverage % = 
   VAR WeightedScoped = 
       SUMX(
           'Full Input Table',
           [Total Amount] * 
           RELATED('FSLi Weights'[Weight])
       )
   VAR WeightedTotal = 
       CALCULATE(
           SUMX(
               'Full Input Table',
               [Total Amount] * 
               RELATED('FSLi Weights'[Weight])
           ),
           ALL()
       )
   RETURN
   DIVIDE(WeightedScoped, WeightedTotal, 0)
   ```

### Feature 5: Drill-Through for Pack Details

**Purpose:** Deep dive into specific pack details

**Setup:**

1. **Create Drill-Through Page:**
   - Add new report page
   - Name: "Pack Details"

2. **Add Drill-Through Field:**
   - Drag `Pack Number Company Table[Pack Name]` to Drill-through wells

3. **Add Detail Visuals:**
   - Table with all FSLis for the pack
   - Amounts and percentages
   - Comparison to average
   - Historical trends (if available)

4. **Configure Drill-Through:**
   - Right-click any pack in main reports
   - Select "Drill through" → "Pack Details"

### Feature 6: Export to Excel for Documentation

**Purpose:** Create audit documentation

**Setup:**

1. **Create Documentation Page:**
   - Table with all scoped packs and FSLis
   - Coverage summary
   - Threshold applied
   - Date and user

2. **Export Options:**
   - Right-click visual → Export Data
   - Choose format (Excel recommended)
   - Save with descriptive name

3. **Automation (Pro/Premium):**
   - Set up scheduled export
   - Email to stakeholders
   - Archive for compliance

### Feature 7: Conditional Formatting

**Purpose:** Highlight important information

**Setup:**

1. **Format Coverage % Card:**
   - Select card
   - Conditional formatting
   - Rules:
     - < 60%: Red
     - 60-80%: Yellow
     - > 80%: Green

2. **Format Amount Columns:**
   - Select table column
   - Conditional formatting
   - Data bars or color scales
   - Highlight top/bottom values

3. **Format Threshold Status:**
   - Icons: ✓ for Yes, ✗ for No
   - Colors: Green for scoped, Red for not scoped

---

## Best Practices

### Data Refresh

1. **Refresh Data:**
   - Click **Home** → **Refresh**
   - Do this after each Excel update

2. **Scheduled Refresh (Pro):**
   - Publish to Power BI Service
   - Configure gateway for Excel file
   - Set refresh schedule

### Performance Optimization

1. **Reduce Visual Count:**
   - Max 10-15 visuals per page
   - Use bookmarks for multiple views

2. **Optimize Measures:**
   - Use variables in DAX
   - Avoid complex calculated columns
   - Use SELECTEDVALUE for parameters

3. **Limit Data:**
   - Filter out zero amounts
   - Consider aggregating small amounts
   - Archive old data

### User Experience

1. **Add Instructions:**
   - Text box with workflow steps
   - Tooltips on visuals
   - Help page with FAQs

2. **Consistent Formatting:**
   - Use theme colors
   - Consistent fonts
   - Aligned visuals

3. **Navigation:**
   - Buttons for page navigation
   - Breadcrumbs showing current location
   - Back to home button

### Documentation

1. **Document Assumptions:**
   - Note default thresholds
   - Explain calculations
   - List data sources

2. **Version Control:**
   - Save report versions
   - Track changes
   - Note update dates

3. **User Guide:**
   - Create PDF guide
   - Include screenshots
   - Step-by-step instructions

---

## Troubleshooting

### Issue: Relationships not working

**Solution:**
- Check data types match (text to text)
- Verify no leading/trailing spaces
- Confirm unique values in "one" side
- Check cross-filter direction

### Issue: Measures show wrong values

**Solution:**
- Verify context (ALL, ALLSELECTED, FILTER)
- Check filter direction on relationships
- Use DAX Studio for debugging
- Add variables to break down calculation

### Issue: Slow performance

**Solution:**
- Reduce data volume (filter in Power Query)
- Optimize measures (use CALCULATE efficiently)
- Minimize custom columns
- Check relationship cardinality

### Issue: Slicers not filtering

**Solution:**
- Check relationship exists
- Verify cross-filter direction
- Confirm field is in correct table
- Check slicer configuration

---

## Conclusion

This Power BI integration enables powerful, interactive scoping analysis that adapts to your audit requirements. The combination of threshold-based and manual selection provides flexibility while the visualizations ensure transparency and documentation of the scoping process.

**Next Steps:**
1. Import data and create relationships
2. Add DAX measures
3. Build initial dashboard
4. Test with actual data
5. Refine based on user feedback
6. Deploy to Power BI Service (if applicable)

For additional help, refer to the main DOCUMENTATION.md file or Power BI official documentation.

---

**Document Version:** 1.0.0  
**Last Updated:** 2024  
**Compatible with:** Power BI Desktop (latest version)
