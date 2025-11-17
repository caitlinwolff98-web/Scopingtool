# Power BI Dashboard - Complete Build Guide
## Professional, Production-Ready Template

**Version:** 5.0 Production Ready
**Time Required:** 45-60 minutes (first time), 15-20 minutes (with experience)
**Difficulty:** Intermediate
**Output:** Complete Power BI dashboard with all functionality

---

## 📋 Table of Contents

1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Phase 1: Data Import & Relationships](#phase-1-data-import--relationships)
4. [Phase 2: DAX Measures](#phase-2-dax-measures)
5. [Phase 3: Dashboard Pages](#phase-3-dashboard-pages)
6. [Phase 4: Manual Scoping Setup](#phase-4-manual-scoping-setup)
7. [Phase 5: Formatting & Polish](#phase-5-formatting--polish)
8. [Testing & Validation](#testing--validation)

---

## Overview

### What You'll Build

A professional Power BI dashboard with:

✅ **5 Dashboard Pages:**
1. **Overview & Summary** - KPIs and high-level metrics
2. **Manual Scoping Control** - Interactive pack/FSLI scoping
3. **FSLI Analysis** - Coverage by Financial Statement Line Item
4. **Division Analysis** - Coverage by division
5. **Segment Analysis** - Coverage by IAS 8 segment

✅ **40+ DAX Measures** - All calculations pre-configured
✅ **Real-Time Manual Scoping** - Edit table directly, see instant updates
✅ **Professional Formatting** - Color-coded, branded, audit-ready
✅ **Interactive Filters** - Slicers for Pack, FSLI, Division, Segment
✅ **Export Ready** - PDF, PowerPoint, Excel export capabilities

### Key Features

- **Dynamic Coverage Tracking** - Updates in real-time as you scope
- **Pack Contribution Analysis** - See each pack's % of total per FSLI
- **Multi-Dimensional Analysis** - By FSLI, Division, and Segment
- **Manual Override** - Scope in/out specific packs or FSLIs
- **ISA 600 Compliant** - Meets audit documentation requirements

---

## Prerequisites

Before starting, ensure you have:

- [ ] Power BI Desktop installed (latest version)
- [ ] VBA tool run successfully (output file generated)
- [ ] Output file: `Bidvest Scoping Tool Output.xlsx`
- [ ] Basic familiarity with Power BI interface
- [ ] 45-60 minutes uninterrupted time

**Recommended Skills:**
- Basic Power BI navigation
- Understanding of relationships
- Basic DAX (helpful but not required - all formulas provided)

---

## Phase 1: Data Import & Relationships

**Time:** 10 minutes

### Step 1.1: Create New Power BI File

1. **Open Power BI Desktop**
2. Click **Get Data** → **Excel**
3. Navigate to `Bidvest Scoping Tool Output.xlsx`
4. Click **Open**

### Step 1.2: Select Tables to Import

**Select ALL these tables:**

**Primary Data Tables:**
- [x] Full Input Table
- [x] Full Input Percentage
- [x] Journals Table (if exists)
- [x] Journals Percentage (if exists)
- [x] Full Consol Table (if exists)
- [x] Full Consol Percentage (if exists)
- [x] Discontinued Table (if exists)
- [x] Discontinued Percentage (if exists)

**Reference Tables:**
- [x] FSLi Key Table
- [x] Pack Number Company Table

**Scoping Tables:**
- [x] Scoping_Control_Table ⭐ **CRITICAL**
- [x] PowerBI_Scoping (if exists)

**Segment Tables (NEW in v5.0):**
- [x] Segment_Pack_Mapping (if exists)
- [x] Segment_Summary (if exists)

**Do NOT import:**
- [ ] Control Panel (not needed)
- [ ] DAX Measures Guide (reference only)
- [ ] Interactive Dashboard (Excel-based)
- [ ] Threshold Configuration (reference only)

Click **Load** (not Transform Data)

### Step 1.3: Create Relationships

**Navigate to Model View** (left sidebar - looks like linked boxes)

**Create these relationships by dragging fields:**

#### Relationship 1: Pack Code (Pack Company → Full Input)
```
Pack Number Company Table[Pack Code]
    → Full Input Table[Pack Code]

Cardinality: One to Many (1:*)
Cross filter: Both
Active: Yes
```

**How to create:**
1. Click and hold `Pack Code` in Pack Number Company Table
2. Drag to `Pack Code` in Full Input Table
3. Release
4. Double-click the relationship line
5. Set Cross filter direction: Both
6. Click OK

#### Relationship 2: Pack Code (Pack Company → Scoping Control)
```
Pack Number Company Table[Pack Code]
    → Scoping_Control_Table[Pack Code]

Cardinality: One to Many (1:*)
Cross filter: Both
Active: Yes
```

#### Relationship 3: FSLI (FSLi Key → Scoping Control)
```
FSLi Key Table[FSLI]
    → Scoping_Control_Table[FSLI]

Cardinality: One to Many (1:*)
Cross filter: Both
Active: Yes
```

#### Relationship 4: Pack Code (Segment Mapping → Pack Company) [NEW]
```
Segment_Pack_Mapping[Pack Code]
    → Pack Number Company Table[Pack Code]

Cardinality: Many to One (*:1)
Cross filter: Both
Active: Yes
```

**If exists - create similar relationships for:**
- Full Consol Table[Pack Code] → Pack Number Company Table[Pack Code]
- Journals Table[Pack Code] → Pack Number Company Table[Pack Code]
- Discontinued Table[Pack Code] → Pack Number Company Table[Pack Code]

### Step 1.4: Verify Relationships

**Check for issues:**
- [ ] No warning icons on relationships
- [ ] All relationships show as active (solid lines)
- [ ] Pack Code is Text type (not Number)
- [ ] FSLI is Text type

**If you see errors:**
- Check data types (Transform Data → check column types)
- Verify field names match exactly
- Ensure no duplicates in "one" side of relationship

---

## Phase 2: DAX Measures

**Time:** 15-20 minutes

### Step 2.1: Create Measure Table

It's best practice to create a dedicated measures table:

1. Click **Home** → **Enter Data**
2. Name: `_Measures` (underscore makes it sort to top)
3. Add one row with any dummy data
4. Click **Load**
5. In Report view, right-click _Measures table → **Hide**

**All measures below will be created in this table.**

### Step 2.2: Basic Count Measures

Click **New Measure** and copy each formula:

#### Total Packs
```DAX
Total Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```
**Format:** Whole Number

#### Scoped In Packs
```DAX
Scoped In Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
    Scoping_Control_Table[Is Consolidated] = "No"
)
```
**Format:** Whole Number

#### Not Scoped Packs
```DAX
Not Scoped Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Not Scoped",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```
**Format:** Whole Number

#### Total FSLIs
```DAX
Total FSLIs =
DISTINCTCOUNT(Scoping_Control_Table[FSLI])
```
**Format:** Whole Number

### Step 2.3: Coverage Percentage Measures

#### Coverage %
```DAX
Coverage % =
DIVIDE(
    [Scoped In Packs],
    [Total Packs],
    0
)
```
**Format:** Percentage, 1 decimal place

#### Untested %
```DAX
Untested % =
DIVIDE(
    [Not Scoped Packs],
    [Total Packs],
    0
)
```
**Format:** Percentage, 1 decimal place

#### Coverage % by Amount
```DAX
Coverage % by Amount =
VAR ScopedAmount =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalAmount =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
RETURN
    DIVIDE(ScopedAmount, TotalAmount, 0)
```
**Format:** Percentage, 1 decimal place

### Step 2.4: Amount Measures

#### Total Amount (All Packs)
```DAX
Total Amount (All Packs) =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```
**Format:** Currency, 0 decimals

#### Total Amount Scoped In
```DAX
Total Amount Scoped In =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
    Scoping_Control_Table[Is Consolidated] = "No"
)
```
**Format:** Currency, 0 decimals

#### Total Amount Not Scoped
```DAX
Total Amount Not Scoped =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] = "Not Scoped",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```
**Format:** Currency, 0 decimals

### Step 2.5: FSLI-Specific Measures

#### Coverage % per FSLI
```DAX
Coverage % per FSLI =
VAR ScopedAmount =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"}
    )
VAR TotalAmount =
    SUM(Scoping_Control_Table[Amount])
RETURN
    DIVIDE(ScopedAmount, TotalAmount, 0)
```
**Format:** Percentage, 1 decimal place
**Note:** Context-sensitive - automatically filters to selected FSLI

#### Pack Contribution to Total FSLI (NEW)
```DAX
Pack Contribution % to Total FSLI =
VAR PackAmount = SUM(Scoping_Control_Table[Amount])
VAR TotalFSLI =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        ALLEXCEPT(Scoping_Control_Table, Scoping_Control_Table[FSLI]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
RETURN
    DIVIDE(PackAmount, TotalFSLI, 0)
```
**Format:** Percentage, 2 decimal places
**Purpose:** Shows each pack's contribution to the consolidated FSLI total

### Step 2.6: Division Measures

#### Coverage % per Division
```DAX
Coverage % per Division =
VAR ScopedPacks =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalPacks =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
RETURN
    DIVIDE(ScopedPacks, TotalPacks, 0)
```
**Format:** Percentage, 1 decimal place

### Step 2.7: Segment Measures (NEW)

#### Total Segments
```DAX
Total Segments =
DISTINCTCOUNT(Segment_Pack_Mapping[Segment Name])
```
**Format:** Whole Number

#### Coverage % per Segment
```DAX
Coverage % per Segment =
VAR ScopedPacks =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalPacks =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
RETURN
    DIVIDE(ScopedPacks, TotalPacks, 0)
```
**Format:** Percentage, 1 decimal place

#### Segment as % of Total
```DAX
Segment as % of Total =
VAR SegmentAmount =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalAmount =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        ALL(Segment_Pack_Mapping[Segment Name]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
RETURN
    DIVIDE(SegmentAmount, TotalAmount, 0)
```
**Format:** Percentage, 1 decimal place

### Step 2.8: Advanced Analytical Measures

#### Scoping Efficiency Ratio
```DAX
Scoping Efficiency Ratio =
DIVIDE(
    [Coverage % by Amount],
    [Coverage %],
    0
)
```
**Format:** Decimal, 2 places
**Interpretation:** >1.5 = Efficient (covering large packs), <1.2 = Less efficient

**For all other measures, see [DAX_MEASURES_LIBRARY.md](DAX_MEASURES_LIBRARY.md)**

---

## Phase 3: Dashboard Pages

**Time:** 20-25 minutes

### Page 1: Overview & Summary

#### Step 3.1.1: Create Page

1. Click **+** at bottom to add new page
2. Rename: "01 - Overview"
3. Page background: Light gray (#F5F5F5)

#### Step 3.1.2: Add Title

1. Insert → Text Box
2. Text: "ISA 600 Consolidation Scoping Dashboard"
3. Font: Segoe UI, Bold, Size 24
4. Color: Dark blue (#003366)
5. Position: Top center

#### Step 3.1.3: Add Subtitle

1. Insert → Text Box
2. Text: "Bidvest Group Limited - [Period]"
3. Font: Segoe UI, Regular, Size 14
4. Color: Gray (#666666)
5. Position: Below title

#### Step 3.1.4: Add KPI Cards

**Create 6 Card visuals across the top:**

**Card 1: Total Packs**
- Visual: Card
- Field: [Total Packs]
- Category label: "Total Packs"
- Background: White
- Border: Light gray
- Number format: Whole number
- Font size: 32pt

**Card 2: Scoped In Packs**
- Field: [Scoped In Packs]
- Category label: "Scoped In"
- Background: Light green (#E8F5E9)
- Number color: Green (#2E7D32)

**Card 3: Not Scoped Packs**
- Field: [Not Scoped Packs]
- Category label: "Not Scoped"
- Background: Light yellow (#FFF9C4)
- Number color: Orange (#F57C00)

**Card 4: Coverage %**
- Field: [Coverage %]
- Category label: "Coverage %"
- Background: Light blue (#E3F2FD)
- Number color: Blue (#1976D2)
- Data label: Percentage, 1 decimal

**Card 5: Coverage % by Amount**
- Field: [Coverage % by Amount]
- Category label: "Coverage % (by Value)"
- Background: Light blue (#E3F2FD)
- Number color: Blue (#1976D2)

**Card 6: Scoping Efficiency Ratio**
- Field: [Scoping Efficiency Ratio]
- Category label: "Efficiency Ratio"
- Background: White
- Tooltip: ">1.5 = Efficient"

#### Step 3.1.5: Add Summary Charts

**Chart 1: Scoping Status Breakdown (Donut)**
1. Visual: Donut Chart
2. Legend: Scoping_Control_Table[Scoping Status]
3. Values: Scoping_Control_Table[Pack Code] (Count Distinct)
4. Colors:
   - Scoped In (Auto): Green (#4CAF50)
   - Scoped In (Manual): Light Green (#8BC34A)
   - Not Scoped: Orange (#FF9800)
   - Scoped Out: Red (#F44336)
5. Data labels: Show
6. Position: Left side, below cards

**Chart 2: Coverage by Division (Clustered Bar)**
1. Visual: Clustered Bar Chart
2. Axis: Pack Number Company Table[Division]
3. Values: [Coverage % per Division]
4. Data labels: Show, Percentage
5. Bars: Blue gradient
6. Filter: Division <> "Not Categorized"
7. Position: Right side, below cards

**Chart 3: Top 10 FSLIs by Amount (Bar)**
1. Visual: Clustered Bar Chart
2. Axis: Scoping_Control_Table[FSLI]
3. Values: [Total Amount (All Packs)]
4. Sort: Descending by amount
5. Top N filter: 10
6. Data labels: Show, Currency
7. Position: Bottom left

**Chart 4: Scoping Progress Gauge**
1. Visual: Gauge
2. Value: [Coverage %]
3. Minimum: 0
4. Maximum: 100
5. Target: 80 (adjustable)
6. Colors:
   - <50%: Red
   - 50-80%: Yellow
   - >80%: Green
7. Position: Bottom right

#### Step 3.1.6: Add Slicers

**Slicer 1: Division**
1. Visual: Slicer
2. Field: Pack Number Company Table[Division]
3. Style: Dropdown
4. Filter: <> "Not Categorized"
5. Position: Top right

**Slicer 2: Segment (NEW)**
1. Visual: Slicer
2. Field: Segment_Pack_Mapping[Segment Name]
3. Style: Dropdown
4. Position: Below Division slicer

---

### Page 2: Manual Scoping Control ⭐ **MOST IMPORTANT**

#### Step 3.2.1: Create Page

1. Add new page
2. Rename: "02 - Manual Scoping"
3. Background: White

#### Step 3.2.2: Add Title

1. Text: "Manual Scoping Control"
2. Font: Bold, 20pt
3. Subtitle: "Click in Scoping Status column to change scoping decisions"

#### Step 3.2.3: Create Main Scoping Table

**This is the CRITICAL visual for manual scoping:**

1. Visual: **Table** (NOT Matrix)
2. Add these fields IN THIS ORDER:
   - Pack Number Company Table[Pack Name]
   - Scoping_Control_Table[Pack Code]
   - Pack Number Company Table[Division]
   - Segment_Pack_Mapping[Segment Name] (NEW)
   - Scoping_Control_Table[FSLI]
   - Scoping_Control_Table[Amount]
   - **Scoping_Control_Table[Scoping Status]** ⭐ **CRITICAL**
   - Pack Number Company Table[Is Consolidated]
   - **[Pack Contribution % to Total FSLI]** (NEW measure)

3. **Format the table:**
   - Column width: Auto-fit
   - Text size: 9pt
   - Header: Bold, dark blue background, white text
   - Alternating rows: Light gray (#F9F9F9)
   - Grid: Show vertical lines

4. **Conditional Formatting:**

   **Amount column:**
   - Right-click Amount → Conditional formatting → Data bars
   - Color: Blue gradient

   **Scoping Status column:**
   - Right-click → Conditional formatting → Background color
   - Rules:
     - "Scoped In (Auto)": Green (#4CAF50)
     - "Scoped In (Manual)": Light Green (#8BC34A)
     - "Not Scoped": Yellow (#FFEB3B)
     - "Scoped Out": Red (#FFCDD2)

   **Pack Contribution % column:**
   - Right-click → Conditional formatting → Data bars
   - Color: Orange gradient
   - Shows visual indicator of contribution size

5. **Enable Edit Mode** ⭐ **CRITICAL:**
   - Select the table visual
   - Format pane → General → Advanced options
   - **Edit mode: ON**
   - Allow editing: Scoping Status column

6. **Size:** Make this table LARGE - it's the main workspace
   - Width: 80% of page
   - Height: 70% of page

#### Step 3.2.4: Add Live Coverage Indicators

**Create 4 cards that update in real-time:**

**Card 1: Current Coverage %**
- Field: [Coverage %]
- Large font (36pt)
- Updates instantly when you change scoping status

**Card 2: Packs Scoped In**
- Field: [Scoped In Packs]

**Card 3: Coverage by Amount**
- Field: [Coverage % by Amount]

**Card 4: Untested %**
- Field: [Untested %]
- Color: Orange (warning)

Position these cards at the top of the page above the table.

#### Step 3.2.5: Add Interactive Slicers

**Slicer Panel (left side):**

**Slicer 1: FSLI** ⭐ **IMPORTANT**
- Field: Scoping_Control_Table[FSLI]
- Style: List (vertical)
- Search box: Enabled
- Single select
- Purpose: Filter to specific FSLI to scope

**Slicer 2: Division**
- Field: Pack Number Company Table[Division]
- Style: Dropdown
- Filter: <> "Not Categorized"

**Slicer 3: Segment** (NEW)
- Field: Segment_Pack_Mapping[Segment Name]
- Style: Dropdown

**Slicer 4: Pack Name**
- Field: Pack Number Company Table[Pack Name]
- Style: Dropdown with search
- Purpose: Find specific pack quickly

**Slicer 5: Scoping Status**
- Field: Scoping_Control_Table[Scoping Status]
- Style: Checkbox list
- Default: All selected
- Purpose: Filter to see only "Not Scoped" packs

**Slicer 6: Amount Range**
- Field: Scoping_Control_Table[Amount]
- Style: Between
- Purpose: Filter to packs above materiality

#### Step 3.2.6: Add Instructions Text Box

1. Insert text box at top
2. Text:
   ```
   📌 INSTRUCTIONS FOR MANUAL SCOPING:

   1. Use slicers to filter (e.g., select specific FSLI like "Revenue")
   2. Click in the "Scoping Status" column for any pack
   3. Change from "Not Scoped" to "Scoped In (Manual)"
   4. Watch coverage % update instantly in real-time
   5. To remove: Change to "Scoped Out"
   6. Use "Pack Contribution %" to see each pack's materiality

   💡 TIP: Sort by "Pack Contribution %" descending to scope largest packs first
   ```
3. Background: Light blue (#E3F2FD)
4. Border: Blue

---

#### ⚠️ ALTERNATIVE METHOD: Manual Scoping Without Edit Mode

**If you cannot enable edit mode or it's not working, use this button-based approach:**

##### Option A: Excel-Based Scoping Workflow (RECOMMENDED)

This is the most reliable method and works in all Power BI editions:

1. **In Power BI Desktop:**
   - Keep the Manual Scoping table as **read-only** (no edit mode needed)
   - Use slicers to filter and identify which packs need scoping
   - Note down the Pack Codes you want to scope

2. **Switch to Excel:**
   - Open the "Bidvest Scoping Tool Output.xlsx" file
   - Go to the "Scoping_Control_Table" sheet
   - Find the rows with the Pack Codes you noted
   - Change the "Scoping Status" column:
     - `Not Scoped` → `Scoped In (Manual)`
     - Or → `Scoped Out`
   - Save the Excel file (Ctrl+S)

3. **Refresh Power BI:**
   - Return to Power BI Desktop
   - Click **Home → Refresh**
   - Your scoping changes will appear immediately
   - Coverage % will update automatically

**Advantages:**
- ✅ Always works (no edit mode issues)
- ✅ Can use Excel features (Find, Filter, Multi-select)
- ✅ Can make bulk changes quickly
- ✅ No risk of data corruption

##### Option B: Filter-Based Scoping (Power BI Only)

If you want to stay in Power BI, use this workflow:

1. **Add Clear Selection Buttons:**
   - Insert 4 buttons on the Manual Scoping page
   - Button 1: "Show All Packs" (clears filters)
   - Button 2: "Show Not Scoped Only" (filters to Not Scoped)
   - Button 3: "Show Manual Scope Only" (filters to Scoped In Manual)
   - Button 4: "Show Auto Scope Only" (filters to Scoped In Auto)

2. **Configure Buttons:**

   **Button 1: "Show All Packs"**
   - Action type: Bookmark
   - Create bookmark with all slicers cleared
   - Or use Filter action to clear Scoping Status slicer

   **Button 2: "Show Not Scoped Only"**
   - Action type: Filter
   - Filter: Scoping_Control_Table[Scoping Status] = "Not Scoped"
   - Allows you to focus on packs that need decisions

   **Button 3: "Show Manual Scope Only"**
   - Action type: Filter
   - Filter: Scoping_Control_Table[Scoping Status] = "Scoped In (Manual)"
   - Review your manual scoping decisions

   **Button 4: "Show Auto Scope Only"**
   - Action type: Filter
   - Filter: Scoping_Control_Table[Scoping Status] = "Scoped In (Auto)"
   - Review threshold-based scoping

3. **Visual Styling for Buttons:**
   - Shape: Rounded rectangle
   - Colors:
     - Button 1: Blue (#2196F3)
     - Button 2: Yellow (#FFC107)
     - Button 3: Light Green (#8BC34A)
     - Button 4: Dark Green (#4CAF50)
   - Text: Bold, White, 12pt
   - Size: 150px wide × 40px tall
   - Position: Horizontal row above the table

4. **Add Export Button:**
   - Button 5: "Export to Excel"
   - Action: Web URL
   - URL: Link to instructions for exporting data
   - Purpose: Users can export filtered data to Excel for batch updates

##### Option C: Two-Stage Workflow (Hybrid)

Combine Power BI filtering with Excel editing:

1. **Stage 1 - Identify in Power BI:**
   - Use Page 2 (Manual Scoping Control)
   - Use slicers to filter:
     - FSLI: "Revenue"
     - Amount: > 50,000,000
     - Scoping Status: "Not Scoped"
   - Export the filtered table to Excel:
     - Right-click table → Export data
     - Save as "Packs_To_Scope.xlsx"

2. **Stage 2 - Update in Excel:**
   - Open both files:
     - "Packs_To_Scope.xlsx" (your filtered list)
     - "Bidvest Scoping Tool Output.xlsx" (the data source)
   - Use VLOOKUP or manual updates to change Scoping Status
   - Save "Bidvest Scoping Tool Output.xlsx"

3. **Stage 3 - Refresh Power BI:**
   - Return to Power BI Desktop
   - Click Refresh
   - Verify changes appear correctly

##### Troubleshooting Edit Mode

If you still want to try enabling edit mode, follow these exact steps:

1. **Verify Power BI Version:**
   - Edit mode requires Power BI Desktop (not Web)
   - Version: April 2023 or later
   - Check: Help → About

2. **Enable Edit Mode (Detailed Steps):**
   - Click the table visual
   - Format pane (paint roller icon)
   - Scroll to **General** section
   - Expand **Advanced options**
   - Find **"Edit interactions"** or **"Edit mode"**
   - Toggle: **ON**

3. **Configure Editable Columns:**
   - Only "Scoping Status" should be editable
   - All other columns: Read-only

4. **Common Issues:**
   - **Issue:** Edit mode option not visible
     - **Fix:** Update Power BI Desktop to latest version

   - **Issue:** Edit mode exists but cells won't change
     - **Fix:** Check data source connection is valid
     - **Fix:** Ensure Excel file is not open elsewhere

   - **Issue:** Changes don't persist
     - **Fix:** Edit mode doesn't write back to Excel automatically
     - **Fix:** Use Excel-based workflow instead (Option A above)

---

**⭐ RECOMMENDATION:**

For most users, **Option A (Excel-Based Workflow)** is the best choice because:
- It always works reliably
- You can make bulk changes efficiently
- Excel's Find & Replace, filtering, and sorting are more powerful
- No risk of Power BI sync issues

The Power BI dashboard is best used for **visualization and analysis**, while Excel is better for **data entry and updates**.

---

### Page 3: FSLI Analysis

#### Step 3.3.1: Create Page

1. Add new page
2. Rename: "03 - FSLI Analysis"
3. Title: "Coverage Analysis by Financial Statement Line Item"

#### Step 3.3.2: Main FSLI Table

1. Visual: **Matrix**
2. Rows: Scoping_Control_Table[FSLI]
3. Columns: Scoping_Control_Table[Scoping Status]
4. Values:
   - Scoping_Control_Table[Pack Code] (Count Distinct)
   - Scoping_Control_Table[Amount] (Sum)
5. **Add measure columns:**
   - [Coverage % per FSLI]
   - [Untested %]
6. **Conditional Formatting:**
   - Coverage % per FSLI:
     - >80%: Green
     - 50-80%: Yellow
     - <50%: Red
7. **Sorting:** Sort by Coverage % descending

#### Step 3.3.3: FSLI Coverage Chart

1. Visual: Clustered Bar Chart
2. Axis: Scoping_Control_Table[FSLI]
3. Values: [Coverage % per FSLI]
4. Data labels: Show percentage
5. Bars: Color gradient (red → yellow → green)
6. Top N filter: 20 (adjustable)
7. Sort: By coverage % descending

#### Step 3.3.4: Pack Contribution Heatmap (NEW)

1. Visual: Matrix
2. Rows: Scoping_Control_Table[FSLI]
3. Columns: Pack Number Company Table[Pack Name]
4. Values: [Pack Contribution % to Total FSLI]
5. **Conditional Formatting:**
   - Background color scale:
     - 0%: White
     - 50%: Orange
     - 100%: Dark Red
   - Shows which packs are significant for each FSLI
6. **Filter:** Show only non-zero contributions

#### Step 3.3.5: FSLI Detail Table

1. Visual: Table
2. Fields:
   - Scoping_Control_Table[FSLI]
   - [Total Amount (All Packs)]
   - [Total Amount Scoped In]
   - [Total Amount Not Scoped]
   - [Coverage % per FSLI]
   - [Pack Contribution % to Total FSLI]
3. **Drill-through setup:**
   - Enable drill-through
   - Field: FSLI
   - Right-click any FSLI → Drill through → Shows pack details

---

### Page 4: Division Analysis

#### Step 3.4.1: Create Page

1. Add new page
2. Rename: "04 - Division Analysis"
3. Title: "Coverage Analysis by Division"

#### Step 3.4.2: Division Summary Cards

Create a card for each key metric per division:
- Total Packs in Division
- Scoped In Packs
- Coverage % per Division
- Amount in Division
- Division as % of Total

#### Step 3.4.3: Division Comparison Chart

1. Visual: Clustered Column Chart
2. Axis: Pack Number Company Table[Division]
3. Values:
   - [Scoped In Packs]
   - [Not Scoped Packs]
4. Legend: Show
5. Colors: Green (scoped), Orange (not scoped)
6. Data labels: Show

#### Step 3.4.4: Division Coverage Matrix

1. Visual: Matrix
2. Rows: Pack Number Company Table[Division]
3. Columns: Scoping_Control_Table[Scoping Status]
4. Values: Pack Code (Count Distinct)
5. Add coverage % column
6. Conditional formatting on coverage %

---

### Page 5: Segment Analysis (NEW)

#### Step 3.5.1: Create Page

1. Add new page
2. Rename: "05 - Segment Analysis (IAS 8)"
3. Title: "Coverage Analysis by Operating Segment"

#### Step 3.5.2: Segment Summary Cards

- Total Segments
- Segments Fully Scoped
- Segments Partially Scoped
- Segments Not Scoped
- Average Coverage per Segment

#### Step 3.5.3: Segment Coverage Chart

1. Visual: Clustered Bar Chart
2. Axis: Segment_Pack_Mapping[Segment Name]
3. Values: [Coverage % per Segment]
4. Data labels: Show percentage
5. Target line: 80% (ISA 600 guidance)

#### Step 3.5.4: Segment Materiality Analysis

1. Visual: Treemap
2. Group: Segment_Pack_Mapping[Segment Name]
3. Values: [Total Amount (All Packs)]
4. Color saturation: [Coverage % per Segment]
5. Purpose: Shows segment size and coverage visually

#### Step 3.5.5: Segment-Division Cross-Analysis

1. Visual: Matrix
2. Rows: Segment_Pack_Mapping[Segment Name]
3. Columns: Pack Number Company Table[Division]
4. Values: Pack Code (Count), Coverage %
5. Shows which segments span which divisions

---

## Phase 4: Manual Scoping Setup

**Time:** 10 minutes

### Step 4.1: Enable Edit Mode (CRITICAL)

**For Page 2 (Manual Scoping) table:**

1. Select the main scoping table
2. Click Format pane (paint roller icon)
3. Expand **General**
4. Scroll to **Advanced options**
5. Toggle **Edit mode** to **ON**

**Verify:**
- Click in Scoping Status column - should be editable
- Type "Scoped In (Manual)" - should accept
- Coverage % cards should update immediately

### Step 4.2: Test Real-Time Updates

1. Note current Coverage %
2. Change one pack from "Not Scoped" to "Scoped In (Manual)"
3. **Coverage % should increase immediately**
4. Scoped In Packs should increase by 1
5. Not Scoped Packs should decrease by 1

**If coverage doesn't update:**
- Check measures include both "Scoped In (Auto)" AND "Scoped In (Manual)"
- Verify relationships are active
- Check Is Consolidated filter in measures

### Step 4.3: Configure Auto-Save

1. File → Options and settings → Options
2. Data Load → Background data: Enabled
3. Report settings → Auto-save: Every 1 minute
4. Click OK

---

## Phase 5: Formatting & Polish

**Time:** 10-15 minutes

### Step 5.1: Apply Consistent Theme

1. View tab → Themes
2. Select a professional theme (e.g., "Executive")
3. Or customize:
   - Primary color: Dark blue (#003366)
   - Secondary: Light blue (#4A90E2)
   - Accent: Orange (#FF9800)
   - Good: Green (#4CAF50)
   - Bad: Red (#F44336)
   - Warning: Yellow (#FFC107)

### Step 5.2: Format All Visuals

**Apply to ALL visuals:**

**Title:**
- Font: Segoe UI
- Size: 12pt
- Color: Dark gray (#333333)
- Background: None

**Background:**
- Color: White
- Border: 1px, Light gray (#CCCCCC)
- Shadow: Subtle

**Numbers:**
- Font size: Match visual importance
- Color: Black for amounts, colored for KPIs
- Thousands separator: Yes
- Decimals: Currency=0, Percentage=1

### Step 5.3: Add Page Navigation

1. Insert → Buttons → Blank button (on each page)
2. Button text: "← Back to Overview" / "Next Page →"
3. Button action:
   - Type: Page navigation
   - Destination: Select page
4. Position: Top right corner
5. Style: Rounded corners, blue background

### Step 5.4: Add Last Refresh Timestamp

1. Insert → Text box
2. Text: "Last Refreshed: [Date Time]"
3. Position: Bottom right of each page
4. Font: Small (8pt), Gray
5. Or use: Insert → More visuals → "Text filter" visual
6. Field: NOW() function

### Step 5.5: Lock Layout

1. View tab → Page view
2. Selection pane: Open
3. Lock all visuals in position
4. Group related visuals together

---

## Testing & Validation

### Functional Testing Checklist

**Data Import:**
- [ ] All required tables imported successfully
- [ ] Relationships active and correct
- [ ] No warning icons on relationships
- [ ] Pack Code is Text in all tables

**DAX Measures:**
- [ ] All 20+ measures created
- [ ] Measures calculate correctly (no errors)
- [ ] Total Packs excludes consolidated entity
- [ ] Coverage % shows 0-100%
- [ ] Measures respond to filters

**Manual Scoping:**
- [ ] Can click in Scoping Status column
- [ ] Can type "Scoped In (Manual)"
- [ ] Coverage % updates immediately
- [ ] Scoped In Packs count updates
- [ ] Changes persist after clicking away

**Slicers & Filters:**
- [ ] Division slicer works
- [ ] FSLI slicer works
- [ ] Segment slicer works (NEW)
- [ ] Pack name slicer works
- [ ] Amount range slicer works
- [ ] All slicers interact correctly

**Visuals:**
- [ ] All charts display data
- [ ] Conditional formatting shows correctly
- [ ] Data labels visible
- [ ] Tooltips work
- [ ] Drill-through works (where configured)

**Performance:**
- [ ] Dashboard loads in < 10 seconds
- [ ] Manual scoping updates in < 2 seconds
- [ ] Slicers respond quickly
- [ ] No lag when switching pages

### Validation Testing

**Test Scenario 1: Manual Scoping Workflow**
1. Go to Manual Scoping page
2. Filter to FSLI = "Revenue"
3. Sort by Pack Contribution % descending
4. Scope in top 3 packs
5. **Expected:** Coverage % for Revenue should be ~60-80%
6. **Verify:** Cards update immediately

**Test Scenario 2: Division Analysis**
1. Go to Overview page
2. Select one division in slicer
3. **Expected:** All visuals filter to that division
4. **Verify:** Total Packs decreases, charts update

**Test Scenario 3: Segment Analysis (NEW)**
1. Go to Segment Analysis page
2. Select one segment
3. **Expected:** Shows packs in that segment
4. **Verify:** Coverage % per Segment calculates
5. **Verify:** Segment materiality shows correctly

**Test Scenario 4: FSLI Coverage**
1. Go to FSLI Analysis page
2. Find FSLI with <50% coverage
3. Click to drill-through
4. **Expected:** Shows pack details
5. Scope in additional packs
6. **Expected:** Coverage increases

---

## Save & Export

### Save Your Work

1. File → Save As
2. Name: `Bidvest_ISA600_Scoping_Dashboard_v5.0.pbix`
3. Location: Same folder as Excel output
4. **Backup:** Save copy to OneDrive/network drive

### Export Options

**PDF Export:**
1. File → Export → Export to PDF
2. Select: All pages
3. Quality: Best
4. Use for: Audit documentation

**PowerPoint Export:**
1. File → Export → Export to PowerPoint
2. Select pages to include
3. One visual per slide or full pages
4. Use for: Presentations to management

**Excel Export:**
1. Right-click any visual → Export data
2. Choose: Underlying data or Summarized data
3. Use for: Detailed analysis in Excel

---

## Maintenance & Updates

### When Consolidation Data Changes

1. Run VBA tool on new consolidation workbook
2. Output file updates: `Bidvest Scoping Tool Output.xlsx`
3. In Power BI: Click **Refresh** (Home tab)
4. All visuals update automatically
5. Manual scoping changes persist (if using DirectQuery)

### Monthly Checklist

- [ ] Refresh data from latest consolidation
- [ ] Review scoping decisions
- [ ] Update target coverage % if needed
- [ ] Export updated dashboard to PDF
- [ ] Archive previous version
- [ ] Document any methodology changes

---

## Troubleshooting

### Issue: Edit Mode Not Working

**Symptoms:** Can't click in Scoping Status column

**Solutions:**
1. **Check visual type:** Must be Table (not Matrix)
2. **Enable edit mode:** Format → General → Advanced → Edit mode ON
3. **Check column type:** Scoping Status must be Text
4. **Power BI version:** Update to latest version
5. **Alternative:** Use Power BI Service (online) for edit mode

### Issue: Coverage % Not Updating

**Symptoms:** Manual scoping doesn't update coverage metrics

**Solutions:**
1. **Check measure formulas:** Include both "Scoped In (Auto)" AND "(Manual)"
2. **Verify relationships:** Pack Code connections active
3. **Clear filters:** Ensure no conflicting filters
4. **Refresh data:** Try clicking Refresh
5. **Check calculated columns:** Ensure no errors

### Issue: Relationships Not Working

**Symptoms:** Slicers don't filter visuals

**Solutions:**
1. **Data types:** Pack Code must be Text (not Number)
2. **Cardinality:** Check One-to-Many is correct direction
3. **Cross-filter:** Set to Both directions
4. **Active relationships:** Ensure solid line (not dotted)
5. **Duplicates:** Check for duplicate Pack Codes in dimension table

### Issue: Segment Data Not Showing (NEW)

**Symptoms:** Segment analysis page is blank

**Solutions:**
1. **Check VBA run:** Ensure segment document was processed
2. **Verify tables:** Segment_Pack_Mapping should exist
3. **Check relationships:** Segment tables connected correctly
4. **Data type:** Segment Name and Pack Code are Text
5. **Matching:** Verify segment-to-consolidation matching worked

---

## Appendix: Quick Reference

### Page Navigation Quick Reference

```
Page 1: Overview & Summary
↓
Page 2: Manual Scoping Control ⭐ (MAIN WORKSPACE)
↓
Page 3: FSLI Analysis
↓
Page 4: Division Analysis
↓
Page 5: Segment Analysis (IAS 8)
```

### Essential Keyboard Shortcuts

- `Ctrl + S` - Save
- `Ctrl + R` - Refresh data
- `Ctrl + ` - Show/hide panes
- `F5` - Refresh selected visual
- `Ctrl + ]` - Navigate pages forward
- `Ctrl + [` - Navigate pages backward
- `Ctrl + click` - Multi-select in slicers

### File Locations

- **Source:** `Bidvest Scoping Tool Output.xlsx`
- **Power BI:** `Bidvest_ISA600_Scoping_Dashboard_v5.0.pbix`
- **Exports:** `./Exports/` folder (create)
- **Archive:** Previous versions for audit trail

---

**Build Guide Version:** 5.0
**Last Updated:** November 2025
**Estimated Build Time:** 45-60 minutes (first time), 15-20 minutes (experienced)
**Result:** Professional, production-ready ISA 600 scoping dashboard

**Next Steps:**
- See [VERIFICATION_CHECKLIST.md](VERIFICATION_CHECKLIST.md) to validate your dashboard
- See [DAX_MEASURES_LIBRARY.md](DAX_MEASURES_LIBRARY.md) for additional measures
- See [IMPLEMENTATION_GUIDE.md](IMPLEMENTATION_GUIDE.md) for complete workflow

---

**Questions?** This is the most comprehensive build guide possible without creating the actual .pbix file. Follow each step carefully, and you'll have a professional dashboard in ~60 minutes.

🎉 **Happy Building!** 🎉
