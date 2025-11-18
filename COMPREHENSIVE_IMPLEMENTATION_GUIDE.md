# ISA 600 Revised Component Scoping Tool - Comprehensive Implementation Guide

**Version 6.0 - Complete Overhaul**
**Bidvest Group Limited**
**Last Updated: November 2025**

---

## Table of Contents

1. [Overview](#1-overview)
2. [System Requirements](#2-system-requirements)
3. [Installation & Setup](#3-installation--setup)
4. [Step-by-Step Usage Guide](#4-step-by-step-usage-guide)
5. [Dashboard Navigation](#5-dashboard-navigation)
6. [Manual Scoping Instructions](#6-manual-scoping-instructions)
7. [Power BI Setup & Integration](#7-power-bi-setup--integration)
8. [Technical Reference](#8-technical-reference)
9. [Troubleshooting](#9-troubleshooting)
10. [ISA 600 Compliance](#10-isa-600-compliance)

---

## 1. Overview

### 1.1 Purpose

The **ISA 600 Revised Component Scoping Tool** is a comprehensive VBA-based solution designed specifically for Bidvest Group Limited to automate the identification and scoping of significant components for group audit engagements.

### 1.2 What This Tool Does

This tool transforms complex consolidation data into structured, actionable insights by:

✅ **Automatically processing** Stripe Packs consolidation workbooks
✅ **Integrating** Segmental Reporting data with fuzzy matching
✅ **Applying** configurable FSLI thresholds for automatic scoping
✅ **Generating** interactive Excel dashboards with 5 comprehensive views
✅ **Creating** Power BI-ready dimension and fact tables
✅ **Enabling** manual scoping at pack or FSLI level
✅ **Calculating** coverage percentages by FSLI, Division, and Segment
✅ **Identifying** untested portions requiring additional audit procedures
✅ **Producing** professional reports and documentation

### 1.3 Key Features

| Feature | Description |
|---------|-------------|
| **No Setup Required** | Works "out of the box" - no manual configuration |
| **Guided Workflow** | Step-by-step prompts guide you through every decision |
| **ISA 600 Compliant** | Built specifically for ISA 600 Revised requirements |
| **Consolidation Entity Handling** | Automatically identifies and excludes consolidation entity |
| **Dual Currency Support** | Processes both Entity and Consolidation currencies |
| **Threshold-Based Scoping** | Configurable automatic scoping (e.g., Revenue > R50M) |
| **Manual Scoping Interface** | Fine-tune scoping at pack or FSLI level |
| **Division-Segment Mapping** | Reconciles Stripe and Segmental reporting |
| **Interactive Dashboards** | 5 comprehensive dashboard views in Excel |
| **Power BI Ready** | Pre-structured dimension and fact tables |
| **Professional Output** | Timestamped, formatted, audit-ready workbooks |

### 1.4 Output Deliverables

When complete, the tool generates:

- **Full Input Table** - Complete pack × FSLI amounts
- **Full Input Percentage Table** - % contribution to consolidated totals
- **Additional Tables** - Discontinued, Journals, Consol (if applicable)
- **FSLi Key Table** - FSLI master reference
- **Pack Number Company Table** - Pack master reference with divisions
- **Division-Segment Mapping** - Reconciliation between Stripe and Segmental
- **Dashboard - Overview** - Executive summary with key metrics
- **Coverage by FSLI** - FSLI-level coverage analysis
- **Coverage by Division** - Division-level analysis
- **Coverage by Segment** - Segment-level analysis
- **Detailed Pack Analysis** - Interactive pack × FSLI table
- **Manual Scoping Interface** - Editable scoping interface
- **Power BI Tables** - Dim_Packs, Dim_FSLIs, Fact_Amounts, Fact_Percentages, Fact_Scoping
- **Scoping Summary** - Pack-level recommendations
- **Threshold Configuration** - Documentation of threshold settings (if applied)

---

## 2. System Requirements

### 2.1 Software Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| **Operating System** | Windows 10 | Windows 10/11 |
| **Microsoft Excel** | Excel 2016 | Excel 2019 or Microsoft 365 |
| **RAM** | 4 GB | 8 GB or more |
| **Disk Space** | 500 MB free | 1 GB free |
| **Power BI Desktop** | Any version (for BI features) | Latest version |

### 2.2 Excel Settings

✅ **Macros must be enabled**
- File → Options → Trust Center → Trust Center Settings → Macro Settings
- Select "Enable all macros" (or "Disable all macros with notification")

✅ **References must be available** (usually enabled by default):
- Microsoft Scripting Runtime

### 2.3 Source Data Requirements

The tool expects:

**Stripe Packs Consolidation Workbook:**
- Row 6: Column type identifiers (Entity/Consolidation currency)
- Row 7: Pack/Entity names
- Row 8: Pack/Entity codes
- Row 9+: FSLI data
- Column B: FSLI names

**Segmental Reporting Workbook (Optional):**
- Row 8: Pack information in format "Pack Name - Pack Code"
- Segment tabs categorized during processing

---

## 3. Installation & Setup

### 3.1 Quick Start (5 Minutes)

**Step 1: Create the Macro-Enabled Workbook**

1. Open Microsoft Excel
2. Create a new blank workbook
3. Save as: `Bidvest_Scoping_Tool.xlsm`
   - File → Save As
   - Choose location (recommend: Desktop or Documents)
   - Save as type: **Excel Macro-Enabled Workbook (*.xlsm)**

**Step 2: Open VBA Editor**

1. Press `Alt + F11` to open the VBA Editor
2. You should see the VBAProject for your workbook in the Project Explorer (left pane)

**Step 3: Import VBA Modules**

1. In VBA Editor, select your project (VBAProject (Bidvest_Scoping_Tool.xlsm))
2. File → Import File...
3. Navigate to the `VBA_Modules` folder
4. Import **all 8 modules** one by one:
   - `Mod1_MainController.bas`
   - `Mod2_TabProcessing.bas`
   - `Mod3_DataExtraction.bas`
   - `Mod4_SegmentalMatching.bas`
   - `Mod5_ScopingEngine.bas`
   - `Mod6_DashboardGeneration.bas`
   - `Mod7_PowerBIExport.bas`
   - `Mod8_Utilities.bas`

5. Verify all 8 modules appear in the Project Explorer under "Modules"

**Step 4: Create Macro Button**

1. Close VBA Editor (press `Alt + Q` or close window)
2. In Excel, go to: Developer tab → Insert → Button (Form Control)
   - *If Developer tab not visible:* File → Options → Customize Ribbon → Check "Developer"
3. Draw a button on the worksheet (anywhere you like)
4. In the "Assign Macro" dialog, select: **StartBidvestScopingTool**
5. Click OK
6. Right-click the button → Edit Text
7. Change text to: **Start Bidvest Scoping Tool**

**Step 5: Test Installation**

1. Save your workbook (`Ctrl + S`)
2. Click the "Start Bidvest Scoping Tool" button
3. You should see the welcome message
4. Click Cancel to exit (we'll do a full run in Section 4)

✅ **Installation Complete!**

---

## 4. Step-by-Step Usage Guide

### 4.1 Preparation Checklist

Before running the tool:

- [ ] Stripe Packs consolidation workbook is open in Excel
- [ ] Segmental Reporting workbook is open (optional)
- [ ] Both workbooks contain current, accurate data
- [ ] You know which entity is the consolidation entity (e.g., BBT-001)
- [ ] You have 10-15 minutes available

### 4.2 Running the Tool

**STEP 1: Launch the Tool**

1. Click the **"Start Bidvest Scoping Tool"** button
2. Read the welcome message
3. Click **OK** to begin

**STEP 2: Select Stripe Packs Workbook**

```
Prompt: "Please enter the exact name of the TGK consolidation workbook."
```

1. Switch to your Stripe Packs workbook
2. Look at the title bar - copy the exact workbook name
3. Example: `Bidvest_Consolidation_2024_Q4.xlsx`
4. Paste into the input box
5. Click **OK**

✅ **Expected Result:** Confirmation message showing workbook loaded

**STEP 3: Categorize Tabs**

```
Prompt: Tab categorization dialog for each tab
```

For each tab in your workbook, select the appropriate category:

| Category | Number | When to Use | Quantity |
|----------|--------|-------------|----------|
| **Division** | 1 | Business segment/division tabs (will prompt for division name) | Multiple |
| **Discontinued Operations** | 2 | Discontinued operations data | Single |
| **Input Continuing** | 3 | **PRIMARY DATA SOURCE - REQUIRED** | Single |
| **Journals Continuing** | 4 | Consolidation journal entries | Single |
| **Consol Continuing** | 5 | Consolidated financial data | Single |
| **Trial Balance** | 6 | Trial balance data | Single |
| **Balance Sheet** | 7 | Balance Sheet statement | Single |
| **Income Statement** | 8 | Income Statement | Single |
| **Uncategorized** | 9 | Ignore this tab | Multiple |

**Critical Notes:**
- At least **ONE** tab must be categorized as "Input Continuing" (category 3)
- Category 3 is your primary data source
- For Division tabs (category 1), you'll be prompted to provide a friendly division name

**Example Categorization:**
```
Tab: "TGK_UK_Division" → Category: 1 (Division) → Division Name: "UK Division"
Tab: "TGK_Discontinued" → Category: 2 (Discontinued Operations)
Tab: "TGK_Input_Cont" → Category: 3 (Input Continuing) ⭐ REQUIRED
Tab: "TGK_Journals" → Category: 4 (Journals Continuing)
Tab: "TGK_Consol" → Category: 5 (Consol Continuing)
Tab: "Summary_Sheet" → Category: 9 (Uncategorized)
```

After categorizing all tabs, you'll see a summary. Confirm or recategorize.

**STEP 4: Assign Division Names**

```
Prompt: For each Division tab, enter a friendly name
```

For each tab categorized as "Division" (category 1):
1. Enter a meaningful division name
2. This name will appear in reports and dashboards
3. Examples: "UK Division", "South Africa Division", "Europe Division"

**STEP 5: Select Currency Type**

```
Prompt: "Which columns would you like to use?"
```

**Recommended: Click YES for Consolidation Currency**

- **YES (Recommended):** Process Consolidation/Consolidable Currency columns
  - This is the group reporting currency
  - Required for ISA 600 scoping

- **NO:** Process Original/Entity Currency columns
  - Local currency of each entity
  - Not typically used for scoping

**STEP 6: Identify Consolidation Entity**

```
Prompt: List of all entities - select the consolidation entity
```

1. Review the list of all packs/entities
2. Identify which one represents the **CONSOLIDATED ENTITY**
   - This is typically the Bidvest Group consolidated pack
   - Example: "BBT-001", "BIDVEST-CONSOL", etc.
   - This entity aggregates all other packs
3. Enter the number corresponding to that entity
4. Confirm your selection

**Important:** The consolidation entity will be:
- Used as the 100% baseline for percentage calculations
- **Excluded** from scoping (it's the aggregate, not a component)

**STEP 7: Process Segmental Reporting (Optional)**

```
Prompt: "Would you like to process the Segmental Reporting workbook?"
```

**If YES:**
1. Enter the Segmental Reporting workbook name
2. Categorize segmental tabs:
   - **1 = Segment Tab** (will prompt for segment name)
   - **2 = Summarized Segment**
   - **3 = Uncategorized**
3. The tool will perform pack matching between Stripe and Segmental

**If NO:**
- Skip segmental processing
- You can still use division-based analysis

**STEP 8: Configure Thresholds (Optional)**

```
Prompt: "Would you like to configure threshold-based automatic scoping?"
```

**If YES:**

1. **Select FSLIs for Threshold Analysis**
   - You'll see a numbered list of all FSLIs
   - Enter numbers separated by commas
   - Example: `1,5,12` (to select FSLIs #1, #5, and #12)
   - Recommended: Revenue, PBT (Profit Before Tax), Total Assets

2. **Enter Threshold Amount for Each FSLI**
   - For each selected FSLI, enter the threshold amount
   - Example: For Revenue, enter `50000000` (R50 million)
   - Packs **exceeding** any threshold will be automatically scoped in

3. **Review Threshold Configuration**
   - Confirm your threshold settings
   - The tool will identify and scope in packs meeting criteria

**Threshold Logic:**
- **ANY threshold exceeded → ENTIRE PACK scoped in**
- Example: Revenue > R50M **OR** Total Assets > R100M
- If a pack has Revenue of R60M → All FSLIs for that pack are scoped in

**If NO:**
- Skip automatic scoping
- You can manually scope packs later using the dashboard

**STEP 9-12: Processing**

The tool now runs automatically:

✅ **Step 9:** Extracting data and generating tables (2-5 minutes)
✅ **Step 10:** Creating interactive dashboard
✅ **Step 11:** Creating Power BI integration assets
✅ **Step 12:** Saving output workbook

**You'll see status updates in the Excel status bar (bottom left)**

**STEP 13: Completion**

```
Success message with output workbook details
```

The tool displays:
- Output workbook name and location
- List of generated assets
- Processing time
- Next steps

The output workbook is saved as:
```
Bidvest Group Scoping [YYYY-MM-DD] [HH-MM-SS].xlsm
Example: Bidvest Group Scoping [2024-11-18] [14-30-45].xlsm
```

---

## 5. Dashboard Navigation

The tool generates **5 comprehensive dashboard views** in Excel.

### 5.1 Dashboard - Overview

**Purpose:** Executive summary with key metrics

**Location:** "Dashboard - Overview" sheet

**What You'll See:**
- Total Packs: Total count of all packs
- Packs Scoped In: Count of automatically scoped packs
- Packs Not Yet Scoped: Count of packs requiring review
- Overall Coverage %: Percentage of packs scoped

**Color Coding:**
- 🟢 Green = Scoped In
- 🟡 Yellow = Not Yet Scoped
- 🔵 Blue = Headers/Titles

**How to Use:**
1. Review overall coverage percentage
2. If coverage is low (<60%), consider:
   - Adjusting thresholds
   - Manually scoping additional packs
   - Reviewing highest contributors

### 5.2 Coverage by FSLI

**Purpose:** Analyze coverage for each Financial Statement Line Item

**Location:** "Coverage by FSLI" sheet

**Columns:**
- FSLI: Name of the financial statement line item
- Total Consolidated Amount: 100% baseline from consolidation entity
- Scoped Amount: Sum of scoped pack amounts for this FSLI
- Untested Amount: Difference (Total - Scoped)
- Coverage %: (Scoped / Total) × 100

**Color Coding:**
- 🟢 Green ≥ 80% coverage (good)
- 🟡 Yellow = 60-79% coverage (review)
- 🔴 Red < 60% coverage (insufficient)

**How to Use:**
1. **Identify low-coverage FSLIs** (red or yellow)
2. **Sort by Coverage %** (lowest to highest) to prioritize
3. **Review "Untested Amount"** to see materiality of gap
4. **Navigate to "Manual Scoping Interface"** to scope additional packs for low-coverage FSLIs

**Example:**
```
FSLI: Total Revenue
Total Consolidated Amount: R 10,000,000,000
Scoped Amount: R 7,500,000,000
Untested Amount: R 2,500,000,000
Coverage %: 75.0% (Yellow - review)

Action: Consider scoping additional revenue-generating packs
```

### 5.3 Coverage by Division

**Purpose:** Analyze scoping by division (from Stripe Packs)

**Location:** "Coverage by Division" sheet

**Columns:**
- Division: Division name (from categorization step)
- Total Packs: Count of packs in this division
- Scoped Packs: Count of scoped packs in this division
- Coverage %: (Scoped / Total) × 100

**How to Use:**
1. **Identify divisions with low coverage**
2. **Ensure proportional coverage** across all divisions
3. **Consider ISA 600 component requirements** by division

### 5.4 Coverage by Segment

**Purpose:** Analyze scoping by segment (from Segmental Reporting)

**Location:** "Coverage by Segment" sheet

**Only available if:** Segmental Reporting was processed

**Columns:**
- Segment: Segment name (from segmental workbook)
- Total Packs: Count of packs in this segment
- Scoped Packs: Count of scoped packs in this segment
- Coverage %: (Scoped / Total) × 100

**How to Use:**
- Similar to Coverage by Division
- Ensures segment-level compliance with ISA 600

### 5.5 Detailed Pack Analysis

**Purpose:** Interactive, filterable table showing all pack × FSLI combinations

**Location:** "Detailed Pack Analysis" sheet

**Columns:**
- Pack Code: Entity code
- Pack Name: Entity name
- FSLI: Financial statement line item
- Amount: Amount for this pack-FSLI combination
- % of Consolidated: Percentage contribution
- Scoping Status: "Scoped In", "Not Scoped", or "Scoped Out"
- Division: Division assignment
- Segment: Segment assignment (if available)

**Features:**
- ✅ **Sortable:** Click column headers to sort
- ✅ **Filterable:** Use AutoFilter to filter by any column
- ✅ **Searchable:** Use Ctrl+F to find specific packs or FSLIs

**How to Use:**

**Example 1: Find packs with high Revenue contribution**
1. Filter FSLI column to "Total Revenue"
2. Sort by "% of Consolidated" (highest to lowest)
3. Review top contributors
4. Check "Scoping Status" - if "Not Scoped", consider scoping in

**Example 2: Review all packs in a specific division**
1. Filter Division column to "UK Division"
2. Review all packs and their scoping status
3. Ensure adequate coverage

**Example 3: Find untested high-value items**
1. Filter "Scoping Status" to "Not Scoped"
2. Sort by "Amount" (highest to lowest)
3. Identify material untested amounts
4. Navigate to Manual Scoping Interface to scope them in

---

## 6. Manual Scoping Instructions

### 6.1 Manual Scoping Interface

**Location:** "Manual Scoping Interface" sheet

**Purpose:** Fine-tune scoping decisions at pack or FSLI level

**Columns:**
- Pack Code
- Pack Name
- FSLI
- Amount
- **Scoping Status** (Editable dropdown)
- Division
- Segment

### 6.2 How to Manually Scope

**Method 1: Scope In Specific FSLI for Specific Pack**

1. Navigate to "Manual Scoping Interface" sheet
2. **Filter** to find the pack-FSLI combination you want
   - Example: Pack = "ABC-123", FSLI = "Inventory"
3. **Change "Scoping Status"** to:
   - **"Scoped In"** → Include in audit scope
   - **"Not Scoped"** → Exclude from audit scope
   - **"Scoped Out"** → Explicitly excluded (for corrections)
4. Coverage percentages will update in dashboard views

**Method 2: Scope In Entire Pack (All FSLIs)**

1. Navigate to "Manual Scoping Interface" sheet
2. **Filter "Pack Code"** to the pack you want to scope in
3. **Select all rows** for that pack (all FSLIs)
4. **Change "Scoping Status"** for all rows to "Scoped In"

**Method 3: Scope In All Packs for Specific FSLI**

1. Navigate to "Manual Scoping Interface" sheet
2. **Filter "FSLI"** to the FSLI you want (e.g., "Inventory")
3. **Sort by "Amount"** (highest to lowest)
4. **Change "Scoping Status"** for top contributors to "Scoped In"
5. Continue until desired coverage % is reached

### 6.3 Coverage Targets

ISA 600 does not specify exact coverage percentages, but best practices suggest:

| Coverage Level | Recommendation |
|----------------|----------------|
| **≥ 90%** | Excellent coverage |
| **80-89%** | Good coverage - consider increasing |
| **60-79%** | Moderate coverage - review untested portions |
| **< 60%** | Insufficient coverage - scope additional packs |

**Factors to Consider:**
- Materiality of untested amounts
- Risk profile of untested components
- Group audit strategy
- Resource constraints

### 6.4 Efficiency Metrics

The dashboard shows:
- **"Adding Pack X increases Revenue coverage from 72% to 81%"**
- Use this to optimize scoping decisions

**Example Workflow:**
1. Note current coverage: Revenue = 72%
2. Review "Detailed Pack Analysis" for highest Revenue contributors not scoped
3. Scope in Pack X (high contributor)
4. Observe coverage increase to 81%
5. Repeat until target coverage reached

---

## 7. Power BI Setup & Integration

### 7.1 Why Power BI?

Power BI provides:
- ✅ **Enhanced Interactivity:** Cross-filtering, drill-through
- ✅ **Advanced Visualizations:** Heatmaps, treemaps, etc.
- ✅ **Mobile Access:** View dashboards on mobile devices
- ✅ **Auto-Refresh:** Update data without re-running tool
- ✅ **Collaboration:** Share dashboards with team

### 7.2 Power BI Tables Generated

The tool creates **6 Power BI-ready tables**:

**Dimension Tables:**
- **Dim_Packs:** Pack master (PackCode, PackName, Division, Segment, IsConsolidated)
- **Dim_FSLIs:** FSLI master (FSLI, Category, AccountNature)
- **Dim_Thresholds:** Threshold configuration (FSLI, ThresholdAmount)

**Fact Tables:**
- **Fact_Amounts:** Unpivoted amounts (PackCode, FSLI, Amount)
- **Fact_Percentages:** Unpivoted percentages (PackCode, FSLI, Percentage)
- **Fact_Scoping:** Scoping decisions (PackCode, FSLI, ScopingStatus, ScopingMethod)

### 7.3 Import to Power BI (15 Minutes)

**Step 1: Open Power BI Desktop**

1. Launch Power BI Desktop (download from Microsoft if needed)
2. Close splash screen

**Step 2: Get Data**

1. Click **Get Data** (Home ribbon)
2. Select **Excel**
3. Click **Connect**
4. Navigate to your output workbook: `Bidvest Group Scoping [timestamp].xlsm`
5. Click **Open**

**Step 3: Select Tables**

In the Navigator window, **check** the following tables:
- ✅ Dim_Packs
- ✅ Dim_FSLIs
- ✅ Dim_Thresholds
- ✅ Fact_Amounts
- ✅ Fact_Percentages
- ✅ Fact_Scoping

Click **Load** (bottom right)

*Wait for data to load (1-2 minutes)*

**Step 4: Create Relationships**

1. Click **Model** view (left sidebar, icon looks like three connected boxes)
2. You should see your 6 tables
3. Create the following relationships by **dragging** fields between tables:

**Relationship 1:**
- Drag `Fact_Amounts[PackCode]` to `Dim_Packs[PackCode]`
- Cardinality: Many-to-One (*)
- Cross filter direction: Single

**Relationship 2:**
- Drag `Fact_Amounts[FSLI]` to `Dim_FSLIs[FSLI]`
- Cardinality: Many-to-One (*)
- Cross filter direction: Single

**Relationship 3:**
- Drag `Fact_Percentages[PackCode]` to `Dim_Packs[PackCode]`
- Cardinality: Many-to-One (*)
- Cross filter direction: Single

**Relationship 4:**
- Drag `Fact_Percentages[FSLI]` to `Dim_FSLIs[FSLI]`
- Cardinality: Many-to-One (*)
- Cross filter direction: Single

**Relationship 5:**
- Drag `Fact_Scoping[PackCode]` to `Dim_Packs[PackCode]`
- Cardinality: Many-to-One (*)
- Cross filter direction: Single

**Relationship 6:**
- Drag `Fact_Scoping[FSLI]` to `Dim_FSLIs[FSLI]`
- Cardinality: Many-to-One (*)
- Cross filter direction: Single

✅ **You should now have a star schema with 3 dimension tables and 3 fact tables**

**Step 5: Create DAX Measures**

1. Click **Data** view (left sidebar, table icon)
2. Right-click on **Fact_Amounts** table
3. Select **New Measure**
4. Enter the following measures one by one:

```dax
// Total Amount
Total Amount = SUM(Fact_Amounts[Amount])
```

```dax
// Scoped Amount
Scoped Amount =
CALCULATE(
    SUM(Fact_Amounts[Amount]),
    Fact_Scoping[ScopingStatus] = "Scoped In"
)
```

```dax
// Untested Amount
Untested Amount = [Total Amount] - [Scoped Amount]
```

```dax
// Coverage %
Coverage % =
DIVIDE(
    [Scoped Amount],
    [Total Amount],
    0
)
```

```dax
// Total Packs
Total Packs = DISTINCTCOUNT(Dim_Packs[PackCode])
```

```dax
// Packs Scoped In
Packs Scoped In =
CALCULATE(
    DISTINCTCOUNT(Fact_Scoping[PackCode]),
    Fact_Scoping[ScopingStatus] = "Scoped In"
)
```

```dax
// Packs Coverage %
Packs Coverage % =
DIVIDE(
    [Packs Scoped In],
    [Total Packs],
    0
)
```

**Step 6: Create Visualizations**

1. Click **Report** view (left sidebar, bar chart icon)
2. Create your first visualization:

**Example: Coverage by FSLI (Bar Chart)**

1. Add visualization: **Clustered Bar Chart**
2. Fields:
   - Y-axis: `Dim_FSLIs[FSLI]`
   - X-axis: `[Coverage %]`
3. Format:
   - Data labels: On
   - Sort: By Coverage % (ascending)

**Example: Pack Scoping Status (Pie Chart)**

1. Add visualization: **Pie Chart**
2. Fields:
   - Legend: `Fact_Scoping[ScopingStatus]`
   - Values: `[Packs Scoped In]`

**Example: Division Coverage (Table)**

1. Add visualization: **Table**
2. Fields:
   - `Dim_Packs[Division]`
   - `[Total Amount]`
   - `[Scoped Amount]`
   - `[Coverage %]`

Continue building visualizations as needed.

**Step 7: Save Power BI File**

1. File → Save As
2. Save as: `Bidvest Scoping Analysis [date].pbix`

### 7.4 Manual Scoping in Power BI

Power BI scoping is **read-only**. To update scoping:

1. Update scoping in Excel (Manual Scoping Interface sheet)
2. Save Excel workbook
3. In Power BI: Home → Refresh
4. Data updates automatically

---

## 8. Technical Reference

### 8.1 Module Architecture

The tool comprises **8 VBA modules**:

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| **Mod1_MainController** | Main entry point, workflow orchestration | `StartBidvestScopingTool()` |
| **Mod2_TabProcessing** | Tab discovery, categorization, validation | `CategorizeAllTabs()` |
| **Mod3_DataExtraction** | Data extraction, table generation | `GenerateFullInputTables()` |
| **Mod4_SegmentalMatching** | Pack matching, fuzzy matching | `ProcessSegmentalWorkbook()` |
| **Mod5_ScopingEngine** | Threshold configuration, automatic scoping | `ConfigureThresholds()`, `ApplyThresholds()` |
| **Mod6_DashboardGeneration** | Dashboard creation, 5 views | `CreateComprehensiveDashboard()` |
| **Mod7_PowerBIExport** | Power BI table generation | `CreatePowerBIAssets()` |
| **Mod8_Utilities** | Common functions, validation | `GetWorkbookByName()`, formatting functions |

### 8.2 Data Flow

```
Stripe Packs Workbook
        ↓
Tab Categorization → g_TabCategories (Dictionary)
        ↓
Currency Selection → g_UseConsolidationCurrency (Boolean)
        ↓
Consolidation Entity ID → g_ConsolidationEntity (String)
        ↓
Data Extraction → Full Input Table, Percentage Table
        ↓
Threshold Config → g_ThresholdFSLIs (Collection)
        ↓
Automatic Scoping → g_ScopedPacks (Dictionary)
        ↓
Dashboard Generation → 5 Dashboard Sheets
        ↓
Power BI Assets → 6 Tables (Dim + Fact)
        ↓
Output Workbook Saved
```

### 8.3 Global Variables

```vba
' Workbook References
g_StripePacksWorkbook As Workbook
g_SegmentalWorkbook As Workbook
g_OutputWorkbook As Workbook

' Categorization & Configuration
g_TabCategories As Object          ' Tab name → Category
g_DivisionNames As Object          ' Tab name → Division name
g_UseConsolidationCurrency As Boolean

' Consolidation Entity
g_ConsolidationEntity As String
g_ConsolidationEntityName As String

' Scoping
g_ThresholdFSLIs As Collection     ' Threshold configs
g_ScopedPacks As Object            ' Scoped pack codes
g_ManualScoping As Object          ' Manual scoping decisions
```

### 8.4 Configuration Constants

All configuration constants are in **Mod8_Utilities**:

```vba
' Row Structure
ROW_CURRENCY_TYPE = 6
ROW_PACK_NAME = 7
ROW_PACK_CODE = 8
ROW_FSLI_START = 9

' Version
TOOL_VERSION = "6.0"
TOOL_NAME = "Bidvest Group ISA 600 Scoping Tool"
```

### 8.5 Customization Points

**To customize threshold FSLIs:**
- Modify `Mod5_ScopingEngine.ConfigureThresholds()`

**To customize dashboard views:**
- Modify `Mod6_DashboardGeneration` module functions

**To add new data tables:**
- Add function in `Mod3_DataExtraction`
- Call from `Mod1_MainController.ExtractAndGenerateTables()`

**To customize Power BI tables:**
- Modify `Mod7_PowerBIExport` module functions

---

## 9. Troubleshooting

### 9.1 Common Issues

**Issue 1: "Could not find workbook"**

**Cause:** Workbook name doesn't match or workbook not open

**Solution:**
1. Ensure the workbook is open in Excel
2. Copy the exact name from the title bar (including extension)
3. Try both with and without .xlsx/.xlsm extension

---

**Issue 2: "Required tabs are missing"**

**Cause:** No tab was categorized as "Input Continuing"

**Solution:**
1. At least ONE tab must be category 3 (Input Continuing)
2. Re-run the tool and ensure you categorize the primary data tab as "Input Continuing"

---

**Issue 3: "No entities found in Input Continuing tab"**

**Cause:** Tab structure doesn't match expected format

**Solution:**
1. Verify Row 6 contains currency type labels
2. Verify Row 7 contains pack names
3. Verify Row 8 contains pack codes
4. Ensure columns start from column C or later

---

**Issue 4: Tool runs slowly or freezes**

**Cause:** Large dataset or insufficient memory

**Solution:**
1. Close other applications to free memory
2. Ensure at least 4GB RAM available
3. For very large datasets (>500 packs), consider splitting into divisions

---

**Issue 5: "Subscript out of range" error**

**Cause:** Referenced worksheet doesn't exist

**Solution:**
1. Ensure all required tabs were categorized
2. Check that tab names don't contain special characters
3. Verify workbook wasn't modified during processing

---

**Issue 6: No Power BI relationships created**

**Cause:** Table structure issue

**Solution:**
1. Verify all 6 tables were imported (Dim_Packs, Dim_FSLIs, Dim_Thresholds, Fact_Amounts, Fact_Percentages, Fact_Scoping)
2. Check that key columns (PackCode, FSLI) exist in all relevant tables
3. Manually create relationships using Model view

---

**Issue 7: Coverage percentages show 0% or N/A**

**Cause:** Consolidation entity not found or no scoping applied

**Solution:**
1. Verify consolidation entity was correctly identified
2. Ensure at least some packs were scoped in (automatic or manual)
3. Check that Fact_Scoping table contains data

---

**Issue 8: Macro security warning**

**Cause:** Macros are disabled

**Solution:**
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Enable all macros (or "Disable with notification")
3. Click "Enable Content" when opening the file

---

### 9.2 Getting Help

If you encounter issues not covered here:

1. **Check VBA Immediate Window**
   - Press `Ctrl + G` in VBA Editor
   - Look for debug messages

2. **Review Error Messages**
   - Note exact error message and error number
   - Check which module/function generated the error

3. **Test with Sample Data**
   - Try running with a smaller, simplified workbook
   - Isolate the problematic step

4. **Check Module Documentation**
   - Each module has detailed comments
   - Review function headers for expected inputs

---

## 10. ISA 600 Compliance

### 10.1 ISA 600 Revised Requirements

**ISA 600 (Revised) - Special Considerations—Audits of Group Financial Statements (Including the Work of Component Auditors)**

The tool addresses key ISA 600 requirements:

**Requirement:** Identify components and assess their significance

✅ **Tool Implementation:**
- Extracts all components (packs) from consolidation workbook
- Calculates materiality as % of consolidated totals
- Identifies significant components via threshold analysis

---

**Requirement:** Determine type of work to be performed on components

✅ **Tool Implementation:**
- Threshold-based automatic scoping for significant components
- Manual scoping interface for judgmental decisions
- Tracks scoping status per component per FSLI

---

**Requirement:** Understand component auditor and their work

✅ **Tool Implementation:**
- Division-Segment mapping for component identification
- Tracks which division/segment each component belongs to
- Enables division/segment-level analysis

---

**Requirement:** Determine appropriate level of involvement in component auditors' work

✅ **Tool Implementation:**
- Coverage analysis by division and segment
- Identifies untested portions by division/segment
- Supports risk-based approach to component selection

---

**Requirement:** Consolidation process understanding

✅ **Tool Implementation:**
- Identifies and excludes consolidation entity (aggregate)
- Processes consolidation journals
- Reconciles division tabs with Input Continuing

---

### 10.2 Component Significance Thresholds

ISA 600 doesn't prescribe exact thresholds. Common approaches:

| Approach | Threshold Examples |
|----------|-------------------|
| **Quantitative** | Revenue > 10% of group revenue<br>OR Total Assets > 10% of group assets<br>OR PBT > 10% of group PBT |
| **Qualitative** | Specific risks<br>Complex transactions<br>New entities |
| **Hybrid** | Combination of quantitative and qualitative |

**Tool Support:**
- Quantitative: Threshold configuration (any % or absolute amount)
- Qualitative: Manual scoping interface for judgment-based decisions

### 10.3 Coverage Targets

Group audit teams typically target:

- **≥ 60%** coverage for each material FSLI (minimum)
- **≥ 80%** coverage for high-risk FSLIs (best practice)
- **≥ 90%** coverage for key performance indicators

The tool's dashboard clearly shows coverage by FSLI, enabling monitoring against targets.

### 10.4 Documentation Requirements

ISA 600 requires documentation of:

✅ **Component selection:** Threshold Configuration sheet + Scoping Summary
✅ **Significance assessments:** Full Input Percentage Table
✅ **Type of work planned:** Fact_Scoping table (ScopingStatus)
✅ **Coverage analysis:** Dashboard views (Coverage by FSLI, Division, Segment)

**Audit File Inclusions:**
1. Output workbook (timestamped filename provides audit trail)
2. Screenshots of dashboard views
3. Power BI report (if used)
4. Threshold configuration documentation

### 10.5 Group Engagement Team Responsibilities

Per ISA 600, the group engagement team must:

1. **Take responsibility for direction, supervision, and performance**
   - Tool provides data to support informed decisions
   - Does not replace professional judgment

2. **Determine materiality**
   - Tool calculates contributions but doesn't set materiality
   - Auditor determines thresholds based on group materiality

3. **Review component auditors' work**
   - Tool identifies which components are scoped
   - Auditor performs actual review procedures

4. **Evaluate audit evidence**
   - Tool provides analysis framework
   - Auditor evaluates sufficiency and appropriateness

**Important:** This tool is a **decision support system**. It does not replace auditor judgment or relieve the group engagement team of their responsibilities under ISA 600.

---

## Appendix A: Quick Reference

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Open VBA Editor | `Alt + F11` |
| Close VBA Editor | `Alt + Q` |
| Immediate Window | `Ctrl + G` |
| Find | `Ctrl + F` |
| Save | `Ctrl + S` |

### File Locations

| Item | Default Location |
|------|------------------|
| Tool Workbook | Desktop or Documents (your choice) |
| Output Workbook | Same directory as Stripe Packs workbook |
| VBA Modules | VBA_Modules folder in repository |

### Color Codes

| Color | Meaning | Usage |
|-------|---------|-------|
| 🟢 Green | Good/Scoped In | Coverage ≥ 80%, scoped packs |
| 🟡 Yellow | Review/Moderate | Coverage 60-79%, not yet scoped |
| 🔴 Red | Insufficient | Coverage < 60%, attention needed |
| 🔵 Blue | Headers | Column/section headers |

### Typical Processing Times

| Dataset Size | Expected Time |
|--------------|---------------|
| Small (5 tabs, 50 packs, 100 FSLIs) | 2-3 minutes |
| Medium (10 tabs, 150 packs, 200 FSLIs) | 5-7 minutes |
| Large (20 tabs, 300 packs, 300 FSLIs) | 10-15 minutes |

---

## Appendix B: Glossary

**Component:** An entity or business unit included in the group financial statements (ISA 600 terminology). In this tool, "component" = "pack" = "entity".

**Consolidation Entity:** The aggregate entity representing the entire group (e.g., Bidvest Group Consolidated). This is the 100% baseline.

**Consolidation Currency:** The group reporting currency (e.g., ZAR for Bidvest). Used for consolidation purposes.

**Entity Currency:** The local functional currency of each individual entity.

**FSLI (Financial Statement Line Item):** Individual line items from the financial statements (e.g., Revenue, Total Assets, Inventory).

**Pack:** Bidvest terminology for an entity/component. Synonymous with "entity" or "component".

**Scoped In:** A pack/FSLI combination included in the audit scope (will be tested).

**Scoped Out:** A pack/FSLI combination explicitly excluded from scope.

**Not Scoped:** A pack/FSLI combination not yet determined (default state).

**Segmental Reporting:** IAS 8 segmental disclosure report showing how packs roll up into segments.

**Stripe Packs:** Bidvest terminology for the consolidation workbook format.

**Threshold:** A quantitative criterion (amount or %) used to automatically identify significant components.

---

## Appendix C: Change Log

### Version 6.0 (November 2025) - Complete Overhaul

**Major Changes:**
- ✅ Complete redesign from v5.x
- ✅ Ensured "Consol" naming (not "Console") throughout
- ✅ 8 modular VBA modules (clean architecture)
- ✅ Comprehensive user guidance at every step
- ✅ 5 interactive dashboard views
- ✅ Full Power BI integration with 6 tables
- ✅ Manual scoping at pack and FSLI level
- ✅ Division-Segment mapping with fuzzy matching
- ✅ Timestamped output filenames
- ✅ Single comprehensive implementation guide

**New Features:**
- Interactive dashboard with 5 views
- Manual scoping interface with dropdowns
- Power BI-ready dimension and fact tables
- Coverage analysis by FSLI, Division, Segment
- Fuzzy matching for pack reconciliation
- Efficiency metrics (impact of scoping decisions)
- Drill-down capability from summary to detail

**Improvements:**
- Clearer prompts and user guidance
- Better error handling
- Professional formatting throughout
- Comprehensive documentation
- Modular code structure (maintainable)
- Performance optimizations

---

## Conclusion

This tool provides a comprehensive, production-ready solution for ISA 600 component scoping at Bidvest Group Limited. It automates complex data processing, provides intuitive dashboards, enables flexible scoping decisions, and produces audit-ready documentation.

**Key Takeaways:**
- ✅ No technical setup required - works immediately
- ✅ Guided workflow with clear prompts
- ✅ Flexible: automatic thresholds + manual fine-tuning
- ✅ Comprehensive: 5 dashboard views + Power BI integration
- ✅ Audit-ready: Professional output with full audit trail
- ✅ ISA 600 compliant: Addresses all key requirements

**Next Steps:**
1. Install the tool (Section 3)
2. Run through the guided workflow (Section 4)
3. Review dashboards (Section 5)
4. Adjust scoping as needed (Section 6)
5. Import to Power BI for enhanced analysis (Section 7)

**Questions or Issues?**
- Consult Section 9 (Troubleshooting)
- Review module documentation (comments in VBA code)
- Check ISA 600 compliance guidance (Section 10)

---

**Document Version:** 1.0
**Tool Version:** 6.0
**Last Updated:** November 2025
**Maintained By:** Bidvest Group Audit Team

---

*End of Comprehensive Implementation Guide*
