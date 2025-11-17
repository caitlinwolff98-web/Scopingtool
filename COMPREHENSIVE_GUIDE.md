# Bidvest Group Limited - ISA 600 Consolidation Scoping Tool
## Comprehensive Guide & Documentation

**Version:** 4.0 (Complete Overhaul)  
**Last Updated:** November 2024  
**Purpose:** ISA 600 Revised Compliance for Bidvest Group Limited Consolidations  
**Platform:** Microsoft Excel VBA + Power BI Desktop

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [System Overview](#2-system-overview)
3. [Installation & Setup](#3-installation--setup)
4. [VBA Tool Usage](#4-vba-tool-usage)
5. [Power BI Integration](#5-power-bi-integration)
6. [Manual Scoping Workflow](#6-manual-scoping-workflow)
7. [ISA 600 Compliance](#7-isa-600-compliance)
8. [Troubleshooting](#8-troubleshooting)
9. [Technical Reference](#9-technical-reference)

## Additional Guides

- **[VISUALIZATION_ALTERNATIVES.md](VISUALIZATION_ALTERNATIVES.md)** - Complete evaluation of Power BI vs. alternatives (Tableau, Qlik, Excel, Python, etc.)
- **[POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md)** - Step-by-step guide to enable manual data entry in Power BI (with troubleshooting)

---

## 1. Executive Summary

### What This Tool Does

The Bidvest Scoping Tool automates the ISA 600 revised compliance process for Bidvest Group Limited consolidation audits by:

✅ **Extracting** all Financial Statement Line Items (FSLIs) from consolidation workbooks  
✅ **Identifying** pack codes, pack names, and divisions automatically  
✅ **Applying** threshold-based automatic scoping (e.g., "Any pack where Revenue > R300M")  
✅ **Generating** Power BI-compatible tables for dynamic analysis  
✅ **Enabling** manual pack/FSLI scoping with real-time coverage updates  
✅ **Tracking** scoped vs. unscoped percentages by FSLI and Division  
✅ **Ensuring** consolidated entity is excluded from scoping calculations  
✅ **Providing** audit-quality documentation and reporting

### Key Benefits

- **Saves Time:** Automates hours of manual consolidation analysis
- **Ensures Accuracy:** Eliminates manual data entry errors
- **ISA 600 Compliant:** Meets revised ISA 600 requirements for group audits
- **Dynamic Analysis:** Real-time updates in Power BI as you scope packs
- **Audit Trail:** Complete documentation of scoping decisions

### Quick Start (3 Steps)

1. **Run VBA Macro** → Select consolidated entity → Configure thresholds (optional)
2. **Import to Power BI** → Load generated Excel tables
3. **Manual Scoping** → Select packs/FSLIs in Power BI to refine coverage

---

## 2. System Overview

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                 CONSOLIDATION WORKBOOK                      │
│  (Excel file with Input Continuing, Journals, Consol tabs) │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│                    VBA SCOPING TOOL                         │
│  • FSLI Extraction (with Notes cutoff)                     │
│  • Pack Identification (names + codes)                     │
│  • Consolidated Entity Selection                           │
│  • Threshold-Based Auto-Scoping                           │
│  • Table Generation (Power BI ready)                       │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│              OUTPUT: Bidvest Scoping Tool Output.xlsx       │
│  • Full Input Table                                         │
│  • Full Input Percentage Table                             │
│  • Journals Table + Percentage                             │
│  • Consol Table + Percentage                               │
│  • Discontinued Table + Percentage                         │
│  • FSLi Key Table                                          │
│  • Pack Number Company Table                               │
│  • Scoping Control Table (for Power BI)                   │
│  • Threshold Configuration (if applied)                    │
│  • Scoping Summary                                         │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│                    POWER BI DESKTOP                         │
│  • Import Excel Tables                                      │
│  • Create Relationships                                     │
│  • Build DAX Measures                                      │
│  • Manual Pack/FSLI Scoping                                │
│  • Real-Time Coverage Analysis                             │
│  • Division-Level Reporting                                │
│  • Export Results                                          │
└─────────────────────────────────────────────────────────────┘
```

### Data Flow

**Input → VBA Processing → Power BI Analysis → Audit Documentation**

### Core Components

1. **VBA Modules (8 total)**
   - ModMain: Orchestration and entry point
   - ModConfig: Configuration and utilities
   - ModTabCategorization: Tab discovery and categorization
   - ModDataProcessing: FSLI extraction and data processing
   - ModTableGeneration: Table creation and formatting
   - ModThresholdScoping: Automatic threshold-based scoping
   - ModInteractiveDashboard: Excel-based analysis dashboard
   - ModPowerBIIntegration: Power BI metadata and scoping control

2. **Excel Output Tables**
   - Data tables: Input, Journals, Consol, Discontinued
   - Percentage tables: Coverage analysis for each data table
   - Reference tables: FSLi Key, Pack Number Company
   - Scoping tables: Summary, Control, Threshold Config
   - Reports: Division-based Scoped In/Out, Detail

3. **Power BI Dashboard**
   - Coverage metrics and KPIs
   - Interactive pack/FSLI selection
   - Division-level analysis
   - Real-time scoping updates

---

## 3. Installation & Setup

### Prerequisites

**Software Requirements:**
- Windows 10 or later
- Microsoft Excel 2016 or later (with VBA enabled)
- Power BI Desktop (latest version recommended)
- 8GB RAM recommended
- Screen resolution: 1920x1080 or higher recommended

**PwC Environment:**
- Standard Microsoft Office environment
- No SQL database required
- No internet access required for VBA tool
- Power BI Desktop approved for use

### Step 1: Install VBA Tool

#### 1.1 Create Macro Workbook

1. Open Microsoft Excel
2. Create a new blank workbook
3. Save as: `Bidvest_Scoping_Tool.xlsm`
   - File Type: **Excel Macro-Enabled Workbook (.xlsm)**
   - Location: Recommended - Documents folder or dedicated audit folder

#### 1.2 Enable VBA Developer Tab

If the Developer tab is not visible:
1. File → Options → Customize Ribbon
2. Check ✅ "Developer" on the right panel
3. Click OK

#### 1.3 Import VBA Modules

1. Press `Alt + F11` to open VBA Editor
2. In VBA Editor: File → Import File
3. Navigate to the `VBA_Modules` folder
4. Import modules **in this order**:
   - `ModConfig.bas` (import FIRST - dependencies)
   - `ModMain.bas`
   - `ModTabCategorization.bas`
   - `ModDataProcessing.bas`
   - `ModTableGeneration.bas`
   - `ModThresholdScoping.bas`
   - `ModInteractiveDashboard.bas`
   - `ModPowerBIIntegration.bas`

**To Import Each Module:**
- File → Import File
- Select the `.bas` file
- Click Open
- Repeat for all 8 modules

#### 1.4 Verify Import

1. In VBA Editor, look at the left panel (Project Explorer)
2. You should see 8 modules listed under "Modules"
3. Double-click each module to verify code is present
4. Debug → Compile VBAProject (should have no errors)

#### 1.5 Create Start Button

1. Return to Excel (close VBA Editor or press Alt + F11)
2. Developer → Insert → Button (Form Control)
3. Draw a button on the worksheet
4. In "Assign Macro" dialog, select: **StartScopingTool**
5. Click OK
6. Right-click button → Edit Text
7. Change text to: "Start Bidvest Scoping Tool"
8. Save the workbook

### Step 2: Test Installation

1. Open a test consolidation workbook (or your actual consolidation file)
2. Keep both workbooks open
3. Click "Start Bidvest Scoping Tool" button
4. If welcome dialog appears, installation is successful
5. Click Cancel to exit test

---

## 4. VBA Tool Usage

### Complete Workflow

#### Step 1: Prepare Your Environment

**Before Running the Tool:**

1. **Open consolidation workbook** (e.g., `Bidvest_Consolidation_2024.xlsx`)
2. **Open VBA tool workbook** (`Bidvest_Scoping_Tool.xlsm`)
3. **Ensure both workbooks are open simultaneously**
4. **Note the exact name** of your consolidation workbook (including .xlsx or .xlsm)

**Consolidation Workbook Requirements:**
- Must contain "Input Continuing" tab (mandatory)
- Row 6: Currency type identifiers
- Row 7: Pack names
- Row 8: Pack codes
- Row 9+: FSLI data
- Column B: FSLI names

#### Step 2: Start the Tool

1. Click "Start Bidvest Scoping Tool" button
2. Read welcome message
3. Click **OK** to continue

#### Step 3: Enter Workbook Name

**Prompt:** "Please enter the exact name of the TGK consolidation workbook"

**Instructions:**
1. Switch to consolidation workbook
2. Look at title bar at top of Excel window
3. Copy the exact filename including extension
4. Examples:
   - ✅ `Bidvest_Consolidation_2024_Q4.xlsx`
   - ✅ `TGK_Consol_Dec2024.xlsm`
   - ❌ `Consolidation` (missing extension)
   - ❌ `bidvest consolidation` (wrong case/spaces)

5. Paste into InputBox
6. Click **OK**

**Troubleshooting:**
- If error "Could not find workbook", verify exact spelling
- Ensure workbook is open
- Check if extension is included

#### Step 4: Categorize Tabs

The tool will discover all tabs and prompt you to categorize each one.

**Available Categories:**

| # | Category | Quantity | Description |
|---|----------|----------|-------------|
| 1 | TGK Segment Tabs | Multiple | Business divisions/segments (e.g., UK, US, Europe) |
| 2 | Discontinued Ops Tab | Single | Discontinued operations |
| 3 | TGK Input Continuing Tab | **Single (REQUIRED)** | Primary input data - MOST IMPORTANT |
| 4 | TGK Journals Continuing Tab | Single | Consolidation journals |
| 5 | TGK Consol Continuing Tab | Single | Consolidated outputs |
| 6 | TGK BS Tab | Single | Balance Sheet |
| 7 | TGK IS Tab | Single | Income Statement |
| 8 | Paul workings | Multiple | Working papers |
| 9 | Trial Balance | Single | Trial balance |
| 10 | Uncategorized | Multiple | Tabs to ignore |

**Categorization Process:**

For each tab, you'll see:
```
Tab: "UK_Division"

Enter category number (1-10):
1. TGK Segment Tabs (multiple allowed)
2. Discontinued Ops Tab (single only)
3. TGK Input Continuing Operations Tab (single only) ★ REQUIRED
...
```

**How to Categorize:**

1. **Input Continuing Tab** (Category 3) - **REQUIRED**
   - This is your primary consolidation data
   - Contains all packs and FSLIs
   - Must categorize exactly ONE tab as this
   - Example tab names: "Input Continuing", "Input_Cont", "TGK Input"

2. **Segment Tabs** (Category 1) - Optional but recommended
   - Each business division (UK, US, Europe, etc.)
   - Can categorize multiple tabs
   - When prompted, enter division name (e.g., "UK Division")
   - If blank, auto-names as "Division_1", "Division_2", etc.

3. **Journals Tab** (Category 4) - Recommended if present
   - Consolidation journal entries
   - One tab only
   - Example names: "Journals", "Consol Journals"

4. **Consol Tab** (Category 5) - Recommended if present
   - Consolidated outputs (Input + Journals)
   - One tab only
   - Example names: "Console", "Consol Continuing"

5. **Other Categories** - Categorize as appropriate
   - Discontinued (if applicable)
   - Balance Sheet / Income Statement (if separate tabs)
   - Working papers (Paul workings)
   - Uncategorized (for tabs to ignore)

**Validation:**

After categorization:
- Tool validates categories
- Ensures Input Continuing exists (mandatory)
- Checks single-tab categories have only one tab
- Shows uncategorized tabs for confirmation

If validation fails:
- Option to restart categorization
- Option to cancel process

#### Step 5: Select Consolidated Entity

**Purpose:** ISA 600 requires excluding the consolidated entity from scoping calculations (it represents totals, not individual entities).

**Prompt:**
```
CONSOLIDATED ENTITY SELECTION

Select which pack represents the CONSOLIDATED entity.
This pack will be EXCLUDED from scoping calculations
as it represents consolidated totals, not individual entities.

Available Packs:
------------------------------------------------------------
1. The Bidvest Group Consolidated (BVT-001)
2. Bidvest UK Limited (BVT-101)
3. Bidvest US Inc (BVT-201)
...

Enter the number of the consolidated pack:
(Or leave blank to include all packs in scoping)
```

**Instructions:**
1. Review list of packs
2. Identify which pack is the consolidated total
   - Usually named "The Bidvest Group Consolidated" or similar
   - Typically has lowest pack code (e.g., BVT-001)
3. Enter the number (e.g., `1`)
4. Confirm selection
5. This pack will be **excluded** from all scoping calculations

**Why This Matters:**
- Consolidated entity = sum of all other entities
- Including it in scoping would double-count
- ISA 600 requires component identification, not consolidated total

#### Step 6: Select Column Type

**Prompt:** "Use Consolidation Currency columns? (Recommended)"

**Options:**
- **YES** (Recommended): Uses consolidation currency (e.g., ZAR, USD)
- **NO**: Uses entity/original currency

**Recommendation:** Always select **YES** for consistency in scoping analysis.

#### Step 7: Configure Threshold Scoping (Optional)

**Prompt:** "Would you like to configure threshold-based automatic scoping?"

**If YES:**

**7a. Select FSLIs for Thresholds**

Dialog will show list of all FSLIs:
```
Select FSLIs for Threshold Analysis

Available FSLIs:
1. Net Revenue
2. Total Revenue
3. Total Assets
4. Total Liabilities
5. Net Profit
...

Enter FSLI name or number (or leave blank to finish):
```

**Instructions:**
1. Enter FSLI name (e.g., "Net Revenue") or number (e.g., "1")
2. Press Enter
3. Repeat for additional FSLIs
4. Leave blank when done

**Example Selections:**
- Net Revenue (for top-line scoping)
- Total Assets (for Balance Sheet scoping)
- Net Profit (for profitability scoping)

**7b. Enter Threshold Values**

For each selected FSLI:
```
Enter threshold for "Net Revenue"

This will automatically scope in any pack where
the absolute value of Net Revenue exceeds the threshold.

Threshold Amount: _______
```

**Instructions:**
1. Enter threshold amount (e.g., `300000000` for R300M)
2. Do NOT include currency symbols or commas
3. Enter as a number only

**What Happens:**
- Tool analyzes all packs
- Any pack where selected FSLI > threshold = **automatically scoped in**
- **ENTIRE pack is scoped in** (all FSLIs for that pack)
- Consolidated entity is excluded from threshold analysis
- Configuration documented in output workbook

**If NO:** Skip to Step 8 (no automatic scoping applied)

#### Step 8: Processing

Tool will now process data:

**Status Updates:**
```
Processing Input Continuing tab...
Processing Journals tab...
Processing Consol tab...
Creating FSLi Key Table...
Creating Pack Number Company Table...
Creating Percentage Tables...
Creating scoping summary...
Creating division-based reports...
Creating interactive dashboard...
Creating Power BI integration assets...
```

**Processing Time:**
- Small workbook (5 tabs, 200 FSLIs): 30-60 seconds
- Medium workbook (10 tabs, 500 FSLIs): 2-4 minutes
- Large workbook (20+ tabs, 1000+ FSLIs): 5-10 minutes

**What's Happening:**
1. Unmerging cells
2. Detecting pack codes and names
3. Extracting FSLIs (stopping at "Notes" section)
4. Creating data tables
5. Calculating percentages
6. Generating scoping metadata
7. Creating Power BI integration tables

#### Step 9: Review Output

**Success Message:**
```
Scoping tool completed successfully!

Output saved as: Bidvest Scoping Tool Output.xlsx
Location: [same folder as source workbook]

Generated assets:
- Data tables for analysis
- Threshold configuration (if applied)
- Scoping summary with recommendations
- Division-based scoping reports
- Interactive Excel dashboard
- Power BI integration metadata

The workbook can be used standalone or with Power BI!
```

**Output Workbook Structure:**

📁 **Bidvest Scoping Tool Output.xlsx**
- Control Panel (information sheet)
- Full Input Table ⭐ (primary data)
- Full Input Percentage ⭐
- Journals Table
- Journals Percentage
- Full Consol Table
- Full Consol Percentage
- Discontinued Table (if applicable)
- Discontinued Percentage
- FSLi Key Table
- Pack Number Company Table
- Scoping Control Table ⭐ (for Power BI)
- Scoping Summary
- Threshold Configuration (if thresholds applied)
- Scoped In by Division
- Scoped Out by Division
- Scoped In Packs Detail
- Interactive Dashboard
- Scoping Calculator
- PowerBI_Metadata
- PowerBI_Scoping
- DAX Measures Guide

**⭐ = Most important for Power BI integration**

### FSLI Extraction Logic (CRITICAL)

**How It Works:**

The tool intelligently extracts FSLIs from Column B of "Input Continuing" tab with the following logic:

**INCLUDED FSLIs:**
✅ Actual line items (e.g., "Revenue", "Cost of Sales")
✅ Items with brackets indicating sub-items (e.g., "(a) Group companies")
✅ Items with numerical data across entity columns
✅ Totals and subtotals (marked in metadata)

**EXCLUDED FSLIs:**
❌ Empty rows (no data across entities)
❌ Statement headers: "INCOME STATEMENT", "BALANCE SHEET"
❌ Section headers: "ASSETS", "LIABILITIES", "EQUITY"
❌ **"NOTES" section and everything below it**

**Dynamic Detection:**
- Automatically detects FSLI hierarchy (indentation levels)
- Identifies totals vs. line items
- Stops at "Notes" row
- Validates row contains financial data before including

**Code Reference:** `ModDataProcessing.AnalyzeFSLiStructure()`

**Testing FSLI Extraction:**
1. Review "FSLi Key Table" in output workbook
2. Verify no statement headers present
3. Confirm "Notes" section is excluded
4. Check that all actual line items are included

---

## 5. Power BI Integration

### Overview

Power BI provides dynamic, interactive scoping analysis with real-time coverage updates.

**📚 Additional Resources:**
- **[VISUALIZATION_ALTERNATIVES.md](VISUALIZATION_ALTERNATIVES.md)** - Why Power BI? Comparison with Tableau, Qlik, Excel, Python, and other tools
- **[POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md)** - Detailed edit mode setup (required for manual scoping)

**Why Power BI?** See VISUALIZATION_ALTERNATIVES.md for complete evaluation showing Power BI is optimal because:
- ✅ Free (Desktop version)
- ✅ PwC-approved
- ✅ Edit mode for manual scoping (critical feature)
- ✅ Excellent Excel integration
- ✅ DAX calculation engine

### Step 1: Import Excel Tables

**Launch Power BI Desktop**

1. Open Power BI Desktop
2. Home → Get Data → Excel
3. Browse to: **Bidvest Scoping Tool Output.xlsx**
4. Select **ALL** of the following tables:

**Primary Tables** (select all):
- ✅ Full_Input_Table
- ✅ Full_Input_Percentage
- ✅ Journals_Table
- ✅ Journals_Percentage
- ✅ Full_Consol_Table
- ✅ Full_Consol_Percentage
- ✅ Discontinued_Table (if present)
- ✅ Discontinued_Percentage (if present)
- ✅ FSLi_Key_Table
- ✅ Pack_Number_Company_Table
- ✅ Scoping_Control_Table ⭐ **CRITICAL for manual scoping**
- ✅ Scoping_Summary_Table

5. Click **Transform Data** (opens Power Query Editor)

### Step 2: Transform Tables (Power Query)

**Important:** Data tables need to be unpivoted for Power BI analysis.

#### 2a. Unpivot Full Input Table

1. In Power Query, select **Full_Input_Table**
2. Select first column (Pack Code and Pack Name)
3. Right-click → **Unpivot Other Columns**
4. Rename columns:
   - "Attribute" → **"FSLI"**
   - "Value" → **"Amount"**
5. Change Amount data type to **Decimal Number**
6. Filter out null/empty amounts (optional)

**Before Unpivot:**
```
| Pack Code | Pack Name | Revenue | Cost of Sales | ... |
|-----------|-----------|---------|---------------|-----|
| BVT-101   | UK Entity | 1000000 | 600000        | ... |
```

**After Unpivot:**
```
| Pack Code | Pack Name | FSLI          | Amount  |
|-----------|-----------|---------------|---------|
| BVT-101   | UK Entity | Revenue       | 1000000 |
| BVT-101   | UK Entity | Cost of Sales | 600000  |
```

#### 2b. Repeat for Other Data Tables

Repeat unpivot process for:
- Full_Input_Percentage
- Journals_Table
- Journals_Percentage
- Full_Consol_Table
- Full_Consol_Percentage
- Discontinued_Table (if present)
- Discontinued_Percentage (if present)

**Tip:** Use Power Query's "Duplicate Query" feature to speed this up.

#### 2c. Leave Reference Tables As-Is

Do NOT unpivot:
- FSLi_Key_Table
- Pack_Number_Company_Table
- Scoping_Control_Table
- Scoping_Summary_Table

#### 2d. Apply & Close

Click **Close & Apply** to load all tables into Power BI data model.

### Step 3: Create Relationships

**Navigate to:** Model View (left sidebar, 3rd icon)

#### Create These Relationships:

**Pack Relationships:**
```
Pack_Number_Company_Table[Pack Code] 
    → Full_Input_Table[Pack Code] (Many-to-One)
    
Pack_Number_Company_Table[Pack Code]
    → Journals_Table[Pack Code] (Many-to-One)
    
Pack_Number_Company_Table[Pack Code]
    → Full_Consol_Table[Pack Code] (Many-to-One)
    
Pack_Number_Company_Table[Pack Code]
    → Scoping_Control_Table[Pack Code] (Many-to-One) ⭐
```

**FSLI Relationships:**
```
FSLi_Key_Table[FSLI]
    → Full_Input_Table[FSLI] (One-to-Many)
    
FSLi_Key_Table[FSLI]
    → Journals_Table[FSLI] (One-to-Many)
    
FSLi_Key_Table[FSLI]
    → Scoping_Control_Table[FSLI] (One-to-Many) ⭐
```

**Important Notes:**
- Use **Pack Code** (not Pack Name) for relationships
- Pack Code is unique identifier
- Pack Name can be used for display purposes
- Ensure cross-filter direction is set correctly (One-way recommended)

### Step 4: Create DAX Measures

Open DAX Measures Guide sheet from Excel for reference, or use these measures:

#### Measure 1: Total Packs

```dax
Total Packs = 
DISTINCTCOUNT(Pack_Number_Company_Table[Pack Code])
```

#### Measure 2: Scoped In Packs (Automatic)

```dax
Scoped In Packs (Auto) = 
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Scoped In (Threshold)"
)
```

#### Measure 3: Scoped In Packs (Manual)

```dax
Scoped In Packs (Manual) = 
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Scoped In (Manual)"
)
```

#### Measure 4: Total Scoped In

```dax
Total Scoped In = 
[Scoped In Packs (Auto)] + [Scoped In Packs (Manual)]
```

#### Measure 5: Not Scoped In

```dax
Not Scoped In = 
[Total Packs] - [Total Scoped In]
```

#### Measure 6: Coverage Percentage (by FSLI)

```dax
Coverage % by FSLI = 
VAR SelectedFSLI = SELECTEDVALUE(FSLi_Key_Table[FSLI])
VAR TotalAmount = 
    CALCULATE(
        SUM(Full_Input_Table[Amount]),
        Full_Input_Table[FSLI] = SelectedFSLI,
        Pack_Number_Company_Table[Is Consolidated] <> "Yes"
    )
VAR ScopedAmount = 
    CALCULATE(
        SUM(Full_Input_Table[Amount]),
        Full_Input_Table[FSLI] = SelectedFSLI,
        Pack_Number_Company_Table[Is Consolidated] <> "Yes",
        Scoping_Control_Table[Scoping Status] IN {
            "Scoped In (Threshold)", 
            "Scoped In (Manual)"
        }
    )
RETURN
DIVIDE(ScopedAmount, TotalAmount, 0)
```

#### Measure 7: Untested Percentage

```dax
Untested % = 
1 - [Coverage % by FSLI]
```

#### Measure 8: Coverage % by Division

```dax
Coverage % by Division = 
VAR SelectedDivision = SELECTEDVALUE(Pack_Number_Company_Table[Division])
VAR TotalAmount = 
    CALCULATE(
        SUM(Full_Input_Table[Amount]),
        Pack_Number_Company_Table[Division] = SelectedDivision,
        Pack_Number_Company_Table[Is Consolidated] <> "Yes"
    )
VAR ScopedAmount = 
    CALCULATE(
        SUM(Full_Input_Table[Amount]),
        Pack_Number_Company_Table[Division] = SelectedDivision,
        Pack_Number_Company_Table[Is Consolidated] <> "Yes",
        Scoping_Control_Table[Scoping Status] IN {
            "Scoped In (Threshold)", 
            "Scoped In (Manual)"
        }
    )
RETURN
DIVIDE(ScopedAmount, TotalAmount, 0)
```

**Important:** All measures automatically exclude consolidated entity using filter:
```dax
Pack_Number_Company_Table[Is Consolidated] <> "Yes"
```

### Step 5: Build Dashboard Pages

#### Page 1: Executive Summary

**Visuals:**

1. **Card Visual - Total Packs**
   - Measure: `[Total Packs]`
   - Format: Large font, bold

2. **Card Visual - Scoped In (Auto)**
   - Measure: `[Scoped In Packs (Auto)]`
   - Format: Green background

3. **Card Visual - Scoped In (Manual)**
   - Measure: `[Scoped In Packs (Manual)]`
   - Format: Blue background

4. **Card Visual - Not Scoped**
   - Measure: `[Not Scoped In]`
   - Format: Red background

5. **Donut Chart - Scoping Status**
   - Legend: Scoping_Control_Table[Scoping Status]
   - Values: DISTINCTCOUNT(Scoping_Control_Table[Pack Code])

6. **Bar Chart - Packs by Division**
   - Axis: Pack_Number_Company_Table[Division]
   - Values: [Total Packs], [Total Scoped In]

#### Page 2: FSLI Analysis

**Visuals:**

1. **Slicer - FSLI Selector**
   - Field: FSLi_Key_Table[FSLI]
   - Style: Dropdown or List
   - Multi-select: OFF (single selection)

2. **Card - Coverage % for Selected FSLI**
   - Measure: `[Coverage % by FSLI]`
   - Format: Percentage, 2 decimals

3. **Card - Untested %**
   - Measure: `[Untested %]`
   - Format: Percentage, 2 decimals

4. **Table - Pack Detail for Selected FSLI**
   - Columns:
     - Pack_Number_Company_Table[Pack Name]
     - Pack_Number_Company_Table[Division]
     - Full_Input_Table[Amount]
     - Scoping_Control_Table[Scoping Status]
   - Conditional Formatting: Color code by status

5. **Clustered Bar Chart - Top 10 Packs by Amount**
   - Axis: Pack_Number_Company_Table[Pack Name]
   - Values: Full_Input_Table[Amount]
   - Filter: TOPN 10

#### Page 3: Manual Scoping Control ⭐

**This is the CRITICAL page for dynamic manual scoping**

**Visuals:**

1. **Slicer - Pack Selector**
   - Field: Pack_Number_Company_Table[Pack Name]
   - Style: Dropdown
   - Multi-select: ON

2. **Slicer - FSLI Selector**
   - Field: FSLi_Key_Table[FSLI]
   - Style: Dropdown
   - Multi-select: ON

3. **Slicer - Division Filter**
   - Field: Pack_Number_Company_Table[Division]
   - Style: Dropdown
   - Multi-select: ON

4. **Table - Scoping Control**
   - Data Source: Scoping_Control_Table
   - Columns (editable):
     - Pack Name
     - Pack Code
     - Division
     - FSLI
     - Amount
     - **Scoping Status** ⭐ (enable editing)
   - Allow editing: YES
   - Users change status here to scope in/out

5. **Card - Current Coverage %**
   - Measure: `[Coverage % by FSLI]`
   - Updates in real-time as scoping status changes

6. **Matrix - Coverage by Division x FSLI**
   - Rows: Pack_Number_Company_Table[Division]
   - Columns: FSLi_Key_Table[FSLI]
   - Values: [Coverage % by FSLI]
   - Conditional Formatting: Color scale (red to green)

#### Page 4: Division Analysis

**Visuals:**

1. **Slicer - Division Selector**
   - Field: Pack_Number_Company_Table[Division]
   - Style: Buttons or Tiles

2. **Card - Coverage % for Selected Division**
   - Measure: `[Coverage % by Division]`

3. **Table - Packs in Selected Division**
   - Pack Name, Pack Code, Scoping Status

4. **Stacked Bar Chart - Coverage by FSLI**
   - Axis: FSLi_Key_Table[FSLI]
   - Values: [Coverage % by FSLI]
   - Filter: Selected Division

### Step 6: Enable Edit Mode for Manual Scoping

**Critical Configuration:**

**📚 Detailed Guide:** See [POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md) for complete step-by-step instructions with screenshots and troubleshooting.

**Quick Steps:**

To enable users to manually change scoping status in Power BI:

1. Click on Scoping Control Table visual
2. Format pane → Values → Enable **Edit** option
3. Users can now click in "Scoping Status" column and change values
4. Options:
   - "Scoped In (Manual)"
   - "Scoped Out"
   - "Not Scoped"

**How It Works:**
- User selects pack/FSLI combination
- Changes status to "Scoped In (Manual)"
- All measures update immediately
- Coverage percentages recalculate
- Dashboard reflects new scope

**If Edit Mode Doesn't Work:**

See [POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md) for:
- 3 alternative methods for manual scoping
- Detailed troubleshooting (10+ common issues)
- Screenshots and visual guides
- Workarounds if edit mode unavailable

**Alternative Method (if Edit not available):**

Create a separate "Scoping Decisions" table in Excel:
```
| Pack Code | FSLI | Manual Status |
|-----------|------|---------------|
| BVT-101   | Revenue | Scoped In |
```

Import to Power BI and create relationship with Scoping_Control_Table.

Full details in [POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md) Method 2.

### Step 7: Publish & Share (Optional)

1. File → Publish → Publish to Power BI Service
2. Select workspace
3. Share with audit team
4. Set up refresh schedule (if source Excel updates)

---

## 6. Manual Scoping Workflow

### Workflow Overview

```
1. Review Automatic Scoping Results
   ↓
2. Identify Coverage Gaps by FSLI
   ↓
3. Select FSLI for Manual Review
   ↓
4. Review Pack List for that FSLI
   ↓
5. Select Packs to Scope In
   ↓
6. Update Scoping Status
   ↓
7. Verify Coverage % Updated
   ↓
8. Repeat for Other FSLIs
   ↓
9. Export Final Results
```

### Method 1: Power BI Edit Mode (Recommended)

**Step-by-Step:**

1. **Open Manual Scoping Control Page**
   - Navigate to "Manual Scoping Control" page in Power BI

2. **Review Current Coverage**
   - Check coverage % cards
   - Identify FSLIs with low coverage

3. **Select FSLI**
   - Use FSLI slicer to select target FSLI (e.g., "Inventory")
   - Table filters to show only that FSLI

4. **Review Packs**
   - Review pack list with amounts
   - Sort by amount (largest first)
   - Identify packs for scoping

5. **Update Status**
   - Click in "Scoping Status" column
   - Change from "Not Scoped" to **"Scoped In (Manual)"**
   - Press Enter

6. **Verify Update**
   - Coverage % card updates immediately
   - Visual feedback confirms change

7. **Repeat**
   - Continue for other packs/FSLIs
   - Monitor coverage targets

### Method 2: Excel Update + Refresh

If Power BI edit mode not available:

1. **Export Current Status**
   - In Power BI, export Scoping Control Table to Excel

2. **Update in Excel**
   - Open exported file
   - Change Scoping Status column for desired packs/FSLIs
   - Save changes

3. **Update Source**
   - Copy changes back to "Scoping_Control_Table" sheet in source Excel
   - Save source Excel file

4. **Refresh Power BI**
   - In Power BI: Home → Refresh
   - All visuals update with new scoping decisions

### Method 3: Pack-Level Scoping

**To scope entire pack (all FSLIs):**

1. Select pack from Pack Name slicer
2. Table shows all FSLIs for that pack
3. Select all rows (Ctrl+A in table)
4. Change status to "Scoped In (Manual)"
5. All FSLIs for pack now scoped in

### Method 4: Division-Level Scoping

**To scope all packs in a division:**

1. Select division from Division slicer
2. Table filters to that division
3. Select all visible rows
4. Change status to "Scoped In (Manual)"
5. All packs in division now scoped in

### Scoping Status Values

**Use these exact values:**

- **"Scoped In (Threshold)"** - Automatically scoped by VBA threshold
- **"Scoped In (Manual)"** - Manually scoped in Power BI
- **"Scoped Out"** - Explicitly excluded from scope
- **"Not Scoped"** - Default, not yet reviewed

**Important:** DAX measures recognize both "Scoped In (Threshold)" and "Scoped In (Manual)" as scoped in.

### Coverage Targets

**ISA 600 Guidelines:**

- **High Risk FSLIs:** 80-90% coverage recommended
  - Revenue, Total Assets, Net Profit
- **Medium Risk FSLIs:** 60-80% coverage
  - Most income statement and balance sheet items
- **Low Risk FSLIs:** 40-60% coverage
  - Smaller balances, lower risk accounts

**Setting Targets:**

1. Identify risk level for each FSLI
2. Set coverage target percentage
3. Use manual scoping to achieve target
4. Document rationale in audit file

### Audit Trail

**Document Your Decisions:**

1. **Take Screenshots**
   - Coverage by FSLI before/after
   - Division analysis
   - Final scoping status

2. **Export Final Tables**
   - Export Scoping Control Table to Excel
   - Save as "Final_Scoping_Decisions_[Date].xlsx"
   - Include in audit documentation

3. **Narrative Documentation**
   - Document why certain packs scoped in/out
   - Explain coverage target rationale
   - Link to risk assessment

---

## 7. ISA 600 Compliance

### ISA 600 Revised Requirements

**Key Requirements:**

1. **Component Identification**
   - ✅ Tool identifies all components (packs)
   - ✅ Excludes consolidated entity from scoping
   - ✅ Links components to divisions

2. **Materiality Thresholds**
   - ✅ Threshold-based automatic scoping
   - ✅ Configurable per FSLI
   - ✅ Audit trail of threshold application

3. **Coverage Analysis**
   - ✅ Coverage % by FSLI
   - ✅ Coverage % by Division
   - ✅ Untested % tracking

4. **Component Scoping**
   - ✅ Automatic threshold-based
   - ✅ Manual risk-based
   - ✅ Combination approach supported

5. **Documentation**
   - ✅ Scoping decisions documented
   - ✅ Threshold configuration recorded
   - ✅ Pack-level detail available
   - ✅ Division-level summaries

### Compliance Checklist

**Use this checklist for each audit:**

- [ ] **1. Component Identification**
  - [ ] All entities identified as components
  - [ ] Consolidated entity identified and excluded
  - [ ] Divisions/segments mapped

- [ ] **2. Materiality Assessment**
  - [ ] Group materiality determined
  - [ ] Component materiality calculated
  - [ ] Thresholds set per FSLI

- [ ] **3. Risk Assessment**
  - [ ] Significant risk FSLIs identified
  - [ ] High-risk components identified
  - [ ] Risk-based scoping applied

- [ ] **4. Scoping Decisions**
  - [ ] Threshold scoping applied
  - [ ] Manual scoping for risk factors
  - [ ] Coverage targets achieved

- [ ] **5. Documentation**
  - [ ] Scoping tool output retained
  - [ ] Power BI dashboard screenshots
  - [ ] Scoping rationale documented
  - [ ] Audit file updated

- [ ] **6. Review**
  - [ ] Senior review of scoping
  - [ ] EQCR review completed
  - [ ] Scoping approved

### ISA 600 Reporting

**Include in Audit File:**

1. **Scoping Tool Output.xlsx**
   - Complete workbook with all tables
   - Threshold configuration
   - Scoping summary

2. **Power BI Dashboard PDF**
   - Export each page as PDF
   - Include coverage metrics
   - Include division analysis

3. **Scoping Memo**
   - Narrative explanation of approach
   - Rationale for thresholds
   - Risk-based scoping decisions
   - Coverage achieved vs. target

4. **Component Work Papers**
   - List of scoped-in components
   - Planned procedures per component
   - Link to risk assessment

### Consolidated Entity Exclusion

**Why It Matters:**

ISA 600 requires identifying **components** for scoping, not the consolidated group total.

**How Tool Handles It:**

1. **VBA Selection:** User selects consolidated entity at start
2. **Flag in Data:** "Is Consolidated" column = "Yes" for that pack
3. **DAX Filters:** All measures exclude where Is Consolidated = "Yes"
4. **Coverage Calc:** Percentages calculated only on component totals

**Verification:**

- Check Pack Number Company Table
- Verify one pack has "Is Consolidated = Yes"
- Verify that pack excluded from all measures
- Test: Coverage should be < 100% even if all other packs scoped

---

## 8. Troubleshooting

### Common Issues & Solutions

#### Issue 1: "Could not find workbook"

**Symptom:** Error message when entering workbook name

**Causes:**
- Workbook not open
- Name mismatch
- Wrong extension

**Solutions:**
1. Verify consolidation workbook is open
2. Copy exact filename from title bar (including .xlsx or .xlsm)
3. Check for typos or extra spaces
4. Ensure workbook is not minimized

#### Issue 2: "Required tabs are missing"

**Symptom:** Validation error after tab categorization

**Cause:** No tab categorized as "Input Continuing"

**Solution:**
1. Restart categorization
2. Ensure exactly ONE tab categorized as category 3 (Input Continuing)
3. This tab is mandatory

#### Issue 3: VBA Runs But No Data in Tables

**Symptom:** Output workbook created but tables are empty or have headers only

**Causes:**
- Incorrect row structure
- Wrong column selected
- Data format issue

**Solutions:**
1. **Verify Structure:**
   - Row 6: Currency type
   - Row 7: Pack names
   - Row 8: Pack codes
   - Row 9+: Data

2. **Check Column Selection:**
   - Rerun tool
   - Select "Consolidation Currency" (YES)

3. **Verify Data:**
   - Open consolidation workbook
   - Check Input Continuing tab
   - Ensure data exists in rows 9+

#### Issue 4: FSLI Headers Appearing in Output

**Symptom:** "INCOME STATEMENT" or "BALANCE SHEET" showing as FSLIs

**Cause:** IsStatementHeader function not working

**Solution:**
1. This should be fixed in current version
2. Verify ModDataProcessing.bas has IsStatementHeader function
3. Check FSLi Key Table - headers should be excluded
4. If still occurring, report as bug

#### Issue 5: Notes Section Not Excluded

**Symptom:** Items below "Notes" appearing in FSLI list

**Cause:** Notes detection logic issue

**Solution:**
1. Check if "Notes" row exists in Column B
2. Verify exact spelling (should be "NOTES" or "Notes")
3. Ensure row is not hidden
4. Check AnalyzeFSLiStructure function for Notes detection

#### Issue 6: Power BI Tables Not Importing

**Symptom:** Can't find tables when importing to Power BI

**Causes:**
- Not saved as Excel tables (ListObjects)
- File not saved
- File path issue

**Solutions:**
1. Open output Excel file
2. Verify each table has filter dropdowns (= Excel Table)
3. Save file and try import again
4. Use "Get Data → Excel" (not "Get Data → Folder")

#### Issue 7: Power BI Relationships Not Creating

**Symptom:** Can't create relationships or relationships show error

**Causes:**
- Data type mismatch
- Duplicate values in "one" side
- Null values

**Solutions:**
1. **Check Data Types:**
   - Pack Code should be Text in both tables
   - FSLI should be Text in both tables

2. **Check for Duplicates:**
   - Pack Code should be unique in Pack_Number_Company_Table
   - FSLI should be unique in FSLi_Key_Table

3. **Remove Nulls:**
   - Filter out blank Pack Codes
   - Filter out blank FSLIs

#### Issue 8: DAX Measures Return Incorrect Values

**Symptom:** Coverage % shows wrong percentage or errors

**Causes:**
- Relationship issues
- Filter context problems
- Data type issues

**Solutions:**
1. **Verify Relationships:** Check all relationships exist and are active
2. **Test Measures:** Create simple measures first (e.g., COUNT)
3. **Check Filters:** Use REMOVEFILTERS() if measures too restrictive
4. **Data Types:** Ensure Amount columns are Decimal Number

#### Issue 9: Manual Scoping Not Updating

**Symptom:** Change scoping status but coverage doesn't update

**Causes:**
- Edit mode not enabled
- Measures not referencing Scoping Status
- Refresh needed

**Solutions:**
1. Verify edit mode enabled on table visual
2. Check DAX measures include both "Scoped In (Threshold)" and "Scoped In (Manual)"
3. Refresh visual (right-click → Refresh)
4. Check spelling of status values (case-sensitive)

#### Issue 10: Excel Crashes or Freezes

**Symptom:** Excel stops responding during VBA processing

**Causes:**
- Large workbook
- Memory limitation
- Complex formulas

**Solutions:**
1. **Free Memory:**
   - Close other applications
   - Restart Excel before running

2. **Disable Features:**
   - Turn off automatic calculation (Formulas → Calculation Options → Manual)
   - Close unneeded workbooks

3. **Process Smaller Sections:**
   - Process one division at a time
   - Combine results manually

4. **Upgrade Hardware:**
   - Minimum 8GB RAM recommended
   - 16GB for large workbooks

### Error Messages Reference

| Error Message | Cause | Solution |
|---------------|-------|----------|
| "Ambiguous name detected" | Function duplicated | Reimport latest VBA modules |
| "Type mismatch" | Data type issue | Check variable types in VBA |
| "Subscript out of range" | Collection/array access error | Check loop bounds |
| "Object required" | Nothing assigned to object | Check Set statements |
| "Method not found" | Wrong object type | Verify object references |

### Performance Optimization

**For Large Workbooks (1000+ FSLIs, 50+ packs):**

1. **Before Running VBA:**
   ```
   - Close all other Excel workbooks
   - Close unnecessary applications
   - Ensure 4GB+ RAM available
   - Disable automatic calculation
   ```

2. **During Processing:**
   ```
   - Do not interact with Excel
   - Let tool run uninterrupted
   - Monitor Task Manager for memory usage
   ```

3. **In Power BI:**
   ```
   - Import only necessary tables initially
   - Add complexity incrementally
   - Use aggregations for large datasets
   - Consider DirectQuery for very large data
   ```

4. **Code Optimization (Advanced):**
   ```vba
   ' Already implemented in tool:
   Application.ScreenUpdating = False
   Application.Calculation = xlCalculationManual
   Application.EnableEvents = False
   ' ... processing ...
   Application.ScreenUpdating = True
   Application.Calculation = xlCalculationAutomatic
   Application.EnableEvents = True
   ```

### Getting Help

**If Issue Persists:**

1. **Review This Guide:** Check relevant section thoroughly
2. **Check VBA Code Comments:** Code is well-documented
3. **Test with Sample Data:** Create small test workbook
4. **Check Excel Version:** Ensure 2016 or later
5. **Verify VBA References:** Tools → References in VBA Editor

**Information to Gather for Support:**

- Excel version
- Output file generated (yes/no)
- Error message (exact text)
- Step where error occurred
- Consolidation workbook size (tabs, rows, columns)

---

## 9. Technical Reference

### VBA Module Documentation

#### ModMain.bas

**Purpose:** Main entry point and orchestration

**Key Functions:**
- `StartScopingTool()` - Main entry point
- `SelectConsolidatedEntity()` - Consolidated entity selection
- `CreateScopingSummarySheet()` - Scoping summary
- `CreateDivisionScopingReports()` - Division reports
- `SaveOutputWorkbook()` - Save with standard name

**Global Variables:**
- `g_SourceWorkbook` - Reference to consolidation workbook
- `g_OutputWorkbook` - Reference to output workbook
- `g_TabCategories` - Dictionary of tab categorizations
- `g_ConsolidatedPackCode` - Consolidated entity pack code
- `g_ConsolidatedPackName` - Consolidated entity name

#### ModConfig.bas

**Purpose:** Configuration and utilities

**Key Constants:**
```vba
CAT_SEGMENT = "TGK Segment Tabs"
CAT_DISCONTINUED = "Discontinued Ops Tab"
CAT_INPUT_CONTINUING = "TGK Input Continuing Operations Tab"
CAT_JOURNALS_CONTINUING = "TGK Journals Continuing Tab"
CAT_CONSOLE_CONTINUING = "TGK Consol Continuing Tab"
```

**Utility Functions:**
- `GetWorkbookByName()` - Find workbook by name
- `GetToolVersion()` - Return version string
- Validation helpers

#### ModTabCategorization.bas

**Purpose:** Tab discovery and categorization

**Key Functions:**
- `CategorizeTabs()` - Main categorization workflow
- `ShowCategorizationDialog()` - User interface
- `ValidateCategories()` - Validation logic
- `GetDivisionName()` - Division name prompt

#### ModDataProcessing.bas

**Purpose:** FSLI extraction and data processing

**Key Functions:**
- `ProcessConsolidationData()` - Main processor
- `AnalyzeFSLiStructure()` - FSLI extraction with Notes cutoff
- `IsStatementHeader()` - Header filtering
- `DetectColumns()` - Column analysis
- `CreateGenericTable()` - Universal table creator

**FSLI Extraction Logic:**
```vba
' Stops at Notes section:
If UCase(fsliName) = "NOTES" Then
    Exit For
End If

' Filters statement headers:
If IsStatementHeader(fsliName) Then
    GoTo NextRow
End If

' Includes items with data
' Detects hierarchy via indentation
' Marks totals/subtotals
```

#### ModTableGeneration.bas

**Purpose:** Table creation and formatting

**Key Functions:**
- `CreateFSLiKeyTable()` - FSLi reference table
- `CreatePackNumberCompanyTable()` - Pack reference
- `CreatePercentageTables()` - Percentage calculations
- `FormatAsTable()` - Excel Table formatting

**Table Structure:**
- All tables as Excel ListObjects
- Standard naming convention
- TableStyleMedium2 applied
- Auto-filter enabled

#### ModThresholdScoping.bas

**Purpose:** Automatic threshold-based scoping

**Key Functions:**
- `ConfigureAndApplyThresholds()` - Threshold wizard
- `ApplyThresholdsToData()` - Threshold analysis
- `CreateThresholdConfigSheet()` - Configuration doc

**Logic:**
```vba
' For each selected FSLI:
' - Compare pack values to threshold
' - If ABS(value) > threshold:
'   - Mark entire pack as "Scoped In (Threshold)"
'   - Track which FSLI triggered scoping
' - Exclude consolidated entity
```

#### ModInteractiveDashboard.bas

**Purpose:** Excel-based dashboard creation

**Key Functions:**
- `CreateInteractiveDashboard()` - Dashboard layout
- `CreateScopingPivotTable()` - Pivot analysis
- `CreateScopingCalculator()` - Coverage calculator

#### ModPowerBIIntegration.bas

**Purpose:** Power BI integration and metadata

**Key Functions:**
- `CreateAllPowerBIAssets()` - Create all assets
- `CreateScopingControlTable()` - Manual scoping table
- `CreateDAXMeasuresGuide()` - DAX documentation
- `CreatePowerBIMetadata()` - Metadata sheet

**Scoping Control Table Structure:**
```
| Pack Name | Pack Code | Division | FSLI | Amount | Scoping Status |
```

### Data Dictionary

#### Full Input Table

| Column | Data Type | Description |
|--------|-----------|-------------|
| Pack Name | Text | Legal entity name |
| Pack Code | Text | Unique identifier (e.g., BVT-101) |
| [FSLI columns] | Number | Financial data per FSLI |

**Notes:**
- Unpivot for Power BI
- Consolidation currency only
- One row per pack

#### Full Input Percentage

Same structure as Full Input Table but values as percentages.

**Calculation:**
```
Percentage = ABS(Cell Value) / SUM(ABS(Column Values)) * 100
```

#### FSLi Key Table

| Column | Data Type | Description |
|--------|-----------|-------------|
| FSLI | Text | Financial statement line item name |
| Statement Type | Text | "Income Statement" or "Balance Sheet" |
| Is Total | Boolean | TRUE if line is a total |
| Level | Number | Indentation level (hierarchy) |

#### Pack Number Company Table

| Column | Data Type | Description |
|--------|-----------|-------------|
| Pack Name | Text | Legal entity name |
| Pack Code | Text | Unique identifier |
| Division | Text | Business division/segment |
| Is Consolidated | Text | "Yes" if consolidated entity |

**Key:** Pack Code (unique)

#### Scoping Control Table

| Column | Data Type | Description |
|--------|-----------|-------------|
| Pack Name | Text | Legal entity name |
| Pack Code | Text | Unique identifier |
| Division | Text | Business division |
| FSLI | Text | Financial statement line item |
| Amount | Number | Value in consolidation currency |
| Scoping Status | Text | "Scoped In (Threshold)", "Scoped In (Manual)", "Not Scoped", "Scoped Out" |

**Purpose:** Enable manual scoping in Power BI

**Editability:** Scoping Status column editable in Power BI

### File Naming Conventions

**VBA Tool:**
- Recommended: `Bidvest_Scoping_Tool.xlsm`

**Output:**
- Always: `Bidvest Scoping Tool Output.xlsx`
- Location: Same folder as source workbook

**Power BI:**
- Recommended: `Bidvest_Scoping_Analysis.pbix`

### Version History

#### v4.0 (Current - Complete Overhaul)
- Consolidated documentation (24 files → 1 guide)
- Verified FSLI extraction logic
- Enhanced Power BI manual scoping documentation
- ISA 600 compliance checklist
- Comprehensive troubleshooting

#### v3.1 (November 2024)
- Consolidated entity selection
- Dynamic Power BI scoping
- Pack Name + Pack Code columns
- Division logic updates
- Scoping Control Table

#### v3.0 (November 2024)
- Division-based reporting
- Text-based FSLI selection
- Autonomous workflow
- Professional Excel output

#### v2.0 (November 2024)
- Threshold-based scoping
- Interactive Excel dashboard
- Scoping summary
- FSLI header filtering fix

#### v1.1 (November 2024)
- Power BI integration
- ModConfig added
- Bug fixes

#### v1.0 (Initial)
- Core functionality
- Tab categorization
- Table generation

### System Requirements

**Minimum:**
- Windows 10
- Excel 2016
- Power BI Desktop
- 4GB RAM
- 500MB disk space

**Recommended:**
- Windows 11
- Excel 2021 / Microsoft 365
- Power BI Desktop (latest)
- 8GB+ RAM
- 1GB disk space
- 1920x1080 display

**PwC Environment:**
- ✅ Works with standard Office installation
- ✅ No SQL database required
- ✅ No internet required for VBA
- ✅ Power BI Desktop approved

### Known Limitations

1. **Language:** English only (consolidation workbooks must be in English)
2. **Format:** Assumes standard TGK format (rows 6-8 structure)
3. **Size:** Practical limit ~1000 FSLIs, ~500 packs
4. **Power BI Edit:** Edit mode may require Power BI Pro license
5. **Formulas:** Complex Excel formulas not analyzed

### Best Practices

1. **Always run on a copy** of consolidation workbook
2. **Test with sample data** before production use
3. **Save frequently** during Power BI setup
4. **Document decisions** in audit file
5. **Review output** before finalizing scoping
6. **Keep tool updated** with latest VBA modules
7. **Back up work** regularly

### Support & Maintenance

**Code Maintenance:**
- Modular design allows independent updates
- Well-commented code
- Consistent naming conventions

**Customization:**
- Add categories in ModConfig.bas
- Modify table structures in ModTableGeneration.bas
- Add DAX measures as needed

**Testing:**
- Test each major change with sample data
- Verify output before production use
- Validate calculations independently

---

## Appendix A: Quick Reference

### VBA Tool - Quick Steps

1. Open consolidation workbook + tool workbook
2. Click "Start Bidvest Scoping Tool"
3. Enter workbook name (exact, with extension)
4. Categorize tabs (3 = Input Continuing, required)
5. Select consolidated entity (usually first pack)
6. Choose Consolidation Currency (YES)
7. Configure thresholds (optional)
8. Wait for processing
9. Review output workbook

### Power BI - Quick Steps

1. Get Data → Excel → Select output file
2. Select all tables → Transform Data
3. Unpivot data tables (keep reference tables)
4. Close & Apply
5. Create relationships (Pack Code, FSLI)
6. Create DAX measures (copy from guide)
7. Build dashboard pages
8. Enable edit mode on Scoping Control Table
9. Manual scope as needed
10. Export results

### DAX Measures - Quick Reference

```dax
// Total Packs
DISTINCTCOUNT(Pack_Number_Company_Table[Pack Code])

// Scoped In
CALCULATE(
    DISTINCTCOUNT(...),
    Scoping_Control_Table[Scoping Status] IN {
        "Scoped In (Threshold)", 
        "Scoped In (Manual)"
    }
)

// Coverage %
DIVIDE([Scoped Amount], [Total Amount], 0)

// Always filter out consolidated:
Pack_Number_Company_Table[Is Consolidated] <> "Yes"
```

### ISA 600 - Quick Checklist

- [ ] All components identified
- [ ] Consolidated entity excluded
- [ ] Thresholds applied
- [ ] Manual scoping completed
- [ ] Coverage targets achieved
- [ ] Documentation complete
- [ ] Senior review done

### Troubleshooting - Quick Fixes

| Problem | Quick Fix |
|---------|-----------|
| Workbook not found | Check exact name, include extension |
| No data in tables | Verify row 6-8 structure |
| Headers in FSLIs | Check IsStatementHeader function |
| Power BI won't import | Verify Excel Tables exist |
| Relationships fail | Check Pack Code data type (Text) |
| Measures return errors | Verify relationships active |

---

## Appendix B: Glossary

**Component** - Individual entity within group structure (ISA 600 term)  
**Consolidated Entity** - Parent company representing group totals  
**Coverage %** - Percentage of FSLI amount scoped for testing  
**Division** - Business segment or geographical region  
**FSLI** - Financial Statement Line Item  
**ISA 600** - International Standard on Auditing for group audits  
**Pack** - Entity or legal entity within consolidation  
**Pack Code** - Unique identifier for entity (e.g., BVT-101)  
**Scoping** - Process of selecting components for testing  
**Threshold** - Materiality level for automatic scoping  
**Untested %** - Percentage of FSLI amount not scoped (100% - Coverage%)

---

## Appendix C: Contact & Support

**Tool Maintained By:** PwC Audit Technology Team  
**Version:** 4.0  
**Last Updated:** November 2024

**For Technical Issues:**
1. Review Troubleshooting section (Section 8)
2. Check VBA code comments
3. Test with sample data
4. Contact audit technology support

**For ISA 600 Questions:**
1. Consult ISA 600 Revised guidance
2. Review Compliance section (Section 7)
3. Discuss with engagement quality control reviewer

---

**End of Comprehensive Guide**

This guide consolidates all previous documentation into a single, comprehensive resource for the Bidvest Scoping Tool. For the latest updates and version information, refer to the version history section.

**Document Control:**
- **Version:** 4.0
- **Date:** November 2024
- **Status:** Complete Overhaul
- **Next Review:** As needed for ISA updates or tool enhancements
