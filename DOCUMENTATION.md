# TGK Consolidation Scoping Tool - Complete Documentation

## Overview

The TGK Consolidation Scoping Tool is a comprehensive VBA-based solution for Microsoft Excel that automates the analysis of TGK consolidation workbooks, creates structured tables for Power BI integration, and provides automated scoping functionality for audit and review purposes.

## Table of Contents

1. [Installation](#installation)
2. [Quick Start Guide](#quick-start-guide)
3. [Module Overview](#module-overview)
4. [Detailed Functionality](#detailed-functionality)
5. [Power BI Integration](#power-bi-integration)
6. [Troubleshooting](#troubleshooting)
7. [Technical Specifications](#technical-specifications)

---

## Installation

### Prerequisites

- Microsoft Excel 2016 or later (Windows)
- Macro-enabled workbook support (.xlsm)
- VBA enabled in Excel (Trust Center settings)

### Installation Steps

1. **Create a new Excel Macro-Enabled Workbook:**
   - Open Excel
   - Create a new workbook
   - Save as `TGK_Scoping_Tool.xlsm`

2. **Import VBA Modules:**
   - Press `Alt + F11` to open the VBA Editor
   - For each `.bas` file in the `VBA_Modules` folder:
     - Go to `File > Import File`
     - Select the module file
     - Click `Open`

3. **Import the modules in this order:**
   - `ModMain.bas`
   - `ModTabCategorization.bas`
   - `ModDataProcessing.bas`
   - `ModTableGeneration.bas`

4. **Create a Button:**
   - Return to Excel (Alt + F11 or close VBA Editor)
   - Go to `Developer` tab (if not visible, enable it in Excel Options)
   - Click `Insert > Button (Form Control)`
   - Draw a button on the worksheet
   - In the "Assign Macro" dialog, select `StartScopingTool`
   - Click `OK`
   - Right-click the button and select `Edit Text`
   - Change the text to "Start TGK Scoping Tool"

5. **Save the workbook**

---

## Quick Start Guide

### Step 1: Prepare Your Environment

1. Open the TGK consolidation workbook you want to analyze
2. Open the `TGK_Scoping_Tool.xlsm` workbook
3. Ensure both workbooks are open simultaneously

### Step 2: Run the Tool

1. Click the "Start TGK Scoping Tool" button
2. Read the welcome message and click `OK`
3. Enter the exact name of your consolidation workbook (e.g., `Consolidation_2024.xlsx`)
4. Click `OK`

### Step 3: Categorize Tabs

The tool will present pop-up dialogs to categorize each tab from your consolidation workbook.

**Available Categories:**

1. **TGK Segment Tabs** (multiple allowed): Individual business segments
2. **Discontinued Ops Tab** (single only): Discontinued operations
3. **TGK Input Continuing Operations Tab** (single only): Primary input data (REQUIRED)
4. **TGK Journals Continuing Tab** (single only): Consolidation journal entries
5. **TGK Consol Continuing Tab** (single only): Consolidated data
6. **TGK BS Tab** (single only): Balance Sheet
7. **TGK IS Tab** (single only): Income Statement
8. **Paul workings** (multiple allowed): Working papers and calculations
9. **Trial Balance** (single only): Trial balance data
10. **Uncategorized**: Tabs to be ignored

**Instructions:**

1. Read the categorization instructions in the initial pop-up
2. For each tab, a pop-up dialog will appear showing the tab name and the full list of categories
3. Enter the number (1-9) corresponding to the desired category
4. For segment tabs (category 1), you'll be prompted to enter a division name (e.g., "UK", "US", "Europe")
   - If you leave the division name blank, it will automatically use "Division_1", "Division_2", etc.
5. If you enter an invalid number, you'll be prompted to try again
6. If you click Cancel on any dialog, you'll be asked if you want to cancel the entire process
7. After all tabs are categorized, the tool validates your selections:
   - Ensures single-only categories have only one tab assigned
   - Checks that required categories are present
8. If validation fails, you can choose to start over or cancel
9. Review any uncategorized tabs and confirm whether to proceed or restart categorization

**Important Notes:**
- All category selection is done through pop-up dialogs (InputBox and MsgBox)
- No worksheet tabs are created for categorization
- You can restart the categorization process at any time if you make a mistake
- The default category suggestion is "3" (Input Continuing Operations) as it is required

### Step 4: Select Column Type

When prompted, choose which columns to process:
- **Consolidation/Consolidation Currency** (RECOMMENDED): Uses consolidated currency values
- **Original/Entity Currency**: Uses original entity currency values

Click `YES` for Consolidation/Consolidation Currency

### Step 5: Wait for Processing

The tool will:
- Unmerge cells
- Analyze FSLi structure
- Extract entity information
- Create all required tables
- Generate percentage calculations

This may take several minutes depending on workbook size.

### Step 6: Review Output

A new workbook will be created with the following tables:
- Full Input Table
- Full Input Percentage
- Journals Table
- Journals Percentage
- Full Consol Table
- Full Consol Percentage
- Discontinued Table
- Discontinued Percentage
- FSLi Key Table
- Pack Number Company Table

---

## Module Overview

### ModMain.bas

**Purpose:** Main entry point and orchestration

**Key Functions:**
- `StartScopingTool()`: Primary entry point called by button
- `GetWorkbookName()`: Prompts user for workbook name
- `SetSourceWorkbook()`: Validates and sets workbook reference
- `DiscoverTabs()`: Lists all worksheets
- `CreateOutputWorkbook()`: Initializes output workbook

### ModTabCategorization.bas

**Purpose:** Tab categorization and validation

**Key Functions:**
- `CategorizeTabs()`: Main categorization orchestrator
- `ShowCategorizationDialog()`: User interface for categorization
- `ValidateSingleTabCategories()`: Ensures single-tab categories have only one tab
- `ValidateCategories()`: Verifies required categories are assigned
- `GetTabsForCategory()`: Retrieves tabs by category
- `GetDivisionName()`: Gets division name for segment tabs

**Category Constants:**
```vba
CAT_SEGMENT = "TGK Segment Tabs"
CAT_DISCONTINUED = "Discontinued Ops Tab"
CAT_INPUT_CONTINUING = "TGK Input Continuing Operations Tab"
CAT_JOURNALS_CONTINUING = "TGK Journals Continuing Tab"
CAT_CONSOLE_CONTINUING = "TGK Consol Continuing Tab"
CAT_BS = "TGK BS Tab"
CAT_IS = "TGK IS Tab"
CAT_PULL_WORKINGS = "Paul workings"
CAT_TRIAL_BALANCE = "Trial Balance"
CAT_UNCATEGORIZED = "Uncategorized"
```

### ModDataProcessing.bas

**Purpose:** Data extraction and analysis

**Key Functions:**
- `ProcessConsolidationData()`: Main processing orchestrator
- `ProcessInputTab()`: Processes Input Continuing tab
- `DetectColumns()`: Analyzes row 6 for column types
- `PromptColumnSelection()`: User selects column type
- `AnalyzeFSLiStructure()`: Identifies FSLi hierarchy
- `CreateFullInputTable()`: Generates primary data table
- `IsRowEmpty()`: Utility to check empty rows
- `DetectIndentationLevel()`: Determines FSLi hierarchy level

**Data Structures:**
```vba
Type ColumnInfo
    ColumnIndex As Long
    ColumnType As String
    PackName As String
    PackCode As String
End Type

Type FSLiInfo
    FSLiName As String
    RowIndex As Long
    IsTotal As Boolean
    IsSubtotal As Boolean
    SubtotalOf As String
    Level As Long
    StatementType As String
End Type
```

### ModTableGeneration.bas

**Purpose:** Generate supporting tables

**Key Functions:**
- `CreateFSLiKeyTable()`: Creates FSLi master table
- `CollectAllFSLiNames()`: Gathers unique FSLi entries
- `CreatePackNumberCompanyTable()`: Creates entity reference table
- `PromptForDivisionName()`: Prompts for segment divisions
- `CreatePercentageTables()`: Generates percentage tables
- `CreatePercentageTable()`: Creates individual percentage table
- `FormatAsTable()`: Applies consistent formatting

---

## Detailed Functionality

### Tab Discovery and Categorization

The tool analyzes the consolidation workbook structure by:

1. **Discovery Phase:**
   - Enumerates all worksheets
   - Creates a list of tab names

2. **Categorization Phase:**
   - Presents interactive interface
   - Validates user selections
   - Ensures mandatory categories are assigned

3. **Validation Phase:**
   - Checks single-tab category constraints
   - Verifies required tabs exist
   - Handles uncategorized tabs

### Data Structure Analysis

#### Row Structure (Standard TGK Layout):

- **Row 6:** Column type identifiers
  - "Original and Entity Currency"
  - "Consolidation and Consolidation Currency"

- **Row 7:** Entity/Pack names
  - Legal entity names
  - Used as primary identifiers

- **Row 8:** Entity/Pack codes
  - Unique entity codes (e.g., "BVT-001")
  - Used for cross-referencing

- **Row 9+:** FSLi data
  - Financial Statement Line Items
  - Hierarchical structure with totals and subtotals

#### Column B Analysis:

The tool analyzes Column B to identify:
- **Statement Headers:** "Income Statement", "Balance Sheet"
- **FSLi Names:** Account names and line items
- **Totals:** Lines containing "total" in the name
- **Subtotals:** Lines containing "subtotal" or detected by indentation
- **Notes Section:** Special section marked by "Notes" header

### Table Generation

#### Full Input Table Structure:

| Pack | Revenue | Cost of Sales | Gross Profit | ... | Total Assets |
|------|---------|---------------|--------------|-----|--------------|
| Entity 1 | 1,000,000 | 600,000 | 400,000 | ... | 5,000,000 |
| Entity 2 | 2,000,000 | 1,200,000 | 800,000 | ... | 10,000,000 |

**Features:**
- First column: Pack/Entity names
- Subsequent columns: Each FSLi
- Metadata tags: "(Total)" or "(Subtotal)" indicators
- Dynamic sizing based on data

#### FSLi Key Table Structure:

| FSLi | FSLi Input | FSLi Input % | FSLi Journal | FSLi Journal % | ... |
|------|------------|--------------|--------------|----------------|-----|
| Revenue | 3,000,000 | 15% | 0 | 0% | ... |
| Cost of Sales | 1,800,000 | 12% | 0 | 0% | ... |

**Features:**
- Consolidates all FSLi entries
- Links to main tables via VLOOKUP
- Includes percentage columns
- Covers Input, Journal, Console, and Discontinued data

#### Pack Number Company Table Structure:

| Pack Name | Pack Code | Division |
|-----------|-----------|----------|
| UK Entity Ltd | BVT-001 | UK |
| US Entity Inc | BVT-002 | US |
| Discontinued Co | BVT-999 | Discontinued |

**Features:**
- Unique entity list
- Division/segment mapping
- Code cross-reference

#### Percentage Tables:

For each main table (Input, Journal, Console, Discontinued), a corresponding percentage table is created:

**Calculation Method:**
```
Percentage = (Absolute Value of Cell / Sum of Absolute Values in Column) × 100
```

**Format:** 
- Same structure as source table
- Values displayed as percentages
- Used for coverage analysis

---

## Power BI Integration

### Data Model Setup

#### Step 1: Import Tables

1. Open Power BI Desktop
2. Get Data > Excel
3. Select the output workbook
4. Import all tables:
   - Full Input Table
   - Full Input Percentage
   - Journals Table
   - Journals Percentage
   - Full Consol Table
   - Full Consol Percentage
   - Discontinued Table
   - Discontinued Percentage
   - FSLi Key Table
   - Pack Number Company Table

#### Step 2: Create Relationships

**Primary Relationships:**
```
Pack Number Company Table[Pack Name] → Full Input Table[Pack]
Pack Number Company Table[Pack Name] → Journals Table[Pack]
Pack Number Company Table[Pack Name] → Full Consol Table[Pack]
Pack Number Company Table[Pack Name] → Discontinued Table[Pack]
```

**FSLi Relationships:**
```
FSLi Key Table[FSLi] → (Unpivoted FSLi columns from tables)
```

#### Step 3: Data Transformation

**Unpivot FSLi Columns:**

For each data table (Input, Journal, Console, Discontinued):

1. Select the table in Power Query Editor
2. Select "Pack" column
3. Right-click > Unpivot Other Columns
4. Rename columns:
   - Attribute → "FSLi"
   - Value → "Amount"

Example transformation:

**Before:**
| Pack | Revenue | Cost of Sales |
|------|---------|---------------|
| Entity 1 | 1000 | 600 |

**After:**
| Pack | FSLi | Amount |
|------|------|--------|
| Entity 1 | Revenue | 1000 |
| Entity 1 | Cost of Sales | 600 |

### Visualization Setup

#### Scoping Dashboard

**Visual 1: FSLi Selector**
- Type: Slicer
- Field: FSLi Key Table[FSLi]
- Settings: Multi-select enabled

**Visual 2: Pack Selector**
- Type: Slicer
- Field: Pack Number Company Table[Pack Name]
- Settings: Multi-select enabled

**Visual 3: Division Filter**
- Type: Slicer
- Field: Pack Number Company Table[Division]

**Visual 4: Coverage Summary**
- Type: Card
- Measure: Total Coverage %
```DAX
Total Coverage % = 
CALCULATE(
    SUM('Full Input Percentage'[Amount]),
    FILTER(
        'Full Input Table',
        'Full Input Table'[Pack] IN VALUES(SelectedPacks)
    )
)
```

**Visual 5: Detailed Analysis**
- Type: Table
- Fields:
  - Pack Number Company Table[Pack Name]
  - Pack Number Company Table[Division]
  - Full Input Table[FSLi]
  - Full Input Table[Amount]
  - Full Input Percentage[Amount]

### Scoping Logic Implementation

#### DAX Measures

**1. Define Threshold Parameter:**
```DAX
Threshold = PARAMETER("Enter threshold value", INTEGER, 300000000)
```

**2. Calculate Scoped In Packs:**
```DAX
Scoped In Packs = 
VAR SelectedFSLi = SELECTEDVALUE('FSLi Key Table'[FSLi])
VAR ThresholdValue = [Threshold]
RETURN
CALCULATE(
    DISTINCTCOUNT('Full Input Table'[Pack]),
    FILTER(
        'Full Input Table',
        'Full Input Table'[FSLi] = SelectedFSLi &&
        ABS('Full Input Table'[Amount]) > ThresholdValue
    )
)
```

**3. Coverage Percentage:**
```DAX
Coverage % = 
VAR ScopedTotal = 
    CALCULATE(
        SUM('Full Input Table'[Amount]),
        FILTER(
            'Full Input Table',
            'Full Input Table'[Pack] IN [Scoped In Packs]
        )
    )
VAR GrandTotal = SUM('Full Input Table'[Amount])
RETURN
DIVIDE(ScopedTotal, GrandTotal, 0)
```

**4. Untested Percentage:**
```DAX
Untested % = 1 - [Coverage %]
```

### Interactive Scoping Workflow

#### Threshold-Based Scoping:

1. User selects FSLi from slicer (e.g., "Net Revenue")
2. User sets threshold parameter (e.g., $300M)
3. Power BI identifies packs where Net Revenue > $300M
4. All FSLis for those packs are marked as "Scoped In"
5. Coverage % updates automatically

#### Manual Pack/FSLi Selection:

1. User selects specific pack from slicer
2. User selects specific FSLi from slicer
3. Selection added to "Scoped In" list
4. Coverage % updates to include new selection

#### Complete Pack Selection:

1. User selects pack from slicer
2. User clicks "Select All FSLis for Pack" button
3. All FSLis for that pack are marked as "Scoped In"
4. Coverage % updates accordingly

### Power BI Reports

#### Report 1: Coverage Analysis
- Coverage % card
- Untested % card
- Scoped In Packs list
- FSLi breakdown by pack

#### Report 2: Threshold Scoping
- FSLi selector
- Threshold parameter
- Packs meeting threshold
- Coverage impact

#### Report 3: Manual Selection
- Pack selector
- FSLi selector
- Current coverage
- Remaining untested items

#### Report 4: Division Analysis
- Division filter
- Pack distribution
- FSLi coverage by division
- Threshold analysis by division

---

## Troubleshooting

### Common Issues

#### Issue 1: "Could not find workbook"

**Cause:** Workbook name doesn't match or workbook is not open

**Solution:**
- Ensure consolidation workbook is open
- Copy exact workbook name including extension
- If name has special characters, verify exact spelling

#### Issue 2: "Required tabs are missing"

**Cause:** Input Continuing tab not categorized

**Solution:**
- Ensure at least one tab is categorized as "TGK Input Continuing Operations Tab"
- This is the minimum required category

#### Issue 3: "Category can only have ONE tab"

**Cause:** Multiple tabs assigned to single-tab category

**Solution:**
- Review categorization
- Ensure only one tab is assigned to:
  - Discontinued Ops Tab
  - TGK Input Continuing Operations Tab
  - TGK Journals Continuing Tab
  - TGK Consol Continuing Tab
  - TGK BS Tab
  - TGK IS Tab
  - Trial Balance

#### Issue 4: Tool runs but no data in tables

**Cause:** Column type mismatch or data structure issue

**Solution:**
- Verify row 6 contains column type identifiers
- Check that row 7 has entity names
- Ensure row 8 has entity codes
- Confirm data starts at row 9

#### Issue 5: Excel freezes or runs slowly

**Cause:** Large workbook or complex formulas

**Solution:**
- Close unnecessary applications
- Ensure adequate memory
- Process smaller workbooks first to test
- Disable automatic calculation before running

#### Issue 6: VBA errors on import

**Cause:** Module dependencies or naming conflicts

**Solution:**
- Import modules in correct order (Main, Categorization, Processing, Generation)
- Check for naming conflicts with existing VBA code
- Verify VBA references are enabled (Tools > References)

### Error Messages

#### "Error in tab categorization"

**Check:**
- Temporary worksheet creation permissions
- Data validation setup
- User canceled operation

#### "Error processing Input tab"

**Check:**
- Tab structure matches expected format
- Rows 6-8 contain header information
- Column B has FSLi names

#### "Error creating Full Input Table"

**Check:**
- Output workbook is accessible
- Sheet name doesn't already exist
- Memory available for large tables

### Performance Optimization

**For Large Workbooks:**

1. **Disable screen updating:**
   ```vba
   Application.ScreenUpdating = False
   ' ... processing ...
   Application.ScreenUpdating = True
   ```

2. **Disable automatic calculation:**
   ```vba
   Application.Calculation = xlCalculationManual
   ' ... processing ...
   Application.Calculation = xlCalculationAutomatic
   ```

3. **Process in batches:**
   - Process one segment at a time
   - Create tables separately
   - Combine results manually

4. **Optimize formulas:**
   - Use values instead of formulas where possible
   - Limit VLOOKUP ranges
   - Consider INDEX/MATCH alternatives

---

## Technical Specifications

### System Requirements

**Minimum:**
- Windows 10
- Excel 2016
- 4GB RAM
- 500MB free disk space

**Recommended:**
- Windows 11
- Excel 2021 or Microsoft 365
- 8GB RAM
- 1GB free disk space

### VBA Requirements

**References (Tools > References):**
- Visual Basic For Applications
- Microsoft Excel 16.0 Object Library
- Microsoft Office 16.0 Object Library
- Microsoft Scripting Runtime (for Dictionary objects)

### Data Limits

**Excel Limits:**
- Maximum rows: 1,048,576
- Maximum columns: 16,384
- Maximum cell characters: 32,767

**Practical Limits:**
- Recommended max entities: 500
- Recommended max FSLis: 1,000
- Recommended max segment tabs: 20

### File Formats

**Input:**
- Excel Workbook (.xlsx)
- Excel Macro-Enabled Workbook (.xlsm)

**Output:**
- Excel Workbook (.xlsx)
- Can be saved as .xlsm if formulas are added

### Security Considerations

**Macro Security:**
- Requires macros to be enabled
- Code is unprotected for customization
- No external data connections
- No internet access required

**Data Security:**
- All processing done locally
- No data transmitted externally
- No logging of sensitive information
- Temporary worksheets deleted after use

### Version History

**v1.0.0 (Initial Release):**
- Core tab categorization
- Data processing engine
- Table generation
- Basic Power BI integration support

### Known Limitations

1. **Language Support:**
   - Currently optimized for English language workbooks
   - Non-English headers may require code modification

2. **Format Variations:**
   - Assumes standard TGK format
   - Custom formats may require adaptation

3. **Formula Complexity:**
   - Complex nested formulas may not be analyzed correctly
   - Circular references not supported

4. **Power BI Integration:**
   - DAX measures must be created manually
   - No automatic Power BI file generation
   - Requires Power BI Desktop or Pro license

### Support and Maintenance

**Code Maintenance:**
- Modular structure allows easy updates
- Each module can be modified independently
- Comment throughout code for clarity

**Customization:**
- Add new categories in ModTabCategorization
- Modify table structure in ModTableGeneration
- Add validation in ModDataProcessing

**Testing:**
- Test with sample data before production use
- Verify output tables before Power BI import
- Validate percentage calculations

---

## Appendix

### A. Glossary

- **FSLi:** Financial Statement Line Item
- **Pack:** Entity or legal entity within consolidation
- **TGK:** Consolidation system name
- **Segment:** Business division or geographical region
- **Console:** Consolidated financial data
- **Discontinued:** Discontinued operations or disposed entities

### B. Sample Data Structure

```
Row 6:  | A        | B          | C (Original) | D (Original) | E (Consol) | F (Consol) |
Row 7:  | FSLi     |            | UK Entity    | US Entity    | UK Entity  | US Entity  |
Row 8:  | Code     |            | BVT-001      | BVT-002      | BVT-001    | BVT-002    |
Row 9:  |          | Revenue    | 1000         | 2000         | 950        | 1900       |
Row 10: |          | COGS       | 600          | 1200         | 570        | 1140       |
```

### C. Quick Reference

**Run Tool:** Click button → Enter workbook name → Categorize tabs → Select columns → Wait for completion

**Required Category:** TGK Input Continuing Operations Tab (must have exactly one)

**Output Location:** New workbook created automatically

**Power BI Import:** Get Data > Excel > Select output workbook > Import all tables

### D. Contact and Support

For issues, questions, or customization requests:
- Review this documentation
- Check troubleshooting section
- Verify VBA code comments
- Test with sample data

---

**Document Version:** 1.0.0  
**Last Updated:** 2024  
**Tool Version:** 1.0.0
