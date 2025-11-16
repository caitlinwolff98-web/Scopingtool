# TGK Consolidation Scoping Tool

[![VBA](https://img.shields.io/badge/VBA-Excel-green)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Power BI](https://img.shields.io/badge/Power%20BI-Compatible-yellow)](https://powerbi.microsoft.com/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

## Overview

The **Bidvest Scoping Tool** (formerly TGK Consolidation Scoping Tool) is a comprehensive, production-ready VBA solution for Microsoft Excel designed to automate the analysis of consolidation workbooks and create structured, interactive outputs for audit scoping. This tool streamlines the entire scoping process by:

- Automatically categorizing worksheets in consolidation workbooks
- Dynamically analyzing Financial Statement Line Items (FSLi) hierarchies with intelligent header filtering
- Extracting entity and pack information across multiple segments
- **NEW:** Threshold-based automatic scoping with user-configured FSLi selection
- **NEW:** Interactive Excel dashboard with pivot tables, charts, and calculators
- Generating structured tables optimized for Power BI import with standardized naming
- Calculating percentage coverage for scoping analysis
- **NEW:** Comprehensive scoping summary with "Suggested for Scope" recommendations
- Supporting both standalone Excel analysis and Power BI integration

## Key Features

### 🎯 Threshold-Based Auto-Scoping (NEW!)
- **Interactive FSLI Selection**: Choose which FSLIs to apply thresholds to
- **Custom Thresholds**: Set individual threshold values for each FSLI
- **Automatic Scoping**: Packs exceeding thresholds automatically marked as "Scoped In"
- **Configuration Tracking**: Threshold settings documented in output workbook
- **Coverage Analysis**: Track scoping coverage based on thresholds

### 🔍 Intelligent Analysis (Enhanced!)
- **Dynamic Tab Discovery**: Automatically identifies and lists all worksheets
- **Flexible Categorization**: User-driven tab categorization with validation
- **Smart FSLi Detection**: **IMPROVED** - Now correctly filters out statement headers like "INCOME STATEMENT" and "BALANCE SHEET"
- **Hierarchy Recognition**: Recognizes totals, subtotals, and indentation levels
- **Multi-Currency Support**: Handles both entity and consolidation currencies

### 📊 Comprehensive Table Generation
- **Full Input Table**: Complete view of input continuing operations
- **Journals Table**: Consolidation journal entries
- **Consol Table**: Consolidated financial data
- **Discontinued Table**: Discontinued operations
- **FSLi Key Table**: Master reference for all FSLi entries with metadata
- **Pack Company Table**: Entity reference with divisions
- **Percentage Tables**: Automatic coverage percentage calculations (4 tables)
- **NEW: Scoping Summary**: Pack-level summary with "Suggested for Scope" column
- **NEW: Threshold Configuration**: Documents applied thresholds and results

### 📈 Interactive Excel Dashboard (NEW!)
- **Standalone Functionality**: Full scoping analysis without Power BI
- **Pivot Tables**: Dynamic pack and FSLI analysis
- **Summary Charts**: Visual representation of scoping status
- **Scoping Calculator**: Coverage calculator with target setting
- **Auto-Filters**: Easy data exploration and filtering
- **Key Metrics Display**: Total packs, scoped in, pending review, coverage %

### 🔄 Power BI Integration (Enhanced!)
- **Standardized Output**: Always saves as "Bidvest Scoping Tool Output.xlsx"
- **Auto-Refresh Ready**: Consistent naming enables Power BI auto-refresh
- **Complete Setup Guide**: Step-by-step Power BI configuration (POWERBI_SETUP_COMPLETE.md)
- **DAX Measures Library**: Pre-built measures for common analyses
- **Relationship Fix**: Clear guidance on Pack Code vs Pack Name connections
- Support for unpivoting and data transformation
- Interactive scoping workflows

### 🛡️ Robust & Reliable
- Comprehensive error handling in all modules
- Validation at each step with user feedback
- Progress indicators (status bar updates)
- Detailed logging for troubleshooting
- Mathematical accuracy checks
- **Enhanced FSLI filtering** to exclude headers

## Quick Start

### Installation

1. **Download the VBA Modules**
   ```
   Clone or download this repository
   ```

2. **Create Macro Workbook**
   - Open Excel
   - Create new workbook
   - Save as `TGK_Scoping_Tool.xlsm`

3. **Import VBA Modules** (in this order)
   - Press `Alt + F11` (VBA Editor)
   - Import all `.bas` files from `VBA_Modules` folder:
     - `ModConfig.bas` (import first - dependencies)
     - `ModMain.bas`
     - `ModTabCategorization.bas`
     - `ModDataProcessing.bas`
     - `ModTableGeneration.bas`
     - `ModPowerBIIntegration.bas`
     - **NEW:** `ModThresholdScoping.bas` (threshold-based scoping)
     - **NEW:** `ModInteractiveDashboard.bas` (Excel dashboard)

4. **Add Button**
   - Return to Excel
   - Insert > Button (Form Control)
   - Assign macro: `StartScopingTool`
   - Label: "Start TGK Scoping Tool"

### Usage

1. **Prepare**
   - Open your TGK consolidation workbook
   - Open the TGK_Scoping_Tool.xlsm

2. **Run**
   - Click "Start TGK Scoping Tool" button
   - Enter consolidation workbook name
   - Categorize tabs using pop-up dialogs (enter numbers 1-9)
   - Select column type (Consolidation recommended)
   - **NEW:** Optionally configure threshold-based scoping:
     - Select FSLIs for threshold analysis
     - Enter threshold values for each FSLI
     - Packs exceeding thresholds automatically scoped in

3. **Review**
   - Output saved as: **"Bidvest Scoping Tool Output.xlsx"**
   - Check **Scoping Summary** sheet for recommendations
   - Review **Threshold Configuration** (if applicable)
   - Use **Interactive Dashboard** for analysis
   - Verify data accuracy in generated tables
   - Use **Scoping Calculator** for coverage planning

4. **Analyze**
   - **Option A - Excel Only**: Use Interactive Dashboard with pivot tables and charts
   - **Option B - Power BI**: Import tables into Power BI
     - File automatically named for easy refresh
     - Follow **POWERBI_SETUP_COMPLETE.md** for step-by-step setup
     - Create comprehensive scoping dashboards

## Documentation

| Document | Description |
|----------|-------------|
| [DOCUMENTATION.md](DOCUMENTATION.md) | Complete user guide, technical specs, troubleshooting |
| [POWERBI_SETUP_COMPLETE.md](POWERBI_SETUP_COMPLETE.md) | **NEW** Comprehensive Power BI setup with DAX measures and relationship fixes |
| [POWERBI_INTEGRATION_GUIDE.md](POWERBI_INTEGRATION_GUIDE.md) | Original Power BI integration guide |
| [CODE_IMPROVEMENTS.md](CODE_IMPROVEMENTS.md) | Version history, bug fixes, and enhancements |
| [INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md) | Detailed installation instructions |
| [USAGE_EXAMPLES.md](USAGE_EXAMPLES.md) | Real-world usage scenarios and examples |
| [VBA_Modules/README.md](VBA_Modules/README.md) | **UPDATED** Complete module documentation with new features |

## Requirements

### System Requirements
- Windows 10 or later
- Microsoft Excel 2016 or later
- Macro-enabled workbook support
- 4GB RAM minimum (8GB recommended)

### Power BI Requirements (for analysis)
- Power BI Desktop (latest version)
- Basic DAX knowledge (helpful)

### Excel Workbook Format
The tool expects TGK consolidation workbooks with:
- Row 6: Column type identifiers
- Row 7: Entity/Pack names
- Row 8: Entity/Pack codes
- Row 9+: FSLi data
- Column B: FSLi names

## Architecture

### Module Structure

```
Bidvest Scoping Tool (8 Modules)
├── ModMain.bas
│   ├── Entry point (StartScopingTool)
│   ├── Workbook validation
│   ├── Orchestration logic
│   ├── NEW: SaveOutputWorkbook (standardized naming)
│   └── NEW: CreateScopingSummarySheet
│
├── ModConfig.bas
│   ├── Configuration constants
│   ├── Utility functions
│   └── Validation helpers
│
├── ModTabCategorization.bas
│   ├── Tab discovery
│   ├── Category assignment
│   └── Validation rules
│
├── ModDataProcessing.bas
│   ├── Cell unmerging
│   ├── Column detection
│   ├── FSLi analysis
│   ├── NEW: IsStatementHeader (filters headers)
│   └── Data extraction
│
├── ModTableGeneration.bas
│   ├── Table creation
│   ├── Percentage calculations
│   └── Formatting
│
├── ModPowerBIIntegration.bas
│   ├── Power BI metadata
│   ├── Scoping configuration
│   └── DAX measures guide
│
├── ModThresholdScoping.bas (NEW)
│   ├── FSLI selection wizard
│   ├── Threshold configuration
│   ├── Automatic scoping logic
│   └── Configuration documentation
│
└── ModInteractiveDashboard.bas (NEW)
    ├── Dashboard creation
    ├── Pivot table generation
    ├── Chart creation
    ├── Scoping calculator
    └── Auto-filter setup
```

### Data Flow

```
Consolidation Workbook
        ↓
Tab Categorization
        ↓
Threshold Configuration (Optional - NEW)
        ↓
Column Selection
        ↓
FSLi Analysis (Enhanced - Filters Headers)
        ↓
Data Extraction
        ↓
Table Generation
        ↓
Scoping Summary Creation (NEW)
        ↓
Interactive Dashboard (NEW)
        ↓
Output Workbook ("Bidvest Scoping Tool Output.xlsx")
        ↓
Power BI Import (Optional)
        ↓
Scoping Analysis
```

## Supported Tab Categories

| Category | Quantity | Required | Description |
|----------|----------|----------|-------------|
| TGK Segment Tabs | Multiple | No | Business segments/divisions |
| Discontinued Ops Tab | Single | No | Discontinued operations |
| TGK Input Continuing Tab | **Single** | **Yes** | Primary input data |
| TGK Journals Continuing Tab | Single | No | Journal entries |
| TGK Consol Continuing Tab | Single | No | Consolidated data |
| TGK BS Tab | Single | No | Balance Sheet |
| TGK IS Tab | Single | No | Income Statement |
| Paul workings | Multiple | No | Working papers |
| Trial Balance | Single | No | Trial balance data |
| Uncategorized | Multiple | No | Ignored tabs |

## Output Tables

### Primary Data Tables
1. **Full Input Table**
   - Packs × FSLis matrix
   - Amounts from input continuing operations
   - Metadata tags for totals/subtotals

2. **Full Input Percentage**
   - Same structure as Full Input Table
   - Percentage of each amount vs. column total

3. **Journals Table** & **Journals Percentage**
   - Similar to Input tables
   - Data from journals continuing tab

4. **Full Consol Table** & **Full Consol Percentage**
   - Consolidated financial data
   - With percentage coverage

5. **Discontinued Table** & **Discontinued Percentage**
   - Discontinued operations data
   - With percentage coverage

### Reference Tables
6. **FSLi Key Table**
   - All unique FSLi entries (headers excluded)
   - Statement type metadata
   - Total/subtotal indicators
   - Indentation level

7. **Pack Number Company Table**
   - Pack name, code, division
   - Master entity reference

### New Interactive Sheets (v2.0)
8. **Scoping Summary** (NEW)
   - Pack-level scoping status
   - "Suggested for Scope" recommendations
   - Color-coded suggestions
   - Summary statistics

9. **Threshold Configuration** (NEW - if thresholds applied)
   - Configured FSLIs and threshold values
   - Packs automatically scoped in
   - Triggering FSLI for each pack

10. **Interactive Dashboard** (NEW)
    - Key metrics display
    - Pivot tables
    - Summary charts
    - Instructions

11. **Scoping Calculator** (NEW)
    - Coverage calculator
    - Target setting
    - Packs needed for target

## Scoping Workflows

### 1. Threshold-Based Scoping (Enhanced - NEW!)
**In VBA Tool:**
- User prompted to select FSLIs for threshold analysis
- Enter threshold value for each FSLI (e.g., $300M for Revenue)
- Tool analyzes data and automatically marks packs exceeding thresholds as "Scoped In"
- Configuration documented in "Threshold Configuration" sheet

**In Power BI:**
- Import "Scoping Summary" table
- Filter by "Scoped In" = "Yes"
- Review "Threshold Configuration" for audit trail
- Use DAX measures for dynamic threshold analysis

### 2. Excel-Based Interactive Analysis (NEW!)
- Use "Interactive Dashboard" sheet
- Explore data with pivot tables and charts
- Use "Scoping Calculator" for coverage planning
- Apply auto-filters for detailed analysis
- No Power BI needed!

### 3. Manual Pack/FSLi Selection (Power BI)
- Select specific packs
- Select specific FSLis
- Fine-tune scope coverage
- Export for documentation

### 4. Hybrid Approach
- Start with threshold-based auto-scoping in VBA
- Review suggestions in Scoping Summary
- Use Interactive Dashboard for initial analysis
- Fine-tune in Power BI if needed
- Optimize coverage percentage
- Document methodology

## Examples

### Example 1: Basic Usage

```vba
' User clicks button
' Enters: "Consolidation_2024_Q4.xlsx"
' Tool discovers 12 tabs:
'   - TGK_UK, TGK_US, TGK_EU (segments)
'   - TGK_Discontinued
'   - TGK_Input_Continuing
'   - TGK_Journals
'   - TGK_Consol
'   - Balance_Sheet, Income_Statement
'   - Pull_Working_1, Pull_Working_2
'   - Summary (uncategorized)
'
' User categorizes appropriately
' Tool generates 11 tables
' User imports to Power BI
```

### Example 2: Threshold Scoping

```
Power BI Workflow:
1. Select "Net Revenue" FSLi
2. Set threshold: $300,000,000
3. Result: 8 packs meet threshold
4. Coverage: 85% of total net revenue
5. Remaining: 15% untested
6. Document: Export scoped pack list
```

See [USAGE_EXAMPLES.md](USAGE_EXAMPLES.md) for more detailed scenarios.

## Troubleshooting

### Common Issues

| Issue | Solution |
|-------|----------|
| "Could not find workbook" | Ensure workbook is open and name matches exactly |
| "Required tabs are missing" | Categorize at least one tab as Input Continuing |
| Tool runs slowly | Disable screen updating, use manual calculation |
| No data in tables | Verify row 6-8 structure in consolidation workbook |
| **"INCOME STATEMENT" showing as FSLI** | **FIXED in v2.0** - Headers now filtered automatically |
| **"Suggested for Scope" column empty** | **FIXED in v2.0** - Now populated with recommendations |
| **Pack Names not connecting in Power BI** | Use Pack Code for relationships (see POWERBI_SETUP_COMPLETE.md) |

See [DOCUMENTATION.md](DOCUMENTATION.md) for comprehensive troubleshooting.

## Performance

### Typical Processing Times
- Small workbook (5 tabs, 200 FSLis): 30-60 seconds
- Medium workbook (10 tabs, 500 FSLis): 2-4 minutes
- Large workbook (20 tabs, 1000 FSLis): 5-10 minutes

### Optimization Tips
- Close unnecessary applications
- Ensure adequate memory (8GB+)
- Disable automatic calculation
- Process smaller segments separately if needed

## Security

- All processing done locally
- No external data transmission
- No internet access required
- Macros must be enabled (standard VBA requirement)
- Code is unprotected for customization

## Contributing

Contributions are welcome! Areas for enhancement:
- Additional language support
- Custom FSLi hierarchy detection
- Automated Power BI file generation
- Integration with other consolidation systems

## License

This project is provided as-is for audit and consolidation scoping purposes.

## Support

For questions, issues, or customization:
1. Review [DOCUMENTATION.md](DOCUMENTATION.md)
2. Check [Troubleshooting section](#troubleshooting)
3. Verify VBA code comments
4. Test with sample data

## Version History

### v2.0.0 (Current - November 2024)
**MAJOR UPDATE - Comprehensive Enhancement Release**

**🎯 New Features:**
- ✨ **Threshold-Based Auto-Scoping**: User-configurable FSLI thresholds with automatic pack scoping
- ✨ **Interactive Excel Dashboard**: Full scoping analysis without Power BI requirement
- ✨ **Scoping Summary**: Complete with "Suggested for Scope" recommendations
- ✨ **Standardized Output**: Always saves as "Bidvest Scoping Tool Output.xlsx"
- ✨ **Scoping Calculator**: Coverage calculator with target planning

**🔧 Bug Fixes & Improvements:**
- ✅ Fixed FSLI detection - now correctly filters out "INCOME STATEMENT" and "BALANCE SHEET" headers
- ✅ Fixed "Suggested for Scope" column - now properly populated
- ✅ Enhanced total/subtotal detection logic
- ✅ Better identification of actual line items vs headers

**📦 New Modules:**
- ModThresholdScoping.bas (350 lines) - Threshold configuration and automation
- ModInteractiveDashboard.bas (285 lines) - Excel-based interactive features

**📚 Documentation:**
- POWERBI_SETUP_COMPLETE.md - Comprehensive Power BI setup guide
- Enhanced DAX measures library
- Fixed Pack Name/Pack Code relationship guidance
- Complete auto-refresh configuration

**📊 Output Enhancement:**
- 20+ sheets generated (was 14)
- Pivot tables and charts included
- Auto-filters enabled
- Color-coded recommendations

### v1.1.0 (November 2024)
**Major Update - Enhanced Power BI Integration**
- 🔧 Fixed ambiguous name error in VBA code
- ✨ Added ModConfig for centralized configuration
- ✨ Added ModPowerBIIntegration for direct Power BI support
- ✨ Added 4 new Power BI integration sheets
- 🚀 Improved code robustness and error handling
- 📚 Enhanced documentation
- See [CODE_IMPROVEMENTS.md](CODE_IMPROVEMENTS.md) for details

### v1.0.0
- Initial release
- Core tab categorization
- Data processing engine
- Table generation
- Power BI integration support
- Comprehensive documentation

## Roadmap

### v2.1.0 (Planned)
- [ ] ✅ Threshold-Based Auto-Scoping (COMPLETED in v2.0)
- [ ] ✅ Interactive Excel Dashboard (COMPLETED in v2.0)
- [ ] ✅ Enhanced FSLI Detection (COMPLETED in v2.0)
- [ ] Direct Power BI .pbix file generation
- [ ] Historical comparison features

### Future Enhancements
- [ ] Multi-language support
- [ ] Custom formula detection
- [ ] Automated testing framework
- [ ] Enhanced error recovery
- [ ] Template library for common structures

## Acknowledgments

Designed for audit professionals working with TGK consolidation systems. Built to handle dynamic structures and support flexible scoping methodologies.

---

**Current Version:** 1.1.0  
**Last Updated:** November 2024  
**Platform:** Microsoft Excel with VBA  
**Integration:** Microsoft Power BI

---

## Quick Links

- [Complete Documentation](DOCUMENTATION.md)
- [Power BI Integration Guide](POWERBI_INTEGRATION_GUIDE.md)
- [Installation Guide](INSTALLATION_GUIDE.md)
- [Usage Examples](USAGE_EXAMPLES.md)

---

**Current Version:** 2.0.0  
**Last Updated:** November 2024  
**Platform:** Microsoft Excel with VBA  
**Integration:** Microsoft Power BI (Optional)  
**Output Format:** Bidvest Scoping Tool Output.xlsx (Standardized)

---

## Quick Links

- [Complete Documentation](DOCUMENTATION.md)
- [Power BI Setup Guide (NEW!)](POWERBI_SETUP_COMPLETE.md)
- [Power BI Integration Guide](POWERBI_INTEGRATION_GUIDE.md)
- [Installation Guide](INSTALLATION_GUIDE.md)
- [Usage Examples](USAGE_EXAMPLES.md)
- [Module Documentation](VBA_Modules/README.md)

---

**Need Help?** Start with the [Quick Start](#quick-start) section above, then refer to the detailed documentation for your specific use case.

**What's New in v2.0?** See [Version History](#version-history) for complete feature list and improvements.