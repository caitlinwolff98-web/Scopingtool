# TGK Consolidation Scoping Tool

[![VBA](https://img.shields.io/badge/VBA-Excel-green)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Power BI](https://img.shields.io/badge/Power%20BI-Compatible-yellow)](https://powerbi.microsoft.com/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

## Overview

The **TGK Consolidation Scoping Tool** is a comprehensive, adaptive VBA script for Microsoft Excel designed to automate the analysis of TGK consolidation workbooks and create structured tables for Power BI integration. This tool streamlines the audit scoping process by:

- Automatically categorizing worksheets in consolidation workbooks
- Dynamically analyzing Financial Statement Line Items (FSLi) hierarchies
- Extracting entity and pack information across multiple segments
- Generating structured tables optimized for Power BI import
- Calculating percentage coverage for scoping analysis
- Supporting threshold-based and manual scoping methodologies

## Key Features

### 🔍 Intelligent Analysis
- **Dynamic Tab Discovery**: Automatically identifies and lists all worksheets
- **Flexible Categorization**: User-driven tab categorization with validation
- **Smart FSLi Detection**: Recognizes totals, subtotals, and hierarchies
- **Multi-Currency Support**: Handles both entity and consolidation currencies

### 📊 Comprehensive Table Generation
- **Full Input Table**: Complete view of input continuing operations
- **Journals Table**: Consolidation journal entries
- **Console Table**: Consolidated financial data
- **Discontinued Table**: Discontinued operations
- **FSLi Key Table**: Master reference for all FSLi entries
- **Pack Company Table**: Entity reference with divisions
- **Percentage Tables**: Automatic coverage percentage calculations

### 🔄 Power BI Integration
- Tables structured for seamless Power BI import
- Support for unpivoting and data transformation
- DAX measure templates included
- Interactive scoping workflows
- Threshold-based automation

### 🛡️ Robust & Reliable
- Comprehensive error handling
- Validation at each step
- Progress indicators
- Detailed logging
- Mathematical accuracy checks

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

3. **Import VBA Modules**
   - Press `Alt + F11` (VBA Editor)
   - Import all `.bas` files from `VBA_Modules` folder:
     - `ModMain.bas`
     - `ModTabCategorization.bas`
     - `ModDataProcessing.bas`
     - `ModTableGeneration.bas`

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

3. **Review**
   - Check generated tables in new workbook
   - Verify data accuracy
   - Export for Power BI

4. **Analyze in Power BI**
   - Import tables into Power BI
   - Follow Power BI Integration Guide
   - Create scoping dashboards

## Documentation

| Document | Description |
|----------|-------------|
| [DOCUMENTATION.md](DOCUMENTATION.md) | Complete user guide, technical specs, troubleshooting |
| [POWERBI_INTEGRATION_GUIDE.md](POWERBI_INTEGRATION_GUIDE.md) | Step-by-step Power BI setup and DAX measures |
| [INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md) | Detailed installation instructions |
| [USAGE_EXAMPLES.md](USAGE_EXAMPLES.md) | Real-world usage scenarios and examples |

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
TGK Scoping Tool
├── ModMain.bas
│   ├── Entry point (StartScopingTool)
│   ├── Workbook validation
│   └── Orchestration logic
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
│   └── Data extraction
│
└── ModTableGeneration.bas
    ├── Table creation
    ├── Percentage calculations
    └── Formatting
```

### Data Flow

```
Consolidation Workbook
        ↓
Tab Categorization
        ↓
Column Selection
        ↓
FSLi Analysis
        ↓
Data Extraction
        ↓
Table Generation
        ↓
Output Workbook
        ↓
Power BI Import
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

4. **Full Console Table** & **Full Console Percentage**
   - Consolidated financial data
   - With percentage coverage

5. **Discontinued Table** & **Discontinued Percentage**
   - Discontinued operations data
   - With percentage coverage

### Reference Tables
6. **FSLi Key Table**
   - All unique FSLi entries
   - Links to all data tables
   - Percentage columns

7. **Pack Number Company Table**
   - Pack name, code, division
   - Master entity reference

## Power BI Scoping Workflows

### 1. Threshold-Based Scoping
- Select key FSLi (e.g., "Net Revenue")
- Set monetary threshold (e.g., $300M)
- Automatically scope in packs exceeding threshold
- Include all FSLis for scoped packs

### 2. Manual Pack/FSLi Selection
- Select specific packs
- Select specific FSLis
- Fine-tune scope coverage
- Export for documentation

### 3. Hybrid Approach
- Start with threshold scoping
- Add manual selections
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
'   - TGK_Console
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

### v1.0.0 (Current)
- Initial release
- Core tab categorization
- Data processing engine
- Table generation
- Power BI integration support
- Comprehensive documentation

## Roadmap

### Future Enhancements
- [ ] Multi-language support
- [ ] Custom formula detection
- [ ] Direct Power BI integration
- [ ] Historical comparison features
- [ ] Automated testing framework
- [ ] Enhanced error recovery
- [ ] Template library for common structures

## Acknowledgments

Designed for audit professionals working with TGK consolidation systems. Built to handle dynamic structures and support flexible scoping methodologies.

---

**Current Version:** 1.0.0  
**Last Updated:** 2024  
**Platform:** Microsoft Excel with VBA  
**Integration:** Microsoft Power BI

---

## Quick Links

- [Complete Documentation](DOCUMENTATION.md)
- [Power BI Integration Guide](POWERBI_INTEGRATION_GUIDE.md)
- [Installation Guide](INSTALLATION_GUIDE.md)
- [Usage Examples](USAGE_EXAMPLES.md)

---

**Need Help?** Start with the [Quick Start](#quick-start) section above, then refer to the detailed documentation for your specific use case.