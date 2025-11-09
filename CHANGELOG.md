# Change Log - TGK Consolidation Scoping Tool

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2024-11-09

### Fixed
- **ActiveX Component Error**: Fixed "ActiveX component can't create object" error
  - Added proper error handling for Scripting.Dictionary creation
  - Provides clear instructions to enable Microsoft Scripting Runtime if needed
  - Improved error messages with step-by-step guidance

### Changed
- **Category Names Updated**:
  - Category 2: "TGK Discontinued Opt Tab" → "Discontinued Ops Tab"
  - Category 5: "TGK Console Continuing Tab" → "TGK Consol Continuing Tab"
  - Category 8: "Pull Workings" → "Paul workings"
- **Division Name Examples**: Updated to "UK Division, Properties Division, BIH division, etc."

### Added
- **New Category**: Trial Balance (Category 9)
  - Single tab only
  - Optional category for trial balance data
- Category 10 is now Uncategorized (previously Category 9)

### Updated
- All documentation files to reflect new category names
- Category selection prompts now accept numbers 1-10
- Validation logic updated to include Trial Balance as single-tab category

## [1.1.1] - 2024-11-09

### Fixed
- **Category Selection Dialog Flow**: Fixed critical bugs that prevented reliable category selection
  - Fixed recursive call issue in `ShowUncategorizedTabs()` that could cause dialog loops
  - Corrected return value handling to ensure proper restart functionality
  - Moved uncategorized tabs check inside validation loop for better user flow
  - Added automatic default division name ("Division_X") when user leaves segment division name empty
  - Enhanced confirmation dialogs for restart decisions to prevent confusion
  - Improved error recovery when users want to recategorize tabs

### Improved
- Better user guidance during category selection process
- More reliable pop-up dialog flow from start to finish
- Enhanced code comments for maintainability

## [1.1.0] - 2024-11-08

### Changed
- **Tab Categorization Interface**: Replaced worksheet-based categorization with pop-up dialog system
  - Removed temporary worksheet creation for categorization
  - Implemented sequential InputBox dialogs for each tab
  - Users now enter numbers (1-9) to select categories
  - Division names prompted immediately for segment tabs
  - Added input validation with retry loop for invalid entries
  - Improved user experience - no more tab switching required

### Fixed
- Resolved issue where users couldn't select categories in the worksheet-based interface
- Eliminated confusion from worksheet-based dropdown selection

## [1.0.0] - 2024-11-08

### Added - Initial Release

#### Core Functionality
- Main orchestration module (`ModMain.bas`) with entry point `StartScopingTool()`
- Tab discovery and workbook validation
- Interactive tab categorization system with 10 predefined categories (updated in v1.2.0)
- Dynamic column detection and user selection (Original vs Consolidation currency)
- FSLi hierarchy analysis with total/subtotal detection
- Entity and pack information extraction from standard TGK rows
- Automatic cell unmerging functionality

#### Table Generation
- Full Input Table generation from Input Continuing tab
- Journals Table generation from Journals Continuing tab
- Full Console Table generation from Consol Continuing tab (name updated in v1.2.0)
- Discontinued Table generation from Discontinued Operations tab
- FSLi Key Table with links to all data tables
- Pack Number Company Table with division mapping
- Automatic percentage table creation for all main tables
- Consistent table formatting with headers, borders, and colors

#### Tab Categorization
- TGK Segment Tabs (multiple allowed)
- Discontinued Ops Tab (single) - updated name in v1.2.0
- TGK Input Continuing Operations Tab (single, required)
- TGK Journals Continuing Tab (single)
- TGK Consol Continuing Tab (single) - updated name in v1.2.0
- TGK BS Tab (single)
- TGK IS Tab (single)
- Paul workings (multiple allowed) - updated name in v1.2.0
- Trial Balance (single) - added in v1.2.0
- Uncategorized (multiple allowed)

#### User Interface
- Welcome dialog with process overview
- Workbook name input prompt
- Tab categorization interface with pop-up dialog validation
- Column type selection dialog
- Progress status bar updates
- Completion confirmation message
- Uncategorized tabs warning and confirmation

#### Validation
- Workbook existence check
- Required category validation (Input Continuing mandatory)
- Single-tab category enforcement
- Division name prompting for segment tabs
- Empty row detection
- Data structure verification

#### Documentation
- Complete user documentation (DOCUMENTATION.md) - 90+ pages
- Power BI Integration Guide (POWERBI_INTEGRATION_GUIDE.md) - comprehensive DAX measures and workflows
- Installation Guide (INSTALLATION_GUIDE.md) - step-by-step setup instructions
- Usage Examples (USAGE_EXAMPLES.md) - 6 detailed real-world scenarios
- Quick Reference Guide (QUICK_REFERENCE.md) - one-page cheat sheet
- Enhanced README.md with architecture and features overview

#### Power BI Integration Support
- Table structure optimized for Power BI import
- Unpivoting guidance and instructions
- DAX measure templates for scoping analysis
- Threshold-based scoping workflow documentation
- Manual selection workflow documentation
- Hybrid approach examples
- Coverage percentage calculations

#### Error Handling
- Comprehensive error handlers in all main procedures
- User-friendly error messages
- Graceful degradation when optional tabs missing
- Status bar updates for progress tracking

### Technical Details

#### Modules Structure
- `ModMain.bas` (5.5 KB) - Entry point and orchestration
- `ModTabCategorization.bas` (10.7 KB) - Tab categorization logic
- `ModDataProcessing.bas` (15.5 KB) - Data extraction and analysis
- `ModTableGeneration.bas` (11.8 KB) - Table creation and formatting

#### Data Structures
- `ColumnInfo` type for column metadata
- `FSLiInfo` type for FSLi hierarchy information
- `TabCategory` type for tab categorization storage
- Dictionary objects for efficient lookups

#### Performance Features
- Screen updating disabled during processing
- Manual calculation mode during intensive operations
- Efficient collection and dictionary usage
- Batch processing of similar operations

### Dependencies
- Microsoft Excel 2016 or later
- VBA 7.0 or later
- Microsoft Scripting Runtime
- Windows OS

### Known Limitations
- English language workbooks only (currently)
- Standard TGK format required (rows 6-8 structure)
- No automatic Power BI file generation
- Manual DAX measure creation required in Power BI
- Limited to Excel file size limits (1M rows)

---

## [Unreleased]

### Planned for Future Releases

#### Version 1.1.0 (Planned)
- Multi-language support
- Custom FSLi hierarchy detection improvements
- Enhanced error recovery mechanisms
- Automated testing framework
- Performance optimizations for large workbooks (100+ tabs)

#### Version 1.2.0 (Planned)
- Direct Power BI file (.pbix) generation
- Automated DAX measure creation
- Historical comparison features
- Template library for common structures
- Advanced mathematical accuracy checks

#### Version 2.0.0 (Planned)
- Support for non-TGK consolidation systems
- Custom format configuration
- API for integration with other tools
- Cloud synchronization support
- Collaborative features

---

## Support and Contributions

### Reporting Issues
When reporting issues, please include:
- Tool version (check this file)
- Excel version and build number
- Description of the problem
- Steps to reproduce
- Error messages (if any)
- Sample data structure (if possible)

### Contributing
Contributions are welcome! Areas for enhancement:
- Additional language support
- Custom format detection
- Performance optimization
- Documentation improvements
- Bug fixes

---

## Version History Summary

| Version | Release Date | Key Features | Status |
|---------|-------------|--------------|--------|
| 1.0.0 | 2024-11-08 | Initial release with full functionality | Released |
| 1.1.0 | TBD | Multi-language, performance improvements | Planned |
| 1.2.0 | TBD | Power BI automation | Planned |
| 2.0.0 | TBD | Multi-system support | Planned |

---

## Breaking Changes

### Version 1.0.0
- Initial release, no breaking changes

---

## Migration Guide

### From Manual Process to Tool

If you've been processing TGK consolidations manually:

1. **Install the tool** following INSTALLATION_GUIDE.md
2. **Test with small workbook** first
3. **Document your categorization** for consistency
4. **Validate output** against manual process
5. **Adjust Power BI** to use new table structure
6. **Update documentation** to reflect new process

**Time to migrate:** 2-4 hours including testing

---

## Acknowledgments

### Credits
- Designed for audit professionals working with TGK systems
- Built to handle dynamic structures and flexible scoping
- Inspired by real-world consolidation scoping challenges

### Technologies
- VBA for Microsoft Excel
- Microsoft Scripting Runtime
- Power BI Desktop for visualization
- DAX for measures and calculations

---

**Maintained by:** TGK Scoping Tool Team  
**License:** See LICENSE file  
**Documentation:** See DOCUMENTATION.md

---

## Appendix: Version Numbering

This project uses [Semantic Versioning](https://semver.org/):

- **MAJOR** version (X.0.0): Incompatible API changes or breaking changes
- **MINOR** version (0.X.0): New functionality in a backwards compatible manner
- **PATCH** version (0.0.X): Backwards compatible bug fixes

### Current Version: 1.0.0
- **1** = Major version (initial release)
- **0** = Minor version (no new features yet)
- **0** = Patch version (no bug fixes yet)
