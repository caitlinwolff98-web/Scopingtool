# Quick Reference Guide - TGK Scoping Tool

## One-Page Quick Start

### Installation (5 minutes)
1. Create new Excel workbook → Save as `TGK_Scoping_Tool.xlsm`
2. Press `Alt+F11` → Import all `.bas` files from `VBA_Modules` folder
3. Insert Button → Assign macro `StartScopingTool`
4. Done!

### Usage (3 steps)
1. **Open** both workbooks (consolidation + tool)
2. **Click** "Start TGK Scoping Tool" button
3. **Follow** prompts → Categorize tabs → Select columns → Wait

### Tab Categories

| Category | Qty | Required | Example |
|----------|-----|----------|---------|
| **TGK Input Continuing** | 1 | ✓ Yes | TGK_Input_Continuing |
| TGK Segment Tabs | Many | No | TGK_UK, TGK_US |
| Discontinued Ops | 1 | No | TGK_Discontinued |
| TGK Journals Continuing | 1 | No | TGK_Journals |
| TGK Consol Continuing | 1 | No | TGK_Consol |
| TGK BS Tab | 1 | No | Balance_Sheet |
| TGK IS Tab | 1 | No | Income_Statement |
| Paul workings | Many | No | Working_Paper_1 |
| Trial Balance | 1 | No | Trial_Balance |

**Note:** Only Input Continuing is required

### Column Selection
- **Consolidation Currency** ← Recommended
- Original Currency ← Use only if needed

### Output Tables
- Full Input Table + Percentage
- Journals Table + Percentage (if exists)
- Console Table + Percentage (if exists)
- Discontinued Table + Percentage (if exists)
- FSLi Key Table
- Pack Number Company Table

### Power BI Import (Quick)
1. Get Data → Excel → Select output workbook
2. Select all tables → Transform Data
3. For each data table: Right-click "Pack" → Unpivot Other Columns
4. Rename: Attribute→FSLi, Value→Amount
5. Close & Apply

### Troubleshooting

| Problem | Solution |
|---------|----------|
| Can't find workbook | Check name includes .xlsx or .xlsm |
| No data in tables | Verify rows 6-8 structure |
| Tool slow | Close other apps, wait patiently |
| VBA error | Check all modules imported |

### Key Keyboard Shortcuts
- `Alt+F11` - Open VBA Editor
- `Ctrl+S` - Save
- `Alt+Q` - Close VBA Editor

### Expected Processing Time
- Small (5 tabs): 1-2 min
- Medium (10 tabs): 3-5 min
- Large (20 tabs): 5-10 min

### Scoping in Power BI
1. **Threshold:** Set FSLi + Amount → Auto-scope packs
2. **Manual:** Select specific pack + FSLi combinations
3. **Complete:** Select entire pack (all FSLis)

### Coverage Formula
```
Coverage % = Scoped Amount ÷ Total Amount
```

### Typical Coverage Targets
- Year-End Audit: 80-90%
- Quarterly Review: 60-70%
- New Entity: 100%

### Support
- Full docs: `DOCUMENTATION.md`
- Install help: `INSTALLATION_GUIDE.md`
- Examples: `USAGE_EXAMPLES.md`
- Power BI: `POWERBI_INTEGRATION_GUIDE.md`

---

## Common Commands

### VBA Editor
```
File → Import File → Select .bas → Open
Tools → References → Check "Microsoft Scripting Runtime"
```

### Excel
```
Developer → Insert → Button
Right-click button → Assign Macro → StartScopingTool
```

### Power BI
```
Transform → Right-click Pack → Unpivot Other Columns
Home → Close & Apply
```

---

## Checklist for Each Run

- [ ] Consolidation workbook is open
- [ ] Tool workbook is open
- [ ] Clicked "Start" button
- [ ] Entered correct workbook name
- [ ] Categorized all required tabs
- [ ] Selected column type
- [ ] Waited for completion
- [ ] Saved output workbook
- [ ] Verified sample data

---

## Quick Tips

**DO:**
✓ Save consolidation workbook before running
✓ Use consistent categorization
✓ Choose Consolidation Currency
✓ Document your scoping methodology
✓ Validate output with sample checks

**DON'T:**
✗ Close workbooks during processing
✗ Edit source workbook while tool runs
✗ Skip validation steps
✗ Ignore error messages
✗ Forget to save output

---

**Version:** 1.0.0 | **Page:** 1 of 1 | Print this page for easy reference!
