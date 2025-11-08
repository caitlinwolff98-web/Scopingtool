# Frequently Asked Questions (FAQ)

## General Questions

### What is the TGK Consolidation Scoping Tool?
The TGK Consolidation Scoping Tool is a VBA-based Excel macro that automates the analysis of TGK consolidation workbooks, extracts financial data, and creates structured tables for Power BI integration. It's designed specifically for audit scoping purposes.

### Who should use this tool?
- Auditors working with consolidated financial statements
- Finance professionals analyzing TGK consolidation data
- Audit managers planning scope coverage
- Data analysts preparing consolidation data for visualization

### Do I need programming knowledge to use this tool?
No. The tool is designed for non-programmers. You just need to:
1. Install the VBA modules (following step-by-step guide)
2. Click a button
3. Follow on-screen prompts

### Is this tool free to use?
Yes, the tool is provided as-is for consolidation scoping purposes. See the LICENSE file for details.

---

## Installation Questions

### What versions of Excel are supported?
- Excel 2016 or later
- Excel 2019
- Excel 2021
- Microsoft 365 Excel

Excel for Mac has limited VBA support and may not work properly.

### Do I need admin rights to install?
No admin rights required to:
- Create the Excel workbook
- Import VBA modules
- Run the tool

However, you may need permission to enable macros in Excel's Trust Center.

### Can I use this on Excel Online?
No. Excel Online has very limited VBA support. You must use Excel Desktop (Windows).

### How much disk space does it need?
- Tool itself: Less than 100 KB
- Output workbooks: Varies (typically 5-50 MB depending on data size)
- Recommended free space: 1 GB

### Why can't I see the Developer tab?
Enable it:
1. File → Options → Customize Ribbon
2. Check "Developer" in the right column
3. Click OK

---

## Usage Questions

### How long does processing take?
Depends on consolidation size:
- Small (5 tabs, 100 FSLis, 10 entities): 1-2 minutes
- Medium (10 tabs, 200 FSLis, 50 entities): 3-5 minutes
- Large (20 tabs, 500 FSLis, 150 entities): 8-12 minutes

### Can I run the tool on multiple workbooks simultaneously?
No. Process one workbook at a time. The tool needs to focus on a single source workbook.

### What if my consolidation workbook has a different structure?
The tool expects standard TGK format:
- Row 6: Column type identifiers
- Row 7: Entity names
- Row 8: Entity codes
- Row 9+: FSLi data
- Column B: FSLi names

If your structure differs, the tool may not work correctly. You may need to restructure the data or modify the VBA code.

### Can I categorize the same tab in multiple categories?
No. Each tab can only belong to one category. However, you can have multiple tabs in categories that allow multiple tabs (like Segment Tabs).

### What happens to uncategorized tabs?
They are ignored. The tool will warn you about uncategorized tabs and ask for confirmation to proceed.

### Can I stop the tool mid-process?
You can press Ctrl+Break to interrupt VBA execution, but this may leave the tool in an inconsistent state. Better to let it complete.

---

## Tab Categorization Questions

### Which category is required?
Only **TGK Input Continuing Operations Tab** is mandatory. All others are optional.

### How many tabs can I assign to each category?
- **Single tab only:**
  - TGK Discontinued Opt Tab
  - TGK Input Continuing Operations Tab
  - TGK Journals Continuing Tab
  - TGK Console Continuing Tab
  - TGK BS Tab
  - TGK IS Tab

- **Multiple tabs allowed:**
  - TGK Segment Tabs
  - Pull Workings
  - Uncategorized

### What if I have 10 segment tabs?
That's fine! The "TGK Segment Tabs" category allows multiple tabs. Categorize all of them and provide a division name for each.

### Do I need to provide division names?
Division names are required for segment tabs. They help organize the data in the Pack Number Company Table.

### What should I name my divisions?
Use clear, concise names:
- Geographic: "UK", "US", "Europe", "Asia Pacific"
- Business: "Retail", "Wholesale", "Manufacturing"
- Whatever makes sense for your organization

### Can I re-categorize if I make a mistake?
Yes, if you haven't clicked the final confirmation. If you've already started processing, you'll need to run the tool again from the beginning.

---

## Data Processing Questions

### Should I choose Original Currency or Consolidation Currency?
**Consolidation Currency is recommended** for most audit purposes. It shows data in the reporting currency after consolidation adjustments.

Use Original Currency only if you specifically need to analyze entity-level currency data.

### Why does the tool unmerge cells?
Merged cells can cause data extraction issues. The tool unmerges everything to ensure reliable data reading.

### What if my FSLi names have special characters?
The tool handles most special characters. However, avoid:
- Control characters
- Very long names (>255 characters)
- Leading/trailing spaces

### How does the tool detect totals vs. subtotals?
It looks for:
- "Total" or "Subtotal" in the FSLi name
- Indentation levels in Column B
- Mathematical relationships between rows

### What if my data has formulas?
The tool reads cell values, not formulas. As long as formulas calculate correctly, the tool will extract the results.

---

## Output Questions

### Where are the output tables created?
In a new workbook that opens automatically when processing completes.

### What do I name the output workbook?
Save it with a descriptive name like:
`[Entity]_Consolidation_[Period]_Output.xlsx`

Example: `Global_Consolidation_Q4_2024_Output.xlsx`

### Can I modify the output tables?
Yes! They're regular Excel tables. You can:
- Add columns
- Add formulas
- Format differently
- Filter data

However, extensive modifications may break Power BI integration.

### What's the difference between "Table" and "Percentage" tables?
- **Table**: Actual amounts (e.g., Full Input Table)
- **Percentage**: Each amount as % of column total (e.g., Full Input Percentage)

Percentage tables are used for coverage analysis in Power BI.

### How are percentages calculated?
```
Percentage = (Absolute Value of Amount) / (Sum of Absolute Values in Column) × 100
```

Absolute values are used so negative amounts (like expenses) don't distort percentages.

---

## Power BI Questions

### Do I need Power BI to use this tool?
No. The tool creates Excel tables that are useful on their own. Power BI is optional but recommended for advanced scoping analysis.

### What version of Power BI do I need?
Power BI Desktop (latest version). Free download from Microsoft.

### Can I use Power BI Service (cloud)?
Yes, after creating the report in Power BI Desktop, you can publish to Power BI Service.

### How do I import the tables into Power BI?
1. Open Power BI Desktop
2. Get Data → Excel
3. Select your output workbook
4. Select all tables
5. Click Transform Data
6. Follow transformation steps in POWERBI_INTEGRATION_GUIDE.md

### What's "unpivoting" and why do I need it?
Unpivoting converts wide tables (many columns) to long tables (fewer columns, more rows). This format is better for Power BI analysis.

Before unpivot:
```
| Pack | Revenue | COGS |
| A    | 100     | 60   |
```

After unpivot:
```
| Pack | FSLi    | Amount |
| A    | Revenue | 100    |
| A    | COGS    | 60     |
```

### Do I need to create DAX measures?
Yes, for advanced scoping functionality. Templates are provided in POWERBI_INTEGRATION_GUIDE.md. You can copy and paste them.

### Can Power BI automatically scope based on thresholds?
Not automatically in real-time, but you can create DAX measures that calculate which packs meet thresholds and visualize them.

---

## Scoping Questions

### What's a good coverage percentage?
Industry standards:
- **Year-end audit:** 80-90%
- **Quarterly review:** 60-70%
- **Risk-based audit:** Varies, typically 70-85%

Your engagement may have different requirements.

### How do I document my scoping?
1. Export tables from Power BI to Excel
2. Create scoping memo documenting:
   - Thresholds used
   - Manual selections and rationale
   - Final coverage percentage
   - Justification for untested items
3. Save for audit file

### Can I use different thresholds for different FSLis?
Yes! In Power BI:
1. Set threshold for Revenue (e.g., $300M)
2. Review which packs meet threshold
3. Change to Total Assets (e.g., $500M)
4. Review which packs meet this threshold
5. Combined scope includes both

### What if my coverage is below target?
Options:
1. Lower thresholds to include more packs
2. Manually select additional packs/FSLis
3. Select complete packs (all FSLis)
4. Document rationale if coverage must stay below target

### How do I handle new entities acquired mid-year?
1. Run tool on updated consolidation
2. Identify new entities in Pack Number Company Table
3. Manually scope ALL FSLis for new entities (100% coverage)
4. Apply normal scoping to existing entities

---

## Troubleshooting Questions

### The tool says "Could not find workbook" - what's wrong?
Check:
1. Is the consolidation workbook actually open?
2. Did you enter the exact name including extension (.xlsx or .xlsm)?
3. Are there any special characters in the name?

### I get "Required tabs are missing" - why?
You didn't categorize any tab as "TGK Input Continuing Operations Tab". This is mandatory. Re-run and ensure at least one tab has this category.

### Tool is very slow - is this normal?
Processing time depends on data size. Factors:
- Number of tabs
- Number of FSLis
- Number of entities
- Computer speed

To improve performance:
- Close other applications
- Disable antivirus temporarily
- Use faster computer if available

### Excel crashes during processing - what do I do?
1. Save all work first
2. Close Excel completely
3. Reopen both workbooks
4. Try again
5. If it keeps crashing:
   - Process smaller sections separately
   - Check if consolidation workbook has errors
   - Ensure enough RAM available (8GB+ recommended)

### Output tables have #N/A or errors - why?
Possible causes:
1. Source data has errors
2. Formulas in source workbook not calculated
3. Missing data in expected rows

Fix:
1. In consolidation workbook: Formulas → Calculate Now
2. Check for errors in source data
3. Re-run the tool

### Some entities are missing from Pack Number Company Table - why?
The tool extracts entities from segment tabs (rows 7-8). If entities only appear in Input Continuing but not in any segment tab, they won't be in this table.

Solution: Ensure all entities appear in at least one segment or discontinued tab.

---

## Performance Questions

### Can I make the tool faster?
Code optimization options:
1. Edit VBA to disable screen updating (already included)
2. Use manual calculation (already included)
3. Process fewer categories at once
4. Reduce data size in source workbook

### What's the maximum workbook size the tool can handle?
Theoretical limit is Excel's limit (1M rows), but practical limits:
- **Recommended max:** 150 entities, 500 FSLis, 20 tabs
- **Beyond this:** Performance degrades significantly

For very large consolidations, consider:
- Processing divisions separately
- Combining results manually
- Using 64-bit Excel for more memory

### Does the tool work on older/slower computers?
Yes, but slower. Minimum:
- 4GB RAM
- Dual-core processor
- Windows 7 or later

For better performance:
- 8GB+ RAM
- Quad-core processor
- SSD drive

---

## Customization Questions

### Can I modify the VBA code?
Yes! The code is unprotected and documented with comments. Common modifications:
- Add new categories
- Change table structure
- Add custom validation
- Modify formatting

### Can I add my own categories?
Yes. Edit `ModTabCategorization.bas`:
1. Add new constant (e.g., `CAT_CUSTOM = "Custom Category"`)
2. Add to validation lists
3. Handle in processing logic

### Can I change the output table format?
Yes. Edit `ModTableGeneration.bas`:
1. Modify `FormatAsTable` subroutine
2. Change colors, fonts, borders
3. Add custom columns

### Will updates break my customizations?
Possibly. If we release updates, you may need to:
1. Back up your modified code
2. Compare with new version
3. Re-apply customizations
4. Test thoroughly

---

## Power BI Advanced Questions

### Can I schedule automatic refresh in Power BI?
Yes, if published to Power BI Service:
1. Set up gateway to access Excel file
2. Configure scheduled refresh
3. Choose frequency (e.g., daily)

Note: Source workbook must be accessible to gateway.

### How do I handle multiple periods in Power BI?
Options:
1. **Separate reports:** One report per period
2. **Combined dataset:** Add "Period" column, combine all periods
3. **Historical comparison:** Import multiple outputs, create relationships

### Can I create a Power BI template?
Yes!
1. Create report with all visuals and measures
2. Delete data connection
3. Save as .pbit (Power BI Template)
4. Share template with team
5. Each user connects to their data

### How do I share my Power BI report?
Options:
1. **Power BI Desktop file:** Share .pbix file directly
2. **Power BI Service:** Publish and share via workspace
3. **PDF Export:** File → Export to PDF
4. **PowerPoint:** Export visuals to PowerPoint

---

## Support Questions

### Where can I get help?
1. Read DOCUMENTATION.md (comprehensive guide)
2. Check this FAQ
3. Review USAGE_EXAMPLES.md for scenarios
4. Check code comments in VBA modules

### How do I report a bug?
Include:
1. Tool version (see CHANGELOG.md)
2. Excel version
3. Steps to reproduce
4. Error message (exact text)
5. Sample data structure (if possible)

### Can I request new features?
Yes! Document your request with:
1. Use case description
2. Expected behavior
3. How it would improve the tool
4. Example scenario

### Is training available?
This package includes extensive documentation:
- INSTALLATION_GUIDE.md - Installation help
- USAGE_EXAMPLES.md - 6 detailed examples
- QUICK_REFERENCE.md - Cheat sheet
- POWERBI_INTEGRATION_GUIDE.md - Power BI training

---

## Security & Privacy Questions

### Is my data secure?
Yes:
- All processing happens locally on your computer
- No data sent to external servers
- No internet connection required
- No data logging

### Can others see my data?
Only if you share:
- The output workbook
- The Power BI report
- Screenshots or exports

Keep files secure according to your organization's policies.

### Does the tool contain viruses?
No. The code is pure VBA with no:
- External connections
- File downloads
- System modifications
- Hidden functionality

You can review all code in the VBA editor (Alt+F11).

### What about macro security warnings?
This is normal for VBA macros. To safely use:
1. Review the code if concerned
2. Enable macros when prompted
3. Add tool location to Trusted Locations if used regularly

---

## Best Practices

### What's the recommended workflow?
1. **Preparation:** Ensure consolidation workbook is finalized
2. **Execution:** Run tool during non-peak hours
3. **Validation:** Check sample data in output
4. **Analysis:** Import to Power BI
5. **Documentation:** Save scoping documentation
6. **Review:** Have colleague review scope
7. **Archive:** Save all files for audit trail

### How often should I run the tool?
- **Year-end:** Once, on final consolidation
- **Quarterly:** Each quarter on quarterly consolidation
- **Monthly:** If monthly consolidations exist
- **Ad-hoc:** Whenever consolidation changes significantly

### Should I keep old output files?
Yes, recommended for:
- Audit trail
- Historical comparison
- Trend analysis
- Documentation

Archive with clear naming: `[Entity]_[Period]_Output_[Date].xlsx`

### What documentation should I maintain?
1. Output workbooks (all periods)
2. Scoping memos (methodology and coverage)
3. Power BI reports (saved as .pbix)
4. Change log (if you modify the tool)
5. Training materials (for team members)

---

**FAQ Version:** 1.0.0  
**Last Updated:** 2024-11-08  
**Tool Version:** 1.0.0

**Can't find your question?** Check DOCUMENTATION.md for more detailed information.
