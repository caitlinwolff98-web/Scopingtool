# Power BI Edit Mode - Complete Setup Guide

**Purpose:** Step-by-step instructions to enable manual data entry in Power BI for the Bidvest Scoping Tool

**Difficulty:** Intermediate  
**Time Required:** 15-20 minutes  
**Version:** 4.0

---

## What is Edit Mode?

**Edit Mode** in Power BI allows you to directly modify data in a table visual. This is CRITICAL for the Bidvest scoping tool because it enables you to:

- Change "Scoping Status" from "Not Scoped" to "Scoped In (Manual)"
- Update scoping decisions in real-time
- See coverage percentages update immediately
- Make manual pack/FSLI scoping decisions directly in Power BI

**Without edit mode,** you would need to:
- Export data to Excel
- Make changes in Excel
- Re-import to Power BI
- Refresh all visuals
- Repeat for every change (very tedious!)

---

## Prerequisites

Before starting, ensure you have:

✅ Power BI Desktop installed (latest version recommended)  
✅ Bidvest Scoping Tool Output.xlsx file  
✅ Tables imported into Power BI (see COMPREHENSIVE_GUIDE.md Section 5 if not done)  
✅ Relationships created between tables  
✅ Basic familiarity with Power BI interface

---

## Method 1: Enable Edit Mode in Power BI Desktop (RECOMMENDED)

This is the primary method for manual scoping in Power BI Desktop.

### Step 1: Create the Scoping Control Page

1. **Open your Power BI file** (the .pbix file with Bidvest data)

2. **Create a new page:**
   - Click the **+** icon at the bottom to add a new page
   - Rename it to **"Manual Scoping Control"**

3. **Add title:**
   - Insert → Text box
   - Type: "Manual Scoping Control"
   - Format: Bold, Size 18, Color Blue
   - Position at top of page

### Step 2: Add the Scoping Control Table Visual

1. **Insert Table Visual:**
   - Click on **Table** icon in Visualizations pane (looks like a grid)
   - Drag to create a large table visual (take up most of the page)

2. **Add fields to the table:**
   
   **From Scoping_Control_Table:**
   - Drag **Pack Name** to the table
   - Drag **Pack Code** to the table
   - Drag **Division** to the table
   - Drag **FSLI** to the table
   - Drag **Amount** to the table
   - Drag **Scoping Status** to the table ⭐ (CRITICAL)
   - Drag **Is Consolidated** to the table

3. **Your table should now have 7 columns:**
   ```
   | Pack Name | Pack Code | Division | FSLI | Amount | Scoping Status | Is Consolidated |
   ```

### Step 3: Format the Table

1. **Select the table visual** (click on it)

2. **Go to Format pane** (paint roller icon)

3. **Grid settings:**
   - Enable **Text size:** 9 or 10
   - Enable **Text wrap**
   - Enable **Row padding:** Small

4. **Column headers:**
   - Enable **Bold**
   - **Background color:** Dark blue (#4472C4)
   - **Text color:** White

5. **Conditional formatting for Amount:**
   - Click on **Amount** field in the table
   - Right-click → **Conditional formatting** → **Data bars**
   - Choose a color gradient (e.g., light blue to dark blue)

### Step 4: Enable Edit Mode (CRITICAL STEP)

This is where many users get stuck. Follow these steps exactly:

#### Option A: Enable Data Input (Power BI Desktop)

1. **Select the table visual**

2. **Go to Visualizations pane**

3. **Click on Format visual** (paint roller icon)

4. **Scroll down to "Values" section**

5. **Look for "Allow data entry" or "Edit mode"**
   - In some versions: **Visual → General → Properties → Allow data input**
   - In newer versions: **Visual → Advanced options → Allow data input**

6. **Toggle it ON** ✅

7. **If you don't see this option:**
   - Your table might not support editing
   - Try recreating the table visual
   - Ensure you're using a standard Table visual (not Matrix)
   - See Method 2 below for alternative

#### Option B: Use Power BI Service (If Desktop doesn't support editing)

If Power BI Desktop doesn't have edit mode for your table:

1. **Publish your report to Power BI Service:**
   - File → Publish → Select workspace
   - Wait for upload to complete

2. **Open report in Power BI Service** (web browser)

3. **Edit the report:**
   - Click **Edit** button at top
   - Select your Scoping Control Table

4. **Enable edit mode:**
   - Format pane → Values → Allow data input → ON

5. **Save the report**

6. **Download the .pbix file back:**
   - File → Download → Download .pbix
   - Replace your local file

### Step 5: Test Edit Mode

1. **Click on a cell in the "Scoping Status" column**

2. **You should see:**
   - Cell becomes editable (cursor appears)
   - Dropdown arrow appears (in some cases)
   - You can type or select a value

3. **Try changing a value:**
   - Click on a cell that says "Not Scoped"
   - Type: **Scoped In (Manual)**
   - Press **Enter**

4. **Verify the change:**
   - The cell should update
   - Other visuals should refresh (if you have measures that reference this)

5. **If it doesn't work:**
   - See Troubleshooting section below
   - Try Method 2 (alternative approach)

### Step 6: Add Slicers for Filtering

Make it easier to find specific packs/FSLIs:

1. **Add Pack Name slicer:**
   - Insert → Slicer
   - Add field: Pack_Number_Company_Table[Pack Name]
   - Position at top-left
   - Format: Dropdown style

2. **Add FSLI slicer:**
   - Insert → Slicer
   - Add field: FSLi_Key_Table[FSLI]
   - Position at top-center
   - Format: Dropdown style

3. **Add Division slicer:**
   - Insert → Slicer
   - Add field: Pack_Number_Company_Table[Division]
   - Position at top-right
   - Format: Dropdown style

### Step 7: Add Coverage Cards

Show the impact of your scoping decisions:

1. **Add "Total Scoped In" card:**
   - Insert → Card
   - Add measure: [Total Scoped In] (you need to create this DAX measure first)
   - Position bottom-left
   - Format: Large font, green background

2. **Add "Coverage %" card:**
   - Insert → Card
   - Add measure: [Coverage % by FSLI]
   - Position bottom-center
   - Format: Large font, blue background, percentage format

3. **Add "Untested %" card:**
   - Insert → Card
   - Add measure: [Untested %]
   - Position bottom-right
   - Format: Large font, red background, percentage format

---

## Method 2: Alternative Approach (If Edit Mode Not Available)

If you cannot enable edit mode in Power BI Desktop, use this workaround:

### Step A: Create Scoping Decisions Table in Excel

1. **Open a new Excel file**

2. **Create a table with these columns:**
   ```
   | Pack Code | FSLI | Manual Scoping Status |
   ```

3. **Add your scoping decisions:**
   - When you want to scope in a pack for a specific FSLI, add a row:
   - Example: `BVT-101 | Revenue | Scoped In`

4. **Save as:** `Scoping_Decisions.xlsx`

5. **Keep this file in the same folder** as your Bidvest Scoping Tool Output.xlsx

### Step B: Import Scoping Decisions to Power BI

1. **In Power BI Desktop:**
   - Home → Get Data → Excel
   - Browse to `Scoping_Decisions.xlsx`
   - Select the table → Load

2. **Create relationships:**
   ```
   Scoping_Decisions[Pack Code] → Scoping_Control_Table[Pack Code]
   Scoping_Decisions[FSLI] → Scoping_Control_Table[FSLI]
   ```

3. **Create a calculated column in Scoping_Control_Table:**
   ```DAX
   Final Scoping Status = 
   VAR ManualStatus = LOOKUPVALUE(
       Scoping_Decisions[Manual Scoping Status],
       Scoping_Decisions[Pack Code], Scoping_Control_Table[Pack Code],
       Scoping_Decisions[FSLI], Scoping_Control_Table[FSLI]
   )
   RETURN
   IF(
       ISBLANK(ManualStatus),
       Scoping_Control_Table[Scoping Status],
       ManualStatus
   )
   ```

4. **Use "Final Scoping Status" in your visuals** instead of "Scoping Status"

### Step C: Update Your Scoping Decisions

1. **When you want to change scoping:**
   - Open `Scoping_Decisions.xlsx`
   - Add/modify rows
   - Save the file

2. **In Power BI:**
   - Click **Refresh** button
   - All visuals update with new scoping decisions

**Pros:** Works around edit mode limitations  
**Cons:** Requires maintaining separate Excel file, less seamless

---

## Method 3: Pack-Level Scoping (Batch Updates)

For quickly scoping entire packs:

### Step 1: Create Pack-Level Decision Table

In Excel (or Power BI if edit mode works):

1. **Create a simple table:**
   ```
   | Pack Code | Pack Scoping Status |
   |-----------|-------------------|
   | BVT-101   | Scoped In         |
   | BVT-102   | Scoped In         |
   | BVT-103   | Not Scoped        |
   ```

2. **Save as:** `Pack_Scoping.xlsx`

### Step 2: Import and Apply in Power BI

1. **Import the table** (Home → Get Data → Excel)

2. **Create relationship:**
   ```
   Pack_Scoping[Pack Code] → Pack_Number_Company_Table[Pack Code]
   ```

3. **Update DAX measures to check Pack Scoping Status**

4. **When pack is "Scoped In":** All FSLIs for that pack are automatically included

---

## Troubleshooting Edit Mode

### Issue 1: "Allow data input" option not visible

**Possible causes:**
- Using Matrix visual instead of Table visual
- Power BI version too old
- Data source not supported for editing

**Solutions:**
1. **Check visual type:**
   - Delete the visual
   - Insert a new **Table** visual (not Matrix)
   - Add fields again

2. **Update Power BI Desktop:**
   - Help → About → Check for updates
   - Download latest version if available

3. **Use Power BI Service:**
   - Publish to Power BI Service
   - Enable edit mode there
   - Download .pbix back

4. **Use Method 2 (Alternative Approach)**

### Issue 2: Changes don't save

**Possible causes:**
- Not pressing Enter after editing
- Table not refreshing
- Data type mismatch

**Solutions:**
1. **Always press Enter after editing**

2. **Check data type:**
   - Scoping Status should be Text type
   - Right-click column → Data type → Text

3. **Refresh the visual:**
   - Right-click visual → Refresh

4. **Save the .pbix file:**
   - File → Save
   - Close and reopen to verify changes persist

### Issue 3: Dropdown not showing options

**Possible causes:**
- Dropdown behavior not enabled
- Values not pre-defined

**Solutions:**
1. **Pre-define valid values:**
   - Create a separate "Scoping Status Options" table
   - Values: "Not Scoped", "Scoped In (Manual)", "Scoped Out"
   - Create relationship to Scoping_Control_Table

2. **Use filter instead:**
   - Add a slicer with Scoping Status options
   - Select value in slicer
   - Then select rows in table to apply

### Issue 4: Edit mode works but measures don't update

**Possible causes:**
- Measures not referencing the edited column
- Relationships broken
- Cached data

**Solutions:**
1. **Verify measures use Scoping Status:**
   ```DAX
   Total Scoped In = 
   CALCULATE(
       DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
       Scoping_Control_Table[Scoping Status] IN {
           "Scoped In (Threshold)", 
           "Scoped In (Manual)"
       }
   )
   ```

2. **Check relationships:**
   - Model view → Verify all relationships active
   - Re-create if needed

3. **Clear cache:**
   - File → Options → Data Load → Clear cache
   - Refresh all data

### Issue 5: Permission errors

**Error:** "You don't have permission to edit this data"

**Solutions:**
1. **Check file permissions:**
   - Ensure you own the .pbix file
   - File is not read-only

2. **Check workspace permissions** (if using Service):
   - You need Edit permissions in the workspace
   - Contact admin if needed

---

## Best Practices for Edit Mode

### 1. Use Consistent Values

**Always use these exact values** (case-sensitive):
- `Not Scoped`
- `Scoped In (Manual)`
- `Scoped In (Threshold)`
- `Scoped Out`

**Tip:** Create a reference table with valid values to prevent typos

### 2. Save Frequently

- File → Save after making scoping changes
- Power BI doesn't auto-save edits
- Create backup copies before major changes

### 3. Document Changes

Keep a log of manual scoping decisions:
- Which packs scoped in manually
- Why they were scoped in
- Date of decision
- Who made the decision

**Tip:** Add a "Notes" column to Scoping_Control_Table for this

### 4. Test Before Production

- Practice with sample data first
- Verify measures update correctly
- Check export functionality works
- Train team on the workflow

### 5. Use Filters for Efficiency

Don't scroll through thousands of rows:
- Use slicers to filter to specific FSLIs
- Filter by Division to focus on one area
- Use search box to find specific packs

---

## Complete Workflow Example

Here's how to use edit mode for a typical scoping task:

### Scenario: Scope in packs for "Revenue" FSLI

**Step 1: Filter to Revenue**
1. Click on FSLI slicer
2. Select "Revenue"
3. Table now shows only Revenue rows

**Step 2: Sort by Amount**
1. Click on Amount column header
2. Sort descending (highest first)
3. Focus on largest amounts

**Step 3: Make Scoping Decisions**
1. Click on Scoping Status cell for BVT-101
2. Type: `Scoped In (Manual)`
3. Press Enter
4. Repeat for other packs you want to scope in

**Step 4: Review Coverage**
1. Look at Coverage % card
2. Target: 80% for high-risk FSLIs
3. Continue scoping packs until target reached

**Step 5: Export Results**
1. File → Export → PDF
2. Save as: `Revenue_Scoping_[Date].pdf`
3. Include in audit documentation

**Step 6: Move to Next FSLI**
1. Clear FSLI slicer (or select next FSLI)
2. Repeat process
3. Build comprehensive scoping across all FSLIs

---

## Video Tutorial (If Available)

**Coming Soon:** Video tutorial showing edit mode setup and usage

**For now:** Screenshots and step-by-step text guide above

---

## Quick Reference

### Enable Edit Mode Checklist

- [ ] Using Table visual (not Matrix)
- [ ] Visual selected
- [ ] Format pane open
- [ ] Values section expanded
- [ ] "Allow data input" toggled ON
- [ ] Power BI version is current
- [ ] If not working → Try Method 2

### Valid Scoping Status Values

✅ Correct:
- `Not Scoped`
- `Scoped In (Manual)`
- `Scoped In (Threshold)`
- `Scoped Out`

❌ Incorrect (will break measures):
- `not scoped` (lowercase)
- `Scoped In` (missing Manual/Threshold)
- `Scoped-In` (wrong separator)
- `YES` (wrong value)

### Keyboard Shortcuts

- **F2:** Edit selected cell (in some versions)
- **Enter:** Confirm edit
- **Esc:** Cancel edit
- **Tab:** Move to next cell
- **Shift+Tab:** Move to previous cell

---

## Additional Resources

**Power BI Documentation:**
- [Data entry in Power BI](https://docs.microsoft.com/en-us/power-bi/create-reports/desktop-data-entry)
- [Table visual documentation](https://docs.microsoft.com/en-us/power-bi/visuals/power-bi-visualization-tables)

**Bidvest Scoping Tool Documentation:**
- COMPREHENSIVE_GUIDE.md Section 5: Power BI Integration
- COMPREHENSIVE_GUIDE.md Section 6: Manual Scoping Workflow
- VISUALIZATION_ALTERNATIVES.md: Why Power BI was chosen

---

## Support

**If edit mode still doesn't work after trying all methods:**

1. Use Method 2 (Scoping Decisions Excel file) - always works
2. Check Power BI community forums for your specific error
3. Consider using Excel standalone for manual scoping
4. Contact Power BI support for technical issues

**Remember:** Even if edit mode doesn't work perfectly, you can still achieve manual scoping using the alternative methods documented above.

---

**Document Version:** 1.0  
**Last Updated:** November 2024  
**Next Update:** When Power BI introduces new features

**Feedback:** If you find a better way to enable edit mode, please share with the team!
