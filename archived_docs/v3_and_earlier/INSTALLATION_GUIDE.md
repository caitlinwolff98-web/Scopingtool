# Installation Guide - TGK Consolidation Scoping Tool

## Table of Contents
1. [Prerequisites](#prerequisites)
2. [Step-by-Step Installation](#step-by-step-installation)
3. [Verification](#verification)
4. [Troubleshooting Installation](#troubleshooting-installation)
5. [Uninstallation](#uninstallation)

---

## Prerequisites

### Software Requirements

#### Required
- **Operating System:** Windows 10 or later
- **Microsoft Excel:** 2016 or later (Desktop version)
  - Excel for Microsoft 365
  - Excel 2021
  - Excel 2019
  - Excel 2016

#### Optional (for full functionality)
- **Power BI Desktop:** Latest version (for analysis)
- **Microsoft Scripting Runtime:** Usually pre-installed

### System Requirements

**Minimum:**
- 4GB RAM
- 500MB free disk space
- 1280x720 display resolution

**Recommended:**
- 8GB RAM or more
- 1GB free disk space
- 1920x1080 display resolution
- SSD for better performance

### Excel Configuration

#### 1. Enable Macros

**Excel 2016/2019/2021/365:**

1. Open Excel
2. Click **File** → **Options**
3. Click **Trust Center** → **Trust Center Settings**
4. Click **Macro Settings**
5. Select one of:
   - **Disable all macros with notification** (Recommended)
   - **Enable all macros** (Not recommended for general use)
6. Click **OK** → **OK**

**Note:** With "Disable all macros with notification", you'll see a security warning when opening the tool. Click "Enable Content" to proceed.

#### 2. Enable Developer Tab

1. Click **File** → **Options**
2. Click **Customize Ribbon**
3. In the right column, check **Developer**
4. Click **OK**

#### 3. Enable VBA References (Advanced)

Usually automatic, but if needed:

1. Press **Alt + F11** to open VBA Editor
2. Click **Tools** → **References**
3. Ensure these are checked:
   - ☑ Visual Basic For Applications
   - ☑ Microsoft Excel 16.0 Object Library
   - ☑ Microsoft Office 16.0 Object Library
   - ☑ Microsoft Scripting Runtime
4. Click **OK**

---

## Step-by-Step Installation

### Method 1: Manual Installation (Recommended)

#### Step 1: Download Files

1. Download or clone this repository
2. Locate the `VBA_Modules` folder
3. Ensure you have all four `.bas` files:
   - `ModMain.bas`
   - `ModTabCategorization.bas`
   - `ModDataProcessing.bas`
   - `ModTableGeneration.bas`

#### Step 2: Create Macro Workbook

1. **Open Microsoft Excel**

2. **Create New Workbook**
   - Click **File** → **New** → **Blank Workbook**

3. **Save as Macro-Enabled**
   - Click **File** → **Save As**
   - Choose location (recommend: Desktop or Documents)
   - Enter filename: `TGK_Scoping_Tool`
   - **Important:** Change **Save as type** to: `Excel Macro-Enabled Workbook (*.xlsm)`
   - Click **Save**

#### Step 3: Import VBA Modules

1. **Open VBA Editor**
   - Press **Alt + F11** (or click **Developer** → **Visual Basic**)

2. **Import First Module**
   - Click **File** → **Import File...**
   - Navigate to `VBA_Modules` folder
   - Select `ModMain.bas`
   - Click **Open**

3. **Verify Import**
   - In Project Explorer (left pane), you should see:
     ```
     VBAProject (TGK_Scoping_Tool.xlsm)
       ├─ Microsoft Excel Objects
       │    └─ Sheet1 (Sheet1)
       │    └─ ThisWorkbook
       └─ Modules
            └─ ModMain
     ```

4. **Import Remaining Modules**
   - Repeat Step 2 for:
     - `ModTabCategorization.bas`
     - `ModDataProcessing.bas`
     - `ModTableGeneration.bas`

5. **Verify All Modules**
   - Project Explorer should now show:
     ```
     └─ Modules
          ├─ ModMain
          ├─ ModTabCategorization
          ├─ ModDataProcessing
          └─ ModTableGeneration
     ```

6. **Close VBA Editor**
   - Press **Alt + Q** or click the **X**

#### Step 4: Create User Interface Button

1. **Return to Excel**
   - You should see your blank workbook

2. **Rename Sheet (Optional)**
   - Right-click **Sheet1** → **Rename**
   - Enter: `Control Panel`
   - Press **Enter**

3. **Add Title (Optional)**
   - In cell A1, type: `TGK Consolidation Scoping Tool`
   - Make it bold and increase font size
   - In cell A2, type: `Click the button below to start`

4. **Insert Button**
   - Click **Developer** tab
   - Click **Insert** → **Button (Form Control)**
     - It's the first icon in the Form Controls section
   - Your cursor will change to a crosshair
   - Click and drag on the worksheet to draw a button
     - Recommended size: 200 pixels wide, 50 pixels high

5. **Assign Macro**
   - The "Assign Macro" dialog will appear automatically
   - Select `StartScopingTool` from the list
   - Click **OK**

6. **Edit Button Text**
   - The button should still be selected (with handles around it)
   - If not, right-click the button → **Edit Text**
   - Delete the default text (e.g., "Button 1")
   - Type: `Start TGK Scoping Tool`
   - Click outside the button to finish editing

7. **Format Button (Optional)**
   - Right-click the button → **Format Control**
   - Font tab: Choose font, size, style
   - Colors and Lines tab: Choose fill color
   - Click **OK**

#### Step 5: Add Instructions (Optional)

1. **Add Usage Instructions**
   - In cell A4, type: `Instructions:`
   - In cells A5-A10, add:
     ```
     1. Open your TGK consolidation workbook
     2. Keep it open alongside this tool
     3. Click the button above
     4. Follow the on-screen prompts
     5. Categorize your tabs
     6. Wait for processing to complete
     ```

2. **Format Sheet**
   - Adjust column widths
   - Add borders if desired
   - Set print area if needed

#### Step 6: Save and Test

1. **Save the Workbook**
   - Press **Ctrl + S**
   - Verify it's saved as `.xlsm`

2. **Initial Test**
   - Click the **Start TGK Scoping Tool** button
   - You should see the welcome message
   - Click **Cancel** (we're just testing)
   - If you see the message, installation is successful!

---

### Method 2: Quick Installation

If you're comfortable with VBA:

1. Create new workbook → Save as `.xlsm`
2. Alt + F11 → Import all four `.bas` files
3. Alt + Q → Insert button → Assign `StartScopingTool`
4. Done!

---

## Verification

### Verify Installation Checklist

Use this checklist to ensure everything is installed correctly:

- [ ] Excel version is 2016 or later
- [ ] Macros are enabled in Trust Center
- [ ] Workbook saved as `.xlsm` format
- [ ] VBA Editor shows all 4 modules:
  - [ ] ModMain
  - [ ] ModTabCategorization
  - [ ] ModDataProcessing
  - [ ] ModTableGeneration
- [ ] Button exists on worksheet
- [ ] Button is assigned to `StartScopingTool` macro
- [ ] Clicking button shows welcome message
- [ ] No VBA compile errors

### Test Basic Functionality

1. **Create Test Workbook**
   - Create a new Excel workbook
   - Name it: `Test_Consolidation.xlsx`
   - Add a few sheets with any names
   - Keep it open

2. **Run Tool**
   - Open `TGK_Scoping_Tool.xlsm`
   - Click the button
   - Enter: `Test_Consolidation.xlsx`
   - Click through categorization
   - Cancel before full processing

3. **Expected Result**
   - Tool finds the test workbook ✓
   - Tab categorization interface appears ✓
   - No error messages ✓

---

## Troubleshooting Installation

### Issue 1: "Macros have been disabled"

**Symptoms:**
- Security warning bar appears at top of Excel
- Button doesn't work

**Solutions:**

**Option A: Enable for this file**
1. Click **Enable Content** in the security warning
2. Try the button again

**Option B: Add location to Trusted Locations**
1. **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Click **Trusted Locations**
3. Click **Add new location**
4. Browse to folder containing the tool
5. Check **Subfolders of this location are also trusted**
6. Click **OK** → **OK** → **OK**
7. Close and reopen the file

### Issue 2: "Cannot import .bas file"

**Symptoms:**
- Import fails
- "File not found" error

**Solutions:**
1. Verify `.bas` files exist in `VBA_Modules` folder
2. Check file extensions are visible (Windows File Explorer → View → File name extensions)
3. Ensure files aren't blocked:
   - Right-click `.bas` file → Properties
   - If you see "Unblock" checkbox, check it
   - Click OK
   - Try importing again

### Issue 3: "Compile error" when clicking button

**Symptoms:**
- Error message about missing references
- Code doesn't run

**Solutions:**
1. Press **Alt + F11** to open VBA Editor
2. Click **Tools** → **References**
3. Look for items marked "MISSING"
4. Uncheck any MISSING items
5. Find and check **Microsoft Scripting Runtime**
6. Click **OK**
7. Try again

### Issue 4: Button doesn't appear

**Symptoms:**
- Successfully inserted button but can't see it

**Solutions:**
1. Check if Developer tab is visible
2. Try Insert → Illustrations → Shapes → Rectangle instead
3. Right-click shape → Assign Macro → `StartScopingTool`

### Issue 5: "Ambiguous name detected"

**Symptoms:**
- VBA compile error
- Mentions duplicate procedures

**Solutions:**
1. You may have imported modules twice
2. **Alt + F11** → Check Modules folder
3. Delete duplicate modules:
   - Right-click duplicate → Remove
   - Click **No** when asked to export
4. Keep only one copy of each module

### Issue 6: Excel version too old

**Symptoms:**
- Features don't work
- Compatibility errors

**Solutions:**
1. Check Excel version: **File** → **Account** → **About Excel**
2. If older than 2016, consider upgrading
3. Or use Excel Online (limited VBA support)

---

## Uninstallation

### Remove the Tool

1. **Delete the File**
   - Close `TGK_Scoping_Tool.xlsm`
   - Navigate to the file location
   - Delete `TGK_Scoping_Tool.xlsm`
   - Empty Recycle Bin

2. **Remove from Recent Files** (optional)
   - Open Excel
   - **File** → **Open** → **Recent**
   - Right-click the tool → **Remove from list**

3. **Remove from Trusted Locations** (if added)
   - **File** → **Options** → **Trust Center** → **Trust Center Settings**
   - **Trusted Locations**
   - Select the location → **Remove**
   - Click **OK**

### Preserve Output

Before uninstalling, if you want to keep generated tables:
1. Save output workbooks separately
2. Copy tables to another workbook
3. Export to Power BI first

---

## Advanced Installation Options

### Option 1: Network Installation

For deploying to multiple users:

1. **Central Location**
   - Place `TGK_Scoping_Tool.xlsm` on network drive
   - Set appropriate permissions (read-only for users)

2. **Shortcuts**
   - Create desktop shortcuts pointing to network file
   - Users open from network location

3. **Considerations**
   - Network speed affects performance
   - Multiple users can't edit simultaneously
   - Consider copying to local machine for better performance

### Option 2: Template Installation

Create as Excel Template:

1. Save file as `TGK_Scoping_Tool.xltm` (Template format)
2. Place in Excel Templates folder:
   - `C:\Users\[Username]\Documents\Custom Office Templates`
3. Access via **File** → **New** → **Personal**

### Option 3: Add-In Installation (Advanced)

Convert to Excel Add-In:

1. **Prepare Code**
   - Modify code to work as add-in
   - Add ribbon customization XML

2. **Save as Add-In**
   - **File** → **Save As**
   - Choose **Excel Add-In (*.xlam)**

3. **Load Add-In**
   - **File** → **Options** → **Add-Ins**
   - Manage: **Excel Add-ins** → **Go**
   - **Browse** → Select your `.xlam` file
   - Check the box → **OK**

**Note:** Add-in installation requires code modifications not included in this version.

---

## Post-Installation

### Next Steps

1. **Read Documentation**
   - Review [DOCUMENTATION.md](DOCUMENTATION.md)
   - Understand tab categories
   - Learn workflow

2. **Test with Sample Data**
   - Create simple test consolidation workbook
   - Run tool end-to-end
   - Verify output tables

3. **Set Up Power BI** (optional)
   - Install Power BI Desktop
   - Follow [POWERBI_INTEGRATION_GUIDE.md](POWERBI_INTEGRATION_GUIDE.md)

4. **Train Users**
   - Share documentation
   - Demonstrate workflow
   - Document any customizations

### Customization

The tool can be customized:
- Modify categories in `ModTabCategorization`
- Adjust table structures in `ModTableGeneration`
- Add validation in `ModDataProcessing`
- Enhance UI in `ModMain`

All code is accessible and documented with comments.

---

## Support

If you encounter issues not covered here:

1. Check [DOCUMENTATION.md](DOCUMENTATION.md) Troubleshooting section
2. Review VBA code comments
3. Test with minimal sample data
4. Verify all prerequisites are met

---

**Installation Guide Version:** 1.0.0  
**Last Updated:** 2024  
**Tool Version:** 1.0.0
