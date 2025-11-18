# EXCEL MACRO WORKBOOK - IMPORT GUIDE

**Date:** 2025-11-18
**Purpose:** Import VBA modules into Excel to create the ISA 600 Scoping Tool macro workbook
**Files Required:** 8 VBA module files (.bas) from VBA_Modules/ folder

---

## WHY WE CAN'T GENERATE .XLSM FILES DIRECTLY

**Technical Limitation:** VBA code exists as TEXT files (.bas modules). Excel macro workbooks (.xlsm) are BINARY files that cannot be generated programmatically without Excel automation.

**Solution:** Import the .bas module files into Excel to create the macro workbook.

---

## METHOD 1: MANUAL IMPORT (RECOMMENDED FOR FIRST-TIME)

### Step 1: Create Blank Macro-Enabled Workbook

1. Open Microsoft Excel
2. Create a new blank workbook
3. Press `Alt + F11` to open VBA Editor
4. File → Save As
5. Save as type: **"Excel Macro-Enabled Workbook (*.xlsm)"**
6. Name: **"ISA600_Bidvest_Scoping_Tool.xlsm"**
7. Location: Choose your preferred location

### Step 2: Import VBA Modules

**In VBA Editor (Alt + F11):**

1. File → Import File (or `Ctrl + M`)

2. Navigate to: `/home/user/Scopingtool/VBA_Modules/`

3. Import ALL 8 modules IN THIS EXACT ORDER:
   ```
   1. Mod1_MainController_Fixed.bas
   2. Mod2_TabProcessing.bas
   3. Mod3_DataExtraction_Fixed.bas
   4. Mod4_SegmentalMatching_Fixed.bas
   5. Mod5_ScopingEngine_Fixed.bas
   6. Mod6_DashboardGeneration_Fixed.bas
   7. Mod7_PowerBIExport.bas
   8. Mod8_Utilities.bas
   ```

4. **Verify:** In VBA Project Explorer (Ctrl + R), you should see:
   ```
   VBAProject (ISA600_Bidvest_Scoping_Tool.xlsm)
   ├── Microsoft Excel Objects
   │   └── ThisWorkbook
   ├── Modules
   │   ├── Mod1_MainController
   │   ├── Mod2_TabProcessing
   │   ├── Mod3_DataExtraction
   │   ├── Mod4_SegmentalMatching
   │   ├── Mod5_ScopingEngine
   │   ├── Mod6_DashboardGeneration
   │   ├── Mod7_PowerBIExport
   │   └── Mod8_Utilities
   ```

### Step 3: Add Button to Run Tool

**In Excel (Alt + F11 to return to worksheet):**

1. Insert a worksheet: Right-click → Insert → Worksheet
2. Rename to: **"Start Here"**
3. Insert Developer Tab (if not visible):
   - File → Options → Customize Ribbon → Check "Developer"

4. Developer → Insert → Button (Form Control)
5. Draw button on worksheet (size: ~3cm x 1.5cm)
6. When "Assign Macro" dialog appears:
   - Select: **"Mod1_MainController.StartBidvestScopingTool"**
   - Click OK

7. Right-click button → Edit Text
   - Change to: **"START SCOPING TOOL"**

### Step 4: Format Start Page

**Make it user-friendly:**

```
Cell A1: "ISA 600 REVISED - BIDVEST GROUP SCOPING TOOL"
         (Font: Arial 18pt, Bold, Blue)

Cell A3: "INSTRUCTIONS:"
Cell A4: "1. Ensure Stripe Packs workbook is open"
Cell A5: "2. Ensure Segmental Reporting workbook is open (optional)"
Cell A6: "3. Click the button below to start the tool"

Cell A8: [START SCOPING TOOL button]

Cell A10: "REQUIREMENTS:"
Cell A11: "- Microsoft Excel 2016 or later"
Cell A12: "- Macro security set to 'Enable all macros' (Developer → Macro Security)"
Cell A13: "- Both source workbooks must be open before running"

Cell A15: "DOCUMENTATION:"
Cell A16: "See COMPREHENSIVE_IMPLEMENTATION_GUIDE.md for detailed instructions"
Cell A17: "See CRITICAL_WORKFLOW_BUG_ANALYSIS.md for technical details"
Cell A18: "See TABLE_NAME_FIXES.md for dashboard fix details"
```

### Step 5: Enable Macros and Test

1. Tools → Macro → Security → **Enable all macros** (temporarily for testing)
   - **IMPORTANT:** Set back to normal after testing for security

2. Save workbook (Ctrl + S)

3. Close and reopen workbook

4. Click **"START SCOPING TOOL"** button to test

5. Verify welcome message appears:
   ```
   ISA 600 REVISED - BIDVEST GROUP SCOPING TOOL

   This comprehensive tool will:
   - Process Stripe Packs consolidation workbook
   - Process Segmental Reporting workbook
   ...
   ```

---

## METHOD 2: AUTOMATED IMPORT (VBScript)

### Using the Provided Import Script

**See: `Import_VBA_Modules.vbs`** (created separately)

1. Double-click `Import_VBA_Modules.vbs`
2. Follow prompts to select Excel workbook
3. Script automatically imports all 8 modules

**OR run from command line:**
```cmd
cscript Import_VBA_Modules.vbs
```

---

## METHOD 3: COMMAND LINE (VBA Automation)

### PowerShell Script

```powershell
# Save as: Import-VBAModules.ps1

$excelPath = "C:\Path\To\ISA600_Bidvest_Scoping_Tool.xlsm"
$modulesPath = "C:\Path\To\VBA_Modules\"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$workbook = $excel.Workbooks.Open($excelPath)
$vbaProject = $workbook.VBProject

# Import modules in order
$modules = @(
    "Mod1_MainController_Fixed.bas",
    "Mod2_TabProcessing.bas",
    "Mod3_DataExtraction_Fixed.bas",
    "Mod4_SegmentalMatching_Fixed.bas",
    "Mod5_ScopingEngine_Fixed.bas",
    "Mod6_DashboardGeneration_Fixed.bas",
    "Mod7_PowerBIExport.bas",
    "Mod8_Utilities.bas"
)

foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath $module
    $vbaProject.VBComponents.Import($modulePath)
    Write-Host "Imported: $module"
}

$workbook.Save()
$workbook.Close()
$excel.Quit()

Write-Host "All modules imported successfully!"
```

**Run:**
```powershell
powershell -ExecutionPolicy Bypass -File Import-VBAModules.ps1
```

---

## VERIFICATION CHECKLIST

After importing modules, verify:

### ✅ Module Verification

Run this VBA code in Immediate Window (Ctrl + G):
```vba
Sub VerifyModules()
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Debug.Print comp.Name & " - " & comp.Type
    Next comp
End Sub
```

**Expected Output:**
```
Mod1_MainController - 1 (Standard Module)
Mod2_TabProcessing - 1
Mod3_DataExtraction - 1
Mod4_SegmentalMatching - 1
Mod5_ScopingEngine - 1
Mod6_DashboardGeneration - 1
Mod7_PowerBIExport - 1
Mod8_Utilities - 1
```

### ✅ Compilation Check

1. In VBA Editor: **Debug → Compile VBAProject**
2. Should complete with **NO ERRORS**
3. If errors appear:
   - Check all 8 modules imported
   - Check module names match exactly
   - Ensure no duplicate modules

### ✅ Function Availability Check

Run in Immediate Window:
```vba
? TypeName(Mod1_MainController.g_StripePacksWorkbook)
```

**Expected:** `"Nothing"` (not error - means global variable exists)

---

## TROUBLESHOOTING

### Error: "Can't find project or library"

**Cause:** Missing reference

**Fix:**
1. VBA Editor → Tools → References
2. Ensure checked:
   - ✅ Visual Basic For Applications
   - ✅ Microsoft Excel 16.0 Object Library
   - ✅ OLE Automation
   - ✅ Microsoft Office 16.0 Object Library
   - ✅ Microsoft Scripting Runtime (for Dictionary)

3. If "Microsoft Scripting Runtime" is MISSING:
   - Browse → C:\Windows\System32\scrrun.dll
   - Select and click Open

### Error: "Programmatic access to VBA project is not trusted"

**Cause:** VBA project security settings

**Fix:**
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Check **"Trust access to the VBA project object model"**
3. Click OK
4. Restart Excel

### Error: "Module names must be unique"

**Cause:** Duplicate modules

**Fix:**
1. Check VBA Project Explorer for duplicates
2. Delete duplicate modules (right-click → Remove)
3. Re-import only once

### Macros Disabled When Opening

**Fix:**
1. File → Options → Trust Center → Trust Center Settings
2. Trusted Locations → Add Location
3. Browse to folder containing your .xlsm file
4. Check **"Subfolders of this location are also trusted"**
5. Click OK

---

## DISTRIBUTION

### Sharing the Macro Workbook

**Option 1: Share .xlsm file**
- Recipient must enable macros
- Recipient must have Excel 2016 or later

**Option 2: Share .bas modules + guide**
- Recipient imports modules following this guide
- More secure (they control the code)

**Option 3: Digital signature (recommended for enterprise)**
1. Obtain code signing certificate
2. VBA Editor → Tools → Digital Signature
3. Select certificate
4. Sign VBA project
5. Recipients will see verified signature

---

## SECURITY CONSIDERATIONS

### Before Deployment

1. **Code Review:** Have another developer review all 8 modules
2. **Test Environment:** Test thoroughly on non-production data
3. **Backup:** Always backup source workbooks before running
4. **User Training:** Train users on proper usage
5. **Documentation:** Ensure all documentation is complete

### Macro Security Settings

**For Development:**
- Enable all macros (temporarily)
- Trust access to VBA project object model

**For Production:**
- Disable all macros except digitally signed macros
- Use Trusted Locations for the .xlsm file
- Sign VBA project with digital certificate

---

## MAINTENANCE

### Updating Modules

**Method A: Replace entire module**
1. VBA Editor → Right-click module → Remove
2. File → Import File → Select updated .bas file

**Method B: Copy-paste code**
1. Open .bas file in text editor
2. Copy all code
3. VBA Editor → Double-click module → Paste

### Version Control

Keep track of versions:
```
ISA600_Bidvest_Scoping_Tool_v7.0.xlsm  (current)
ISA600_Bidvest_Scoping_Tool_v6.0.xlsm  (previous - for rollback)
```

---

## QUICK START SUMMARY

1. Open Excel → Create new workbook
2. Save as `ISA600_Bidvest_Scoping_Tool.xlsm`
3. Alt + F11 → Open VBA Editor
4. File → Import File → Import all 8 .bas modules
5. Create "Start Here" worksheet
6. Add button → Assign macro: StartBidvestScopingTool
7. Save, close, reopen
8. Enable macros
9. Click button to run!

---

## SUPPORT

**Issues? Check:**
1. CRITICAL_WORKFLOW_BUG_ANALYSIS.md - Workflow details
2. TABLE_NAME_FIXES.md - Dashboard issues
3. RUNTIME_ERROR_FIXES.md - Error fixes
4. MODULE_VERIFICATION_REPORT.md - Module details
5. COMPREHENSIVE_IMPLEMENTATION_GUIDE.md - Full guide

**Still stuck?**
- Check Immediate Window (Ctrl + G) for Debug.Print messages
- Enable error reporting: Tools → Options → General → Error Trapping: Break on All Errors
- Review error messages carefully

---

*End of Excel Import Guide*
