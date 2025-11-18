'========================================================================
' VBA MODULE IMPORT AUTOMATION SCRIPT
' ISA 600 Bidvest Scoping Tool
'
' PURPOSE: Automatically import all 8 VBA modules into an Excel workbook
'
' USAGE:
'   Method 1: Double-click this file
'   Method 2: cscript Import_VBA_Modules.vbs
'   Method 3: wscript Import_VBA_Modules.vbs
'
' REQUIREMENTS:
'   - Microsoft Excel installed
'   - Trust access to VBA project object model enabled
'     (File → Options → Trust Center → Macro Settings)
'
' AUTHOR: ISA 600 Scoping Tool Team
' DATE: 2025-11-18
'========================================================================

Option Explicit

Dim objExcel, objWorkbook, objVBProject
Dim strWorkbookPath, strModulesPath
Dim arrModules, strModule
Dim objFSO, objShell
Dim intCount, intSuccess, intFailed

' Initialize
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

intSuccess = 0
intFailed = 0

' ===== STEP 1: Get script directory =====
strModulesPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\VBA_Modules\"

If Not objFSO.FolderExists(strModulesPath) Then
    MsgBox "ERROR: VBA_Modules folder not found!" & vbCrLf & vbCrLf & _
           "Expected location: " & strModulesPath & vbCrLf & vbCrLf & _
           "Please ensure this script is in the same folder as VBA_Modules/", _
           vbCritical, "Folder Not Found"
    WScript.Quit
End If

' ===== STEP 2: Prompt for Excel workbook =====
strWorkbookPath = InputBox( _
    "EXCEL WORKBOOK SELECTION" & vbCrLf & vbCrLf & _
    "Enter the FULL PATH to the Excel workbook where modules should be imported:" & vbCrLf & vbCrLf & _
    "Examples:" & vbCrLf & _
    "  C:\Users\YourName\Documents\ISA600_Tool.xlsm" & vbCrLf & _
    "  C:\Projects\Bidvest\Scoping_Tool.xlsm" & vbCrLf & vbCrLf & _
    "Or leave blank to create a NEW workbook in the current directory.", _
    "Select Excel Workbook", "")

If strWorkbookPath = "" Then
    ' Create new workbook
    strWorkbookPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ISA600_Bidvest_Scoping_Tool.xlsm"

    Dim intResult
    intResult = MsgBox( _
        "Create new workbook at:" & vbCrLf & vbCrLf & _
        strWorkbookPath & vbCrLf & vbCrLf & _
        "Click YES to create, NO to cancel.", _
        vbYesNo + vbQuestion, "Create New Workbook")

    If intResult <> vbYes Then
        WScript.Quit
    End If
End If

' ===== STEP 3: Define modules to import (IN ORDER) =====
arrModules = Array( _
    "Mod1_MainController_Fixed.bas", _
    "Mod2_TabProcessing.bas", _
    "Mod3_DataExtraction_Fixed.bas", _
    "Mod4_SegmentalMatching_Fixed.bas", _
    "Mod5_ScopingEngine_Fixed.bas", _
    "Mod6_DashboardGeneration_Fixed.bas", _
    "Mod7_PowerBIExport.bas", _
    "Mod8_Utilities.bas" _
)

' ===== STEP 4: Create Excel instance =====
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "ERROR: Could not create Excel application!" & vbCrLf & vbCrLf & _
           "Please ensure Microsoft Excel is installed.", _
           vbCritical, "Excel Not Found"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

' ===== STEP 5: Open or create workbook =====
On Error Resume Next

If objFSO.FileExists(strWorkbookPath) Then
    ' Open existing workbook
    Set objWorkbook = objExcel.Workbooks.Open(strWorkbookPath)
    If Err.Number <> 0 Then
        MsgBox "ERROR: Could not open workbook!" & vbCrLf & vbCrLf & _
               "Path: " & strWorkbookPath & vbCrLf & _
               "Error: " & Err.Description, _
               vbCritical, "Workbook Open Failed"
        objExcel.Quit
        WScript.Quit
    End If
Else
    ' Create new workbook
    Set objWorkbook = objExcel.Workbooks.Add
    objWorkbook.SaveAs strWorkbookPath, 52 ' xlOpenXMLWorkbookMacroEnabled = 52
    If Err.Number <> 0 Then
        MsgBox "ERROR: Could not create workbook!" & vbCrLf & vbCrLf & _
               "Path: " & strWorkbookPath & vbCrLf & _
               "Error: " & Err.Description, _
               vbCritical, "Workbook Create Failed"
        objExcel.Quit
        WScript.Quit
    End If
End If

On Error GoTo 0

' ===== STEP 6: Get VBA project =====
On Error Resume Next
Set objVBProject = objWorkbook.VBProject
If Err.Number <> 0 Then
    MsgBox "ERROR: Cannot access VBA project!" & vbCrLf & vbCrLf & _
           "You must enable 'Trust access to the VBA project object model':" & vbCrLf & vbCrLf & _
           "1. Open Excel" & vbCrLf & _
           "2. File → Options → Trust Center → Trust Center Settings" & vbCrLf & _
           "3. Macro Settings → Check 'Trust access to VBA project object model'" & vbCrLf & _
           "4. Click OK and restart Excel" & vbCrLf & vbCrLf & _
           "Then run this script again.", _
           vbCritical, "VBA Project Access Denied"
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit
End If
On Error GoTo 0

' ===== STEP 7: Import modules =====
WScript.Echo "Starting VBA module import..." & vbCrLf & _
             "Workbook: " & strWorkbookPath & vbCrLf & _
             "Modules: " & UBound(arrModules) + 1 & vbCrLf

For intCount = 0 To UBound(arrModules)
    strModule = arrModules(intCount)
    Dim strFullPath
    strFullPath = strModulesPath & strModule

    If objFSO.FileExists(strFullPath) Then
        On Error Resume Next
        objVBProject.VBComponents.Import strFullPath
        If Err.Number = 0 Then
            WScript.Echo "[OK] Imported: " & strModule
            intSuccess = intSuccess + 1
        Else
            WScript.Echo "[FAIL] Error importing: " & strModule & " - " & Err.Description
            intFailed = intFailed + 1
        End If
        On Error GoTo 0
    Else
        WScript.Echo "[FAIL] File not found: " & strModule
        intFailed = intFailed + 1
    End If
Next

' ===== STEP 8: Save and close =====
objWorkbook.Save
objWorkbook.Close
objExcel.Quit

' ===== STEP 9: Display results =====
If intFailed = 0 Then
    MsgBox "MODULE IMPORT COMPLETED SUCCESSFULLY!" & vbCrLf & vbCrLf & _
           "Workbook: " & strWorkbookPath & vbCrLf & _
           "Modules imported: " & intSuccess & "/" & (UBound(arrModules) + 1) & vbCrLf & vbCrLf & _
           "NEXT STEPS:" & vbCrLf & _
           "1. Open the workbook in Excel" & vbCrLf & _
           "2. Enable macros when prompted" & vbCrLf & _
           "3. Press Alt+F11 to verify modules imported" & vbCrLf & _
           "4. Run Mod1_MainController.StartBidvestScopingTool" & vbCrLf & vbCrLf & _
           "See EXCEL_IMPORT_GUIDE.md for detailed instructions.", _
           vbInformation, "Import Complete"
Else
    MsgBox "MODULE IMPORT COMPLETED WITH ERRORS!" & vbCrLf & vbCrLf & _
           "Workbook: " & strWorkbookPath & vbCrLf & _
           "Successful: " & intSuccess & vbCrLf & _
           "Failed: " & intFailed & vbCrLf & vbCrLf & _
           "Please check:" & vbCrLf & _
           "1. All .bas files exist in VBA_Modules/ folder" & vbCrLf & _
           "2. File names match exactly (case-sensitive)" & vbCrLf & _
           "3. Files are not corrupted" & vbCrLf & vbCrLf & _
           "See EXCEL_IMPORT_GUIDE.md for troubleshooting.", _
           vbExclamation, "Import Errors"
End If

' Cleanup
Set objWorkbook = Nothing
Set objVBProject = Nothing
Set objExcel = Nothing
Set objFSO = Nothing
Set objShell = Nothing

WScript.Echo vbCrLf & "Script completed. Press any key to exit."
WScript.Quit
