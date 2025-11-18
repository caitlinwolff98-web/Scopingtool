Attribute VB_Name = "Mod8_Utilities"
Option Explicit

' ============================================================================
' MODULE 8: UTILITIES & HELPERS
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.0 - Complete Overhaul
' ============================================================================
' PURPOSE: Centralized utility functions and configuration
' DESCRIPTION: Provides common functions, validation, error handling,
'              and configuration constants used across all modules
' ============================================================================

' ==================== VERSION INFORMATION ====================
Public Const TOOL_VERSION As String = "6.0"
Public Const TOOL_NAME As String = "Bidvest Group ISA 600 Scoping Tool"
Public Const TOOL_DATE As String = "2025-11"
Public Const TOOL_FULL_NAME As String = "ISA 600 Revised Component Scoping Tool - Complete Overhaul"

' ==================== ROW CONSTANTS ====================
Public Const ROW_CURRENCY_TYPE As Long = 6
Public Const ROW_PACK_NAME As Long = 7
Public Const ROW_PACK_CODE As Long = 8
Public Const ROW_FSLI_START As Long = 9

' ==================== WORKBOOK FUNCTIONS ====================
Public Function GetWorkbookByName(workbookName As String) As Workbook
    '------------------------------------------------------------------------
    ' Get workbook by name (case-insensitive, handles extensions)
    ' Returns Workbook object or Nothing if not found
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim wb As Workbook
    Dim nameWithoutExt As String
    Dim testName As String

    ' Try exact name first
    Set wb = Workbooks(workbookName)

    If wb Is Nothing Then
        ' Try without extension
        nameWithoutExt = workbookName
        nameWithoutExt = Replace(nameWithoutExt, ".xlsx", "")
        nameWithoutExt = Replace(nameWithoutExt, ".xlsm", "")
        nameWithoutExt = Replace(nameWithoutExt, ".xls", "")

        ' Try with different extensions
        Set wb = Workbooks(nameWithoutExt & ".xlsx")
        If wb Is Nothing Then Set wb = Workbooks(nameWithoutExt & ".xlsm")
        If wb Is Nothing Then Set wb = Workbooks(nameWithoutExt & ".xls")
        If wb Is Nothing Then Set wb = Workbooks(nameWithoutExt)
    End If

    Set GetWorkbookByName = wb
    On Error GoTo 0
End Function

Public Function WorkbookIsOpen(workbookName As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if a workbook is currently open
    '------------------------------------------------------------------------
    Dim wb As Workbook

    Set wb = GetWorkbookByName(workbookName)
    WorkbookIsOpen = Not (wb Is Nothing)
End Function

' ==================== STRING FUNCTIONS ====================
Public Function SafeTrim(value As Variant) As String
    '------------------------------------------------------------------------
    ' Safe string trim that handles null/empty values
    '------------------------------------------------------------------------
    If IsNull(value) Or IsEmpty(value) Then
        SafeTrim = ""
    ElseIf IsNumeric(value) Then
        SafeTrim = Trim(CStr(value))
    Else
        SafeTrim = Trim(CStr(value))
    End If
End Function

Public Function IsNullOrEmpty(value As Variant) As Boolean
    '------------------------------------------------------------------------
    ' Check if value is null, empty, or blank string
    '------------------------------------------------------------------------
    IsNullOrEmpty = (IsNull(value) Or IsEmpty(value) Or Trim(CStr(value)) = "")
End Function

Public Function RemoveSpecialCharacters(inputStr As String) As String
    '------------------------------------------------------------------------
    ' Remove special characters from string (keep only alphanumeric, space, dash, underscore)
    '------------------------------------------------------------------------
    Dim i As Long
    Dim char As String
    Dim result As String

    result = ""

    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)

        If (char >= "A" And char <= "Z") Or _
           (char >= "a" And char <= "z") Or _
           (char >= "0" And char <= "9") Or _
           char = " " Or char = "-" Or char = "_" Then
            result = result & char
        End If
    Next i

    RemoveSpecialCharacters = result
End Function

' ==================== NUMBER FUNCTIONS ====================
Public Function IsValidNumber(value As Variant) As Boolean
    '------------------------------------------------------------------------
    ' Check if value is a valid number (not null, empty, or error)
    '------------------------------------------------------------------------
    IsValidNumber = False

    If Not IsNull(value) And Not IsEmpty(value) And Not IsError(value) Then
        If IsNumeric(value) Then
            IsValidNumber = True
        End If
    End If
End Function

Public Function SafeNumeric(value As Variant, Optional defaultValue As Double = 0) As Double
    '------------------------------------------------------------------------
    ' Safely convert value to number with fallback default
    '------------------------------------------------------------------------
    If IsValidNumber(value) Then
        SafeNumeric = CDbl(value)
    Else
        SafeNumeric = defaultValue
    End If
End Function

Public Function FormatCurrency(value As Variant) As String
    '------------------------------------------------------------------------
    ' Format currency value for display
    '------------------------------------------------------------------------
    If IsValidNumber(value) Then
        FormatCurrency = Format(CDbl(value), "#,##0.00")
    Else
        FormatCurrency = "N/A"
    End If
End Function

' ==================== DATE/TIME FUNCTIONS ====================
Public Function GetTimestamp() As String
    '------------------------------------------------------------------------
    ' Get current timestamp in format: YYYY-MM-DD HH-MM-SS
    ' Used for filename generation
    '------------------------------------------------------------------------
    GetTimestamp = Format(Now, "yyyy-mm-dd hh-mm-ss")
End Function

Public Function GetDateStamp() As String
    '------------------------------------------------------------------------
    ' Get current date stamp in format: YYYY-MM-DD
    '------------------------------------------------------------------------
    GetDateStamp = Format(Now, "yyyy-mm-dd")
End Function

' ==================== ERROR HANDLING ====================
Public Sub ShowError(errorTitle As String, errorMessage As String, Optional errorNumber As Long = 0)
    '------------------------------------------------------------------------
    ' Display formatted error message
    '------------------------------------------------------------------------
    Dim msg As String

    msg = errorMessage

    If errorNumber <> 0 Then
        msg = msg & vbCrLf & vbCrLf & "Error Number: " & errorNumber
    End If

    MsgBox msg, vbCritical, errorTitle
End Sub

Public Sub ShowInfo(infoTitle As String, infoMessage As String)
    '------------------------------------------------------------------------
    ' Display formatted information message
    '------------------------------------------------------------------------
    MsgBox infoMessage, vbInformation, infoTitle
End Sub

Public Function ShowConfirmation(confirmTitle As String, confirmMessage As String) As Boolean
    '------------------------------------------------------------------------
    ' Display confirmation dialog and return user choice
    '------------------------------------------------------------------------
    ShowConfirmation = (MsgBox(confirmMessage, vbYesNo + vbQuestion, confirmTitle) = vbYes)
End Function

' ==================== VALIDATION FUNCTIONS ====================
Public Function ValidateWorksheetExists(wb As Workbook, sheetName As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if a worksheet exists in a workbook
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)

    ValidateWorksheetExists = Not (ws Is Nothing)

    On Error GoTo 0
End Function

Public Function ValidateRowStructure(ws As Worksheet) As Boolean
    '------------------------------------------------------------------------
    ' Validate that worksheet has expected row structure (Rows 6-8)
    '------------------------------------------------------------------------
    On Error Resume Next

    ValidateRowStructure = False

    If ws Is Nothing Then Exit Function

    ' Check Row 6 (Currency Type)
    If ws.Cells(ROW_CURRENCY_TYPE, 3).Value = "" Then Exit Function

    ' Check Row 7 (Pack Name)
    If ws.Cells(ROW_PACK_NAME, 3).Value = "" Then Exit Function

    ' Check Row 8 (Pack Code)
    If ws.Cells(ROW_PACK_CODE, 3).Value = "" Then Exit Function

    ValidateRowStructure = True

    On Error GoTo 0
End Function

' ==================== FORMATTING FUNCTIONS ====================
Public Sub ApplyHeaderFormatting(headerRange As Range)
    '------------------------------------------------------------------------
    ' Apply standard header formatting to a range
    '------------------------------------------------------------------------
    With headerRange
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub ApplyTableFormatting(ws As Worksheet, startRow As Long, startCol As Long, endRow As Long, endCol As Long)
    '------------------------------------------------------------------------
    ' Apply table formatting to a range
    '------------------------------------------------------------------------
    Dim tbl As ListObject
    Dim tableRange As Range

    On Error Resume Next

    Set tableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))

    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)

    If Not tbl Is Nothing Then
        tbl.TableStyle = "TableStyleMedium2"
    End If

    On Error GoTo 0
End Sub

Public Sub FreezePanesAt(ws As Worksheet, row As Long, col As Long)
    '------------------------------------------------------------------------
    ' Freeze panes at specified row and column
    '------------------------------------------------------------------------
    On Error Resume Next

    ws.Activate
    ws.Cells(row, col).Select
    ActiveWindow.FreezePanes = True

    On Error GoTo 0
End Sub

' ==================== DICTIONARY HELPERS ====================
Public Function CreateDictionary() As Object
    '------------------------------------------------------------------------
    ' Create new Scripting.Dictionary object
    '------------------------------------------------------------------------
    On Error Resume Next

    Set CreateDictionary = CreateObject("Scripting.Dictionary")

    If Err.Number <> 0 Then
        ShowError "Missing Library", "Microsoft Scripting Runtime is not available. Please enable it in VBA References."
        Set CreateDictionary = Nothing
    End If

    On Error GoTo 0
End Function

Public Function DictionaryToArray(dict As Object) As Variant()
    '------------------------------------------------------------------------
    ' Convert dictionary keys to array
    '------------------------------------------------------------------------
    Dim result() As Variant
    Dim i As Long
    Dim key As Variant

    If dict.Count = 0 Then
        DictionaryToArray = result
        Exit Function
    End If

    ReDim result(0 To dict.Count - 1)

    i = 0
    For Each key In dict.Keys
        result(i) = key
        i = i + 1
    Next key

    DictionaryToArray = result
End Function

' ==================== COLLECTION HELPERS ====================
Public Function CollectionContains(coll As Collection, value As Variant) As Boolean
    '------------------------------------------------------------------------
    ' Check if collection contains a value
    '------------------------------------------------------------------------
    Dim item As Variant

    CollectionContains = False

    For Each item In coll
        If item = value Then
            CollectionContains = True
            Exit Function
        End If
    Next item
End Function

Public Function CollectionToArray(coll As Collection) As Variant()
    '------------------------------------------------------------------------
    ' Convert collection to array
    '------------------------------------------------------------------------
    Dim result() As Variant
    Dim i As Long

    If coll.Count = 0 Then
        CollectionToArray = result
        Exit Function
    End If

    ReDim result(0 To coll.Count - 1)

    For i = 1 To coll.Count
        result(i - 1) = coll(i)
    Next i

    CollectionToArray = result
End Function

' ==================== PROGRESS INDICATORS ====================
Public Sub UpdateStatusBar(message As String)
    '------------------------------------------------------------------------
    ' Update Excel status bar with progress message
    '------------------------------------------------------------------------
    Application.StatusBar = message
    DoEvents
End Sub

Public Sub ClearStatusBar()
    '------------------------------------------------------------------------
    ' Clear Excel status bar
    '------------------------------------------------------------------------
    Application.StatusBar = False
End Sub

' ==================== FILE SYSTEM HELPERS ====================
Public Function GetDesktopPath() As String
    '------------------------------------------------------------------------
    ' Get user's Desktop folder path
    '------------------------------------------------------------------------
    GetDesktopPath = Environ("USERPROFILE") & "\Desktop"
End Function

Public Function GetDocumentsPath() As String
    '------------------------------------------------------------------------
    ' Get user's Documents folder path
    '------------------------------------------------------------------------
    GetDocumentsPath = Environ("USERPROFILE") & "\Documents"
End Function

Public Function FileExists(filePath As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if a file exists
    '------------------------------------------------------------------------
    FileExists = (Dir(filePath) <> "")
End Function

' ==================== VERSION INFO ====================
Public Function GetToolVersionInfo() As String
    '------------------------------------------------------------------------
    ' Get complete tool version information
    '------------------------------------------------------------------------
    GetToolVersionInfo = TOOL_NAME & " v" & TOOL_VERSION & " (" & TOOL_DATE & ")"
End Function

Public Sub DisplayAboutDialog()
    '------------------------------------------------------------------------
    ' Display "About" dialog with tool information
    '------------------------------------------------------------------------
    Dim msg As String

    msg = TOOL_FULL_NAME & vbCrLf & _
          "Version: " & TOOL_VERSION & vbCrLf & _
          "Release Date: " & TOOL_DATE & vbCrLf & vbCrLf & _
          "Designed for Bidvest Group Limited" & vbCrLf & _
          "ISA 600 Revised Compliance" & vbCrLf & vbCrLf & _
          "This tool automates component scoping for group audits," & vbCrLf & _
          "processes consolidation data, and generates interactive" & vbCrLf & _
          "dashboards and Power BI-ready datasets."

    MsgBox msg, vbInformation, "About " & TOOL_NAME
End Sub

' ==================== DEBUG HELPERS ====================
Public Sub DebugPrint(message As String)
    '------------------------------------------------------------------------
    ' Print debug message to Immediate Window
    '------------------------------------------------------------------------
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & message
End Sub

Public Sub DebugPrintDictionary(dict As Object, dictName As String)
    '------------------------------------------------------------------------
    ' Print dictionary contents to Immediate Window (for debugging)
    '------------------------------------------------------------------------
    Dim key As Variant

    Debug.Print "=== " & dictName & " (Count: " & dict.Count & ") ==="

    For Each key In dict.Keys
        Debug.Print "  " & key & " -> " & dict(key)
    Next key

    Debug.Print "========================================"
End Sub

Public Sub DebugPrintCollection(coll As Collection, collName As String)
    '------------------------------------------------------------------------
    ' Print collection contents to Immediate Window (for debugging)
    '------------------------------------------------------------------------
    Dim item As Variant
    Dim i As Long

    Debug.Print "=== " & collName & " (Count: " & coll.Count & ") ==="

    i = 1
    For Each item In coll
        Debug.Print "  " & i & ": " & item
        i = i + 1
    Next item

    Debug.Print "========================================"
End Sub

' ==================== PERFORMANCE HELPERS ====================
Public Sub OptimizePerformance(enable As Boolean)
    '------------------------------------------------------------------------
    ' Enable/disable Excel performance optimizations
    '------------------------------------------------------------------------
    If enable Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
End Sub
