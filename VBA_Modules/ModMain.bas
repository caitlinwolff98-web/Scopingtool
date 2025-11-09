Option Explicit

' ============================================================================
' MODULE: ModMain
' PURPOSE: Main entry point for the TGK Consolidation Scoping Tool
' DESCRIPTION: Orchestrates the entire process of analyzing TGK consolidation
'              workbooks, categorizing tabs, and creating structured tables
'              for Power BI integration
' ============================================================================

' Global variables for workbook references
Public g_SourceWorkbook As Workbook
Public g_OutputWorkbook As Workbook
Public g_TabCategories As Object ' Dictionary for tab categorization

' Main entry point - called when user clicks the button
Public Sub StartScopingTool()
    On Error GoTo ErrorHandler
    
    Dim workbookName As String
    Dim result As VbMsgBoxResult
    
    ' Display welcome message
    result = MsgBox("Welcome to the TGK Consolidation Scoping Tool!" & vbCrLf & vbCrLf & _
                    "This tool will:" & vbCrLf & _
                    "1. Analyze your TGK consolidation workbook" & vbCrLf & _
                    "2. Categorize tabs for processing" & vbCrLf & _
                    "3. Create structured tables for Power BI" & vbCrLf & _
                    "4. Perform mathematical accuracy checks" & vbCrLf & vbCrLf & _
                    "Click OK to continue or Cancel to exit.", _
                    vbOKCancel + vbInformation, "TGK Scoping Tool")
    
    If result = vbCancel Then Exit Sub
    
    ' Step 1: Get the workbook name from user
    workbookName = GetWorkbookName()
    If workbookName = "" Then
        MsgBox "No workbook name provided. Process cancelled.", vbExclamation
        Exit Sub
    End If
    
    ' Step 2: Validate and set source workbook
    If Not SetSourceWorkbook(workbookName) Then
        MsgBox "Could not find workbook '" & workbookName & "'. Please ensure it is open.", vbCritical
        Exit Sub
    End If
    
    ' Step 3: Discover and list all tabs
    Dim tabList As Collection
    Set tabList = DiscoverTabs()
    
    If tabList.count = 0 Then
        MsgBox "No tabs found in the workbook.", vbExclamation
        Exit Sub
    End If
    
    ' Step 4: Categorize tabs
    If Not CategorizeTabs(tabList) Then
        MsgBox "Tab categorization was cancelled. Process terminated.", vbInformation
        Exit Sub
    End If
    
    ' Step 5: Validate required categories
    If Not ValidateCategories() Then
        MsgBox "Required tabs are missing. Please ensure all mandatory categories are assigned.", vbCritical
        Exit Sub
    End If
    
    ' Step 6: Create output workbook for tables
    CreateOutputWorkbook
    
    ' Step 7: Process data and create tables
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ProcessConsolidationData
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Step 8: Display completion message
    MsgBox "Scoping tool completed successfully!" & vbCrLf & vbCrLf & _
           "Tables have been created in: " & g_OutputWorkbook.Name, _
           vbInformation, "Process Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Error"
End Sub

' Get workbook name from user
Private Function GetWorkbookName() As String
    Dim userInput As String
    
    userInput = InputBox( _
        "Please enter the exact name of the TGK consolidation workbook." & vbCrLf & vbCrLf & _
        "Instructions:" & vbCrLf & _
        "1. Open the consolidation workbook" & vbCrLf & _
        "2. Copy the workbook name from the title bar" & vbCrLf & _
        "3. Paste it below (include .xlsx or .xlsm extension)", _
        "Enter Workbook Name", _
        "")
    
    GetWorkbookName = Trim(userInput)
End Function

' Set the source workbook reference
Private Function SetSourceWorkbook(workbookName As String) As Boolean
    On Error Resume Next
    
    ' Try to find workbook by exact name
    Set g_SourceWorkbook = Workbooks(workbookName)
    
    If g_SourceWorkbook Is Nothing Then
        ' Try without extension
        Dim nameWithoutExt As String
        nameWithoutExt = Replace(Replace(workbookName, ".xlsx", ""), ".xlsm", "")
        Set g_SourceWorkbook = Workbooks(nameWithoutExt & ".xlsx")
        
        If g_SourceWorkbook Is Nothing Then
            Set g_SourceWorkbook = Workbooks(nameWithoutExt & ".xlsm")
        End If
    End If
    
    SetSourceWorkbook = Not (g_SourceWorkbook Is Nothing)
    On Error GoTo 0
End Function

' Discover all tabs in the source workbook
Private Function DiscoverTabs() As Collection
    Dim tabs As New Collection
    Dim ws As Worksheet
    
    For Each ws In g_SourceWorkbook.Worksheets
        tabs.Add ws.Name
    Next ws
    
    Set DiscoverTabs = tabs
End Function

' Create the output workbook for generated tables
Private Sub CreateOutputWorkbook()
    Set g_OutputWorkbook = Workbooks.Add
    g_OutputWorkbook.Worksheets(1).Name = "Control Panel"
    
    ' Add informational sheet
    With g_OutputWorkbook.Worksheets("Control Panel")
        .Range("A1").Value = "TGK Scoping Tool - Output Tables"
        .Range("A2").Value = "Source: " & g_SourceWorkbook.Name
        .Range("A3").Value = "Generated: " & Now()
        .Range("A1:A3").Font.Bold = True
    End With
End Sub
