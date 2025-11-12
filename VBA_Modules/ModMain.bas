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
    
    ' Step 7: Configure threshold-based scoping (optional)
    Dim thresholds As Collection
    Dim scopedPacks As Object
    Dim applyThresholds As VbMsgBoxResult
    
    applyThresholds = MsgBox("Would you like to configure threshold-based automatic scoping?" & vbCrLf & vbCrLf & _
                             "This will allow you to:" & vbCrLf & _
                             "- Select specific FSLIs for threshold analysis" & vbCrLf & _
                             "- Set threshold values for each FSLI" & vbCrLf & _
                             "- Automatically mark packs as 'Scoped In' based on thresholds" & vbCrLf & vbCrLf & _
                             "Click YES to configure thresholds, NO to skip.", _
                             vbYesNo + vbQuestion, "Threshold-Based Scoping")
    
    If applyThresholds = vbYes Then
        Set thresholds = ModThresholdScoping.ConfigureAndApplyThresholds()
        If thresholds.Count > 0 Then
            Set scopedPacks = ModThresholdScoping.ApplyThresholdsToData(thresholds)
        End If
    End If
    
    ' Step 8: Process data and create tables
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ProcessConsolidationData
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Step 9: Create threshold configuration sheet if thresholds were applied
    If applyThresholds = vbYes And thresholds.Count > 0 Then
        ModThresholdScoping.CreateThresholdConfigSheet thresholds, scopedPacks
    End If
    
    ' Step 10: Create scoping summary sheet
    Application.StatusBar = "Creating scoping summary..."
    CreateScopingSummarySheet scopedPacks
    Application.StatusBar = False
    
    ' Step 11: Create interactive Excel dashboard
    Application.StatusBar = "Creating interactive dashboard..."
    ModInteractiveDashboard.CreateInteractiveDashboard
    ModInteractiveDashboard.AddInteractiveFilters
    ModInteractiveDashboard.CreateScopingCalculator
    Application.StatusBar = False
    
    ' Step 12: Create Power BI integration assets
    Application.StatusBar = "Creating Power BI integration assets..."
    ModPowerBIIntegration.CreateAllPowerBIAssets
    Application.StatusBar = False
    
    ' Step 13: Save the output workbook with standardized name
    SaveOutputWorkbook
    
    ' Step 14: Display completion message
    MsgBox "Scoping tool completed successfully!" & vbCrLf & vbCrLf & _
           "Output saved as: " & g_OutputWorkbook.Name & vbCrLf & _
           "Location: " & g_OutputWorkbook.Path & vbCrLf & vbCrLf & _
           "Generated assets:" & vbCrLf & _
           "- Data tables for analysis" & vbCrLf & _
           "- Threshold configuration (if applied)" & vbCrLf & _
           "- Scoping summary with recommendations" & vbCrLf & _
           "- Interactive Excel dashboard" & vbCrLf & _
           "- Scoping calculator" & vbCrLf & _
           "- Power BI integration metadata" & vbCrLf & vbCrLf & _
           "The workbook can be used standalone or with Power BI!" & vbCrLf & _
           "See POWERBI_INTEGRATION_GUIDE.md for next steps.", _
           vbInformation, "Process Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Error"
End Sub

' Save output workbook with standardized name
Private Sub SaveOutputWorkbook()
    On Error GoTo ErrorHandler
    
    Dim savePath As String
    Dim fileName As String
    
    ' Standard output file name
    fileName = "Bidvest Scoping Tool Output.xlsx"
    
    ' Use the same directory as the source workbook
    savePath = g_SourceWorkbook.Path & Application.PathSeparator & fileName
    
    ' Save the workbook
    Application.DisplayAlerts = False
    g_OutputWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    ' If save fails, just leave it unsaved for user to manually save
    Debug.Print "Could not auto-save output workbook: " & Err.Description
End Sub

' Create scoping summary sheet with recommendations
Private Sub CreateScopingSummarySheet(scopedPacks As Object)
    On Error GoTo ErrorHandler
    
    Dim summaryWs As Worksheet
    Dim row As Long
    Dim inputTab As Worksheet
    Dim packDict As Object
    Dim col As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim packName As String
    Dim isScopedIn As String
    Dim suggestedForScope As String
    Dim packKey As Variant
    
    ' Check if sheet already exists
    On Error Resume Next
    Set summaryWs = g_OutputWorkbook.Worksheets("Scoping Summary")
    On Error GoTo ErrorHandler
    
    If summaryWs Is Nothing Then
        Set summaryWs = g_OutputWorkbook.Worksheets.Add
        summaryWs.Name = "Scoping Summary"
    Else
        summaryWs.Cells.Clear
    End If
    
    ' Get input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then Exit Sub
    
    ' Create pack dictionary
    Set packDict = CreateObject("Scripting.Dictionary")
    
    ' Collect all packs from input tab
    lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
    For col = 3 To lastCol
        packCode = Trim(inputTab.Cells(8, col).Value)
        packName = Trim(inputTab.Cells(7, col).Value)
        
        If packCode <> "" And packName <> "" Then
            If Not packDict.Exists(packCode) Then
                Dim packInfo As Object
                Set packInfo = CreateObject("Scripting.Dictionary")
                packInfo("Name") = packName
                packInfo("Code") = packCode
                
                ' Check if scoped in by threshold
                If Not scopedPacks Is Nothing Then
                    If scopedPacks.Exists(packCode) Then
                        packInfo("ScopedIn") = "Yes (Threshold)"
                        packInfo("Suggested") = "Yes"
                    Else
                        packInfo("ScopedIn") = "No"
                        packInfo("Suggested") = "Review Required"
                    End If
                Else
                    packInfo("ScopedIn") = "Not Yet Determined"
                    packInfo("Suggested") = "Review Required"
                End If
                
                packDict.Add packCode, packInfo
            End If
        End If
    Next col
    
    ' Write header
    row = 1
    With summaryWs
        .Cells(row, 1).Value = "SCOPING SUMMARY"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        row = row + 2
        
        ' Column headers
        .Cells(row, 1).Value = "Pack Code"
        .Cells(row, 2).Value = "Pack Name"
        .Cells(row, 3).Value = "Scoped In"
        .Cells(row, 4).Value = "Suggested for Scope"
        .Range("A" & row & ":D" & row).Font.Bold = True
        .Range("A" & row & ":D" & row).Interior.Color = RGB(68, 114, 196)
        .Range("A" & row & ":D" & row).Font.Color = RGB(255, 255, 255)
        row = row + 1
        
        ' Write pack data
        For Each packKey In packDict.Keys
            Set packInfo = packDict(packKey)
            .Cells(row, 1).Value = packInfo("Code")
            .Cells(row, 2).Value = packInfo("Name")
            .Cells(row, 3).Value = packInfo("ScopedIn")
            .Cells(row, 4).Value = packInfo("Suggested")
            
            ' Color code the Suggested column
            If packInfo("Suggested") = "Yes" Then
                .Cells(row, 4).Interior.Color = RGB(198, 239, 206) ' Light green
            Else
                .Cells(row, 4).Interior.Color = RGB(255, 235, 156) ' Light yellow
            End If
            
            row = row + 1
        Next packKey
        
        ' Summary statistics
        row = row + 2
        .Cells(row, 1).Value = "SUMMARY STATISTICS"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "Total Packs:"
        .Cells(row, 2).Value = packDict.Count
        row = row + 1
        
        If Not scopedPacks Is Nothing Then
            .Cells(row, 1).Value = "Automatically Scoped In:"
            .Cells(row, 2).Value = scopedPacks.Count
            row = row + 1
            
            .Cells(row, 1).Value = "Requiring Review:"
            .Cells(row, 2).Value = packDict.Count - scopedPacks.Count
        Else
            .Cells(row, 1).Value = "Threshold-Based Scoping:"
            .Cells(row, 2).Value = "Not Applied"
        End If
        
        ' Auto-fit columns
        .Columns.AutoFit
        
        ' Create table
        Dim lastDataRow As Long
        lastDataRow = 3 + packDict.Count ' Header rows + data rows
        
        On Error Resume Next
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(3, 1), .Cells(lastDataRow, 4)), , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = "Scoping_Summary_Table"
            tbl.TableStyle = "TableStyleMedium2"
        End If
        On Error GoTo ErrorHandler
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error creating scoping summary: " & Err.Description
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
    
    ' Use centralized function from ModConfig
    Set g_SourceWorkbook = ModConfig.GetWorkbookByName(workbookName)
    
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
