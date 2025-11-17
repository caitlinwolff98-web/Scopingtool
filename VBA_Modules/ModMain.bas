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
Public g_ConsolidatedPackCode As String ' Pack code for consolidated entity (excluded from scoping)
Public g_ConsolidatedPackName As String ' Pack name for consolidated entity

' Main entry point - called when user clicks the button
Public Sub StartScopingTool()
    On Error GoTo ErrorHandler
    
    Dim workbookName As String
    Dim result As VbMsgBoxResult
    
    ' Display welcome message
    result = MsgBox("Welcome to the TGK Consolidation Scoping Tool v5.0!" & vbCrLf & vbCrLf & _
                    "This tool will:" & vbCrLf & _
                    "1. Analyze your TGK consolidation workbook" & vbCrLf & _
                    "2. Categorize tabs for processing" & vbCrLf & _
                    "3. Create structured tables for Power BI" & vbCrLf & _
                    "4. Process IAS 8 segment reporting (optional)" & vbCrLf & _
                    "5. Generate scoping analysis and recommendations" & vbCrLf & vbCrLf & _
                    "Click OK to continue or Cancel to exit.", _
                    vbOKCancel + vbInformation, "TGK Scoping Tool v5.0")
    
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
    
    ' Step 5a: Select consolidated entity (to exclude from scoping)
    If Not SelectConsolidatedEntity() Then
        MsgBox "Consolidated entity selection was cancelled. Process terminated.", vbInformation
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
    
    ' Step 10a: Create Division-based reporting sheets
    Application.StatusBar = "Creating division-based reports..."
    CreateDivisionScopingReports scopedPacks
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

    ' Step 12a: Process IAS 8 Segment Reporting Document (optional) - NEW in v5.0
    Application.StatusBar = "Checking for segment reporting document..."
    Dim segmentProcessed As Boolean
    segmentProcessed = ModSegmentAnalysis.ProcessSegmentDocument()
    Application.StatusBar = False

    ' Step 13: Save the output workbook with standardized name
    SaveOutputWorkbook
    
    ' Step 14: Display completion message
    Dim completionMsg As String
    completionMsg = "Scoping tool v5.0 completed successfully!" & vbCrLf & vbCrLf & _
                   "Output saved as: " & g_OutputWorkbook.Name & vbCrLf & _
                   "Location: " & g_OutputWorkbook.Path & vbCrLf & vbCrLf & _
                   "Generated assets:" & vbCrLf & _
                   "- Data tables for analysis" & vbCrLf & _
                   "- Threshold configuration (if applied)" & vbCrLf & _
                   "- Scoping summary with recommendations" & vbCrLf & _
                   "- Division-based scoping reports" & vbCrLf & _
                   "- Scoped In Packs Detail" & vbCrLf & _
                   "- Interactive Excel dashboard" & vbCrLf & _
                   "- Scoping calculator" & vbCrLf & _
                   "- Power BI integration metadata" & vbCrLf

    ' Add segment tables message if processed
    If segmentProcessed Then
        completionMsg = completionMsg & "- IAS 8 Segment Pack Mapping (NEW)" & vbCrLf & _
                       "- IAS 8 Segment Summary (NEW)" & vbCrLf
    End If

    completionMsg = completionMsg & vbCrLf & _
                   "The workbook can be used standalone or with Power BI!" & vbCrLf & _
                   "See IMPLEMENTATION_GUIDE.md for next steps."

    MsgBox completionMsg, vbInformation, "Process Complete - v5.0"
    
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
    
    ' Add professional informational sheet
    With g_OutputWorkbook.Worksheets("Control Panel")
        ' Title
        .Range("A1").Value = "BIDVEST SCOPING TOOL - OUTPUT WORKBOOK"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Color = RGB(0, 112, 192)
        
        ' Subtitle
        .Range("A3").Value = "Generated Data Tables for Audit Scoping Analysis"
        .Range("A3").Font.Size = 12
        .Range("A3").Font.Italic = True
        
        ' Information section
        .Range("A5").Value = "Source Information:"
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Size = 11
        
        .Range("A6").Value = "Source Workbook:"
        .Range("B6").Value = g_SourceWorkbook.Name
        .Range("A7").Value = "Source Path:"
        .Range("B7").Value = g_SourceWorkbook.Path
        .Range("A8").Value = "Generated Date/Time:"
        .Range("B8").Value = Now()
        .Range("B8").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Range("A9").Value = "Tool Version:"
        .Range("B9").Value = ModConfig.GetToolVersion()
        
        ' Usage instructions
        .Range("A11").Value = "How to Use This Workbook:"
        .Range("A11").Font.Bold = True
        .Range("A11").Font.Size = 11
        .Range("A11").Interior.Color = RGB(217, 225, 242)
        
        .Range("A12").Value = "1. Review 'Scoping Summary' sheet for pack-level recommendations"
        .Range("A13").Value = "2. Check 'Scoped In by Division' and 'Scoped Out by Division' for division analysis"
        .Range("A14").Value = "3. Use 'Scoped In Packs Detail' to see FSLi-level amounts for scoped packs"
        .Range("A15").Value = "4. Review 'Threshold Configuration' if threshold-based scoping was applied"
        .Range("A16").Value = "5. For PowerBI integration, see POWERBI_COMPLETE_SETUP.md"
        
        ' Generated sheets list
        .Range("A18").Value = "Generated Sheets:"
        .Range("A18").Font.Bold = True
        .Range("A18").Font.Size = 11
        .Range("A18").Interior.Color = RGB(217, 225, 242)
        
        Dim sheetRow As Long
        sheetRow = 19
        .Range("A" & sheetRow).Value = "✓ Full Input Table (primary data)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ Scoping Summary (recommendations)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ Scoped In by Division (division breakdown)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ Scoped Out by Division (coverage gaps)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ Scoped In Packs Detail (FSLi amounts)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ FSLi Key Table (FSLi reference)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ Pack Number Company Table (pack reference)"
        sheetRow = sheetRow + 1
        .Range("A" & sheetRow).Value = "✓ Additional data tables as applicable"
        
        ' Format columns
        .Columns("A:B").AutoFit
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 50
        
        ' Add borders
        .Range("A6:B9").Borders.LineStyle = xlContinuous
        .Range("A6:B9").Borders.Weight = xlThin
        
        ' Color coding
        .Range("A6:A9").Interior.Color = RGB(242, 242, 242)
        .Range("A6:A9").Font.Bold = True
    End With
End Sub

' Create Division-based scoping reports showing scoped in/out per division
Private Sub CreateDivisionScopingReports(scopedPacks As Object)
    On Error GoTo ErrorHandler
    
    ' Create "Scoped In by Division" sheet
    CreateScopedInByDivision scopedPacks
    
    ' Create "Scoped Out by Division" sheet
    CreateScopedOutByDivision scopedPacks
    
    ' Create "Scoped In Packs Detail" sheet with FSLi amounts
    CreateScopedInPacksDetail scopedPacks
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error creating division reports: " & Err.Description
End Sub

' Create "Scoped In by Division" summary
Private Sub CreateScopedInByDivision(scopedPacks As Object)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim inputTab As Worksheet
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim packName As String
    Dim division As String
    Dim divisionDict As Object
    Dim divKey As Variant
    
    Set divisionDict = CreateObject("Scripting.Dictionary")
    
    ' Get input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then Exit Sub
    
    ' Collect scoped-in packs by division
    lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
    For col = 3 To lastCol
        packCode = Trim(inputTab.Cells(8, col).Value)
        packName = Trim(inputTab.Cells(7, col).Value)
        
        If packCode <> "" And packName <> "" Then
            ' Check if this pack is scoped in
            If Not scopedPacks Is Nothing Then
                If scopedPacks.Exists(packCode) Then
                    ' Get division from Pack Number Company Table
                    division = GetPackDivision(packCode)
                    
                    If Not divisionDict.Exists(division) Then
                        Set divisionDict(division) = CreateObject("Scripting.Dictionary")
                    End If
                    
                    If Not divisionDict(division).Exists(packCode) Then
                        divisionDict(division).Add packCode, packName
                    End If
                End If
            End If
        End If
    Next col
    
    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Scoped In by Division"
    
    ' Write headers
    row = 1
    With ws
        .Cells(row, 1).Value = "PACKS SCOPED IN BY DIVISION"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        .Cells(row, 1).Font.Color = RGB(0, 112, 192)
        row = row + 2
        
        ' Write data by division
        For Each divKey In divisionDict.Keys
            .Cells(row, 1).Value = "Division: " & divKey
            .Cells(row, 1).Font.Bold = True
            .Cells(row, 1).Interior.Color = RGB(217, 225, 242)
            row = row + 1
            
            .Cells(row, 1).Value = "Pack Code"
            .Cells(row, 2).Value = "Pack Name"
            .Cells(row, 1).Font.Bold = True
            .Cells(row, 2).Font.Bold = True
            row = row + 1
            
            Dim packDict As Object
            Set packDict = divisionDict(divKey)
            
            Dim packKey As Variant
            For Each packKey In packDict.Keys
                .Cells(row, 1).Value = packKey
                .Cells(row, 2).Value = packDict(packKey)
                row = row + 1
            Next packKey
            
            .Cells(row, 1).Value = "Count: " & packDict.Count
            .Cells(row, 1).Font.Italic = True
            row = row + 2
        Next divKey
        
        ' Summary
        .Cells(row, 1).Value = "TOTAL SCOPED IN: " & IIf(scopedPacks Is Nothing, 0, scopedPacks.Count)
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        
        .Columns("A:B").AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error creating Scoped In by Division: " & Err.Description
End Sub

' Create "Scoped Out by Division" summary
Private Sub CreateScopedOutByDivision(scopedPacks As Object)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim inputTab As Worksheet
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim packName As String
    Dim division As String
    Dim divisionDict As Object
    Dim divKey As Variant
    Dim totalPacks As Long
    
    Set divisionDict = CreateObject("Scripting.Dictionary")
    totalPacks = 0
    
    ' Get input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then Exit Sub
    
    ' Collect scoped-out packs by division
    lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
    For col = 3 To lastCol
        packCode = Trim(inputTab.Cells(8, col).Value)
        packName = Trim(inputTab.Cells(7, col).Value)
        
        If packCode <> "" And packName <> "" Then
            ' Check if this pack is NOT scoped in
            Dim isScoped As Boolean
            isScoped = False
            
            If Not scopedPacks Is Nothing Then
                If scopedPacks.Exists(packCode) Then
                    isScoped = True
                End If
            End If
            
            If Not isScoped Then
                totalPacks = totalPacks + 1
                ' Get division
                division = GetPackDivision(packCode)
                
                If Not divisionDict.Exists(division) Then
                    Set divisionDict(division) = CreateObject("Scripting.Dictionary")
                End If
                
                If Not divisionDict(division).Exists(packCode) Then
                    divisionDict(division).Add packCode, packName
                End If
            End If
        End If
    Next col
    
    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Scoped Out by Division"
    
    ' Write headers
    row = 1
    With ws
        .Cells(row, 1).Value = "PACKS NOT SCOPED IN (BY DIVISION)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        .Cells(row, 1).Font.Color = RGB(192, 0, 0)
        row = row + 2
        
        ' Write data by division
        For Each divKey In divisionDict.Keys
            .Cells(row, 1).Value = "Division: " & divKey
            .Cells(row, 1).Font.Bold = True
            .Cells(row, 1).Interior.Color = RGB(255, 242, 204)
            row = row + 1
            
            .Cells(row, 1).Value = "Pack Code"
            .Cells(row, 2).Value = "Pack Name"
            .Cells(row, 1).Font.Bold = True
            .Cells(row, 2).Font.Bold = True
            row = row + 1
            
            Dim packDict As Object
            Set packDict = divisionDict(divKey)
            
            Dim packKey As Variant
            For Each packKey In packDict.Keys
                .Cells(row, 1).Value = packKey
                .Cells(row, 2).Value = packDict(packKey)
                row = row + 1
            Next packKey
            
            .Cells(row, 1).Value = "Count: " & packDict.Count
            .Cells(row, 1).Font.Italic = True
            row = row + 2
        Next divKey
        
        ' Summary
        .Cells(row, 1).Value = "TOTAL NOT SCOPED: " & totalPacks
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        
        .Columns("A:B").AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error creating Scoped Out by Division: " & Err.Description
End Sub

' Create detailed report of scoped-in packs with FSLi amounts
Private Sub CreateScopedInPacksDetail(scopedPacks As Object)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim inputTab As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim col As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim packCode As String
    Dim packName As String
    Dim fsliName As String
    Dim amount As Variant
    Dim packCol As Long
    
    ' Get input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then Exit Sub
    
    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Scoped In Packs Detail"
    
    ' Write headers
    row = 1
    With ws
        .Cells(row, 1).Value = "SCOPED IN PACKS - DETAILED FSLi AMOUNTS"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        .Cells(row, 1).Font.Color = RGB(0, 112, 192)
        row = row + 2
        
        .Cells(row, 1).Value = "Pack Code"
        .Cells(row, 2).Value = "Pack Name"
        .Cells(row, 3).Value = "FSLi"
        .Cells(row, 4).Value = "Amount"
        .Cells(row, 5).Value = "% of Pack Total"
        .Range("A" & row & ":E" & row).Font.Bold = True
        .Range("A" & row & ":E" & row).Interior.Color = RGB(68, 114, 196)
        .Range("A" & row & ":E" & row).Font.Color = RGB(255, 255, 255)
        row = row + 1
        
        ' Process each scoped-in pack
        If Not scopedPacks Is Nothing Then
            lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
            lastRow = inputTab.Cells(inputTab.Rows.Count, 2).End(xlUp).row
            
            For col = 3 To lastCol
                packCode = Trim(inputTab.Cells(8, col).Value)
                packName = Trim(inputTab.Cells(7, col).Value)
                
                If packCode <> "" And packName <> "" Then
                    If scopedPacks.Exists(packCode) Then
                        packCol = col
                        
                        ' Calculate pack total
                        Dim packTotal As Double
                        packTotal = 0
                        For dataRow = 9 To lastRow
                            If IsNumeric(inputTab.Cells(dataRow, packCol).Value) Then
                                packTotal = packTotal + Abs(CDbl(inputTab.Cells(dataRow, packCol).Value))
                            End If
                        Next dataRow
                        
                        ' Write each FSLi for this pack
                        For dataRow = 9 To lastRow
                            fsliName = Trim(inputTab.Cells(dataRow, 2).Value)
                            amount = inputTab.Cells(dataRow, packCol).Value
                            
                            If fsliName <> "" And IsNumeric(amount) Then
                                .Cells(row, 1).Value = packCode
                                .Cells(row, 2).Value = packName
                                .Cells(row, 3).Value = fsliName
                                .Cells(row, 4).Value = CDbl(amount)
                                .Cells(row, 4).NumberFormat = "#,##0.00"
                                
                                ' Calculate percentage
                                If packTotal > 0 Then
                                    .Cells(row, 5).Value = Abs(CDbl(amount)) / packTotal
                                    .Cells(row, 5).NumberFormat = "0.00%"
                                End If
                                
                                row = row + 1
                            End If
                        Next dataRow
                    End If
                End If
            Next col
        End If
        
        ' Auto-fit columns
        .Columns("A:E").AutoFit
        
        ' Add table if data exists
        If row > 4 Then
            Dim tbl As ListObject
            On Error Resume Next
            Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(3, 1), .Cells(row - 1, 5)), , xlYes)
            If Not tbl Is Nothing Then
                tbl.Name = "Scoped_In_Packs_Detail_Table"
                tbl.TableStyle = "TableStyleMedium2"
            End If
            On Error GoTo ErrorHandler
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error creating Scoped In Packs Detail: " & Err.Description
End Sub

' Helper function to get division for a pack code
Private Function GetPackDivision(packCode As String) As String
    On Error Resume Next
    
    Dim packWs As Worksheet
    Dim lastRow As Long
    Dim row As Long
    
    ' Try to find in Pack Number Company Table
    Set packWs = g_OutputWorkbook.Worksheets("Pack Number Company Table")
    
    If Not packWs Is Nothing Then
        lastRow = packWs.Cells(packWs.Rows.Count, 1).End(xlUp).row
        
        For row = 2 To lastRow
            If Trim(packWs.Cells(row, 2).Value) = packCode Then ' Column 2 is Pack Code
                GetPackDivision = Trim(packWs.Cells(row, 3).Value) ' Column 3 is Division
                Exit Function
            End If
        Next row
    End If
    
    ' Default division if not found
    GetPackDivision = "Unknown Division"
    
    On Error GoTo 0
End Function

' Select the consolidated entity (to exclude from scoping)
Private Function SelectConsolidatedEntity() As Boolean
    On Error GoTo ErrorHandler
    
    Dim inputTab As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim packName As String
    Dim packList As String
    Dim packDict As Object
    Dim packKey As Variant
    Dim userInput As String
    Dim selectedIndex As Long
    Dim counter As Long
    
    ' Initialize
    g_ConsolidatedPackCode = ""
    g_ConsolidatedPackName = ""
    
    ' Get the input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then
        SelectConsolidatedEntity = True ' Continue without consolidated selection
        Exit Function
    End If
    
    ' Create dictionary to store unique packs
    Set packDict = CreateObject("Scripting.Dictionary")
    
    ' Collect all packs from input tab
    lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
    counter = 0
    
    For col = 3 To lastCol
        packCode = Trim(inputTab.Cells(8, col).Value)
        packName = Trim(inputTab.Cells(7, col).Value)
        
        If packCode <> "" And packName <> "" Then
            If Not packDict.Exists(packCode) Then
                counter = counter + 1
                packDict.Add packCode, Array(packName, counter)
            End If
        End If
    Next col
    
    If packDict.Count = 0 Then
        SelectConsolidatedEntity = True ' No packs found, continue
        Exit Function
    End If
    
    ' Build message with pack list
    packList = "CONSOLIDATED ENTITY SELECTION" & vbCrLf & vbCrLf
    packList = packList & "Select which pack represents the CONSOLIDATED entity." & vbCrLf
    packList = packList & "This pack will be EXCLUDED from scoping calculations" & vbCrLf
    packList = packList & "as it represents consolidated totals, not individual entities." & vbCrLf & vbCrLf
    packList = packList & "Available Packs:" & vbCrLf
    packList = packList & String(60, "-") & vbCrLf
    
    ' List all packs with numbers
    For Each packKey In packDict.Keys
        Dim packInfo As Variant
        packInfo = packDict(packKey)
        packList = packList & packInfo(1) & ". " & packInfo(0) & " (" & packKey & ")" & vbCrLf
    Next packKey
    
    packList = packList & vbCrLf & "Enter the number of the consolidated pack:"
    packList = packList & vbCrLf & "(Or leave blank to include all packs in scoping)"
    
    ' Get user input
    userInput = InputBox(packList, "Select Consolidated Entity", "")
    
    ' Parse user input
    If Trim(userInput) = "" Then
        ' User chose to include all packs
        SelectConsolidatedEntity = True
        Exit Function
    End If
    
    ' Validate input is a number
    If Not IsNumeric(userInput) Then
        MsgBox "Invalid input. Please enter a number from the list.", vbExclamation
        SelectConsolidatedEntity = SelectConsolidatedEntity() ' Recursive call to try again
        Exit Function
    End If
    
    selectedIndex = CLng(userInput)
    
    ' Find the selected pack
    For Each packKey In packDict.Keys
        packInfo = packDict(packKey)
        If packInfo(1) = selectedIndex Then
            g_ConsolidatedPackCode = CStr(packKey)
            g_ConsolidatedPackName = packInfo(0)
            
            ' Confirm selection
            Dim confirmMsg As String
            confirmMsg = "You selected:" & vbCrLf & vbCrLf
            confirmMsg = confirmMsg & "Pack Name: " & g_ConsolidatedPackName & vbCrLf
            confirmMsg = confirmMsg & "Pack Code: " & g_ConsolidatedPackCode & vbCrLf & vbCrLf
            confirmMsg = confirmMsg & "This pack will be EXCLUDED from scoping." & vbCrLf & vbCrLf
            confirmMsg = confirmMsg & "Is this correct?"
            
            If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Selection") = vbYes Then
                SelectConsolidatedEntity = True
                Exit Function
            Else
                ' User wants to reselect
                SelectConsolidatedEntity = SelectConsolidatedEntity() ' Recursive call
                Exit Function
            End If
        End If
    Next packKey
    
    ' If we get here, invalid selection
    MsgBox "Invalid selection. Please enter a number from 1 to " & packDict.Count, vbExclamation
    SelectConsolidatedEntity = SelectConsolidatedEntity() ' Recursive call to try again
    Exit Function
    
ErrorHandler:
    MsgBox "Error selecting consolidated entity: " & Err.Description, vbCritical
    SelectConsolidatedEntity = False
End Function
