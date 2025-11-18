Attribute VB_Name = "Mod1_MainController"
Option Explicit

' =================================================================================
' MODULE 1: MAIN CONTROLLER AND USER INTERFACE
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 7.0 - Complete Fix with Symbol Removal
' =================================================================================
' PURPOSE:
'   Main entry point and workflow orchestration for the scoping tool
'   Manages the entire process from workbook selection through final output
'
' CRITICAL FIXES:
'   1. Removed all Unicode symbols (checkmarks, bullets) from prompts
'   2. Replaced with plain text equivalents for better compatibility
'   3. Professional clean appearance
'   4. Updated version to 7.0
'
' AUTHOR: ISA 600 Scoping Tool Team
' DATE: 2025-11-18
' =================================================================================

' ==================== GLOBAL VARIABLES ====================
Public g_StripePacksWorkbook As Workbook        ' Stripe Packs consolidation workbook
Public g_SegmentalWorkbook As Workbook          ' Segmental reporting workbook
Public g_OutputWorkbook As Workbook             ' Generated output workbook
Public g_TabCategories As Object                ' Dictionary of tab categorizations
Public g_DivisionNames As Object                ' Dictionary of division names (key = tab name, value = division name)
Public g_ConsolidationEntity As String          ' Consolidation entity pack code
Public g_ConsolidationEntityName As String      ' Consolidation entity pack name
Public g_ThresholdFSLIs As Collection           ' Collection of threshold configurations
Public g_ScopedPacks As Object                  ' Dictionary of scoped packs
Public g_ManualScoping As Object                ' Dictionary for manual scoping (packCode_FSLI -> status)
Public g_UseConsolidationCurrency As Boolean   ' User preference for currency type

' ==================== MAIN ENTRY POINT ====================
Public Sub StartBidvestScopingTool()
    '------------------------------------------------------------------------
    ' MAIN ENTRY POINT - Button Click Handler
    ' Orchestrates the entire scoping tool workflow with user guidance
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    ' Display welcome message (NO SYMBOLS)
    Dim result As VbMsgBoxResult
    result = MsgBox( _
        "ISA 600 REVISED - BIDVEST GROUP SCOPING TOOL" & vbCrLf & vbCrLf & _
        "This comprehensive tool will:" & vbCrLf & _
        "- Process Stripe Packs consolidation workbook" & vbCrLf & _
        "- Process Segmental Reporting workbook" & vbCrLf & _
        "- Generate structured data tables" & vbCrLf & _
        "- Create interactive dashboard" & vbCrLf & _
        "- Prepare Power BI-ready datasets" & vbCrLf & _
        "- Apply threshold-based scoping" & vbCrLf & _
        "- Enable manual scoping interface" & vbCrLf & vbCrLf & _
        "Estimated time: 5-10 minutes" & vbCrLf & vbCrLf & _
        "Click OK to begin or Cancel to exit.", _
        vbOKCancel + vbInformation, "Bidvest Scoping Tool v7.0")

    If result = vbCancel Then Exit Sub

    ' Initialize global objects
    InitializeGlobalObjects

    ' PART 1: STRIPE PACKS CONSOLIDATION WORKBOOK
    Application.StatusBar = "Step 1/12: Selecting Stripe Packs consolidation workbook..."
    If Not SelectStripePacksWorkbook() Then
        MsgBox "Stripe Packs workbook selection cancelled. Process terminated.", vbInformation
        CleanupAndExit
        Exit Sub
    End If

    ' PART 2: TAB CATEGORIZATION
    Application.StatusBar = "Step 2/12: Analyzing tabs..."
    If Not CategorizeTabs() Then
        MsgBox "Tab categorization cancelled. Process terminated.", vbInformation
        CleanupAndExit
        Exit Sub
    End If

    ' PART 3: DIVISION NAME ASSIGNMENT
    Application.StatusBar = "Step 3/12: Assigning division names..."
    If Not AssignDivisionNames() Then
        MsgBox "Division name assignment cancelled. Process terminated.", vbInformation
        CleanupAndExit
        Exit Sub
    End If

    ' PART 4: CURRENCY TYPE SELECTION
    Application.StatusBar = "Step 4/12: Selecting currency type..."
    If Not SelectCurrencyType() Then
        MsgBox "Currency selection cancelled. Process terminated.", vbInformation
        CleanupAndExit
        Exit Sub
    End If

    ' PART 5: CONSOLIDATION ENTITY IDENTIFICATION
    Application.StatusBar = "Step 5/12: Identifying consolidation entity..."
    If Not IdentifyConsolidationEntity() Then
        MsgBox "Consolidation entity identification cancelled. Process terminated.", vbInformation
        CleanupAndExit
        Exit Sub
    End If

    ' PART 6: SEGMENTAL REPORTING WORKBOOK (OPTIONAL)
    Application.StatusBar = "Step 6/12: Processing segmental reporting..."
    ProcessSegmentalReporting ' Optional - continues even if cancelled

    ' PART 7: CREATE OUTPUT WORKBOOK
    Application.StatusBar = "Step 7/12: Creating output workbook..."
    CreateOutputWorkbook

    ' PART 8: EXTRACT AND GENERATE TABLES
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "Step 8/12: Extracting data and generating tables..."
    ExtractAndGenerateTables

    ' PART 9: THRESHOLD CONFIGURATION (OPTIONAL)
    Application.StatusBar = "Step 9/12: Configuring thresholds..."
    ConfigureThresholds ' Optional

    ' PART 10: CREATE INTERACTIVE DASHBOARD
    Application.StatusBar = "Step 10/12: Creating interactive dashboard..."
    Mod6_DashboardGeneration.CreateComprehensiveDashboard

    ' PART 11: CREATE POWER BI ASSETS
    Application.StatusBar = "Step 11/12: Creating Power BI integration assets..."
    Mod7_PowerBIExport.CreatePowerBIAssets

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' PART 12: SAVE OUTPUT WORKBOOK
    Application.StatusBar = "Step 12/12: Saving output workbook..."
    SaveOutputWorkbook

    Application.StatusBar = False

    ' Display completion message (NO SYMBOLS)
    Dim elapsedTime As Long
    elapsedTime = CInt(Timer - startTime)

    MsgBox _
        "SCOPING TOOL COMPLETED SUCCESSFULLY!" & vbCrLf & vbCrLf & _
        "Output workbook: " & g_OutputWorkbook.Name & vbCrLf & _
        "Location: " & g_OutputWorkbook.Path & vbCrLf & _
        "Processing time: " & elapsedTime & " seconds" & vbCrLf & vbCrLf & _
        "GENERATED ASSETS:" & vbCrLf & _
        "[DONE] Full Input Table and Percentages" & vbCrLf & _
        "[DONE] Discontinued, Journals, Consol Tables" & vbCrLf & _
        "[DONE] Division-Segment Mapping" & vbCrLf & _
        "[DONE] Interactive Dashboard (6 views)" & vbCrLf & _
        "[DONE] Manual Scoping Interface" & vbCrLf & _
        "[DONE] Power BI-Ready Tables" & vbCrLf & _
        "[DONE] Scoping Summary and Reports" & vbCrLf & vbCrLf & _
        "NEXT STEPS:" & vbCrLf & _
        "1. Review Dashboard - Overview sheet" & vbCrLf & _
        "2. Use Manual Scoping Interface to adjust scoping" & vbCrLf & _
        "3. Import to Power BI for advanced analysis" & vbCrLf & vbCrLf & _
        "See COMPREHENSIVE_IMPLEMENTATION_GUIDE.md for details.", _
        vbInformation, "Process Complete - Bidvest Scoping Tool v7.0"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "An error occurred during processing:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Please check your source data and try again.", _
           vbCritical, "Error"

    CleanupAndExit
End Sub

' ==================== INITIALIZATION ====================
Private Sub InitializeGlobalObjects()
    '------------------------------------------------------------------------
    ' Initialize all global objects/collections
    '------------------------------------------------------------------------
    Set g_TabCategories = CreateObject("Scripting.Dictionary")
    Set g_DivisionNames = CreateObject("Scripting.Dictionary")
    Set g_ThresholdFSLIs = New Collection
    Set g_ScopedPacks = CreateObject("Scripting.Dictionary")
    Set g_ManualScoping = CreateObject("Scripting.Dictionary")

    g_ConsolidationEntity = ""
    g_ConsolidationEntityName = ""
    g_UseConsolidationCurrency = True
End Sub

' ==================== STRIPE PACKS WORKBOOK SELECTION ====================
Private Function SelectStripePacksWorkbook() As Boolean
    '------------------------------------------------------------------------
    ' STEP 1: Prompt user to select Stripe Packs consolidation workbook
    ' Returns True if successful, False if cancelled
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim workbookName As String
    Dim promptMsg As String

    promptMsg = "STRIPE PACKS CONSOLIDATION WORKBOOK" & vbCrLf & vbCrLf & _
                "Please provide the name of the Stripe Packs consolidation workbook." & vbCrLf & vbCrLf & _
                "Instructions:" & vbCrLf & _
                "1. Open the Stripe Packs consolidation workbook in Excel" & vbCrLf & _
                "2. Copy the exact workbook name from the title bar" & vbCrLf & _
                "3. Paste it below (include .xlsx or .xlsm extension)" & vbCrLf & vbCrLf & _
                "Example: Bidvest_Consolidation_2024.xlsx"

    workbookName = InputBox(promptMsg, "Select Stripe Packs Workbook", "")

    If Trim(workbookName) = "" Then
        SelectStripePacksWorkbook = False
        Exit Function
    End If

    ' Validate workbook exists
    Set g_StripePacksWorkbook = Mod8_Utilities.GetWorkbookByName(workbookName)

    If g_StripePacksWorkbook Is Nothing Then
        MsgBox "Could not find workbook '" & workbookName & "'." & vbCrLf & vbCrLf & _
               "Please ensure:" & vbCrLf & _
               "- The workbook is open in Excel" & vbCrLf & _
               "- The name is spelled correctly" & vbCrLf & _
               "- The file extension is included", _
               vbExclamation, "Workbook Not Found"
        SelectStripePacksWorkbook = False
        Exit Function
    End If

    ' Success
    MsgBox "Stripe Packs workbook loaded successfully:" & vbCrLf & vbCrLf & _
           g_StripePacksWorkbook.Name & vbCrLf & vbCrLf & _
           "Tabs found: " & g_StripePacksWorkbook.Worksheets.Count, _
           vbInformation, "Workbook Loaded"

    SelectStripePacksWorkbook = True
    Exit Function

ErrorHandler:
    SelectStripePacksWorkbook = False
End Function

' ==================== TAB CATEGORIZATION ====================
Private Function CategorizeTabs() As Boolean
    '------------------------------------------------------------------------
    ' STEP 2: Categorize all tabs in the Stripe Packs workbook
    ' Uses Mod2_TabProcessing module
    ' Returns True if successful, False if cancelled
    '------------------------------------------------------------------------
    CategorizeTabs = Mod2_TabProcessing.CategorizeAllTabs(g_StripePacksWorkbook, g_TabCategories)
End Function

' ==================== DIVISION NAME ASSIGNMENT ====================
Private Function AssignDivisionNames() As Boolean
    '------------------------------------------------------------------------
    ' STEP 3: Prompt for division names for each Division category tab
    ' Returns True if successful, False if cancelled
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim tabName As Variant
    Dim divisionName As String
    Dim divisionCount As Long

    divisionCount = 0

    ' Count Division tabs
    For Each tabName In g_TabCategories.Keys
        If g_TabCategories(tabName) = "Division" Then
            divisionCount = divisionCount + 1
        End If
    Next tabName

    If divisionCount = 0 Then
        ' No divisions, continue without division names
        AssignDivisionNames = True
        Exit Function
    End If

    ' Prompt for each division name
    For Each tabName In g_TabCategories.Keys
        If g_TabCategories(tabName) = "Division" Then
            divisionName = InputBox( _
                "Division Tab: " & tabName & vbCrLf & vbCrLf & _
                "Please enter a friendly name for this division:" & vbCrLf & _
                "(This will be used in reports and dashboards)" & vbCrLf & vbCrLf & _
                "Example: UK Division, South Africa Division", _
                "Division Name Assignment", _
                CStr(tabName))

            If Trim(divisionName) = "" Then
                divisionName = CStr(tabName) ' Use tab name as fallback
            End If

            g_DivisionNames(tabName) = divisionName
        End If
    Next tabName

    AssignDivisionNames = True
    Exit Function

ErrorHandler:
    AssignDivisionNames = False
End Function

' ==================== CURRENCY TYPE SELECTION ====================
Private Function SelectCurrencyType() As Boolean
    '------------------------------------------------------------------------
    ' STEP 4: Prompt user to select currency type for processing
    ' Returns True if successful, False if cancelled
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim result As VbMsgBoxResult

    result = MsgBox( _
        "CURRENCY TYPE SELECTION" & vbCrLf & vbCrLf & _
        "The consolidation workbook contains data in multiple currencies." & vbCrLf & vbCrLf & _
        "Row 6 typically shows:" & vbCrLf & _
        "- Original/Entity Currency (local currency of each entity)" & vbCrLf & _
        "- Consolidation Currency (group reporting currency)" & vbCrLf & vbCrLf & _
        "For ISA 600 scoping, you should use CONSOLIDATION CURRENCY." & vbCrLf & vbCrLf & _
        "Click YES to use Consolidation Currency (recommended)" & vbCrLf & _
        "Click NO to use Entity Currency", _
        vbYesNoCancel + vbQuestion, "Currency Selection")

    If result = vbCancel Then
        SelectCurrencyType = False
        Exit Function
    End If

    g_UseConsolidationCurrency = (result = vbYes)

    Dim confirmMsg As String
    confirmMsg = "Currency selection confirmed:" & vbCrLf & vbCrLf

    If g_UseConsolidationCurrency Then
        confirmMsg = confirmMsg & "[SELECTED] Consolidation Currency" & vbCrLf & vbCrLf & _
                     "The tool will process columns identified as" & vbCrLf & _
                     "'Consolidation' or 'Consolidable' in Row 6."
    Else
        confirmMsg = confirmMsg & "[SELECTED] Entity Currency" & vbCrLf & vbCrLf & _
                     "The tool will process columns identified as" & vbCrLf & _
                     "'Original' or 'Entity' in Row 6."
    End If

    MsgBox confirmMsg, vbInformation, "Currency Confirmed"

    SelectCurrencyType = True
    Exit Function

ErrorHandler:
    SelectCurrencyType = False
End Function

' ==================== CONSOLIDATION ENTITY IDENTIFICATION ====================
Private Function IdentifyConsolidationEntity() As Boolean
    '------------------------------------------------------------------------
    ' STEP 5: Identify which entity is the consolidation entity
    ' This entity represents the aggregate and should be excluded from scoping
    ' Returns True if successful, False if cancelled
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    ' Use Mod3_DataExtraction to get list of all entities
    Dim entities As Object ' Dictionary: code -> name
    Set entities = Mod3_DataExtraction.GetAllEntitiesFromInputContinuing(g_TabCategories, g_UseConsolidationCurrency)

    If entities.Count = 0 Then
        MsgBox "No entities found in Input Continuing tab. Cannot continue.", vbExclamation
        IdentifyConsolidationEntity = False
        Exit Function
    End If

    ' Build selection prompt
    Dim promptMsg As String
    Dim entityKey As Variant
    Dim counter As Long
    Dim entityList As Object ' Dictionary: index -> code

    Set entityList = CreateObject("Scripting.Dictionary")
    counter = 1

    promptMsg = "CONSOLIDATION ENTITY IDENTIFICATION" & vbCrLf & vbCrLf & _
                "Select which entity represents the CONSOLIDATED entity." & vbCrLf & _
                "This entity (typically BBT-001 or similar) aggregates all packs" & vbCrLf & _
                "and will be used as the 100% baseline for percentage calculations." & vbCrLf & vbCrLf & _
                "Available Entities:" & vbCrLf & _
                String(60, "-") & vbCrLf

    For Each entityKey In entities.Keys
        promptMsg = promptMsg & counter & ". " & entities(entityKey) & " (" & entityKey & ")" & vbCrLf
        entityList(counter) = entityKey
        counter = counter + 1
    Next entityKey

    promptMsg = promptMsg & vbCrLf & "Enter the number of the consolidation entity:"

    Dim userInput As String
    Dim selectedIndex As Long

    userInput = InputBox(promptMsg, "Select Consolidation Entity", "1")

    If Trim(userInput) = "" Then
        IdentifyConsolidationEntity = False
        Exit Function
    End If

    If Not IsNumeric(userInput) Then
        MsgBox "Invalid input. Please enter a number.", vbExclamation
        IdentifyConsolidationEntity = False
        Exit Function
    End If

    selectedIndex = CLng(userInput)

    If Not entityList.exists(selectedIndex) Then
        MsgBox "Invalid selection. Please select a valid number from the list.", vbExclamation
        IdentifyConsolidationEntity = False
        Exit Function
    End If

    g_ConsolidationEntity = CStr(entityList(selectedIndex))
    g_ConsolidationEntityName = entities(g_ConsolidationEntity)

    ' Confirm selection
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox( _
        "Consolidation Entity Selected:" & vbCrLf & vbCrLf & _
        "Name: " & g_ConsolidationEntityName & vbCrLf & _
        "Code: " & g_ConsolidationEntity & vbCrLf & vbCrLf & _
        "This entity will be:" & vbCrLf & _
        "- Used as 100% baseline for all FSLIs" & vbCrLf & _
        "- Excluded from scoping (it's the aggregate)" & vbCrLf & vbCrLf & _
        "Is this correct?", _
        vbYesNo + vbQuestion, "Confirm Selection")

    If confirmResult <> vbYes Then
        IdentifyConsolidationEntity = False
        Exit Function
    End If

    IdentifyConsolidationEntity = True
    Exit Function

ErrorHandler:
    IdentifyConsolidationEntity = False
End Function

' ==================== SEGMENTAL REPORTING ====================
Private Sub ProcessSegmentalReporting()
    '------------------------------------------------------------------------
    ' STEP 6: Optional - Process Segmental Reporting workbook
    ' Performs pack matching between Stripe and Segmental
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim result As VbMsgBoxResult

    result = MsgBox( _
        "SEGMENTAL REPORTING INTEGRATION" & vbCrLf & vbCrLf & _
        "Would you like to process the Segmental Reporting workbook?" & vbCrLf & vbCrLf & _
        "This will:" & vbCrLf & _
        "- Match packs between Stripe and Segmental reports" & vbCrLf & _
        "- Create Division-Segment mapping table" & vbCrLf & _
        "- Enable segment-based analysis in dashboard" & vbCrLf & vbCrLf & _
        "Click YES to process Segmental Reporting" & vbCrLf & _
        "Click NO to skip (you can still use division-based analysis)", _
        vbYesNo + vbQuestion, "Segmental Reporting")

    If result <> vbYes Then
        Exit Sub ' User chose to skip
    End If

    ' Select segmental workbook
    Dim workbookName As String
    workbookName = InputBox( _
        "SEGMENTAL REPORTING WORKBOOK" & vbCrLf & vbCrLf & _
        "Please provide the name of the Segmental Reporting workbook." & vbCrLf & vbCrLf & _
        "Instructions:" & vbCrLf & _
        "1. Open the Segmental Reporting workbook in Excel" & vbCrLf & _
        "2. Copy the exact workbook name from the title bar" & vbCrLf & _
        "3. Paste it below (include .xlsx or .xlsm extension)", _
        "Select Segmental Workbook", "")

    If Trim(workbookName) = "" Then
        MsgBox "Segmental reporting skipped.", vbInformation
        Exit Sub
    End If

    Set g_SegmentalWorkbook = Mod8_Utilities.GetWorkbookByName(workbookName)

    If g_SegmentalWorkbook Is Nothing Then
        MsgBox "Could not find workbook '" & workbookName & "'." & vbCrLf & _
               "Segmental reporting will be skipped.", vbExclamation
        Exit Sub
    End If

    ' Process segmental workbook using Mod4_SegmentalMatching
    Mod4_SegmentalMatching.ProcessSegmentalWorkbook g_SegmentalWorkbook, g_TabCategories, g_DivisionNames

    MsgBox "Segmental reporting processed successfully!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error processing segmental reporting:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Segmental analysis will be skipped.", vbExclamation
End Sub

' ==================== OUTPUT WORKBOOK CREATION ====================
Private Sub CreateOutputWorkbook()
    '------------------------------------------------------------------------
    ' STEP 7: Create the output workbook for all generated tables
    '------------------------------------------------------------------------
    Set g_OutputWorkbook = Workbooks.Add
    g_OutputWorkbook.Worksheets(1).Name = "ReadMe"

    ' Create informational ReadMe sheet
    With g_OutputWorkbook.Worksheets("ReadMe")
        .Range("A1").Value = "BIDVEST GROUP - ISA 600 SCOPING TOOL OUTPUT"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(0, 112, 192)

        .Range("A3").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        .Range("A4").Value = "Source (Stripe): " & g_StripePacksWorkbook.Name
        If Not g_SegmentalWorkbook Is Nothing Then
            .Range("A5").Value = "Source (Segmental): " & g_SegmentalWorkbook.Name
        End If
        .Range("A6").Value = "Tool Version: 7.0 - Complete Overhaul"

        .Range("A8").Value = "GENERATED CONTENT:"
        .Range("A8").Font.Bold = True

        Dim row As Long
        row = 9
        .Cells(row, 1).Value = "[DONE] Full Input Table (amounts)": row = row + 1
        .Cells(row, 1).Value = "[DONE] Full Input Percentage Table": row = row + 1
        .Cells(row, 1).Value = "[DONE] Discontinued, Journals, Consol Tables": row = row + 1
        .Cells(row, 1).Value = "[DONE] Dim FSLIs Table (reference)": row = row + 1
        .Cells(row, 1).Value = "[DONE] Pack Number Company Table (reference)": row = row + 1
        .Cells(row, 1).Value = "[DONE] Division-Segment Mapping (if applicable)": row = row + 1
        .Cells(row, 1).Value = "[DONE] Interactive Dashboard (6 views)": row = row + 1
        .Cells(row, 1).Value = "[DONE] Manual Scoping Interface": row = row + 1
        .Cells(row, 1).Value = "[DONE] Power BI-Ready Tables": row = row + 1
        .Cells(row, 1).Value = "[DONE] Fact Scoping Table": row = row + 1

        .Columns("A:A").AutoFit
    End With
End Sub

' ==================== EXTRACT AND GENERATE TABLES ====================
Private Sub ExtractAndGenerateTables()
    '------------------------------------------------------------------------
    ' STEP 8: Extract data and generate all required tables
    ' Uses Mod3_DataExtraction
    '------------------------------------------------------------------------

    ' Generate Full Input Table and Percentage Table
    Mod3_DataExtraction.GenerateFullInputTables g_TabCategories, g_UseConsolidationCurrency, g_ConsolidationEntity

    ' Generate Discontinued, Journals, Consol tables (if tabs exist)
    If TabCategoryExists("Discontinued Operations") Then
        Mod3_DataExtraction.GenerateDiscontinuedTables g_TabCategories, g_UseConsolidationCurrency, g_ConsolidationEntity
    End If

    If TabCategoryExists("Journals Continuing") Then
        Mod3_DataExtraction.GenerateJournalsTables g_TabCategories, g_UseConsolidationCurrency, g_ConsolidationEntity
    End If

    If TabCategoryExists("Consol Continuing") Then
        Mod3_DataExtraction.GenerateConsolTables g_TabCategories, g_UseConsolidationCurrency, g_ConsolidationEntity
    End If

    ' Generate reference tables
    Mod3_DataExtraction.GenerateFSLiKeyTable g_TabCategories
    Mod3_DataExtraction.GeneratePackCompanyTable g_TabCategories, g_DivisionNames, g_ConsolidationEntity
End Sub

' ==================== THRESHOLD CONFIGURATION ====================
Private Sub ConfigureThresholds()
    '------------------------------------------------------------------------
    ' STEP 9: Optional - Configure threshold-based automatic scoping
    ' Uses Mod5_ScopingEngine
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim result As VbMsgBoxResult

    result = MsgBox( _
        "THRESHOLD-BASED AUTOMATIC SCOPING" & vbCrLf & vbCrLf & _
        "Would you like to configure threshold-based scoping?" & vbCrLf & vbCrLf & _
        "This allows you to:" & vbCrLf & _
        "- Select specific FSLIs as threshold criteria" & vbCrLf & _
        "- Set threshold amounts for each FSLI" & vbCrLf & _
        "- Automatically scope in packs exceeding thresholds" & vbCrLf & vbCrLf & _
        "Example: Revenue > R50M OR Total Assets > R100M" & vbCrLf & vbCrLf & _
        "Click YES to configure thresholds" & vbCrLf & _
        "Click NO to skip (you can manually scope later)", _
        vbYesNo + vbQuestion, "Threshold Configuration")

    If result <> vbYes Then
        ' Create empty Fact_Scoping table even if no thresholds
        Dim emptyPacks As Object
        Dim emptyThresholds As Collection
        Set emptyPacks = CreateObject("Scripting.Dictionary")
        Set emptyThresholds = New Collection

        Mod5_ScopingEngine.GenerateFactScopingTable emptyPacks, emptyThresholds, g_ConsolidationEntity
        Exit Sub ' User chose to skip
    End If

    ' Configure thresholds using Mod5_ScopingEngine
    Set g_ThresholdFSLIs = Mod5_ScopingEngine.ConfigureThresholds()

    If g_ThresholdFSLIs.Count > 0 Then
        ' Apply thresholds and identify scoped packs
        Set g_ScopedPacks = Mod5_ScopingEngine.ApplyThresholds(g_ThresholdFSLIs, g_ConsolidationEntity)

        ' Generate Fact_Scoping table
        Mod5_ScopingEngine.GenerateFactScopingTable g_ScopedPacks, g_ThresholdFSLIs, g_ConsolidationEntity

        ' Generate Dim_Thresholds table
        Mod5_ScopingEngine.GenerateDimThresholdsTable g_ThresholdFSLIs

        ' Generate Scoping Summary
        Mod5_ScopingEngine.GenerateScopingSummary g_ScopedPacks, g_ThresholdFSLIs

        MsgBox "Threshold scoping applied successfully!" & vbCrLf & vbCrLf & _
               "Packs scoped in: " & g_ScopedPacks.Count, _
               vbInformation, "Thresholds Applied"
    Else
        ' Create empty Fact_Scoping table
        Dim emptyDict As Object
        Set emptyDict = CreateObject("Scripting.Dictionary")
        Mod5_ScopingEngine.GenerateFactScopingTable emptyDict, g_ThresholdFSLIs, g_ConsolidationEntity
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error configuring thresholds: " & Err.Description, vbExclamation
End Sub

' ==================== SAVE OUTPUT WORKBOOK ====================
Private Sub SaveOutputWorkbook()
    '------------------------------------------------------------------------
    ' STEP 12: Save output workbook with timestamped filename
    ' Format: Bidvest Group Scoping [YYYY-MM-DD] [HH-MM-SS].xlsm
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim saveDirectory As String
    Dim fileName As String
    Dim savePath As String
    Dim timestamp As String

    ' Generate timestamp: YYYY-MM-DD HH-MM-SS
    timestamp = Format(Now, "yyyy-mm-dd hh-mm-ss")

    ' Standard filename with timestamp
    fileName = "Bidvest Group Scoping [" & timestamp & "].xlsm"

    ' Determine save directory (same as Stripe Packs workbook)
    If g_StripePacksWorkbook.Path <> "" Then
        saveDirectory = g_StripePacksWorkbook.Path
    Else
        ' Fallback to Documents folder
        saveDirectory = Environ("USERPROFILE") & "\Documents"
    End If

    savePath = saveDirectory & Application.PathSeparator & fileName

    ' Save as macro-enabled workbook
    Application.DisplayAlerts = False
    g_OutputWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

    MsgBox "Output workbook saved successfully:" & vbCrLf & vbCrLf & savePath, vbInformation, "File Saved"

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Could not save output workbook:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Please save manually using File > Save As.", vbExclamation, "Save Error"
End Sub

' ==================== HELPER FUNCTIONS ====================
Private Function TabCategoryExists(categoryName As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if any tab is categorized as the specified category
    '------------------------------------------------------------------------
    Dim tabName As Variant

    For Each tabName In g_TabCategories.Keys
        If g_TabCategories(tabName) = categoryName Then
            TabCategoryExists = True
            Exit Function
        End If
    Next tabName

    TabCategoryExists = False
End Function

Private Sub CleanupAndExit()
    '------------------------------------------------------------------------
    ' Cleanup and reset state
    '------------------------------------------------------------------------
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
