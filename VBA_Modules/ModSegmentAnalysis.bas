Attribute VB_Name = "ModSegmentAnalysis"
Option Explicit

' ============================================================================
' MODULE: ModSegmentAnalysis
' PURPOSE: Handle IAS 8 operating segment analysis and mapping
' DESCRIPTION: Processes segment reporting documents, extracts pack-to-segment
'              mappings, and creates analysis tables for Power BI integration
' VERSION: 5.0 - NEW MODULE
' ============================================================================

' Global variables for segment analysis
Public g_SegmentWorkbook As Workbook
Public g_SegmentTabCategories As Object ' Dictionary for segment tab categorization

' ============================================================================
' MAIN ENTRY POINT
' ============================================================================

' ProcessSegmentDocument - Main orchestrator for segment reporting analysis
' Called from ModMain after consolidation document is processed
' Returns: True if successful, False if cancelled or error
Public Function ProcessSegmentDocument() As Boolean
    On Error GoTo ErrorHandler

    Dim segmentWorkbookName As String
    Dim result As VbMsgBoxResult

    ' Ask user if they want to process segment reporting document
    result = MsgBox("Do you have an IAS 8 Operating Segment reporting document?" & vbCrLf & vbCrLf & _
                    "This optional document allows you to:" & vbCrLf & _
                    "• Map packs to their operating segments" & vbCrLf & _
                    "• Analyze scoping coverage by segment" & vbCrLf & _
                    "• Create segment-level reporting in Power BI" & vbCrLf & vbCrLf & _
                    "Click YES if you have a segment document to process." & vbCrLf & _
                    "Click NO to skip segment analysis.", _
                    vbYesNo + vbQuestion, "Segment Reporting Document")

    If result = vbNo Then
        ProcessSegmentDocument = True ' Skip is not an error
        Exit Function
    End If

    ' Step 1: Get segment workbook name
    segmentWorkbookName = GetSegmentWorkbookName()
    If segmentWorkbookName = "" Then
        MsgBox "No segment workbook name provided. Skipping segment analysis.", vbInformation
        ProcessSegmentDocument = True ' Skip is not an error
        Exit Function
    End If

    ' Step 2: Validate and set segment workbook reference
    If Not SetSegmentWorkbook(segmentWorkbookName) Then
        MsgBox "Could not find segment workbook '" & segmentWorkbookName & "'." & vbCrLf & _
               "Please ensure it is open in Excel." & vbCrLf & vbCrLf & _
               "Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True ' Skip is not an error
        Exit Function
    End If

    ' Step 3: Discover segment tabs
    Dim segmentTabList As Collection
    Set segmentTabList = DiscoverSegmentTabs()

    If segmentTabList.Count = 0 Then
        MsgBox "No tabs found in segment workbook. Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True
        Exit Function
    End If

    ' Step 4: Categorize segment tabs
    If Not CategorizeSegmentTabs(segmentTabList) Then
        MsgBox "Segment tab categorization was cancelled. Skipping segment analysis.", vbInformation
        ProcessSegmentDocument = True
        Exit Function
    End If

    ' Step 5: Extract segment pack mappings
    Application.StatusBar = "Extracting segment pack mappings..."
    Dim segmentMappings As Collection
    Set segmentMappings = ExtractSegmentPackMappings()

    If segmentMappings.Count = 0 Then
        MsgBox "No segment pack mappings could be extracted." & vbCrLf & _
               "Please verify the segment document structure." & vbCrLf & vbCrLf & _
               "Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True
        Exit Function
    End If

    ' Step 6: Match segment packs to consolidation packs
    Application.StatusBar = "Matching segment packs to consolidation document..."
    Dim matchedMappings As Collection
    Set matchedMappings = MatchSegmentToConsolidationPacks(segmentMappings)

    If matchedMappings.Count = 0 Then
        MsgBox "Could not match any segment packs to consolidation packs." & vbCrLf & _
               "Please verify pack names and codes match between documents." & vbCrLf & vbCrLf & _
               "Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True
        Exit Function
    End If

    ' Step 7: Create segment analysis tables in output workbook
    Application.StatusBar = "Creating segment analysis tables..."
    CreateSegmentPackMappingTable matchedMappings
    CreateSegmentSummaryTable matchedMappings

    ' Success
    Application.StatusBar = False
    MsgBox "Segment analysis completed successfully!" & vbCrLf & vbCrLf & _
           "Created tables:" & vbCrLf & _
           "• Segment_Pack_Mapping" & vbCrLf & _
           "• Segment_Summary" & vbCrLf & vbCrLf & _
           "These tables can be imported into Power BI for segment-level scoping analysis.", _
           vbInformation, "Segment Analysis Complete"

    ProcessSegmentDocument = True
    Exit Function

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error processing segment document: " & Err.Description & vbCrLf & vbCrLf & _
           "Segment analysis will be skipped.", vbCritical
    ProcessSegmentDocument = True ' Don't fail entire process
End Function

' ============================================================================
' WORKBOOK MANAGEMENT FUNCTIONS
' ============================================================================

' GetSegmentWorkbookName - Prompt user for segment workbook name
' Returns: Workbook name as string, or empty string if cancelled
Private Function GetSegmentWorkbookName() As String
    Dim workbookName As String

    workbookName = InputBox("Enter the name of the SEGMENT REPORTING workbook:" & vbCrLf & vbCrLf & _
                           "Example: ""Bidvest_Segment_Reporting_2024.xlsx""" & vbCrLf & vbCrLf & _
                           "IMPORTANT:" & vbCrLf & _
                           "• The workbook must be OPEN in Excel" & vbCrLf & _
                           "• Enter the EXACT name including extension (.xlsx or .xlsm)" & vbCrLf & _
                           "• This is the document showing IAS 8 operating segments", _
                           "Segment Reporting Workbook Name", "")

    GetSegmentWorkbookName = Trim(workbookName)
End Function

' SetSegmentWorkbook - Validate and set reference to segment workbook
' Parameters:
'   workbookName - Name of segment workbook to find
' Returns: True if found and set, False otherwise
Private Function SetSegmentWorkbook(workbookName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim wb As Workbook

    ' Try to find the workbook
    For Each wb In Application.Workbooks
        If wb.Name = workbookName Then
            Set g_SegmentWorkbook = wb
            SetSegmentWorkbook = True
            Exit Function
        End If
    Next wb

    ' Not found
    SetSegmentWorkbook = False
    Exit Function

ErrorHandler:
    SetSegmentWorkbook = False
End Function

' ============================================================================
' TAB DISCOVERY AND CATEGORIZATION
' ============================================================================

' DiscoverSegmentTabs - Get list of all worksheets in segment workbook
' Returns: Collection of worksheet names
Private Function DiscoverSegmentTabs() As Collection
    On Error GoTo ErrorHandler

    Dim tabList As New Collection
    Dim ws As Worksheet

    For Each ws In g_SegmentWorkbook.Worksheets
        tabList.Add ws.Name
    Next ws

    Set DiscoverSegmentTabs = tabList
    Exit Function

ErrorHandler:
    Set DiscoverSegmentTabs = New Collection
End Function

' CategorizeSegmentTabs - Prompt user to categorize each segment tab
' Parameters:
'   tabList - Collection of tab names to categorize
' Returns: True if successful, False if cancelled
' Tab Categories for Segment Document:
'   1 = Segment Tab (contains pack data for specific segment)
'   2 = Segment Summary (summary/consolidation of all segments)
'   9 = Uncategorized (ignore this tab)
Private Function CategorizeSegmentTabs(tabList As Collection) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim tabName As String
    Dim category As String
    Dim segmentName As String
    Dim msg As String
    Dim categoryDict As Object

    Set g_SegmentTabCategories = CreateObject("Scripting.Dictionary")
    Set categoryDict = CreateObject("Scripting.Dictionary")

    ' Category mapping for segment document
    categoryDict.Add "1", "Segment Tab"
    categoryDict.Add "2", "Segment Summary Tab"
    categoryDict.Add "9", "Uncategorized (Skip)"

    ' Build instruction message
    msg = "SEGMENT TAB CATEGORIZATION" & vbCrLf & vbCrLf
    msg = msg & "For each tab, enter the category number:" & vbCrLf & vbCrLf
    msg = msg & "1 = Segment Tab (contains pack data for a specific segment)" & vbCrLf
    msg = msg & "2 = Segment Summary Tab (summary of all segments)" & vbCrLf
    msg = msg & "9 = Uncategorized (skip this tab)" & vbCrLf & vbCrLf
    msg = msg & "NOTE: For Segment Tabs (1), you will also enter the segment name." & vbCrLf
    msg = msg & "Example segment names: ""Food Services"", ""Freight"", ""Office Products""" & vbCrLf & vbCrLf
    msg = msg & "Click OK to begin categorization."

    MsgBox msg, vbInformation, "Segment Tab Categorization"

    ' Categorize each tab
    For i = 1 To tabList.Count
        tabName = tabList(i)

        ' Prompt for category
        category = InputBox("Tab " & i & " of " & tabList.Count & ": """ & tabName & """" & vbCrLf & vbCrLf & _
                           "Enter category number:" & vbCrLf & _
                           "1 = Segment Tab" & vbCrLf & _
                           "2 = Segment Summary" & vbCrLf & _
                           "9 = Uncategorized" & vbCrLf & vbCrLf & _
                           "Enter category:", "Categorize: " & tabName, "1")

        ' Validate category
        If Not categoryDict.Exists(category) Then
            MsgBox "Invalid category '" & category & "'. Please enter 1, 2, or 9.", vbExclamation
            i = i - 1 ' Retry this tab
            GoTo NextTab
        End If

        ' If Segment Tab, prompt for segment name
        segmentName = ""
        If category = "1" Then
            segmentName = InputBox("Enter the SEGMENT NAME for this tab:" & vbCrLf & vbCrLf & _
                                  "Examples:" & vbCrLf & _
                                  "• Food Services" & vbCrLf & _
                                  "• Freight" & vbCrLf & _
                                  "• Office Products" & vbCrLf & _
                                  "• Automotive" & vbCrLf & vbCrLf & _
                                  "This name will be used in segment analysis.", _
                                  "Segment Name for: " & tabName, "")

            If Trim(segmentName) = "" Then
                MsgBox "Segment name is required for Segment Tabs. Please re-enter.", vbExclamation
                i = i - 1 ' Retry this tab
                GoTo NextTab
            End If
        End If

        ' Store categorization (format: "Category|SegmentName")
        g_SegmentTabCategories.Add tabName, category & "|" & segmentName

NextTab:
    Next i

    CategorizeSegmentTabs = True
    Exit Function

ErrorHandler:
    MsgBox "Error during segment tab categorization: " & Err.Description, vbCritical
    CategorizeSegmentTabs = False
End Function

' ============================================================================
' SEGMENT PACK EXTRACTION
' ============================================================================

' ExtractSegmentPackMappings - Extract pack names/codes from segment tabs
' Returns: Collection of mapping dictionaries
' Each mapping contains: SegmentName, PackNameCode (combined), PackName, PackCode
Private Function ExtractSegmentPackMappings() As Collection
    On Error GoTo ErrorHandler

    Dim mappings As New Collection
    Dim tabName As Variant
    Dim categoryInfo As String
    Dim categoryParts() As String
    Dim category As String
    Dim segmentName As String
    Dim ws As Worksheet

    ' Iterate through categorized tabs
    For Each tabName In g_SegmentTabCategories.Keys
        categoryInfo = g_SegmentTabCategories(tabName)
        categoryParts = Split(categoryInfo, "|")
        category = categoryParts(0)

        ' Process only Segment Tabs (category 1)
        If category = "1" Then
            segmentName = categoryParts(1)
            Set ws = g_SegmentWorkbook.Worksheets(CStr(tabName))

            ' Extract packs from this segment tab
            ExtractPacksFromSegmentTab ws, segmentName, mappings
        End If
    Next tabName

    Set ExtractSegmentPackMappings = mappings
    Exit Function

ErrorHandler:
    MsgBox "Error extracting segment pack mappings: " & Err.Description, vbCritical
    Set ExtractSegmentPackMappings = New Collection
End Function

' ExtractPacksFromSegmentTab - Extract pack info from a single segment tab
' Parameters:
'   ws - Worksheet to process
'   segmentName - Name of the segment
'   mappings - Collection to add mappings to
' Row 8 format: "Top Turf - LS-0714" (PackName - PackCode)
Private Sub ExtractPacksFromSegmentTab(ws As Worksheet, segmentName As String, mappings As Collection)
    On Error GoTo ErrorHandler

    Dim col As Long
    Dim lastCol As Long
    Dim cellValue As String
    Dim packNameCode As String
    Dim parsedPack As Object
    Dim mapping As Object

    ' Find last column with data in row 8
    lastCol = ws.Cells(8, ws.Columns.Count).End(xlToLeft).Column

    ' Scan row 8 for pack entries
    For col = 1 To lastCol
        cellValue = Trim(ws.Cells(8, col).Value)

        ' Skip empty cells
        If cellValue <> "" Then
            ' Parse the pack name and code from "Name - Code" format
            Set parsedPack = ParseSegmentPackNameCode(cellValue)

            If Not parsedPack Is Nothing Then
                ' Create mapping entry
                Set mapping = CreateObject("Scripting.Dictionary")
                mapping("SegmentName") = segmentName
                mapping("PackNameCode") = cellValue ' Original combined format
                mapping("PackName") = parsedPack("PackName")
                mapping("PackCode") = parsedPack("PackCode")
                mapping("ColumnIndex") = col
                mapping("SourceTab") = ws.Name

                ' Add to mappings collection
                mappings.Add mapping
            End If
        End If
    Next col

    Exit Sub

ErrorHandler:
    ' Log error but continue processing other tabs
    Debug.Print "Error extracting packs from segment tab " & ws.Name & ": " & Err.Description
End Sub

' ParseSegmentPackNameCode - Parse "Name - Code" format from segment doc
' Parameters:
'   combined - String in format "Top Turf - LS-0714"
' Returns: Dictionary with PackName and PackCode, or Nothing if parse fails
Private Function ParseSegmentPackNameCode(combined As String) As Object
    On Error GoTo ErrorHandler

    Dim parts() As String
    Dim packName As String
    Dim packCode As String
    Dim result As Object

    ' Check if contains separator " - "
    If InStr(combined, " - ") > 0 Then
        parts = Split(combined, " - ")

        If UBound(parts) >= 1 Then
            packName = Trim(parts(0))
            packCode = Trim(parts(1))

            ' Create result dictionary
            Set result = CreateObject("Scripting.Dictionary")
            result("PackName") = packName
            result("PackCode") = packCode

            Set ParseSegmentPackNameCode = result
            Exit Function
        End If
    End If

    ' If no separator found, try alternative formats
    ' Format: "LS-0714" only (code only)
    If Len(combined) > 0 And InStr(combined, "-") > 0 Then
        Set result = CreateObject("Scripting.Dictionary")
        result("PackName") = combined ' Use full value as name
        result("PackCode") = combined ' Use full value as code
        Set ParseSegmentPackNameCode = result
        Exit Function
    End If

    ' Parse failed
    Set ParseSegmentPackNameCode = Nothing
    Exit Function

ErrorHandler:
    Set ParseSegmentPackNameCode = Nothing
End Function

' ============================================================================
' MATCHING LOGIC: Segment Document → Consolidation Document
' ============================================================================

' MatchSegmentToConsolidationPacks - Match segment packs to consolidation packs
' Parameters:
'   segmentMappings - Collection of segment pack mappings
' Returns: Collection of matched mappings with consolidation pack info added
' Matching Strategy:
'   1. Try exact Pack Code match
'   2. Try Pack Code with/without spaces or dashes
'   3. Try Pack Name fuzzy match
'   4. Try combined Name-Code match
Private Function MatchSegmentToConsolidationPacks(segmentMappings As Collection) As Collection
    On Error GoTo ErrorHandler

    Dim matchedMappings As New Collection
    Dim consolidationPacks As Object ' Dictionary of consolidation packs
    Dim i As Long
    Dim mapping As Object
    Dim matchedPack As Object
    Dim matchCount As Long
    Dim unmatchedCount As Long

    ' Build dictionary of consolidation packs for fast lookup
    Set consolidationPacks = BuildConsolidationPacksDictionary()

    If consolidationPacks.Count = 0 Then
        MsgBox "No consolidation packs available for matching." & vbCrLf & _
               "Please ensure consolidation document was processed first.", vbExclamation
        Set MatchSegmentToConsolidationPacks = matchedMappings
        Exit Function
    End If

    ' Match each segment pack to consolidation pack
    matchCount = 0
    unmatchedCount = 0

    For i = 1 To segmentMappings.Count
        Set mapping = segmentMappings(i)

        ' Try to find match in consolidation
        Set matchedPack = FindConsolidationPackMatch(mapping, consolidationPacks)

        If Not matchedPack Is Nothing Then
            ' Add matched consolidation pack info to mapping
            mapping("ConsolPackName") = matchedPack("PackName")
            mapping("ConsolPackCode") = matchedPack("PackCode")
            mapping("ConsolDivision") = matchedPack("Division")
            mapping("MatchMethod") = matchedPack("MatchMethod")
            mapping("IsMatched") = True
            matchCount = matchCount + 1
        Else
            ' No match found
            mapping("ConsolPackName") = "[Not Matched]"
            mapping("ConsolPackCode") = "[Not Matched]"
            mapping("ConsolDivision") = ""
            mapping("MatchMethod") = "No Match"
            mapping("IsMatched") = False
            unmatchedCount = unmatchedCount + 1
        End If

        matchedMappings.Add mapping
    Next i

    ' Report matching statistics
    MsgBox "Segment Pack Matching Results:" & vbCrLf & vbCrLf & _
           "Total segment packs: " & segmentMappings.Count & vbCrLf & _
           "Successfully matched: " & matchCount & vbCrLf & _
           "Unmatched: " & unmatchedCount & vbCrLf & vbCrLf & _
           IIf(unmatchedCount > 0, "Review unmatched packs in Segment_Pack_Mapping table.", "All packs matched successfully!"), _
           IIf(unmatchedCount = 0, vbInformation, vbExclamation), _
           "Matching Results"

    Set MatchSegmentToConsolidationPacks = matchedMappings
    Exit Function

ErrorHandler:
    MsgBox "Error matching segment packs: " & Err.Description, vbCritical
    Set MatchSegmentToConsolidationPacks = New Collection
End Function

' BuildConsolidationPacksDictionary - Build dictionary of consolidation packs
' Returns: Dictionary with pack codes/names from Pack Number Company Table
Private Function BuildConsolidationPacksDictionary() As Object
    On Error GoTo ErrorHandler

    Dim packsDict As Object
    Set packsDict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim packName As String
    Dim packCode As String
    Dim division As String
    Dim packInfo As Object

    ' Find Pack Number Company Table in output workbook
    On Error Resume Next
    Set ws = g_OutputWorkbook.Worksheets("Pack Number Company Table")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Set BuildConsolidationPacksDictionary = packsDict
        Exit Function
    End If

    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Start from row 2 (skip header)
    For row = 2 To lastRow
        packName = Trim(ws.Cells(row, 1).Value) ' Column A: Pack Name
        packCode = Trim(ws.Cells(row, 2).Value) ' Column B: Pack Code
        division = Trim(ws.Cells(row, 3).Value) ' Column C: Division

        If packCode <> "" Then
            ' Store pack info
            Set packInfo = CreateObject("Scripting.Dictionary")
            packInfo("PackName") = packName
            packInfo("PackCode") = packCode
            packInfo("Division") = division

            ' Add with multiple keys for flexible matching
            If Not packsDict.Exists(packCode) Then
                packsDict.Add packCode, packInfo
            End If

            ' Also add normalized version (no spaces, dashes)
            Dim normalizedCode As String
            normalizedCode = NormalizePackCode(packCode)
            If Not packsDict.Exists(normalizedCode) And normalizedCode <> packCode Then
                packsDict.Add normalizedCode, packInfo
            End If

            ' Add by pack name for fuzzy matching
            Dim normalizedName As String
            normalizedName = NormalizePackName(packName)
            If Not packsDict.Exists(normalizedName) Then
                packsDict.Add normalizedName, packInfo
            End If
        End If
    Next row

    Set BuildConsolidationPacksDictionary = packsDict
    Exit Function

ErrorHandler:
    Set BuildConsolidationPacksDictionary = CreateObject("Scripting.Dictionary")
End Function

' FindConsolidationPackMatch - Find matching consolidation pack
' Parameters:
'   segmentMapping - Dictionary with segment pack info
'   consolidationPacks - Dictionary of consolidation packs
' Returns: Matched pack dictionary or Nothing
Private Function FindConsolidationPackMatch(segmentMapping As Object, consolidationPacks As Object) As Object
    On Error GoTo ErrorHandler

    Dim packCode As String
    Dim packName As String
    Dim matchedPack As Object
    Dim normalizedCode As String
    Dim normalizedName As String

    packCode = segmentMapping("PackCode")
    packName = segmentMapping("PackName")

    ' Strategy 1: Exact pack code match
    If consolidationPacks.Exists(packCode) Then
        Set matchedPack = consolidationPacks(packCode)
        matchedPack("MatchMethod") = "Exact Code"
        Set FindConsolidationPackMatch = matchedPack
        Exit Function
    End If

    ' Strategy 2: Normalized pack code match (remove spaces, dashes)
    normalizedCode = NormalizePackCode(packCode)
    If consolidationPacks.Exists(normalizedCode) Then
        Set matchedPack = consolidationPacks(normalizedCode)
        matchedPack("MatchMethod") = "Normalized Code"
        Set FindConsolidationPackMatch = matchedPack
        Exit Function
    End If

    ' Strategy 3: Pack name match
    normalizedName = NormalizePackName(packName)
    If consolidationPacks.Exists(normalizedName) Then
        Set matchedPack = consolidationPacks(normalizedName)
        matchedPack("MatchMethod") = "Pack Name"
        Set FindConsolidationPackMatch = matchedPack
        Exit Function
    End If

    ' Strategy 4: Fuzzy match on pack name (contains)
    Dim key As Variant
    For Each key In consolidationPacks.Keys
        If InStr(1, CStr(key), normalizedName, vbTextCompare) > 0 Or _
           InStr(1, normalizedName, CStr(key), vbTextCompare) > 0 Then
            Set matchedPack = consolidationPacks(key)
            matchedPack("MatchMethod") = "Fuzzy Name"
            Set FindConsolidationPackMatch = matchedPack
            Exit Function
        End If
    Next key

    ' No match found
    Set FindConsolidationPackMatch = Nothing
    Exit Function

ErrorHandler:
    Set FindConsolidationPackMatch = Nothing
End Function

' NormalizePackCode - Normalize pack code for matching
' Parameters:
'   packCode - Original pack code
' Returns: Normalized code (uppercase, no spaces/dashes)
Private Function NormalizePackCode(packCode As String) As String
    Dim normalized As String
    normalized = UCase(Trim(packCode))
    normalized = Replace(normalized, " ", "")
    normalized = Replace(normalized, "-", "")
    normalized = Replace(normalized, "_", "")
    NormalizePackCode = normalized
End Function

' NormalizePackName - Normalize pack name for matching
' Parameters:
'   packName - Original pack name
' Returns: Normalized name (uppercase, no extra spaces)
Private Function NormalizePackName(packName As String) As String
    Dim normalized As String
    normalized = UCase(Trim(packName))
    ' Remove multiple spaces
    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop
    NormalizePackName = normalized
End Function

' ============================================================================
' OUTPUT TABLE GENERATION
' ============================================================================

' CreateSegmentPackMappingTable - Create Segment_Pack_Mapping table
' Parameters:
'   matchedMappings - Collection of matched segment-to-consolidation mappings
Private Sub CreateSegmentPackMappingTable(matchedMappings As Collection)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim row As Long
    Dim i As Long
    Dim mapping As Object
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long

    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Segment_Pack_Mapping"

    ' Write headers
    ws.Cells(1, 1).Value = "Segment Name"
    ws.Cells(1, 2).Value = "Pack Name (Segment Doc)"
    ws.Cells(1, 3).Value = "Pack Code (Segment Doc)"
    ws.Cells(1, 4).Value = "Pack Name (Consol Doc)"
    ws.Cells(1, 5).Value = "Pack Code"
    ws.Cells(1, 6).Value = "Division"
    ws.Cells(1, 7).Value = "Match Status"
    ws.Cells(1, 8).Value = "Match Method"
    ws.Cells(1, 9).Value = "Source Tab"

    ' Write data
    row = 2
    For i = 1 To matchedMappings.Count
        Set mapping = matchedMappings(i)

        ws.Cells(row, 1).Value = mapping("SegmentName")
        ws.Cells(row, 2).Value = mapping("PackName")
        ws.Cells(row, 3).Value = mapping("PackCode")
        ws.Cells(row, 4).Value = mapping("ConsolPackName")
        ws.Cells(row, 5).Value = mapping("ConsolPackCode")
        ws.Cells(row, 6).Value = mapping("ConsolDivision")
        ws.Cells(row, 7).Value = IIf(mapping("IsMatched"), "Matched", "Unmatched")
        ws.Cells(row, 8).Value = mapping("MatchMethod")
        ws.Cells(row, 9).Value = mapping("SourceTab")

        row = row + 1
    Next i

    ' Get dimensions
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 9

    ' Create Excel Table
    If lastRow > 1 Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        tbl.Name = "Segment_Pack_Mapping"
        tbl.TableStyle = "TableStyleMedium2"
    End If

    ' Auto-fit columns
    ws.Columns("A:I").AutoFit

    ' Conditional formatting for Match Status
    With ws.Range("G2:G" & lastRow)
        .FormatConditions.Add Type:=xlTextString, String:="Matched", TextOperator:=xlContains
        .FormatConditions(1).Interior.Color = RGB(200, 255, 200) ' Light green

        .FormatConditions.Add Type:=xlTextString, String:="Unmatched", TextOperator:=xlContains
        .FormatConditions(2).Interior.Color = RGB(255, 200, 200) ' Light red
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Segment Pack Mapping table: " & Err.Description, vbCritical
End Sub

' CreateSegmentSummaryTable - Create Segment_Summary table with statistics
' Parameters:
'   matchedMappings - Collection of matched mappings
Private Sub CreateSegmentSummaryTable(matchedMappings As Collection)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim segmentStats As Object ' Dictionary to collect stats per segment
    Dim mapping As Object
    Dim i As Long
    Dim segmentName As String
    Dim stats As Object
    Dim row As Long
    Dim tbl As ListObject
    Dim lastRow As Long

    ' Build statistics dictionary
    Set segmentStats = CreateObject("Scripting.Dictionary")

    For i = 1 To matchedMappings.Count
        Set mapping = matchedMappings(i)
        segmentName = mapping("SegmentName")

        ' Initialize stats for this segment if not exists
        If Not segmentStats.Exists(segmentName) Then
            Set stats = CreateObject("Scripting.Dictionary")
            stats("TotalPacks") = 0
            stats("MatchedPacks") = 0
            stats("UnmatchedPacks") = 0
            segmentStats.Add segmentName, stats
        End If

        ' Update stats
        Set stats = segmentStats(segmentName)
        stats("TotalPacks") = stats("TotalPacks") + 1
        If mapping("IsMatched") Then
            stats("MatchedPacks") = stats("MatchedPacks") + 1
        Else
            stats("UnmatchedPacks") = stats("UnmatchedPacks") + 1
        End If
    Next i

    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Segment_Summary"

    ' Write headers
    ws.Cells(1, 1).Value = "Segment Name"
    ws.Cells(1, 2).Value = "Total Packs"
    ws.Cells(1, 3).Value = "Matched Packs"
    ws.Cells(1, 4).Value = "Unmatched Packs"
    ws.Cells(1, 5).Value = "Match Rate %"

    ' Write data
    row = 2
    Dim segment As Variant
    For Each segment In segmentStats.Keys
        Set stats = segmentStats(segment)

        ws.Cells(row, 1).Value = segment
        ws.Cells(row, 2).Value = stats("TotalPacks")
        ws.Cells(row, 3).Value = stats("MatchedPacks")
        ws.Cells(row, 4).Value = stats("UnmatchedPacks")

        ' Calculate match rate
        If stats("TotalPacks") > 0 Then
            ws.Cells(row, 5).Value = stats("MatchedPacks") / stats("TotalPacks")
            ws.Cells(row, 5).NumberFormat = "0.0%"
        Else
            ws.Cells(row, 5).Value = 0
        End If

        row = row + 1
    Next segment

    ' Get dimensions
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Create Excel Table
    If lastRow > 1 Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 5)), , xlYes)
        tbl.Name = "Segment_Summary"
        tbl.TableStyle = "TableStyleMedium2"
    End If

    ' Auto-fit columns
    ws.Columns("A:E").AutoFit

    ' Add totals row
    row = lastRow + 2
    ws.Cells(row, 1).Value = "TOTAL"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Formula = "=SUM(B2:B" & lastRow & ")"
    ws.Cells(row, 3).Formula = "=SUM(C2:C" & lastRow & ")"
    ws.Cells(row, 4).Formula = "=SUM(D2:D" & lastRow & ")"
    ws.Cells(row, 5).Formula = "=C" & row & "/B" & row
    ws.Cells(row, 5).NumberFormat = "0.0%"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Segment Summary table: " & Err.Description, vbCritical
End Sub
