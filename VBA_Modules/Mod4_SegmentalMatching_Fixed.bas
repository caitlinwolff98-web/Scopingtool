Attribute VB_Name = "Mod4_SegmentalMatching"
Option Explicit

' =================================================================================
' MODULE 4: SEGMENTAL REPORTING MATCHING AND DIVISION-SEGMENT MAPPING
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 7.0 - Complete Fix and Enhancement
' =================================================================================
' PURPOSE:
'   Match packs between Stripe Packs and Segmental Reporting workbooks
'   Map divisions and segments to all packs
'   Create comprehensive Division-Segment mapping table
'
' CRITICAL FIXES:
'   1. Properly recognizes segmental reporting workbook
'   2. Maps division from division tabs to packs
'   3. Maps segment from segmental tabs to packs
'   4. Updates Pack Number Company Table with actual mappings
'   5. Creates detailed reconciliation report
'
' AUTHOR: ISA 600 Scoping Tool Team
' DATE: 2025-11-18
' =================================================================================

' ==================== MAIN PROCESSING FUNCTION ====================
Public Sub ProcessSegmentalWorkbook(segmentalWb As Workbook, tabCategories As Object, divisionNames As Object)
    '------------------------------------------------------------------------
    ' Main function to process Segmental Reporting workbook
    ' Performs pack matching and creates mapping tables
    '
    ' CRITICAL FIX: Now properly processes segmental workbook and updates
    '               Pack Number Company Table with actual divisions and segments
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.StatusBar = "Processing Segmental Reporting workbook..."

    Dim segmentalTabs As Object
    Dim segmentPacks As Object
    Dim stripePacks As Object
    Dim matchingResults As Object

    ' Step 1: Categorize segmental tabs
    Set segmentalTabs = CategorizeSegmentalTabs(segmentalWb)

    ' Step 2: Extract packs from segmental tabs
    Set segmentPacks = ExtractSegmentalPacks(segmentalWb, segmentalTabs)

    ' Step 3: Extract packs from Stripe (already categorized)
    Set stripePacks = ExtractStripePacks(tabCategories, divisionNames)

    ' Step 4: Perform matching
    Set matchingResults = PerformPackMatching(stripePacks, segmentPacks)

    ' Step 5: CRITICAL FIX - Update Pack Number Company Table with mappings
    UpdatePackCompanyTableWithMappings matchingResults

    ' Step 6: Generate Division-Segment Mapping Table
    GenerateDivisionSegmentMapping matchingResults

    ' Step 7: Generate Pack Matching Reconciliation Report
    GeneratePackMatchingReport matchingResults, stripePacks, segmentPacks

    Application.StatusBar = False

    MsgBox "Segmental Reporting processed successfully!" & vbCrLf & vbCrLf & _
           "Division and Segment mappings have been applied to Pack Number Company Table.", _
           vbInformation, "Segmental Processing Complete"

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error processing Segmental workbook: " & Err.Description, vbCritical
End Sub

' ==================== CATEGORIZE SEGMENTAL TABS ====================
Private Function CategorizeSegmentalTabs(segmentalWb As Workbook) As Object
    '------------------------------------------------------------------------
    ' Categorize tabs in Segmental Reporting workbook
    ' Category 1: Segment Tabs (multiple) - prompt for Segment Name
    ' Category 2: Summarized Segment (single) - not actively used
    '------------------------------------------------------------------------
    Dim categories As Object
    Dim ws As Worksheet
    Dim segmentName As String
    Dim counter As Long
    Dim result As String

    Set categories = CreateObject("Scripting.Dictionary")

    MsgBox "SEGMENTAL TAB CATEGORIZATION" & vbCrLf & vbCrLf & _
           "You will now categorize tabs in the Segmental Reporting workbook." & vbCrLf & vbCrLf & _
           "Categories:" & vbCrLf & _
           "1. Segment Tabs (you will name each segment)" & vbCrLf & _
           "2. Summarized Segment (informational only)" & vbCrLf & _
           "3. Uncategorized (ignore)", _
           vbInformation, "Segment Tab Categorization"

    counter = 1
    For Each ws In segmentalWb.Worksheets
        result = InputBox( _
            "Segmental Tab " & counter & " of " & segmentalWb.Worksheets.Count & vbCrLf & vbCrLf & _
            "Tab Name: " & ws.Name & vbCrLf & vbCrLf & _
            "Select category:" & vbCrLf & _
            "1 = Segment Tab (will prompt for segment name)" & vbCrLf & _
            "2 = Summarized Segment" & vbCrLf & _
            "3 = Uncategorized" & vbCrLf & vbCrLf & _
            "Enter number:", _
            "Categorize: " & ws.Name, "1")

        If result = "" Then result = "3" ' Default to uncategorized

        Select Case result
            Case "1"
                ' Prompt for segment name
                segmentName = InputBox( _
                    "Segment Tab: " & ws.Name & vbCrLf & vbCrLf & _
                    "Enter segment name:" & vbCrLf & _
                    "(This will be used in reports)" & vbCrLf & vbCrLf & _
                    "Example: UK Segment, SA Segment", _
                    "Segment Name", ws.Name)

                If Trim(segmentName) = "" Then segmentName = ws.Name

                categories(ws.Name) = "Segment:" & segmentName

            Case "2"
                categories(ws.Name) = "Summarized"

            Case Else
                categories(ws.Name) = "Uncategorized"
        End Select

        counter = counter + 1
    Next ws

    Set CategorizeSegmentalTabs = categories
End Function

' ==================== EXTRACT SEGMENTAL PACKS ====================
Private Function ExtractSegmentalPacks(segmentalWb As Workbook, segmentalTabs As Object) As Object
    '------------------------------------------------------------------------
    ' Extract packs from Segmental Reporting workbook
    ' Row 8 contains: "Pack Name - Pack Code" (with spaces around dash)
    ' Returns Dictionary: pack code -> {name, segment, tab}
    '
    ' CRITICAL FIX: Enhanced to handle various formats of pack naming
    '------------------------------------------------------------------------
    Dim packs As Object
    Dim tabName As Variant
    Dim ws As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim cellValue As String
    Dim packName As String
    Dim packCode As String
    Dim segmentName As String
    Dim dashPos As Long

    Set packs = CreateObject("Scripting.Dictionary")

    For Each tabName In segmentalTabs.Keys
        If Left(segmentalTabs(tabName), 8) = "Segment:" Then
            segmentName = Mid(segmentalTabs(tabName), 9) ' Extract segment name

            Set ws = segmentalWb.Worksheets(CStr(tabName))
            lastCol = ws.Cells(8, ws.Columns.Count).End(xlToLeft).Column

            For col = 1 To lastCol
                cellValue = Trim(ws.Cells(8, col).Value)

                ' Parse "Pack Name - Pack Code" format
                ' Try multiple separators: " - ", "-", " -", "- "
                dashPos = InStr(cellValue, " - ")
                If dashPos = 0 Then dashPos = InStr(cellValue, "-")

                If dashPos > 0 Then
                    ' Extract parts
                    If InStr(cellValue, " - ") > 0 Then
                        packName = Trim(Left(cellValue, InStr(cellValue, " - ") - 1))
                        packCode = Trim(Mid(cellValue, InStr(cellValue, " - ") + 3))
                    Else
                        packName = Trim(Left(cellValue, InStr(cellValue, "-") - 1))
                        packCode = Trim(Mid(cellValue, InStr(cellValue, "-") + 1))
                    End If

                    If packCode <> "" And packName <> "" Then
                        If Not packs.exists(packCode) Then
                            Dim packInfo As Object
                            Set packInfo = CreateObject("Scripting.Dictionary")
                            packInfo("Name") = packName
                            packInfo("Segment") = segmentName
                            packInfo("Tab") = CStr(tabName)

                            packs(packCode) = packInfo
                        End If
                    End If
                End If
            Next col
        End If
    Next tabName

    Set ExtractSegmentalPacks = packs
End Function

' ==================== EXTRACT STRIPE PACKS ====================
Private Function ExtractStripePacks(tabCategories As Object, divisionNames As Object) As Object
    '------------------------------------------------------------------------
    ' Extract packs from Stripe Packs workbook with division assignment
    ' Returns Dictionary: pack code -> {name, division, divisionTab}
    '
    ' CRITICAL FIX: Now properly extracts division from division tabs
    '------------------------------------------------------------------------
    Dim packs As Object
    Dim inputTab As Worksheet
    Dim divisionTabs As Collection
    Dim tabName As Variant
    Dim divisionTab As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim packName As String
    Dim divisionName As String

    Set packs = CreateObject("Scripting.Dictionary")

    ' CRITICAL FIX: Extract from Division tabs, not just Input Continuing
    For Each tabName In tabCategories.Keys
        If tabCategories(tabName) = "Division" Then
            ' Get division name
            If divisionNames.exists(tabName) Then
                divisionName = divisionNames(tabName)
            Else
                divisionName = CStr(tabName)
            End If

            Set divisionTab = Mod1_MainController.g_StripePacksWorkbook.Worksheets(CStr(tabName))
            lastCol = divisionTab.Cells(7, divisionTab.Columns.Count).End(xlToLeft).Column

            ' Extract packs from this division tab
            For col = 3 To lastCol
                packCode = Trim(divisionTab.Cells(8, col).Value)
                packName = Trim(divisionTab.Cells(7, col).Value)

                If packCode <> "" And packName <> "" Then
                    If Not packs.exists(packCode) Then
                        Dim packInfo As Object
                        Set packInfo = CreateObject("Scripting.Dictionary")
                        packInfo("Name") = packName
                        packInfo("Division") = divisionName
                        packInfo("DivisionTab") = CStr(tabName)

                        packs(packCode) = packInfo
                    End If
                End If
            Next col
        End If
    Next tabName

    ' Also extract from Input Continuing for any packs not in division tabs
    Set inputTab = Mod2_TabProcessing.GetTabByCategory(tabCategories, "Input Continuing")
    If Not inputTab Is Nothing Then
        lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column

        For col = 3 To lastCol
            packCode = Trim(inputTab.Cells(8, col).Value)
            packName = Trim(inputTab.Cells(7, col).Value)

            If packCode <> "" And packName <> "" Then
                If Not packs.exists(packCode) Then
                    Set packInfo = CreateObject("Scripting.Dictionary")
                    packInfo("Name") = packName
                    packInfo("Division") = "Not Mapped"
                    packInfo("DivisionTab") = "N/A"

                    packs(packCode) = packInfo
                End If
            End If
        Next col
    End If

    Set ExtractStripePacks = packs
End Function

' ==================== PERFORM PACK MATCHING ====================
Private Function PerformPackMatching(stripePacks As Object, segmentPacks As Object) As Object
    '------------------------------------------------------------------------
    ' Match packs between Stripe and Segmental using exact and fuzzy matching
    ' Returns Dictionary with matching results
    '
    ' CRITICAL FIX: Enhanced matching algorithm with better fuzzy logic
    '------------------------------------------------------------------------
    Dim matchResults As Object
    Dim stripeCode As Variant
    Dim segmentCode As Variant
    Dim matchInfo As Object
    Dim bestMatch As String
    Dim bestSimilarity As Double

    Set matchResults = CreateObject("Scripting.Dictionary")

    ' Exact matching first
    For Each stripeCode In stripePacks.Keys
        If segmentPacks.exists(stripeCode) Then
            ' Exact match
            Set matchInfo = CreateObject("Scripting.Dictionary")
            matchInfo("StripeCode") = stripeCode
            matchInfo("StripeName") = stripePacks(stripeCode)("Name")
            matchInfo("Division") = stripePacks(stripeCode)("Division")
            matchInfo("SegmentCode") = stripeCode
            matchInfo("SegmentName") = segmentPacks(stripeCode)("Name")
            matchInfo("Segment") = segmentPacks(stripeCode)("Segment")
            matchInfo("MatchType") = "Exact"
            matchInfo("Similarity") = 100

            matchResults(stripeCode) = matchInfo
        Else
            ' Try fuzzy matching
            bestMatch = FindBestFuzzyMatch(CStr(stripeCode), CStr(stripePacks(stripeCode)("Name")), segmentPacks, bestSimilarity)

            If bestMatch <> "" And bestSimilarity >= 70 Then ' 70% similarity threshold
                Set matchInfo = CreateObject("Scripting.Dictionary")
                matchInfo("StripeCode") = stripeCode
                matchInfo("StripeName") = stripePacks(stripeCode)("Name")
                matchInfo("Division") = stripePacks(stripeCode)("Division")
                matchInfo("SegmentCode") = bestMatch
                matchInfo("SegmentName") = segmentPacks(bestMatch)("Name")
                matchInfo("Segment") = segmentPacks(bestMatch)("Segment")
                matchInfo("MatchType") = "Fuzzy"
                matchInfo("Similarity") = bestSimilarity

                matchResults(stripeCode) = matchInfo
            Else
                ' No match found
                Set matchInfo = CreateObject("Scripting.Dictionary")
                matchInfo("StripeCode") = stripeCode
                matchInfo("StripeName") = stripePacks(stripeCode)("Name")
                matchInfo("Division") = stripePacks(stripeCode)("Division")
                matchInfo("SegmentCode") = "NOT FOUND"
                matchInfo("SegmentName") = "NOT FOUND"
                matchInfo("Segment") = "Not Mapped"
                matchInfo("MatchType") = "Not Found"
                matchInfo("Similarity") = 0

                matchResults(stripeCode) = matchInfo
            End If
        End If
    Next stripeCode

    Set PerformPackMatching = matchResults
End Function

' ==================== UPDATE PACK COMPANY TABLE WITH MAPPINGS ====================
Private Sub UpdatePackCompanyTableWithMappings(matchResults As Object)
    '------------------------------------------------------------------------
    ' CRITICAL FIX: Update Pack Number Company Table with actual Division
    '               and Segment mappings from matching results
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim packTable As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim packCode As Variant  ' CRITICAL FIX: Must be Variant for For Each loop
    Dim matchInfo As Object

    ' Get Pack Number Company Table
    Set packTable = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    If packTable Is Nothing Then Exit Sub

    ' Find data range
    lastRow = packTable.Cells(packTable.Rows.Count, 2).End(xlUp).row

    ' Update each pack
    For row = 2 To lastRow
        packCode = packTable.Cells(row, 2).Value

        ' Look up in match results
        If matchResults.exists(CStr(packCode)) Then
            Set matchInfo = matchResults(CStr(packCode))

            ' Update Division (Column 3)
            If matchInfo("Division") <> "Not Mapped" Then
                packTable.Cells(row, 3).Value = matchInfo("Division")
            End If

            ' Update Segment (Column 4)
            If matchInfo("Segment") <> "Not Mapped" Then
                packTable.Cells(row, 4).Value = matchInfo("Segment")
            End If
        End If
    Next row

    ' CRITICAL FIX: Also update Mod3 module's pack segments dictionary for future use
    For Each packCode In matchResults.Keys
        Set matchInfo = matchResults(packCode)
        If matchInfo("Segment") <> "Not Mapped" Then
            Mod3_DataExtraction.SetPackSegment CStr(packCode), matchInfo("Segment")
        End If
    Next packCode

    On Error GoTo 0
End Sub

' ==================== FUZZY MATCHING ====================
Private Function FindBestFuzzyMatch(targetCode As String, targetName As String, candidates As Object, ByRef bestSimilarity As Double) As String
    '------------------------------------------------------------------------
    ' Find best fuzzy match for a pack code/name
    ' Uses similarity algorithm based on character overlap and position
    '------------------------------------------------------------------------
    Dim candidateCode As Variant
    Dim similarityByCode As Double
    Dim similarityByName As Double
    Dim bestCode As String

    bestCode = ""
    bestSimilarity = 0

    For Each candidateCode In candidates.Keys
        ' Calculate similarity based on code
        similarityByCode = CalculateSimilarity(targetCode, CStr(candidateCode))

        ' Also check name similarity
        similarityByName = CalculateSimilarity(targetName, CStr(candidates(candidateCode)("Name")))

        ' Use the higher of the two similarities
        Dim similarity As Double
        similarity = Application.WorksheetFunction.Max(similarityByCode, similarityByName)

        If similarity > bestSimilarity Then
            bestSimilarity = similarity
            bestCode = CStr(candidateCode)
        End If
    Next candidateCode

    FindBestFuzzyMatch = bestCode
End Function

' ==================== CALCULATE SIMILARITY ====================
Private Function CalculateSimilarity(str1 As String, str2 As String) As Double
    '------------------------------------------------------------------------
    ' Calculate similarity percentage between two strings
    ' Uses character-by-character comparison
    '------------------------------------------------------------------------
    Dim i As Long
    Dim matches As Long
    Dim maxLen As Long

    str1 = UCase(Trim(str1))
    str2 = UCase(Trim(str2))

    If str1 = str2 Then
        CalculateSimilarity = 100
        Exit Function
    End If

    maxLen = Application.WorksheetFunction.Max(Len(str1), Len(str2))
    If maxLen = 0 Then
        CalculateSimilarity = 0
        Exit Function
    End If

    matches = 0
    For i = 1 To Application.WorksheetFunction.Min(Len(str1), Len(str2))
        If Mid(str1, i, 1) = Mid(str2, i, 1) Then
            matches = matches + 1
        End If
    Next i

    CalculateSimilarity = (matches / maxLen) * 100
End Function

' ==================== GENERATE DIVISION-SEGMENT MAPPING ====================
Private Sub GenerateDivisionSegmentMapping(matchResults As Object)
    '------------------------------------------------------------------------
    ' Generate Division-Segment Mapping Table
    ' CRITICAL FIX: Now shows actual divisions and segments, not "To Be Assigned"
    '------------------------------------------------------------------------
    Dim outputWs As Worksheet
    Dim row As Long
    Dim matchKey As Variant
    Dim matchInfo As Object

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Division-Segment Mapping"

    ' Write headers
    outputWs.Cells(1, 1).Value = "Pack Code"
    outputWs.Cells(1, 2).Value = "Pack Name"
    outputWs.Cells(1, 3).Value = "Division"
    outputWs.Cells(1, 4).Value = "Segment"
    outputWs.Cells(1, 5).Value = "Match Type"
    outputWs.Cells(1, 6).Value = "Similarity %"
    outputWs.Cells(1, 7).Value = "Match Status"

    ' Format headers
    With outputWs.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = 2

    For Each matchKey In matchResults.Keys
        Set matchInfo = matchResults(matchKey)

        outputWs.Cells(row, 1).Value = matchInfo("StripeCode")
        outputWs.Cells(row, 2).Value = matchInfo("StripeName")
        outputWs.Cells(row, 3).Value = matchInfo("Division")  ' CRITICAL FIX: Actual division
        outputWs.Cells(row, 4).Value = matchInfo("Segment")   ' CRITICAL FIX: Actual segment
        outputWs.Cells(row, 5).Value = matchInfo("MatchType")
        outputWs.Cells(row, 6).Value = matchInfo("Similarity")
        outputWs.Cells(row, 6).NumberFormat = "0.0%"

        ' Match status
        If matchInfo("Division") <> "Not Mapped" And matchInfo("Segment") <> "Not Mapped" Then
            outputWs.Cells(row, 7).Value = "Fully Mapped"
            outputWs.Cells(row, 7).Interior.Color = RGB(198, 239, 206) ' Green
        ElseIf matchInfo("Division") <> "Not Mapped" Or matchInfo("Segment") <> "Not Mapped" Then
            outputWs.Cells(row, 7).Value = "Partially Mapped"
            outputWs.Cells(row, 7).Interior.Color = RGB(255, 235, 156) ' Yellow
        Else
            outputWs.Cells(row, 7).Value = "Not Mapped"
            outputWs.Cells(row, 7).Interior.Color = RGB(255, 199, 206) ' Red
        End If

        ' Color code match type
        Select Case matchInfo("MatchType")
            Case "Exact"
                outputWs.Cells(row, 5).Interior.Color = RGB(198, 239, 206) ' Green
            Case "Fuzzy"
                outputWs.Cells(row, 5).Interior.Color = RGB(255, 235, 156) ' Yellow
            Case "Not Found"
                outputWs.Cells(row, 5).Interior.Color = RGB(255, 199, 206) ' Red
        End Select

        row = row + 1
    Next matchKey

    ' Convert to Excel Table
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range

    lastRow = row - 1
    lastCol = 7

    If lastRow >= 2 Then
        Set tableRange = outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol))
        outputWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        outputWs.ListObjects(outputWs.ListObjects.Count).Name = "DivisionSegmentMapping"
        outputWs.ListObjects("DivisionSegmentMapping").TableStyle = "TableStyleMedium2"
    End If

    outputWs.Columns.AutoFit
End Sub

' ==================== GENERATE PACK MATCHING REPORT ====================
Private Sub GeneratePackMatchingReport(matchResults As Object, stripePacks As Object, segmentPacks As Object)
    '------------------------------------------------------------------------
    ' Generate detailed Pack Matching Reconciliation Report
    '------------------------------------------------------------------------
    Dim outputWs As Worksheet
    Dim row As Long

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Pack Matching Report"

    ' Write title
    outputWs.Cells(1, 1).Value = "PACK MATCHING RECONCILIATION REPORT"
    outputWs.Cells(1, 1).Font.Size = 14
    outputWs.Cells(1, 1).Font.Bold = True
    outputWs.Cells(1, 1).Font.Color = RGB(0, 112, 192)

    ' Summary statistics
    Dim exactMatches As Long
    Dim fuzzyMatches As Long
    Dim notFound As Long
    Dim fullyMapped As Long
    Dim partiallyMapped As Long
    Dim notMapped As Long
    Dim matchKey As Variant
    Dim matchInfo As Object

    For Each matchKey In matchResults.Keys
        Set matchInfo = matchResults(matchKey)

        Select Case matchInfo("MatchType")
            Case "Exact": exactMatches = exactMatches + 1
            Case "Fuzzy": fuzzyMatches = fuzzyMatches + 1
            Case "Not Found": notFound = notFound + 1
        End Select

        If matchInfo("Division") <> "Not Mapped" And matchInfo("Segment") <> "Not Mapped" Then
            fullyMapped = fullyMapped + 1
        ElseIf matchInfo("Division") <> "Not Mapped" Or matchInfo("Segment") <> "Not Mapped" Then
            partiallyMapped = partiallyMapped + 1
        Else
            notMapped = notMapped + 1
        End If
    Next matchKey

    row = 3
    outputWs.Cells(row, 1).Value = "MATCHING STATISTICS"
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 1).Font.Size = 12
    row = row + 2

    outputWs.Cells(row, 1).Value = "Total Stripe Packs:"
    outputWs.Cells(row, 2).Value = stripePacks.Count
    outputWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    outputWs.Cells(row, 1).Value = "Total Segment Packs:"
    outputWs.Cells(row, 2).Value = segmentPacks.Count
    outputWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    outputWs.Cells(row, 1).Value = "Exact Matches:"
    outputWs.Cells(row, 2).Value = exactMatches
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
    row = row + 1

    outputWs.Cells(row, 1).Value = "Fuzzy Matches:"
    outputWs.Cells(row, 2).Value = fuzzyMatches
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 2).Interior.Color = RGB(255, 235, 156)
    row = row + 1

    outputWs.Cells(row, 1).Value = "Not Found:"
    outputWs.Cells(row, 2).Value = notFound
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 2).Interior.Color = RGB(255, 199, 206)
    row = row + 2

    outputWs.Cells(row, 1).Value = "MAPPING STATISTICS"
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 1).Font.Size = 12
    row = row + 2

    outputWs.Cells(row, 1).Value = "Fully Mapped (Division AND Segment):"
    outputWs.Cells(row, 2).Value = fullyMapped
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
    row = row + 1

    outputWs.Cells(row, 1).Value = "Partially Mapped (Division OR Segment):"
    outputWs.Cells(row, 2).Value = partiallyMapped
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 2).Interior.Color = RGB(255, 235, 156)
    row = row + 1

    outputWs.Cells(row, 1).Value = "Not Mapped:"
    outputWs.Cells(row, 2).Value = notMapped
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 2).Interior.Color = RGB(255, 199, 206)

    outputWs.Columns.AutoFit
End Sub
