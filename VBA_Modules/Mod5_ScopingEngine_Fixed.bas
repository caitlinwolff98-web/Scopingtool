Attribute VB_Name = "Mod5_ScopingEngine"
Option Explicit

' =================================================================================
' MODULE 5: SCOPING ENGINE AND FACT TABLE GENERATION
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 7.0 - Complete Fix with Fact_Scoping Table
' =================================================================================
' PURPOSE:
'   Threshold configuration, automatic scoping, manual scoping interface
'   Generate Fact_Scoping table for dashboard formulas and Power BI
'
' CRITICAL FIXES:
'   1. Generate Fact_Scoping table with proper structure
'   2. Populate with threshold-based scoping results
'   3. Enable manual scoping updates to Fact_Scoping
'   4. Create proper Excel Table for dashboard references
'
' AUTHOR: ISA 600 Scoping Tool Team
' DATE: 2025-11-18
' =================================================================================

' ==================== CONFIGURE THRESHOLDS ====================
Public Function ConfigureThresholds() As Collection
    '------------------------------------------------------------------------
    ' Configure threshold-based scoping criteria
    ' Returns Collection of threshold configurations
    ' Each item: {FSLI, ThresholdAmount}
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim thresholds As Collection
    Dim fslis As Collection
    Dim fsli As Variant
    Dim selectedFSLIs As String
    Dim fsliList As String
    Dim counter As Long
    Dim thresholdAmount As String
    Dim fsliArray() As String
    Dim i As Long

    Set thresholds = New Collection

    ' Step 1: Get list of available FSLIs from Full Input Table
    Set fslis = GetAvailableFSLIs()

    If fslis.Count = 0 Then
        MsgBox "No FSLIs available for threshold configuration.", vbExclamation
        Set ConfigureThresholds = thresholds
        Exit Function
    End If

    ' Step 2: Display FSLI selection prompt (NO SYMBOLS)
    fsliList = "AVAILABLE FSLIs FOR THRESHOLD CRITERIA:" & vbCrLf & String(60, "-") & vbCrLf

    counter = 1
    For Each fsli In fslis
        fsliList = fsliList & counter & ". " & fsli & vbCrLf
        counter = counter + 1
    Next fsli

    fsliList = fsliList & vbCrLf & "Enter FSLI numbers separated by commas:" & vbCrLf & _
               "Example: 1,5,12 (to select FSLIs 1, 5, and 12)" & vbCrLf & vbCrLf & _
               "Recommended: Select Revenue, PBT, Total Assets"

    selectedFSLIs = InputBox(fsliList, "Select Threshold FSLIs", "")

    If Trim(selectedFSLIs) = "" Then
        Set ConfigureThresholds = thresholds
        Exit Function
    End If

    ' Step 3: Parse selected FSLIs
    fsliArray = Split(selectedFSLIs, ",")

    For i = LBound(fsliArray) To UBound(fsliArray)
        If IsNumeric(Trim(fsliArray(i))) Then
            Dim fsliIndex As Long
            fsliIndex = CLng(Trim(fsliArray(i)))

            If fsliIndex >= 1 And fsliIndex <= fslis.Count Then
                Dim selectedFSLI As String
                selectedFSLI = fslis(fsliIndex)

                ' Step 4: Prompt for threshold amount for this FSLI
                thresholdAmount = InputBox( _
                    "THRESHOLD FOR: " & selectedFSLI & vbCrLf & vbCrLf & _
                    "Enter threshold amount:" & vbCrLf & _
                    "(Packs exceeding this amount will be scoped in)" & vbCrLf & vbCrLf & _
                    "Example: 50000000 for R50 million" & vbCrLf & vbCrLf & _
                    "Enter amount:", _
                    "Threshold Amount", "")

                If Trim(thresholdAmount) <> "" And IsNumeric(thresholdAmount) Then
                    ' Add to thresholds collection
                    Dim thresholdConfig As Object
                    Set thresholdConfig = CreateObject("Scripting.Dictionary")
                    thresholdConfig("FSLI") = selectedFSLI
                    thresholdConfig("Amount") = CDbl(thresholdAmount)

                    thresholds.Add thresholdConfig
                End If
            End If
        End If
    Next i

    ' Step 5: Confirm threshold configuration
    If thresholds.Count > 0 Then
        Dim confirmMsg As String
        confirmMsg = "THRESHOLD CONFIGURATION SUMMARY" & vbCrLf & vbCrLf

        For i = 1 To thresholds.Count
            confirmMsg = confirmMsg & i & ". " & thresholds(i)("FSLI") & ": " & _
                        Format(thresholds(i)("Amount"), "#,##0.00") & vbCrLf
        Next i

        confirmMsg = confirmMsg & vbCrLf & "Rule: If ANY threshold is exceeded, ENTIRE PACK is scoped in." & vbCrLf & vbCrLf & _
                    "Proceed with this configuration?"

        If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Thresholds") <> vbYes Then
            Set thresholds = New Collection ' Clear and restart
        End If
    End If

    Set ConfigureThresholds = thresholds
    Exit Function

ErrorHandler:
    MsgBox "Error configuring thresholds: " & Err.Description, vbCritical
    Set ConfigureThresholds = New Collection
End Function

' ==================== APPLY THRESHOLDS ====================
Public Function ApplyThresholds(thresholds As Collection, consolEntity As String) As Object
    '------------------------------------------------------------------------
    ' Apply threshold criteria to identify scoped packs
    ' Returns Dictionary of scoped pack codes -> triggering FSLI
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim scopedPacks As Object
    Dim inputWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim packCode As String
    Dim packName As String
    Dim fsliName As String
    Dim amount As Double
    Dim thresholdItem As Object
    Dim i As Long

    Set scopedPacks = CreateObject("Scripting.Dictionary")

    ' Get Full Input Table
    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo ErrorHandler

    If inputWs Is Nothing Then
        Set ApplyThresholds = scopedPacks
        Exit Function
    End If

    lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).row
    lastCol = inputWs.Cells(1, inputWs.Columns.Count).End(xlToLeft).Column

    ' Process each pack
    For row = 2 To lastRow
        packName = inputWs.Cells(row, 1).Value

        ' Skip consolidation entity
        If InStr(UCase(packName), UCase(consolEntity)) > 0 Then GoTo NextRow

        ' Extract pack code from "Name (Code)" format
        packCode = ExtractPackCode(packName)

        ' Check each threshold FSLI for this pack
        For i = 1 To thresholds.Count
            Set thresholdItem = thresholds(i)
            fsliName = thresholdItem("FSLI")

            ' Find column for this FSLI
            col = FindFSLIColumn(inputWs, fsliName)

            If col > 0 Then
                amount = 0
                If IsNumeric(inputWs.Cells(row, col).Value) Then
                    amount = Abs(CDbl(inputWs.Cells(row, col).Value))
                End If

                ' Check if threshold exceeded
                If amount > thresholdItem("Amount") Then
                    If Not scopedPacks.exists(packCode) Then
                        scopedPacks(packCode) = fsliName ' Store triggering FSLI
                    End If
                End If
            End If
        Next i

NextRow:
    Next row

    Set ApplyThresholds = scopedPacks
    Exit Function

ErrorHandler:
    MsgBox "Error applying thresholds: " & Err.Description, vbCritical
    Set ApplyThresholds = CreateObject("Scripting.Dictionary")
End Function

' ==================== GENERATE FACT_SCOPING TABLE ====================
Public Sub GenerateFactScopingTable(scopedPacks As Object, thresholds As Collection, consolEntity As String)
    '------------------------------------------------------------------------
    ' CRITICAL NEW FUNCTION: Generate Fact_Scoping table
    ' This table is used by dashboards for formula-driven coverage calculations
    '
    ' TABLE STRUCTURE:
    '   - PackCode
    '   - PackName
    '   - FSLI
    '   - ScopingStatus (Scoped In / Not Scoped)
    '   - ScopingMethod (Automatic (Threshold) / Manual / Not Scoped)
    '   - ThresholdFSLI (which FSLI triggered threshold)
    '   - ScopedDate
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.StatusBar = "Generating Fact Scoping table..."

    Dim outputWs As Worksheet
    Dim packTable As Worksheet
    Dim fsliTable As Worksheet
    Dim row As Long
    Dim packRow As Long
    Dim fsliRow As Long
    Dim lastPackRow As Long
    Dim lastFsliRow As Long
    Dim packCode As String
    Dim packName As String
    Dim fsliName As String
    Dim scopingStatus As String
    Dim scopingMethod As String
    Dim thresholdFSLI As String

    ' Create output worksheet
    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Fact Scoping"

    ' Write headers
    outputWs.Cells(1, 1).Value = "PackCode"
    outputWs.Cells(1, 2).Value = "PackName"
    outputWs.Cells(1, 3).Value = "FSLI"
    outputWs.Cells(1, 4).Value = "ScopingStatus"
    outputWs.Cells(1, 5).Value = "ScopingMethod"
    outputWs.Cells(1, 6).Value = "ThresholdFSLI"
    outputWs.Cells(1, 7).Value = "ScopedDate"

    ' Format headers
    With outputWs.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = 2

    ' Get Pack Number Company Table
    On Error Resume Next
    Set packTable = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    Set fsliTable = Mod1_MainController.g_OutputWorkbook.Worksheets("Dim FSLIs")
    On Error GoTo ErrorHandler

    If packTable Is Nothing Or fsliTable Is Nothing Then
        MsgBox "Error: Required tables not found. Cannot generate Fact Scoping table.", vbExclamation
        Exit Sub
    End If

    ' Find last rows
    lastPackRow = packTable.Cells(packTable.Rows.Count, 2).End(xlUp).row
    lastFsliRow = fsliTable.Cells(fsliTable.Rows.Count, 1).End(xlUp).row

    ' Loop through each pack
    For packRow = 2 To lastPackRow
        packCode = packTable.Cells(packRow, 2).Value
        packName = packTable.Cells(packRow, 1).Value

        ' Skip consolidation entity
        If packCode = consolEntity Then GoTo NextPack

        ' Check if this pack is scoped in by threshold
        If scopedPacks.exists(packCode) Then
            thresholdFSLI = scopedPacks(packCode)

            ' Loop through each FSLI and mark entire pack as scoped in
            For fsliRow = 2 To lastFsliRow
                fsliName = fsliTable.Cells(fsliRow, 1).Value

                outputWs.Cells(row, 1).Value = packCode
                outputWs.Cells(row, 2).Value = packName
                outputWs.Cells(row, 3).Value = fsliName
                outputWs.Cells(row, 4).Value = "Scoped In"
                outputWs.Cells(row, 5).Value = "Automatic (Threshold)"
                outputWs.Cells(row, 6).Value = thresholdFSLI
                outputWs.Cells(row, 7).Value = Now
                outputWs.Cells(row, 7).NumberFormat = "yyyy-mm-dd hh:mm:ss"

                row = row + 1
            Next fsliRow
        Else
            ' Pack not scoped in - create placeholder rows for manual scoping
            For fsliRow = 2 To lastFsliRow
                fsliName = fsliTable.Cells(fsliRow, 1).Value

                outputWs.Cells(row, 1).Value = packCode
                outputWs.Cells(row, 2).Value = packName
                outputWs.Cells(row, 3).Value = fsliName
                outputWs.Cells(row, 4).Value = "Not Scoped"
                outputWs.Cells(row, 5).Value = "Not Scoped"
                outputWs.Cells(row, 6).Value = ""
                outputWs.Cells(row, 7).Value = ""

                row = row + 1
            Next fsliRow
        End If

NextPack:
    Next packRow

    ' Convert to Excel Table
    If row > 2 Then
        Dim lastRow As Long
        Dim lastCol As Long
        Dim tableRange As Range

        lastRow = row - 1
        lastCol = 7

        Set tableRange = outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol))

        outputWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        outputWs.ListObjects(outputWs.ListObjects.Count).Name = "FactScoping"
        outputWs.ListObjects("FactScoping").TableStyle = "TableStyleMedium2"
    End If

    outputWs.Columns.AutoFit
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error generating Fact Scoping table: " & Err.Description, vbCritical
End Sub

' ==================== GENERATE DIM_THRESHOLDS TABLE ====================
Public Sub GenerateDimThresholdsTable(thresholds As Collection)
    '------------------------------------------------------------------------
    ' Generate Dim_Thresholds dimension table for Power BI
    '
    ' TABLE STRUCTURE:
    '   - FSLI
    '   - ThresholdAmount
    '   - ConfiguredDate
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    If thresholds.Count = 0 Then Exit Sub

    Application.StatusBar = "Creating Dim Thresholds table..."

    Dim outputWs As Worksheet
    Dim row As Long
    Dim i As Long

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Dim Thresholds"

    ' Write headers
    outputWs.Cells(1, 1).Value = "FSLI"
    outputWs.Cells(1, 2).Value = "ThresholdAmount"
    outputWs.Cells(1, 3).Value = "ConfiguredDate"

    ' Format headers
    With outputWs.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = 2

    For i = 1 To thresholds.Count
        outputWs.Cells(row, 1).Value = thresholds(i)("FSLI")
        outputWs.Cells(row, 2).Value = thresholds(i)("Amount")
        outputWs.Cells(row, 2).NumberFormat = "#,##0.00"
        outputWs.Cells(row, 3).Value = Now
        outputWs.Cells(row, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        row = row + 1
    Next i

    ' Convert to Excel Table
    If row > 2 Then
        Dim tableRange As Range
        Set tableRange = outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(row - 1, 3))
        outputWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        outputWs.ListObjects(outputWs.ListObjects.Count).Name = "DimThresholds"
        outputWs.ListObjects("DimThresholds").TableStyle = "TableStyleMedium2"
    End If

    outputWs.Columns.AutoFit
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error generating Dim Thresholds table: " & Err.Description, vbCritical
End Sub

' ==================== MANUAL SCOPING FUNCTIONS ====================
Public Function ScopeInPack(packCode As String) As Boolean
    '------------------------------------------------------------------------
    ' Manually scope in entire pack (all FSLIs)
    ' Updates Fact_Scoping table
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim factWs As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim currentPackCode As String

    Set factWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    If factWs Is Nothing Then
        MsgBox "Fact Scoping table not found.", vbExclamation
        ScopeInPack = False
        Exit Function
    End If

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row

    ' Update all rows for this pack
    For row = 2 To lastRow
        currentPackCode = factWs.Cells(row, 1).Value

        If currentPackCode = packCode Then
            factWs.Cells(row, 4).Value = "Scoped In"
            factWs.Cells(row, 5).Value = "Manual"
            factWs.Cells(row, 7).Value = Now
        End If
    Next row

    ScopeInPack = True
    Exit Function

ErrorHandler:
    ScopeInPack = False
End Function

Public Function ScopeInPackFSLI(packCode As String, fsli As String) As Boolean
    '------------------------------------------------------------------------
    ' Manually scope in specific FSLI for specific pack
    ' Updates Fact_Scoping table
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim factWs As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim currentPackCode As String
    Dim currentFSLI As String

    Set factWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    If factWs Is Nothing Then
        MsgBox "Fact Scoping table not found.", vbExclamation
        ScopeInPackFSLI = False
        Exit Function
    End If

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row

    ' Find and update the specific pack-FSLI combination
    For row = 2 To lastRow
        currentPackCode = factWs.Cells(row, 1).Value
        currentFSLI = factWs.Cells(row, 3).Value

        If currentPackCode = packCode And currentFSLI = fsli Then
            factWs.Cells(row, 4).Value = "Scoped In"
            factWs.Cells(row, 5).Value = "Manual"
            factWs.Cells(row, 7).Value = Now
            Exit For
        End If
    Next row

    ScopeInPackFSLI = True
    Exit Function

ErrorHandler:
    ScopeInPackFSLI = False
End Function

Public Function ScopeOutPackFSLI(packCode As String, fsli As String) As Boolean
    '------------------------------------------------------------------------
    ' Remove pack-FSLI from scoping
    ' Updates Fact_Scoping table
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim factWs As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim currentPackCode As String
    Dim currentFSLI As String

    Set factWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    If factWs Is Nothing Then
        MsgBox "Fact Scoping table not found.", vbExclamation
        ScopeOutPackFSLI = False
        Exit Function
    End If

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row

    ' Find and update the specific pack-FSLI combination
    For row = 2 To lastRow
        currentPackCode = factWs.Cells(row, 1).Value
        currentFSLI = factWs.Cells(row, 3).Value

        If currentPackCode = packCode And currentFSLI = fsli Then
            factWs.Cells(row, 4).Value = "Not Scoped"
            factWs.Cells(row, 5).Value = "Not Scoped"
            factWs.Cells(row, 6).Value = ""
            factWs.Cells(row, 7).Value = ""
            Exit For
        End If
    Next row

    ScopeOutPackFSLI = True
    Exit Function

ErrorHandler:
    ScopeOutPackFSLI = False
End Function

' ==================== HELPER FUNCTIONS ====================
Private Function GetAvailableFSLIs() As Collection
    '------------------------------------------------------------------------
    ' Get list of all FSLIs from Full Input Table headers
    '------------------------------------------------------------------------
    Dim fslis As Collection
    Dim inputWs As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim fsliName As String

    Set fslis = New Collection

    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If inputWs Is Nothing Then
        Set GetAvailableFSLIs = fslis
        Exit Function
    End If

    lastCol = inputWs.Cells(1, inputWs.Columns.Count).End(xlToLeft).Column

    For col = 2 To lastCol ' Start from column B (column A is pack names)
        fsliName = Trim(inputWs.Cells(1, col).Value)
        If fsliName <> "" Then
            fslis.Add fsliName
        End If
    Next col

    Set GetAvailableFSLIs = fslis
End Function

Private Function FindFSLIColumn(ws As Worksheet, fsliName As String) As Long
    '------------------------------------------------------------------------
    ' Find column number for a specific FSLI in the worksheet
    '------------------------------------------------------------------------
    Dim col As Long
    Dim lastCol As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 2 To lastCol
        If Trim(ws.Cells(1, col).Value) = fsliName Then
            FindFSLIColumn = col
            Exit Function
        End If
    Next col

    FindFSLIColumn = 0 ' Not found
End Function

Private Function ExtractPackCode(packNameWithCode As String) As String
    '------------------------------------------------------------------------
    ' Extract pack code from "Name (Code)" format
    '------------------------------------------------------------------------
    Dim openParen As Long
    Dim closeParen As Long

    openParen = InStrRev(packNameWithCode, "(")
    closeParen = InStrRev(packNameWithCode, ")")

    If openParen > 0 And closeParen > openParen Then
        ExtractPackCode = Trim(Mid(packNameWithCode, openParen + 1, closeParen - openParen - 1))
    Else
        ExtractPackCode = packNameWithCode
    End If
End Function

' ==================== GENERATE SCOPING SUMMARY ====================
Public Sub GenerateScopingSummary(scopedPacks As Object, thresholds As Collection)
    '------------------------------------------------------------------------
    ' Generate Scoping Summary sheet with comprehensive details
    '------------------------------------------------------------------------
    Dim outputWs As Worksheet
    Dim row As Long
    Dim packCode As Variant
    Dim packTable As Worksheet
    Dim packRow As Long
    Dim lastPackRow As Long
    Dim packName As String

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Scoping Summary"

    ' Write title
    outputWs.Cells(1, 1).Value = "ISA 600 SCOPING SUMMARY"
    outputWs.Cells(1, 1).Font.Size = 16
    outputWs.Cells(1, 1).Font.Bold = True
    outputWs.Cells(1, 1).Font.Color = RGB(0, 112, 192)

    ' Subtitle
    outputWs.Cells(2, 1).Value = "Automatically Scoped Packs (Threshold-Based)"
    outputWs.Cells(2, 1).Font.Size = 12
    outputWs.Cells(2, 1).Font.Bold = True

    ' Write headers
    row = 4
    outputWs.Cells(row, 1).Value = "Pack Code"
    outputWs.Cells(row, 2).Value = "Pack Name"
    outputWs.Cells(row, 3).Value = "Scoping Status"
    outputWs.Cells(row, 4).Value = "Triggering FSLI"
    outputWs.Cells(row, 5).Value = "Scoping Rationale"

    With outputWs.Range("A4:E4")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = 5

    ' Get Pack Number Company Table for pack names
    On Error Resume Next
    Set packTable = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    On Error GoTo 0

    For Each packCode In scopedPacks.Keys
        ' Find pack name
        packName = ""
        If Not packTable Is Nothing Then
            lastPackRow = packTable.Cells(packTable.Rows.Count, 2).End(xlUp).row
            For packRow = 2 To lastPackRow
                If packTable.Cells(packRow, 2).Value = packCode Then
                    packName = packTable.Cells(packRow, 1).Value
                    Exit For
                End If
            Next packRow
        End If

        outputWs.Cells(row, 1).Value = packCode
        outputWs.Cells(row, 2).Value = packName
        outputWs.Cells(row, 3).Value = "Scoped In"
        outputWs.Cells(row, 4).Value = scopedPacks(packCode)
        outputWs.Cells(row, 5).Value = "Exceeded threshold for " & scopedPacks(packCode)

        outputWs.Cells(row, 3).Interior.Color = RGB(198, 239, 206) ' Green

        row = row + 1
    Next packCode

    ' Summary statistics
    row = row + 2
    outputWs.Cells(row, 1).Value = "SUMMARY STATISTICS"
    outputWs.Cells(row, 1).Font.Bold = True
    outputWs.Cells(row, 1).Font.Size = 12
    row = row + 2

    outputWs.Cells(row, 1).Value = "Total Packs Automatically Scoped In:"
    outputWs.Cells(row, 2).Value = scopedPacks.Count
    outputWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    outputWs.Cells(row, 1).Value = "Threshold FSLIs Used:"
    outputWs.Cells(row, 2).Value = thresholds.Count
    outputWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    outputWs.Cells(row, 1).Value = "Scoping Date:"
    outputWs.Cells(row, 2).Value = Now
    outputWs.Cells(row, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    outputWs.Cells(row, 1).Font.Bold = True

    outputWs.Columns.AutoFit
End Sub
