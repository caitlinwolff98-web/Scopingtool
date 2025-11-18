Attribute VB_Name = "Mod3_DataExtraction"
Option Explicit

' ============================================================================
' MODULE 3: DATA EXTRACTION & TABLE GENERATION
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.1 - Enhanced with FSLI Type Detection and Formula-Driven Tables
' ============================================================================
' PURPOSE: Extract financial data and generate structured Excel Tables
' DESCRIPTION: Processes Input Continuing, Discontinued, Journals, and Consol
'              tabs to create data tables with formula-driven percentages
'              and automatic FSLI type detection (Income Statement vs Balance Sheet)
' ============================================================================

' ==================== MODULE-LEVEL VARIABLES ====================
Private m_FSLITypes As Object ' Dictionary: FSLI -> Type (Income Statement / Balance Sheet)

' ==================== ROW CONSTANTS ====================
Private Const ROW_CURRENCY_TYPE As Long = 6    ' Row containing currency type identifiers
Private Const ROW_PACK_NAME As Long = 7        ' Row containing pack/entity names
Private Const ROW_PACK_CODE As Long = 8        ' Row containing pack/entity codes
Private Const ROW_FSLI_START As Long = 9       ' First row of FSLI data

' ==================== GET FSLI TYPES ====================
Public Function GetFSLITypes() As Object
    '------------------------------------------------------------------------
    ' Return the FSLI types dictionary (Income Statement vs Balance Sheet)
    ' Must call ExtractFSLITypesFromInput first
    '------------------------------------------------------------------------
    If m_FSLITypes Is Nothing Then
        Set m_FSLITypes = CreateObject("Scripting.Dictionary")
    End If
    Set GetFSLITypes = m_FSLITypes
End Function

' ==================== EXTRACT FSLI TYPES ====================
Public Sub ExtractFSLITypesFromInput(tabCategories As Object)
    '------------------------------------------------------------------------
    ' Scan Input Continuing Column B to identify FSLI types
    ' Detects "INCOME STATEMENT" and "BALANCE SHEET" headers
    ' Maps each FSLI to its statement type
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim inputTab As Worksheet
    Dim row As Long
    Dim lastRow As Long
    Dim cellValue As String
    Dim currentType As String
    Dim fsliValue As String

    Set m_FSLITypes = CreateObject("Scripting.Dictionary")
    currentType = "Unknown"

    ' Get Input Continuing tab
    Set inputTab = Mod2_TabProcessing.GetTabByCategory(tabCategories, "Input Continuing")
    If inputTab Is Nothing Then Exit Sub

    ' Find last row in Column B
    lastRow = inputTab.Cells(inputTab.Rows.Count, 2).End(xlUp).row

    ' Scan Column B to detect statement headers and categorize FSLIs
    For row = ROW_FSLI_START To lastRow
        cellValue = Trim(inputTab.Cells(row, 2).Value)

        ' Stop at Notes
        If UCase(cellValue) = "NOTES" Then Exit For

        ' Skip empty rows
        If cellValue = "" Then GoTo NextRow

        ' Check for statement headers
        If IsIncomeStatementHeader(cellValue) Then
            currentType = "Income Statement"
            GoTo NextRow
        End If

        If IsBalanceSheetHeader(cellValue) Then
            currentType = "Balance Sheet"
            GoTo NextRow
        End If

        ' If this is a valid FSLI (not a header), store its type
        If Not IsStatementHeader(cellValue) Then
            m_FSLITypes(cellValue) = currentType
        End If

NextRow:
    Next row

    Exit Sub

ErrorHandler:
    Debug.Print "Error extracting FSLI types: " & Err.Description
End Sub

' ==================== GET ALL ENTITIES ====================
Public Function GetAllEntitiesFromInputContinuing(tabCategories As Object, useConsolCurrency As Boolean) As Object
    '------------------------------------------------------------------------
    ' Extract all entities from Input Continuing tab
    ' Returns Dictionary: pack code -> pack name
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim entities As Object
    Dim inputTab As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim packCode As String
    Dim packName As String
    Dim currencyType As String

    Set entities = CreateObject("Scripting.Dictionary")

    ' Get Input Continuing tab
    Set inputTab = Mod2_TabProcessing.GetTabByCategory(tabCategories, "Input Continuing")
    If inputTab Is Nothing Then
        Set GetAllEntitiesFromInputContinuing = entities
        Exit Function
    End If

    ' Find last column
    lastCol = inputTab.Cells(ROW_PACK_NAME, inputTab.Columns.Count).End(xlToLeft).Column

    ' Extract entities from columns matching currency criteria
    For col = 3 To lastCol ' Start from column C (assuming A-B are labels)
        currencyType = Trim(UCase(inputTab.Cells(ROW_CURRENCY_TYPE, col).Value))

        ' Check currency match
        If IsConsolidationCurrency(currencyType) = useConsolCurrency Then
            packCode = Trim(inputTab.Cells(ROW_PACK_CODE, col).Value)
            packName = Trim(inputTab.Cells(ROW_PACK_NAME, col).Value)

            If packCode <> "" And packName <> "" Then
                If Not entities.exists(packCode) Then
                    entities(packCode) = packName
                End If
            End If
        End If
    Next col

    Set GetAllEntitiesFromInputContinuing = entities
    Exit Function

ErrorHandler:
    Set GetAllEntitiesFromInputContinuing = CreateObject("Scripting.Dictionary")
End Function

' ==================== GENERATE FULL INPUT TABLES ====================
Public Sub GenerateFullInputTables(tabCategories As Object, useConsolCurrency As Boolean, consolEntity As String)
    '------------------------------------------------------------------------
    ' Generate Full Input Table and Full Input Percentage Table
    ' Creates proper Excel Tables with formula-driven percentages
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.StatusBar = "Generating Full Input Tables..."
    Application.ScreenUpdating = False

    Dim inputTab As Worksheet
    Dim outputWs As Worksheet
    Dim percentWs As Worksheet
    Dim fslis As Collection
    Dim packs As Object
    Dim outRow As Long
    Dim outCol As Long

    ' Get Input Continuing tab
    Set inputTab = Mod2_TabProcessing.GetTabByCategory(tabCategories, "Input Continuing")
    If inputTab Is Nothing Then Exit Sub

    ' Extract FSLI types first
    ExtractFSLITypesFromInput tabCategories

    ' Create output worksheets
    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Full Input Table"

    Set percentWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    percentWs.Name = "Full Input Percentage"

    ' Extract FSLIs (Column B, starting from ROW_FSLI_START, stop at "Notes")
    Set fslis = ExtractFSLIs(inputTab)

    ' Extract packs matching currency type
    Set packs = ExtractPacks(inputTab, useConsolCurrency)

    ' Write headers - Row 1: FSLIs, Column A: Packs
    outputWs.Cells(1, 1).Value = "Pack Name (Code)"
    percentWs.Cells(1, 1).Value = "Pack Name (Code)"

    outCol = 2 ' Column B onwards for FSLIs
    Dim fsli As Variant
    For Each fsli In fslis
        outputWs.Cells(1, outCol).Value = fsli
        percentWs.Cells(1, outCol).Value = fsli
        outCol = outCol + 1
    Next fsli

    ' Write pack names in Column A
    outRow = 2
    Dim packCode As Variant
    For Each packCode In packs.Keys
        outputWs.Cells(outRow, 1).Value = packs(packCode) & " (" & packCode & ")"
        percentWs.Cells(outRow, 1).Value = packs(packCode) & " (" & packCode & ")"
        outRow = outRow + 1
    Next packCode

    ' Extract amounts and populate table
    PopulateAmountTable outputWs, inputTab, fslis, packs, useConsolCurrency

    ' Create formula-driven percentage table
    CreateFormulaDrivenPercentageTable percentWs, outputWs, consolEntity, fslis.Count, packs.Count

    ' Convert to proper Excel Tables
    ConvertToExcelTable outputWs, "FullInputTable"
    ConvertToExcelTable percentWs, "FullInputPercentageTable"

    ' Format tables
    FormatTable outputWs
    FormatTable percentWs

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error generating Full Input Tables: " & Err.Description, vbCritical
End Sub

' ==================== CREATE FORMULA-DRIVEN PERCENTAGE TABLE ====================
Private Sub CreateFormulaDrivenPercentageTable(percentWs As Worksheet, amountWs As Worksheet, _
                                               consolEntity As String, fsliCount As Long, packCount As Long)
    '------------------------------------------------------------------------
    ' Create formula-driven percentage table that references amount table
    ' Formulas automatically update when amounts change
    '------------------------------------------------------------------------
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim consolRow As Long
    Dim packName As String
    Dim amountTableName As String
    Dim formula As String

    ' Get table dimensions
    lastRow = packCount + 1 ' +1 for header row
    lastCol = fsliCount + 1 ' +1 for pack name column

    ' Find consolidation entity row
    For row = 2 To lastRow
        packName = percentWs.Cells(row, 1).Value
        If InStr(packName, consolEntity) > 0 Then
            consolRow = row
            Exit For
        End If
    Next row

    If consolRow = 0 Then
        MsgBox "Warning: Consolidation entity not found. Percentages may be incorrect.", vbExclamation
        Exit Sub
    End If

    ' Build formulas for each cell
    For row = 2 To lastRow
        For col = 2 To lastCol
            ' Formula: =IF('Full Input Table'!ConsolRow<>0, 'Full Input Table'!CurrentRow/'Full Input Table'!ConsolRow, 0)
            formula = "=IFERROR(" & _
                     "'" & amountWs.Name & "'!" & Cells(row, col).Address & "/" & _
                     "'" & amountWs.Name & "'!" & Cells(consolRow, col).Address & _
                     ",0)"

            percentWs.Cells(row, col).formula = formula
            percentWs.Cells(row, col).NumberFormat = "0.00%"
        Next col
    Next row
End Sub

' ==================== EXTRACT FSLIs ====================
Private Function ExtractFSLIs(ws As Worksheet) As Collection
    '------------------------------------------------------------------------
    ' Extract all FSLIs from Column B, starting from ROW_FSLI_START
    ' Stops at row containing "Notes" (case-insensitive)
    ' Excludes statement headers like "INCOME STATEMENT", "BALANCE SHEET"
    '------------------------------------------------------------------------
    Dim fslis As Collection
    Dim row As Long
    Dim fsliValue As String
    Dim lastRow As Long

    Set fslis = New Collection

    ' Find last row in Column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row

    ' Extract FSLIs
    For row = ROW_FSLI_START To lastRow
        fsliValue = Trim(ws.Cells(row, 2).Value)

        ' Stop at "Notes"
        If UCase(fsliValue) = "NOTES" Then Exit For

        ' Skip empty rows
        If fsliValue = "" Then GoTo NextRow

        ' Skip statement headers
        If IsStatementHeader(fsliValue) Then GoTo NextRow

        ' Add to collection
        fslis.Add fsliValue

NextRow:
    Next row

    Set ExtractFSLIs = fslis
End Function

' ==================== EXTRACT PACKS ====================
Private Function ExtractPacks(ws As Worksheet, useConsolCurrency As Boolean) As Object
    '------------------------------------------------------------------------
    ' Extract all packs from Row 7 & 8 matching currency criteria
    ' Returns Dictionary: pack code -> pack name
    '------------------------------------------------------------------------
    Dim packs As Object
    Dim lastCol As Long
    Dim col As Long
    Dim packCode As String
    Dim packName As String
    Dim currencyType As String

    Set packs = CreateObject("Scripting.Dictionary")

    lastCol = ws.Cells(ROW_PACK_NAME, ws.Columns.Count).End(xlToLeft).Column

    For col = 3 To lastCol
        currencyType = Trim(UCase(ws.Cells(ROW_CURRENCY_TYPE, col).Value))

        If IsConsolidationCurrency(currencyType) = useConsolCurrency Then
            packCode = Trim(ws.Cells(ROW_PACK_CODE, col).Value)
            packName = Trim(ws.Cells(ROW_PACK_NAME, col).Value)

            If packCode <> "" And packName <> "" Then
                If Not packs.exists(packCode) Then
                    packs(packCode) = packName
                End If
            End If
        End If
    Next col

    Set ExtractPacks = packs
End Function

' ==================== POPULATE AMOUNT TABLE ====================
Private Sub PopulateAmountTable(outputWs As Worksheet, sourceWs As Worksheet, _
                                fslis As Collection, packs As Object, useConsolCurrency As Boolean)
    '------------------------------------------------------------------------
    ' Populate the amount table with data from source worksheet
    '------------------------------------------------------------------------
    Dim fsliRow As Long
    Dim packCol As Long
    Dim outRow As Long
    Dim outCol As Long
    Dim packCode As Variant
    Dim fsli As Variant
    Dim sourceCol As Long
    Dim sourceFsliRow As Long
    Dim amount As Variant

    outRow = 2 ' Start from row 2 (row 1 is headers)

    For Each packCode In packs.Keys
        ' Find pack column in source
        packCol = FindPackColumn(sourceWs, CStr(packCode), useConsolCurrency)

        If packCol > 0 Then
            outCol = 2 ' Start from column B

            For Each fsli In fslis
                ' Find FSLI row in source
                sourceFsliRow = FindFSLIRow(sourceWs, CStr(fsli))

                If sourceFsliRow > 0 Then
                    amount = sourceWs.Cells(sourceFsliRow, packCol).Value

                    If IsNumeric(amount) Then
                        outputWs.Cells(outRow, outCol).Value = CDbl(amount)
                        outputWs.Cells(outRow, outCol).NumberFormat = "#,##0.00"
                    End If
                End If

                outCol = outCol + 1
            Next fsli
        End If

        outRow = outRow + 1
    Next packCode
End Sub

' ==================== GENERATE PACK COMPANY TABLE ====================
Public Sub GeneratePackCompanyTable(tabCategories As Object, divisionNames As Object, consolEntity As String)
    '------------------------------------------------------------------------
    ' Generate Pack Number Company Table with pack master data
    ' Includes: Pack Name, Pack Code, Division, Segment, Is Consolidated flag
    ' Creates proper Excel Table for Power BI
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.StatusBar = "Creating Pack Number Company Table..."

    Dim outputWs As Worksheet
    Dim inputTab As Worksheet
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim packName As String
    Dim division As String

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Pack Number Company Table"

    ' Write headers
    outputWs.Cells(1, 1).Value = "Pack Name"
    outputWs.Cells(1, 2).Value = "Pack Code"
    outputWs.Cells(1, 3).Value = "Division"
    outputWs.Cells(1, 4).Value = "Segment"
    outputWs.Cells(1, 5).Value = "Is Consolidated"

    ' Format headers
    With outputWs.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = 2

    ' Extract pack data from Input Continuing
    Set inputTab = Mod2_TabProcessing.GetTabByCategory(tabCategories, "Input Continuing")
    If Not inputTab Is Nothing Then
        lastCol = inputTab.Cells(ROW_PACK_NAME, inputTab.Columns.Count).End(xlToLeft).Column

        For col = 3 To lastCol
            packCode = Trim(inputTab.Cells(ROW_PACK_CODE, col).Value)
            packName = Trim(inputTab.Cells(ROW_PACK_NAME, col).Value)

            If packCode <> "" And packName <> "" Then
                outputWs.Cells(row, 1).Value = packName
                outputWs.Cells(row, 2).Value = packCode
                outputWs.Cells(row, 3).Value = "To Be Mapped" ' Will be updated by segmental matching
                outputWs.Cells(row, 4).Value = "To Be Mapped" ' Will be updated by segmental matching
                outputWs.Cells(row, 5).Value = IIf(packCode = consolEntity, "Yes", "No")
                row = row + 1
            End If
        Next col
    End If

    ' Convert to Excel Table
    If row > 2 Then ' Only if we have data
        ConvertToExcelTable outputWs, "PackNumberCompanyTable"
    End If

    outputWs.Columns.AutoFit
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error generating Pack Company Table: " & Err.Description, vbCritical
End Sub

' ==================== CONVERT TO EXCEL TABLE ====================
Private Sub ConvertToExcelTable(ws As Worksheet, tableName As String)
    '------------------------------------------------------------------------
    ' Convert range to proper Excel Table (ListObject)
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range
    Dim existingTable As ListObject

    ' Delete existing table with same name if it exists
    For Each existingTable In ws.ListObjects
        If existingTable.Name = tableName Then
            existingTable.Delete
        End If
    Next existingTable

    ' Find data range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 1 Then Exit Sub ' No data

    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Create ListObject
    ws.ListObjects.Add xlSrcRange, tableRange, , xlYes
    ws.ListObjects(ws.ListObjects.Count).Name = tableName
    ws.ListObjects(tableName).TableStyle = "TableStyleMedium2"

    On Error GoTo 0
End Sub

' ==================== HELPER FUNCTIONS ====================
Private Function IsConsolidationCurrency(currencyType As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if currency type is consolidation/consolidable currency
    '------------------------------------------------------------------------
    IsConsolidationCurrency = (InStr(currencyType, "CONSOL") > 0)
End Function

Private Function IsStatementHeader(fsliValue As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if FSLI value is a statement header that should be excluded
    '------------------------------------------------------------------------
    Dim upperValue As String
    upperValue = UCase(Trim(fsliValue))

    IsStatementHeader = IsIncomeStatementHeader(fsliValue) Or IsBalanceSheetHeader(fsliValue)
End Function

Private Function IsIncomeStatementHeader(fsliValue As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if value is an Income Statement header
    '------------------------------------------------------------------------
    Dim upperValue As String
    upperValue = UCase(Trim(fsliValue))

    IsIncomeStatementHeader = (upperValue = "INCOME STATEMENT" Or _
                               upperValue = "STATEMENT OF COMPREHENSIVE INCOME" Or _
                               upperValue = "PROFIT OR LOSS" Or _
                               InStr(upperValue, "INCOME STATEMENT") > 0)
End Function

Private Function IsBalanceSheetHeader(fsliValue As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if value is a Balance Sheet header
    '------------------------------------------------------------------------
    Dim upperValue As String
    upperValue = UCase(Trim(fsliValue))

    IsBalanceSheetHeader = (upperValue = "BALANCE SHEET" Or _
                            upperValue = "STATEMENT OF FINANCIAL POSITION" Or _
                            InStr(upperValue, "BALANCE SHEET") > 0)
End Function

Private Function FindPackColumn(ws As Worksheet, packCode As String, useConsolCurrency As Boolean) As Long
    '------------------------------------------------------------------------
    ' Find column number for a specific pack code
    '------------------------------------------------------------------------
    Dim col As Long
    Dim lastCol As Long
    Dim currencyType As String

    lastCol = ws.Cells(ROW_PACK_CODE, ws.Columns.Count).End(xlToLeft).Column

    For col = 3 To lastCol
        currencyType = Trim(UCase(ws.Cells(ROW_CURRENCY_TYPE, col).Value))

        If IsConsolidationCurrency(currencyType) = useConsolCurrency Then
            If Trim(ws.Cells(ROW_PACK_CODE, col).Value) = packCode Then
                FindPackColumn = col
                Exit Function
            End If
        End If
    Next col

    FindPackColumn = 0 ' Not found
End Function

Private Function FindFSLIRow(ws As Worksheet, fsli As String) As Long
    '------------------------------------------------------------------------
    ' Find row number for a specific FSLI
    '------------------------------------------------------------------------
    Dim row As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row

    For row = ROW_FSLI_START To lastRow
        If Trim(ws.Cells(row, 2).Value) = fsli Then
            FindFSLIRow = row
            Exit Function
        End If
    Next row

    FindFSLIRow = 0 ' Not found
End Function

Private Sub FormatTable(ws As Worksheet)
    '------------------------------------------------------------------------
    ' Apply professional formatting to table
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    ' Format headers
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Freeze panes
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True

    ' Auto-fit
    ws.Columns.AutoFit

    On Error GoTo 0
End Sub
