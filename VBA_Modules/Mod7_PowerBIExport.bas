Attribute VB_Name = "Mod7_PowerBIExport"
Option Explicit

' ============================================================================
' MODULE 7: POWER BI EXPORT
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.0 - Complete Overhaul
' ============================================================================
' PURPOSE: Create Power BI-ready data tables
' DESCRIPTION: Generates properly structured dimension and fact tables
'              for seamless Power BI integration
' ============================================================================

' ==================== CREATE POWER BI ASSETS ====================
Public Sub CreatePowerBIAssets()
    '------------------------------------------------------------------------
    ' Main function to create all Power BI-ready tables
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.StatusBar = "Creating Power BI assets..."

    ' Create dimension tables
    CreateDimPacks
    CreateDimFSLIs

    ' Create fact tables
    CreateFactAmounts
    CreateFactPercentages
    CreateFactScoping
    CreateDimThresholds

    ' Create metadata sheet
    CreatePowerBIMetadata

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "Error creating Power BI assets: " & Err.Description
End Sub

' ==================== DIM_PACKS ====================
Private Sub CreateDimPacks()
    '------------------------------------------------------------------------
    ' Create Dim_Packs table - Pack master data
    ' Columns: PackCode, PackName, Division, Segment, IsConsolidated
    '------------------------------------------------------------------------
    Dim dimWs As Worksheet
    Dim sourceWs As Worksheet
    Dim row As Long
    Dim srcRow As Long
    Dim lastRow As Long

    Set dimWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dimWs.Name = "Dim_Packs"

    ' Headers
    dimWs.Cells(1, 1).Value = "PackCode"
    dimWs.Cells(1, 2).Value = "PackName"
    dimWs.Cells(1, 3).Value = "Division"
    dimWs.Cells(1, 4).Value = "Segment"
    dimWs.Cells(1, 5).Value = "IsConsolidated"

    dimWs.Range("A1:E1").Font.Bold = True
    dimWs.Range("A1:E1").Interior.Color = RGB(68, 114, 196)
    dimWs.Range("A1:E1").Font.Color = RGB(255, 255, 255)

    ' Get data from Pack Number Company Table
    On Error Resume Next
    Set sourceWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    On Error GoTo 0

    If Not sourceWs Is Nothing Then
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).row
        row = 2

        For srcRow = 2 To lastRow
            dimWs.Cells(row, 1).Value = sourceWs.Cells(srcRow, 2).Value ' Pack Code
            dimWs.Cells(row, 2).Value = sourceWs.Cells(srcRow, 1).Value ' Pack Name
            dimWs.Cells(row, 3).Value = sourceWs.Cells(srcRow, 3).Value ' Division
            dimWs.Cells(row, 4).Value = "" ' Segment (from mapping table)
            dimWs.Cells(row, 5).Value = sourceWs.Cells(srcRow, 4).Value ' Is Consolidated

            row = row + 1
        Next srcRow
    End If

    ' Convert to table
    ConvertToTable dimWs, "Dim_Packs"

    dimWs.Columns.AutoFit
End Sub

' ==================== DIM_FSLIS ====================
Private Sub CreateDimFSLIs()
    '------------------------------------------------------------------------
    ' Create Dim_FSLIs table - FSLI master data
    ' Columns: FSLI, Category (Income Statement/Balance Sheet), AccountNature
    '------------------------------------------------------------------------
    Dim dimWs As Worksheet
    Dim inputWs As Worksheet
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim fsli As String

    Set dimWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dimWs.Name = "Dim_FSLIs"

    ' Headers
    dimWs.Cells(1, 1).Value = "FSLI"
    dimWs.Cells(1, 2).Value = "Category"
    dimWs.Cells(1, 3).Value = "AccountNature"

    dimWs.Range("A1:C1").Font.Bold = True
    dimWs.Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    dimWs.Range("A1:C1").Font.Color = RGB(255, 255, 255)

    ' Get FSLIs from Full Input Table
    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If Not inputWs Is Nothing Then
        lastCol = inputWs.Cells(1, inputWs.Columns.Count).End(xlToLeft).Column
        row = 2

        For col = 2 To lastCol ' Start from column B
            fsli = inputWs.Cells(1, col).Value

            If fsli <> "" Then
                dimWs.Cells(row, 1).Value = fsli
                dimWs.Cells(row, 2).Value = DetermineFSLICategory(fsli)
                dimWs.Cells(row, 3).Value = DetermineFSLIAccountNature(fsli)

                row = row + 1
            End If
        Next col
    End If

    ' Convert to table
    ConvertToTable dimWs, "Dim_FSLIs"

    dimWs.Columns.AutoFit
End Sub

' ==================== FACT_AMOUNTS ====================
Private Sub CreateFactAmounts()
    '------------------------------------------------------------------------
    ' Create Fact_Amounts table - Unpivoted amounts
    ' Columns: PackCode, FSLI, Amount
    '------------------------------------------------------------------------
    Dim factWs As Worksheet
    Dim inputWs As Worksheet
    Dim outRow As Long
    Dim dataRow As Long
    Dim col As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim fsli As String
    Dim amount As Variant

    Set factWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    factWs.Name = "Fact_Amounts"

    ' Headers
    factWs.Cells(1, 1).Value = "PackCode"
    factWs.Cells(1, 2).Value = "FSLI"
    factWs.Cells(1, 3).Value = "Amount"

    factWs.Range("A1:C1").Font.Bold = True
    factWs.Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    factWs.Range("A1:C1").Font.Color = RGB(255, 255, 255)

    ' Get data from Full Input Table
    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If inputWs Is Nothing Then Exit Sub

    lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).row
    lastCol = inputWs.Cells(1, inputWs.Columns.Count).End(xlToLeft).Column

    outRow = 2

    ' Unpivot data
    For dataRow = 2 To lastRow
        packCode = ExtractPackCode(inputWs.Cells(dataRow, 1).Value)

        For col = 2 To lastCol
            fsli = inputWs.Cells(1, col).Value
            amount = inputWs.Cells(dataRow, col).Value

            If IsNumeric(amount) Then
                factWs.Cells(outRow, 1).Value = packCode
                factWs.Cells(outRow, 2).Value = fsli
                factWs.Cells(outRow, 3).Value = CDbl(amount)

                outRow = outRow + 1
            End If
        Next col
    Next dataRow

    ' Convert to table
    If outRow > 2 Then
        ConvertToTable factWs, "Fact_Amounts"
    End If

    factWs.Columns.AutoFit
End Sub

' ==================== FACT_PERCENTAGES ====================
Private Sub CreateFactPercentages()
    '------------------------------------------------------------------------
    ' Create Fact_Percentages table - Unpivoted percentages
    ' Columns: PackCode, FSLI, Percentage
    '------------------------------------------------------------------------
    Dim factWs As Worksheet
    Dim percentWs As Worksheet
    Dim outRow As Long
    Dim dataRow As Long
    Dim col As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim packCode As String
    Dim fsli As String
    Dim percentage As Variant

    Set factWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    factWs.Name = "Fact_Percentages"

    ' Headers
    factWs.Cells(1, 1).Value = "PackCode"
    factWs.Cells(1, 2).Value = "FSLI"
    factWs.Cells(1, 3).Value = "Percentage"

    factWs.Range("A1:C1").Font.Bold = True
    factWs.Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    factWs.Range("A1:C1").Font.Color = RGB(255, 255, 255)

    ' Get data from Full Input Percentage table
    On Error Resume Next
    Set percentWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Percentage")
    On Error GoTo 0

    If percentWs Is Nothing Then Exit Sub

    lastRow = percentWs.Cells(percentWs.Rows.Count, 1).End(xlUp).row
    lastCol = percentWs.Cells(1, percentWs.Columns.Count).End(xlToLeft).Column

    outRow = 2

    ' Unpivot data
    For dataRow = 2 To lastRow
        packCode = ExtractPackCode(percentWs.Cells(dataRow, 1).Value)

        For col = 2 To lastCol
            fsli = percentWs.Cells(1, col).Value
            percentage = percentWs.Cells(dataRow, col).Value

            If IsNumeric(percentage) Or percentage = "N/A" Then
                factWs.Cells(outRow, 1).Value = packCode
                factWs.Cells(outRow, 2).Value = fsli

                If IsNumeric(percentage) Then
                    factWs.Cells(outRow, 3).Value = CDbl(percentage)
                Else
                    factWs.Cells(outRow, 3).Value = 0
                End If

                outRow = outRow + 1
            End If
        Next col
    Next dataRow

    ' Convert to table
    If outRow > 2 Then
        ConvertToTable factWs, "Fact_Percentages"
    End If

    factWs.Columns.AutoFit
End Sub

' ==================== FACT_SCOPING ====================
Private Sub CreateFactScoping()
    '------------------------------------------------------------------------
    ' Create Fact_Scoping table - Scoping decisions
    ' Columns: PackCode, FSLI, ScopingStatus, ScopingMethod
    '------------------------------------------------------------------------
    Dim factWs As Worksheet
    Dim row As Long
    Dim packCode As Variant
    Dim fsliKey As Variant

    Set factWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    factWs.Name = "Fact_Scoping"

    ' Headers
    factWs.Cells(1, 1).Value = "PackCode"
    factWs.Cells(1, 2).Value = "PackName"
    factWs.Cells(1, 3).Value = "FSLI"
    factWs.Cells(1, 4).Value = "FSLIName"
    factWs.Cells(1, 5).Value = "ScopingStatus"
    factWs.Cells(1, 6).Value = "ScopingMethod"
    factWs.Cells(1, 7).Value = "ScopingReason"

    factWs.Range("A1:G1").Font.Bold = True
    factWs.Range("A1:G1").Interior.Color = RGB(68, 114, 196)
    factWs.Range("A1:G1").Font.Color = RGB(255, 255, 255)

    row = 2

    ' Add scoped packs (automatic/threshold scoping)
    For Each packCode In Mod1_MainController.g_ScopedPacks.Keys
        Dim packName As String
        Dim scopingReason As String

        packName = GetPackName(CStr(packCode))
        scopingReason = Mod1_MainController.g_ScopedPacks(packCode)

        factWs.Cells(row, 1).Value = packCode
        factWs.Cells(row, 2).Value = packName
        factWs.Cells(row, 3).Value = "ALL"
        factWs.Cells(row, 4).Value = "All FSLIs"
        factWs.Cells(row, 5).Value = "Scoped In"
        factWs.Cells(row, 6).Value = "Automatic (Threshold)"
        factWs.Cells(row, 7).Value = scopingReason

        row = row + 1
    Next packCode

    ' Add manual scoping decisions
    For Each fsliKey In Mod1_MainController.g_ManualScoping.Keys
        Dim parts() As String
        parts = Split(CStr(fsliKey), "|")

        Dim manualPackCode As String
        Dim manualPackName As String
        Dim manualFSLI As String

        manualPackCode = parts(0)
        manualFSLI = parts(1)
        manualPackName = GetPackName(manualPackCode)

        factWs.Cells(row, 1).Value = manualPackCode
        factWs.Cells(row, 2).Value = manualPackName
        factWs.Cells(row, 3).Value = manualPackCode
        factWs.Cells(row, 4).Value = manualFSLI
        factWs.Cells(row, 5).Value = Mod1_MainController.g_ManualScoping(fsliKey)
        factWs.Cells(row, 6).Value = "Manual"
        factWs.Cells(row, 7).Value = "Manually scoped by user"

        row = row + 1
    Next fsliKey

    ' Convert to table
    If row > 2 Then
        ConvertToTable factWs, "Fact_Scoping"
    End If

    factWs.Columns.AutoFit
End Sub

' ==================== DIM_THRESHOLDS ====================
Private Sub CreateDimThresholds()
    '------------------------------------------------------------------------
    ' Create Dim_Thresholds table - Threshold configuration
    ' Columns: FSLI, ThresholdAmount
    '------------------------------------------------------------------------
    Dim dimWs As Worksheet
    Dim row As Long
    Dim i As Long

    Set dimWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dimWs.Name = "Dim_Thresholds"

    ' Headers
    dimWs.Cells(1, 1).Value = "FSLI"
    dimWs.Cells(1, 2).Value = "ThresholdAmount"

    dimWs.Range("A1:B1").Font.Bold = True
    dimWs.Range("A1:B1").Interior.Color = RGB(68, 114, 196)
    dimWs.Range("A1:B1").Font.Color = RGB(255, 255, 255)

    row = 2

    ' Add threshold data
    If Not Mod1_MainController.g_ThresholdFSLIs Is Nothing Then
        For i = 1 To Mod1_MainController.g_ThresholdFSLIs.Count
            dimWs.Cells(row, 1).Value = Mod1_MainController.g_ThresholdFSLIs(i)("FSLI")
            dimWs.Cells(row, 2).Value = Mod1_MainController.g_ThresholdFSLIs(i)("Amount")
            dimWs.Cells(row, 2).NumberFormat = "#,##0.00"

            row = row + 1
        Next i
    End If

    ' Convert to table
    If row > 2 Then
        ConvertToTable dimWs, "Dim_Thresholds"
    End If

    dimWs.Columns.AutoFit
End Sub

' ==================== POWER BI METADATA ====================
Private Sub CreatePowerBIMetadata()
    '------------------------------------------------------------------------
    ' Create metadata sheet for Power BI integration
    ' Includes relationship definitions and DAX measures
    '------------------------------------------------------------------------
    Dim metaWs As Worksheet

    Set metaWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    metaWs.Name = "PowerBI_Integration_Guide"

    ' Title
    metaWs.Cells(1, 1).Value = "POWER BI INTEGRATION GUIDE"
    metaWs.Cells(1, 1).Font.Size = 14
    metaWs.Cells(1, 1).Font.Bold = True

    Dim row As Long
    row = 3

    ' Relationships section
    metaWs.Cells(row, 1).Value = "RELATIONSHIPS TO CREATE:"
    metaWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    metaWs.Cells(row, 1).Value = "1. Fact_Amounts[PackCode] → Dim_Packs[PackCode] (Many-to-One)"
    row = row + 1
    metaWs.Cells(row, 1).Value = "2. Fact_Amounts[FSLI] → Dim_FSLIs[FSLI] (Many-to-One)"
    row = row + 1
    metaWs.Cells(row, 1).Value = "3. Fact_Percentages[PackCode] → Dim_Packs[PackCode] (Many-to-One)"
    row = row + 1
    metaWs.Cells(row, 1).Value = "4. Fact_Percentages[FSLI] → Dim_FSLIs[FSLI] (Many-to-One)"
    row = row + 1
    metaWs.Cells(row, 1).Value = "5. Fact_Scoping[PackCode] → Dim_Packs[PackCode] (Many-to-One)"
    row = row + 1
    metaWs.Cells(row, 1).Value = "6. Fact_Scoping[FSLI] → Dim_FSLIs[FSLI] (Many-to-One)"
    row = row + 2

    ' DAX Measures section
    metaWs.Cells(row, 1).Value = "RECOMMENDED DAX MEASURES:"
    metaWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    metaWs.Cells(row, 1).Value = "Total Amount = SUM(Fact_Amounts[Amount])"
    row = row + 1
    metaWs.Cells(row, 1).Value = "Scoped Amount = CALCULATE(SUM(Fact_Amounts[Amount]), Fact_Scoping[ScopingStatus] = ""Scoped In"")"
    row = row + 1
    metaWs.Cells(row, 1).Value = "Coverage % = DIVIDE([Scoped Amount], [Total Amount], 0)"
    row = row + 1
    metaWs.Cells(row, 1).Value = "Packs Scoped In = DISTINCTCOUNT(Fact_Scoping[PackCode])"
    row = row + 2

    ' Import instructions
    metaWs.Cells(row, 1).Value = "IMPORT INSTRUCTIONS:"
    metaWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    metaWs.Cells(row, 1).Value = "1. Open Power BI Desktop"
    row = row + 1
    metaWs.Cells(row, 1).Value = "2. Get Data → Excel → Select this workbook"
    row = row + 1
    metaWs.Cells(row, 1).Value = "3. Select all tables starting with 'Dim_' and 'Fact_'"
    row = row + 1
    metaWs.Cells(row, 1).Value = "4. Load data"
    row = row + 1
    metaWs.Cells(row, 1).Value = "5. Go to Model view and create relationships as listed above"
    row = row + 1
    metaWs.Cells(row, 1).Value = "6. Create DAX measures as recommended"
    row = row + 1
    metaWs.Cells(row, 1).Value = "7. Build visualizations"
    row = row + 1

    metaWs.Columns("A:A").AutoFit
End Sub

' ==================== HELPER FUNCTIONS ====================
Private Sub ConvertToTable(ws As Worksheet, tableName As String)
    '------------------------------------------------------------------------
    ' Convert range to Excel Table (ListObject)
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow > 1 And lastCol > 0 Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        tbl.Name = tableName
        tbl.TableStyle = "TableStyleMedium2"
    End If

    On Error GoTo 0
End Sub

Private Function ExtractPackCode(packNameWithCode As String) As String
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

Private Function DetermineFSLICategory(fsli As String) As String
    '------------------------------------------------------------------------
    ' Determine if FSLI belongs to Income Statement or Balance Sheet
    ' Uses FSLI types dictionary from Mod3
    '------------------------------------------------------------------------
    Dim fsliTypes As Object
    Set fsliTypes = Mod3_DataExtraction.GetFSLITypes()

    If Not fsliTypes Is Nothing Then
        If fsliTypes.exists(fsli) Then
            DetermineFSLICategory = fsliTypes(fsli)
            Exit Function
        End If
    End If

    ' Fallback: Unknown
    DetermineFSLICategory = "Unknown"
End Function

Private Function DetermineFSLIAccountNature(fsli As String) As String
    ' Determine if FSLI is Debit or Credit
    Dim fsliUpper As String
    fsliUpper = UCase(fsli)

    If InStr(fsliUpper, "REVENUE") > 0 Or InStr(fsliUpper, "INCOME") > 0 Or _
       InStr(fsliUpper, "LIABILITY") > 0 Or InStr(fsliUpper, "EQUITY") > 0 Then
        DetermineFSLIAccountNature = "Credit"
    Else
        DetermineFSLIAccountNature = "Debit"
    End If
End Function

Private Function GetPackName(packCode As String) As String
    '------------------------------------------------------------------------
    ' Look up pack name from Pack Number Company Table
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim packWs As Worksheet
    Dim lastRow As Long
    Dim row As Long

    Set packWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")

    If Not packWs Is Nothing Then
        lastRow = packWs.Cells(packWs.Rows.Count, 2).End(xlUp).row

        For row = 2 To lastRow
            If Trim(packWs.Cells(row, 2).Value) = packCode Then
                GetPackName = packWs.Cells(row, 1).Value
                Exit Function
            End If
        Next row
    End If

    ' Fallback
    GetPackName = packCode

    On Error GoTo 0
End Function
