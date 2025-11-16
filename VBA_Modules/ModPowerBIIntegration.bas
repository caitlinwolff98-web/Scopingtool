Attribute VB_Name = "ModPowerBIIntegration"
Option Explicit

' ============================================================================
' MODULE: ModPowerBIIntegration
' PURPOSE: Enhanced Power BI integration features for entity scoping
' DESCRIPTION: Provides direct Power BI integration helpers including
'              metadata export, entity scoping configurations, and
'              threshold-based analysis support
' ============================================================================

' Create Power BI Metadata Sheet with tool information and configuration
Public Sub CreatePowerBIMetadata()
    On Error GoTo ErrorHandler
    
    Dim metaWs As Worksheet
    Dim row As Long
    
    ' Check if metadata sheet already exists
    On Error Resume Next
    Set metaWs = g_OutputWorkbook.Worksheets(ModConfig.POWERBI_METADATA_SHEET)
    On Error GoTo ErrorHandler
    
    If metaWs Is Nothing Then
        Set metaWs = g_OutputWorkbook.Worksheets.Add
        metaWs.Name = ModConfig.POWERBI_METADATA_SHEET
    Else
        metaWs.Cells.Clear
    End If
    
    ' Write metadata information
    row = 1
    With metaWs
        .Cells(row, 1).Value = "Metadata Property"
        .Cells(row, 2).Value = "Value"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "Tool Name"
        .Cells(row, 2).Value = ModConfig.TOOL_NAME
        row = row + 1
        
        .Cells(row, 1).Value = "Tool Version"
        .Cells(row, 2).Value = ModConfig.TOOL_VERSION
        row = row + 1
        
        .Cells(row, 1).Value = "Generated Date"
        .Cells(row, 2).Value = Now()
        .Cells(row, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        row = row + 1
        
        .Cells(row, 1).Value = "Source Workbook"
        .Cells(row, 2).Value = g_SourceWorkbook.Name
        row = row + 1
        
        .Cells(row, 1).Value = "Source Path"
        .Cells(row, 2).Value = g_SourceWorkbook.FullName
        row = row + 1
        
        ' Add table count information
        row = row + 1
        .Cells(row, 1).Value = "Tables Generated"
        .Cells(row, 2).Value = g_OutputWorkbook.Worksheets.count - 1 ' Exclude Control Panel
        row = row + 1
        
        ' Add category information
        row = row + 1
        .Cells(row, 1).Value = "Category"
        .Cells(row, 2).Value = "Tab Count"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Font.Bold = True
        row = row + 1
        
        Dim cat As Variant
        For Each cat In ModConfig.GetAllCategories()
            If g_TabCategories.Exists(cat) Then
                .Cells(row, 1).Value = cat
                .Cells(row, 2).Value = g_TabCategories(cat).count
                row = row + 1
            End If
        Next cat
        
        ' Add Power BI integration notes
        row = row + 1
        .Cells(row, 1).Value = "Power BI Integration Notes"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "1. Import all tables into Power BI Desktop"
        row = row + 1
        .Cells(row, 1).Value = "2. Use Power Query to unpivot data tables"
        row = row + 1
        .Cells(row, 1).Value = "3. Create relationships between tables"
        row = row + 1
        .Cells(row, 1).Value = "4. Use DAX measures for scoping analysis"
        row = row + 1
        .Cells(row, 1).Value = "5. See POWERBI_INTEGRATION_GUIDE.md for details"
        
        ' Auto-fit columns
        .Columns.AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    ModConfig.ShowError "Power BI Metadata Error", "Error creating Power BI metadata: " & Err.Description, Err.Number
End Sub

' Create Power BI Scoping Configuration Sheet
Public Sub CreatePowerBIScopingConfig()
    On Error GoTo ErrorHandler
    
    Dim scopeWs As Worksheet
    Dim row As Long
    
    ' Check if scoping sheet already exists
    On Error Resume Next
    Set scopeWs = g_OutputWorkbook.Worksheets(ModConfig.POWERBI_SCOPING_SHEET)
    On Error GoTo ErrorHandler
    
    If scopeWs Is Nothing Then
        Set scopeWs = g_OutputWorkbook.Worksheets.Add
        scopeWs.Name = ModConfig.POWERBI_SCOPING_SHEET
    Else
        scopeWs.Cells.Clear
    End If
    
    ' Write scoping configuration template
    row = 1
    With scopeWs
        .Cells(row, 1).Value = "Entity/Pack Name"
        .Cells(row, 2).Value = "Entity Code"
        .Cells(row, 3).Value = "Division"
        .Cells(row, 4).Value = "In Scope"
        .Cells(row, 5).Value = "Scope Reason"
        .Cells(row, 6).Value = "Threshold Met"
        .Cells(row, 7).Value = "Manual Selection"
        .Cells(row, 8).Value = "Comments"
        
        ' Format headers
        .Range(.Cells(1, 1), .Cells(1, 8)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 8)).Interior.Color = RGB(68, 114, 196)
        .Range(.Cells(1, 1), .Cells(1, 8)).Font.Color = RGB(255, 255, 255)
        
        row = row + 1
        
        ' Populate with entities from Pack Number Company Table
        Dim packWs As Worksheet
        On Error Resume Next
        Set packWs = g_OutputWorkbook.Worksheets("Pack Number Company Table")
        On Error GoTo ErrorHandler
        
        If Not packWs Is Nothing Then
            Dim lastRow As Long
            lastRow = packWs.Cells(packWs.Rows.count, 1).End(xlUp).row
            
            Dim i As Long
            For i = 2 To lastRow
                .Cells(row, 1).Value = packWs.Cells(i, 1).Value ' Pack Name
                .Cells(row, 2).Value = packWs.Cells(i, 2).Value ' Pack Code
                .Cells(row, 3).Value = packWs.Cells(i, 3).Value ' Division
                .Cells(row, 4).Value = "No" ' Default not in scope
                .Cells(row, 5).Value = "" ' Scope reason
                .Cells(row, 6).Value = "No" ' Threshold met
                .Cells(row, 7).Value = "No" ' Manual selection
                .Cells(row, 8).Value = "" ' Comments
                row = row + 1
            Next i
        End If
        
        ' Add instructions
        row = row + 2
        .Cells(row, 1).Value = "Instructions:"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "1. Use Power BI to identify entities meeting scoping thresholds"
        row = row + 1
        .Cells(row, 1).Value = "2. Update 'In Scope' column to 'Yes' for selected entities"
        row = row + 1
        .Cells(row, 1).Value = "3. Document reason in 'Scope Reason' column"
        row = row + 1
        .Cells(row, 1).Value = "4. Use 'Threshold Met' and 'Manual Selection' for tracking"
        row = row + 1
        .Cells(row, 1).Value = "5. Export this table back to Power BI for dashboard integration"
        
        ' Auto-fit columns
        .Columns.AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    ModConfig.ShowError "Power BI Scoping Error", "Error creating Power BI scoping config: " & Err.Description, Err.Number
End Sub

' Create DAX Measures documentation sheet
Public Sub CreateDAXMeasuresGuide()
    On Error GoTo ErrorHandler
    
    Dim daxWs As Worksheet
    Dim row As Long
    
    ' Create worksheet
    Set daxWs = g_OutputWorkbook.Worksheets.Add
    daxWs.Name = "DAX Measures Guide"
    
    row = 1
    With daxWs
        .Cells(row, 1).Value = "DAX Measure Templates for Scoping Analysis"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        row = row + 2
        
        ' Total Amount measure
        .Cells(row, 1).Value = "Measure 1: Total Amount"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "Total Amount = SUM('Full Input Table'[Amount])"
        .Cells(row, 1).Font.Name = "Consolas"
        row = row + 2
        
        ' Entity Count measure
        .Cells(row, 1).Value = "Measure 2: Entity Count"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "Entity Count = DISTINCTCOUNT('Full Input Table'[Pack])"
        .Cells(row, 1).Font.Name = "Consolas"
        row = row + 2
        
        ' Threshold Flag measure
        .Cells(row, 1).Value = "Measure 3: Threshold Flag"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "Threshold Flag = IF([Total Amount] > 300000000, ""Yes"", ""No"")"
        .Cells(row, 1).Font.Name = "Consolas"
        row = row + 2
        
        ' Coverage Percentage measure
        .Cells(row, 1).Value = "Measure 4: Coverage %"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "Coverage % = DIVIDE([Total Amount], CALCULATE([Total Amount], ALL('Full Input Table'[Pack])))"
        .Cells(row, 1).Font.Name = "Consolas"
        row = row + 2
        
        ' Scoped Entities measure
        .Cells(row, 1).Value = "Measure 5: Scoped Entities"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "Scoped Entities = CALCULATE([Entity Count], 'PowerBI_Scoping'[In Scope] = ""Yes"")"
        .Cells(row, 1).Font.Name = "Consolas"
        row = row + 2
        
        ' Scoping Percentage measure
        .Cells(row, 1).Value = "Measure 6: Scoping %"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "Scoping % = DIVIDE([Scoped Entities], [Entity Count])"
        .Cells(row, 1).Font.Name = "Consolas"
        row = row + 2
        
        ' Add notes
        row = row + 1
        .Cells(row, 1).Value = "Notes:"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        .Cells(row, 1).Value = "- Adjust threshold values (300000000) based on your scoping methodology"
        row = row + 1
        .Cells(row, 1).Value = "- Use these measures in Power BI visuals for interactive scoping"
        row = row + 1
        .Cells(row, 1).Value = "- Create calculated columns for more complex logic"
        row = row + 1
        .Cells(row, 1).Value = "- See POWERBI_INTEGRATION_GUIDE.md for complete examples"
        
        ' Auto-fit columns
        .Columns("A:A").ColumnWidth = 120
    End With
    
    Exit Sub
    
ErrorHandler:
    ModConfig.ShowError "DAX Guide Error", "Error creating DAX measures guide: " & Err.Description, Err.Number
End Sub

' Export entity scoping summary
Public Sub CreateEntityScopingSummary()
    On Error GoTo ErrorHandler
    
    Dim summaryWs As Worksheet
    Dim packWs As Worksheet
    Dim inputWs As Worksheet
    Dim row As Long
    Dim i As Long
    Dim lastRow As Long
    Dim packName As String
    Dim totalAmount As Double
    Dim packDict As Object
    
    ' Create summary worksheet
    Set summaryWs = g_OutputWorkbook.Worksheets.Add
    summaryWs.Name = "Entity Scoping Summary"
    
    ' Get Pack Number Company Table
    On Error Resume Next
    Set packWs = g_OutputWorkbook.Worksheets("Pack Number Company Table")
    Set inputWs = g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo ErrorHandler
    
    If packWs Is Nothing Or inputWs Is Nothing Then
        ModConfig.ShowWarning "Missing Tables", "Cannot create entity scoping summary. Required tables not found."
        Exit Sub
    End If
    
    ' Write headers
    row = 1
    With summaryWs
        .Cells(row, 1).Value = "Entity/Pack Name"
        .Cells(row, 2).Value = "Entity Code"
        .Cells(row, 3).Value = "Division"
        .Cells(row, 4).Value = "Total Amount"
        .Cells(row, 5).Value = "% of Total"
        .Cells(row, 6).Value = "Suggested for Scope"
        
        ' Format headers
        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 6)).Interior.Color = RGB(68, 114, 196)
        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Color = RGB(255, 255, 255)
        
        row = row + 1
        
        ' Calculate totals for each entity from Full Input Table
        lastRow = packWs.Cells(packWs.Rows.count, 1).End(xlUp).row
        
        For i = 2 To lastRow
            packName = packWs.Cells(i, 1).Value
            
            ' Calculate total amount for this pack
            totalAmount = CalculatePackTotal(inputWs, packName)
            
            .Cells(row, 1).Value = packName
            .Cells(row, 2).Value = packWs.Cells(i, 2).Value ' Pack Code
            .Cells(row, 3).Value = packWs.Cells(i, 3).Value ' Division
            .Cells(row, 4).Value = totalAmount
            .Cells(row, 4).NumberFormat = "#,##0.00"
            ' Calculate percentage in next pass
            .Cells(row, 6).Value = "" ' Will be filled by user or threshold logic
            
            row = row + 1
        Next i
        
        ' Calculate percentages
        Dim grandTotal As Double
        grandTotal = 0
        lastRow = .Cells(.Rows.count, 1).End(xlUp).row
        
        For i = 2 To lastRow
            If IsNumeric(.Cells(i, 4).Value) Then
                grandTotal = grandTotal + Abs(.Cells(i, 4).Value)
            End If
        Next i
        
        For i = 2 To lastRow
            If IsNumeric(.Cells(i, 4).Value) And grandTotal > 0 Then
                .Cells(i, 5).Value = Abs(.Cells(i, 4).Value) / grandTotal
                .Cells(i, 5).NumberFormat = "0.00%"
            End If
        Next i
        
        ' Auto-fit columns
        .Columns.AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    ModConfig.ShowError "Entity Summary Error", "Error creating entity scoping summary: " & Err.Description, Err.Number
End Sub

' Helper function to calculate total for a pack
Private Function CalculatePackTotal(inputWs As Worksheet, packName As String) As Double
    On Error Resume Next
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim total As Double
    Dim headerRow As Long
    
    total = 0
    headerRow = 1
    
    ' Find the column for this pack
    lastCol = inputWs.Cells(headerRow, inputWs.Columns.count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If inputWs.Cells(headerRow, col).Value = packName Then
            ' Sum all values in this column
            lastRow = inputWs.Cells(inputWs.Rows.count, col).End(xlUp).row
            
            For row = 2 To lastRow
                If IsNumeric(inputWs.Cells(row, col).Value) Then
                    total = total + Abs(inputWs.Cells(row, col).Value)
                End If
            Next row
            
            Exit For
        End If
    Next col
    
    CalculatePackTotal = total
    On Error GoTo 0
End Function

' Main entry point to create all Power BI integration assets
Public Sub CreateAllPowerBIAssets()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Creating Power BI integration assets..."
    
    ' Create metadata
    CreatePowerBIMetadata
    
    ' Create scoping configuration
    CreatePowerBIScopingConfig
    
    ' Create DAX measures guide
    CreateDAXMeasuresGuide
    
    ' Create entity scoping summary
    CreateEntityScopingSummary
    
    ' Create scoping control table for PowerBI
    CreateScopingControlTable
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ModConfig.ShowInfo "Power BI Assets Created", _
        "All Power BI integration assets have been created successfully!" & vbCrLf & vbCrLf & _
        "New sheets:" & vbCrLf & _
        "- " & ModConfig.POWERBI_METADATA_SHEET & vbCrLf & _
        "- " & ModConfig.POWERBI_SCOPING_SHEET & vbCrLf & _
        "- DAX Measures Guide" & vbCrLf & _
        "- Entity Scoping Summary" & vbCrLf & _
        "- Scoping Control Table (for dynamic PowerBI scoping)" & vbCrLf & vbCrLf & _
        "Import these into Power BI for enhanced scoping analysis."
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ModConfig.ShowError "Power BI Assets Error", "Error creating Power BI assets: " & Err.Description, Err.Number
End Sub

' Create Scoping Control Table for dynamic PowerBI scoping
Public Sub CreateScopingControlTable()
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
    Dim packDict As Object
    Dim division As String
    
    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Scoping Control Table"
    
    ' Get input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then Exit Sub
    
    ' Set up headers
    row = 1
    With ws
        .Cells(row, 1).Value = "Pack Name"
        .Cells(row, 2).Value = "Pack Code"
        .Cells(row, 3).Value = "Division"
        .Cells(row, 4).Value = "FSLi"
        .Cells(row, 5).Value = "Amount"
        .Cells(row, 6).Value = "Scoping Status"
        .Cells(row, 7).Value = "Is Consolidated"
        
        ' Format headers
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        row = row + 1
        
        ' Get dimensions
        lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
        lastRow = inputTab.Cells(inputTab.Rows.Count, 2).End(xlUp).row
        
        ' Create pack dictionary to get divisions
        Set packDict = CreateObject("Scripting.Dictionary")
        
        ' Iterate through each pack (column)
        For col = 3 To lastCol
            packCode = Trim(inputTab.Cells(8, col).Value)
            packName = Trim(inputTab.Cells(7, col).Value)
            
            If packCode <> "" And packName <> "" Then
                ' Get division for this pack
                division = GetPackDivisionFromTable(packCode)
                
                ' Iterate through each FSLi (row)
                For dataRow = 9 To lastRow
                    fsliName = Trim(inputTab.Cells(dataRow, 2).Value)
                    amount = inputTab.Cells(dataRow, col).Value
                    
                    ' Only include rows with FSLi names
                    If fsliName <> "" And Not ModDataProcessing.IsStatementHeader(fsliName) Then
                        .Cells(row, 1).Value = packName
                        .Cells(row, 2).Value = packCode
                        .Cells(row, 3).Value = division
                        .Cells(row, 4).Value = fsliName
                        
                        If IsNumeric(amount) Then
                            .Cells(row, 5).Value = CDbl(amount)
                            .Cells(row, 5).NumberFormat = "#,##0.00"
                        Else
                            .Cells(row, 5).Value = 0
                        End If
                        
                        ' Initial scoping status (to be updated in PowerBI)
                        .Cells(row, 6).Value = "Not Scoped"
                        
                        ' Mark if consolidated
                        If packCode = g_ConsolidatedPackCode Then
                            .Cells(row, 7).Value = "Yes"
                        Else
                            .Cells(row, 7).Value = "No"
                        End If
                        
                        row = row + 1
                    End If
                Next dataRow
            End If
        Next col
        
        ' Auto-fit columns
        .Columns("A:G").AutoFit
        
        ' Create table
        If row > 2 Then
            Dim tbl As ListObject
            On Error Resume Next
            Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(row - 1, 7)), , xlYes)
            If Not tbl Is Nothing Then
                tbl.Name = "Scoping_Control_Table"
                tbl.TableStyle = "TableStyleMedium2"
            End If
            On Error GoTo ErrorHandler
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error creating Scoping Control Table: " & Err.Description
End Sub

' Helper function to get division from Pack Number Company Table
Private Function GetPackDivisionFromTable(packCode As String) As String
    On Error Resume Next
    
    Dim packWs As Worksheet
    Dim lastRow As Long
    Dim row As Long
    
    ' Try to find in Pack Number Company Table
    Set packWs = g_OutputWorkbook.Worksheets("Pack Number Company Table")
    
    If Not packWs Is Nothing Then
        lastRow = packWs.Cells(packWs.Rows.Count, 2).End(xlUp).row
        
        For row = 2 To lastRow
            If Trim(packWs.Cells(row, 2).Value) = packCode Then
                GetPackDivisionFromTable = Trim(packWs.Cells(row, 3).Value)
                Exit Function
            End If
        Next row
    End If
    
    ' Default if not found
    GetPackDivisionFromTable = "Unknown"
    
    On Error GoTo 0
End Function
