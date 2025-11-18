Attribute VB_Name = "Mod5_ScopingEngine"
Option Explicit

' ============================================================================
' MODULE 5: SCOPING ENGINE
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.0 - Complete Overhaul
' ============================================================================
' PURPOSE: Threshold configuration, automatic scoping, and manual scoping
' DESCRIPTION: Manages threshold-based automatic scoping and provides
'              interface for manual pack/FSLI scoping decisions
' ============================================================================

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

    ' Step 2: Display FSLI selection prompt
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
        If InStr(packName, consolEntity) > 0 Then GoTo NextRow

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

' ==================== MANUAL SCOPING INTERFACE ====================
Public Sub EnableManualScoping()
    '------------------------------------------------------------------------
    ' Enable manual scoping interface in Dashboard
    ' Allows users to scope in/out specific packs or pack-FSLI combinations
    '------------------------------------------------------------------------
    ' This is implemented in Mod6_DashboardGeneration
    ' Manual scoping is done through interactive controls in the dashboard
End Sub

Public Function ScopeInPack(packCode As String) As Boolean
    '------------------------------------------------------------------------
    ' Manually scope in entire pack (all FSLIs)
    '------------------------------------------------------------------------
    ' Implementation: Add pack to manual scoping dictionary
    ' Mark all FSLIs for this pack as "Scoped In"
    ScopeInPack = True
End Function

Public Function ScopeInPackFSLI(packCode As String, fsli As String) As Boolean
    '------------------------------------------------------------------------
    ' Manually scope in specific FSLI for specific pack
    '------------------------------------------------------------------------
    ' Implementation: Add pack-FSLI combination to manual scoping dictionary
    Dim key As String
    key = packCode & "|" & fsli

    If Not Mod1_MainController.g_ManualScoping.exists(key) Then
        Mod1_MainController.g_ManualScoping(key) = "Scoped In"
    End If

    ScopeInPackFSLI = True
End Function

Public Function ScopeOutPackFSLI(packCode As String, fsli As String) As Boolean
    '------------------------------------------------------------------------
    ' Remove pack-FSLI from scoping
    '------------------------------------------------------------------------
    Dim key As String
    key = packCode & "|" & fsli

    If Mod1_MainController.g_ManualScoping.exists(key) Then
        Mod1_MainController.g_ManualScoping.Remove key
    End If

    ScopeOutPackFSLI = True
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

' ==================== GENERATE SCOPING REPORTS ====================
Public Sub GenerateScopingSummary(scopedPacks As Object, thresholds As Collection)
    '------------------------------------------------------------------------
    ' Generate Scoping Summary sheet with recommendations
    '------------------------------------------------------------------------
    Dim outputWs As Worksheet
    Dim row As Long
    Dim packCode As Variant

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Scoping Summary"

    ' Write title
    outputWs.Cells(1, 1).Value = "SCOPING SUMMARY"
    outputWs.Cells(1, 1).Font.Size = 14
    outputWs.Cells(1, 1).Font.Bold = True

    ' Write headers
    row = 3
    outputWs.Cells(row, 1).Value = "Pack Code"
    outputWs.Cells(row, 2).Value = "Scoping Status"
    outputWs.Cells(row, 3).Value = "Triggering FSLI"
    outputWs.Cells(row, 4).Value = "Recommendation"

    outputWs.Range("A3:D3").Font.Bold = True
    outputWs.Range("A3:D3").Interior.Color = RGB(68, 114, 196)
    outputWs.Range("A3:D3").Font.Color = RGB(255, 255, 255)

    row = 4

    For Each packCode In scopedPacks.Keys
        outputWs.Cells(row, 1).Value = packCode
        outputWs.Cells(row, 2).Value = "Automatically Scoped In"
        outputWs.Cells(row, 3).Value = scopedPacks(packCode)
        outputWs.Cells(row, 4).Value = "Include in audit scope"

        outputWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206) ' Green

        row = row + 1
    Next packCode

    ' Summary statistics
    row = row + 2
    outputWs.Cells(row, 1).Value = "Total Packs Scoped In:"
    outputWs.Cells(row, 2).Value = scopedPacks.Count
    outputWs.Cells(row, 1).Font.Bold = True

    outputWs.Columns.AutoFit
End Sub

Public Sub GenerateThresholdConfigSheet(thresholds As Collection)
    '------------------------------------------------------------------------
    ' Document threshold configuration in output workbook
    '------------------------------------------------------------------------
    Dim outputWs As Worksheet
    Dim row As Long
    Dim i As Long

    Set outputWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Threshold Configuration"

    ' Write title
    outputWs.Cells(1, 1).Value = "THRESHOLD CONFIGURATION"
    outputWs.Cells(1, 1).Font.Size = 14
    outputWs.Cells(1, 1).Font.Bold = True

    ' Write headers
    row = 3
    outputWs.Cells(row, 1).Value = "FSLI"
    outputWs.Cells(row, 2).Value = "Threshold Amount"

    outputWs.Range("A3:B3").Font.Bold = True
    outputWs.Range("A3:B3").Interior.Color = RGB(68, 114, 196)
    outputWs.Range("A3:B3").Font.Color = RGB(255, 255, 255)

    row = 4

    For i = 1 To thresholds.Count
        outputWs.Cells(row, 1).Value = thresholds(i)("FSLI")
        outputWs.Cells(row, 2).Value = thresholds(i)("Amount")
        outputWs.Cells(row, 2).NumberFormat = "#,##0.00"
        row = row + 1
    Next i

    outputWs.Columns.AutoFit
End Sub
