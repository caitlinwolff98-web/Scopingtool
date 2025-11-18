Attribute VB_Name = "Mod2_TabProcessing"
Option Explicit

' ============================================================================
' MODULE 2: WORKBOOK & TAB PROCESSING
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.0 - Complete Overhaul
' ============================================================================
' PURPOSE: Tab discovery, categorization, and validation
' DESCRIPTION: Manages the categorization of worksheets into predefined
'              categories with user guidance and confirmation
' ============================================================================

' ==================== CATEGORY CONSTANTS ====================
Private Const CAT_DIVISION As String = "Division"
Private Const CAT_DISCONTINUED As String = "Discontinued Operations"
Private Const CAT_INPUT_CONTINUING As String = "Input Continuing"
Private Const CAT_JOURNALS_CONTINUING As String = "Journals Continuing"
Private Const CAT_CONSOL_CONTINUING As String = "Consol Continuing"
Private Const CAT_TRIAL_BALANCE As String = "Trial Balance"
Private Const CAT_BALANCE_SHEET As String = "Balance Sheet"
Private Const CAT_INCOME_STATEMENT As String = "Income Statement"
Private Const CAT_UNCATEGORIZED As String = "Uncategorized"

' ==================== MAIN CATEGORIZATION FUNCTION ====================
Public Function CategorizeAllTabs(sourceWorkbook As Workbook, ByRef tabCategories As Object) As Boolean
    '------------------------------------------------------------------------
    ' Main function to categorize all tabs in the workbook
    ' Parameters:
    '   sourceWorkbook - The Stripe Packs consolidation workbook
    '   tabCategories - Dictionary to store categorizations (output)
    ' Returns: True if successful, False if cancelled
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim tabList As Collection
    Dim tabName As Variant
    Dim category As String
    Dim counter As Long

    ' Step 1: Discover all tabs
    Set tabList = New Collection
    For Each ws In sourceWorkbook.Worksheets
        tabList.Add ws.Name
    Next ws

    If tabList.Count = 0 Then
        MsgBox "No tabs found in the workbook.", vbExclamation
        CategorizeAllTabs = False
        Exit Function
    End If

    ' Step 2: Display tab discovery message
    MsgBox "TAB DISCOVERY" & vbCrLf & vbCrLf & _
           "Found " & tabList.Count & " tabs in the workbook." & vbCrLf & vbCrLf & _
           "You will now categorize each tab." & vbCrLf & vbCrLf & _
           "Available categories:" & vbCrLf & _
           "1. Division" & vbCrLf & _
           "2. Discontinued Operations" & vbCrLf & _
           "3. Input Continuing" & vbCrLf & _
           "4. Journals Continuing" & vbCrLf & _
           "5. Consol Continuing" & vbCrLf & _
           "6. Trial Balance" & vbCrLf & _
           "7. Balance Sheet" & vbCrLf & _
           "8. Income Statement" & vbCrLf & _
           "9. Uncategorized", _
           vbInformation, "Tab Discovery"

    ' Step 3: Categorize each tab
    counter = 1
    For Each tabName In tabList
        category = PromptForCategory(CStr(tabName), counter, tabList.Count)

        If category = "CANCEL" Then
            CategorizeAllTabs = False
            Exit Function
        End If

        tabCategories(tabName) = category
        counter = counter + 1
    Next tabName

    ' Step 4: Display categorization summary and confirm
    If Not ConfirmCategorization(tabCategories) Then
        ' User wants to recategorize
        tabCategories.RemoveAll
        CategorizeAllTabs = CategorizeAllTabs(sourceWorkbook, tabCategories) ' Recursive call
        Exit Function
    End If

    ' Step 5: Validate required categories exist
    If Not ValidateRequiredCategories(tabCategories) Then
        MsgBox "VALIDATION FAILED" & vbCrLf & vbCrLf & _
               "At least one tab must be categorized as 'Input Continuing'." & vbCrLf & _
               "This is required for processing.", _
               vbExclamation, "Missing Required Category"
        CategorizeAllTabs = False
        Exit Function
    End If

    CategorizeAllTabs = True
    Exit Function

ErrorHandler:
    MsgBox "Error during tab categorization: " & Err.Description, vbCritical
    CategorizeAllTabs = False
End Function

' ==================== CATEGORY PROMPT ====================
Private Function PromptForCategory(tabName As String, currentTab As Long, totalTabs As Long) As String
    '------------------------------------------------------------------------
    ' Prompt user to select category for a single tab
    ' Returns category name or "CANCEL"
    '------------------------------------------------------------------------
    Dim promptMsg As String
    Dim userInput As String
    Dim categoryNumber As Long

    promptMsg = "TAB CATEGORIZATION (" & currentTab & " of " & totalTabs & ")" & vbCrLf & vbCrLf & _
                "Tab Name: " & tabName & vbCrLf & _
                String(60, "-") & vbCrLf & _
                "Select category:" & vbCrLf & vbCrLf & _
                "1. Division (multiple allowed, will prompt for division name)" & vbCrLf & _
                "2. Discontinued Operations (single tab only)" & vbCrLf & _
                "3. Input Continuing (single tab only) *** REQUIRED ***" & vbCrLf & _
                "4. Journals Continuing (single tab only)" & vbCrLf & _
                "5. Consol Continuing (single tab only)" & vbCrLf & _
                "6. Trial Balance (single tab only)" & vbCrLf & _
                "7. Balance Sheet (single tab only)" & vbCrLf & _
                "8. Income Statement (single tab only)" & vbCrLf & _
                "9. Uncategorized (multiple allowed, will be ignored)" & vbCrLf & vbCrLf & _
                "Enter category number (or 'C' to cancel):"

    Do
        userInput = UCase(Trim(InputBox(promptMsg, "Categorize Tab: " & tabName, "3")))

        If userInput = "C" Then
            PromptForCategory = "CANCEL"
            Exit Function
        End If

        If userInput = "" Then
            userInput = "9" ' Default to Uncategorized if blank
        End If

        If IsNumeric(userInput) Then
            categoryNumber = CLng(userInput)

            If categoryNumber >= 1 And categoryNumber <= 9 Then
                ' Valid category selected
                Exit Do
            Else
                MsgBox "Invalid category number. Please enter 1-9.", vbExclamation
            End If
        Else
            MsgBox "Invalid input. Please enter a number from 1-9.", vbExclamation
        End If
    Loop

    ' Return category name based on number
    Select Case categoryNumber
        Case 1: PromptForCategory = CAT_DIVISION
        Case 2: PromptForCategory = CAT_DISCONTINUED
        Case 3: PromptForCategory = CAT_INPUT_CONTINUING
        Case 4: PromptForCategory = CAT_JOURNALS_CONTINUING
        Case 5: PromptForCategory = CAT_CONSOL_CONTINUING
        Case 6: PromptForCategory = CAT_TRIAL_BALANCE
        Case 7: PromptForCategory = CAT_BALANCE_SHEET
        Case 8: PromptForCategory = CAT_INCOME_STATEMENT
        Case 9: PromptForCategory = CAT_UNCATEGORIZED
    End Select
End Function

' ==================== CATEGORIZATION CONFIRMATION ====================
Private Function ConfirmCategorization(tabCategories As Object) As Boolean
    '------------------------------------------------------------------------
    ' Display summary of categorizations and confirm with user
    ' Returns True if confirmed, False if user wants to recategorize
    '------------------------------------------------------------------------
    Dim summaryMsg As String
    Dim tabName As Variant
    Dim category As String
    Dim categoryCount As Object
    Dim catName As Variant
    Dim uncategorizedList As String
    Dim result As VbMsgBoxResult

    ' Count tabs per category
    Set categoryCount = CreateObject("Scripting.Dictionary")
    uncategorizedList = ""

    For Each tabName In tabCategories.Keys
        category = tabCategories(tabName)

        If Not categoryCount.exists(category) Then
            categoryCount(category) = 0
        End If
        categoryCount(category) = categoryCount(category) + 1

        If category = CAT_UNCATEGORIZED Then
            If uncategorizedList <> "" Then uncategorizedList = uncategorizedList & ", "
            uncategorizedList = uncategorizedList & tabName
        End If
    Next tabName

    ' Build summary message
    summaryMsg = "CATEGORIZATION SUMMARY" & vbCrLf & vbCrLf & _
                 "Please review the categorization:" & vbCrLf & _
                 String(60, "-") & vbCrLf

    For Each catName In categoryCount.Keys
        summaryMsg = summaryMsg & catName & ": " & categoryCount(catName) & " tab(s)" & vbCrLf
    Next catName

    If uncategorizedList <> "" Then
        summaryMsg = summaryMsg & vbCrLf & "Uncategorized tabs:" & vbCrLf & uncategorizedList & vbCrLf
    End If

    summaryMsg = summaryMsg & vbCrLf & String(60, "-") & vbCrLf & _
                 "Are you happy with these categories?" & vbCrLf & vbCrLf & _
                 "Click YES to proceed" & vbCrLf & _
                 "Click NO to recategorize all tabs"

    result = MsgBox(summaryMsg, vbYesNo + vbQuestion, "Confirm Categorization")

    ConfirmCategorization = (result = vbYes)
End Function

' ==================== VALIDATION ====================
Private Function ValidateRequiredCategories(tabCategories As Object) As Boolean
    '------------------------------------------------------------------------
    ' Validate that all required categories have been assigned
    ' Required: At least one "Input Continuing" tab
    '------------------------------------------------------------------------
    Dim tabName As Variant
    Dim hasInputContinuing As Boolean

    hasInputContinuing = False

    For Each tabName In tabCategories.Keys
        If tabCategories(tabName) = CAT_INPUT_CONTINUING Then
            hasInputContinuing = True
            Exit For
        End If
    Next tabName

    ValidateRequiredCategories = hasInputContinuing
End Function

' ==================== PUBLIC HELPER FUNCTIONS ====================
Public Function GetTabByCategory(tabCategories As Object, categoryName As String) As Worksheet
    '------------------------------------------------------------------------
    ' Get the first tab assigned to a specific category
    ' Returns worksheet object or Nothing if not found
    '------------------------------------------------------------------------
    On Error Resume Next

    Dim tabName As Variant

    For Each tabName In tabCategories.Keys
        If tabCategories(tabName) = categoryName Then
            Set GetTabByCategory = Mod1_MainController.g_StripePacksWorkbook.Worksheets(CStr(tabName))
            Exit Function
        End If
    Next tabName

    Set GetTabByCategory = Nothing
End Function

Public Function GetAllTabsByCategory(tabCategories As Object, categoryName As String) As Collection
    '------------------------------------------------------------------------
    ' Get all tabs assigned to a specific category
    ' Returns collection of worksheet names
    '------------------------------------------------------------------------
    Dim result As Collection
    Dim tabName As Variant

    Set result = New Collection

    For Each tabName In tabCategories.Keys
        If tabCategories(tabName) = categoryName Then
            result.Add CStr(tabName)
        End If
    Next tabName

    Set GetAllTabsByCategory = result
End Function

Public Function CategoryExists(tabCategories As Object, categoryName As String) As Boolean
    '------------------------------------------------------------------------
    ' Check if any tab is categorized as the specified category
    '------------------------------------------------------------------------
    Dim tabName As Variant

    For Each tabName In tabCategories.Keys
        If tabCategories(tabName) = categoryName Then
            CategoryExists = True
            Exit Function
        End If
    Next tabName

    CategoryExists = False
End Function
