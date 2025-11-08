# Contributing to TGK Consolidation Scoping Tool

Thank you for your interest in contributing! This document provides guidelines and instructions for contributing to this project.

## Table of Contents
1. [Code of Conduct](#code-of-conduct)
2. [How Can I Contribute?](#how-can-i-contribute)
3. [Development Setup](#development-setup)
4. [Coding Standards](#coding-standards)
5. [Submitting Changes](#submitting-changes)
6. [Documentation](#documentation)

---

## Code of Conduct

### Our Pledge
We are committed to providing a welcoming and inspiring community for all.

### Expected Behavior
- Be respectful and considerate
- Provide constructive feedback
- Focus on what's best for the community
- Show empathy towards others

### Unacceptable Behavior
- Harassment or discrimination
- Trolling or insulting comments
- Publishing private information
- Unprofessional conduct

---

## How Can I Contribute?

### Reporting Bugs

**Before submitting a bug report:**
- Check the FAQ.md for common issues
- Check existing issues to avoid duplicates
- Collect information about the bug

**What to include in a bug report:**
- Tool version (see CHANGELOG.md)
- Excel version and build number
- Operating system
- Clear description of the issue
- Steps to reproduce
- Expected vs actual behavior
- Error messages (full text)
- Sample data structure (if possible, anonymized)

**Template:**
```markdown
**Tool Version:** 1.0.0
**Excel Version:** Excel 2021 Build 16.0.14332.20447
**OS:** Windows 11

**Description:**
[Clear description of the bug]

**Steps to Reproduce:**
1. Open consolidation workbook with...
2. Run tool and select...
3. At step X, see error...

**Expected Behavior:**
[What should happen]

**Actual Behavior:**
[What actually happens]

**Error Message:**
```
[Exact error text]
```

**Additional Context:**
[Any other relevant information]
```

### Suggesting Enhancements

**Before suggesting an enhancement:**
- Check if it already exists in a newer version
- Check if it's already been suggested
- Consider if it fits the tool's scope

**What to include:**
- Clear description of the enhancement
- Use case / business value
- Example scenario
- Impact on existing functionality
- Alternative approaches considered

**Template:**
```markdown
**Enhancement Title:** [Clear, concise title]

**Use Case:**
[Description of the business need]

**Proposed Solution:**
[How you think it should work]

**Benefits:**
- Benefit 1
- Benefit 2

**Example Scenario:**
[Step-by-step example of using the enhancement]

**Alternatives Considered:**
[Other approaches you thought about]

**Additional Context:**
[Any other relevant information]
```

### Contributing Code

We welcome code contributions! Areas where contributions are particularly valuable:

**Priority Areas:**
- Multi-language support
- Performance optimizations
- Enhanced error handling
- Additional table formats
- Custom format detection
- Test automation

**Secondary Areas:**
- Documentation improvements
- Code refactoring
- Bug fixes
- UI enhancements

---

## Development Setup

### Prerequisites
- Microsoft Excel 2016 or later
- VBA Editor access
- Basic VBA knowledge
- Git installed

### Getting Started

1. **Fork the Repository**
   ```bash
   # On GitHub, click "Fork" button
   ```

2. **Clone Your Fork**
   ```bash
   git clone https://github.com/YOUR-USERNAME/Scopingtool.git
   cd Scopingtool
   ```

3. **Set Up Excel Development Environment**
   - Open Excel
   - Create new macro-enabled workbook
   - Alt+F11 to open VBA Editor
   - Import existing .bas files
   - Tools → References → Enable Microsoft Scripting Runtime

4. **Create a Branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b bugfix/your-bugfix-name
   ```

5. **Make Your Changes**
   - Edit VBA code in Excel VBA Editor
   - Export modules: Right-click module → Export File
   - Save to VBA_Modules folder
   - Update documentation if needed

6. **Test Your Changes**
   - Test with various consolidation structures
   - Test edge cases
   - Verify no regression in existing functionality
   - Document test cases

7. **Commit Changes**
   ```bash
   git add VBA_Modules/ModuleName.bas
   git commit -m "Add: Brief description of changes"
   ```

8. **Push to Your Fork**
   ```bash
   git push origin feature/your-feature-name
   ```

9. **Create Pull Request**
   - Go to original repository on GitHub
   - Click "New Pull Request"
   - Select your branch
   - Fill in pull request template

---

## Coding Standards

### VBA Style Guide

#### Naming Conventions

**Modules:**
```vba
' Use descriptive names with Mod prefix
Mod Main
ModTabCategorization
ModDataProcessing
ModTableGeneration
```

**Functions and Subroutines:**
```vba
' Use PascalCase for public procedures
Public Sub StartScopingTool()
Public Function GetWorkbookName() As String

' Use PascalCase for private procedures
Private Sub ProcessInputTab()
Private Function ValidateData() As Boolean
```

**Variables:**
```vba
' Use camelCase for local variables
Dim workbookName As String
Dim rowIndex As Long
Dim isValid As Boolean

' Use descriptive names
Dim packList As Collection  ' Good
Dim pl As Collection        ' Bad

' Use prefixes for module-level variables
Private m_TabCategories() As TabCategory
Private m_TabCount As Long

' Use g_ prefix for global variables
Public g_SourceWorkbook As Workbook
Public g_OutputWorkbook As Workbook
```

**Constants:**
```vba
' Use UPPER_CASE for constants
Public Const CAT_SEGMENT = "TGK Segment Tabs"
Public Const MAX_ENTITIES = 500
Private Const DEFAULT_TIMEOUT = 30
```

#### Code Structure

**Module Header:**
```vba
Attribute VB_Name = "ModuleName"
Option Explicit

' ============================================================================
' MODULE: ModuleName
' PURPOSE: Brief description of module purpose
' DESCRIPTION: Detailed description of what this module does and how it fits
'              into the overall architecture
' ============================================================================
```

**Function Header:**
```vba
' Brief description of function
' Parameters:
'   paramName - Description of parameter
' Returns:
'   Description of return value
Public Function FunctionName(paramName As String) As Boolean
    ' Implementation
End Function
```

**Error Handling:**
```vba
Public Sub SubroutineName()
    On Error GoTo ErrorHandler
    
    ' Main code here
    
    Exit Sub
    
ErrorHandler:
    ' User-friendly error message
    MsgBox "Error in SubroutineName: " & Err.Description, vbCritical
    ' Log or handle error appropriately
End Sub
```

**Comments:**
```vba
' Use comments to explain WHY, not WHAT
' Good:
' Calculate percentage using absolute values to handle negative amounts
percentValue = (Abs(amount) / total) * 100

' Bad:
' Calculate percentage
percentValue = (Abs(amount) / total) * 100
```

#### Best Practices

**Always Use Option Explicit:**
```vba
' At top of every module
Option Explicit
```

**Avoid Magic Numbers:**
```vba
' Bad:
If rowIndex = 6 Then

' Good:
Const HEADER_ROW = 6
If rowIndex = HEADER_ROW Then
```

**Use Meaningful Variable Names:**
```vba
' Bad:
Dim x As Long
For x = 1 To 10
    ' ...
Next x

' Good:
Dim entityIndex As Long
For entityIndex = 1 To entityCount
    ' ...
Next entityIndex
```

**Limit Function Length:**
- Aim for functions under 50 lines
- Extract complex logic to separate functions
- One function, one purpose

**Avoid Deep Nesting:**
```vba
' Bad: Deep nesting (4+ levels)
If condition1 Then
    If condition2 Then
        If condition3 Then
            If condition4 Then
                ' Do something
            End If
        End If
    End If
End If

' Good: Early returns
If Not condition1 Then Exit Sub
If Not condition2 Then Exit Sub
If Not condition3 Then Exit Sub
If Not condition4 Then Exit Sub
' Do something
```

### Performance Considerations

**Disable Screen Updating for Intensive Operations:**
```vba
Application.ScreenUpdating = False
' ... intensive operations ...
Application.ScreenUpdating = True
```

**Use Manual Calculation:**
```vba
Application.Calculation = xlCalculationManual
' ... operations ...
Application.Calculation = xlCalculationAutomatic
```

**Avoid Select and Activate:**
```vba
' Bad:
Sheets("Sheet1").Select
Range("A1").Select
Selection.Value = "Hello"

' Good:
Sheets("Sheet1").Range("A1").Value = "Hello"
```

**Use With Blocks:**
```vba
' Bad:
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Size = 12
ws.Range("A1").Interior.Color = RGB(200, 200, 200)

' Good:
With ws.Range("A1")
    .Font.Bold = True
    .Font.Size = 12
    .Interior.Color = RGB(200, 200, 200)
End With
```

---

## Submitting Changes

### Pull Request Process

1. **Update Documentation**
   - Update README.md if needed
   - Update relevant documentation files
   - Add entry to CHANGELOG.md

2. **Ensure Quality**
   - Code follows style guide
   - No compilation errors
   - Tested with sample data
   - No breaking changes to existing functionality

3. **Write Good Commit Messages**
   ```
   Add: Feature description
   Fix: Bug description
   Update: Documentation description
   Refactor: Code improvement description
   ```

   Example:
   ```
   Add: Multi-language support for French consolidations
   
   - Added language detection in ModDataProcessing
   - Updated column detection to handle French headers
   - Added new constants for French terminology
   - Updated documentation with French examples
   ```

4. **Fill Pull Request Template**
   - Clear title
   - Description of changes
   - Related issues
   - Testing performed
   - Breaking changes (if any)

### Review Process

- Maintainers will review your pull request
- May request changes or clarifications
- Be responsive to feedback
- Once approved, will be merged

### After Merge

- Your contribution will be included in next release
- You'll be credited in CHANGELOG.md
- Consider contributing more!

---

## Documentation

### Documentation Standards

**When to Update Documentation:**
- Adding new features
- Changing existing behavior
- Fixing bugs that affect usage
- Adding examples
- Improving clarity

**Files to Consider:**
- `README.md` - Overview and quick start
- `DOCUMENTATION.md` - Complete technical documentation
- `INSTALLATION_GUIDE.md` - Installation instructions
- `USAGE_EXAMPLES.md` - Usage scenarios
- `POWERBI_INTEGRATION_GUIDE.md` - Power BI specifics
- `FAQ.md` - Common questions
- `CHANGELOG.md` - Version history
- VBA code comments - Inline documentation

### Documentation Style

**Clear and Concise:**
```markdown
# Good
To install, press Alt+F11 and import the .bas files.

# Bad
In order to proceed with the installation process, you need to open the VBA Editor by pressing the Alt and F11 keys simultaneously, and then navigate to the File menu...
```

**Use Examples:**
```markdown
Set threshold to $300,000,000 to automatically scope in high-revenue entities.

Example: If Net Revenue > $300M, include all FSLis for that pack.
```

**Include Screenshots (when helpful):**
- UI elements
- Configuration screens
- Expected output
- Power BI dashboards

---

## Recognition

Contributors will be recognized in:
- CHANGELOG.md (for significant contributions)
- Git commit history
- GitHub contributors page

---

## Questions?

If you have questions about contributing:
1. Check this guide
2. Review existing code and documentation
3. Open an issue with your question
4. Tag it as "question"

---

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

**Thank you for contributing to the TGK Consolidation Scoping Tool!**

Your contributions help make consolidation scoping easier for audit professionals worldwide.
