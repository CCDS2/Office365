1. **Loop through all worksheets and perform an action**:
   This script loops through all worksheets in the active workbook and performs a specific action, such as formatting or data manipulation.

```vba
Sub LoopThroughWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Your code here
        ' For example:
        ' ws.Cells.Font.Bold = True
    Next ws
End Sub
```

2. **Find and replace values**:
   This script finds a specific value in a worksheet and replaces it with another value.

```vba
Sub FindAndReplace()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    ws.Cells.Replace What:="oldValue", Replacement:="newValue", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False
End Sub
```

3. **Sort data in a range**:
   This script sorts data in a specific range based on a chosen criteria.

```vba
Sub SortData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    ws.Range("A1:B10").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub
```

4. **Insert a new worksheet and name it**:
   This script inserts a new worksheet into the workbook and renames it.

```vba
Sub AddNewWorksheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "NewSheet"
End Sub
```

5. **Protect and unprotect worksheets**:
   This script protects or unprotects a worksheet with a specified password.

```vba
Sub ProtectWorksheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    ws.Protect Password:="password"
End Sub

Sub UnprotectWorksheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    ws.Unprotect Password:="password"
End Sub
```

6. **Copy data from one range to another**:
   This script copies data from one range to another within the same worksheet.

```vba
Sub CopyData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    ws.Range("A1:B10").Copy Destination:=ws.Range("C1")
End Sub
```

7. **Delete blank rows in a range**:
   This script deletes blank rows within a specified range.

```vba
Sub DeleteBlankRows()
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    Set rng = ws.Range("A1:A10") ' Change to your range
    rng.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub
```

8. **Calculate the sum, average, and count of a range**:
   This script calculates the sum, average, and count of values within a specified range.

```vba
Sub CalculateStatistics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    Dim rng As Range
    Set rng = ws.Range("A1:A10") ' Change to your range
    MsgBox "Sum: " & Application.WorksheetFunction.Sum(rng) & vbCrLf & _
           "Average: " & Application.WorksheetFunction.Average(rng) & vbCrLf & _
           "Count: " & Application.WorksheetFunction.Count(rng)
End Sub
```

9. **Create a chart**:
   This script creates a chart based on data in a specified range.

```vba
Sub CreateChart()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cht As ChartObject
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change to your sheet name
    Set rng = ws.Range("A1:B10") ' Change to your data range
    Set cht = ws.ChartObjects.Add(Left:=100, Width:=375, Top:=75, Height:=225)
    With cht.Chart
        .SetSourceData Source:=rng
        .ChartType = xlColumnClustered ' Change to your desired chart type
    End With
End Sub
```

10. **Run a macro automatically when opening workbook**:
    This script runs a specified macro automatically when the workbook is opened.

```vba
Private Sub Workbook_Open()
    ' Call your macro here
    ' For example:
    ' Call MyMacro
End Sub
```

These VBA scripts cover a range of common tasks and can be adapted to suit your specific needs within Excel.
