Attribute VB_Name = "Module1"

Option Explicit

' Some missing movement/selection commands in Excel:

' Select current column(s), based on .CurrentRegion, including blank cells,
' also checking left and right for data
Private Sub Cmn_SelectColumn_Do(also_header As Boolean)
    Dim a As Range, b As Range, sh As Worksheet
    Dim firstrow As Long, lastrow As Long
    Dim i As Long, j As Long
    
    If (ActiveCell Is Nothing) Then
        Exit Sub
    End If
    
    ' find first and last row (track to left until a non-empty CurrentRegion is found)
    Set b = ActiveCell
    For i = 1 To 100
        Set a = b.CurrentRegion
        If (Not IsEmpty(b.Value)) Or (a.Count > 1) Then Exit For
        If (b.Column > 1) Then Set b = b.Offset(0, -1) Else Exit For
    Next i

    firstrow = a.Cells(1, 1).Row
    lastrow = firstrow + a.Rows.Count - 1

    ' now track 100 cells to the right to see if we get a CurrentRegion there, if we
    ' get a larger range, use that one instead
    Set b = ActiveCell
    For i = 1 To 100
        Set b = b.Offset(0, 1)
        Set a = b.CurrentRegion
        If (Not IsEmpty(b.Value)) Or (a.Count > 1) Then Exit For
    Next i
    
    firstrow = WorksheetFunction.Min(firstrow, a.Cells(1, 1).Row)
    lastrow = WorksheetFunction.Max(lastrow, a.Cells(1, 1).Row + a.Rows.Count - 1)

    ' header?
    If (Not also_header) And (lastrow - firstrow > 0) Then firstrow = firstrow + 1

    ' find all columns in current selection
    Set sh = Selection.Parent
    Set b = Nothing
    For i = 1 To Selection.Areas.Count
        For j = 1 To Selection.Areas(i).Columns.Count
            If (b Is Nothing) Then
                Set b = sh.Columns(Selection.Areas(i).Columns(j).Column)
            Else
                Set b = Union(b, sh.Columns(Selection.Areas(i).Columns(j).Column))
            End If
        Next j
    Next i
    Set b = Intersect(b, Range(sh.Rows(firstrow), sh.Rows(lastrow)))
    b.Select
End Sub

Public Sub Cmn_SelectColumn_OnlyData()
    Cmn_SelectColumn_Do (False)
End Sub

Public Sub Cmn_SelectColumn_WithHeader()
    Cmn_SelectColumn_Do (True)
End Sub

' Make pivot table tabular layout, no subtotals, show 0's
Public Sub Cmn_TabularPivot()
    Dim pt As PivotTable, fl As PivotField
    Set pt = Nothing
    On Error GoTo err
    Set pt = ActiveCell.PivotTable
    For Each fl In pt.PivotFields
        fl.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next fl
    pt.RowAxisLayout xlTabularRow
    pt.NullString = 0
err:
End Sub

' Replace formulas in visible cells with values
Public Sub Cmn_SelectionAsValues()
    Dim sareas() As Variant, sh As Worksheet, i As Long, r As Range
    On Error GoTo err
    Selection.SpecialCells(xlCellTypeVisible).Select
    Set sh = Selection.Parent
    ReDim sareas(1 To Selection.Areas.Count)
    For i = 1 To UBound(sareas)
        With Selection.Areas(i)
            sareas(i) = Array(.Row, .Column, .Rows.Count, .Columns.Count)
        End With
    Next i
    If (UBound(sareas) > 10) Then
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.ScreenUpdating = False
    End If
    For i = 1 To UBound(sareas)
        With sh.Cells
            Application.StatusBar = "Copy/Paste " & i & "/" & UBound(sareas)
            Set r = Range(.Item(sareas(i)(0), sareas(i)(1)), .Item(sareas(i)(0) + sareas(i)(2) - 1, sareas(i)(1) + sareas(i)(3) - 1))
            r.Copy
            r.PasteSpecial xlPasteValues
        End With
    Next i
err:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Custom VLOOKUP - check if cells to the left contains "Cmn_VLookup" and if yes, offset the columns
Function Cmn_VLookup(key As Range, table As Range, Optional start_at_offset As Long = 2) As Variant
    Dim clr As Range
    Set clr = Application.Caller
    Do While clr.Column() > 1
        Set clr = clr.Offset(0, -1)
        If (InStr(1, LCase(clr.Formula), LCase("Cmn_VLookup"), vbTextCompare) >= 1) Then
            start_at_offset = start_at_offset + 1
        Else
            Exit Do
        End If
    Loop
    Cmn_VLookup = WorksheetFunction.VLookup(key, table, start_at_offset, 0)
End Function

' Custom Index - acts as shorthand for =INDEX(result, MATCH(value, table, 0))
Function Cmn_Index(key As Range, table As Range, result As Range) As Variant
    Cmn_Index = WorksheetFunction.Index(result, WorksheetFunction.Match(key, table, 0))
End Function

