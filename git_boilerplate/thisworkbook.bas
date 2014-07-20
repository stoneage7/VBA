
Option Explicit

' Import/Export Excel modules when saving/opening .xlsm in Excel (for version control of sources)
' Belongs in ThisWorkbook module
' Requires "Microsoft Visual Basic for Application Extensibility" in Tools/References

Private Const IsTesting = True
Private Const SourceDir = "src_myproject"
Private Const ButtonName = "WTF"
Private Const ButtonAction = "DOWTF()"

Private Sub ExportModules(to_dir As String)
    Dim comp As VBComponent
    Dim ext As String
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule: ext = ".bas"
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_MSForm: ext = ".frm"
            Case Else: ext = vbNullString
        End Select
        If (ext <> vbNullString) Then
            Debug.Print "export: " & comp.name & ext
            comp.Export to_dir & "\" & comp.name & ext
        End If
    Next comp
End Sub

Private Sub ImportModules(from_dir As String)
    Dim i As Long, fn As String
    Dim comp As VBComponent
    With ThisWorkbook.VBProject.VBComponents
        For i = .Count To 1 Step -1
            Select Case .Item(i).Type
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm: .Remove .Item(i)
                Case Else: ' skip
            End Select
        Next i
        fn = Dir(from_dir & "\", vbNormal)
        Do While fn <> vbNullString
            Select Case Right(fn, 4)
                Case ".bas", ".cls":
                    .Import from_dir & "\" & fn
                Case ".frm":
                    .Import from_dir & "\" & fn
                    With ThisWorkbook.VBProject.VBComponents(Left(fn, Len(fn) - 4)).CodeModule
                        If (.Lines(1, 1) = vbNullString) Then .DeleteLines (1)
                    End With
                Case Else: ' skip
            End Select
            fn = Dir()
        Loop
    End With
End Sub

Private Sub Workbook_Open()
    Dim cControl As CommandBarButton
    On Error Resume Next
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls(ButtonName)
    On Error GoTo 0
    If (cControl Is Nothing) Then
        Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
        With cControl
            .Caption = ButtonName
            .Style = msoButtonCaption
            .OnAction = ButtonAction 'Macro stored in a Standard Module
        End With
    End If
    If (IsTesting) Then
        Dim d As String
        d = ThisWorkbook.path & "\" & SourceDir
        If (Dir(d, vbDirectory) <> vbNullString) Then ImportModules (d)
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal gui As Boolean, cancel As Boolean)
    If (IsTesting) Then
        Dim d As String
        Debug.Print "Export Modules"
        d = ThisWorkbook.path & "\" & SourceDir
        If (Dir(d, vbDirectory) <> vbNullString) Then ExportModules (d)
    End If
End Sub

Private Sub Workbook_BeforeClose(cancel As Boolean)
    On Error Resume Next 'In case it has already gone.
    err.Clear
    Do While err.Number = 0
        Application.CommandBars("Worksheet Menu Bar").Controls(ButtonName).Delete
    Loop
End Sub

Public Sub ForceExport()
    ExportModules ThisWorkbook.path & "\" & SourceDir
End Sub
