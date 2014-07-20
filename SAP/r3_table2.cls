VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "r3_table2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type tbl_column
    header As String
    y() As String
End Type

Private x() As tbl_column
Private y_extents As r3_table2

Public Sub alloc(columns As Long, rows As Long)
    Dim i As Long
    ReDim x(1 To columns)
    If (rows > 0) Then For i = 1 To columns: ReDim x(i).y(1 To rows): Next i
End Sub

Public Sub erase_table()
    Set y_extents = Nothing
    Erase x
End Sub

Public Sub resize(rows As Long)
    Dim i As Long
    If (rows > priv_ysz()) Then
        If Not (y_extents Is Nothing) Then
            y_extents.resize (rows - priv_ysz())
            Exit Sub
        End If
    ElseIf (rows = 0) Then
        For i = LBound(x) To UBound(x): Erase x(i).y: Next i
    Else
        For i = LBound(x) To UBound(x): ReDim Preserve x(i).y(1 To rows): Next i
    End If
End Sub

Public Property Get row_count() As Long
    Dim i As Long
    row_count = priv_ysz()
    If Not (y_extents Is Nothing) Then row_count = row_count + y_extents.row_count
End Property

Public Property Get col_count() As Long
    If Not Not x Then
        col_count = UBound(x)
    Else
        col_count = 0
    End If
End Property

Public Property Get header(x_ As Long) As String
    header = x(x_).header
End Property

Public Property Let header(x_ As Long, value As String)
    x(x_).header = value
End Property

Public Property Get header_x(h As String) As Long
    Dim i As Long
    For i = LBound(x) To UBound(x)
        If (x(i).header = h) Then
            header_x = i
            Exit Function
        End If
    Next i
    Err.Raise 9
End Property

Public Sub set_headers(h As Variant)
    Dim i As Long
    If (UBound(h) - LBound(h) + 1) <> UBound(x) Then Err.Raise 9
    For i = LBound(x) To UBound(x)
        x(i).header = h(LBound(h) + i - LBound(x))
    Next i
End Sub

Public Property Get item(x_ As Long, y_ As Long) As String
    Dim e As r3_table2, offset As Long
    priv_getextent y_, e, offset
    If (e Is Nothing) Then Err.Raise 5
    item = e.priv_xy(x_, y_ - offset)
End Property

Public Property Let item(x_ As Long, y_ As Long, value As String)
    Dim e As r3_table2, offset As Long
    priv_getextent y_, e, offset
    If (e Is Nothing) Then Err.Raise 5
    e.priv_xy(x_, y_ - offset) = value
End Property

Private Sub priv_getextent(y As Long, ByRef o_extent As r3_table2, ByRef o_offset As Long)
    Set o_extent = Me: o_offset = 0
    Do While y > o_offset + o_extent.priv_ysz()
        o_offset = o_offset + o_extent.priv_ysz()
        Set o_extent = o_extent.priv_yext()
    Loop
End Sub

Public Property Get priv_xy(x_ As Long, y_ As Long) As String
    priv_xy = x(x_).y(y_)
End Property

Public Property Let priv_xy(x_ As Long, y_ As Long, value As String)
    x(x_).y(y_) = value
End Property

Public Property Get item_h(header As String, y_ As Long) As String
    item_h = item(header_x(header), y_)
End Property

Public Property Let item_h(header As String, y_ As Long, value As String)
    item(header_x(header), y_) = value
End Property

Public Sub add_column(Optional t2 As r3_table2, Optional header As String)
    Dim i As Long, nx As Long
    nx = UBound(x) + 1
    ReDim Preserve x(1 To nx)
    If (priv_ysz() > 0) Then ReDim x(nx).y(1 To priv_ysz())
    If Not (y_extents Is Nothing) Then y_extents.add_column header:=header
    x(nx).header = header
    If Not (t2 Is Nothing) Then
        If (t2.col_count <> 1) Or (t2.row_count <> Me.row_count) Then Err.Raise 9
        If x(nx).header = vbNullString Then x(nx).header = t2.header(1)
        For i = 1 To row_count(): item(nx, i) = t2.item(1, i): Next i
    End If
End Sub

Public Sub append(t2 As r3_table2)
    If (t2 Is Nothing) Then Exit Sub
    If (t2.row_count = 0) Then Exit Sub
    If (Me.col_count = 0) Then
        Dim i As Long
        Me.alloc t2.col_count, 0
        For i = 1 To t2.col_count
            x(i).header = t2.header(i)
        Next i
    ElseIf (t2.col_count <> Me.col_count) Then
        Exit Sub
    End If
    If (y_extents Is Nothing) Then
        Set y_extents = t2
    Else
        y_extents.append t2
    End If
End Sub

Public Property Get appendc(t2 As r3_table2) As r3_table2
    Set appendc = Me
    Me.append t2
End Property

Public Sub union(c As Collection)
    Dim v As Variant, t As r3_table2
    If (c Is Nothing) Then Exit Sub
    For Each v In c
        Set t = v
        Me.append t
    Next v
End Sub

Public Property Get unionc(c As Collection) As r3_table2
    Set unionc = Me
    Me.union c
End Property

' if nosimple is true -> will duplicate rows on left side if there's more than one match on right side
Public Sub equijoin(t2 As r3_table2, l_eq As Variant, r_eq As Variant, r_add As Variant, Optional nosimple As Boolean)
    Dim r_index As Collection, row As Variant, i As Long, y As Long, after As Long, r_ln As Variant
    Dim ndups As Long, dups As Collection

    If (t2 Is Nothing) Then Exit Sub
    'If (t2.row_count = 0) Or (Me.row_count = 0) Then Exit Sub
    'Debug.Assert (LBound(l_eq) = LBound(r_eq)) And (UBound(l_eq) = UBound(r_eq))

    after = Me.col_count
    For i = LBound(r_add) To UBound(r_add)
        Me.add_column header:=t2.header(CLng(r_add(i)))
    Next i
    
    If (t2.row_count = 0) Or (Me.row_count = 0) Then Exit Sub

    If (nosimple) Then
        Set r_index = t2.priv_mkindex(r_eq)
    Else
        Set r_index = t2.select_distinct(r_eq)
    End If

    Set dups = New Collection
    ReDim row(LBound(l_eq) To UBound(r_eq))
    For y = 1 To Me.row_count
        For i = LBound(l_eq) To UBound(l_eq): row(i) = Replace(Me.item(CLng(l_eq(i)), y), "_", "__"): Next i
        r_ln = Empty
        On Error Resume Next
        r_ln = r_index.item(Join(row, "_"))
        On Error GoTo 0
        If IsArray(r_ln) Then
            ndups = ndups + UBound(r_ln) - LBound(r_ln)
            For i = LBound(r_ln) + 1 To UBound(r_ln): dups.add Array(CLng(y), CLng(r_ln(i))): Next i
            r_ln = r_ln(0)
        End If
        If (r_ln <> 0) Then
            For i = LBound(r_add) To UBound(r_add)
                    Me.item(after + i - LBound(r_add) + 1, y) = t2.item(CLng(r_add(i)), CLng(r_ln))
            Next i
        End If
    Next y
    If (dups.Count > 0) Then
        Dim add As r3_table2
        Set add = New r3_table2
        add.alloc Me.col_count, dups.Count
        y = 1
        For Each r_ln In dups
            For i = 1 To Me.col_count
                add.item(i, y) = Me.item(i, CLng(r_ln(0)))
            Next i
            For i = LBound(r_add) To UBound(r_add)
                add.item(i - LBound(r_add) + Me.col_count - UBound(r_add), y) = t2.item(CLng(r_add(i)), CLng(r_ln(1)))
            Next i
            y = y + 1
        Next r_ln
        Me.append add
    End If
End Sub

Public Function priv_mkindex(cols As Variant) As Collection
    Dim v As Variant, s As String, y As Long, i As Long, row As Variant
    Set priv_mkindex = New Collection
    ReDim row(LBound(cols) To UBound(cols))
    For y = 1 To Me.row_count
        For i = LBound(cols) To UBound(cols): row(i) = Replace(Me.item(CLng(cols(i)), y), "_", "__"): Next i
        s = Join(row, "_")
        Err.Clear
        On Error Resume Next
        priv_mkindex.add y, s
        If (Err.Number <> 0) Then
            On Error GoTo 0
            v = priv_mkindex.item(s)
            If IsArray(v) Then ' this is slow but doesn't happen much
                priv_mkindex.Remove s
                ReDim Preserve v(LBound(v) To UBound(v) + 1)
                v(UBound(v)) = y
                priv_mkindex.add v, s
            Else
                priv_mkindex.Remove s
                v = Array(v, y)
                priv_mkindex.add (v), s
            End If
        End If
        On Error GoTo 0
    Next y
End Function

Public Property Get select_distinct(columns As Variant) As Collection
    Dim y As Long, row As Variant, i As Long
    ReDim row(LBound(columns) To UBound(columns))
    Set select_distinct = New Collection
    On Error Resume Next
    For y = 1 To Me.row_count
        For i = LBound(columns) To UBound(columns): row(i) = Replace(Me.item(CLng(columns(i)), y), "_", "__"): Next i
        select_distinct.add y, Join(row, "_")
    Next y
End Property

Public Property Get select_eq(idxs As Variant, values As Variant) As Collection
    Dim y As Long, i As Long, values2 As Variant
    Set select_eq = New Collection
    ReDim values2(LBound(values) To UBound(values))

    For i = LBound(values) To UBound(values)
        If IsArray(values(i)) Then
            Set values2(i) = New Collection
            For y = LBound(values(i)) To UBound(values(i))
                values2(i).add True, values(i)(y)
            Next y
        Else
            values2(i) = values(i)
        End If
    Next i
    For y = 1 To Me.row_count
        For i = LBound(values2) To UBound(values2)
            If IsArray(values2(i)) Then
                On Error Resume Next
                If Not (values2(i).item(Me.item(CLng(idxs(i)), y))) Then
                    On Error GoTo 0
                    Exit For
                Else
                    On Error GoTo 0
                End If
            Else
                If values2(i) <> Me.item(CLng(idxs(i)), y) Then Exit For
            End If
        Next i
        If (i > UBound(values2)) Then select_eq.add y
    Next y
End Property

Public Sub sort(idxs As Variant)
    Dim i As Long, d As Date
    d = Time
    Debug.Print "sorting " & Me.row_count & " items..."
    qsort_ idxs, 1, Me.row_count, 0
    Debug.Print "sorting took " & DateDiff("s", d, Time) & " secs"
End Sub

Private Sub qsort_(ByRef colidxs As Variant, ystart As Long, yend As Long, depth As Long)
    Dim ymid As Long
    If ystart >= yend Then Exit Sub
    ymid = (yend + ystart) / 2
    ymid = partition_(colidxs, ystart, yend, ymid)
    qsort_ colidxs, ystart, ymid - 1, depth + 1
    qsort_ colidxs, ymid + 1, yend, depth + 1
End Sub

Private Function cmplt_(ByRef colidxs As Variant, y0 As Long, y1 As Long) As Boolean
    Dim i As Long
    For i = LBound(colidxs) To UBound(colidxs)
        If Me.item(CLng(colidxs(i)), y0) < Me.item(CLng(colidxs(i)), y1) Then
            cmplt_ = True
            Exit Function
        End If
    Next i
    cmplt_ = False
End Function

Private Sub swap_(y0 As Long, y1 As Long)
    Dim i As Long, tmp As String
    For i = 1 To Me.col_count
        tmp = Me.item(i, y0)
        Me.item(i, y0) = Me.item(i, y1)
        Me.item(i, y1) = tmp
    Next i
End Sub

Private Function partition_(ByRef colidxs As Variant, ystart As Long, yend As Long, ypivot As Long) As Long
    Dim cur As Long, i As Long
    swap_ ypivot, yend
    cur = ystart
    For i = ystart To (yend - 1)
        If cmplt_(colidxs, i, yend) Then
            swap_ cur, i
            cur = cur + 1
        End If
    Next i
    swap_ cur, yend
    partition_ = cur
End Function

Public Sub import_range(r As Range, has_headers As Boolean)
    Dim rr As Range, y_ As Long, x_ As Long
    Set rr = r.Areas(1)
    If (has_headers) Then
        Me.alloc rr.columns.Count, rr.rows.Count - 1
        For x_ = 1 To Me.col_count: Me.header(x_) = rr.Cells(1, x_).value: Next x_
        Set rr = rr.offset(1).resize(rr.rows.Count - 1)
    Else
        Me.alloc rr.columns.Count, rr.rows.Count
    End If
        For y_ = 1 To rr.rows.Count
        For x_ = 1 To rr.columns.Count: Me.item(x_, y_) = rr.Cells(y_, x_).Text: Next x_
    Next y_
End Sub

Public Sub export_range(r As Range, has_headers As Boolean)
    Dim rr As Range, x_ As Long
    Set rr = r.Cells(1, 1)
    If (has_headers) Then
        For x_ = 1 To Me.col_count: rr.offset(0, x_ - 1).value = header(x_): Next x_
        Set rr = rr.offset(1, 0)
    End If
    If (priv_ysz > 0) Then
        For x_ = 1 To Me.col_count
            If (priv_ysz > 65535) Then
                Dim yy As Long
                For yy = 1 To priv_ysz
                    rr.offset(yy - 1, x_ - 1).value = x(x_).y(yy)
                Next yy
            Else
                Range(rr.offset(0, x_ - 1), rr.offset(priv_ysz() - 1, x_ - 1)).value = Application.Transpose(x(x_).y)
            End If
        Next x_
    End If
    If Not (y_extents Is Nothing) Then
        y_extents.export_range rr.offset(priv_ysz(), 0), False
    End If
End Sub

Public Property Get priv_yext() As r3_table2
    Set priv_yext = y_extents
End Property

Public Property Get priv_ysz() As Long
    If (Not Not x) Then
        If (Not Not x(1).y) Then
            priv_ysz = UBound(x(1).y)
        Else
            priv_ysz = 0
        End If
    Else
        priv_ysz = 0
    End If
End Property
