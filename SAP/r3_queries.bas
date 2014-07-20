Attribute VB_Name = "r3_queries"

Option Explicit

Private Const max_opts = 5000
Private Const dump_queries = False

Public Enum r3_selection_operator
    r3_eq
    r3_lt
    r3_gt
    r3_lteq
    r3_gteq
    r3_between
    r3_notbetween
    r3_inset
    r3_notinset
End Enum

Public Type r3_selection_spec
    field As String
    value As Variant ' as string, except with between and notbetween: array of 2 values, inset: array of many values
    format As Variant ' will call Format(.value, .format) when building the query
    operator As r3_selection_operator
End Type

Public Type r3_query_spec
    table As String
    selection() As r3_selection_spec
    fields As Variant ' fields to be read from sap table - string array
    output_table_name As String ' will be prepended to headers
End Type

Public Enum r3_equijoin_hints
    r3eq_merge_queries = 1
End Enum

Public Type r3_equijoin_spec
    left_tbl As r3_table2
    right_tbl_in_sap As String
    right_tbl_in_output As String ' use a different table name in output on our side
    right_fields As Variant ' as string array
    left_join_fields As Variant
    right_join_fields As Variant
    'hints As r3_equijoin_hints
    hint_merge_on_field As String ' string ar
    const_condition_fields As Variant
    const_contition_values As Variant
End Type

Private Function fmt_(val As Variant, fm As Variant) As String
    If (fm <> vbNullString) Then
        fmt_ = format(val, fm)
    Else
        fmt_ = val
    End If
End Function


Public Function r3_query(spec As r3_query_spec) As r3_table2
    Dim opts As Variant, y As Long, i As Long
    
    ReDim opts(0 To query_opts_length(spec) - 1)
    y = 0
    For i = LBound(spec.selection) To UBound(spec.selection)
        With spec.selection(i)
            Select Case .operator
                Case r3_eq: opts(y) = .field & " eq '" & fmt_(.value, .format) & "' "
                Case r3_lt: opts(y) = .field & " < '" & fmt_(.value, .format) & "' "
                Case r3_gt: opts(y) = .field & " > '" & fmt_(.value, .format) & "' "
                Case r3_lteq: opts(y) = .field & " <= '" & fmt_(.value, .format) & "' "
                Case r3_gteq: opts(y) = .field & " >= '" & fmt_(.value, .format) & "' "
                Case r3_between:
                    opts(y) = .field & " BETWEEN '" & fmt_(.value(0), .format) & "' AND '" & fmt_(.value(1), .format) & "' "
                Case r3_notbetween:
                    opts(y) = .field & " NOT BETWEEN '" & fmt_(.value(0), .format) & "' AND '" & fmt_(.value(1), .format) & "' "
                Case r3_inset, r3_notinset:
                    Dim sety As Long
                    opts(y) = .field & IIf(.operator = r3_inset, " IN ( ", " NOT IN ( ")
                    For sety = LBound(.value) To UBound(.value) - 1
                        y = y + 1
                        opts(y) = " '" & fmt_(.value(sety), .format) & "', "
                    Next sety
                    y = y + 1
                    opts(y) = " '" & fmt_(.value(UBound(.value)), .format) & "' ) "
                Case Else:
                    Stop
            End Select
            y = y + 1
            If (i < UBound(spec.selection)) Then
                opts(y) = "AND "
                y = y + 1
            End If
        End With
    Next i

    If (dump_queries) Then
        For i = LBound(opts) To UBound(opts)
            ActiveCell.value = "'" & opts(i)
            ActiveCell.offset(1, 0).Activate
        Next i
    End If

    Set r3_query = r3_read_table2(spec.table, spec.fields, opts)
    If (spec.output_table_name <> vbNullString) Then
        For i = 1 To r3_query.col_count
            r3_query.header(i) = spec.output_table_name & "-" & r3_query.header(i)
        Next i
    End If
End Function

' TODO split this function
Public Function r3_equijoin(j As r3_equijoin_spec, Optional only_download As Boolean) As r3_table2
    Dim o As r3_table2
    Dim opts As Variant, q As r3_query_spec, i As Long
    Dim left_idx As Variant, right_idx As Variant, add_idx As Variant, distincts As Collection, merged As Long, y As Long

    q.table = j.right_tbl_in_sap
    q.fields = j.right_fields

    ReDim left_idx(LBound(j.left_join_fields) To UBound(j.left_join_fields))
    ReDim q.selection(LBound(left_idx) To UBound(left_idx))
    For i = LBound(left_idx) To UBound(left_idx)
        left_idx(i) = j.left_tbl.header_x(CStr(j.left_join_fields(i)))
        With q.selection(i)
            .field = j.right_join_fields(i)
            If (.field = j.hint_merge_on_field) Then
                .operator = r3_inset
                merged = i
            Else
                .operator = r3_eq
            End If
        End With
    Next i
    If Not IsEmpty(j.const_condition_fields) Then
        y = UBound(q.selection) + 1
        ReDim Preserve q.selection(LBound(q.selection) To (y + UBound(j.const_condition_fields) - LBound(j.const_condition_fields)))
        For i = LBound(j.const_condition_fields) To UBound(j.const_condition_fields)
            With q.selection(y + i - LBound(j.const_condition_fields))
                .field = j.const_condition_fields(i)
                .operator = IIf(IsArray(j.const_contition_values(i)), r3_inset, r3_eq)
                .value = j.const_contition_values(i)
            End With
        Next i
    End If

    Set o = New r3_table2
    o.alloc (UBound(j.right_fields) - LBound(j.right_fields) + 1), 0
    o.set_headers j.right_fields

    If (UBound(left_idx) = LBound(left_idx)) Then
        y = 1
        merged = LBound(q.selection)
        q.selection(merged).operator = r3_inset
        Do While y < j.left_tbl.row_count
            With q.selection(merged)
                ReDim .value(0 To WorksheetFunction.Min(max_opts - 1, j.left_tbl.row_count - y))
                For i = 0 To UBound(.value)
                    .value(i) = j.left_tbl.item(CLng(left_idx(merged)), y)
                    y = y + 1
                Next i
            End With
            o.append r3_query(q)
        Loop
    ElseIf j.hint_merge_on_field <> vbNullString Then
        Dim ll As Variant, v As Variant, dd As Collection, distidx As Variant
        Set distincts = get_distincts_for_merged_query(j.left_tbl, left_idx, distidx, q)
        For Each ll In distincts
            For i = LBound(left_idx) To UBound(left_idx)
                q.selection(i).value = j.left_tbl.item(CLng(left_idx(i)), CLng(ll))
            Next i
            v = Empty
            ReDim v(LBound(distidx) To UBound(distidx))
            For i = LBound(distidx) To UBound(distidx)
                v(i) = j.left_tbl.item(CLng(distidx(i)), CLng(ll))
            Next i
            Set dd = j.left_tbl.select_eq(distidx, v)
            i = 0: y = 0
            For Each v In dd
                If (i = 0) Then
                    q.selection(merged).value = Empty
                    ReDim q.selection(merged).value(0 To WorksheetFunction.Min(max_opts - 1, dd.Count - 1 - y))
                End If
                q.selection(merged).value(i) = j.left_tbl.item(CLng(left_idx(merged)), CLng(v))
                If (i = UBound(q.selection(merged).value)) Then
                    Debug.Print "r3_equijoin request " & i + 1 & " items"
                    o.append r3_query(q)
                    y = y + i + 1
                    i = 0
                Else
                    i = i + 1
                End If
            Next v
        Next ll
    Else
        For y = 1 To j.left_tbl.row_count
            For i = LBound(left_idx) To UBound(left_idx)
                q.selection(i).value = j.left_tbl.item(CLng(left_idx(i)), y)
            Next i
            o.append r3_query(q)
        Next y
    End If

    ReDim right_idx(LBound(left_idx) To UBound(left_idx))
    For i = LBound(right_idx) To UBound(right_idx)
        right_idx(i) = o.header_x(CStr(j.right_join_fields(i)))
    Next i
    For i = LBound(j.right_fields) To UBound(j.right_fields)
        For y = LBound(j.right_join_fields) To UBound(j.right_join_fields)
            If j.right_fields(i) = j.right_join_fields(y) Then Exit For
        Next y
        If (y > UBound(j.right_join_fields)) Then
            If (IsEmpty(add_idx)) Then
                ReDim add_idx(0 To 0)
            Else
                ReDim Preserve add_idx(0 To UBound(add_idx) + 1)
            End If
            add_idx(UBound(add_idx)) = o.header_x(CStr(j.right_fields(i)))
        End If
    Next i
    If (j.right_tbl_in_output <> vbNullString) Then
        For i = 1 To o.col_count: o.header(i) = j.right_tbl_in_output & "-" & o.header(i): Next i
    End If
    If (Not only_download) Then
        j.left_tbl.equijoin o, left_idx, right_idx, add_idx
    End If
    Set r3_equijoin = o
End Function

Private Function get_distincts_for_merged_query(l As r3_table2, li As Variant, didx As Variant, q As r3_query_spec)
    Dim j As Long, k As Long
    ReDim didx(LBound(li) To UBound(li) - 1)
    j = LBound(li): Do While (q.selection(j).operator <> r3_inset): didx(j) = li(j): j = j + 1: Loop
    j = j + 1: Do While (j <= UBound(li)): didx(j - 1) = li(j): j = j + 1: Loop
    Set get_distincts_for_merged_query = l.select_distinct(didx)
End Function

Private Function selection_opts_length(spec As r3_selection_spec) As Long
    If (spec.operator = r3_inset) Or (spec.operator = r3_notinset) Then
        selection_opts_length = (UBound(spec.value) - LBound(spec.value) + 1) + 1
    Else
        selection_opts_length = 1
    End If
End Function

Private Function query_opts_length(spec As r3_query_spec) As Long
    Dim i As Long, l As Long
    If ((Not spec.selection) = -1) Then Exit Function
    For i = LBound(spec.selection) To UBound(spec.selection)
        l = l + selection_opts_length(spec.selection(i))
    Next i
    l = l + (UBound(spec.selection) - LBound(spec.selection)) ' for "AND" between fields
    query_opts_length = l
End Function

