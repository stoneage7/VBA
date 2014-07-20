Attribute VB_Name = "r3_base2"
' SAP functions
' Version 2014-02-24 / Peter.Cernak@jci.com
' functions for transferring raw data from SAP, not report specific

Option Explicit

Function r3_fn(Optional kill As Boolean) As Object
    Static f As Object
    If kill Then
        Set f = Nothing
        Set r3_fn = Nothing
        Exit Function
    End If
    If (f Is Nothing) Then
        Set f = CreateObject("SAP.Functions")
    End If
    Set r3_fn = f
End Function

Function r3_connected() As Boolean
    Dim conn As Object
    Set conn = r3_fn().Connection
    r3_connected = conn.IsConnected
End Function

Sub r3_logon(Optional user As String, Optional pass As String, Optional client As String, _
 Optional server As String, Optional lang As String, Optional autologin As Boolean)
    Dim conn As Object
    Dim r As Boolean

    Set conn = r3_fn().Connection
    If (conn.IsConnected) Then
        conn.logoff
        Set conn = Nothing
        Set conn = r3_fn().Connection
    End If
    
    If user <> "" Then conn.user = user
    If pass <> "" Then conn.password = pass
    If client <> "" Then conn.client = client
    If server <> "" Then conn.applicationserver = server
    If lang <> "" Then conn.Language = lang

    Debug.Print "r3_logon: user " & user & ", server " & server & ", client " & client
    r = conn.logon(0, autologin)
    If r Then
        r3_fn().Connection = conn
    Else
        Err.Raise vbObjectError + 1, "r3_logon", "Failed to login to SAP r3"
    End If
End Sub

Sub r3_logon_with_destination(destination As String, user As String, pass As String, _
    client As String, lang As String, autologin As Boolean)
    Dim conn As Object, r As Boolean
    Set conn = r3_fn().Connection
    If (conn.IsConnected) Then
        conn.logoff
        Set conn = Nothing
        Set conn = r3_fn().Connection
    End If
    
    conn.destination = destination
    conn.user = user
    conn.password = pass
    conn.client = client
    conn.Language = lang
    Debug.Print "r3_logon_with_destination: " & destination & ", user " & user & ", client " & client
    r = conn.logon(0, autologin)
    If r Then
        r3_fn().Connection = conn
    Else
        Err.Raise vbObjectError + 1, "r3_logon_with_destination", "Failed to login to SAP r3 '" & _
            destination & "'"
    End If
End Sub

Sub r3_logoff()
    Debug.Print "r3_logoff"
    r3_fn().Connection.logoff
    Call r3_fn(True)
End Sub

Function r3_read_table2(i_tbl As String, ByRef i_flds As Variant, _
    ByRef i_opts As Variant) As r3_table2
    Dim o As r3_table2
    Dim rfc As Object, tbl As Object, deli As Object
    Dim t_opts As Object, t_data As Object, t_flds As Object
    Dim debugs As String

    debugs = "r3_read_table " & i_tbl & ": "
    Set rfc = r3_fn().add("RFC_READ_TABLE")
    Set tbl = rfc.exports("QUERY_TABLE")
    Set deli = rfc.exports("DELIMITER")
    Set t_opts = rfc.tables("OPTIONS")
    Set t_flds = rfc.tables("FIELDS")
    
    tbl.value = i_tbl
    deli.value = ""
    
    Dim x As Long, y As Long
    For x = LBound(i_flds) To UBound(i_flds)
        t_flds.appendrow
        t_flds(t_flds.rowcount, "FIELDNAME") = i_flds(x)
    Next x
    For x = LBound(i_opts) To UBound(i_opts)
        t_opts.appendrow
        t_opts(t_opts.rowcount, "TEXT") = i_opts(x)
        If t_opts.rowcount <= 8 Then debugs = debugs & i_opts(x)
    Next x
    
    Debug.Print debugs
    If rfc.Call = True Then
        Debug.Print "r3_read_table rfc call done"
        Set t_flds = rfc.tables("FIELDS")
        Set t_data = rfc.tables("DATA")
        Set t_opts = rfc.tables("OPTIONS")
        Set o = New r3_table2

        If (t_data.rowcount = 0) Then
            o.alloc t_flds.rowcount, 0
            o.set_headers i_flds
            Set r3_read_table2 = o
            r3_fn().Remove (rfc)
            Exit Function
        End If
        
        Dim offsets() As Long
        Dim wa As String
        o.alloc t_flds.rowcount, t_data.rowcount
        ReDim offsets(1 To t_flds.rowcount)
        For x = 1 To t_flds.rowcount
            offsets(x) = t_flds(x, "OFFSET") + 1
        Next x
        For y = 1 To t_data.rowcount
            wa = t_data(y, "WA")
            For x = 1 To UBound(offsets) - 1
                o.item(x, y) = Mid(wa, offsets(x), offsets(x + 1) - offsets(x))
            Next x
            o.item(x, y) = Mid(wa, offsets(x))
        Next y
        o.set_headers i_flds
        Set r3_read_table2 = o
        r3_fn().Remove (rfc)
        DoEvents
    Else
        Dim emsg As String
        Set r3_read_table2 = Nothing
        emsg = rfc.exception
        r3_fn().Remove (rfc)
        DoEvents
        Err.Raise vbObjectError + 2, "r3_read_table", "RFC_READ_TABLE failed: " & emsg
    End If
End Function

Function r3_convert_date(idt As String) As Date
    r3_convert_date = DateSerial(CInt(Mid(idt, 1, 4)), CInt(Mid(idt, 5, 2)), CInt(Mid(idt, 7, 2)))
End Function


Function r3_convert_currency(raw As String, decimals As Long, shkzg As String) As Double
    raw = Trim(raw)
    If (raw = vbNullString) Then Exit Function
    Select Case decimals:
        Case 2:
            r3_convert_currency = Round( _
                IIf(shkzg = "H", -1, 1) * _
                (CDbl(Left(raw, Len(raw) - 3)) + 0.01 * CDbl(Right(raw, 2))) _
                , 2)
        Case 0:
            r3_convert_currency = Round( _
                IIf(shkzg = "H", -1, 1) * CDbl(Replace(raw, ".", "")) _
                , 2)
        Case Else:
            Err.Raise vbObjectError, "r3_convert_currency"
    End Select
End Function
