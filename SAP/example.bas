Attribute VB_Name = "example"

Option Explicit


' example for query and join
' load billing documents from VBRK by billing date, add material level and customer name, export

Sub example()
    Dim tbl As r3_table2, q As r3_query_spec, j1 As r3_equijoin_spec, j2 As r3_equijoin_spec
    r3_logon_with_destination "destination from saplogon.ini", "USER", "PASSWORD", "client", "EN", True
    
    With q
        .table = "VBRK"
        .fields = Split("VBELN,FKART,VKORG,FKDAT,KUNAG", ",")
        ReDim .selection(1 To 1)
    End With
    With q.selection(1)
        .field = "FKDAT"
        .operator = r3_between
        .value = Array(#7/14/2014#, #7/20/2014#)
        .format = "YYYYMMDD"
    End With
    Set tbl = r3_query(q)

    With j1
        Set .left_tbl = tbl
        .left_join_fields = Array("VBELN")
        .right_tbl_in_sap = "VBRP"
        .right_join_fields = Array("VBELN")
        .right_fields = Split("VBELN,POSNR,MATNR,FKIMG,VRKME,NETWR", ",")
    End With
    r3_equijoin j1
    
    With j2
        Set .left_tbl = tbl
        .left_join_fields = Array("KUNAG")
        .right_tbl_in_sap = "KNA1"
        .right_join_fields = Array("KUNNR")
        .right_fields = Split("KUNNR,NAME1", ",")
    End With
    r3_equijoin j2
    
    tbl.export_range Workbooks.add().Sheets(1).Cells(1, 1), True
End Sub
