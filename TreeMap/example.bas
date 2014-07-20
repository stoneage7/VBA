Attribute VB_Name = "example"

Option Explicit

Sub t()
    Dim t As New TreeMap, C As New Collection
    Dim i As Long, start As Date, k As Long
    Randomize Now
    
    k = 20
    
    Debug.Print "TreeMap " & k & " additions"
    start = Now
    For i = 1 To k
        t.add WorksheetFunction.RandBetween(1, 10000000), i
    Next i
    Debug.Print "TreeMap done after " & DateDiff("s", start, Now) & "s"
    t.dump
    
    Debug.Print "TreeMap " & k & " searches"
    start = Now
    For i = 1 To k
        Call t.find(WorksheetFunction.RandBetween(1, 10000000))
    Next i
    Debug.Print "TreeMap done after " & DateDiff("s", start, Now) & "s"
    Set t = Nothing

    Debug.Print "Collection " & k & " additions"
    start = Now
    On Error Resume Next
    For i = 1 To k
        C.add i, CStr(WorksheetFunction.RandBetween(1, 10000000))
    Next i
    Debug.Print "Collection done after " & DateDiff("s", start, Now) & "s"

    Debug.Print "Collection " & k & " searches"
    start = Now
    For i = 1 To k
        Call C.Item(CStr(WorksheetFunction.RandBetween(1, 10000000)))
    Next i
    Debug.Print "Collection done after " & DateDiff("s", start, Now) & "s"


End Sub

Sub tt()
    Dim i As Long, j As Long, k As Long, tmp As Variant
    Dim t As New TreeMap, N As TreeNode, C As TreeInorderCursor
    Dim rm() As Variant
    
    k = 200000
    For i = 1 To k
        't.add Array(WorksheetFunction.RandBetween(1, 100), WorksheetFunction.RandBetween(1, 100), WorksheetFunction.RandBetween(1, 100)), i
        t.add WorksheetFunction.RandBetween(1, 1000000), i
    Next i
    
    't.dump
    Set C = t.inorder(): i = 1
    Set N = C.first()

    ReDim rm(1 To t.count)
    Do Until N Is Nothing
        'Debug.Print i & ": " & N.key
        rm(i) = N.key
        i = i + 1
        Set N = C.next_node()
        DoEvents
    Loop
    
    For i = 1 To UBound(rm) - 10000
        j = WorksheetFunction.RandBetween(1, UBound(rm) - i + 1)
        tmp = rm(UBound(rm) - i + 1)
        rm(UBound(rm) - i + 1) = rm(j)
        rm(j) = tmp
        'Debug.Print "rm: " & rm(UBound(rm) - i + 1)
        t.remove rm(UBound(rm) - i + 1)
    Next i
    Debug.Print t.count
    
    For i = 1 To 100000
        t.add WorksheetFunction.RandBetween(1, 1000000), "newi" & i
    Next i
    'Stop
    k = t.count
    Debug.Print "full count: " & k
    For i = 1 To k - 200
        t.remove t.inorder.first().key
    Next i

    For i = 1 To 1000
        t.add WorksheetFunction.RandBetween(1, 1000000), "newi" & i
    Next i

    t.dump
End Sub


