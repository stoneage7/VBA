VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum TMDupesMode
    TMIgnoreDupes
    TMErrorDupes
    TMOverwriteDupes
End Enum

Private root As TreeNode
Private dupemode As TMDupesMode
Private lcount As Long

Public Function _
add(key As Variant, value As Variant) As TreeNode
    Dim N As TreeNode
    'Debug.Print "add key " & key
    Set N = rb_insert(key)
    If Not (N Is Nothing) Then
        lcount = count + 1
        If (IsObject(value)) Then
            Set N.payload = value
        Else
            N.payload = value
        End If
        Set add = N
    End If
End Function

Public Function _
find(key As Variant) As TreeNode
    Dim at As TreeNode
    Set at = root
    Do While Not (at Is Nothing)
        Select Case cmp_key(key, at.key)
            Case 0:
                Set find = at
                Exit Function
            Case -1:
                Set at = at.left
            Case 1:
                Set at = at.right
        End Select
    Loop
End Function

Public Property Get _
count() As Long
    count = lcount
End Property

Public Property Get _
duplicates() As TMDupesMode
    duplicates = dupemode
End Property

Public Property Let _
duplicates(value As TMDupesMode)
    dupemode = value
End Property

Public Sub _
remove(key As Variant)
    Dim N As TreeNode
    Set N = find(key)
    If (N Is Nothing) Then Err.Raise 9, "TreeMap.Remove", "key not found"
    remove_node N
    lcount = lcount - 1
End Sub

Public Function _
inorder(Optional from_key As Variant) As TreeInorderCursor
    Set inorder = New TreeInorderCursor
    If (Not IsError(from_key)) Then
        inorder.start find(from_key)
    Else
        inorder.start root
    End If
End Function

Public Sub _
dump(Optional N As TreeNode)
    Open ThisWorkbook.Path & "\tree.txt" For Output As #1
    do_dump N, 0
    Close #1
End Sub

Private Sub _
do_dump(Optional N As TreeNode, Optional depth As Long = 0)
    Dim spacer As String, M As TreeNode, last As Long, keystr As String
    Dim blacks As Long
    If (N Is Nothing) Then Set N = root
    If Not (N.left Is Nothing) Then
        do_dump N.left, depth + 1
    End If
    If (N.parent Is Nothing) Then
        spacer = "ROOT+ "
    ElseIf (N Is N.parent.left) Then
        last = 0
        spacer = "/---+ "
    Else
        last = 1
        spacer = "\---+ "
    End If
    Set M = N.parent
    blacks = IIf(N.color = 0, 3, 2)
    If Not (M Is Nothing) Then
        Do While Not M.parent Is Nothing:
            If (last = 0) <> (M Is M.parent.left) Then
                spacer = "|   " & spacer
            Else
                spacer = "    " & spacer
            End If
            last = IIf(M Is M.parent.left, 0, 1)
            If (M.color = 0) Then blacks = blacks + 1
            Set M = M.parent
        Loop
        If (N.left Is Nothing) And (N.right Is Nothing) Then
            spacer = Format(blacks, "0000") & spacer
        Else
            spacer = "----" & spacer
        End If
    End If
    If IsArray(N.key) Then
        keystr = "(" & N.key(LBound(N.key))
        For last = LBound(N.key) + 1 To UBound(N.key)
            keystr = keystr & "," & N.key(last)
        Next last
        keystr = keystr & ")"
    ElseIf IsObject(N.key) Then
        keystr = ObjPtr(N.key)
    Else
        keystr = CStr(N.key)
    End If
    Write #1, spacer & IIf(N.color = 0, "blk:", "red:") & keystr
    If Not (N.right Is Nothing) Then
        do_dump N.right, depth + 1
    End If
End Sub


' left rotation, ie.:
' (N N.left (M M.left M.right)) -> (M (N N.left M.left) M.right)
Private Sub left_rotate(N As TreeNode)
    Dim M As TreeNode
    Set M = N.right
    Set N.right = M.left
    If Not (M.left Is Nothing) Then
        Set M.left.parent = N
    End If
    Set M.parent = N.parent
    If (N.parent Is Nothing) Then
        Set root = M
    Else
        If (N Is N.parent.left) Then
            Set N.parent.left = M
        Else
            Set N.parent.right = M
        End If
    End If
    Set M.left = N
    Set N.parent = M
End Sub

' right rotation, ie.:
' (N (M M.left M.right) N.right) -> (M M.left (N M.right N.right))
Private Sub right_rotate(N As TreeNode)
    Dim M As TreeNode
    Set M = N.left
    Set N.left = M.right
    If Not (M.right Is Nothing) Then
        Set M.right.parent = N
    End If
    Set M.parent = N.parent
    If (N.parent Is Nothing) Then
        Set root = M
    Else
        If (N Is N.parent.left) Then
            Set N.parent.left = M
        Else
            Set N.parent.right = M
        End If
    End If
    Set M.right = N
    Set N.parent = M
End Sub

Private Function _
cmp_key(key1 As Variant, key2 As Variant) As Long
    Dim i As Long, offset As Long
    If (IsArray(key1)) Then
        If (IsArray(key2)) Then
            offset = LBound(key2) - LBound(key1)
            For i = LBound(key1) To WorksheetFunction.Min(UBound(key1), UBound(key2) + offset)
                If (key1(i) < key2(i + offset)) Then
                    cmp_key = -1
                    Exit Function
                ElseIf (key1(i) > key2(i + offset)) Then
                    cmp_key = 1
                    Exit Function
                End If
            Next i
            If (UBound(key1) < UBound(key2) + offset) Then
                cmp_key = -1
            ElseIf (UBound(key1) > UBound(key2) + offset) Then
                cmp_key = 1
            Else
                cmp_key = 0
            End If
        Else
            cmp_key = -1
        End If
    Else
        If (IsArray(key2)) Then
            cmp_key = 1
        Else
            If (key1 < key2) Then
                cmp_key = -1
            ElseIf (key1 = key2) Then
                cmp_key = 0
            Else
                cmp_key = 1
            End If
        End If
    End If
End Function

Private Function _
tree_insert(key As Variant) As TreeNode
    Dim at As TreeNode, last As TreeNode
    Set at = root: Set last = Nothing
    Do While Not (at Is Nothing)
        Set last = at
        Select Case cmp_key(key, at.key)
            Case -1: Set at = at.left
            Case 1: Set at = at.right
            Case 0:
                Select Case (dupemode)
                    Case TMErrorDupes:
                        Err.Raise 9, "TreeMap.tree_insert()", "Key already exists"
                    Case TMIgnoreDupes:
                        Exit Function
                    Case TMOverwriteDupes:
                        Set tree_insert = at
                        Exit Function
                End Select
        End Select
    Loop
    If (last Is Nothing) Then
        Set root = New TreeNode
        root.key = key
        Set tree_insert = root
        Exit Function
    Else
        Set at = New TreeNode
        Set at.parent = last
        at.key = key
        Set tree_insert = at
        Select Case cmp_key(key, last.key)
            Case -1: Set last.left = at
            Case 1: Set last.right = at
            Case Else: Stop
        End Select
    End If
End Function

Private Function _
grandparent(N As TreeNode) As TreeNode
    Set grandparent = Nothing
    If Not (N.parent Is Nothing) Then
        Set grandparent = N.parent.parent
    End If
End Function

Private Function _
uncle(N As TreeNode) As TreeNode
    Dim gp As TreeNode
    Set gp = grandparent(N)
    Set uncle = Nothing
    If Not (gp Is Nothing) Then
        If (N.parent Is gp.left) Then
            Set uncle = gp.right
        Else
            Set uncle = gp.left
        End If
    End If
End Function

Private Function _
sibling(N As TreeNode) As TreeNode
    If (N Is N.parent.left) Then
        Set sibling = N.parent.right
    Else
        Set sibling = N.parent.left
    End If
End Function

Private Function _
rb_insert(key As Variant) As TreeNode
    Dim N As TreeNode, U As TreeNode, G As TreeNode

    Set N = tree_insert(key)
    If (N Is Nothing) Then Exit Function
    Set rb_insert = N
    
    ' node already existed and payload was overwritten by tree_insert()
    If (Not N.left Is Nothing) Or (Not N.right Is Nothing) Then Exit Function

    N.color = 1
    Do While True
        If (N Is root) Then Exit Do
        If (N.parent.color <> 1) Then Exit Do
        Set U = uncle(N)
        Set G = grandparent(N)
        Select Case True:
            Case U Is Nothing, U.color = 0:
                ' U in black
                If (N Is N.parent.right) And (N.parent Is G.left) Then
                    left_rotate N.parent
                    Set N = N.left
                ElseIf (N Is N.parent.left) And (N.parent Is G.right) Then
                    right_rotate N.parent
                    Set N = N.right
                End If
                N.parent.color = 0
                G.color = 1
                If (N Is N.parent.left) Then
                    right_rotate G
                Else
                    left_rotate G
                End If
            Case Else:
                ' U is red
                N.parent.color = 0
                U.color = 0
                G.color = 1
                Set N = G
        End Select
    Loop
    root.color = 0
End Function

Private Sub _
replace_node(dest As TreeNode, src As TreeNode)
    If (src Is Nothing) Then
        dest.key = Empty
        dest.payload = Empty
    Else
        dest.key = src.key
        dest.payload = Empty
        If (IsObject(src.payload)) Then
            Set dest.payload = src.payload
        Else
            dest.payload = src.payload
        End If
    End If
End Sub

Private Sub _
unlink_node(N As TreeNode)
    If (N Is Nothing) Then Exit Sub
    If (N.parent Is Nothing) Then
        Set root = Nothing
        Exit Sub
    End If
    If (N Is N.parent.left) Then
        Set N.parent.left = Nothing
    Else
        Set N.parent.right = Nothing
    End If
    Set N.parent = Nothing
End Sub

Private Function _
get_color(N As TreeNode) As Long
    If (N Is Nothing) Then
        get_color = 0
    Else
        get_color = N.color
    End If
End Function

Private Sub _
set_color(N As TreeNode, color As Long)
    If (N Is Nothing) Then
        If (color = 0) Then Exit Sub Else Stop 'trying to repaint a leaf red
    End If
    N.color = color
End Sub

Private Sub _
remove_node(N As TreeNode)
    Dim C As New TreeInorderCursor, child As TreeNode, S As TreeNode
    
    If ((Not N.left Is Nothing) And (Not N.right Is Nothing)) Then
        Set C = New TreeInorderCursor
        C.start N
        Set child = C.prev_node()
        replace_node N, child
        remove_node child
    Else
        'Me.dump
        If (N.left Is Nothing) Then Set child = N.right Else Set child = N.left
        replace_node N, child
        unlink_node child
        If (get_color(N) = 0) Then
            If (get_color(child) = 1) Then
                set_color child, 0
                If (IsEmpty(N.key)) Then unlink_node N
                Exit Sub
            Else
                ' by wikipedia :P http://en.wikipedia.org/wiki/Red%E2%80%93black_tree
                ' case #1: N is the new root -> exit sub
                Do While (Not N.parent Is Nothing)
                    ' case #2: S is red (-> N.parent, S.left, S.right are black)
                    Set S = sibling(N)
                    If (S.color = 1) Then
                        N.parent.color = 1
                        S.color = 0
                        If (N Is N.parent.left) Then
                            left_rotate N.parent
                        Else
                            right_rotate N.parent
                        End If
                    End If
                    ' case #3: S, Sl, Sr, P are black
                    Set S = sibling(N)
                    If ((N.parent.color = 0) And _
                        (S.color = 0) And _
                        (get_color(S.left) = 0) And _
                        (get_color(S.right) = 0)) Then
                        S.color = 1
                        Set child = N
                        Set N = N.parent
                        If (IsEmpty(child.key)) Then unlink_node child
                        'Stop ' will loop to case #1
                    Else
                        ' case #4
                        'Set S = sibling(N)
                        If ((N.parent.color = 1) And _
                            (S.color = 0) And _
                            (get_color(S.left) = 0) And _
                            (get_color(S.right) = 0)) Then
                            S.color = 1
                            N.parent.color = 0
                            If (IsEmpty(N.key)) Then unlink_node N
                            Exit Sub
                        End If
                        ' case #5
                        If (S.color = 0) Then
                            If ((N Is N.parent.left) And _
                                (get_color(S.right) = 0) And _
                                (get_color(S.left) = 1)) Then
                                S.color = 1
                                set_color S.left, 0
                                right_rotate S
                            ElseIf ((N Is N.parent.right) And _
                                (get_color(S.left) = 0) And _
                                (get_color(S.right) = 1)) Then
                                S.color = 1
                                set_color S.right, 0
                                left_rotate S
                            End If
                        End If
                        ' case #6
                        Set S = sibling(N)
                        S.color = N.parent.color
                        N.parent.color = 0
                        If (N Is N.parent.left) Then
                            set_color S.right, 0
                            left_rotate N.parent
                        Else
                            set_color S.left, 0
                            right_rotate N.parent
                        End If
                        If (IsEmpty(N.key)) Then unlink_node N
                        Exit Sub
                    End If
                Loop
            End If
            'If (IsEmpty(N.key)) Then unlink_node N
        Else
            'If (IsEmpty(N.key)) Then unlink_node N
        End If
        If (IsEmpty(N.key)) Then unlink_node N
    End If
End Sub
