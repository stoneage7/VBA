VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeInorderCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private cursor As TreeNode

Public Function _
next_node() As TreeNode
    If (Not cursor.right Is Nothing) Then
        Set cursor = leftmost(cursor.right)
    Else
        If (cursor.parent Is Nothing) Then
            Set cursor = Nothing
            Exit Function
        End If
        Do While (cursor Is cursor.parent.right)
            Set cursor = cursor.parent
            If (cursor.parent Is Nothing) Then Exit Do
        Loop
        Set cursor = cursor.parent
    End If
    Set next_node = cursor
End Function

Public Function _
prev_node() As TreeNode
    If (Not cursor.left Is Nothing) Then
        Set cursor = rightmost(cursor.left)
    Else
        If (cursor.parent Is Nothing) Then
            Set cursor = Nothing
            Exit Function
        End If
        Do While (cursor Is cursor.parent.left)
            Set cursor = cursor.parent
            If (cursor.parent Is Nothing) Then Exit Do
        Loop
        Set cursor = cursor.parent
    End If
    Set prev_node = cursor
End Function

Public Sub _
start(start_at As TreeNode)
    Set cursor = start_at
End Sub

Public Function _
first() As TreeNode
    Set cursor = leftmost(root(cursor))
    Set first = cursor
End Function

Public Function _
last() As TreeNode
    Set cursor = rightmost(root(cursor))
    Set last = cursor
End Function

Private Function _
root(N As TreeNode) As TreeNode
    Do While (Not N.parent Is Nothing)
        Set N = N.parent
    Loop
    Set root = N
End Function

Private Function _
leftmost(N As TreeNode) As TreeNode
    Do While (Not N.left Is Nothing)
        Set N = N.left
    Loop
    Set leftmost = N
End Function

Private Function _
rightmost(N As TreeNode) As TreeNode
    Do While (Not N.right Is Nothing)
        Set N = N.right
    Loop
    Set rightmost = N
End Function

