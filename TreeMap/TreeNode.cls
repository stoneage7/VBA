VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public parent As TreeNode
Public left As TreeNode
Public right As TreeNode
Public color As Long ' 0 for black, 1 for red
Public key As Variant
Public payload As Variant

