VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CanCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Cast(ByRef State As Object, ByVal Spell As String)
    Debug.Print State.Name & " casts " & Spell & "!"
    State.Mana = State.Mana - 1
End Sub
