VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CanFight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Fight(ByRef State As Object)
    Debug.Print State.Name & " slashes at the foe!"
    State.Stamina = State.Stamina - 1
End Sub
