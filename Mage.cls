VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CastAbility As CanCast
Private pName As String
Private pMana As Long

Private Sub Class_Initialize()
    pMana = 100
    Set CastAbility = New CanCast
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal Value As String)
    pName = Value
End Property

Public Property Get Mana() As String
    Mana = pMana
End Property

Public Property Let Mana(ByVal Value As String)
    pMana = Value
End Property

Public Sub Cast(ByVal Spell As String)
    CastAbility.Cast Me, Spell
End Sub
