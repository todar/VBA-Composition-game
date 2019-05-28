VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FightAbility As CanFight
Private pName As String
Private pStamina As Long

Private Sub Class_Initialize()
    pStamina = 100
    Set FightAbility = New CanFight
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal Value As String)
    pName = Value
End Property

Public Property Get Stamina() As String
    Stamina = pStamina
End Property

Public Property Let Stamina(ByVal Value As String)
    pStamina = Value
End Property

'This is the function that uses the ability to fight.
Public Sub Fight()
    FightAbility.Fight Me
End Sub
