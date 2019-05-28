Attribute VB_Name = "Game"
Option Explicit

'MOST EXCITING GAME OF ALL TIME!
Private Sub StartGame()
    
    'CAN ONLY FIGHT (ONLY HAS STAMINA)
    Dim Slasher As Fighter
    Set Slasher = New Fighter
    Slasher.Name = "Slasher"
    Slasher.Fight '-> Slasher slashes at the foe!
    Debug.Print Slasher.Stamina '-> 99
    
    'MAGES CAN ONLY CAST (ONLY HAS MANA)
    Dim Scorcher As Mage
    Set Scorcher = New Mage
    Scorcher.Name = "Scorcher"
    Scorcher.Cast "fireball" '->Scorcher casts fireball!
    Debug.Print Scorcher.Mana '-> 99
    
    'CAN BOTH FIGHT & CAST (HAS BOTH STAMINA & MANA)
    Dim Roland As Paladin
    Set Roland = New Paladin
    Roland.Name = "Roland"
    Roland.Fight '-> Roland slashes at the foe!
    Roland.Cast "Holy Light" '-> Roland casts Holy Light!
    
End Sub


