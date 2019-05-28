# VBA-Composition-game
Trying to learn how to do composition in VBA. This is a testing example of what I have come up with so far.

[Here](https://stackoverflow.com/q/56347881/8309643) is a link to my question on stackoverflow as to how to do Object Composition in VBA. 

Here is what the game looks like so far:

```vba
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
```

Goal of adding this is to seek answers on the best way of approching this. So feel free to send a pull request, or comment on stackoverflow.
