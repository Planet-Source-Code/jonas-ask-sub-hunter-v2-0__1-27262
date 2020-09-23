Attribute VB_Name = "modBoss"
Public Type aBoss
 X As Integer
 Y As Integer
 Ty As Integer
 Tx As Integer
 Act As Boolean
 Life As Integer
 FlagFire As Byte
 FireTag1 As Byte
End Type

Public Boss As aBoss
Public OkToMakeBoss As Boolean
Public BossBombs(1 To 70) As DropBombs
Public Const BossBredde As Integer = 70
Public Const BossHoyde As Integer = 34
Public Sub CheckBoss()
    If P1.Killed Mod 20 = 0 And OkToMakeBoss And Boss.Act = False Then
        'if the player has killed 20 subs since the last time, and it's falged green and there is not a boss active as it it,
        MakeBoss 'make the boss
        OkToMakeBoss = False 'and flag red
    End If
End Sub

Public Sub MakeBoss()
    Boss.Act = True
    Select Case Int(Rnd * 2)
    Case 0
        Boss.X = 0 - BossBredde
    Case 1
        Boss.X = Bredde
    End Select
    Boss.Y = Int((Rnd * 200) + 130) 'Give a random Y point to enter
    Boss.Ty = Int((Rnd * 200) + 130) 'Give a random Y point to target
    Boss.Tx = P1.X
    Boss.Life = 15
    PlaySound "boss"
End Sub

Public Sub MoveBoss()
    BossMoveShotSmall 'Move small shots
    'This HAS to be done, no matter what
    If Boss.Act = False Then Exit Sub
    
    BossCheckTarget 'Target finding and checking
    
    BossFireSmallThink 'Fire smaller gunns
    
    If Boss.X <> Boss.Tx Then 'moving
        If Boss.X > Boss.Tx Then
            Boss.X = Boss.X - 2
        Else
            Boss.X = Boss.X + 2
        End If
    End If
    If Boss.X + 1 = Boss.Tx Then Boss.X = Boss.X + 1 'Prevent Missing the target
    If Boss.X - 1 = Boss.Tx Then Boss.X = Boss.X - 1
    
    If Boss.Y <> Boss.Ty Then 'moving
        If Boss.Y > Boss.Ty Then
            Boss.Y = Boss.Y - 2
        Else
            Boss.Y = Boss.Y + 2
        End If
    End If
    If Boss.Y + 1 = Boss.Ty Then Boss.Y = Boss.Y + 1 'Prevent Missing the target
    If Boss.Y - 1 = Boss.Ty Then Boss.Y = Boss.Y - 1
End Sub
Sub BossCheckTarget()
    If Boss.Ty = Boss.Y Then
        Boss.Ty = Int((Rnd * 200) + 130)
    End If
    If Boss.Tx = Boss.X Then
        Boss.Tx = P1.X - 20 + Int(Rnd * 40) + 1
    End If
End Sub
Public Sub Hitboss()
    If Boss.Life > 1 Then 'Wounded the boss
        Boss.Life = Boss.Life - 1
    Else 'Killed the boss
        P1.Score = P1.Score + 1000 'add score
        P1.Health = P1.Health + 5 'add some health too
        MakeSign 1000, Boss.X + 20, Boss.Y 'add sign
        
        MakeExplo Boss.X + 30, Boss.Y + 5
        MakeExplo Boss.X + 60, Boss.Y
        Boss.Act = False
        Boss.Tx = 0
        Boss.Ty = 0
        Boss.X = 0
        Boss.Y = 0
    End If
End Sub
Sub BossFireSmall()
    If Boss.FlagFire Mod 4 = 0 Then 'only fire every 2nd time
        'primary gun
        For A = 1 To UBound(BossBombs)
            If BossBombs(A).Act = False Then
                BossBombs(A).Act = True
                BossBombs(A).X = Boss.X + 17
                BossBombs(A).Y = Boss.Y + 1 + Int(Rnd * 3)
                Exit For
            End If
        Next A
        'secondary gun
        For A = 1 To UBound(BossBombs)
            If BossBombs(A).Act = False Then
                BossBombs(A).Act = True
                BossBombs(A).X = Boss.X + 55
                BossBombs(A).Y = Boss.Y + 1 + Int(Rnd * 3)
                Exit For
            End If
        Next A
        PlaySound "BossFire"
    End If
    Boss.FlagFire = Boss.FlagFire - 1
End Sub
Sub BossFireSmallThink()
    If Boss.FlagFire > 0 Then BossFireSmall: Exit Sub 'already fireing
    If Boss.FireTag1 > 0 Then Boss.FireTag1 = Boss.FireTag1 - 1: Exit Sub
    If Not (Boss.X <= P1.X + ShipBredde + 10 And Boss.X > P1.X - 10) Then Exit Sub
    'If Rnd < 0.9 Then Exit Sub 'a small random
    Boss.FlagFire = Int(Rnd * 15) + 10
    Boss.FireTag1 = 25
End Sub

Sub BossMoveShotSmall()
    For A = 1 To UBound(BossBombs)
        If BossBombs(A).Act Then
            BossBombs(A).Y = BossBombs(A).Y - 3.2
            
            If BossBombs(A).X <= P1.X + ShipBredde And BossBombs(A).X >= P1.X Then
            If BossBombs(A).Y <= P1.Y + ShipHoyde And BossBombs(A).Y >= P1.Y Then
                MakeExplo BossBombs(A).X, BossBombs(A).Y
                
                P1.Health = P1.Health - 1 'Remove health
                If P1.Health <= 0 Then 'check if died
                    MainPause = True
                    P1.Health = 0
                    MsgBox "Game Over", vbOKOnly, Form1.Caption
                    Form1.PicExit_Click
                End If
                
                BossBombs(A).Act = False 'deactiavte bomb
                BossBombs(A).X = 0
                BossBombs(A).Y = 0
            End If
            End If
            
            If BossBombs(A).Y <= 117 Then 'Reached the water surface
                BossBombs(A).Act = False
                BossBombs(A).X = 0
                BossBombs(A).Y = 0
            End If
        End If
    Next A
End Sub
