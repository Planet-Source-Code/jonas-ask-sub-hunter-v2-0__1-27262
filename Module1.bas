Attribute VB_Name = "Publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020



Public Type Coor
 X As Integer
 Y As Integer
 Act As Boolean
 Tag As Integer
End Type

Public Type Player
 X As Integer
 Y As Integer
 Ammo As Integer
 Dire As Byte
 Health As Integer
 Score As Long
 Speed As Currency
 Firetime As Currency
 Killed As Integer
End Type

Public Type SubMarine
 X As Integer
 Y As Integer
 Act As Boolean
 Score As Integer
 Dire As Integer
 Speed As Integer
 Damaged As Integer
End Type

Public Type DropBombs
 X As Integer
 Y As Integer
 Act As Boolean
 Speed As Currency
End Type

Public Type Bomber
 X As Integer
 Y As Integer
 Act As Boolean
 Dire As Integer
 Speed As Integer
 BombLoad As Integer
 Droped As Boolean
End Type

Public Type HighScore
 PlName As String
 plScore As Long
 plDate As String
End Type

Public Type PointsSign
 Score As Integer
 X As Integer
 Y As Integer
 Tag As Byte
End Type

Public Const Bredde As Integer = 400
Public Const Hoyde As Integer = 360
Public Const ShipBredde = 53
Public Const ShipHoyde = 14
Public Const SubBredde = 36
Public Const SubHoyde = 10
Public Const PlaneBredde = 29
Public Const PlaneHoyde = 12
Public Const MaxAmmo = 7
Public P1 As Player
Public Shot(1 To 30) As Coor
Public Subs(1 To 30) As SubMarine
Public HighS(1 To 10) As HighScore
Public Planes(1 To 10) As Bomber
Public Bombs(1 To 30) As DropBombs
Public Explo(1 To 10) As Coor
Public Signs(1 To 10) As PointsSign
Public TheKing As Coor
Public NumPlanes As Integer
Public NumSubs As Integer
Public NumShots As Integer
Public NumBombs As Integer
Public DontClose As Boolean
Public MainPause As Boolean


Public Function PlaySound(File As String)
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Svar = sndPlaySound(App.Path & "\" & File & ".wav", wFlags%) 'Send the sound to the big world
End Function

Public Sub Fire() 'This is what happens when the player fires
Dim A As Integer
    
    If (GetTickCount - P1.Firetime) < 300 Then Exit Sub 'A delay of 300 millisec between each shot
    If P1.Ammo = 0 Then Exit Sub 'out of ammo, don't fire
    If NumShots = 30 Then Exit Sub 'to many bombs active, don't fire
    
    P1.Firetime = GetTickCount 'set time till next fire
    P1.Ammo = P1.Ammo - 1 'remove one in ammo
    NumShots = NumShots + 1 'add a shot to the counting variable
    
    
    A = 1
    Do Until Not Shot(A).Act 'find an open shot slot
        A = A + 1
    Loop
    
    With Shot(A)
        .Act = True 'atctivate it
        .Y = P1.Y + ShipHoyde 'put it at the bottom of the ship
        .X = P1.X + (ShipBredde / 2) 'an in the middle of it
    End With
End Sub

Public Sub MakeSub() 'Triggerd every tick, it makes subs
Dim A As Integer
    If NumSubs = 30 Then Exit Sub 'Too many subs are active
    If Boss.Act Then Exit Sub 'Don't create subs when at the boss
    
    Randomize
    temp = (Rnd * 130)
    If temp < 2 + 30 - NumSubs Then 'if a small chance occures it creats a sub
    
        NumSubs = NumSubs + 1 'adds a sub to the subcounter
        
        A = 1
        Do Until Not Subs(A).Act Or A = 30 'finds a free slot
            A = A + 1
        Loop
        
        With Subs(A)
        .Act = True 'activates the sub
        
        If Int((Rnd * 2) + 1) = 1 Then 'put's it eighter at the left or right side of the screen
            .X = 0 - SubBredde - 2
            .Dire = 2
        Else
            .X = Bredde + 2
            .Dire = 1
        End If
        
        .Y = Int((Rnd * 200) + 130) 'Give a random Y point to enter
        Randomize
        temp = Int((Rnd * 100) + 1) 'Give a random speed
        Select Case temp
        Case 80 To 100
            .Speed = 3
        Case 50 To 80
            .Speed = 2
        Case Else
            .Speed = 1
        End Select
        
        
        .Score = (.Speed * 2) * (.Y / 8) 'calculate score for killing sub using speed and depth
        
        End With
    End If
End Sub
Public Sub Movesubs() 'moves all the subs
    For A = 1 To 30
    With Subs(A)
        If .Act Then
        'M is the speed at which the subs move
        If .Dire = 2 Then M = .Speed 'going right
        If .Dire = 1 Then M = -1 * .Speed 'going left
        
        If .Damaged <> 0 Then 'if it's hit and sinking:
            .Damaged = .Damaged + 1
            .Y = .Y + 3
            If .Damaged = 10 Then 'sunk deep enough, deactivate
                .Damaged = 0
                .X = 0
                .Y = 0
                .Dire = 0
                .Act = False
            End If
        Else 'if not, move it
            .X = .X + M
        End If
        'Reach end of screen?
        If .X < 0 - SubBredde - 2 Or .X > Bredde + 2 Then
            .Act = False
            .Dire = 0
            .Score = 0
            .Speed = 0
            .Damaged = 0
            .X = 0
            .Y = 0
            NumSubs = NumSubs - 1
        End If
                    
        End If
    End With
    Next A
End Sub
Public Sub MovePlanes()
    For A = 1 To 10
    With Planes(A)
        If .Act Then
        
        If .Dire = 2 Then M = .Speed
        If .Dire = 1 Then M = -1 * .Speed
        
        .X = .X + M
        
        'Reach the edge
        If .X < 0 - PlaneBredde - 2 Or .X > Bredde + 2 Then
            .Act = False
            .Dire = 0
            .Speed = 0
            .Droped = False
            .X = 0
            .Y = 0
            .BombLoad = 0
            NumPlanes = NumPlanes - 1
        End If
                    
        End If
    End With
    Next A
End Sub
Public Sub MoveShots()
    For A = 1 To 30
        With Shot(A)
        If .Act Then
            .Y = .Y + 1.8
            
            'Hit a sub?
            HitSubCheck (A)
            
            If .Y >= Hoyde Then 'Reach the bottom?
                .Act = False
                .X = 0
                .Y = 0
                NumShots = NumShots - 1
            End If
        End If
        End With
        
        
        With Bombs(A) 'Bombs dropped from planes
        If .Act Then
            .Speed = .Speed - (.Speed * 0.1)
            .X = .X + .Speed
            .Y = .Y + 2
            
            HitShipCheck (A)
            
            If .Y >= 117 Then 'Reach surface, deactivate
                .Act = False
                .X = 0
                .Y = 0
                .Speed = 0
                NumBombs = NumBombs - 1
            End If
            
        End If
        End With
    Next A
    'The King
    If TheKing.Act Then
        TheKing.X = TheKing.X - 1
        If TheKing.X <= -24 Then 'Deactivate
            PlaySound "elvis"
            TheKing.Act = False
            TheKing.X = 0
            TheKing.Y = 0
            TheKing.Tag = 0
        End If
        If TheKing.Tag = 0 Then
            TheKing.Tag = 2
        Else: TheKing.Tag = TheKing.Tag - 1
        End If
        
        If Rnd > 0.98 Then PlaySound "elvis2"
    End If
End Sub

Public Sub HitSubCheck(M)
Dim Svar(1 To 4)
    
    Svar(1) = GetPixel(Form1.Pic2.hdc, Shot(M).X, Shot(M).Y)
    Svar(2) = GetPixel(Form1.Pic2.hdc, Shot(M).X + 6, Shot(M).Y)
    Svar(3) = GetPixel(Form1.Pic2.hdc, Shot(M).X + 6, Shot(M).Y + 6)
    Svar(4) = GetPixel(Form1.Pic2.hdc, Shot(M).X, Shot(M).Y + 6)
    For A = 1 To 4
        If Svar(A) <> vbWhite And Not Svar(A) = -1 Then
            'Check the boss
            If Boss.Act Then
                Select Case A
                Case 1
                    If Boss.X <= Shot(M).X And Boss.X + BossBredde >= Shot(M).X Then
                        If Shot(M).Y >= Boss.Y And Shot(M).Y <= Boss.Y + BossHoyde Then
                            Hitboss
                        End If
                    End If
                Case 2
                    If Boss.X <= Shot(M).X + 6 And Boss.X + BossBredde >= Shot(M).X + 6 Then
                        If Shot(M).Y >= Boss.Y And Shot(M).Y <= Boss.Y + BossHoyde Then
                            Hitboss
                        End If
                    End If
                Case 3
                    If Boss.X <= Shot(M).X + 6 And Boss.X + BossBredde >= Shot(M).X + 6 Then
                        If Shot(M).Y + 6 >= Boss.Y And Shot(M).Y + 6 <= Boss.Y + BossHoyde Then
                            Hitboss
                        End If
                    End If
                Case 4
                    If Boss.X <= Shot(M).X And Boss.X + BossBredde >= Shot(M).X Then
                        If Shot(M).Y + 6 >= Boss.Y And Shot(M).Y + 6 <= Boss.Y + BossHoyde Then
                            Hitboss
                        End If
                    End If
                End Select
            End If
            'Now check the subs
            For s = 1 To 30
                Select Case A
                Case 1
                    If Subs(s).X <= Shot(M).X And Subs(s).X + SubBredde >= Shot(M).X Then
                        If Shot(M).Y >= Subs(s).Y And Shot(M).Y <= Subs(s).Y + SubHoyde Then
                            Killsub (s)
                        End If
                    End If
                Case 2
                    If Subs(s).X <= Shot(M).X + 6 And Subs(s).X + SubBredde >= Shot(M).X + 6 Then
                        If Shot(M).Y >= Subs(s).Y And Shot(M).Y <= Subs(s).Y + SubHoyde Then
                            Killsub (s)
                        End If
                    End If
                Case 3
                    If Subs(s).X <= Shot(M).X + 6 And Subs(s).X + SubBredde >= Shot(M).X + 6 Then
                        If Shot(M).Y + 6 >= Subs(s).Y And Shot(M).Y + 6 <= Subs(s).Y + SubHoyde Then
                            Killsub (s)
                        End If
                    End If
                Case 4
                    If Subs(s).X <= Shot(M).X And Subs(s).X + SubBredde >= Shot(M).X Then
                        If Shot(M).Y + 6 >= Subs(s).Y And Shot(M).Y + 6 <= Subs(s).Y + SubHoyde Then
                            Killsub (s)
                        End If
                    End If
                End Select
            Next s
            If Shot(M).Act Then MakeExplo Shot(M).X, Shot(M).Y
            Shot(M).Act = 0
            Shot(M).X = 0
            Shot(M).Y = 0
            NumShots = NumShots - 1
        End If
    Next A
    
End Sub

Public Sub HitShipCheck(M)
Dim Svar
    
    Svar = GetPixel(Form1.Pic2.hdc, Bombs(M).X, Bombs(M).Y)
    
    If Svar <> vbWhite Then
        P1.Health = P1.Health - 1
        With Bombs(M)
        .Act = False
        .Speed = 0
        .X = 0
        .Y = 0
        End With
        PlaySound "hit" 'play the sound
        If P1.Health <= 0 Then
            MainPause = True
            P1.Health = 0
            MsgBox "Game Over", vbOKOnly, Form1.Caption
            Form1.PicExit_Click
        End If
    End If
End Sub
Public Sub Killsub(A)
    With Subs(A)
        If Not .Act Then Exit Sub
        P1.Score = P1.Score + Subs(A).Score 'Add the subs score to the players
        MakeSign Subs(A).Score, Subs(A).X, Subs(A).Y  'Make the Score Sign
        .Score = 0
        .Speed = 0
        .Damaged = 1
        P1.Killed = P1.Killed + 1 'Increase number of killed subs
    End With
    NumSubs = NumSubs - 1
    OkToMakeBoss = True
End Sub
Sub MakeSign(Score, X, Y)
    For A = 1 To UBound(Signs)
        If Signs(A).Tag = 0 Then
            Signs(A).Score = Score
            Signs(A).Tag = 100
            Signs(A).X = X + 10
            Signs(A).Y = Y
            Exit For
        End If
    Next A
End Sub

Public Sub LoadScore()
    Open App.Path & "\data.dat" For Random As #1 Len = 18
    For A = 3 To 30 Step 3
        Get #1, A - 2, HighS(A / 3).PlName
        Get #1, A - 1, HighS(A / 3).plScore
        Get #1, A, HighS(A / 3).plDate
    Next A
    Close #1
End Sub

Public Sub SaveScore()
    On Error Resume Next
    Kill App.Path & "\data.dat"
    Open App.Path & "\data.dat" For Random As #1 Len = 18
    For A = 3 To 30 Step 3
        Put #1, A - 2, HighS(A / 3).PlName
        Put #1, A - 1, HighS(A / 3).plScore
        Put #1, A, HighS(A / 3).plDate
    Next A
    Close #1
End Sub
Public Sub CheckKing()
    If P1.X = 0 And GetAsyncKeyState(vbKeyE) And TheKing.Act = False Then
        'Activate HIM
        PlaySound "elvis2"
        TheKing.Act = True
        TheKing.Tag = 0
        TheKing.X = Bredde + 1
        TheKing.Y = Int((Rnd * 150) + 130)
    End If
End Sub
Public Sub UpdateScore()
Dim MyName As String
Dim Score As Long
    Score = P1.Score
    
    If Score = 0 Then Exit Sub
    
    For A = 1 To 10
        If Score > HighS(A).plScore Then GoTo FantEn
    Next A
    ' No highscore, exit sub
    Exit Sub
FantEn:
    
    'Wanna save?
    Svar = MsgBox("Congratulations! " & P1.Score & " points is a new highscore!" & vbNewLine & "Do you want to write it down?", vbYesNo, "New HighScore: " & A & ". place!")
    If Svar = vbNo Then Exit Sub
    
    'Move previous scores down
    For b = 10 To A + 1 Step -1
        HighS(b).plDate = HighS(b - 1).plDate
        HighS(b).PlName = HighS(b - 1).PlName
        HighS(b).plScore = HighS(b - 1).plScore
    Next b

NewName:
    MyName = InputBox("Please input your name (Max 16 characters)", "New HighScore: " & A & ". place!")
    If Len(MyName) > 16 Then GoTo NewName
    If Len(MyName) = 0 Then GoTo NewName
    
    HighS(A).plDate = Date
    HighS(A).PlName = MyName
    HighS(A).plScore = P1.Score
    frmHigh.Show , Form1
    DontClose = True
End Sub

Public Sub MakePlane()
    If NumPlanes = 10 Then Exit Sub
    
    Randomize
    temp = (Rnd * 130)
    If temp < 20 Then
    
        NumPlanes = NumPlanes + 1
        
        A = 1
        Do Until Not Planes(A).Act Or A = 10
            A = A + 1
        Loop
        With Planes(A)
        
        .Act = True
        
        If Int((Rnd * 2) + 1) = 1 Then
            .X = 0 - PlaneBredde - 2
            .Dire = 2
        Else
            .X = Bredde + 2
            .Dire = 1
        End If
        
        .Y = Int((Rnd * 35) + 5)
        
        .Droped = False
        .Speed = 4
        
        End With
    End If
End Sub

Public Sub DropBombs()
Dim PL As Integer
    For PL = 1 To 10
        If Planes(PL).BombLoad > 0 Then GoTo AllClear
        
        If Planes(PL).Act And Not Planes(PL).Droped Then
            
            'Check if it is smart to drop bombs
            If Planes(PL).X < P1.X + ShipBredde And Planes(PL).X > P1.X Then
                
                If Planes(PL).BombLoad = 0 Then
                    Randomize
                    Planes(PL).BombLoad = Int((Rnd * 7) + 3)
                    Planes(PL).Droped = True
                End If
AllClear:
                
                If NumBombs = 30 Then Exit Sub

                NumBombs = NumBombs + 1
                Planes(PL).BombLoad = Planes(PL).BombLoad - 1
                
                A = 1
                Do Until Not Bombs(A).Act Or A = 30
                    A = A + 1
                Loop
                
                With Bombs(A)
                .Act = True
                
                If Planes(PL).Dire = 1 Then
                    .Speed = Planes(PL).Speed * -1
                Else
                    .Speed = Planes(PL).Speed
                End If
                
                Select Case Planes(PL).Dire
                Case 1: .X = Planes(PL).X + 20
                Case 2: .X = Planes(PL).X + 4
                End Select
                
                .Y = Planes(PL).Y + 14
                
                End With
            End If
        End If
    Next PL
End Sub

Sub MakeExplo(X, Y)
    'play a sound
    PlaySound "explo"
    X = X - 30
    Y = Y - 25
    A = 1
    Do Until Not Explo(A).Act Or A = UBound(Explo)
        A = A + 1
    Loop
    With Explo(A)
        .X = X
        .Y = Y
        .Tag = 0
        .Act = True
    End With
End Sub
Public Sub DoExplo()
    For A = 1 To UBound(Explo)
        If Explo(A).Act Then
            If Explo(A).Tag < 11 Then
                Explo(A).Tag = Explo(A).Tag + 1
            Else
                Explo(A).Act = False
                Explo(A).X = 0
                Explo(A).Y = 0
                Explo(A).Tag = 0
            End If
        End If
    Next A
End Sub
Public Sub DoSigns()
    For A = 1 To UBound(Signs)
        If Signs(A).Tag > 0 Then
            Signs(A).Tag = Signs(A).Tag - 1
            Signs(A).Y = Signs(A).Y - 1
            If (Signs(A).X Mod 5) = 2 Then Signs(A).X = Signs(A).X + 1 'a hack
            Signs(A).X = Signs(A).X - 2 + (Signs(A).X Mod 5) 'wiggle the letters
        Else
            Signs(A).Score = 0
            Signs(A).X = 0
            Signs(A).Y = 0
        End If
    Next A
End Sub
