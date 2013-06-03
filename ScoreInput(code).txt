
Private Sub CommandButton1_Click()

End Sub

Public Sub UserForm_Initialize()
    Player1.Caption = Range("Names").Cells(1, 1)
    Player2.Caption = Range("Names").Cells(1, 2)
    Player3.Caption = Range("Names").Cells(1, 3)
    Player4.Caption = Range("Names").Cells(1, 4)
    ScoringSystem.win = "nobody"
    ScoringSystem.score = 0
    ScoringSystem.feeder = "nobody"
End Sub

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Public Sub OkBtn_Click()
Dim msgPrompt As String, msgTitle As String, Feeder1 As String
Dim msgButtons As Integer, msgResult As Integer
    If feeder = "Self" Then
        Feeder1 = "(Self Game)"
        feeder = ""
    Else
        Feeder1 = "Feeder is " & feeder
    End If
    
    msgPrompt = "Winner is " & win & "," & vbNewLine & "with " & score & " Tai." & vbNewLine & Feeder1
    msgButtons = vbYesNo + vbQuestion + vbDefaultButton2
    msgTitle = "Is this OK?"

    msgResult = MsgBox(msgPrompt, msgButtons, msgTitle)

    If msgResult = vbYes Then
        CalculateScore
    End If
    ResetBtn_Click
End Sub
Public Sub ResetBtn_Click()
Dim Ctrl As Control
For Each Ctrl In ScoreInput.Controls
    If TypeName(Ctrl) = "ToggleButton" Then
        Ctrl.Enabled = True
        Ctrl.Value = False
    End If
Next Ctrl
    Feeder1.Caption = ""
    Feeder2.Caption = ""
    Feeder3.Caption = ""
    Feeder4.Caption = ""
    ScoreInput.OpenGG.Enabled = True
    ScoreInput.ClosedGG.Enabled = True
    UserForm_Initialize
End Sub

Private Sub Player1_Click()
    Player2.Enabled = False
    Player3.Enabled = False
    Player4.Enabled = False
    
    Player1.Enabled = True
    Player1.Value = False
    ScoringSystem.win = Player1.Caption
    Feeder1.Caption = "Self"
    Feeder2.Caption = Player2.Caption
    Feeder3.Caption = Player3.Caption
    Feeder4.Caption = Player4.Caption
End Sub

Private Sub Player2_Click()
    Player1.Enabled = False
    Player3.Enabled = False
    Player4.Enabled = False
    
    Player2.Enabled = True
    Player2.Value = False
    ScoringSystem.win = Player2.Caption
    Feeder1.Caption = Player1.Caption
    Feeder2.Caption = "Self"
    Feeder3.Caption = Player3.Caption
    Feeder4.Caption = Player4.Caption
End Sub

Private Sub Player3_Click()
    Player1.Enabled = False
    Player2.Enabled = False
    Player4.Enabled = False
    
    Player3.Enabled = True
    Player3.Value = False
    ScoringSystem.win = Player3.Caption
    Feeder1.Caption = Player1.Caption
    Feeder2.Caption = Player2.Caption
    Feeder3.Caption = "Self"
    Feeder4.Caption = Player4.Caption
End Sub

Private Sub Player4_Click()
    Player1.Enabled = False
    Player2.Enabled = False
    Player3.Enabled = False
    
    Player4.Enabled = True
    Player4.Value = False
    ScoringSystem.win = Player4.Caption
    Feeder1.Caption = Player1.Caption
    Feeder2.Caption = Player2.Caption
    Feeder3.Caption = Player3.Caption
    Feeder4.Caption = "Self"
End Sub

Private Sub Tai1_Click()
    OpenGG.Enabled = False
    ClosedGG.Enabled = False
    Tai2.Enabled = False
    Tai3.Enabled = False
    Tai4.Enabled = False
    Tai5.Enabled = False
    
    Tai1.Enabled = True
    Tai1.Value = False
    ScoringSystem.score = 1
End Sub
Private Sub Tai2_Click()
    OpenGG.Enabled = False
    ClosedGG.Enabled = False
    Tai1.Enabled = False
    Tai3.Enabled = False
    Tai4.Enabled = False
    Tai5.Enabled = False
    
    Tai2.Enabled = True
    Tai2.Value = False
    ScoringSystem.score = 2
End Sub
Private Sub Tai3_Click()
    OpenGG.Enabled = False
    ClosedGG.Enabled = False
    Tai1.Enabled = False
    Tai2.Enabled = False
    Tai4.Enabled = False
    Tai5.Enabled = False
    
    Tai3.Enabled = True
    Tai3.Value = False
    ScoringSystem.score = 3
End Sub
Private Sub Tai4_Click()
    OpenGG.Enabled = False
    ClosedGG.Enabled = False
    Tai1.Enabled = False
    Tai2.Enabled = False
    Tai3.Enabled = False
    Tai5.Enabled = False
    
    Tai4.Enabled = True
    Tai4.Value = False
    ScoringSystem.score = 4
End Sub
Private Sub Tai5_Click()
    OpenGG.Enabled = False
    ClosedGG.Enabled = False
    Tai1.Enabled = False
    Tai2.Enabled = False
    Tai3.Enabled = False
    Tai4.Enabled = False
    
    Tai5.Enabled = True
    Tai5.Value = False
    ScoringSystem.score = 5
End Sub

Private Sub Feeder1_Click()
    
    Feeder1.Enabled = True
    Feeder1.Value = False
    feeder = Feeder1.Caption
    OkBtn_Click
End Sub
Private Sub Feeder2_Click()
    
    Feeder2.Enabled = True
    Feeder2.Value = False
    feeder = Feeder2.Caption
    OkBtn_Click
End Sub
Private Sub Feeder3_Click()
    
    Feeder3.Enabled = True
    Feeder3.Value = False
    feeder = Feeder3.Caption
    OkBtn_Click
End Sub
Private Sub Feeder4_Click()
    
    Feeder4.Enabled = True
    Feeder4.Value = False
    feeder = Feeder4.Caption
    OkBtn_Click
End Sub
Private Sub OpenGG_Click()
    feeder = win
    score = 1
    
    msgPrompt = win & " Open Gang/Ga?"
    msgButtons = vbYesNo + vbQuestion + vbDefaultButton2
    msgTitle = win & " Open GG?"
    msgResult = MsgBox(msgPrompt, msgButtons, msgTitle)
    
    If msgResult = vbYes Then
        CalculateScore
    End If
    ResetBtn_Click
End Sub
Private Sub ClosedGG_Click()
    feeder = ""
    score = 1
    
    msgPrompt = win & " Open Gang/Ga?"
    msgButtons = vbYesNo + vbQuestion + vbDefaultButton2
    msgTitle = win & " Closed GG?"
    msgResult = MsgBox(msgPrompt, msgButtons, msgTitle)
    
    If msgResult = vbYes Then
        CalculateScore
    End If
    ResetBtn_Click
End Sub

