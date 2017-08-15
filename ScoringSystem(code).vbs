Public win As String
Public score As Integer
Public feeder As String

Sub Rectangle1_Click()
ScoreInput.Show
End Sub

Function CalculateScore()
If (win = "nobody" Or feeder = "nobody" Or score = 0) Then Exit Function
Dim LstCell As Range
Set LstCell = ActiveSheet.Range("F4")
    If (ActiveSheet.Range("F5") = "") Then Set LstCell = ActiveSheet.Range("F3")
LstCell.End(xlDown).Offset(1, 0).Activate
ActiveCell.Value = win
ActiveCell.Offset(0, 1).Value = score
ActiveCell.Offset(0, 2).Value = feeder
End Function

Function ClearSheets()
ActiveSheet.Range("B3", "E3").Copy
ActiveSheet.Range("B100", "E100").PasteSpecial Paste:=xlPasteValues
ActiveSheet.Range("F5", "H100").ClearContents
ActiveSheet.Range("F4").Activate
End Function

Function AllClear()
ActiveSheet.Range("F5", "H100").ClearContents
ActiveSheet.Range("B5", "E5").Copy (ActiveSheet.Range("B6", "E100"))
End Function
