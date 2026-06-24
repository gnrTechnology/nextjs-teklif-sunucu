Private Sub CBox21_Click()
On Error Resume Next
If CBox21.Value = True Then
Worksheets.Item("Veriler").Visible = True
Worksheets.Item("Veriler").Select
Else
Worksheets.Item("Veriler").Visible = False
End If
End Sub
Private Sub Label251_Click()
TBAA5 = Format(Now, "DD MMMM YYYY hh:mm:ss")
End Sub
Private Sub Label252_Click()
TBAA4 = "TD-" & Format(Format(Now, "DD MMMM YYYY hh:mm:ss"), "DDMMYY-hhmmss")
End Sub
Private Sub Label253_Click()
Call DOSYAADI_R1
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Set V1 = Worksheets.Item("Veriler")
For i = 1 To 9
Controls("TBAA" & i) = V1.Cells(i, 2).Value
Next i
TBAAB8 = V1.Cells(8, 3).Value: TBAAC8 = V1.Cells(8, 4).Value
If Worksheets.Item("Veriler").Visible = True Then CBox21.Value = True
End Sub
Private Sub CBOK1_Click()
Set V1 = Worksheets.Item("Veriler")
For i = 1 To 9
V1.Cells(i, 2).Value = Controls("TBAA" & i)
Next i
V1.Cells(8, 3).Value = TBAAB8: V1.Cells(8, 4).Value = TBAAC8
End Sub
Private Sub CBIP1_Click()
Unload Me
End Sub