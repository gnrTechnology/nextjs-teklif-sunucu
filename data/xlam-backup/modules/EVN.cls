Public WithEvents Resimler As MSForms.Image
Attribute Resimler.VB_VarHelpID = -1
Private Sub Resimler_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Local Error Resume Next
If Resimler.BackColor = &H8000000F Then

 If Left(Resimler.Name, 3) = "EGO" Then
 Call renkeo
 Resimler.BackColor = &HBCCAF3
 UFOPAN11.CB1.Value = True
 UFOPAN11.LBP.Caption = Resimler.Tag
 Else
 Resimler.BackColor = &HCAE3BF
 If Left(Resimler.Name, 3) = "EGO" Then UFOPAN11.CB1.Value = False
 End If
Else
 Resimler.BackColor = &H8000000F
 If Left(Resimler.Name, 3) = "EGO" Then UFOPAN11.CB1.Value = False
End If
End Sub