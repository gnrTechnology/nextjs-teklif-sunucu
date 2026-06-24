Private Sub CommandButton12_Click()
On Error Resume Next
Unload Me
ActiveWindow.SelectedSheets.PrintOut
End Sub
Private Sub CommandButton13_Click()
Unload Me
End Sub
Private Sub CommandButton14_Click()
On Error Resume Next
Unload UserForm2
UFmd.Hide
'UserFormANA.Hide
Application.Visible = True
ActiveWindow.SelectedSheets.PrintPreview
Sheets("Sayfa1").Select
UFmd.Show
'UFmy.Show
'UFmy.MultiPage2.Value = 3
End Sub
Private Sub UserForm_QueryClosexx(Cancel As Integer, CloseMode As Integer) 'TAMAM
On Error Resume Next
'printreset
End Sub