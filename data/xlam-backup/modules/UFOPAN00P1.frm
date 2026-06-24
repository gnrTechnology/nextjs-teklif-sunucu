Dim ir As String, ir2 As Byte
Private Sub Label61_Click()
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\CEMEX\Resimler\" & ir
End Sub
Private Sub Label62_Click()
If TBRS01 = "" Or UFOPAN00.ListBoxP5.listIndex < 0 Then Exit Sub
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\PDF\" & UFOPAN00.TBPM01 & "\" & TBRS01 & ".pdf")
If a = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\CEMEX\PDF\" & UFOPAN00.TBPM01 & "\" & TBRS01 & ".pdf"
Else
MsgBox ("C:\Belgelerim\CEMEX\PDF\" & UFOPAN00.TBPM01 & "\" & TBRS01 & ".pdf" & " dosyası mevcut değil! "), vbInformation, "scngnr@hotmail.com": Exit Sub
End If
End Sub
Private Sub TBRS01_Change()
On Error GoTo hata
Dim rd, a
ir = Trim(Left(UFOPAN00.TBPM01, 3))
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TBRS01 & ".jpg")
If a = True Then
UFOPAN00P1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TBRS01 & ".jpg")
ir2 = 2
Else
 If Not ir2 = 1 Then
  a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ir & ".jpg")
  If a = True Then
  UFOPAN00P1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ir & ".jpg")
  ir2 = 1
  Else
  If Not ir2 = 0 Then UFOPAN00P1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
  End If
 End If
End If
hata:
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '
On Error Resume Next
UFOPAN00.Toolbar1.Buttons.Item(6).Image = UFOPAN00.ImageList1.ListImages.Item(6).Index
UFOPAN00.prs = 0
End Sub