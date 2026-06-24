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
Private Sub TB1_Change()
On Error GoTo hata
If UFOPAN00.ListBoxPT.listIndex >= 0 Then R1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TB1 & ".jpg"): Exit Sub
hata: R1.Picture = LoadPicture()
End Sub
Private Sub TB2_Change()
On Error GoTo hata
If UFOPAN00.ListBoxPTD1.listIndex >= 0 Then R2.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TB2 & ".jpg"): Exit Sub
hata: R2.Picture = LoadPicture()
End Sub
Private Sub TB3_Change()
On Error GoTo hata
If UFOPAN00.ListBoxPTD2.listIndex >= 0 Then R3.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TB3 & ".jpg"): Exit Sub
hata: R3.Picture = LoadPicture()
End Sub
Private Sub TB4_Change()
On Error GoTo hata
If UFOPAN00.ListBoxPTD7.listIndex >= 0 Then R4.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TB4 & ".jpg"): Exit Sub
hata: R4.Picture = LoadPicture()
End Sub
Private Sub TB5_Change()
On Error GoTo hata
If UFOPAN00.ListBoxPTD9.listIndex >= 0 Then R5.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TB5 & ".jpg"): Exit Sub
hata: R5.Picture = LoadPicture()
End Sub
Private Sub TB6_Change()
On Error GoTo hata
If UFOPAN00.ListBoxPTD8.listIndex >= 0 Then R6.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & TB6 & ".jpg"): Exit Sub
hata: R6.Picture = LoadPicture()
End Sub
Private Sub TBRS01_Change()
On Error Resume Next
ir = Left(UFOPAN00.CBBPM1, 3) & "\" & UFOPAN00.KDtip
TPKOD = UFOPAN00.KDtip
TB1 = UFOPAN00.TBtip: TB2 = UFOPAN00.YPOEK1
If UFOPAN00.ListBoxPTD2.listIndex >= 0 Then TB3 = UFOPAN00.ListBoxPTD2.Text
TB4 = UFOPAN00.KDok1
TB5 = UFOPAN00.KDdyp1 & UFOPAN00.KDdyp2: TB6 = UFOPAN00.KDiyp1
End Sub
Private Sub TPKOD_Change()
TB1 = "": TB2 = "": TB3 = "": TB4 = "": TB5 = "": TB6 = ""
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '
On Error Resume Next
UFOPAN00.Toolbar1.Buttons.Item(6).Image = UFOPAN00.ImageList1.ListImages.Item(6).Index
UFOPAN00.prs = 0
End Sub