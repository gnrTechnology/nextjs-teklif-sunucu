Dim rs
Dim ns As Integer
Private Sub UserForm_Activate()
'UF2P.Picture = LoadPicture()
rs = Trim(Left(Replace(mlz, "+", ""), 3))
UF2P.Caption = UF2.TextBox8
UF2P.Picture = UF2.Image1.Picture

Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
For i = 1 To 10
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & UF2P.Caption & "_" & i & ".jpg")
If a = True Then TBRS01.Tag = 1 + TBRS01.Tag
Next
UF2P.Tag = UF2.LBR1.Tag
TBRS01 = UF2P.Tag + 1 & " / " & TBRS01.Tag
End Sub
Private Sub UserForm_Click()
UF2P.Tag = UF2P.Tag + 1: Call rsileri
End Sub
Sub rsileri()
On Error Resume Next
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & UF2P.Caption & "_" & UF2P.Tag & ".jpg")
If a = True Then
UF2P.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & UF2P.Caption & "_" & UF2P.Tag & ".jpg")
TBRS01 = UF2P.Tag + 1 & " / " & TBRS01.Tag
Exit Sub
Else
UF2P.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & UF2P.Caption & ".jpg"): UF2P.Tag = 0
TBRS01 = UF2P.Tag + 1 & " / " & TBRS01.Tag
Exit Sub
End If
UF2P.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
End Sub