Dim pno As Integer
Dim z As Integer
Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image1.Tag = 1 Then Image0.Picture = Image1.Picture: z = "1": TBRS01.Text = Image0.ControlTipText & "_1"
Call resimboyut2
End Sub
Private Sub Image2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image2.Tag = 1 Then Image0.Picture = Image2.Picture: z = "2": TBRS01.Text = Image0.ControlTipText & "_2"
Call resimboyut2
End Sub
Private Sub Image3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image3.Tag = 1 Then Image0.Picture = Image3.Picture: z = "3": TBRS01.Text = Image0.ControlTipText & "_3"
Call resimboyut2
End Sub
Private Sub Image4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image4.Tag = 1 Then Image0.Picture = Image4.Picture: z = "4": TBRS01.Text = Image0.ControlTipText & "_4"
Call resimboyut2
End Sub
Private Sub Image10_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image10.Tag = 1 Then Image0.Picture = Image10.Picture: z = "10": LBRX1 = 1
For k = 1 To 4
If Controls("Image1" & k).Tag = 1 Then
Controls("Image" & k).Picture = Controls("Image1" & k).Picture: Controls("Image" & k).Tag = 1
Else
Controls("Image" & k).Picture = LoadPicture(): Controls("Image" & k).Tag = 2
End If
Next
TBRS01.Text = Image10.ControlTipText: Image0.ControlTipText = TBRS01.Text
Call resimboyut2
End Sub
Private Sub Image20_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image20.Tag = 1 Then Image0.Picture = Image20.Picture: z = "20": LBRX1 = 2
For k = 1 To 4
If Controls("Image" & LBRX1 & k).Tag = 1 Then
Controls("Image" & k).Picture = Controls("Image" & LBRX1 & k).Picture: Controls("Image" & k).Tag = 1
Else
Controls("Image" & k).Picture = LoadPicture(): Controls("Image" & k).Tag = 2
End If
Next
TBRS01.Text = Image20.ControlTipText: Image0.ControlTipText = TBRS01.Text
Call resimboyut2
End Sub
Private Sub Image30_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image30.Tag = 1 Then Image0.Picture = Image30.Picture: z = "30": LBRX1 = 3
For k = 1 To 4
If Controls("Image3" & k).Tag = 1 Then
Controls("Image" & k).Picture = Controls("Image3" & k).Picture: Controls("Image" & k).Tag = 1
Else
Controls("Image" & k).Picture = LoadPicture(): Controls("Image" & k).Tag = 2
End If
Next
TBRS01.Text = Image30.ControlTipText: Image0.ControlTipText = TBRS01.Text
Call resimboyut2
End Sub
Private Sub Image40_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Image40.Tag = 1 Then Image0.Picture = Image40.Picture: z = "40": LBRX1 = 4
For k = 1 To 4
If Controls("Image" & LBRX1 & k).Tag = 1 Then
Controls("Image" & k).Picture = Controls("Image" & LBRX1 & k).Picture: Controls("Image" & k).Tag = 1
Else
Controls("Image" & k).Picture = LoadPicture(): Controls("Image" & k).Tag = 2
End If
Next
TBRS01.Text = Image40.ControlTipText: Image0.ControlTipText = TBRS01.Text
Call resimboyut2
End Sub
Sub resimboyut2()
Dim k As Integer
For k = 1 To 4
If Controls("LBRX1") = k Then Controls("Image" & k & "0").BorderColor = &H7379EC Else Controls("Image" & k & "0").BorderColor = &H80000000
Controls("Image" & k & "0").Height = Controls("Image" & k & "0").Height + 1
Controls("Image" & k & "0").Height = Controls("Image" & k & "0").Height - 1
Controls("Image" & k).Height = Controls("Image" & k).Height + 1
Controls("Image" & k).Height = Controls("Image" & k).Height - 1
Next
Image0.Height = Image0.Height + 1: Image0.Height = Image0.Height - 1
End Sub
Private Sub Label61_Click()
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3))
'CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\JPG"
End Sub
Private Sub Label62_Click()
If TBRS01 = "" Then Exit Sub
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\PDF\" & TBRS01 & ".pdf")
If a = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\PDF\" & TBRS01 & ".pdf"
Else
MsgBox ("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\PDF\" & TBRS01 & ".pdf" & " dosyası mevcut değil! "), vbInformation, "scngnr@hotmail.com": Exit Sub
End If
End Sub
Private Sub Label63_Click()
On Error GoTo hata
Dim RetStat
     'Me.Zoom = 200
RetStat = Application.Dialogs(xlDialogPrinterSetup).Show
If RetStat Then Me.PrintForm
hata:
     'Me.Zoom = 100
End Sub
Private Sub Label64_Click()
If TBRS01 = "" Then Exit Sub
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\DXF\" & TBRS01 & ".dxf")
If a = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\DXF\" & TBRS01 & ".dxf"
Else
MsgBox ("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & Trim(Left(sl1, 3)) & "\DXF\" & TBRS01 & ".dxf" & " dosyası mevcut değil! "), vbInformation, "scngnr@hotmail.com": Exit Sub
End If
End Sub
Private Sub SBmk_Change()
resimboyut
End Sub
Sub resimboyut()
kt1 = ((SBmk.Value - SBmk.Min + 270) / 270)
UFKWP.Width = SBmk.Min * kt1
UFKWP.Height = SBmk.Max * kt1
Image0.Width = UFKWP.Width - 14
Image0.Height = UFKWP.Height - 200
Frame1.Left = (UFKWP.Width - Frame1.Width) / 2
TBRS01.Width = UFKWP.Width - 90
Frame4.Left = UFKWP.Width - 50
Frame2.Top = UFKWP.Height - 85
Frame2.Left = (UFKWP.Width - Frame2.Width) / 2
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '
On Error Resume Next
UFKW.Toolbar1.Buttons.Item(8).Image = UFKW.ImageList1.ListImages.Item(8).Index
sUF = 1
End Sub
Private Sub CommandButton11_Click()
CreateObject("WScript.Network").SetDefaultPrinter "Microsoft Print to PDF"
fileName = ActiveWorkbook.path & Application.PathSeparator & "ac"
Me.PrintForm
End Sub