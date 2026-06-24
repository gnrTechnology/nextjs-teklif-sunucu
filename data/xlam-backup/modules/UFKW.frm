Option Explicit '
Dim kur, mkur As String
Dim n As Integer, z As Integer, i As Integer, Y As Integer, m As Integer
Dim ad As Integer, son As Integer
Dim X
Dim t
Dim fco
Dim ft1 As Byte
Dim fyt, tfyt, gek, prd, tire
Dim rs As String
Dim rsy As String
Public bfyt As String
Sub Resimler() 'Resimler Giriş
If ListBoxMG.listIndex >= 0 Then
TextBox8.Text = ListBoxMG.List(ListBoxMG.listIndex): Call Resimler1
Else
UFKWP.Image10.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\Logo.jpg")
UFKWP.Image0.Picture = UFKWP.Image10.Picture
Exit Sub
End If
If ListBoxMG2.listIndex >= 0 Then TextBox8.Text = ListBoxMG2.List(ListBoxMG2.listIndex): Call Resimler2 Else Exit Sub
If LBdbg.listIndex >= 0 Then TextBox8.Text = LBdbg.List(LBdbg.listIndex): Call Resimler3
If LBL01 = 1 Then
TextBox8.Text = ListViewA.SelectedItem: prd = TextBox8.Text
rsy = Left(ListViewA.SelectedItem.ListSubItems(4), 3): rsy = Trim(rsy)
Call Resimler4
End If
If LBL01 = 3 Then Lb_Click
End Sub
Sub Resimler1() 'Otomatik Seçim Dosyaları1
On Error GoTo hata
Dim rd, a
Dim k As Integer
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & TextBox8.Text & ".jpg")
If a = True Then
UFKWP.Image10.Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & TextBox8.Text & ".jpg")
Else
UFKWP.Image10.Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & "\Logo.jpg")
End If
UFKWP.Image10.Tag = 1: UFKWP.LBRX1 = 1: GoTo atla
hata: UFKWP.Image10.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg"): UFKWP.Image10.Tag = 2
atla:
'---
UFKWP.Image20.Tag = 2: UFKWP.Image20.Picture = LoadPicture()
UFKWP.Image30.Tag = 2: UFKWP.Image30.Picture = LoadPicture()
UFKWP.Image40.Tag = 2: UFKWP.Image40.Picture = LoadPicture()
For k = 1 To 4
UFKWP.Controls("Image1" & k).Picture = LoadPicture(): UFKWP.Controls("Image1" & k).Tag = 2
UFKWP.Controls("Image2" & k).Picture = LoadPicture(): UFKWP.Controls("Image2" & k).Tag = 2
UFKWP.Controls("Image3" & k).Picture = LoadPicture(): UFKWP.Controls("Image3" & k).Tag = 2
UFKWP.Controls("Image4" & k).Picture = LoadPicture(): UFKWP.Controls("Image4" & k).Tag = 2
UFKWP.Controls("Image" & k).Picture = LoadPicture(): UFKWP.Controls("Image" & k).Tag = 2
Next
UFKWP.Image0.Picture = UFKWP.Image10.Picture
'resim seçim
prd = TextBox8.Text
z = 1: tire = "_"
For k = 1 To 4
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & prd & "_" & k & ".jpg")
If a = True Then
UFKWP.Controls("Image1" & z).Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & prd & "_" & k & ".jpg")
UFKWP.Controls("Image1" & z).Tag = 1: UFKWP.Controls("Image" & z).Tag = 1
UFKWP.Controls("Image" & z).Picture = UFKWP.Controls("Image1" & z).Picture
z = z + 1
End If
Next
Call resimboyut1
UFKWP.TBRS02 = TextBox8.Text: UFKWP.TBRS01 = TextBox8.Text
UFKWP.Image0.ControlTipText = UFKWP.TBRS01.Text: UFKWP.Image10.ControlTipText = UFKWP.TBRS01.Text
Set rd = Nothing
End Sub
Sub Resimler2() 'Otomatik Seçim Dosyaları2
On Error GoTo hata
Dim rd, a
Dim k As Integer
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & TextBox8.Text & ".jpg")
If a = True Then
UFKWP.Image20.Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & TextBox8.Text & ".jpg")
UFKWP.Image20.Tag = 1
Else
UFKWP.Image20.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
UFKWP.Image20.Tag = 2: GoTo atla
hata: UFKWP.Image20.Tag = 2: UFKWP.Image20.Picture = LoadPicture()
End If
atla:
'---
UFKWP.Image40.Tag = 2: UFKWP.Image40.Picture = LoadPicture()
For k = 1 To 4
UFKWP.Controls("Image2" & k).Picture = LoadPicture(): UFKWP.Controls("Image2" & k).Tag = 2
'UFKWP.Controls("Image3" & k).Picture = LoadPicture(): UFKWP.Controls("Image3" & k).Tag = 2
UFKWP.Controls("Image4" & k).Picture = LoadPicture(): UFKWP.Controls("Image4" & k).Tag = 2
UFKWP.Controls("Image" & k).Picture = LoadPicture(): UFKWP.Controls("Image" & k).Tag = 2
Next
UFKWP.Image0.Picture = UFKWP.Image20.Picture: UFKWP.LBRX1 = 2
'resim seçim
prd = TextBox8.Text
z = 1: tire = "_"
For k = 1 To 4
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & prd & "_" & k & ".jpg")
If a = True Then
UFKWP.Controls("Image2" & z).Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & prd & "_" & k & ".jpg")
UFKWP.Controls("Image2" & z).Tag = 1: UFKWP.Controls("Image" & z).Tag = 1
UFKWP.Controls("Image" & z).Picture = UFKWP.Controls("Image2" & z).Picture
z = z + 1
End If
Next
Call resimboyut1
UFKWP.TBRS01 = UFKWP.TBRS02 & " " & TextBox8.Text
UFKWP.Image0.ControlTipText = UFKWP.TBRS01.Text: UFKWP.Image20.ControlTipText = UFKWP.TBRS01.Text
Set rd = Nothing
End Sub
Sub Resimler3() 'ek grup Resimler
On Error GoTo hata
Dim rd, a
Dim k As Integer
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & TextBox8.Text & ".jpg")
If a = True Then
UFKWP.Image30.Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & TextBox8.Text & ".jpg")
UFKWP.Image30.Tag = 1
Else
UFKWP.Image30.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
UFKWP.Image30.Tag = 2: GoTo atla
hata: UFKWP.Image30.Tag = 2: UFKWP.Image30.Picture = LoadPicture()
End If
atla:
'---
UFKWP.Image40.Tag = 2: UFKWP.Image40.Picture = LoadPicture()
For k = 1 To 4
UFKWP.Controls("Image3" & k).Picture = LoadPicture(): UFKWP.Controls("Image3" & k).Tag = 2
UFKWP.Controls("Image4" & k).Picture = LoadPicture(): UFKWP.Controls("Image4" & k).Tag = 2
UFKWP.Controls("Image" & k).Picture = LoadPicture(): UFKWP.Controls("Image" & k).Tag = 2
Next
UFKWP.Image0.Picture = UFKWP.Image30.Picture: UFKWP.LBRX1 = 3
'resim seçim
prd = TextBox8.Text
z = 1: tire = "_"
For k = 1 To 4
a = rd.FileExists("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & prd & "_" & k & ".jpg")
If a = True Then
UFKWP.Controls("Image3" & z).Picture = LoadPicture("C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\" & rs & "\JPG\" & prd & "_" & k & ".jpg")
UFKWP.Controls("Image3" & z).Tag = 1: UFKWP.Controls("Image" & z).Tag = 1
UFKWP.Controls("Image" & z).Picture = UFKWP.Controls("Image3" & z).Picture
z = z + 1
End If
Next
Call resimboyut1
UFKWP.TBRS01 = UFKWP.TBRS02 & " " & TextBox8.Text
UFKWP.Image0.ControlTipText = UFKWP.TBRS01.Text: UFKWP.Image30.ControlTipText = UFKWP.TBRS01.Text
Set rd = Nothing
End Sub
Sub Resimler4() 'Malzeme Resimleri
On Error Resume Next
Dim rd, a
Dim z As Integer, k As Integer
Set rd = CreateObject("Scripting.FileSystemObject")
For k = 1 To 4
UFKWP.Controls("Image4" & k).Picture = LoadPicture(): UFKWP.Controls("Image4" & k).Tag = 2
UFKWP.Controls("Image" & k).Picture = LoadPicture(): UFKWP.Controls("Image" & k).Tag = 2
Next
'resim seçim
z = 0: tire = ""
For k = 0 To 4
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rsy & "\" & prd & tire & ".jpg")
If a = True Then
UFKWP.Controls("Image4" & z).Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rsy & "\" & prd & tire & ".jpg")
UFKWP.Controls("Image4" & z).Tag = 1
If k > 0 Then UFKWP.Controls("Image" & z).Picture = UFKWP.Controls("Image4" & z).Picture: UFKWP.Controls("Image" & z).Tag = 1
z = z + 1
End If
tire = "_" & k + 1
Next
'---
UFKWP.LBRX1 = 4
If z = 0 Then UFKWP.Image40.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
UFKWP.Image0.Picture = UFKWP.Image40.Picture
Call resimboyut1
UFKWP.TBRS01 = prd
UFKWP.Image40.ControlTipText = prd: UFKWP.Image0.ControlTipText = prd
Set rd = Nothing
End Sub
Sub resimboyut1()
Dim k As Integer
For k = 1 To 4
If UFKWP.Controls("LBRX1") = k Then UFKWP.Controls("Image" & k & "0").BorderColor = &H7379EC Else UFKWP.Controls("Image" & k & "0").BorderColor = &H80000000
UFKWP.Controls("Image" & k & "0").Height = UFKWP.Controls("Image" & k & "0").Height + 1
UFKWP.Controls("Image" & k & "0").Height = UFKWP.Controls("Image" & k & "0").Height - 1
UFKWP.Controls("Image" & k).Height = UFKWP.Controls("Image" & k).Height + 1
UFKWP.Controls("Image" & k).Height = UFKWP.Controls("Image" & k).Height - 1
Next
UFKWP.Image0.Height = UFKWP.Image0.Height + 1: UFKWP.Image0.Height = UFKWP.Image0.Height - 1
End Sub
Private Sub CK1_Click()
If CK1.Value = True Then CKB1.Value = False
End Sub
Private Sub CKB1_Click()
If CKB1.Value = True Then CK1.Value = False
End Sub
Private Sub CommandButton7_Click()
If LB.ListCount < 1 Then Exit Sub
MultiPage3.Value = 0
z = ListViewA.ListItems.Count
If z < 0 Then Exit Sub
Call grupekle
Call tool1
End Sub
Private Sub CB50_Click()
If ListViewA.ListItems.Count > 0 And MultiPage3.Value = 0 Then
  If gir.BackColor = &H86B57D Then
  CB50.Caption = "Adet"
  gir.BackColor = &H7379EC
  CB50.BackColor = &H7379EC
  Y = 1
  Do Until Y > ListViewA.ListItems.Count
  ListViewA.ListItems(Y).SmallIcon = 8
  Y = Y + 1
  Loop
  Else
  gir.Caption = "Teklife Gir"
  CB50.Caption = "Grup"
  gir.BackColor = &H86B57D
  CB50.BackColor = &H86B57D
  End If
End If
End Sub
Private Sub Image442_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim otmdetay As String
otmdetay = "1. AÇIKLAMA METİNLERİ (Ürün Grupları Seçimi)" & vbLf & "2. EK-? (Ürün Grup Seçimi)" & vbLf & "    Örnek: EK-1 (1.Ürün Grupları İlavelerini Gösterir)" _
& vbLf & "3. LİSTE-EK (Ürün Grup İlavesi Oluşturma Ve Liste Seçimleri)" & vbLf & "4. LİSTE (Liste Seçimleri)"
MsgBox "Otomatik Dosya Oluşturma Format Detayları" & vbLf & "A SUTUNUNA YAZILACAK METİNLER:" & vbLf & otmdetay, vbInformation, "scngnr@hotmail.com"
End Sub
Private Sub SpinButton2_SpinUp()
On Error GoTo hata
TBADET.Value = SpinButton2.Value
hata:
End Sub
Private Sub SpinButton2_SpinDown()
On Error GoTo hata
TBADET.Value = SpinButton2.Value
hata:
End Sub
Private Sub TBADET_Change()
On Error Resume Next
SpinButton2.Value = CDbl(TBADET.Value)
If gir.BackColor = &H7379EC Then
ListViewA.SelectedItem.ListSubItems(2).Text = TBADET.Value
Else
tfyt = 0
i = 1
z = ListViewA.ListItems.Count
Do Until i > z
    ListViewA.ListItems(i).ListSubItems(2).Text = (ListViewA.ListItems(i).ListSubItems(2).Text / ad) * SpinButton2.Value
    i = i + 1
Loop
ad = CDbl(TBADET.Value)
End If
Call toplamFiyat
End Sub
Sub toplamFiyat()
tfyt = 0
i = 1
z = ListViewA.ListItems.Count
Do Until i > z
    fyt = 0
    If ListViewA.ListItems(i).ListSubItems(1).ForeColor = &H0& Then
     If ListViewA.ListItems(i).ListSubItems(5).Text <> "" Then
     fyt = ListViewA.ListItems(i).ListSubItems(5).Text * ListViewA.ListItems(i).ListSubItems(2).Text
     End If
    End If
    tfyt = tfyt + fyt
    i = i + 1
Loop
LBGF.Caption = Format(tfyt, "#,##0.00") & " TL"
End Sub
Private Sub UserForm_Initialize() 'form yükleme +1
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
'--
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
Toolbar1.Buttons.Item(4).Image = ImageList1.ListImages.Item(4).Index
Toolbar1.Buttons.Item(5).Image = ImageList1.ListImages.Item(5).Index
Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(6).Index
Toolbar1.Buttons.Item(7).Image = ImageList1.ListImages.Item(7).Index
Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(8).Index
Toolbar1.Buttons.Item(9).Image = ImageList1.ListImages.Item(9).Index
Toolbar1.Buttons.Item(10).Image = ImageList1.ListImages.Item(10).Index
'--
'Application.ScreenUpdating = False
ActiveWorkbook.Worksheets("Sayfa1").Select
ListViewA.SmallIcons = ImageList2: ListViewA.Gridlines = False
ListViewA.Width = 456
On Error GoTo hata
UFKW.Caption = Workbooks(sl1).Name
fco = 30
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
n = ssayfa.Range("B65536").End(xlUp).row
'***MALZEME SEÇİM--
If ssayfa.CodeName = "OTMM" Then Toolbar1.Buttons.Item(5).Visible = True 'motor seçim
If ssayfa.CodeName = "OTMOG" Then CK1.Value = True 'OG seçim
ListBoxMG.ListStyle = fmListStyleOption
Dim mss
For i = 2 To n
If ssayfa.Cells(i, 2) <> "BÖLÜM ADI/NO:" Then GoTo atla
If Left(ssayfa.Cells(i, 1), 5) = "LİSTE" Then GoTo atla
If Left(ssayfa.Cells(i, 1), 2) = "EK" Then GoTo atla
mss = WorksheetFunction.CountIf(ssayfa.Range("A2:A" & i), ssayfa.Cells(i, 1).Value)
   If mss = 1 Then
     If ssayfa.Cells(i, 1) <> "" Then ListBoxMG.AddItem ssayfa.Cells(i, 1)
   End If
atla:
Next i
rs = Left(sl1, 3): rs = Trim(rs)
hata:
'Application.ScreenUpdating = True
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button) '+1
On Error Resume Next
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
n = ssayfa.Range("B65536").End(xlUp).row
z = ListViewA.ListItems.Count
'--
Select Case Button.Index
Case 1
MultiPage3.Value = 0
ListViewA.Width = 455.2
LB.Selected(LB.listIndex) = False
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index Then
'UFKWP.resimboyut21
End If

Case 2 'gruba ek ürün seçmek
'Toolbar1.Buttons.item(3).Enabled = False
Dim EK
ComboBoxL.Visible = False: Label47.Visible = True
MultiPage3.Value = 1
MultiPage4.Value = 0
'If ss1 = 1 Then
LBdbg.Clear
i = 1
Do Until i > n
EK = ListBoxMG.listIndex + 1
If ssayfa.Cells(i, 2) = "BÖLÜM ADI/NO:" And ssayfa.Cells(i, 1) = "EK-" & EK Then
z = LBdbg.ListCount
LBdbg.AddItem ssayfa.Cells(i, 3): LBdbg.List(z, 1) = ssayfa.Cells(i, 3).row 'satırıncı
z = z + 1
End If
i = i + 1
Loop
'End If
'--
Case 3 'gruba ürün ekleme
If z < 0 Then Exit Sub
Call grupekle
'--
Case 4 'liste ürün ekleme
MultiPage3.Value = 2
Toolbar1.Buttons.Item(3).Enabled = False
'If ss1 = 1 Then
LBdb3.Clear
i = 1
Do Until i > n
If ssayfa.Cells(i, 2) = "BÖLÜM ADI/NO:" Then
If Left(ssayfa.Cells(i, 1), 5) = "LİSTE" Then
z = LBdb3.ListCount
LBdb3.AddItem ssayfa.Cells(i, 3): LBdb3.List(z, 1) = ssayfa.Cells(i, 3).row 'satırıncı
z = z + 1
End If
End If
i = i + 1
Loop
'End If
'--
Case 5 'termik röle ekleme
If Toolbar1.Buttons.Item(5).Value = tbrPressed Then
Call ListViewA.ListItems.Add((TBLTER.Value + 1), , ssayfa.Cells(TBSTER.Value, 2), , 1)
Call ListViewA.ListItems(TBLTER.Value + 1).ListSubItems.Add(1, , ssayfa.Cells(TBSTER.Value, 3))
Call ListViewA.ListItems(TBLTER.Value + 1).ListSubItems.Add(2, , ssayfa.Cells(TBSTER.Value, 5).Text * TBADET.Value)
Call ListViewA.ListItems(TBLTER.Value + 1).ListSubItems.Add(3, , ssayfa.Cells(TBSTER.Value, 2).row)
Call ListViewA.ListItems(TBLTER.Value + 1).ListSubItems.Add(4, , ssayfa.Cells(TBSTER.Value, 4))
Toolbar1.Buttons.Item(5).Caption = "Termik Röle Çıkar"
'fiyat--
Y = CDbl(TBSTER.Value)
fyt = ssayfa.Cells(Y, 6) - (ssayfa.Cells(Y, 6) * ssayfa.Cells(Y, 7))
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Usd") '$
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$€-1]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Eur") '£
Call ListViewA.ListItems(TBLTER.Value + 1).ListSubItems.Add(5, , Format(fyt, "#,##0.00"))
'--
Else
ListViewA.ListItems.Remove (TBLTER.Value + 1)
Toolbar1.Buttons.Item(5).Caption = "Termik Röle Ekle"
End If
TBLSON.Value = ListViewA.ListItems.Count
Call toplamFiyat
ListViewA.Refresh
'--
Case 6 'ürün aktif
Y = 1
Do Until Y > z
If ListViewA.ListItems(Y).SmallIcon = 2 Then ListViewA.ListItems(Y).SmallIcon = 1
If ListViewA.ListItems(Y).SmallIcon = 5 Then ListViewA.ListItems(Y).SmallIcon = 4
If ListViewA.ListItems(Y).SmallIcon = 9 Then ListViewA.ListItems(Y).SmallIcon = 8
If ListViewA.ListItems(Y).SmallIcon = 7 Then ListViewA.ListItems(Y).SmallIcon = 3
If ListViewA.ListItems(Y).SmallIcon = 15 Then ListViewA.ListItems(Y).SmallIcon = 14
ListViewA.ListItems(Y).ForeColor = &H0&
ListViewA.ListItems(Y).ListSubItems(1).ForeColor = &H0&
ListViewA.ListItems(Y).ListSubItems(2).ForeColor = &H0&
ListViewA.ListItems(Y).ListSubItems(4).ForeColor = &H0&
Y = Y + 1
Loop
Y = CDbl(TBSTER.Value)
Call toplamFiyat
'--
Case 7 'ürün pasif
Y = 1
Do Until Y > z
If ListViewA.ListItems(Y).SmallIcon = 1 Then ListViewA.ListItems(Y).SmallIcon = 2
If ListViewA.ListItems(Y).SmallIcon = 4 Then ListViewA.ListItems(Y).SmallIcon = 5
If ListViewA.ListItems(Y).SmallIcon = 8 Then ListViewA.ListItems(Y).SmallIcon = 9
If ListViewA.ListItems(Y).SmallIcon = 3 Then ListViewA.ListItems(Y).SmallIcon = 7
If ListViewA.ListItems(Y).SmallIcon = 14 Then ListViewA.ListItems(Y).SmallIcon = 15
ListViewA.ListItems(Y).ForeColor = &H7379EC
ListViewA.ListItems(Y).ListSubItems(1).ForeColor = &H7379EC
ListViewA.ListItems(Y).ListSubItems(2).ForeColor = &H7379EC
ListViewA.ListItems(Y).ListSubItems(4).ForeColor = &H7379EC
Y = Y + 1
Loop
Call toplamFiyat
'--
Case 8
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(8).Index Then
Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index
sUF = 2
UFKWP.Show
Call Resimler
Else
Unload UFKWP
End If
Case 9
If UFKW.Height > 240 Then
UFKW.Height = UFKW.Height - UFKW.InsideHeight + Toolbar1.Height
Else
UFKW.Height = 295
End If
GoTo atla:
Case 10
Unload Me: Exit Sub
End Select
UFKW.Height = 295
'UFKW.Height = 260
atla:
Call tool1
'If MultiPage1.Value = 2 Or MultiPage1.Value = 1 Then Toolbar1.Buttons.item(3).Enabled = False Else Toolbar1.Buttons.item(3).Enabled = True
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Tag
Case 1
If ListViewA.ListItems.Count > 0 And MultiPage3.Value = 0 Then
CB50.Caption = "Adet": CB50.BackColor = &H7379EC: gir.BackColor = &H7379EC
Y = 1
Do Until Y > ListViewA.ListItems.Count
ListViewA.ListItems(Y).SmallIcon = 8
'If ListViewA.ListItems(Y).SmallIcon = 1 Or ListViewA.ListItems(Y).SmallIcon = 2 Then ListViewA.ListItems(Y).SmallIcon = 8
Y = Y + 1
Loop
End If
'--
Case 2 'GRUP-LİSTE
ComboBoxL.Visible = True: Label47.Visible = False
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
n = ssayfa.Range("B65536").End(xlUp).row
Toolbar1.Buttons.Item(3).Enabled = False
MultiPage3.Value = 1
MultiPage4.Value = 1
'If ss1 = 1 Then
'LBdb.Clear
LBdl.Clear
ComboBoxL.Clear
ComboBoxL.Value = ""
i = 1
Do Until i > n
If ssayfa.Cells(i, 2) = "BÖLÜM ADI/NO:" And Left(ssayfa.Cells(i, 1), 8) = "LİSTE-EK" Then
'z = LBdb.ListCount
z = ComboBoxL.ListCount
'LBdb.AddItem ssayfa.Cells(i, 3): LBdb.List(z, 1) = ssayfa.Cells(i, 3).Row 'satırıncı
ComboBoxL.AddItem ssayfa.Cells(i, 3): ComboBoxL.List(z, 1) = ssayfa.Cells(i, 3).row 'satırıncı
z = z + 1
End If
i = i + 1
Loop
ComboBoxL.SelText = "Ürün Grubunu Seçiniz.."
'End If
'--
Case 3 'GRUP-LİSTE çıkar
Call grupeksil
Case 4 'veri dosyasını aç
Workbooks(sl1).Activate: Application.Windows(sl1).Visible = True: End
End Select
Call tool1
End Sub
Sub tool1()
If MultiPage3.Value = 0 And ListViewA.ListItems.Count > 0 Then Frame3.Enabled = True: _
gir.Enabled = True: CK1.Enabled = True Else Frame3.Enabled = False: gir.Enabled = False: CK1.Enabled = False
If MultiPage3.Value = 0 Then
If ListViewA.ListItems.Count > 0 And LB.ListCount > 0 Then Toolbar1.Buttons.Item(3).Enabled = True Else Toolbar1.Buttons.Item(3).Enabled = False
If ListViewA.ListItems.Count > 0 And ft1 = 1 Then Toolbar1.Buttons.Item(5).Enabled = True Else Toolbar1.Buttons.Item(5).Enabled = False
If ListViewA.ListItems.Count > 0 Then Toolbar1.Buttons.Item(6).Enabled = True Else Toolbar1.Buttons.Item(6).Enabled = False
If ListViewA.ListItems.Count > 0 Then Toolbar1.Buttons.Item(7).Enabled = True Else Toolbar1.Buttons.Item(7).Enabled = False
Else
Toolbar1.Buttons.Item(3).Enabled = False
Toolbar1.Buttons.Item(5).Enabled = False
Toolbar1.Buttons.Item(6).Enabled = False
Toolbar1.Buttons.Item(7).Enabled = False
End If
If ListViewA.ListItems.Count > 0 Then Toolbar1.Buttons.Item(2).Enabled = True Else Toolbar1.Buttons.Item(2).Enabled = False
End Sub
Sub grupekle() 'gruba ürün ekleme
On Error Resume Next
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
z = ListViewA.ListItems.Count
i = CDbl(TBLSON.Value)
Do Until i >= z
ListViewA.ListItems.Remove (z)
    z = z - 1
Loop
Dim ymlst
ymlst = LB.ListCount - 1
i = 0
z = ListViewA.ListItems.Count
Do Until i > ymlst

If LB.List(i, 4) = "" Then Call ListViewA.ListItems.Add((z + 1), , LB.List(i, 1), , 14) Else _
 Call ListViewA.ListItems.Add((z + 1), , LB.List(i, 1), , 4)
'Call ListViewA.ListItems.Add((z + 1), , LB.List(i, 1), , 4)
Call ListViewA.ListItems(z + 1).ListSubItems.Add(1, , LB.List(i, 2))
If LB.List(i, 4) = "" Then ListViewA.ListItems(z + 1).ListSubItems(1).ReportIcon = 13
Call ListViewA.ListItems(z + 1).ListSubItems.Add(2, , LB.List(i) * TBADET.Value)
Call ListViewA.ListItems(z + 1).ListSubItems.Add(3, , LB.List(i, 3)) 'satır no
Call ListViewA.ListItems(z + 1).ListSubItems.Add(4, , LB.List(i, 4))
'fiyat--
Y = LB.List(i, 3) * 1
fyt = ssayfa.Cells(Y, 6) - (ssayfa.Cells(Y, 6) * ssayfa.Cells(Y, 7))
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Usd") '$
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$€-1]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Eur") '£
Call ListViewA.ListItems(z + 1).ListSubItems.Add(5, , Format(fyt, "#,##0.00"))
'--
    z = z + 1
    i = i + 1
Loop
gek = ListViewA.ListItems.Count - TBLSON.Value
Call toplamFiyat
ListViewA.Refresh
TextBoxGA1.Tag = 1
End Sub
Sub grupeksil()
z = ListViewA.ListItems.Count
If gek = 0 Then Exit Sub
i = z - gek
Do Until i >= z
ListViewA.ListItems.Remove (z)
    z = z - 1
Loop
gek = 0
Call toplamFiyat
ListViewA.Refresh
TextBoxGA1.Tag = 0
End Sub
Private Sub ComboBoxL_Click()
On Error Resume Next
LBdl.Clear
'TBYADET.Value = 1
SpinButton1.Value = 1
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
'--
Dim t
t = ComboBoxL.listIndex
z = 0
Y = ComboBoxL.List(t, 1)
n = ssayfa.Cells(Y + 1, 4).End(xlDown).row
If n > ssayfa.Range("B65536").End(xlUp).row Then n = ssayfa.Range("B65536").End(xlUp).row
Y = Y + 1
Do Until Y > n
If Y > 5000 Then Exit Sub
 LBdl.AddItem ssayfa.Cells(Y, 2)
 LBdl.List(z, 1) = ssayfa.Cells(Y, 3)
 LBdl.List(z, 2) = ssayfa.Cells(Y, 5).Text
 LBdl.List(z, 3) = ssayfa.Cells(Y, 2).row
 LBdl.List(z, 4) = ssayfa.Cells(Y, 4).Text
 If ssayfa.Cells(Y, 5) = Int(ssayfa.Cells(Y, 5)) Then LBdl.List(z, 5) = 1 Else LBdl.List(z, 5) = ssayfa.Cells(Y, 5).Text
z = z + 1
'End If
Y = Y + 1
Loop
LBdl.SetFocus
End Sub
Private Sub LBdbg_Click() 'EK GRUP
On Error Resume Next
If LBdbg.ListCount = 0 Then Exit Sub
LB.Clear
'TBYADET.Value = 1
SpinButton1.Value = 1
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
'--
Dim t
t = LBdbg.listIndex
Y = LBdbg.List(t, 1)
'n = ssayfa.Cells(Y + 1, 4).End(xlDown).Row
TextBoxGA1.Text = " - " & ssayfa.Cells(Y, 3)
Call grupeksil
'--
Y = Y + 1
'--
n = ssayfa.Range("B65536").End(xlUp).row
For i = Y To n
If ssayfa.Cells(i, 2) = "BÖLÜM ADI/NO:" Then n = i - 1: Exit For
Next i
'--
z = 0
Do Until Y > n
If Y > 5000 Then Exit Sub
'If ssayfa.Cells(Y, 4) = "" Then Exit Do
 LB.AddItem ssayfa.Cells(Y, 5).Text
 LB.List(z, 1) = ssayfa.Cells(Y, 2)
 LB.List(z, 2) = ssayfa.Cells(Y, 3).Text
 LB.List(z, 3) = ssayfa.Cells(Y, 2).row
 LB.List(z, 4) = ssayfa.Cells(Y, 4).Text
 If ssayfa.Cells(Y, 5) = Int(ssayfa.Cells(Y, 5)) Then LB.List(z, 5) = 1 Else LB.List(z, 5) = ssayfa.Cells(Y, 5).Text
z = z + 1
'End If
Y = Y + 1
Loop
LB.SetFocus
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index Then
prd = LBdbg.Text
LBL01 = 2
TextBox8.Text = prd '2023
Call Resimler3
End If
End Sub
Private Sub LBdb3_Click() 'LİSTE
On Error Resume Next
LBdl3.Clear
TBEADET.Value = 1
SpinButton3.Value = 1
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
'--
Dim t
t = LBdb3.listIndex
z = 0
Y = CDbl(LBdb3.List(t, 1))
n = ssayfa.Cells(Y + 1, 4).End(xlDown).row
If n > ssayfa.Range("B65536").End(xlUp).row Then n = ssayfa.Range("B65536").End(xlUp).row
Y = Y + 1
Do Until Y > n
If Y > 5000 Then Exit Sub
 LBdl3.AddItem ssayfa.Cells(Y, 2)
 LBdl3.List(z, 1) = ssayfa.Cells(Y, 3)
 LBdl3.List(z, 2) = ssayfa.Cells(Y, 5).Text
 LBdl3.List(z, 3) = ssayfa.Cells(Y, 2).row
 LBdl3.List(z, 4) = ssayfa.Cells(Y, 4).Text
 'b. fiyat--
fyt = ssayfa.Cells(Y, 6)
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Usd") '$
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$€-1]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Eur") '£
 LBdl3.List(z, 5) = Format(fyt, "#,##0.00")
'--
z = z + 1
'End If
Y = Y + 1
Loop
AlignListColumn LBdl3, 5, True
Frame5.Enabled = False: girEK.Enabled = False
End Sub
Sub AlignListColumn(LBox As MSForms.ListBox, WhichColumn As Integer, AlignRight As Boolean)
Dim vntColWidths As Variant
Dim sngWidth As Single
Dim strTemp As String
Dim labTester As MSForms.Label
Dim intItem As Integer
If WhichColumn > LBox.ColumnCount Then Exit Sub
Set labTester = Me.LabelP31
labTester.WordWrap = False
With LBox
If .ColumnWidths <> "" Then
vntColWidths = Split(.ColumnWidths, ";")
sngWidth = val(57) - 1
Else
sngWidth = (.Width / .ColumnCount) - ((.ColumnCount - 1) * 3)
End If
If sngWidth <= 0 Then Exit Sub
For intItem = 0 To .ListCount - 1
strTemp = Trim(.List(intItem, WhichColumn))
labTester.AutoSize = False
labTester.Width = .Width
labTester.Caption = strTemp
labTester.AutoSize = True
Do While labTester.Width <= sngWidth
labTester.Caption = " " & labTester.Caption
Loop
.List(intItem, WhichColumn) = labTester.Caption
Next
End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '
On Error Resume Next
Application.ScreenUpdating = False
    Windows(dt).Activate
    'Sheets("Sayfa1").Select
    Windows(sl1).Close False
    Application.Windows(sl1).Visible = True
    sl1 = Empty
Application.ScreenUpdating = True
If sUF = 2 Then Unload UFKWP
sUF = 0
End Sub
Private Sub CommandButton5_Click() 'yardımcı ürün seçileni sil
On Error Resume Next
If LB.listIndex < 0 Then Exit Sub
LB.RemoveItem (LB.listIndex)
End Sub
Private Sub CommandButton6_Click() 'tüm yardımcı ürünleri sil
LB.Clear
End Sub
Private Sub Lb_Click()
LBL01 = 3
If LB.ListCount - 1 < 0 Then Exit Sub
TBYADET.Value = LB.List(LB.listIndex, 0)
rsy = Left(LB.List(LB.listIndex, 4), 3): rsy = Trim(rsy)
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index Then
prd = LB.List(LB.listIndex, 1)
Call Resimler4
End If
End Sub
Private Sub SpinButton1_SpinUp()
On Error GoTo hata
If LB.List(LB.listIndex, 0) = 999 Then Exit Sub
If LB.List(LB.listIndex, 5) >= 1 Then
SpinButton1.Value = LB.List(LB.listIndex, 0)
LB.List(LB.listIndex, 0) = SpinButton1.Value + 1
TBYADET.Value = SpinButton1.Value + 1
Else
LB.List(LB.listIndex, 0) = Replace(LB.List(LB.listIndex, 0) * 1 + LB.List(LB.listIndex, 5) * 1, ".", ",")
End If
hata:
'Call yardmalzeme
End Sub
Private Sub SpinButton1_SpinDown()
On Error GoTo hata
If LB.List(LB.listIndex, 0) - LB.List(LB.listIndex, 5) <= 0 Then Exit Sub
If LB.List(LB.listIndex, 5) >= 1 Then
SpinButton1.Value = LB.List(LB.listIndex, 0)
LB.List(LB.listIndex, 0) = SpinButton1.Value - 1
TBYADET.Value = SpinButton1.Value - 1
Else
LB.List(LB.listIndex, 0) = Replace(LB.List(LB.listIndex, 0) * 1 - LB.List(LB.listIndex, 5) * 1, ".", ",")
End If
hata:
'Call yardmalzeme
End Sub
Private Sub TBYADET_Change()
If LB.listIndex >= 0 Then LB.List(LB.listIndex, 0) = TBYADET
End Sub
Private Sub ListBoxMG_Click()
On Error Resume Next
Label2.BackColor = &H7379EC: Label5.BackColor = &HD7BBA2
ListViewA.ListItems.Clear: gek = 0
gir.Caption = "Teklife Gir": gir.BackColor = &H86B57D: CB50.Caption = "Grup": CB50.BackColor = &H86B57D
ListBoxMG2.Clear: LBdbg.Clear
LB.Clear
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
n = ssayfa.Range("B65536").End(xlUp).row
t = ListBoxMG.Text
i = 1
Do Until i > n
 If ssayfa.Cells(i, 1) = t Then
 z = ListBoxMG2.ListCount
 ListBoxMG2.AddItem ssayfa.Cells(i, 3): ListBoxMG2.List(z, 1) = ssayfa.Cells(i, 3).row 'satırıncı
     i = i + 1
     Do Until i > n
     If ssayfa.Cells(i, 2) = "BÖLÜM ADI/NO:" Then
      If ssayfa.Cells(i, 1) = "" Then
      z = ListBoxMG2.ListCount
      ListBoxMG2.AddItem ssayfa.Cells(i, 3): ListBoxMG2.List(z, 1) = ssayfa.Cells(i, 3).row '
      Else
      i = i - 1
      Exit Do
      End If
     End If
     i = i + 1
     Loop
End If
i = i + 1
Loop
bitti:
'If ss1 <> 1 Then XX Else XX
LBGF.Caption = "-"
Call tool1
TextBox8.Text = ListBoxMG.Text
LBL01 = 0
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index Then
UFKWP.Image40.Tag = 2: UFKWP.Image40.Picture = LoadPicture()
Call Resimler1
End If
End Sub
Private Sub ListBoxMG2_Click()
On Error Resume Next
Label5.BackColor = &H7379EC: Label2.BackColor = &HD7BBA2
'--
ListViewA.ListItems.Clear: gek = 0
'--
'ListViewA.ListItems.Clear
gir.Caption = "Teklife Gir": gir.BackColor = &H86B57D: CB50.Caption = "Grup": CB50.BackColor = &H86B57D
ListViewA.ColumnHeaders(2).Text = ListBoxMG2.Text & " için Seçimler"
'Label4.Caption = "Motor Gücü(kW)=" & ListBoxMG2.Text
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
'Windows(mlz).Activate
'Güç seçim--
Dim t
t = ListBoxMG2.listIndex
z = 0
tfyt = 0
'--
Y = ListBoxMG2.List(t, 1) + 1: son = ssayfa.Range("B65536").End(xlUp).row
For i = Y To son
If ssayfa.Cells(i, 2) = "BÖLÜM ADI/NO:" Then n = i - 1: Exit For
Next i
'--
'If n >= ssayfa.Range("B65536").End(xlUp).Row Then n = son
ft1 = 0
If Y > n Then n = son
Do Until Y > n
If ssayfa.Cells(Y, 1) = "FT" Then
TBSTER.Value = ssayfa.Cells(Y, 2).row: ft1 = 1
TBLTER.Value = z
If Toolbar1.Buttons.Item(5).Value = tbrUnpressed Then GoTo atla11:
End If
 If ssayfa.Cells(Y, 4) = "" Then Call ListViewA.ListItems.Add((z + 1), , ssayfa.Cells(Y, 2), , 14) Else _
Call ListViewA.ListItems.Add((z + 1), , ssayfa.Cells(Y, 2), , 1)
Call ListViewA.ListItems(z + 1).ListSubItems.Add(1, , ssayfa.Cells(Y, 3))
Call ListViewA.ListItems(z + 1).ListSubItems.Add(2, , ssayfa.Cells(Y, 5).Text)
Call ListViewA.ListItems(z + 1).ListSubItems.Add(3, , ssayfa.Cells(Y, 2).row)
Call ListViewA.ListItems(z + 1).ListSubItems.Add(4, , ssayfa.Cells(Y, 4).Text)
'fiyat--
If ssayfa.Cells(Y, 4) = "" Then
Call ListViewA.ListItems(z + 1).ListSubItems.Add(5, , "")
ListViewA.ListItems(z + 1).ListSubItems(1).ReportIcon = 13: z = z + 1: GoTo atla11
End If
fyt = ssayfa.Cells(Y, 6) - (ssayfa.Cells(Y, 6) * ssayfa.Cells(Y, 7))
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Usd") '$
If ssayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$€-1]" Then fyt = fyt * Workbooks(dt).Sheets("Sayfa3").Range("Eur") '£
Call ListViewA.ListItems(z + 1).ListSubItems.Add(5, , Format(fyt, "#,##0.00"))

tfyt = tfyt + (fyt * ssayfa.Cells(Y, 5))
z = z + 1
'--
atla11:
Y = Y + 1
Loop
ListBoxMG.SetFocus
ad = 1
TBADET.Value = 1
TBLSON.Value = ListViewA.ListItems.Count
'If ss1 <> 1 Then
 'If ListBoxMG.ListIndex >= 0 Then ListBoxMG.Selected(ListBoxMG.ListIndex) = False
'End If
'TOPLAM fiyat--
LBGF.Caption = Format(tfyt, "#,##0.00") & " TL"
'--
Call tool1
TextBox8.Text = ListBoxMG2.Text
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index Then Call Resimler2
ListViewA.Refresh
TextBoxGA1.Tag = 0
LBL01 = 0
End Sub
Private Sub ListViewA_DblClick()
On Error Resume Next
If ListViewA.ListItems.Count > 0 Then
Y = ListViewA.SelectedItem.Index
If ListViewA.ListItems(Y).SmallIcon = 2 Then ListViewA.ListItems(Y).SmallIcon = 1: GoTo git1
If ListViewA.ListItems(Y).SmallIcon = 5 Then ListViewA.ListItems(Y).SmallIcon = 4: GoTo git1
If ListViewA.ListItems(Y).SmallIcon = 9 Then ListViewA.ListItems(Y).SmallIcon = 8: GoTo git1
If ListViewA.ListItems(Y).SmallIcon = 15 Then ListViewA.ListItems(Y).SmallIcon = 14: GoTo git1

If ListViewA.ListItems(Y).SmallIcon = 1 Then ListViewA.ListItems(Y).SmallIcon = 2: GoTo git1
If ListViewA.ListItems(Y).SmallIcon = 4 Then ListViewA.ListItems(Y).SmallIcon = 5: GoTo git1
If ListViewA.ListItems(Y).SmallIcon = 8 Then ListViewA.ListItems(Y).SmallIcon = 9: GoTo git1
If ListViewA.ListItems(Y).SmallIcon = 14 Then ListViewA.ListItems(Y).SmallIcon = 15: GoTo git1
git1:
If ListViewA.ListItems(Y).SmallIcon = 1 Or ListViewA.ListItems(Y).SmallIcon = 4 Or _
ListViewA.ListItems(Y).SmallIcon = 8 Or ListViewA.ListItems(Y).SmallIcon = 14 Then
ListViewA.SelectedItem.ForeColor = &H0& 'yeşil
ListViewA.SelectedItem.ListSubItems(1).ForeColor = &H0& 'yeşil
ListViewA.SelectedItem.ListSubItems(2).ForeColor = &H0& 'yeşil
ListViewA.SelectedItem.ListSubItems(4).ForeColor = &H0& 'yeşil
End If
If ListViewA.ListItems(Y).SmallIcon = 2 Or ListViewA.ListItems(Y).SmallIcon = 5 Or _
ListViewA.ListItems(Y).SmallIcon = 9 Or ListViewA.ListItems(Y).SmallIcon = 15 Then
ListViewA.SelectedItem.ForeColor = &H7379EC 'kırmızı
ListViewA.SelectedItem.ListSubItems(1).ForeColor = &H7379EC 'kırmızı
ListViewA.SelectedItem.ListSubItems(2).ForeColor = &H7379EC 'kırmızı
End If
ListViewA.SelectedItem.Selected = False
End If
Call toplamFiyat
End Sub
Private Sub ListViewA_Click()
On Error Resume Next
If gir.BackColor = &H7379EC Then
TBADET.Value = ListViewA.SelectedItem.ListSubItems(2).Text
If TBADET.Value Like "*,*" Then SpinButton2.Enabled = False Else SpinButton2.Enabled = True
End If
LBL01 = 1
rsy = Left(ListViewA.SelectedItem.ListSubItems(4), 3): rsy = Trim(rsy)
If Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(11).Index Then
prd = ListViewA.SelectedItem
Call Resimler4
End If
End Sub
Private Sub gir_Click() ' teklif sayfasına veri girişi
On Error Resume Next
z = ListViewA.ListItems.Count
If z <= 0 Then Exit Sub
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
Dim nameo
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
Workbooks(dt).Worksheets("Sayfa1").Activate
'--
Set X = Workbooks(dt).Worksheets("Sayfa1").Columns("B:B").Find("", LookAt:=xlWhole)
n = X.row
If CK1.Value = True Then baslıkM1
If CKB1.Value = True Then baslıkM2
'--
If Cells(2, 2).End(xlDown).Font.ColorIndex = 30 Then fco = 49 Else fco = 30 ' renk seçimleri
Y = 1
Do Until Y > z
If ListViewA.ListItems(Y).SmallIcon = 2 Or ListViewA.ListItems(Y).SmallIcon = 5 Then GoTo git:
If ListViewA.ListItems(Y).SmallIcon = 9 Or ListViewA.ListItems(Y).SmallIcon = 15 Then GoTo git:
'--
Set X = Workbooks(dt).Worksheets("Sayfa1").Columns("B:B").Find("", LookAt:=xlWhole)
n = X.row
m = ListViewA.ListItems(Y).ListSubItems(3).Text
'--
If ListViewA.ListItems(Y).SmallIcon = 14 Then
Range("B" & n) = "."
Range("C" & n) = ListViewA.ListItems(Y).ListSubItems(1).Text
Range("A" & n & ":U" & n).Borders.LineStyle = xlContinuous: Range("W" & n & ":X" & n).Borders.LineStyle = xlContinuous
Range("B" & n & ":E" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("W" & n & ":X" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("A" & n & ":X" & n).Interior.Pattern = xlNone
Range("A" & n & ":X" & n).RowHeight = 12.75: Range("A" & n & ":X" & n).Font.Size = 9
Range("A" & n & ":X" & n).Font.Bold = True: Range("A" & n & ":X" & n).Font.ColorIndex = 47
GoTo git:
End If
Workbooks(dt).Worksheets("Sayfa1").Range("B" & n).EntireRow.Insert
'ssayfa.Range("B" & m).EntireRow.Copy
Dim k As Integer
For k = 1 To 9
Workbooks(dt).Worksheets("Sayfa1").Cells(n, k) = ssayfa.Cells(m, k)
Next k
'Genel Biçimlemeler'--
Range("A" & n & ":U" & n).RowHeight = 12.75
Range("A" & n & ":U" & n).Borders.LineStyle = xlContinuous: Range("W" & n & ":X" & n).Borders.LineStyle = xlContinuous
Range("A" & n & ":X" & n).Interior.Pattern = xlNone
Range("A" & n & ":X" & n).Font.Bold = False
Range("A" & n & ":X" & n).Font.ColorIndex = fco: Range("A" & n & ":X" & n).Font.Size = 9
Range("A" & n & ":D" & n).HorizontalAlignment = xlLeft: Range("E" & n & ":X" & n).HorizontalAlignment = xlRight
Range("A" & n & ":D" & n).NumberFormat = "@"
Range("E" & n).NumberFormat = ssayfa.Range("E" & m).NumberFormat

Range("G" & n).NumberFormat = "0.0%"
Range("H" & n).NumberFormat = "#,##0"
Range("J" & n & ":X" & n).NumberFormat = "#,##0.00"
'KUR'--
Range("F" & n).NumberFormat = ssayfa.Range("F" & m).NumberFormat
kur = ""
If Range("F" & n).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & n).Font.ColorIndex = 3: kur = "*Usd" ' Sayfa3 de $ kuru
If Range("F" & n).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & n).Font.ColorIndex = 5: kur = "*Eur" ' Sayfa3 de € kuru
'ORTAK VERİLER'--
Range("E" & n).FormulaR1C1 = ListViewA.ListItems(Y).ListSubItems(2).Text * 1
If Left(Workbooks(dt).ActiveSheet.CodeName, 3) = "OTM" Then GoTo git:
Range("J" & n).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("K" & n).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
'ÜRÜN GRUPLARI--
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
If Left(Range("A" & n), 5) = "PM-MP" Then Range("L" & n).FormulaR1C1 = bfyt & "*Oisci/100" & mkur: GoTo devam1 'İŞÇ.
If Left(Range("A" & n), 5) = "PM-MS" Then Range("L" & n).FormulaR1C1 = bfyt & "*Osarf/100" & mkur: GoTo devam1 'SARF
If Left(Range("A" & n), 5) = "PM-MA" Then Range("L" & n).FormulaR1C1 = bfyt & "*Oamb/100" & mkur: GoTo devam1 'AMB.
If Left(Range("A" & n), 5) = "PM-MN" Then Range("L" & n).FormulaR1C1 = bfyt & "*Onak/100" & mkur: GoTo devam1 'NAK.
If Left(Range("A" & n), 5) = "PM-MB" Then Range("L" & n).FormulaR1C1 = bfyt & "*Obara/100" & mkur: GoTo devam1 'Bara
If Left(Range("A" & n), 3) = "PP-" Then Range("L" & n).FormulaR1C1 = bfyt & "*Opano/100" & mkur: GoTo devam1 'Pano
If Left(Range("A" & n), 3) = "PS-" Then 'Pano sac & aksesuarlar
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then Range("L" & n).FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else Range("L" & n).FormulaR1C1 = bfyt & "*Opano/100" & mkur
GoTo devam1
End If
Range("L" & n).FormulaR1C1 = bfyt & "*Osalt/100" & mkur 'Mlz.Kar+1
devam1:
'--
Range("M" & n).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("N" & n).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("O" & n).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & n).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & n).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & n).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & n).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("U" & n).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
Range("T" & n).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.
'TOPLAMLAR'--
'Range("W" & n).FormulaR1C1 = "=((RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb)*Tpbs/Tpb" 'Satış kuru ilave
Range("W" & n).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
Range("X" & n).FormulaR1C1 = "=RC[-19]*RC[-1]"
'--
Dim tS3
Set tS3 = Workbooks(dt).Worksheets("Sayfa3")
If tS3.Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & n, "X" & n).NumberFormat = "#,##0.00"
If tS3.Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & n, "X" & n).NumberFormat = "#,##0.00 [$€-1]"
If tS3.Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & n, "X" & n).NumberFormat = "#,##0.00 [$$-C0C]"
'--
git:
Y = Y + 1
Loop
If CK1.Value = True Then Call toplamM1
    'Cells.Replace What:="[*]", Replacement:="", LookAt:= _
        xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'n = Cells(y, 3).End(xlDown).Row
Workbooks(dt).Worksheets("Sayfa1").Range("B" & n + 1).Select
Set tS3 = Nothing: Set X = Nothing
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub girEK_Click() ' teklif sayfasına veri girişi ek
On Error Resume Next
z = LBdl3.ListCount
If z <= 0 Then Exit Sub
m = LBdl3.List(LBdl3.listIndex, 3)
Application.Calculation = xlCalculationManual
Dim ssayfa
Set ssayfa = Workbooks(sl1).ActiveSheet
Workbooks(dt).Worksheets("Sayfa1").Activate
'--
'If Cells(2, 2).End(xlDown).Font.ColorIndex = 30 Then fco = 49 Else fco = 30 ' renk seçimleri
Y = 1
ssayfa.Range("B" & m).EntireRow.Copy
Set X = Workbooks(dt).Worksheets("Sayfa1").Columns("B:B").Find("", LookAt:=xlWhole)
n = X.row
'--
Workbooks(dt).Worksheets("Sayfa1").Range("B" & n).EntireRow.Insert
'Workbooks(dt).Worksheets("Sayfa1").Range("A" & n & ":U" & n & ",W" & n & ":X" & n).Font.ColorIndex = fco
'KUR'--
Range("F" & n).NumberFormat = ssayfa.Range("F" & m).NumberFormat
kur = ""
If Range("F" & n).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & n).Font.ColorIndex = 3: kur = "*Usd" ' Sayfa3 de $ kuru
If Range("F" & n).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & n).Font.ColorIndex = 5: kur = "*Eur" ' Sayfa3 de € kuru
'ORTAK VERİLER'--
Range("E" & n).FormulaR1C1 = CDbl(TBEADET.Value)
Range("J" & n).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("K" & n).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
'ÜRÜN GRUPLARI--
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
If Left(Range("A" & n), 5) = "PM-MP" Then Range("L" & n).FormulaR1C1 = bfyt & "*Oisci/100" & mkur: GoTo devam1 'İŞÇ.
If Left(Range("A" & n), 5) = "PM-MS" Then Range("L" & n).FormulaR1C1 = bfyt & "*Osarf/100" & mkur: GoTo devam1 'SARF
If Left(Range("A" & n), 5) = "PM-MA" Then Range("L" & n).FormulaR1C1 = bfyt & "*Oamb/100" & mkur: GoTo devam1 'AMB.
If Left(Range("A" & n), 5) = "PM-MN" Then Range("L" & n).FormulaR1C1 = bfyt & "*Onak/100" & mkur: GoTo devam1 'NAK.
If Left(Range("A" & n), 5) = "PM-MB" Then Range("L" & n).FormulaR1C1 = bfyt & "*Obara/100" & mkur: GoTo devam1 'Bara
If Left(Range("A" & n), 3) = "PP-" Then Range("L" & n).FormulaR1C1 = bfyt & "*Opano/100" & mkur: GoTo devam1 'Pano
If Left(Range("A" & n), 3) = "PS-" Then 'Pano sac & aksesuarlar
Dim nameo
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then Range("L" & n).FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else Range("L" & n).FormulaR1C1 = bfyt & "*Opano/100" & mkur
GoTo devam1
End If
Range("L" & n).FormulaR1C1 = bfyt & "*Osalt/100" & mkur 'Mlz.Kar+1
devam1:
Range("M" & n).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("N" & n).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("O" & n).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & n).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & n).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & n).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & n).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("U" & n).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
Range("T" & n).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.
'TOPLAMLAR'--
'Range("W" & n).FormulaR1C1 = "=((RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb)*Tpbs/Tpb" 'Satış kuru ilave
Range("W" & n).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
Range("X" & n).FormulaR1C1 = "=RC[-19]*RC[-1]" '--
Dim tS3
Set tS3 = Workbooks(dt).Worksheets("Sayfa3")
If tS3.Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & n, "X" & n).NumberFormat = "#,##0.00"
If tS3.Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & n, "X" & n).NumberFormat = "#,##0.00 [$€-1]"
If tS3.Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & n, "X" & n).NumberFormat = "#,##0.00 [$$-C0C]"
'--
Application.CutCopyMode = False
Workbooks(dt).Worksheets("Sayfa1").Range("B" & n + 1).Select
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub SpinButton3_Change()
On Error Resume Next
TBEADET.Value = SpinButton3.Value
End Sub
Private Sub TBEADET_Change() 'MİKTAR DEĞİŞİMİ TAMAM
On Error Resume Next
SpinButton3.Value = TBEADET.Value
End Sub
Private Sub LBdl3_Click()
If LBdl3.ListCount > 0 Then Frame5.Enabled = True: girEK.Enabled = True
End Sub
Private Sub LBdl3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
z = LBdl3.ListCount
If z <= 0 Then Exit Sub
girEK_Click
End Sub
Private Sub LBdl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call yardmalzeme
On Error Resume Next
z = LB.ListCount
i = LBdl.listIndex
If i < 0 Then Exit Sub
'--
m = 0
Do Until m >= z
If LB.List(m, 1) = LBdl.List(LBdl.listIndex, 0) Then MsgBox ("Bu ürün daha önce eklenmiş!"), vbInformation, "Ekli Ürün": Exit Sub
    m = m + 1
Loop
'--
    LB.AddItem LBdl.List(i, 2)
    LB.List(z, 1) = LBdl.List(i)
    LB.List(z, 2) = LBdl.List(i, 1)
    LB.List(z, 3) = LBdl.List(i, 3)
    LB.List(z, 4) = LBdl.List(i, 4)
    LB.List(z, 5) = LBdl.List(i, 5)
End Sub
Sub baslıkM2() '06.02.2019 başlık ekle
Range("B" & n) = "."
If TextBoxGA1.Tag = 0 Then
Range("C" & n) = ListBoxMG2.Text
Else
Range("C" & n) = ListBoxMG2.Text & TextBoxGA1.Text
End If
Range("E" & n) = "(" & TBADET & " Grup)"
Range("A" & n & ":U" & n).Borders.LineStyle = xlContinuous
Range("W" & n & ":X" & n).Borders.LineStyle = xlContinuous
Range("B" & n & ":E" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("W" & n & ":X" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("A" & n & ":X" & n).Interior.Pattern = xlNone
Range("A" & n & ":X" & n).RowHeight = 12.75
Range("A" & n & ":X" & n).Font.Size = 9
Range("A" & n & ":X" & n).Font.Bold = True
Range("A" & n & ":X" & n).Font.ColorIndex = 47
n = n + 1
Range("B" & n).Select
Workbooks(dt).Worksheets("Sayfa1").Range("B" & n).EntireRow.Insert
End Sub
Sub baslıkM1() '06.02.2019 başlık ekle
'Biçimlemeler'--
Range("B" & n) = "BÖLÜM ADI/NO:"
If TextBoxGA1.Tag = 0 Then
Range("C" & n) = ListBoxMG2.Text
Else
Range("C" & n) = "BÖLÜM ADI/NO:" & TextBoxGA1.Text & " - (" & TBADET & " Grup)"
End If
Range("A" & n & ":U" & n).Borders.LineStyle = xlContinuous
Range("W" & n & ":X" & n).Borders.LineStyle = xlContinuous
Range("B" & n & ":E" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("W" & n & ":X" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("A" & n & ":X" & n).Interior.Pattern = xlNone
Range("A" & n & ":X" & n).RowHeight = 12.75
Range("A" & n & ":X" & n).Font.Size = 9
Range("A" & n & ":X" & n).Font.Bold = True
Range("A" & n & ":X" & n).Font.ColorIndex = 11
n = n + 1
Range("B" & n).Select
Workbooks(dt).Worksheets("Sayfa1").Range("B" & n).EntireRow.Insert
End Sub
Sub toplamM1() '06.02.2019 toplam ekle
n = n + 1
Range("B" & n) = "BÖLÜM TOPLAMI:"
'Biçimlemeler'--
Range("A" & n & ":U" & n).Borders.LineStyle = xlContinuous
Range("W" & n & ":X" & n).Borders.LineStyle = xlContinuous
Range("B" & n & ":E" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("W" & n & ":X" & n).Borders(xlInsideVertical).LineStyle = xlNone
Range("A" & n & ":X" & n).Interior.Pattern = xlNone
Range("A" & n & ":X" & n).RowHeight = 12.75
Range("A" & n & ":X" & n).Font.Size = 9
Range("A" & n & ":X" & n).Font.Bold = True
Range("A" & n & ":X" & n).Font.ColorIndex = 11
Call AraToplamlar
End Sub