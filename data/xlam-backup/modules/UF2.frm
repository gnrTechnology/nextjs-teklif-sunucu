Option Explicit '2007 icin İLHAN
Dim msayfa
Dim rs As String, rs1 As Byte
Dim ug1 As String
Dim kur, mkur As String
Public bfyt As String
Dim z, f, df, ds
Dim e As Integer, zy As Integer, ss As Integer, zx As Integer
Dim X As Integer
Dim Y As Integer
'Dim c_hForm
Dim il As Integer
Dim LXM As Integer
Dim P31L
Private kopya As New DataObject
Sub MakePopUp()
'Remove any old instance of MyPopUp
On Error Resume Next
CommandBars("MyPopUp").Delete
On Error GoTo 0
With CommandBars.Add(Name:="MyPopUp", Position:=msoBarPopup)
.Controls.Add(Type:=msoControlButton, Id:=19).OnAction = "Textbox_Copy2"
End With
End Sub
Sub MakePopUp1()
On Error Resume Next
CommandBars("MyPopUp1").Delete
On Error GoTo 0
With CommandBars.Add(Name:="MyPopUp1", Position:=msoBarPopup)
.Controls.Add(Type:=msoControlButton, Id:=22).OnAction = "Textbox_Paste"
End With
End Sub
Private Sub Label62_Click()
If TextBox2.Text = "" Then Exit Sub
Dim Link
Link = "https://www.google.com/search?q=" & TextBox2.Value & "+" & TextBox4.Value & "&source=lnms&tbm=isch"
On Error GoTo NoCanDo
ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
'tam ekran olarak açmak Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & Link, vbMaximizedFocus
'--
 TextBox2.SetFocus
 TextBox2.SelStart = 0
 TextBox2.SelLength = TextBox2.TextLength
 TextBox2.Copy
'--
Exit Sub
NoCanDo:
MsgBox Link & " açılamıyor. İnternet bağlantınızı kontrol ediniz.", vbInformation
End Sub
Private Sub Label72_Click()
Application.ScreenUpdating = False
If TextBox2.Text = "" Then Exit Sub
Dim yolm As String
Dim d1m
d1m = GetSetting("ilhan", "Settings", "deposabitdosya")
yolm = GetSetting("ilhan", "Settings", "depodizini") & "\"
Workbooks.Open (yolm & d1m): Application.Windows(ActiveWorkbook.Name).Visible = False
Dim tms
tms = WorksheetFunction.SumIf(Workbooks(d1m).ActiveSheet.Range("B:B"), TextBox2, Workbooks(d1m).ActiveSheet.Range("E:E"))
Workbooks(d1m).Close (False)
MsgBox "Bakılan Dosya Adı : " & d1m & vbLf & "Bakılan Ürün : " & TextBox2 & vbLf & "Mevcut Stok Ad. : " & tms
Application.ScreenUpdating = True
End Sub
Private Sub Label63_Clickxxx()
If TextBox2.Text = "" Then Exit Sub
Dim Link
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
If TextBox4 = "ABB" Then Link = "https://new.abb.com/products/" & TextBox2.Value: GoTo git:
If TextBox4 = "EATON" Then Link = "http://datasheet.moeller.net/datasheet.php?model=" & msayfa.Range("L" & z) & "&locale=en_GB&_lt=": GoTo git:
If TextBox4 = "SIEMENS" Then Link = "https://mall.industry.siemens.com/mall/en/tr/Catalog/Product/" & TextBox2.Value: GoTo git:
If TextBox4 = "SCHNEIDER" Then Link = "https://www.schneider-electric.com.tr/tr/product/" & TextBox2.Value: GoTo git:
If TextBox4 = "PHOENIX" Then Link = "https://www.phoenixcontact.com/online/portal/tr?uri=pxc-oc-itemdetail:pid=" & TextBox2.Value: GoTo git:
git:
On Error GoTo NoCanDo
ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
'tam ekran olarak açmak Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & Link, vbMaximizedFocus
 TextBox2.SetFocus
 TextBox2.SelStart = 0
 TextBox2.SelLength = TextBox2.TextLength
 TextBox2.Copy
'--
Exit Sub
NoCanDo:
MsgBox Link & " Bağlantı hatası", vbInformation
'MsgBox Link & " açılamıyor. İnternet bağlantınızı kontrol ediniz.", vbInformation
End Sub
Private Sub Label63_Click()
If TextBox2 = "" Then Exit Sub
If ListBoxMK.ListCount = 0 Then Call malzemelink
'TextBoxMK1 = ""
Dim i As Byte
For i = 0 To ListBoxMK.ListCount - 1
If ListBoxMK.List(i, 0) = TextBox4 Then TextBoxMK1 = ListBoxMK.List(i, 1) & TextBox2 & ListBoxMK.List(i, 2): Exit For
Next
If TextBoxMK1 = "" Then MsgBox " Link bağlantı dosyasını kontrol ediniz.", vbInformation: Exit Sub
Dim Link
Link = TextBoxMK1
On Error GoTo NoCanDo
ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
Exit Sub
NoCanDo:
MsgBox Link & " açılamıyor. İnternet bağlantınızı kontrol ediniz.", vbInformation
End Sub
Sub malzemelink()
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, satır2 As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Malzeme Linkleri\Malzeme Linkleri.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
    MsgBox "Malzeme Linkleri.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxMK.Clear
    Do While Not EOF(Ert)
        Line Input #Ert, Rky
        If Left(Rky, 1) = "#" Then GoTo git1
        ayır = Split(Rky, ";")
        ListBoxMK.AddItem ayır(i)
'--
'kaçıncı = InStr(1, Rky, ";")
Dim tsay As Byte, n As Byte
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxMK.List(satır - 1, n) = ayır(i + n)
Next n
        satır = satır + 1
git1:
    Loop
    Close #Ert
End Sub
Private Sub TextBox2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBoxPPT.Text = TextBox2.Text: If TextBoxPPT.Text = "" Then Exit Sub
MakePopUp
If Button = 2 Then
Application.CommandBars("MyPopUp").ShowPopup
End If
End Sub
Private Sub TextBox8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
TextBoxPPT.Text = TextBox8.Text: If TextBoxPPT.Text = "" Then Exit Sub
MakePopUp
If Button = 2 Then
Application.CommandBars("MyPopUp").ShowPopup
End If
End Sub
Private Sub TextBox22_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
MakePopUp1
If Button = 2 Then
Application.CommandBars("MyPopUp1").ShowPopup
End If
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
UF2.Height = 450
If Not dtx = 0 Then MultiPage1M.Value = 0: GoTo gitd1
If MultiPage1.Value = 0 Then
MultiPage1M.Value = 0
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\Logo.jpg")
MultiPage2.Value = 0
Else
If MultiPage1M.Value <> 0 Then MultiPage1M.Value = 0: GoTo gitd1 Else MultiPage1.Value = 0
End If
Label18.Caption = "    Yapılacak İşin Cinsi": Label18.BackColor = &H80000002: Label18.ForeColor = &HFFFFFF: Label18.Font.Bold = True
gitd1:
If LXM >= 0 Then ListBox1.Selected(LXM) = True
Case 2
UF2.Height = 450
If Not dtx = 0 Then GoTo gitd2
If MultiPage1M.Value <> 1 Then GoTo gitd2 Else MultiPage1.Value = 0
gitd2:
MultiPage1M.Value = 1
TextBox1.Text = "": TextBox2.Text = "": TextBox4.Text = "": TextBox6.Text = ""
TextBox7.Text = "": TextBox9.Text = ""
CommandButton1.Enabled = False: CommandButton3.Enabled = False: CommandButton8.Enabled = False
SpinButton1.Enabled = False: 'ScrollBarP21.Enabled = False
TextBox5.Enabled = False: ListBox1.Selected(ListBox1.listIndex) = False
'ListBoxA.Selected(ListBoxA.ListIndex) = False
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\Logo.jpg")
MultiPage2.Value = 0
ListBox21.Selected(ListBox21.listIndex) = False
Case 3
UF2.Height = 450
ListViewP31.SelectedItem.EnsureVisible: ListViewP31.Width = 643
MultiPage2.Value = 1
Case 4
If UF2.Height > 400 Then
UF2.Height = UF2.Height - UF2.InsideHeight + Toolbar2.Height: Toolbar2.Buttons(5).Value = 1
Else
UF2.Height = 450
End If
Case 5
If Not dtx = 0 Then UF2.Hide Else Unload Me
End Select
End Sub
Private Sub Toolbar2_DblClick()
If UF2.Height > 400 Then
UF2.Height = UF2.Height - UF2.InsideHeight + Toolbar2.Height: Toolbar2.Buttons(5).Value = 1
Else
UF2.Height = 450
End If
End Sub
Private Sub UserForm_Initialize() '2007 icin TAMAM İLHAN
On Error Resume Next
'--
il = 1: If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then il = 0
'--
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
If il = 0 Then
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) <> "TM" Then CommandButton1.Visible = False: Toolbar2.Buttons(3).Enabled = False: _
Toolbar2.Buttons(4).Enabled = False
Else: dt = ActiveWorkbook.Name
End If
'--
If mlz = Empty Then End
'Application.ScreenUpdating = False:Application.EnableEvents = False
Toolbar2.ImageList = ImageList2
Toolbar2.Buttons.Item(1).Image = ImageList2.ListImages.Item(1).Index
Toolbar2.Buttons.Item(2).Image = ImageList2.ListImages.Item(2).Index
Toolbar2.Buttons.Item(3).Image = ImageList2.ListImages.Item(3).Index
Toolbar2.Buttons.Item(4).Image = ImageList2.ListImages.Item(4).Index
Toolbar2.Buttons.Item(5).Image = ImageList2.ListImages.Item(5).Index
Caption = mlz
ss = Workbooks(mlz).Worksheets("Sayfa1").Range("B65536").End(xlUp).row
'c_hForm = FindWindow(vbNullString, Me.Caption)
''SetWindowLong c_hForm, -16, &H20000 Or &H10000 Or &H84C80080 'simgeler komple
'SetWindowLong c_hForm, -16, &H20000 Or &H84C80080
rs = Trim(Left(Replace(mlz, "+", ""), 3))
'rs = Trim(Left(mlz, 3))
MultiPage1.Value = 0
Call form
'--
If Not dtx = 0 Then
Toolbar2.Buttons(3).Visible = False
Label18.Caption = "  " & UFMZ.TextBox1.Value & " " & UFMZ.TextBox2.Value & " (" & UFMZ.TextBox3.Value & ")"
Label18.BackColor = &H80C0FF: Label18.ControlTipText = ""
If UFMZ.MultiPage1.Value = 1 Then MultiPage1.Value = 3 Else MultiPage1.Value = 1
End If
'--
'TextBoxkur1 = Sheets("Sayfa3").Range("Eur") & vbCr & Sheets("Sayfa3").Range("Usd")
Application.ScreenUpdating = True: Application.EnableEvents = True
''If UFDB.C3.Value = True Then UF2.Caption = mlz: mlz = ""
End Sub
Sub form() '19.03.2016 REV1
On Error Resume Next
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
If msayfa.Range("A1") = "ANA ŞALT MALZEME" Then ug1 = "+" Else ug1 = "" '**ürün grubu
Dim k, t, a
k = ""
e = msayfa.Range("B65536").End(xlUp).row + 1
Y = 1
z = 0
tekrar:
t = "B" & Y & ":" & "B" & e
Set a = msayfa.Range(t).Find(k, LookIn:=xlValues, LookAt:=xlPart)
Y = a.row
If Y >= e Then GoTo son:
If Not a Is Nothing Then
ListBoxA.AddItem msayfa.Cells(Y, 1) 'ListBoxA için
ListBoxA.List(z, 1) = msayfa.Cells(Y, 1).row 'ListBoxA için
z = z + 1
If z > 10000 Then MsgBox ("Kodları kontrol et!  "), vbCritical, "Uyarı": Exit Sub
GoTo tekrar

End If
son:
    Windows(dt).Activate
    Sheets("Sayfa1").Select '
    Range("A2").Select
    TextBox10.Text = Sheets("Sayfa1").Range("D65536").End(xlUp).row - 1
Set z = Columns("B:B").Find("", LookAt:=xlWhole)
    TextBox19.Text = z.row
    TextBox5.Text = ""
    SpinButton1.Enabled = False
    ListBoxB.Visible = False
    ListBox1.Visible = False
    TextBox5.Enabled = False
    CommandButton1.Enabled = False
    CommandButton3.Enabled = False
    CommandButton8.Enabled = False
    TextBox8.Value = ""
    Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\Logo.jpg")
End Sub
Private Sub ListBoxA_Click() '19.03.2016  REV1
'Application.ScreenUpdating = False: Application.EnableEvents = False
On Error Resume Next
LXM = -1: TextBox1.Text = "": TextBox2.Text = "": TextBox6.Text = "": TextBox9.Text = ""
TextBox5.Enabled = False
ListBox1.Visible = False
ListBoxB.Visible = True
ListBox1.Clear
CommandButton1.Enabled = False
DT1.Enabled = False
CommandButton3.Enabled = False
CommandButton8.Enabled = False
ListBoxB.Clear
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
rs = Trim(Left(Replace(mlz, "+", ""), 3))
'rs = Trim(Left(mlz, 3))
'Windows(mlz).Activate ' Ana tablosu seçtim
Dim t As Integer
t = ListBoxA.listIndex
z = 0
If Not t = ListBoxA.ListCount - 1 Then X = ListBoxA.List(t + 1, 1) Else X = ss
Y = ListBoxA.List(t, 1) + 1
TextBox8.Text = Replace(msayfa.Cells(ListBoxA.List(t, 1), 1), "/", " ") 'RESİM SEÇİMLERİ
Do Until Y > X - 1
     If msayfa.Cells(Y, 1) <> "" Then
        ListBoxB.AddItem msayfa.Cells(Y, 1)
        ListBoxB.List(z, 1) = msayfa.Cells(Y, 2).row
        z = z + 1
     End If
        Y = Y + 1
Loop

Windows(dt).Activate
'Application.EnableEvents = True: Application.ScreenUpdating = True
ListBoxA.Selected(t) = True
LabelB2.Caption = "   " & TextBox8.Text
End Sub
Private Sub ListBoxB_Click() '19.03.2016  REV1
On Error Resume Next
LXM = -1: TextBox2.Text = "": TextBox6.Text = "": TextBox9.Text = ""
TextBox5.Enabled = False
TextBox5.Value = 1
SpinButton1.Value = 1
CommandButton1.Enabled = False
DT1.Enabled = False
CommandButton3.Enabled = False
CommandButton8.Enabled = False
SpinButton1.Enabled = False
ListBox1.Visible = True
ListBox1.Clear
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
rs = Trim(Left(Replace(mlz, "+", ""), 3))
'rs = Trim(Left(mlz, 3))
Dim t
t = ListBoxB.listIndex
z = 0
Y = ListBoxB.List(t, 1)
TextBox8.Text = Replace(msayfa.Cells(ListBoxB.List(t, 1), 1), "/", " ") 'RESİM SEÇİMLERİ
If msayfa.Cells(Y + 1, 1) = "" Then X = msayfa.Cells(Y, 1).End(xlDown).row Else X = msayfa.Cells(Y + 1, 1).row
If Y >= X Then X = msayfa.Cells(Y, 2).End(xlDown).row + 1
'X = ListBoxB.List(t + 1, 1) + 0
'If Y >= X Then X = msayfa.Cells(Y, 2).End(xlDown).Row + 1
TextBox18.Text = Y + 1
If msayfa.Range("K" & Y) <> "" Then TextBox1.Text = msayfa.Range("K" & Y) Else TextBox1.Text = "XX" 'yeni liste için 2016
Do Until Y > X - 1
If msayfa.Cells(Y, 3) <> "" And msayfa.Cells(Y, 2) <> "" Then
ListBox1.AddItem msayfa.Cells(Y, 2)
ListBox1.List(z, 1) = msayfa.Cells(Y, 3)
ListBox1.List(z, 2) = msayfa.Cells(Y, 6).Text
z = z + 1
End If
Y = Y + 1
Loop
AlignListColumn ListBox1, 2, True
ListBox1.SetFocus
Windows(dt).Activate
End Sub
Private Sub XXXListBoxB_Click() 'silmeeeeeee
'Application.ScreenUpdating = False
On Error Resume Next
TextBox5.Enabled = False
TextBox5.Value = 1
SpinButton1.Value = 1
CommandButton1.Enabled = False
DT1.Enabled = False
CommandButton3.Enabled = False
CommandButton8.Enabled = False
SpinButton1.Enabled = False
ListBox1.Visible = True
ListBox1.Clear
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
'Windows(mlz).Activate
Dim t
t = ListBoxB.listIndex
z = 0
Y = ListBoxB.List(t, 1)
X = ListBoxB.List(t + 1, 1) + 0
If Y >= X Then X = msayfa.Cells(Y, 2).End(xlDown).row + 1
TextBox18.Text = Y + 1
TextBox8.Text = Replace(msayfa.Cells(ListBoxB.List(t, 1), 3), "/", " ") 'RESİM SEÇİMLERİ
Do Until Y > X
If msayfa.Cells(Y, 4) <> "" And msayfa.Cells(Y, 2) <> "" Then
ListBox1.AddItem msayfa.Cells(Y, 2)
ListBox1.List(z, 1) = msayfa.Cells(Y, 3)
ListBox1.List(z, 2) = msayfa.Cells(Y, 6).Text
z = z + 1
End If
Y = Y + 1
Loop
AlignListColumn ListBox1, 2, True
ListBox1.SetFocus
Windows(dt).Activate
'Application.ScreenUpdating = True
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
Dim abi
Do While labTester.Width <= sngWidth
labTester.Caption = " " & labTester.Caption
Loop
.List(intItem, WhichColumn) = labTester.Caption
Next
End With
End Sub
Private Sub ListBox1_Click()
On Error Resume Next
Windows(dt).Activate
Sheets("Sayfa1").Select '
rs = Trim(Left(Replace(mlz, "+", ""), 3))
TextBox6.Value = ""
TextBox9.Value = ""
TextBox5.Enabled = True
CommandButton1.Enabled = True
DT1.Enabled = True
CommandButton3.Enabled = True
CommandButton8.Enabled = True
SpinButton1.Enabled = True
X = ListBox1.listIndex
LXM = ListBox1.listIndex
'z = TextBox18.Text + x 'eski liste için
z = TextBox18.Text + X - 1
TextBox17.Text = z
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
'TextBox1.Text = msayfa.Range("A" & z)'eski liste için
If msayfa.Range("K" & z) = "" Then TextBox1.Text = msayfa.Range("K" & z + 1).End(3) Else TextBox1.Text = msayfa.Range("K" & z)
'If TextBox1.Text = "" Then TextBox1.Text = msayfa.Range("K" & z + 1).End(3)
TextBox2.Text = msayfa.Range("B" & z)
TextBox4.Text = msayfa.Range("D" & z)
rs = Trim(Left(Replace(TextBox4.Text, "+", ""), 3))
'rs = TextBox4.Text
TextBox6.Value = msayfa.Range("F" & z).Text 'birim fiyat
TextBox7.Value = Format((msayfa.Range("G" & z) * 100), "#,##0.0") 'iskonto
TextBox8.Text = Replace(msayfa.Range("B" & z), "/", " ") 'RESİM SEÇİMLERİ
ScrollBarP21.Value = TextBox7.Value * 2 'iskonto
'--
Dim fnet
fnet = (msayfa.Range("F" & z).Value - ((msayfa.Range("F" & z).Value * TextBox7.Text) / 100)) * TextBox5.Value 'net fiyat
TextBox9.ControlTipText = ""
TextBox9.Value = Format(fnet * TextBox5.Value, "#,##0.00") 'net fiyat
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$€-1]" Then TextBox9.Value = TextBox9.Value & " €": TextBox9.ControlTipText = Format(fnet * Sheets("Sayfa3").Range("Eur"), "#,##0.00") & " TL"
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$$-C0C]" Then TextBox9.Value = TextBox9.Value & " $": TextBox9.ControlTipText = Format(fnet * Sheets("Sayfa3").Range("Usd"), "#,##0.00") & " TL"
'--
If ListBox1.Text = "•" Or ListBox1.Text = "BAŞLIK" Then Frame1.Enabled = False: Exit Sub Else Frame1.Enabled = True
'--
CommandButton1.BackColor = &H86B57D
'If TextBox4.Text = "SIEMENS" Or TextBox4.Text = "EATON" Or TextBox4.Text = "ABB" Or TextBox4.Text = "SCHNEIDER" Or TextBox4.Text = "PHOENIX" Then _
Label63.Visible = True Else Label63.Visible = False
'--
End Sub
Private Sub TextBox8_Change() 'RESİM DEĞİŞİMİ TAMAM+
On Error GoTo hata
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & ".jpg")
If a = True Then Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & ".jpg"): GoTo git1
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rs & "\Logo.jpg")
If a = True Then Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\Logo.jpg"): Exit Sub
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg"): Exit Sub
git1:
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & "_1" & ".jpg")
If a = True Then LBR1.Visible = True: LBR2.Visible = True Else LBR1.Visible = False: LBR2.Visible = False
hata:
End Sub
Private Sub LBR1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
LBR1.Tag = LBR1.Tag + 1: Call rsileri
End Sub
Private Sub Frame3_Click()
'LBR1.Tag = LBR1.Tag + 1: Call rsileri
UF2P.Show
End Sub
Private Sub LBR2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
LBR1.Tag = LBR1.Tag - 1: Call rsileri
End Sub
Sub rsileri()
On Error Resume Next
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & "_" & LBR1.Tag & ".jpg")
If a = True Then
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & "_" & LBR1.Tag & ".jpg"): Exit Sub
Else
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & ".jpg"): LBR1.Tag = 0
Image1.Height = Image1.Height + 1: Image1.Height = Image1.Height - 1: Exit Sub
End If
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
End Sub
Private Sub TextBox8_Changexxx() 'RESİM DEĞİŞİMİ TAMAM
On Error GoTo hata
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & ".jpg")
Exit Sub
hata:
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
End Sub
Private Sub ScrollBarP21_Change()
On Error Resume Next
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
If TextBox7.Text = ScrollBarP21.Value / 2 Then Exit Sub
TextBox7.Text = Format(ScrollBarP21.Value * 0.5, "#,##0.0")
'--
Dim fnet
fnet = (msayfa.Range("F" & z).Value - ((msayfa.Range("F" & z).Value * TextBox7.Text) / 100)) * TextBox5.Value 'net fiyat
TextBox9.ControlTipText = ""
TextBox9.Value = Format(fnet * TextBox5.Value, "#,##0.00") 'net fiyat
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$€-1]" Then
TextBox9.Value = TextBox9.Value & " €": TextBox9.ControlTipText = Format(fnet * Sheets("Sayfa3").Range("Eur"), "#,##0.00") & " TL"
TextBox9.ForeColor = 2
ScrollBarP21.Value = TextBox7.Text * 2
Exit Sub
End If
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$$-C0C]" Then
TextBox9.Value = TextBox9.Value & " $": TextBox9.ControlTipText = Format(fnet * Sheets("Sayfa3").Range("Usd"), "#,##0.00") & " TL"
TextBox9.ForeColor = 2
ScrollBarP21.Value = TextBox7.Text * 2
End If
'--
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'TAMAM
On Error Resume Next
If ListBox1.Text = "•" Or ListBox1.Text = "BAŞLIK" Then Exit Sub
If MultiPage1.Value = 1 Then Call CommandButton3_Click: GoTo son1
'--
If il = 0 Then
 If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call verigir2 Else Call msgteklif2: GoTo son1
Else
If MultiPage1.Value = 0 Then Call verigir
If Left(ActiveSheet.CodeName, 3) <> "OTM" Then Call ToplamT2
If MultiPage1.Value = 2 Then Call DT1_Click
If MultiPage1.Value = 3 Then Call CommandButton8_Click
End If
son1:
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error Resume Next
If ListBox1.Text = "" Then Exit Sub
If KeyAscii = 13 Then
If ListBox1.Text = "•" Or ListBox1.Text = "BAŞLIK" Then Exit Sub
If il = 0 Then
 If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call verigir2 Else Call msgteklif2: GoTo son1
Else
If MultiPage1.Value = 0 Then Call verigir
If Left(ActiveSheet.CodeName, 3) <> "OTM" Then Call ToplamT2
If MultiPage1.Value = 1 Then Call CommandButton3_Click
'If MultiPage1.Value = 2 Then Call DT1_Click
End If
End If
son1:
Application.Calculation = xlCalculationAutomatic
End Sub
Sub msgteklif2()
Dim msg
MsgBox ("    Teklif dosyası açınız !  "), vbInformation, "scngnr@gmail.com"
End Sub
Private Sub CommandButton1_Click() 'TAMAM
On Error Resume Next
Windows(dt).Activate
Sheets("Sayfa1").Select '
If il = 0 Then
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call verigir2: GoTo son1
Else
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
Call verigir
End If
son1:
Toolbar2.Buttons(3).Caption = "Malzeme Değişimi"
Application.Calculation = xlCalculationAutomatic
End Sub
Sub verigir() ' teklif sayfasına veri girişi
On Error Resume Next
Application.Calculation = xlCalculationManual
'Application.EnableEvents = False
Dim mliste, mlisteS3, Y
Dim nameo
Set mliste = Workbooks(mlz).Worksheets("Sayfa1")
If MultiPage1.Value = 1 Then Y = TextBoxT3.Value: GoTo ekle1
If MultiPage1.Value = 2 Then Y = TextBoxT3.Value: GoTo ekle1
Dim a, son
If Cells(2, 2) = "" Then Y = 2: GoTo ekle1
Y = Cells(1, 2).End(xlDown).row + 1
If Y > 10000 Then Y = 2
ekle1:
'Genel Biçimlemeler'--
Range("A" & Y & ":U" & Y).Borders.LineStyle = xlContinuous
Range("W" & Y & ":X" & Y).Borders.LineStyle = xlContinuous
Range("A" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Font.Bold = False
Range("A" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Font.ColorIndex = xlAutomatic
Range("A" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Font.Size = 9
Range("A" & Y & ":X" & Y).Font.Name = "Arial"
Range("A" & Y & ":D" & Y).HorizontalAlignment = xlLeft
Range("E" & Y & ":U" & Y & ",W" & Y & ":X" & Y).HorizontalAlignment = xlRight
Range("A" & Y & ":D" & Y).NumberFormat = "@"
Range("E" & Y).NumberFormat = "#,##0"
Range("F" & Y).NumberFormat = "#,##0.00"
Range("G" & Y).NumberFormat = "0.0%"
Range("H" & Y).NumberFormat = "#,##0"
Range("J" & Y & ":X" & Y).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & Y & ",X" & Y).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & Y & ",X" & Y).NumberFormat = "#,##0.00 [$€-1]"
If Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & Y & ",X" & Y).NumberFormat = "#,##0.00 [$$-C0C]"
'GENEL VERİLER'--
z = CDbl(TextBox17.Text)
Range("A" & Y) = TextBox1.Text & ug1 '**ürün grubu
Range("B" & Y) = TextBox2.Text 'Sipariş Kd.
Range("C" & Y) = mliste.Range("C" & z) 'Yapılacak İşin Cinsi
Range("D" & Y) = TextBox4.Text 'Üretici

If MultiPage1.Value = 2 Then GoTo zıpla:
If MultiPage1.Value = 0 Then Range("E" & Y) = TextBox5.Value 'Miktar
zıpla:

Range("F" & Y) = mliste.Range("F" & z) 'Mlz. Br. Fiyat
Range("F" & Y).NumberFormat = mliste.Range("F" & z).NumberFormat
Range("G" & Y) = mliste.Range("G" & z) 'mlz.isk.
Range("H" & Y) = mliste.Range("H" & z) 'Adam/dk

Range("I" & Y) = mliste.Range("I" & z) 'Boyut
If Left(ActiveSheet.CodeName, 3) = "OTM" Then GoTo zıpla1:
'KUR'--
kur = ""
If Range("F" & Y).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & Y).Font.ColorIndex = 3: kur = "*Usd" ' Sayfa3 de $ kuru
If Range("F" & Y).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & Y).Font.ColorIndex = 5: kur = "*Eur" ' Sayfa3 de € kuru
'ÜRÜN GRUPLARI--
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
If Left(Range("A" & Y), 5) = "PM-MP" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Oisci/100" & mkur: GoTo devam1 'İŞÇ.
If Left(Range("A" & Y), 5) = "PM-MS" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Osarf/100" & mkur: GoTo devam1 'SARF
If Left(Range("A" & Y), 5) = "PM-MA" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Oamb/100" & mkur: GoTo devam1 'AMB.
If Left(Range("A" & Y), 5) = "PM-MN" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Onak/100" & mkur: GoTo devam1 'NAK.
If Left(Range("A" & Y), 5) = "PM-MB" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Obara/100" & mkur: GoTo devam1 'Bara
If Left(Range("A" & Y), 3) = "PP-" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Opano/100" & mkur: GoTo devam1 'Pano
If Left(Range("A" & Y), 3) = "PS-" Then 'Pano sac & aksesuarlar
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then Range("L" & Y).FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else Range("L" & Y).FormulaR1C1 = bfyt & "*Opano/100" & mkur
GoTo devam1
End If
Range("L" & Y).FormulaR1C1 = bfyt & "*Osalt/100" & mkur 'Mlz.Kar+1
devam1:
'--
Range("J" & Y).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("N" & Y).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("K" & Y).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
Range("M" & Y).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("O" & Y).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & Y).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & Y).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & Y).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & Y).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("T" & Y).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.* işçilik katsayılı+1
Range("U" & Y).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
'TOPLAMLAR'--
git:
'Range("W" & Y).FormulaR1C1 = "=((RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb)*Tpbs/Tpb" 'Döviz satış kuru ilave
Range("W" & Y).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb" 'Dövize göre Birim Fiyat TL+1
Range("X" & Y).FormulaR1C1 = "=RC[-19]*RC[-1]"  'Toplam Fiyat TL+1
'--
zıpla1:
TextBox10.Text = Workbooks(dt).Sheets("Sayfa1").Range("D65536").End(xlUp).row - 1
z = Y + 1
    TextBox19.Text = z.row
CommandButton1.BackColor = &H7379EC
Range("B" & Y).Select
'Application.EnableEvents = False
End Sub
Sub verigir2() ' malzeme sayfasına veri girişi
On Error Resume Next
Application.Calculation = xlCalculationManual
'Application.EnableEvents = False
Dim mliste, mlisteS3, Y
Dim nameo
Set mliste = Workbooks(mlz).Worksheets("Sayfa1")
If MultiPage1.Value = 1 Then Y = TextBoxT3.Value: GoTo ekle1
If MultiPage1.Value = 2 Then Y = TextBoxT3.Value: GoTo ekle1
Dim a, son
If Cells(2, 2) = "" Then Y = 2: GoTo ekle1
Y = Cells(1, 2).End(xlDown).row + 1
If Y > 10000 Then Y = 2
ekle1:
'Genel Biçimlemeler'--
Range("A" & Y & ":J" & Y).Borders.LineStyle = xlContinuous
Range("A" & Y & ":J" & Y).Font.Bold = False
Range("A" & Y & ":J" & Y).Font.ColorIndex = xlAutomatic
Range("A" & Y & ":J" & Y).Font.Size = 9
Range("A" & Y & ":X" & Y).Font.Name = "Arial"
Range("A" & Y & ":D" & Y).HorizontalAlignment = xlLeft
Range("E" & Y & ":J" & Y).HorizontalAlignment = xlRight
Range("A" & Y & ":D" & Y).NumberFormat = "@"
Range("E" & Y).NumberFormat = "#,##0"
Range("F" & Y & ":J" & Y).NumberFormat = "#,##0.00"
Range("H" & Y).NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
Range("I" & Y).NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
'GENEL VERİLER'--
z = CDbl(TextBox17.Text)
Range("A" & Y) = TextBox1.Text 'Referans
Range("B" & Y) = TextBox2.Text 'Sipariş Kd.
Range("C" & Y) = mliste.Range("C" & z) 'Yapılacak İşin Cinsi
Range("D" & Y) = TextBox4.Text 'Üretici
If MultiPage1.Value = 2 Then GoTo zıpla:
If MultiPage1.Value = 0 Then Range("E" & Y) = TextBox5.Value 'Miktar
zıpla:
Range("F" & Y).FormulaR1C1 = "": Range("G" & Y).FormulaR1C1 = "": Range("J" & Y).FormulaR1C1 = ""
Range("H" & Y).FormulaR1C1 = "=IF(RC[-2]-RC[-1]<=0,""-"",RC[-2]-RC[-1])"
Range("I" & Y).FormulaR1C1 = "=IF(RC[-2]-RC[-3]<=0,""-"",RC[-2]-RC[-3])"
'--
TextBox10.Text = Workbooks(dt).Sheets("Sayfa1").Range("D65536").End(xlUp).row - 1
z = Y + 1
    TextBox19.Text = z.row
CommandButton1.BackColor = &H7379EC
Range("B" & Y).Select
'Application.EnableEvents = False
End Sub
Sub ToplamT2() 'TOPLAM ALMA TAMAM
On Error Resume Next
Sheets("Sayfa1").Select
Dim hcr As Range, eskihcr As Range
On Error GoTo hata
Set hcr = Columns("B:B").Find("BÖLÜM TOPLAMI:", LookAt:=xlWhole)
  Cells(hcr.row, "X") = "=Sum(X2:X" & hcr.row - 1 & ")"
Set eskihcr = hcr.Offset(2, 0)
Do
Set hcr = Range(hcr.Offset(1, 0), [B65000]).FindNext
  Cells(hcr.row, "X") = "=Sum(X" & eskihcr.row - 1 & ":X" & hcr.row - 1 & ")"
Set eskihcr = hcr.Offset(2, 0)
Loop
Dim gtp As Range
Set gtp = Columns("B:B").Find("GENEL TOPLAM:", LookAt:=xlWhole)
  Cells(gtp.row, "X") = "=SUM(R2C24:R[-1]C)-SUMIF(R2C2:R[-1]C[-22],""=BÖLÜM TOPLAMI:"",R2C24:R[-1]C)"
hata:
End Sub
Private Sub TextBox5_Change() 'MİKTAR DEĞİŞİMİ TAMAM
On Error Resume Next
SpinButton1.Value = CDbl(TextBox5.Text)
End Sub
Private Sub Label61_Click()
On Error Resume Next
If dir("C:\Belgelerim\CEMEX\Resimler\" & rs & "\", vbDirectory) = "" Then _
MkDir "C:\Belgelerim\CEMEX\Resimler\" & rs & "\"
Shell "C:\Windows\Explorer.exe C:\Belgelerim\CEMEX\Resimler\" & rs & "\", vbNormalFocus
End Sub
Private Sub SpinButton1_Change() 'MİKTAR DEĞİŞİMİ TAMAM
On Error Resume Next
z = CDbl(TextBox17.Text)
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
TextBox5.Text = SpinButton1.Value
'--
TextBox9.Value = Format((msayfa.Range("F" & z).Value - ((msayfa.Range("F" & z).Value * TextBox7.Text) / 100)) * TextBox5.Value, "#,##0.00") 'net fiyat
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$€-1]" Then TextBox9.Value = TextBox9.Value & " €"   '€ kuru
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$$-C0C]" Then TextBox9.Value = TextBox9.Value & " $" '$ kuru
End Sub
Private Sub CommandButton8_Click() 'ÇOKLU MALZEME DEĞİŞTİRME ???
On Error Resume Next
If ListBox1.listIndex < 0 And ListBox21.listIndex < 0 Then Exit Sub
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
Dim n As Integer
UFMZ.LBLS1.AddItem TextBox1.Value
For n = 1 To 8
    UFMZ.LBLS1.List(UFMZ.LBLS1.ListCount - 1, n) = msayfa.Cells(z, n + 1)
Next n
UFMZ.LBLS1.List(UFMZ.LBLS1.ListCount - 1, 9) = "YTL"
If msayfa.Cells(z, 6).NumberFormat = "#,##0.00 [$€-1]" Then UFMZ.LBLS1.List(UFMZ.LBLS1.ListCount - 1, 9) = "EUR"
If msayfa.Cells(z, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then UFMZ.LBLS1.List(UFMZ.LBLS1.ListCount - 1, 9) = "USD"
End Sub
Private Sub CommandButton3_Click() 'MALZEME DEĞİŞTİRME ???
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
Call tdegistir
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
End Sub
Sub tdegistir()
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error Resume Next
Dim k, s, t, a
Dim sh
Dim tds1
sh = Selection.row
If Range("A1").Interior.Color = 16751103 Then
TextBoxT3.Value = UFMZ.Tbk
k = UFMZ.TextBox1
GoTo git:
End If
If TextBox2.Value = "" Then Exit Sub
git:
If k = TextBox2.Value Then MsgBox ("Aynı Malzeme Kodları!  "), vbCritical, "Uyarı": Exit Sub
z = 2: zx = 2: tds1 = "Sayfa1"
If dtx = 2 Then tds1 = ActiveWorkbook.ActiveSheet.Name: zx = Selection.Column

If CheckBoxP32.Value = True Then
s = Workbooks(dt).Worksheets(tds1).Range("B65536").End(xlUp).row
tekrar:
 Set t = Workbooks(dt).Worksheets(tds1).Range(Cells(z, zx), Cells(s, zx)) 't = "B" & z & ":" & "B" & s
 Set a = t.Find(k, LookIn:=xlValues, LookAt:=xlPart)
 If Not a Is Nothing Then
 TextBoxT3.Value = a.row
 If dtx = 1 Then Call verigir Else verigirx
 Cells(a.row, zx).Font.ColorIndex = 7
 GoTo tekrar
 End If
Else
If dtx = 1 Then Call verigir Else verigirx
Cells(a.row, zx).Font.ColorIndex = 7
End If
Cells(sh, zx).Select
'Range("B" & sh).Select
If Range("A1").Interior.Color = 16751103 Then UF2.Hide: Exit Sub
MultiPage1.Value = 0
UF2.Hide
If Range("A1").Interior.Color = 16764108 Then GoTo atla:
Workbooks(dt).Application.Range("A1").Interior.Color = 65535
atla:
Workbooks(dt).Application.Range("A1").Interior.Color = 49407 'X
End Sub
Sub verigirx()
Application.Calculation = xlCalculationManual
Dim dtxs, ayır
dtxs = GetSetting("ilhan", "Settings", "dtxs")
ayır = Split(dtxs, "-")
Dim mliste, Y
Set mliste = Workbooks(mlz).Worksheets("Sayfa1")
Y = TextBoxT3.Value
z = CDbl(TextBox17.Text)
Cells(Y, zx) = mliste.Range("B" & z)
Range(ayır(0) & Y) = mliste.Range("C" & z) 'tanım
Range(ayır(1) & Y) = mliste.Range("D" & z) 'Marka
Range(ayır(2) & Y) = mliste.Range("F" & z) 'Mlz. Br. Fiyat
Range(ayır(2) & Y).NumberFormat = mliste.Range("F" & z).NumberFormat
'Range("G" & Y) = mliste.Range("G" & z) 'mlz.isk.
kur = ""
If Range(ayır(2) & Y).NumberFormat = "#,##0.00 [$$-C0C]" Then Range(ayır(2) & Y).Font.ColorIndex = 3
If Range(ayır(2) & Y).NumberFormat = "#,##0.00 [$€-1]" Then Range(ayır(2) & Y).Font.ColorIndex = 5
z = Y + 1
End Sub
Sub malzemeler() '08.12.2013 malzeme değişimi için malzeme dökümü alma
Windows(dt).Activate: Sheets("Sayfa1").Select
ListBoxP31.Clear
Dim n, skd, BR
Dim son As Integer
son = Range("D65536").End(xlUp).row
X = 0
'kd1 = "PM" 'işçilik,bara,ambalaj vs.
'kd2 = "AD" 'pano
'kd3 = "AS" 'pano
'kd4 = "AP" 'pano
'kd5 = "AH" 'pano
'kd6 = "AK" 'pano
For n = 2 To son
    If WorksheetFunction.CountIf(Range("D2:D" & n), Cells(n, 4).Value) = 1 And Cells(n, 3).Value <> "" Then
    skd = Left(Cells(n, "a"), 2)
    'If skd = kd1 Or skd = kd2 Or skd = kd3 Or skd = kd4 Or skd = kd5 Or skd = kd6 Then GoTo atla
    ListBoxP31.AddItem Cells(n, 4)
    
    If Left(Cells(n, "a"), 2) = "PB" Then
    BR = "Kg."
    Else
    BR = "Adet"
    End If
    ListBoxP31.List(X, 1) = (WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("E1:E65536")) & " " & BR)
    X = X + 1
   End If
'atla:
Next n
'--
Dim liste As Variant ' a dan z ye sıralama
liste = ListBoxP31.List 'Değişkenimize ListBox'taki listeyi aldık
ListBoxP31.List = Sirala(liste, ListBoxP31.ColumnCount, 1) ' a dan z ye sıralama burada 1. sutunu aldık
'--
AlignListColumn ListBoxP31, 1, True
End Sub
Private Function Sirala(liste As Variant, Sutun_Adedi As Byte, Siralanacak_Sutun_No As Byte) ' a dan z ye sıralama
Dim i As Integer, j As Integer, say As Byte, X As Variant
For i = LBound(liste) To UBound(liste) - 1
For j = i + 1 To UBound(liste)
If StrComp(liste(i, Siralanacak_Sutun_No - 1), liste(j, Siralanacak_Sutun_No - 1), vbTextCompare) = 1 Then
For say = 0 To Sutun_Adedi - 1
X = liste(j, say)
liste(j, say) = liste(i, say)
liste(i, say) = X
Next
End If
Next j
Next i
Sirala = liste
End Function
Private Sub ListBoxP31_Click() '08.12.2013 seçilen için malzeme dökümü alma
On Error Resume Next
'Application.ScreenUpdating = False
'Application.EnableEvents = False
If ListBoxP31.Text = "" Then Exit Sub
ListViewP31.ListItems.Clear
ListViewP31.ColumnHeaders.Clear
ListViewP31.Width = 643
ListViewP31.View = lvwReport
ListViewP31.OLEDragMode = ccOLEDragAutomatic
ListViewP31.FullRowSelect = True
ListViewP31.Sorted = False
Call ListViewP31.ColumnHeaders.Add(1, , "Referans", 100)
Call ListViewP31.ColumnHeaders.Add(2, , "Açıklama", 415)
Call ListViewP31.ColumnHeaders.Add(3, , "Adet", 52, 1)
Call ListViewP31.ColumnHeaders.Add(4, , "Fiyat", 75, 1)
Dim Y, z, a, i, p As Single
Dim s As Double
Dim listson As Long
Dim SayfaMarka, formmarka, formkod, mm As String
'''''
        a = WorksheetFunction.CountA(Sheets("Sayfa1").Range("B:B"), xlDown) - 1
        If a < 2 Then Exit Sub
        a = 0
ali:
''''
      formmarka = ListBoxP31.Text
''''
        p = WorksheetFunction.CountIf(Sheets("Sayfa1").Range("D:D"), formmarka)
''''
        Y = 0
        z = 0
Do Until Y = p

If Y > p Then Exit Sub
yok:
    If Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1) = "" Then
                Do Until Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1) <> ""
                z = z + 1
                If z > Range("B65536").End(xlUp).row Then Exit Sub
                If Y > p Then Exit Sub
                Loop
    Else:
                SayfaMarka = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 3)
''''
                If UCase(formmarka) = UCase(SayfaMarka) Then GoTo devam:
                z = z + 1
GoTo yok:
devam:
''''
                Dim kt As String
                kt = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1)

''''
                    i = 0
                listson = ListViewP31.ListItems.Count
                Do Until listson = i
                formkod = ListViewP31.ListItems(i + 1)
                    If formkod = kt Then GoTo var:
                    i = i + 1
                Loop
'c'''''
listson = ListViewP31.ListItems.Count
Call ListViewP31.ListItems.Add((listson + 1), , Range("A1").Offset(Y + z + 1, 1))
Call ListViewP31.ListItems(listson + 1).ListSubItems.Add(1, , Range("A1").Offset(Y + z + 1, 2))
mm = (WorksheetFunction.SumIf(Range("B1:B65536"), kt, Range("E1:E65536")))
Call ListViewP31.ListItems(listson + 1).ListSubItems.Add(2, , mm)
Call ListViewP31.ListItems(listson + 1).ListSubItems.Add(3, , Range("A1").Offset(Y + z + 1, 10).Text)
var:
        Y = Y + 1
    End If
Loop
ListViewP31.Sorted = True
'Application.ScreenUpdating = True
'Application.EnableEvents = True
Exit Sub
GoTo ali
End Sub
Private Sub CommandButton5_Click()
Call malzemeler
CommandButton5.BackColor = &HDBE8DB
End Sub
Private Sub ListViewP31_dblClick()
On Error GoTo hata
TextBoxkod.Value = ListViewP31.SelectedItem
TextBoxacık.Value = ListViewP31.SelectedItem.ListSubItems(1)
TextBoxadet.Value = ListViewP31.SelectedItem.ListSubItems(2)
P31L = ListViewP31.SelectedItem.Index
MultiPage2.Value = 0
MultiPage1.Value = 2
Label18.Caption = "  " & TextBoxacık.Value & " (" & ListBoxP31.Text & ")": Label18.BackColor = &HE1F9F8
Label18.ForeColor = &HB97E4A: Label18.Font.Bold = False
hata:
End Sub
Private Sub Label18_Click()
If Label18.ForeColor = &HB97E4A Then MsgBox " » Değişecek Ürün Bilgileri" & vbLf _
& " Sipariş Kd. : " & ListViewP31.ListItems(P31L) & vbLf _
& " Ürün Açıklaması : " & ListViewP31.ListItems(P31L).SubItems(1) & vbLf _
& " Toplam Adet : " & ListViewP31.ListItems(P31L).SubItems(2) & vbLf _
& " Malzeme Net Alış Fiyatı : " & ListViewP31.ListItems(P31L).SubItems(3) & " TL" & vbLf _
& " Malzeme Net Toplam Fiyatı : " & Format(ListViewP31.ListItems(P31L).SubItems(3) * ListViewP31.ListItems(P31L).SubItems(2), "#,##0.00") & " TL"
End Sub
Private Sub DT1_Click()
'Application.ScreenUpdating = False
On Error Resume Next
Application.Calculation = xlCalculationManual
ListViewP31.Width = 643: MultiPage2.Value = 1
Call malzemedegisim
ListBoxP31_Click
Application.Calculation = xlCalculationAutomatic
CommandButton5.BackColor = &HBCCAF3
If ListViewP31.ListItems.Count > 0 Then ListViewP31.ListItems(P31L).Selected = True: ListViewP31.ListItems(P31L + 1).EnsureVisible
'Application.ScreenUpdating = True
End Sub
Private Sub CommandButton6_Click()
ListViewP31_dblClick
End Sub
Sub malzemedegisim() '08.12.2013 malzeme değişimi için
On Error Resume Next
Dim k, s, t, a
If Range("A1").Interior.Color = 16751103 Then
TextBoxT3.Value = UFMZ.Tbk
GoTo git:
End If
If TextBox2.Value = "" Then Exit Sub
k = TextBoxkod
git:
If CheckBox1.Value = True Then GoTo akod
If k = TextBox2.Value Then MsgBox ("Aynı Malzeme Kodları!  "), vbCritical, "Uyarı": Exit Sub
akod:
'UF2.Height = 50
'Toolbar2.Buttons(5).Value = 1
s = Workbooks(dt).Worksheets("Sayfa1").Range("B65536").End(xlUp).row
z = 2
tekrar:
t = "B" & z & ":" & "B" & s
Set a = Workbooks(dt).Worksheets("Sayfa1").Range(t).Find(k, LookIn:=xlValues, LookAt:=xlWhole) 'LookAt:=xlWholeBİREBİR EŞLEME
If Not a Is Nothing Then
TextBoxT3.Value = a.row: z = TextBoxT3.Value + 1
'---
If il = 0 Then
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call verigir2 Else Call msgteklif2: GoTo son1
Else
Call verigir
End If
Cells(TextBoxT3.Value, "B").Font.ColorIndex = 7
GoTo tekrar
End If
'Toolbar2.Buttons(5).Value = 0
son1:
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub Y1_Click()
ListBoxP31_Click
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'TAMAM
On Error Resume Next
'If Not MultiPage1.Value = 0 Then CommandButton1.Visible = True: Exit Sub
'If Range("A1").Interior.Color = 16751103 Then UF2.Hide: Exit Sub
    Windows(mlz).Close False: mlz = Empty
    UnhookListBoxScroll
    Windows(dt).Activate
    Sheets("Sayfa1").Select
    If Range("A1").Interior.Color = 16751103 Then Exit Sub
    Range("A1").Interior.Color = 65535
End Sub
Private Sub TextBox21_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
On Error Resume Next
'TextBox211.Text = ""
If Len(TextBox21) <= 3 Then Exit Sub
Call TBox21
End Sub
Sub TBox21()
On Error Resume Next
Dim ad, deg
Dim b, c
ad = TextBox21.Text
ListBox21.Clear: TextBox22.Text = ""
X = 0
deg = ""
Set c = Workbooks(mlz).Worksheets("Sayfa1").Range("C2:C65000").Find(ad, LookAt:=xlPart)
If Not c Is Nothing Then
b = c.Address
Do
If c.row <> deg Then
ListBox21.AddItem Workbooks(mlz).Worksheets("Sayfa1").Range("B" & c.row)
ListBox21.List(ListBox21.ListCount - 1, 1) = Workbooks(mlz).Worksheets("Sayfa1").Range("C" & c.row)
ListBox21.List(ListBox21.ListCount - 1, 2) = Workbooks(mlz).Worksheets("Sayfa1").Range("F" & c.row).Text
ListBox21.List(ListBox21.ListCount - 1, 3) = c.row
deg = c.row
Set c = Workbooks(mlz).Worksheets("Sayfa1").Range("C2:C65000").FindNext(c)
End If
Loop While Not c Is Nothing And c.Address <> b
AlignListColumn ListBox21, 2, True
End If
End Sub
Private Sub CommandButton7_Click()
If Me.ListBox21.ListCount <> 0 Then
ListBox21.List = Diz(ListBox21.List, 3)
Else: MsgBox "ListBox boş", vbCritical
End If
End Sub
Private Function Diz(ByVal Dizim As Variant, Stn As Integer) As Variant
    Dim i, j, k As Long
    Dim Tmp As Variant
    Stn = Stn - 1
    For i = LBound(Dizim, 1) To UBound(Dizim, 1)
        For j = i + 1 To UBound(Dizim, 1)
            If Dizim(i, Stn) > Dizim(j, Stn) Then
                For k = LBound(Dizim, 2) To UBound(Dizim, 2)
                    Tmp = Dizim(j, k)
                    Dizim(j, k) = Dizim(i, k)
                    Dizim(i, k) = Tmp
                Next
            End If
        Next
    Next
    Diz = Dizim
End Function
Private Sub TextBox211xxx_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If Len(TextBox21) <= 3 Then Exit Sub
Call TBox21
On Error Resume Next
Dim i As Integer
Dim son As Integer
i = 0
Do Until i = Me.ListBox21.ListCount
       If Not Me.ListBox21.List(i, 1) Like "*" & TextBox211.Value & "*" Then
      UF2.ListBox21.RemoveItem (i)
      i = i - 1
      son = son + 1
    End If
    i = i + 1
Loop
End Sub
Private Sub TextBox22_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Call TBox22
End Sub
Sub TBox22()
If Len(TextBox22) <= 3 Then Exit Sub
On Error Resume Next
Dim ad, deg
Dim b, c
ad = TextBox22.Text
ListBox21.Clear: TextBox21.Text = "" ': TextBox211.Text = ""
X = 0
deg = ""
Set c = Workbooks(mlz).Worksheets("Sayfa1").Range("B2:B65000").Find(ad, LookAt:=xlPart)
If Not c Is Nothing Then
b = c.Address
Do
If c.row <> deg Then
ListBox21.AddItem Workbooks(mlz).Worksheets("Sayfa1").Range("B" & c.row)
ListBox21.List(ListBox21.ListCount - 1, 1) = Workbooks(mlz).Worksheets("Sayfa1").Range("C" & c.row)
ListBox21.List(ListBox21.ListCount - 1, 2) = Workbooks(mlz).Worksheets("Sayfa1").Range("F" & c.row).Text
ListBox21.List(ListBox21.ListCount - 1, 3) = c.row
deg = c.row
Set c = Workbooks(mlz).Worksheets("Sayfa1").Range("B2:B65000").FindNext(c)
End If
Loop While Not c Is Nothing And c.Address <> b
AlignListColumn ListBox21, 2, True
End If
End Sub
Private Sub ListBox21_Click()
On Error Resume Next
Windows(dt).Activate
'Sheets("Sayfa1").Select '
TextBox6.Value = ""
TextBox9.Value = ""
TextBox5.Enabled = True
TextBox5.Value = 1
CommandButton1.Enabled = True
DT1.Enabled = True
CommandButton3.Enabled = True
CommandButton8.Enabled = True
SpinButton1.Enabled = True
X = ListBox21.listIndex
z = ListBox21.List(X, 3)
TextBox17.Text = z
Set msayfa = Workbooks(mlz).Worksheets("Sayfa1")
TextBox1.Text = msayfa.Range("K" & z + 1).End(3) 'Referans
If z <= 3 Then TextBox1.Text = msayfa.Range("K" & z) 'Referans
TextBox2.Text = msayfa.Range("B" & z)
TextBox4.Text = msayfa.Range("D" & z)
TextBox6.Value = msayfa.Range("F" & z).Text 'birim fiyat
TextBox7.Value = Format((msayfa.Range("G" & z) * 100), "#,##0.0") 'iskonto
TextBox8.Text = Replace(msayfa.Range("B" & z), "/", " ") 'RESİM SEÇİMLERİ
ScrollBarP21.Value = TextBox7.Value * 2 'iskonto
'--
TextBox9.Value = Format((msayfa.Range("F" & z).Value - ((msayfa.Range("F" & z).Value * TextBox7.Text) / 100)) * TextBox5.Value, "#,##0.00") 'net fiyat
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$€-1]" Then TextBox9.Value = TextBox9.Value & " €"   '€ kuru
If msayfa.Range("F" & z).NumberFormat = "#,##0.00 [$$-C0C]" Then TextBox9.Value = TextBox9.Value & " $" '$ kuru
'--
If ListBox21.Text = "•" Or ListBox21.Text = "BAŞLIK" Then Frame1.Enabled = False: Exit Sub Else Frame1.Enabled = True
'--
'If TextBox4.Text = "SIEMENS" Or TextBox4.Text = "EATON" Or TextBox4.Text = "ABB" Or TextBox4.Text = "SCHNEIDER" Or TextBox4.Text = "PHOENIX" Then _
Label63.Visible = True Else Label63.Visible = False
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & rs & "\" & TextBox8.Value & ".jpg")
End Sub
Private Sub ListBox21_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'TAMAM
On Error Resume Next
If ListBox21.Text = "•" Or ListBox21.Text = "BAŞLIK" Then Exit Sub
'--
If il = 0 Then
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call verigir2 Else Call msgteklif2: GoTo son1
Else
If MultiPage1.Value = 0 Then Call verigir
If Left(ActiveSheet.CodeName, 3) <> "OTM" Then Call ToplamT2
If MultiPage1.Value = 1 Then Call CommandButton3_Click
If MultiPage1.Value = 2 Then Call DT1_Click
End If
son1:
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub ListBoxA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If ListBoxA.ListCount > 15 Then
    HookListBoxScroll Me, Me.ListBoxA
End If
End Sub
Private Sub ListBoxB_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If ListBoxB.ListCount > 15 Then
    HookListBoxScroll Me, Me.ListBoxB
End If
End Sub
Private Sub ListBox21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If ListBox21.ListCount > 32 Then
    HookListBoxScroll Me, Me.ListBox21
End If
End Sub
Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If ListBox1.ListCount > 18 Then
    HookListBoxScroll Me, Me.ListBox1
End If
End Sub