Dim oran As Byte 'form oranı+++
Dim a As Integer, b As Integer, pad As Integer 'pano ürün aralığı+++
Dim m, od, os As Integer 'object vs sayıları+++
Dim oetop As Double 'örtü sacı üst ölçü+++
Dim EECT As Byte, PCAD As Byte 'pano cerceve numarası
Dim OEMLZ As Byte 'MALZEME YATAY(0)-DİKEY(1)
Dim rswidth As Double, malz_modulgen As Double, malz_modulsayısı As Double, pano_modulgen As Double 'ürün gen+++
Dim msg As Byte
Dim yt As Byte
Dim H15, H20, H30, G60, GY60, YY60 As Integer 'örtü sacı ebat
Dim mg 'var
Dim QLG00, QLG01, QLG02, QLGEN, LS00 As Integer 'ürün gen.
Dim QL00ad, QL01ad, QL02ad, ls00ad  As Integer 'XXXXXXX
Dim panoyarıkyük As Integer, OEYUK As Integer, OEDYUK  As Integer
Dim pano_kalanyarıkgen, panoyarık
Dim PCYK, PCYUK, PCGEN, PDGEN1, PDGEN2, TPAYUK As Integer 'pano iç-DIŞ boyutlar
Dim qg, ms, qs, Y As Integer
Dim bclor As String
Private Sub CBGN_Click()
If CB1.Value = False Then Exit Sub
If Controls("LBP").Caption = "1" And CB1.Value = True Then
If PDGEN1 <= 400 Then Exit Sub
PDGEN1 = PDGEN1 - 100 'pano dış gen.
End If
If Controls("LBP").Caption = "2" And CB1.Value = True Then
If PDGEN2 <= 400 Then Exit Sub
PDGEN2 = PDGEN2 - 100 'pano dış gen.
End If
Call panosil
PCYUK = PDYUK - 50 'pano iç yük.
Call boyutlandır
MultiPage1.Value = 0
FRP.ScrollTop = 0
Controls((LBO)).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1
End Sub
Private Sub CBGP_Click()
On Error Resume Next
If CB1.Value = False Then Exit Sub
If Controls("LBP").Caption = "1" And CB1.Value = True Then
If PDGEN1 >= 1000 Then Exit Sub
PDGEN1 = PDGEN1 + 100 'pano dış gen.
End If
If Controls("LBP").Caption = "2" And CB1.Value = True Then
If PDGEN2 >= 1000 Then Exit Sub
PDGEN2 = PDGEN2 + 100 'pano dış gen.
End If
Call panosil
PCYUK = PDYUK - 50 'pano iç yük.
Call boyutlandır
MultiPage1.Value = 0
FRP.ScrollTop = 0
Controls((LBO)).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1
End Sub
Private Sub CBN_Click()
On Error Resume Next
If CB1 = False Then Exit Sub
For i = 1 To os
If Controls("EGOM" & i).BackColor = &HBCCAF3 Then
Controls("EGOM" & i).Height = Controls("EGOM" & i).Height - (50 / oran)
egs = i
Call boyutlandır
Controls("EGOM" & egs).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1: Exit Sub
End If
Next i
For i = 1 To od
If Controls("EGOD" & i).BackColor = &HBCCAF3 Then
'Exit Sub 'düz saclarda işlem yapmasın
'Controls("EGOD" & i).Height = Controls("EGOD" & i).Height - (50 / oran)
  If i = 1 Then USTBOSLUK = USTBOSLUK - 50 Else ALTBOSLUK = ALTBOSLUK - 50
  If ALTBOSLUK < 50 Then ALTBOSLUK = 50
egs = i
Call boyutlandır
Controls("EGOD" & egs).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1: Exit Sub
End If
Next i
End Sub
Private Sub CBP_Click()
On Error Resume Next
If CB1 = False Then Exit Sub
For i = 1 To os
If Controls("EGOM" & i).BackColor = &HBCCAF3 Then
Controls("EGOM" & i).Height = Controls("EGOM" & i).Height + (50 / oran)
egs = i
Call boyutlandır
Controls("EGOM" & egs).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1: Exit Sub
'CB1 = False
End If
Next i
For i = 1 To od
If Controls("EGOD" & i).BackColor = &HBCCAF3 Then
'Exit Sub 'düz saclarda işlem yapmasın
'Controls("EGOD" & i).Height = Controls("EGOD" & i).Height + (50 / oran)
  If i = 1 Then USTBOSLUK = USTBOSLUK + 50 Else ALTBOSLUK = TPAYUK + 50
egs = i
Call boyutlandır
Controls("EGOD" & egs).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1: Exit Sub
End If
Next i
End Sub
Private Sub CommandButton1_Click()
CBS1.Value = False
Call tekliftenpanoaktar
End Sub
Sub tekliftenpanoaktar()
PDGEN1 = 600  'pano dış gen.
PDGEN2 = 600  'pano dış gen.
PCGEN = 500  'pano iç gen.
G60 = PDGEN2 - 115: GY60 = G60 - 53 'örtü sacı gen.için
CB1 = False
ListViewmzmtip.ListItems.Clear
Call panosil
PCYUK = PDYUK - 50 'pano iç yük.
Call panoara
Call boyutlandır
MultiPage1.Value = 0
FRP.ScrollTop = 0
If msg = 1 Then Call Msg1
End Sub
Private Sub CommandButton3_Click()
CBS1.Value = True
Call tekliftenpanoaktar
End Sub
Sub panosil()
On Error Resume Next
For i = 1 To m
Me.Controls.Remove ("MLZ" & i)
Next i
On Error GoTo hata1
For i = 1 To os
If Controls("EGOM" & i).BackColor = &HBCCAF3 Then LBO = "EGOM" & i
Me.Controls.Remove ("EGAOM" & i)
Me.Controls.Remove ("EGOM" & i)
Next i
hata1:
On Error GoTo hata2
For i = 1 To od
If Controls("EGOD" & i).BackColor = &HBCCAF3 Then LBO = "EGOD" & i
Me.Controls.Remove ("EGOD" & i)
Next i
hata2:
End Sub
Sub panoyenile1()
PCYUK = PDYUK - 50 'pano iç yük.
Call panoara
Call boyutlandır
MultiPage1.Value = 0
End Sub
Sub formayarlama()
If FRP.Zoom = 200 Then kat = 2 Else kat = 1
If EECT > 1 Then
UFOPAN11.Width = (FREPP1.Width + FREPP2.Width) * kat + 100
Else
UFOPAN11.Width = FREPP1.Width * kat + 100
End If
MultiPage1.Width = UFOPAN11.Width - 30
FRP.Width = UFOPAN11.Width - 75
FRP.ScrollTop = 0: FRP.ScrollLeft = 0
LabelPO.Width = MultiPage1.Width
'..
If FREPP1.Height + 100 < 400 Then
UFOPAN11.Height = 400
Else
If FREPP1.Height + 110 < 615 Then UFOPAN11.Height = FREPP1.Height + 110 Else UFOPAN11.Height = 615: FRP.ScrollHeight = UFOPAN11.Height + 75
End If

MultiPage1.Height = UFOPAN11.Height - 65
FRP.Height = MultiPage1.Height - 25
If EECT = 1 Then LabelPO = UFOPAN11.Caption & " = " & "Yük." & PDPYUK & " x Gen." & PDGEN1 & "mm "
If EECT > 1 Then LabelPO = UFOPAN11.Caption & " = " & "Yük." & PDPYUK & " x Gen.(" & PDGEN1 & " + " & PDGEN2 & ") mm "
End Sub
Sub yazdır()
Application.ScreenUpdating = False
On Error GoTo hata
Dim RetStat
     'Me.Zoom = 200
RetStat = Application.Dialogs(xlDialogPrinterSetup).Show
UFOPAN11.Hide
A1 = UFOPAN11.Height: A2 = FRP.Height: a3 = MultiPage1.Height
UFOPAN11.Height = FREPP1.Height + 120: FRP.Height = FREPP1.Height + 80
MultiPage1.Height = FREPP1.Height + 90: FRP.ScrollTop = 0
FRP.ScrollBars = 0: FrameR.Visible = False: Toolbar1.Visible = False
If RetStat Then Me.PrintForm
hata:
UFOPAN11.Height = A1:   FRP.Height = A2: MultiPage1.Height = a3
FRP.ScrollBars = 3: FrameR.Visible = True: Toolbar1.Visible = True
UFOPAN11.Show
Application.ScreenUpdating = True
     'Me.Zoom = 100
End Sub
Private Sub CommandButton2_Click()
On Error Resume Next
Dim mds As Integer, mad As Integer
If MDY1 - MD >= 0 Then
ms = ListViewmzmtip.ListItems.Count
YDMD = "YEDEK" & "-" & "MD" & "-" & "M10"
mad = MDY1 - MD
Set ydmi = ListViewmzmtip.FindItem(YDMD)
 If Not ydmi Is Nothing Then
 ListViewmzmtip.ListItems(ListViewmzmtip.FindItem(YDMD).Index).ListSubItems(1).Text = mad
 ListViewmzmtip.ListItems(ListViewmzmtip.FindItem(YDMD).Index).ListSubItems(2).Text = 10 * mad / 10
 GoTo atla1
 End If
Call ListViewmzmtip.ListItems.Add(ms + 1, , YDMD)
Call ListViewmzmtip.ListItems(ms + 1).ListSubItems.Add(1, , mad)
Call ListViewmzmtip.ListItems(ms + 1).ListSubItems.Add(2, , 10 * mad / 10)
Call resimgen
Call ListViewmzmtip.ListItems(ms + 1).ListSubItems.Add(3, , rswidth)
atla1:
Call yenile1
End If
Set ydmi = Empty
If msg = 1 Then Call Msg1
End Sub
Private Sub Label284_Click()
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Modül Hesabı\Pano Yüksekliği.txt"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button) '+1
On Error Resume Next
Select Case Button.Index
Case 1
Call yenile1
Toolbar1.Buttons.Item(2).Caption = "Büyült"
If msg = 1 Then Call Msg1
Case 2
MultiPage1.Value = 0
If FRP.Zoom = 100 Then
FRP.Zoom = 200
Toolbar1.Buttons.Item(2).Caption = "Küçült"
formayarlama
Else
FRP.Zoom = 100
Toolbar1.Buttons.Item(2).Caption = "Büyült"
formayarlama
End If
Case 3
MultiPage1.Value = 1
UFOPAN11.Width = 500
MultiPage1.Width = 475
Call panoebatlar
Case 4
Call yazdır
Case 5
Call bilgi1
Case 6
Unload Me
End Select
End Sub
Sub bilgi1()
If LBEA1.Visible = True Then ea1 = "Ölçü Aleti :96x96 "
MsgBox (ea1 & vbCr & "Boş Yer : %" & Application.WorksheetFunction.RoundUp((TBPMDB.Value * 100) / (CDbl(TBPMDB.Value) + MDY1), 0) & vbCr & _
"Boş Yer (Modül) :" & TBPMDB.Value) & vbCr & "Kullanılan (Modül) :" & MDY1, vbInformation
End Sub
Sub yenile1()
Call panosil
PCYUK = PDYUK - 50 'pano iç yük.
FRP.Zoom = 100
Call boyutlandır
MultiPage1.Value = 0
FRP.ScrollTop = 0
FRP.ScrollLeft = 0
Controls((LBO)).BackColor = &HBCCAF3
If msg = 1 Then Call Msg1
End Sub
Private Sub UserForm_Initialize()
If Selection.row = 1 Then Exit Sub
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
Toolbar1.Buttons.Item(4).Image = ImageList1.ListImages.Item(10).Index
Toolbar1.Buttons.Item(5).Image = ImageList1.ListImages.Item(11).Index
Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(5).Index
'malzeme giriş verileri--
QLG00 = 76.2
QLG01 = 106
QLG02 = 140
LS00 = 70
'pano giriş verileri--
PDGEN1 = 600  'pano dış gen.
PDGEN2 = 600  'pano dış gen.
'örtü sacı giriş verileri--
G60 = PDGEN2 - 115 'örtü sacı gen.için
GY60 = G60 - 53 'örtü sacı yarık gen.için
YY60 = 47 'örtü sacı yarık yük.
H15 = 150 'sigorta örtü sacı 150 yük.
H20 = 200 'kompakt örtü sacı 200 yük.
H30 = 300 'kompakt örtü sacı 200 yük.
ALTBOSLUK = 250
USTBOSLUK = 0
'oran--
oran = 3
Call panoara
Call boyutlandır
If msg = 1 Then Call Msg1
End Sub
Sub boyutlandır()
TBPMDB = 0
PCYUK = PDYUK - 50 'pano iç yük.
EECT = 1: oetop = 0: malz_left = 0: m = 0: od = 0: os = 0: msg = 0: PCYK = 0: PCAD = 1
PDGEN = PDGEN1: PCGEN = PDGEN - 100 'pano iç gen.
G60 = PDGEN1 - 115: GY60 = G60 - 53 'örtü sacı gen.için
panoyarık = GY60 '600 örtü sacı yarık gen.
FREPP1.Width = PDGEN / oran: FREPC1.Width = PCGEN / oran 'form pano iç-dış gen.
FREPP2.Visible = False
If USTBOSLUK >= 0 And USTBOSLUK <> "" Then Call ustbosluk1
ms = ListViewmzmtip.ListItems.Count
If ms > 0 Then Call moduler1: ms = 0
If QL00Dad = "" Then QL00Dad = 0
If QL00Yad = "" Then QL00Yad = 0
If QL01Dad = "" Then QL01Dad = 0
If QL01Yad = "" Then QL01Yad = 0
If LS00Dad = "" Then LS00Dad = 0

If CDbl(QL00Dad) > 0 Then yt = 10: Call kompaktdikeyB00: yt = 0
If CDbl(QL00Yad) > 0 Then yt = 1: Call kompaktyatayB00: yt = 0
If CDbl(QL01Dad) > 0 Then yt = 11: Call kompaktdikeyB01: yt = 0
If CDbl(QL01Yad) > 0 Then yt = 1: Call kompaktyatayB01: yt = 0
If CDbl(LS00Dad) > 0 Then Call akımtrfdikeyB00
Call bitti21
End Sub
Sub yerlestir()
On Error Resume Next
mbasla1:
Y = 1 'modül kontrol
If CB1.Value = True Then OEYUK = Controls("EGOM" & os + 1).Height * oran

If CDbl(malz_modulgen) > pano_kalanyarıkgen Then
oeyuk1 = Application.WorksheetFunction.RoundUp(oetop + (OEYUK / oran), 0)
oeyuk2 = Application.WorksheetFunction.RoundUp((PCYUK - ALTBOSLUK) / oran, 0)
If oeyuk1 > oeyuk2 And EECT > 1 Then PCAD = 0: Call bitti21
If PCAD = 2 Then Exit Sub
If oeyuk1 > oeyuk2 Then
If ALTBOSLUK > 0 And ALTBOSLUK <> "" Then Call altbosluk1  'pano sonu ekleme
EECT = EECT + 1
PDGEN = PDGEN2: PCGEN = PDGEN - 100 'pano iç gen.
G60 = PDGEN2 - 115: GY60 = G60 - 53 'örtü sacı gen.için
panoyarık = GY60 '600 örtü sacı yarık gen.
If yt = 1 Then malz_modulgen = panoyarık
UFOPAN11.Controls("FREPP" & EECT).Width = PDGEN / oran
UFOPAN11.Controls("FREPP" & EECT).Left = _
UFOPAN11.Controls("FREPP" & EECT - 1).Left + UFOPAN11.Controls("FREPP" & EECT - 1).Width
UFOPAN11.Controls("FREPC" & EECT).Width = PCGEN / oran
UFOPAN11.Controls("FREPP" & EECT).Visible = True
PCYK = 0: oetop = 0
If USTBOSLUK > 0 And USTBOSLUK <> "" Then Call ustbosluk1  ': OEYUK = H15
End If
If OEYUK + USTBOSLUK + ALTBOSLUK > PCYUK And EECT > 1 Then PCAD = 0: Call bitti21
'---
If ms > 0 Then
pkmd = Application.WorksheetFunction.RoundDown(pano_kalanyarıkgen / 18, 0)
Controls("EGAOM" & os).ControlTipText = "Boş Yer:" & pkmd & " Modül (" & pano_kalanyarıkgen & "mm)"
TBPMDB.Value = CDbl(TBPMDB.Value) + pkmd
End If
'---
Call sacekle 'örtü sacı ekleme
End If
Do Until Y > CDbl(malz_modulsayısı)
'malz_modulgen = panoyarık
pano_modulgen = malz_modulgen * Y
If pano_modulgen > pano_kalanyarıkgen Then
pano_modulgen = pano_modulgen - malz_modulgen: Exit Do
End If
Y = Y + 1
Loop
Call resimekle 'modüle göre resim ekleme
pano_kalanyarıkgen = pano_kalanyarıkgen - CDbl(pano_modulgen)
If Y < CDbl(malz_modulsayısı) + 1 Then
malz_modulsayısı = CDbl(malz_modulsayısı) + 1 - Y
GoTo mbasla1
End If
End Sub
Sub moduler1() 'modüler ürünler--
ms = ListViewmzmtip.ListItems.Count
panoyarıkyük = YY60 '600 örtü sacı yarık yük.
pano_kalanyarıkgen = 0
For i = 1 To ms
malz_kod = ListViewmzmtip.ListItems(i)
malz_modulgen = Right(ListViewmzmtip.ListItems(i), 2) * 1.8
malz_modulsayısı = CDbl(ListViewmzmtip.ListItems(i).ListSubItems(1))
malz_modultopgen = CDbl(malz_modulgen * malz_modulsayısı)
pano_modulgen = malz_modultopgen
OEYUK = H15: OEMLZ = 0
bclor = &H80000016
Call yerlestir
Next i
pkmd = Application.WorksheetFunction.RoundDown(pano_kalanyarıkgen / 18, 0)
Controls("EGAOM" & os).ControlTipText = "Boş Yer:" & pkmd & " Modül (" & pano_kalanyarıkgen & "mm)"
TBPMDB.Value = CDbl(TBPMDB.Value) + pkmd
End Sub
Sub kompaktdikeyB00() 'kompakt şalter dikey B00--
panoyarıkyük = YY60 '600 örtü sacı yarık yük.
modulgen = QLG00 'kompak gen.
malz_kod = QL00KOD
malz_modulsayısı = CDbl(QL00Dad)
malz_modulgen = modulgen
malz_left = 0
pano_kalanyarıkgen = 0
OEYUK = H20: OEMLZ = 0
'If CB1.Value = True Then OEYUK = Controls("EGOM" & os + 1).Height * oran
bclor = &H80000016
Call yerlestir
End Sub
Sub kompaktyatayB00() 'kompakt şalter yatay B00--
modulgen = panoyarık 'kompak gen.
malz_kod = QL00KOD & "-Y"
malz_modulsayısı = CDbl(QL00Yad)
malz_modulgen = modulgen
malz_left = CDbl(modulgen) / 2
pano_kalanyarıkgen = 0
OEYUK = H15: OEMLZ = 1: QLGEN = QLG00
'If CB1.Value = True Then OEYUK = Controls("EGOM" & os + 1).Height * oran
bclor = &H80000016
Call yerlestir
End Sub
Sub akımtrfdikeyB00() 'akım trf--
panoyarıkyük = YY60 '600 örtü sacı yarık yük.
modulgen = LS00 'akım trf.
malz_kod = LS00KOD
malz_modulsayısı = CDbl(LS00Dad)
malz_modulgen = modulgen
malz_left = 0
pano_kalanyarıkgen = 0
OEYUK = H20: OEMLZ = 0
'If CB1.Value = True Then OEYUK = Controls("EGOM" & os + 1).Height * oran
bclor = &H8000000F
Call yerlestir
End Sub
Sub kompaktdikeyB01() 'kompakt şalter dikey B01--
panoyarıkyük = YY60 '600 örtü sacı yarık yük.
modulgen = QLG01 'kompak gen.
malz_kod = QL01KOD
malz_modulsayısı = CDbl(QL01Dad)
malz_modulgen = modulgen
malz_left = 0
pano_kalanyarıkgen = 0
OEYUK = H30: OEMLZ = 0
'If CB1.Value = True Then OEYUK = Controls("EGOM" & os + 1).Height * oran
bclor = &H80000016
Call yerlestir
End Sub
Sub kompaktyatayB01() 'kompakt şalter yatay B01--
modulgen = panoyarık 'kompak gen.
malz_kod = QL01KOD & "-Y"
malz_modulsayısı = CDbl(QL01Yad)
malz_modulgen = modulgen
malz_left = CDbl(modulgen) / 2
pano_kalanyarıkgen = 0
OEYUK = H15: OEMLZ = 1: QLGEN = QLG01
'If CB1.Value = True Then OEYUK = Controls("EGOM" & os + 1).Height * oran
bclor = &H80000016
Call yerlestir
End Sub
Sub bitti21()
If PCAD = 2 Then Exit Sub
Call panosonu
Call formayarlama
Call isimler1
If PCAD = 0 Then
msg = 1: PCAD = 2: Exit Sub
End If
End Sub
Sub resimekle()
On Error GoTo hata
'EECT örtü çerceveye göre resim oluşturur
m = m + 1
Set Img = UFOPAN11.Controls("EGAOM" & os).Controls.Add("Forms.Image.1")
 With Img
        .Name = "MLZ" & m
        .BorderStyle = 0
        .BorderColor = &HB1B4B8
        .PictureSizeMode = 3
        .PictureAlignment = 0
        .PictureTiling = True
        .Height = Controls("EGAOM" & os).Height - (2 / oran)
        .Width = CDbl(pano_modulgen) / oran
        .Left = CDbl(malz_left) / oran
        .Top = 0
        .SpecialEffect = 0
        .ControlTipText = malz_kod & " - " & Y - 1 & " Adet"
 End With

malz_left = CDbl(pano_modulgen) + CDbl(malz_left)
malz_kod = Replace(malz_kod, "-grup", "")
malz_kod = Replace(malz_kod, "-devre", "")
Controls("MLZ" & m).Picture = LoadPicture("C:\Belgelerim\Cemex\Görünüşler\" & malz_kod & ".bmp")
Exit Sub
hata:
Controls("MLZ" & m).BorderStyle = 1
End Sub
Sub ustbosluk1()
OEDYUK = USTBOSLUK
Call düzsacekle
oetop = (OEDYUK / oran) + oetop
PCYK = OEDYUK + PCYK
End Sub
Sub altbosluk1()
OEDYUK = ALTBOSLUK
If EECT > 2 Then Exit Sub
If EECT > 1 Then
PCYUK = PDPYUK - 50 'pano iç yük.
UFOPAN11.Controls("FREPP" & EECT).Height = FREPP1.Height
UFOPAN11.Controls("FREPC" & EECT).Height = FREPC1.Height
UFOPAN11.Controls("FREPC" & EECT).Top = (FREPP1.Height - FREPC1.Height) / 2
UFOPAN11.Controls("FREPC" & EECT).Left = (FREPP1.Width - FREPC1.Width) / 2
Else
PCYUK = PCYK + ALTBOSLUK: PDPYUK = PCYUK + 50
dizi1 = PDSTDBOY 'STANDART pano yük.
For i = 0 To 8
 pcyukb1 = Split(dizi1, ";")(i)
 If CDbl(pcyukb1) >= CDbl(PDPYUK) Then PDPYUK = pcyukb1: Exit For
Next i
PCYUK = PDPYUK - 50 'pano iç yük.
FREPP1.Height = PDPYUK / oran
FREPC1.Height = FREPP1.Height - 50 / oran
FREPC1.Top = (FREPP1.Height - FREPC1.Height) / 2
FREPC1.Left = (FREPP1.Width - FREPC1.Width) / 2
TPAYUK = PCYUK - PCYK
End If
yeni1:
If PCYUK - PCYK > 0 Then
 OEDYUK = PCYUK - PCYK
 If PCYUK - PCYK <= 300 Then
 Call düzsacekle
 oetop = (OEDYUK / oran) + oetop
 PCYK = OEDYUK + PCYK
 Else
yeni2:
 If OEDYUK = 0 Then Exit Sub
 If OEDYUK > 300 Then OEDYUK = OEDYUK - 300: GoTo yeni2
 Call düzsacekle
 oetop = (OEDYUK / oran) + oetop
 PCYK = OEDYUK + PCYK
 GoTo yeni1
 End If
End If
End Sub
Sub panosonu()
Call altbosluk1
Call sil1: Call sil2
End Sub
Sub sil1()
On Error GoTo hata
osad = os
kaldır:
Me.Controls.Remove ("EGAOM" & os + 1): Me.Controls.Remove ("EGOM" & os + 1)
os = os + 1
GoTo kaldır:
hata:
os = osad
If Left(LBO.Caption, 4) = "EGOM" Then
If Replace(LBO.Caption, "EGOM", "") * 1 > os Then LBO.Caption = "EGOM" & os
End If
End Sub
Sub sil2()
On Error GoTo hata
odad = od
kaldır:
Me.Controls.Remove ("EGOD" & od + 1): Me.Controls.Remove ("EGOD" & od + 1)
od = od + 1
GoTo kaldır:
hata:
od = odad
If Left(LBO.Caption, 4) = "EGOD" Then
If Replace(LBO.Caption, "EGOD", "") * 1 > od Then LBO.Caption = "EGOD" & od
End If
End Sub
Sub sacekle()
On Error Resume Next
os = os + 1
'EECT örtü çerceveye göre EGOM ORTU SACI oluşturur
If CB1.Value = True Then
OEYUK = Controls("EGOM" & os).Height * oran
Me.Controls.Remove ("EGAOM" & os): Me.Controls.Remove ("EGOM" & os)
End If
Set EGOMS = UFOPAN11.Controls("FREPC" & EECT).Controls.Add("Forms.Image.1")
 With EGOMS
        .Name = "EGOM" & os
        .SpecialEffect = 1
        .Height = OEYUK / oran
        .Width = G60 / oran 'örtü sacı 600 gen. x/oran
        .Left = 5 / oran
        .Top = oetop
        .Tag = EECT
        .ControlTipText = "EG OM " & OEYUK / 10 & "06" & "A - h=" & OEYUK

 End With
'--
Set EGOAMS = UFOPAN11.Controls("FREPC" & EECT).Controls.Add("Forms.Frame.1")
 With EGOAMS
        .Name = "EGAOM" & os
        .SpecialEffect = 0
        .BorderStyle = 1
        .BorderColor = &HB1B4B8
        .Height = panoyarıkyük / oran
        .Width = GY60 / oran 'örtü sacı 600 gen. x/oran
        .Left = 30 / oran
        .Top = oetop + ((OEYUK - panoyarıkyük) / 2) / oran
        .BackColor = bclor
 End With
'--
If OEMLZ = 1 Then 'şalter yatay ise (1)
Controls("EGAOM" & os).Height = (QLGEN + 4) / oran
Controls("EGAOM" & os).Width = (panoyarıkyük + 4) / oran
Controls("EGAOM" & os).Left = GY60 / 2 / oran
Controls("EGAOM" & os).Top = oetop + ((OEYUK - QLGEN) / 2) / oran
Controls("EGAOM" & os).SpecialEffect = 5
End If
'--
malz_left = 0
pano_kalanyarıkgen = panoyarık
oetop = (OEYUK / oran) + oetop
PCYK = OEYUK + PCYK
End Sub
Sub düzsacekle()
On Error Resume Next
od = od + 1
If CB1.Value = True Then Me.Controls.Remove ("EGOD" & od)
'EECT örtü çerceveye göre EGOD DUZ ORTU SACI oluşturur
Set EGODS = UFOPAN11.Controls("FREPC" & EECT).Controls.Add("Forms.Image.1")
 With EGODS
        .Name = "EGOD" & od
        .SpecialEffect = 1
        .Height = OEDYUK / oran
        .Width = G60 / oran 'örtü sacı 600 gen. x/oran
        .Left = 5 / oran
        .Top = oetop
        .Tag = EECT
        .ControlTipText = "EG OM " & OEDYUK / 10 & "06" & " - h=" & OEDYUK
 End With
'--
End Sub
Sub panoara()
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
'--
LBEA1.Visible = False: LBEA1.ControlTipText = ""
If CBS1.Value = True Then
a = ActiveWindow.RangeSelection.row:  b = a + Selection.Cells.Count - 1
GoTo git2:
End If
'--
i = Selection.row: n = 2
If i > Range("B65536").End(xlUp).row Then Exit Sub
Do Until i < n
If Cells(i, 2) = "BÖLÜM ADI/NO:" And Cells(i, 6) = "" Then UFOPAN11.Caption = Cells(i, 3): a = Cells(i, 5).row: pad = Cells(i, 5): GoTo git1:
i = i - 1
Loop
git1:
If pad = 0 Then pad = 1
TBPAD = pad
i = Selection.row: n = Range("B65536").End(xlUp).row
Do Until i > n
If Cells(i, 2) = "BÖLÜM TOPLAMI:" And Cells(i, 6) = "" Then b = Cells(i, 5).row: GoTo git2:
i = i + 1
Loop
git2:
'--
Dim mds As Integer
Dim mad As Integer
Dim QL00Ds As Integer
Dim QL01Ds As Integer
Dim QL02Ds As Integer
Dim LS00Ds As Integer
Do Until a > b
'--
If Cells(a, 9) = "96x96" And LBEA1.ControlTipText = "" Then LBEA1.Visible = True: LBEA1.ControlTipText = "96x96" _
'--
mad = Cells(a, 5) / pad
If Left(Cells(a, 9), 1) = "M" Then
mzmtip = Cells(a, 4) & "-" & Cells(a, 1) & "-" & Cells(a, 9)
ms = ListViewmzmtip.ListItems.Count
Call ListViewmzmtip.ListItems.Add(ms + 1, , mzmtip)
Call ListViewmzmtip.ListItems(ms + 1).ListSubItems.Add(1, , mad)
Call ListViewmzmtip.ListItems(ms + 1).ListSubItems.Add(2, , Right(Cells(a, 9), 2) * mad / 10)
Call resimgen
Call ListViewmzmtip.ListItems(ms + 1).ListSubItems.Add(3, , rswidth)
mds = Right(Cells(a, 9), 2) * mad / 10 + mds
End If
ms = ListViewmzmtip.ListItems.Count
'--
If Left(Cells(a, 1), 2) = "QL" Then
If Left(Cells(a, 9), 3) = "B00" Then QL00Ds = mad + QL00Ds: QL00KOD = Left(Cells(a, 4), 3) & "-" & Cells(a, 1) & "-" & Cells(a, 9)
If Left(Cells(a, 9), 3) = "B01" Then QL01Ds = mad + QL01Ds: QL01KOD = Left(Cells(a, 4), 3) & "-" & Cells(a, 1) & "-" & Cells(a, 9)
If Left(Cells(a, 9), 3) = "B02" Then QL02Ds = mad + QL02Ds: QL02KOD = Left(Cells(a, 4), 3) & "-" & Cells(a, 1) & "-" & Cells(a, 9)
End If
If Left(Cells(a, 1), 2) = "LS" Then
If Left(Cells(a, 1), 2) = "LS" Then LS00Ds = mad + LS00Ds: LS00KOD = Left(Cells(a, 4), 3) & "-" & Cells(a, 1) & "-" & Cells(a, 1) & "00"
End If
a = a + 1
Loop
MD = mds: mg = mds: MDY = MD: MDY1 = MD: LBEA1.ControlTipText = ""
QL00Dad = QL00Ds: qgD0 = QL00Ds: QL00Yad = 0
QL01Dad = QL01Ds: qgD1 = QL01Ds: QL01Yad = 0
QL02Dad = QL02Ds: qgD2 = QL02Ds: QL02Yad = 0
LS00Dad = LS00Ds: qgL0 = LS00Ds
End Sub
Sub resimgen()
On Error Resume Next
Dim objShell As Object, objFolder As Object, objFile As Object
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace("C:\Belgelerim\Cemex\Görünüşler")
Set objFile = objFolder.ParseName(ListViewmzmtip.ListItems(ms + 1).Text & ".bmp")
sDims = objFile.ExtendedProperty("Dimensions")
ayır = Split(Right(sDims, Len(sDims) - 1), "x")
rswidth = ayır(0) / 2
Set objFile = Nothing
End Sub
Sub Msg1()
MsgBox ("Yerleştirilemeyen ürün var!"), vbInformation
msg = 0
End Sub
Sub panoebatlar() 'varsayılanlar TXT 2021
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Modül Hesabı\Pano Yüksekliği.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        'MsgBox "Pano Yüksekliği.txt / Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    sat1 = 1
    ListBoxPY.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
If Not Rky = "" Then
ListBoxPY.AddItem Rky
End If
sat1 = sat1 + 1
Loop
Close #Ert
End Sub
Private Sub ListBoxPY_Click()
If ListBoxPY.ListCount > 0 Then PDSTDBOY = ListBoxPY.List(ListBoxPY.listIndex)
End Sub