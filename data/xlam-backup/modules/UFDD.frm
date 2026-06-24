Dim Y, i
Public bfyt As String
Dim kur As String
Dim CTM1

Private Sub CA1_Change()

End Sub

Private Sub CheckBoxsa_Click()

End Sub

Private Sub CommandButton3_Click()
MultiPage1.Value = 0
End Sub
Private Sub CommandButton4_Click()
On Error Resume Next
ListBoxcml.Clear
For i = 1 To ListViewP1.ListItems.Count
 If ListViewP1.ListItems(i).Checked = True Then ListBoxcml.AddItem ListViewP1.ListItems(i)
Next
For i = 1 To ListViewP2.ListItems.Count
 If ListViewP2.ListItems(i).Checked = True Then ListBoxcml.AddItem ListViewP2.ListItems(i)
Next
For i = 1 To ListViewP3.ListItems.Count
 If ListViewP3.ListItems(i).Checked = True Then ListBoxcml.AddItem ListViewP3.ListItems(i)
Next
For i = 1 To ListViewP4.ListItems.Count
 If ListViewP4.ListItems(i).Checked = True Then ListBoxcml.AddItem ListViewP4.ListItems(i)
Next
MultiPage1.Value = 0
If ListBoxcml.ListCount < 1 Then Exit Sub
T2.Visible = True: ComboBox1.Visible = False
T2.Value = ListBoxcml.List(0)
CLK.Value = True
If ListViewp11.Tag = 1 Then CA1.Text = "B": CA2.Text = "F": CA3.Text = "B": CA4.Text = "F": CKM1.Value = False
If ListViewp11.Tag = 2 Then CA1.Text = "A": CA2.Text = "C": CA3.Text = "B": CA4.Text = "C": CKM1.Value = True
End Sub
Private Sub CommandButton5_Click()
MultiPage1.Value = 0
End Sub
Private Sub CommandButton6_Click()
MultiPage1.Value = 0
ListBoxcml.AddItem CTM1
If ListBoxcml.ListCount < 1 Then Exit Sub
T2.Visible = True: ComboBox1.Visible = False
T2.Value = ListBoxcml.List(0)
CLK.Value = True
If ListViewp11.Tag = 1 Then CA1.Text = "B": CA2.Text = "F": CA3.Text = "B": CA4.Text = "F": CKM1.Value = False
If ListViewp11.Tag = 2 Then CA1.Text = "A": CA2.Text = "C": CA3.Text = "B": CA4.Text = "C": CKM1.Value = True
End Sub

Private Sub Frame29_Click()

End Sub

Private Sub Frame30_Click()

End Sub

Private Sub ListViewP1_ItemCheck(ByVal Item As MSComctlLib.listItem)
If ListViewP1.ListItems.Count = 0 Then Exit Sub
For i = 1 To ListViewP1.ListItems.Count
 If ListViewP1.ListItems(i).Checked = False Then CBB1.Tag = 0: CBB1.Value = False: CBB1.Tag = 1: Exit For
Next
End Sub
Private Sub ListViewP2_ItemCheck(ByVal Item As MSComctlLib.listItem)
If ListViewP2.ListItems.Count = 0 Then Exit Sub
For i = 1 To ListViewP2.ListItems.Count
 If ListViewP2.ListItems(i).Checked = False Then CBB2.Tag = 0: CBB2.Value = False: CBB2.Tag = 1: Exit For
Next
End Sub
Private Sub ListViewP3_ItemCheck(ByVal Item As MSComctlLib.listItem)
If ListViewP3.ListItems.Count = 0 Then Exit Sub
For i = 1 To ListViewP3.ListItems.Count
 If ListViewP3.ListItems(i).Checked = False Then CBB3.Tag = 0: CBB3.Value = False: CBB3.Tag = 1: Exit For
Next
End Sub
Private Sub ListViewP4_ItemCheck(ByVal Item As MSComctlLib.listItem)
If ListViewP4.ListItems.Count = 0 Then Exit Sub
For i = 1 To ListViewP4.ListItems.Count
 If ListViewP4.ListItems(i).Checked = False Then CBB4.Tag = 0: CBB4.Value = False: CBB4.Tag = 1: Exit For
Next
End Sub
Private Sub CBB1_Click()
If CBB1.Tag = 0 Then Exit Sub
For i = 1 To ListViewP1.ListItems.Count
 If CBB1.Value = True Then ListViewP1.ListItems(i).Checked = True Else ListViewP1.ListItems(i).Checked = False
Next
End Sub
Private Sub CBB2_Click()
If CBB2.Tag = 0 Then Exit Sub
For i = 1 To ListViewP2.ListItems.Count
 If CBB2.Value = True Then ListViewP2.ListItems(i).Checked = True Else ListViewP2.ListItems(i).Checked = False
Next
End Sub
Private Sub CBB3_Click()
If CBB3.Tag = 0 Then Exit Sub
For i = 1 To ListViewP3.ListItems.Count
 If CBB3.Value = True Then ListViewP3.ListItems(i).Checked = True Else ListViewP3.ListItems(i).Checked = False
Next
End Sub
Private Sub CBB4_Click()
If CBB4.Tag = 0 Then Exit Sub
For i = 1 To ListViewP4.ListItems.Count
 If CBB4.Value = True Then ListViewP4.ListItems(i).Checked = True Else ListViewP4.ListItems(i).Checked = False
Next
End Sub
Private Sub LabelP10_Click()
On Error Resume Next
CBB1.Value = True: CBB2.Value = True: CBB3.Value = True: CBB4.Value = True
End Sub
Private Sub LabelP11_Click()
On Error Resume Next
CBB1.Value = True: CBB2.Value = True: CBB3.Value = True: CBB4.Value = True
CBB1.Value = False: CBB2.Value = False: CBB3.Value = False: CBB4.Value = False
End Sub
Private Sub LabelP12_Click()
On Error Resume Next
ListBoxM1.Clear
Dim son As Integer
son = Sheets("Sayfa1").Range("D65536").End(xlUp).row
For n = 2 To son
mss = WorksheetFunction.CountIf(Sheets("Sayfa1").Range("D2:D" & n), Sheets("Sayfa1").Cells(n, 4).Value)
If Sheets("Sayfa1").Cells(n, 1) = "" And Sheets("Sayfa1").Cells(n, 4) = "" Then GoTo atla
   If mss = 1 Then
       ListBoxM1.AddItem Sheets("Sayfa1").Cells(n, 4) 'ListBox
   End If
atla:
Next n
If Not ListBoxM1.ListCount > 0 Then Exit Sub
LabelP11_Click
For z = 0 To ListBoxM1.ListCount - 1
 mrk = UCase(Replace(Replace(ListBoxM1.List(z), "ı", "I"), "i", "İ"))
 For i = 1 To ListViewP1.ListItems.Count
 If UCase(Replace(Replace(ListViewP1.ListItems(i), "ı", "I"), "i", "İ")) Like "*" & mrk & "*" Then ListViewP1.ListItems(i).Checked = True
 Next
 For i = 1 To ListViewP2.ListItems.Count
 If UCase(Replace(Replace(ListViewP2.ListItems(i), "ı", "I"), "i", "İ")) Like "*" & mrk & "*" Then ListViewP2.ListItems(i).Checked = True
 Next
 For i = 1 To ListViewP3.ListItems.Count
 If UCase(Replace(Replace(ListViewP3.ListItems(i), "ı", "I"), "i", "İ")) Like "*" & mrk & "*" Then ListViewP3.ListItems(i).Checked = True
 Next
 For i = 1 To ListViewP4.ListItems.Count
 If UCase(Replace(Replace(ListViewP4.ListItems(i), "ı", "I"), "i", "İ")) Like "*" & mrk & "*" Then ListViewP4.ListItems(i).Checked = True
 Next
Next z
End Sub
Private Sub UserForm_Activate()
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(5).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
Toolbar2.ImageList = ImageList1
Toolbar2.Buttons.Item(1).Image = ImageList1.ListImages.Item(2).Index
Toolbar2.Buttons.Item(2).Image = ImageList1.ListImages.Item(5).Index
Toolbar2.Buttons.Item(3).Image = ImageList1.ListImages.Item(4).Index
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
ProgressBar1.Value = 1
dizi1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
For i = 1 To 26
        CA1.AddItem Mid(dizi1, i, 1)
        CA2.AddItem Mid(dizi1, i, 1)
        CA3.AddItem Mid(dizi1, i, 1)
        CA4.AddItem Mid(dizi1, i, 1)
Next
dizi1 = ""
dizi2 = "AAABACADAEAFAGAHAIAJAKALAMANAOAPAQAUAVAWARASATAXAYAZBABBBCBDBEBFBGBHBIBJBKBLBMBNBOBPBQBUBVBWBRBSBTBXBYBZCACBCCCDCECFCGCHCICJCKCLCMCNCOCPCQCUCVCWCRCSCTCXCYCZ"
m = 1
For i = 1 To 78
        CA1.AddItem Mid(dizi2, m, 2)
        CA2.AddItem Mid(dizi2, m, 2)
        CA3.AddItem Mid(dizi2, m, 2)
        CA4.AddItem Mid(dizi2, m, 2)
        m = m + 2
Next
dizi2 = ""
'--sayfaları listeleme--
'For n = 1 To ActiveWorkbook.Sheets.Count
        'With ListBox1
            '.AddItem ActiveWorkbook.Sheets(n).Name
        'End With
    'Next n
'--
Dim wb As Workbook 'Populate list box with names of open workbooks.
For Each wb In Workbooks
ComboBox1.AddItem wb.Name
Next wb
'--
End Sub
Private Sub ComboBox1_Change()
T2.Value = ComboBox1.Text
ListBoxcml.Clear
ListBoxcml.AddItem ComboBox1.Text
CLK.Value = False
If CA1.Text = "" Then CA1.Text = "B" '.TextBox2.Value
If CA2.Text = "" Then CA2.Text = "F" 'TextBox4.Value
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
CBB1.Value = False: CBB2.Value = False: CBB3.Value = False: CBB4.Value = False
ListBoxcml.Clear
Select Case Button.Index
Case 1
'--
ComboBox1.Clear
Dim wb As Workbook
For Each wb In Workbooks
ComboBox1.AddItem wb.Name
Next wb
'--
ComboBox1.Text = ActiveWorkbook.Name
'T2.Value = ActiveWorkbook.Name
Workbooks(T2.Value).Activate
T2.Visible = False: ComboBox1.Visible = True
Case 2
Call listeler: ListViewp11.Tag = 1: Image3.Visible = False
Case 3
MultiPage1.Value = 2
Call cevlisteler1: ListViewp11.Tag = 2: Image3.Visible = True
End Select
'Workbooks(T3.Value).Activate
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Tag
Case 1
Call dosyaac
If wd = 0 Then Exit Sub
T2.Value = ActiveWorkbook.Name
ComboBox1.AddItem T2.Value: ComboBox1.Text = T2.Value
T2.Visible = False: ComboBox1.Visible = True
End Select
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
If Range("B" & "65536").End(xlUp).row < 2 And Range("A" & "65536").End(xlUp).row < 2 Then MsgBox (" Dosya boş."), vbInformation, "scngnr@hotmail.com": Exit Sub
T3.Value = ActiveWorkbook.Name
Workbooks(T3.Value).Activate
CA3.Text = "B": CA4.Text = "F" '.TextBox2.Value:TextBox4.Value
Case 2
DL1.ListViewp11.Tag = 1
DL1.Show
ComboBox1.Clear
Case 3
Workbooks(T3.Value).Close
T3.Value = ""
End Select
End Sub
Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Tag
Case 1
Call dosyaac
If wd = 0 Then Exit Sub
T3.Value = ActiveWorkbook.Name
End Select
End Sub
Private Sub CommandButton2_Click()
On Error Resume Next
Workbooks(T2.Value).Close
Unload Me
End Sub
Private Sub CommandButton1xxx_Click() 'döngü yöntemi ile
On Error Resume Next
Application.ScreenUpdating = False
If T2.Value = "" Or T3.Value = "" Then Exit Sub
UFDD.Height = 45
Dim b
Dim n As String
Dim m As String
Dim kod As String
Dim aranan
Dim ara
Dim s
'''''
n = CA3.Text '.TextBox1.Value
p = CA4.Text 'TextBox3.Value
m = CA1.Text '.TextBox2.Value
r = CA2.Text 'TextBox4.Value

eski = T3.Value
yeni = T2.Value

Workbooks(eski).Activate
bul = Workbooks(eski).ActiveSheet.Range(n & "65536").End(xlUp).row
ara = Workbooks(yeni).ActiveSheet.Range(m & "65536").End(xlUp).row

ProgressBar1.Max = bul
ProgressBar1.Min = 1

For i = 2 To bul
Workbooks(T3.Value).Activate
Range(n & i).Select
    kod = ActiveCell.Text
    If kod = "" Or Cells(i, "d") = "" Then GoTo bos
        'For Each aranan In Workbooks(Yeni).Worksheets(1).Range(m & "2 :" & m & ara)
        For Each aranan In Workbooks(yeni).ActiveSheet.Range(m & "2 :" & m & ara)
            If aranan = kod Then
            s = aranan.row
            'Workbooks(Yeni).Worksheets(1).Range(m & s).Interior.ColorIndex = 7
            Workbooks(yeni).ActiveSheet.Range(m & s).Interior.ColorIndex = 7
            'b = Workbooks(Yeni).Worksheets(1).Range(r & s)
            b = Workbooks(yeni).ActiveSheet.Range(r & s)
            Range(p & i) = b: Range(n & i).Font.ColorIndex = 7: Range(p & i).Font.ColorIndex = 7
            End If
        Next
bos:
    ProgressBar1.Value = i
Next i
Range("A2").Select
Application.ScreenUpdating = True
Unload Me
End Sub
Private Sub CommandButton1_Click() ' ara bul yöntemi ile
On Error Resume Next
If ListBoxcml.ListCount < 1 Or T3.Value = "" Then Exit Sub
Workbooks(T3.Value).Activate
UFDD.Height = 45
Dim n As String
Dim p As String
Dim m As String
Dim r As String
Dim kod As String
Dim b, ara
'TextBox1.Value eski liste
n = CA3.Text: p = CA4.Text
'TextBox2.Value yeni liste
m = CA1.Text: r = CA2.Text
'...
lbs = ListBoxcml.ListCount - 1
Application.ScreenUpdating = False
'''''
For k = 0 To lbs
    ListBoxcml.Selected(k) = True ': ListBoxcml.RemoveItem (i)
    If k > 0 Then CK1.Value = True
    If CLK.Value = True Then Call mlzara
'--
Set eski = Workbooks(T3.Value).ActiveSheet
Set yeni = Workbooks(T2.Value).ActiveSheet
yeni.Cells.FormatConditions.Delete 'sayfadaki kuralları kaldır

ara = yeni.Range(m & "65536").End(xlUp).row + 1
UFDD.Caption = "(" & k + 1 & " / " & lbs + 1 & ") " & T2.Value & " dosyası kontrol ediliyor."
eski.Activate
'--
If CheckBoxsa.Value = True Then
i = ActiveWindow.RangeSelection.row:  bul = i + Selection.Cells.Count
Else
i = 2: bul = eski.Range(n & "65536").End(xlUp).row + 1
End If
'--
ProgressBar1.Max = bul
ProgressBar1.Min = 1
'--
Application.Calculation = xlCalculationManual

tekrar:
kod = eski.Range(n & i)
'marka = Workbooks(yeni).ActiveSheet.Range("D" & i)
'If kod = "" Or Cells(i, "d") = "" Then GoTo son:'?***
Left(ssayfa.Cells(i, 1), 5) = "LİSTE"
If kod = "" Or Left(kod, 5) = "BÖLÜM" Then GoTo son:
If kod = "" Or kod = "." Then GoTo son:
If CK1.Value = True Then
If Range(n & i).Interior.Color = &HD7F4DA Then GoTo son:
End If

t = m & "2 :" & m & ara
Set a = yeni.Range(t).Find(kod, LookIn:=xlValues, LookAt:=xlWhole) ' 31.10.2013 tarihinde xlPart iken xlWhole olarak değiştirdin.
Y = a.row
If i >= bul Then GoTo bitti:
If Not a Is Nothing Then
b = yeni.Range(r & Y)
If CKM1.Value = False Then eski.Range(p & i) = CDbl(b) 'sadece sayı olarak atamak için
If CKM1.Value = True Then eski.Range(p & i) = b: eski.Range(p & i).NumberFormat = "@" 'sadece metin vs.
'KUR FORMAT'--
'kur = ""
'If CKM1.Value = False Then eski.Range(p & i).NumberFormat = "#,##0.00"
'If yeni.Range(r & Y).NumberFormat = "#,##0.00 [$€-1]" Then eski.Range(p & i).NumberFormat = "#,##0.00 [$€-1]": kur = "*Eur" ' € kuru
'If yeni.Range(r & Y).NumberFormat = "#,##0.00 [$$-C0C]" Then eski.Range(p & i).NumberFormat = "#,##0.00 [$$-C0C]": kur = "*Usd" '$ kuru



'############################## Kur formatını güncelledim ############################
'eskiden kaynak hücre değeri olduğu gibi yazılırken
' Artık kaynak hücre kur değeri ile yazılıyor

'KUR FORMAT'--
kur = "" ' 'formuller' sub'ı için kullanılacak değişkeni sıfırla

If CKM1.Value = False Then
    ' 1. Değeri SAYI olarak yaz
    eski.Range(p & i) = CDbl(b)
    
    ' 2. SAYI FORMATINI kaynaktan (yeni) hedefe (eski) kopyala.
    ' Kaynak hücre €, $, TL veya standart sayı da olsa, bu satır
    ' formatı doğru şekilde hedef hücreye kopyalar.
    eski.Range(p & i).NumberFormat = yeni.Range(r & Y).NumberFormat
    
    ' 3. 'kur' değişkenini 'formuller' sub'ı için ayarla (Eski mantıkla aynı)
    If yeni.Range(r & Y).NumberFormat = "#,##0.00 [$€-1]" Then
        kur = "*Eur" ' € kuru
    ElseIf yeni.Range(r & Y).NumberFormat = "#,##0.00 [$$-C0C]" Then
        kur = "*Usd" '$ kuru
    End If
    
Else ' CKM1.Value = True ise (Metin olarak yaz)
    eski.Range(p & i) = b
    eski.Range(p & i).NumberFormat = "@" 'Hücre formatını METİN yap
End If
'################################################################################## kur sonu ###############################

'yeni.Range(r & y).Copy
'eski.Range(p & i).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'--
eski.Range(n & i).Interior.Color = &HD7F4DA       'ürün var eski kod
eski.Range(p & i).Interior.Color = &HD7F4DA       'ürün var eski kod
yeni.Range(m & Y).Interior.Color = &HD7F4DA       'ürün var yeni kod
yeni.Range(r & Y).Interior.Color = &HD7F4DA       'ürün var yeni kod

If CKF1.Value = True Then formuller
If i > 15000 Then GoTo bitti:
Else
eski.Range(n & i).Interior.Color = 13434879       'ürün yok eski kod
eski.Range(p & i).Interior.Color = 13434879       'ürün yok eski kod
End If
son:
ProgressBar1.Value = i
i = i + 1
If i >= bul Then GoTo bitti:
GoTo tekrar
bitti:
If CLK.Value = True Then Workbooks(T2.Value).Close False ': CLK.Value = False
'--
Next
'CK1.Value = False
UFDD.Caption = "Veri Güncelleme"
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
If CBUFDD.Value <> True Then Unload Me Else ProgressBar1.Value = 1: UFDD.Height = 205
End Sub
Sub listeler() 'cemex listeler
On Error Resume Next
MultiPage1.Value = 1
ListViewp11.ListItems.Clear
ListViewP1.ListItems.Clear: ListViewP2.ListItems.Clear
ListViewP3.ListItems.Clear: ListViewP4.ListItems.Clear
Dim itmX As listItem
  Dim dosya
  dosya = dir(fm1 & "\Malzeme Listeleri\1\*.xlsb")
Do While dosya <> ""
Set itmX = ListViewP1.ListItems.Add(, , dosya): 'itmX.Bold = True
   dosya = dir
Loop
  dosya = dir(fm1 & "\Malzeme Listeleri\2\*.xlsb")
Do While dosya <> ""
Set itmX = ListViewP2.ListItems.Add(, , dosya)
    dosya = dir
Loop
  dosya = dir(fm1 & "\Malzeme Listeleri\3\*.xlsb")
Do While dosya <> ""
Set itmX = ListViewP3.ListItems.Add(, , dosya)
    dosya = dir
Loop
  dosya = dir(fm1 & "\Malzeme Listeleri\4\*.xlsb")
Do While dosya <> ""
Set itmX = ListViewP4.ListItems.Add(, , dosya)
    dosya = dir
Loop
ListViewp11.StartLabelEdit
ListViewP1.StartLabelEdit: ListViewP2.StartLabelEdit: ListViewP3.StartLabelEdit: ListViewP4.StartLabelEdit:
ListViewP1.Refresh: ListViewP2.Refresh: ListViewP3.Refresh: ListViewP4.Refresh
ListViewp11.Refresh
End Sub
Sub cevlisteler1() 'ceviri listeler
On Error Resume Next
MultiPage1.Value = 2
ListViewp11.ListItems.Clear
Dim itmX As listItem
  Dim dosya
  dosya = dir("C:\Belgelerim\CEMEX\Çeviri Dosyaları\Yabancı Listeler\*.xls*")
Do While dosya <> ""
Set itmX = ListViewp11.ListItems.Add(, , dosya)
itmX.Bold = True
   dosya = dir
Loop
ListViewp11.StartLabelEdit: ListViewp11.Refresh
End Sub
Private Sub Image3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If ListViewp11.ListItems.Count < 1 Or ListViewp11.Tag = 1 Then Exit Sub
Call dosyabicim
End Sub
Sub dosyabicim()
On Error Resume Next
yold = "C:\Belgelerim\CEMEX\Çeviri Dosyaları\Yabancı Listeler\" & "[" & ListViewp11.SelectedItem & "]"
Dim str As Integer
For stt = 1 To 12
dkl = Application.ExecuteExcel4Macro("'" & yold & "Sayfa1" & "'!R1" & "C" & stt) '
If dkl = "" Or dkl = 0 Then GoTo git2
dkb = dkb & vbLf & dkl
Next
git2:
MsgBox ListViewp11.SelectedItem & " Satır Açıklamaları :" & vbLf & dkb, vbInformation, "scngnr@hotmail.com"
End Sub
Private Sub ListViewP11_dblClick()
CommandButton6_Click
End Sub
Private Sub ListViewP11_Click()
If ListViewp11.ListItems.Count < 1 Then Exit Sub
CTM1 = ListViewp11.SelectedItem.Text
End Sub
Sub mlzara()
On Error Resume Next
Application.ScreenUpdating = False
T3.Value = ActiveWorkbook.Name
Dim ds, f, f1, fc
'--
Set ds = CreateObject("Scripting.FileSystemObject")
If ListViewp11.Tag = 1 Then Set f = ds.GetFolder(fm1 & "\Malzeme Listeleri")
If ListViewp11.Tag = 2 Then Set f = ds.GetFolder("C:\Belgelerim\CEMEX\Çeviri Dosyaları")
Set fc = f.SubFolders
ds = ListBoxcml.List(ListBoxcml.listIndex)
Dim mlz As listItem
For Each f1 In fc
'--
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists(f1 & "\" & ds)
If a = True Then Workbooks.Open fileName:=f1 & "\" & ds: Exit For
'--
Next
If CLK.Value = True Then Application.Windows(ds).Visible = False
'--
T2.Value = ds
'CT1 = ds
MultiPage1.Value = 0
Application.ScreenUpdating = True
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'TAMAM
On Error Resume Next
'--
    Windows(T3.Value).Activate
    'Sheets("Sayfa1").Select
    If Range("A1").Interior.Color = 16751103 Then Exit Sub
    Range("A1").Interior.Color = 65535
    'Windows((CT1)).Close False
End Sub
Sub formuller()
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Set eski = Workbooks(T3.Value).ActiveSheet
Set yeni = Workbooks(T2.Value).ActiveSheet
Dim nameo
'--
Dim Ckar
Ckar = Workbooks(T3.Value).names("CkarO").RefersToR1C1
If Ckar = Empty Then Workbooks(T3.Value).names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'GENEL VERİLER'--
If Not Left(Range("A" & i), 3) = "PP-" Then Range("I" & i) = yeni.Range("I" & Y) 'Boyut

If yeni.Range("K" & Y) = "" Then Range("A" & i) = yeni.Range("K" & Y + 1).End(3) Else Range("A" & i) = yeni.Range("K" & Y)

'Range("A" & i) = yeni.Range("K" & Y + 1).End(3)
'Range("B" & i) = yeni.Range("B" & Y) 'Sipariş Kd.
Range("C" & i) = yeni.Range("C" & Y) 'Yapılacak İşin Cinsi
Range("D" & i) = yeni.Range("D" & Y) 'Üretici
'Range("E" & i) = yeni.Range("E" & Y) 'Miktar
'Range("F" & i) = yeni.Range("F" & Y) 'Mlz. Br. Fiyat
Range("G" & i) = yeni.Range("G" & Y) 'mlz.isk.
Range("H" & i) = yeni.Range("H" & Y) 'Adam/dk
If Not Left(Range("A" & i), 3) = "PP-" Then Range("I" & i) = yeni.Range("I" & Y) 'Boyut
If Left(ActiveSheet.CodeName, 3) = "OTM" Then GoTo zıpla1:
Range("J" & i).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt
Range("K" & i).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
'ÜRÜN GRUPLARI--
If Left(Range("A" & i), 5) = "PM-MP" Then Range("L" & i).FormulaR1C1 = bfyt & "*Oisci/100": GoTo devam1 'İŞÇ.
If Left(Range("A" & i), 5) = "PM-MS" Then Range("L" & i).FormulaR1C1 = bfyt & "*Osarf/100": GoTo devam1  'SARF
If Left(Range("A" & i), 5) = "PM-MA" Then Range("L" & i).FormulaR1C1 = bfyt & "*Oamb/100": GoTo devam1  'AMB.
If Left(Range("A" & i), 5) = "PM-MN" Then Range("L" & i).FormulaR1C1 = bfyt & "*Onak/100": GoTo devam1  'NAK.
If Left(Range("A" & i), 5) = "PM-MB" Then Range("L" & i).FormulaR1C1 = bfyt & "*Obara/100": GoTo devam1  'Bara
If Left(Range("A" & i), 3) = "PP-" Then Range("L" & i).FormulaR1C1 = bfyt & "*Opano/100": GoTo devam1  'Pano
If Left(Range("A" & i), 3) = "PS-" Then 'Pano sac & aksesuarlar
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then Range("L" & i).FormulaR1C1 = bfyt & "*Opsac/100" Else Range("L" & i).FormulaR1C1 = bfyt & "*Opano/100"
GoTo devam1
End If
Range("L" & i).FormulaR1C1 = bfyt & "*Osalt/100" 'Mlz.Kar+1
devam1:
Range("M" & i).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("N" & i).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("O" & i).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & i).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & i).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & i).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & i).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("T" & i).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.
Range("U" & i).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
'TOPLAMLAR'--
Range("W" & i).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
Range("X" & i).FormulaR1C1 = "=RC[-19]*RC[-1]"
zıpla1:
'Genel Biçimlemeler'--
Range("A" & i & ":U" & i).Borders.LineStyle = xlContinuous
Range("W" & i & ":X" & i).Borders.LineStyle = xlContinuous
Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.Bold = False
Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.ColorIndex = xlAutomatic
Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.Size = 9
Range("A" & i & ":D" & i).HorizontalAlignment = xlLeft
Range("E" & i & ":U" & i & ",W" & i & ":X" & i).HorizontalAlignment = xlRight
Range("A" & i & ":D" & i).NumberFormat = "@"
Range("E" & i).NumberFormat = "#,##0"
'Range("F" & i).NumberFormat = "#,##0.00"
Range("G" & i).NumberFormat = "0.0%"
Range("H" & i).NumberFormat = "#,##0"
Range("J" & i & ":X" & i).NumberFormat = "#,##0.00"
If Range("F" & i).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & i).Font.ColorIndex = 3
If Range("F" & i).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & i).Font.ColorIndex = 5
Set mlisteS3 = Workbooks(T3.Value).Worksheets("Sayfa3")
If mlisteS3.Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & i & ",X" & i).NumberFormat = "#,##0.00"
If mlisteS3.Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & i & ",X" & i).NumberFormat = "#,##0.00 [$€-1]"
If mlisteS3.Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & i & ",X" & i).NumberFormat = "#,##0.00 [$$-C0C]"
'--
End Sub