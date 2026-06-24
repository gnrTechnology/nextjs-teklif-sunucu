Dim a As Integer, sts As Integer
Dim nft As String

Private Sub Frame37_Click()

End Sub

Private Sub ListBox2_Click() 'VAR
TBKD1.Text = ListBox2.List(ListBox2.listIndex, 0)
End Sub
Private Sub UserForm_Initialize() 'VAR
On Error Resume Next
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
Toolbar1.Buttons.Item(4).Image = ImageList1.ListImages.Item(4).Index
ListBox1.AddItem "Şalt Malzeme": ListBox1.List(0, 1) = "XX"
ListBox1.AddItem "Pano": ListBox1.List(1, 1) = "PP-"
ListBox1.AddItem "Bakır Bara": ListBox1.List(2, 1) = "PM-MB"
ListBox1.AddItem "Sarf Malzeme": ListBox1.List(3, 1) = "PM-MS"
ListBox1.AddItem "İşçilik": ListBox1.List(4, 1) = "PM-MP"
ListBox1.AddItem "Ambalaj": ListBox1.List(5, 1) = "PM-MA"
ListBox1.AddItem "Nakliye, Vinç, Depolama vs.": ListBox1.List(6, 1) = "PM-MN"
'malzeme dosyası
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then MultiPage1.Pages.Item(0).Visible = False: MultiPage1.Value = 2: _
Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Call imalatdosya1: Exit Sub
Call malzkod2
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then _
TextBox4.Text = Worksheets("Sayfa3").Range("C7"): Frame41.Enabled = False 'TEKLİF NO
End Sub
Private Sub ListBox1_Click() 'VAR
ListBox2.Clear
TBKD1.Text = ""
CommandButton19.Enabled = True: CommandButton191.Enabled = True: CheckBox3.Enabled = True
If ListBox1.listIndex = 0 Then
For n = 1 To ListBoxMK.ListCount '
pkd = Left(ListBoxMK.List(n - 1, 0), 3)
If Not pkd = "PM-" And Not pkd = "PP-" And Not pkd = "PS-" And Not pkd = "" Then _
ListBox2.AddItem ListBoxMK.List(n - 1, 0): ListBox2.List(ListBox2.ListCount - 1, 1) = ListBoxMK.List(n - 1, 1)
Next n
Exit Sub
End If
If ListBox1.listIndex = 1 Then
For n = 1 To ListBoxMK.ListCount '
pkd = Left(ListBoxMK.List(n - 1, 0), 3)
If pkd = "PP-" Or pkd = "PS-" Then _
ListBox2.AddItem ListBoxMK.List(n - 1, 0): ListBox2.List(ListBox2.ListCount - 1, 1) = ListBoxMK.List(n - 1, 1)
Next n
End If
End Sub
Sub malzkod2() 'VAR
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, satır2 As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
    MsgBox "Malzeme Kodları.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxMK.Clear
    Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxMK.AddItem ayır(i)
'--
'kaçıncı = InStr(1, Rky, ";")
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxMK.List(satır - 1, n) = ayır(i + n)
Next n
        satır = satır + 1
    Loop
    Close #Ert
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button) 'VAR
On Error GoTo hata
TBKD1.Text = ""
Dim r As Range
Select Case Button.Index
Case 1
MultiPage1.Value = 0
Case 2
MultiPage1.Value = 3
If Not ListBoxG1.ListCount > 0 Then
For n = 1 To ListBoxMK.ListCount '
ListBoxG1.AddItem ListBoxMK.List(n - 1, 0): ListBoxG1.List(ListBoxG1.ListCount - 1, 1) = ListBoxMK.List(n - 1, 1)
Next n
End If
Case 3
MultiPage1.Value = 1
Case 4
MultiPage1.Value = 2
Call imalatdosya1
End Select
hata:
End Sub
Sub malzkod()
ListBoxG1.AddItem ListBoxmG1.List
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, satır2 As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
    MsgBox "Malzeme Kodları.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxG1.Clear
    Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxG1.AddItem ayır(i)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxG1.List(satır - 1, n) = ayır(i + n)
Next n
        satır = satır + 1
    Loop
    Close #Ert
End Sub
Private Sub OptionButton1_Click()
ListBoxG1.Clear
If Not ListBoxG1.ListCount > 0 Then
For n = 1 To ListBoxMK.ListCount '
ListBoxG1.AddItem ListBoxMK.List(n - 1, 0): ListBoxG1.List(ListBoxG1.ListCount - 1, 1) = ListBoxMK.List(n - 1, 1)
Next n
End If
End Sub
Private Sub OptionButton2_Click()
ListBoxG1.Clear
If Not ListBoxG1.ListCount > 0 Then
Dim son As Integer
son = Sheets("Sayfa1").Range("D65536").End(xlUp).row
For n = 2 To son
mss = WorksheetFunction.CountIf(Sheets("Sayfa1").Range("D2:D" & n), Sheets("Sayfa1").Cells(n, 4).Value) 'Malzeme markası sayısı
If Sheets("Sayfa1").Cells(n, 4) = "" Then GoTo atla
   If mss = 1 Then ListBoxG1.AddItem Sheets("Sayfa1").Cells(n, 4)
atla:
Next n
End If
End Sub
Private Sub CommandButton198_Click()
On Error Resume Next
If Not WorksheetFunction.CountA(Range("A:A"), xlDown) > 3 Then Exit Sub
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
Dim a, Y, b As Integer
a = WorksheetFunction.CountA(Range("B:B"), xlDown) + 1
b = Cells(1, 256).End(xlToLeft).Column ' kolon sayısı
Y = 2
Do Until Y > a
 If Cells(a, 1) = "" Then Cells(a, 2).EntireRow.Delete
a = a - 1
Loop
ProgressBarTHB.Visible = True: ProgressBarTHB.Value = 1
Y = 2
Range("A1").Select
If OptionButton1.Value = True Then
    ActiveCell.CurrentRegion.Offset(1, 0).Resize(ActiveCell.CurrentRegion.Rows.Count - 1, b).Select
    Selection.Sort key1:=Range("A2"), Key2:=Range("B2"), order1:=xlAscending, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
Else
    ActiveCell.CurrentRegion.Offset(1, 0).Resize(ActiveCell.CurrentRegion.Rows.Count - 1, b).Select
    Selection.Sort key1:=Range("D2"), Key2:=Range("A2"), order1:=xlAscending, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End If
Call adettopla1
Range("A1").Select
a = WorksheetFunction.CountA(Range("B:B"), xlDown)
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If OptionButton1.Value = True Then
Cells(Y, 2).EntireRow.Insert Shift:=xlDown: Cells(Y, 2) = "BÖLÜM ADI/NO:": Cells(Y, 3) = "GRUPLANDIRILMIŞ MALZEME LİSTESİ": a = a + 1
Cells(a, 2) = "BÖLÜM TOPLAMI:": Cells(a, "X") = "=Sum(X2:X" & a - 1 & ")": a = a + 1
End If
ProgressBarTHB.Max = a
Do Until Y >= a
If Cells(Y, 1) = "" Then GoTo atla1
sayfakod = Replace(Cells(Y, 1), "-auto", "")
ayır = Split(sayfakod, "."): sayfakod = ayır(0)
mugson = ListBoxMK.ListCount
  For X = 0 To mugson - 1
  kodL = ListBoxMK.List(X, 0)
  If sayfakod = kodL Then mlzb = ListBoxMK.List(X, 1): GoTo git
  Next X
  mlzb = "DİĞER ÜRÜNLER"
git:
Cells(Y, 1).EntireRow.Insert Shift:=xlDown
Cells(Y, 2) = "Ürün Grup Adı:": Cells(Y, 3) = mlzb: Y = Y + 1
a = WorksheetFunction.CountA(Range("B:B"), xlDown)
     Do While Y < a
If Cells(Y + 1, 1) <> Cells(Y, 1) Then Exit Do
     Y = Y + 1
     Loop
atla1:
Y = Y + 1
ProgressBarTHB.Value = Y
Loop
Y = 2
Do Until Y >= a
If Cells(Y, 2) = "Ürün Grup Adı:" Or Cells(Y, 2) = "BÖLÜM ADI/NO:" Or Cells(Y, 2) = "BÖLÜM TOPLAMI:" Then
    Range(Cells(Y, 1), Cells(Y, b)).Interior.Pattern = xlNone
    Range(Cells(Y, 1), Cells(Y, b)).Borders.LineStyle = xlContinuous
    Range("B" & Y & ":E" & Y & ",W" & Y & ":X" & Y).Borders(xlInsideVertical).LineStyle = xlNone
    Range(Cells(Y, 1), Cells(Y, b)).RowHeight = 12.75
    Range(Cells(Y, 1), Cells(Y, b)).Font.Size = 9
    Range(Cells(Y, 1), Cells(Y, b)).Font.Bold = True
    If Cells(Y, 2) = "Ürün Grup Adı:" Then Range(Cells(Y, 1), Cells(Y, b)).Font.ColorIndex = 53 Else _
    Range(Cells(Y, 1), Cells(Y, b)).Font.ColorIndex = 11
 End If
Y = Y + 1
Loop
ProgressBarTHB.Value = 1
Range("B2").Select
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
End Sub
Sub adettopla1()
On Error Resume Next
Application.ScreenUpdating = False
Dim say As Integer
 sat2 = WorksheetFunction.CountA(Range("B:B"), xlDown): ProgressBarTHB.Max = sat2
    If sat2 > 15000 Then sat2 = ActiveCell.SpecialCells(xlLastCell).row
    Cells(2, 2).Select
For i = 2 To sat2
    kod1 = Cells(i, 2)
      For n = sat2 To i + 1 Step -1
      If kod1 = Cells(n, 2) Then
       If Cells(i, 6) = Cells(n, 6) Then
       Cells(i, 5) = Cells(i, 5) + Cells(n, 5)
       Cells(n, 2).EntireRow.Delete
       sat2 = sat2 - 1
       End If
      End If
      Next
git1:
ProgressBarTHB.Value = i
Next
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton14_Click() 'VAR
Call baslıkformatlarıdüzelt
End Sub
Sub baslıkformatlarıdüzelt() ' ara bul yöntemi ile 'VAR
If Not Cells(1, "X") = "Toplam Fiyat" Then Exit Sub
'Sheets("Sayfa1").Select
Application.ScreenUpdating = False
son = Sheets("Sayfa1").Range("B65536").End(xlUp).row
Dim hcr As Range, eskihcr As Range
ProgressBarTHB.Visible = True: ProgressBarTHB.Value = 1: ProgressBarTHB.Max = son
On Error GoTo hata
Set hcr = Columns("B:B").Find(TBB1.Text, LookAt:=xlWhole)
  a = hcr.row: Cells(a, 2) = "BÖLÜM ADI/NO:": Call ad_toplam_format
Set eskihcr = hcr.Offset(2, 0)
Do
ProgressBarTHB.Value = a
Set hcr = Range(hcr.Offset(1, 0), [B65000]).FindNext
  a = hcr.row: Cells(a, 2) = "BÖLÜM ADI/NO:": Call ad_toplam_format
Set eskihcr = hcr.Offset(2, 0)
Loop
hata:
ProgressBarTHB.Value = 1: Application.ScreenUpdating = True
End Sub
Private Sub CommandButton192_Click() 'VAR
If Not Cells(1, "X") = "Toplam Fiyat" Then Exit Sub
Worksheets("Sayfa3").Range("C7") = TextBox4.Text 'Teklifin Numarası
Dim say As Integer
say = 1
Dim hcr As Range, eskihcr As Range
On Error GoTo hata
Set hcr = Columns("B:B").Find(TBB1.Text, LookAt:=xlWhole)
a = hcr.row
If TextBox3 = "" Then Cells(a, 1) = "" & TextBox4.Text & "-" & say & "" Else Cells(a, 1) = TextBox3 & say & "-" & TextBox4.Text & "-" & say & ""
Cells(a, 1).Font.ThemeColor = 6
Set eskihcr = hcr.Offset(2, 0)
atla2:
say = say + 1
Do
Set hcr = Range(hcr.Offset(1, 0), [B65000]).FindNext
a = hcr.row
If TextBox3 = "" Then Cells(a, 1) = "" & TextBox4.Text & "-" & say & "" Else Cells(a, 1) = TextBox3 & say & "-" & TextBox4.Text & "-" & say & ""
Cells(a, 1).Font.ThemeColor = 6
Set eskihcr = hcr.Offset(2, 0)
say = say + 1
Loop
hata:
End Sub
Private Sub CommandButton197_Click() 'VAR
If Not Cells(1, "X") = "Toplam Fiyat" Then Exit Sub
Dim say As Integer
Dim hcr As Range, eskihcr As Range
On Error GoTo hata
Set hcr = Columns("B:B").Find(TBB1.Text, LookAt:=xlWhole)
a = hcr.row
Cells(a, 1) = " "
Set eskihcr = hcr.Offset(2, 0)
Do
Set hcr = Range(hcr.Offset(1, 0), [B65000]).FindNext
a = hcr.row
Cells(a, 1) = " "
Set eskihcr = hcr.Offset(2, 0)
Loop
hata:
End Sub
Private Sub CommandButton15_Click() 'VAR
Call toplamformatlarıdüzelt
End Sub
Sub toplamformatlarıdüzelt() ' ara bul yöntemi ile 'VAR
If Not Cells(1, "X") = "Toplam Fiyat" Then Exit Sub
'Sheets("Sayfa1").Select
Application.ScreenUpdating = False
son = Sheets("Sayfa1").Range("B65536").End(xlUp).row
Dim hcr As Range, eskihcr As Range
ProgressBarTHB.Visible = True: ProgressBarTHB.Value = 1: ProgressBarTHB.Max = son
On Error GoTo hata
Set hcr = Columns("B:B").Find(TBT1.Text, LookAt:=xlWhole)
  Cells(hcr.row, "X") = "=Sum(X2:X" & hcr.row - 1 & ")"
  a = hcr.row: Cells(a, 2) = "BÖLÜM TOPLAMI:": Cells(a, 3) = "": Call ad_toplam_format
Set eskihcr = hcr.Offset(2, 0)
Do
ProgressBarTHB.Value = a
Set hcr = Range(hcr.Offset(1, 0), [B65000]).FindNext
  Cells(hcr.row, "X") = "=Sum(X" & eskihcr.row - 1 & ":X" & hcr.row - 1 & ")"
  a = hcr.row: Cells(a, 2) = "BÖLÜM TOPLAMI:": Cells(a, 3) = "": Call ad_toplam_format
Set eskihcr = hcr.Offset(2, 0)
Loop
hata:
ProgressBarTHB.Value = 1: Application.ScreenUpdating = True
End Sub
Sub ad_toplam_format() 'VAR
  'Biçimlemeler'
  Dim slc As Range
Set slc = Range("A" & a & ",B" & a & ":E" & a & ",F" & a & ":U" & a & ",W" & a & ":X" & a)
    slc.Borders.LineStyle = xlContinuous
    Range("B" & a & ":E" & a & ",W" & a & ":X" & a).Borders(xlInsideVertical).LineStyle = xlNone
    'slc.Borders(xlInsideVertical).LineStyle = xlNone
    slc.Interior.Pattern = xlNone
    slc.RowHeight = 12.75
    slc.Font.Size = 9
    slc.Font.Bold = True
    slc.Font.ColorIndex = 11
End Sub
Private Sub CBB0_Click()
If CBB0 = True Then
CBB1.Value = True: CBB2.Value = True: CBB3.Value = True
Else
CBB1.Value = False: CBB2.Value = False: CBB3.Value = False
End If
End Sub
Private Sub CommandButton17_Click() 'VAR
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If CBB2.Value = False And CBB1.Value = False And CBB3.Value = False Then Exit Sub
Application.Calculation = xlCalculationManual
Sheets("Sayfa1").Select
son = Sheets("Sayfa1").Range("C65536").End(xlUp).row
'malzeme listesi
If Left(ActiveWorkbook.ActiveSheet.CodeName, 3) = "CML" Then
ActiveSheet.Cells.Replace What:="[*]", Replacement:=""
For n = 1 To ActiveWorkbook.names.Count
    ActiveWorkbook.names(1).Delete
Next n
Range("A" & 2 & ":K" & son).Interior.ColorIndex = 0
Range("A" & 2 & ":K" & son).Font.ColorIndex = 56
ProgressBarTHB.Visible = True: ProgressBarTHB.Max = son
For a = 2 To son
ProgressBarTHB.Value = a
    If Cells(a, "A") <> "" Then
    If Cells(a, "B") = "" Then Range("A" & a & ":K" & a).Interior.ColorIndex = 37 Else _
    Range("A" & a).Interior.ColorIndex = 35: Range("B" & a & ":K" & a).Interior.Pattern = xlNone
    End If
    If Cells(a, "F") <> "" Then
    If Cells(a, "F").NumberFormat = "#,##0.00 [$$-C0C]" Then Cells(a, "F").Font.ColorIndex = 3
    If Cells(a, "F").NumberFormat = "#,##0.00 [$€-1]" Then Cells(a, "F").Font.ColorIndex = 5
    End If
Next a
GoTo bitti
End If
'otm.listeler
If Left(ActiveSheet.CodeName, 3) = "OTM" Then
ActiveSheet.Cells.Replace What:="[*]", Replacement:=""
For n = 1 To ActiveWorkbook.names.Count
    'ActiveWorkbook.Names(n).Delete
ad = ActiveWorkbook.names(n).RefersToR1C1
A2 = InStr(1, ad, "!")
If A2 > 0 Then
abi = "=Sayfa3!" & Split(ad, "!")(1)
    'ActiveWorkbook.Names(n).RefersToR1C1 = Application.WorksheetFunction.Replace(ad, a1, a2 - a1, "")
     ActiveWorkbook.names(n).RefersToR1C1 = abi
End If
If ActiveWorkbook.names(n).Name = "Tpb" Then ActiveWorkbook.names(n).RefersToR1C1 = 1
Next n
End If
'teklif dosyası
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then
ProgressBarTHB.Visible = True: ProgressBarTHB.Max = son
For a = 2 To son
ProgressBarTHB.Value = a
If CBB2.Value = True Or CBB3.Value = True Then Call bicimler1
If CBB1.Value = True Then Call formuller1
'--
Next a
End If
'son
bitti:
a = ""
Range("A2").Select
ProgressBarTHB.Visible = False
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub formuller1() 'VAR
If Range("D" & a) = "" Then Exit Sub
kur = ""
If Range("F" & a).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & a).Font.ColorIndex = 3: kur = "*Usd" ' Sayfa3 de $ kuru
If Range("F" & a).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & a).Font.ColorIndex = 5: kur = "*Eur" ' Sayfa3 de € kuru
If Left(ActiveSheet.CodeName, 3) = "OTM" Or Left(ActiveSheet.CodeName, 3) = "CML" Or Left(ActiveSheet.CodeName, 5) = "Icmal" Then
Exit Sub
End If
If CBB1.Value = False Then Exit Sub
Dim nameo
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'ÜRÜN GRUPLARI--
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
If Left(Range("A" & a), 5) = "PM-MP" Then Range("L" & a).FormulaR1C1 = bfyt & "*Oisci/100" & mkur: GoTo devam1 'İŞÇ.
If Left(Range("A" & a), 5) = "PM-MS" Then Range("L" & a).FormulaR1C1 = bfyt & "*Osarf/100" & mkur: GoTo devam1 'SARF
If Left(Range("A" & a), 5) = "PM-MA" Then Range("L" & a).FormulaR1C1 = bfyt & "*Oamb/100" & mkur: GoTo devam1 'AMB.
If Left(Range("A" & a), 5) = "PM-MN" Then Range("L" & a).FormulaR1C1 = bfyt & "*Onak/100" & mkur: GoTo devam1 'NAK.
If Left(Range("A" & a), 5) = "PM-MB" Then Range("L" & a).FormulaR1C1 = bfyt & "*Obara/100" & mkur: GoTo devam1 'Bara
If Left(Range("A" & a), 3) = "PP-" Then Range("L" & a).FormulaR1C1 = bfyt & "*Opano/100" & mkur: GoTo devam1 'Pano
If Left(Range("A" & a), 3) = "PS-" Then 'Pano sac & aksesuarlar
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then Range("L" & a).FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else Range("L" & a).FormulaR1C1 = bfyt & "*Opano/100" & mkur
GoTo devam1
End If
Range("L" & a).FormulaR1C1 = bfyt & "*Osalt/100" & mkur 'Mlz.Kar+1
devam1:
Range("J" & a).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("N" & a).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("K" & a).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
Range("M" & a).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("O" & a).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & a).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & a).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & a).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & a).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("T" & a).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.* işçilik katsayılı+1
Range("U" & a).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
Range("W" & a).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb" 'Dövize göre Birim Fiyat
Range("X" & a).FormulaR1C1 = "=RC[-19]*RC[-1]"  'Toplam Fiyat TL+1
End Sub
Private Sub CommandButton25_Click() 'VAR
Application.Calculation = xlCalculationManual
Call kurformat2
If Selection.Rows.Count > 1 Then Selection.SpecialCells(xlCellTypeVisible).Select
For Each deger In Selection
a = deger.row
If Not Range("D" & deger.row) = "" Then
Range("F" & deger.row).NumberFormat = nft
Range("F" & deger.row).Font.ColorIndex = xlAutomatic
CBB1.Value = True
Call formuller1
End If
Next deger
CBB1.Value = False
Application.Calculation = xlCalculationAutomatic
End Sub
Sub kurformat2() 'VAR
If OP11.Value = True Then nft = "#,##0.00"
If OP21.Value = True Then nft = "#,##0.00 [$$-C0C]"
If OP31.Value = True Then nft = "#,##0.00 [$€-1]"
End Sub
Private Sub CommandButton19_Click() 'VAR
On Error GoTo hata
Application.Calculation = xlCalculationManual
If ListBox1.listIndex < 0 Then Exit Sub
Call kurformat
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Call stkur
hata:
cdm = ActiveWorkbook.ActiveSheet.CodeName
If Left(cdm, 3) = "CML" Then Call stkurmzEK2
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub CommandButton191_Click() 'VAR
On Error GoTo hata
Sheets("Sayfa1").Select
'If ListBox1.ListIndex < 0 Or ListBox2.ListIndex < 0 Then Exit Sub
If ListBox1.listIndex < 0 Then Exit Sub
Application.Calculation = xlCalculationManual
Range("B" & Selection.row).EntireRow.Insert
ActiveCell.Offset(0, 0).Select
CheckBox3.Value = True
Call kurformat
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Call stkur
hata:
cdm = ActiveWorkbook.ActiveSheet.CodeName
If Left(cdm, 3) = "CML" Then Call stkurmzEK1
CheckBox3.Value = False
Application.Calculation = xlCalculationAutomatic
End Sub
Sub kurformat() 'VAR
If OP01.Value = True Then nft = "#,##0.00"
If OP02.Value = True Then nft = "#,##0.00 [$$-C0C]"
If OP03.Value = True Then nft = "#,##0.00 [$€-1]"
End Sub
Sub stkurmzEK1() 'VAR
On Error Resume Next
If Selection.Rows.Count > 1 Then Selection.SpecialCells(xlCellTypeVisible).Select
For Each deger In Selection
a = deger.row
Range("F" & deger.row).NumberFormat = nft
Range("B" & a) = "YENİ REFERANS"
If ListBox2.listIndex >= 0 Then Range("C" & a) = ListBox2.List(ListBox2.listIndex, 1) Else Range("C" & a) = ListBox1.List(ListBox1.listIndex, 0)
Range("D" & a) = "MARKA": Range("E" & a) = "Ad.": Range("F" & a) = 0: Range("G" & a) = 0
Range("A" & a & ":K" & a).Borders.LineStyle = xlContinuous
If Range("F" & a).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & a).Font.ColorIndex = 3
If Range("F" & a).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & a).Font.ColorIndex = 5
Range("A" & a & ":K" & a).Font.Size = 10
Range("A" & a & ":K" & a).Font.Bold = False
Range("G" & a).NumberFormat = "0.0%"
Next deger
End Sub
Sub stkurmzEK2() 'VAR
On Error Resume Next
If Selection.Rows.Count > 1 Then Selection.SpecialCells(xlCellTypeVisible).Select
For Each deger In Selection
a = deger.row
Range("F" & deger.row).NumberFormat = nft
If Not Range("A" & deger.row) = "" Then
If Not TBKD1 = "" Then Range("K" & a) = TBKD1 Else Range("K" & a) = ListBox1.List(ListBox1.listIndex, 1)
Range("A" & a).Interior.ColorIndex = 35
End If
Range("A" & a & ":K" & a).Borders.LineStyle = xlContinuous
If Range("F" & a).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & a).Font.ColorIndex = 3
If Range("F" & a).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & a).Font.ColorIndex = 5
Next deger
End Sub
Sub stkur() 'VAR
On Error Resume Next
If Selection.Rows.Count > 1 Then Selection.SpecialCells(xlCellTypeVisible).Select
For Each deger In Selection
a = deger.row
'--
Range("F" & deger.row).NumberFormat = nft
If Not TBKD1 = "" Then Range("A" & a) = TBKD1 Else Range("A" & a) = ListBox1.List(ListBox1.listIndex, 1)
If CheckBox3.Value = True Then
Range("B" & a) = "YENİ REFERANS"
If ListBox2.listIndex >= 0 Then Range("C" & a) = ListBox2.List(ListBox2.listIndex, 1) Else Range("C" & a) = ListBox1.List(ListBox1.listIndex, 0)
Range("D" & a) = "MARKA": Range("E" & a) = 1: Range("F" & a) = 0
End If
'--
CBB1.Value = True
Call bicimler1
Call formuller1
Next deger
CBB1.Value = False
End Sub
Sub bicimler1() 'VAR
'If CBB3.Value = True Then Range("A" & a & ":X" & a).Interior.ColorIndex = 0
If CBB3.Value = True Then
 Rows(a & ":" & a).Interior.ColorIndex = 0
 If Cells(a, "D") = "" Then
 If Left(Range("A" & a), 3) = "EK-" Then Range("A" & a).Interior.ColorIndex = 37 Else _
 Range("A" & a).Interior.ColorIndex = 24
 End If
End If
If CBB2.Value = True Then
 If Cells(a, "D") = "" Then
 Range("A" & a & ":X" & a).Font.Color = 8388608: _
 Range("A" & a & ":X" & a).Font.Bold = True
 Range("A" & a & ":U" & a).Borders.LineStyle = xlContinuous
 Range("W" & a & ":X" & a).Borders.LineStyle = xlContinuous
 Exit Sub
 End If
Range("A" & a & ":X" & a).Font.ColorIndex = xlAutomatic
If Range("F" & a).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & a).Font.ColorIndex = 3
If Range("F" & a).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & a).Font.ColorIndex = 5
Range("A" & a & ":U" & a).Borders.LineStyle = xlContinuous
Range("W" & a & ":X" & a).Borders.LineStyle = xlContinuous
Range("A" & a & ":X" & a).Font.Bold = False
Range("A" & a & ":X" & a).Font.Size = 9
Range("A" & a & ":D" & a).HorizontalAlignment = xlLeft
Range("E" & a & ":U" & a & ",W" & a & ":X" & a).HorizontalAlignment = xlRight
Range("A" & a & ":D" & a).NumberFormat = "@"
Range("E" & a).NumberFormat = "#,##0"
Range("G" & a).NumberFormat = "0.0%"
Range("H" & a).NumberFormat = "#,##0"
Range("J" & a & ":X" & a).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & a & ",X" & a).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & a & ",X" & a).NumberFormat = "#,##0.00 [$€-1]"
If Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & a & ",X" & a).NumberFormat = "#,##0.00 [$$-C0C]"
End If
End Sub
Private Sub Image442_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021++
On Error Resume Next
Dim ds, pds
Set ds = CreateObject("Scripting.FileSystemObject")
pdosya = "C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt"
pds = ds.FileExists(pdosya)
If pds = True Then CreateObject("Shell.Application").Open pdosya Else MsgBox pdosya & " Dosyası Bulunamadı !", vbCritical, "Hata !"
End Sub
Private Sub CommandButton31_Click() 'ceviri dosyaları
On Error Resume Next
ListViewP21.ListItems.Clear: ListViewP31.ListItems.Clear: TextBoxP21 = ""
Dim itmX As listItem
  Dim dosya
  dosya = dir("C:\Belgelerim\CEMEX\Çeviri Dosyaları\Kelime Dosyaları\*.xls*")
Do While dosya <> ""
Set itmX = ListViewP21.ListItems.Add(, , dosya)
itmX.Bold = True
   dosya = dir
Loop
ListViewP21.StartLabelEdit: ListViewP21.Refresh
End Sub
Private Sub ListViewp21_Click()
If ListViewP21.ListItems.Count > 0 Then TextBoxP21 = ListViewP21.SelectedItem.Text
End Sub
Private Sub CommandButton32_Click()
On Error Resume Next
If ListViewP21.ListItems.Count < 1 Or TextBoxP21 = "" Then Exit Sub
Application.ScreenUpdating = False
Dim ds, f
Set ds = CreateObject("Scripting.FileSystemObject")
Set f = ds.GetFolder("C:\Belgelerim\Cemex\Çeviri Dosyaları\Kelime Dosyaları")
Workbooks.Open fileName:=f & "\" & TextBoxP21
Application.Windows(TextBoxP21.Text).Visible = False
ListViewP31.ListItems.Clear: ListViewP31.ColumnHeaders.Clear
cls = Workbooks(TextBoxP21.Text).ActiveSheet.Cells(1, 256).End(xlToLeft).Column
For k = 1 To cls
Call ListViewP31.ColumnHeaders.Add(k, , Workbooks(TextBoxP21.Text).ActiveSheet.Cells(1, k), 95)
Next
Dim son As Integer
son = Workbooks(TextBoxP21.Text).ActiveSheet.Range("A" & "65536").End(xlUp).row
For n = 2 To son
    If Workbooks(TextBoxP21.Text).ActiveSheet.Range("A" & n).Value <> "" Then
    m = ListViewP31.ListItems.Count
    Call ListViewP31.ListItems.Add(m + 1, , Workbooks(TextBoxP21.Text).ActiveSheet.Range("A" & n))
    Call ListViewP31.ListItems(m + 1).ListSubItems.Add(1, , Workbooks(TextBoxP21.Text).ActiveSheet.Range("B" & n))
    End If
Next n
Workbooks(TextBoxP21.Text).Close False
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton33_Click()
On Error Resume Next
For i = 1 To ListViewP31.ListItems.Count
a = ListViewP31.ListItems(i + 1): b = ListViewP31.ListItems(i + 1).ListSubItems(1)
ActiveSheet.Cells.Replace What:=a, Replacement:=b, LookAt:=xlPart
Next
End Sub
Private Sub CommandButton194_Click()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then
Dim ds, a
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists("C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif-İmalat V1.2.xltx")
If a <> True Then msg = MsgBox("Cemex\Yeni Teklif Şablonları\Yeni Teklif-İmalat V1.2.xltx" & vbCr & "bu dosya mevcut değil!", vb, "scngnr@hotmail.com"): Exit Sub
'--
ActiveWorkbook.Worksheets("Sayfa1").Select
dss1 = ActiveWorkbook.Name
Columns("A:E").Select: Selection.Copy
Range("A1").Select
Workbooks.Open "C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif-İmalat V1.2.xltx"
dss2 = ActiveWorkbook.Name
Worksheets("Sayfa1").Select
Range("A1").Select
ActiveSheet.Paste: Application.CutCopyMode = False
Range("A1").Select
Workbooks(dss2).Activate
For n = 1 To ActiveWorkbook.names.Count
    ActiveWorkbook.names(1).Delete
Next n
Columns("E:J").ColumnWidth = 13.9: Columns("F:J").HorizontalAlignment = xlRight:
Columns("F:J").Font.Name = "Arial": Columns("F:J").Font.Size = 9
Columns("F:F").NumberFormat = "#,##0"
Columns("F:J").NumberFormat = "#,##0.00"
Columns("H:H").NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
Columns("I:I").NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
'Genel Biçimlemeler'--
Dim Y As Integer
Y = Cells(1, 2).End(xlDown).row: If Y > 12000 Then Y = 3
Range("F" & 1 & ":J" & Y).Borders.LineStyle = xlContinuous
Range("F" & 1 & ":J" & 1).Interior.ColorIndex = 19
'--
  If ListBox3.ListCount > 0 Then
  Range("E1") = "Teklif Miktarı": Range("F1") = ListBox3.List(0, 1): Range("G1") = ListBox3.List(1, 1)
  Range("H1") = ListBox3.List(2, 1): Range("I1") = ListBox3.List(3, 1): Range("J1") = ListBox3.List(4, 1)
  Else
  Range("E1") = "Teklif Miktarı": Range("F1") = "Planlanan": Range("G1") = "Stok/Teslim"
  Range("H1") = "Kalan": Range("I1") = "Fazla-İade": Range("J1") = "Kontrol"
  End If
End If
For i = 2 To Y
If Not Range("E" & i) = "" Then
Range("H" & i).FormulaR1C1 = "=IF(RC[-2]-RC[-1]<=0,""-"",RC[-2]-RC[-1])"
Range("I" & i).FormulaR1C1 = "=IF(RC[-2]-RC[-3]<=0,""-"",RC[-2]-RC[-3])"
End If
Next
'--
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
Unload Me
Range("F1").Select
End Sub
Private Sub CommandButton194xxx_Click()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then
'--
ActiveWorkbook.Worksheets("Sayfa1").Select
For i = 1 To Worksheets.Count
If Worksheets(i).Name = "Notlar" Then Sheets(Array("Notlar", "Sayfa1")).Copy: GoTo git1
Next i
Sheets("Sayfa1").Copy
git1:
ActiveWorkbook.Worksheets("Sayfa1").Select
With ActiveSheet
.Parent.VBProject.VBComponents(.CodeName).Properties("_CodeName") = "TM1"
End With
ActiveWorkbook.Worksheets("Sayfa1").Select
With ActiveSheet
.Parent.VBProject.VBComponents(.CodeName).Properties("_CodeName") = "TM1"
End With
'--
For n = 1 To ActiveWorkbook.names.Count
    ActiveWorkbook.names(1).Delete
Next n
Columns("K:X").Delete: Columns("F:J").Clear
Columns("E:J").ColumnWidth = 13.9: Columns("F:J").HorizontalAlignment = xlRight:
Columns("F:J").Font.Name = "Arial": Columns("F:J").Font.Size = 9
Columns("F:F").NumberFormat = "#,##0"
Columns("F:J").NumberFormat = "#,##0.00"
Columns("H:H").NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
Columns("I:I").NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
'Genel Biçimlemeler'--
Dim Y As Integer
Y = Cells(1, 2).End(xlDown).row: If Y > 12000 Then Y = 3
Range("F" & 1 & ":J" & Y).Borders.LineStyle = xlContinuous
Range("F" & 1 & ":J" & 1).Interior.ColorIndex = 19
'--
  If ListBox3.ListCount > 0 Then
  Range("E1") = "Teklif Miktarı": Range("F1") = ListBox3.List(0, 1): Range("G1") = ListBox3.List(1, 1)
  Range("H1") = ListBox3.List(2, 1): Range("I1") = ListBox3.List(3, 1): Range("J1") = ListBox3.List(4, 1)
  Else
  Range("E1") = "Teklif Miktarı": Range("F1") = "Planlanan": Range("G1") = "Stok/Teslim"
  Range("H1") = "Kalan": Range("I1") = "Fazla-İade": Range("J1") = "Kontrol"
  End If
End If
For i = 2 To Y
If Not Range("E" & i) = "" Then
Range("H" & i).FormulaR1C1 = "=IF(RC[-2]-RC[-1]<=0,""-"",RC[-2]-RC[-1])"
Range("I" & i).FormulaR1C1 = "=IF(RC[-2]-RC[-3]<=0,""-"",RC[-2]-RC[-3])"
End If
Next
'--
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
Unload Me
Range("F1").Select
End Sub
Private Sub CommandButton195_Click()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If CB1.Value = False And CB2.Value = False And CB3.Value = False And _
CB4.Value = False And CB5.Value = False And CB6.Value = False Then Exit Sub
ActiveSheet.ShowAllData
Dim Y As Integer
Dim by As Integer
Dim p, q As Byte
Y = Range("B65536").End(xlUp).row
ProgressBarTHB.Visible = True: ProgressBarTHB.Max = Y
p = 0
For i = 2 To Y
q = 0
ProgressBarTHB.Value = i
If Range("B" & i) = "BÖLÜM ADI/NO:" Then by = i
Range("A" & i & ":J" & i).Borders.LineStyle = xlContinuous
Range("A" & i & ":J" & i).Interior.Pattern = xlNone
If Not Range("D" & i) = "" Then
Range("A" & i & ":J" & i).Font.Size = 9
Range("A" & i & ":J" & i).Font.Name = "Arial"
Range("F" & i & ":J" & i).NumberFormat = "#,##0.00"
Range("H" & i).NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
Range("I" & i).NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then
Range("H" & i).FormulaR1C1 = "=IF(RC[-2]-RC[-1]<=0,""-"",RC[-2]-RC[-1])"
Range("I" & i).FormulaR1C1 = "=IF(RC[-2]-RC[-3]<=0,""-"",RC[-2]-RC[-3])"
End If
'satır
 If CB1.Value = True Then If Range("G" & i) = Range("F" & i) Then Range("A" & i & ":J" & i).Interior.Color = 15594477: q = 1
 If CB2.Value = True Then If Range("F" & i) > Range("G" & i) Then Range("A" & i & ":J" & i).Interior.Color = 14540287: q = 1
 If CB3.Value = True Then If Range("G" & i) > Range("F" & i) Then Range("A" & i & ":J" & i).Interior.Color = 14548991: q = 1
'hücre
 If CB4.Value = True Then If Range("F" & i) > Range("E" & i) Then Range("F" & i).Interior.Color = 16764415: q = 1
 If CB5.Value = True Then If Range("F" & i) > Range("G" & i) Then Range("H" & i).Interior.Color = 13487615: q = 1
 If CB6.Value = True Then If Range("G" & i) > Range("F" & i) Then Range("I" & i).Interior.Color = 16764365: q = 1
'bölümadı
If p = 0 Then: If q = 1 Then p = 1
 If CBP1.Value = True Then
 If q = 1 Then Range("B" & by).Interior.Color = 15594477: Range("B" & i).Interior.Color = 15594477
 End If
'---
End If
Next
 If CBP1.Value = True And p = 1 Then
    ActiveSheet.Range("B" & Y).AutoFilter Field:=2, Criteria1:=RGB(237, 243, 237), Operator:=xlFilterCellColor
 End If
ProgressBarTHB.Value = 1
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
 If p = 1 Then MsgBox ("Eksik/Fazla Ürün Mevcut! "), vbInformation, "scngnr@hotmail.com"
 If p = 0 Then MsgBox ("Veri Bulunamadı! "), vbInformation, "scngnr@hotmail.com"
End Sub
Private Sub CommandButton196_Click()
On Error Resume Next
Application.ScreenUpdating = False
Dim Y As Integer
Y = Range("B65536").End(xlUp).row
ProgressBarTHB.Visible = True: ProgressBarTHB.Value = 1: ProgressBarTHB.Max = Y
For i = 2 To Y
ProgressBarTHB.Value = i
Range("A" & i & ":J" & i).Borders.LineStyle = xlContinuous
'If Not Range("D" & i) = "" Then
Range("A" & i & ":J" & i).Interior.Pattern = xlNone
'End If
Next
ActiveSheet.ShowAllData
ProgressBarTHB.Value = 1
Application.ScreenUpdating = True
End Sub
Sub imalatdosya1()
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Sayfa Düzenleme\İmalat Dosyası.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "İmalat Dosyası.txt / Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBox3.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBox3.AddItem ayır(i)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBox3.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
    kt = 1
Close #Ert
End Sub
Private Sub Image44_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Sayfa Düzenleme\İmalat Dosyası.txt"
End Sub
Private Sub Image441_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021
On Error Resume Next
Call imalatdosya1
End Sub