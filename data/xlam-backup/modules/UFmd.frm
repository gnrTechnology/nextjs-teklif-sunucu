Public bfyt As String
Dim kur As String
Dim ad As String
Private Sub CheckBoxP32_Click()
CheckBoxP31.Enabled = True
End Sub
Private Sub CheckBoxP35_Click()
CheckBoxP31.Enabled = True
End Sub
Private Sub UserForm_Initialize() 'form yükleme xxxxxxxxxxxxxxxxxx
On Error Resume Next
'--
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then CheckBoxP31.Visible = False: OptionButtonP32.Enabled = False
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
ProgressBarP21.Visible = False
ToolbarP3.ImageList = ImageList2
ToolbarP3.Buttons.Item(1).Image = ImageList2.ListImages.Item(1).Index
ToolbarP3.Buttons.Item(2).Image = ImageList2.ListImages.Item(2).Index
ToolbarP3.Buttons.Item(3).Image = ImageList2.ListImages.Item(3).Index
ToolbarP3.Buttons.Item(4).Image = ImageList2.ListImages.Item(4).Index
Call malzkod
End Sub
Private Sub ToolbarP3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
UFmd.Height = 400
MultiPageP2.Value = 0
Windows(dt).Activate
Case 2
MultiPageP2.Value = 1
Case 3
If UFmd.Height >= 350 Then
UFmd.Height = UFmd.Height - UFmd.InsideHeight + ToolbarP3.Height
Else
UFmd.Height = 400
End If
Case 4
Unload Me
End Select
End Sub
Private Sub ToolbarP3_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt"
End Sub
Sub malzkod()
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
If Left(Rky, 1) = "#" Then GoTo git1
        ListBoxMK.AddItem ayır(i)
        TBM1.Text = TBM1.Text & "," & ayır(i)
'--
'kaçıncı = InStr(1, Rky, ";")
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxMK.List(satır - 1, n) = ayır(i + n)
Next n
        satır = satır + 1
git1:
    Loop
If Left(TBM1.Text, 1) = "," Then TBM1.Text = Right(TBM1.Text, Len(TBM1.Text) - 1)
If Right(TBM1.Text, 1) = "," Then TBM1.Text = Left(TBM1.Text, Len(TBM1.Text) - 1)
    Close #Ert
End Sub
Private Sub CommandButtonP31_Click()
son = ListBoxP31.ListCount
For i = 0 To son - 1
    If ListBoxP31.Selected(i) = False Then
    ListBoxP31.Selected(i) = True
    End If
Next
End Sub
Private Sub CommandButtonP32_Click()
son = ListBoxP31.ListCount
For i = 0 To son - 1
    If ListBoxP31.Selected(i) = True Then
    ListBoxP31.Selected(i) = False
    End If
Next
End Sub
Private Sub CheckBoxP33_Click()
Call tümmalzeme
End Sub
Private Sub CheckBoxP34_Click()
Call tümmalzeme
End Sub
Sub tümmalzeme()
If CheckBoxP33.Enabled = False Or CheckBoxP34.Enabled = False Then Exit Sub
If CheckBoxP33.Value = False And CheckBoxP34.Value = False Then ListBoxP31.Clear: Exit Sub
Windows(dt).Activate: Sheets("Sayfa1").Select
ListBoxP31.Clear
Dim son As Integer
son = Range("D65536").End(xlUp).row
X = 0
If CheckBoxP33.Value = True And CheckBoxP34.Value = False Then
For n = 2 To son
If WorksheetFunction.CountIf(Range("D2:D" & n), Cells(n, 4).Value) = 1 And Cells(n, 3).Value <> "" Then
    If Left(Cells(n, "a"), 3) <> "PP-" Then
    If Left(Cells(n, "a"), 4) = "PM-M" Then
       If Left(Cells(n, "a"), 5) <> "PM-MB" Then GoTo atla2
    End If
    ListBoxP31.AddItem Cells(n, 4)
    ListBoxP31.List(X, 1) = WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("E1:E65536")) & " Ad. "
    ListBoxP31.List(X, 2) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("O1:O65536")), "#,##0.00")
    ListBoxP31.List(X, 3) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("P1:P65536")), "#,##0.00")
    X = X + 1
    End If: End If
atla2:
Next n
End If
If CheckBoxP33.Value = False And CheckBoxP34.Value = True Then
For n = 2 To son
If WorksheetFunction.CountIf(Range("D2:D" & n), Cells(n, 4).Value) = 1 And Cells(n, 3).Value <> "" Then
    If Left(Cells(n, "a"), 3) = "PP-" Then
    ListBoxP31.AddItem Cells(n, 4)
    ListBoxP31.List(X, 1) = WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("E1:E65536")) & " Ad. "
    ListBoxP31.List(X, 2) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("O1:O65536")), "#,##0.00")
    ListBoxP31.List(X, 3) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("P1:P65536")), "#,##0.00")
    X = X + 1
    End If: End If
Next n
End If
If CheckBoxP33.Value = True And CheckBoxP34.Value = True Then
For n = 2 To son
If WorksheetFunction.CountIf(Range("D2:D" & n), Cells(n, 4).Value) = 1 And Cells(n, 3).Value <> "" Then
    If Left(Cells(n, "a"), 3) = "PM-" Then
    If Left(Cells(n, "a"), 5) <> "PM-MB" Then GoTo atla1
    End If
    ListBoxP31.AddItem Cells(n, 4)
    ListBoxP31.List(X, 1) = WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("E1:E65536")) & " Ad. "
    ListBoxP31.List(X, 2) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("O1:O65536")), "#,##0.00")
    ListBoxP31.List(X, 3) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("P1:P65536")), "#,##0.00")
    X = X + 1
    End If
atla1:
Next n
End If
AlignListColumn ListBoxP31, 2, True
AlignListColumn ListBoxP31, 3, True
End Sub
Sub saltmalzeme()
Windows(dt).Activate: Sheets("Sayfa1").Select
ListBoxP31.Clear
Dim son As Integer
son = Range("D65536").End(xlUp).row
X = 0
kd1 = "PP-" 'PANO
kd2 = "PM-" 'işçilik,bara,ambalaj vs.
For n = 2 To son
    If WorksheetFunction.CountIf(Range("D2:D" & n), Cells(n, 4).Value) = 1 And Cells(n, 3).Value <> "" Then
    skd = Left(Cells(n, "a"), 3)
    If skd = kd1 Or skd = kd2 Then GoTo atla
    ListBoxP31.AddItem Cells(n, 4)
    ListBoxP31.List(X, 1) = WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("E1:E65536")) & " Ad. "
    ListBoxP31.List(X, 2) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("O1:O65536")), "#,##0.00")
    ListBoxP31.List(X, 3) = Format(WorksheetFunction.SumIf(Range("D1:D65536"), Cells(n, 4), Range("P1:P65536")), "#,##0.00")
    X = X + 1
   End If
atla:
Next n
AlignListColumn ListBoxP31, 2, True
AlignListColumn ListBoxP31, 3, True
End Sub
Private Sub CheckBoxP321_Click()
If CheckBoxP321.Value = True Then
ListBoxP31.MultiSelect = 0
CheckBoxP32.Value = 0: CheckBoxP32.Enabled = False: CheckBoxP35.Value = 0: CheckBoxP35.Enabled = False
Else
ListBoxP31.MultiSelect = 1
CheckBoxP32.Enabled = True: CheckBoxP35.Enabled = True
End If
If CheckBoxP321.Value = True Then CheckBoxP31.Enabled = False: CheckBoxP31.Value = False
End Sub
Private Sub OptionButtonP31_Click()
On Error Resume Next
OptionButtonP32.Value = False
ListBoxP31.Clear
Frame27.Caption = "Teklifde Bulunan Malzemeler"
Label2.Caption = " Marka": Label4.Caption = "  Mlz. Liste Top." & " (TL)": Label5.Caption = "  Mlz. Net Top." & " (TL)"
CheckBoxP31.Enabled = True: CheckBoxP31.Value = True
CommandButtonP30.Enabled = True
CheckBoxP32.Value = False: CheckBoxP32.Enabled = False: CheckBoxP35.Value = False: CheckBoxP35.Enabled = False
CheckBoxP321.Value = False: CheckBoxP321.Enabled = False: CheckBoxP33.Enabled = True: CheckBoxP34.Enabled = True
Windows(dt).Activate
'Sheets("Malzeme Listesi").Visible = -1
'Windows(dt).Activate
Sheets("Sayfa1").Select '
TextBoxP33.Value = 0
TextBoxP32.Value = 0
'Call saltmalzeme
End Sub
Private Sub OptionButtonP32_Click()
On Error Resume Next
Windows(dt).Activate: Sheets("Sayfa1").Select
OptionButtonP31.Value = False
ListBoxP31.Clear
'Sheets("Malzeme Listesi").Visible = 0
CheckBoxP31.Enabled = True: CheckBoxP31.Value = True
CheckBoxP33.Enabled = False: CheckBoxP33.Value = False
CheckBoxP34.Enabled = False: CheckBoxP34.Value = False
CheckBoxP321.Enabled = True
CommandButtonP30.Enabled = False
CheckBoxP32.Enabled = True: CheckBoxP35.Enabled = True
TextBoxP32.Value = 0: TextBoxP33.Value = 0
pkur = " " & Right(Range("Tpbr"), 5)
Dim ad As Integer
Dim son As Integer
son = Range("B65536").End(xlUp).row
X = 0
For n = 1 To son
    If Cells(n, 2) = "BÖLÜM ADI/NO:" Then
    ListBoxP31.AddItem Cells(n, 3)
    If Cells(n, 5) > 0 Then ad = Cells(n, 5) / 1 Else ad = 1
    ListBoxP31.List(ListBoxP31.ListCount - 1, 1) = ad & " Ad. "
    ListBoxP31.List(ListBoxP31.ListCount - 1, 4) = Cells(n, 2).row
    adt = adt + ad
    X = X + 1
   End If
    If Cells(n + 1, 2) = "BÖLÜM TOPLAMI:" Then
        'ListBoxP31.List(ListBoxP31.ListCount - 1, 1) = Format(Cells(n + 1, 24), "#,##0.00")
        ListBoxP31.List(ListBoxP31.ListCount - 1, 2) = Format(Cells(n + 1, 24) / ad, "#,##0.00")
        ListBoxP31.List(ListBoxP31.ListCount - 1, 3) = Format(Cells(n + 1, 24), "#,##0.00")
        ListBoxP31.List(ListBoxP31.ListCount - 1, 5) = Cells(n + 1, 2).row
    End If
Next n
Frame27.Caption = "Teklifte Bulunan Panolar" & " '" & adt & " Adet "
Label2.Caption = " Pano Adı": Label4.Caption = "  Birim Fiyat" & pkur: Label5.Caption = "  Toplam Fiyat" & pkur
TextBoxPT3.Value = ListBoxP31.ListCount
AlignListColumn ListBoxP31, 2, True
AlignListColumn ListBoxP31, 3, True
End Sub
Private Sub CommandButtonP33_Click()
Application.Calculation = xlManual: Application.ScreenUpdating = False
On Error Resume Next
Windows(dt).Activate
If CheckBoxP321.Value = True Then GoTo git:
son = ListBoxP31.ListCount
For i = 0 To son - 1
    If ListBoxP31.Selected(i) = True Then GoTo git:
Next
msg = MsgBox("   Listelenecek ürünü seçmediniz ! ", vbYes, "scngnr@hotmail.com")
GoTo bitti:
git:
If OptionButtonP32.Value = True Then
If CheckBoxP32.Value = False And CheckBoxP35.Value = False And CheckBoxP321.Value = False Then GoTo bitti
Call MA
If CheckBoxP321.Value = True Then Call Makro122 Else Call Makro121
'UserFormP1.Show
'UserForm2.Show
GoTo bitti:
End If
If OptionButtonP31.Value = True And CheckBoxP31.Value = True Then
'--
Call OMA: Call MA
Call mlzfullaktar
'UserForm2.Show
GoTo bitti:
End If
If OptionButtonP31.Value = True And CheckBoxP31.Value = False Then
Call OMA
Call MA
Call mlzfullaktar
'UserForm2.Show
GoTo bitti:
End If
bitti:
ad = ""
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub OMA()
On Error Resume Next
If OptionButtonP32.Value = True Then Exit Sub
Windows(dt).Activate: Sheets("Sayfa1").Select: ActiveSheet.ShowAllData
TextBoxP31.Value = WorksheetFunction.CountA(Range("A:A"), xlDown) - 2: TextBoxP32.Value = 0: TextBoxP33.Value = 0
listson = ListBoxP31.ListCount
For i = 0 To listson - 1
    If ListBoxP31.Selected(i) = True Then
    TextBoxP32.Value = TextBoxP31.Value - TextBoxP33.Value
    TextBoxP33.Value = WorksheetFunction.CountIf(Range("D:D"), (ListBoxP31.List(i, 0))) + TextBoxP33.Value
End If
Next
End Sub
Sub MA()
On Error Resume Next
Dim sf As Worksheet
If OptionButtonP32.Value = True Then
If CheckBoxP321.Value = True Then ad = "Pano İcmal" Else ad = "Pano Listesi"
GoTo git:
End If
son = ListBoxP31.ListCount
ads = ""
For i = 0 To son
    If ListBoxP31.Selected(i) = True Then
    ad = ad & tire & Left(ListBoxP31.List(i), 3): tire = "-"
    'ad = ad & tire2 & ListBoxP31.List(i, 0): tire2 = "-"
    End If
Next
If Len(ad) > 30 Then ad = "M.Listesi+"
git:
Dim m
        For Each sf In Worksheets
        If sf.Name = ad Then GoTo var
        Next sf
        m = Sheets.Count + 1
        Set NewSheet = Worksheets.Add
        NewSheet.Name = ad
        Sheets(ad).Move After:=Sheets(m)
var:
        Sheets(ad).Select
        Call ChangeCodeName
        'ActiveSheet.PageSetup.PrintArea = ""
End Sub
Sub ChangeCodeName()
If Left(ActiveSheet.CodeName, 6) = "Icmal_" Then Exit Sub
ActiveSheet.Parent.VBProject.VBComponents(ActiveSheet.CodeName).Properties("_CodeName") = "Icmal_" & Sheets.Count + 1
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
Sub malzemegrup2() '2014 Kümülatif Malzeme Fiyat listesini sıraya koyar ve başlık atar.
On Error Resume Next
Dim Y, a, b As Single
    a = WorksheetFunction.CountA(Range("B:B"), xlDown)
        If a = 0 Then
        Exit Sub
        Else
'''''
    k = 2
ProgressBarP21.Value = 0: ProgressBarP21.Max = a
Do Until k = a
ProgressBarP21.Value = k
    Cells(k, 1) = Replace(Cells(k, 1), "-grup", "") ' -grup silme
    Cells(k, 1) = Replace(Cells(k, 1), "-devre", "") ' -devre silme
    'Cells(k, 1) = Left(Cells(k, 1), 10) ' son rakam bakılacak harf sayısı
    'sn = Cells(k, 1)
    'Modüler ürünler
    'If sn = "FA" Or sn = "FI" Or sn = "FM" Or sn = "KM" Or sn = "US" Or sn = "KS" Then Cells(k, 1) = "MD": GoTo sonraki:
sonraki:
    k = k + 1
Loop
'''''
'tüm malzemeleri seçer
     ActiveCell.CurrentRegion.Offset(1, 0).Resize(ActiveCell.CurrentRegion.Rows.Count - 1, ActiveCell.CurrentRegion.Columns.Count).Select
'seçilen malzemeleri sıralar eski
    'Selection.Sort key1:=Range("A2"), Key2:=Range("B2"), order1:=xlAscending, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
'''''seçilen malzemeleri sıralar yeni
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=(TBM1), DataOption:=xlSortNor
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Selection
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'''''
    b = 0
    Range("A1").Select
'''''
Y = 2
'ürünler1--
Do Until Y >= a
 ProgressBarP21.Value = Y
Set sn = Cells(Y, 1)
sayfakod = Replace(sn, "-auto", ""): sayfakod = Replace(sayfakod, "-grup", ""): sayfakod = Replace(sayfakod, "-devre", "")
ayır = Split(sayfakod, "."): sayfakod = ayır(0)
mugson = ListBoxMK.ListCount
  For X = 0 To mugson - 1
  kodL = ListBoxMK.List(X, 0)
  If Left(kodL, 1) = "#" Then GoTo git
  If sayfakod = kodL Then mlzb = ListBoxMK.List(X, 1): GoTo git
  Next X
  mlzb = "DİĞER ÜRÜNLER"
git:
sn.EntireRow.Insert Shift:=xlDown: Cells(Y, 1) = mlzb
Y = Y + 1
a = WorksheetFunction.CountA(Range("A:A"), xlDown)
     Do While Y < a
If Cells(Y + 1, 1) <> sn Then Exit Do
     Y = Y + 1
     Loop
Y = Y + 1
Loop
'''''
End If
End Sub
Sub malzemebicim2() '201' biçimlendirmeler
On Error Resume Next
     Range("B2").Select
'e = WorksheetFunction.CountA(Range("B:B"), xlDown)
e = Range("B65536").End(xlUp).row
Dlsonkolon = Cells(1, 256).End(xlToLeft).Column ' sablon kolon sayısını verir
If Dlsonkolon > 5 Then e = e + 1
Y = 1
z = 0
ProgressBarP21.Value = 0: ProgressBarP21.Max = e
Range(Cells(1, 1), Cells(1, Dlsonkolon)).Borders.LineStyle = xlContinuous
Range(Cells(1, 1), Cells(1, Dlsonkolon)).Interior.Color = 13434828
Range(Cells(1, 1), Cells(1, Dlsonkolon)).Font.Bold = False
Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.Size = 10
Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.Name = "Arial"
Y = Y + 1
Do Until Y > e
ProgressBarP21.Value = Y
'''''
If OptionButtonP32.Value = True Then
    If Cells(Y, 2) = "BÖLÜM ADI/NO:" Then
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Borders.LineStyle = xlContinuous
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Borders(xlInsideVertical).LineStyle = xlNone
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Interior.ColorIndex = 2
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.Bold = True
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.ColorIndex = 11
    Y = Y + 1
    End If: End If
'''''
    If Cells(Y, 2) = "" Then
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Borders.LineStyle = xlContinuous
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Borders(xlInsideVertical).LineStyle = xlNone
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Interior.ColorIndex = 2
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.Bold = True
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.ColorIndex = 11
    Else:
    If OptionButtonP32.Value = False Then Cells(Y, 1) = z + 1
    Range("A" & Y & ":E" & Y).Borders.LineStyle = xlContinuous
   Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Borders.LineStyle = xlContinuous
    Cells(Y, 1).HorizontalAlignment = xlLeft
    z = z + 1
    End If
    Range("A" & Y & ":E" & Y).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.Name = "Arial"
    Range(Cells(Y, 1), Cells(Y, Dlsonkolon)).Font.Size = 9
    Range("B" & Y).RowHeight = 14.4
    Y = Y + 1
    Loop
    If CheckBoxP321.Value = False Then Range(Cells(1, 1), Cells(e, 4)).HorizontalAlignment = xlLeft
    'Application.PrintCommunication = True
    'ActiveSheet.PageSetup.PrintArea = "$A:$E"
    'Application.PrintCommunication = True
ProgressBarP21.Visible = False
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
msg = MsgBox("    Seçilen ürünler sayfaya aktarıldı. ", vb, "scngnr@hotmail.com")
End Sub
Sub mlzfullaktar() ' Malzeme listesi fiyatlı ve iskontolu 22.09.2020
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlManual
'printreset
Dim Y, z, k, a, i, p As Single
Dim l, s As Double
Dim kt, ky, b, msg As String
'''''
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then CheckBoxP31.Value = False 'malzeme dosyası
    p = TextBoxP33.Value
    a = WorksheetFunction.CountA(Range("A:A"), xlDown) - 1
    If a > 0 Then
    msg = MsgBox("    Mevcut  liste  silinerek  yeni   liste   düzenlenecek ", vbYesNo, "scngnr@hotmail.com ")
    If msg = vbYes Then
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.UsedRange.Delete
    a = 0
ProgressBarP21.Visible = True: ProgressBarP21.Value = 0: ProgressBarP21.Max = p
ali:
For X = 1 To 5
    Cells(1, X).Value = Sheets("Sayfa1").Cells(1, X).Value
Next
''''
    Range("A1").FormulaR1C1 = "S.No"
    Range("A1:E1").Font.Size = 10
    Range("A1").RowHeight = 18.6
    Range("A1").ColumnWidth = 4.56: Range("B1").ColumnWidth = 18: Range("C1").ColumnWidth = 58: Range("D1").ColumnWidth = 11
    Range("E1").ColumnWidth = 8
If OptionButtonP31.Value = True And CheckBoxP31.Value = True Then 'isk.
    Range("F1").ColumnWidth = 14: Range("G1").ColumnWidth = 6: Range("H1").ColumnWidth = 14: Range("I1").ColumnWidth = 15
    Cells(1, 6).Value = Sheets("Sayfa1").Cells(1, 6).Value
    Cells(1, 7).Value = Sheets("Sayfa1").Cells(1, 7).Value
    Cells(1, 8).Value = Sheets("Sayfa1").Cells(1, 11).Value
    Cells(1, 9).Value = Sheets("Sayfa1").Cells(1, 16).Value
End If
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then 'malzeme dosyası
Columns("E:J").ColumnWidth = 13.9
Columns("E:J").NumberFormat = "#,##0.00": Columns("F:J").HorizontalAlignment = xlRight
Columns("H:H").NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
Columns("I:I").NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
    For X = 1 To 5
    Cells(1, X + 5).Value = Sheets("Sayfa1").Cells(1, X + 5).Value
    Next
End If
''''
    Y = 0
    z = 0
    l = 0
 ProgressBarP21.Visible = True: ProgressBarP21.Value = 0: ProgressBarP21.Max = p
Do Until Y = p
yok:
    If Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0) = "" Then
       If Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0) = "" And Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 22) <> "" Then _
       Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0) = "XX": Exit Do
       Do Until Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0) <> ""
       z = z + 1
       s1b = WorksheetFunction.CountA(Sheets("Sayfa1").Range("B:B"), xlDown) ' sayfa1 Bde son hücre
       If Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0).row > s1b Then GoTo s1son: ' DÖNGÜ DURDURMA 10.06.2015
       Loop
    Else:
       SayfaMarka = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 3)
       i = 0
       son = ListBoxP31.ListCount
       Do Until son = i
       formmarka = (ListBoxP31.List(i, 0))
       If ListBoxP31.Selected(i) = True And formmarka = SayfaMarka Then GoTo devam:
       i = i + 1
       Loop
       z = z + 1
ProgressBarP21.Value = Y
GoTo yok:
devam:
'''''
    kt = Format((Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1)), "@") 'formatı metin yapar
    k = 0
       Do Until k = Y + z + 1
       ky = Format((Sheets("Sayfa1").Range("A1").Offset(k + 1, 1)), "@") 'formatı metin yapar
       If ky = kt Then
       Exit Do
       Else:
       k = k + 1
       End If
       Loop
     If k >= Y + z Then ' önceki satırları tarıyor ve z ile boşlukları da sayıyor .
'''''
       For X = 1 To 5
       Cells(a + 2, X).Value = Sheets("Sayfa1").Cells(k + 2, X).Value
       Next
       Cells(a + 2, 5).FormulaR1C1 = "=SUMIF(Sayfa1!C[-3],RC[-3],Sayfa1!C)" 'malzeme kodu için
       Cells(a + 2, 5).NumberFormat = "#,##0"
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then 'malzeme dosyası
    Cells(a + 2, 6).FormulaR1C1 = "=SUMIF(Sayfa1!C[-4],RC[-4],Sayfa1!C)"
    Cells(a + 2, 7).FormulaR1C1 = "=SUMIF(Sayfa1!C[-5],RC[-5],Sayfa1!C)"
    Cells(a + 2, 8).FormulaR1C1 = "=SUMIF(Sayfa1!C[-6],RC[-6],Sayfa1!C)"
    Cells(a + 2, 9).FormulaR1C1 = "=SUMIF(Sayfa1!C[-7],RC[-7],Sayfa1!C)"
    Cells(a + 2, 10).FormulaR1C1 = "=SUMIF(Sayfa1!C[-8],RC[-8],Sayfa1!C)"
End If
'''''
If OptionButtonP31.Value = True And CheckBoxP31.Value = True Then 'isk.
    Cells(a + 2, 6).FormulaR1C1 = Sheets("Sayfa1").Cells(k + 2, 6).Value
    Cells(a + 2, 6).NumberFormat = Sheets("Sayfa1").Cells(k + 2, 6).NumberFormat
    Cells(a + 2, 7).FormulaR1C1 = Format(Sheets("Sayfa1").Cells(k + 2, 7).Value, "0%")
    kur = ""
    If Cells(a + 2, 6).NumberFormat = "#,##0.00 [$€-1]" Then kur = "*Eur" ' Sayfa3 de € kuru
    If Cells(a + 2, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then kur = "*Usd" ' Sayfa3 de $ kuru
                
    Cells(a + 2, 8).FormulaR1C1 = "=(RC[-2]-RC[-2]*RC[-1])" & kur ' Formül yaz
    Cells(a + 2, 8).NumberFormat = "#,##0.00"
    Cells(a + 2, 9).FormulaR1C1 = "=RC[-4]*RC[-1]" ' Formül yaz
    Cells(a + 2, 9).NumberFormat = "#,##0.00"
End If
    a = a + 1
    End If
'''''
    Y = Y + 1
    End If
Loop
Else
Exit Sub
End If
Else
GoTo ali
End If
If OptionButtonP31.Value = True And CheckBoxP31.Value = True Then ''''İSK
Cells(a + 2, 8).FormulaR1C1 = "TOPLAM :"

Cells(a + 2, 9) = "=Sum(I2:I" & a + 1 & ")" 'toplam alma formüllü
Range("A1").Select
End If '''''
s1son:

Call malzemegrup2
Call malzemebicim2
ProgressBarP21.Value = 0
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub Makro121() ' Pano listesi
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlManual
Dim Y, z, k, a, p, o, m, n As Single
Dim s As Double
Dim msg As String
'''''
        Sheets(ad).Visible = xlSheetVisible
        Sheets(ad).Select
        a = WorksheetFunction.CountA(Range("B:B"), xlDown) - 1
        If a > 0 Then
        msg = MsgBox("    Mevcut  liste  silinerek  yeni   liste   düzenlenecek ", vbYesNo, "scngnr@hotmail.com ")
        If msg = vbYes Then
        Cells.Clear
        Else
        GoTo bitti2
        End If
        End If
''''
For X = 1 To 5
    Sheets(ad).Cells(1, X).Value = Sheets("Sayfa1").Cells(1, X).Value
Next
''''
        Range("A1").FormulaR1C1 = "S.No"
        Range("A1:E1").Font.Size = 10
        Range("A1").RowHeight = 18.6
        Range("A1").ColumnWidth = 4.56
        Range("B1").ColumnWidth = 18
        Range("C1").ColumnWidth = 58
        Range("D1").ColumnWidth = 11
        Range("E1").ColumnWidth = 8
If OptionButtonP32.Value = True And CheckBoxP31.Value = True Then ''''İSK
        Range("F1").ColumnWidth = 14
        Range("G1").ColumnWidth = 6
        Range("H1").ColumnWidth = 14
        Range("I1").ColumnWidth = 15
        Cells(1, 6).Value = Sheets("Sayfa1").Cells(1, 6).Value
        Cells(1, 7).Value = Sheets("Sayfa1").Cells(1, 7).Value
        Cells(1, 8).Value = Sheets("Sayfa1").Cells(1, 11).Value
        Cells(1, 9).Value = Sheets("Sayfa1").Cells(1, 16).Value
End If ''''
'Sheets("Sayfa1").Select
son = Sheets("Sayfa1").Range("A65536").End(xlUp).row
Y = 1
i = 0
listson = ListBoxP31.ListCount
 ProgressBarP21.Visible = True: ProgressBarP21.Value = 0: ProgressBarP21.Max = son
Do Until i = listson
    i = i + 1
    Do Until Sheets("Sayfa1").Cells(Y, 2) = "BÖLÜM ADI/NO:"
    Y = Y + 1
    If Y > son Then GoTo bitti:
    Loop
    Sayfaisim = Sheets("Sayfa1").Cells(Y, 3)
    listisim = ListBoxP31.List(i - 1, 0)
    If ListBoxP31.Selected(i - 1) = True And listisim = Sayfaisim Then: Else: Y = Y + 1: GoTo git:
    a = WorksheetFunction.CountA(Sheets(ad).Range("B:B"), xlDown)
      For X = 1 To 5
      Sheets(ad).Cells(a, X).Value = Sheets("Sayfa1").Cells(Y, X).Value
      Next
    Y = Y + 1
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    s = 1
    Do Until Sheets("Sayfa1").Cells(Y, 2) = "BÖLÜM ADI/NO:"
    If Sheets("Sayfa1").Cells(Y, 3) = "" Then GoTo var:
''''
kod1 = "PP-": kod2 = "PM-"
    If CheckBoxP32.Value = True Or CheckBoxP35.Value = True Then ' *** ürün kodlaması üzerinden ***
     sayfakod = Left(Sheets("Sayfa1").Cells(Y, 1), 3)
     If sayfakod = kod2 Then GoTo var
     If sayfakod = kod1 And CheckBoxP35.Value = True Then
      GoTo devam1
      Else
      If sayfakod = kod1 And CheckBoxP35.Value = False Then GoTo var
      If sayfakod <> kod1 And CheckBoxP32.Value = False Then GoTo var
    End If
End If
''''
devam1:
        a = WorksheetFunction.CountA(Sheets(ad).Range("B:B"), xlDown)
      For X = 1 To 5
      Sheets(ad).Cells(a, X).Value = Sheets("Sayfa1").Cells(Y, X).Value
      Next
''''''
If OptionButtonP32.Value = True And CheckBoxP31.Value = True Then ''''İSK
      Sheets(ad).Cells(a, 6).FormulaR1C1 = Sheets("Sayfa1").Cells(Y, 6).Value
      Sheets(ad).Cells(a, 6).NumberFormat = Sheets("Sayfa1").Cells(Y, 6).NumberFormat
      Sheets(ad).Cells(a, 7).FormulaR1C1 = Format(Sheets("Sayfa1").Cells(Y, 7).Value, "0%")
      kur = ""
      If Sheets(ad).Cells(a, 6).NumberFormat = "#,##0.00 [$€-1]" Then kur = "*Eur" ' Sayfa3 de € kuru
      If Sheets(ad).Cells(a, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then kur = "*Usd" ' Sayfa3 de $ kuru
      Sheets(ad).Cells(a, 8).FormulaR1C1 = "=(RC[-2]-RC[-2]*RC[-1])" & kur ' Formül yaz
      Sheets(ad).Cells(a, 8).NumberFormat = "#,##0.00"
      Sheets(ad).Cells(a, 9).FormulaR1C1 = "=RC[-4]*RC[-1]" ' Formül yaz
      Sheets(ad).Cells(a, 9).NumberFormat = "#,##0.00"
End If
'''''
    Cells(a, 1) = s
    s = s + 1
var:
    Y = Y + 1
    'Sheets("Sayfa1").Select
    ProgressBarP21.Value = Y
    If Y > son Then GoTo bitti:
    Loop
git:
Loop
bitti:
'ProgressBarP21.Value = son
Sheets(ad).Select
'ProgressBarP21.Visible = False
Call malzemebicim2
'Unload UFmd
'Application.ScreenUpdating = True
'ActiveSheet.PageSetup.CenterFooter = "Sayfa &P / &N"
'ActiveWindow.SelectedSheets.PrintPreview
'Range("A1").Select
'Msg = MsgBox("    Seçilen panolar sayfaya aktarıldı. ", vbOKOnly, "scngnr@hotmail.com")
bitti2:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub Makro122() ' Pano İcmal
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlManual
Dim a As Integer
Dim Y As Integer, X As Integer
Dim msg As String
        Sheets(ad).Visible = xlSheetVisible: Sheets(ad).Select
        a = WorksheetFunction.CountA(Range("B:B"), xlDown) - 1
        If a > 0 Then
        msg = MsgBox("    Mevcut  liste  silinerek  yeni   liste   düzenlenecek ", vbYesNo, "scngnr@hotmail.com ")
        If msg = vbYes Then
        Cells.Clear
        Else
        GoTo bitti2
        End If
        End If
Range("A1") = "S.No": Range("B1") = " Pano Adı": Range("C1") = "Miktar": Range("D1") = "Birim Fiyat": Range("E1") = "Toplam Fiyat"
        Range("A1:E1").Font.Size = 10
        Range("A1").RowHeight = 18.6
        Range("A1").ColumnWidth = 4.56
        Range("B1").ColumnWidth = 58
        Range("C1").ColumnWidth = 8
        Range("D1").ColumnWidth = 13
        Range("E1").ColumnWidth = 16
''''
son = Sheets("Sayfa1").Range("A65536").End(xlUp).row
listson = ListBoxP31.ListCount
 ProgressBarP21.Visible = True: ProgressBarP21.Value = 0: ProgressBarP21.Max = listson
For i = 0 To listson - 1
    a = WorksheetFunction.CountA(Sheets(ad).Range("B:B"), xlDown)
    Cells(a, 1) = i + 1
    Cells(a, 2) = ListBoxP31.List(i, 0)
    X = Sheets("Sayfa1").Cells(ListBoxP31.List(i, 4), 5)
    If X = 0 Then
    Cells(a, 3) = 1
    Cells(a, 4) = Sheets("Sayfa1").Cells(ListBoxP31.List(i, 5), 24)
    Else
    Cells(a, 3) = Sheets("Sayfa1").Cells(ListBoxP31.List(i, 4), 5)
    Cells(a, 4) = Sheets("Sayfa1").Cells(ListBoxP31.List(i, 5), 24) / Cells(a, 3)
    End If
    Cells(a, 5).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range(Cells(a, 4), Cells(a, 5)).NumberFormat = Sheets("Sayfa1").Cells(ListBoxP31.List(i, 5), 24).NumberFormat
ProgressBarP21.Value = i + 1
Next
Cells(a + 1, 2).FormulaR1C1 = "TOPLAM :"
Cells(a + 1, 5) = "=Sum(E2:E" & a & ")"
    Range(Cells(a + 1, 1), Cells(a + 1, 5)).Font.ColorIndex = 11
    Range(Cells(a + 1, 1), Cells(a + 1, 5)).Font.Bold = True
'Sheets(ad).Select
Call malzemebicim2
bitti2:
ProgressBarP21.Visible = False
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub