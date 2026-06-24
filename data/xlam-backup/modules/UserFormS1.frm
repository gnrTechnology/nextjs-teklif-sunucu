Dim fiyat
Dim p 'parabiçim
Dim dil 'dil
Dim ad 'seçilen şablon
Dim U, H, son As Integer
Dim ds1, dd
Dim DT1
Dim S1, s3, sn1
Dim nson As Integer
Private Sub CommandButton1_Click()
On Error Resume Next
UFFirma.Show
End Sub
Private Sub CommandButtonP33_Click()
If LBSAD.Caption = "" Then MsgBox ("Şablon seçin !  "), vbInformation, "scngnr@hotmail.com": Exit Sub
ListViewP1_dblClick
End Sub
Private Sub CommandButtonsd1_Click()
If LBSAD.Caption = "" Then MsgBox ("Şablon seçin !  "), vbInformation, "scngnr@hotmail.com": Exit Sub
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Şablonlar\" & ListBoxDS.Text & "\" & LBSAD.Caption
End Sub
Private Sub Image2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Şablonlar"
End Sub
Private Sub Image3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call dosyabicim
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Set S1 = Worksheets.Item("Sayfa1")
Set s3 = Worksheets.Item("Sayfa3")
Set sn1 = Worksheets.Item("Notlar")
'--
Dim ds, fs, f1, fc
Set ds = CreateObject("Scripting.FileSystemObject")
Set fs = ds.GetFolder("C:\Belgelerim\Cemex\Şablonlar")
Set fc = fs.SubFolders
For Each f1 In fc
ListBoxDS.AddItem f1.Name
Next
If ListBoxDS.ListCount > 0 Then ListBoxDS.Selected(0) = True
'--
If s3.Range("Tpbr") = "Teklif Para Birimi (TL)" Then p = "#,##0.00 $"
If s3.Range("Tpbr") = "Teklif Para Birimi (EUR)" Then p = "#,##0.00 [$€-1]"
If s3.Range("Tpbr") = "Teklif Para Birimi (USD)" Then p = "#,##0.00 [$$-C0C]"
ListViewP1.HotTracking = True
ssno = 2
End Sub
Private Sub ListBoxDS_Click()
On Error Resume Next
If ListBoxDS.ListCount < 1 Then Exit Sub
Dim itmX As listItem
Dim dosya
Dim n As Integer
Dim ic As Byte
ListViewP1.ListItems.Clear
ListViewP1.Arrange = lvwAutoTop
ListViewP1.Icons = ImageList1
dd = ListBoxDS.Text
ssno = 3
dosya = dir("C:\Belgelerim\Cemex\Şablonlar\" & dd & "\*.xltx")
If ListBoxDS.listIndex = 0 Then: ic = 3: MultiPage1.Value = 0
If ListBoxDS.listIndex > 0 And ListBoxDS.listIndex < 4 Then: ic = 1: MultiPage1.Value = 1
If ListBoxDS.listIndex > 3 Then: ic = 4: MultiPage1.Value = 2
If ListBoxDS.listIndex = 1 Or ListBoxDS.listIndex = 2 Then ssno = 2
Do While dosya <> ""
Set itmX = ListViewP1.ListItems.Add()
itmX.Text = dosya: itmX.Icon = ic
    dosya = dir
    n = n + 1
Loop
ListViewP1.SelectedItem.Selected = False
LBSAD.Caption = ""
End Sub
Private Sub ListViewP1_Click()
LBSAD.Caption = ListViewP1.SelectedItem
yold = "C:\Belgelerim\Cemex\Şablonlar\" & dd & "\[" & LBSAD.Caption & "]"
dil = Application.ExecuteExcel4Macro("'" & yold & "Veriler" & "'!R" & 19 & "C2") '
End Sub
Sub dosyabicim()
On Error Resume Next
yold = "C:\Belgelerim\Cemex\Şablonlar\" & dd & "\[" & LBSAD.Caption & "]"
ssf = "Kapak"
Dim str As Integer
For str = 1 To 30
sno = Application.ExecuteExcel4Macro("'" & yold & ssf & "'!R" & str & "C1") '
ssft = "KAPAK"
If sno = "S.No" Then GoTo git1
Next
ssf = "DETAYLI LİSTE"
For str = 1 To 10
sno = Application.ExecuteExcel4Macro("'" & yold & ssf & "'!R" & str & "C1") '
ssft = "KAPAK .. DETAYLI LİSTE"
If sno = "S.No" Then GoTo git1
Next
Exit Sub
git1:
For stt = 1 To 12
dkl = Application.ExecuteExcel4Macro("'" & yold & ssf & "'!R" & str & "C" & stt) '
If dkl = "" Or dkl = 0 Then GoTo git2
dkb = dkb & "-" & dkl
Next
git2:
MsgBox ssft & vbLf & dkb, vbInformation, "scngnr@hotmail.com"
End Sub
Private Sub ListViewP1_dblClick()
On Error Resume Next
DT1 = ActiveWorkbook.Name
'ActiveWorkbook.Sheets("Sayfa1").Select
son = Range("B65536").End(xlUp).row
If son < 2 Then MsgBox ("    Dosyaya veri girimemiş !  "), vbInformation, "scngnr@hotmail.com": Exit Sub
ad = ListViewP1.SelectedItem.Text

'--
If ListBoxDS.listIndex = 0 Then Call teklifyap1
If ListBoxDS.listIndex > 0 Then
  If Not Cells(1, 1) = "S.No" Then
  MsgBox ("   Bu işlem malzeme listesinin olduğu sayfada yapılabilir !  "), vbInformation, "scngnr@hotmail.com": GoTo bitti
  End If
  If ListBoxDS.listIndex = 4 Then CBox212.Value = True
  Call teklifyap2
End If
'--
'ActiveWorkbook.Sheets("Kapak").Select
Call sayfasay
ActiveWorkbook.Sheets("Kapak").Select
Call dilbul

' **Yeni Kayıt Kısmı Başlangıç**
Dim Yol As String, dosyaAdi As String

' Çağrılan Excel dosyasının klasör yolunu al
'yol = ThisWorkbook.path 'ThisWorkbook Userformun içinde bulunduğu excel dosyasıdır. Eğer çağırdığınız excel ise "Workbooks(DT1).Path"

Yol = Workbooks(DT1).path
' Yeni dosya adını oluştur (Klasör adı + "Teklif")

'dosyaAdi = Left(DT1, InStrRev(DT1, ".", -1, vbTextCompare) - 1) & " - Teklif.xlsx"
dosyaAdi = Left(DT1, InStrRev(DT1, ".", -1, vbTextCompare) - 1) & " - " & ListBoxDS.Value & ".xlsx"

' Yeni tam dosya yolunu oluştur
ad2 = Yol & "\" & dosyaAdi

' Mevcut aktif çalışma kitabını yeni adıyla kaydet
Workbooks(ds1).SaveAs fileName:=ad2, FileFormat:=xlOpenXMLWorkbook

' **Yeni Kayıt Kısmı Bitiş**

msg = MsgBox("Teklif dosyası aktarıldı: " & ad2, vbOKOnly, "Dosya Oluştuma")

Workbooks(ds1).Activate
Unload Me
bitti:
End Sub
Sub teklifyap1() 'veriler sayfası
On Error Resume Next
Application.Calculation = xlCalculationManual: Application.ScreenUpdating = False: Application.EnableEvents = False
UserFormS1.Caption = "Pano iç malzemeli detaylı liste oluşturuluyor.."
Sheets("Sayfa1").Activate
son = Range("B65536").End(xlUp).row
nson = Worksheets("Notlar").Range("B65536").End(xlUp).row
'--
Workbooks.Open "C:\Belgelerim\Cemex\Şablonlar\" & dd & "\" & ad
ds1 = ActiveWorkbook.Name
Call teklifveri1
'If Left(ActiveSheet.CodeName, 3) = "MTS" Then Call teklifyap2: GoTo git
'--
Sheets("DETAYLI LİSTE").Select
Range("A65536").End(xlUp).Offset(1, 0).Select
U = Selection.row
H = Selection.row
S1.Range("A2:E" & son & ",W2:X" & son).Copy
ActiveSheet.Paste: Application.CutCopyMode = False
Selection.Interior.Pattern = xlNone
Selection.Font.ColorIndex = xlAutomatic
If CBox4.Value = True Then Selection.VerticalAlignment = xlTop
'Selection.HorizontalAlignment = xlGeneral
Range("A1").Select
'--
If CBox1.Value = False Then Call düzenleF0 Else Call düzenleF0: Call AraToplamlarF1
'--
Call panolisteleF0
If CBox1.Value = False Then Call fytsilF1
'--
Call bicim1
git:
If CBox4.Value = True Then Sheets("DETAYLI LİSTE").Columns("C:C").WrapText = True: Sheets("DETAYLI LİSTE").Columns("C:C").Rows.AutoFit '
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub düzenleF0()
On Error Resume Next
Application.Calculation = xlCalculationManual
d = Range("B65536").End(xlUp).row
s = 1
ProgressBarP21.Min = 0: ProgressBarP21.Max = d: ProgressBarP21.Visible = True
For X = U To d
    If Cells(X, 4) = "" And Cells(X, 3) <> "" And Cells(X, 2) = "BÖLÜM ADI/NO:" Then
     Range(Cells(X, 1), Cells(X, 7)).Font.Color = -4165632: s = 1 'Başlık renk
     If CBox5.Value = True Then
     If Not Cells(X, 1) = "" Then Cells(X, 3) = Cells(X, 3) & " [Pano Ref." & Cells(X, 1) & "]"
     End If
     Cells(X, 1) = ""
    End If
If CBox1.Value = True Then
 If Cells(X, 4) <> "" Then Cells(X, 7).FormulaR1C1 = "=RC[-2]*RC[-1]"
 Range(Cells(X, 6), Cells(X, 7)).NumberFormat = p
End If
'----b
If CBox6.Value = True Then
If Left(Cells(X, 1), 3) = "PP-" Then
Pptp = Pptp + Cells(X, 7): ppx = X 'pano
End If
If ppz <> ppx Then
       LBPP.AddItem X
       LBPP.List(LBPP.ListCount - 1, 1) = Cells(X, 7)
       ppz = ppx
End If
If Left(Cells(X, 1), 3) = "PM-" Then
If Not Left(Cells(X, 1), 5) = "PM-MB" Then
Pptm = Pptm + Cells(X, 7): Cells(X, 7).EntireRow.Delete: X = X - 1 'montaj
s = s - 1
End If
End If

End If
'----s
If Cells(X, 4) <> "" Then Cells(X, 1) = s: Cells(X, 1).HorizontalAlignment = xlLeft: _
    Range(Cells(X, 6), Cells(X, 7)).NumberFormat = p: s = s + 1
If Cells(X, 2) = "BÖLÜM TOPLAMI:" Then
'----b
If CBox6.Value = True Then
 For i = 0 To LBPP.ListCount - 1
 ' kontrol et
ppi = 0: ppi = Cells(LBPP.List(i), 7) / Pptp * Pptm
Cells(LBPP.List(i), 6) = (ppi + Cells(LBPP.List(i), 7)) / Cells(LBPP.List(i), 5)
Cells(LBPP.List(i), 7).FormulaR1C1 = "=RC[-2]*RC[-1]"
 Next i
End If
    LBPP.Clear: ppx = "": ppz = "": Pptp = 0: Pptm = 0: ppi = 0
'----s
 If CBox1.Value = False Then
    Cells(X, 4) = Cells(X, 7): Range(Cells(X, 4), Cells(X, 5)).Merge
    'Cells(X, 7).Copy: Cells(X, 4).PasteSpecial xlPasteAll: Range(Cells(X, 4), Cells(X, 5)).Merge
    'Cells(X, 4).HorizontalAlignment = xlRight: Cells(X, 4) = WorksheetFunction.RoundUp(Cells(X, 4), 0)
    Cells(X, 4).HorizontalAlignment = xlRight
    Cells(X, 4).NumberFormat = p
    Range(Cells(X, 1), Cells(X, 7)).Font.Color = -4165632 'toplam renk
 End If

End If
    'Range(Cells(X, 1), Cells(X, 1)).Merge
ProgressBarP21.Value = X
Next
    If Cells(X - 1, 2) = "GENEL TOPLAM:" Then
    Range(Cells(X - 1, 4), Cells(X - 1, 5)).Merge
    Cells(X - 1, 4).HorizontalAlignment = xlRight: Cells(X - 1, 4).NumberFormat = p
    Cells(X - 1, 4) = "=SUMIF(R2C2:R[-1]C[-2],""=BÖLÜM TOPLAMI:"",R2C:R[-1]C4)"
    Range(Cells(X - 1, 1), Cells(X - 1, 7)).Font.Color = -4165632 'g.toplam renk
    End If
ProgressBarP21.Value = 0
Application.Calculation = xlCalculationAutomatic
End Sub
Sub düzenleF1() 'iptal edilebilir
On Error Resume Next
Application.Calculation = xlCalculationManual ': Application.ScreenUpdating = False: Application.EnableEvents = False
d = Range("B65536").End(xlUp).row
ProgressBarP21.Min = 0: ProgressBarP21.Max = d: ProgressBarP21.Visible = True
For X = U To d
    If Cells(X, 1) = "" Then: Range(Cells(X, 1), Cells(X, 7)).Font.Color = -4165632: s = 1
    If Cells(X, 4) <> "" Then Cells(X, 1) = s: Cells(X, 1).HorizontalAlignment = xlLeft: s = s + 1
    If Cells(X, 4) <> "" Then Cells(X, 7).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range(Cells(X, 6), Cells(X, 7)).NumberFormat = p
ProgressBarP21.Value = X
Next
ProgressBarP21.Value = 0
Application.Calculation = xlCalculationAutomatic ': Application.EnableEvents = True: Application.ScreenUpdating = True
End Sub
Sub panolisteleF0()
On Error Resume Next
Application.Calculation = xlCalculationManual
UserFormS1.Caption = "Pano açılımlı icmal oluşturuluyor.."
Set ek1 = Workbooks(ds1).Worksheets("İCMAL")
Set ek4 = Workbooks(ds1).Worksheets("DETAYLI LİSTE")
ek1.Select
ek1.Range("A65536").End(xlUp).Offset(1, 0).Select
U = Selection.row: Y = 1: s = 1
son = ek4.Range("B65536").End(xlUp).row + 1
tekrar:
ProgressBarP21.Value = Y
    Do Until ek4.Cells(Y, 2) = "BÖLÜM ADI/NO:"
    Y = Y + 1
    If Y > son Then GoTo bitti:
    Loop
    ads = ek4.Range("E" & Y): If ads = "" Then ads = 1
    a = ek1.Range("B65536").End(xlUp).row + 1
    ek1.Cells(a, 2) = ek4.Range("C" & Y)
    ek1.Cells(a, 3) = ek4.Range("E" & Y)
    ek1.Range(Cells(a, 1), Cells(a, 5)).Borders.LineStyle = xlContinuous
    ek1.Range(Cells(a, 1), Cells(a, 5)).Font.Size = 9
    ek1.Range(Cells(a, 1), Cells(a, 5)).Font.Name = "Arial"
    ek1.Range(Cells(a, 1), Cells(a, 2)).HorizontalAlignment = xlLeft
    ek1.Range(Cells(a, 3), Cells(a, 5)).HorizontalAlignment = xlRight
    ek1.Cells(a, 3).NumberFormat = "#,##0"
    ek1.Range(Cells(a, 4), Cells(a, 5)).NumberFormat = p
    Y = Y + 1
    Do Until ek4.Cells(Y, 2) = "BÖLÜM TOPLAMI:"
    Y = Y + 1
    If Y > son Then GoTo bitti:
    Loop
    ek1.Cells(a, 4) = ek4.Cells(Y, 7)
    'Cells(a, 4) = WorksheetFunction.RoundUp(Cells(a, 4) / ads, 0)
    ek1.Cells(a, 4) = ek1.Cells(a, 4) / ads
    If ek1.Cells(a, 3) = "" Then ek1.Cells(a, 3) = 1
    If Not IsNumeric(ek1.Cells(a, 3)) Or ek1.Cells(a, 3) < 1 Then ek1.Cells(a, 3) = 1
    ek1.Cells(a, 5).FormulaR1C1 = "=RC[-2]*RC[-1]"
    ek1.Cells(a, 1) = s
    s = s + 1
GoTo tekrar:
bitti:
    a = Range("B65536").End(xlUp).row + 1
    Cells(a, 1) = "GENEL TOPLAM:"
    Cells(a, 5) = "=Sum(E2:E" & a - 1 & ")"
    fiyat = Cells(a, 5).Text
    ek1.Range(Cells(a, 1), Cells(a, 5)).Borders.LineStyle = xlContinuous
    ek1.Range(Cells(a, 1), Cells(a, 4)).Borders(xlInsideVertical).LineStyle = NONE
    ek1.Range(Cells(a, 1), Cells(a, 5)).Font.Size = 9
    ek1.Range(Cells(a, 1), Cells(a, 5)).Font.Name = "Arial"
    ek1.Range(Cells(a, 1), Cells(a, 5)).Font.Bold = True
    ek1.Range(Cells(a, 1), Cells(a, 5)).Font.ColorIndex = 11
    ek1.Range(Cells(a, 4), Cells(a, 5)).NumberFormat = p
    'Worksheets.item("Kapak").Range("AB15") = fiyat
    ek1.Cells(a + 2, 1) = Worksheets("Veriler").Range("B15").Value: ek1.Cells(a + 2, 1).Font.ColorIndex = 11
    'ek1.Cells(a + 2, 1) = "Fiyatlarımıza KDV dahil edilmemiştir."
Range("A1").Select
ProgressBarP21.Value = son
Application.Calculation = xlCalculationAutomatic
End Sub
Sub AraToplamlarF1()
Application.Calculation = xlCalculationManual ': Application.ScreenUpdating = False: Application.EnableEvents = False
Dim hcr As Range, eskihcr As Range
On Error GoTo hata
Set hcr = Columns("B:B").Find("BÖLÜM TOPLAMI:", LookAt:=xlWhole)
  Cells(hcr.row, "G") = "=Sum(G2:G" & hcr.row - 1 & ")"
Set eskihcr = hcr.Offset(2, 0)
Do
Set hcr = Range(hcr.Offset(1, 0), [B65000]).FindNext
  Cells(hcr.row, "G") = "=Sum(G" & eskihcr.row - 1 & ":G" & hcr.row - 1 & ")"
Set eskihcr = hcr.Offset(2, 0)
Loop
hata:
Application.Calculation = xlCalculationAutomatic ': Application.EnableEvents = True: Application.ScreenUpdating = True
End Sub
Sub fytsilF1()
On Error Resume Next
Application.Calculation = xlCalculationManual ': Application.EnableEvents = False
Set ek4 = Workbooks(ds1).Worksheets("DETAYLI LİSTE")
Set ek2 = Workbooks(ds1).Worksheets("İCMAL")
If ek4.Range("G1").Value <> "" Then b1 = 1 Else b1 = 2
    ek4.Range("G" & b1 & ":G" & H - 2).Cut Destination:=ek4.Range("E" & b1 & ":E" & H - 2)
    Application.CutCopyMode = False
    ek4.Columns("F:G").Delete Shift:=xlToLeft
If ek4.PageSetup.Orientation = xlPortrait Then
    ek4.Range("B1").ColumnWidth = 17
    ek4.Range("C1").ColumnWidth = 52
    ek4.Range("D1").ColumnWidth = 10
    ek4.Range("E1").ColumnWidth = 10
Else
    ek4.Range("B1").ColumnWidth = 21
    ek4.Range("C1").ColumnWidth = 79.56
    ek4.Range("D1").ColumnWidth = 19.22
    ek4.Range("E1").ColumnWidth = 10.67
End If
    ek4.Range("E" & U & ":E" & son - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
Application.Calculation = xlCalculationAutomatic ': Application.EnableEvents = True
End Sub
Sub sayfasay()
Dim a, s As Integer
a = 0: n = 0
 For i = 1 To ActiveWorkbook.Worksheets.Count
 If Worksheets(i).Name = "Veriler" Then i = i + 1
      a = ActiveWorkbook.Worksheets(i).HPageBreaks.Count + 1
      Worksheets.Item("Veriler").Range("A" & 20 + n) = Worksheets(i).Name
      Worksheets.Item("Veriler").Range("B" & 20 + n) = Worksheets(i).HPageBreaks.Count + 1
      s = a + s
      n = n + 1
 Next i
End Sub
Sub bicim1()
On Error Resume Next
Application.Calculation = xlCalculationManual ': Application.ScreenUpdating = False
If CBox2.Value = True Or CBox3.Value = True Then
 Set ek4 = Workbooks(ds1).Worksheets("DETAYLI LİSTE")
ProgressBarP21.Max = ek4.Range("B65536").End(xlUp).row + 1
i = 2
Do
ProgressBarP21.Value = i
  If ek4.Cells(i, 2) = "BÖLÜM ADI/NO:" Then
  If CBox1.Value = True Then ek4.Range("A" & i & ":G" & i).Interior.ColorIndex = 40 Else _
    ek4.Range("A" & i & ":E" & i).Interior.ColorIndex = 40
  End If
  If ek4.Cells(i, 2) = "BÖLÜM TOPLAMI:" Or ek4.Cells(i, 2) = "GENEL TOPLAM:" Then
    If CBox2.Value = True And ek4.Cells(i + 1, 2) <> "GENEL TOPLAM:" Then
    ek4.Cells(i + 1, 2).EntireRow.Insert
    ek4.Range("A" & i + 1).Borders(xlEdgeLeft).LineStyle = xlNone
    ek4.Range("A" & i + 1 & ":H" & i + 1).Borders(xlInsideVertical).LineStyle = xlNone
    End If
    If CBox3.Value = True Then
    ek4.Range("B" & i).Font.ColorIndex = 15: ek4.Range("D" & i) = "": ek4.Range("G" & i) = ""
    If CBox1.Value = True Then ek4.Range("A" & i & ":G" & i).Interior.ColorIndex = 15 Else _
    ek4.Range("A" & i & ":E" & i).Interior.ColorIndex = 15
    End If
    End If
 i = i + 1
Loop Until i = ek4.Range("B65536").End(xlUp).row + 1
End If
    If CBox3.Value = True Then If ek4.Cells(i - 1, 2) = "GENEL TOPLAM:" Then ek4.Cells(i - 1, 2).EntireRow.Delete
End Sub
Sub teklifyap2() 'veriler sayfası
On Error Resume Next
Application.Calculation = xlCalculationManual: Application.ScreenUpdating = False ': Application.EnableEvents = False
UserFormS1.Caption = "Pano iç malzemeli detaylı liste oluşturuluyor.."
Dim i As Long
Sonkolon = Cells(1, 256).End(xlToLeft).Column ' kolon sayısını verir
son = Cells(Rows.Count, Sonkolon).End(xlUp).row
'Range(Cells(2, 1), Cells(son, Sonkolon)).Select
ssdd = 1
nson = Worksheets("Notlar").Range("B65536").End(xlUp).row
'--
Workbooks.Open "C:\Belgelerim\Cemex\Şablonlar\" & dd & "\" & ad
ds1 = ActiveWorkbook.Name
  Call mlzveri1
For i = 1 To Sheets.Count
    If Sheets(i).Name = "DETAYLI LİSTE" Then Sheets("DETAYLI LİSTE").Select: GoTo var1
Next i
Worksheets("KAPAK").Select
var1:
smf = Workbooks(ds1).ActiveSheet.Name
'--
Set sno = Columns("A:A").Find("S.No", LookAt:=xlWhole)
U = sno.row + 1
H = sno.row + 1
Range("A" & U).Select
Dlsonkolon = Cells(U - 1, 256).End(xlToLeft).Column ' sablon kolon sayısını verir
ssf = Workbooks(DT1).ActiveSheet.Name
Set smx = Workbooks(DT1).Worksheets(ssf)
If Dlsonkolon < 6 Then son = smx.Cells(Rows.Count, Dlsonkolon).End(xlUp).row
If CBoxkdv = False Then son = smx.Cells(Rows.Count, 2).End(xlUp).row
'--
smx.Range(smx.Cells(2, 1), smx.Cells(son, Dlsonkolon)).Copy
'Workbooks(ds1).Worksheets(smf).Range("A" & U).Insert
Workbooks(ds1).Worksheets(smf).Range("A" & U).EntireRow.Insert Shift:=xlDown
'Range("A" & U).Insert
'--
 a = U + son - 1
If Dlsonkolon > Sonkolon Then
Workbooks(ds1).Worksheets(smf).Range(Cells(U, Sonkolon), Cells(a - 1, Dlsonkolon)).Borders.LineStyle = xlContinuous
End If
Workbooks(ds1).Worksheets(smf).Range(Cells(U, 1), Cells(a - 1, Dlsonkolon)).Borders(xlEdgeRight).LineStyle = xlContinuous
Workbooks(ds1).Worksheets(smf).Range(Cells(U, Sonkolon), Cells(a - 1, Dlsonkolon)).RowHeight = 12.75
If CBoxkdv = False Then GoTo atla1
If Dlsonkolon < 6 Or Sonkolon < 6 Then GoTo atla1
    Range(Cells(a, 1), Cells(a + 1, 1)).EntireRow.Insert
    Range(Cells(a - 1, 1), Cells(a - 1, 9)).Copy
    Range(Cells(a, 1), Cells(a, 9)).PasteSpecial Paste:=xlPasteFormats
    Range(Cells(a + 1, 1), Cells(a + 1, 9)).PasteSpecial Paste:=xlPasteFormats
    Cells(a, 8) = "%18 KDV:": Cells(a, 9).FormulaR1C1 = "=R[-1]C*0.18"
    Cells(a + 1, 8) = "GENEL TOPLAM:": Cells(a + 1, 9).FormulaR1C1 = "=R[-2]C+R[-1]C"
    Cells(a - 1, 8).HorizontalAlignment = xlRight
    Cells(a, 8).HorizontalAlignment = xlRight
    Cells(a + 1, 8).HorizontalAlignment = xlRight
If Dlsonkolon > Sonkolon Then
Workbooks(ds1).Worksheets(smf).Range(Cells(U, Sonkolon + 1), Cells(a + 1, Dlsonkolon)).Borders.LineStyle = xlContinuous
End If
atla1:
Application.CutCopyMode = False
'--
If CBox21 = True Then
Do Until U > a + 1
    If Cells(U, 4) = "" And Cells(U, 8) = "" Then Cells(U, 4).EntireRow.Delete: a = a - 1
U = U + 1
Loop
End If
Do Until H > a
    If Not Cells(H, 4) = "" Then
    Cells(H, 5) = Cells(H, 5).Value
     If CBox212 = True Then
     Range(Cells(H, 6), Cells(H, Dlsonkolon)) = ""
     Range(Cells(H, 6), Cells(H, Dlsonkolon)).Font.Size = 9: Range(Cells(H, 6), Cells(H, Dlsonkolon)).Font.Name = "Arial"
     Cells(H, 7).FormulaR1C1 = "=IF(RC[-1]>RC[-2],0,RC[-2]-RC[-1])"
     Range(Cells(H, 6), Cells(H, Dlsonkolon)).NumberFormat = "#,##0"
     Range(Cells(H, 6), Cells(H, Dlsonkolon)).Font.ColorIndex = 3
     End If
    End If
    H = H + 1
Loop
Range("A1").Select
'--
git:
'--
'If Worksheets("KAPAK").HPageBreaks.Count + 1 > 1 Then
    'Range("A" & h).Select: ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
'End If
Workbooks(ds1).Worksheets(smf).Refresh
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True ': Application.EnableEvents = True
End Sub
Sub teklifveri1() 'Veriler sayfası--
Application.Calculation = xlCalculationManual
For i = 1 To Worksheets.Count
If Sheets(i).Name = "Veriler" Then GoTo sfvar:
Next i
Sheets.Add Before:=Sheets(1): Sheets(1).Name = "Veriler"
sfvar:
Worksheets("Veriler").Visible = xlSheetHidden
'--
Set veriler = Worksheets.Item("Veriler")
'Veriler.Visible = True
veriler.Range("A1") = s3.Range("B3"): veriler.Range("B1") = s3.Range("C3")
veriler.Range("A2") = s3.Range("B4"): veriler.Range("B2") = s3.Range("C4")
veriler.Range("A3") = s3.Range("B5"): veriler.Range("B3") = s3.Range("C5"): veriler.Range("C3") = s3.Range("E5")
veriler.Range("A4") = s3.Range("B7"): veriler.Range("B4") = s3.Range("C7")
veriler.Range("A5") = s3.Range("B8"): veriler.Range("B5") = s3.Range("C8")
veriler.Range("A6") = s3.Range("B10"): veriler.Range("B6") = s3.Range("C10"): veriler.Range("C6") = s3.Range("F10")
veriler.Range("A7") = s3.Range("B11"): veriler.Range("B7") = s3.Range("C11")
veriler.Range("A8") = s3.Range("B12"): veriler.Range("B8") = s3.Range("C12"):
veriler.Range("D8") = s3.Range("D12"): veriler.Range("C8") = s3.Range("F12")
veriler.Range("A9") = s3.Range("B13"): veriler.Range("B9") = s3.Range("C13")
veriler.Range("A10") = s3.Range("Tpbr")
veriler.Range("A11") = "USD/TRY": veriler.Range("B11") = s3.Range("Usd")
veriler.Range("A12") = "EUR/TRY": veriler.Range("B12") = s3.Range("Eur")
veriler.Range("A30") = "Dosya Yolu": veriler.Range("B30") = Workbooks(DT1).path
veriler.Range("A31") = "Dosya Adı": veriler.Range("B31") = Replace(DT1, ".xlsx", "")
'--
If nson > 0 Then 'notlar sayfası
For si = 2 To nson
veriler.Range("A" & 33 + si) = sn1.Range("A" & si): veriler.Range("B" & 33 + si) = sn1.Range("B" & si)
Next si
End If
'--
Application.Calculation = xlCalculationAutomatic
'Sheets("Notlar").Select
Sheets("Notlar").Range("A" & 2 & ":B" & 100).Rows.AutoFit
End Sub
Sub mlzveri1() 'Veriler sayfası--
On Error Resume Next
Application.Calculation = xlCalculationManual
'Veriler sayfası--
For i = 1 To Worksheets.Count
If Sheets(i).Name = "Veriler" Then GoTo sfvar:
Next i
Sheets.Add Before:=Sheets(1): Sheets(1).Name = "Veriler"
sfvar:
Worksheets("Veriler").Visible = xlSheetHidden
'--
Set veriler = Worksheets.Item("Veriler")
'Veriler.Visible = True
veriler.Range("A2") = ListBoxFB.List(1, 0): veriler.Range("B2") = ListBoxFB.List(1, 1)
veriler.Range("A3") = LabelFAD1: veriler.Range("B3") = LabelFAD2
veriler.Range("A7") = ListBoxFB.List(4, 0): veriler.Range("B7") = ListBoxFB.List(4, 1)
veriler.Range("A8") = ListBoxFB.List(2, 0): veriler.Range("B8") = ListBoxFB.List(2, 1)
veriler.Range("C8") = ListBoxFB.List(6, 1): veriler.Range("D8") = ListBoxFB.List(3, 1)
veriler.Range("A9") = ListBoxFB.List(5, 0): veriler.Range("B9") = ListBoxFB.List(5, 1)
veriler.Range("A5") = s3.Range("B8"): veriler.Range("B5") = s3.Range("C8")
veriler.Range("A4") = s3.Range("B7"): veriler.Range("B4") = s3.Range("C7")
'--
veriler.Range("A30") = "Dosya Yolu": veriler.Range("B30") = Workbooks(DT1).path
veriler.Range("A31") = "Dosya Adı": veriler.Range("B31") = Replace(DT1, ".xlsx", "")
If sn1 <> "" Then
For si = 2 To nson
veriler.Range("A" & 33 + si) = sn1.Range("A" & si): veriler.Range("B" & 33 + si) = sn1.Range("B" & si)
Next si
End If
Application.Calculation = xlCalculationAutomatic
End Sub
Sub dilbul() 'dil ceviri--
On Error Resume Next
Application.Calculation = xlCalculationManual: Application.ScreenUpdating = False ': Application.EnableEvents = False
 If dil = "İNGİLİZCE" Then
Dim Tabelle As Worksheet
For Each Tabelle In ActiveWorkbook.Worksheets
 If Tabelle.Name = "KAPAK" Then Tabelle.Name = "COVER"
 If Tabelle.Name = "İCMAL" Then Tabelle.Name = "PANELBOARD LIST"
 If Tabelle.Name = "DETAYLI LİSTE" Then Tabelle.Name = "DETAILED LIST"
 If Tabelle.Name = "NOTLAR" Then Tabelle.Name = "NOTES"
 Set tsayfa = Workbooks(ds1).Worksheets(Tabelle.Name)
 'Worksheets.Item(Tabelle.Name).Select
    tsayfa.Cells.Replace What:="BÖLÜM ADI/NO:", Replacement:="SECTION NAME/NO:", LookAt:=xlPart
    tsayfa.Cells.Replace What:="BÖLÜM TOPLAMI:", Replacement:="SECTION TOTAL:", LookAt:=xlPart
    tsayfa.Cells.Replace What:="GENEL TOPLAM:", Replacement:="GENERAL TOTAL:", LookAt:=xlPart
    tsayfa.Cells.Replace What:="Fiyatlarımıza KDV dahil edilmemiştir.", Replacement:="VAT is not included in our prices.", LookAt:=xlPart
Next Tabelle
 End If
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub