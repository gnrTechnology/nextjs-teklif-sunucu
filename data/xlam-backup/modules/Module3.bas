Public wd
Public sk1
Sub dosyaac()
On Error Resume Next
    Dim wb As Workbook
    Dim i%
    Dim dName
    Dim dFilter$
    dFilter = "Excel Dosyaları (*.xl*), *.xl*"
    ChDrive "c"
    ChDir "c:\"
    dName = Application.GetOpenFilename(dFilter)
    If dName = False Then wd = 0: Exit Sub
    wd = 1
    Set wb = Workbooks.Open(dName)
    dsa = ActiveWorkbook.Name
    'Call Dosyalar
    'ComboBoxP12.Value = dsa
    'Call bilgilerdosya
End Sub
Sub formulyap() 'yeni formül yapma 25.10.2014
On Error GoTo hata1
Application.DisplayAlerts = False
Sheets("Sayfa1").Select
son = Range("B65536").End(xlUp).row + 1
m = 2
Do Until son < m
b = Cells(m, "b"): d = Cells(m, "d")
If b = "" Then GoTo git1:
Cells(m, "a").NumberFormat = "General"
If d = "" Then sn = b: Cells(m, "a") = sn
Cells(m, "a") = sn
git1:
m = m + 1
Loop
hata1:
Application.DisplayAlerts = True
End Sub
Sub formulyapxxx() 'yeni formül yapma eski a sutunu referans satırlı
On Error GoTo hata1
Application.DisplayAlerts = False
Sheets("Sayfa1").Select
son = Range("B65536").End(xlUp).row + 1
m = 2
Do Until son < m
b = Cells(m, "b")
d = Cells(m, "d")
If b = "" Then GoTo git1:
If d = "" Then Cells(m, "a").NumberFormat = "General": Cells(m, "a") = Cells(m, "b"): z = b: GoTo git1:
Cells(m, "a") = z & "/" & b
git1:
m = m + 1
Loop
hata1:
Application.DisplayAlerts = True
End Sub

Sub yazırenk() '
On Error Resume Next
son = Sheets("Sayfa1").Range("C65536").End(xlUp).row
Range("A" & 2 & ":K" & son).Select
a = Application.Dialogs(xlDialogFontProperties).Show(Arg1:="Arial", Arg2:="Normal", Arg3:=10)
Range("A2").Select
End Sub
Sub AraToplamlar()
If Not Cells(1, "X") = "Toplam Fiyat" Then Exit Sub
'Sheets("Sayfa1").Select
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
Dim Y
hata:
'genel toplam kontrolü
Y = Range("B65536").End(xlUp).row
     If Cells(Y, 2) = "GENEL TOPLAM:" Then
     Cells(Y, 24) = "=SUM(R2C24:R[-1]C)-SUMIF(R2C2:R[-1]C[-22],""=BÖLÜM TOPLAMI:"",R2C24:R[-1]C)"
     End If
'genel toplam kontrolü sonu
End Sub
Sub Hatabul() 'Hata Sınaması
On Error Resume Next
Dim kisk As Double
Sheets("Sayfa1").Select
Range("A1").Select

X = Range("B65536").End(xlUp).row + 1
        Y = 0
        z = 0
        k = 0
Do Until Y = X
            Range("A1").Select
       If Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0) = "" Then
            Do Until Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 0) <> ""
                z = z + 1
            If z > X Then Exit Do
            Loop
        End If
       If Left(Range("A1").Offset(Y + z + 1, 0), 4) = "PM-M" Then
       If Left(Range("A1").Offset(Y + z + 1, 0), 5) = "PM-MB" Then GoTo bakır
            Do Until Left(Range("A1").Offset(Y + z + 1, 0), 4) <> "PM-M"
                z = z + 1
            Loop
bakır:
        End If
                kRef = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1)
                kad = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 4)
                kMlz = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 5)
                kisk = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 6)
                kAdk = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 7)
                kBr = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 22)
                Y = Y + 1
        Do Until kRef = Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 1)
                   k = k + 1
                   If k > X + z + Y Then
                   Exit Do
                   End If
        Loop
        kRef1 = Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 1)
        kRef1s = "(" & Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 1).row & ".Satır) " + kRef1
        If kRef = kRef1 Then
            If Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 5) <> kMlz Then
                msg = MsgBox(kRef1s + " Satırda Malzeme liste Fiyatı farklılığı var !", vbExclamation, "  scngnr@gmail.com ")
                Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 5).Select
                ActiveCell.Range("A1").Font.ColorIndex = 3
                k = 0
                'Exit Do
                Exit Sub
            ElseIf Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 6) <> kisk Then
                msg = MsgBox(kRef1s + " Satırda Malzeme İskonto oranı farklılığı var !", vbExclamation, "  scngnr@gmail.com ")
                Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 6).Select
                ActiveCell.Range("A1").Font.ColorIndex = 3
                k = 0
                'Exit Do
                Exit Sub
                ElseIf Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 7) <> kAdk Then
                msg = MsgBox(kRef1s + " Satırda Malzeme Ad/dk  farklılığı var !", vbExclamation, "  scngnr@gmail.com ")
                Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 7).Select
                ActiveCell.Range("A1").Font.ColorIndex = 3
                k = 0
                'Exit Do
                Exit Sub
                
            ElseIf Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 22) <> kBr Then
            
                msg = MsgBox(kRef1s + " Satırda Formülasyonlarda farklılık var !", vbExclamation, "  scngnr@gmail.com ")
                Sheets("Sayfa1").Range("A1").Offset(Y + z + k + 1, 0).Select
                ''ActiveCell.Range("A1:AA1").Font.ColorIndex = 3
                ActiveCell.Range("A1:X1").Font.ColorIndex = 3
                'ActiveCell.Range("A1").Select
                k = 0
                'Exit Do
                Exit Sub
            End If
        End If
            k = 0
Loop
msg = MsgBox("Hata sınaması başarıyla tamamlanmıştır. ", vbExclamation, "  scngnr@gmail.com ")
End Sub
Sub MakroDOLAR()
On Error GoTo hata
Application.ScreenUpdating = False
Usd = Sheets("Sayfa3").Range("Usd")
Sheets("Sayfa1").Columns("W:X").NumberFormat = "#,##0.00 [$$-C0C]"
    Sheets("Sayfa3").Range("M29").NumberFormat = "#,##0.0000 [$$-C0C]"
    Sheets("Sayfa3").Range("Tpb") = Sheets("Sayfa3").Range("Usd")
'Range("A1").Select
Sheets("Sayfa3").Range("Tpbr") = "Teklif Para Birimi (USD)"
Application.ScreenUpdating = True
hata:
End Sub
Sub MakroEURO()
On Error GoTo hata
Application.ScreenUpdating = False
euro = Sheets("Sayfa3").Range("Eur").Value
    Sheets("Sayfa1").Columns("W:X").NumberFormat = "#,##0.00 [$€-1]"
    Sheets("Sayfa3").Range("M29").NumberFormat = "#,##0.00 [$€-1]"
    Sheets("Sayfa3").Range("Tpb") = Sheets("Sayfa3").Range("Eur")
    'Range("A1").Select
Sheets("Sayfa3").Range("Tpbr") = "Teklif Para Birimi (EUR)"
Application.ScreenUpdating = True
hata:
End Sub
Sub MakroTL()
On Error GoTo hata
Application.ScreenUpdating = False
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
    Sheets("Sayfa1").Columns("W:X").NumberFormat = "#,##0.00"
    Sheets("Sayfa3").Range("M29").NumberFormat = "#,##0.00"
    Sheets("Sayfa3").Range("Tpb") = 1
'Range("A1").Select
Sheets("Sayfa3").Range("Tpbr") = "Teklif Para Birimi (TL)"
Application.ScreenUpdating = True
hata:
End Sub
Sub kateklehepsi()
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
i = 2: n = Range("B65536").End(xlUp).row
Application.ScreenUpdating = False: Application.Calculation = xlManual

Do Until i > n
If Cells(i, 2) = "BÖLÜM ADI/NO:" And Cells(i, 5) > 1 Then
k = Cells(i, 5).Address: ks = Cells(i, 5)
    Cells(i, 5) = ks: Cells(i, 5).NumberFormat = "#,##0 ""Adet"""
bas1:
    i = i + 1
    If Cells(i, 2) = "BÖLÜM TOPLAMI:" Or i > Range("B65536").End(xlUp).row Then: GoTo git1

    a = Cells(i, 5).Formula: a = Replace(a, "=", "")
    For j = 1 To Len(a)
    If Not Right(a, 1) = "*" Then a = Left(a, Len(a) - 1) Else Cells(i, 5).Formula = Left(a, Len(a) - 1): GoTo git2
    Next
git2:
    Dim s As String
    s = Replace(Cells(i, 5), ",", ".")
    's = Format(Application.WorksheetFunction.RoundUp(Cells(i, 5), 0), "#,##0.00")
    Cells(i, 5).Font.ColorIndex = 32
    If Not Cells(i, 5).HasFormula Then Cells(i, 5).Formula = "=" & s & "*" & k: GoTo bas1
End If
git1:
i = i + 1
Loop
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub katekle()
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
i = Selection.row: n = 2
If i > Range("B65536").End(xlUp).row Then Exit Sub
Do Until i < n
If Cells(i, 2) = "BÖLÜM ADI/NO:" Then
If Cells(i, 5) > 1 Then k = Cells(i, 5).Address: ks = Cells(i, 5): GoTo git1 Else Exit Sub
End If
i = i - 1
Loop
git1:
Application.ScreenUpdating = False: Application.Calculation = xlManual
Call katsil
Cells(i, 5) = ks: Cells(i, 5).NumberFormat = "#,##0 ""Adet"""
i = i + 1: n = Range("B65536").End(xlUp).row
Dim s As String
Do Until i > n
    If Cells(i, 2) = "BÖLÜM TOPLAMI:" Then GoTo git3
    If Cells(i, 5) <> "" Then
    s = Replace(Cells(i, 5), ",", ".")
    's = Format(Application.WorksheetFunction.RoundUp(Cells(i, 5), 2), "#,##0.00")
    Cells(i, 5).Font.ColorIndex = 32
    If Not Cells(i, 5).HasFormula Then Cells(i, 5).Formula = "=" & s & "*" & k
    End If
i = i + 1
Loop
git3:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub katsilhepsi()
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
i = 2: n = Range("B65536").End(xlUp).row
Application.ScreenUpdating = False: Application.Calculation = xlManual
Do Until i > n
If Cells(i, 2) = "BÖLÜM ADI/NO:" And Cells(i, 5) > 1 Then
    Cells(i, 5) = "": Cells(i, 5).NumberFormat = "#,##0"
bas1:
    i = i + 1
    If Cells(i, 2) = "BÖLÜM TOPLAMI:" Or i > Range("B65536").End(xlUp).row Then: GoTo git1
    a = Cells(i, 5).Formula: a = Replace(a, "=", "")
    Cells(i, 5).Font.ColorIndex = 1
    For j = 1 To Len(a)
    If Not Right(a, 1) = "*" Then a = Left(a, Len(a) - 1) Else Cells(i, 5).Formula = Left(a, Len(a) - 1): GoTo bas1
    Next
End If
git1:
i = i + 1
Loop
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub katsil()
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
i = Selection.row: n = 2
If i > Range("B65536").End(xlUp).row Then Exit Sub
Do Until i < n
If Cells(i, 2) = "BÖLÜM ADI/NO:" Then
Cells(i, 5) = "": Cells(i, 5).NumberFormat = "#,##0": GoTo git1
End If
i = i - 1
Loop
git1:
i = i + 1: n = Range("B65536").End(xlUp).row
Do Until i > n
    If Cells(i, 2) = "BÖLÜM TOPLAMI:" Then Exit Sub
    a = Cells(i, 5).Formula: a = Replace(a, "=", "")
    Cells(i, 5).Font.ColorIndex = 1
    For j = 1 To Len(a)
    If Not Right(a, 1) = "*" Then a = Left(a, Len(a) - 1) Else Cells(i, 5).Formula = Left(a, Len(a) - 1): GoTo git2
    Next
git2:
i = i + 1
Loop
End Sub
Sub katsilxx()
On Error GoTo hata
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
n = Range("B65536").End(xlUp).row
i = 1
Do Until i > n
If Cells(i, 2) = "BÖLÜM ADI/NO:" And Cells(i, 5) > 1 Then
k = Cells(i, 5)
git:
i = i + 1
If i > n Then Exit Sub
    If Not Cells(i, 2) = "BÖLÜM TOPLAMI:" Then
    'If Cells(i, 2) = "PMB" Then GoTo git
s = Cells(i, 5)
    If Cells(i, 5).HasFormula Then Cells(i, 5) = s / k
    'If Cells(i, 5).HasFormula Then Cells(i, 5) = s / k
GoTo git
End If
End If
i = i + 1
Loop
hata:
End Sub
Sub printresetxxx()
    son = Range("B65536").End(xlUp).row
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$W1:$X" & son
End Sub
Sub Macro75(control As IRibbonControl) ' Montaj Fiyatlarını Günceller
On Error GoTo hata
If Not ActiveWorkbook.ActiveSheet.CodeName = "CML" Then Exit Sub
Dim b
Dim aranan
Dim s
eski = ActiveWorkbook.Name
If Not WorkbookOpen("Montaj Fiyatları.xls") Then
Workbooks.Open "C:\Belgelerim\CEMEX\Parametreler\Montaj Fiyatları.xls"
End If
yeni = "Montaj Fiyatları.xls"
Workbooks(yeni).Activate
Worksheets("Parametreler").Select
ara = Range("A65536").End(xlUp).row
Windows(eski).Activate
bul = Range("A65536").End(xlUp).row
Range("A2").Select
For i = 2 To bul
'Workbooks(eski).Activate
Range("A" & i).Select
kod = Left(ActiveCell.Value, 2) & "-" & Range("I" & i)

    'kod = Left(Range("A" & i), 6)
    If kod = "" Or Cells(i, "d") = "" Then GoTo git
        For Each aranan In Workbooks(yeni).Worksheets("Parametreler").Range("A" & "2 :" & "A" & ara)
            If aranan = kod Then

            s = aranan.row
            b = Workbooks(yeni).Worksheets("Parametreler").Range("H" & s)
            Range("H" & i) = b: Range("A" & i).Font.ColorIndex = 3: Range("H" & i).Font.ColorIndex = 3: Range("I" & i).Font.ColorIndex = 3
            GoTo git
            End If
        Next
git:
Next i
Range("A2").Select
Workbooks(yeni).Close
hata:
End Sub
Sub teklifformat(control As IRibbonControl)
On Error Resume Next
Call ilhan
Call yeniteklif
End Sub
Sub Macro170(control As IRibbonControl)
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "T1" Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then UFObara1.Show: Exit Sub
Call ilhan: UFObara1.Show
End Sub
Sub PrintPDf()
On Error Resume Next
For i = 1 To Worksheets.Count
If Sheets(i).Name = "Veriler" Then GoTo sfvar:
Next i
msg = MsgBox("Teklif şablon dosyası seçiniz.", vbQuestion, "scngnr@gmail.com")
Exit Sub
sfvar:
Call pdfkaydet
End Sub
Sub Macro171(control As IRibbonControl)
On Error Resume Next
Call ilhan
    dt = ActiveWorkbook.Name
Application.ScreenUpdating = False
    sk1 = "Kablo ve Klemensler.xlsb"
    Dim fn, fm1
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
    fn = fm1 & "\Otomatik Seçim\Kablo ve Klemensler.xlsb"
    If Not WorkbookOpen((sk1)) Then Workbooks.Open fn: Application.Windows(sk1).Visible = False
    Workbooks(dt).Activate
UFOSARF.Show
End Sub
Sub teklifsablonkaydet()
On Error Resume Next
For i = 1 To Worksheets.Count
If Sheets(i).CodeName = "T1" Then Call sablonkaydet1: Exit Sub
Next i
For i = 1 To Worksheets.Count
If Sheets(i).Name = "Veriler" Then GoTo sfvar:
Next i
msg = MsgBox("Teklif şablon dosyası seçiniz.", vbQuestion, "scngnr@gmail.com")
Exit Sub
sfvar:
dsa = Sheets(i).Range("B31") & "-" & "Şablonu"
If Sheets(i).Range("B30") = "" Then Sheets(i).Range("B30") = Environ("USERPROFILE") & "\Desktop"
ds = InputBox("Kaydedilecek klasör : " & Sheets(i).Range("B30"), "Dosya Adı", dsa)
If ds = "" Then Exit Sub
'..
kl = Sheets(i).Range("B30") & "\" & ds & ".xlsx"
ActiveWorkbook.SaveAs kl
End Sub
Sub sablonkaydet1() 'Yeni Teklif şablonu
On Error Resume Next
Worksheets("Sayfa1").Select
ds = InputBox("C:\Users\İlhan Şirin\OneDrive\Belgeler\Özel Office Şablonları\", "Şablon Adı", "Yeni Teklif V1.2.xltx")
If ds = "" Then Exit Sub
Application.ScreenUpdating = False: Application.DisplayAlerts = False
    ChDir "C:\Users\İlhan Şirin\OneDrive\Belgeler\Özel Office Şablonları"
    ActiveWorkbook.SaveAs fileName:="C:\Users\İlhan Şirin\OneDrive\Belgeler\Özel Office Şablonları\" & ds _
        , FileFormat:=xlOpenXMLTemplate, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False, AccessMode:=xlNoChange
    ChDir "C:\Belgelerim\Cemex\Yeni Teklif Şablonları"
    ActiveWorkbook.SaveAs fileName:="C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx" _
        , FileFormat:=xlOpenXMLTemplate, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False, AccessMode:=xlNoChange
Application.ScreenUpdating = True: Application.DisplayAlerts = True
MsgBox "Şablon dosyası kaydedildi. ", vb, "scngnr@gmail.com"
End Sub
Sub pdfkaydet()
Dim varResult As Variant
Set dc = CreateObject("Scripting.FileSystemObject")
If ActiveWorkbook.path = "" Then
dyol = Sheets("Veriler").Range("B30")
If Not Left(dyol, 1) = "\" Then dyol = Sheets("Veriler").Range("B30") & "\"
Yol = dyol & Sheets("Veriler").Range("B31") & ".pdf"
Else
Yol = ActiveWorkbook.path & "\" & Split(ActiveWorkbook.Name, ".")(0) & ".pdf"
End If
varResult = Application.GetSaveAsFilename(Yol, FileFilter:="PDF Dosyası (*.pdf), *.pdf", Title:="PDF Kaydet")
If varResult = False Then Exit Sub
a = dc.FileExists(varResult)
If a = True Then
cvp = MsgBox("Bu isimde bir dosya mevcut, devam edilsin mi ?", vbYesNo)
If cvp = vbYes Then GoTo devam Else Exit Sub
Else
GoTo devam
End If
devam:
If varResult <> False Then
Sheets("Veriler").Visible = False
    With ActiveWorkbook
    .ExportAsFixedFormat _
    Type:=xlTypePDF, _
    fileName:=varResult, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
    End With
    CreateObject("Shell.Application").Open (varResult)
End If
End Sub