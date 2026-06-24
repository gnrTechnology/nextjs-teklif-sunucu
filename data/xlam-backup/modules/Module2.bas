Option Explicit
Public tdosya
Public pfc As Byte
Public dtx As Byte
Public tp1 As Byte
Public tsi1 As Byte
Public sa As Byte
Public ssno As Integer
'Public dt
Function WorkbookOpen(workbookName As String) As Boolean
    WorkbookOpen = False
    On Error GoTo WorkBookNotOpen
    If Len(Application.Workbooks(workbookName).Name) > 0 Then
        WorkbookOpen = True
        Exit Function
    End If
WorkBookNotOpen:
End Function
Sub Macro2(control As IRibbonControl)
    Call dosyaac
End Sub
Sub Macro32(control As IRibbonControl)
On Error GoTo hata
ssno = 1
UFTH.Show
hata:
End Sub
Sub Macropanofiyat(control As IRibbonControl) 'Pano giriş
On Error Resume Next
Call ilhan
pfc = GetSetting("ilhan", "Settings", "pfc")
If pfc = 0 Then Call panogir0
If pfc = 1 Then Call panogir1
End Sub
Sub Macro311(control As IRibbonControl) 'Pano carpandan giriş
Call ilhan
SaveSetting "ilhan", "Settings", "pfc", 1
Call panogir1
End Sub
Sub Macro312(control As IRibbonControl) 'Pano listeden giriş
Call ilhan
SaveSetting "ilhan", "Settings", "pfc", 0
pfc = 0: Call UFOPAN02.Show
End Sub
Sub panogir1()
pfc = 1: UFOPAN00.Show
End Sub
Sub panogir0()
On Error GoTo hata
Dim marka As String, dsmyol As String
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
dt = ActiveWorkbook.Name
Application.ScreenUpdating = False
mlz = GetSetting("ilhan", "Settings", "panodizini")
dsmyol = fm1 & "\Malzeme Listeleri\4\" & mlz
If Not WorkbookOpen(dsmyol) Then
    Dim ds, da
    Set ds = CreateObject("Scripting.FileSystemObject")
    da = ds.FileExists(dsmyol)
    If da = False Then Exit Sub
    Workbooks.Open dsmyol
End If
Application.Windows(mlz).Visible = False
Workbooks(mlz).Worksheets("Sayfa1").AutoFilterMode = False
    UFOPAN00.Show
hata:
End Sub
Sub teklifkur(control As IRibbonControl) 'döviz kurları
On Error GoTo hata
Call ilhan
    UFKur.Show
hata:
End Sub
Sub Macro4(control As IRibbonControl) 'hata denetimi
On Error GoTo hata
Call ilhan
    Call Hatabul
hata:
End Sub
Sub Macro5(control As IRibbonControl) 'Teklif Bilgileri
On Error GoTo hata
Call ilhan
ssno = 2
UFTH.Show
'UFTH.MultiPage1.Value = 0
Call UFTH.toolbarbutton2
hata:
End Sub
Sub Macro72(control As IRibbonControl) 'formülleri düzenler
On Error GoTo hata
If Not ActiveWorkbook.ActiveSheet.CodeName = "CML" Then Exit Sub
    Call formulyap
hata:
End Sub
Sub Macro73(control As IRibbonControl) 'renkleri düzenler
On Error GoTo hata
If Not ActiveWorkbook.ActiveSheet.CodeName = "CML" Then Exit Sub
    Call yazırenk
hata:
End Sub
Sub Macro10(control As IRibbonControl)
    MsgBox "This is Macro10"
End Sub
Sub Macro11(control As IRibbonControl)
    MsgBox "This is Macro11"
End Sub
Sub Macro13(control As IRibbonControl)
On Error GoTo hata
Call ilhan
'dt = ActiveWorkbook.Name
    'Sheets("Sayfa1").Select
UFmy.MultiPageP2.Value = 0
    UFmy.Show
hata:
End Sub
Sub Macro14(control As IRibbonControl)
On Error GoTo hata
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then GoTo atla1
Call ilhan
atla1:
dt = ActiveWorkbook.Name: Sheets("Sayfa1").Select: UFmd.Show
hata:
End Sub
Sub Macro15(control As IRibbonControl) 'analiz sayfası
On Error GoTo hata
Dim giris
Call ilhan
If Sheets("Sayfa3").Visible = False Then
giris = InputBox("Şifre giriniz.", "Analiz Sayfası Şifre Penceresi", "")
ActiveWorkbook.Unprotect Password:=giris
ActiveWorkbook.Unprotect
Sheets("Sayfa3").Visible = True
End If
If ActiveSheet.Name = "Sayfa3" Then
Application.Sheets("Sayfa1").Activate
Else
Worksheets("Sayfa3").Activate
End If
hata:
End Sub
Sub Macro500(control As IRibbonControl) 'analiz sayfası
On Error GoTo hata
Dim giris
Call ilhan
If Sheets("Sayfa3").Visible = False Then Exit Sub
giris = InputBox("Şifre giriniz.", "Analiz Sayfası Şifre Penceresi", "")
If giris = "" Then Exit Sub
Sheets("Sayfa3").Visible = False
ActiveWorkbook.Protect Password:=giris, Structure:=True, Windows:=True
'Sheets("Sayfa3").Visible = True
Application.Sheets("Sayfa1").Activate
hata:
End Sub
Sub Macro16(control As IRibbonControl) 'malzeme değişim dosyaları
On Error GoTo hata
Dim cod1
cod1 = ActiveWorkbook.ActiveSheet.CodeName
dtx = 1
If Left(cod1, 3) = "CML" Or Left(cod1, 2) = "T1" Or Left(cod1, 3) = "OTM" Then
 dt = ActiveWorkbook.Name: Sheets("Sayfa1").Select
 If Cells(Selection.row, "B") = "" Or Selection.row = 1 Or Cells(Selection.row, "D") = "" Then
 MsgBox "    Malzemenin olduğu bir satırı seçin. ", vbCritical, "scngnr@gmail.com"
 dt = "": Exit Sub
 End If
Else
 If Selection.Cells = "" Or Selection.row = 1 Or Cells(Selection.row, "D") = "" Then
 MsgBox "    Malzemenin olduğu bir satırı seçin. ", vbCritical, "scngnr@gmail.com"
 Exit Sub
 End If
Dim dtxs
dtxs = GetSetting("ilhan", "Settings", "dtxs")
If dtxs = "" Then dtxs = "C-E-F"
dtxs = InputBox("Sutun Seçim Sıralaması :" & vbCr & "Açıklama - Marka - Liste Fiyatı", "Sutun Seçimi", dtxs)
If dtxs = "" Then Exit Sub
SaveSetting "ilhan", "Settings", "dtxs", dtxs
dt = ActiveWorkbook.Name
dtx = 2
End If
UFMZ.Show
hata:
End Sub
Sub Macro17(control As IRibbonControl) 'kayıtlar
On Error GoTo hata
Application.ScreenUpdating = False
If Not WorkbookOpen("Kayıtlar.xlsb") Then
    Application.ScreenUpdating = False
    Workbooks.Open "C:\Belgelerim\CEMEX\Kayıtlar\Kayıtlar.xlsb"
End If
Windows("Kayıtlar.xlsb").Activate
Windows("Kayıtlar.xlsb").Visible = False
    UserFormT4.Show
hata:
Application.ScreenUpdating = True
End Sub
Sub Macro18(control As IRibbonControl) 'toplamlar
On Error GoTo hata
Call ilhan
If Not Cells(1, "X") = "Toplam Fiyat" Then Exit Sub
    Call AraToplamlar
    MsgBox ("    Mevcut  listede  bulunan  bölüm toplamları düzeltildi.  "), , "scngnr@gmail.com"
hata:
End Sub
Sub Macro19(control As IRibbonControl) 'araya kopyala
On Error GoTo hata
'Call ilhan
If ActiveSheet.Name = "Sayfa3" Then Exit Sub
Selection.EntireRow.Copy
Dim msg
msg = MsgBox("    Seçili olan bölümü boş satırın üzerine eklemek istiyor musunuz? ", vbYesNo, "İlave Etme/Yapıştırma  ")
If msg = vbNo Then GoTo hata
Application.ScreenUpdating = False
Cells(2, "B").End(xlDown).Offset(1, 0).EntireRow.Select
Selection.Insert Shift:=xlDown 'satır ilave ederek yapıştırır.
Application.ScreenUpdating = True
ActiveCell.Offset(1, 1).Select
hata:
Application.CutCopyMode = False
End Sub
Sub Macro190(control As IRibbonControl) 'kopyala
Call ilhan
If ActiveSheet.Name = "Sayfa3" Then Exit Sub
Selection.EntireRow.Copy
End Sub
Sub Macro191(control As IRibbonControl) 'yapıştır
Call ilhan
On Error GoTo hata
Dim a
Application.ScreenUpdating = False: Application.DisplayAlerts = False
a = Selection.row
Range("A" & a).Select
Selection.Insert Shift:=xlDown 'satır ilave ederek yapıştırır.
ActiveCell.Offset(1, 1).Select
hata:
Application.CutCopyMode = False
Application.ScreenUpdating = True: Application.DisplayAlerts = True
End Sub
Sub Macro1621(control As IRibbonControl) 'işçilik ekle2
Call ilhan
Call isciliksatır_gir1
End Sub
Sub Macro1641(control As IRibbonControl) 'sarf gir1
Call ilhan
Call sarfsatır_gir1
End Sub
Sub MacroadT(control As IRibbonControl) 'adet değiştirme
Call ilhan
UFadT.Show
End Sub
Sub MacroDAD(control As IRibbonControl) 'teklif biçimlendirme
On Error GoTo hata
Dim i As Integer
For i = 1 To Worksheets.Count
If Worksheets(i).Name = "Sayfa3" Then GoTo git1
Next i
For i = 1 To Worksheets.Count
If Worksheets(i).Name = "Sayfa1" Then GoTo git2
Next i
GoTo git3
git1:
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then UFDAD.Show: Exit Sub
git2:
If Left(ActiveWorkbook.Sheets("Sayfa1").CodeName, 2) = "TM" Then UFDAD.Show: Exit Sub
git3:
Dim cdm
cdm = ActiveWorkbook.ActiveSheet.CodeName
If Left(cdm, 3) = "CML" Or Left(cdm, 3) = "OTM" Then UFDAD.Show: Exit Sub
hata:
Call MsgBox("Bu dosyada işlem yapamazsınız.", vbQuestion, "scngnr@gmail.com")
End Sub
Sub MacroDAD0(control As IRibbonControl) 'sayfa düzenleme
UFDAD0.Show
End Sub
Sub MacroDAD01(control As IRibbonControl) 'sayfa düzenleme
UFDAD1.Show
End Sub
Sub Macro169(control As IRibbonControl) 'bakır satır ekleme
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "T1" Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call bakir_gir
End Sub
Sub Macro21(control As IRibbonControl) 'kat ekle
Call katekle
End Sub
Sub Macro211(control As IRibbonControl) 'kat ekle hepsi
Call kateklehepsi
End Sub
Sub Macro22(control As IRibbonControl) 'kat sil
Call katsil
End Sub
Sub Macro221(control As IRibbonControl) 'kat sil hepsi
Call katsilhepsi
End Sub
Sub Macro23(control As IRibbonControl) 'para birimi TL
Call ilhan
Sheets("Sayfa3").Range("Tpbr") = "Teklif Para Birimi (TL)" 'Teklif Para Birimi
Call MakroTL
End Sub
Sub Macro24(control As IRibbonControl) 'para birimi $
Call ilhan
Sheets("Sayfa3").Range("Tpbr") = "Teklif Para Birimi (USD)" 'Teklif Para Birimi
Call MakroDOLAR
End Sub
Sub Macro25(control As IRibbonControl) 'para birimi €
Call ilhan
Sheets("Sayfa3").Range("Tpbr") = "Teklif Para Birimi (EUR)" 'Teklif Para Birimi
Call MakroEURO
End Sub
Sub Macro26(control As IRibbonControl) 'kurlar
Shell ("C:\Program files\Internet Explorer\Iexplore.exe http://www.tcmb.gov.tr/wps/wcm/connect/tr/tcmb+tr/main+page+site+area/bugun"), vbNormalFocus
End Sub
Sub baslık1(control As IRibbonControl) '23.08.2012 başlık ekle BÖLÜM ADI/NO
On Error Resume Next
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "T1" Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" _
Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "OT" Then
Dim a As Integer
Dim msg
a = Selection.row
Cells(a, 2).Range("A1").Select
    'If ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM ADI/NO:" Or ActiveCell.Offset(-1, 0).FormulaR1C1 = "BÖLÜM ADI/NO:" Then
    If ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM ADI/NO:" Then
    msg = MsgBox("Burada bu işlemi yapamazsınız.", vbQuestion, "scngnr@gmail.com")
    Exit Sub
    Else
    'If TextBox2.Value = "" Then TextBox2.Value = Selection.Row - 1
    UserFormAD.Caption = "Bölüm Adı / No:"
    If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "OT" Then UserFormAD.Frame4.BackColor = &H96A446
    UserFormAD.Show
    End If
End If
End Sub
Sub baslık2(control As IRibbonControl) '23.08.2012 başlık ekle PROJE ADI/NO
On Error Resume Next
'Call ilhan
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "T1" Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then
Dim a
Dim msg
a = Selection.row
Cells(a, 2).Range("A1").Select
    If ActiveCell.Offset(0, 0).FormulaR1C1 = "PROJE ADI/NO:" Or ActiveCell.Offset(-1, 0).FormulaR1C1 = "PROJE ADI/NO:" Then
    msg = MsgBox("Burada bu işlemi yapamazsınız.", vbQuestion, "scngnr@gmail.com")
    Exit Sub
    Else
    UserFormAD.Caption = "Proje Adı / No:"
    UserFormAD.Show
 End If
End If
End Sub
Sub baslık3(control As IRibbonControl) '23.08.2012 başlık ekle GRUP ADI
On Error Resume Next
'Call ilhan
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "T1" Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" _
Or Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "OT" Then
Dim ds
ds = InputBox("Grup Adı / No Giriniz.", "Klasör / Grup Adı / No", "")
If ds = "" Then Exit Sub
If Selection.row = 1 Then Range("B2").Select
Selection.EntireRow.Insert
Dim Y As Integer
Y = Selection.row
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then
Range("A" & Y & ":J" & Y).Select: Selection.Borders.LineStyle = xlContinuous
Else
Range("A" & Y & ",B" & Y & ":E" & Y & ",F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Select: Selection.Borders.LineStyle = xlContinuous
Range("F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Borders(xlInsideVertical).LineStyle = xlNone
Range("A" & Y & ":D" & Y).NumberFormat = "@"
End If
Range("B" & Y & ":E" & Y).Borders(xlInsideVertical).LineStyle = xlNone
Selection.Interior.Pattern = xlNone: Selection.RowHeight = 12.75: Selection.Font.Size = 9: Selection.Font.Bold = True
Selection.Font.ColorIndex = 53
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.Offset(0, 1).FormulaR1C1 = ds
ActiveCell(1, 1).Value = ".": ActiveCell(1, 0).Value = "Grup": ActiveCell(1, 4).Value = "(1 Grup)"
End If
End Sub
Sub toplam1(control As IRibbonControl) '23.08.2012 Toplam ekle
On Error Resume Next
'Call ilhan
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
Dim a
Dim msg
Dim Y As Integer
Y = Selection.row
Cells(Y, 2).Range("A1").Select
    If ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM TOPLAMI:" Or ActiveCell.Offset(-1, 0).FormulaR1C1 = "BÖLÜM TOPLAMI:" Then
               msg = MsgBox(" Burada bu işlemi yapamazsınız.", vbQuestion, "scngnr@gmail.com")
               Exit Sub
    End If
    Selection.EntireRow.Insert
    ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM TOPLAMI:"
'Biçimlemeler'--
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then
Range("A" & Y & ":J" & Y).Select: Selection.Borders.LineStyle = xlContinuous
Else
Range("A" & Y & ",B" & Y & ":E" & Y & ",F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Select: Selection.Borders.LineStyle = xlContinuous
Range("F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Borders(xlInsideVertical).LineStyle = xlNone
Range("W" & Y & ":X" & Y).HorizontalAlignment = xlRight
End If
Range("B" & Y & ":E" & Y).Borders(xlInsideVertical).LineStyle = xlNone
Selection.Interior.Pattern = xlNone: Selection.RowHeight = 12.75: Selection.Font.Size = 9: Selection.Font.Bold = True
Selection.Font.ColorIndex = 11
Call AraToplamlar
End Sub
Sub toplam2(control As IRibbonControl) 'genel toplam ekle
On Error Resume Next
Dim Y
Call ilhan
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
Y = Range("B65536").End(xlUp).row + 1
     If Cells(Y - 1, 2) = "GENEL TOPLAM:" Then Cells(Y - 1, 2).EntireRow.Delete: Y = Y - 1
     Cells(Y, 2).Select
     Cells(Y, 2).FormulaR1C1 = "GENEL TOPLAM:"
     Cells(Y, 24) = "=SUM(R2C24:R[-1]C)-SUMIF(R2C2:R[-1]C[-22],""=BÖLÜM TOPLAMI:"",R2C24:R[-1]C)"
'Biçimlemeler'--
Range("A" & Y & ",B" & Y & ":E" & Y & ",F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Select
    Selection.Borders.LineStyle = xlContinuous
    Range("B" & Y & ":E" & Y & ",W" & Y & ":X" & Y).Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Interior.Pattern = xlNone
    Selection.RowHeight = 12.75
    Selection.Font.Size = 9
    Selection.Font.Bold = True
    Selection.Font.ColorIndex = 11
Call AraToplamlar
    Cells(Y, 2).Select
End Sub
Sub Macro28(control As IRibbonControl) 'sablonlar
On Error Resume Next
Call ilhan
ssno = 2
UserFormS1.Show
End Sub
Sub Macro74(control As IRibbonControl)
UFDD.Show
End Sub
Sub Macro741(control As IRibbonControl)
UFDD1.Show
End Sub
Sub Macro742(control As IRibbonControl)
On Error GoTo hata
If Date > CDate("08.03.2026") Then Exit Sub
UFDD2.Show
Exit Sub
hata:
MsgBox "Dosya Hatası!", vbCritical
End Sub
Sub panoboyut(control As IRibbonControl) 'pano boyut
On Error Resume Next
If Selection.row > Range("B65536").End(xlUp).row Or Selection.row < 2 Then Exit Sub
UFOPAN11.Show
hata:
UFOPAN11.FRP.ScrollTop = 0
End Sub
Sub Macro43(control As IRibbonControl)
On Error Resume Next
Sheets("Sayfa1").Select '
tp1 = 1: sa = 1: Call iscilik_gir1: Call sarf_gir1: Call amb_gir1: tp1 = 0
End Sub
Sub Macro431(control As IRibbonControl)
On Error Resume Next
Sheets("Sayfa1").Select '
tp1 = 1:  tsi1 = 1: Call iscilik_gir1: Call sarf_gir1: tp1 = 0: tsi1 = 0
End Sub
Sub Macro44(control As IRibbonControl)
Sheets("Sayfa1").Select '
Call isciliksarfambsil
End Sub
Sub Macro45(control As IRibbonControl)
Sheets("Sayfa1").Select '
Call pfkod
End Sub
Sub Macro501(control As IRibbonControl)
On Error Resume Next
Call PrintPDf
End Sub
Sub Macro502(control As IRibbonControl)
On Error Resume Next
Call teklifsablonkaydet
End Sub
Sub Macro503(control As IRibbonControl)
On Error Resume Next
Dim i As Integer
For i = 1 To Worksheets.Count
If Sheets(i).Name = "Veriler" Then UserFormS2.Show: Exit Sub
Next i
End Sub
Sub Stokkontrol1(control As IRibbonControl) 'stok kontrol
On Error Resume Next
Dim d1m, yolm, msg
d1m = GetSetting("ilhan", "Settings", "deposabitdosya")
yolm = GetSetting("ilhan", "Settings", "depodizini") & "\"
Dim ds, a, n
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists(yolm & d1m)
If a <> True Then msg = MsgBox("Stok dosyası mevcut değil!", vbInformation, "scngnr@gmail.com"): Exit Sub
n = Selection.row
If Range("B" & n).Value = "" Or Range("B" & n) = "BÖLÜM TOPLAMI:" Or Range("B" & n) = "BÖLÜM ADI/NO:" Or Range("B" & n) = "GENEL TOPLAM:" Then
msg = MsgBox("Sipariş kodu olan satırı seçiniz!", vbInformation, "scngnr@gmail.com"): Exit Sub: End If
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
Workbooks.Open((yolm & d1m), False) = True: Application.Windows(ActiveWorkbook.Name).Visible = False
Dim tms
tms = WorksheetFunction.SumIf(Workbooks(d1m).ActiveSheet.Range("B:B"), Range("B" & n).Value, Workbooks(d1m).ActiveSheet.Range("E:E"))
Workbooks(d1m).Close (False)
MsgBox "Bakılan Dosya Adı : " & d1m & vbLf & "Bakılan Ürün : " & ActiveSheet.Range("B" & n).Value & vbLf & "Mevcut Stok Ad. : " & tms
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub