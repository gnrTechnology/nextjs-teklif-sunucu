Dim a 'bölüm adı
Dim b 'boş satır
Dim kod '
Dim kdi '
Dim mdt As Integer
Private Sub CBkablo_Click()
Windows(dt).Activate: Sheets("Sayfa1").Select '
Call aralık
'Call modulara
If mdt > 0 Then Call gir1
End Sub
Private Sub CBklemens_Click()
Windows(dt).Activate: Sheets("Sayfa1").Select '
Call aralık
'Call modulara
If mdt > 0 Then Call gir2
End Sub
Private Sub CommandButton1_Click()
Call aralık
Call modulara
Call kablometre
End Sub
Private Sub TextBoxM1_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM2_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM3_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM4_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM5_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM6_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM7_Change()
Call uf_modulara
End Sub
Private Sub TextBoxM8_Change()
Call uf_modulara
End Sub
Sub uf_modulara()
TextBoxMDT = 0
For i = 1 To 8 '
If Controls("TextBoxM" & i) <> "" Then TextBoxMDT = val(TextBoxMDT) + val(Controls("TextBoxM" & i).Value)
mdt = TextBoxMDT
Next i
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Application.ScreenUpdating = False
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
Call aralık
Call modulara
Dim dosya
Dim n As Integer
dosya = dir("C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Klemensler\*.txt")
ListBoxKM.SetFocus
Do While dosya <> ""
dosyaad = Replace(dosya, ".txt", "")
   ListBoxKM.AddItem dosyaad
    dosya = dir
    n = n + 1
Loop
If ListBoxKM.ListCount > 0 Then ListBoxKM.Selected(0) = True
Call kablokesit
Call kablometre
Application.ScreenUpdating = True
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '
On Error Resume Next
Application.ScreenUpdating = False
    Windows(dt).Activate
    Windows(sk1).Close False
    Application.Windows(sk1).Visible = True
    sl1 = Empty
End Sub
Sub aralık()
On Error Resume Next
Application.ScreenUpdating = False
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then Exit Sub
i = Selection.row: n = 2: a = 2: b = Empty
'If i > Range("B65536").End(xlUp).Row Then Exit Sub
Do Until i < n
If Cells(i, 2) = "BÖLÜM ADI/NO:" And Cells(i, 1) = "" Then: UFOSARF.Caption = Cells(i, 3): a = Cells(i, 5).row: GoTo git1:
i = i - 1
Loop
git1:
n = Range("B65536").End(xlUp).row: 'i = Selection.Row
Do Until i > n
If Cells(i, 2) = "BÖLÜM TOPLAMI:" And Cells(i, 1) = "" Then: b = Cells(i, 5).row: GoTo git2:
i = i + 1
Loop
If b = Empty Then b = n + 1
git2:
'--
End Sub
Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
If ListBoxKM.ListCount = 0 Then Exit Sub
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Klemensler\" & ListBoxKM & ".txt"
End Sub
Private Sub Image2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Klemensler"
End Sub
Private Sub Image3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Kablolar\Kablo Kesitleri.txt"
End Sub
Private Sub Image4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Kablolar\Kablo Uzunluk.txt"
End Sub
Private Sub ListBoxKM_Click()
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, satır2 As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Klemensler\" & ListBoxKM & ".txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox ListBoxKM & ".txt /" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxKT.Clear
    Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxKT.AddItem ayır(i)
'--
'kaçıncı = InStr(1, Rky, ";")
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxKT.List(satır - 1, n) = ayır(i + n)
Next n
'--
        'If UBound(ayır) <> 0 Then ListBoxKT.List(satır - 1, 1) = ayır(i + 1)
        'If UBound(ayır) <> 0 Then ListBoxKT.List(satır - 1, 2) = ayır(i + 2)
'--
        satır = satır + 1
    Loop
    kt = 1
For i = 1 To 7
Controls("SBK" & i).Max = ListBoxKT.ListCount - 1
pta = Controls("LabelPTA" & i).Caption
    Do Until kt = ListBoxKT.ListCount
    If CDbl(ListBoxKT.List(kt - 1, 0)) >= pta Then: GoTo git1
    kt = kt + 1
    Loop
git1:
  Controls("TextBoxK" & i) = ListBoxKT.List(kt - 1, 1)
  Controls("TextBoxK" & i).ControlTipText = ListBoxKT.List(kt - 1, 2)
  Controls("SBK" & i).Value = kt - 1
Next i
  Controls("SBK" & 8).Max = ListBoxKT.ListCount - 1
  Controls("TextBoxK" & 8) = ListBoxKT.List(0, 1)
  Controls("TextBoxK" & 8).ControlTipText = ListBoxKT.List(0, 2)
  Controls("SBK" & 8).Value = 0
  
    Close #Ert
    Label238.Caption = "         " & ListBoxKM
End Sub
Sub kablokesit()
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Kablolar\Kablo Kesitleri.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Kablo Kesitleri.txt / Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        '--
        Controls("LabelPT" & satır).Caption = ayır(i)
        Controls("LabelPTA" & satır).Caption = ayır(i + 1)
        Controls("LabelPTA" & satır).ControlTipText = ayır(i + 2)
        '--
        satır = satır + 1
    Loop
    Close #Ert
End Sub
Sub kablometre()
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\Kablolar\Kablo Uzunluk.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Kablo Uzunluk.txt / Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    Do While Not EOF(Ert)
        Line Input #Ert, Rky
        Controls("TextBoxPTM" & satır) = Format(Rky, "#,##0.0")
        satır = satır + 1
    Loop
    Close #Ert
End Sub
Sub gir1()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
For i = 1 To 7
If Controls("TextBoxM" & i) <> "" Then
Cells(b, 2).EntireRow.Insert
Range("A" & b).FormulaR1C1 = "KB-auto"
Range("B" & b).FormulaR1C1 = Controls("LabelPTA" & i).ControlTipText
Range("E" & b).FormulaR1C1 = Controls("TextBoxPTM" & i) * Controls("TextBoxM" & i)
kod = Range("B" & b)
kdi = b
b = b + 1
Call formuller
End If
Next i
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
End Sub
Sub gir2()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
For i = 1 To 8
If Controls("TextBoxM" & i) <> "" Then
Cells(b, 2).EntireRow.Insert
Range("A" & b).FormulaR1C1 = "X1-auto"
Range("B" & b).FormulaR1C1 = Controls("TextBoxK" & i).ControlTipText
Range("E" & b).FormulaR1C1 = Controls("TextBoxM" & i)
kod = Range("B" & b)
kdi = b
b = b + 1
Call formuller
End If
Next i
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub SBK1_Change()
TextBoxK1 = ListBoxKT.List(SBK1.Value, 1): TextBoxK1.ControlTipText = ListBoxKT.List(SBK1.Value, 2)
End Sub
Private Sub SBK2_Change()
TextBoxK2 = ListBoxKT.List(SBK2.Value, 1): TextBoxK2.ControlTipText = ListBoxKT.List(SBK2.Value, 2)
End Sub
Private Sub SBK3_Change()
TextBoxK3 = ListBoxKT.List(SBK3.Value, 1): TextBoxK3.ControlTipText = ListBoxKT.List(SBK3.Value, 2)
End Sub
Private Sub SBK4_Change()
TextBoxK4 = ListBoxKT.List(SBK4.Value, 1): TextBoxK4.ControlTipText = ListBoxKT.List(SBK4.Value, 2)
End Sub
Private Sub SBK5_Change()
TextBoxK5 = ListBoxKT.List(SBK5.Value, 1): TextBoxK5.ControlTipText = ListBoxKT.List(SBK5.Value, 2)
End Sub
Private Sub SBK6_Change()
TextBoxK6 = ListBoxKT.List(SBK6.Value, 1): TextBoxK6.ControlTipText = ListBoxKT.List(SBK6.Value, 2)
End Sub
Private Sub SBK7_Change()
TextBoxK7 = ListBoxKT.List(SBK7.Value, 1): TextBoxK7.ControlTipText = ListBoxKT.List(SBK7.Value, 2)
End Sub
Private Sub SBK8_Change()
TextBoxK8 = ListBoxKT.List(SBK8.Value, 1): TextBoxK8.ControlTipText = ListBoxKT.List(SBK8.Value, 2)
End Sub
Private Sub SBmk_SpinUp()
On Error GoTo hata
For X = 1 To 7
Controls("TextBoxPTM" & X).Value = Format(Controls("TextBoxPTM" & X) + 0.2, "#,##0.0")
Next
hata:
End Sub
Private Sub SBmk_SpinDown()
On Error GoTo hata
For X = 1 To 7
If Controls("TextBoxPTM" & X).Value <= 0 Then Exit Sub
Controls("TextBoxPTM" & X).Value = Format(Controls("TextBoxPTM" & X) - 0.2, "#,##0.0")
Next
hata:
End Sub
Sub formuller()
On Error Resume Next
Application.Calculation = xlCalculationManual
'--
Set yeni = Workbooks(sk1).ActiveSheet
son = yeni.Range("B65536").End(xlUp).row
'--
i = kdi
Workbooks(dt).Worksheets("Sayfa1").Activate

Set aa = yeni.Range("B2:B" & son).Find(kod, LookIn:=xlValues, LookAt:=xlWhole)
Y = aa.row
'KUR'--
Range("F" & i).NumberFormat = yeni.Range("F" & Y).NumberFormat
kur = ""
If Range("F" & i).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & i).Font.ColorIndex = 3: kur = "*Usd" ' Sayfa3 de $ kuru
If Range("F" & i).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & i).Font.ColorIndex = 5: kur = "*Eur" ' Sayfa3 de € kuru
'GENEL VERİLER'--
Range("C" & i) = yeni.Range("C" & Y) 'Yapılacak İşin Cinsi
Range("D" & i) = yeni.Range("D" & Y) 'Üretici
Range("F" & i) = yeni.Range("F" & Y) 'Mlz. Br. Fiyat
Range("F" & i).NumberFormat = yeni.Range("F" & Y).NumberFormat
Range("G" & i) = yeni.Range("G" & Y) 'mlz.isk.
Range("H" & i) = yeni.Range("H" & Y) 'Adam/dk
Range("I" & i) = yeni.Range("I" & Y) 'Boyut
Range("J" & i).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt
Range("K" & i).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
Range("L" & i).FormulaR1C1 = bfyt & "*Osalt/100" & kur 'Mlz.Kar+1
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
'--
Dim tS3
Set tS3 = Workbooks(dt).Worksheets("Sayfa3")
If tS3.Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & i, "X" & i).NumberFormat = "#,##0.00"
If tS3.Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & i, "X" & i).NumberFormat = "#,##0.00 [$€-1]"
If tS3.Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & i, "X" & i).NumberFormat = "#,##0.00 [$$-C0C]"
'Genel Biçimlemeler'--
Range("A" & i & ":U" & i).Borders.LineStyle = xlContinuous
Range("W" & i & ":X" & i).Borders.LineStyle = xlContinuous
Range("A" & a & ":X" & i).Interior.Pattern = xlNone: Range("A" & i & ":U" & i).RowHeight = 12.75
Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.Bold = False
Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.ColorIndex = xlAutomatic
Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.Size = 9
Range("A" & i & ":D" & i).HorizontalAlignment = xlLeft
Range("E" & i & ":U" & i & ",W" & i & ":X" & i).HorizontalAlignment = xlRight
Range("A" & i & ":D" & i).NumberFormat = "@"
Range("E" & i).NumberFormat = "#,##0"
Range("F" & i).NumberFormat = "#,##0.00"
Range("G" & i).NumberFormat = "0.0%"
Range("H" & i).NumberFormat = "#,##0"
Range("J" & i & ":X" & i).NumberFormat = "#,##0.00"
If Range("F" & i).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & i).Font.ColorIndex = 3
If Range("F" & i).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & i).Font.ColorIndex = 5
'--
Application.Calculation = xlCalculationAutomatic
End Sub
Sub modulara()
On Error Resume Next
Application.ScreenUpdating = False
For i = 1 To 8 '
Controls("TextBoxM" & i) = ""
Next i
mdt = 0
Dim mds As Integer, mdm As Integer
Do Until a > b
If Left(Cells(a, 9), 1) = "M" Then
mdm = Right(Cells(a, 9), 2)
  If Left(Cells(a, 1), 2) = "FA" Then
    If mdm = 15 Then mdm = 10
    If mdm = 25 Then mdm = 20
    If mdm = 45 Then mdm = 30
  End If
'--
mds = mdm * Cells(a, 5) / 10
If Cells(a, "C") Like "*" & " 1A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x1A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 2A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x2A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 3A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x2A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 4A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x4A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 5A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x5A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 6A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x6A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 8A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x8A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 10A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x10A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 16A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x16A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 20A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x20A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 25A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & "x25A" & "*" Then TextBoxM1.Value = mds + val(TextBoxM1.Value)
If Cells(a, "C") Like "*" & " 32A" & "*" Then TextBoxM2.Value = mds + val(TextBoxM2.Value)
If Cells(a, "C") Like "*" & "x32A" & "*" Then TextBoxM2.Value = mds + val(TextBoxM2.Value)
If Cells(a, "C") Like "*" & " 40A" & "*" Then TextBoxM3.Value = mds + val(TextBoxM3.Value)
If Cells(a, "C") Like "*" & "x40A" & "*" Then TextBoxM3.Value = mds + val(TextBoxM3.Value)
If Cells(a, "C") Like "*" & " 50A" & "*" Then TextBoxM4.Value = mds + val(TextBoxM4.Value)
If Cells(a, "C") Like "*" & "x50A" & "*" Then TextBoxM4.Value = mds + val(TextBoxM4.Value)
If Cells(a, "C") Like "*" & " 63A" & "*" Then TextBoxM4.Value = mds + val(TextBoxM4.Value)
If Cells(a, "C") Like "*" & "x63A" & "*" Then TextBoxM4.Value = mds + val(TextBoxM4.Value)
If Cells(a, "C") Like "*" & " 80A" & "*" Then TextBoxM5.Value = mds + val(TextBoxM5.Value)
If Cells(a, "C") Like "*" & "x80A" & "*" Then TextBoxM5.Value = mds + val(TextBoxM5.Value)
If Cells(a, "C") Like "*" & " 90A" & "*" Then TextBoxM5.Value = mds + val(TextBoxM5.Value)
If Cells(a, "C") Like "*" & "x90A" & "*" Then TextBoxM5.Value = mds + val(TextBoxM5.Value)
If Cells(a, "C") Like "*" & " 100A" & "*" Then TextBoxM6.Value = mds + val(TextBoxM6.Value)
If Cells(a, "C") Like "*" & "x100A" & "*" Then TextBoxM6.Value = mds + val(TextBoxM6.Value)
If Cells(a, "C") Like "*" & " 125A" & "*" Then TextBoxM6.Value = mds + val(TextBoxM6.Value)
If Cells(a, "C") Like "*" & "x125A" & "*" Then TextBoxM6.Value = mds + val(TextBoxM6.Value)
mdt = mdt + mds
End If
mds = 0
a = a + 1
Loop
'r = &HCAE3BF 'yeşil,
Dim r
r = &HDBE8DB 'açık yeşil
For i = 1 To 7 'renk dolgusu
If Controls("TextBoxM" & i) <> "" Then Controls("TextBoxM" & i).BackColor = r Else _
Controls("TextBoxM" & i).BackColor = &HFFFFFF
Next i
TextBoxMDT = mdt
End Sub
Private Sub CommandButtonkd1_Click()
On Error GoTo hata
Workbooks(sk1).Activate: Application.Windows(sk1).Visible = True: End
hata:
End Sub