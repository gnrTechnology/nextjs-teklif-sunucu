Sub yeniteklif()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
de = ActiveWorkbook.Name
e = Sheets("Sayfa1").Range("B65536").End(xlUp).row + 1
Workbooks.Open "C:\Belgelerim\CEMEX\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx"
dt = ActiveWorkbook.Name
Application.DisplayAlerts = False
Workbooks(de).Sheets("Sayfa1").Range("A1:X" & e).Copy Workbooks(dt).Sheets("Sayfa1").Range("A1")
Application.DisplayAlerts = True
e = Workbooks(de).names.Count
For n = 1 To e
Ad1 = Workbooks(de).names.Item(n).Name
ad2 = Workbooks(dt).names(Ad1).Name
If Len(ad2) < 6 Then
Workbooks(dt).Sheets("Sayfa3").Range(Ad1) = Workbooks(de).Sheets("Sayfa3").Range(Ad1)
End If
Next n
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
Call pfkod
End Sub
Sub pfkod()
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
n = 2
e = Sheets("Sayfa1").Range("B65536").End(xlUp).row + 1
Do Until n > e
If Left(Range("A" & n), 3) = "PP-" Then
  If Left(Range("A" & n), 5) = "PP-PD" Then
    Set pf2 = Range("A" & n & ":C" & n).Find("FORM2", LookAt:=xlPart)
    If Not pf2 Is Nothing Then Range("A" & n).Replace What:="PP-PD", Replacement:="PP-F2", LookAt:=xlPart: GoTo git
    Set pf3 = Range("A" & n & ":C" & n).Find("FORM3", LookAt:=xlPart)
    If Not pf3 Is Nothing Then Range("A" & n).Replace What:="PP-PD", Replacement:="PP-F3", LookAt:=xlPart: GoTo git
    Set pf4 = Range("A" & n & ":C" & n).Find("FORM4", LookAt:=xlPart)
    If Not pf4 Is Nothing Then Range("A" & n).Replace What:="PP-PD", Replacement:="PP-F4", LookAt:=xlPart: GoTo git
    Range("A" & n).Replace What:="PP-PD", Replacement:="PP-DD", LookAt:=xlPart: GoTo git
  End If
  If Left(Range("A" & n), 5) = "PP-PS" Then
    Set sp1 = Range("A" & n & ":C" & n).Find("HARİCİ", LookAt:=xlPart)
    If Not sp1 Is Nothing Then Range("A" & n).Replace What:="PP-PS", Replacement:="PP-SH", LookAt:=xlPart: GoTo git
    Set SP2 = Range("A" & n & ":C" & n).Find("PASLANMAZ", LookAt:=xlPart)
    If Not SP2 Is Nothing Then Range("A" & n).Replace What:="PP-PS", Replacement:="PP-SX", LookAt:=xlPart: GoTo git
    Set sp3 = Range("A" & n & ":C" & n).Find("ALTI", LookAt:=xlPart)
    If Not sp3 Is Nothing Then Range("A" & n).Replace What:="PP-PS", Replacement:="PP-SA", LookAt:=xlPart: GoTo git
    Range("A" & n).Replace What:="PP-PS", Replacement:="PP-SD", LookAt:=xlPart: GoTo git
  End If
  If Left(Range("A" & n), 5) = "PP-PH" Then Range("A" & n).Replace What:="PP-PH", Replacement:="PP-DH", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-PX" Then Range("A" & n).Replace What:="PP-PX", Replacement:="PP-DX", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-PK" Then Range("A" & n).Replace What:="PP-PK", Replacement:="PP-KK", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-PT" Then Range("A" & n).Replace What:="PP-PT", Replacement:="PP-TK", LookAt:=xlPart: GoTo git

  If Left(Range("A" & n), 5) = "PP-MM" Then Range("A" & n).Replace What:="PP-MM", Replacement:="PM-MP", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-MS" Then Range("A" & n).Replace What:="PP-MS", Replacement:="PM-MS", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-MB" Then Range("A" & n).Replace What:="PP-MB", Replacement:="PM-MB", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-MH" Then Range("A" & n).Replace What:="PP-MH", Replacement:="PM-MA", LookAt:=xlPart: GoTo git
  If Left(Range("A" & n), 5) = "PP-MN" Then Range("A" & n).Replace What:="PP-MN", Replacement:="PM-MN", LookAt:=xlPart: GoTo git
End If
git:
n = n + 1
Loop
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub iscilik_gir2() '20 OCAK 2019  'adam saat üzerinden
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'--
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
'katsayılar varmı'''''
'If dkat = "" Or skat = "" Then Msg = MsgBox("İşçilik katsayıları girilmemiş.", vbInformation, "scngnr@gmail.com"): GoTo bitti
Dim hcr As Range, eskihcr As Range
Dim z, t
On Error GoTo hata
Set baslık = Columns("B:B").Find("BÖLÜM ADI/NO:", LookAt:=xlWhole)
If baslık Is Nothing Then GoTo son
z = baslık.row
Set toplam = baslık.Offset(-1, 0)
Do
If toplam Is Nothing Then GoTo son

Set baslık = Range(toplam, [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If baslık Is Nothing Then GoTo son
z = baslık.row
Set baslıkmı = Range(toplam.Offset(0, 0), [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
zm = baslıkmı.row
Set toplam = Range(baslık, [B65000]).Find("BÖLÜM TOPLAMI:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
t = toplam.row
'işçilik varmı'''''
Dim k
Set hcr = Range("A" & z & ":A" & t).Find("PM-MP-*", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
 If hcr Is Nothing Then
toplam.EntireRow.Insert
k = toplam.row - 1
 Else
k = hcr.row
 End If
'--
Range("A" & k).FormulaR1C1 = "PM-MP-auto"
Range("B" & k).FormulaR1C1 = "PM-MP"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "misia")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "misi")
Range("E" & k).FormulaR1C1 = "1"
Cells(k, "F") = "=Sum(Q" & baslık.row & ":Q" & toplam.row & ")"
Range("G" & k).FormulaR1C1 = "0"
If Left(ActiveSheet.CodeName, 3) = "OTM" Then GoTo zıpla1:
Range("L" & k).FormulaR1C1 = bfyt & "*Oisci/100" & kur 'Kar+1+1
Range("J" & k).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("N" & k).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("K" & k).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
Range("M" & k).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("O" & k).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & k).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & k).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & k).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & k).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("T" & k).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.
Range("U" & k).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
'TOPLAMLAR'--
Range("W" & k).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
Range("X" & k).FormulaR1C1 = "=RC[-19]*RC[-1]"
'Biçimlemeler'--
zıpla1:
Range("A" & k & ":U" & k & ",W" & k & ":X" & k).Select
    Selection.Borders.LineStyle = xlContinuous
    Selection.Interior.Pattern = xlNone
    Selection.RowHeight = 12.75
    Selection.Font.Size = 9
    Selection.Font.Bold = False
    Selection.Font.ColorIndex = xlAutomatic
    Selection.HorizontalAlignment = xlRight
'..
Range("A" & k & ":D" & k).HorizontalAlignment = xlLeft
Range("A" & k & ":D" & k).NumberFormat = "@"
Range("E" & k).NumberFormat = "#,##0"
Range("F" & k).NumberFormat = "#,##0.00"
Range("G" & k).NumberFormat = "0.0%"
Range("H" & k).NumberFormat = "#,##0"
Range("J" & k & ":X" & k).NumberFormat = "#,##0.00"
'..
Range("W" & k & ",X" & k).Select
If Range("Tpbr") = "Teklif Para Birimi (TL)" Then Selection.NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Selection.NumberFormat = "#,##0.00 [$€-1]"
If Range("Tpbr") = "Teklif Para Birimi (USD)" Then Selection.NumberFormat = "#,##0.00 [$$-C0C]"
'--
 Set baslık = toplam
hata:
Loop
Range("B1").Select
son:
Application.ScreenUpdating = True
If t = 0 Or z = 0 Then MsgBox "    BÖLÜM ADI/NO: ile BÖLÜM TOPLAMI: aralık girilmemiş. ", vbInformation, "scngnr@gmail.com": GoTo bitti
MsgBox "  İşçilikler girildi. ", vbOKOnly, "scngnr@gmail.com"
Call AraToplamlar
bitti:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub