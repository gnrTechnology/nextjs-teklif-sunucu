Public bfyt As String
Dim kur, mkur As String
Dim k
Dim s
Dim tfy
Dim adr
Sub isciliksarfambsil() ' +1
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo hata
msg = MsgBox("Teklifte bulunan işçilik, sarf, ambalaj satırları (PM-M*-auto) silinecek. ", vbOKCancel, " Silme İşlemi")
If msg = vbCancel Then Exit Sub
Dim iscilik As Range
Set iscilik = Columns("A:A").Find("PM-M*-auto", LookAt:=xlWhole)
iscilik.EntireRow.Delete
Do
Set iscilik = Columns("A:A").Find("PM-M*-auto", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
iscilik.EntireRow.Delete
Loop
hata:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
Range("B2").Select
End Sub
Sub isciliksil() ' +1
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo hata
mnkod = "PM-MP*": mnkod1 = "PM-MP , PM-MP-auto"
If tp1 <> 1 Then If UFOPAN00.CKPKOD.Value = False Then mnkod1 = "PM-MP-auto": mnkod = "PM-MP-*" 'pano kodu tanıma
msg = MsgBox("Teklifte bulunan (" & mnkod1 & ") işçilik satırları silinecek.", vbOKCancel, "İşçilik Silme İşlemi")
If msg = vbCancel Then Exit Sub
Dim iscilik As Range
Set iscilik = Columns("A:A").Find(mnkod, LookAt:=xlWhole)
iscilik.EntireRow.Delete
Do
Set iscilik = Columns("A:A").Find(mnkod, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
iscilik.EntireRow.Delete
Loop
hata:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
Range("B2").Select
End Sub
Sub sarfsil() ' +1
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo hata
mskod = "PM-MS*": mskod1 = "PM-MS , PM-MS-auto"
If tp1 <> 1 Then If UFOPAN00.CKPKOD.Value = False Then mskod1 = "PM-MS-auto": mskod = "PM-MS-*" 'pano kodu tanıma
msg = MsgBox("Teklifte bulunan (" & mskod1 & ") sarf malzeme satırları silinecek.", vbOKCancel, "Sarf Malzeme Silme İşlemi")
If msg = vbCancel Then Exit Sub
Dim sarf As Range
Set sarf = Columns("A:A").Find(mskod, LookAt:=xlWhole)
sarf.EntireRow.Delete
Do
Set sarf = Columns("A:A").Find(mskod, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
sarf.EntireRow.Delete
Loop
hata:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
Range("B2").Select
End Sub
Sub ambsil() ' +1
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo hata
mhkod = "PM-MA*": mhkod1 = "PM-MA , PM-MA-auto"
If tp1 <> 1 Then If UFOPAN00.CKPKOD.Value = False Then mhkod1 = "PM-MA-auto":  mhkod = "PM-MA-*" 'pano kodu tanıma
msg = MsgBox("Teklifte bulunan (" & mhkod1 & ") Ambalaj satırları silinecek.", vbOKCancel, "Ambalaj Malzeme Silme İşlemi")
If msg = vbCancel Then Exit Sub
Dim sarf As Range
Set amb = Columns("A:A").Find(mhkod, LookAt:=xlWhole)
amb.EntireRow.Delete
Do
Set amb = Columns("A:A").Find(mhkod, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
amb.EntireRow.Delete
Loop
hata:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
Range("B2").Select
End Sub
Sub isciliksatır_gir1() '04 mayıs 2018 ' +1
On Error Resume Next
If Selection.row < 2 Then Exit Sub
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
On Error GoTo hata
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
'bara = GetSetting("ilhan", "Settings", "bara")
Dim msg
k = Selection.row
Cells(k, 2).Range("A1").Select
                If ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM ADI/NO:" Or ActiveCell.Offset(-1, 0).FormulaR1C1 = "BÖLÜM TOPLAMI:" Then
                msg = MsgBox("Burada bu işlemi yapamazsınız.", vbCritical, "scngnr@gmail.com")
                Exit Sub
                Else
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
Selection.EntireRow.Insert
'--
Range("A" & k).FormulaR1C1 = "PM-MP"
Call verilergir
End If
hata:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub sarfsatır_gir1() '04 mayıs 2018 ' +1
On Error Resume Next
If Selection.row < 2 Then Exit Sub
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
On Error GoTo hata
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
'bara = GetSetting("ilhan", "Settings", "bara")
Dim msg
k = Selection.row
Cells(k, 2).Range("A1").Select
                If ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM ADI/NO:" Or ActiveCell.Offset(-1, 0).FormulaR1C1 = "BÖLÜM TOPLAMI:" Then
                msg = MsgBox("Burada bu işlemi yapamazsınız.", vbCritical, "scngnr@gmail.com")
                Exit Sub
                Else
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
Selection.EntireRow.Insert
'--
Range("A" & k).FormulaR1C1 = "PM-MS"
Call verilergir
End If
hata:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub iscilik_gir1() '04 mayıs 2018 ' +1 tüm liste +1 2021
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = 0: UFOPAN00.ProgressBarP21.Visible = True ' ###########
'--
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
pnkod = "PP-*-auto*"
If tp1 = 0 Then
If UFOPAN00.CKPKOD.Value = True Then pnkod = "PP-*" Else pnkod = "PP-*-auto*" 'pano kodu tanıma
End If
'--
Dim Chtv, Cmtv
Chtv = ActiveWorkbook.names("Chtv").RefersToR1C1: Cmtv = ActiveWorkbook.names("Cmtv").RefersToR1C1
If Chtv = Empty Or Cmtv = Empty Then
 'txt dosyasından ad yöneticisine veri aktar.txt yoksa yok deyip çık.
End If
 Chtv = ActiveWorkbook.names("Chtv").Comment: Cmtv = ActiveWorkbook.names("Cmtv").Comment 'harfler & katlar
 dizi0 = Chtv: dizi1 = Cmtv
'--
'katsayılar varmı
If dizi1 = "" Then msg = MsgBox("İşçilik katsayılarını giriniz.", vbInformation, "scngnr@gmail.com"): GoTo atla1
''''
Dim hcr As Range, eskihcr As Range
Dim z, t
On Error GoTo hata
Set baslık = Columns("B:B").Find("BÖLÜM ADI/NO:", LookAt:=xlWhole)
If baslık Is Nothing Then GoTo son
z = baslık.row
Set toplam = baslık.Offset(-1, 0)
Do
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Max = Sheets("Sayfa1").Range("B65536").End(xlUp).row
If toplam Is Nothing Then GoTo son
Set baslık = Range(toplam, [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If baslık Is Nothing Then GoTo son
z = baslık.row
s = 0
If Cells(baslık.row, 5) > 1 And Cells(baslık.row, 5).NumberFormat = "#,##0 ""Adet""" Then s = Cells(baslık.row, 5): _
adr = Cells(baslık.row, 5).Address 'kat ekleme
Set baslıkmı = Range(toplam.Offset(0, 0), [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
zm = baslıkmı.row
Set toplam = Range(baslık, [B65000]).Find("BÖLÜM TOPLAMI:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
t = toplam.row
'pano işçilik varmı
Set hpr = Range("A" & z & ":A" & t).Find(pnkod, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False) 'pano varmı
Set hcr = Range("A" & z & ":A" & t).Find("PM-MP-*", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False) 'işçilik
If Not hcr Is Nothing Then Range("F" & hcr.row) = "0"

If Not hpr Is Nothing Then
tfy = 0
'pano tipleri
For i = z To t
If Range("A" & i) Like pnkod Then 'If Left(Range("A" & i), 3) = "PP-" Then
'harf ve çarpan bulma--
mharf = "": mdfc = 0
For sy = 1 To Len(Range("I" & i))
If IsNumeric(Mid(Range("I" & i), sy, 1)) = True Then Exit For Else mharf = mharf & Mid(Range("I" & i), sy, 1)
Next
 If mharf = "" Then mharf = "A": Range("I" & i) = "A" & Range("I" & i)
'--
arr1 = Split(dizi0, "-"): arr2 = Split(dizi1, "-")
aranan = Application.Match(mharf, arr1, False)
If Not IsError(aranan) Then kat = arr2(aranan - 1)
'sayı başlangıç--
deg = ""
For sy = 1 To Len(Range("I" & i))
If IsNumeric(Mid(Range("I" & i), sy, 1)) = True Then deg = deg & Mid(Range("I" & i), sy, 1)
Next
If deg = "" Or kat = "" Then
Msgsay = Msgsay + 1
If Msgsay < 11 Then mdfc = 0: msg = msg & Range("C" & z) & " /"
Else
ptalan = CDbl(deg): mdfc = (ptalan / 10) * kat * Range("E" & i)
End If
tfy = tfy + mdfc
'--
End If
Next i
tfy = Application.WorksheetFunction.RoundUp(tfy, 0)
If tfy = 0 Then GoTo atipi
'işçilik varmı
Set hcr = Range("A" & z & ":A" & t).Find("PM-MP-*", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)

If hcr Is Nothing Then
'''''
toplam.EntireRow.Insert
k = toplam.row - 1
Else
k = hcr.row
End If
'--
Range("A" & k).FormulaR1C1 = "PM-MP-auto"
Call verilergir ' ###########
'--
atipi:
Set baslık = toplam
End If 'pano varmı son

If hpr Is Nothing Then
Msg1say = Msg1say + 1
If Msg1say < 11 Then Msg1 = Msg1 & Range("C" & z) & " /"
End If
hata:
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = t ' ###########
Loop
Range("B1").Select
son:
If t = 0 Or z = 0 Then MsgBox "    BÖLÜM ADI/NO: ile BÖLÜM TOPLAMI: aralık girilmemiş. ", vbInformation, "scngnr@gmail.com": GoTo atla1
Call AraToplamlar
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Visible = False ' ###########
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
If msg <> "" Or Msg1 <> "" Then
If tp1 = 1 Then: GoTo atla1
'--
If Msgsay > 10 Then msg = msg & " + " & Msgsay
If Msg1say > 10 Then Msg1 = Msg1 & " + " & Msg1say
'MsgBox "Tanımlı Harf & Ebat Eksikler:" & vbCr & Msg & vbCr & "Pano Eksikler:" & vbCr & Msg1, vbInformation, "scngnr@gmail.com"
MsgBox "Harf & Ebat (m2) Eksiklerin Bulunduğu Panolar:" & vbCr & msg & vbCr & Msg1, vbInformation, "scngnr@gmail.com"
Else
If tp1 = 1 Then: GoTo atla1
MsgBox "  İşçilikler tüm panolara girildi. ", vbOKOnly, "scngnr@gmail.com"
End If
atla1:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
'If t = 0 Or z = 0 Or Msg <> "" And tp1 = 1 Then GoTo atla1 'End
End Sub
Sub sarf_gir1() '04 mayıs 2018 ' +1 tüm liste
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = 0: UFOPAN00.ProgressBarP21.Visible = True ' ###########
'--
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
pnkod = "PP-*-auto*"
If tp1 = 0 Then
If UFOPAN00.CKPKOD.Value = True Then pnkod = "PP-*" Else pnkod = "PP-*-auto*" 'pano kodu tanıma
End If
'--
Dim Chtv, Cstv
Chtv = ActiveWorkbook.names("Chtv").RefersToR1C1: Cstv = ActiveWorkbook.names("Cstv").RefersToR1C1
If Chtv = Empty Or Cstv = Empty Then
 'txt dosyasından ad yöneticisine veri aktar.txt yoksa yok deyip çık.
End If
 Chtv = ActiveWorkbook.names("Chtv").Comment: Cstv = ActiveWorkbook.names("Cstv").Comment 'harfler & katlar
 dizi0 = Chtv: dizi2 = Cstv
'katsayılar varmı
If dizi2 = "" Then msg = MsgBox("Montaj katsayılarını giriniz.", vbInformation, "scngnr@gmail.com"): GoTo atla1
'''''
Dim hcr As Range, eskihcr As Range
Dim z, t
On Error GoTo hata
Set baslık = Columns("B:B").Find("BÖLÜM ADI/NO:", LookAt:=xlWhole)
If baslık Is Nothing Then GoTo son
z = baslık.row
Set toplam = baslık.Offset(-1, 0)
Do
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Max = Sheets("Sayfa1").Range("B65536").End(xlUp).row
If toplam Is Nothing Then GoTo son
Set baslık = Range(toplam, [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If baslık Is Nothing Then GoTo son
z = baslık.row
Set baslıkmı = Range(toplam.Offset(0, 0), [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
zm = baslıkmı.row
s = 0
If Cells(baslık.row, 5) > 1 And Cells(baslık.row, 5).NumberFormat = "#,##0 ""Adet""" Then s = Cells(baslık.row, 5): _
adr = Cells(baslık.row, 5).Address 'kat ekleme
Set toplam = Range(baslık, [B65000]).Find("BÖLÜM TOPLAMI:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
t = toplam.row
'pano işçilik varmı
Set hpr = Range("A" & z & ":A" & t).Find(pnkod, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
Set hcr = Range("A" & z & ":A" & t).Find("PM-MS-*", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False) 'sarf
If Not hcr Is Nothing Then Range("F" & hcr.row) = "0"
If Not hpr Is Nothing Then
tfy = 0
'pano tipleri'--
For i = z To t
If Range("A" & i) Like pnkod Then 'If Left(Range("A" & i), 3) = "PP-" Then
'katsayılar--
mharf = "": mdfc = 0
For sy = 1 To Len(Range("I" & i))
If IsNumeric(Mid(Range("I" & i), sy, 1)) = True Then Exit For Else mharf = mharf & Mid(Range("I" & i), sy, 1)
Next
If mharf = "" Then mharf = "A": Range("I" & i) = "A" & Range("I" & i)
arr1 = Split(dizi0, "-"): arr2 = Split(dizi2, "-")
aranan = Application.Match(mharf, arr1, False)
If Not IsError(aranan) Then kat = arr2(aranan - 1)
'sayı başlangıç--
deg = ""
For sy = 1 To Len(Range("I" & i))
If IsNumeric(Mid(Range("I" & i), sy, 1)) = True Then deg = deg & Mid(Range("I" & i), sy, 1)
Next
If deg = "" Or kat = "" Then
Msgsay = Msgsay + 1
If Msgsay < 11 Then mdfc = 0: msg = msg & Range("C" & z) & " /"
Else
ptalan = CDbl(deg): mdfc = (ptalan / 10) * kat * Range("E" & i)
End If
tfy = tfy + mdfc
'--
End If
Next i
tfy = Application.WorksheetFunction.RoundUp(tfy, 0)
If tfy = 0 Then GoTo atipi
'sarf varmı
Set hcr = Range("A" & z & ":A" & t).Find("PM-MS-*", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If hcr Is Nothing Then
''''
toplam.EntireRow.Insert
k = toplam.row - 1
Else
k = hcr.row
End If
'--
Range("A" & k).FormulaR1C1 = "PM-MS-auto"
Call verilergir ' ###########
'--
atipi:
Set baslık = toplam
End If

If hpr Is Nothing Then
Msg1say = Msg1say + 1
If Msg1say < 11 Then Msg1 = Msg1 & Range("C" & z) & " /"
End If
hata:
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = t ' ###########
Loop
Range("B1").Select
son:
If t = 0 Or z = 0 Then MsgBox "    BÖLÜM ADI/NO: ile BÖLÜM TOPLAMI: aralık girilmemiş. ", vbInformation, "scngnr@gmail.com": GoTo atla1
Call AraToplamlar
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Visible = False ' ###########
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
If msg <> "" Or Msg1 <> "" Then
 If tp1 = 1 And tsi1 = 0 Then: GoTo atla1
'--
 If Msgsay > 10 Then msg = msg & " + " & Msgsay
 If Msg1say > 10 Then Msg1 = Msg1 & " + " & Msg1say
 If tp1 = 0 Then UFOPAN00.ProgressBarP21.Visible = False ' ###########
 'MsgBox "Tanımlı Harf & Ebat Eksikler:" & vbCr & Msg & vbCr & "Pano Eksikler:" & vbCr & Msg1, vbInformation, "scngnr@gmail.com"
 MsgBox "Harf & Ebat (m2) Eksiklerin Bulunduğu Panolar:" & vbCr & msg & vbCr & Msg1, vbInformation, "scngnr@gmail.com"
Else
 If tp1 = 1 Then
 If tsi1 = 1 Then MsgBox "  İşçilik + Sarf (" & pnkod & ") teklife girildi. ", vbOKOnly, "scngnr@gmail.com"
 GoTo atla1
 End If
 MsgBox "  Sarf malzeme tüm panolara girildi. ", vbOKOnly, "scngnr@gmail.com"
End If
atla1:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
'If t = 0 Or z = 0 Or Msg <> "" And tp1 = 1 Then GoTo atla1 'End
End Sub
Sub amb_gir1() '04 mayıs 2018 ' +1
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = 0: UFOPAN00.ProgressBarP21.Visible = True ' ###########
'--
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'carpan alma'''''
Dim Catv
Catv = ActiveWorkbook.names("Catv").RefersToR1C1
If Catv = Empty Then
  dkat = GetSetting("ilhan", "Settings", "amb"): skat = GetSetting("ilhan", "Settings", "amb")
  'dkat = CDbl(UFOPAN00.TextBoxamb.Value): skat = CDbl(UFOPAN00.TextBoxamb.Value)
Else
  Catv = ActiveWorkbook.names("Catv").Comment
  dizi1 = Catv
  dkat = Split(dizi1, "-")(0): skat = dkat
End If
If tp1 <> 1 Then
  tx = InputBox("Pano yüzey alanı (Yük.x Gen.)m2 x ? ", "Ambalaj Çarpan Değerini Giriniz.", dkat)
  If tx = "" Then GoTo atla1
  dkat = tx
  ActiveWorkbook.names.Add Name:="Catv", RefersToR1C1:="Amb.Çarpanlar": ActiveWorkbook.names("Catv").Comment = tx
Else
  If dkat = "" Then dkat = 6
End If
'--
pnkod = "PP-*-auto*"
If tp1 = 0 Then
If UFOPAN00.CKPKOD.Value = True Then pnkod = "PP-*" Else pnkod = "PP-*-auto*" 'pano kodu tanıma
End If
'--
'pano aralık bulma'''''
Dim hcr As Range, eskihcr As Range
Dim z, t
On Error GoTo hata
Set baslık = Columns("B:B").Find("BÖLÜM ADI/NO:", LookAt:=xlWhole)
If baslık Is Nothing Then GoTo son1
z = baslık.row
Set toplam = baslık.Offset(-1, 0)
Do 'aralık döngüsü başlatma
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Max = Sheets("Sayfa1").Range("B65536").End(xlUp).row
If toplam Is Nothing Then GoTo son1
Set baslık = Range(toplam, [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If baslık Is Nothing Then GoTo son1
z = baslık.row
Set baslıkmı = Range(toplam.Offset(0, 0), [B65000]).Find("BÖLÜM ADI/NO:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
zm = baslıkmı.row
s = 0
If Cells(baslık.row, 5) > 1 And Cells(baslık.row, 5).NumberFormat = "#,##0 ""Adet""" Then s = Cells(baslık.row, 5): _
adr = Cells(baslık.row, 5).Address 'kat ekleme
Set toplam = Range(baslık, [B65000]).Find("BÖLÜM TOPLAMI:", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
t = toplam.row
'pano varmı'''''
Set hpr = Range("A" & z & ":A" & t).Find(pnkod, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If Not hpr Is Nothing Then
tfy = 0
'pano tiplerine göre ambalaj fiyatları'''''
For i = z To t
If Range("A" & i) Like pnkod Then 'If Left(Range("A" & i), 3) = "PP-" Then
'yenikatsayılar--
mdfc = 0
ptip = Mid(Range("A" & i), 4, 1)
kat = dkat
'burada istersen pano tipi belirle
deg = ""
For sy = 1 To Len(Range("I" & i)) 'sayı başlangıç
If IsNumeric(Mid(Range("I" & i), sy, 1)) = True Then deg = deg & Mid(Range("I" & i), sy, 1)
Next
If deg = "" Then
mdfc = 0: msg = msg & vbCr & Range("C" & z) & " (Harf + Ebat) girilmemiş."
Else
ptalan = CDbl(deg): mdfc = (ptalan / 10) * CDbl(kat) * Range("E" & i)
End If
tfy = tfy + mdfc
End If
Next i
'--
tfy = Application.WorksheetFunction.RoundUp(tfy, -1)
If tfy = 0 Then GoTo atipi
'ambalaj varmı'
Set hcr = Range("A" & z & ":A" & t).Find("PM-MA-*", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
If hcr Is Nothing Then
toplam.EntireRow.Insert
k = toplam.row - 1
Else
k = hcr.row
End If
Range("A" & k).FormulaR1C1 = "PM-MA-auto"
Call verilergir ' ###########
'--
atipi:
Set baslık = toplam
End If
If hpr Is Nothing Then Msg1 = Msg1 & vbCr & Range("C" & z) & " (" & pnkod & ") pano girilmemiş."
hata:
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = t ' ###########
Loop
Range("B1").Select
son1:
If t = 0 Or z = 0 Then
MsgBox "    BÖLÜM ADI/NO: ile BÖLÜM TOPLAMI: aralık girilmemiş. ", vbInformation, "scngnr@gmail.com": GoTo atla1
End If
Call AraToplamlar
If tp1 = 0 Then UFOPAN00.ProgressBarP21.Visible = False ' ###########
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
If msg <> "" Or Msg1 <> "" Then
  If tp1 = 0 Then UFOPAN00.ProgressBarP21.Visible = False
  MsgBox msg & Msg1, vbInformation, "scngnr@gmail.com"
Else
  If tp1 = 1 Then: Application.ScreenUpdating = True: _
  MsgBox "  İşçilik + Sarf + Ambalaj (" & pnkod & ") teklife girildi. ", vbOKOnly, "scngnr@gmail.com": GoTo atla1
  MsgBox "  Ambalaj bedelleri tüm panolara girildi. ", vbOKOnly, "scngnr@gmail.com"
End If
atla1:
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
'If t = 0 Or z = 0 Or Msg <> "" And tp1 = 1 Then End
End Sub
Sub bakir_gir() 'yeni 14 nisan 2018 +1
On Error Resume Next
If Selection.row < 2 Then Exit Sub
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
On Error GoTo hata
If Not ActiveSheet.Name = "Sayfa1" Then Exit Sub
k = Selection.row
Cells(k, 2).Range("A1").Select
                'If ActiveCell.Offset(0, 0).FormulaR1C1 = "BÖLÜM ADI/NO:" Or ActiveCell.Offset(-1, 0).FormulaR1C1 = "BÖLÜM TOPLAMI:" Then
                'Msg = MsgBox("Burada bu işlemi yapamazsınız.", vbCritical, "scngnr@gmail.com")
                'Exit Sub
'Application.ScreenUpdating = False:
Selection.EntireRow.Insert
Range("A" & k).FormulaR1C1 = "PM-MB"
Call verilergir ' ###########
hata:
End Sub
Sub verilergir()
On Error Resume Next
Application.Calculation = xlCalculationManual
'--
'If tp1 = 0 Then UFOPAN00.ProgressBarP21.Value = k
'Ortak
kur = ""
Range("E" & k) = "1"
Range("F" & k) = "0": Range("G" & k) = "0"
'Seçimler
If Range("A" & k) = "PM-MP-auto" Then 'işçilik-auto
Range("B" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "skdi")
If Range("B" & k).FormulaR1C1 = "" Then Range("B" & k).FormulaR1C1 = "PM-MP"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "misia")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "misi")
Range("L" & k).FormulaR1C1 = bfyt & "*Oisci/100" & kur 'işç.kar
Range("F" & k).FormulaR1C1 = Application.WorksheetFunction.RoundUp(tfy, -1)
If s > 0 Then Range("E" & k).Formula = "=" & "1" & "*" & adr
If s > 0 Then Range("F" & k).FormulaR1C1 = tfy / s
GoTo git:
End If
If Range("A" & k) = "PM-MS-auto" Then 'sarf-auto
Range("B" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "skds")
If Range("B" & k).FormulaR1C1 = "" Then Range("B" & k).FormulaR1C1 = "PM-MS"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "msasa")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "msas")
Range("L" & k).FormulaR1C1 = bfyt & "*Osarf/100" & kur 'sarf kar
Range("F" & k).FormulaR1C1 = Application.WorksheetFunction.RoundUp(tfy, -1)
If s > 0 Then Range("E" & k).Formula = "=" & "1" & "*" & adr
If s > 0 Then Range("F" & k).FormulaR1C1 = tfy / s
GoTo git:
End If
If Range("A" & k) = "PM-MA-auto" Then 'amb-auto
Range("B" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "skda")
If Range("B" & k).FormulaR1C1 = "" Then Range("B" & k).FormulaR1C1 = "PM-MA"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "mamaa")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "mama")
Range("L" & k).FormulaR1C1 = bfyt & "*Oamb/100" & kur 'amb.kar
Range("F" & k).FormulaR1C1 = Application.WorksheetFunction.RoundUp(tfy, -1)
GoTo git:
End If
If Range("A" & k) = "PM-MP" Then 'işçilik-mono
Range("B" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "skdi")
If Range("B" & k).FormulaR1C1 = "" Then Range("B" & k).FormulaR1C1 = "PM-MP"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "misia")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "misi")
Range("L" & k).FormulaR1C1 = bfyt & "*Oisci/100" & kur 'işç.kar
GoTo git:
End If
If Range("A" & k) = "PM-MS" Then 'sarf-mono
Range("B" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "skds")
If Range("B" & k).FormulaR1C1 = "" Then Range("B" & k).FormulaR1C1 = "PM-MS"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "msasa")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "msas")
Range("L" & k).FormulaR1C1 = bfyt & "*Osarf/100" & kur 'sarf kar
GoTo git:
End If
If Range("A" & k) = "PM-MB" Then 'bara-mono
Range("B" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "skdb")
If Range("B" & k).FormulaR1C1 = "" Then Range("B" & k).FormulaR1C1 = "PM-MB"
Range("C" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "mbaba")
Range("D" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "mbab")
Range("E" & k).FormulaR1C1 = "1"
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then Call malzemedosya1: GoTo son
Range("F" & k).FormulaR1C1 = GetSetting("ilhan", "Settings", "bara")
Range("L" & k).FormulaR1C1 = bfyt & "*Obara/100" & kur 'bakır kar
GoTo git:
End If
'Ortak
git:
If Left(ActiveSheet.CodeName, 3) = "OTM" Then GoTo zıpla1
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
'Toplamlar
Range("W" & k).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
Range("X" & k).FormulaR1C1 = "=RC[-19]*RC[-1]"
'Biçimlemeler
zıpla1:
Range("A" & k & ":U" & k).RowHeight = 12.75
Range("A" & k & ":U" & k).Borders.LineStyle = xlContinuous
Range("W" & k & ":X" & k).Borders.LineStyle = xlContinuous
Range("A" & k & ":X" & k).Interior.Pattern = xlNone
Range("A" & k & ":X" & k).Font.Bold = False
Range("A" & k & ":X" & k).Font.ColorIndex = xlAutomatic
Range("A" & k & ":X" & k).Font.Size = 9
Range("A" & k & ":D" & k).HorizontalAlignment = xlLeft
Range("E" & k & ":X" & k).HorizontalAlignment = xlRight
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
son:
Application.Calculation = xlCalculationAutomatic
End Sub
Sub malzemedosya1()
Range("F" & k).FormulaR1C1 = "": Range("G" & k).FormulaR1C1 = "": Range("J" & k).FormulaR1C1 = ""
Range("H" & k).FormulaR1C1 = "=IF(RC[-2]-RC[-1]<=0,""-"",RC[-2]-RC[-1])"
Range("I" & k).FormulaR1C1 = "=IF(RC[-2]-RC[-3]<=0,""-"",RC[-2]-RC[-3])"
Range("A" & k & ":J" & k).Borders.LineStyle = xlContinuous
Range("A" & k & ":J" & k).Font.Bold = False
Range("A" & k & ":J" & k).Font.ColorIndex = xlAutomatic
Range("A" & k & ":J" & k).Font.Size = 9
Range("A" & k & ":X" & k).Font.Name = "Arial"
Range("A" & k & ":D" & k).HorizontalAlignment = xlLeft
Range("E" & k & ":J" & k).HorizontalAlignment = xlRight
Range("A" & k & ":D" & k).NumberFormat = "@"
Range("E" & k).NumberFormat = "#,##0"
Range("F" & k & ":J" & k).NumberFormat = "#,##0.00"
Range("H" & k).NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
Range("I" & k).NumberFormat = "[Red]#,##0.00;[Blue]-#,##0.00;[Blue] #,##0.00"
End Sub