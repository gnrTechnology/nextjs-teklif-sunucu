Dim i As Integer, n As Integer
Public bfyt As String
Dim kur As String

Private Sub Frame33_Click()

End Sub

Private Sub ListView1_DblClick()
G = ListView1.SelectedItem.Index
With ListView1.ListItems(G).ListSubItems(1)
 ug = .Text
 If ug = "Şalt Malzeme" Then .Text = "Pano": Exit Sub
 If ug = "Pano" Then .Text = "Bara": Exit Sub
 If ug = "Bara" Then .Text = "İşçilik": Exit Sub
 If ug = "İşçilik" Then .Text = "Sarf Malzeme": Exit Sub
 If ug = "Sarf Malzeme" Then .Text = "Ambalaj": Exit Sub
 If ug = "Ambalaj" Then .Text = "Nakliye": Exit Sub
 If ug = "Nakliye" Then .Text = "Şalt Malzeme": Exit Sub
End With
End Sub
Private Sub UserForm_Activate()
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
dizi1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
For i = 1 To 27
        CATH1.AddItem Mid(dizi1, i, 1)
        CATH2.AddItem Mid(dizi1, i, 1)
        CATH3.AddItem Mid(dizi1, i, 1)
        CATH4.AddItem Mid(dizi1, i, 1)
        CATH5.AddItem Mid(dizi1, i, 1)
        CATH6.AddItem Mid(dizi1, i, 1)
        CATH7.AddItem Mid(dizi1, i, 1)
        CATH8.AddItem Mid(dizi1, i, 1)
        CATH9.AddItem Mid(dizi1, i, 1)
Next
dizi2 = "-BOŞ-DOLU-B-BAŞLIK-BÖLÜM-BÖLÜM ADI-BÖLÜM ADI/NO:-PANO-PANO ADI"
ayır2 = Split(dizi2, "-")
For n = 1 To 10 '
    CATKB1.AddItem ayır2(n - 1)
    CATKB2.AddItem ayır2(n - 1)
Next n
dizi3 = "-BOŞ-DOLU-T-TOPLAM-ARA TOPLAM-BÖLÜM TOPLAMI:"
ayır3 = Split(dizi3, "-")
For n = 1 To 7 '
    CATKT1.AddItem ayır3(n - 1)
    CATKT2.AddItem ayır3(n - 1)
Next n
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
If Range("B" & "65536").End(xlUp).row < 2 And Range("A" & "65536").End(xlUp).row < 2 Then MsgBox (" Dosya boş."), vbInformation, "scngnr@hotmail.com": Exit Sub
T1.Value = ActiveWorkbook.Name
Workbooks(T1.Value).Activate
CATS1 = 2
CATS2 = Workbooks(T1.Value).ActiveSheet.Range("B65536").End(xlUp).row
Case 2
MultiPage1.Value = 1
Case 3
Unload Me
End Select
End Sub
Private Sub CommandButtonA1_Click()
If T1.Value = "" Then MsgBox (" Veri dosyasını seçiniz."), vbInformation, "scngnr@hotmail.com": Exit Sub
If CATH3.Text = "" Then MsgBox (" Üretici boş olamaz."), vbInformation, "scngnr@hotmail.com": Exit Sub
saltmalzeme1
End Sub
Sub saltmalzeme1()
'Application.ScreenUpdating = False
Workbooks(T1.Value).Activate
ListView1.ListItems.Clear
Dim son As Integer
If CATH3.Text = "" Then Exit Sub
son = ActiveSheet.Range(CATH3 & "65536").End(xlUp).row
For n = CATS1 To son
   'If WorksheetFunction.CountIf(Range(CATH3 & n & ":" & CATH3 & CATS2), Range(CATH3 & n).Value) = 1 Then
    If ActiveSheet.Range(CATH3 & n).Value <> "" Then
    ms = ListView1.ListItems.Count
    If ms = 0 Then GoTo atla1
    Set marka = ListView1.FindItem(Range(CATH3 & n).Value)
    If Not marka Is Nothing Then GoTo atla2
atla1:
    Call ListView1.ListItems.Add(ms + 1, , Range(CATH3 & n))
    Call ListView1.ListItems(ms + 1).ListSubItems.Add(1, , "Şalt Malzeme")
    End If
atla2:
Next n
'Application.ScreenUpdating = True
End Sub
Private Sub CommandButtonM1_Click()
If CATH3.Text = "" Then MsgBox (" Üretic boş olamaz."), vbInformation, "scngnr@hotmail.com": Exit Sub
Call teklifaktar1
End Sub
Sub teklifaktar1()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
ProgressBar1.Value = 1: ProgressBar1.Max = CATS2
If Not WorkbookOpen("Yeni Teklif V1.2.xltx") Then
    Workbooks.Open "C:\Belgelerim\CEMEX\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx"
End If
dt = ActiveWorkbook.Name
Set eski = Workbooks(T1.Value).ActiveSheet
Set yeni = Workbooks(dt).ActiveSheet
Workbooks(dt).Sheets("Sayfa1").Select
'--
Dim Ckar
Ckar = Workbooks(dt).names("CkarO").RefersToR1C1
If Ckar = Empty Then Workbooks(dt).names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
If CATH3.Text = "" Then Exit Sub
For i = CATS1 To CATS2
'GENEL VERİLER'--
yeni.Range("B" & i) = eski.Range(CATH1 & i) 'Sipariş Kd.
yeni.Range("C" & i) = eski.Range(CATH2 & i) 'Yapılacak İşin Cinsi
yeni.Range("D" & i) = eski.Range(CATH3 & i) 'Üretici
yeni.Range("E" & i) = eski.Range(CATH4 & i) 'Miktar
If Not yeni.Range("D" & i) = "" Then
yeni.Range("A" & i) = "XX"
If CATH1.Text = "" Then yeni.Range("B" & i) = "Sipariş Kd."
If CATH2.Text = "" Then yeni.Range("C" & i) = "Açıklama yazın"
'ÜRÜN GRUPLARI--
Set marka = ListView1.FindItem(Range("D" & i))
If Not marka Is Nothing Then ug = ListView1.ListItems(marka.Index).ListSubItems(1) Else GoTo atla2
If ug = "Şalt Malzeme" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Osalt/100": GoTo devam1 'Mlz.Kar+1
If ug = "İşçilik" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Oisci/100": yeni.Range("A" & i) = "PM-MP": GoTo devam1 'İŞÇ.
If ug = "Sarf Malzeme" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Osarf/100": yeni.Range("A" & i) = "PM-MS": GoTo devam1 'SARF
If ug = "Ambalaj" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Oamb/100": yeni.Range("A" & i) = "PM-MA": GoTo devam1 'AMB.
If ug = "Nakliye" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Onak/100": yeni.Range("A" & i) = "PM-MN": GoTo devam1 'NAK.
If ug = "Bara" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Obara/100": yeni.Range("A" & i) = "PM-MB": GoTo devam1 'Bara
If ug = "Pano" Then yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Opano/100": yeni.Range("A" & i) = "PP-": GoTo devam1 'Pano
yeni.Range("L" & i).FormulaR1C1 = bfyt & "*Osalt/100" 'Mlz.Kar+1
devam1:
 If CKF1.Value = True Then
 yeni.Range("F" & i) = "" 'fiyat
 yeni.Range("G" & i) = 0 'mlz.isk.
 yeni.Range("H" & i) = "" 'Adam/dk
 yeni.Range("I" & i) = "" 'Boyut
 yeni.Range("J" & i).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt
 yeni.Range("K" & i).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" 'Net Mlz. Alış+1
 yeni.Range("M" & i).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
 yeni.Range("N" & i).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
 yeni.Range("O" & i).FormulaR1C1 = "=RC[-10]*RC[-9]" 'Mlz. List Top.+1
 yeni.Range("P" & i).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
 yeni.Range("Q" & i).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
 yeni.Range("R" & i).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
 yeni.Range("S" & i).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
 yeni.Range("T" & i).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.
 yeni.Range("U" & i).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
'TOPLAMLAR'--
 yeni.Range("W" & i).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
 yeni.Range("X" & i).FormulaR1C1 = "=RC[-19]*RC[-1]"
'Genel Biçimlemeler'--
 'Range("A" & i & ":U" & i & ",W" & i & ":X" & i).Font.Size = 9
 yeni.Range("A" & i & ":D" & i).HorizontalAlignment = xlLeft
 yeni.Range("E" & i & ":U" & i & ",W" & i & ":X" & i).HorizontalAlignment = xlRight
 yeni.Range("A" & i & ":D" & i).NumberFormat = "@"
 yeni.Range("E" & i).NumberFormat = "#,##0"
 yeni.Range("F" & i).NumberFormat = "#,##0.00"
 yeni.Range("G" & i).NumberFormat = "0.0%"
 yeni.Range("H" & i).NumberFormat = "#,##0"
 yeni.Range("J" & i & ":X" & i).NumberFormat = "#,##0.00"
 yeni.Range("W" & i & ",X" & i).NumberFormat = "#,##0.00"
End If
yeni.Range("A" & i & ":U" & i).Borders.LineStyle = xlContinuous
yeni.Range("W" & i & ":X" & i).Borders.LineStyle = xlContinuous
End If
'--
atla2:
ProgressBar1.Value = i
Next i
ProgressBar1.Value = 1
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
Workbooks(T1.Value).Activate
End Sub
Private Sub CommandButtonM2_Click()
On Error Resume Next
If CATH5.Text = "" Then Exit Sub
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
ProgressBar1.Value = 1: ProgressBar1.Max = CATS2
Set eski = Workbooks(T1.Value).ActiveSheet
Set yeni = Workbooks(dt).ActiveSheet
Workbooks(dt).Sheets("Sayfa1").Select
For i = CATS1 To CATS2
If CATH5.Text <> "" Then
   If CATH6.Text <> "" And CATH7.Text <> "" Then bk1 = CATKB1: bk2 = CATKB2
   If bk1 = "DOLU" Then bk1 = "D"
   If bk1 = "BOŞ" Then bk1 = ""
   If bk2 = "DOLU" Then bk2 = "D"
   If bk2 = "BOŞ" Then bk2 = ""
   bk11 = eski.Range(CATH6 & i): bk21 = eski.Range(CATH7 & i)
   If bk11 = "" Then Else: If bk1 = "D" Then bk11 = "D"
   If bk21 = "" Then Else: If bk2 = "D" Then bk21 = "D"
'--
   If bk11 = bk1 And bk21 = bk2 Then
   yeni.Range("B" & i & ":E" & i & ",W" & i & ":X" & i).Borders(xlInsideVertical).LineStyle = xlNone
   yeni.Range("A" & i & ":X" & i).ClearContents
   yeni.Range("B" & i) = "BÖLÜM ADI/NO:"
   yeni.Range("C" & i) = eski.Range(CATH5 & i) 'Yapılacak İşin Cinsi
   Call ad_toplam_format
   End If
End If
If CATH8.Text <> "" And CATH9.Text <> "" Then bt1 = CATKT1: bt2 = CATKT2
   If bt1 = "DOLU" Then bt1 = "D"
   If bt1 = "BOŞ" Then bt1 = ""
   If bt2 = "DOLU" Then bt2 = "D"
   If bt2 = "BOŞ" Then bt2 = ""
   bt11 = eski.Range(CATH8 & i): bt21 = eski.Range(CATH9 & i)
   If bt11 = "" Then Else: If bt1 = "D" Then bt11 = "D"
   If bt21 = "" Then Else: If bt2 = "D" Then bt21 = "D"
   If bt11 = bt1 And bt21 = bt2 Then
   yeni.Range("A" & i & ":X" & i).ClearContents
   yeni.Range("B" & i) = "BÖLÜM TOPLAMI:"
   Call ad_toplam_format
End If
ProgressBar1.Value = i
Next i
ProgressBar1.Value = 1
i = ""
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub ad_toplam_format()
  'Biçimlemeler'
  Dim slc As Range
Set slc = Workbooks(dt).ActiveSheet.Range("A" & i & ",B" & i & ":E" & i & ",F" & i & ":U" & i & ",W" & i & ":X" & i)
    slc.Borders.LineStyle = xlContinuous
    Range("B" & i & ":E" & i & ",W" & i & ":X" & i).Borders(xlInsideVertical).LineStyle = xlNone
    'slc.Borders(xlInsideVertical).LineStyle = xlNone
    slc.Interior.Pattern = xlNone
    slc.RowHeight = 12.75
    slc.Font.Size = 9
    slc.Font.Bold = True
    slc.Font.ColorIndex = 11
End Sub
Private Sub CommandButtonM21_Click()
MultiPage1.Value = 0
End Sub
Private Sub CommandButtonM11_Click()
Unload Me
End Sub