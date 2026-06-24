Dim tsayfa
Dim Y As Integer, z As Integer
Private Sub CBGN_Click()
MultiPage1.Value = 0
End Sub
Private Sub CBGP_Click()
MultiPage1.Value = 1
End Sub
Private Sub CommandButton11_Click()
Unload Me
End Sub
Private Sub CommandButton12_Click()
LBLS1.Clear
End Sub
Private Sub Label64_Click()
ComboBoxL.DropDown
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
Workbooks(dt).Application.Range("A1").Interior.Color = 65535
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(fm1 & "\Malzeme Listeleri")
Set colSubfolders = objFolder.SubFolders
Tbk = Selection.row
 TextBox2.Value = Cells(Selection.row, "C")
 TextBox3.Value = Cells(Selection.row, "D")
 TextBox1.Value = Cells(Selection.row, "B")
For Each objSubfolder In colSubfolders
    ds = objSubfolder.Name
  Dim dosya
  Dim n As Integer
  dosya = dir(fm1 & "\Malzeme Listeleri\" & ds & "\*.xlsb")
  n = 1
  Do While dosya <> ""
    ComboBoxL.AddItem dosya
    ComboBoxL.List(ComboBoxL.ListCount - 1, 1) = objSubfolder.path & "\" & dosya
    dosya = dir
    n = n + 1
  Loop
Next

End Sub
Private Sub ComboBoxL_Click()
On Error Resume Next
CommandButton1.Visible = True
 If Selection.Cells = "" Or Selection.row = 1 Or Cells(Selection.row, "D") = "" Then _
 MsgBox "    Malzemenin olduğu bir satırı seçin. ", vbCritical, "scngnr@hotmail.com": Exit Sub
Application.ScreenUpdating = False
MLZ1 = ComboBoxL.Value
Tbk = Selection.row
If Not mlz = Empty Then
If mlz <> MLZ1 Then Windows(mlz).Close (False): Unload UF2
End If
TextBox2.Value = Cells(Selection.row, "C")
TextBox3.Value = Cells(Selection.row, "D")
TextBox1.Value = Cells(Selection.row, "B")
mlz = MLZ1
If Not WorkbookOpen((mlz)) Then
Workbooks.Open fileName:=ComboBoxL.List(ComboBoxL.listIndex, 1)
Application.Windows(mlz).Visible = False
'--
'TextBox1.SetFocus
Workbooks(dt).Activate
Workbooks(dt).Application.Range("A1").Interior.Color = 16751103
If Workbooks(dt).Application.Range("A1").Interior.Color = 16751103 Then
UF2.Show
End If
Cells(Tbk, "B").Select
TextBox2.Value = Cells(Selection.row, "C")
TextBox3.Value = Cells(Selection.row, "D")
TextBox1.Value = Cells(Selection.row, "B")
End If
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton1_Click()
On Error Resume Next
 If Selection.Cells = "" Or Selection.row = 1 Or Cells(Selection.row, "D") = "" Then _
 MsgBox "    Malzemenin olduğu bir satırı seçin. ", vbCritical, "scngnr@hotmail.com": Exit Sub
Application.ScreenUpdating = False
Tbk = Selection.row
TextBox2.Value = Cells(Selection.row, "C")
TextBox3.Value = Cells(Selection.row, "D")
TextBox1.Value = Cells(Selection.row, "B")
Workbooks(dt).Activate
mlz = ComboBoxL.Value
If Not WorkbookOpen((mlz)) Then
Workbooks.Open fileName:=ComboBoxL.List(ComboBoxL.listIndex, 1)
Application.Windows(mlz).Visible = False
End If
Workbooks(dt).Activate
UF2.Show
UF2.Label18.Caption = "  " & TextBox1.Value & " " & TextBox2.Value & " (" & TextBox3.Value & ")"
Cells(Tbk, 2).Select
If UFMZ.MultiPage1.Value = 1 Then UF2.MultiPage1.Value = 3 Else UF2.MultiPage1.Value = 1
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton10_Click()
If Cells(Selection.row, "D") = "" Or Selection.row = 1 Then Exit Sub
Tbk = Selection.row
Cells(Tbk, 2).Select
TextBox2.Value = Cells(Selection.row, "C")
TextBox3.Value = Cells(Selection.row, "D")
TextBox1.Value = Cells(Selection.row, "B")
End Sub
Private Sub Label63_Click()
If TextBox1.Text = "" Then Exit Sub
Dim Link
Dim tsayfa
Set tsayfa = Workbooks(dt).Worksheets("Sayfa1")
If TextBox3 = "ABB" Then Link = "https://new.abb.com/products/ABB" & TextBox1.Value: GoTo git:
If TextBox3 = "EATON" Then Link = "http://datasheet.moeller.net/datasheet.php?model=" & tsayfa.Range("L" & z) & "&locale=en_GB&_lt=": GoTo git:
If TextBox3 = "SIEMENS" Then Link = "https://mall.industry.siemens.com/mall/en/tr/Catalog/Product/" & TextBox1.Value: GoTo git:
If TextBox3 = "SCHNEIDER" Then Link = "https://www.schneider-electric.com.tr/tr/product/" & TextBox1.Value: GoTo git:
If TextBox3 = "PHOENIX" Then Link = "https://www.phoenixcontact.com/online/portal/tr?uri=pxc-oc-itemdetail:pid=" & TextBox1.Value: GoTo git:
git:
On Error GoTo NoCanDo
ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
'tam ekran olarak açmak Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & Link, vbMaximizedFocus
 TextBox1.SetFocus
 TextBox1.SelStart = 0
 TextBox1.SelLength = TextBox1.TextLength
 TextBox1.Copy
Exit Sub
NoCanDo:
MsgBox Link & " Bağlantı hatası", vbInformation
'MsgBox Link & " açılamıyor. İnternet bağlantınızı kontrol ediniz.", vbInformation
End Sub
Private Sub TextBox1_Change() 'RESİM DEĞİŞİMİ TAMAM+
On Error GoTo hata
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & Left(TextBox3.Value, 3) & "\" & TextBox1.Value & ".jpg")
If a = True Then
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(TextBox3.Value, 3) & "\" & TextBox1.Value & ".jpg")
Else
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(TextBox3.Value, 3) & "\Logo.jpg")
End If
Exit Sub
hata:
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'TAMAM
On Error Resume Next
Windows(mlz).Close (False)
mlz = Empty
Workbooks(dt).Application.Range("A1").Interior.Color = 65535
End
Unload UF2
End Sub
Private Sub CommandButton3_Click()
If LBLS1.ListCount < 1 Then Exit Sub
Set tsayfa = Workbooks(dt).Worksheets("Sayfa1")
k = TextBox1
z = 2: zx = 2
If CheckBoxLS1.Value = True Then
tekrar:
s = tsayfa.Range("B65536").End(xlUp).row
 Set t = tsayfa.Range(Cells(z, zx), Cells(s, zx))
 Set a = t.Find(k, LookIn:=xlValues, LookAt:=xlPart)
 If Not a Is Nothing Then
 Tbk = a.row
 Y = CDbl(Tbk)
 Call tCdegistir
 GoTo tekrar
 End If
Else
 Y = CDbl(Tbk)
Call tCdegistir
End If
zx = 2
End Sub
Sub tCdegistir() '08.12.2013 malzeme değişimi için
On Error Resume Next
Application.Calculation = xlCalculationManual
'GENEL VERİLER'--
Dim nameo
ad = tsayfa.Cells(Y, 5)
If ad = "" Then ad = 1
For ns = 0 To LBLS1.ListCount - 1
tsayfa.Cells(Y, 1) = LBLS1.List(ns, 0)
tsayfa.Cells(Y, 2) = LBLS1.List(ns, 1)
tsayfa.Cells(Y, 3) = LBLS1.List(ns, 2)
tsayfa.Cells(Y, 4) = LBLS1.List(ns, 3)
tsayfa.Cells(Y, 5) = ad
tsayfa.Cells(Y, 6) = LBLS1.List(ns, 5)
tsayfa.Cells(Y, 7) = LBLS1.List(ns, 6)
tsayfa.Cells(Y, 8) = LBLS1.List(ns, 7)
tsayfa.Cells(Y, 9) = LBLS1.List(ns, 8)
tsayfa.Cells(Y, 2).Font.ColorIndex = 7
'KUR'--
tsayfa.Cells(Y, 6).Font.ColorIndex = xlAutomatic: tsayfa.Cells(Y, 6).NumberFormat = "#,##0.00"
If LBLS1.List(ns, 9) = "USD" Then tsayfa.Cells(Y, 6).Font.ColorIndex = 3: tsayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$$-C0C]"
If LBLS1.List(ns, 9) = "EUR" Then tsayfa.Cells(Y, 6).Font.ColorIndex = 5: tsayfa.Cells(Y, 6).NumberFormat = "#,##0.00 [$€-1]"
If Left(ActiveSheet.CodeName, 3) = "OTM" Then GoTo zıpla1:
kur = ""
If LBLS1.List(ns, 9) = "USD" Then kur = "*Usd"
If LBLS1.List(ns, 9) = "EUR" Then kur = "*Eur"
'ÜRÜN GRUPLARI--
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
If Left(tsayfa.Cells(Y, 1), 5) = "PM-MP" Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Oisci/100" & mkur: GoTo devam1 'İŞÇ.
If Left(tsayfa.Cells(Y, 1), 5) = "PM-MS" Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Osarf/100" & mkur: GoTo devam1 'SARF
If Left(tsayfa.Cells(Y, 1), 5) = "PM-MA" Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Oamb/100" & mkur: GoTo devam1 'AMB.
If Left(tsayfa.Cells(Y, 1), 5) = "PM-MN" Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Onak/100" & mkur: GoTo devam1 'NAK.
If Left(tsayfa.Cells(Y, 1), 5) = "PM-MB" Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Obara/100" & mkur: GoTo devam1 'Bara
If Left(tsayfa.Cells(Y, 1), 3) = "PP-" Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Opano/100" & mkur: GoTo devam1 'Pano
If Left(tsayfa.Cells(Y, 1), 3) = "PS-" Then 'Pano sac & aksesuar
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Opano/100" & mkur
GoTo devam1
End If
tsayfa.Cells(Y, 12).FormulaR1C1 = bfyt & "*Osalt/100" & mkur 'Mlz.Kar+1
devam1:
Range("J" & Y).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt
Range("N" & Y).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("K" & Y).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış
Range("M" & Y).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1
Range("O" & Y).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.
Range("P" & Y).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.
Range("Q" & Y).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.
Range("R" & Y).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.
Range("S" & Y).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.
Range("T" & Y).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.* işçilik katsayılı
Range("U" & Y).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd
Range("W" & Y).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb" 'Dövize göre Birim Fiyat TL+1
Range("X" & Y).FormulaR1C1 = "=RC[-19]*RC[-1]"  'Toplam Fiyat TL
'--
zıpla1:
Y = Y + 1
If LBLS1.ListCount - ns > 1 Then tsayfa.Cells(Y, 1).EntireRow.Insert: z = Y
Next ns
Application.Calculation = xlCalculationAutomatic
End Sub