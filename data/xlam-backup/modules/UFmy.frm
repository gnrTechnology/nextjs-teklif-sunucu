Public bfyt As String
Option Compare Text
Private Sub CommandButton10_Click()
Call teklifisk
End Sub
Private Sub CheckBox2_Click()
ListBoxP26.Clear
If CheckBox2.Value = True Then CheckBox1.Enabled = False: ListBoxP27.Clear: ListBoxP28.Clear Else CheckBox1.Enabled = True
End Sub
Private Sub CommandButton11_Click()
Call teklifisk
End Sub
Private Sub CommandButtonP23_Click()
On Error Resume Next
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
'--
baslax = Timer
ProgressBarP21.Visible = True
If Frame23.Visible = False Then GoTo basla1
If Frame23.Visible = True Then GoTo basla2
'--
basla1:
If ListViewP21.ListItems.Count = 0 Then
msg = MsgBox("    Lütfen önce aktarılacak verileri seçiniz. ", vb, "scngnr@hotmail.com")
Else:
'--
son = Sheets("Sayfa1").Range("B65536").End(xlUp).row + 1
listson = ListViewP21.ListItems.Count
'--
'If CheckBoxP21.Value = False And CheckBoxP22.Value = False Then
'If CheckBoxP20.Value = False Then Lmarka = ListViewP21.SelectedItem.ListSubItems(5).Text: GoTo git1
'End If
ProgressBarP21.Max = son
Y = 2
Do Until Y >= son
       ProgressBarP21.Value = Y
lst = 1
       Do Until lst > ListViewP21.ListItems.Count - 1
       sayfakod = Left(Sheets("Sayfa1").Cells(Y, 1), 10)
       sayfakod = Replace(sayfakod, "-auto", "")
       ayır = Split(sayfakod, "."): sayfakod = ayır(0) 'deneme ..
       If sayfakod = "" Then GoTo git11
       smarka = Sheets("Sayfa1").Cells(Y, 4)
       Lkod = ListViewP21.ListItems(lst).ListSubItems(2).Text
       Lmarka = ListViewP21.ListItems(lst).ListSubItems(5).Text
       isk = ListViewP21.ListItems(lst).ListSubItems(1).Text
       
       If sayfakod = Lkod And smarka = Lmarka Then
       ListViewP21.ListItems(lst).Bold = False
       ListViewP21.ListItems(lst).ListSubItems(1).Bold = False
       Sheets("Sayfa1").Cells(Y, 7) = isk / 100: GoTo git11
       End If
       
       lst = lst + 1
       Loop
git11:
       Y = Y + 1
Loop
ProgressBarP21.Value = son
ProgressBarP21.Visible = False
bitirx = Timer
süre = Format(bitirx - baslax, "Fixed") & " sn."
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
msg = MsgBox(süre & "de  Seçilen markaya ait iskontolar teklife aktarıldı. ", vb, "scngnr@hotmail.com")
End If: GoTo bitir
'basla2:
basla2:
If ListViewP21.ListItems.Count = 0 Then
msg = MsgBox("    Lütfen önce aktarılacak verileri seçiniz. ", vb, "scngnr@hotmail.com")
Else:
mlzs = TextBoxP21.Value
Dim msayfa
Set msayfa = Workbooks(mlzs).Worksheets("Sayfa1")
listson = ListViewP21.ListItems.Count
son = msayfa.Range("D65536").End(xlUp).row + 1
ProgressBarP21.Max = son

Y = 2
Do Until Y >= son
       ProgressBarP21.Value = Y
       If msayfa.Cells(Y, 4) = "" Then GoTo git12
       sayfakod = Left(msayfa.Cells(Y, 11), 10)
       If sayfakod = "" Then
       msayfa.Cells(Y, 7) = isk: GoTo git12
       End If
i = 1
       Do Until i > ListViewP21.ListItems.Count
                ListKod = ListViewP21.ListItems(i).ListSubItems(2).Text
                If sayfakod = ListKod Then
                msayfa.Cells(Y, 7) = ListViewP21.ListItems(i).ListSubItems(1).Text / 100
                isk = msayfa.Cells(Y, 7)
                Exit Do
                End If
                i = i + 1
        Loop
git12:
'--
       Y = Y + 1
Loop
ProgressBarP21.Value = son
ProgressBarP21.Visible = False
msg = MsgBox("    İskontolar malzeme listesine aktarıldı. ", vb, "scngnr@hotmail.com")
End If
Workbooks(mlzs).SaveChanges = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
'ListViewP21.ListItems.Clear
'--
bitir:
End Sub
Private Sub Image3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt"
End Sub
Private Sub Image4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Transfer"
End Sub
Private Sub Label68_Click()
If Me.ListBoxMK.ListCount <> 0 Then
ListBoxMK.List = Diz(ListBoxMK.List, 1)
Else: MsgBox "ListBox boş", vbCritical
End If
End Sub
Private Sub Label67_Click()
If Me.ListBoxMK.ListCount <> 0 Then
ListBoxMK.List = Diz(ListBoxMK.List, 2)
Else: MsgBox "ListBox boş", vbCritical
End If
End Sub
Private Function Diz(ByVal Dizim As Variant, Stn As Integer) As Variant
    Dim i, j, k As Long
    Dim Tmp As Variant
    Stn = Stn - 1
    For i = LBound(Dizim, 1) To UBound(Dizim, 1)
        For j = i + 1 To UBound(Dizim, 1)
            If Dizim(i, Stn) > Dizim(j, Stn) Then
                For k = LBound(Dizim, 2) To UBound(Dizim, 2)
                    Tmp = Dizim(j, k)
                    Dizim(j, k) = Dizim(i, k)
                    Dizim(i, k) = Tmp
                Next
            End If
        Next
    Next
    Diz = Dizim
End Function
Private Sub UserForm_Initialize() 'form yükleme xxx
On Error Resume Next
'--
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
ToolbarP2.ImageList = UFmy.ImageList2
ToolbarP2.Buttons.Item(1).Image = ImageList2.ListImages.Item(1).Index
ToolbarP2.Buttons.Item(2).Image = ImageList2.ListImages.Item(2).Index
ToolbarP2.Buttons.Item(3).Image = ImageList2.ListImages.Item(3).Index
ToolbarP2.Buttons.Item(4).Image = ImageList2.ListImages.Item(4).Index
ToolbarP2.Buttons.Item(5).Image = ImageList2.ListImages.Item(5).Index
ToolbarP2.Buttons.Item(6).Image = ImageList2.ListImages.Item(6).Index
ToolbarP2.Buttons.Item(7).Image = ImageList2.ListImages.Item(7).Index
ListViewP21.ColumnHeaders.Clear
ListViewP21.SmallIcons = ImageList4
ListViewP21.View = lvwReport
ListViewP21.OLEDragMode = ccOLEDragAutomatic
ListViewP21.FullRowSelect = True
Call ListViewP21.ColumnHeaders.Add(1, , "Yapılacak İşin Cinsi", 310)
Call ListViewP21.ColumnHeaders.Add(2, , "İsk.%", 42, 1)
Call ListViewP21.ColumnHeaders.Add(3, , "Kod.", 0)
Call ListViewP21.ColumnHeaders.Add(4, , "Mlz.List Top.", 76, 1)
Call ListViewP21.ColumnHeaders.Add(5, , "Mlz.Net Top.", 76, 1)
Call ListViewP21.ColumnHeaders.Add(6, , "Marka", 0)
Call teklifisk
ListBoxP21.SetFocus
Call malzkod
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'TAMAM
On Error Resume Next
Application.DisplayAlerts = False
Application.Windows(TextBoxP21.Text).WindowState = xlMinimized
Windows(TextBoxP21.Text).Close SaveChanges:=True
Application.Windows(TBMLZ1.Text).WindowState = xlMinimized
Windows(TBMLZ1.Text).Close SaveChanges:=True
Application.DisplayAlerts = True
End Sub
Private Sub MultiPageP2_Change() 'xxx
ListViewP21.Visible = False
ListViewP21.Visible = True
End Sub
Private Sub ToolbarP2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
'--
Application.DisplayAlerts = False
Windows(TextBoxP21.Text).Close SaveChanges:=True
Application.DisplayAlerts = True
'--
If Frame18.Visible = False Then ListViewP21.ListItems.Clear
MultiPageisk1.Value = 0: Frame18.Visible = True: Frame23.Visible = False
'CommandButtonP21.Enabled = True
CheckBoxP21.Value = False
ListViewP21.ColumnHeaders.Item(4).Width = 76
ListViewP21.ColumnHeaders.Item(5).Width = 76
'CommandButtonP21.Enabled = False
MultiPageP2.Value = 0
UFmy.Height = 400
ListBoxP21.SetFocus
'ListBoxP21.Selected(ListBoxP21.ListIndex) = False
Case 2
'Call degistir
'If CheckBox1 = False Then Call tbmmalzemeler1
UFmy.Height = 400
'--
Application.DisplayAlerts = False
Windows(TextBoxP21.Text).Close SaveChanges:=True
Application.DisplayAlerts = True
'--
MultiPageP2.Value = 1
Case 3
UFmy.Height = 400
MultiPageP2.Value = 3
If CBMLZ1.ListCount > 1 Then Exit Sub
CBMLZ1.Clear
Call mlzdosyalar
Case 4
MultiPageisk1.Value = 1: Frame23.Visible = True: Frame18.Visible = False
'CommandButtonP21.Enabled = False
CheckBoxP21.Enabled = False
ComboBoxP22.Clear
'ListBoxP21.Clear
'--
ListViewP21.ColumnHeaders.Item(4).Width = 0
ListViewP21.ColumnHeaders.Item(5).Width = 0
ListViewP21.ListItems.Clear
For i = 0 To ListBoxP21.ListCount - 1
    ListBoxP21.Selected(i) = False
Next i
For i = 0 To ListBoxP22.ListCount - 1
    ListBoxP22.Selected(i) = False
Next i
'--
Call malzisk
UFmy.Height = 400
MultiPageP2.Value = 0
Case 5
UFmy.Height = 400
Call malzkod
MultiPageP2.Value = 2
Case 6
If UFmy.Height >= 350 Then
UFmy.Height = UFmy.Height - UFmy.InsideHeight + ToolbarP2.Height
Else
UFmy.Height = 400
End If
Case 7
Unload Me
End Select
End Sub
Private Sub ToolbarP2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Tag
Case 1
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Transfer\" & TBMLZTS
Case 2
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt"
End Select
End Sub
Sub teklifisk() '2019
On Error Resume Next
ListViewP21.ListItems.Clear
ComboBoxP22.Clear
ListBoxP21.Clear: ListBoxP22.Clear: ListBoxP26.Clear: CBTKL1.Clear
Windows(dt).Activate
'Sheets("Sayfa1").Select '
Dim son As Integer
son = Sheets("Sayfa1").Range("D65536").End(xlUp).row
For n = 2 To son
mss = WorksheetFunction.CountIf(Sheets("Sayfa1").Range("D2:D" & n), Sheets("Sayfa1").Cells(n, 4).Value) 'Malzeme markası sayısı
pkd = Left(Sheets("Sayfa1").Cells(n, 1), 3) 'pano ve montaj ile ilgili referans
sbt = Sheets("Sayfa1").Cells(n, 2) 'satırbasları
mmc = Sheets("Sayfa1").Cells(n, 3).Value 'yapılacak işin cinsi
If Sheets("Sayfa1").Cells(n, 1) = "" And Sheets("Sayfa1").Cells(n, 4) = "" Then GoTo atla
If sbt = "BÖLÜM TOPLAMI:" Or sbt = "BÖLÜM ADI/NO:" Or sbt = "GENEL TOPLAM:" Then GoTo atla
If Sheets("Sayfa1").Cells(n, 1) = "" Then Sheets("Sayfa1").Cells(n, 1) = "XX"
   If mss = 1 Then
       ListBoxP26.AddItem Sheets("Sayfa1").Cells(n, 4) 'ListBox için yazdım
       CBTKL1.AddItem Cells(n, 4)
       If pkd <> "PP-" And pkd <> "PM-" Then ListBoxP21.AddItem Sheets("Sayfa1").Cells(n, 4) _
       Else ListBoxP22.AddItem Sheets("Sayfa1").Cells(n, 4) 'ListBox için yazdım
   End If
atla:
Next n
ListBoxP21.IntegralHeight = False: ListBoxP21.Height = 126: ListBoxP21.IntegralHeight = True
CommandButtonP21.Enabled = True
End Sub
Sub malzisk()
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(fm1 & "\Malzeme Listeleri")
Set colSubfolders = objFolder.SubFolders
For Each objSubfolder In colSubfolders
    ds = objSubfolder.Name
  Dim dsm
  Dim n As Integer
  dsm = dir(fm1 & "\Malzeme Listeleri\" & ds & "\*.xlsb")
  n = 1
  Do While dsm <> ""
    ComboBoxP22.AddItem dsm
    dsm = dir
    n = n + 1
  Loop
Next
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
Private Sub ComboBoxP22_Change()
On Error Resume Next
If ComboBoxP22.Text <> TextBoxP21.Text Then
Application.DisplayAlerts = False
Windows(TextBoxP21.Text).Close SaveChanges:=True
Application.DisplayAlerts = True
ListViewP21.ListItems.Clear
End If
End Sub
Private Sub CommandButtonP22_Click() '2019
On Error Resume Next
If ComboBoxP22.listIndex = -1 Then Exit Sub
If ComboBoxP22.Text = "" Then Exit Sub
If ListViewP21.ListItems.Count > 0 Then Exit Sub
ListViewP21.SmallIcons = UFmy.ImageList4
Dim ds, f, f1, fc, mlzs
'--
If WorkbookOpen(ComboBoxP22.Text) Then GoTo atla1:
'--
Application.ScreenUpdating = False
Set ds = CreateObject("Scripting.FileSystemObject")
Set f = ds.GetFolder(fm1 & "\Malzeme Listeleri")
Set fc = f.SubFolders
Dim mlz As listItem
'--
For Each f1 In fc
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists(f1 & "\" & ComboBoxP22.Text)
If a = True Then Workbooks.Open fileName:=f1 & "\" & ComboBoxP22.Text: Exit For
Next
'--
atla1:
mlzs = ComboBoxP22.Text
Application.Windows(mlzs).Visible = False
Dim msayfa
Set msayfa = Workbooks(mlzs).Worksheets("Sayfa1")
TextBoxP21.Text = mlzs

Application.Windows(mlzs).WindowState = xlMinimized
'Application.ScreenUpdating = True
'--
ListViewP21.ListItems.Clear
son = Workbooks(mlzs).Worksheets("Sayfa1").Range("B65536").End(xlUp).row
X = 0
For n = 2 To son
    If msayfa.Cells(n, 11) = "" Then
    Do Until msayfa.Cells(n, 11) <> ""
    n = n + 1
    If n > son Then GoTo bitti:
    Loop
End If

    sayfakod = Left(msayfa.Cells(n, 11), 10)
    listson = ListViewP21.ListItems.Count
    i = 0
    Do Until i = listson
    ListKod = ListViewP21.ListItems(i + 1).ListSubItems(2).Text
    
    If sayfakod = ListKod Then GoTo devam:
    i = i + 1
    Loop
Call ListViewP21.ListItems.Add((X + 1), , sayfakod, , 1)
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(1, , Format(msayfa.Cells(n, 7) * 100, "#,##0.0"))
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(2, , sayfakod)
    X = X + 1
devam:
Next n
bitti:
mug = 2
TextBoxP22.Text = ListViewP21.ListItems(1).ListSubItems(1).Text
TBLS1.Value = ""
Call malzemegrup
'Windows(mlzs).Activate
Application.ScreenUpdating = True
End Sub
Private Sub CommandButtonP21_Click()
On Error Resume Next
If ListBoxP21.ListCount = 0 Then Exit Sub
CheckBoxP21.Value = False
ListViewP21.ListItems.Clear
ListBoxP21.Selected(ListBoxP21.listIndex) = False
ListViewP21.SmallIcons = UFmy.ImageList4
Windows(dt).Activate
'Sheets("Sayfa1").Select '
'Range("A2").Select
son = Sheets("Sayfa1").Range("B65536").End(xlUp).row
markason = ListBoxP21.ListCount - 1
For ms = 0 To markason
For n = 2 To son
    If Sheets("Sayfa1").Cells(n, 4) = "" Then ' tüm malzemeler için toplam iskonto
    Do Until Sheets("Sayfa1").Cells(n, 4) <> ""
    n = n + 1
    If n > son Then GoTo bitti:
    Loop
    End If
    SayfaMarka = Sheets("Sayfa1").Cells(n, 4)
    sayfakod = Left(Sheets("Sayfa1").Cells(n, 1), 10)
    ayır = Split(sayfakod, "."): sayfakod = ayır(0) 'deneme ..
    If Sheets("Sayfa1").Cells(n, 4) = ListBoxP21.List(ms) Then Else GoTo devam:
    listson = ListViewP21.ListItems.Count
    k = ListViewP21.ListItems.Count
    Dim H
    H = 0
    Do Until k = H
    ListKod = ListViewP21.ListItems(k).Text
    If sayfakod = ListKod Then GoTo bak:
    k = k - 1
    Loop
    i = 0
    Do Until i = listson
    ListMarka = ListViewP21.ListItems(i + 1).Text
    If SayfaMarka = ListMarka Then GoTo git:
    i = i + 1
    Loop
UP:
    mz = ListViewP21.ListItems.Count
    Call ListViewP21.ListItems.Add((mz + 1), , ListBoxP21.List(ms), , 3)
    ListViewP21.ListItems(mz + 1).Bold = True
git:
    mz = ListViewP21.ListItems.Count
    Call ListViewP21.ListItems.Add((mz + 1), , sayfakod)
 
    Call ListViewP21.ListItems(mz + 1).ListSubItems.Add(1, , Format(Sheets("Sayfa1").Cells(n, 7) * 100, "#,##0.0"))
    Call ListViewP21.ListItems(mz + 1).ListSubItems.Add(2, , sayfakod)
    m = 1
    lt = 0
    mt = 0
    Do Until son = m
    sayfakod1 = Left(Sheets("Sayfa1").Cells(m + 1, 1), 10)
    ayır = Split(sayfakod1, "."): sayfakod1 = ayır(0) 'deneme..
    If sayfakod1 = sayfakod And Sheets("Sayfa1").Cells((m + 1), 4) = SayfaMarka Then
    lt = Format(lt + Sheets("Sayfa1").Cells(m + 1, 15), "#,##0.00")
    mt = Format(mt + Sheets("Sayfa1").Cells(m + 1, 16), "#,##0.00")
    End If
    m = m + 1
    Loop
    Call ListViewP21.ListItems(mz + 1).ListSubItems.Add(3, , lt)
    Call ListViewP21.ListItems(mz + 1).ListSubItems.Add(4, , mt)
    Call ListViewP21.ListItems(mz + 1).ListSubItems.Add(5, , SayfaMarka)
GoTo devam:
bak:
    i = 1
    Do Until i = listson
    ListMarka = ListViewP21.ListItems(i).Text
    If SayfaMarka = ListMarka And i < k Then GoTo devam:
    If SayfaMarka = ListMarka And i > k Then GoTo git:
    i = i + 1
    Loop
GoTo UP:
devam:
Next n
bitti:
Next ms
mug = 2
Call malzemegrup
Call listtoplam
ListViewP21.Refresh
End Sub
Private Sub CommandButtonP24_Click() 'montaj fiyatları
On Error Resume Next
Application.ScreenUpdating = False
If ListViewP21.ListItems.Count = 0 Then Exit Sub
Dim b, s
Dim aranan
Set eski = Workbooks(ComboBoxP22.Text).Worksheets("Sayfa1")
If Not WorkbookOpen("Montaj Fiyatları.xlsb") Then
Workbooks.Open "C:\Belgelerim\CEMEX\Parametreler\Montaj Fiyatları.xlsb"
End If
Set yeni = Workbooks("Montaj Fiyatları.xlsb").Worksheets("Parametreler")
ara = yeni.Range("A65536").End(xlUp).row
TextBoxP21.Text = ComboBoxP22.Text
bul = eski.Range("B65536").End(xlUp).row
ProgressBarP21.Visible = True: ProgressBarP21.Min = 0: ProgressBarP21.Max = bul
For i = 2 To bul
ProgressBarP21.Value = i
If eski.Cells(i, "D") = "" Then GoTo git
If eski.Cells(i, "K") <> "" Then kd1 = Left(eski.Range("K" & i), 2)
     kod = kd1 & "-" & eski.Range("I" & i)
        For Each aranan In yeni.Range("A" & "2 :" & "A" & ara)
            If aranan = kod Then
            s = aranan.row
            b = yeni.Range("G" & s)
            eski.Range("H" & i) = b: eski.Range("H" & i).Font.ColorIndex = 5
            GoTo git
            End If
        Next
git:
Next i
Range("A2").Select
ProgressBarP21.Visible = False
Workbooks(yeni).Close
Application.ScreenUpdating = True
End Sub
Private Sub CommandButtonP25_Click()
Call teklifisk
End Sub
Private Sub ListViewp21_Click()
On Error Resume Next
TextBoxP22.Enabled = True
SpinButtonP21.Enabled = True
TextBoxP22.Text = ListViewP21.SelectedItem.ListSubItems(1).Text
SpinButtonP21.Value = (ListViewP21.SelectedItem.ListSubItems(1).Text) * 10
TBLS1.Value = ListViewP21.SelectedItem.Index
ListViewP21.Refresh
End Sub
Private Sub ListViewP21_dblClick()
On Error GoTo git
If ListViewP21.SelectedItem.ListSubItems(1).Text = "" Then Exit Sub
UserFormISK.Show
git:
End Sub
Private Sub ListViewP21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
On Error GoTo git
If TBLS1.Value = "" Then
   If ListViewP21.ListItems.Count > 0 Then
   SpinButtonP21.Value = ListViewP21.SelectedItem.ListSubItems(1).Text * 10
   Else
   Exit Sub
   End If
End If
git:
End Sub
Private Sub SpinButton1_SpinUp()
On Error Resume Next
SpinButtonP21.Value = SpinButtonP21.Value + 50
End Sub
Private Sub SpinButton1_SpinDown()
On Error Resume Next
SpinButtonP21.Value = SpinButtonP21.Value - 50
End Sub
Private Sub SpinButtonP21_Change()
On Error Resume Next
    Dim nk As Single
    Dim ik As Single
'--
TextBoxP22.Text = Format(SpinButtonP21.Value * 0.1, "#,##0.0")
i = 0
listson = ListViewP21.ListItems.Count

    Do Until listson = i
    If ListViewP21.ListItems(i + 1).Selected = True Then
    sico = ListViewP21.ListItems(i + 1).SmallIcon
        If sico <> 3 And sico <> 5 Then
    ListViewP21.ListItems(i + 1).SmallIcon = 1
    'ListViewP21.ListItems(i + 1).Bold = True
    'ListViewP21.ListItems(i + 1).ListSubItems(1).Bold = True
    ListViewP21.ListItems(i + 1).ListSubItems(1).Text = TextBoxP22.Text
    isk = ListViewP21.ListItems(i + 1).ListSubItems(1).Text
    lt = ListViewP21.ListItems(i + 1).ListSubItems(3).Text
    ListViewP21.ListItems(i + 1).ListSubItems(4).Text = Format(lt - ((lt * isk) / 100), "#,##0.00")
        End If
    End If
    i = i + 1
    Loop

listson = ListViewP21.ListItems.Count - 1
i = 0
nt = 0
it = 0
Do Until listson = i
    If ListViewP21.ListItems(i + 1).ListSubItems(3) = "" Then GoTo git
    If i > listson Then GoTo bitti:

    'nk = ListViewP21.ListItems(i + 1).ListSubItems(3)
    ik = ListViewP21.ListItems(i + 1).ListSubItems(4)
    'nt = Format(nt + nk, "#,##0.00")
    it = Format(it + ik, "#,##0.00")
git:
    i = i + 1
    Loop
bitti:
'ListViewP21.ListItems(listson + 1).ListSubItems(3).Text = nt
ListViewP21.ListItems(listson + 1).ListSubItems(4).Text = it
CommandButtonP23.Enabled = True
End Sub
Private Sub ListBoxP21_Click() '2019
On Error Resume Next
CheckBoxP21.Value = True: CheckBoxP22.Value = False
Call iskaktar
For i = 0 To ListBoxP22.ListCount - 1
    ListBoxP22.Selected(i) = False
Next i
TextBoxP22.Enabled = True
SpinButtonP21.Enabled = True
TextBoxP22.Text = ListViewP21.SelectedItem.ListSubItems(1).Text
SpinButtonP21.Value = (ListViewP21.SelectedItem.ListSubItems(1).Text) * 10
End Sub
Private Sub ListBoxP22_Click() '2019
CheckBoxP22.Value = True: CheckBoxP21.Value = False
Call iskaktar
For i = 0 To ListBoxP21.ListCount - 1
    ListBoxP21.Selected(i) = False
Next i
End Sub
Sub iskaktar() '2019
On Error Resume Next
ListViewP21.ListItems.Clear
'For i = 1 To ListViewP21.ListItems.Count
    'ListViewP21.ListItems.Remove (1)
'Next i
TextBoxP22.Text = ""
TextBoxP22.Enabled = False
SpinButtonP21.Enabled = False
ListViewP21.SmallIcons = UFmy.ImageList4
If CheckBoxP21.Value = True Then Lmarka = ListBoxP21.Text
If CheckBoxP22.Value = True Then Lmarka = ListBoxP22.Text

'Sheets("Sayfa1").Select '
son = Sheets("Sayfa1").Range("B65536").End(xlUp).row
X = 0
z = 2
'--
For n = z To son
If Sheets("Sayfa1").Cells(n, 4) = "" Then
    Do Until Sheets("Sayfa1").Cells(n, 4) <> ""
    n = n + 1
    If n > son Then GoTo bitti:
    Loop
End If
    If Sheets("Sayfa1").Cells(n, 4) <> Lmarka Then
    Do Until Sheets("Sayfa1").Cells(n, 4) = Lmarka
    n = n + 1
    If n > son Then GoTo bitti:
    Loop
End If
    sayfakod = Replace(Sheets("Sayfa1").Cells(n, 1), "-auto", "")
    ayır = Split(sayfakod, "."): sayfakod = ayır(0)
    listson = ListViewP21.ListItems.Count
    i = 0
    Do Until listson = i
    ListKod = ListViewP21.ListItems(i + 1)
    If sayfakod = ListKod Then GoTo devam:
    i = i + 1
    Loop
Call ListViewP21.ListItems.Add((X + 1), , sayfakod)
isk1 = Sheets("Sayfa1").Cells(n, 7)
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(1, , Format((Sheets("Sayfa1").Cells(n, 7) * 100), "#,##0.0"))
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(2, , sayfakod)
m = 1
lt = 0
mt = 0
    Do Until son = m
    sayfakod1 = Left(Sheets("Sayfa1").Cells(m + 1, 1), 10)
    sayfakod1 = Replace(sayfakod1, "-auto", "")
    ayır = Split(sayfakod1, "."): sayfakod1 = ayır(0) 'deneme..
    If sayfakod1 = sayfakod And Sheets("Sayfa1").Cells(m + 1, 4) = Lmarka Then
    isk2 = Sheets("Sayfa1").Cells(m + 1, 7)
    If isk1 <> isk2 Then msg = 1: ListViewP21.ListItems(X + 1).SmallIcon = 11
    lt = lt + Sheets("Sayfa1").Cells(m + 1, 15)
    mt = mt + Sheets("Sayfa1").Cells(m + 1, 16)
    End If
    m = m + 1
    Loop
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(3, , Format(lt, "#,##0.00"))
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(4, , Format(mt, "#,##0.00"))
Call ListViewP21.ListItems(X + 1).ListSubItems.Add(5, , Lmarka)
    X = X + 1
    z = z + 1
devam:
Next n

bitti:
mug = 2
Call malzemegrup
Call listtoplam
ListViewP21.Refresh

If msg = 1 Then MsgBox Lmarka & " markalı aynı ürün gruplarının iskontolarında faklılıklar var. Düzeltmek için 'Aktar' Butonuna basınız! ", vbInformation, "scngnr@hotmail.com"
End Sub
Private Sub listtoplam()
On Error Resume Next
listson = ListViewP21.ListItems.Count
i = 0
nt = 0
it = 0
Do Until listson = i
    Dim nk As Single
    Dim ik As Single
    If ListViewP21.ListItems(i + 1).ListSubItems(3) = "" Then GoTo git
    If i > listson Then GoTo bitti:
    nk = ListViewP21.ListItems(i + 1).ListSubItems(3)
    ik = ListViewP21.ListItems(i + 1).ListSubItems(4)
    nt = Format(nt + nk, "#,##0.00")
    it = Format(it + ik, "#,##0.00")
git:
    i = i + 1
    Loop
bitti:
Call ListViewP21.ListItems.Add((listson + 1), , "TOPLAM", , 5)
Call ListViewP21.ListItems(listson + 1).ListSubItems.Add(1, , "")
Call ListViewP21.ListItems(listson + 1).ListSubItems.Add(2, , "")
Call ListViewP21.ListItems(listson + 1).ListSubItems.Add(3, , nt)
Call ListViewP21.ListItems(listson + 1).ListSubItems.Add(4, , it)
ListViewP21.ListItems(listson + 1).Bold = True: ListViewP21.ListItems(listson + 1).ListSubItems(3).Bold = True: ListViewP21.ListItems(listson + 1).ListSubItems(4).Bold = True
End Sub
Private Sub ListBoxP26_Click() '2023
On Error Resume Next
If CheckBox1.Value = False Then TBMLZTS = "Transfer.xlsb"
'TBMLZ2 = ""
CBTM1NO = ListBoxP26.listIndex + 1
If CheckBox1.Value = True Then
TBMLZTS = ListBoxP26.Text & ".xlsb"
Call malzemelisteleri3
End If
End Sub
Private Sub CommandButtonP26_Click()
If CheckBox2.Value = False Then
Call tbmmalzemeler1
ListBoxP27.Clear: ListBoxP28.Clear
Else
Call tbmmalzemeler2
ListBoxP27.Clear: ListBoxP28.Clear
End If
End Sub
Sub tbmmalzemeler2()
On Error Resume Next
ListBoxP26.Clear
Windows(dt).Activate
If Not Cells(1, 1) = "S.No" Then
  MsgBox ("   Bu işlem malzeme listesinin olduğu sayfada yapılabilir !  "), vbInformation, "scngnr@hotmail.com"
Else
tsayfa = ActiveSheet.Name
son = Sheets(tsayfa).Range("D65536").End(xlUp).row
For n = 2 To son
   If WorksheetFunction.CountIf(Sheets(tsayfa).Range("D2:D" & n), Sheets(tsayfa).Cells(n, 4).Value) = 1 And Sheets(tsayfa).Cells(n, 3).Value <> "" Then
    ListBoxP26.AddItem Sheets(tsayfa).Cells(n, 4) 'ListBox için yazdım
   End If
Next n
End If
End Sub
Sub tbmmalzemeler2XXX()
On Error Resume Next
ListBoxP26.Clear
Windows(dt).Activate
If Left(ActiveSheet.CodeName, 6) = "Icmal_" Then tsayfa = ActiveSheet.Name Else _
msg = MsgBox("Malzeme icmal sayfasında işlem yapın! ", vb, "scngnr@hotmail.com"): Exit Sub
Dim son As Integer
son = Sheets(tsayfa).Range("D65536").End(xlUp).row
For n = 2 To son
   If WorksheetFunction.CountIf(Sheets(tsayfa).Range("D2:D" & n), Sheets(tsayfa).Cells(n, 4).Value) = 1 And Sheets(tsayfa).Cells(n, 3).Value <> "" Then
    ListBoxP26.AddItem Sheets(tsayfa).Cells(n, 4) 'ListBox için yazdım
   End If
Next n
End Sub
Private Sub CommandButtonP261_Click()
If CheckBox1.Value = False Then
TBMLZTS = "Transfer.xlsb"
Call malzemelisteleri1
Else
If ListBoxP26.listIndex < 0 Then: msg = MsgBox(" Teklifte bulunan malzeme markası seçiniz! ", _
vb, "scngnr@hotmail.com"): Exit Sub
Call malzemeserilisteleri1
TBMLZTS = ListBoxP26.Text & ".xlsb"
End If
TBMLZ3 = ""
End Sub
Private Sub CheckBox1_Click()
ListBoxP27.Clear
ListBoxP28.Clear
TBMLZ2 = "": TBMLZ3 = ""
End Sub
Sub malzemeserilisteleri1()
On Error Resume Next
ListBoxP27.Clear
ListBoxP28.Clear
ListBoxPP28.Clear
For str1 = 0 To ListBoxP26.ListCount - 1
marka = ListBoxP26.List(str1, 0)
Dim ds, a
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists("C:\Belgelerim\Cemex\Transfer\" & marka & ".xlsb")
If a <> True Then GoTo git:
markads1 = Application.ExecuteExcel4Macro("'" & "C:\Belgelerim\Cemex\Transfer\[" & marka & ".xlsb]" & marka & "'!R1C" & 4)
Dim str As Integer
For str = 4 To 30
markads = Application.ExecuteExcel4Macro("'" & "C:\Belgelerim\Cemex\Transfer\[" & marka & ".xlsb]" & marka & "'!R1C" & str)
If markads = "" Or markads = 0 Then markads = markads1
markaseri = Application.ExecuteExcel4Macro("'" & "C:\Belgelerim\Cemex\Transfer\[" & marka & ".xlsb]" & marka & "'!R2C" & str)
If markaseri <> 0 And markaseri <> "" Then
       ListBoxPP28.AddItem markaseri
       ListBoxPP28.List(ListBoxPP28.ListCount - 1, 1) = markads
       ListBoxPP28.List(ListBoxPP28.ListCount - 1, 2) = marka
End If
Next
git:
Next
ListBoxP26_Click
'ListBoxP26.Selected(ListBoxP26.ListIndex) = False
End Sub
Sub tbmmalzemeler1()
On Error Resume Next
ListBoxP26.Clear
Windows(dt).Activate
'Sheets("Sayfa1").Select '
Dim son As Integer
son = Sheets("Sayfa1").Range("D65536").End(xlUp).row
For n = 2 To son
   If WorksheetFunction.CountIf(Sheets("Sayfa1").Range("D2:D" & n), Sheets("Sayfa1").Cells(n, 4).Value) = 1 And Sheets("Sayfa1").Cells(n, 3).Value <> "" Then
    ListBoxP26.AddItem Sheets("Sayfa1").Cells(n, 4) 'ListBox için yazdım
   End If
Next n
End Sub
Sub malzemelisteleri3() '2023
On Error Resume Next
ListBoxP27.Clear
ListBoxP28.Clear
Call transferliste1
mal = ListBoxP26.List(ListBoxP26.listIndex)
Dim str As Integer
For str = 0 To ListBoxPP28.ListCount - 1
markaad = ListBoxPP28.List(str, 0): markads = ListBoxPP28.List(str, 1): marka = ListBoxPP28.List(str, 2)
If marka = mal Then
ListBoxP27.AddItem markaad
ListBoxP27.List(ListBoxP27.ListCount - 1, 1) = markads
ListBoxP28.AddItem markaad
ListBoxP28.List(ListBoxP28.ListCount - 1, 1) = markads
TBMLZ2 = markads & ".xlsb"
End If
Next
If ListBoxP27.ListCount = 0 Then TBMLZ2 = ""
End Sub
Sub malzemelisteleri4x()
On Error Resume Next
ListBoxP28.Clear
mal = ListBoxP27.List(ListBoxP27.listIndex)
Dim str As Integer
For str = 0 To ListBoxP27.ListCount - 1
markaad = ListBoxP27.List(str, 0): markads = ListBoxP27.List(str, 1)
'If markaad <> mal Then
ListBoxP28.AddItem markaad
ListBoxP28.List(ListBoxP28.ListCount - 1, 1) = markads
'End If
Next
End Sub
Sub malzemelisteleri1()
On Error Resume Next
ListBoxP27.Clear
ListBoxP28.Clear
ListBoxPP28.Clear
Dim ds, a
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists("C:\Belgelerim\Cemex\Transfer\Transfer.xlsb")
If a <> True Then msg = MsgBox("Transfer dosyası mevcut değil!", vb, "scngnr@hotmail.com"): Exit Sub
Dim str As Integer
For str = 4 To 60
markads = Application.ExecuteExcel4Macro("'" & "C:\Belgelerim\Cemex\Transfer\[Transfer.xlsb]" & "Markalar'!R1C" & str)
markaad = Application.ExecuteExcel4Macro("'" & "C:\Belgelerim\Cemex\Transfer\[Transfer.xlsb]" & "Markalar'!R2C" & str)
If markaad <> 0 And markaad <> "" Then
 If Not mal = Empty Then
       If markaad Like "*" & mal & "*" Then
       ListBoxPP28.AddItem markaad
       ListBoxPP28.List(ListBoxPP28.ListCount - 1, 1) = markads
       GoTo git:
       End If
  End If
ListBoxP27.AddItem markaad
ListBoxP27.List(ListBoxP27.ListCount - 1, 1) = markads
mal = markaad
git:
End If
Next
TBMLZ2 = ""
End Sub
Sub malzemelisteleri2()
On Error Resume Next
ListBoxP28.Clear
Dim str As Integer
mal = ListBoxP27.List(ListBoxP27.listIndex)
For str = 0 To ListBoxPP28.ListCount - 1
markaad = ListBoxPP28.List(str, 0): markads = ListBoxPP28.List(str, 1)
If markaad Like "*" & mal & "*" Then
ListBoxP28.AddItem markaad
ListBoxP28.List(ListBoxP28.ListCount - 1, 1) = markads
End If
Next
End Sub
Private Sub ListBoxP27_Click()
On Error Resume Next
Call transferliste1
CBMTD1NO = ListBoxP27.listIndex + 1
If ListBoxP27.listIndex < 0 Then Exit Sub
If CheckBox1.Value = True Then
'Call malzemelisteleri4
If ListBoxP27.listIndex >= 0 Then TBMLZ3 = ListBoxP27.List(ListBoxP27.listIndex, 0)
Exit Sub
End If
Call malzemelisteleri2
If ListBoxP28.listIndex >= 0 Then ListBoxP28.Selected(ListBoxP28.listIndex) = False
If ListBoxP27.listIndex >= 0 Then TBMLZ3 = ListBoxP27.List(ListBoxP27.listIndex, 0)
End Sub
Sub transferliste1()
On Error Resume Next
trf1 = ListBoxP27.List(ListBoxP27.listIndex, 1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(fm1 & "\Malzeme Listeleri")
Set colSubfolders = objFolder.SubFolders
For Each objSubfolder In colSubfolders
    ds = objSubfolder.Name
  Dim dsm
  Dim n As Integer
  dsm = dir(fm1 & "\Malzeme Listeleri\" & ds & "\*.xlsb")
  n = 1
  Do While dsm <> ""
     If dsm Like "*" & trf1 & "*" Then TBMLZ2 = dsm: Exit Sub
    dsm = dir
    n = n + 1
  Loop
Next
End Sub
Private Sub ListBoxP28_Click()
'If Len(ListBoxP28.List(ListBoxP28.ListIndex, 1)) > 1 Then
'TBMLZ2 = ListBoxP28.List(ListBoxP28.ListIndex, 1)
'End If
If CheckBox1.Value = True Then
If ListBoxP28.listIndex >= 0 Then TBMLZ3 = ListBoxP28.List(ListBoxP28.listIndex, 0)
Exit Sub
End If
If ListBoxP27.listIndex >= 0 Then ListBoxP27.Selected(ListBoxP27.listIndex) = False
If ListBoxP28.listIndex >= 0 Then TBMLZ3 = ListBoxP28.List(ListBoxP28.listIndex, 0)
End Sub
Private Sub CommandButtonP28_Click()
If CheckBox2.Value = False Then
Windows(dt).Activate: Sheets("Sayfa1").Select
Call markaaktar1
Else
Call markaaktar2
End If
End Sub
Sub markaaktar2()
On Error Resume Next
dtskol = Cells(1, 256).End(xlToLeft).Column ' kolon sayısı
For k = 1 To dtskol
If Not Cells(1, k) <> "" Then Exit For
Next
dtskol0 = Cells(1, k).Column - 1
If dtskol0 < 9 Then msg = MsgBox(" İcmal sayfasında iskonto ve fiyat satırlarında eksikler var ! ", vb, "scngnr@hotmail.com"): Exit Sub
''''
If ListBoxP26.listIndex < 0 Then msg = MsgBox(" Teklifte bulunan malzeme markası seçiniz!", vb, "scngnr@hotmail.com"): Exit Sub
If ListBoxP27.listIndex < 0 And ListBoxP28.listIndex < 0 Then msg = MsgBox(" 2. ve/veya 3. Listeden seçim yapınız! ", vb, "scngnr@hotmail.com"): Exit Sub
''''
If ListBoxP27.listIndex >= 0 Then malzs2 = ListBoxP27.List(ListBoxP27.listIndex)
If ListBoxP28.listIndex >= 0 Then malzs2 = ListBoxP28.List(ListBoxP28.listIndex)
 malzs1 = ListBoxP26.List(ListBoxP26.listIndex)
''''
If CheckBox1.Value = True Then
 If ListBoxP27.listIndex < 0 Then msg = MsgBox(" 2.Listeden seçim yapınız! ", vb, "scngnr@hotmail.com"): Exit Sub
 If ListBoxP28.listIndex < 0 Then msg = MsgBox(" 3.Listeden seçim yapınız! ", vb, "scngnr@hotmail.com"): Exit Sub
 malzs1 = ListBoxP27.List(ListBoxP27.listIndex)
 malzs2 = ListBoxP28.List(ListBoxP28.listIndex)
End If
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
dt = ActiveWorkbook.Name
dts = ActiveSheet.Name
Dim b, c, d, arateklif, aramalzeme
arateklif = Workbooks(dt).Worksheets(dts).Range("B65536").End(xlUp).row
mlz = ListBoxP26.List(ListBoxP26.listIndex)
malzListe = TBMLZ2
''''
ProgressBarP21.Min = 0: ProgressBarP21.Value = 0: ProgressBarP21.Visible = True
Transferdosya = TBMLZTS
Workbooks.Open ("C:\Belgelerim\Cemex\Transfer\" & Transferdosya)
Application.Windows(Transferdosya).Visible = False
sayfa = Workbooks(Transferdosya).ActiveSheet.Name
Dim trfliste
Set trfliste = Workbooks(Transferdosya).Worksheets(sayfa)
''''
aranan1 = malzs1: aranan2 = malzs2
''''
Set aranan11 = trfliste.Rows("2:2").Find(aranan1, LookAt:=xlWhole).Rows
Set aranan21 = trfliste.Rows("2:2").Find(aranan2, LookAt:=xlWhole).Rows

If aranan11 = Empty Or aranan21 = Empty Then
msg = MsgBox("    Seçilen markaya ait malzeme bilgileri mevcut değil! ", vb, "scngnr@hotmail.com")
GoTo SayfaYok
End If
a1row = aranan11.Column: a2row = aranan21.Column
Dim ds, f, f1, fc
Set ds = CreateObject("Scripting.FileSystemObject")
Set f = ds.GetFolder(fm1 & "\Malzeme Listeleri")
Set fc = f.SubFolders
'--
For Each f1 In fc
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists(f1 & "\" & malzListe)
If a = True Then Workbooks.Open fileName:=f1 & "\" & malzListe: Exit For
Next
'--
Application.Windows(malzListe).Visible = False
Dim mliste
Set mliste = Workbooks(malzListe).Worksheets("Sayfa1")
ProgressBarP21.Value = 0
aramalzeme = mliste.Range("B65536").End(xlUp).row
kodson = trfliste.Cells(Rows.Count, a1row).End(xlUp).row
Transfer = Range(Cells(2, a1row), Cells(kodson, a1row)).Address

sutun = a2row - a1row
Workbooks(dt).Activate
''''
If CheckBoxsi.Value = True Then si2 = ActiveCell.row: arateklif2 = si2 + Selection.Cells.Count - 1
''''
Worksheets(dts).Range(Columns(1), Columns(dtskol0)).Copy
Cells(1, dtskol + 2).Select: ActiveSheet.Paste
Application.CutCopyMode = False: Cells(2, dtskol0 + 2).Select
ActiveWindow.FreezePanes = True
Worksheets(dts).Range(Cells(1, dtskol + 2), Cells(1, dtskol0 + dtskol + 1)).Interior.ColorIndex = 40
dtskol2 = Cells(1, 256).End(xlToLeft).Column + 1 ' kolon sayısı
''''
ProgressBarP21.Max = arateklif
For i = 2 To arateklif
ProgressBarP21.Value = i
If CheckBoxsi.Value = True Then
If i < si2 Then GoTo atla
If i > arateklif2 Then GoTo atla
End If
Set s1mark = Worksheets(dts).Range("B" & i)
If Worksheets(dts).Range("D" & i) = mlz Then
    kod1 = s1mark
    Set aranan = trfliste.Range(Transfer).Find(kod1, LookIn:=xlValues, LookAt:=xlWhole)
If Not aranan Is Nothing Then
    kod2 = aranan.Offset(0, sutun)
    If kod2 = "" Then Worksheets(dts).Range(Cells(i, 1), Cells(i, dtskol - 1)).Font.ColorIndex = 46: GoTo atla 'turuncu
''''
    Set bulunan = mliste.Range("B2:" & "B" & aramalzeme).Find(kod2, LookIn:=xlValues, LookAt:=xlWhole)
    If bulunan Is Nothing Then GoTo git:
    b = bulunan.Offset(0, 1) 'açıklama
    c = bulunan.Offset(0, 2) 'üretici
    d = bulunan.Offset(0, 4) 'br.fiyat
    e = bulunan.Offset(0, 5) 'isk
    Worksheets(dts).Cells(i, dtskol0).Interior.ColorIndex = 36 'SARI
''''
    Worksheets(dts).Cells(i, dtskol + 3) = kod2
    Worksheets(dts).Cells(i, dtskol + 4) = b
    Worksheets(dts).Cells(i, dtskol + 5) = c
    Worksheets(dts).Cells(i, dtskol + 7) = d
    Worksheets(dts).Cells(i, dtskol + 7).NumberFormat = mliste.Range("F" & bulunan.row).NumberFormat
    Worksheets(dts).Cells(i, dtskol + 8) = e
    Worksheets(dts).Range(Cells(i, 1), Cells(i, dtskol - 1)).Font.ColorIndex = 10 'yeşil
    Worksheets(dts).Range(Cells(i, dtskol), Cells(i, dtskol2 - 1)).Font.ColorIndex = 51 'yeşil
    'KUR'--
    kur = ""
    If Worksheets(dts).Cells(i, dtskol + 7).NumberFormat = "#,##0.00 [$$-C0C]" Then ' Sayfa3 de $ kuru
    Worksheets(dts).Cells(i, dtskol + 7).Font.ColorIndex = 3: kur = "*Usd": GoTo kurson
    End If
    If Worksheets(dts).Cells(i, dtskol + 7).NumberFormat = "#,##0.00 [$€-1]" Then ' Sayfa3 de € kuru
    Worksheets(dts).Cells(i, dtskol + 7).Font.ColorIndex = 5: kur = "*Eur": GoTo kurson
    End If
kurson:
    Worksheets(dts).Cells(i, dtskol + 9).FormulaR1C1 = "=(RC[-2]-RC[-2]*RC[-1])" & kur
    Worksheets(dts).Cells(i, dtskol + 10).FormulaR1C1 = "=RC[-4]*RC[-1]" 'Mlz. List Top.
    GoTo atla
Else
git:
Worksheets(dts).Range(Cells(i, 1), Cells(i, dtskol - 1)).Font.ColorIndex = 46
Worksheets(dts).Range(Cells(i, dtskol), Cells(i, dtskol2 - 1)).Font.ColorIndex = 53
End If
End If
atla:
If Not Worksheets(dts).Range("E" & i) = "" Then
Worksheets(dts).Cells(i, dtskol + 6).FormulaR1C1 = "=RC5"
If Worksheets(dts).Cells(i, dtskol + 10) < Worksheets(dts).Range("I" & i) Then Worksheets(dts).Cells(i, dtskol + 10).Interior.ColorIndex = 35
If Worksheets(dts).Cells(i, dtskol + 10) > Worksheets(dts).Range("I" & i) Then Worksheets(dts).Cells(i, dtskol + 10).Interior.ColorIndex = 38
End If
ProgressBarP21.Value = i
Next i

SayfaYok:
Workbooks(malzListe).Close (False)
Workbooks(Transferdosya).Close (False)
ProgressBarP21.Visible = False
'TBMLZ2 = ""
If CheckBox1.Value = False Then ListBoxP26.listIndex = CBTM1NO - 1
Set trfliste = Nothing: Set s1mark = Nothing: Set aranan11 = Nothing
Set aranan21 = Nothing: Set aranan = Nothing: Set bulunan = Nothing
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub markaaktar1()
On Error Resume Next
If ListBoxP26.listIndex < 0 Then msg = MsgBox(" Teklifte bulunan malzeme markası seçiniz!", vb, "scngnr@hotmail.com"): Exit Sub
If ListBoxP27.listIndex < 0 And ListBoxP28.listIndex < 0 Then msg = MsgBox(" 2. ve/veya 3. Listeden seçim yapınız! ", vb, "scngnr@hotmail.com"): Exit Sub
''''
If ListBoxP27.listIndex >= 0 Then malzs2 = ListBoxP27.List(ListBoxP27.listIndex)
If ListBoxP28.listIndex >= 0 Then malzs2 = ListBoxP28.List(ListBoxP28.listIndex)
 malzs1 = ListBoxP26.List(ListBoxP26.listIndex)
''''
If CheckBox1.Value = True Then
 If ListBoxP27.listIndex < 0 Then msg = MsgBox(" 2.Listeden seçim yapınız! ", vb, "scngnr@hotmail.com"): Exit Sub
 If ListBoxP28.listIndex < 0 Then msg = MsgBox(" 3.Listeden seçim yapınız! ", vb, "scngnr@hotmail.com"): Exit Sub
 malzs1 = ListBoxP27.List(ListBoxP27.listIndex)
 malzs2 = ListBoxP28.List(ListBoxP28.listIndex)
End If
''''
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual

dt = ActiveWorkbook.Name
Dim b, c, d, arateklif, aramalzeme
arateklif = Workbooks(dt).Worksheets("Sayfa1").Range("B65536").End(xlUp).row
mlz = ListBoxP26.List(ListBoxP26.listIndex)
malzListe = TBMLZ2
''''
ProgressBarP21.Min = 0: ProgressBarP21.Value = 0: ProgressBarP21.Visible = True
Transferdosya = TBMLZTS

Workbooks.Open ("C:\Belgelerim\Cemex\Transfer\" & Transferdosya)
Application.Windows(Transferdosya).Visible = False
'sayfa = mlz & "-" & Left(malzListe, 3):'sayfa = "Markalar" 'sayfa seçimi
sayfa = Workbooks(Transferdosya).ActiveSheet.Name
Dim trfliste
Set trfliste = Workbooks(Transferdosya).Worksheets(sayfa)
''''
aranan1 = malzs1: aranan2 = malzs2
''''
Set aranan11 = trfliste.Rows("2:2").Find(aranan1, LookAt:=xlWhole).Rows
Set aranan21 = trfliste.Rows("2:2").Find(aranan2, LookAt:=xlWhole).Rows

If aranan11 = Empty Or aranan21 = Empty Then
msg = MsgBox("    Seçilen markaya ait malzeme bilgileri mevcut değil! ", vb, "scngnr@hotmail.com")
GoTo SayfaYok
End If
a1row = aranan11.Column: a2row = aranan21.Column
Dim ds, f, f1, fc
Set ds = CreateObject("Scripting.FileSystemObject")
Set f = ds.GetFolder(fm1 & "\Malzeme Listeleri")
Set fc = f.SubFolders
'Dim mlz As ListItem
'--
For Each f1 In fc
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists(f1 & "\" & malzListe)
If a = True Then Workbooks.Open fileName:=f1 & "\" & malzListe: Exit For
Next
'--
Application.Windows(malzListe).Visible = False
Dim mliste
Set mliste = Workbooks(malzListe).Worksheets("Sayfa1")
ProgressBarP21.Value = 0
aramalzeme = mliste.Range("B65536").End(xlUp).row
kodson = trfliste.Cells(Rows.Count, a1row).End(xlUp).row
Transfer = Range(Cells(2, a1row), Cells(kodson, a1row)).Address

sutun = a2row - a1row
Workbooks(dt).Activate
''''
si = 2
If CheckBoxsi.Value = True Then si = ActiveCell.row: arateklif = si + Selection.Cells.Count - 1
ProgressBarP21.Max = arateklif
For i = si To arateklif
ProgressBarP21.Value = i
Set s1mark = Worksheets("Sayfa1").Range("B" & i)
If Worksheets("Sayfa1").Range("D" & i) = mlz Then
    kod1 = s1mark
    Set aranan = trfliste.Range(Transfer).Find(kod1, LookIn:=xlValues, LookAt:=xlWhole)
If Not aranan Is Nothing Then
    kod2 = aranan.Offset(0, sutun)
    If kod2 = "" Then Worksheets("Sayfa1").Range("B" & i, "G" & i).Font.ColorIndex = 46: GoTo atla 'turuncu
''''
    Set bulunan = mliste.Range("B2:" & "B" & aramalzeme).Find(kod2, LookIn:=xlValues, LookAt:=xlWhole)
    If bulunan Is Nothing Then GoTo git:
    a = mliste.Range("K" & bulunan.row + 1).End(3) 'Referans
    b = bulunan.Offset(0, 1) 'açıklama
    c = bulunan.Offset(0, 2) 'üretici
    d = bulunan.Offset(0, 4) 'br.fiyat
    e = bulunan.Offset(0, 5) 'isk
    f = bulunan.Offset(0, 6) 'adam/dak.
''''
    Worksheets("Sayfa1").Range("A" & i) = a
    Worksheets("Sayfa1").Range("B" & i) = kod2
    Worksheets("Sayfa1").Range("C" & i) = b
    Worksheets("Sayfa1").Range("D" & i) = c
    Worksheets("Sayfa1").Range("F" & i) = d
    Worksheets("Sayfa1").Range("F" & i).NumberFormat = mliste.Range("F" & bulunan.row).NumberFormat
    Worksheets("Sayfa1").Range("G" & i) = e
    Worksheets("Sayfa1").Range("H" & i) = f
    Worksheets("Sayfa1").Range("B" & i, "G" & i).Font.ColorIndex = 10 'yeşil
    'KUR'--
    If Worksheets("Sayfa1").Range("F" & i).NumberFormat = "#,##0.00 [$$-C0C]" Then ' Sayfa3 de $ kuru
    Worksheets("Sayfa1").Range("F" & i).Font.ColorIndex = 3: kur = "*Usd"
    Worksheets("Sayfa1").Range("K" & i).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur
    Worksheets("Sayfa1").Range("O" & i).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
    End If
    If Worksheets("Sayfa1").Range("F" & i).NumberFormat = "#,##0.00 [$€-1]" Then ' Sayfa3 de € kuru
    Worksheets("Sayfa1").Range("F" & i).Font.ColorIndex = 5: kur = "*Eur"
    Worksheets("Sayfa1").Range("K" & i).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur
    Worksheets("Sayfa1").Range("O" & i).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
    End If
    GoTo atla
Else
git:
Worksheets("Sayfa1").Range("B" & i, "G" & i).Font.ColorIndex = 46
End If
'Else
'Worksheets("Sayfa1").Range("B" & i, "G" & i).Font.ColorIndex = 1
End If
atla:
ProgressBarP21.Value = i
Next i
If CheckBox1.Value = False Then Call tbmmalzemeler1
SayfaYok:
'Range("A2").Select
'ListBoxP27.Selected(ListBoxP27.ListIndex) = False
Workbooks(malzListe).Close (False)
Workbooks(Transferdosya).Close (False)
ProgressBarP21.Visible = False
'TBMLZ2 = ""
If CheckBox1.Value = False Then ListBoxP26.listIndex = CBTM1NO - 1
TBMLZ2 = malzListe
Set trfliste = Nothing: Set s1mark = Nothing: Set aranan11 = Nothing
Set aranan21 = Nothing: Set aranan = Nothing: Set bulunan = Nothing
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub malzemegrup() ' Ürün Grupları
On Error Resume Next
Dim listson As Integer
Dim i, X, mugson As Integer
Dim kod, kodL
listson = ListViewP21.ListItems.Count
mugson = ListBoxMK.ListCount
i = 0
Do Until i = listson
If ListViewP21.ListItems(i + 1).ListSubItems(2).Text = "" And ListViewP21.ListItems(i + 1).Text <> "" Then GoTo git1:
kod = ListViewP21.ListItems(i + 1).ListSubItems(2).Text
  For X = 0 To mugson - 1
  kodL = ListBoxMK.List(X, 0)
  If kod = kodL Then ListViewP21.ListItems(i + 1).Text = ListBoxMK.List(X, 1): Exit For
  Next X
git1:
    i = i + 1
Loop
End Sub
Sub tkldosyalar()
Windows(dt).Activate: Sheets("Sayfa1").Select
CBTKL1.Clear
Dim n
Dim son As Integer
son = Range("D65536").End(xlUp).row
X = 0
For n = 2 To son
    If WorksheetFunction.CountIf(Range("D2:D" & n), Cells(n, 4).Value) = 1 And Cells(n, 3).Value <> "" Then
    CBTKL1.AddItem Cells(n, 4)
    End If
    X = X + 1
Next n
End Sub
Private Sub CBTKL1_Change()
On Error Resume Next
If CBTKL1.Text = "" Then Exit Sub
tkldosya
TextBoxTKL1 = "": TextBoxTKL2 = "": TBTKL1 = "": TBTKLF01 = ""
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(CBTKL1.Text, 3) & "\Logo.jpg")
End Sub
Sub tkldosya()
ListBoxTKL1.Clear
Dim Y, z, a, i, p As Single
Dim s As Double
Dim listson As Long
Dim SayfaMarka, formmarka, formkod, mm As String
        a = WorksheetFunction.CountA(Sheets("Sayfa1").Range("B:B"), xlDown) - 1
        If a < 2 Then Exit Sub
        a = 0
ali:
      formmarka = CBTKL1.Text
        p = WorksheetFunction.CountIf(Sheets("Sayfa1").Range("D:D"), formmarka)
        Y = 0
        z = 0
Do Until Y = p
If Y > p Then Exit Sub
yok:
    If Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1) = "" Then
                Do Until Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1) <> ""
                z = z + 1
                If z > Range("B65536").End(xlUp).row Then Exit Sub
                If Y > p Then Exit Sub
                Loop
    Else:
                SayfaMarka = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 3)
                If formmarka = SayfaMarka Then GoTo devam:
                z = z + 1
GoTo yok:
devam:
                Dim kt As String
                kt = Sheets("Sayfa1").Range("A1").Offset(Y + z + 1, 1)
                    i = 0
                listson = ListBoxTKL1.ListCount
                Do Until listson = i
                formkod = ListBoxTKL1.List(i, 0)
                    If formkod = kt Then GoTo var:
                    i = i + 1
                Loop
        ListBoxTKL1.AddItem Range("A1").Offset(Y + z + 1, 1)
        ListBoxTKL1.List(ListBoxTKL1.ListCount - 1, 1) = Range("A1").Offset(Y + z + 1, 2)
        ListBoxTKL1.List(ListBoxTKL1.ListCount - 1, 3) = Range("A1").Offset(Y + z + 1, 5).Text
var:
        Y = Y + 1
    End If
Loop
CBTKL1AD = ListBoxTKL1.ListCount
'Application.ScreenUpdating = True:'Application.EnableEvents = True
If Me.ListBoxTKL1.ListCount <> 0 Then
ListBoxTKL1.List = Diz(ListBoxTKL1.List, 1)
End If
Exit Sub
GoTo ali
End Sub
Private Sub TextBoxTKL1_Change()
On Error Resume Next
If CBTKL1 = "" Then Exit Sub
If Len(TextBoxTKL1) > 3 Then
Call tkldosya2: Exit Sub
Else
If Len(TextBoxTKL2) > 3 Then Call tkldosya: Call tkldosya2: Exit Sub
If ListBoxTKL1.ListCount <> CBTKL1AD Then
tkldosya
Else
Exit Sub
End If
End If
End Sub
Private Sub TextBoxTKL2_Change()
If CBTKL1 = "" Then Exit Sub
On Error Resume Next
If Len(TextBoxTKL2) > 3 Then
Call tkldosya2: Exit Sub
Else
If Len(TextBoxTKL1) > 3 Then Call tkldosya: Call tkldosya2: Exit Sub
If ListBoxTKL1.ListCount <> CBTKL1AD Then
tkldosya
Else
Exit Sub
End If
End If
End Sub
Private Sub CBTKL_Click()
If TextBoxMLZ1.Enabled = False Then Exit Sub
TextBoxMLZ1.Text = TextBoxTKL1.Text
TextBoxMLZ2.Text = TextBoxTKL2.Text
End Sub
Sub tkldosya2()
On Error Resume Next
        For i = ListBoxTKL1.ListCount - 1 To 0 Step -1
            If Not (ListBoxTKL1.List(i, 1) Like "*" & TextBoxTKL1.Text & "*" And ListBoxTKL1.List(i, 1) Like "*" & TextBoxTKL2.Text & "*") Then
                ListBoxTKL1.RemoveItem (i)
            End If
        Next
End Sub
Sub mlzdosyalar()
CBMLZ1.Clear
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(fm1 & "\Malzeme Listeleri")
Set colSubfolders = objFolder.SubFolders
For Each objSubfolder In colSubfolders
    ds = objSubfolder.Name
  Dim dsm
  Dim n As Integer
  dsm = dir(fm1 & "\Malzeme Listeleri\" & ds & "\*.xlsb")
  n = 1
  Do While dsm <> ""
    CBMLZ1.AddItem dsm
    dsm = dir
    n = n + 1
  Loop
Next
End Sub
Private Sub CBMLZ1_Change()
On Error Resume Next
If CBMLZ1.Text <> TBMLZ1.Text Then
Application.DisplayAlerts = False
Windows(TBMLZ1.Text).Close SaveChanges:=True
Application.DisplayAlerts = True
ListBoxMLZ1.Clear
End If
Application.ScreenUpdating = False
Set ds = CreateObject("Scripting.FileSystemObject")
Set f = ds.GetFolder(fm1 & "\Malzeme Listeleri")
Set fc = f.SubFolders
Dim mlz As listItem
'--
For Each f1 In fc
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists(f1 & "\" & CBMLZ1.Text)
If a = True Then Workbooks.Open fileName:=f1 & "\" & CBMLZ1.Text: Application.Windows(CBMLZ1.Text).Visible = False: Exit For
Next
'--
TBMLZ1.Text = CBMLZ1.Text
TextBoxMLZ1.Enabled = True: TextBoxMLZ2.Enabled = True
TextBoxMLZ1 = "": TextBoxMLZ2 = "": TBMLZ01 = "": TBMLZF01 = ""
Image2.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(CBMLZ1.Text, 3) & "\Logo.jpg")
End Sub
Private Sub TextBoxMLZ1_Change()
Call mlzarama
End Sub
Private Sub TextBoxMLZ2_Change()
Call mlzarama
End Sub
Sub mlzarama()
If Len(TextBoxMLZ2) <= 3 And Len(TextBoxMLZ1) <= 3 Then
 If ListBoxMLZ1.ListCount > 0 Then ListBoxMLZ1.Clear
TBMLZ0 = "": TBMLZ01 = "": TBMLZF01 = "": Exit Sub
End If
If Len(TextBoxMLZ1) > 3 Then
 If Len(TextBoxMLZ2) > CDbl(TBMLZ2L.Value) Then Call mlzdosya2 Else malzdosya: Call mlzdosya2
Else
 If Len(TextBoxMLZ2) > 3 Then Call malzdosya: Call mlzdosya2
End If
TBMLZ2L = Len(TextBoxMLZ2)
End Sub
Sub mlzdosya2()
On Error Resume Next
        For i = ListBoxMLZ1.ListCount - 1 To 0 Step -1
            If Not (ListBoxMLZ1.List(i, 1) Like "*" & TextBoxMLZ1.Text & "*" And ListBoxMLZ1.List(i, 1) Like "*" & TextBoxMLZ2.Text & "*") Then
                ListBoxMLZ1.RemoveItem (i)
            End If
        Next
End Sub
Private Sub LBTM1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If LBTM1.BackColor = &HE6D3C4 Then
LBTM1.BackColor = &HCAE3BF
Else
LBTM1.BackColor = &HE6D3C4
End If
'mlzdosya21
End Sub
Sub mlzdosya21() 'XXXXX
On Error Resume Next
ListBoxMLZ1.Clear
Set mzs = Workbooks(CBMLZ1.Text).Worksheets("Sayfa1")
son = mzs.Range("B65536").End(xlUp).row + 1
Y = 2
Do Until Y >= son
If Not mzs.Range("D" & Y) = "" Then
If (mzs.Cells(Y, 3) Like "*" & LBTM1 & "*" And mzs.Cells(Y, 3) Like "*" & LBTM2 & "*" _
And mzs.Cells(Y, 3) Like "*" & LBTM3 & "*" And mzs.Cells(Y, 3) Like "*" & LBTM4 & "*") Then
ListBoxMLZ1.AddItem mzs.Range("B" & Y)
ListBoxMLZ1.List(ListBoxMLZ1.ListCount - 1, 1) = mzs.Range("C" & Y)
ListBoxMLZ1.List(ListBoxMLZ1.ListCount - 1, 2) = c.row
ListBoxMLZ1.List(ListBoxMLZ1.ListCount - 1, 3) = mzs.Range("F" & Y).Text
End If
End If
Y = Y + 1
Loop
End Sub
Sub malzdosya()
On Error Resume Next
Dim ad, deg
Dim b, c
If Len(TextBoxMLZ1) > 3 Then ad = TextBoxMLZ1.Text: GoTo git:
If Len(TextBoxMLZ2) > 3 Then ad = TextBoxMLZ2.Text
git:
ListBoxMLZ1.Clear
X = 0
deg = ""
Set c = Workbooks(CBMLZ1.Text).Worksheets("Sayfa1").Range("C2:C65000").Find(ad, LookAt:=xlPart)
If Not c Is Nothing Then
b = c.Address
Do
If c.row <> deg Then
ListBoxMLZ1.AddItem Workbooks(CBMLZ1.Text).Worksheets("Sayfa1").Range("B" & c.row)
ListBoxMLZ1.List(ListBoxMLZ1.ListCount - 1, 1) = Workbooks(CBMLZ1.Text).Worksheets("Sayfa1").Range("C" & c.row)
ListBoxMLZ1.List(ListBoxMLZ1.ListCount - 1, 2) = c.row
ListBoxMLZ1.List(ListBoxMLZ1.ListCount - 1, 3) = Workbooks(CBMLZ1.Text).Worksheets("Sayfa1").Range("F" & c.row).Text
deg = c.row
Set c = Workbooks(CBMLZ1.Text).Worksheets("Sayfa1").Range("C2:C65000").FindNext(c)
End If
Loop While Not c Is Nothing And c.Address <> b
End If
End Sub
Private Sub ListBoxTKL1_Click()
On Error Resume Next
CBTKL1NO = ListBoxTKL1.listIndex + 1
TBTKL0 = ListBoxTKL1.List(ListBoxTKL1.listIndex, 0)
TBTKL1 = ListBoxTKL1.List(ListBoxTKL1.listIndex, 1)
TBTKLF01 = ListBoxTKL1.List(ListBoxTKL1.listIndex, 3)
'XXXXX
'LBTM1 = "": LBTM2 = "": LBTM3 = "": LBTM4 = "": LBTM5 = ""
'ayır = Split(TBTKL1, " ")
'tsay = Len(TBTKL1) - Len(Replace(TBTKL1, " ", ""))
'z = 1
'For n = 0 To tsay '
'If UBound(ayır) <> 0 Then Controls("LBTM" & z) = Replace(ayır(n), ",", ""): z = z + 1
'Next n
End Sub
Private Sub TBTKL0_Change()
On Error GoTo hata
Dim rd, a
Static q
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & Left(CBTKL1.Text, 3) & "\" & TBTKL0 & ".jpg")
If a = True Then
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(CBTKL1.Text, 3) & "\" & TBTKL0 & ".jpg")
q = 0
Else
If q = 0 Then Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(CBTKL1.Text, 3) & "\Logo.jpg")
q = 1
End If
Exit Sub
hata:
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
End Sub
Private Sub ListBoxMLZ1_Click()
On Error Resume Next
TBMLZ0 = ListBoxMLZ1.List(ListBoxMLZ1.listIndex, 0)
TBMLZ01 = ListBoxMLZ1.List(ListBoxMLZ1.listIndex, 1)
TBMLZF01 = ListBoxMLZ1.List(ListBoxMLZ1.listIndex, 3)
End Sub
Private Sub TBMLZ0_Change()
On Error GoTo hata
Dim rd, a
Static q
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & Left(CBMLZ1.Text, 3) & "\" & TBMLZ0 & ".jpg")
If a = True Then
Image2.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(CBMLZ1.Text, 3) & "\" & TBMLZ0 & ".jpg")
q = 0
Else
If q = 0 Then Image2.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(CBMLZ1.Text, 3) & "\Logo.jpg")
q = 1
End If
Exit Sub
hata:
Image2.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimage.jpg")
End Sub
Private Sub ListBoxMLZ1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton8_Click
End Sub
Private Sub ListBoxMLZ1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error Resume Next
If KeyAscii = 13 Then
CommandButton8_Click
End If
End Sub
Private Sub CommandButton8_Click()
On Error Resume Next
'Application.Calculation = xlCalculationManual
If ListBoxTKL1.List(ListBoxTKL1.listIndex, 0) = "" Or ListBoxMLZ1.List(ListBoxMLZ1.listIndex, 0) = "" Then Exit Sub
Call malzemedegisim
Call tkldosya: Call tkldosya2
'ListBoxP31_Click
'Application.Calculation = xlCalculationAutomatic
'If ListViewP31.ListItems.Count > 0 Then ListViewP31.ListItems(P31L).Selected = True: ListViewP31.ListItems(P31L + 1).EnsureVisible
If ListBoxTKL1.ListCount <= 0 Then Call tkldosyalar
If CDbl(CBTKL1AD) > CDbl(CBTKL1NO) Then ListBoxTKL1.listIndex = CDbl(CBTKL1NO) - 1 Else ListBoxTKL1.listIndex = CDbl(CBTKL1AD) - 1
End Sub
Sub malzemedegisim() '08.12.2013 malzeme değişimi için
On Error Resume Next
Dim teklifkod, mlzkod
Dim nameo
teklifkod = ListBoxTKL1.List(ListBoxTKL1.listIndex, 0)
mlzkod = ListBoxMLZ1.List(ListBoxMLZ1.listIndex, 0)
m = CDbl(ListBoxMLZ1.List(ListBoxMLZ1.listIndex, 2))
If teklifkod = "" Or mlzkod = "" Then Exit Sub
If mlzkod = teklifkod Then MsgBox ("Aynı Malzeme Kodları!  "), vbCritical, "Uyarı": Exit Sub

ProgressBarP21.Visible = True
s = Workbooks(dt).Worksheets("Sayfa1").Range("B65536").End(xlUp).row
z = 2
ProgressBarP21.Max = s
tekrar:
t = "B" & z & ":" & "B" & s
Set a = Workbooks(dt).Worksheets("Sayfa1").Range(t).Find(teklifkod, LookIn:=xlValues, LookAt:=xlWhole) 'LookAt:=xlWholeBİREBİR EŞLEME

If Not a Is Nothing Then
'*--verigir
Application.Calculation = xlCalculationManual
'Application.EnableEvents = False
Dim mliste, mlisteS3, Y
Set mliste = Workbooks((TBMLZ1)).Worksheets("Sayfa1")
Y = a.row
'Genel Biçimlemeler'--
Range("A" & Y & ":U" & Y).Borders.LineStyle = xlContinuous
Range("W" & Y & ":X" & Y).Borders.LineStyle = xlContinuous
Range("A" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Font.Bold = False
Range("A" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Font.ColorIndex = xlAutomatic
Range("A" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Font.Size = 9
Range("A" & Y & ":D" & Y).HorizontalAlignment = xlLeft
Range("E" & Y & ":U" & Y & ",W" & Y & ":X" & Y).HorizontalAlignment = xlRight
Range("A" & Y & ":D" & Y).NumberFormat = "@"
Range("E" & Y).NumberFormat = "#,##0"
Range("F" & Y).NumberFormat = "#,##0.00"
Range("G" & Y).NumberFormat = "0.0%"
Range("H" & Y).NumberFormat = "#,##0"
Range("J" & Y & ":X" & Y).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & Y & ",X" & Y).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & Y & ",X" & Y).NumberFormat = "#,##0.00 [$€-1]"
If Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & Y & ",X" & Y).NumberFormat = "#,##0.00 [$$-C0C]"
'GENEL VERİLER'--
'z = CDbl(TextBox17.Text)
Range("A" & Y) = mliste.Range("K" & m)
If ("A" & Y) = "" Then mliste.Range("K" & m + 1).End (3) 'Referans
Range("B" & Y) = mliste.Range("B" & m) 'Sipariş Kd.
Range("C" & Y) = mliste.Range("C" & m) 'Yapılacak İşin Cinsi
Range("D" & Y) = mliste.Range("D" & m) 'Üretici
Range("F" & Y) = mliste.Range("F" & m) 'Mlz. Br. Fiyat
Range("F" & Y).NumberFormat = mliste.Range("F" & m).NumberFormat
Range("G" & Y) = mliste.Range("G" & m) 'mlz.isk.
Range("H" & Y) = mliste.Range("H" & m) 'Adam/dk
If Not Left(Range("A" & Y), 3) = "PP-" Then mliste.Range ("I" & m) 'Boyut
'Range("I" & Y) = mliste.Range("I" & m) 'Boyut
'KUR'--
kur = ""
If Range("F" & Y).NumberFormat = "#,##0.00 [$$-C0C]" Then Range("F" & Y).Font.ColorIndex = 3: kur = "*Usd" ' Sayfa3 de $ kuru
If Range("F" & Y).NumberFormat = "#,##0.00 [$€-1]" Then Range("F" & Y).Font.ColorIndex = 5: kur = "*Eur" ' Sayfa3 de € kuru
'ÜRÜN GRUPLARI--
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
If Left(Range("A" & Y), 5) = "PM-MP" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Oisci/100" & mkur: GoTo devam1 'İŞÇ.
If Left(Range("A" & Y), 5) = "PM-MS" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Osarf/100" & mkur: GoTo devam1 'SARF
If Left(Range("A" & Y), 5) = "PM-MA" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Oamb/100" & mkur: GoTo devam1 'AMB.
If Left(Range("A" & Y), 5) = "PM-MN" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Onak/100" & mkur: GoTo devam1 'NAK.
If Left(Range("A" & Y), 5) = "PM-MB" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Obara/100" & mkur: GoTo devam1 'Bara
If Left(Range("A" & Y), 3) = "PP-" Then Range("L" & Y).FormulaR1C1 = bfyt & "*Opano/100" & mkur: GoTo devam1 'Pano
If Left(Range("A" & Y), 3) = "PS-" Then 'Pano sac & aksesuarlar
nameo = ActiveWorkbook.names("Opsac").RefersToR1C1
If Not nameo = Empty Then Range("L" & Y).FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else Range("L" & Y).FormulaR1C1 = bfyt & "*Opano/100" & mkur
GoTo devam1
End If
Range("L" & Y).FormulaR1C1 = bfyt & "*Osalt/100" & mkur 'Mlz.Kar+1
devam1:
'--
Range("J" & Y).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("N" & Y).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("K" & Y).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
Range("M" & Y).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("O" & Y).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & Y).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & Y).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & Y).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & Y).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("T" & Y).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.* işçilik katsayılı+1
Range("U" & Y).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
'TOPLAMLAR'--
Range("W" & Y).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb" 'Dövize göre Birim Fiyat TL+1
Range("X" & Y).FormulaR1C1 = "=RC[-19]*RC[-1]"  'Toplam Fiyat TL+1
'--
Cells(Y, "B").Font.ColorIndex = 7
z = a.row + 1
ProgressBarP21.Value = Y
GoTo tekrar
End If
ProgressBarP21.Visible = False
Application.Calculation = xlCalculationAutomatic
End Sub