Dim ara As String
Public tsaat As String
Public bfyt As String
Dim f, fd
Dim tkl
Dim dosya_yolu
Dim kur, mkur As String
Public fds
Sub kar_carpan() 'kar çarpanı uygula
On Error Resume Next
Application.Calculation = xlCalculationManual: Application.ScreenUpdating = False
Windows(dt).Activate
'--
ActiveWorkbook.names("CkarO").RefersToR1C1 = ComboBoxCRP1.Text 'kar çarpanı
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
Sheets("Sayfa1").Select '
son = Sheets("Sayfa1").Range("L65536").End(xlUp).row
Dim nameo
For Y = 2 To son
    If Sheets("Sayfa1").Cells(Y, "F") <> "" Then
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
     End If
Next Y
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub

Private Sub Frame5_Click()

End Sub

Private Sub Image7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open Labelyol & "\"
End Sub
Private Sub Labels1_Click()
On Error Resume Next
ds = Labels1
Workbooks.Open ds
dsi = ActiveWorkbook.Name
If CreateObject("Scripting.FileSystemObject").FileExists(ds) = True Then GoTo git:
If Err Then MsgBox ("    Dosya bulunamadı !  "), vbInformation, "scngnr@hotmail.com": Exit Sub
git:
Unload Me
msg = MsgBox("Açılan dosya: " & ds, vbOKCancel, "Dosya Açma İşlemi")
Windows(dsi).Activate
If msg = vbCancel Then Windows(dsi).Close False: Exit Sub
End Sub
Sub folders() '++
On Error Resume Next
Set df = CreateObject("Scripting.FileSystemObject")
fx = df.GetFolder(ActiveWorkbook.path)
If fx = Empty Then Exit Sub
ListBoxF1.AddItem fx
n = 1
Do Until Y > 10
fn = df.GetFolder(fx).ParentFolder
If fn = f Then Exit Do
ListBoxF1.AddItem fn
fx = fn
Y = Y + 1
Loop
z = ListBoxF1.ListCount - 1

Do Until z < 0
i = TreeView1.Nodes(ListBoxF1.List(z)).Index
a = ListBoxF1.List(z) & "\"
With TreeView1.Nodes
For Each kad In df.GetFolder(a).SubFolders
   If df.GetFolder(kad).SubFolders.Count > 0 Or dir(kad & "\*.xlsx", vbDirectory) <> "" Then Img = 1 Else Img = 2
   .Add i, tvwChild, kad, kad.Name, Img
   Next
fad = dir(a & "*.xlsx", vbDirectory)

Do While fad <> ""
If fad = ThisWorkbook.Name Then GoTo ResumeSub:
   .Add i, tvwChild, a & fad, fad, Image:=3
ResumeSub:
fad = dir
Loop
End With
TreeView1.Nodes(i).Expanded = True
z = z - 1
Loop
TreeView1.Nodes(ActiveWorkbook.fullName).Selected = True: TreeView1.SelectedItem.Image = 5
Labelyol = ActiveWorkbook.fullName
'..
End Sub
Sub klasorler1() '++
On Error Resume Next
If f = "" Then Exit Sub
Set df = CreateObject("Scripting.FileSystemObject")
For Each kad In df.GetFolder(f).SubFolders
a = a + 1
With TreeView1.Nodes
   If df.GetFolder(kad).SubFolders.Count > 0 Or dir(kad & "\*.xlsx", vbDirectory) <> "" Then Img = 1 Else Img = 2
.Add , a, kad, kad.Name, Img
b = a + c
'..
GoTo atla ' klasör uzun süren deneme*****2023
   For Each A1 In df.GetFolder(kad.path).SubFolders
   c = c + 1
   nn = A1 & "\"
   dsm = dir(A1)
   
   If df.GetFolder(A1).SubFolders.Count > 0 Or dir(A1 & "\*.xlsx", vbDirectory) <> "" Then Img = 1 Else Img = 2
   .Add b, tvwChild, A1, A1.Name, Img
   Next
atla:
'..
ks = kad & "\"
fad = dir(ks & "*.xlsx", vbDirectory)
Do While fad <> ""
If fad = ThisWorkbook.Name Then GoTo ResumeSub:
    c = c + 1
   .Add b, tvwChild, ks & fad, fad, Image:=3
ResumeSub:
fad = dir
Loop
'..
'For Each dosya In CreateObject("Scripting.FileSystemObject").GetFolder(klasoradi.Path).Files
'D = D + 1
'.Add "ana" & C, tvwChild, C & "ana_alt" & D, dosya.Name, Image:=2
'Next
End With
'D = 0
Next
End Sub
Private Sub Labelyol_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
CreateObject("Shell.Application").Open Labelyol & "\"
End Sub
Private Sub UserForm_Initialize() '++
On Error Resume Next
Call CreateCmdBar
Toolbar1.ImageList = ImageList3
Toolbar1.Buttons.Item(1).Image = ImageList3.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList3.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList3.ListImages.Item(6).Index
Toolbar1.Buttons.Item(4).Image = ImageList3.ListImages.Item(7).Index
Toolbar1.Buttons.Item(5).Image = ImageList3.ListImages.Item(9).Index
Toolbar1.Buttons.Item(6).Image = ImageList3.ListImages.Item(10).Index
Toolbar1.Buttons.Item(7).Image = ImageList3.ListImages.Item(12).Index
Toolbar1.Buttons.Item(8).Image = ImageList3.ListImages.Item(11).Index
PH = ActiveWorkbook.Worksheets("Sayfa3").Range("I55555")
If PH <> "Programı Hazırlayan: İlhan Şirin" Then: Toolbar1.Buttons.Item(2).Enabled = False
'ListView1.SmallIcons = ImageList3
TByol = ActiveWorkbook.path: f = GetSetting("ilhan", "Settings", "teklifdizini")
'TextBox1 = "C:\Belgelerim\Cemex\P"
TextBox1 = f
TBnoek1.Text = GetSetting("ilhan", "Settings", "tbnoek")
'..
ComboBoxP1.Clear: ComboBoxP2.Clear
ComboBoxP1.AddItem "Teklif Para Birimi (TL)"
ComboBoxP1.AddItem "Teklif Para Birimi (EUR)"
ComboBoxP1.AddItem "Teklif Para Birimi (USD)"
ComboBoxP2.AddItem "Takipte"
ComboBoxP2.AddItem "Alındı"
ComboBoxP2.AddItem "Alınamadı"
ComboBoxCRP1.AddItem "Net Fiyatı"
ComboBoxCRP1.AddItem "Liste Fiyatı"
'--
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'..
If TByol = "" Then
dsi = 0
Else
dsi = ActiveWorkbook.fullName
End If
'..
TBds = ActiveWorkbook.Name
If Not Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Then Toolbar1.Buttons.Item(5).Enabled = False
'If ActiveWorkbook.Name = TreeView1.SelectedItem Then copykaydetbuton
If ssno = 2 Then MultiPage1.Value = 1: Toolbar1.Buttons.Item(1).Enabled = False
'...............................................
If ssno = 1 Then
TreeView1.ImageList = ImageList1: Call klasorler1: Call folders
If Not dsi = 0 Then TreeView1.Nodes(dsi).Selected = True Else TreeView1.Nodes(1).Selected = True
'..
LB = 1
LB = TreeView1.SelectedItem.Index
TBds = TreeView1.SelectedItem.Text
'..
 If Not Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Then
 TBds = "": Labelyol = TreeView1.SelectedItem.key
 Else
 Set df = CreateObject("Scripting.FileSystemObject")
 Labelyol = f & "\" & df.GetParentFolderName(TreeView1.SelectedItem.FullPath)
 End If
End If
'...............................................
Labels1 = GetSetting("ilhan", "Settings", "sonteklif") 'ilave son dosya için
TreeView1.Width = 791
End Sub
Sub copykaydetbuton()
If Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Then
    If Toolbar1.Buttons.Item(5).Caption = "Kopyala" Then Toolbar1.Buttons.Item(5).Enabled = True
    If Toolbar1.Buttons.Item(5).Caption = "Farklı Kaydet" Then Toolbar1.Buttons.Item(5).Enabled = False
    'Toolbar1.Buttons.item(6).Enabled = False
Else
    If Toolbar1.Buttons.Item(5).Caption = "Kopyala" Then Toolbar1.Buttons.Item(5).Enabled = False
    If Toolbar1.Buttons.Item(5).Caption = "Farklı Kaydet" Then Toolbar1.Buttons.Item(5).Enabled = True
    If Right(ActiveWorkbook.Name, 5) = ".xlsx" Or Right(ActiveWorkbook.Name, 4) = ".xls" Then Exit Sub
    If Toolbar1.Buttons.Item(5).Enabled = True Then Toolbar1.Buttons.Item(6).Enabled = False Else Toolbar1.Buttons.Item(6).Enabled = True
    
End If
End Sub
Private Sub UserForm_Terminate()
    Call DestroyCmdBar
End Sub
Sub DestroyCmdBar()
    On Error Resume Next
    Application.CommandBars("flexgrid_rc").Delete
    On Error GoTo 0
End Sub
Sub klasorsec1()
On Error Resume Next
Yol = InputBox("Tekliflerin bulunduğu dizini yazınız. ", "Teklif Dizini", f)
If Yol = "" Then Exit Sub
If TextBox1 = Yol Then Exit Sub Else TextBox1 = Yol
f = Yol
SaveSetting "ilhan", "Settings", "teklifdizini", Yol
Unload Me
UFTH.Show
'TreeView1.Nodes.Clear
'klasorler1
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
ListBox1.Clear
i = TreeView1.SelectedItem.Index
fd = f & "\" & TreeView1.SelectedItem.FullPath
a = fd & "\"
With TreeView1.Nodes
ekey = .Item(i).Child.key
Set df = CreateObject("Scripting.FileSystemObject")

If Not ekey = Empty Then GoTo git1:
   For Each kad In df.GetFolder(a).SubFolders
   If df.GetFolder(kad).SubFolders.Count > 0 Or dir(kad & "\*.xlsx", vbDirectory) <> "" Then Img = 1 Else Img = 2
   .Add i, tvwChild, kad, kad.Name, Img
   Next
'..
fad = dir(a & "*.xlsx", vbDirectory)
Do While fad <> ""
If fad = ThisWorkbook.Name Then GoTo ResumeSub:
   .Add i, tvwChild, a & fad, fad, Image:=3
ResumeSub:
fad = dir
Loop
git1:
'alternatif..
   'For Each fad In CreateObject("Scripting.FileSystemObject").GetFolder(a).Files
   'If VBA.Right(fad.Name, 3) = "xls" Or VBA.Right(fad.Name, 4) = "xlsx" Or VBA.Right(fad.Name, 4) = "xlsm" Then
   '.Add i, tvwChild, fad, fad.Name, Image:=2
   'End If
   'Next
End With
'..
LB = TreeView1.SelectedItem.Index
TBds = TreeView1.SelectedItem.Text
'..
If Not Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Then
TBds = "": Labelyol = TreeView1.SelectedItem.key
Else
Labelyol = f & "\" & df.GetParentFolderName(TreeView1.SelectedItem.FullPath)
End If
'..
If TreeView1.Nodes(CDbl(LB)).Image = 5 Then Labelyol.ForeColor = &H459650 Else Labelyol.ForeColor = &H7379EC
'..
If ActiveWorkbook.Name = "" Then Call copykaydetbuton: Exit Sub
If ActiveWorkbook.Name = TreeView1.SelectedItem Then
'Veri_al
Else
'Call textBOS
'Call textdisabled
End If
Call copykaydetbuton
'..
'If Right(TreeView1.SelectedItem, 5) = ".xlsx" Or Right(TreeView1.SelectedItem, 4) = ".xls" Then Call Veri_al

End Sub
Sub Expanded()
    On Error Resume Next
For a = 1 To TreeView1.Nodes.Count
TreeView1.Nodes(a).Expanded = False
Next
End Sub
Private Sub TreeView1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
If TextBox1.Text = "" Then MsgBox (" Teklif klasörlerini seçiniz! "), vbInformation, "scngnr@hotmail.com": Exit Sub
'fd = f & "\" & TreeView1.SelectedItem.FullPath
''Set objFSO = CreateObject("scripting.filesystemobject")
''Set f = objFSO.getfolder()
''ad = f.parentfolder & "\" & UFTH.TreeView1.SelectedItem.Text
'..
'Lb = TreeView1.SelectedItem.index
'TreeView1.SelectedItem.Key = fd
'Labelyol = fd
'..
    Dim Item As MSComctlLib.listItem
    If (Button <> xlSecondaryButton) Or (Shift <> 0) Then Exit Sub
        'MsgBox "Item: " & TreeView1.SelectedItem.Key & " has been right-clicked!", vbInformation, "Capture Right-Click"
Application.CommandBars("flexgrid_rc").ShowPopup
    ''Just pass along the event's x and y arguments.
    'Set item = lvwTest.HitTest(x, y)

    'MsgBox "Item: " & item.ListSubItems(1) & " has been right-clicked!", vbInformation, "Capture Right-Click"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
dosyaveri.Enabled = False
FRTB.Enabled = True
Select Case Button.Index
Case 1
Toolbar1.Buttons.Item(3).Enabled = False
Toolbar1.Buttons.Item(4).Enabled = False
'Toolbar1.Buttons.item(5).Enabled = True
Toolbar1.Buttons.Item(6).Enabled = True
'Toolbar1.Buttons.item(7).Enabled = True
Toolbar1.Buttons.Item(5).Caption = "Kopyala"
If Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Then Toolbar1.Buttons.Item(5).Enabled = True

MultiPage1.Value = 0
'..
LB = TreeView1.SelectedItem.Index
TBds = TreeView1.SelectedItem.Text
'..
If Not Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Then
TBds = "": Labelyol = TreeView1.SelectedItem.key
Else
Labelyol = f & "\" & df.GetParentFolderName(TreeView1.SelectedItem.FullPath)
End If
'..
If TreeView1.Nodes(CDbl(LB)).Image = 5 Then Labelyol.ForeColor = &H459650 Else Labelyol.ForeColor = &H7379EC
TreeView1.Nodes(CDbl(LB)).EnsureVisible
TreeView1.Width = 791
Case 2
LB = TreeView1.SelectedItem.Index
Call toolbarbutton2
TBnoek1.Text = GetSetting("ilhan", "Settings", "tbnoek")
Case 3
Call uygula
'Call textdisabled
Case 4
Call textBOS
Case 5
MultiPage1.Value = 0
'If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then
If Toolbar1.Buttons.Item(5).Caption = "Farklı Kaydet" Then Call menu21: Exit Sub
msg = MsgBox("Kopyalanacak klasörü seçiniz", vbOKCancel, "Kopyalama İşlemi")
If msg = vbCancel Then Exit Sub
tkl = Labelyol & "\" & TBds
'Toolbar1.Buttons.item(5).Value = 1
Toolbar1.Buttons.Item(5).Caption = "Farklı Kaydet"
Toolbar1.Buttons.Item(5).Image = ImageList3.ListImages.Item(10).Index
'Else
'MsgBox ("    Teklif dosyası açmalısınız !  "), vbInformation, "scngnr@hotmail.com"
'End If

Case 6
TByol = ActiveWorkbook.path
If TByol.Value <> "" Then:  ActiveWorkbook.Save: MsgBox (ActiveWorkbook.Name & " dosyası kaydedildi."), , "Dosya Kaydetme İşlemi": Exit Sub
'--
If Labelyol = "" Then msg = MsgBox("Kopyalanacak klasörü seçiniz", vbInformation, "Kopyalama İşlemi"): Exit Sub
msg = MsgBox("Kaydetmek için seçilen klasör " & Labelyol & vbCrLf & _
"Farklı bir klasör seçmek için iptale basın.", vbOKCancel, "Kaydetme İşlemi : " & ActiveWorkbook.Name)
If msg = vbCancel Then Exit Sub
Call menu22: Exit Sub

Case 7
If ActiveWorkbook.Name <> TBds Then
Windows((TBds)).Close True
MsgBox (TBds & " kapatıldı! "), vbInformation, "scngnr@hotmail.com"
End If
Case 8
Unload Me
End Select
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Tag
Case 1
Call klasorsec1
Case 2
Set Klasor = CreateObject("Shell.Application").BrowseForFolder(0, "Lütfen bir klasör seçin !", &H100)
Yol = Klasor.items.Item.path
If Yol = "" Then Exit Sub
TextBox1 = Yol
f = Yol
SaveSetting "ilhan", "Settings", "teklifdizini", Yol
Unload Me
UFTH.Show
'TreeView1.Nodes.Clear
'klasorler1
End Select
End Sub
Sub toolbarbutton2()
    On Error GoTo hata
    Toolbar1.Buttons.Item(5).Enabled = False: Toolbar1.Buttons.Item(6).Enabled = False
    'Toolbar1.Buttons.item(7).Enabled = False
    '--
    If Not LB = 0 Then TBds = TreeView1.Nodes(CDbl(LB))
        If Not Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Or Right(TBds, 5) = ".xltx" Then
            TBds = ActiveWorkbook.Name: Labelyol = ActiveWorkbook.fullName
            Labelyol.ForeColor = &H459650
            GoTo atla1
hata:
            TBds = ActiveWorkbook.Name
atla1:
            dosya_yolu = ActiveWorkbook.path & "\[" & ActiveWorkbook.Name & "]"
            If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then _
            dosya_yolu = ""

            Call Veri_al
            MultiPage1.Value = 1
            Toolbar1.Buttons.Item(3).Enabled = True: Toolbar1.Buttons.Item(4).Enabled = True
            Exit Sub
        Else
            If ssno = 2 Then Labelyol = ActiveWorkbook.path
                fd = Labelyol & "\" & TBds
                Set ds = CreateObject("Scripting.FileSystemObject")
                Yol = ds.GetParentFolderName(fd)
                dosya = "\[" & TBds & "]"
                dosya_yolu = Yol & dosya
                
                'MsgBox Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R55555C9")
                
                On Error GoTo hata2
                i1 = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R55555C9") 'ilhan
                'If i1 = "Programı Hazırlayan: İlhan Şirin" Then
                If True Then
                    Call Veri_al
                    MultiPage1.Value = 1
                    Toolbar1.Buttons.Item(3).Enabled = True: Toolbar1.Buttons.Item(4).Enabled = True
                    Exit Sub
                    Else
                        MsgBox (" Bu bir teklif dosyası değils! "), vbInformation, "scngnr@hotmail.com"
                    Exit Sub
                    End If
hata2:

                MsgBox (" Bu bir teklif dosyası değilss! "), vbInformation, "scngnr@hotmail.com"
                End If
'--
'If ActiveWorkbook.Name = TBds Then
FRTB.Enabled = False
'dosyaveri.Enabled = True
End Sub
Sub toolbarbutton2xxx()
On Error GoTo hata
Toolbar1.Buttons.Item(5).Enabled = False: Toolbar1.Buttons.Item(6).Enabled = False
'Toolbar1.Buttons.item(7).Enabled = False
'--
If Not LB = 0 Then TBds = TreeView1.Nodes(CDbl(LB))
If Not Right(TBds, 5) = ".xlsx" Or Right(TBds, 4) = ".xls" Or Right(TBds, 5) = ".xltx" Then
TBds = ActiveWorkbook.Name: Labelyol = ActiveWorkbook.fullName
Labelyol.ForeColor = &H459650
'Lb = 1
'TBds = Labelyol
GoTo atla1
hata:
TBds = ActiveWorkbook.Name
atla1:
dosya_yolu = ActiveWorkbook.path & "\[" & ActiveWorkbook.Name & "]"
If ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then _

dosya_yolu = ""

Call Veri_al
MultiPage1.Value = 1
Toolbar1.Buttons.Item(3).Enabled = True
Toolbar1.Buttons.Item(4).Enabled = True
Exit Sub
Else
If ssno = 2 Then Labelyol = ActiveWorkbook.path
fd = Labelyol & "\" & TBds
Set ds = CreateObject("Scripting.FileSystemObject")
Yol = ds.GetParentFolderName(fd)
dosya = "\[" & TBds & "]"
dosya_yolu = Yol & dosya
On Error GoTo hata2
i1 = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R55555C9") 'ilhan
If i1 = "Programı Hazırlayan: İlhan Şirin" Then
Call Veri_al
MultiPage1.Value = 1
Toolbar1.Buttons.Item(3).Enabled = True
Toolbar1.Buttons.Item(4).Enabled = True
Exit Sub
Else
MsgBox (" Bu bir teklif dosyası değil2! "), vbInformation, "scngnr@hotmail.com"
Exit Sub
End If
hata2:
MsgBox (" Bu bir teklif dosyası değil2! "), vbInformation, "scngnr@hotmail.com"
End If
'--
'If ActiveWorkbook.Name = TBds Then
FRTB.Enabled = False
'dosyaveri.Enabled = True
End Sub
Sub menu22() ' 13.09.2013
On Error Resume Next
If TBtarih = Empty Or TBtarih = "-" Then TBtarih = Format(Now, "DD MMMM YYYY hh:mm:ss")
tsaat = Format(TBtarih, "DDMMYY-hhmmss")
dsa = "Yeni Teklif" & "-" & tsaat
ds = InputBox("Kaydedilecek klasör : " & Labelyol, "Dosya Adı", dsa)
If ds = "" Then Exit Sub
Yol = Labelyol
If Yol = "" Then Exit Sub
Dim s As Integer
s = TreeView1.Nodes(Yol).Index
'..
kl = Yol & "\" & ds & ".xlsx": kkl = ds & ".xlsx"
ActiveWorkbook.SaveAs kl
TreeView1.Nodes(kl).Image = 3
Call TreeView1.Nodes.Add(s, tvwChild, kl, kkl, Image:=3)
If ActiveWorkbook.Name = kkl Then TreeView1.Nodes(kl).Image = 5
TreeView1.Nodes(kl).Selected = True
Sheets("Sayfa3").Range("C8") = TBtarih 'Teklif Tarihi
End Sub
Sub menu21() ' 13.09.2013
On Error Resume Next
If TBtarih = Empty Or TBtarih = "-" Then TBtarih = Format(Now, "DD MMMM YYYY hh:mm:ss")
tsaat = Format(TBtarih, "DDMMYY-hhmmss")
dsa = "Yeni Teklif" & "-" & tsaat
ds = InputBox("Kaydedilecek klasör : " & UFTH.Labelyol, "Dosya Adı", dsa)
Yol = Labelyol
If ds = "" Then Toolbar1.Buttons.Item(5).Caption = "Kopyala": Exit Sub
If Yol = "" Then Exit Sub
Dim s As Integer
s = TreeView1.Nodes(Yol).Index
'..
kl = Yol & "\" & ds & ".xlsx": kkl = ds & ".xlsx"
'..
FileCopy tkl, kl
If Err.Description <> "" Then ActiveWorkbook.SaveAs kl
Call TreeView1.Nodes.Add(s, tvwChild, kl, kkl, Image:=3)
TreeView1.Nodes(tkl).Image = 3
If ActiveWorkbook.Name = kkl Then TreeView1.Nodes(kl).Image = 5
TreeView1.Nodes(kl).Selected = True
Toolbar1.Buttons.Item(5).Value = 0
Toolbar1.Buttons.Item(5).Caption = "Kopyala"
Toolbar1.Buttons.Item(5).Image = ImageList3.ListImages.Item(9).Index
'Sheets("Sayfa3").Range("C8") = TBtarih 'Teklif Tarihi
End Sub
Private Sub CommandButton1_Click()
    tsaat = Format(TBtarih, "DDMMYY-hhmmss")
    'Call teklifnoal 'Sercan güngör tarafından kapatıldı
    'Call zMrpi.MrpApi_Example_TeklifNextNumberDdMmYy ' Mrpden teklif numarası al
    MrpApi_Example_Configure
    TBno.Text = zMrpi.MrpApi_TeklifNextNumberDdMmYy ' Mrpden teklif numarası al
End Sub
Private Sub CommandButton83_Click()
TBtarih = Format(Now, "DD MMMM YYYY hh:mm:ss")
tsaat = Format(TBtarih, "DDMMYY-hhmmss")
Call teklifnoal
End Sub
Private Sub CommandButton3_Click()
SaveSetting "ilhan", "Settings", "tbnoek", TBnoek1.Text
End Sub
Private Sub CommandButton2_Click()
TBtarih = Format(Now, "DD MMMM YYYY hh:mm:ss")
tsaat = Format(TBtarih, "DDMMYY-hhmmss")
End Sub
Private Sub CommandButton9_Click()
SaveSetting "ilhan", "Settings", "TBveren", TBveren.Text
End Sub
Private Sub CommandButton10_Click()
TBveren.Text = GetSetting("ilhan", "Settings", "TBveren")
End Sub
Sub Veri_al() 'dosyadan veri al
On Error Resume Next
TBPAD.Text = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R3C3") 'Proje Adı
TBpadres.Text = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R4C3") 'Proje Kısa Adresi
TBisveren.Text = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R5C3") 'İşveren
'..
TBno = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R7C3") 'Teklifin Numarası
TBtarih = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R8C3") 'Teklifin İlk Veriliş Tarihi
Tadam.Text = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!Ads") 'adam/saat
TextBoxEURO.Text = Format(Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!Eur"), "#,####0.0000") 'Döviz kuru E
TextBoxUSD.Text = Format(Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!Usd"), "#,####0.0000") 'Döviz kuru $
ComboBoxP1.Value = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!Tpbr") 'Teklif Para Birimi
ComboBoxCRP1.Value = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!CkarO") 'kar çarpanı
'..
TBveren = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R10C3") 'Hazırlayan
ComboBoxP2.Value = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R10C6")  'Teklifin Durumu
TBilgili = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R11C3") 'Teklifle İlgili Kişi
TBtno = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R12C3") 'Telefon Numarası
TBtno2 = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R12C4") 'Telefon Numarası2
TBfax = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R12C6") 'Fax Numarası
TBemail = Application.ExecuteExcel4Macro("'" & dosya_yolu & "Sayfa3'!R13C3") 'E.mail Adresi
'..
If TBPAD.Text = "0" Then TBPAD.Text = "-"
If TBpadres.Text = "0" Then TBpadres.Text = "-"
If TBisveren.Text = "0" Then TBisveren.Text = "-"
If TBisvadres.Text = "0" Then TBisvadres.Text = "-"
If TBilgili.Text = "0" Then TBilgili.Text = "-"
If TBemail.Text = "0" Then TBemail.Text = "-"
If TBtno.Text = "0" Then TBtno.Text = "-"
If TBtno2.Text = "0" Then TBtno2.Text = "-"
If TBfax.Text = "0" Then TBfax.Text = "-"
If TBveren.Text = "0" Then TBveren.Text = "-"
If TBno.Text = "0" Then TBno.Text = "-"
If TBtarih.Text = "0" Then TBtarih.Text = "-"
If Tadam.Text = "-" Then TBtarih.Text = "0"
'..
'Toolbar1.Buttons.item(6).Enabled = True
Labelyol.Enabled = True
FRTB.Enabled = True
End Sub
Sub teklifnoal()
Dim saa
If Not TBnoek1.Text = "" Then GoTo atla1
If TBveren.Text = "" Or TBveren.Text = "-" Then TBveren.Text = Environ("USERNAME")
git:
On Error GoTo hata
b = Split(TBveren.Text, " ")
If UBound(b) > 1 Then
ab = b(0) & " " & b(1)
sa = b(2)
Else
ab = b(0)
saa = b(1)
End If
atla1:
If TBtarih = Empty Or TBtarih = "-" Then TBtarih = Format(Now, "DD MMMM YYYY hh:mm:ss")
'If TBno <> "" Then tsaat = Format(TBtarih, "DDMMYY") & "-" & Format(Now, "hhmmss")
If tsaat = Empty Then tsaat = Format(TBtarih, "DDMMYY-hhmmss")

If TBnoek1.Text = "" Then TBno = Left(ab, 1) & Left(saa, 1) & "-" & tsaat
If Not TBnoek1.Text = "" Then TBno = TBnoek1.Text & "-" & "XXXX" 'tsaat
Exit Sub
hata:
TBveren.Text = Environ("USERNAME") & " "
GoTo git:
End Sub
Sub teklifnoalXXX()
If TBveren.Text = "" Or TBveren.Text = "-" Then TBveren.Text = Environ("USERNAME")
git:
On Error GoTo hata
b = Split(TBveren.Text, " ")
If UBound(b) > 1 Then
ab = b(0) & " " & b(1)
sa = b(2)
Else
ab = b(0)
sa = b(1)
End If
If tsaat = Empty Then tsaat = Format(Now, "DDMMYY") & "-" & Format(Now, "hhmmss")
TBno = Left(ab, 1) & Left(sa, 1) & "-" & tsaat
Exit Sub
hata:
TBveren.Text = Environ("USERNAME") & " "
GoTo git:
'TextBox1.Text = Environ("USERNAME") & "-" & Format(Now, "DDMMYY") & "-" & Format(Now, "hhmm")
'TextBox1.Text = Application.UserName & Format(Now, "ddmmyy")
'TextBox1.Text = Environ("USERNAME") & Format(Now, "ddmmyy")
End Sub
Sub textBOS()
TBPAD.Text = "-": TBpadres.Text = "-": TBisveren.Text = "-": TBisvadres.Text = "-" 'Proje Adı'Proje Kısa Adresi'İşveren
TBilgili.Text = "-": TBtno.Text = "-": TBfax.Text = "-" 'T.İlgili Kişi'Tel.No'Fax No
TBemail.Text = "-": TBtarih.Text = "-" 'E.mail Adresi:'T.No'T.İlk Veriliş Tarihi
TBveren.Text = "-": 'TBno.Text = "-": Tadam.Text = "-" 'Hazırlayan'adam/saat
'TextBoxUSD.Text = "-": TextBoxEURO.Text = "-" 'Döviz kuru $'Döviz kuru €
'ComboBoxP1.Value = "-": ComboBoxP2.Value = "-" 'Teklif Para Birimi'Teklifin Durumu
End Sub
Sub uygula()
On Error Resume Next
Set ds = CreateObject("Scripting.FileSystemObject")
ds = TBds
If Not WorkbookOpen((ds)) Then
    Application.ScreenUpdating = False
    Workbooks.Open (Labelyol & "\" & ds)
    Call uygula_bilgiler
    Call kar_carpan
Workbooks((ds)).Close True
Application.ScreenUpdating = True
Exit Sub
End If
Call uygula_bilgiler
Call kar_carpan
End Sub
Sub p1()
Index = ComboBoxP1.listIndex
Select Case Index
Case 0: Call MakroTL
Case 1: Call MakroEURO
Case 2: Call MakroDOLAR
End Select
End Sub
Sub uygula_bilgiler()
ds = TBds
Set s3 = Workbooks(ds).Sheets("Sayfa3")
s3.Range("C3") = TBPAD.Text 'Proje Adı
s3.Range("C4") = TBpadres.Text 'Proje Kısa Adresi
s3.Range("C5") = TBisveren.Text 'İşveren
s3.Range("E5") = TBisvadres.Text 'İşveren adres
s3.Range("C7") = TBno.Text  'Teklifin Numarası
s3.Range("C8") = TBtarih.Text  'Teklifin İlk Veriliş Tarihi
s3.Range("C10") = TBveren.Text  'Hazırlayan
s3.Range("C11") = TBilgili.Text  'Teklifle İlgili Kişi
s3.Range("C12") = TBtno.Text  'Telefon Numarası
s3.Range("D12") = TBtno2.Text  'Telefon Numarası2
s3.Range("F12") = TBfax.Text  'Fax Numarası
s3.Range("C13") = TBemail.Text  'E.mail Adresi
s3.Range("F10") = ComboBoxP2.Value  'Teklifin Durumu

s3.Range("Ads") = Tadam.Text  'adam/saat
s3.Range("Tpbr") = ComboBoxP1.Text 'Teklif Para Birimi
s3.Range("Usd").Value = CDbl(TextBoxUSD.Value) 'Döviz kuru $
s3.Range("Eur").Value = CDbl(TextBoxEURO.Value)  'Döviz kuru €
'--
Workbooks(ds).names("CkarO").RefersToR1C1 = ComboBoxCRP1.Text 'kar çarpanı
If Workbooks(ds).names("CkarO").RefersToR1C1 = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
Call p1
End Sub
Private Sub TextBoxEURO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
    Case 46
        'noktayı virgül ile değiştir.
        ' 44 virgül'ün, 46 nokta'nın ASCII kodu
        KeyAscii = 44
    Case 44, 48 To 57
        'basılan tuş virgül veya sayıysa
        'Tuş kodunda bir değişiklik yapma
    Case Else
        'Diğer her tuş basımını iptal et
        KeyAscii = 0
    End Select
End Sub
Private Sub TextBoxUSD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
    Case 46
        'noktayı virgül ile değiştir.
        ' 44 virgül'ün, 46 nokta'nın ASCII kodu
        KeyAscii = 44
    Case 44, 48 To 57
        'basılan tuş virgül veya sayıysa
        'Tuş kodunda bir değişiklik yapma
    Case Else
        'Diğer her tuş basımını iptal et
        KeyAscii = 0
    End Select
End Sub
Sub dosyaac()
On Error GoTo hata
ds = ActiveWorkbook.Name
Dim td
td = f & "\" & TreeView1.SelectedItem.FullPath
If Not Right(td, 5) = ".xlsx" Or Right(td, 4) = ".xls" Then GoTo hata
msg = MsgBox("Açılan dosya: " & td, vbOKCancel, "Dosya Açma İşlemi")
If msg = vbCancel Then: Exit Sub
'If Msg = vbCancel Then Windows(tdi).Close False ': Exit Sub
Application.ScreenUpdating = False
Unload UFTH
'If TreeView1.SelectedItem.Sorted = True Then TreeView1.SelectedItem.Image = 3 Else TreeView1.SelectedItem.Image = 2
Workbooks.Open td
tdi = ActiveWorkbook.Name
TByol = ActiveWorkbook.path
SaveSetting "ilhan", "Settings", "sonteklif", td
'Workbooks(td).Activate
'Workbooks(ds).Activate
'TreeView1.SelectedItem.BackColor = &H80C0FF
UFTH.Show
'Unload UFTH
hata:
Application.ScreenUpdating = True
'Msg = MsgBox("Açılan dosya: " & td, vbOKOnly, "Dosya Açma İşlemi")
End Sub
Private Sub TreeView1_DblClick()
On Error GoTo hata1
If TreeView1.SelectedItem.Children > 0 Then
'klasör resimleri
If TreeView1.SelectedItem.Expanded = False Then TreeView1.SelectedItem.Image = 1 Else TreeView1.SelectedItem.Image = 4
'--
Exit Sub
Else
ds = TreeView1.SelectedItem.Text
If Not WorkbookOpen((ds)) Then
Call dosyaac
dt = ds
Else
If ActiveWorkbook.Name = ds Then Workbooks(ds).Activate: TreeView1.SelectedItem.Image = 5: Exit Sub
Unload UFTH
Workbooks(ds).Activate
UFTH.Show
End If
Exit Sub
End If
hata1:
End Sub
Private Sub CBenparaxxx_Click()
On Error GoTo hata
Set IE = CreateObject("InternetExplorer.Application")
IE.navigate "https://www.qnbfinansbank.enpara.com/hesaplar/doviz-ve-altin-kurlari"
Do: DoEvents: Loop Until IE.ReadyState = 4
LSTKUR.ColumnHeaders.Clear
LSTKUR.ListItems.Clear
LSTKUR.View = lvwReport
LSTKUR.OLEDragMode = ccOLEDragAutomatic
LSTKUR.FullRowSelect = True

Call LSTKUR.ColumnHeaders.Add(1, , "Döviz ve altın", 100)
Call LSTKUR.ColumnHeaders.Add(2, , "Enpara.com alış", 100, 1)
Call LSTKUR.ColumnHeaders.Add(3, , "Enpara.com satış", 100, 1)
i = 1
d = 0
    Do Until 5 = i
Call LSTKUR.ListItems.Add((i), , IE.document.getElementsByClassName("flex")(d).innerText)
Call LSTKUR.ListItems(i).ListSubItems.Add(1, , IE.document.getElementsByClassName("flex")(d + 1).innerText)
Call LSTKUR.ListItems(i).ListSubItems.Add(2, , IE.document.getElementsByClassName("flex")(d + 2).innerText)
    d = d + 3
    i = i + 1
    Loop
hata:
IE.Quit
End Sub
Private Sub CommandButtontcmb_Click()
Shell ("C:\Program files\Internet Explorer\Iexplore.exe http://www.tcmb.gov.tr/wps/wcm/connect/tr/tcmb+tr/main+page+site+area/bugun"), vbNormalFocus
End Sub
Private Sub CBenpara_Click()
On Error GoTo HataYakala

LSTKUR.ListItems.Clear

Dim path As String
Dim icerik As String, xmlhttp As Object
Dim evn1, evn2, evn3, evn4 As Variant
Dim i As Integer, Y As Long
Dim Dovtip As String

' XMLHTTP nesnesini oluştur
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
If xmlhttp Is Nothing Then
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    If xmlhttp Is Nothing Then
        MsgBox "XMLHTTP nesnesi oluşturulamadı!", vbCritical
        Exit Sub
    End If
End If

Application.Volatile
path = "https://www.tcmb.gov.tr/kurlar/today.xml" ' Yeni TCMB URL'si
xmlhttp.Open "GET", path, False
xmlhttp.send

If xmlhttp.Status = 200 Then
    icerik = xmlhttp.responseText
    temizlik = Split(icerik, "<Currency CrossOrder=")

    For i = 1 To 2 ' Hem USD hem de EUR için

        Dovtip = IIf(i = 1, "USD", "EUR") ' Döviz tipini ayarla
        Dovtip = UCase(Dovtip)

        For Y = 0 To UBound(temizlik)
            If temizlik(Y) Like "*=""" & Dovtip & "*" Then
                sonuclar = Split(temizlik(Y), "</CurrencyName>")
                evn1 = Split(sonuclar(1), "<ForexBuying>"): evn2 = Split(sonuclar(1), "<ForexSelling>")
                evn3 = Split(sonuclar(1), "<BanknoteBuying>"): evn4 = Split(sonuclar(1), "<BanknoteSelling>")

                'sonuçlar
                evn11 = Split(evn1(1), "</"): evn21 = Split(evn2(1), "</")
                evn31 = Split(evn3(1), "</"): evn41 = Split(evn4(1), "</")

                ' LSTKUR'a ekle
                Call LSTKUR.ListItems.Add(i, , Dovtip)
                Call LSTKUR.ListItems(i).ListSubItems.Add(1, , Replace(evn11(0), ".", ","))
                Call LSTKUR.ListItems(i).ListSubItems.Add(2, , Replace(evn21(0), ".", ","))
                Call LSTKUR.ListItems(i).ListSubItems.Add(3, , Replace(evn31(0), ".", ","))
                Call LSTKUR.ListItems(i).ListSubItems.Add(4, , Replace(evn41(0), ".", ","))
                Exit For ' Döviz tipi bulundu, iç döngüden çık
            End If
        Next Y

    Next i ' Sonraki döviz tipi
Else
    MsgBox "XML dosyası alınamadı. HTTP Durumu: " & xmlhttp.Status, vbCritical
End If

Exit Sub ' Alt yordamdan normal çıkış

HataYakala:
    MsgBox "Hata oluştu: " & Err.Description & " (Hata Kodu: " & Err.Number & ")", vbCritical
End Sub
Private Sub CommandButton82_Click()
If LSTKUR.ListItems.Count < 1 Then Exit Sub
TextBoxUSD.Text = LSTKUR.ListItems(1).ListSubItems(2)
TextBoxEURO.Text = LSTKUR.ListItems(2).ListSubItems(2)
End Sub
Private Sub CommandButton8_Click()
On Error Resume Next
    MrpApi_Example_Configure
    Call zMrpi.MrpApi_ImportCompaniesContactsToWorkbook
Application.ScreenUpdating = False
If Not WorkbookOpen("Teklif Firma Bilgileri.xlsb") Then
    Dim ds, a
    Set ds = CreateObject("Scripting.FileSystemObject")
    a = ds.FileExists("C:\Belgelerim\Cemex\Parametreler\Teklif Firma Bilgileri.xlsb")
    If a = False Then Exit Sub
    Workbooks.Open "C:\Belgelerim\Cemex\Parametreler\Teklif Firma Bilgileri.xlsb"
End If
Application.Windows("Teklif Firma Bilgileri.xlsb").Visible = False
    fds = "Teklif Firma Bilgileri.xlsb"
Application.ScreenUpdating = True
ssno = 0
UFFirma.Show
End Sub