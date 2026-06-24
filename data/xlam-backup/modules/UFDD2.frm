Private Sub CA12_Change()
If CheckBoxsa.Value = False Then Range(CA12 & 1).Select
If CA11 = CA12 Then CA12.BackColor = &HC0C0FF Else CA12.BackColor = &HDBE8DB
End Sub
Private Sub CommandButtonM1_Click()
On Error Resume Next
If Toolbar2.Tag = 2 Then
Dim ds, a
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists(TextBoxM1)
If a <> True Then msg = MsgBox("Dosya mevcut değil!", vb, "scngnr@hotmail.com"): Exit Sub
End If
CATS1 = 2
CATS2 = Workbooks(T3.Value).ActiveSheet.Range(CA11.Text & "65536").End(xlUp).row
'--
If CheckBoxsa.Value = True Then
CATS1 = ActiveWindow.RangeSelection.row
'If CATS1 = 1 Then CATS1 = 2
CATS2 = CATS1 + Selection.Cells.Count - 1
If CATS2 > 30000 Then CATS2 = Workbooks(T3.Value).ActiveSheet.Range(CA11.Text & "65536").End(xlUp).row + 1
End If
'--
If CATS2 - CATS1 < 0 Then Exit Sub

ProgressBar1.Value = 1: ProgressBar1.Max = CATS2
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
dd = LBS1.Caption
If Not Toolbar2.Tag = 1 Then Application.Workbooks.Open (TextBoxM1): Application.Windows(dd).Visible = False
Dim dmt As Workbook
Set dmt = Workbooks(dd)
If dmt Is Nothing Then
            MsgBox "Dosya erişim hatası!", vbInformation, "scngnr@hotmail.com"
            Set wBook = Nothing
            On Error GoTo 0: Exit Sub
End If
Set DTM = dmt.ActiveSheet
'-------
Workbooks(T3.Value).Activate
C1 = CA11.listIndex + 1: C2 = CA12.listIndex + 1: C3 = CA21.listIndex + 1: C4 = CA22.listIndex
If CKR2.Value = True Then
For n = CATS1 To CATS2
    If ActiveSheet.Range(CA11 & n).Value <> "" Then
    If Range(CA11 & n) = "BÖLÜM TOPLAMI:" Or Range(CA11 & n) = "BÖLÜM ADI/NO:" Or Range(CA11 & n) = "GENEL TOPLAM:" Then GoTo atla1
    If Range(CA12 & n).Value > 0 Then GoTo atla1
    tms = WorksheetFunction.SumIf(DTM.Range(CA21 & 2 & ":" & CA21 & 65536), Range(CA11 & n), DTM.Range(CA22 & 2 & ":" & CA22 & 65536))
    Range(CA12 & n).Value = tms
    If Range(CA12 & n).Value > 0 Then Range(CA12 & n).Font.ColorIndex = 5 Else Range(CA12 & n).Font.ColorIndex = 3
    End If
atla1:
ProgressBar1.Value = n
Next n
GoTo atla3
End If
For n = CATS1 To CATS2
    If ActiveSheet.Range(CA11 & n).Value <> "" Then
    If Range(CA11 & n) = "BÖLÜM TOPLAMI:" Or Range(CA11 & n) = "BÖLÜM ADI/NO:" Or Range(CA11 & n) = "GENEL TOPLAM:" Then GoTo atla2
    tms = WorksheetFunction.SumIf(DTM.Range(CA21 & 2 & ":" & CA21 & 65536), Range(CA11 & n), DTM.Range(CA22 & 2 & ":" & CA22 & 65536))
    If CKR1.Value = True Then Range(CA12 & n).Value = Range(CA12 & n).Value + tms Else Range(CA12 & n).Value = tms
'formülle
    'fm = "=SUMIF('[" & dd & "]Sayfa1'!C" & C & C3 & ",RC[" & C1 - C2 & "],'[" & dd & "]Sayfa1'!C" & C4 & ")"
    'Range(CA12 & n).FormulaR1C1 = fm : Range(CA12 & n) = Range(CA12 & n)
'-------
tms = Empty
    If Range(CA12 & n).Value > 0 Then Range(CA12 & n).Font.ColorIndex = 5 Else Range(CA12 & n).Font.ColorIndex = 3
    End If
atla2:
ProgressBar1.Value = n
Next n
atla3:
ProgressBar1.Value = 1
'-------
If Toolbar2.Tag = 2 Then Windows(dd).Close False
 BB = Empty
'Application.Windows(dd).Visible = True
dd = Empty: fm = Empty
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Private Sub CommandButtonM11_Click()
Unload Me
End Sub
Private Sub LBR2_Click()
If CA12.listIndex > 0 Then CA12 = CA12.List(CA12.listIndex - 1)
End Sub
Private Sub Label1_Click()
On Error Resume Next
Workbooks.Open fileName:=TextBoxM1.Text
If CreateObject("Scripting.FileSystemObject").FileExists(TextBoxM1.Text) = True Then Exit Sub
If Err Then MsgBox ("    Dosya bulunamadı !  "), vbInformation, "scngnr@hotmail.com"
End Sub
Private Sub Labelm1_Click()
On Error Resume Next
Call klasorsecm11
End Sub
Private Sub Labelm2_Click()
On Error Resume Next
If LBS1 = "-" Or LBS1 = "" Then Exit Sub Else SaveSetting "ilhan", "Settings", "deposabitdosya", LBS1
End Sub
Sub klasorsecm11()
On Error Resume Next
yold = GetSetting("ilhan", "Settings", "depodizini")
Yol = InputBox("Dosyaların bulunduğu dizini yazınız. ", "Stok Dosyaları Dizini", yold)
If Yol = "" Or Yol = TextBoxM1 Then Exit Sub
SaveSetting "ilhan", "Settings", "depodizini", Yol
Labelyol = Yol: TextBoxM1 = Yol
End Sub
Private Sub ListBoxDS_Click()
On Error Resume Next
If ListBoxDS.listIndex < 0 Then Exit Sub
LBS1 = ListBoxDS.Text
If Toolbar2.Tag = 2 Then
TextBoxM1 = Labelyol & "\" & LBS1
Else
TextBoxM1 = Workbooks((LBS1)).path
End If
End Sub
Private Sub SpinButton2_SpinUp()
On Error GoTo hata
If CA12.listIndex < CA12.ListCount Then CA12 = CA12.List(CA12.listIndex + 1)
hata:
End Sub
Private Sub SpinButton2_SpinDown()
On Error GoTo hata
If CA12.listIndex > 0 Then CA12 = CA12.List(CA12.listIndex - 1)
hata:
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
If Range("B" & "65536").End(xlUp).row < 2 And Range("A" & "65536").End(xlUp).row < 2 Then MsgBox (" Dosya boş."), vbInformation, "scngnr@hotmail.com": Exit Sub
T3.Value = ActiveWorkbook.Name
Workbooks(T3.Value).Activate
MultiPage1.Value = 0
Case 2
Dim ds, a
'''
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists(TextBoxM1)
If a <> True Then msg = MsgBox("Dosya mevcut değil!", vb, "scngnr@hotmail.com"): Exit Sub
n = Selection.row
If ActiveSheet.Range(CA11 & n).Value <> "" Then
If Range(CA11 & n) = "BÖLÜM TOPLAMI:" Or Range(CA11 & n) = "BÖLÜM ADI/NO:" Or Range(CA11 & n) = "GENEL TOPLAM:" Then Exit Sub
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
dd = LBS1.Caption
Application.Workbooks.Open (TextBoxM1): Application.Windows(dd).Visible = False
Set DTM = Workbooks(dd).ActiveSheet
    tms = WorksheetFunction.SumIf(DTM.Range(CA21 & 2 & ":" & CA21 & 65536), Range(CA11 & n), DTM.Range(CA22 & 2 & ":" & CA22 & 65536))
    MsgBox (" Bakılan Dosya Adı : " & dd & vbLf & " Bu üründen  " & tms & "    Ad. mevcut  "), vbInformation, "scngnr@hotmail.com"
If Toolbar2.Tag = 2 Then Windows(dd).Close False
DTM = Empty
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End If
'MultiPage1.Value = 1
End Select
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1
ListBoxDS.Clear
Dim wb As Workbook
For Each wb In Workbooks
If wb.Name <> T3.Value Then ListBoxDS.AddItem wb.Name
Next wb
Toolbar2.Tag = 1: TextBoxM1 = "": LBS1 = "": Yol = "":
ListBoxDS.Selected(0) = True
Case 2
ListBoxDS.Clear
Toolbar2.Tag = 2
dmsabit = GetSetting("ilhan", "Settings", "deposabitdosya")
If dmsabit <> "" Then ListBoxDS.AddItem GetSetting("ilhan", "Settings", "deposabitdosya"): ListBoxDS.Selected(0) = True
Case 3
Toolbar2.Tag = 2: TextBoxM1 = Labelyol: LBS1 = ""
ListBoxDS.Clear
  Dim dosya
  dosya = dir(Labelyol & "\*.xls*")
Do While dosya <> ""
ListBoxDS.AddItem dosya
   dosya = dir
Loop
If ListBoxDS.ListCount > 0 Then ListBoxDS.Selected(0) = True
End Select
End Sub
Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Tag
Case 1
If LBS1 = "-" Or LBS1 = "" Then Exit Sub Else SaveSetting "ilhan", "Settings", "deposabitdosya", LBS1
Case 2
On Error Resume Next
Set Klasor = CreateObject("Shell.Application").BrowseForFolder(0, "Lütfen bir klasör seçin !", &H100)
Yol = Klasor.items.Item.path
If Yol = "" Then Exit Sub
Labelyol = Yol: TextBoxM1 = Yol
SaveSetting "ilhan", "Settings", "depodizini", Yol
ListBoxDS.Clear
  Dim dosya
  dosya = dir(Yol & "\*.xls*")
Do While dosya <> ""
ListBoxDS.AddItem dosya
   dosya = dir
Loop
'If ListBoxDS.ListCount > 0 Then ListBoxDS.Selected(0) = True
End Select
End Sub
Private Sub UserForm_Initialize()
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(6).Index
Toolbar2.ImageList = ImageList1
Toolbar2.Buttons.Item(1).Image = ImageList1.ListImages.Item(2).Index
Toolbar2.Buttons.Item(2).Image = ImageList1.ListImages.Item(3).Index
Toolbar2.Buttons.Item(3).Image = ImageList1.ListImages.Item(1).Index
Labelyol = GetSetting("ilhan", "Settings", "depodizini")
TextBoxM1 = Labelyol
dmsabit = GetSetting("ilhan", "Settings", "deposabitdosya")
If dmsabit <> "" Then ListBoxDS.AddItem GetSetting("ilhan", "Settings", "deposabitdosya"): ListBoxDS.Selected(0) = True

'TextBoxm1 = "C:\Users\N550\Desktop\Depo\Z Depo Stok02.01.2023.xlsb"
T3 = ActiveWorkbook.Name
'Dim filePath As String
'filePath = TextBoxm1: LBS1.Caption = Dir(filePath)
'-------
dizi1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
For i = 1 To 26
        CA11.AddItem Mid(dizi1, i, 1): CA12.AddItem Mid(dizi1, i, 1): CA21.AddItem Mid(dizi1, i, 1): CA22.AddItem Mid(dizi1, i, 1)
Next
dizi1 = ""
dizi2 = "AAABACADAEAFAGAHAIAJAKALAMANAOAPAQAUAVAWARASATAXAYAZBABBBCBDBEBFBGBHBIBJBKBLBMBNBOBPBQBUBVBWBRBSBTBXBYBZCACBCCCDCECFCGCHCICJCKCLCMCNCOCPCQCUCVCWCRCSCTCXCYCZ"
m = 1
For i = 1 To 78
        CA11.AddItem Mid(dizi2, m, 2): CA12.AddItem Mid(dizi2, m, 2): CA21.AddItem Mid(dizi2, m, 2): CA22.AddItem Mid(dizi2, m, 2)
        m = m + 2
Next
dizi2 = ""
'-------
CA11.Text = "B": CA21.Text = "B": CA22.Text = "E"
sk = Cells(1, 256).End(xlToLeft).Column ' kolon sayısını verir
Cells(1, sk + 1).Select
CA12.Text = Replace(ActiveCell.Address(0, 0), ActiveCell.row, "")
End Sub