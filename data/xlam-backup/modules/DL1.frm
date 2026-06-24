Private Sub CommandButton2_Click()
SaveSetting "ilhan", "Settings", "malzemedizini", TextBox1
End Sub
Private Sub CommandButton3_Click()
If CBB1.Value = True Then
Const OverWriteFiles = True
Set fso = CreateObject("Scripting.FileSystemObject")
cemexds = TextBox1 & "\Malzeme Listeleri"
If Not fso.FolderExists(cemexds) Then Msg1 = "Malzeme listeleri mevcut değil!": GoTo git1
fso.CopyFolder cemexds, fm1 & "\Malzeme Listeleri", OverWriteFiles
Set fso = Nothing
Msg1 = "Malzeme listeleri aktarıldı."
End If
git1:
msg = Msg1
If Not msg = "" Then MsgBox msg, vbInformation
End Sub
Private Sub Image7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open TextBox1.Text
End Sub
Private Sub Label6_Click()
Workbooks.Open fileName:=ListViewp11.SelectedItem.key
Application.Windows((ListViewp11.SelectedItem)).Visible = False
Windows((ListViewp11.SelectedItem)).Close True
End Sub
Private Sub ListViewP11_Click()
On Error Resume Next
TextBox1 = ListViewp11.SelectedItem
End Sub
Private Sub Label7_Click()
Call ListViewP11_dblClick
End Sub
Private Sub ListViewP11_dblClick()
On Error Resume Next
td = ActiveWorkbook.Name
ds = ListViewp11.SelectedItem
msg = MsgBox("Açılan dosya: " & ds, vbOKCancel, "Dosya Açma İşlemi")
If msg = vbCancel Then Exit Sub
Workbooks.Open fileName:=ListViewp11.SelectedItem.key
Application.Windows(ds).Visible = True
Unload Me
Windows(td).ActivateNext
If Not ListViewp11.Tag = "" Then
Unload UFDD
UFDD.Show
UFDD.T3.Text = ds: UFDD.CA3 = "B": UFDD.CA4 = "F"
End If
End Sub
Private Sub UserForm_Initialize()
ControlTipText = "Hazırlayan: İlhan Şirin"
'---
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
TextBoxM1 = fm1
'---
ListViewp11.ListItems.Clear
'ListViewp11.Icons = ImageList1
ListViewp11.SmallIcons = ImageList1
'ListViewP11.Arrange = lvwAutoLeft
Dim f
Dim itmX As listItem
'--
f = fm1 & "\Malzeme Listeleri\1\"
dosya = dir(f & "*.xlsb")
Do While dosya <> ""
Set itmX = ListViewp11.ListItems.Add(, f & dosya, dosya, , 1)
itmX.Bold = True
   dosya = dir
Loop
'--
f = fm1 & "\Malzeme Listeleri\2\"
dosya = dir(f & "*.xlsb")
Do While dosya <> ""
Set itmX = ListViewp11.ListItems.Add(, f & dosya, dosya, , 2)
   dosya = dir
Loop
'--
f = fm1 & "\Malzeme Listeleri\3\"
dosya = dir(f & "*.xlsb")
Do While dosya <> ""
Set itmX = ListViewp11.ListItems.Add(, f & dosya, dosya, , 2)
   dosya = dir
Loop
'--
f = fm1 & "\Malzeme Listeleri\4\"
dosya = dir(f & "*.xlsb")
Do While dosya <> ""
Set itmX = ListViewp11.ListItems.Add(, f & dosya, dosya, , 2)
   dosya = dir
Loop
'--
f = fm1 & "\Otomatik Seçim\"
dosya = dir(f & "*.xlsb")
Do While dosya <> ""
Set itmX = ListViewp11.ListItems.Add(, f & dosya, dosya, , 3)
   dosya = dir
Loop
'--
ListViewp11.StartLabelEdit
ListViewp11.Refresh
TextBox1 = GetSetting("ilhan", "Settings", "malzemedizini")
LBS1.Caption = " Mevcut Listeler " & ListViewp11.ListItems.Count & " Adet"
End Sub
Private Sub Label1_Click()
On Error Resume Next
CreateObject("Shell.Application").Open fm1 & "\Malzeme Listeleri\1\"
End Sub
Private Sub Label2_Click()
On Error Resume Next
CreateObject("Shell.Application").Open fm1 & "\Malzeme Listeleri\2\"
End Sub
Private Sub Label3_Click()
On Error Resume Next
CreateObject("Shell.Application").Open fm1 & "\Malzeme Listeleri\3\"
End Sub
Private Sub Label4_Click()
On Error Resume Next
CreateObject("Shell.Application").Open fm1 & "\Malzeme Listeleri\4\"
End Sub
Private Sub Label5_Click()
On Error Resume Next
CreateObject("Shell.Application").Open fm1 & "\Otomatik Seçim\"
End Sub
Private Sub Labelm1_Click()
On Error Resume Next
Call klasorsecm11
End Sub
Sub klasorsecm11()
On Error Resume Next
yolm1 = InputBox("Fiyat listelerinin bulunduğu dizini yazınız. ", "Fiyat Listeleri Dizini", "C:\Belgelerim\Cemex")
If yolm1 = "" Or yolm1 = TextBoxM1 Then Exit Sub
TextBoxM1 = yolm1: fm1 = yolm1
SaveSetting "ilhan", "Settings", "malzemedizini", yolm1
End Sub