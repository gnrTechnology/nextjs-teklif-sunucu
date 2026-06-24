Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub CommandButton31_Click() 'ceviri dosyaları
On Error Resume Next
ListViewP21.ListItems.Clear: ListViewP31.ListItems.Clear: TextBoxP21 = ""
Dim itmX As listItem
  Dim dosya
  dosya = dir("C:\Belgelerim\CEMEX\Çeviri Dosyaları\Kelime Dosyaları\*.xls*")
Do While dosya <> ""
Set itmX = ListViewP21.ListItems.Add(, , dosya)
itmX.Bold = True
   dosya = dir
Loop
ListViewP21.StartLabelEdit: ListViewP21.Refresh
End Sub

Private Sub Frame41_Click()

End Sub

Private Sub Image2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Çeviri Dosyaları\Kelime Dosyaları"
End Sub
Private Sub ListViewp21_Click()
If ListViewP21.ListItems.Count > 0 Then TextBoxP21 = ListViewP21.SelectedItem.Text
End Sub
Private Sub CommandButton32_Click()
On Error Resume Next
If ListViewP21.ListItems.Count < 1 Or TextBoxP21 = "" Then Exit Sub
Application.ScreenUpdating = False
Dim ds, f
Set ds = CreateObject("Scripting.FileSystemObject")
Set f = ds.GetFolder("C:\Belgelerim\Cemex\Çeviri Dosyaları\Kelime Dosyaları")
Workbooks.Open fileName:=f & "\" & TextBoxP21
Application.Windows(TextBoxP21.Text).Visible = False
ListViewP31.ListItems.Clear: ListViewP31.ColumnHeaders.Clear
cls = Workbooks(TextBoxP21.Text).ActiveSheet.Cells(1, 256).End(xlToLeft).Column
For k = 1 To cls
Call ListViewP31.ColumnHeaders.Add(k, , Workbooks(TextBoxP21.Text).ActiveSheet.Cells(1, k), 95)
Next
Dim son As Integer
son = Workbooks(TextBoxP21.Text).ActiveSheet.Range("A" & "65536").End(xlUp).row
For n = 2 To son
    If Workbooks(TextBoxP21.Text).ActiveSheet.Range("A" & n).Value <> "" Then
    m = ListViewP31.ListItems.Count
    Call ListViewP31.ListItems.Add(m + 1, , Workbooks(TextBoxP21.Text).ActiveSheet.Range("A" & n))
    Call ListViewP31.ListItems(m + 1).ListSubItems.Add(1, , Workbooks(TextBoxP21.Text).ActiveSheet.Range("B" & n))
    End If
Next n
Workbooks(TextBoxP21.Text).Close False
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton33_Click()
On Error Resume Next
For i = 1 To ListViewP31.ListItems.Count
a = ListViewP31.ListItems(i): b = ListViewP31.ListItems(i).ListSubItems(1)
ActiveSheet.Cells.Replace What:=a, Replacement:=b, LookAt:=xlPart
Next
End Sub