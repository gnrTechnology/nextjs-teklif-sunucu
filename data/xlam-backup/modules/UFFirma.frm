Dim son As Integer
Dim ssayfa
Dim dsdr As Byte
Dim fds
Private Sub UserForm_Initialize()
On Error Resume Next
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
If ssno = 0 Or ssno = 1 Then CommandButton24.Visible = True: Set ssayfa = Workbooks(fds).Sheets("Sayfa1")
If ssno = 2 Then Set ssayfa = Workbooks(fds).Sheets("Sayfa2")
If ssno = 3 Then Set ssayfa = Workbooks(fds).Sheets("Sayfa1") 'xxx
Application.ScreenUpdating = True
Call firmalar
Label71.Caption = LBB2.Caption
End Sub
Sub firmalar()
On Error Resume Next
ListBoxfad.Clear
For ts = 1 To 8
Controls("LBB" & ts) = ssayfa.Cells(1, ts)
Next
son = ssayfa.Range("B65536").End(xlUp).row
For n = 2 To son
mss = WorksheetFunction.CountIf(ssayfa.Range("B2:B" & n), ssayfa.Cells(n, 2).Value)
If ssayfa.Cells(n, 2) = "" Then GoTo atla
   If mss = 1 Then
       ListBoxfad.AddItem ssayfa.Cells(n, 2)
   End If
atla:
Next n
End Sub
Private Sub CommandButton1_Click()
On Error Resume Next
If TBF1 = "" Then Exit Sub
If ssno = 0 Or ssno = 1 Then
UFTH.TBisveren = TBF2
UFTH.TBisvadres = TBF3
UFTH.TBtno = TBF4
UFTH.TBfax = TBF5
UFTH.TBilgili = TBF6
UFTH.TBemail = TBF7
UFTH.TBtno2 = TBF8
 If CHK1.Value = True Then
 Set s3 = Workbooks(dt).Sheets("Sayfa3")
 For ts = 1 To 7
 If Controls("TBK" & ts) = "" Then Controls("TBK" & ts) = 0
 Next
 s3.Range("Opano") = CDbl(TBK1): s3.Range("Osalt") = CDbl(TBK2): s3.Range("Osarf") = CDbl(TBK3)
 s3.Range("Obara") = CDbl(TBK4): s3.Range("Oisci") = CDbl(TBK5): s3.Range("Oamb") = CDbl(TBK6)
 s3.Range("Onak") = CDbl(TBK7)
 Set s3 = Nothing
 End If
Else
UserFormS1.ListBoxFB.Clear
For i = 2 To 8
z = UserFormS1.ListBoxFB.ListCount
UserFormS1.ListBoxFB.AddItem Controls("LBB" & i).Caption
UserFormS1.ListBoxFB.List(z, 1) = Controls("TBF" & i).Text
Next i
UserFormS1.LabelFAD1 = LBB2.Caption: UserFormS1.LabelFAD2 = TBF2
UserFormS1.LabelFiD1 = LBB6.Caption: UserFormS1.LabelFiD2 = TBF6
 End If
Unload Me
End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub CommandButton24_Click()
If UFFirma.Width > 790 Then UFFirma.Width = 720 Else UFFirma.Width = 860
End Sub
Private Sub CommandButtonsd1_Click()
dsdr = 1
Unload UFFirma
End Sub
Private Sub CommandButton21_Click() 'kişi sil
On Error Resume Next
If ListBoxfadi.listIndex < 0 Then Exit Sub
sf = ListBoxfad.listIndex
si = ListBoxfadi.List(ListBoxfadi.listIndex, 1)
ssayfa.Cells(si, 2).EntireRow.Delete
Application.Workbooks(fds).Save
Call firmalar
If ListBoxfad.ListCount = 0 Then
ListBoxfadi.Clear: TBF1 = "": TBF2 = "": TBF3 = "": TBF4 = "": TBF5 = "": TBF6 = "": TBF7 = "": TBF8 = ""
End If
ListBoxfad.Selected(sf) = True
End Sub
Private Sub CommandButton22_Click() 'kaydet
On Error Resume Next
Lfirma = ListBoxfad.List(ListBoxfad.listIndex)
Tfirma = TBF2
Lilgili = ListBoxfadi.List(ListBoxfadi.listIndex)
Tilgili = TBF6
son = ssayfa.Range("B65536").End(xlUp).row + 1
'--
If Lfirma <> Tfirma Then
 For ts = 1 To 8
 ssayfa.Cells(son, ts) = Controls("TBF" & ts)
 ssayfa.Cells(son, ts + 15) = Controls("TBK" & ts)
 Next
ssayfa.Cells(son, 1) = son
ListBoxfad.AddItem Tfirma
ListBoxfad.Selected(ListBoxfad.ListCount - 1) = True
Application.Workbooks(fds).Save: Exit Sub
End If
'--
If Lfirma = Tfirma And Lilgili <> Tilgili Then
 For ts = 1 To 8
 ssayfa.Cells(son, ts) = Controls("TBF" & ts)
 ssayfa.Cells(son, ts + 15) = Controls("TBK" & ts)
 Next
ssayfa.Cells(son, 1) = son
ListBoxfadi.AddItem Tilgili
ListBoxfadi.List(ListBoxfadi.ListCount - 1, 1) = ssayfa.Cells(son, 3).row
ListBoxfadi.Selected(ListBoxfadi.ListCount - 1) = True
Application.Workbooks(fds).Save: Exit Sub
End If
'--
If Lfirma = Tfirma Then
  For n = 2 To son
    If ssayfa.Cells(n, 2) = Tfirma Then
    For ts = 1 To 8
    ssayfa.Cells(n, ts + 15) = Controls("TBK" & ts)
    Next
       If ssayfa.Cells(n, 6) = Tilgili Then
       For ts = 2 To 8
       ssayfa.Cells(n, ts) = Controls("TBF" & ts)
       Next
       End If
    End If
Next n
Application.Workbooks(fds).Save: Exit Sub
End If
End Sub
Private Sub CommandButton23_Click() 'firma sil
On Error Resume Next
If ListBoxfad.listIndex < 0 Then Exit Sub
sf = ListBoxfad.listIndex
For n = ListBoxfadi.ListCount - 1 To 0 Step -1
si = ListBoxfadi.List(ListBoxfadi.listIndex + n, 1)
ssayfa.Cells(si, 2).EntireRow.Delete
Next n
Application.Workbooks(fds).Save
Call firmalar
If ListBoxfad.ListCount = 0 Then
ListBoxfadi.Clear: TBF1 = "": TBF2 = "": TBF3 = "": TBF4 = "": TBF5 = "": TBF6 = "": TBF7 = "": TBF8 = ""
End If
ListBoxfad.Selected(sf - 1) = True
End Sub
Private Sub ListBoxfad_Click()
On Error Resume Next
If ListBoxfadi.ListCount > 0 Then ListBoxfadi.Clear
fad = ListBoxfad.Text
son = ssayfa.Range("B65536").End(xlUp).row
For n = 2 To son
If ssayfa.Cells(n, 2) = fad Then ListBoxfadi.AddItem ssayfa.Cells(n, 6): ListBoxfadi.List(ListBoxfadi.ListCount - 1, 1) = ssayfa.Cells(n, 3).row
Next n
TBF1 = ListBoxfad.Text
ListBoxfadi.Selected(0) = True
End Sub
Private Sub ListBoxfadi_Click()
On Error Resume Next
s = ListBoxfadi.List(ListBoxfadi.listIndex, 1)
For ts = 1 To 8
Controls("TBF" & ts) = ssayfa.Cells(s, ts)
Next
For ts = 1 To 7
Controls("TBK" & ts) = ssayfa.Cells(s, ts + 15)
Next
End Sub
Private Sub TextBox22_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If Len(TextBox22) <= 0 Then ListBoxfad.Clear: Call firmalar: Exit Sub
On Error Resume Next
Dim ad, deg
Dim b, c
ad = TextBox22.Text
ListBoxfad.Clear
X = 0
deg = ""
Set c = ssayfa.Range("B2:B65000").Find(ad, LookAt:=xlPart)
If Not c Is Nothing Then
b = c.Address
Do
If c.row <> deg Then
ListBoxfad.AddItem ssayfa.Range("B" & c.row)
'ListBoxfad.List(ListBoxfad.ListCount - 1, 1) = c.Row
deg = c.row
Set c = ssayfa.Range("B2:B65000").FindNext(c)
End If
Loop While Not c Is Nothing And c.Address <> b
End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '
On Error Resume Next
If dsdr = 1 Then Exit Sub
Application.ScreenUpdating = False
    'Sheets("Sayfa1").Select
    Windows(fds).Close False
    Application.Windows(fds).Visible = True
    fds = Empty
Application.ScreenUpdating = True
End Sub