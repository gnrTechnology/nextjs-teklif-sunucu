Sub CreateCmdBar()
    Dim St As CommandBar
     'delete the pop-up if it exists
    On Error Resume Next
    Application.CommandBars("flexgrid_rc").Delete
     'Disables enabled error handler in the current procedure and resets it to Nothing. On Error GoTo 0
    On Error GoTo 0
    Set St = CommandBars.Add(Name:="flexgrid_rc", Position:=msoBarPopup, Temporary:=False)
    With St
        .Controls.Add Type:=msoControlButton
        .Controls(1).Caption = "Yeni Dizin"
        .Controls(1).OnAction = "menu1"
    End With
        With St
        .Controls.Add Type:=msoControlButton
        .Controls(2).Caption = "Yeni Teklif"
        .Controls(2).OnAction = "menu2"
    End With
    With St
        .Controls.Add Type:=msoControlButton
        .Controls(3).Caption = "Yeniden Adlandır"
        .Controls(3).OnAction = "menu3"
    End With
    With St
        .Controls.Add Type:=msoControlButton
        .Controls(4).Caption = "Sil"
        .Controls(4).OnAction = "menu4"
    End With
    St.Enabled = True
End Sub
Sub DestroyCmdBar()
    On Error Resume Next
    Application.CommandBars("flexgrid_rc").Delete
    On Error GoTo 0
End Sub
Sub Menu1() 'Yeni Dizin
On Error Resume Next
Dim s As Integer
Static q
s = UFTH.LB
dsa = Format(Now, "dd/mm/yyyy")
If s > 0 Then
ds = InputBox("Adı Giriniz.", "Klasör / Dosya Adı", dsa)
If ds = "" Then Exit Sub
'ad = "Yeni Klasör" & q
Yol = UFTH.Labelyol
'..
If Right(Yol, 5) = ".xlsx" Or Right(Yol, 4) = ".xls" Then
Set objFSO = CreateObject("scripting.filesystemobject")
Set X = objFSO.GetFile(Yol): Yol = X.ParentFolder
s = UFTH.TreeView1.Nodes(Yol).Index
End If
'..
kl = Yol & "\" & ds

Call UFTH.TreeView1.Nodes.Add(s, tvwChild, kl, ds, Image:=2)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(kl)
UFTH.Labelyol = kl
'q = q + 1
UFTH.TreeView1.SelectedItem.Expanded = True
UFTH.TreeView1.SelectedItem.EnsureVisible
UFTH.TreeView1.Nodes(kl).Selected = True
UFTH.LB = UFTH.TreeView1.Nodes(kl).Index
'UFTH.TreeView1.StartLabelEdit
End If
End Sub
Sub menu2() 'Yeni Teklif
On Error Resume Next
Dim s As Integer
s = UFTH.LB
dsa = Format(Now, "DD MMMM YYYY hh:mm:ss")
ds = InputBox("Teklif İsmini Giriniz.", "Dosya Adı", "Yeni Teklif" & "-" & Format(dsa, "DDMMYY-hhmmss"))
Yol = UFTH.Labelyol
If ds = "" Then Exit Sub
Application.ScreenUpdating = False: Application.DisplayAlerts = False
Workbooks.Open "C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx"
'..
If Right(UFTH.TBds, 5) = ".xlsx" Or Right(UFTH.TBds, 4) = ".xls" Then
Set objFSO = CreateObject("scripting.filesystemobject")
Set X = objFSO.GetFile(Yol): Yol = X.ParentFolder
s = UFTH.TreeView1.Nodes(Yol).Index
End If
'..
Sheets("Sayfa3").Range("C8") = dsa 'Teklif Tarihi
kl = Yol & "\" & ds & ".xlsx"
ActiveWorkbook.SaveAs kl
Call UFTH.TreeView1.Nodes.Add(s, tvwChild, kl, ActiveWorkbook.Name, Image:=3)
UFTH.TreeView1.Nodes(Yol).Image = 1
c = UFTH.TreeView1.SelectedItem
UFTH.TreeView1.Nodes(kl).Selected = True
UFTH.LB = UFTH.TreeView1.SelectedItem.Index
'UFTH.Labelyol = kl
UFTH.TBds = ds & ".xlsx"
'UFTH.Veri_al
Workbooks(ds & ".xlsx").Close False
Application.ScreenUpdating = True: Application.DisplayAlerts = True
'Windows(kl).Close False
'Workbooks(kl).Close False
'ActiveWindow.Close
End Sub
Sub menu3() 'Yeniden Adlandır
On Error Resume Next
kA = UFTH.TreeView1.SelectedItem
kl = UFTH.Labelyol & "\" & kA
Set objFSO = CreateObject("scripting.filesystemobject")

If Right(kA, 5) = ".xlsx" Or Right(kA, 4) = ".xls" Then
Set X = objFSO.GetFile(kl)
hd = X.Name
dg = "dsy"
Else
kl = UFTH.Labelyol
Set X = objFSO.GetFolder(kl)
dg = "kls"
End If

Dim s As Integer
s = UFTH.LB
ds = InputBox("Adı Giriniz.", "Klasör / Dosya Adı", Replace(kA, ".xlsx", ""))
If ds & ".xlsx" = kA Or ds = "" Then Exit Sub Else ds = ds & ".xlsx"
'..
ky = X.ParentFolder & "\" & ds
UFTH.TreeView1.SelectedItem.key = ky: UFTH.TreeView1.SelectedItem.Text = ds
'UFTH.Labelyol = ky
'..eğer dosya ise
If dg = "dsy" Then
'..eğer dosya açık ise
If WorkbookOpen((kA)) Then
Application.ScreenUpdating = False: Application.DisplayAlerts = False: Unload UFTH
    Workbooks(kA).ChangeFileAccess Mode:=xlReadOnly
    objFSO.movefile kl, ky
    'Name ka.FullName As ds
    Workbooks(kA).Close SaveChanges:=False
    Workbooks.Open (ky)
Application.ScreenUpdating = True: Application.DisplayAlerts = True
UFTH.Show
Exit Sub
End If
'..eğer dosya açık değilse
objFSO.movefile kl, ky
End If
'..eğer klasör ise
If dg = "kls" Then
objFSO.MoveFolder kl, ky
   For Each kad In CreateObject("Scripting.FileSystemObject").GetFolder(ky).SubFolders
   ekey = kl & "\" & kad.Name: ykey = ky & "\" & kad.Name
   UFTH.TreeView1.Nodes.Item(ekey).key = ykey
 Next
End If
End Sub
Sub menu4xx() 'sil
On Error GoTo hata
msg = MsgBox("Kalıcı olarak silmek istediğinizden emin misiniz?", vbYesNo, "Silme İşlemi")
If msg = vbNo Then Exit Sub
'--
Dim s As Integer
s = UFTH.LB
kA = UFTH.TreeView1.SelectedItem
kl = UFTH.TreeView1.SelectedItem.key
'--
    If Not WorkbookOpen((kA)) Then kA = Empty
    If kA <> Empty Then msg = MsgBox("Dosya açık olduğundan bu işlem yapılamaz!", vbCritical): Exit Sub
'--
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set X = objFSO.GetFile(kl)
hd = X.Name
Workbooks(hd).Close False
'--
objFSO.DeleteFolder (kl) 'klasör sil
objFSO.DeleteFile (kl) 'dosya sil
UFTH.TreeView1.Nodes.Remove (s)
hata:
End Sub
Sub menu4() 'sil
On Error GoTo hata
kA = UFTH.TreeView1.SelectedItem
msg = MsgBox(kA & " kalıcı olarak silmek istediğinizden emin misiniz?", vbYesNo, "Silme İşlemi")
If msg = vbNo Then Exit Sub
'--
Dim s As Integer
s = UFTH.LB
kl = UFTH.TreeView1.SelectedItem.key
'--
    If Not WorkbookOpen((kA)) Then kA = Empty
    If kA <> Empty Then msg = MsgBox("Dosya açık olduğundan bu işlem yapılamaz!", vbCritical): Exit Sub
'--
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set x = objFSO.GetFile(kl)
'hd = x.Name
'Workbooks(hd).Close False
'--
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Right(kl, 5) = ".xlsx" Or Right(kl, 4) = ".xls" Then
objFSO.DeleteFile (kl) 'dosya sil
UFTH.TreeView1.Nodes.Remove (s)
UFTH.Labelyol = ""
UFTH.TBds = ActiveWorkbook.Name
UFTH.Toolbar1.Buttons.Item(5).Enabled = False
Else
objFSO.DeleteFolder (kl) 'klasör sil
UFTH.TreeView1.Nodes.Remove (s)
UFTH.TBds = ActiveWorkbook.Name
UFTH.Labelyol = ""
UFTH.Toolbar1.Buttons.Item(5).Enabled = False
End If
UFTH.Labelyol = UFTH.TreeView1.SelectedItem.key
UFTH.LB = UFTH.TreeView1.SelectedItem.Index
Exit Sub
hata:
msg = MsgBox("Dosya açık olduğundan bu işlem yapılamaz!", vbCritical)
End Sub
Sub DOSYAADI_R1() 'Dosya adı değiştir
On Error Resume Next
If Not UserFormS2.TBAA4 = "" Then
dy01 = ActiveWorkbook.fullName
da01 = ActiveWorkbook.Name: da02 = UserFormS2.TBAA4 & ".xlsx"
If Not da01 = da02 Then
Set objFSO = CreateObject("scripting.filesystemobject")
dy02 = ActiveWorkbook.path & Application.PathSeparator & da02
Worksheets.Item("Veriler").Range("B4") = UserFormS2.TBAA4
Application.ScreenUpdating = False: Application.DisplayAlerts = False: Unload UserFormS2
    Workbooks(da01).ChangeFileAccess Mode:=xlReadOnly
    objFSO.movefile dy01, dy02
    Workbooks(da01).Close SaveChanges:=False
    Workbooks.Open (dy02)
Application.ScreenUpdating = True: Application.DisplayAlerts = True
'---
Worksheets.Item("Veriler").Range("b30") = ActiveWorkbook.path
Worksheets.Item("Veriler").Range("b31") = Replace(ActiveWorkbook.Name, ".xlsx", "")
'---
UserFormS2.Show
End If
End If
End Sub