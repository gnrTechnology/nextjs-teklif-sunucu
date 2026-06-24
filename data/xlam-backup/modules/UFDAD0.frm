Dim nft As String
Private Sub UserForm_Initialize() 'tamam 2021 xxx
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
Toolbar1.Buttons.Item(4).Image = ImageList1.ListImages.Item(4).Index
Toolbar1.Buttons.Item(5).Image = ImageList1.ListImages.Item(6).Index
Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(7).Index
Toolbar1.Buttons.Item(7).Image = ImageList1.ListImages.Item(8).Index
Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(9).Index
Toolbar1.Buttons.Item(9).Image = ImageList1.ListImages.Item(10).Index
End Sub
Sub formatcevir()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlManual
Dim deger As Range
If Selection.Rows.Count > 1 Then Selection.SpecialCells(xlCellTypeVisible).Select
For Each deger In Selection
If Not deger = "" Then
deger.Value = CDbl(deger.Value): deger.NumberFormat = nft
If nft = "#,##0.00" Then deger.Font.ColorIndex = xlAutomatic
If nft = "#,##0.00 [$$-C0C]" Then deger.Font.ColorIndex = 3
If nft = "#,##0.00 [$€-1]" Then deger.Font.ColorIndex = 5
End If
Next deger
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo hata
sat = ActiveCell.SpecialCells(xlLastCell).row
Dim deger As Range
If Selection.Rows.Count > 1 Then Selection.SpecialCells(xlCellTypeVisible).Select
Select Case Button.Index
Case 1
nft = "#,##0.00": Call formatcevir
Case 2
nft = "#,##0.00 [$$-C0C]": Call formatcevir
Case 3
nft = "#,##0.00 [$€-1]": Call formatcevir
Case 4
    For Each deger In Selection
    If deger.row > sat Then Exit For
    If Not deger = "" Then deger.Value = UCase(Replace(Replace(deger.Value, "ı", "I"), "i", "İ"))
    Next
Case 5
    For Each deger In Selection
    If deger.row > sat Then Exit For
    If Not deger = "" Then deger.Value = Application.Proper(deger.Value)
    Next
Case 6
    For Each deger In Selection
    If deger.row > sat Then Exit For
    If Not deger = "" Then deger.Value = WorksheetFunction.Trim(deger.Value)
    Next
    MsgBox " Seçilen sutunda bulunan başta ve sondaki boşluklar kırpıldı."
Case 7
Call adetsay1
Case 8
Call aynıbul1
Case 9
Call adettopla1
End Select
hata:
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
sat = ActiveCell.SpecialCells(xlLastCell).row
Dim deger As Range
Select Case ButtonMenu.Tag
Case 1
    For Each deger In Selection
    If deger.row > sat Then Exit For
        deger.Value = Evaluate("=LOWER(" & """" & deger.Value & """" & ")")
    Next
Case 2
    Selection.Cells.Replace What:="¤", Replacement:="ğ"
    Selection.Cells.Replace What:="›", Replacement:="ı"
    Selection.Cells.Replace What:="‹", Replacement:="i"
    Selection.Cells.Replace What:="fl", Replacement:="ş"
Case 3
    Selection.Cells.Replace What:=vbLf, Replacement:=" "
    Selection.Cells.Replace What:="  ", Replacement:=" "
Case 4
Call aynıbul2
End Select
End Sub
Sub adetsay1()
On Error Resume Next
Dim say As Integer
    sa = Selection.Column
    sat1 = Selection.row: sat2 = sat1 + Selection.Cells.Count - 1
    If sat2 > 15000 Then sat2 = ActiveCell.SpecialCells(xlLastCell).row
    For i = sat1 To sat2
    If Not Cells(i, sa) = "" And Not Cells(i, sa) = "-" Then
     If WorksheetFunction.CountIf(Range(Cells(sat1, sa), Cells(i, sa)), Cells(i, sa).Value) = 1 Then
     msg = Msga & Cells(i, sa).Value & " = " & WorksheetFunction.CountIf(Range(Cells(sat1, sa), Cells(sat2, sa)), Cells(i, sa).Value) & " Ad."
     Msga = msg & vbCr
     End If
    End If
    Next
    Dim Myfile As String
    If Not ActiveWorkbook.path = "" Then Myfile = ActiveWorkbook.path & "\" & "Miktar" & ".txt" _
    Else Myfile = "C:\Belgelerim\Cemex" & "\" & "Miktar" & ".txt"
    If Len(dir(Myfile)) > 0 Then
    Close Myfile
    Kill Myfile
    End If
    Open Myfile For Append As #1
    Print #1, msg
    Close #1
CreateObject("Shell.Application").Open (Myfile)
    'MsgBox Msg
End Sub
Sub adettopla1()
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlManual
Dim say As Integer
    sa = 2
    sat1 = Selection.row: sat2 = sat1 + Selection.Cells.Count - 1
    If sat2 > 15000 Then sat2 = ActiveCell.SpecialCells(xlLastCell).row
    Cells(sat1, sa).Select
For i = sat1 To sat2
    If Cells(i, sa) = "" Or Cells(i, sa) = "." Then Cells(i, sa).EntireRow.Delete: sat2 = sat2 - 1: GoTo git1
    If Cells(i, sa) = "BÖLÜM ADI/NO:" Or Cells(i, sa) = "BÖLÜM TOPLAMI:" Then GoTo git1
    kod1 = Cells(i, sa)
      For n = sat2 To i + 1 Step -1
      If kod1 = Cells(n, sa) Then
      Cells(i, 5) = val(Cells(i, 5)) + val(Cells(n, 5))
      Cells(n, sa).EntireRow.Delete
      sat2 = sat2 - 1
      End If
      Next
git1:
Next
Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
End Sub
Sub aynıbul1()
On Error Resume Next
kod1 = Selection.Cells
sa = Selection.Column
sr = Selection.row
son = Cells(Rows.Count, sa).End(xlUp).row
If kod1 = "" Or kod1 = "." Then Exit Sub
If kod1 = "BÖLÜM ADI/NO:" Or kod1 = "BÖLÜM TOPLAMI:" Then Exit Sub
Cells(sr, sa).Interior.ColorIndex = 40
For i = 2 To son
    If Cells(i, sa).Interior.ColorIndex = 38 And kod1 <> Cells(i, sa) Then
    Cells(i, sa).Interior.ColorIndex = 0
    End If
    If kod1 = Cells(i, sa) And i <> sr Then
    Cells(i, sa).Interior.ColorIndex = 38
    Cells(sr, sa).Interior.ColorIndex = 38
    End If
Next
Application.ScreenUpdating = True
End Sub
Sub aynıbul2()
On Error Resume Next
kod1 = Selection.Cells
sa = Selection.Column
sr = Selection.row
If kod1 = "" Or kod1 = "." Then Exit Sub
If kod1 = "BÖLÜM ADI/NO:" Or kod1 = "BÖLÜM TOPLAMI:" Then Exit Sub
bst1 = InputBox(sa & ". Sutun için Eşleştirme Yapılacak Sutun:" & vbCr & "", "Sutun Seçimi", "F")
If bst1 = "" Then Exit Sub
skd1 = Range(bst1 & sr)
sr2 = Range(bst1 & sr).Column
son = Cells(Rows.Count, sr2).End(xlUp).row
Cells(sr, sa).Interior.ColorIndex = 40: Range(bst1 & sr).Interior.ColorIndex = 40
For i = 2 To son
    If Cells(i, sa).Interior.ColorIndex = 38 And kod1 <> Cells(i, sa) Then
    Cells(i, sa).Interior.ColorIndex = 0
    End If
    If kod1 = Cells(i, sa) And i <> sr Then
    Cells(i, sa).Interior.ColorIndex = 38: Cells(i, sr2).Interior.ColorIndex = 38
    Cells(sr, sa).Interior.ColorIndex = 38: Cells(sr, sr2).Interior.ColorIndex = 38
    Cells(i, sr2) = Cells(sr, sr2).Text
    End If
Next
End Sub