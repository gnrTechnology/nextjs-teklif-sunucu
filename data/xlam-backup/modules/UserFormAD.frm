Private Sub UserForm_Activate()
If UserFormAD.Caption = "Proje Adı / No:" Then
TextBox21.Text = "Proje Adı": TextBox19.Width = 272: TextBox20.Visible = False
Label484.Visible = False: CBPC1.Visible = False
End If
If UserFormAD.Frame4.BackColor = &H96A446 Then
TextBox21.Text = ".": TextBox19.Width = 272: TextBox20.Visible = False
Label484.Visible = False: CBPC1.Visible = False
End If
End Sub
Private Sub CommandButton12_Click()
If TextBox19.Value = "" Then Exit Sub
Selection.EntireRow.Insert
If UserFormAD.Caption = "Proje Adı / No:" Then ActiveCell.FormulaR1C1 = "PROJE ADI/NO:" Else ActiveCell.FormulaR1C1 = "BÖLÜM ADI/NO:"
'Biçimlemeler'--
Y = Selection.row
If Left(ActiveWorkbook.ActiveSheet.CodeName, 2) = "TM" Then
Range("A" & Y & ":J" & Y).Select: Selection.Borders.LineStyle = xlContinuous
Else
Range("A" & Y & ",B" & Y & ":E" & Y & ",F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Select: Selection.Borders.LineStyle = xlContinuous
Range("F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Borders(xlInsideVertical).LineStyle = xlNone
End If
Range("B" & Y & ":E" & Y).Borders(xlInsideVertical).LineStyle = xlNone
Selection.Interior.Pattern = xlNone: Selection.RowHeight = 12.75: Selection.Font.Size = 9: Selection.Font.Bold = True
Range("C" & Y) = TextBox19.Value
Range("A" & Y & ":D" & Y).NumberFormat = "@"
If UserFormAD.Caption = "Proje Adı / No:" Or UserFormAD.Frame4.BackColor = &H96A446 Then
Selection.Font.ColorIndex = 54: GoTo son1
Else
Selection.Font.ColorIndex = 11
If Not TextBox20.Value = "" Then Range("E" & Y).FormulaR1C1 = TextBox20.Value: Range("E" & Y).NumberFormat = "#,##0 ""Adet"""
End If
If CBPC1.Value = True Then
If TextBox21.Value = "" Then Range("A" & Y).Value = "Pano Ref." Else Range("A" & Y).Value = TextBox21.Value
End If
son1:
Unload Me
End Sub
Private Sub CommandButton13_Click()
Unload Me
End Sub
Private Sub CommandButton14_Click()
TextBox19.Value = "": TextBox20.Value = 1
End Sub
Private Sub TextBox20_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
    Case 48 To 57
    Case Else
        KeyAscii = 0
    End Select
End Sub