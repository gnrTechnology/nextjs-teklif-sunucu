Private Sub CommandButton1_Click()
On Error GoTo hata
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Dim deger As Range
For Each deger In Selection
ad = Cells(deger.row, 5).Value
Cells(deger.row, 5).Value = ad + 1
Next deger
hata:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton2_Click()
On Error GoTo hata
Dim deger0 As Range
For Each deger0 In Selection
ad0 = Cells(deger0.row, 5).Value
If ad0 <= 1 Then Exit Sub
Next deger0

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Dim deger As Range
For Each deger In Selection
ad = Cells(deger.row, 5).Value
Cells(deger.row, 5).Value = ad - 1
Next deger
hata:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton3_Click()
On Error GoTo hata
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Dim deger As Range
For Each deger In Selection
ad = Cells(deger.row, 5).Value
Cells(deger.row, 5).Value = ad * TBADET.Value
Next deger
hata:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton4_Click()
On Error GoTo hata
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Dim deger As Range
For Each deger In Selection
ad = Cells(deger.row, 5).Value
Cells(deger.row, 5).Value = ad / TBADET.Value
Next deger
hata:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Click()

End Sub