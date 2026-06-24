Private Sub UserForm_Activate()
Set s3 = Sheets("Sayfa3")
TextBoxEURO.Text = Format(s3.Range("Eur").Value, "#,####0.0000")  'Döviz kuru €
TextBoxUSD.Text = Format(s3.Range("Usd").Value, "#,####0.0000")  'Döviz kuru $
End Sub
Private Sub CommandButton12_Click()
On Error Resume Next
Set s3 = Sheets("Sayfa3")
s3.Range("Eur").Value = CDbl(TextBoxEURO.Value)
s3.Range("Usd").Value = CDbl(TextBoxUSD.Value)
If s3.Range("Tpbr") = "Teklif Para Birimi (USD)" Then s3.Range("Tpb").Value = CDbl(TextBoxUSD.Value)
If s3.Range("Tpbr") = "Teklif Para Birimi (EUR)" Then s3.Range("Tpb").Value = CDbl(TextBoxEURO.Value)
If s3.Range("Tpbr") = "Teklif Para Birimi (TL)" Then s3.Range("Tpb").Value = 1
Unload Me
End Sub
Private Sub CommandButton13_Click()
Unload Me
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