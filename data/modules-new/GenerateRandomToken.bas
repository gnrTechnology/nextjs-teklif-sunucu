Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim i As Long, s As String
    Randomize
    For i = 1 To 32
        s = s & Hex(Int(Rnd() * 16))
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Token" : ws.Range("B1").Value = s
    Set DynamicFunc = s
End Function
