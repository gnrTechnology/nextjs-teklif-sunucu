Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "EncryptCellRange", targetWb, param
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Cozuldu" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function
