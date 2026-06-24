Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim$(CStr(param))
    targetWb.ActiveSheet.Shapes.AddPicture url, False, True, 50, 50, 200, 200
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Resim" : ws.Range("B1").Value = url
    Set DynamicFunc = Nothing
End Function
