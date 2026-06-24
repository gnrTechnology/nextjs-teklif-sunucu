Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "ConvertSheetToJson", targetWb, ""
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A3").Value = "PivotJSON" : ws.Range("B3").Value = "headers+rows format"
    Set DynamicFunc = Nothing
End Function
