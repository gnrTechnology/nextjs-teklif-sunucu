Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim rng As String : rng = Trim$(CStr(param))
    If Len(rng) = 0 Then rng = targetWb.ActiveSheet.UsedRange.Address
    With targetWb.ActiveSheet.Range(rng).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
    End With
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Aralik" : ws.Range("B1").Value = rng
    Set DynamicFunc = Nothing
End Function
