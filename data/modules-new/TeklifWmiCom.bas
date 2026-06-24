Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Wmi")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing
End Function
