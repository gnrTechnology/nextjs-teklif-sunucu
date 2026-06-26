Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Tick64" : ws.Range("B1").Value = GetTickCount64()
    Set DynamicFunc = Nothing
End Function
