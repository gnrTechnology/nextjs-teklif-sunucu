Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Ms" : ws.Range("B1").Value = timeGetTime()
    Set DynamicFunc = Nothing
End Function
