Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim db As String : db = Trim$(parts(0))
    Dim q As String : q = IIf(UBound(parts) >= 1, parts(1), "SELECT 1")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "SQLite icin D71 SqliteQueryLocal DLL onerilir"
    ws.Range("A2").Value = "DB" : ws.Range("B2").Value = db
    Set DynamicFunc = Nothing
End Function
