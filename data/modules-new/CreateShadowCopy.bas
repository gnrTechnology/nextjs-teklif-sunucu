Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim drive As String : drive = Trim$(CStr(param)) : If Len(drive) = 0 Then drive = "C:"
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("vssadmin create shadow /for=" & drive, 0, True)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Surucu" : ws.Range("B1").Value = drive
    ws.Range("A2").Value = "Sonuc" : ws.Range("B2").Value = rc
    Set DynamicFunc = Nothing
End Function
