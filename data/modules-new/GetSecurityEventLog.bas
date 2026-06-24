Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error GoTo Fail
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long : rc = sh.Run("powershell -NoProfile -Command ""Get-EventLog -LogName Security -Newest 20 | Format-Table -AutoSize | Out-String"", 0, True)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "GetSecurityEventLog" : ws.Range("B1").Value = rc
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing
End Function
