Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cmd As String : cmd = CStr(param)
    Dim sh As Object : Set sh = CreateObject("Shell.Application")
    sh.ShellExecute "powershell.exe", "-NoProfile -Command " & Chr(34) & cmd & Chr(34), "", "runas", 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Komut" : ws.Range("B1").Value = Left$(cmd, 32000)
    Set DynamicFunc = Nothing
End Function
