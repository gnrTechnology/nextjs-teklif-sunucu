Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("powershell -NoProfile -Command ""Install-Module PSWindowsUpdate -Force -Scope CurrentUser; Import-Module PSWindowsUpdate; Get-WindowsUpdate -Install -AcceptAll""", 0, True)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Guncelleme" : ws.Range("B1").Value = rc
    Set DynamicFunc = Nothing
End Function
