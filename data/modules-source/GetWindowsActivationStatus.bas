Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim exec As Object
    Set exec = wsh.Exec("powershell -NonInteractive -Command ""(Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -ne $null -and $_.Name -like '*Windows*' } | Select-Object -First 1).LicenseStatus""")
    Dim out As String
    Do While exec.Status = 0 : Application.Wait Now + TimeValue("00:00:01") : Loop
    out = Trim(exec.StdOut.ReadAll)
    Dim status As String
    Select Case out
        Case "1" : status = "Aktif (Lisanslı)"
        Case "0" : status = "Lisanssız"
        Case "2" : status = "Ek Bilgi Gerekli"
        Case "5" : status = "Bildirim Modu"
        Case Else : status = "Durum: " & out
    End Select
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Windows Aktivasyon Durumu"
    ws.Range("B1").Value = status
    ws.Range("A1").Font.Bold = True
    ws.Columns.AutoFit
    MsgBox "Aktivasyon: " & status, vbInformation, "GetWindowsActivationStatus"
    Set DynamicFunc = Nothing
End Function