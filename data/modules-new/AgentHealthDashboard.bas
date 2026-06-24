Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error GoTo Fail
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long : rc = sh.Run("powershell -NoProfile -Command ""$os=Get-CimInstance Win32_OperatingSystem; $cpu=(Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue; @{UptimeH=[math]::Round(((Get-Date)-(Get-CimInstance Win32_OperatingSystem).LastBootUpTime).TotalHours,1); CpuPct=[math]::Round($cpu,1); FreeMemGB=[math]::Round($os.FreePhysicalMemory/1MB,1)} | ConvertTo-Json"", 0, True)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Health" : ws.Range("B1").Value = rc
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing
End Function
