Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim exec As Object
    Set exec = wsh.Exec("powershell -NonInteractive -Command ""[PSCustomObject]@{Locale=$([System.Globalization.CultureInfo]::CurrentCulture.Name);Language=$([System.Globalization.CultureInfo]::CurrentCulture.DisplayName);Currency=$([System.Globalization.CultureInfo]::CurrentCulture.NumberFormat.CurrencySymbol);DateFormat=$([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.ShortDatePattern);TimeZone=$([System.TimeZoneInfo]::Local.DisplayName)} | ConvertTo-Json""")
    Do While exec.Status = 0 : Application.Wait Now + TimeValue("00:00:01") : Loop
    Dim out As String : out = Trim(exec.StdOut.ReadAll)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Alan", "Değer")
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2:B2").Value = Array("Sistem Dili", ExtractJsonValue(out, "Language"))
    ws.Range("A3:B3").Value = Array("Locale Kodu", ExtractJsonValue(out, "Locale"))
    ws.Range("A4:B4").Value = Array("Para Birimi", ExtractJsonValue(out, "Currency"))
    ws.Range("A5:B5").Value = Array("Tarih Formatı", ExtractJsonValue(out, "DateFormat"))
    ws.Range("A6:B6").Value = Array("Saat Dilimi", ExtractJsonValue(out, "TimeZone"))
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function

Private Function ExtractJsonValue(json As String, key As String) As String
    Dim sk As String : sk = """" & key & """:"
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then ExtractJsonValue = "" : Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1 : p2 = InStr(p1, json, """")
        ExtractJsonValue = Mid(json, p1, p2 - p1)
    Else
        p2 = p1
        Do While p2 <= Len(json)
            If InStr(",}" & Chr(13) & Chr(10), Mid(json, p2, 1)) > 0 Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function