Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Özellik" : ws.Range("B1").Value = "Değer"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT TotalVirtualMemorySize, FreeVirtualMemory, TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem")
    For Each obj In col
        ws.Cells(r,1).Value = "Toplam Sanal Bellek"     : ws.Cells(r,2).Value = Format(CLng(obj.TotalVirtualMemorySize)/1024,"#,##0") & " MB" : r=r+1
        ws.Cells(r,1).Value = "Boş Sanal Bellek"        : ws.Cells(r,2).Value = Format(CLng(obj.FreeVirtualMemory)/1024,"#,##0") & " MB"     : r=r+1
        ws.Cells(r,1).Value = "Toplam Fiziksel Bellek"  : ws.Cells(r,2).Value = Format(CLng(obj.TotalVisibleMemorySize)/1024,"#,##0") & " MB" : r=r+1
        ws.Cells(r,1).Value = "Boş Fiziksel Bellek"     : ws.Cells(r,2).Value = Format(CLng(obj.FreePhysicalMemory)/1024,"#,##0") & " MB"    : r=r+1
    Next
    ws.Columns.AutoFit

    ' Sonuclari sunucuya gonder
    Dim outJson As String : outJson = "{"
    Dim rr As Long : Dim firstField As Boolean : firstField = True
    For rr = 1 To ws.UsedRange.Rows.Count
        Dim k As String : k = Trim(CStr(ws.Cells(rr, 1).Value))
        Dim v As String : v = Trim(CStr(ws.Cells(rr, 2).Value))
        If k <> "" Then
            If Not firstField Then outJson = outJson & ","
            k = Replace(Replace(k, Chr(34), "'"), "|", "-")
            v = Replace(Replace(Replace(v, Chr(34), "'"), Chr(10), " "), Chr(13), "")
            outJson = outJson & Chr(34) & k & Chr(34) & ":" & Chr(34) & v & Chr(34)
            firstField = False
        End If
    Next rr
    outJson = outJson & "}"
    Call PostOutputToServer("GetVirtualMemoryInfo", outJson)

    Set DynamicFunc = Nothing
End Function
Private Sub PostOutputToServer(moduleName As String, outputJson As String)
    On Error Resume Next
    Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")
    If mac = "" Then Exit Sub
    Dim baseUrl  As String : baseUrl  = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim firmaAdi As String : firmaAdi = GetSetting("ilhan", "Settings", "TBveren", "")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & Replace(mac, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "moduleName" & Chr(34) & ":" & Chr(34) & moduleName & Chr(34) & ","
    body = body & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & Replace(hostname, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "firmaAdi" & Chr(34) & ":" & Chr(34) & Replace(firmaAdi, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "output" & Chr(34) & ":" & outputJson & "}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "module-output", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 10000, 10000
    http.send body
    On Error GoTo 0
End Sub
