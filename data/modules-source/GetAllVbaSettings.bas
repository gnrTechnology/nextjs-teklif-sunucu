Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sections() As String : sections = Split("ilhan,scngnr,sercan", ",")
    Dim keys() As String : keys = Split("mac,mdip,TBveren,teklifYolu,startingAddin,ihlalDosyaYolu,ihlalDosyaAdi,apiBaseUrl", ",")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Bölüm" : ws.Range("B1").Value = "Anahtar" : ws.Range("C1").Value = "Değer"
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    Dim sec As Variant, k As Variant
    For Each sec In sections
        For Each k In keys
            Dim v As String : v = GetSetting(CStr(sec), "Settings", CStr(k), "(yok)")
            ws.Cells(r,1).Value = sec : ws.Cells(r,2).Value = k : ws.Cells(r,3).Value = v : r=r+1
        Next k
    Next sec
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
    Call PostOutputToServer("GetAllVbaSettings", outJson)

    Set DynamicFunc = Nothing
End Function
Private Sub PostOutputToServer(moduleName As String, outputJson As String)
    On Error Resume Next
    Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")
    If mac = "" Then mac = GetMacFromWmi()
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
    http.Open "POST", baseUrl & "module-output/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 10000, 10000
    http.send body
    On Error GoTo 0
End Sub
Private Function GetMacFromWmi() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then
            GetMacFromWmi = o.MACAddress
            Exit Function
        End If
    Next
End Function

