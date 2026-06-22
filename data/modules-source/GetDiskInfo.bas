Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT DeviceID, VolumeName, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType=3")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Surucu", "Etiket", "Toplam (GB)", "Bos (GB)", "Dosya Sistemi")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.DeviceID
        ws.Cells(r, 2).Value = obj.VolumeName
        ws.Cells(r, 3).Value = Format(obj.Size / 1073741824, "0.00")
        ws.Cells(r, 4).Value = Format(obj.FreeSpace / 1073741824, "0.00")
        ws.Cells(r, 5).Value = obj.FileSystem
        r = r + 1
    Next obj
    ws.Columns.AutoFit

    Dim outJson As String : outJson = "{"
    Dim rr As Long : Dim firstField As Boolean : firstField = True
    For rr = 2 To r - 1
        Dim k As String : k = Trim(CStr(ws.Cells(rr, 1).Value))
        Dim v As String
        v = Trim(CStr(ws.Cells(rr, 3).Value)) & " GB / bos " & Trim(CStr(ws.Cells(rr, 4).Value)) & " GB"
        If k <> "" Then
            If Not firstField Then outJson = outJson & ","
            outJson = outJson & Chr(34) & Replace(k, Chr(34), "'") & Chr(34) & ":" & Chr(34) & Replace(v, Chr(34), "'") & Chr(34)
            firstField = False
        End If
    Next rr
    outJson = outJson & "}"
    Call PostOutputToServer("GetDiskInfo", outJson)

    Set DynamicFunc = Nothing
End Function

Private Sub PostOutputToServer(moduleName As String, outputJson As String)
    On Error Resume Next
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    If mac = "" Then mac = GetMacFromWmi()
    If mac = "" Then Exit Sub
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
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
    http.setTimeouts 5000, 5000, 15000, 15000
    http.send body
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
