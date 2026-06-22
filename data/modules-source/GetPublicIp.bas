Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error Resume Next
    http.Open "GET", "https://api.ipify.org", False
    http.setTimeouts 5000, 10000, 15000, 15000
    http.send
    Dim ip As String
    If http.Status = 200 Then
        ip = Trim(http.responseText)
    Else
        ip = "Erişilemedi"
    End If
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = "Dış IP Adresi"
    targetWb.Sheets(1).Range("B1").Value = ip
    targetWb.Sheets(1).Range("A1").Font.Bold = True
    targetWb.Sheets(1).Columns.AutoFit
    MsgBox "Dış IP: " & ip, vbInformation, "GetPublicIp"
    Set DynamicFunc = Nothing
End Function