Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim uploadUrl As String : uploadUrl = Trim$(parts(0))
    Dim filePath As String : filePath = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.LoadFromFile filePath
    Dim bytes() As Byte : bytes = stm.Read : stm.Close
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "PUT", uploadUrl, False
    http.send bytes
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    Set DynamicFunc = Nothing
End Function
