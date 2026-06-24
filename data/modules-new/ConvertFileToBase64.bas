Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim$(CStr(param))
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open
    stm.LoadFromFile fp
    Dim bytes() As Byte : bytes = stm.Read
    stm.Close
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.nodeTypedValue = bytes
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Dosya" : ws.Range("B1").Value = fp
    ws.Range("A2").Value = "Base64" : ws.Range("B2").Value = Left$(el.Text, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
