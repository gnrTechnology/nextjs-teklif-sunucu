Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim b64 As String, outPath As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    b64 = parts(0) : outPath = IIf(UBound(parts) >= 1, Trim$(parts(1)), Environ("TEMP") & "\decoded.bin")
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.Text = b64
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.Write el.nodeTypedValue
    stm.SaveToFile outPath, 2 : stm.Close
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Cikti" : ws.Range("B1").Value = outPath
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
