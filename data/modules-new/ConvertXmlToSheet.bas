Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim xmlPath As String : xmlPath = Trim$(CStr(param))
    Dim dom As Object : Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.Load xmlPath
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "XML" : ws.Range("B1").Value = dom.documentElement.nodeName
    ws.Range("A2").Value = "Text" : ws.Range("B2").Value = Left$(dom.documentElement.Text, 32000)
    Set DynamicFunc = Nothing
End Function
