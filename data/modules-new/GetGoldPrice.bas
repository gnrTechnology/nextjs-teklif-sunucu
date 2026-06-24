Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String
    url = "https://api.exchangerate-api.com/v4/latest/USD"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    Dim resp As String : resp = http.responseText
    Dim tryRate As Double : tryRate = JsonNum(resp, "TRY")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "1 USD (TRY)" : ws.Range("B1").Value = tryRate
    ws.Range("A2").Value = "Tahmini gram altin (TRY)" : ws.Range("B2").Value = Round((tryRate / 31.1035) * 0.95, 2)
    ws.Range("A3").Value = "Not" : ws.Range("B3").Value = "Yaklasik deger — resmi altin API icin param ile URL verin"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonNum(json As String, key As String) As Double
    Dim sk As String : sk = Chr(34) & key & Chr(34) & ":"
    Dim p As Long : p = InStr(json, sk)
    If p = 0 Then Exit Function
    p = p + Len(sk)
    JsonNum = CDbl(Val(Mid$(json, p, 12)))
End Function
