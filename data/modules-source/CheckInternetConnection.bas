Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "http://www.msftconnecttest.com/connecttest.txt", False
    http.setTimeouts 3000, 5000, 8000, 8000
    http.send
    Dim ok As Boolean : ok = (http.Status = 200)
    On Error GoTo 0
    ws.Range("A1").Value = "İnternet Bağlantısı"
    ws.Range("B1").Value = IIf(ok, "✅ Bağlı", "❌ Bağlı Değil")
    ' Genel gecikme (ping-benzeri)
    Dim t1 As Double : t1 = Timer
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://api.ipify.org", False
    http.setTimeouts 3000, 5000, 8000, 8000
    http.send
    Dim t2 As Double : t2 = Timer
    On Error GoTo 0
    ws.Range("A2").Value = "Yanıt Süresi" : ws.Range("B2").Value = Format((t2-t1)*1000,"0") & " ms"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function