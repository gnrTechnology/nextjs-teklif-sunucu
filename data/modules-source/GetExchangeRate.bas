Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Döviz" : ws.Range("B1").Value = "Alış" : ws.Range("C1").Value = "Satış"
    ws.Range("A1:C1").Font.Bold = True
    ' exchangerate-api (ücretsiz, kayıt gerektirmeyen)
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://api.exchangerate-api.com/v4/latest/TRY", False
    http.setTimeouts 5000, 10000, 20000, 20000
    http.send
    If http.Status = 200 Then
        Dim resp As String : resp = http.responseText
        ' Basit JSON parse - rates içinden USD, EUR, GBP
        Dim currencies() As String : currencies = Split("USD,EUR,GBP,CHF,JPY", ",")
        Dim r As Long : r = 2
        Dim c As Variant
        For Each c In currencies
            Dim key As String : key = Chr(34) & c & Chr(34) & ":"
            Dim p1 As Long : p1 = InStr(resp, key)
            If p1 > 0 Then
                p1 = p1 + Len(key)
                Dim p2 As Long : p2 = p1
                Do While InStr(",}", Mid(resp,p2,1)) = 0 : p2=p2+1 : Loop
                Dim rate As Double : rate = CDbl(Trim(Mid(resp,p1,p2-p1)))
                ws.Cells(r,1).Value = c
                ws.Cells(r,2).Value = Format(1/rate,"0.0000") & " TL"
                r = r + 1
            End If
        Next c
        ws.Cells(r,1).Value = "Güncelleme" : ws.Cells(r,2).Value = Now()
    Else
        ws.Range("A2").Value = "Hata: HTTP " & http.Status
    End If
    On Error GoTo 0
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function