Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim savePath As String : savePath = Environ("TEMP") & "\downloaded_" & CLng(Timer*1000) & ".xlsx"
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.setTimeouts 5000,10000,60000,60000 : http.send
    If http.Status = 200 Then
        Dim st As Object : Set st = CreateObject("ADODB.Stream")
        st.Type = 1 : st.Open : st.Write http.responseBody
        st.SaveToFile savePath, 2 : st.Close
        Workbooks.Open savePath
        targetWb.Sheets(1).Range("A1").Value = "✅ Açıldı: " & savePath
    Else
        targetWb.Sheets(1).Range("A1").Value = "❌ HTTP " & http.Status
    End If
    On Error GoTo 0
    Set DynamicFunc = Nothing
End Function