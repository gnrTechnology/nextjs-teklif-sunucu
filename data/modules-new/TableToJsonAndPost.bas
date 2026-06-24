Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim apiUrl As String : apiUrl = Trim$(CStr(param))
    Application.Run "ConvertSheetToJson", targetWb, ""
    Dim json As String : json = CStr(targetWb.Sheets(1).Range("B1").Value)
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send json
    targetWb.Sheets(1).Range("A2").Value = "POST" : targetWb.Sheets(1).Range("B2").Value = http.Status
    Set DynamicFunc = Nothing
End Function
