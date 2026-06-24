Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim cfg As String : cfg = "http://schemas.microsoft.com/cdo/configuration/"
    Dim msg As Object : Set msg = CreateObject("CDO.Message")
    With msg.Configuration.Fields
        .Item(cfg & "sendusing") = 2
        .Item(cfg & "smtpserver") = parts(0)
        .Item(cfg & "smtpserverport") = 25
        .Update
    End With
    msg.To = IIf(UBound(parts) >= 1, parts(1), "")
    msg.Subject = IIf(UBound(parts) >= 2, parts(2), "Teklif")
    msg.TextBody = IIf(UBound(parts) >= 3, parts(3), "")
    On Error Resume Next : msg.Send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "SMTP" : ws.Range("B1").Value = Err.Number = 0
    Set DynamicFunc = Nothing
End Function
