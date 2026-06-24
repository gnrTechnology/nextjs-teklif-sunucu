Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim toAddr As String : toAddr = Trim$(parts(0))
    Dim subj As String : subj = IIf(UBound(parts) >= 1, parts(1), "Teklif")
    On Error GoTo Fail
    Dim ol As Object : Set ol = CreateObject("Outlook.Application")
    Dim mail As Object : Set mail = ol.CreateItem(0)
    mail.To = toAddr : mail.Subject = subj
    targetWb.Save
    mail.Attachments.Add targetWb.FullName
    mail.Send
    ws.Range("A1").Value = "Gonderildi" : ws.Range("B1").Value = toAddr
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing
End Function
