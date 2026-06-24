Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error GoTo Fail
    Dim parts() As String : parts = Split(CStr(param), "|", 3)
    Dim ol As Object : Set ol = CreateObject("Outlook.Application")
    Dim mail As Object : Set mail = ol.CreateItem(0)
    mail.To = Trim$(parts(0))
    If UBound(parts) >= 1 Then mail.Subject = parts(1)
    If UBound(parts) >= 2 Then mail.Body = parts(2)
    targetWb.Save : mail.Attachments.Add targetWb.FullName
    mail.Send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Gonderildi" : ws.Range("B1").Value = parts(0)
    GoTo Done
Fail: ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done: Set DynamicFunc = Nothing
End Function
