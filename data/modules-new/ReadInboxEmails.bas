Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error GoTo Fail
    Dim n As Long : n = CLng(Val(CStr(param))) : If n < 1 Then n = 10
    Dim ol As Object : Set ol = CreateObject("Outlook.Application")
    Dim ns As Object : Set ns = ol.GetNamespace("MAPI")
    Dim inbox As Object : Set inbox = ns.GetDefaultFolder(6)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Cells.ClearContents
    ws.Range("A1").Value = "Konu" : ws.Range("B1").Value = "Gonderen"
    Dim it As Object, r As Long : r = 2, i As Long : i = 0
    For Each it In inbox.Items
        i = i + 1 : If i > n Then Exit For
        ws.Cells(r, 1).Value = it.Subject : ws.Cells(r, 2).Value = it.SenderEmailAddress : r = r + 1
    Next
    GoTo Done
Fail: ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done: Set DynamicFunc = Nothing
End Function
