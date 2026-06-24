Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim msg As String : msg = CStr(param)
    Dim ans As Long : ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "Teklif")
    Dim result As String
    If ans = vbYes Then result = "Yes" ElseIf ans = vbNo Then result = "No" Else result = "Cancel"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Secim" : ws.Range("B1").Value = result
    Set DynamicFunc = result
End Function
