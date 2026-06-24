Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Dim voice As Object : Set voice = CreateObject("SAPI.SpVoice")
    voice.Speak CStr(param)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Okundu" : ws.Range("B1").Value = Left$(CStr(param), 200)
    Set DynamicFunc = Nothing
End Function
