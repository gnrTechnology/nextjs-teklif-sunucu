Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String : path = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    Dim val As String : val = wsh.RegRead(path)
    Dim errTxt As String : If Err.Number <> 0 Then errTxt = "HATA: " & Err.Description
    On Error GoTo 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Yol"   : ws.Range("B1").Value = path
    ws.Range("A2").Value = "Değer" : ws.Range("B2").Value = IIf(errTxt <> "", errTxt, val)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function