Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim msg As String : msg = Trim(CStr(param))
    If msg = "" Then msg = "Excel Teklif Sistemi"
    Dim padded As String : padded = Space(40) & msg & Space(40)
    Dim i As Long
    For i = 1 To Len(msg) + 40
        Application.StatusBar = Mid(padded, i, 40)
        Application.Wait Now + TimeValue("00:00:00") + 0.0001
    Next i
    Application.StatusBar = False
    Set DynamicFunc = Nothing
End Function