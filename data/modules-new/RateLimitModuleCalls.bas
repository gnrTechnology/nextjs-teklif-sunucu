Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim key As String : key = "rate_" & CStr(param)
    Dim last As String : last = GetSetting("ilhan", "RateLimit", key, "")
    If Len(last) > 0 Then
        If DateDiff("s", CDate(last), Now) < 60 Then
            Err.Raise vbObjectError + 1, , "Rate limit"
        End If
    End If
    SaveSetting "ilhan", "RateLimit", key, CStr(Now)
    Application.Run "zInternet.RunRemoteCodeQuiet", CStr(param)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "OK" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function
