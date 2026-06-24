Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    If Application.CalculationState <> xlDone Then
        Dim q As String : q = Environ("ProgramData") & "\TeklifAgent\queue\" & Format(Now, "yyyymmddhhnnss") & ".cmd"
        Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim ts As Object : Set ts = fso.CreateTextFile(q, True)
        ts.Write CStr(param) : ts.Close
    Else
        Application.Run "zInternet.RunRemoteCodeQuiet", CStr(param)
    End If
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Queued" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function
