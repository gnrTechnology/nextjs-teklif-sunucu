Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), ",")
    Dim freq As Long : freq = IIf(Len(Trim$(parts(0))) > 0, CLng(Val(parts(0))), 800)
    Dim dur As Long : dur = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 200)
    Beep freq, dur
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Beep" : ws.Range("B1").Value = freq & "Hz " & dur & "ms"
    Set DynamicFunc = Nothing
End Function
