Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String, timeoutSec As Long
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    folderPath = Trim$(parts(0))
    timeoutSec = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 60)
    If Right(folderPath, 1) <> Chr(92) Then folderPath = folderPath & Chr(92)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim start As Double : start = Timer
    Dim found As String : found = ""
    Do While Timer - start < timeoutSec
        Dim fn As String : fn = Dir(folderPath & "*.*")
        Do While Len(fn) > 0
            If fn <> "." And fn <> ".." Then found = folderPath & fn : Exit Do
            fn = Dir()
        Loop
        If Len(found) > 0 Then Exit Do
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Klasor" : ws.Range("B1").Value = folderPath
    ws.Range("A2").Value = "Bulunan" : ws.Range("B2").Value = IIf(Len(found) > 0, found, "(yok)")
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
