Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String
    fp = Trim(CStr(param))
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim exists As Boolean
    exists = fso.FileExists(fp)

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    If ws Is Nothing Then Set ws = targetWb.ActiveSheet
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Range("A1").Value = "Dosya Yolu"
        ws.Range("B1").Value = "Mevcut"
        ws.Range("A2").Value = fp
        ws.Range("B2").Value = IIf(exists, "Evet", "Hayir")
        ws.Range("A1:B2").Columns.AutoFit
    End If

    Debug.Print "[FileExists] " & fp & Chr(10) & IIf(exists, "Evet", "Hayir")
    Set DynamicFunc = Nothing
End Function
