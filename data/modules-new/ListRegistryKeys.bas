Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim rootPath As String
    rootPath = Trim$(CStr(param))
    If Len(rootPath) = 0 Then rootPath = "HKCU\Software\ilhan"

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Anahtar"
    ws.Range("B1").Value = "Alt anahtar"
    ws.Range("C1").Value = "Deger adi"
    ws.Range("D1").Value = "Deger"

    Dim row As Long
    row = 2
    Call WalkRegistry(shell, rootPath, rootPath, ws, row)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function

Private Sub WalkRegistry(shell As Object, displayRoot As String, current As String, ws As Worksheet, ByRef row As Long)
    On Error Resume Next
    Dim subKeys As Variant
    subKeys = shell.RegRead(current & "\")
    If Err.Number = 0 And IsArray(subKeys) Then
        Dim i As Long
        For i = LBound(subKeys) To UBound(subKeys)
            If row > 500 Then Exit Sub
            WalkRegistry shell, displayRoot, current & "\" & subKeys(i), ws, row
        Next i
    End If
    Err.Clear

    Dim values As Variant
    values = shell.RegRead(current)
    If Err.Number = 0 And IsArray(values) Then
        Dim j As Long
        For j = LBound(values) To UBound(values)
            If row > 500 Then Exit Sub
            If Len(values(j)) > 0 Then
                ws.Cells(row, 1).Value = displayRoot
                ws.Cells(row, 2).Value = Mid$(current, Len(displayRoot) + 2)
                ws.Cells(row, 3).Value = values(j)
                On Error Resume Next
                ws.Cells(row, 4).Value = CStr(shell.RegRead(current & "\" & values(j)))
                If Err.Number <> 0 Then ws.Cells(row, 4).Value = "(okunamadi)"
                Err.Clear
                row = row + 1
            End If
        Next j
    End If
    Err.Clear
End Sub
