Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String
    fp = Trim(CStr(param))
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fp) Then
        Debug.Print "[ReadCsvToSheet] Dosya bulunamadi: " & fp
        Set DynamicFunc = Nothing
        Exit Function
    End If

    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim fNum As Integer
    fNum = FreeFile
    Open fp For Input As #fNum
    Dim r As Long
    r = 1
    Dim rowLine As String
    Do While Not EOF(fNum)
        Line Input #fNum, rowLine
        Dim cols() As String
        cols = Split(rowLine, ",")
        Dim c As Integer
        For c = 0 To UBound(cols)
            ws.Cells(r, c + 1).Value = Replace(cols(c), """", "")
        Next c
        r = r + 1
        If r > 5000 Then Exit Do
    Loop
    Close #fNum
    ws.Rows(1).Font.Bold = True
    ws.UsedRange.Columns.AutoFit
    Debug.Print "[ReadCsvToSheet] " & (r - 1) & " satir aktarildi."
    Set DynamicFunc = Nothing
    Exit Function
Fail:
    Debug.Print "[ReadCsvToSheet] Hata: " & Err.Description
    Set DynamicFunc = Nothing
End Function
