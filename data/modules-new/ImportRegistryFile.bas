Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim regFile As String
    regFile = Trim$(CStr(param))
    If Len(regFile) = 0 Then
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("reg import """ & regFile & """", 0, True)

    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Dosya"
    ws.Range("B1").Value = regFile
    ws.Range("A2").Value = "Sonuc"
    ws.Range("B2").Value = IIf(rc = 0, "Ice aktarildi", "HATA kod " & rc)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
