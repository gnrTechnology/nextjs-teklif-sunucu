Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cmd As String
    cmd = Trim$(CStr(param))
    If Len(cmd) = 0 Then
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    On Error Resume Next
    sh.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce\" & cmd
    sh.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce\" & cmd
    Dim errN As Long
    errN = Err.Number
    Dim errD As String
    errD = Err.Description
    Err.Clear
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "RunOnce adi"
    ws.Range("B1").Value = cmd
    ws.Range("A2").Value = "Sonuc"
    If errN = 0 Then
        ws.Range("B2").Value = "Silindi (HKCU/HKLM)"
    Else
        ws.Range("B2").Value = "HATA: " & errD
    End If
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
