Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim hivePath As String
    Dim outFile As String
    Dim parts() As String
    hivePath = Trim$(CStr(param))
    outFile = Environ("TEMP") & "\registry-export.reg"
    If InStr(hivePath, "|") > 0 Then
        parts = Split(hivePath, "|", 2)
        hivePath = Trim$(parts(0))
        outFile = Trim$(parts(1))
    End If
    If Len(hivePath) = 0 Then hivePath = "HKCU\Software\ilhan"

    Dim cmd As String
    cmd = "reg export """ & hivePath & """ """ & outFile & """ /y"
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run(cmd, 0, True)

    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Kaynak"
    ws.Range("B1").Value = hivePath
    ws.Range("A2").Value = "Dosya"
    ws.Range("B2").Value = outFile
    ws.Range("A3").Value = "Sonuc"
    ws.Range("B3").Value = IIf(rc = 0, "OK", "HATA kod " & rc)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
