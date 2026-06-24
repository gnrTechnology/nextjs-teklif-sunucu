Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String
    If InStr(CStr(param), "|") > 0 Then parts = Split(CStr(param), "|", 2) Else ReDim parts(0 To 1) : parts(0) = CStr(param) : parts(1) = ""
    Dim fileA As String : fileA = Trim$(parts(0))
    Dim fileB As String : fileB = Trim$(parts(1))
    If Len(fileA) = 0 Then fileA = Environ("LOCALAPPDATA") & "\TeklifAgent\reg-snap-a.txt"
    If Len(fileB) = 0 Then fileB = Environ("LOCALAPPDATA") & "\TeklifAgent\reg-snap-b.txt"
  Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim a As String, b As String
    If fso.FileExists(fileA) Then a = fso.OpenTextFile(fileA, 1).ReadAll
    If fso.FileExists(fileB) Then b = fso.OpenTextFile(fileB, 1).ReadAll
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Dosya A" : ws.Range("B1").Value = fileA
    ws.Range("A2").Value = "Dosya B" : ws.Range("B2").Value = fileB
    ws.Range("A3").Value = "Ayni" : ws.Range("B3").Value = (a = b)
    ws.Range("A4").Value = "Fark" : ws.Range("B4").Value = IIf(a = b, "-", "Farkli icerik")
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
