Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim outDir As String : outDir = Environ("TEMP") & "\split\"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    Dim dict As Object : Set dict = CreateObject("Scripting.Dictionary")
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long
    For r = 2 To ur.Rows.Count
        Dim k As String : k = CStr(ur.Cells(r, col).Value)
        If Not dict.Exists(k) Then dict.Add k, 1
    Next r
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Grup" : ws.Range("B1").Value = dict.Count
    Set DynamicFunc = Nothing
End Function
