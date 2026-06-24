Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim attached As Boolean : attached = False
    On Error Resume Next
    attached = (Application.VBE.ActiveVBProject.Name <> "")
    On Error GoTo 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Debugger" : ws.Range("B1").Value = attached
    Set DynamicFunc = Nothing
End Function
