Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim txt As String : txt = CStr(param) : If Len(txt) = 0 Then txt = "GIZLI"
    With targetWb.ActiveSheet.PageSetup
        .CenterFooter = txt
        .PrintTitleRows = ""
    End With
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Baski filigran" : ws.Range("B1").Value = txt
    Set DynamicFunc = Nothing
End Function
