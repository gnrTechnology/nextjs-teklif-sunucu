Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim hidden As Worksheet
    On Error Resume Next
    Set hidden = targetWb.Worksheets("_fx")
    If hidden Is Nothing Then Set hidden = targetWb.Worksheets.Add : hidden.Name = "_fx" : hidden.Visible = xlSheetVeryHidden
    hidden.Range("A1").Value = targetWb.ActiveSheet.UsedRange.Formula
    targetWb.ActiveSheet.UsedRange.Value = targetWb.ActiveSheet.UsedRange.Value
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Gizlendi" : ws.Range("B1").Value = "_fx"
    Set DynamicFunc = Nothing
End Function
