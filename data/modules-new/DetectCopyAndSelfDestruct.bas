Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim expected As String : expected = GetSetting("ilhan", "Settings", "startingAddin", "")
    Dim cur As String : cur = Application.AddIns(1).Name
    Dim violation As Boolean : violation = (Len(expected) > 0 And InStr(cur, expected) = 0)
    If violation Then targetWb.Close SaveChanges:=False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Ihlal" : ws.Range("B1").Value = violation
    Set DynamicFunc = Nothing
End Function
