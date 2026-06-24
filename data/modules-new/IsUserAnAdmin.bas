Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Object : Set sh = CreateObject("Shell.Application")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Admin" : ws.Range("B1").Value = sh.IsUserAnAdmin
    Set DynamicFunc = Nothing
End Function
