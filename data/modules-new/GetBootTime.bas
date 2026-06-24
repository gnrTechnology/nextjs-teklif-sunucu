Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    Dim o As Object : For Each o In col
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Boot" : ws.Range("B1").Value = o.LastBootUpTime : Exit For
    Next
    Set DynamicFunc = Nothing
End Function
