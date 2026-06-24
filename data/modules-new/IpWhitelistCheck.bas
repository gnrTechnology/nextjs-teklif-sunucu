Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim allowed As String : allowed = LCase$(CStr(param))
    Dim ip As String : ip = ""
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\.ootcimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim o As Object : For Each o In col : ip = o.IPAddress(0) : Exit For
    Dim ok As Boolean : ok = (InStr(allowed, LCase$(ip)) > 0)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "IP" : ws.Range("B1").Value = ip
    ws.Range("A2").Value = "Izinli" : ws.Range("B2").Value = ok
    Set DynamicFunc = Nothing
End Function
