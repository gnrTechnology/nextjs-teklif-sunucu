Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cpu As String, mac As String
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\.ootcimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT ProcessorId FROM Win32_Processor")
    Dim o As Object : For Each o In col : cpu = o.ProcessorId : Exit For : Next
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col : If Not IsNull(o.MACAddress) Then mac = o.MACAddress : Exit For
    Dim raw As String : raw = UCase$(cpu & mac)
    Dim i As Long, hx As String : hx = ""
    For i = 1 To Len(raw)
        hx = hx & Right$("0" & Hex(Asc(Mid$(raw, i, 1)) Mod 256), 2)
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "HWID" : ws.Range("B1").Value = Left$(hx, 32)
    Set DynamicFunc = Nothing
End Function
