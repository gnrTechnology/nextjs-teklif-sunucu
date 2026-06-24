Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim vm As Boolean : vm = False
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\.ootcimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT Model,Manufacturer FROM Win32_ComputerSystem")
    Dim o As Object
    For Each o In col
        Dim s As String : s = UCase$(o.Model & " " & o.Manufacturer)
        If InStr(s, "VMWARE") > 0 Or InStr(s, "VIRTUAL") > 0 Or InStr(s, "HYPER-V") > 0 Or InStr(s, "QEMU") > 0 Then vm = True
    Next
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "VM" : ws.Range("B1").Value = vm
    Set DynamicFunc = Nothing
End Function
