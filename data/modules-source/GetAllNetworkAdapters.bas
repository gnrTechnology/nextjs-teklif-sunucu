Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Description, MACAddress, IPAddress, IPEnabled, DHCPEnabled FROM Win32_NetworkAdapterConfiguration")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Adaptör", "MAC", "IP Adresi", "IP Etkin", "DHCP")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        Dim ipStr As String : ipStr = ""
        If Not IsNull(obj.IPAddress) Then
            Dim ip As Variant
            For Each ip In obj.IPAddress
                ipStr = ipStr & ip & " "
            Next ip
        End If
        ws.Cells(r, 1).Value = obj.Description
        ws.Cells(r, 2).Value = IIf(IsNull(obj.MACAddress), "-", obj.MACAddress)
        ws.Cells(r, 3).Value = Trim(ipStr)
        ws.Cells(r, 4).Value = IIf(obj.IPEnabled, "Evet", "Hayır")
        ws.Cells(r, 5).Value = IIf(obj.DHCPEnabled, "Evet", "Hayır")
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function