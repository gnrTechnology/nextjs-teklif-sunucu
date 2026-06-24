Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim mac As String : mac = Replace(UCase$(CStr(param)), ":", "")
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell -Command ""$m='" & mac & "'; $b=[byte[]](,0)*6; for($i=0;$i -lt 6;$i++){$b[$i]=0xFF}; for($i=0;$i -lt 16;$i+=2){$b+=[byte]('0x'+$m.Substring($i,2))}; $u=New-Object Net.Sockets.UdpClient; $u.Connect('255.255.255.255',9); $u.Send($b,$b.Length); $u.Close()""", 0, True
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "WOL" : ws.Range("B1").Value = mac
    Set DynamicFunc = Nothing
End Function
