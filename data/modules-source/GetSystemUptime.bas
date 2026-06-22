Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Alan", "Değer")
    ws.Range("A1:B1").Font.Bold = True
    For Each obj In col
        Dim raw As String : raw = obj.LastBootUpTime
        Dim bootYear  As Integer : bootYear  = CInt(Left(raw, 4))
        Dim bootMonth As Integer : bootMonth = CInt(Mid(raw, 5, 2))
        Dim bootDay   As Integer : bootDay   = CInt(Mid(raw, 7, 2))
        Dim bootHour  As Integer : bootHour  = CInt(Mid(raw, 9, 2))
        Dim bootMin   As Integer : bootMin   = CInt(Mid(raw, 11, 2))
        Dim bootSec   As Integer : bootSec   = CInt(Mid(raw, 13, 2))
        Dim bootDT As Date
        bootDT = DateSerial(bootYear, bootMonth, bootDay) + TimeSerial(bootHour, bootMin, bootSec)
        Dim diff  As Double : diff  = Now - bootDT
        Dim days  As Long   : days  = Int(diff)
        Dim hours As Long   : hours = Int((diff - days) * 24)
        Dim mins  As Long   : mins  = Int(((diff - days) * 24 - hours) * 60)
        ws.Range("A2:B2").Value = Array("Son Açılış", Format(bootDT, "dd.mm.yyyy HH:MM:SS"))
        ws.Range("A3:B3").Value = Array("Çalışma Süresi", days & " gün " & hours & " saat " & mins & " dk")
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function