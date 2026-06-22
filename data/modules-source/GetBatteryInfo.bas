Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, BatteryStatus, EstimatedChargeRemaining, EstimatedRunTime FROM Win32_Battery")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Pil Adı", "Durum", "Şarj (%)", "Tahmini Süre (dk)")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    Dim found As Boolean : found = False
    For Each obj In col
        found = True
        Dim bStatus As String
        Select Case obj.BatteryStatus
            Case 1 : bStatus = "Deşarj"
            Case 2 : bStatus = "AC - Şarj Değil"
            Case 3 : bStatus = "Tam Dolu"
            Case 4 : bStatus = "Düşük"
            Case 5 : bStatus = "Kritik"
            Case 6 : bStatus = "Şarj Oluyor"
            Case Else : bStatus = "Bilinmiyor"
        End Select
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = bStatus
        ws.Cells(r, 3).Value = obj.EstimatedChargeRemaining
        ws.Cells(r, 4).Value = IIf(obj.EstimatedRunTime = 71582788, "AC'de", obj.EstimatedRunTime)
        r = r + 1
    Next obj
    If Not found Then
        ws.Range("A2").Value = "Pil bulunamadı (masaüstü bilgisayar)"
    End If
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function