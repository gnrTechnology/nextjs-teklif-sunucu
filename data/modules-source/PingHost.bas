Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim host As String : host = Trim(CStr(param))
    If host = "" Then host = "8.8.8.8"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Host" : ws.Range("B1").Value = host
    ws.Range("A1:B1").Font.Bold = True
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address='" & host & "'")
    For Each obj In col
        ws.Cells(r,1).Value = "Durum"       : ws.Cells(r,2).Value = IIf(obj.StatusCode=0,"✅ Başarılı","❌ Başarısız (" & obj.StatusCode & ")") : r=r+1
        ws.Cells(r,1).Value = "Yanıt ms"    : ws.Cells(r,2).Value = obj.ResponseTime & " ms" : r=r+1
        ws.Cells(r,1).Value = "TTL"         : ws.Cells(r,2).Value = obj.TimeToLive : r=r+1
        ws.Cells(r,1).Value = "Paket Boyutu": ws.Cells(r,2).Value = obj.BufferSize & " byte" : r=r+1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function