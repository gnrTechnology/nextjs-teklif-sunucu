Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Version, Vendor, InstallDate FROM Win32_Product")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Yazılım", "Sürüm", "Üretici", "Kurulum Tarihi")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.Version
        ws.Cells(r, 3).Value = obj.Vendor
        Dim dt As String : dt = CStr(obj.InstallDate)
        If Len(dt) >= 8 Then dt = Left(dt, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2)
        ws.Cells(r, 4).Value = dt
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    MsgBox (r - 2) & " yazılım listelendi.", vbInformation, "GetInstalledSoftwareList"
    Set DynamicFunc = Nothing
End Function