Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Özellik" : ws.Range("B1").Value = "Değer"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT Name, IdentifyingNumber, Vendor, Version, UUID FROM Win32_ComputerSystemProduct")
    For Each obj In col
        ws.Cells(r,1).Value = "Ürün Adı"      : ws.Cells(r,2).Value = obj.Name             : r=r+1
        ws.Cells(r,1).Value = "Seri No"        : ws.Cells(r,2).Value = obj.IdentifyingNumber : r=r+1
        ws.Cells(r,1).Value = "Üretici"        : ws.Cells(r,2).Value = obj.Vendor           : r=r+1
        ws.Cells(r,1).Value = "Sürüm"          : ws.Cells(r,2).Value = obj.Version          : r=r+1
        ws.Cells(r,1).Value = "UUID"            : ws.Cells(r,2).Value = obj.UUID             : r=r+1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function