Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Özellik" : ws.Range("B1").Value = "Değer"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT TotalVirtualMemorySize, FreeVirtualMemory, TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem")
    For Each obj In col
        ws.Cells(r,1).Value = "Toplam Sanal Bellek"     : ws.Cells(r,2).Value = Format(CLng(obj.TotalVirtualMemorySize)/1024,"#,##0") & " MB" : r=r+1
        ws.Cells(r,1).Value = "Boş Sanal Bellek"        : ws.Cells(r,2).Value = Format(CLng(obj.FreeVirtualMemory)/1024,"#,##0") & " MB"     : r=r+1
        ws.Cells(r,1).Value = "Toplam Fiziksel Bellek"  : ws.Cells(r,2).Value = Format(CLng(obj.TotalVisibleMemorySize)/1024,"#,##0") & " MB" : r=r+1
        ws.Cells(r,1).Value = "Boş Fiziksel Bellek"     : ws.Cells(r,2).Value = Format(CLng(obj.FreePhysicalMemory)/1024,"#,##0") & " MB"    : r=r+1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function