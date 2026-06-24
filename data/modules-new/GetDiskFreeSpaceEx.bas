#If VBA7 Then
Private Declare PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExW" (ByVal lpDirectoryName As LongPtr, FreeBytesAvailableToCaller As Currency, TotalNumberOfBytes As Currency, TotalNumberOfFreeBytes As Currency) As Long
#Else
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExW" (ByVal lpDirectoryName As Long, FreeBytesAvailableToCaller As Currency, TotalNumberOfBytes As Currency, TotalNumberOfFreeBytes As Currency) As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim drive As String : drive = Trim$(CStr(param))
    If Len(drive) = 0 Then drive = Left$(Application.ActiveWorkbook.Path, 3)
    If Right$(drive, 1) <> "\" Then drive = drive & "\"
    Dim freeB As Currency, totalB As Currency, freeAll As Currency
    GetDiskFreeSpaceEx StrPtr(drive), freeB, totalB, freeAll
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Drive" : ws.Range("B1").Value = drive
    ws.Range("A2").Value = "FreeBytes" : ws.Range("B2").Value = freeB
    ws.Range("A3").Value = "TotalBytes" : ws.Range("B3").Value = totalB
    Set DynamicFunc = Nothing
End Function
