Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Application.Run "GetDiskFreeSpaceEx", targetWb, param
    If Err.Number <> 0 Then Application.Run "GetDiskInfo", targetWb, param
    Set DynamicFunc = Nothing
End Function
