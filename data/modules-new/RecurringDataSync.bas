Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "ExportSheetToServer", targetWb, CStr(param)
    Set DynamicFunc = Nothing
End Function
