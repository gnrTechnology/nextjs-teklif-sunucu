Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "AdoQueryToSheet", targetWb, CStr(param) & "|SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES"
    Set DynamicFunc = Nothing
End Function
