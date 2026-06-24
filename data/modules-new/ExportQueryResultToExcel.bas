Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "AdoQueryToSheet", targetWb, param
    targetWb.SaveAs Environ("TEMP") & "\query-result.xlsx"
    Set DynamicFunc = Nothing
End Function
