Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "GenerateQrCodeImage", targetWb, param
    Set DynamicFunc = Nothing
End Function
