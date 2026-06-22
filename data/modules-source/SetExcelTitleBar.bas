Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim title As String : title = Trim(CStr(param))
    If title = "" Then title = "Excel - " & targetWb.Name
    Application.Caption = title
    targetWb.Sheets(1).Range("A1").Value = "✅ Başlık: " & title
    Set DynamicFunc = Nothing
End Function