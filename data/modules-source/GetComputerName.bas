Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim name As String : name = Environ("COMPUTERNAME")
    targetWb.Sheets(1).Range("A1").Value = "Bilgisayar Adı"
    targetWb.Sheets(1).Range("B1").Value = name
    targetWb.Sheets(1).Columns.AutoFit
    MsgBox "Bilgisayar Adı: " & name, vbInformation, "GetComputerName"
    Set DynamicFunc = Nothing
End Function