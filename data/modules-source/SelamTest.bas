Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "[SelamTest] DynamicFunc basladi. Kitap: " & targetWb.Name
    MsgBox "selam"
    Debug.Print "[SelamTest] DynamicFunc tamamlandi."
    Set DynamicFunc = Nothing
End Function
