Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "[getLicense] DynamicFunc basladi. Kitap: " & targetWb.Name
    Debug.Print "[getLicense] zInternet.TestInternetConnection cagriliyor..."
    Application.Run "zInternet.TestInternetConnection"
    Debug.Print "[getLicense] zInternet.TestInternetConnection tamamlandi."
    Set DynamicFunc = Nothing
End Function
